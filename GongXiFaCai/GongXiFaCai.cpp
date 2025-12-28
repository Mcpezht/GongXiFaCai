#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <filesystem>
#include <iostream>
#include <comdef.h>
#include <uiautomationclient.h>
#include <tlhelp32.h>

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "Advapi32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Uiautomationcore.lib")

using namespace std;
namespace fs = std::filesystem;

bool IsRunningElevated()
{
    BOOL fRet = FALSE;
    HANDLE hToken = NULL;
    if (OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken))
    {
        TOKEN_ELEVATION Elevation;
        DWORD cbSize = sizeof(TOKEN_ELEVATION);
        if (GetTokenInformation(hToken, TokenElevation, &Elevation, sizeof(Elevation), &cbSize))
        {
            fRet = Elevation.TokenIsElevated;
        }
    }
    if (hToken)
        CloseHandle(hToken);
    return fRet != FALSE;
}

void RelaunchElevated()
{
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(NULL, path, MAX_PATH);
    SHELLEXECUTEINFOW sei = { 0 };
    sei.cbSize = sizeof(sei);
    sei.lpVerb = L"runas";
    sei.lpFile = path;
    sei.nShow = SW_SHOWNORMAL;
    if (!ShellExecuteExW(&sei))
    {
    }
}

void SendUnicodeString(const wstring &s)
{
    std::vector<INPUT> inputs;
    for (wchar_t ch : s)
    {
        INPUT in = { 0 };
        in.type = INPUT_KEYBOARD;
        in.ki.wScan = ch;
        in.ki.dwFlags = KEYEVENTF_UNICODE;
        inputs.push_back(in);
        INPUT out = { 0 };
        out.type = INPUT_KEYBOARD;
        out.ki.wScan = ch;
        out.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
        inputs.push_back(out);
    }
    SendInput((UINT)inputs.size(), inputs.data(), sizeof(INPUT));
}

void SendKey(WORD vk, bool down)
{
    INPUT in = { 0 };
    in.type = INPUT_KEYBOARD;
    in.ki.wVk = vk;
    in.ki.dwFlags = down ? 0 : KEYEVENTF_KEYUP;
    SendInput(1, &in, sizeof(in));
}

HWND FindTopLevelWindowForProcess(DWORD pid)
{
    struct EnumData { DWORD pid; HWND result; } data{ pid, NULL };
    EnumWindows([](HWND hwnd, LPARAM lParam)->BOOL {
        EnumData* d = (EnumData*)lParam;
        DWORD pid = 0;
        GetWindowThreadProcessId(hwnd, &pid);
        if (pid == d->pid)
        {
            if (IsWindowVisible(hwnd) && GetWindow(hwnd, GW_OWNER) == NULL)
            {
                d->result = hwnd;
                return FALSE;
            }
        }
        return TRUE;
    }, (LPARAM)&data);
    return data.result;
}

wstring GetDesktopPath()
{
    PWSTR path = NULL;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Desktop, 0, NULL, &path)))
    {
        wstring p(path);
        CoTaskMemFree(path);
        return p;
    }
    return L"";
}

bool LaunchEdgeShortcutFound(const wstring &desktop)
{
    for (auto &entry : fs::directory_iterator(desktop))
    {
        if (!entry.is_regular_file()) continue;
        wstring name = entry.path().filename().wstring();
        // case-insensitive search for "Microsoft Edge"
        wstring lower = name;
        for (auto &c : lower) c = towlower(c);
        if (lower.find(L"microsoft edge") != wstring::npos && entry.path().extension() == L".lnk")
        {
            // shell execute the .lnk will open edge
            ShellExecuteW(NULL, L"open", entry.path().c_str(), NULL, NULL, SW_SHOWNORMAL);
            return true;
        }
    }
    return false;
}

bool ClickElement(IUIAutomationElement *elem)
{
    if (!elem) return false;
    IUIAutomationInvokePattern *invoke = nullptr;
    HRESULT hr = elem->GetCurrentPatternAs(UIA_InvokePatternId, IID_PPV_ARGS(&invoke));
    if (SUCCEEDED(hr) && invoke)
    {
        invoke->Invoke();
        invoke->Release();
        return true;
    }
    VARIANT var;
    VariantInit(&var);
    hr = elem->GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, &var);
    if (SUCCEEDED(hr) && var.vt == (VT_ARRAY | VT_R8))
    {
        SAFEARRAY *sa = var.parray;
        double *data = nullptr;
        if (SUCCEEDED(SafeArrayAccessData(sa, (void**)&data)) && data)
        {
            double left = data[0];
            double top = data[1];
            double width = data[2];
            double height = data[3];
            SafeArrayUnaccessData(sa);
            VariantClear(&var);
            int x = (int)(left + width / 2);
            int y = (int)(top + height / 2);
            // simulate mouse click
            SetCursorPos(x, y);
            INPUT inputs[2] = {};
            inputs[0].type = INPUT_MOUSE;
            inputs[0].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
            inputs[1].type = INPUT_MOUSE;
            inputs[1].mi.dwFlags = MOUSEEVENTF_LEFTUP;
            SendInput(2, inputs, sizeof(INPUT));
            return true;
        }
    }
    VariantClear(&var);
    return false;
}

bool FindAndClickByName(IUIAutomation *automation, HWND hwnd, const wstring &name)
{
    if (!automation) return false;
    HRESULT hr;
    IUIAutomationElement *root = NULL;
    hr = automation->ElementFromHandle(hwnd, &root);
    if (FAILED(hr) || !root) return false;
    VARIANT varName;
    varName.vt = VT_BSTR;
    varName.bstrVal = SysAllocStringLen(name.c_str(), (UINT)name.size());
    IUIAutomationCondition *cond = NULL;
    hr = automation->CreatePropertyCondition(UIA_NamePropertyId, varName, &cond);
    VariantClear(&varName);
    if (FAILED(hr) || !cond) { root->Release(); return false; }
    IUIAutomationElement *found = NULL;
    hr = root->FindFirst(TreeScope_Subtree, cond, &found);
    cond->Release();
    root->Release();
    if (FAILED(hr) || !found) return false;
    bool ok = ClickElement(found);
    found->Release();
    return ok;
}

struct LangSelection { int choice; };

LRESULT CALLBACK LangWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    LangSelection* data = (LangSelection*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    switch (msg)
    {
    case WM_CREATE:
        {
            CREATESTRUCTW* cs = (CREATESTRUCTW*)lParam;
            if (cs && cs->lpCreateParams)
            {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)cs->lpCreateParams);
                data = (LangSelection*)cs->lpCreateParams;
            }
            CreateWindowW(L"STATIC", L"请选择语言 / Please choose language:",
                WS_VISIBLE | WS_CHILD | SS_CENTER, 10, 10, 360, 20, hwnd, NULL, NULL, NULL);
            CreateWindowW(L"BUTTON", L"中文", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                60, 50, 100, 30, hwnd, (HMENU)1001, NULL, NULL);
            CreateWindowW(L"BUTTON", L"English", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
                200, 50, 100, 30, hwnd, (HMENU)1002, NULL, NULL);
        }
        break;
    case WM_COMMAND:
        if (LOWORD(wParam) == 1001 || LOWORD(wParam) == 1002)
        {
            if (data)
                data->choice = (LOWORD(wParam) == 1001) ? 1 : 2;
            PostMessageW(hwnd, WM_CLOSE, 0, 0);
        }
        break;
    case WM_CLOSE:
        DestroyWindow(hwnd);
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR pCmdLine, int nCmdShow)
{
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    if (!IsRunningElevated())
    {
        RelaunchElevated();
        return 0;
    }

    bool chinese = true;
    {
        LangSelection* sel = new LangSelection(); sel->choice = 0;
        WNDCLASSW wc = { 0 };
        wc.lpfnWndProc = LangWndProc;
        wc.hInstance = hInstance;
        wc.lpszClassName = L"LangClass";
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        RegisterClassW(&wc);

        HWND langWnd = CreateWindowExW(0, L"LangClass", L"选择语言", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
            CW_USEDEFAULT, CW_USEDEFAULT, 380, 140, NULL, NULL, hInstance, sel);
        if (langWnd)
        {
            RECT rc = { 0 };
            GetWindowRect(langWnd, &rc);
            int w = rc.right - rc.left;
            int h = rc.bottom - rc.top;
            int sx = (GetSystemMetrics(SM_CXSCREEN) - w) / 2;
            int sy = (GetSystemMetrics(SM_CYSCREEN) - h) / 2;
            SetWindowPos(langWnd, NULL, sx, sy, 0, 0, SWP_NOZORDER | SWP_NOSIZE);

            ShowWindow(langWnd, SW_SHOW);
            UpdateWindow(langWnd);

            MSG msg;
            while (GetMessage(&msg, NULL, 0, 0))
            {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }

            chinese = (sel->choice == 1);
        }
        else
        {
            chinese = true;
        }

        UnregisterClassW(L"LangClass", hInstance);
        delete sel;
    }

    INPUT inputs[4] = {};
    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = VK_LWIN;
    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = 'D';
    inputs[2] = inputs[1]; inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;
    inputs[3] = inputs[0]; inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(4, inputs, sizeof(INPUT));

    wstring desktop = GetDesktopPath();
    if (!desktop.empty())
    {
        LaunchEdgeShortcutFound(desktop);
    }

    Sleep(2000);

    DWORD edgePid = 0;
    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snap != INVALID_HANDLE_VALUE)
    {
        PROCESSENTRY32W pe;
        ZeroMemory(&pe, sizeof(pe));
        pe.dwSize = sizeof(pe);
        if (Process32FirstW(snap, &pe))
        {
            do {
                wstring exe = pe.szExeFile;
                for (auto &c : exe) c = towlower(c);
                if (exe == L"msedge.exe" || exe == L"microsoftedge.exe")
                {
                    edgePid = pe.th32ProcessID;
                    break;
                }
            } while (Process32NextW(snap, &pe));
        }
        CloseHandle(snap);
    }

    HWND edgeHwnd = NULL;
    if (edgePid != 0)
    {
        edgeHwnd = FindTopLevelWindowForProcess(edgePid);
        if (edgeHwnd) SetForegroundWindow(edgeHwnd);
    }

    SendKey(VK_CONTROL, true);
    SendKey('L', true);
    SendKey('L', false);
    SendKey(VK_CONTROL, false);
    Sleep(100);

    wstring url = chinese ? L"https://www.bilibili.com/video/BV1ad4y1V7wb" : L"https://www.bilibili.com/video/BV1AHmRBCEsq";
    SendUnicodeString(url);
    SendKey(VK_RETURN, true);
    SendKey(VK_RETURN, false);

    Sleep(800);

    IUIAutomation *automation = NULL;
    HRESULT hr = CoCreateInstance(CLSID_CUIAutomation, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&automation));

    if (automation && edgeHwnd)
    {
        const int maxAttempts = 6;
        auto tryClick = [&](const wchar_t* name)->bool {
            for (int i = 0; i < maxAttempts; i++) {
                SetForegroundWindow(edgeHwnd);
                Sleep(150 + i * 80);
                if (FindAndClickByName(automation, edgeHwnd, name)) return true;
            }
            return false;
        };

        Sleep(2000);

        bool ok = tryClick(L"从头播放");
        Sleep(500);
        bool ok2 = tryClick(L"点击恢复音量");

        automation->Release();
    }

    if (edgeHwnd) SetForegroundWindow(edgeHwnd);
    Sleep(50);
    SendKey('F', true);
    SendKey('F', false);

    CoUninitialize();
    return 0;
}

int main()
{
    return wWinMain(GetModuleHandle(NULL), NULL, GetCommandLineW(), SW_SHOW);
}
