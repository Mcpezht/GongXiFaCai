// GongXiFaCai.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include <iostream>
#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <tlhelp32.h>
#include <string>
#include <vector>

#pragma comment(lib, "Shell32.lib")

// 简单的 Win32 程序：提升权限 -> 语言选择窗口（"中文"/"English"）-> 回到桌面 -> 在桌面搜索名为 "Microsoft Edge" 的快捷方式并打开（模拟双击）
// -> 等待 2 秒 -> 将焦点切到 Edge 地址栏 (Ctrl+L) -> 输入对应 URL -> 回车

static const wchar_t* URL_CN = L"https://www.bilibili.com/video/BV1ad4y1V7wb";
static const wchar_t* URL_EN = L"https://www.bilibili.com/video/BV1AHmRBCEsq";

bool IsRunAsAdmin()
{
    BOOL isAdmin = FALSE;
    PSID adminGroup = NULL;
    SID_IDENTIFIER_AUTHORITY ntAuthority = SECURITY_NT_AUTHORITY;
    if (AllocateAndInitializeSid(&ntAuthority, 2,
        SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS,
        0, 0, 0, 0, 0, 0, &adminGroup))
    {
        CheckTokenMembership(NULL, adminGroup, &isAdmin);
        FreeSid(adminGroup);
    }
    return isAdmin == TRUE;
}

void RelaunchElevatedAndExit()
{
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(NULL, path, (DWORD)std::size(path));
    SHELLEXECUTEINFOW sei{ sizeof(sei) };
    sei.lpVerb = L"runas";
    sei.lpFile = path;
    sei.nShow = SW_SHOWNORMAL;
    if (!ShellExecuteExW(&sei))
    {
        // 用户取消了提升或失败，继续以当前权限运行或直接退出
    }
    ExitProcess(0);
}

// Minimal language selection window
struct LangSelection {
    HWND hwnd;
    int choice; // 0=none, 1=中文, 2=English
};

LRESULT CALLBACK LangWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    LangSelection* data = (LangSelection*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    switch (msg)
    {
    case WM_CREATE:
        {
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

int ShowLanguageSelectionDialog()
{
    const wchar_t CLASS_NAME[] = L"LangSelectWndClass";
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = LangWndProc;
    wc.hInstance = GetModuleHandleW(NULL);
    wc.lpszClassName = CLASS_NAME;
    RegisterClassExW(&wc);

    LangSelection data{};
    data.choice = 0;

    HWND hwnd = CreateWindowExW(0, CLASS_NAME, L"语言选择 / Language",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 400, 140,
        NULL, NULL, GetModuleHandleW(NULL), NULL);

    if (!hwnd) return 0;
    // center window
    RECT rc;
    GetWindowRect(hwnd, &rc);
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;
    int sx = GetSystemMetrics(SM_CXSCREEN);
    int sy = GetSystemMetrics(SM_CYSCREEN);
    SetWindowPos(hwnd, NULL, (sx - w) / 2, (sy - h) / 2, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

    SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)&data);
    ShowWindow(hwnd, SW_SHOW);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0) > 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    return data.choice;
}

void ShowDesktop()
{
    // 模拟 Win + D
    INPUT inputs[4] = {};
    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = VK_LWIN;
    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = 'D';
    inputs[2] = inputs[1];
    inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;
    inputs[3] = inputs[0];
    inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(4, inputs, sizeof(INPUT));
}

std::wstring GetSpecialDesktopPath(bool allUsers)
{
    wchar_t path[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPathW(NULL,
        allUsers ? CSIDL_COMMON_DESKTOPDIRECTORY : CSIDL_DESKTOPDIRECTORY,
        NULL, 0, path)))
    {
        return std::wstring(path);
    }
    return std::wstring();
}

std::wstring FindEdgeShortcutOnDesktops()
{
    std::vector<std::wstring> desktops;
    desktops.push_back(GetSpecialDesktopPath(false));
    desktops.push_back(GetSpecialDesktopPath(true));

    for (auto& d : desktops)
    {
        if (d.empty()) continue;
        WIN32_FIND_DATAW fd;
        std::wstring pattern = d + L"\\*";
        HANDLE h = FindFirstFileW(pattern.c_str(), &fd);
        if (h == INVALID_HANDLE_VALUE) continue;
        do
        {
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
            std::wstring name = fd.cFileName;
            // 看文件名是否包含 "Microsoft Edge"（不区分大小写），并且通常为 .lnk
            std::wstring lower = name;
            for (auto& c : lower) c = towlower(c);
            if (lower.find(L"microsoft edge") != std::wstring::npos)
            {
                std::wstring full = d + L"\\" + name;
                FindClose(h);
                return full;
            }
        } while (FindNextFileW(h, &fd));
        FindClose(h);
    }
    return L"";
}

HWND FindEdgeWindow()
{
    struct EnumContext { HWND found; };
    EnumContext ctx{ 0 };

    EnumWindows([](HWND hwnd, LPARAM lParam) -> BOOL {
        EnumContext* ctx = (EnumContext*)lParam;
        if (!IsWindowVisible(hwnd)) return TRUE;
        DWORD pid = 0;
        GetWindowThreadProcessId(hwnd, &pid);
        if (pid == 0) return TRUE;

        HANDLE hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_VM_READ, FALSE, pid);
        if (!hProc) return TRUE;

        wchar_t exePath[MAX_PATH] = {};
        DWORD len = (DWORD)std::size(exePath);
        if (QueryFullProcessImageNameW(hProc, 0, exePath, &len))
        {
            std::wstring s = exePath;
            for (auto& c : s) c = towlower(c);
            if (s.find(L"msedge.exe") != std::wstring::npos)
            {
                ctx->found = hwnd;
                CloseHandle(hProc);
                return FALSE; // stop enumeration
            }
        }
        CloseHandle(hProc);
        return TRUE;
    }, (LPARAM)&ctx);

    return ctx.found;
}

void SendCtrlLThenUrlAndEnter(HWND target, const std::wstring& url)
{
    if (!target) return;
    // 尝试将目标置于前台
    DWORD curThread = GetCurrentThreadId();
    DWORD tgtThread = GetWindowThreadProcessId(target, NULL);
    AttachThreadInput(curThread, tgtThread, TRUE);
    SetForegroundWindow(target);
    SetFocus(target);
    AttachThreadInput(curThread, tgtThread, FALSE);

    // 发送 Ctrl+L
    INPUT in[4] = {};
    in[0].type = INPUT_KEYBOARD; in[0].ki.wVk = VK_CONTROL;
    in[1].type = INPUT_KEYBOARD; in[1].ki.wVk = 'L';
    in[2] = in[1]; in[2].ki.dwFlags = KEYEVENTF_KEYUP;
    in[3] = in[0]; in[3].ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(4, in, sizeof(INPUT));

    // 发送 URL（使用 Unicode 输入）
    std::vector<INPUT> inputs;
    for (wchar_t ch : url)
    {
        INPUT iu{};
        iu.type = INPUT_KEYBOARD;
        iu.ki.wScan = ch;
        iu.ki.dwFlags = KEYEVENTF_UNICODE;
        inputs.push_back(iu);
        INPUT iuUp = iu;
        iuUp.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
        inputs.push_back(iuUp);
    }

    // 回车
    INPUT retDown{}; retDown.type = INPUT_KEYBOARD; retDown.ki.wVk = VK_RETURN;
    INPUT retUp = retDown; retUp.ki.dwFlags = KEYEVENTF_KEYUP;
    inputs.push_back(retDown);
    inputs.push_back(retUp);

    if (!inputs.empty())
        SendInput((UINT)inputs.size(), inputs.data(), sizeof(INPUT));
}

int wmain()
{
    // 1) 提升权限
    if (!IsRunAsAdmin())
    {
        RelaunchElevatedAndExit();
        return 0;
    }

    // 2) 语言选择
    int choice = ShowLanguageSelectionDialog();
    const wchar_t* url = (choice == 2) ? URL_EN : URL_CN;

    // 3) 回到桌面
    ShowDesktop();
    Sleep(300); // give Windows a moment

    // 4) 搜索桌面快捷方式名为 "Microsoft Edge"
    std::wstring lnk = FindEdgeShortcutOnDesktops();
    BOOL launched = FALSE;
    if (!lnk.empty())
    {
        // 用 ShellExecute 打开 .lnk，相当于双击
        HINSTANCE h = ShellExecuteW(NULL, L"open", lnk.c_str(), NULL, NULL, SW_SHOWNORMAL);
        launched = (INT_PTR)h > 32;
    }

    if (!launched)
    {
        // 回退：使用协议直接打开 Edge（尽量模仿用户双击）
        std::wstring proto = L"microsoft-edge:";
        HINSTANCE h = ShellExecuteW(NULL, L"open", proto.c_str(), NULL, NULL, SW_SHOWNORMAL);
        launched = (INT_PTR)h > 32;
    }

    // 5) 等待 2 秒
    Sleep(2000);

    // 6) 找到 Edge 窗口并输入 URL
    HWND edgeWnd = FindEdgeWindow();
    if (edgeWnd)
    {
        SendCtrlLThenUrlAndEnter(edgeWnd, url);
    }
    else
    {
        // 备用：直接用协议一次性打开带 URL（当无法找到窗口时）
        std::wstring full = std::wstring(L"microsoft-edge:") + url;
        ShellExecuteW(NULL, L"open", full.c_str(), NULL, NULL, SW_SHOWNORMAL);
    }

    return 0;
}
