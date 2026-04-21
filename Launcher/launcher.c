#define UNICODE
#define _UNICODE
#include <windows.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <stdio.h>

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")

static void build_path(wchar_t *out, size_t out_chars, const wchar_t *dir, const wchar_t *name) {
    lstrcpynW(out, dir, (int)out_chars);
    PathAppendW(out, name);
}

static BOOL file_exists(const wchar_t *path) {
    DWORD attrs = GetFileAttributesW(path);
    return attrs != INVALID_FILE_ATTRIBUTES && !(attrs & FILE_ATTRIBUTE_DIRECTORY);
}

static BOOL run_and_wait(const wchar_t *command_line, const wchar_t *working_dir, DWORD *exit_code) {
    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    wchar_t cmd[4096];

    ZeroMemory(&si, sizeof(si));
    ZeroMemory(&pi, sizeof(pi));
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;

    lstrcpynW(cmd, command_line, (int)(sizeof(cmd) / sizeof(cmd[0])));
    if (!CreateProcessW(NULL, cmd, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, working_dir, &si, &pi)) {
        return FALSE;
    }

    WaitForSingleObject(pi.hProcess, INFINITE);
    if (exit_code) {
        GetExitCodeProcess(pi.hProcess, exit_code);
    }
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return TRUE;
}

static BOOL run_detached(const wchar_t *command_line, const wchar_t *working_dir) {
    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    wchar_t cmd[4096];

    ZeroMemory(&si, sizeof(si));
    ZeroMemory(&pi, sizeof(pi));
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;

    lstrcpynW(cmd, command_line, (int)(sizeof(cmd) / sizeof(cmd[0])));
    if (!CreateProcessW(NULL, cmd, NULL, NULL, FALSE, CREATE_NO_WINDOW | DETACHED_PROCESS, NULL, working_dir, &si, &pi)) {
        return FALSE;
    }

    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return TRUE;
}

static void show_error(const wchar_t *message) {
    MessageBoxW(NULL, message, L"启动错误", MB_OK | MB_ICONERROR);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow) {
    wchar_t exe_path[MAX_PATH];
    wchar_t app_dir[MAX_PATH];
    wchar_t py_exe[MAX_PATH];
    wchar_t pyw_exe[MAX_PATH];
    wchar_t app_config[MAX_PATH];
    wchar_t app_config_template[MAX_PATH];
    wchar_t ui_script[MAX_PATH];
    wchar_t cmd[4096];
    DWORD exit_code = 0;
    const wchar_t *fallback_py = L"C:\\Users\\adimn\\AppData\\Local\\Programs\\Python\\Python313\\python.exe";
    const wchar_t *fallback_pyw = L"C:\\Users\\adimn\\AppData\\Local\\Programs\\Python\\Python313\\pythonw.exe";

    if (!GetModuleFileNameW(NULL, exe_path, MAX_PATH)) {
        show_error(L"无法确定启动器路径。");
        return 1;
    }

    lstrcpynW(app_dir, exe_path, MAX_PATH);
    PathRemoveFileSpecW(app_dir);
    SetCurrentDirectoryW(app_dir);

    build_path(app_config, MAX_PATH, app_dir, L"app_config.json");
    build_path(app_config_template, MAX_PATH, app_dir, L"app_config.template.json");
    build_path(ui_script, MAX_PATH, app_dir, L"tr_raman_ui.py");

    if (!file_exists(ui_script)) {
        show_error(L"未找到 tr_raman_ui.py。");
        return 1;
    }

    if (!file_exists(app_config) && file_exists(app_config_template)) {
        CopyFileW(app_config_template, app_config, FALSE);
    }

    if (file_exists(fallback_py)) {
        lstrcpynW(py_exe, fallback_py, MAX_PATH);
    } else {
        lstrcpynW(py_exe, L"python", MAX_PATH);
    }

    if (file_exists(fallback_pyw)) {
        lstrcpynW(pyw_exe, fallback_pyw, MAX_PATH);
    } else {
        lstrcpynW(pyw_exe, L"pythonw", MAX_PATH);
    }

    _snwprintf(
        cmd,
        sizeof(cmd) / sizeof(cmd[0]),
        L"\"%s\" -c \"import pyvisa, pyvisa_py\"",
        py_exe
    );
    cmd[(sizeof(cmd) / sizeof(cmd[0])) - 1] = L'\0';

    if (!run_and_wait(cmd, app_dir, &exit_code)) {
        show_error(L"无法启动 Python 进行依赖检查。");
        return 1;
    }

    if (exit_code != 0) {
        _snwprintf(
            cmd,
            sizeof(cmd) / sizeof(cmd[0]),
            L"\"%s\" -m pip install pyvisa pyvisa-py psutil zeroconf",
            py_exe
        );
        cmd[(sizeof(cmd) / sizeof(cmd[0])) - 1] = L'\0';

        if (!run_and_wait(cmd, app_dir, &exit_code) || exit_code != 0) {
            show_error(L"Python 依赖安装失败。");
            return 1;
        }
    }

    _snwprintf(
        cmd,
        sizeof(cmd) / sizeof(cmd[0]),
        L"\"%s\" \"%s\"",
        pyw_exe,
        ui_script
    );
    cmd[(sizeof(cmd) / sizeof(cmd[0])) - 1] = L'\0';

    if (!run_detached(cmd, app_dir)) {
        show_error(L"无法启动图形界面。");
        return 1;
    }

    return 0;
}
