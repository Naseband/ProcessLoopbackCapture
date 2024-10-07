#pragma once

#include <string>
#include <vector>
#include <filesystem>

#include <Windows.h>
#include <tlhelp32.h>
#include <Psapi.h>

// Window related functions require Psapi.lib

// ------------------------------------------------------------ 

// Counts all processes with the given name
inline bool CountProcesses(const std::wstring& executable_name, size_t& count)
{
    count = 0;

    HANDLE handle = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS | TH32CS_SNAPMODULE, 0);

    if (!handle)
        return false;
    
    PROCESSENTRY32W  process_info{ sizeof(PROCESSENTRY32W) };

    if (Process32FirstW(handle, &process_info))
    {
        do
        {
            if (executable_name.compare(process_info.szExeFile) == 0)
            {
                ++count;
            }
        }
        while (Process32NextW(handle, &process_info));
    }

    CloseHandle(handle);

    return true;
}

// Finds all processes with the given name
inline bool FindProcessIDs(const std::wstring& executable_name, std::vector<DWORD>& process_id_list)
{
    process_id_list.clear();

    HANDLE handle = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS | TH32CS_SNAPMODULE, 0);

    if (!handle)
        return false;
    
    PROCESSENTRY32W  process_info{ sizeof(PROCESSENTRY32W) };

    if (Process32FirstW(handle, &process_info))
    {
        do
        {
            if (executable_name.compare(process_info.szExeFile) == 0)
            {
                process_id_list.emplace_back(process_info.th32ProcessID);
            }
        }
        while (Process32NextW(handle, &process_info));
    }

    CloseHandle(handle);

    return true;
}

// Finds all processes with the same name, but filters out child processes before returning
// Used to find each parent process for multi-process applications
inline bool FindParentProcessIDs(std::wstring executable_name, std::vector<DWORD> &process_id_list)
{
    process_id_list.clear();

	HANDLE handle = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS | TH32CS_SNAPMODULE, 0);

    if (!handle)
        return false;
    
    struct s_process_pair
    {
		s_process_pair(DWORD process_id, DWORD parent_id) :
			process_id(process_id),
			parent_id(parent_id)
		{ }

        DWORD process_id;
        DWORD parent_id;
    };

    std::vector<s_process_pair> process_candidates;

    PROCESSENTRY32W  process_info = { sizeof(PROCESSENTRY32W) };

    if (Process32FirstW(handle, &process_info))
    {
        do
        {
            if (executable_name.compare(process_info.szExeFile) == 0)
            {
                process_candidates.emplace_back(process_info.th32ProcessID, process_info.th32ParentProcessID);
            }
        }
        while (Process32NextW(handle, &process_info));
    }

    CloseHandle(handle);

    for (auto &a : process_candidates)
    {
        for (auto &b : process_candidates)
        {
            if (a.process_id != b.process_id && a.parent_id == b.process_id)
            {
                goto next;
            }
        }

        process_id_list.emplace_back(a.process_id);

    next:
        continue;
    }

    return true;
}

// Gets the executable file path of the specified process id
inline bool GetProcessExecutablePath(DWORD process_id, std::filesystem::path& path)
{
    HANDLE process_handle = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, process_id);

    if (!process_handle)
        return false;

    wchar_t process_file_name[MAX_PATH]{};

    DWORD process_file_name_len = GetModuleFileNameExW(process_handle, NULL, process_file_name, MAX_PATH);

    CloseHandle(process_handle);

    if (!process_file_name_len)
        return false;

    path = process_file_name;

    return true;
}

// Gets all visible window handles associated with the given executable name
inline bool GetVisibleWindowsFromProcessName(const std::wstring& executable_name, std::vector<HWND>& window_list)
{
    struct s_window_info
    {
        std::wstring exe_name;
        std::vector<HWND>* window_list;
    };

	s_window_info window_info{ executable_name, &window_list };

    auto callback = [](HWND hwnd, LPARAM lparam) -> BOOL
    {
        if (!IsWindowVisible(hwnd))
            return true;

        auto window_info = reinterpret_cast<s_window_info*>(lparam);

        DWORD process_id;

        if (!GetWindowThreadProcessId(hwnd, &process_id))
            return true;

        HANDLE process_handle = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, process_id);

        if (!process_handle)
            return true;

        wchar_t process_file_name_c[MAX_PATH]{};

        DWORD process_file_name_len = GetModuleFileNameExW(process_handle, NULL, process_file_name_c, MAX_PATH);

        CloseHandle(process_handle);

        if (!process_file_name_len)
            return true;

        std::filesystem::path process_file_name{ process_file_name_c };

        if (process_file_name.filename().wstring().compare(window_info->exe_name) != 0)
            return true;

        window_info->window_list->emplace_back(hwnd);

        return true;
    };

    EnumWindows(callback, reinterpret_cast<LPARAM>(&window_info));

    return !window_list.empty();
}

// ------------------------------------------------------------ EOF