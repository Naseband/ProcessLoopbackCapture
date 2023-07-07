#pragma once

#include <string>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <tlhelp32.h>
#undef WIN32_LEAN_AND_MEAN

// Counts all processes with the given name
inline bool CountAllProcesses(std::wstring Name, size_t &iCount)
{
    iCount = 0;

    HANDLE hndl = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS | TH32CS_SNAPMODULE, 0);

    if (!hndl)
        return false;
    
    PROCESSENTRY32W  ProcessInfo = { sizeof(PROCESSENTRY32W) };

    if (Process32FirstW(hndl, &ProcessInfo))
    {
        do
        {
            if (wcscmp(ProcessInfo.szExeFile, Name.c_str()) == 0)
            {
                ++iCount;
            }
        }
        while (Process32NextW(hndl, &ProcessInfo));
    }

    CloseHandle(hndl);

    return true;
}

// Finds all processes with the given name
inline bool FindAllProcesses(std::wstring Name, std::vector<DWORD> &vecProcessIdList)
{
    vecProcessIdList.clear();

    HANDLE hndl = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS | TH32CS_SNAPMODULE, 0);

    if (!hndl)
        return false;
    
    PROCESSENTRY32W  ProcessInfo = { sizeof(PROCESSENTRY32W) };

    if (Process32FirstW(hndl, &ProcessInfo))
    {
        do
        {
            if (wcscmp(ProcessInfo.szExeFile, Name.c_str()) == 0)
            {
                vecProcessIdList.push_back(ProcessInfo.th32ProcessID);
            }
        }
        while (Process32NextW(hndl, &ProcessInfo));
    }

    CloseHandle(hndl);

    return true;
}

// Finds all processes with the same name, but filters out child processes before returning
// Used to find each parent process for multi-process applications
inline bool FindAllParentProcesses(std::wstring Name, std::vector<DWORD> &vecProcessIdList)
{
    vecProcessIdList.clear();

	HANDLE hndl = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS | TH32CS_SNAPMODULE, 0);

    if (!hndl)
        return false;
    
    struct stProcessPair
    {
        DWORD dwProcessId;
        DWORD dwParentId;
    };

    std::vector<stProcessPair> vecProcessCandidates;

    PROCESSENTRY32W  ProcessInfo = { sizeof(PROCESSENTRY32W) };

    if (Process32FirstW(hndl, &ProcessInfo))
    {
        do
        {
            if (wcscmp(ProcessInfo.szExeFile, Name.c_str()) == 0)
            {
                vecProcessCandidates.push_back({ ProcessInfo.th32ProcessID, ProcessInfo.th32ParentProcessID });
            }
        }
        while (Process32NextW(hndl, &ProcessInfo));
    }

    CloseHandle(hndl);

    for (auto &a : vecProcessCandidates)
    {
        for (auto &b : vecProcessCandidates)
        {
            if (a.dwProcessId != b.dwProcessId && a.dwParentId == b.dwProcessId)
            {
                goto next;
            }
        }

        vecProcessIdList.push_back(a.dwProcessId);

    next:
        continue;
    }

    return true;
}