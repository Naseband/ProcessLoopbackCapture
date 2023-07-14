/*

Example for ProcessLoopbackCapture

Captures the audio of the specified Process (by name), with pausing/unpausing and saving feature.
It will also find the parent process with the given name. This allows recording chrome, firefox, or other multiprocess applications.

*/

// If true there will be no RIFF/WAV header in the output file.
#define WRITE_RAW_FILE      false

// If true a helper thread will be spawned to preview the volume of each recording channel in the title bar of the program
#define VOLUME_DISPLAY      true

#include <Windows.h>
#include <iostream>
#include <psapi.h>
#include <tlhelp32.h>
#include <vector>
#include <mutex>
#include <thread>
#include <string>
#include <format>

#include <ProcessLoopbackCapture.h>

#include "ParentProcess.h"

#if VOLUME_DISPLAY
#include "PCMFormatConverter.h"
#endif

#ifndef FCC
#define FCC(ch4) ((((DWORD)(ch4) & 0xFF) << 24) |     \
                  (((DWORD)(ch4) & 0xFF00) << 8) |    \
                  (((DWORD)(ch4) & 0xFF0000) >> 8) |  \
                  (((DWORD)(ch4) & 0xFF000000) >> 24))
#endif

ProcessLoopbackCapture g_LoopbackCapture;

std::mutex g_AudioDataLock;
std::vector<unsigned char> g_Data;

#if VOLUME_DISPLAY
std::atomic<bool> g_bRunVolumeDisplayThread;
#endif

void WriteWavFile(std::wstring FileName, std::vector<unsigned char> &Data, WAVEFORMATEX &Format)
{
    HANDLE hFile;

    hFile = CreateFile(FileName.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    
    if (!hFile)
        return;

#if !WRITE_RAW_FILE

    // 1. RIFF chunk descriptor
    DWORD header[] = {
                        FCC('RIFF'),
                        0,
                        FCC('WAVE'),
                        FCC('fmt '),
                        sizeof(WAVEFORMATEX)
    };
    DWORD dwBytesWritten = 0;
    DWORD dwHeaderSize = 0;
    WriteFile(hFile, header, sizeof(header), &dwBytesWritten, NULL);
    dwHeaderSize += dwBytesWritten;

    WriteFile(hFile, &Format, sizeof(WAVEFORMATEX), &dwBytesWritten, NULL);
    dwHeaderSize += dwBytesWritten;

    DWORD data[] = { FCC('data'), 0 };
    WriteFile(hFile, data, sizeof(data), &dwBytesWritten, NULL);
    dwHeaderSize += dwBytesWritten;

    WriteFile(hFile, Data.data(), (DWORD)Data.size(), &dwBytesWritten, NULL);

    DWORD dwSize = (DWORD)Data.size();

    SetFilePointer(hFile, dwHeaderSize - sizeof(DWORD), NULL, FILE_BEGIN);
    WriteFile(hFile, &dwSize, sizeof(DWORD), &dwBytesWritten, NULL);

    SetFilePointer(hFile, sizeof(DWORD), NULL, FILE_BEGIN);

    DWORD cbTotalSize = dwSize + dwHeaderSize - 8;
    WriteFile(hFile, &cbTotalSize, sizeof(DWORD), &dwBytesWritten, NULL);

#else

    DWORD dwBytesWritten = 0;

    WriteFile(hFile, pData, dwSize, &dwBytesWritten, NULL);

#endif

    CloseHandle(hFile);
}

void OnDataCapture(std::vector<unsigned char>::iterator &i1, std::vector<unsigned char>::iterator &i2, void* pUserData)
{
    std::scoped_lock lock(g_AudioDataLock);

    g_Data.insert(g_Data.end(), i1, i2);
}

#if VOLUME_DISPLAY
void VolumeDisplay(HWND hWnd, WAVEFORMATEX Format)
{
    std::vector<unsigned char> vecData;
    std::vector<float> vecData_F;

    size_t iCount[2];
    float fVolume[2];

    std::string Text;
    int iChannel;

    while (g_bRunVolumeDisplayThread)
    {
        // Attempts to get the latest packet

        if (g_LoopbackCapture.GetState() == eCaptureState::CAPTURING && g_LoopbackCapture.GetLastPacket(vecData) == eCaptureError::NONE)
        {
            for (auto &v : iCount)
                v = 0;

            for (auto &v : fVolume)
                v = 0.0f;

            // Convert from unsigned char to float

            PCMByteToFloat(Format, vecData, vecData_F);

            // For channel 0 and 1 (L & R) we just get an average (ignoring values around 0)
            // This is far from an actual volume representation but it does the job for now

            iChannel = 0;

            for (auto &v : vecData_F)
            {
                if (iChannel <= 1 && abs(v) > 0.01)
                {
                    ++iCount[iChannel];
                    fVolume[iChannel] += abs(v);
                }

                if ((++iChannel) == Format.nChannels)
                    iChannel = 0;
            }

            // Build two bars from the values

            Text.clear();

            for (int i = 0; i < 2; ++i)
            {
                if (i != 0)
                    Text.append("   ");

                Text.append("[");

                if (iCount[i] != 0)
                    fVolume[i] = 2.0f * fVolume[i] / iCount[i]; // Approx.

                for (int j = 0; j < 12; ++j)
                {
                    if (j / 12.0f < fVolume[i])
                        Text.append("|");
                    else
                        Text.append(" ");
                }

                Text.append("]");
            }

            // Console window title

            SetWindowTextA(hWnd, Text.c_str());
        }

        Sleep(30);
    }
}
#endif

int wmain(int argc, wchar_t* argv[])
{
__loop:

    g_Data.clear();
    g_Data.shrink_to_fit();
    DWORD processId = 0;

    do
    {
        // Repeatedly ask the user for a process name (incl. .exe) until one is valid

        processId = 0;

        std::wcout << L"Enter the Process Name to listen to (incl. .exe):" << std::endl << L"  ";

        std::wstring processName;

        std::getline(std::wcin, processName);

        if (processName.length() == 0)
            continue;

        // Finds all parent processes with the given name
        // For Chrome-based browsers and Firefox there is only one child process responsible for playing audio
        // ProcessLoopbackCapture automatically includes all children, so recording the highest parent automatically includes the one playing audio
        // Here we just take the first one, but in a real app you might have the user select one

        std::vector<DWORD> vecProcessIds;
        FindAllParentProcesses(processName, vecProcessIds);

        if (vecProcessIds.size())
            processId = vecProcessIds[0];
    }
    while (processId == 0);

    std::wcout << "PID: " << processId << std::endl;

    // Set up format and some parameters

    g_LoopbackCapture.SetCaptureFormat(48000, 16, 2);
    g_LoopbackCapture.SetTargetProcess(processId, true);
    g_LoopbackCapture.SetCallback(&OnDataCapture);
    g_LoopbackCapture.SetIntermediateBufferEnabled(true); // Use intermediate thread because our vector will probably reallocate if it reaches sizes like 2000MB
    g_LoopbackCapture.SetLastPacketEnabled(true); // For Volume Display

    // For saving to Wav and VolumeDisplay

    WAVEFORMATEX Format{};
    g_LoopbackCapture.CopyCaptureFormat(Format);

    eCaptureError eError = g_LoopbackCapture.StartCapture();
    std::string tmp{};

    // Failed to start Capture, show the error as text and a HRESULT
    // Note that not all error codes give back a valid HRESULT

    if (eError != eCaptureError::NONE)
    {
        std::cout << "ERROR (" << (int)eError << "): " << LoopbackCaptureConst::GetErrorText(eError) << std::endl;
        std::cout << "HR: " << std::hex << g_LoopbackCapture.GetLastErrorResult() << std::dec << std::endl << std::endl;
        std::cout << "Press Enter to retry" << std::endl;

        std::getline(std::cin, tmp);

        goto __loop;
    }

#if VOLUME_DISPLAY

    g_bRunVolumeDisplayThread = true;

    std::thread VolumeDisplayThread(VolumeDisplay, GetConsoleWindow(), Format);

#endif

    std::cout << "Capturing audio." << std::endl;
    std::cout << "Press Enter to stop and save." << std::endl;
    std::cout << "Type \"pause\" to pause or resume capture." << std::endl;
    std::cout << "Type \"exit\" to stop without saving." << std::endl;

    while (1)
    {
        std::getline(std::cin, tmp);

        if (tmp.compare("exit") == 0)
            break;

        if (tmp.compare("pause") == 0)
        {
            if (g_LoopbackCapture.GetState() == eCaptureState::CAPTURING)
            {
                eCaptureError eError = g_LoopbackCapture.PauseCapture();

                if (eError != eCaptureError::NONE)
                {
                    std::cout << "ERROR (" << (int)eError << "): " << LoopbackCaptureConst::GetErrorText(eError) << std::endl;
                    std::cout << "HR: " << std::hex << g_LoopbackCapture.GetLastErrorResult() << std::dec << std::endl;
                }
                else
                {
                    std::cout << "Capture paused" << std::endl;
                }
            }
            else if (g_LoopbackCapture.GetState() == eCaptureState::PAUSED)
            {
                eCaptureError eError = g_LoopbackCapture.ResumeCapture();

                if (eError != eCaptureError::NONE)
                {
                    std::cout << "ERROR (" << (int)eError << "): " << LoopbackCaptureConst::GetErrorText(eError) << std::endl;
                    std::cout << "HR: " << std::hex << g_LoopbackCapture.GetLastErrorResult() << std::dec << std::endl;
                }
                else
                {
                    std::cout << "Capture resumed" << std::endl;
                }
                
            }

            continue;
        }

        if (tmp.length() == 0)
        {
            std::wstring Name = std::format(L"out-{}.wav", GetTickCount64());

            std::wcout << L"Saving Audio to \"" << Name << "\" ..." << std::endl;

            std::scoped_lock lock(g_AudioDataLock);

            WriteWavFile(Name, g_Data, Format);

            std::cout << "Done" << std::endl;

            break;
        }
    }

    g_LoopbackCapture.StopCapture();

#if VOLUME_DISPLAY

    g_bRunVolumeDisplayThread = false;

    VolumeDisplayThread.join();

#endif

    std::cout << "Going back to the beginning ..." << std::endl << std::endl;

    goto __loop;

    return 0;
}

