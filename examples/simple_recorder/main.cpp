/*

Example for ProcessLoopbackCapture

Captures the audio of the specified Process (by name), with pausing/unpausing and saving feature.
It will also find the parent process with the given name. This allows recording chrome, firefox, or other multiprocess applications.

*/

// If true there will be no RIFF/WAV header in the output file.
#define WRITE_RAW_FILE      false

#include <Windows.h>
#include <comdef.h> // HRESULT texts
#include <iostream>
#include <vector>
#include <mutex>
#include <string>
#include <format>

#include <ProcessLoopbackCapture.h>

#include <ParentProcess.h>

#ifndef FCC
#define FCC(ch4) ((((DWORD)(ch4) & 0xFF) << 24) |     \
                  (((DWORD)(ch4) & 0xFF00) << 8) |    \
                  (((DWORD)(ch4) & 0xFF0000) >> 8) |  \
                  (((DWORD)(ch4) & 0xFF000000) >> 24))
#endif

ProcessLoopbackCapture g_LoopbackCapture;

std::mutex g_AudioDataLock;
std::vector<unsigned char> g_Data;

void WriteWavFile(std::wstring FileName, std::vector<unsigned char> &Data, WAVEFORMATEX &Format);
void OnDataCapture(std::vector<unsigned char>::iterator &i1, std::vector<unsigned char>::iterator &i2, void* pUserData);

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

    g_LoopbackCapture.SetCaptureFormat(48000, 16, 2, WAVE_FORMAT_PCM);
    g_LoopbackCapture.SetTargetProcess(processId, true);
    g_LoopbackCapture.SetCallback(&OnDataCapture);
    g_LoopbackCapture.SetIntermediateBufferEnabled(true); // Use intermediate thread because our vector will probably reallocate if it reaches sizes like 2000MB
    g_LoopbackCapture.SetCallbackInterval(40);

    eCaptureError eError = g_LoopbackCapture.StartCapture();
    std::string tmp{};

    // Failed to start Capture, show the error as text and a HRESULT
    // Note that not all eCaptureErrors give back a valid HRESULT

    if (eError != eCaptureError::NONE)
    {
        HRESULT hr = g_LoopbackCapture.GetLastErrorResult();

        std::cout << std::endl;
        std::cout << "ERROR (" << (int)eError << "): " << LoopbackCaptureConst::GetErrorText(eError) << std::endl;
        std::cout << "HR: 0x" << std::hex << hr << std::dec << std::endl;
        std::wcout << L"HR Text: " << _com_error(hr).ErrorMessage() << std::endl;
        std::cout << std::endl;

        std::cout << "Press Enter to retry" << std::endl;

        std::getline(std::cin, tmp);

        goto __loop;
    }

    std::cout << "Capturing audio." << std::endl;
    std::cout << "Press Enter to stop and save." << std::endl;
    std::cout << "Type \"pause\" to pause or resume capture." << std::endl;
    std::cout << "Type \"hang\" to simulate a long hang in the callback." << std::endl;
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
                    HRESULT hr = g_LoopbackCapture.GetLastErrorResult();

                    std::cout << std::endl;
                    std::cout << "ERROR (" << (int)eError << "): " << LoopbackCaptureConst::GetErrorText(eError) << std::endl;
                    std::cout << "HR: 0x" << std::hex << hr << std::dec << std::endl;
                    std::wcout << L"HR Text: " << _com_error(hr).ErrorMessage() << std::endl;
                    std::cout << std::endl;
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
                    HRESULT hr = g_LoopbackCapture.GetLastErrorResult();

                    std::cout << std::endl;
                    std::cout << "ERROR (" << (int)eError << "): " << LoopbackCaptureConst::GetErrorText(eError) << std::endl;
                    std::cout << "HR: 0x" << std::hex << hr << std::dec << std::endl;
                    std::wcout << L"HR Text: " << _com_error(hr).ErrorMessage() << std::endl;
                    std::cout << std::endl;
                }
                else
                {
                    std::cout << "Capture resumed" << std::endl;
                }
            }

            continue;
        }

        // Hang test
        // If intermediate buffer is disabled, 15s of audio will be missing

        if (tmp.compare("hang") == 0)
        {
            std::cout << "Hanging Callback Thread for 15 seconds ..." << std::endl;
            std::scoped_lock lock(g_AudioDataLock);
            Sleep(15000);
            std::cout << "Done." << std::endl;
            continue;
        }

        if (tmp.length() == 0)
        {
            std::wstring Name = std::format(L"out-{}.wav", GetTickCount64());

            std::wcout << L"Saving Audio to \"" << Name << "\" ..." << std::endl;

            std::scoped_lock lock(g_AudioDataLock);

            WAVEFORMATEX Format{};
            g_LoopbackCapture.CopyCaptureFormat(Format);

            WriteWavFile(Name, g_Data, Format);

            std::cout << "Done" << std::endl;

            break;
        }
    }

    g_LoopbackCapture.StopCapture();

    std::cout << std::endl << "---------------" << std::endl << std::endl;

    goto __loop;

    return 0;
}


