/*

Example for ProcessLoopbackCapture

Captures the audio of the specified Process (by name), with pausing/unpausing and saving feature.
It will also find the parent process with the given name.

This allows recording chrome, firefox and other multiprocess applications where audio is emitted from a child process.

*/

#include <Windows.h>
#include <combaseapi.h> // CoInitialize(Ex)
#include <comdef.h> // HRESULT texts

#include <iostream>
#include <vector>
#include <thread>
#include <atomic>
#include <string>
#include <format>

#include <ProcessLoopbackCapture.h>
#include <ProcessInfo.h> // For FindParentProcessIDs

// ------------------------------------------------------------ 

// If true there will be no RIFF/WAV header in the output file.
#define WRITE_RAW_FILE      false

constexpr unsigned int DEFAULT_SAMPLE_RATE = 44100U; // Can be changed by passing the three values in this order as command line arguments
constexpr unsigned int DEFAULT_BIT_DEPTH = 16U;
constexpr unsigned int DEFAULT_CHANNEL_COUNT = 2U;

// ------------------------------------------------------------ 

ProcessLoopbackCapture g_LoopbackCapture;

CRITICAL_SECTION g_AudioDataLock;
std::vector<unsigned char> g_Data;

// ------------------------------------------------------------ 

int main(int argc, char* argv[]);
void OnDataCapture(const std::vector<unsigned char>::iterator& i1, const std::vector<unsigned char>::iterator& i2, void* pUserData);
void WriteWavFile(std::wstring FileName, std::vector<unsigned char>& Data, WAVEFORMATEX& Format);

// ------------------------------------------------------------ 

int main(int argc, char* argv[])
{
    if (CoInitializeEx(NULL, COINIT_MULTITHREADED) != S_OK)
    {
        std::cout << "Failed to init COM" << std::endl;
        exit(1);
    }

    unsigned int sample_rate{ DEFAULT_SAMPLE_RATE };
    unsigned int bit_depth{ DEFAULT_BIT_DEPTH };
    unsigned int channel_count{ DEFAULT_CHANNEL_COUNT};

    if (argc >= 2)
    {
        try
        {
            sample_rate = static_cast<unsigned int>(std::stoull(argv[1]));
        }
        catch (...)
        {
            sample_rate = DEFAULT_SAMPLE_RATE;
        }
    }

    if (argc >= 3)
    {
        try
        {
            bit_depth = static_cast<unsigned int>(std::stoull(argv[2]));
        }
        catch (...)
        {
            bit_depth = DEFAULT_BIT_DEPTH;
        }
    }

    if (argc >= 4)
    {
        try
        {
            channel_count = static_cast<unsigned int>(std::stoull(argv[3]));
        }
        catch (...)
        {
            channel_count = DEFAULT_CHANNEL_COUNT;
        }
    }

    std::wcout << std::format(L"Sample Rate: {}", sample_rate) << std::endl;
    std::wcout << std::format(L"Bit Depth  : {}", bit_depth) << std::endl;
    std::wcout << std::format(L"Channels   : {}", channel_count) << std::endl;
    std::wcout << std::endl;

    InitializeCriticalSection(&g_AudioDataLock);

    bool run_application{ true };

    while (run_application)
    {
        g_Data.clear();
        g_Data.shrink_to_fit();
        DWORD processId{ 0 };

        do
        {
            // Repeatedly ask the user for a process name (incl. .exe) until one is valid

            processId = 0;

            std::wcout << L"Enter the Process Name to listen to (incl. .exe):" << std::endl << L"  ";

            std::wstring processName;

            std::getline(std::wcin, processName);

            if (processName.empty())
                continue;

            // Finds all parent processes with the given name
            // For Chrome-based browsers and Firefox there is only one child process responsible for playing audio
            // ProcessLoopbackCapture automatically includes all children, so recording the highest parent automatically includes the one playing audio
            // Here we just take the first one, but in a real app you might have the user select one

            std::vector<DWORD> vecProcessIds;
            FindParentProcessIDs(processName, vecProcessIds);

            if (vecProcessIds.size())
                processId = vecProcessIds[0];
        }
        while (processId == 0);

        std::wcout << "PID: " << processId << std::endl;

        // Set up format and some parameters

        g_LoopbackCapture.SetCaptureFormat(sample_rate, bit_depth, channel_count, WAVE_FORMAT_PCM);
        g_LoopbackCapture.SetTargetProcess(processId, true);
        g_LoopbackCapture.SetCallback(&OnDataCapture);
        g_LoopbackCapture.SetIntermediateThreadEnabled(true); // Use intermediate thread because our vector reallocation might take some time when it grows
        g_LoopbackCapture.SetCallbackInterval(40);

        eCaptureError eError = g_LoopbackCapture.StartCapture();

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

            continue;
        }

        std::cout << "Capturing audio." << std::endl;
        std::cout << "Press Enter to stop and save." << std::endl;
        std::cout << "Type \"discard\" to stop without saving." << std::endl;
        std::cout << "Type \"pause\" to pause or resume capture." << std::endl;
        std::cout << "Type \"hang\" to simulate a long hang in the callback." << std::endl;
        std::cout << "Type \"exit\" to exit the application." << std::endl;

        while (1)
        {
            std::string input;
            std::getline(std::cin, input);

            if (input.compare("discard") == 0)
            {
                break;
            }

            if (input.compare("pause") == 0)
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
            // If the intermediate thread is disabled, 10s of audio will be missing

            if (input.compare("hang") == 0)
            {
                std::cout << "Hanging Callback Thread for 10 seconds ..." << std::endl;
                
                EnterCriticalSection(&g_AudioDataLock);

                std::this_thread::sleep_for(std::chrono::seconds(10));

                LeaveCriticalSection(&g_AudioDataLock);

                std::cout << "Done." << std::endl;
                continue;
            }

            if (input.compare("exit") == 0)
            {
                run_application = false;
                break;
            }

            if (input.empty())
            {
                g_LoopbackCapture.StopCapture();

                std::wstring Name = std::format(L"out-{}.wav", GetTickCount64());

                std::wcout << L"Saving Audio to \"" << Name << "\" ..." << std::endl;

                WAVEFORMATEX Format{};
                g_LoopbackCapture.CopyCaptureFormat(Format);

                WriteWavFile(Name, g_Data, Format);

                std::cout << "Done" << std::endl;

                break;
            }
        }

        g_LoopbackCapture.StopCapture();
    }

    DeleteCriticalSection(&g_AudioDataLock);

    CoUninitialize();

    return 0;
}

// ------------------------------------------------------------ 

void OnDataCapture(const std::vector<unsigned char>::iterator& i1, const std::vector<unsigned char>::iterator& i2, void* pUserData)
{
    EnterCriticalSection(&g_AudioDataLock);

    g_Data.insert(g_Data.end(), i1, i2);

    LeaveCriticalSection(&g_AudioDataLock);
}

void WriteWavFile(std::wstring FileName, std::vector<unsigned char>& Data, WAVEFORMATEX& Format)
{
    HANDLE hFile;

    hFile = CreateFile(FileName.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    if (!hFile)
        return;

#if !WRITE_RAW_FILE

    // 1. RIFF chunk descriptor
    DWORD header[]{
        1179011410, // 'RIFF'
        0,
        1163280727, // 'WAVE'
        544501094, // 'fmt '
        sizeof(WAVEFORMATEX)
    };

    DWORD dwBytesWritten = 0;
    DWORD dwHeaderSize = 0;
    WriteFile(hFile, header, sizeof(header), &dwBytesWritten, NULL);
    dwHeaderSize += dwBytesWritten;

    WriteFile(hFile, &Format, sizeof(WAVEFORMATEX), &dwBytesWritten, NULL);
    dwHeaderSize += dwBytesWritten;

    DWORD data[] = { 1635017060, 0 };
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

// ------------------------------------------------------------ EOF