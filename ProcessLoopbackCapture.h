#pragma once

/*

Relatively simple class that uses Windows' AudioClient in loopback capture mode to capture audio coming from one process only.

Calls a user-provided callback with the resulting audio data in byte-by-byte PCM format.

Uses cameron314's readerwriterqueue
https://github.com/cameron314/readerwriterqueue

Does not use WIL/WRL, which provides better compatibility than the original sample.

Supports restarting one or more clients on the same process. The original ApplicationLoopback Sample leaked an interface pointer
that prevented another operation from being started.
https://github.com/microsoft/windows-classic-samples/tree/main/Samples/ApplicationLoopback

Settings (ie Capture Format, Process ID, etc) can only be modified if the capture is stopped (READY state).

When the callback function provided is called, you are expected to retrieve all audio data from it.
After each call, the internal buffer is cleared.

The user callback is not subject to timer glitches if you take longer to execute.
However, longer execution times may cause the internal buffer to grow in size.
The buffer is resized to a minimum after a capture is stopped.

*/

#define WIN32_LEAN_AND_MEAN
#define VC_EXTRALEAN
#define NOMINMAX
#include <windows.h>
#undef WIN32_LEAN_AND_MEAN
#undef VC_EXTRALEAN
#undef NOMINMAX

#include <AudioClient.h>

#include <atomic>
#include <thread>
#include <vector>

#include <readerwriterqueue.h>

// ----------------------------------------------------------------------- 

enum class eCaptureState : int
{
    READY = 0,
    CAPTURING,
    PAUSED
};

enum class eCaptureError : int
{
    NONE = 0, // Success
    PARAM,
    STATE,
    FORMAT,
    PROCESSID,
    DEVICE,
    ACTIVATION,
    INITIALIZE,
    CAPTURE_SERVICE,
    START,
    STOP,
    EVENT
};

namespace LoopbackCaptureConst
{
    constexpr const char* GetErrorText(eCaptureError eID)
    {
        switch (eID)
        {
        case eCaptureError::NONE: return "Success";
        case eCaptureError::PARAM: return "Invalid parameter";
        case eCaptureError::STATE: return "Invalid operation for current state";
        case eCaptureError::FORMAT: return "Format is invalid or not initialized";
        case eCaptureError::PROCESSID: return "ProcessId is invalid (0/not set)";
        case eCaptureError::DEVICE: return "Failed to get device";
        case eCaptureError::ACTIVATION: return "Failed to activate device";
        case eCaptureError::INITIALIZE: return "Failed to init device";
        case eCaptureError::CAPTURE_SERVICE: return "Failed to get capture service";
        case eCaptureError::START: return "Failed to start capture";
        case eCaptureError::STOP: return "Failed to stop capture";
        case eCaptureError::EVENT: return "Failed to create and set event";
        }

        return "Unknown";
    }

    constexpr unsigned int VALID_SAMPLE_RATES[] =
    {
        8000,
        11025,
        16000,
        22050,
        44100,
        48000,
        88200,
        96,000,
        176400,
        192000,
        352800,
        384000
    };

    constexpr unsigned int MIN_CHANNEL_COUNT = 1;
    constexpr unsigned int MAX_CHANNEL_COUNT = 8;
}

// ----------------------------------------------------------------------- 

class ProcessLoopbackCapture
{
public:

    ProcessLoopbackCapture();
    ~ProcessLoopbackCapture();

    // Supported formats are listed in the LoopbackCaptureConst namespace.
    // iBitDepth must be 8, 16, 24 or 32.
    eCaptureError SetCaptureFormat(unsigned int iSampleRate, unsigned int iBitDepth, unsigned int iChannelCount);
    bool CopyCaptureFormat(WAVEFORMATEX &Format);

    // dwProcessId is the ID of the target process, but its children will always be considered part of it.
    // If bInclusive is true, only the sound of this process will be captured.
    // If it is false, the audio of this process will be excluded from all other sounds on the same device.
    eCaptureError SetTargetProcess(DWORD dwProcessId, bool bInclusive = true);

    eCaptureError SetCallback(void (*pCallbackFunc)(std::vector<unsigned char>::iterator&, std::vector<unsigned char>::iterator&, void*), void *pUserData = nullptr);

    // The interval is subject to Windows scheduling. Usually, the wait time is at least 16 and often a multiple of 16.
    // Default: 100
    eCaptureError SetCallbackInterval(DWORD dwInterval);

    eCaptureState GetState();

    // If StartCapture fails, everything is reset to initial state.
    // Can be called again after failure.
    eCaptureError StartCapture();

    // Stops capture and resets everything to initial state.
    eCaptureError StopCapture();

    // Pauses capture. Note that there may still be some samples in the queue to process at the time of calling this.
    eCaptureError PauseCapture();
    eCaptureError ResumeCapture();

    // Gets the last error code returned by Windows interface functions. Does not apply to eCaptureError::PARAM and eCaptureError::STATE.
    HRESULT GetLastErrorResult();

private:

    void Reset();

    void StartThreads();

    void ProcessMain();
    void ProcessQueue();

    bool                            m_bInitMF;

    HRESULT                         m_hrLastError;

    std::atomic<eCaptureState>      m_CaptureState;

    IAudioClient                    *m_pAudioClient;
    IAudioCaptureClient             *m_pAudioCaptureClient; // Accessed from capture thread
    HANDLE                          m_hSampleReadyEvent;

    bool                            m_bCaptureFormatInitialized;
    WAVEFORMATEX                    m_CaptureFormat{};
    DWORD                           m_dwProcessId;
    bool                            m_bProcessInclusive;

    void                            (*m_pCallbackFunc)(std::vector<unsigned char>::iterator&, std::vector<unsigned char>::iterator&, void*);
    void                            *m_pCallbackFuncUserData;
    DWORD                           m_dwCallbackInterval;

    std::thread                     *m_pMainAudioThread;
    std::atomic<bool>               m_bRunMainAudioThread;

    std::thread                     *m_pQueueAudioThread;
    std::atomic<bool>               m_bRunQueueAudioThread;

    moodycamel::ReaderWriterQueue<unsigned char>
                                    m_Queue;

    std::vector<unsigned char>      m_vecIntermediateBuffer; // Used to align the audio data and serves as storage if you use longer intervals.
};

// ----------------------------------------------------------------------- EOF