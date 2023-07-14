#pragma once

/*

Relatively simple class that uses Windows' AudioClient in loopback capture mode to capture audio coming from one process only.

Calls a user-provided callback with the resulting audio data in byte-by-byte PCM format.

Uses cameron314's readerwriterqueue
https://github.com/cameron314/readerwriterqueue

Define PROCESS_LOOPBACK_CAPTURE_NO_QUEUE before including ProcessLoopbackCapture (or as a global) to not use the Queue.
This will force ProcessLoopbackCapture to run without intermediate buffer mode.

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

#if defined PROCESS_LOOPBACK_CAPTURE_NO_QUEUE
#define PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE        false
#else
#define PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE        true
#endif

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
#include <string>
#include <mutex>

#if PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE
#include <readerwriterqueue.h>
#endif

// ----------------------------------------------------------------------- 

enum class eCaptureState : int
{
    READY = 0,
    CAPTURING,
    PAUSED
};

enum class eCaptureError : int
{
    // Errors without associated HRESULT
    NONE = 0, // Success
    PARAM,
    STATE,
    NOT_AVAILABLE,
    FORMAT,
    PROCESSID,

    // Errors with associated HRESULT (GetLastErrorResult)
    DEVICE,
    ACTIVATION,
    INITIALIZE,
    SERVICE,
    START,
    STOP,
    EVENT,
    INTERFACE
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
        case eCaptureError::NOT_AVAILABLE: return "Feature not available";
        case eCaptureError::FORMAT: return "CaptureFormat is invalid or not initialized";
        case eCaptureError::PROCESSID: return "ProcessId is invalid (0/not set)";

        case eCaptureError::DEVICE: return "Failed to get device";
        case eCaptureError::ACTIVATION: return "Failed to activate device";
        case eCaptureError::INITIALIZE: return "Failed to init device";
        case eCaptureError::SERVICE: return "Failed to get interface pointer via service";
        case eCaptureError::START: return "Failed to start capture";
        case eCaptureError::STOP: return "Failed to stop capture";
        case eCaptureError::EVENT: return "Failed to create and set event";
        case eCaptureError::INTERFACE: return "Failed to call Windows interface function";
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
        96000,
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

    // dwProcessId is the ID of the target process, but its children will always be considered part of it and the exclusion chain.
    // If bInclusive is true, only the sound of this process will be captured.
    // If it is false, the audio of this process will be excluded from all other sounds on the same device.
    eCaptureError SetTargetProcess(DWORD dwProcessId, bool bInclusive = true);

    eCaptureError SetCallback(void (*pCallbackFunc)(std::vector<unsigned char>::iterator&, std::vector<unsigned char>::iterator&, void*), void *pUserData = nullptr);

    // The interval is subject to Windows scheduling. Usually, the wait time is at least 16 and often a multiple of 16.
    // Only used if the intermediate buffer is active.
    // Default: 100
    eCaptureError SetCallbackInterval(DWORD dwInterval);

    // If you define PROCESS_LOOPBACK_CAPTURE_NO_QUEUE the intermediate buffer will never be used and this function returns an error (NOT_AVAILABLE).
    // If true, the intermediate buffer will be passed to the user callback from a seperate thread.
    // This will result in a non-time-critical callback, but it may delay buffer data by the given interval (see SetCallbackInterval).
    // If false, the user callback will be called from the main audio thread.
    // In this mode, you are responsible for handling the audio data in a fast manner.
    // Usually, the internal buffer is cleared every 10ms and the callback should take no longer than this period to execute.
    // Default: true
    eCaptureError SetIntermediateBufferEnabled(bool bEnable);

    eCaptureError SetLastPacketEnabled(bool bEnable);
    eCaptureError GetLastPacket(std::vector<unsigned char> &vecPacket, size_t *pPacketID = nullptr);

    eCaptureState GetState();

    // If StartCapture fails, everything is reset to initial state.
    // Can be called again after failure.
    eCaptureError StartCapture();

    // Stops capture and resets everything to initial state.
    eCaptureError StopCapture();

    // Pauses capture. Note that there may still be some samples in the queue to process at the time of calling this.
    eCaptureError PauseCapture();
    // When resuming, you want to skip 0.1 seconds of the buffer. WASAPI keeps some of the old buffer when restarting capture.
    eCaptureError ResumeCapture(double fInitialDurationToSkip = 0.1);

    // Gets the last error code returned by Windows interface functions. Does not apply to eCaptureError::PARAM and eCaptureError::STATE.
    HRESULT GetLastErrorResult();

    // Returns/Resets the max execution time of the main audio thread (milliseconds).
    double GetMaxExecutionTime();
    void ResetMaxExecutionTime();

private:

    void Reset();

    // Duration in seconds of the initial buffer duration to skip. Used in Resume because some digital devices have leftover frames after resuming capture (Stop, then Start).
    void StartThreads(double fInitialDurationToSkip);
    void StopThreads();

    void ProcessMainToCallback();

#if PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE
    void ProcessMainToQueue();
    void ProcessIntermediate();
#endif

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
    bool                            m_bUseIntermediateBuffer;
    std::atomic<bool>               m_bStoreLastPacket;

    void                            (*m_pCallbackFunc)(std::vector<unsigned char>::iterator&, std::vector<unsigned char>::iterator&, void*);
    void                            *m_pCallbackFuncUserData;
    DWORD                           m_dwCallbackInterval;

    bool                            m_bThreadsStarted;

    std::thread                     *m_pMainAudioThread;
    std::atomic<bool>               m_bRunMainAudioThread;
    DWORD                           m_dwMainThreadBytesToSkip;
    std::atomic<double>             m_fMaxExecutionTime;

    std::thread                     *m_pQueueAudioThread;
    std::atomic<bool>               m_bRunQueueAudioThread;

#if PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE
    moodycamel::ReaderWriterQueue<unsigned char>
                                    m_Queue;
#endif

    std::vector<unsigned char>      m_vecIntermediateBuffer; // Used to align the audio data. Serves as storage if you use longer intervals.

    std::mutex                      m_xLastPacketLock;
    std::vector<unsigned char>      m_vecLastPacket;
    size_t                          m_iLastPacketIndex;
};

// ----------------------------------------------------------------------- EOF