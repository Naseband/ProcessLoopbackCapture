#pragma once

/*

Relatively simple class that uses Windows' AudioClient in loopback capture mode to capture audio coming from one process only.

Calls a user-provided callback with the resulting audio data in byte-by-byte PCM format.

Supports restarting one or more clients on the same process. The original ApplicationLoopback Sample leaked an interface pointer
that prevented another operation from being started.
https://github.com/microsoft/windows-classic-samples/tree/main/Samples/ApplicationLoopback

Settings (ie Capture Format, Process ID, etc) can only be modified if the capture is stopped (READY state).

When the callback function provided is called, you are expected to retrieve all audio data from it.
After each call, the internal buffer is cleared.

Optionally uses cameron314's readerwriterqueue to provide a more lenient user callback (https://github.com/cameron314/readerwriterqueue).
To use it, define PROCESS_LOOPBACK_CAPTURE_USE_QUEUE as a precompiler macro and enable it via the SetIntermediateThreadEnabled member.

Link against mfplat.lib, mmdevapi.lib and avrt.lib.
#pragma comment(lib, "mmdevapi.lib")
#pragma comment(lib, "avrt.lib")
*/

#include <windows.h>

#include <AudioClient.h>

#include <atomic>
#include <thread>
#include <vector>

#if defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE
#include <readerwriterqueue.h>
#endif

// ------------------------------------------------------------ 

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
}

// ------------------------------------------------------------ 

class ProcessLoopbackCapture
{
public:

    ProcessLoopbackCapture();
    ~ProcessLoopbackCapture();

    // Supported format values are listed in the LoopbackCaptureConst namespace.
    // The format tag must be either WAVE_FORMAT_IEEE_FLOAT or WAVE_FORMAT_PCM.
    // When using WAVE_FORMAT_IEEE_FLOAT, the stream will always have a bit depth equal to the size of float (32).
    eCaptureError SetCaptureFormat(unsigned int iSampleRate, unsigned int iBitDepth, unsigned int iChannelCount, unsigned int iFormatTag = WAVE_FORMAT_PCM);
    bool CopyCaptureFormat(WAVEFORMATEX &Format);

    // dwProcessId is the ID of the target process, but its children will always be considered part of it and the exclusion chain.
    // If bInclusive is true, only the sound of this process will be captured.
    // If it is false, the audio of this process will be excluded from all other sounds on the same device.
    eCaptureError SetTargetProcess(DWORD dwProcessId, bool bInclusive = true);

    eCaptureError SetCallback(void (*pCallbackFunc)(const std::vector<unsigned char>::iterator&, const std::vector<unsigned char>::iterator&, void*), void *pUserData = nullptr);

    // The interval (in milliseconds) is subject to Windows scheduling. Usually, the wait time is at least 16 and often a multiple of 16.
    // Only used if the intermediate thread is active.
    // Default: 100
    eCaptureError SetCallbackInterval(DWORD dwInterval);

    // If PROCESS_LOOPBACK_CAPTURE_USE_QUEUE is not defined (default), the intermediate thread will not be available and this function fails.
    // 
    // If bEnable is set to true, the audio data will be passed to the user callback from a seperate thread.
    // This will result in a non-time-critical callback, but it may delay buffer data by the given interval (see SetCallbackInterval).
    // 
    // If false, the user callback will be called directly from the main audio thread.
    // In this mode, you are responsible for handling the audio data quickly.
    // Usually, the internal buffer is cleared every 10ms and the callback should take no longer than this period to execute.
    eCaptureError SetIntermediateThreadEnabled(bool bEnable);

    // Returns the current state of the capture (ready, paused or started).
    // Safe to call from any thread.
    eCaptureState GetState();

    // If StartCapture fails, everything is reset to initial state.
    // Can be called again after failure.
    eCaptureError StartCapture();

    // Stops capture and resets everything to initial state.
    eCaptureError StopCapture();

    // Pauses capture. Note that there may still be some samples in the queue to process at the time of calling this.
    // Those samples will always be discarded.
    eCaptureError PauseCapture();
    // When resuming, you typically want to skip 0.1 seconds of the buffer. WASAPI keeps some of the old buffer when restarting capture.
    eCaptureError ResumeCapture(double fInitialDurationToSkip = 0.1);

    // Gets the last error code returned by Windows interface functions. Does not apply to eCaptureError::PARAM and eCaptureError::STATE.
    HRESULT GetLastErrorResult();

    // Returns/Resets the max execution time of the main audio thread (in milliseconds).
    double GetMaxExecutionTime();
    void ResetMaxExecutionTime();

    // Gets the current approx. size of the intermediate queue if the intermediate thread is available and in use.
    // Fails if PROCESS_LOOPBACK_CAPTURE_USE_QUEUE is not defined or SetIntermediateThreadEnabled was not set to true.
    eCaptureError GetQueueSize(size_t& iSize);

private:

    void Reset();

    // Duration in seconds of the initial buffer duration to skip. Used in Resume because some digital devices have leftover frames after resuming capture (IAudioClient::Stop, IAudioClient::Start).
    void StartThreads(double fInitialDurationToSkip);
    void StopThreads();

    void ProcessMainToCallback();

#if defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE
    void ProcessMainToQueue();
    void ProcessIntermediate();
#endif

    HRESULT                         m_hrLastError;

    std::atomic<eCaptureState>      m_CaptureState;

    IAudioClient                    *m_pAudioClient;
    IAudioCaptureClient             *m_pAudioCaptureClient; // Accessed from main audio thread
    HANDLE                          m_hSampleReadyEvent;

    bool                            m_bCaptureFormatInitialized;
    WAVEFORMATEX                    m_CaptureFormat{};
    DWORD                           m_dwProcessId;
    bool                            m_bProcessInclusive;
    bool                            m_bUseIntermediateThread;

    void                            (*m_pCallbackFunc)(const std::vector<unsigned char>::iterator&, const std::vector<unsigned char>::iterator&, void*);
    void                            *m_pCallbackFuncUserData;
    DWORD                           m_dwCallbackInterval;

    std::atomic<bool>               m_bRunAudioThreads;

    std::thread                     *m_pMainAudioThread;
    DWORD                           m_dwMainThreadBytesToSkip;
    std::atomic<double>             m_fMaxExecutionTime;

    std::thread                     *m_pQueueAudioThread;

#if defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE
    moodycamel::ReaderWriterQueue<unsigned char, 8192>
                                    m_Queue;
#endif

    std::vector<unsigned char>      m_AudioData; // Used to align the audio data in intermediate mode.
};

// ------------------------------------------------------------ EOF