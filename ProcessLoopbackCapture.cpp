#include <ProcessLoopbackCapture.h>

#include <chrono>

#include <mmdeviceapi.h>
#include <mfapi.h>
#include <audioclientactivationparams.h>
#include <audiopolicy.h>

using namespace std;

// ----------------------------------------------------------------------- 

// Simple class that turns ActivateAudioInterfaceAsync into a blocking operation.
// Use an instance of this class for ActivateAudioInterfaceAsync.
// Afterwards call AsyncCallback::Wait, which blocks until completed.
// AgileObject inheritance not strictly needed but the API might change in the future.

// If successful, you will need to release the returned IAudioClient.
// The IActivateAudioInterfaceAsyncOperation interface pointer will also need to be released, regardless of success.
// Call Reset() between ActivateAudioInterfaceAsync calls if you want to use the callback more than once.

class ActivateAudioInterfaceAsyncCallback : 
    public IActivateAudioInterfaceCompletionHandler,
    public IAgileObject
{
public:

    ActivateAudioInterfaceAsyncCallback()
    {
        Reset();
    }

    void Reset()
    {
        m_bActivateFinished = false;
    }

    void Wait()
    {
        m_bActivateFinished.wait(false);
    }

protected:

    atomic<bool> m_bActivateFinished;

    STDMETHOD(ActivateCompleted)(IActivateAudioInterfaceAsyncOperation *activateOperation)
    {
        m_bActivateFinished = true;
        m_bActivateFinished.notify_one();

        return S_OK;
    }

    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObject)
    {
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IAgileObject))
        {
            *ppvObject = this;
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    STDMETHOD_(ULONG, AddRef)() // Ignored
    {
        return 1;
    }

    STDMETHOD_(ULONG, Release)() // Ignored
    {
        return 0;
    }
};

// ----------------------------------------------------------------------- ProcessLoopbackCapture

// public

ProcessLoopbackCapture::ProcessLoopbackCapture() :
    m_bInitMF(false),
    m_hrLastError(S_OK),

    m_CaptureState(eCaptureState::READY),
    m_pAudioClient(nullptr),
    m_pAudioCaptureClient(nullptr),
    m_hSampleReadyEvent(NULL),
    m_bCaptureFormatInitialized(false),
    m_dwProcessId(0),
    m_bProcessInclusive(false),
    m_bUseIntermediateBuffer(true),

    m_pCallbackFunc(nullptr),
    m_pCallbackFuncUserData(nullptr),
    m_dwCallbackInterval(100),

    m_bThreadsStarted(false),

    m_pMainAudioThread(nullptr),
    m_bRunMainAudioThread(false),
    m_dwMainThreadBytesToSkip(0),
    m_fMaxExecutionTime(0.0),

    m_pQueueAudioThread(nullptr),
    m_bRunQueueAudioThread(false)
{
    
}

ProcessLoopbackCapture::~ProcessLoopbackCapture()
{
    StopCapture();

    if (m_bInitMF)
    {
        MFShutdown();
        m_bInitMF = false;
    }
}

eCaptureError ProcessLoopbackCapture::SetCaptureFormat(unsigned int iSampleRate, unsigned int iBitDepth, unsigned int iChannelCount, unsigned int iFormatTag)
{
    if (m_CaptureState != eCaptureState::READY)
        return eCaptureError::STATE;

    if (iSampleRate < LoopbackCaptureConst::MIN_SAMPLE_RATE || iSampleRate > LoopbackCaptureConst::MAX_SAMPLE_RATE)
        return eCaptureError::PARAM;

    if (iBitDepth == 0 || iBitDepth > LoopbackCaptureConst::MAX_BIT_DEPTH || (iBitDepth % CHAR_BIT) != 0)
        return eCaptureError::PARAM;

    if (iChannelCount < LoopbackCaptureConst::MIN_CHANNEL_COUNT || iChannelCount > LoopbackCaptureConst::MAX_CHANNEL_COUNT)
        return eCaptureError::PARAM;

    if (iFormatTag == WAVE_FORMAT_IEEE_FLOAT)
        iBitDepth = sizeof(float) * CHAR_BIT;
    else if (iFormatTag != WAVE_FORMAT_PCM)
        return eCaptureError::PARAM;

    m_CaptureFormat.wFormatTag = (WORD)iFormatTag;
    m_CaptureFormat.nChannels = (WORD)iChannelCount;
    m_CaptureFormat.nSamplesPerSec = (DWORD)iSampleRate;
    m_CaptureFormat.wBitsPerSample = (WORD)iBitDepth;
    m_CaptureFormat.nBlockAlign = m_CaptureFormat.nChannels * m_CaptureFormat.wBitsPerSample / CHAR_BIT;
    m_CaptureFormat.nAvgBytesPerSec = m_CaptureFormat.nSamplesPerSec * m_CaptureFormat.nBlockAlign;

    m_bCaptureFormatInitialized = true;

    return eCaptureError::NONE;
}

bool ProcessLoopbackCapture::CopyCaptureFormat(WAVEFORMATEX &Format)
{
    if (!m_bCaptureFormatInitialized)
        return false;

    Format = m_CaptureFormat;

    return true;
}

eCaptureError ProcessLoopbackCapture::SetTargetProcess(DWORD dwProcessId, bool bInclusive)
{
    if (m_CaptureState != eCaptureState::READY)
        return eCaptureError::STATE;

    if (dwProcessId == 0)
        return eCaptureError::PARAM;

    m_dwProcessId = dwProcessId;
    m_bProcessInclusive = bInclusive;

    return eCaptureError::NONE;
}

eCaptureError ProcessLoopbackCapture::SetCallback(void (*pCallbackFunc)(std::vector<unsigned char>::iterator&, std::vector<unsigned char>::iterator&, void*), void *pUserData)
{
    if (m_CaptureState != eCaptureState::READY)
        return eCaptureError::STATE;

    m_pCallbackFunc = pCallbackFunc;
    m_pCallbackFuncUserData = pUserData;

    return eCaptureError::NONE;
}

eCaptureError ProcessLoopbackCapture::SetCallbackInterval(DWORD dwInterval)
{
    if (m_CaptureState != eCaptureState::READY)
        return eCaptureError::STATE;

    if (dwInterval < 1)
        m_dwCallbackInterval = 1;
    else
        m_dwCallbackInterval = dwInterval;

    return eCaptureError::NONE;
}

eCaptureError ProcessLoopbackCapture::SetIntermediateBufferEnabled(bool bEnable)
{
#if !PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE
    return eCaptureError::NOT_AVAILABLE;
#endif

    if (m_CaptureState != eCaptureState::READY)
        return eCaptureError::STATE;

    m_bUseIntermediateBuffer = !!bEnable;

    return eCaptureError::NONE;
}

eCaptureState ProcessLoopbackCapture::GetState()
{
    return m_CaptureState.load();
}

eCaptureError ProcessLoopbackCapture::StartCapture()
{
    if (m_CaptureState != eCaptureState::READY)
        return eCaptureError::STATE;

    if (!m_bCaptureFormatInitialized)
        return eCaptureError::FORMAT;

    if (!m_dwProcessId)
        return eCaptureError::PROCESSID;

    if (!m_bInitMF)
    {
        MFStartup(MF_VERSION, MFSTARTUP_LITE);

        m_bInitMF = true;
    }

    // Set up Params

    AUDIOCLIENT_ACTIVATION_PARAMS Blob{};
    Blob.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
    Blob.ProcessLoopbackParams.ProcessLoopbackMode = m_bProcessInclusive ? PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE : PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE;
    Blob.ProcessLoopbackParams.TargetProcessId = m_dwProcessId;

    PROPVARIANT ActivateParams = {};
    ActivateParams.vt = VT_BLOB;
    ActivateParams.blob.cbSize = sizeof(AUDIOCLIENT_ACTIVATION_PARAMS);
    ActivateParams.blob.pBlobData = (BYTE*)&Blob;

    // Activate ("Async")

    IActivateAudioInterfaceAsyncOperation* pActivateOp = nullptr;
    ActivateAudioInterfaceAsyncCallback AsyncCallback;

    m_hrLastError = ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, __uuidof(IAudioClient), &ActivateParams, &AsyncCallback, &pActivateOp);

    if (m_hrLastError != S_OK)
    {
        Reset();
        return eCaptureError::DEVICE;
    }

    AsyncCallback.Wait();

    pActivateOp->GetActivateResult(&m_hrLastError, reinterpret_cast<IUnknown**>(&m_pAudioClient));
    pActivateOp->Release();

    if (m_hrLastError != S_OK)
    {
        Reset();
        return eCaptureError::ACTIVATION;
    }
    
    // Initialize (AudioClient is valid)

    m_hrLastError = m_pAudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY,
        0, // 100ns units, 10 = 1 micro, 10000 = 1 milli. Does not seem to do anything for this mode (Win10).
        0, // Do not use for Capture Clients.
        &m_CaptureFormat,
        nullptr);

    if (m_hrLastError != S_OK)
    {
        Reset();
        return eCaptureError::INITIALIZE;
    }

    // Get CaptureClient pointer (used to get the samples)

    m_hrLastError = m_pAudioClient->GetService(IID_PPV_ARGS(&m_pAudioCaptureClient));

    if (m_hrLastError != S_OK)
    {
        Reset();
        return eCaptureError::SERVICE;
    }

    // Create and set event

    m_hSampleReadyEvent = CreateEventW(NULL, false, false, NULL);

    m_hrLastError = m_pAudioClient->SetEventHandle(m_hSampleReadyEvent); // Fails if event fails to create

    if(m_hrLastError != S_OK || m_hSampleReadyEvent == NULL) // NULL check so VS doesn't cry
    {
        Reset();
        return eCaptureError::EVENT;
    }

    // Start

    m_hrLastError = m_pAudioClient->Start();

    if (m_hrLastError != S_OK)
    {
        Reset();
        return eCaptureError::START;
    }

    StartThreads(0.0);

    m_CaptureState = eCaptureState::CAPTURING;

    return eCaptureError::NONE;
}

eCaptureError ProcessLoopbackCapture::StopCapture()
{
    if (m_CaptureState == eCaptureState::READY)
        return eCaptureError::STATE;

    Reset();

    return eCaptureError::NONE;
}

eCaptureError ProcessLoopbackCapture::PauseCapture()
{
    if (m_CaptureState != eCaptureState::CAPTURING)
        return eCaptureError::STATE;

    m_CaptureState = eCaptureState::PAUSED;

    m_hrLastError = m_pAudioClient->Stop();

    if (m_hrLastError != S_OK)
        return eCaptureError::STOP;

    StopThreads();

    return eCaptureError::NONE;
}

eCaptureError ProcessLoopbackCapture::ResumeCapture(double fInitialDurationToSkip)
{
    if (m_CaptureState != eCaptureState::PAUSED)
        return eCaptureError::STATE;

    m_CaptureState = eCaptureState::CAPTURING;

    ResetEvent(m_hSampleReadyEvent);
    m_hrLastError = m_pAudioClient->Start();

    if (m_hrLastError != S_OK)
        return eCaptureError::START;

    StartThreads(fInitialDurationToSkip);

    return eCaptureError::NONE;
}

HRESULT ProcessLoopbackCapture::GetLastErrorResult()
{
    return m_hrLastError;
}

double ProcessLoopbackCapture::GetMaxExecutionTime()
{
    return m_fMaxExecutionTime;
}

void ProcessLoopbackCapture::ResetMaxExecutionTime()
{
    m_fMaxExecutionTime = 0.0;
}

eCaptureError ProcessLoopbackCapture::GetIntermediateBuffer(std::vector<unsigned char> *&pVector)
{
    if (m_CaptureState == eCaptureState::READY)
        return eCaptureError::STATE;

    pVector = &m_vecIntermediateBuffer;

    return eCaptureError::NONE;
}

// private

void ProcessLoopbackCapture::Reset()
{
    if (m_CaptureState == eCaptureState::CAPTURING) // Prevent another event trigger 
    {
        m_pAudioClient->Stop();
    }

    StopThreads();

    if (m_pAudioCaptureClient != nullptr)
    {
        m_pAudioCaptureClient->Release();
        m_pAudioCaptureClient = nullptr;
    }

    if (m_pAudioClient != nullptr)
    {
        m_pAudioClient->Reset();
        m_pAudioClient->Release();
        m_pAudioClient = nullptr;
    }

    if (m_hSampleReadyEvent != NULL)
    {
        CloseHandle(m_hSampleReadyEvent);
        m_hSampleReadyEvent = NULL;
    }

    m_CaptureState = eCaptureState::READY;
}

void ProcessLoopbackCapture::StartThreads(double fInitialDurationToSkip)
{
    if (m_bThreadsStarted)
        return;

    if (fInitialDurationToSkip < 0.0)
        fInitialDurationToSkip = 0.0;

    m_dwMainThreadBytesToSkip = (DWORD)(m_CaptureFormat.nSamplesPerSec * fInitialDurationToSkip) * (DWORD)m_CaptureFormat.nBlockAlign;

#if PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE

    if (m_bUseIntermediateBuffer)
    {
        m_bRunMainAudioThread = true;
        m_pMainAudioThread = new thread(&ProcessLoopbackCapture::ProcessMainToQueue, this);

        m_bRunQueueAudioThread = true;
        m_pQueueAudioThread = new thread(&ProcessLoopbackCapture::ProcessIntermediate, this);
    }
    else
    {
        m_bRunMainAudioThread = true;
        m_pMainAudioThread = new thread(&ProcessLoopbackCapture::ProcessMainToCallback, this);

        m_bRunQueueAudioThread = false;
        m_pQueueAudioThread = nullptr;
    }

#else

    m_bRunMainAudioThread = true;
    m_pMainAudioThread = new thread(&ProcessLoopbackCapture::ProcessMainToCallback, this);

    m_bRunQueueAudioThread = false;
    m_pQueueAudioThread = nullptr;

#endif

    m_bThreadsStarted = true;
}

void ProcessLoopbackCapture::StopThreads()
{
    if (!m_bThreadsStarted)
        return;

    m_bRunMainAudioThread = false;
    m_bRunQueueAudioThread = false;

    SetEvent(m_hSampleReadyEvent);

#if PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE

    m_pMainAudioThread->join();
    delete m_pMainAudioThread;
    m_pMainAudioThread = nullptr;

    if (m_bUseIntermediateBuffer)
    {
        m_pQueueAudioThread->join();
        delete m_pQueueAudioThread;
        m_pQueueAudioThread = nullptr;
    }

    while (m_Queue.pop()); // Flush

#else

    m_pMainAudioThread->join();
    delete m_pMainAudioThread;
    m_pMainAudioThread = nullptr;

#endif

    m_vecIntermediateBuffer.clear();
    m_vecIntermediateBuffer.shrink_to_fit();

    m_bThreadsStarted = false;
}

void ProcessLoopbackCapture::ProcessMainToCallback()
{
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);

    BYTE* pData = nullptr;
    UINT32 iFramesAvailable;
    UINT64 iBytesAvailable;
    DWORD dwCaptureFlags;
    // UINT64 u64DevicePosition = 0; // Not currently used
    // UINT64 u64QPCPosition = 0;

    while (m_bRunMainAudioThread)
    {
        // The event is signaled if either a sample is ready or the capture was stopped.

        if (WaitForSingleObject(m_hSampleReadyEvent, 5000) == WAIT_OBJECT_0)
        {
            // The capture was stopped, exit thread.
            if (!m_bRunMainAudioThread)
                return;

            auto tick_start = chrono::steady_clock::now();

            while (m_pAudioCaptureClient->GetBuffer(&pData, &iFramesAvailable, &dwCaptureFlags, nullptr, nullptr) == S_OK)
            {
                iBytesAvailable = (UINT64)iFramesAvailable * (UINT64)m_CaptureFormat.nBlockAlign;

                for (UINT64 i = 0; i < iBytesAvailable; ++i)
                {
                    if (m_dwMainThreadBytesToSkip != 0)
                    {
                        --m_dwMainThreadBytesToSkip;
                    }
                    else
                    {
                        m_vecIntermediateBuffer.push_back(pData[i]);
                    }
                }

                m_pAudioCaptureClient->ReleaseBuffer(iFramesAvailable);
            }

            if (m_vecIntermediateBuffer.size() > 0)
            {
                if (m_pCallbackFunc != nullptr)
                {
                    auto i1 = m_vecIntermediateBuffer.begin();
                    auto i2 = m_vecIntermediateBuffer.begin() + m_vecIntermediateBuffer.size();
                    m_pCallbackFunc(i1, i2, m_pCallbackFuncUserData);
                }

                m_vecIntermediateBuffer.clear();
            }

            auto dur = chrono::duration_cast<chrono::nanoseconds>(chrono::steady_clock::now() - tick_start).count() / 1e6;

            if (dur > m_fMaxExecutionTime)
                m_fMaxExecutionTime = dur;
        }
    }
}

#if PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE

void ProcessLoopbackCapture::ProcessMainToQueue()
{
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);

    BYTE* pData = nullptr;
    UINT32 iFramesAvailable;
    UINT64 iBytesAvailable;
    DWORD dwCaptureFlags;
    // UINT64 u64DevicePosition = 0; // Not currently used
    // UINT64 u64QPCPosition = 0;

    while (m_bRunMainAudioThread)
    {
        // The event is signaled if either a sample is ready or the capture was stopped.

        if (WaitForSingleObject(m_hSampleReadyEvent, 5000) == WAIT_OBJECT_0)
        {
            // The capture was stopped, exit thread.
            if (!m_bRunMainAudioThread)
                return;

            auto tick_start = chrono::steady_clock::now();

            while (m_pAudioCaptureClient->GetBuffer(&pData, &iFramesAvailable, &dwCaptureFlags, nullptr, nullptr) == S_OK)
            {
                iBytesAvailable = (UINT64)iFramesAvailable * (UINT64)m_CaptureFormat.nBlockAlign;

                for (UINT64 i = 0; i < iBytesAvailable; ++i)
                {
                    if (m_dwMainThreadBytesToSkip != 0)
                    {
                        --m_dwMainThreadBytesToSkip;
                    }
                    else
                    {
                        m_Queue.enqueue(pData[i]);
                    }
                }

                m_pAudioCaptureClient->ReleaseBuffer(iFramesAvailable);
            }

            auto dur = chrono::duration_cast<chrono::nanoseconds>(chrono::steady_clock::now() - tick_start).count() / 1e6;

            if (dur > m_fMaxExecutionTime)
                m_fMaxExecutionTime = dur;
        }
    }
}

void ProcessLoopbackCapture::ProcessIntermediate()
{
    while (m_bRunQueueAudioThread)
    {
        // Get all data from queue into intermediate buffer and pass it to the callback
        // At the same time, we make sure the data stays aligned for the callback

        unsigned char Data;

        while (m_Queue.try_dequeue(Data))
        {
            m_vecIntermediateBuffer.push_back(Data);
        }

        // Get aligned buffer size and send to callback, then delete

        size_t iAlignedSize = m_vecIntermediateBuffer.size() / m_CaptureFormat.nBlockAlign * m_CaptureFormat.nBlockAlign;

        if (iAlignedSize > 0)
        {
            if (m_pCallbackFunc != nullptr)
            {
                auto i1 = m_vecIntermediateBuffer.begin();
                auto i2 = m_vecIntermediateBuffer.begin() + iAlignedSize;
                m_pCallbackFunc(i1, i2, m_pCallbackFuncUserData);
            }

            m_vecIntermediateBuffer.erase(m_vecIntermediateBuffer.begin(), m_vecIntermediateBuffer.begin() + iAlignedSize);
        }

        Sleep(m_dwCallbackInterval);
    }
}

#endif // #if PROCESS_LOOPBACK_CAPTURE_QUEUE_AVAILABLE

// ----------------------------------------------------------------------- EOF