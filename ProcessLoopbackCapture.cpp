#include <ProcessLoopbackCapture.h>

#include <mmdeviceapi.h>
#include <mfapi.h>
#include <audioclientactivationparams.h>
#pragma comment(lib, "mfplat.lib")
#pragma comment(lib, "mmdevapi.lib")

using namespace std;

// ----------------------------------------------------------------------- 

// Simple class that turns ActivateAudioInterfaceAsync into a blocking operation.
// Use an instance of this class for ActivateAudioInterfaceAsync.
// Afterwards call AsyncCallback::GetResult, which blocks until completed.

// If successful, you will need to release the returned IAudioClient.

class ActivateAudioInterfaceAsyncCallback : public IActivateAudioInterfaceCompletionHandler
{
public:

    atomic<bool> m_bActivateFinished;
    HRESULT m_hrActivateResult;
    IAudioClient* m_pAudioClient;

    ActivateAudioInterfaceAsyncCallback() :
        m_hrActivateResult(E_UNEXPECTED),
        m_pAudioClient(nullptr)
    {
        Init();
    }

    STDMETHOD(ActivateCompleted)(IActivateAudioInterfaceAsyncOperation *activateOperation)
    {
        IUnknown* pUnk = nullptr;

        m_hrActivateResult = E_UNEXPECTED;
        HRESULT hr = activateOperation->GetActivateResult(&m_hrActivateResult, &pUnk);

        if (m_hrActivateResult == S_OK)
            m_pAudioClient = reinterpret_cast<IAudioClient*>(pUnk);
        else
            m_pAudioClient = nullptr;

        m_bActivateFinished = true;
        m_bActivateFinished.notify_one();

        return S_OK;
    }

    // Will only ever be called by the MF worker thread. No need to check for riid.
    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObject)
    {
        *ppvObject = this;
        return S_OK;
    }

    STDMETHOD_(ULONG,AddRef)() // Ignored
    {
        return 0;
    }

    STDMETHOD_(ULONG, Release)() // Ignored
    {
        return 0;
    }

    // Blocking call until operation is finished.
    // Make sure that the operation was started successfully before calling this!
    void WaitForResult(HRESULT& hr, IAudioClient** pAudioClient)
    {
        m_bActivateFinished.wait(false);

        hr = m_hrActivateResult;
        *pAudioClient = m_pAudioClient;
    }

    // If you want to reuse the callback, call this first.
    void Init()
    {
        m_bActivateFinished = false;
        m_hrActivateResult = E_UNEXPECTED;
        m_pAudioClient = nullptr;
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

    m_pCallbackFunc(nullptr),
    m_pCallbackFuncUserData(nullptr),
    m_dwCallbackInterval(100),

    m_pMainAudioThread(nullptr),
    m_bRunMainAudioThread(false),
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

eCaptureError ProcessLoopbackCapture::SetCaptureFormat(unsigned int iSampleRate, unsigned int iBitDepth, unsigned int iChannelCount)
{
    if (m_CaptureState != eCaptureState::READY)
        return eCaptureError::STATE;

    bool valid = false;

    for (auto &v : LoopbackCaptureConst::VALID_SAMPLE_RATES)
    {
        if (iSampleRate == v)
        {
            valid = true;
            break;
        }
    }

    if(!valid)
        return eCaptureError::PARAM;

    if (iBitDepth == 0 || iBitDepth > 32 || (iBitDepth % 8) != 0)
        return eCaptureError::PARAM;

    if (iChannelCount < LoopbackCaptureConst::MIN_CHANNEL_COUNT || iChannelCount > LoopbackCaptureConst::MAX_CHANNEL_COUNT)
        return eCaptureError::PARAM;

    m_CaptureFormat.wFormatTag = WAVE_FORMAT_PCM;
    m_CaptureFormat.nChannels = (WORD)iChannelCount;
    m_CaptureFormat.nSamplesPerSec = (DWORD)iSampleRate;
    m_CaptureFormat.wBitsPerSample = (WORD)iBitDepth;
    m_CaptureFormat.nBlockAlign = m_CaptureFormat.nChannels * m_CaptureFormat.wBitsPerSample / 8;
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
    ActivateParams.blob.cbSize = sizeof(Blob);
    ActivateParams.blob.pBlobData = (BYTE*)&Blob;

    // Activate ("Async")

    IActivateAudioInterfaceAsyncOperation* Op;
    ActivateAudioInterfaceAsyncCallback AsyncCallback;

    m_hrLastError = ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, __uuidof(IAudioClient), &ActivateParams, &AsyncCallback, &Op);

    if (m_hrLastError != S_OK)
    {
        Reset();
        return eCaptureError::DEVICE;
    }

    AsyncCallback.WaitForResult(m_hrLastError, &m_pAudioClient);
    Op->Release(); // Enables re-running the capture on the same process infinite times (missing in original sample).

    if (m_hrLastError != S_OK)
    {
        Reset();
        return eCaptureError::ACTIVATION;
    }
    
    // Initialize (AudioClient is valid)

    m_hrLastError = m_pAudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY,
        0, // 100ns units, 10 = 1 micro, 10000 = 1 milli. Does not seem to do anything for this mode though (Win10).
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
        return eCaptureError::CAPTURE_SERVICE;
    }

    // Create and set event

    m_hSampleReadyEvent = CreateEventW(NULL, false, false, NULL);

    m_hrLastError = m_pAudioClient->SetEventHandle(m_hSampleReadyEvent);

    if(m_hrLastError != S_OK)
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

    StartThreads();

    m_CaptureState = eCaptureState::CAPTURING;

    return eCaptureError::NONE;
}

eCaptureError ProcessLoopbackCapture::StopCapture()
{
    if (m_CaptureState != eCaptureState::CAPTURING && m_CaptureState != eCaptureState::PAUSED)
        return eCaptureError::STATE;

    Reset();

    m_CaptureState = eCaptureState::READY;

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

    return eCaptureError::NONE;
}

eCaptureError ProcessLoopbackCapture::ResumeCapture()
{
    if (m_CaptureState != eCaptureState::PAUSED)
        return eCaptureError::STATE;

    m_CaptureState = eCaptureState::CAPTURING;

    m_hrLastError = m_pAudioClient->Start();

    if (m_hrLastError != S_OK)
        return eCaptureError::START;

    return eCaptureError::NONE;
}

HRESULT ProcessLoopbackCapture::GetLastErrorResult()
{
    return m_hrLastError;
}

// private

void ProcessLoopbackCapture::Reset()
{
    m_bRunMainAudioThread = false;
    m_bRunQueueAudioThread = false;

    if (m_pAudioClient != nullptr) // Prevent another event trigger 
    {
        m_pAudioClient->Stop();
    }

    if (m_pMainAudioThread != nullptr)
    {
        SetEvent(m_hSampleReadyEvent); // If the thread is currently (or will be) waiting for the event, we need to unblock it.

        m_pMainAudioThread->join();

        delete m_pMainAudioThread;
        m_pMainAudioThread = nullptr;
    }

    if (m_pQueueAudioThread != nullptr)
    {
        m_pQueueAudioThread->join();

        delete m_pQueueAudioThread;
        m_pQueueAudioThread = nullptr;
    }

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

    while (m_Queue.pop()); // Flush
    m_vecIntermediateBuffer.clear();
    m_vecIntermediateBuffer.shrink_to_fit();
}

void ProcessLoopbackCapture::StartThreads()
{
    m_bRunMainAudioThread = true;
    m_pMainAudioThread = new thread(&ProcessLoopbackCapture::ProcessMain, this);

    m_bRunQueueAudioThread = true;
    m_pQueueAudioThread = new thread(&ProcessLoopbackCapture::ProcessQueue, this);
}

void ProcessLoopbackCapture::ProcessMain()
{
    HANDLE hThread = GetCurrentThread();

    if(hThread)
        SetThreadPriority(hThread, THREAD_PRIORITY_TIME_CRITICAL);

    UINT32 FramesAvailable = 0;
    BYTE* pData = nullptr;
    DWORD dwCaptureFlags;
    UINT64 u64DevicePosition = 0;
    UINT64 u64QPCPosition = 0;
    DWORD cbBytesToCapture = 0;

    while (m_bRunMainAudioThread)
    {
        // The event is signaled if either a sample is ready or the capture was stopped.

        if (WaitForSingleObject(m_hSampleReadyEvent, 5000) == WAIT_OBJECT_0)
        {
            // The capture was stopped, exit thread.
            if (!m_bRunMainAudioThread)
                return;

            while (m_pAudioCaptureClient->GetBuffer(&pData, &FramesAvailable, &dwCaptureFlags, &u64DevicePosition, &u64QPCPosition) == S_OK)
            {
                cbBytesToCapture = FramesAvailable * m_CaptureFormat.nBlockAlign;

                for (DWORD i = 0; i < cbBytesToCapture; ++i)
                {
                    m_Queue.enqueue(pData[i]);
                }

                m_pAudioCaptureClient->ReleaseBuffer(FramesAvailable);
            }
        }
    }
}

void ProcessLoopbackCapture::ProcessQueue()
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

        size_t iSize = m_vecIntermediateBuffer.size() / m_CaptureFormat.nBlockAlign * m_CaptureFormat.nBlockAlign;

        if (iSize > 0)
        {
            if (m_pCallbackFunc != nullptr)
            {
                auto i1 = m_vecIntermediateBuffer.begin();
                auto i2 = m_vecIntermediateBuffer.begin() + iSize;
                m_pCallbackFunc(i1, i2, m_pCallbackFuncUserData);
            }

            m_vecIntermediateBuffer.erase(m_vecIntermediateBuffer.begin(), m_vecIntermediateBuffer.begin() + iSize);
        }

        Sleep(m_dwCallbackInterval);
    }
}

// ----------------------------------------------------------------------- EOF