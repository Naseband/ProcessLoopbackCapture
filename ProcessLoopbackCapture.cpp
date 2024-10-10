#include <ProcessLoopbackCapture.h>

#include <chrono>

#include <mmdeviceapi.h>
#include <mfapi.h>
#include <audioclientactivationparams.h>
#include <audiopolicy.h>
#include <avrt.h>

using namespace std;

// ------------------------------------------------------------ 

// Simple class that turns ActivateAudioInterfaceAsync into a blocking operation.
// Use an instance of this class for ActivateAudioInterfaceAsync.
// Afterwards call AsyncCallback::Wait, which blocks until completed.
// The AudioInterface pointer can be obtained from the IActivateAudioInterfaceAsyncOperation pointer before releasing it.

class ActivateAudioInterfaceAsyncCallback : 
    public IActivateAudioInterfaceCompletionHandler
{
public:

    ActivateAudioInterfaceAsyncCallback() :
        m_bActivateCompleted(false)
    {

    }

    void Wait()
    {
        m_bActivateCompleted.wait(false);
        m_bActivateCompleted = false;
    }

protected:

    atomic<bool> m_bActivateCompleted;

    STDMETHOD(ActivateCompleted)(IActivateAudioInterfaceAsyncOperation *activateOperation)
    {
        m_bActivateCompleted = true;
        m_bActivateCompleted.notify_all();

        return S_OK;
    }

    STDMETHOD(QueryInterface)(const IID &riid, void **ppvObject)
    {
        if (riid == __uuidof(IAgileObject)) // Removed IUnknown
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

// ------------------------------------------------------------ ProcessLoopbackCapture

// public

ProcessLoopbackCapture::ProcessLoopbackCapture() :
    m_hrLastError(S_OK),

    m_CaptureState(eCaptureState::READY),
    m_pAudioClient(nullptr),
    m_pAudioCaptureClient(nullptr),
    m_hSampleReadyEvent(NULL),
    m_bCaptureFormatInitialized(false),
    m_dwProcessId(0),
    m_bProcessInclusive(false),
    m_bUseIntermediateThread(false),

    m_pCallbackFunc(nullptr),
    m_pCallbackFuncUserData(nullptr),
    m_dwCallbackInterval(100),

    m_bRunAudioThreads(false),

    m_pMainAudioThread(nullptr),
    m_dwMainThreadBytesToSkip(0),
    m_fMaxExecutionTime(0.0),

    m_pQueueAudioThread(nullptr)

#if defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE
    ,
    m_Queue(8192)
#endif
{
    
}

ProcessLoopbackCapture::~ProcessLoopbackCapture()
{
    StopCapture();
}

eCaptureError ProcessLoopbackCapture::SetCaptureFormat(unsigned int iSampleRate, unsigned int iBitDepth, unsigned int iChannelCount, unsigned int iFormatTag)
{
    if (m_CaptureState != eCaptureState::READY)
        return eCaptureError::STATE;

    if (iSampleRate < 1000)
        return eCaptureError::PARAM;

    if (iBitDepth == 0 || iBitDepth > 32 || (iBitDepth % 8) != 0)
        return eCaptureError::PARAM;

    if (iChannelCount < 1 || iChannelCount > 1024)
        return eCaptureError::PARAM;

    if (iFormatTag == WAVE_FORMAT_IEEE_FLOAT)
        iBitDepth = 32;
    else if (iFormatTag != WAVE_FORMAT_PCM)
        return eCaptureError::PARAM;

    m_CaptureFormat.wFormatTag = (WORD)iFormatTag;
    m_CaptureFormat.nChannels = (WORD)iChannelCount;
    m_CaptureFormat.nSamplesPerSec = (DWORD)iSampleRate;
    m_CaptureFormat.wBitsPerSample = (WORD)iBitDepth;
    m_CaptureFormat.nBlockAlign = m_CaptureFormat.wBitsPerSample / 8 * m_CaptureFormat.nChannels;
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

eCaptureError ProcessLoopbackCapture::SetCallback(void (*pCallbackFunc)(const std::vector<unsigned char>::iterator&, const std::vector<unsigned char>::iterator&, void*), void *pUserData)
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

eCaptureError ProcessLoopbackCapture::SetIntermediateThreadEnabled(bool bEnable)
{
#if !defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE
    return eCaptureError::NOT_AVAILABLE;
#endif

    if (m_CaptureState != eCaptureState::READY)
        return eCaptureError::STATE;

    m_bUseIntermediateThread = bEnable;

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

    // Set up Params

    AUDIOCLIENT_ACTIVATION_PARAMS blob{};
    blob.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
    blob.ProcessLoopbackParams.ProcessLoopbackMode = m_bProcessInclusive ? PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE : PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE;
    blob.ProcessLoopbackParams.TargetProcessId = m_dwProcessId;

    PROPVARIANT activation_params{};
    activation_params.vt = VT_BLOB;
    activation_params.blob.cbSize = sizeof(AUDIOCLIENT_ACTIVATION_PARAMS);
    activation_params.blob.pBlobData = (BYTE*)&blob;

    // Activate ("Async")

    IActivateAudioInterfaceAsyncOperation* activation_operation = nullptr;
    ActivateAudioInterfaceAsyncCallback async_callback;

    m_hrLastError = ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, __uuidof(IAudioClient), &activation_params, &async_callback, &activation_operation);

    if (m_hrLastError != S_OK)
    {
        Reset();
        return eCaptureError::DEVICE;
    }

    async_callback.Wait();

    m_hrLastError = activation_operation->GetActivateResult(&m_hrLastError, reinterpret_cast<IUnknown**>(&m_pAudioClient));
    activation_operation->Release();

    if (m_hrLastError != S_OK)
    {
        Reset();
        return eCaptureError::ACTIVATION;
    }
    
    // Initialize (AudioClient is valid)

    m_hrLastError = m_pAudioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY,
        0, // buffer duration (100ns units), 10 = 1 micro, 10000 = 1 milli. Does not seem to do anything for this mode (Win10).
        0, // device periodicty, do not use for Capture Clients.
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

eCaptureError ProcessLoopbackCapture::GetQueueSize(size_t& iSize)
{
#if defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE

    if (!m_bUseIntermediateThread)
    {
        iSize = 0U;
        return eCaptureError::NOT_AVAILABLE;
    }

    iSize = m_Queue.size_approx();
    return eCaptureError::NONE;

#else

    iSize = 0U;
    return eCaptureError::NOT_AVAILABLE;

#endif
}

// private

void ProcessLoopbackCapture::Reset()
{
    StopThreads();

    if (m_CaptureState == eCaptureState::CAPTURING)
    {
        m_pAudioClient->Stop();
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

    m_CaptureState = eCaptureState::READY;
}

void ProcessLoopbackCapture::StartThreads(double fInitialDurationToSkip)
{
    if (m_bRunAudioThreads)
        return;

    if (fInitialDurationToSkip < 0.0)
        fInitialDurationToSkip = 0.0;

    m_dwMainThreadBytesToSkip = (DWORD)(m_CaptureFormat.nSamplesPerSec * fInitialDurationToSkip) * (DWORD)m_CaptureFormat.nBlockAlign;

    m_bRunAudioThreads = true;

#if defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE

    if (m_bUseIntermediateThread)
    {
        m_pMainAudioThread = new thread(&ProcessLoopbackCapture::ProcessMainToQueue, this);
        m_pQueueAudioThread = new thread(&ProcessLoopbackCapture::ProcessIntermediate, this);
    }
    else
    {
        m_pMainAudioThread = new thread(&ProcessLoopbackCapture::ProcessMainToCallback, this);
        m_pQueueAudioThread = nullptr;
    }

#else

    m_pMainAudioThread = new thread(&ProcessLoopbackCapture::ProcessMainToCallback, this);
    m_pQueueAudioThread = nullptr;

#endif
}

void ProcessLoopbackCapture::StopThreads()
{
    if (!m_bRunAudioThreads)
        return;

    m_bRunAudioThreads = false;

#if defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE

    m_pMainAudioThread->join();

    if (m_bUseIntermediateThread)
    {
        m_pQueueAudioThread->join();
        delete m_pQueueAudioThread;
        m_pQueueAudioThread = nullptr;
    }

    delete m_pMainAudioThread;
    m_pMainAudioThread = nullptr;

    while (m_Queue.pop()); // Flush

#else

    m_pMainAudioThread->join();
    delete m_pMainAudioThread;
    m_pMainAudioThread = nullptr;

#endif

    m_AudioData.clear();
    m_AudioData.shrink_to_fit();
}

void ProcessLoopbackCapture::ProcessMainToCallback()
{
    DWORD dwTaskIndex = 0;
    HANDLE hTaskHandle = AvSetMmThreadCharacteristicsW(L"Pro Audio", &dwTaskIndex);

    BYTE* pData = nullptr;
    UINT32 iFramesAvailable;
    UINT64 iBytesAvailable;
    DWORD dwCaptureFlags;

    while (m_bRunAudioThreads)
    {
        // The event is signaled if either a sample is ready or the capture was stopped.

        if (WaitForSingleObject(m_hSampleReadyEvent, 50) == WAIT_OBJECT_0)
        {
            // The capture was stopped, exit thread.
            if (!m_bRunAudioThreads)
                break;

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
                        m_AudioData.emplace_back(pData[i]);
                    }
                }

                m_pAudioCaptureClient->ReleaseBuffer(iFramesAvailable);
            }

            if (m_AudioData.size() > 0)
            {
                if (m_pCallbackFunc != nullptr)
                {
                    auto i1 = m_AudioData.begin();
                    auto i2 = m_AudioData.begin() + m_AudioData.size();
                    m_pCallbackFunc(i1, i2, m_pCallbackFuncUserData);
                }

                m_AudioData.clear();
            }

            auto dur = chrono::duration_cast<chrono::nanoseconds>(chrono::steady_clock::now() - tick_start).count() / 1e6;

            if (dur > m_fMaxExecutionTime)
                m_fMaxExecutionTime = dur;
        }
    }

    if (hTaskHandle)
        AvRevertMmThreadCharacteristics(hTaskHandle);
}

#if defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE

void ProcessLoopbackCapture::ProcessMainToQueue()
{
    DWORD dwTaskIndex = 0;
    HANDLE hTaskHandle = AvSetMmThreadCharacteristicsW(L"Pro Audio", &dwTaskIndex);

    BYTE* pData = nullptr;
    UINT32 iFramesAvailable;
    UINT64 iBytesAvailable;
    DWORD dwCaptureFlags;

    while (m_bRunAudioThreads)
    {
        // The event is signaled if either a sample is ready or the capture was stopped.

        if (WaitForSingleObject(m_hSampleReadyEvent, 50) == WAIT_OBJECT_0)
        {
            // The capture was stopped, exit thread.
            if (!m_bRunAudioThreads)
                break;

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

    if (hTaskHandle)
        AvRevertMmThreadCharacteristics(hTaskHandle);
}

void ProcessLoopbackCapture::ProcessIntermediate()
{
    while (m_bRunAudioThreads)
    {
        // Get all data from the queue into intermediate buffer and pass it to the callback

        unsigned char Data;

        while (m_Queue.try_dequeue(Data))
        {
            m_AudioData.emplace_back(Data);
        }

        // Get buffer size and send to callback, then delete
        // Make sure the temp. buffer is aligned, the queue can contain stray bytes

        size_t iAlignedSize = m_AudioData.size() / m_CaptureFormat.nBlockAlign * m_CaptureFormat.nBlockAlign;

        if (iAlignedSize > 0)
        {
            if (m_pCallbackFunc != nullptr)
            {
                auto i1 = m_AudioData.begin();
                auto i2 = m_AudioData.begin() + iAlignedSize;
                m_pCallbackFunc(i1, i2, m_pCallbackFuncUserData);
            }

            m_AudioData.erase(m_AudioData.begin(), m_AudioData.begin() + iAlignedSize);
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(m_dwCallbackInterval));
    }
}

#endif // #if defined PROCESS_LOOPBACK_CAPTURE_USE_QUEUE

// ------------------------------------------------------------ EOF