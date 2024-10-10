# Windows Process Loopback Capture

A relatively simple class that uses Windows' AudioClient in loopback capture mode to capture audio coming from one process and its children only (Windows 10/11).

Calls a user-provided callback with the resulting audio data in little endian unsigned char PCM format.
Supports 8-32 bit PCM Int/Float and any Sample Rate above 1 kHz supported by WASAPI's SRC.

Based on the [original ApplicationLoopbackCapture example](https://github.com/microsoft/windows-classic-samples/tree/main/Samples/ApplicationLoopback) but does not require WIL/WRL to be installed or included.
Also fixes a few bugs and leaks. This version allows the AudioClient to be restarted at any point, including the same process.

(Optionally) uses [cameron314's readerwriterqueue](https://github.com/cameron314/readerwriterqueue) to transport samples from the main audio thread to the user callback via a helper thread.
This allows the user callback to be non-time-critical at the cost of a delay between the actual audio output and the callback. This is usually the best option if you plan on storing audio data without worrying about audio glitches due to allocation or io operations.

By default, the queue is not used and there is no dependency on readerwriterqueue.

If you wish to enable it, define PROCESS_LOOPBACK_CAPTURE_USE_QUEUE as a preprocessor macro and use SetIntermediateThreadEnabled(true).

In order to successfully start an AudioClient you need to initialize COM (call CoInitialize(Ex), found in combaseapi.h).
Note that MFStartup also initializes COM if you do already use it, however the actual MF functionality is not required.

You will also need to link against mmdevapi.lib (for ActivateAudioInterfaceAsync) and avrt.lib (for setting thread characteristics).

# Usage

To start capturing an application, create an instance of the ProcessLoopbackCapture class and define a callback that it can use.

``` 
#include <combaseapi.h> // For CoInitialie(Ex)
#include <ProcessLoopbackCapture.h>

ProcessLoopbackCapture LoopbackCapture;

std::vector<unsigned char> g_MySampleData;

void MyCallback(const std::vector<unsigned char>::iterator &i1, const std::vector<unsigned char>::iterator &i2, void *pUserData)
{
    // By default, this callback is time critical. There will be audio glitches if you take longer than the device interval.
    // If the use of the queue is enabled, you can take as long as you wish at the cost of a growing queue size.

    // Some synchronization ...

    g_MySampleData.insert(g_MySampleData.end(), i1, i2);
}
```

Somewhere in your code you can then set up the format, process id and callback and start the capture.

```
CoInitializeEx(NULL, COINIT_MULTITHREADED); // Initialize COM

DWORD dwProcessId = GetMeSomeProcessId(); // Your Process ID

LoopbackCapture.SetCaptureFormat(44100, 16, 2, WAVE_FORMAT_IEEE_FLOAT); // Supports WAVE_FORMAT_PCM and WAVE_FORMAT_IEEE_FLOAT
LoopbackCapture.SetTargetProcess(dwProcessId);
LoopbackCapture.SetCallback(&MyCallback);

eCaptureError eError = LoopbackCapture.StartCapture();

if(eError != eCaptureError::NONE)
{
    // Get Error Text rather than just a code:
    std::cout << "Failed to Start Capture: " << LoopbackCaptureConst::GetErrorText(eError) << std::endl;

    // Get result if the cause for the error is a Windows interface
    std::cout << "HRESULT: " << std::hex << LoopbackCapture.GetLastErrorResult() << std::dec << std::endl;

    return 0; // Fail
}

// We are now capturing!
// ...

LoopbackCature.StopCapture();

CoUninitialize();
``` 

You can use StopCapture, PauseCapture, ResumeCapture and GetState to control the capture. GetState is thread safe, but all other functions must be called on the same thread that StartCapture was called on.

If StartCapture returns an error, the capture client is reset completely and you are free to try again.

A ProcessLoopbackCapture instance can also be destroyed at any time. The AudioClient will be stopped properly.
Again, make sure that if destroying the object would cause the capture to be stopped, you will need to destroy it on the same thread that *started* the capture.
This is a requirement from the WASAPI engine. Not doing so can lead to random errors and crashes.

# Notes

StartCapture is a blocking operation. On some systems, depending on load and device activity, there may be times where it takes a few hundred milliseconds to execute.

For fast capture starting/stopping without changing any settings, you can use PauseCapture and ResumeCapture.

For all functions, their parameters and notes see comments in the header file.
