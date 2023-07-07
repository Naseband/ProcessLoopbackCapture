# Windows Process Loopback Capture

A relatively simple class that uses Windows' AudioClient in loopback capture mode to capture audio coming from one process only.

Calls a user-provided callback with the resulting audio data in little endian unsigned char PCM format.
Supports 8, 16, 24 and 32 bit and any common Sample Rate up to 384,000Hz.

Based on the [original ApplicationLoopbackCapture example](https://github.com/microsoft/windows-classic-samples/tree/main/Samples/ApplicationLoopback) but does not require WIL/WRL to be installed or included.
Also fixes a few bugs and leaks. This version allows the AudioClient to be restarted at any point, including the same process.

Uses [cameron314's readerwriterqueue](https://github.com/cameron314/readerwriterqueue) to transport samples from the main audio thread to the user callback via a helper thread.
This allows the end-user callback to be non-time-critical.

# Usage

To start capturing an application, create an instance of the ProcessLoopbackCapture class and a callback that it can use.

The ProcessLoopbackCapture class can be instantiated globally, locally or allocated and deleted dynamically.

``` 
#include <readerwriterqueue.h>
#include <ProcessLoopbackCapture.h>

ProcessLoopbackCapture LoopbackCapture;

std::vector<unsigned char> g_MySampleData;

void MyCallback(std::vector<unsigned char>::iterator &i1, std::vector<unsigned char>::iterator &i2, void *pUserData)
{
    // This callback is not time critical. There will be no audio glitches if you take longer than the callback interval.
    // In a real application you would also lock a mutex here.

    g_MySampleData.insert(g_MySampleData.end(), i1, i2);
}
```

Somewhere in your code you can then set up the format, process id and callback and start the capture:

The format, callback and process id can be changed only while the capture is stopped.
Unlike the original Windows Sample, the AudioClient will start correctly even if you start it on the same process again.

```
DWORD dwProcessId = GetMeSomeProcessId(); // Your Process ID

LoopbackCapture.SetCaptureFormat(44100, 16, 2);
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
``` 

From there you can use StopCapture, PauseCapture, ResumeCapture and GetState to control the capture. GetState is thread safe, but all other functions must be called on the same thread that StartCapture was called on.

If StartCapture returns an error, the capture client is reset completely and you are free to try again.

A ProcessLoopbackCapture instance can also be destroyed at any time. The AudioClients will be unitialized properly.

# Notes

StartCapture is a blocking operation. On some systems, depending on load and device activity, there may be times where it takes a few seconds to finish.

For all functions and their parameters see ProcessLoopbackCapture.h.
