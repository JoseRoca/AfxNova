# CDSAudio Class

The `CDSAudio` class allows playing audio files of a variety of formats using **Direct Show**.

**Include file**: Afx/CDSAudio.inc

### Constructor

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructor) | Creates an instance of the **CDSAudio** class and loads the specified audio file. |

### Methods

| Name       | Description |
| ---------- | ----------- |
| [GetBalance](#getbalance) | Gets the balance for the audio signal. |
| [CurrentPosition](#currentposition) | Gets the current position, relative to the total duration of the stream. |
| [GetDuration](#getduration) | Gets the duration of the stream, in 100-nanosecond units. |
| [GetEvent](#getevent) | Retrieves the next event notification from the event queue. |
| [GetVolume](#getvolume) | Gets the volume (amplitude) of the audio signal. |
| [Load](#load) | Builds a filter graph that renders the specified file. |
| [Pause](#pause) | Pauses all the filters in the filter graph. |
| [Run](#run) | Runs all the filters in the filter graph. |
| [SetBalance](#setbalance) | Sets the balance for the audio signal. |
| [SetNotifyWindow](#setnotifywindow) | Registers a window to process event notifications. |
| [SetPositions](#setpositions) | Sets the current position and the stop position. |
| [SetVolume](#setvolume) | Sets the volume (amplitude) of the audio signal. |
| [Stop](#stop) | Stops all the filters in the filter graph. |
| [WaitForCompletion](#waitforcompletion) | Waits for the filter graph to render all available data. |

---

## Constructor

Creates an instance of the `CDSAudio` class and loads the specified audio file.

```
CONSTRUCTOR CDSAudio (BYREF wszFileName AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | The file to load. |

You can create an instance of the class and load the file at the same time with:

```
DIM pCDSAudio AS CDSAudio = "MyAudioFile.mp3"
```

or you can use the default constructor and then load the file:

```
DIM pCDSAudio AS CDSAudio
pCDSAudio.Load("MyAudioFile.mp3")
```

With the **Load** method you can change the audio file on the fly.

To receive event messages, you can define a custom message

```
#define WM_GRAPHNOTIFY  WM_APP + 1
```

and pass it to the **SetNotifyWindow** method, together with the handle of the window that will process the message:

```
pCDSAudio.SetNotifyWindow(pWindow.hWindow, WM_GRAPHNOTIFY, 0)
```

The messages will be processed in the window callback procedure:

```
      CASE WM_GRAPHNOTIFY
         DIM AS LONG lEventCode
         DIM AS LONG_PTR lParam1, lParam2
         WHILE (SUCCEEDED(pCDSAudio.GetEvent(lEventCode, lParam1, lParam2)))
            SELECT CASE lEventCode
               CASE EC_COMPLETE:     ' Handle event
               CASE EC_USERABORT:    ' Handle event
               CASE EC_ERRORABORT:   ' Handle event
            END SELECT
         WEND
```

Once loaded, you can start playing the audio file calling the **Run** method:

```
pCDSAudio.Run
```

There are other methods to get/set the volume and balance, to get the duration and current position, to set the positions, and to pause or stop.

#### Example

```
' ########################################################################################
' Microsoft Windows
' File: CW_CDSAudio_01.bas
' Contents: CWindow - Direct Show audio
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
#INCLUDE ONCE "AfxNova/CDSAudio.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_Play = 1001
CONST IDC_Pause = 1002
CONST IDC_Stop = 1003

#define WM_GRAPHNOTIFY  WM_APP + 1

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   ' // Creates the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - CDSAudio Test", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center

   ' // Adds buttons
   pWindow.AddControl("Button", hWin, IDC_Play, "Play", 50, 30, 100, 23)
   pWindow.AddControl("Button", hWin, IDC_Pause, "Pause", 50, 60, 100, 23)
   pWindow.AddControl("Button", hWin, IDC_Stop, "Stop", 50, 90, 100, 23)
   pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", 270, 155, 75, 30)

   ' // Anchors the cancel button to the bottom and the right side of the main window
   pWindow.AnchorControl(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)
         
   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

STATIC pCDSAudio AS CDSAudio PTR

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         ' // Create an instance of the CDSAudio and load an audio file
         pCDSAudio = NEW CDSAudio("Kalimba.mp3")
         ' // Registers a window to process event notifications
         pCDSAudio->SetNotifyWindow(hwnd, WM_GRAPHNOTIFY, 0)
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
            CASE IDC_Play
               ' // Play the audio
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN pCDSAudio->Run
            CASE IDC_Pause
               ' // Pause the audio
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN pCDSAudio->Pause
            CASE IDC_Stop   : pCDSAudio->Stop
               ' // Stop the audio
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN pCDSAudio->Stop
         END SELECT
         RETURN 0

      CASE WM_GRAPHNOTIFY
         DIM evCode AS LONG, lParam1 AS LONG_PTR, lParam2 AS LONG_PTR
         WHILE (SUCCEEDED(pCDSAudio->GetEvent(evCode, lParam1, lParam2)))
            SELECT CASE evCode
               CASE EC_COMPLETE:   print evCode ' Handle event
               CASE EC_USERABORT:  print evCode   ' Handle event
               CASE EC_ERRORABORT: print evCode  ' Handle event
            END SELECT
         WEND
         RETURN 0

      CASE WM_DESTROY
         ' // Destroy the class
         IF pCDSAudio THEN DELETE pCDSAudio
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
```

---

## GetBalance

Gets the balance for the audio signal.

```
FUNCTION GetBalance () AS LONG
```

#### Return value

The balance for the audio signal.

#### Remarks

The balance ranges from -10,000 to 10,000. The value -10,000 means the right channel is attenuated by 100 dB and is effectively silent. The value 10,000 means the left channel is silent. The neutral value is 0, which means that both channels are at full volume. When one channel is attenuated, the other remains at full volume.

---

## CurrentPosition

Gets the current position, relative to the total duration of the stream.

```
FUNCTION CurrentPosition () AS LONG
```
---

## GetDuration

Gets the duration of the stream, in 100-nanosecond units.

```
FUNCTION GetDuration () AS LONG
```
---

## GetEvent

Retrieves the next event notification from the event queue.

```
FUNCTION GetEvent(BYREF lEventCode AS LONG, BYREF lParam1 AS LONG_PTR, _
   BYREF lParam2 AS LONG_PTR, BYVAL msTimeout AS LONG = 0) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *lEventCode* | Out. Variable that receives the event code. |
| *lParam1* | Out. Variable that receives the first event parameter. |
| *lParam2* | Out. Variable that receives the second event parameter. |
| *msTimeout* | In. Time-out interval, in milliseconds. Use INFINITE to block until there is an event. |

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| E_ABORT | Timeout expired. |
| E_POINTER | The **IMediaEventEx** interface pointer is null. |

---

## GetVolume

Gets the volume (amplitude) of the audio signal.

```
FUNCTION GetVolume () AS LONG
```

#### Return value

The volume (amplitude) of the audio signal. Specifies the volume, as a number from –10,000 to 0, inclusive. Full volume is 0, and –10,000 is silence. Multiply the desired decibel level by 100. For example, –10,000 = –100 dB.

---

## Load

Builds a filter graph that renders the specified file.

```
FUNCTION Load (BYREF wszFileName AS WSTRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFilename* | The file to load. |

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| VFW_S_AUDIO_NOT_RENDERED | Partial success; the audio was not rendered. |
| VFW_S_DUPLICATE_NAME | Success; the Filter Graph Manager modified the filter name to avoid duplication. |
| VFW_S_PARTIAL_RENDER | Some of the streams in this movie are in an unsupported format. |
| VFW_S_VIDEO_NOT_RENDERED | Partial success; some of the streams in this movie are in an unsupported format. |
| E_ABORT | Operation aborted. |
| E_FAIL | Failure. |
| E_INVALIDARG | Argument is invalid. |
| E_OUTOFMEMORY | Insufficient memory. |
| E_POINTER | NULL pointer argument.. |
| VFW_E_CANNOT_CONNECT | No combination of intermediate filters could be found to make the connection.. |
| VFW_E_CANNOT_LOAD_SOURCE_FILTER | The source filter for this file could not be loaded.. |
| VFW_E_CANNOT_RENDER | No combination of filters could be found to render the stream.. |
| VFW_E_INVALID_FILE_FORMAT | The file format is invalid.. |
| VFW_E_NOT_FOUND | An object or name was not found.. |
| VFW_E_UNKNOWN_FILE_TYPE | The media type of this file is not recognized.. |
| VFW_E_UNSUPPORTED_STREAM | Cannot play back the file: the format is not supported.. |

---

## Pause

Pauses all the filters in the filter graph.

```
FUNCTION Pause () AS HRESULT
```

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | All filters in the graph completed the transition to a paused state. |
| S_FALSE | The graph paused successfully, but some filters have not completed the state transition. |
| E_POINTER | The **IMediaControl** interface pointer is null. |

#### Remarks

Pausing the filter graph cues the graph for immediate rendering when the graph is next run. While the graph is paused, filters process data but do not render it. Data is pushed through the graph and processed by transform filters as far as buffering permits, but renderer filters do not render the data. However, video renderers display a static poster frame of the first sample.

If the method returns S_FALSE, call the **GetState** method to wait for the state transition to complete, or to check if the transition has completed. When you call **Pause** to display the first frame of a video file, always follow it immediately with a call to **GetState** to ensure that the state transition has completed. Failure to do this can result in the video rectangle being painted black.

If the method fails, it stops the graph before returning.

---

## Run

Runs all the filters in the filter graph.

```
FUNCTION Run () AS HRESULT
```

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | All filters in the graph completed the transition to a running state. |
| S_FALSE | The graph is preparing to run, but some filters have not completed the transition to a running state. |
| E_POINTER | The **IMediaControl** interface pointer is null. |

#### Remarks

If the filter graph is stopped, this method pauses the graph before running. If the graph is already running, the method returns S_OK but has no effect.

The graph runs until the application calls the **Pause** or **Stop** method. When playback reaches the end of the stream, the graph continues to run, but the filters do not stream any more data. At that point, the application can pause or stop the graph.

This method does not seek to the beginning of the stream. Therefore, if you run the graph, pause it, and then run it again, playback resumes from the paused position. If you run the graph after it has reached the end of the stream, nothing is rendered. To seek the graph, use the the **GetCurrentPosition** and **SetPositions** methods.

If method returns S_FALSE, it means that the method returned before all of the filters switched to a running state. The filters will complete the transition after the method returns. Optionally, you can wait for the transition to complete by calling the **GetState** method with a timeout value. However, this is not required.

If the **Run** method returns an error code, it means that one or more filters failed to run. However, some filters might be in a running state. In a multistream graph, entire streams might be playing successfully. Typically the application would tear down the graph and report an error in this case.

---

## SetBalance

Sets the balance for the audio signal.

```
FUNCTION SetBalance (BYVAL nBalance AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nBalance* | The balance ranges from -10,000 to 10,000. The value -10,000 means the right channel is attenuated by 100 dB and is effectively silent. The value 10,000 means the left channel is silent. The neutral value is 0, which means that both channels are at full volume. When one channel is attenuated, the other remains at full volume. |

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| E_FAIL | The underlying audio device returned an error. |
| E_INVALIDARG | The value of *nBalance* is invalid. |
| E_NOTIMPL | The filter graph does not contain an audio renderer filter. (Possibly the source does not contain an audio stream.) |
| E_POINTER | The **IBasicAudio** interface pointer is null. |

---

## SetNotifyWindow

The **SetNotifyWindow** method registers a window to process event notifications.

```
FUNCTION SetNotifyWindow (BYVAL hwnd AS HWND, BYVAL lMsg AS LONG, BYVAL lParam AS LONG_PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window. |
| *lMsg* | Window message to be passed as the notification. |
| *lInstanceData* | Value to be passed as the lParam parameter for the lMsg message. |

#### Return value

Returns S_OK if successful or E_INVALIDARG if the hwnd parameter is not a valid handle to a window.

---

## SetPositions

Sets the current position and the stop position.

```
FUNCTION SetPositions (BYREF pCurrent AS LONGLONG, BYREF pStop AS LONGLONG, _
   BYVAL bAbsolutePositioning AS BOOLEAN) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCurrent* | In/Out. Pointer to a variable that specifies the current position, in units of the current time format. |
| *pStop* | In/Out. Pointer to a variable that specifies the stop time, in units of the current time format. |
| *bAbsolutePositioning* | In. TRUE or FALSE. If TRUE, the specified position is absolute. If FALSE, the specified position is relative. |

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| S_FALSE | No position change. (Both flags specify no seeking.) |
| E_INVALIDARG | Invalid argument. |
| E_NOTIMPL | Method is not supported. |
| E_POINTER | NULL pointer argument. |

---

## SetVolume

Sets the volume (amplitude) of the audio signal.

```
FUNCTION SetVolume (BYVAL nVol AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nVol* | The volume (amplitude) of the audio signal. Specifies the volume, as a number from –10,000 to 0, inclusive. Full volume is 0, and –10,000 is silence. Multiply the desired decibel level by 100. For example, –10,000 = –100 dB. |

---

## Stop

Stops all the filters in the filter graph.

```
FUNCTION Stop () AS HRESULT
```

#### Return value

Returns S_OK if successful, or an HRESULT value that indicates the cause of the error.

#### Remarks

If the graph is running, this method pauses the graph before stopping it. While paused, video renderers can copy the current frame to display as a poster frame.

This method does not seek to the beginning of the stream. If you call this method and then call the **Run** method, playback resumes from the stopped position. To seek, use the **GetCurrentPosition** and **SetPositions** methods.

The Filter Graph Manager pauses all the filters in the graph, and then calls the **Stop** method on all filters, without waiting for the pause operations to complete. Therefore, some filters might have their **Stop** method called before they complete their pause operation. If you develop a custom rendering filter, you might need to handle this case by pausing the filter if it receives a stop command while in a running state. However, most filters do not need to take any special action in this regard.

---

## WaitForCompletion

Waits for the filter graph to render all available data.

```
FUNCTION WaitForCompletion (BYVAL msTimeout AS LONG, BYREF EvCode AS LONG) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *msTimeout* | Duration of the time-out, in milliseconds. Pass zero to return immediately. To block indefinitely, pass INFINITE. |
| *EvCode* | Out. Event that terminated the wait. This value can be one of the following:<br>**EC_COMPLETE** : Operation completed.<br>**EC_ERRORABORT** : Error. Playback cannot continue.<br>**EC_USERABORT** : User terminated the operation.<br>**Zero** (0) = Operation has not completed. |

#### Return value

| Result code | Description |
| ----------- | ----------- |
| S_OK | Success. |
| E_ABORT | Time-out expired. |
| E_POINTER | The **IMediaEventEx** interface is null. |
| VFW_E_WRONG_STATE | The filter graph is not running. |

#### Remarks

This method blocks until the time-out expires, or one of the following events occurs:

**EC_COMPLETE**<br>
**EC_ERRORABORT**<br>
**EC_USERABORT**

During the wait, the method discards all other event notifications.

If the return value is S_OK, the *EvCode* parameter receives the event code that ended the wait. When the method returns, the filter graph is still running. The application can pause or stop the graph, as appropriate.

---
