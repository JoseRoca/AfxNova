# About GDI+

GDI+ is the successor to GDI, the graphics device interface included with earlier versions of Windows.

A graphics device interface, such as GDI+, allows application programmers to display information on a screen or printer without having to be concerned about the details of a particular display device. The application programmer makes calls to the provided functions and those functions in turn make the appropriate calls to specific device drivers. GDI+ insulates the application from the graphics hardware, and it is this insulation that allows developers to create device-independent applications.

Microsoft Windows GDI+ exposes a flat application programming interface (API) that consists of about 600 functions, which are implemented in Gdiplus.dll and declared in Gdiplusflat.h. Microsoft has wrapped these functions with a collection of about 40 C++ classes and the Microsoft Product Support Services will not provide support for code that calls the flat API directly.

The functions in the flat API return a status value that indicates whether the call has succeeded or not.

Documentation MSDN Documentation: [GDI+](https://learn.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-about-gdi--about)

# GDI+ Flat Api

Microsoft Windows GDI+ exposes a flat application programming interface (API) that consists of about 600 functions, which are implemented in Gdiplus.dll and declared in Gdiplusflat.h.

The functions in the flat API return a value of type **GpStatus**, identical to the **Status** enumeration.

#### Template for the examples

```
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.bi"
#INCLUDE ONCE "AfxNova/ARGB_Colors.bi"
#INCLUDE ONCE "AfxNova/CGraphCtx.inc"
USING AfxNova

CONST IDC_GRCTX = 1001

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG

   END wWinMain(GetModuleHandleW(NULL), NULL, wCOMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Put here the example code
' ========================================================================================
SUB Example_XXXX (BYVAL hdc AS HDC)
END SUB
' ========================================================================================

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

   ' // Create the main window
   DIM pWindow AS CWindow = "MyClassName"
'   pWindow.CreateOverlapped(NULL, "GDI+ CreateAdjustableArrowCap", @WndProc)
   pWindow.Create(NULL, "GDI+ CreateAdjustableArrowCap", @WndProc)
   ' // Size it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 250)
   ' // Center the window
   pWindow.Center

   ' // Add a graphic control
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   pGraphCtx.Clear RGB_FLORALWHITE
   ' // Anchor the control
   pWindow.AnchorControl(pGraphCtx.hWindow, AFX_ANCHOR_HEIGHT_WIDTH)
   
   ' // Get the memory device context of the graphic control
   DIM hdc AS HDC = pGraphCtx.GetMemDc

   ' // Initialize GDI+
   DIM token AS ULONG_PTR = AfxGdipInit

   ' // Draw the graphics
   Example_CreateAdjustableArrowCap(hdc)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

   ' // Shutdown GDI+
   AfxGdipShutdown token

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      ' // Theme has changed
      CASE WM_THEMECHANGED
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  RETURN 0
               END IF
         END SELECT

    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
```
---

# Functions

| Name       | Description |
| ---------- | ----------- |
| [GdiplusStartup](#gdiplusstartup) | Initializes Microsoft Windows GDI+. |
| [GdiplusShutdown](#gdiplusshutdown) | Cleans up resources used by Windows GDI+. |
| [GdipAlloc](#gdipalloc) | Allocates memory for GDI+ objects. |
| [GdipFree](#gdipfree) | Frees memory allocated for GDI+ objects. |

---

#  Wrapper Functions

| Name       | Description |
| ---------- | ----------- |
| [AfxGdipInit](#afxgdipinit) | Initilizes GDI+. |
| [AfxGdipShutdown](#afxgdipshutdown) | Cleans up resources used by Windows GDI+. |
| [AfxGdipStatusStr](#afxgdipstatusstr) | Returns the description of a GdiPlus status code. |
| [ObjectTypeIsValid](#objecttypeisvalid) | Determines whether an element of the **ObjectType** enumeration represents a valid object type. |

---

## GdiplusStartup

The **GdiplusStartup** function initializes Microsoft Windows GDI+. Call **GdiplusStartup** before making any other GDI+ calls, and call **GdiplusShutdown** when you have finished using GDI+.

```
FUNCTION GdiplusStartup (BYVAL token AS ULONG_PTR PTR, BYVAL input AS CONST GdiplusStartupInput PTR, _
   BYVAL output AS GdiplusStartupOutput PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *token* | [out] Pointer to a **ULONG_PTR** that receives a token. Pass the token to GdiplusShutdown when you have finished using GDI+. |
| *input* | [in] Pointer to a **GdiplusStartupInput** structure that contains the GDI+ version, a pointer to a debug callback function, a Boolean value that specifies whether to suppress the background thread, and a Boolean value that specifies whether to suppress external image codecs. |
| *output* | [out] Pointer to a **GdiplusStartupOutput** structure that receives a pointer to a notification hook function and a pointer to a notification unhook function. If the **SuppressBackgroundThread** data member of the input parameter is **FALSE**, then this parameter can be **NULL**. |

#### Return Value 

If the function succeeds, it returns **Ok**, which is an element of the **Status** enumeration.

If the function fails, it returns one of the other elements of the **Status** enumeration.

#### Remarks 

You must call **GdiplusStartup** before you create any GDI+ objects, and you must delete all of your GDI+ objects before you call GdiplusShutdown.

You can call **GdiplusStartup** on one thread and call **GdiplusShutdown** on another thread as long as you delete all of your GDI+ objects (or have them go out of scope) before you call **GdiplusShutdown**.

Do not call **GdiplusStartup** or **GdiplusShutdown** in **DllMain** or in any function that is called by **DllMain**. If you want to create a dynamic-link library (DLL) that uses GDI+, you should use one of the following techniques to initialize GDI+: 

* Require your clients to call **GdiplusStartup** before they call the functions in your DLL and to call **GdiplusShutdown** when they have finished using your DLL.
* Export your own startup function that calls **GdiplusStartup** and your own shutdown function that calls **GdiplusShutdown**. Require your clients to call your startup function before they call other functions in your DLL and to call your shutdown function when they have finished using your DLL.
* Call **GdiplusStartup** and **GdiplusShutdown** in each of your functions that make GDI+ calls.

**Warning**  For info about how to use dynamic data exchange (DDE) with GDI+, see [Special CWinApp Services](https://learn.microsoft.com/en-us/cpp/mfc/special-cwinapp-services?view=msvc-170).

The GdiplusShutdown function cleans up resources used by Microsoft Windows GDI+. Each call to GdiplusStartup should be paired with a call to GdiplusShutdown.

---

## GdiplusShutdown 

The **GdiplusShutdown** function cleans up resources used by Windows GDI+. Each call to **GdiplusStartup** should be paired with a call to **GdiplusShutdown**.

```
SUB GdiplusShutdown (BYVAL token AS ULONG_PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *token* | [in] Token returned by a previous call to **GdiplusStartup**. |

#### Remarks

You must call **GdiplusStartup** before you create any GDI+ objects, and you must delete all of your GDI+ objects before you call **GdiplusShutdown**.

---

## GdipAlloc

Allocates memory for GDI+ objects.

```
FUNCTION GdipAlloc (BYVAL size AS size_t) AS ANY PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *size* | [in] [in] Number of bytes to be allocated. |

#### Return Value

If the function succeeds, the return value is a pointer to the allocated memory block.

If the function fails, then return value is **NULL**.

---

## GdipFree

Frees memory allocated for GDI+ objects.

```
SUB GdipFree (BYVAL ptr AS ANY PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *ptr* | [in] Pointer to the memory block to free. |

---

## ObjectTypeIsValid

The **ObjectTypeIsValid** function determines whether an element of the ObjectType enumeration represents a valid object type.

```
FUNCTION ObjectTypeIsValid (BYVAL type_ as ObjectType) AS BOOLEAN
	FUNCTION = -((type_ >= ObjectTypeMin) AND (type_ <= ObjectTypeMax))
END FUNCTION
```

| Parameter  | Description |
| ---------- | ----------- |
| *type* | [in] Element of the **ObjectType** enumeration to be tested. |

#### Return value

If *objectType* is equal to *ObjectTypeInvalid*, this function returns FALSE.

If *objectType* is equal to any other element of the **ObjectType** enumeration, this function returns TRUE.

---

## AfxGdipInit

Initilizes GDI+

```
FUNCTION AfxGdipInit () AS ULONG_PTR
```

####  Remarks

You must call **AfxGdipInit** before you create any GDI+ objects, and you must delete all of your GDI+ objects (or have them go out of scope) before you call **AfxGdipShutdown**.

You can call **AfxGdipIni**" on one thread and call **AfxGdipShutdown"" on another thread as long as you delete all of your GDI+ objects (or have them go out of scope) before you call **AfxGdipShutdown**.

Do not call **AfxGdipIni** or **AfxGdipShutdown** in **DllMain** or in any function that is called by **DllMain**. If you want to create a DLL that uses GDI+, you should use one of the following techniques to initialize GDI+:

* Require your clients to call **AfxGdipInit** before they call the functions in your DLL and to call **AfxGdipShutdown** when they have finished using your DLL.
* Export your own startup function that calls **AfxGdipIni** and your own shutdown function that calls **AfxGdipShutdown**. Require your clients to call your startup function before they call other functions in your DLL and to call your shutdown function when they have finished using your DLL.
* Call **AfxGdipIni** and **AfxGdipShutdown** in each of your functions that make GDI+ calls.

---

## AfxGdipShutdown

Cleans up resources used by Windows GDI+. Each call to **AfxGdipInit** should be paired with a call to **AfxGdipShutdown**.

```
FUNCTION AfxGdipShutdown (BYVAL token AS ULONG_PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *token* | [in] Token returned by a previous call to **AfxGdipInit**. |

You must call **AfxGdipInit** before you create any GDI+ objects, and you must delete all of your GDI+ objects (or have them go out of scope) before you call **AfxGdipShutdown**.

---

## AfxGdipStatusStr

Returns the description of a GdiPlus status code.

```
PRIVATE FUNCTION AfxGdipStatusStr (BYVAL status AS GpStatus) AS DWSTRING

   DIM wszMsg AS WSTRING * 260 = "Unknown"

   SELECT CASE status
      CASE Ok                        : wszMsg = "Ok"
      CASE GenericError              : wszMsg = "Generic error"
      CASE InvalidParameter          : wszMsg = "Invalid parameter"
      CASE OutOfMemory               : wszMsg = "Out of memory"
      CASE ObjectBusy                : wszMsg = "Object busy"
      CASE InsufficientBuffer        : wszMsg = "Insufficient buffer"
      CASE NotImplemented            : wszMsg = "Not implemented"
      CASE Win32Error                : wszMsg = "Win 32 error"
      CASE WrongState                : wszMsg = "Wrong state"
      CASE Aborted                   : wszMsg = "Aborted"
      CASE FileNotFound              : wszMsg = "File not found"
      CASE ValueOverflow             : wszMsg = "Value overflow"
      CASE AccessDenied              : wszMsg = "Access denied"
      CASE UnknownImageFormat        : wszMsg = "Unknown image format"
      CASE FontFamilyNotFound        : wszMsg = "Font family not found"
      CASE FontStyleNotFound         : wszMsg = "Font style not found"
      CASE NotTrueTypeFont           : wszMsg = "Not TrueType font"
      CASE UnsupportedGdiplusVersion : wszMsg = "Unsupported GdiPlus version"
      CASE GdiplusNotInitialized     : wszMsg = "GdiPlus not initialized"
      CASE PropertyNotFound          : wszMsg = "Property not found"
      CASE PropertyNotSupported      : wszMsg = "Property not supported"
'   #IF (GDIPVER >= &H0110)
'      CASE ProfileNotFound           : wszMsg = "Profile not found"
'   #ENDIF   ' //(GDIPVER >= &H0110)
   END SELECT

   RETURN wszMsg

END FUNCTION
```
