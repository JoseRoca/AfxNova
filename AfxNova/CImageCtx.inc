' ########################################################################################
' Platform: Microsoft Windows
' File: CImageCtx.inc
' Contents: GDI+ image control
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2025 Jos� Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the �Software�), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED �AS IS�, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "AfxNova/CWindow.inc"
#include once "AfxNova/AfxGdiPlus.inc"

NAMESPACE AfxNova

CONST GDIP_IMAGECTX_AUTOSIZE    = 0   ' // Autoadjusts the image to the width or height of the control
CONST GDIP_IMAGECTX_ACTUALSIZE  = 1   ' // Shows the image with its actual size
CONST GDIP_IMAGECTX_FITTOWIDTH  = 2   ' // Adjusts the image to the width of the control
CONST GDIP_IMAGECTX_FITTOHEIGHT = 3   ' // Adjusts the image to the height of the control
CONST GDIP_IMAGECTX_STRETCH     = 4   ' // Adjusts the image to the height and width of the control

' // Notification messages
'NM_CLICK     = NM_FIRST - 2
'NM_DBLCLK    = NM_FIRST - 3
'NM_RCLICK    = NM_FIRST - 5
'NM_RDBLCLK   = NM_FIRST - 6
'NM_SETFOCUS  = NM_FIRST - 7
'NM_KILLFOCUS = NM_FIRST - 8

' // Process them in the main window callback as follows:
'
' CASE WM_NOTIFY
'    DIM phdr AS NMHDR PTR
'    phdr = CAST(NMHDR PTR, lParam)
'    IF wParam = IDC_GRCTX THEN
'       SELECT CASE phdr->code
'          CASE NM_CLICK
'             ' Left button clicked
'          CASE NM_RCLICK
'             ' Right button clicked
'          CASE NM_SETFOCUS
'             ' The control has gained focus
'          CASE NM_KILLFOCUS
'             ' The control has lost focus
'       END SELECT
'    END IF

' ========================================================================================
' CImageCtx class
' ========================================================================================
TYPE CImageCtx

   Private:
      m_hCtl        As HWnd             ' // Handle to the control
      m_hFont       As HFONT            ' // Handle to the font used by the control
      m_pImage      As GpImage Ptr      ' // Image object pointer
      m_token       As ULONG_PTR        ' // Token returned by GdiplusStartup
      m_BkColor     As COLORREF         ' // RGB background color
      m_BkColorHot  As COLORREF         ' // RGB background hot color
      m_ImageWidth  AS LONG             ' // Image width
      m_ImageHeight AS LONG             ' // Image height
      m_Fit         As Long             ' // Fit to width/height
      m_InterpolationMode As Long       ' // Interpolation mode

   Public:

      DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszTitle AS WSTRING = "", _
         BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
         BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION hWindow () AS HWND
      DECLARE FUNCTION GetImagePtr () AS GpImage PTR
      DECLARE FUNCTION DrawImage (BYVAL hwnd AS HWND, BYVAL hDC AS HDC, BYVAL pImage AS GpImage PTR) AS LONG
      DECLARE FUNCTION LoadImageFromFile (BYREF wszFileName AS WSTRING) AS LONG
      DECLARE FUNCTION LoadBitmapFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
      DECLARE FUNCTION LoadImageFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
      DECLARE FUNCTION SetImageAdjustment (BYVAL ImageAdjustment AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
      DECLARE FUNCTION GetImageAdjustment () AS LONG
      DECLARE FUNCTION GetImageWidth () AS LONG
      DECLARE FUNCTION GetImageHeight () AS LONG
      DECLARE FUNCTION SetBkColor (BYVAL clr AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
      DECLARE FUNCTION GetBkColor () AS LONG
      Declare Function SetBkColorHot (ByVal clr As Long, ByVal fRedraw As BOOLEAN = False) As Long
      Declare Function GetBkColorHot () As Long
      DECLARE SUB SetImageSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetImageWidth (BYVAL nWidth AS LONG, BYVAL fRedraw AS LONG = FALSE)
      DECLARE SUB SetImageHeight (BYVAL nHeight AS LONG, BYVAL fRedraw AS LONG = FALSE)
      Declare Sub SetInterpolationMode (ByVal InterpolationMode As Long)
      DECLARE FUNCTION GetInterpolationMode () AS LONG
      DECLARE SUB Redraw
      DECLARE SUB Clear
      DECLARE STATIC FUNCTION CImageCtxProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CImageCtx class constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CImageCtx (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, BYREF wszTitle AS WSTRING = "", _
   BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, _
   BYVAL dwStyle AS DWORD = 0, BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)

   ' // Register the class
   DIM wAtom AS ATOM
   DIM wcexw AS WNDCLASSEXW
   DIM wszClassName AS WSTRING * 260 = "AFX_GDIPIMAGECTX"
   IF .GetClassInfoExW(.GetModuleHandleW(NULL), @wszClassName, @wcexw) = 0 THEN
      ' // Fill the WNDCLASSEXW structure
      WITH wcexw
         .cbSize        = SIZEOF(wcexw)
         .style         = CS_DBLCLKS OR CS_HREDRAW OR CS_VREDRAW
         .lpfnWndProc   = @CImageCtxProc
         .cbClsExtra    = 0
         .cbWndExtra    = SIZEOF(HANDLE)   ' // make room to store a pointer to the class
         .hInstance     = ..GetModuleHandleW(NULL)
         .hCursor       = ..LoadCursorW(NULL, CAST(LPCWSTR, IDC_ARROW))
         .hbrBackground = CAST(HBRUSH, WHITE_BRUSH)
         .lpszMenuName  = NULL
         .lpszClassName = @wszClassName
         .hIcon         = NULL
         .hIconSm       = NULL
      END WITH
      wAtom = .RegisterClassExW(@wcexw)
   END IF

   ' // Set default values for background normal and hot colors
   m_BkColor    = GetSysColor(COLOR_BTNFACE)
   m_BkColorHot = m_BkColor
   
   ' // Create the control
   IF dwStyle = 0 THEN dwStyle = WS_VISIBLE OR WS_CHILD OR WS_TABSTOP
   IF pWindow THEN m_hCtl = pWindow->AddControl(wszClassName, pWindow->hWindow, cID, wszTitle, x, y, nWidth, nHeight, dwStyle, dwExStyle, lpParam)
   IF m_hCtl THEN
      .SetWindowLongPtrW m_hCtl, 0, CAST(LONG_PTR, @this)
      ' // Set the same font used by the parent
      DIM lfw AS LOGFONTW
      IF pWindow->Font THEN
         IF .GetObjectW(pWindow->Font, SIZEOF(lfw), @lfw) THEN m_hFont = CreateFontIndirectW(@lfw)
      END IF
   END IF

   ' // Initialize GDI+
   m_token = AfxGdipInit
   ' // Default interpolation mode
   m_InterpolationMode = InterpolationModeHighQualityBicubic

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CImageCtx class destructor
' ========================================================================================
PRIVATE DESTRUCTOR CImageCtx
   ' // Delete the font used by the control
   IF m_hFont THEN DeleteObject m_hFont
   ' // Destroy the Image object
   IF m_pImage THEN GdipDisposeImage(m_pImage)
   ' // Shutdown GDI+
   IF m_token THEN GdiplusShutdown m_token
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Redraws the control.
' ========================================================================================
PRIVATE SUB CImageCtx.Redraw()
   .InvalidateRect m_hCtl, NULL, 0
   .UpdateWindow m_hCtl
END SUB
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
PRIVATE FUNCTION CImageCtx.CImageCtxProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM hDC AS HDC, ps AS PAINTSTRUCT, hdr AS NMHDR, pImageCtx AS CImageCtx PTR

   SELECT CASE uMsg

      CASE WM_CREATE
         EXIT FUNCTION

      CASE WM_COMMAND, WM_NOTIFY
         ' // Forwards the message to the parent window
         .SendMessageW .GetParent(hwnd), uMsg, wParam, lParam

      CASE WM_GETDLGCODE
        ' // Ensures that the control will process all the keystrokes by itself
        FUNCTION = DLGC_WANTALLKEYS
        EXIT FUNCTION

      CASE WM_ERASEBKGND
         ' // Don't erase the background to avoid flicker
         FUNCTION = 1
         EXIT FUNCTION

      CASE WM_PRINTCLIENT
         ' // Get a pointer to the class
         pImageCtx = CAST(CImageCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pImageCtx = NULL THEN EXIT FUNCTION
         ' // Draws the image in the provided device context
         hDC = CAST(.HDC, wParam)
         pImageCtx->DrawImage(hwnd, hDC, pImageCtx->m_pImage)
         EXIT FUNCTION

      CASE WM_PAINT
         ' // Get a pointer to the class
         pImageCtx = CAST(CImageCtx PTR, .GetWindowLongPtrW(hwnd, 0))
         IF pImageCtx = NULL THEN EXIT FUNCTION
         ' // Prepares the window for painting
         hDC = .BeginPaint(hwnd, @ps)
         ' // Draws the image
         pImageCtx->DrawImage(hwnd, hDC, pImageCtx->m_pImage)
         ' // Ends the painting of the window
         .EndPaint hwnd, @ps
         EXIT FUNCTION

      CASE WM_KEYDOWN
         SELECT CASE (wParam)
            ' // Must handle tab key by ourselves, since we have taken full control
            CASE VK_TAB
               IF HIWORD(.GetKeyState(VK_SHIFT)) = 0 THEN
                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, 0)
               ELSE
                  .SetFocus .GetNextDlgTabItem(GetParent(hwnd), hwnd, -1)
               END IF
         END SELECT

      CASE WM_LBUTTONDOWN
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = GetDlgCtrlId(hwnd)
         hdr.code = NM_CLICK
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_RBUTTONDOWN
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_RCLICK
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_LBUTTONDBLCLK
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_DBLCLK
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_RBUTTONDBLCLK
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_RDBLCLK
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_SETFOCUS
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_SETFOCUS
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_KILLFOCUS
         ' // Forwards the message to the parent window
         hdr.hwndFrom = hwnd
         hdr.idFrom = .GetDlgCtrlId(hwnd)
         hdr.code = NM_KILLFOCUS
         .SendMessageW .GetParent(hwnd), WM_NOTIFY, .GetDlgCtrlId(hwnd), CAST(.LPARAM, @hdr)
         EXIT FUNCTION

      CASE WM_ENABLE
         ' // Redraws the control
         .InvalidateRect hwnd, NULL, 0
         .UpdateWindow hwnd
         EXIT FUNCTION

      CASE WM_MOUSEMOVE
         ' // Tracks the mouse movement and stores the hot state
         Dim trackMouse As TrackMouseEvent
         If .GetPropW(HWnd, "HOT") = 0 Then
            trackMouse.cbSize = Sizeof(trackMouse)
            trackMouse.dwFlags = TME_LEAVE
            trackMouse.hwndTrack = HWnd
            trackMouse.dwHoverTime = 1
            .TrackMouseEvent(@trackMouse)
            .SetPropW HWnd, "HOT", Cast(HANDLE, CTRUE)
            .InvalidateRect hwnd, NULL, 0
            .UpdateWindow hwnd
         End If
         Exit Function

      Case WM_MOUSELEAVE
         ' // Removes the hot state and redraws the button
         .RemovePropW HWnd, "HOT"
         .InvalidateRect hwnd, NULL, 0
         .UpdateWindow hwnd
         Exit Function

   End Select

   ' // Default processing for other messages.
   FUNCTION = .DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the handle of the control
' ========================================================================================
PRIVATE FUNCTION CImageCtx.hWindow () AS HWND
   FUNCTION = m_hCtl
END FUNCTION
' ========================================================================================

' ========================================================================================
' Draws the image
' Returns StatusOk or an error code.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.DrawImage (BYVAL hwnd AS HWND, BYVAL _hDC AS HDC, BYVAL pImage AS GpImage PTR) AS LONG

   DIM pImageCtx AS CImageCtx PTR    ' // Pointer to the CImageCtx class
   DIM hStatus AS GpStatus           ' // Result code
   DIM pGraphics AS GpGraphics PTR   ' // Graphics object pointer
   DIM hBgBrush AS HBRUSH            ' // Background brush
   DIM ImageWidth AS LONG            ' // Image width
   DIM ImageHeight AS LONG           ' // Image heigth
   DIM p AS LONG                     ' // Percent
   DIM w AS LONG                     ' // Width
   DIM h AS LONG                     ' // Height
   DIM cx AS LONG                    ' // x-coordinate for center fill
   DIM cy AS LONG                    ' // y-coordinate for center fill
   DIM nFit AS LONG                  ' // Fit
   DIM rc AS RECT                    ' // client area
   Dim memDC as HDC                 ' // Double buffering
   Dim hbit As HBITMAP               ' // Double buffering

   ' // Check parameters validity
   IF hWnd = NULL OR _hDC = NULL THEN RETURN E_POINTER
   ' // Get a pointer to the class
   pImageCtx = CAST(CImageCtx PTR, GetWindowLongPtrW(hwnd, 0))
   IF pImageCtx = NULL THEN EXIT FUNCTION
   
   ' // Retrieves the coordinates of the window's client area
   GetClientRect hwnd, @rc
   memDC = CreateCompatibleDC( _hDC )
   hbit  = CreateCompatibleBitmap( _hDC, rc.right, rc.bottom )
   If hbit Then hbit = SelectObject( memDC, hbit )

   ' // Create the background brush based on normal or hot.
   ' // Need to brush the entire background in order to accommodate any
   ' // images that may have transparent backgrounds.
   If GetPropW(m_hCtl, "HOT") Then 
      hBgBrush = CreateSolidBrush(pImageCtx->m_BkColorHot)
   Else
      hBgBrush = CreateSolidBrush(pImageCtx->m_BkColor)
   End If
   FillRect memDC, @rc, hBgBrush
   DeleteObject hBgBrush
   
   ' // Retrieve the width and height of the image
   ImageWidth  = m_ImageWidth
   ImageHeight = m_ImageHeight
   
   ' // Creates a Graphics object
   hStatus = .GdipCreateFromHDC(memDC, @pGraphics)
   
   ' // Draws the image
   IF ( pGraphics ) andalso ( pImage )THEN
      nFit = pImageCtx->m_Fit
      IF nFit = GDIP_IMAGECTX_AUTOSIZE THEN
         IF (ImageWidth \ (rc.Right - rc.Left)) > (ImageHeight \ (rc.Bottom - rc.Top)) THEN
            nFit = GDIP_IMAGECTX_FITTOWIDTH
         ELSE
            nFit = GDIP_IMAGECTX_FITTOHEIGHT
         END IF
      END IF
      IF nFit = GDIP_IMAGECTX_ACTUALSIZE THEN
         ' // If the width of the image is smaller than the width of the control
         ' // calculate the position for a centered image
         w = ImageWidth: h = ImageHeight
         IF w < (rc.Right - rc.Left) THEN
            cx = ((rc.Right - rc.Left) - w) / 2
         END IF
         IF h < (rc.Bottom - rc.Top) THEN
            cy = ((rc.Bottom - rc.Top) - h) / 2
         END IF
         ' // Draw the image
         hStatus = GdipDrawImageRectI(pGraphics, pImageCtx->m_pImage, cx, cy, ImageWidth, ImageHeight)
      ELSEIF nFit = GDIP_IMAGECTX_STRETCH THEN
        ' Stretch or shrunk to fill the controls client area
         w = rc.Right - rc.Left
         h = rc.Bottom - rc.Top
         IF m_InterpolationMode THEN GdipSetInterpolationMode(pGraphics, m_InterpolationMode)
         hStatus = GdipDrawImageRectI(pGraphics, pImageCtx->m_pImage, 0, 0, w, h)
      ELSEIF nFit = GDIP_IMAGECTX_FITTOWIDTH THEN
         ' // Calculate the percent of reduction or ampliation
         ' // and adjust the image height
         p = ((rc.Right - rc.Left) * 100) / ImageWidth
         w = rc.Right - rc.Left
         h = ImageHeight * p / 100
         ' // If the height of the image is smaller than the height of the control
         ' // calculate the position for a centered image
         IF h < (rc.Bottom - rc.Top) THEN
            cy = ((rc.Bottom - rc.Top) - h) / 2
         END IF
         IF m_InterpolationMode THEN GdipSetInterpolationMode(pGraphics, m_InterpolationMode)
         hStatus = GdipDrawImageRectI(pGraphics, pImageCtx->m_pImage, 0, cy, w, h)
      ELSE
         ' // Calculate the percent of reduction or ampliation
         ' // and adjust the image width
         IF ImageHeight <> rc.Bottom - rc.Top THEN
            p = ((rc.Bottom - rc.Top) * 100) /  ImageHeight
            h = rc.Bottom - rc.Top
            w = ImageWidth * p / 100
         ELSE
            h = ImageHeight
            w = ImageWidth
         END IF
         ' // If the width of the image is smaller than the width of the control
         ' // calculate the position for a centered image
         IF w < (rc.Right - rc.Left) THEN
            cx = ((rc.Right - rc.Left) - w) / 2
         END IF
         ' // Draws the image
         IF m_InterpolationMode THEN GdipSetInterpolationMode(pGraphics, m_InterpolationMode)
         hStatus = GdipDrawImageRectI(pGraphics, pImageCtx->m_pImage, cx, 0, w, h)
      END IF
      ' // Deletes the Graphics object
      GdipDeleteGraphics(pGraphics)
   End If

   ' Copy the entire memory bitmap over to the visual display
   GetClientRect hwnd, @rc
   BitBlt _hDC, 0, 0, rc.right, rc.bottom, memDC, 0, 0, SRCCOPY 

   ' Cleanup
   If hbit  Then DeleteObject SelectObject(memDC, hbit)
   If memDC Then DeleteDC memDC
End Function
' ========================================================================================

' ========================================================================================
' Loads an image from file using GDI+
' Parameters:
' - wszFileName = Fully qualified path and filename of the image file to load.
' Return value:
' - Ok (0) or an error code.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.LoadImageFromFile (BYREF wszFileName AS WSTRING) AS LONG
   IF LEN(wszFileName) = 0 THEN RETURN E_INVALIDARG
   ' // Dispose a previous instance of the Image object, if any
   IF m_pImage THEN
      .GdipDisposeImage(m_pImage)
      m_pImage = NULL
   END IF
   ' // Loads the image
   .GdipLoadImageFromFile(wszFileName, @m_pImage)
   .GdipGetImageWidth(m_pImage, @m_ImageWidth)
   .GdipGetImageHeight(m_pImage, @m_ImageHeight)
   this.Redraw
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads a bitmap from a resource file using GDI+
' Parameters:
' - hInstance = Instance handle
' - wszResourceName = Name of the resource.
' Return value:
'   Ok or an error code.
' Note: Because GdipCreateBitmapFromResource uses LoadImageW to load the resource, it only
' works with icons or bitmaps. It works with .bmp and .png files, but fails with .jpg and .tif files.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.LoadBitmapFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
   IF LEN(wszResourceName) = 0 THEN RETURN E_INVALIDARG
   ' // Dispose a previous instance of the Image object, if any
   IF m_pImage THEN
      .GdipDisposeImage(m_pImage)
      m_pImage = NULL
   END IF
   DIM hStatus AS LONG = .GdipCreateBitmapFromResource(hInstance, wszResourceName, cast(GpBitmap PTR PTR, @m_pImage))
   .GdipGetImageWidth(m_pImage, @m_ImageWidth)
   .GdipGetImageHeight(m_pImage, @m_ImageHeight)
   this.Redraw
   FUNCTION = hStatus
END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads an image from a resource file using GDI+
' Parameters:
' - hInstance = Instance handle
' - wszResourceName = Name of the resource image.
' Return value:
'   Ok or an error code.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.LoadImageFromResource (BYVAL hInstance AS HINSTANCE, BYREF wszResourceName AS WSTRING) AS LONG
   IF LEN(wszResourceName) = 0 THEN RETURN E_INVALIDARG
   ' // Dispose a previous instance of the Image object, if any
   IF m_pImage THEN
      .GdipDisposeImage(m_pImage)
      m_pImage = NULL
   END IF
   DIM hStatus AS LONG, hDC AS HDC
   ' // Find the resource and lock it
   DIM hResource AS HRSRC = .FindResourceW(hInstance, wszResourceName, CAST(LPCWSTR, RT_RCDATA))
   IF hResource = NULL THEN RETURN E_INVALIDARG
   DIM imageSize AS DWORD = .SizeofResource(hInstance, hResource)
   IF imageSize = 0 THEN RETURN E_INVALIDARG
   DIM pResourceData AS LPVOID = .LockResource(LoadResource(hInstance, hResource))
   IF pResourceData = NULL THEN RETURN E_INVALIDARG
   ' // Allocate memory to hold the image
   DIM hGlobal AS HGLOBAL = .GlobalAlloc(GMEM_MOVEABLE, imageSize)
   IF hGlobal THEN
      ' // Lock the memory
      DIM pGlobalBuffer AS LPVOID = .GlobalLock(hGlobal)
      IF pGlobalBuffer THEN
         ' // Copy the image from the resource file to global memory
         .memcpy pGlobalBuffer, pResourceData, imageSize
         ' // Create an stream in global memory
         DIM pImageStream AS LPSTREAM, pGraphics AS GpGraphics PTR
         IF .CreateStreamOnHGlobal(hGlobal, FALSE, @pImageStream) = S_OK THEN
            ' // Create a bitmap from the data contained in the stream
            hStatus = .GdipCreateBitmapFromStream(pImageStream, cast(GpBitmap PTR PTR, @m_pImage))
            IF hStatus = 0 THEN
               ' // Gets the device context handle
               hDC = .GetDc(m_hCtl)
               ' // Creates a graphics object from it
               IF hDC THEN hStatus = .GdipCreateFromHDC(hDC, @pGraphics)
               ' // Draws the image (required to keep it in memory, since we are
               ' // going to unlock and free the resource)
               IF pGraphics THEN hStatus = .GdipDrawImageI(pGraphics, m_pImage, 0, 0)
               .GdipGetImageWidth(m_pImage, @m_ImageWidth)
               .GdipGetImageHeight(m_pImage, @m_ImageHeight)
               ' // Deletes the graphics object
               IF pGraphics THEN .GdipDeleteGraphics(pGraphics)
               ' // Releases the device context handle
               IF hDC THEN .DeleteDc hDC
            END IF
            pImageStream->lpVtbl->Release(pImageStream)
         END IF
         ' // Unlock the memory
         .GlobalUnlock pGlobalBuffer
      END IF
      ' // Free the memory
      .GlobalFree hGlobal
   END IF
   this.Redraw
   FUNCTION = hStatus
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the image adjustment.
' Parameters:
' - ImageAdjustment =
'     GDIP_IMAGECTX_AUTOSIZE = 0
'        Autoadjusts the image to the width or height of the control
'     GDIP_IMAGECTX_ACTUALSIZE = 1
'        Shows the image with its actual size
'     GDIP_IMAGECTX_FITTOWIDTH = 2
'        Adjusts the image to the width of the control
'     GDIP_IMAGECTX_FITTOHEIGHT = 3
'        Adjusts the image to the height of the control
' - fRedraw = [optional] TRUE or FALSE (redraws the control to reflect the changes)
' Return value:
' * Returns the previous setting.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.SetImageAdjustment (BYVAL ImageAdjustment AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
   ' // Returns the old setting
   FUNCTION = m_Fit
   ' // Store the new setting
   m_Fit = ImageAdjustment
   ' // Redraws the control
   IF fRedraw THEN this.Redraw
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the image adjustment setting.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetImageAdjustment () AS LONG
   FUNCTION = m_Fit
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the width of the image.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetImageWidth () AS LONG
   function = m_ImageWidth
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the height of the image.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetImageHeight () AS LONG
   function = m_ImageHeight
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color of the control.
' Parameters:
' * clr = RGB color.
' * fRedraw = [optional] TRUE or FALSE (redraws the control to reflect the changes)
' Return value:
' * Returns the previous background color.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.SetBkColor (BYVAL clr AS LONG, BYVAL fRedraw AS BOOLEAN = FALSE) AS LONG
   ' // Returns the old setting
   FUNCTION = m_BkColor
   ' // Stores the new color
   m_BkColor = clr
   ' // Redraws the control
   IF fRedraw THEN this.Redraw
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the background color of the control's window.
' ========================================================================================
Private Function CImageCtx.GetBkColor () As Long
   Function = m_BkColor
End Function
' ========================================================================================

' ========================================================================================
' Sets the background hot color of the control (when mouse is over the control).
' Parameters:
' * clr = RGB color.
' * fRedraw = [optional] TRUE or FALSE (redraws the control to reflect the changes)
' Return value:
' * Returns the previous background color.
' ========================================================================================
Private Function CImageCtx.SetBkColorHot (ByVal clr As Long, ByVal fRedraw As BOOLEAN = False) As Long
   ' // Returns the old setting
   Function = m_BkColorHot
   ' // Stores the new color
   m_BkColorHot = clr
   ' // Redraws the control
   If fRedraw Then this.Redraw
End Function
' ========================================================================================

' ========================================================================================
' Gets the background hot color of the control's window.
' ========================================================================================
Private Function CImageCtx.GetBkColorHot () As Long
   Function = m_BkColorHot
End Function
' ========================================================================================

' ========================================================================================
' Sets the interpolation mode used by GDI+. The interpolation mode determines the algorithm
' that is used when images are scaled or rotated.
' ========================================================================================
PRIVATE SUB CImageCtx.SetInterpolationMode (BYVAL InterpolationMode AS LONG)
   m_InterpolationMode = InterpolationMode
END SUB
' ========================================================================================

' ========================================================================================
' Gets the interpolation mode used by GDI+.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetInterpolationMode () AS LONG
   FUNCTION = m_InterpolationMode
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Image object pointer.
' Return value: The Image object pointer.
' NOTE: Don't dispose the returned pointer. The control keeps a copy of it and calls
' GdipDisposeImage when it is destroyed.
' ========================================================================================
PRIVATE FUNCTION CImageCtx.GetImagePtr () AS GpImage PTR
   FUNCTION = m_pImage
END FUNCTION
' ========================================================================================

' ========================================================================================
' Clears the contents of the control
' Return value:
' * If the function succeeds, the return value is TRUE.
'   If the function fails, the return value is FALSE.
' ========================================================================================
PRIVATE SUB CImageCtx.Clear
   IF m_pImage THEN
      .GdipDisposeImage(m_pImage)
      m_pImage = 0
   END IF
   this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Returns a pointer to the class given the handle of its associated window handle.
' hCtl = Handle of the image control.
' ========================================================================================
PRIVATE FUNCTION AfxCImageCtxPtr OVERLOAD (BYVAL hCtl AS HWND) AS CImageCtx PTR
   FUNCTION = CAST(CImageCtx PTR, .GetWindowLongPtrW(hCtl, 0))
END FUNCTION
' ========================================================================================
' ========================================================================================
' hParent = Handle of the parent window of the image control.
' cID = Identifier of the image control.
' ========================================================================================
PRIVATE FUNCTION AfxCImageCtxPtr OVERLOAD (BYVAL hParent AS HWND, BYVAL cID AS LONG) AS CImageCtx PTR
   FUNCTION = CAST(CImageCtx PTR, .GetWindowLongPtrW(GetDlgItem(hParent, cID), 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the size of the image
' Parameters:
' - nWidth = Width of the image, in pixels.
' - nHeight = Height of the image, in pixels.
' - fRedraw = TRUE or FALSE (redraws the image to reflect the changes)
' ========================================================================================
PRIVATE SUB CImageCtx.SetImageSize (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL fRedraw AS LONG = FALSE)
   m_ImageWidth = nWidth
   m_ImageHeight = nHeight
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Sets the width of the image
' Parameters:
' - nWidth = Width of the image, in pixels.
' - fRedraw = TRUE or FALSE (redraws the image to reflect the changes)
' ========================================================================================
PRIVATE SUB CImageCtx.SetImageWidth (BYVAL nWidth AS LONG, BYVAL fRedraw AS LONG = FALSE)
   m_ImageWidth = nWidth
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

' ========================================================================================
' Sets the height of the image
' Parameters:
' - nHeight = Height of the image, in pixels.
' - fRedraw = TRUE or FALSE (redraws the image to reflect the changes)
' ========================================================================================
PRIVATE SUB CImageCtx.SetImageHeight (BYVAL nHeight AS LONG, BYVAL fRedraw AS LONG = FALSE)
   m_ImageHeight = nHeight
   IF fRedraw THEN this.Redraw
END SUB
' ========================================================================================

END NAMESPACE
