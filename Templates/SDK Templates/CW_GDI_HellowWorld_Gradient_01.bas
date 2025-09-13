' ########################################################################################
' Microsoft Windows
' File: CW_GDI_HelloWorld_Gradient_01.bas
' Contents: CWindow Hello Word with gradient example (High DPI)
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Hello World - Gradient", @WndProc)
   ' // Change the background to white
   pWindow.Brush = GetStockObject(WHITE_BRUSH)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center

   ' // add buttons
   pWindow.AddControl("Button", hWin, IDOK, "&Ok", pWindow.ClientWidth - 185, pWindow.ClientHeight - 35, 75, 23)
   pWindow.AddControl("Button", hWin, IDCANCEL, "&Quit", pWindow.ClientWidth - 95, pWindow.ClientHeight - 35, 75, 23)
   ' // Anchor buttons
   pWindow.AnchorControl(IDOK, AFX_ANCHOR_BOTTOM_RIGHT)
   pWindow.AnchorControl(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gradient procedure
' ========================================================================================
SUB DrawGradient (BYVAL hDC AS HDC)

   DIM rcFill   AS RECT
   DIM rcClient AS RECT
   DIM fStep    AS SINGLE
   DIM hBrush   AS HBRUSH
   DIM lOnBand  AS LONG

   GetClientRect WindowFromDC(hDC), @rcClient
   fStep = rcClient.Bottom / 200

   FOR lOnBand = 0 TO 199
      SetRect @rcFill, 0, lOnBand * fStep, rcClient.Right + 1, (lOnBand + 1) * fStep
      ' // Note: The C macro RGB has been renamed as BGR to avoid conflicts with the Free Basic RGB intrinsic function
      hBrush = CreateSolidBrush(BGR(0, 0, (255 - lOnBand)))
      FillRect hDC, @rcFill, hBrush
      DeleteObject hBrush
   NEXT

END SUB
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC hNewFont AS HFONT

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         AfxEnableDarkModeForWindow(hwnd)
         ' // Get a pointer to the CWindow class from the CREATESTRUCT structure
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(lParam)
         ' // Create a new font scaled according the DPI ratio
         hNewFont = pWindow->CreateFont("Times New Roman", 16)
         RETURN 0

      ' // Theme has changed
      CASE WM_THEMECHANGED
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
         END SELECT
         RETURN 0

      CASE WM_PAINT
         ' // Draw the text
         DIM ps AS PAINTSTRUCT, hOldFont AS HFONT
         DIM hDC AS HDC = BeginPaint(hwnd, @ps)
         IF hNewFont THEN hOldFont = CAST(HFONT, SelectObject(hDC, CAST(HGDIOBJ, hNewFont)))
         DIM rc AS RECT
         GetClientRect hwnd, @rc
         SetBkMode hDC, TRANSPARENT
         SetTextColor hDC, &HFFFFFF
         DrawTextW hDC, "Hello, World!", -1, @rc, DT_SINGLELINE OR DT_CENTER OR DT_VCENTER
         IF hNewFont THEN SelectObject(hDC, CAST(HGDIOBJ, CAST(HFONT, hOldFont)))
         EndPaint hwnd, @ps
         RETURN TRUE

      CASE WM_ERASEBKGND
         ' // Draw the gradient
         DIM hDC AS HDC = CAST(HDC, wParam)
         DrawGradient hDC
         FUNCTION = CTRUE
         RETURN 0

      CASE WM_DESTROY
         ' // Destroy the new font
         IF hNewFont THEN DeleteObject(CAST(HGDIOBJ, hNewFont))
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
