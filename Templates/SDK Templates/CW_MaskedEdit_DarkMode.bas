' ########################################################################################
' Microsoft Windows
' File: CW_MaskedEdit_DarkMode.bas
' Contents: CWindow - Masked Edit Control
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
#INCLUDE ONCE "AfxNova/CMaskedEdit.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG

   END wWinMain(GetModuleHandleW(NULL), NULL, WCOMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_MASKED = 1001

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest file
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   ' // Creates the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   pWindow.Create(NULL, "Masked edit control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(300, 150)
   ' // Centers the window
   pWindow.Center
   ' // Set the main window background color
   pWindow.SetBackColor(RGB_BLACK)

   ' // Add a masked edit control
   DIM pMaskedEdit AS CMaskedEdit = CMaskedEdit(@pWindow, IDC_MASKED, 50, 30, 200, 23)
   ' // Colors of the edit text
   pWindow.SetCtlColors(pMaskedEdit.hWindow, RGB_WHITE, RGB_BLACK)

   SetFocus pMaskedEdit.hWindow
   pMaskedEdit.EnableMask(" ddd  ddd dddd", "(___) ___-____", "_")
   pMaskedEdit.SetText("(123) 123-1212"), TRUE

   ' // Adds a button
   DIM hButton AS HWND = pWindow.AddControl("Button", , IDCANCEL, "&Close", 190, 100, 75, 23)
   ' // Set button dark mode
   SetWindowTheme(hButton, "DarkMode_Explorer", NULL)

   ' // Anchor controls
   pWindow.AnchorControl(IDC_MASKED, AFX_ANCHOR_WIDTH)
   pWindow.AnchorControl(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

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
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
            CASE IDC_MASKED
               IF CBCTLMSG(wPAram, lParam) = EN_CHANGE THEN
                  DIM hEdit AS HWND = GetDlgItem(hwnd, LOWORD(wParam))
                  DIM pMaskedEdit AS CMaskedEdit PTR = AfxCMaskedEditPtr(hEdit)
                  IF pMaskedEdit THEN
                     DIM dwsText AS DWSTRING
                     ' // Get the text including the mask
                     dwsText = pMaskedEdit->GetText(FALSE)
                     ' // Get the text excluding the mask
                     dwsText = pMaskedEdit->GetText(TRUE)
                  END IF
               END IF
         END SELECT


    	CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
