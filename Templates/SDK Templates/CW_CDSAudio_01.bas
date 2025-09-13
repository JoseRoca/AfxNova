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
         pCDSAudio = NEW CDSAudio(ExePath & "/resources/Kalimba.mp3")
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
