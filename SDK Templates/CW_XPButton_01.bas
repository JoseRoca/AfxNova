' ########################################################################################
' Microsoft Windows
' File: CW_XPButton_01.bas
' Contents: CWindow - XPButton
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/CWindow.inc"
#INCLUDE ONCE "AfxNova/CXpButton.inc"
USING AfxNova

CONST IDC_BUTTON1 = 1001
CONST IDC_BUTTON2 = 1002
CONST IDC_BUTTON3 = 1003

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwsszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG

   END wWinMain(GetModuleHandleW(NULL), NULL, WCOMMAND(), SW_NORMAL)

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

   ' // Create the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "XpButton example", @WndProc)
   pWindow.SetClientSize(215, 142)
   pWindow.Center

   ' // Create the first button
   DIM pXpButton1 AS CXpButton = CXpButton(@pWindow, IDC_BUTTON1, "&Hot", 50, 10, 114, 34)
   ' // Load a png icon from file (remember to change the path of the file)
   pXpButton1.SetImageFromFile(ExePath & "/resources/arrow_left_64.png", XPBI_NORMAL)
   ' // Load a Windows predefined icon as the "hot" image
   pXpButton1.SetIcon LoadIcon(NULL, IDI_QUESTION), XPBI_HOT
   ' // Change the cursor shape
   pXpButton1.SetCursor LoadCursor(NULL, IDC_CROSS)
   ' // Set a margin of 10 dpi units (pixels x scaling ratio)
   pXpButton1.SetImageMargin 10

   ' // Create the second button
   DIM pXpButton2 AS CXpButton = CXpButton(@pWindow, IDC_BUTTON2, "&Cancel", 50, 50, 114, 34)
   ' // Load a Windows predefined icon
   pXpButton2.SetIcon LoadIcon(NULL, IDI_ERROR), XPBI_NORMAL
   pXpButton2.SetImagePos XPBI_RIGHT OR XPBI_VCENTER
   pXpButton2.SetTextFormat DT_LEFT OR DT_VCENTER OR DT_SINGLELINE
'   EnableWindow pXpButton2.hWindow, FALSE   ' Disable the button

   ' // Create a coloured button
   DIM pXpButton3 AS CXpButton = CXpButton(@pWindow, IDC_BUTTON3, "&Coloured Button", 50, 90, 114, 34, WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER OR BS_FLAT)
   ' // Disable theming
   pXpButton3.DisableTheming
   ' // Set the background color to blue
   pXpButton3.SetButtonBkColor BGR(0, 0, 255)
   ' // Set the background color when pressed to yellow
   pXpButton3.SetButtonBkColorDown BGR(255, 255, 0)
   ' // Set the text foreground color to yellow
   pXpButton3.SetTextForeColor BGR(255, 255, 0)
   ' // Set the text background color to blue
   pXpButton3.SetTextBkColor BGR(0, 0, 255)
   ' // Set the text foreground color when pressed to white
   pXpButton3.SetTextForeColorDown BGR(255, 255, 255)
   ' // Set the text background color when pressed to red
   pXpButton3.SetTextBkColorDown BGR(255, 0, 0)
   ' // Set the text background color when pressed to red
   pXpButton3.SetTextBkColorDown BGR(255, 0, 0)
   ' // Set the text format (centered and single line)
   pXpButton3.TextFormat = DT_CENTER OR DT_VCENTER OR DT_SINGLELINE
' pXpButton3.SetToggle TRUE   ' make it a toggle button

   ' // Set the focus in the first button
   SetFocus pXpButton1.hWindow

   ' // Dispatch Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window callback procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDC_BUTTON1, IDC_BUTTON2, IDC_BUTTON3
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  ' AfxMsg "Button clicked"
                  EXIT FUNCTION
               END IF

            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================