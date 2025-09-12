' ########################################################################################
' Microsoft Windows
' File: CW_MonthValendar_MonthDayState.bas
' Contents: CWindow - Month calendar - Month day state
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
' // MCN_GETDAYSTATE is wrongly defined in the FreeBasic headers
#undef MCN_GETDAYSTATE
#define MCN_GETDAYSTATE culng(-747)
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST IDC_MONTHCAL = 1001
#define BOLDDAY(ds, iDay) ds = ds OR (1 SHL (iDay - 1))

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
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Month calendar", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center

   ' // Adds the control
   pWindow.AddControl("SysMonthCal32", hWin, IDC_MONTHCAL, "", _
      10, 10, pWindow.ClientWidth - 20, pWindow.ClientHeight - 20, WS_VISIBLE OR WS_TABSTOP OR MCS_DAYSTATE)
   ' // Anchors the control
   pWindow.AnchorControl(IDC_MONTHCAL, AFX_ANCHOR_HEIGHT_WIDTH)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   ' Array to hold day states
   STATIC rgDayState(11) AS MONTHDAYSTATE

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

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
         END SELECT
         RETURN 0

     CASE WM_NOTIFY
         ' // Process notification messages
         DIM ptnmsc AS NMSELCHANGE PTR = CAST(NMSELCHANGE PTR, lParam)
         ' // Get selected date
         IF ptnmsc->nmhdr.code = MCN_SELCHANGE THEN
            DIM wszDate AS WSTRING * 260
            wszDate = "Day: " & WSTR(ptnmsc->stSelStart.wDay) & " " & _
                      "Month: " & WSTR(ptnmsc->stSelStart.wMonth) & " " & _
                      "Year: " & WSTR(ptnmsc->stSelStart.wYear)
            SendMessageW hwnd, WM_SETTEXT, 0, CAST(LPARAM, @wszDate)
            ' Note: Don't use MessageBox here or you will get non-stop messages!
            EXIT FUNCTION
         END IF

         ' // Set the day state
         DIM pnmhdr AS LPNMHDR = CAST(LPNMHDR, lParam)
         IF pnmhdr->code = MCN_GETDAYSTATE THEN
            DIM pDayState AS LPNMDAYSTATE = CAST(LPNMDAYSTATE, lParam)
            FOR i AS INTEGER = 0 TO pDayState->cDayState - 1
                rgDayState(i) = 0
                BOLDDAY(rgDayState(i), 1)
                BOLDDAY(rgDayState(i), 15)
            NEXT
            pDayState->prgDayState = @rgDayState(0)
            RETURN TRUE
         END IF


      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
