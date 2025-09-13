'#TEMPLATE DDT Dialog with a StatusBar and a Progress Bar
#define UNICODE
#include once "AfxNova/DDT.inc"
USING AfxNova

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_START = 1001
CONST IDC_STATUSBAR = 1002
CONST IDC_PROGRESSBAR = 1003

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   ' // Create a new dialog using dialog units
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog - StatusBar", 50, 50, 190, 85, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a status bar
   ControlAddStatusBar, hDlg, IDC_STATUSBAR
   ' // Set the parts
   StatusBarSetParts(hDlg, IDC_STATUSBAR, "90, -1")

   ' // Add a progress bar
   ControlAddProgressBar, hDlg, IDC_PROGRESSBAR, "", 0, 74, 89, 11
   ' // Disable theme for the control when using dialog units;
   ' // otherwise it does not work properly.
   DIM hProgressBar AS HWND = ControlHandle(hDlg, IDC_PROGRESSBAR)
   SetWindowTheme(hProgressBar, "", "")
   ' // Set the range
   ProgressBarSetRange(hDlg, IDC_PROGRESSBAR, 0, 100)
   ' // Set the initial position
   ProgressBarSetPos(hDlg, IDC_PROGRESSBAR, 0)

   ' // Add a start button
   ControlAddButton, hDlg, IDC_START, "&Start", 10, 30, 50, 12

   ' // Anchor the controls
   ControlAnchor(hDlg, IDC_START, AFX_ANCHOR_NONE)
   ControlAnchor(hDlg, IDC_STATUSBAR, AFX_ANCHOR_BOTTOM_WIDTH)
   ControlAnchor(hDlg, IDC_PROGRESSBAR, AFX_ANCHOR_BOTTOM)

   ' // Display and activate the dialog as modal
   DialogShowModal(hDlg, @DlgProc)

   RETURN DialogEndResult(hDlg)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Dialog callback procedure
' ========================================================================================
FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR
   
   SELECT CASE uMsg

      CASE WM_INITDIALOG
         RETURN TRUE

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hDlg, WM_CLOSE, 0, 0
               END IF
            CASE IDC_START
               ' *** Test code ***
               ' // Sets the step increment
               ProgressBarSetStep(hDlg, IDC_PROGRESSBAR, 1)
               ' // Draws the bar
               FOR i AS LONG = 1 TO 100
                  ' // Advances the current position for a progress bar by the step
                  ' // increment and redraws the bar to reflect the new position
                  ProgressBarStep(hDlg, IDC_PROGRESSBAR)
                  SLEEP 10
               NEXT
               ' // Clears the bar by reseting its position to 0
               ProgressBarSetPos(hDlg, IDC_PROGRESSBAR, 0)
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
