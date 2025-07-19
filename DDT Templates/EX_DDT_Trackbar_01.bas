'#TEMPLATE DDT Dialog with a Date Time Picker
#define UNICODE
#define _WIN32_WINNT &h0602
#include once "AfxNova/DDT.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_TRACKBAR = 1001

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

   ' // Create a new dialog using dialog units
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog with a Trackbar", 50, 50, 190, 85, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a date time picker control to the dialog
   DIM hTrackBar AS HWND = ControlAdd("Trackbar", hDlg, IDC_TRACKBAR, "", 25, 30, 145, 12)
   ' // Set the range values
   Trackbar_SetRangeMin hTrackBar, 0, TRUE
   Trackbar_SetRangeMax hTrackBar, 20, TRUE
   ' // Set the page size
   Trackbar_SetPageSize hTrackBar, 1

   ' // Add a cancel button
   ControlAddButton, hDlg, IDCANCEL, "&Close", 120, 60, 50, 12

   ' // Anchor the controls
   ControlAnchor(hDlg, IDC_TRACKBAR, AFX_ANCHOR_WIDTH)
   ControlAnchor(hDlg, IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

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
         END SELECT

      CASE WM_HSCROLL
         IF GetWindowID(cast(HWND, lParam)) = IDC_TRACKBAR THEN
            ' Put your code here
            ' The LOWORD of wParam specifies a scroll bar value that indicates the user's scrolling request.
            ' The HIWORD specifies the current position of the scroll box if the LOWORD is SB_THUMBPOSITION
            ' or SB_THUMBTRACK; otherwise, this word is not used.
         END IF

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
