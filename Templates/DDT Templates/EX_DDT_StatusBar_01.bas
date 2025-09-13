'#TEMPLATE DDT Dialog with a StatusBar
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

CONST IDC_STATUSBAR = 1001
CONST IDC_LABEL = 1002

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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog - StatusBar", 50, 50, 190, 85, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a label control
   ControlAddLabel, hDlg, IDC_LABEL, "Click one of the sections of the status bar", 20, 29, 157, 30

   ' // Add a status bar
   ControlAddStatusBar, hDlg, IDC_STATUSBAR
   ' // Set the parts
   StatusBarSetParts(hDlg, IDC_STATUSBAR, "50, 50, 50, -1")

   ControlAnchor(hDlg, IDC_LABEL, AFX_ANCHOR_WIDTH)
   ControlAnchor(hDlg, IDC_STATUSBAR, AFX_ANCHOR_BOTTOM_WIDTH)

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

      CASE WM_NOTIFY
         ' // Detect if the user has clicked the mouse in one of the status bar parts
         DIM tMouse AS NMMOUSE
         CBNMTYPESET(tMouse, wParam, lParam)
         IF tMouse.hdr.idFrom = IDC_STATUSBAR THEN
            IF tMouse.hdr.code = NM_CLICK THEN
               MsgBox hDlg, "You have clicked part " & WSTR(tMouse.dwItemSpec + 1), MB_OK, "Message"
            END IF
         END IF

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
