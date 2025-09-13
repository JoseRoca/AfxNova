'#TEMPLATE DDT Dialog with a Line control
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

CONST IDC_LINE1 = 1001
CONST IDC_LINE2 = 1002
CONST IDC_LINE3 = 1003

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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog with Line controls", , , 195, 100, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add the Line controls
   ControlAddLine, hDlg, IDC_LINE1, "", 10, 10, 175, 1, WS_VISIBLE OR SS_ETCHEDFRAME OR SS_NOTIFY
   ControlAddLine, hDlg, IDC_LINE2, "", 10, 20, 175, 10, WS_VISIBLE OR SS_ETCHEDFRAME OR SS_NOTIFY
   ControlAddLine, hDlg, IDC_LINE3, "", 10, 40, 175, 25, WS_VISIBLE OR SS_ETCHEDFRAME OR SS_NOTIFY

   ' // Add a button
   ControlAddButton, hDlg, IDCANCEL, "&Close", 135, 75, 50, 12

   ' // Anchor the controls
   ControlAnchor(hDlg, IDC_LINE1, AFX_ANCHOR_WIDTH)
   ControlAnchor(hDlg, IDC_LINE2, AFX_ANCHOR_WIDTH)
   ControlAnchor(hDlg, IDC_LINE3, AFX_ANCHOR_WIDTH)
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
            CASE IDC_LINE1
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  MsgBox hDlg, "Line 1 clicked", MB_OK, "Line example"
               END IF
            CASE IDC_LINE2
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  MsgBox hDlg, "Line 2 clicked", MB_OK, "Line example"
               END IF
            CASE IDC_LINE3
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  MsgBox hDlg, "Line 3 clicked", MB_OK, "Line example"
               END IF
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
