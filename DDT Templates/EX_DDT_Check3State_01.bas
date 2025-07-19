'#TEMPLATE DDT Dialog with a 3 state checkbox
#define _WIN32_WINNT &h0602
#include once "AfxNova/DDT.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_CHECK3STATE = 1001

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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog with a check button", 50, 50, 175, 65, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a check button to the dialog
   ControlAddCheck3State, hDlg, IDC_CHECK3STATE, "Click me", 105, 40, 60, 12

   ' // Anchor the button to the bottom and the right side of the dialog
   ControlAnchor(hDlg, IDC_CHECK3STATE, AFX_ANCHOR_BOTTOM_RIGHT)

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
            CASE IDC_CHECK3STATE
               SELECT CASE ControlGetCheck(hDlg, IDC_CHECK3STATE)
                  CASE BST_CHECKED
                     MsgBox hDlg, "Checked", MB_OK, "Message"
                  CASE BST_UNCHECKED
                     MsgBox hDlg, "Unchecked", MB_OK, "Message"
                  CASE BST_INDETERMINATE
                     MsgBox hDlg, "Indeterminate", MB_OK, "Message"
               END SELECT
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
