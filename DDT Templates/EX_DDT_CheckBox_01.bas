'#TEMPLATE DDT Dialog with a checkbox
#define UNICODE
#define _WIN32_WINNT &h0602
#include once "AfxNova/DDT.inc"
USING AfxNova

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG
END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_CHECKBOX = 1001

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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog with a checkbox", 50, 50, 175, 65, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a check button to the dialog
   ControlAddCheckBox, hDlg, IDC_CHECKBOX, "Click me", 105, 40, 60, 12

   ' // Anchor the button to the bottom and the right side of the dialog
   ControlAnchor(hDlg, IDC_CHECKBOX, AFX_ANCHOR_BOTTOM_RIGHT)

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
            CASE IDC_CHECKBOX
               SELECT CASE ControlGetCheck(hDlg, IDC_CHECKBOX)
                  CASE BST_CHECKED
                     MsgBox hDlg, "Checked", MB_OK, "Message"
                  CASE BST_UNCHECKED
                     MsgBox hDlg, "Unchecked", MB_OK, "Message"
               END SELECT
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
