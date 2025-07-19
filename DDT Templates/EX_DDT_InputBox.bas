'#TEMPLATE DDT Dialog with a button
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

CONST IDC_OK = 1001
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
   DIM hDlg AS HWND = DialogNew("MyClassName", 0, "Input Box dialog", 50, 50, 220, 80, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a label
   ControlAddLabel, hDlg, IDC_LABEL, "", 20, 15, 180, 12, WS_BORDER

   ' // Add a button to the dialog
   ControlAddButton, hDlg, IDC_OK, "&Click me", 105, 50, 50, 12, BS_DEFPUSHBUTTON

   ' // Anchor the controls
   ControlAnchor(hDlg, IDC_OK, AFX_ANCHOR_BOTTOM_RIGHT)
   ControlAnchor(hDlg, IDC_LABEL, AFX_ANCHOR_WIDTH)

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
                  RETURN TRUE
               END IF
            CASE IDC_OK
               DIM dwsDate AS DWSTRING = AfxInputBox(hDlg, 0, 0, "Input Box", "Which is your name?", "My name is Nobody")
               ' Don't use MsgBox here
               ControlSetText(hDlg, IDC_LABEL, dwsDate)
               RETURN TRUE
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
' ========================================================================================
