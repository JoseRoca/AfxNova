'#TEMPLATE DDT Dialog with an IPAddress control
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

CONST IDC_IPADDRESS = 1001
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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog - IPAddress", 50, 50, 220, 80, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a label
   ControlAddLabel, hDlg, IDC_LABEL, "", 20, 15, 180, 12, WS_BORDER

   ' // Add an IPAddress control
   ControlAddIPAddress, hDlg, IDC_IPADDRESS, "", 75, 40, 70, 12

   ' // Anchor the controls
   ControlAnchor(hDlg, IDC_IPADDRESS, AFX_ANCHOR_WIDTH)
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
               END IF
            CASE IDC_IPADDRESS
               IF CBCTLMSG(wParam, lParam) = EN_CHANGE THEN
                  ' Smple code: Get the address in string form
                  DIM wszText AS WSTRING * 260 = ControlGetText(hDlg, IDC_IPADDRESS)
                  ControlSetText(hDlg, IDC_LABEL, wszText)
               END IF
         END SELECT
      CASE WM_NOTIFY
         ' // Notification messages
         DIM ptnmhdr AS NMHDR PTR = CAST(NMHDR PTR, lParam)
         SELECT CASE ptnmhdr->idFrom
            CASE IDC_IPADDRESS
               IF ptnmhdr->code = IPN_FIELDCHANGED THEN
                  ' *** Put your code here ***
               END IF
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
