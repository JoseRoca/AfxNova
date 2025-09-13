'#TEMPLATE DDT Dialog with an UpDown Control
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

CONST IDC_UPDOWN = 1001
CONST IDC_TEXTBOX = 1002

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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog UpDown Control", 50, 50, 190, 85, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a label control
   ControlAddLabel, hDlg, -1, "How many lines?", 46, 29, 57, 12
   ' // Add an edit control
   ControlAddTextBox, hDlg, IDC_TEXTBOX, "", 114, 29, 29, 12
   ' // Add an UpDown control (the size and position will be automatically adjusted to the buddy control)
   ControlAddUpDown, hDlg, IDC_UPDOWN

   ' // Set the edit control as budy of the updown control
   DIM hUpDown AS HWND = ControlHandle(hDlg, IDC_UPDOWN)
   DIM hTextBox AS HWND = ControlHandle(hDlg, IDC_TEXTBOX)
   UpDown_SetBuddy(hUpDown, hTextBox)
   ' // Set the base
   UpDown_SetBase(hUpDown, 10)
   ' // Set the range
   UpDown_SetRange32(hUpDown, 0, 100)
   ' // Set the default initial value
   UpDown_SetPos32(hUpDown, 1)

   ' // Add a cancel button
   ControlAddButton, hDlg, IDCANCEL, "&Close", 120, 60, 50, 12

   ' // Anchor the controls
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

      CASE WM_NOTIFY
         ' // Process notification messages
         DIM tUD AS NMUPDOWN
         CBNMTYPESET(tUD, wParam, lParam)
         IF tUD.hdr.idFrom = IDC_UPDOWN THEN
            IF tUD.hdr.code = UDN_DELTAPOS THEN
               ' Put your code here, e.g.
               ' tUD.iPos = ? ' Up-down control's current position.
               ' tUD.iDelta = ? ' Proposed change in the up-down control's position.
            END IF
            ' Return TRUE to prevent the change in the control's position, or FALSE to allow the change.
            RETURN TRUE
         END IF

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
