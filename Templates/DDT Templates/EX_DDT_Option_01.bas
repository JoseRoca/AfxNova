'#TEMPLATE DDT Dialog with option buttons
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

CONST IDC_GROUPBOX = 1001
CONST IDC_OPTION1 = 1002
CONST IDC_OPTION2 = 1003
CONST IDC_OPTION3 = 1004

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
   DIM hDlg AS HWND = DialogNewPixels(0, "DDT Dialog with option buttons", 0, 0, 275, 180, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a frame control
   ControlAddFrame, hDlg, IDC_GROUPBOX, "Options", 10, 10, 255, 110
   ' // Add three radio buttons (the first one should have the WS_GROUP style)
   ControlAddOption, hDlg, IDC_OPTION1, "Option 1", 60, 40, 75, 23, WS_GROUP
   ControlAddOption, hDlg, IDC_OPTION2, "Option 2", 60, 60, 75, 23
   ControlAddOption, hDlg, IDC_OPTION3, "Option 3", 60, 80, 75, 23
   ' // Check the second radio button
   ControlSetOption hDlg, IDC_OPTION2, IDC_OPTION1, IDC_OPTION3

   ' // Add a cancel button to the dialog
   ControlAddButton, hDlg, IDCANCEL, "&Cancel", 95, 140, 80, 23, BS_DEFPUSHBUTTON

   ' // Anchor the button to the bottom and the right side of the dialog
   ControlAnchor(hDlg, IDC_GROUPBOX, AFX_ANCHOR_WIDTH)
   ControlAnchor(hDlg, IDCANCEL, AFX_ANCHOR_CENTER_HORZ_BOTTOM)

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

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
