'#TEMPLATE DDT Dialog with a ComboBox
#define UNICODE
#include once "AfxNova/DDT.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_COMBOBOX = 1001

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
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog with a ComboBox", 50, 50, 190, 80, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a check button to the dialog
   ControlAddComboBox, hDlg, IDC_COMBOBOX, "", 50, 20, 85, 12

   ' // Fill the comboxbox with some data
   DIM wszText AS WSTRING * 260
   FOR i AS LONG = 1 TO 10
      wszText = "Item " & RIGHT("00" & STR(i), 2)
      ComboBoxAdd(hDlg, IDC_COMBOBOX, @wszText)
   NEXT

   ' // Select the second item
   ComboboxSelect hDlg, IDC_COMBOBOX, 2

   ' // Anchor the button to the bottom and the right side of the dialog
   ControlAnchor(hDlg, IDC_COMBOBOX, AFX_ANCHOR_WIDTH)

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
            CASE IDC_COMBOBOX
               IF CBCTLMSG(wParam, lParam) = LBN_SELCHANGE THEN
                  DIM item AS LONG = ComboBoxGetSelect(hDlg, IDC_COMBOBOX)
                  DIM dws AS DWSTRING = ComboBoxGetText(hDlg, IDC_COMBOBOX, item)
                  MsgBox hDlg, dws, MB_OK, "Message"
               END IF
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
