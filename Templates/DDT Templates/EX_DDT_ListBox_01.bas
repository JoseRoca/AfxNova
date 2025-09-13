'#TEMPLATE DDT Dialog with a ListBox
#define UNICODE
#define _WIN32_WINNT &h0602
#include once "AfxNova/DDT.inc"
USING AfxNova

#define IDC_LISTBOX 1001

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles
   AfxEnableVisualStyles

   ' // Create a new dialog
   DIM hDlg AS HWND = DialogNewPixels(0, "DDT - ListBox", 0, 0, 320, 335, WS_OVERLAPPEDWINDOW)
   ' // Center the window
   DialogCenter(hDlg)

   ' // Add a listbox
   ControlAddListBox, hDlg, IDC_LISTBOX, "", 8, 8, 300, 280

   ' // Fill the list box
   DIM dwsText AS DWSTRING
   FOR i AS LONG = 1 TO 50
      dwsText = "Item " & RIGHT("00" & WSTR(i), 2)
      ListBoxAdd(hDlg, IDC_LISTBOX, dwsText)
   NEXT

   ' // Add a cancel button
   ControlAddButton, hDlg, IDCANCEL, "&Cancel", 233, 298, 75, 23

  ' // Anchor the controls
  ControlAnchor hDlg, IDC_LISTBOX, AFX_ANCHOR_HEIGHT_WIDTH
  ControlAnchor hDlg, IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT

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
                  DialogSend hDlg, WM_CLOSE, 0, 0
               END IF
            CASE IDC_LISTBOX
               IF CBCTLMSG(wParam, lParam) = LBN_DBLCLK THEN
                  ' // Get the current selection
                  DIM curSel AS LONG = ListBoxGetSelect(hDlg, IDC_LISTBOX)
                  ' // Get the length of the ListBox item text
                  DIM dwsText AS DWSTRING = ListBoxGetText(hDlg, IDC_LISTBOX, curSel)
                  MsgBox(hDlg, dwsText, MB_OK,  "ListBox test")
               END IF
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
