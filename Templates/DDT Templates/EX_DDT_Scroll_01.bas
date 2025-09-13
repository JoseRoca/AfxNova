'#TEMPLATE DDT Scrollable Dialog with a ListBox
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
   DIM hDlg AS HWND = DialogNewPixels(0, "DDT - Scrollable window",,, 320, 335, WS_OVERLAPPEDWINDOW)

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

   ' // Make the dialog scrollable
   DialogSetViewPort(hDlg, 250, 260)

   ' // Center the window (must be done after shrinking the client size)
   DialogCenter(hDlg)

   ' // Display and activate the dialog as modal
'   DialogShowModal(hDlg, @DlgProc)

   ' // Display and activate the dialog as modeless
   DialogShowModeless(hDlg, @DlgProc)

   ' // Message handler loop
   DO
      DialogDoEvents
   LOOP WHILE IsWin(hDlg)

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
               ' // Do something with the ListBox
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
