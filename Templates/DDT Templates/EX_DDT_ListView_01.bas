'#TEMPLATE DDT Dialog - ListView Demo
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

#define IDC_LISTVIEW 1001
#define IDC_TEST 1002
#define IDC_OK 1003

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

   ' // Create dialog using dialog units
   DIM hDlg AS HWND = DialogNew(0, "ListView Demo",,, 440, 195, WS_OVERLAPPEDWINDOW OR DS_CENTER)
'   ' // Create custom dialog using dialog units
'   DIM hDlg AS HWND = DialogNew("MyClassName", 0, "ListView Demo",,, 440, 195, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ControlAddListView, hDlg, IDC_LISTVIEW, "", 5, 5, 430, 165, WS_VISIBLE OR WS_TABSTOP OR LVS_REPORT OR LVS_SHOWSELALWAYS OR LVS_SHAREIMAGELISTS OR LVS_AUTOARRANGE OR LVS_EDITLABELS OR LVS_ALIGNTOP
   ControlAddButton, hDlg, IDC_TEST, "&Test", 5, 176, 50, 12
   ControlAddButton, hDlg, IDC_OK, "&Ok", 60, 176, 50, 12, BS_DEFPUSHBUTTON

   ' // Add some extended styles
   DIM dwExStyle AS DWORD = ListviewGetStyleXX(hDlg, IDC_LISTVIEW)
   dwExStyle OR= LVS_EX_FULLROWSELECT OR LVS_EX_GRIDLINES
   ListViewSetStyleXX(hDlg, IDC_LISTVIEW, dwExStyle)

   ' // Set the text color of the ListView
   SendDlgItemMessageW(hDlg, IDC_LISTVIEW, LVM_SETTEXTCOLOR, 0, RGB_BLUE)


   ' // Add the header's column names
   DIM lvc AS LVCOLUMNW
   lvc.mask = LVCF_FMT OR LVCF_WIDTH OR LVCF_TEXT OR LVCF_SUBITEM
   DIM wszText AS WSTRING * 260
   FOR i AS LONG = 0 TO 9
      wszText = "Column " & STR(i)
      ListViewInsertColumn(hDlg, IDC_LISTVIEW, i, wszText, ScaleForDpiX(hDlg, 70))
   NEXT

   ' // Populate the ListView with some data
   FOR i AS LONG = 0 to 29
      wszText = "Column 1 Row " + WSTR(i)
      ListViewInsertItem(hDlg, IDC_LISTVIEW, i, 0, wszText)
      FOR x AS LONG = 0 TO 9
         wszText = "Column " & WSTR(x) & " Row " + WSTR(i)
         ListViewSetText(hDlg, IDC_LISTVIEW, i, x, wszText)
      NEXT
   NEXT

   ' // Anchor the controls for automatic resizing or reposition
   ControlAnchor(hDlg, IDC_LISTVIEW, AFX_ANCHOR_HEIGHT_WIDTH)
   ControlAnchor(hDlg, IDC_TEST, AFX_ANCHOR_BOTTOM)
   ControlAnchor(hDlg, IDC_OK, AFX_ANCHOR_BOTTOM)

   ' // Set the focus in the listview control
   ControlSetFocus(hDlg, IDC_LISTVIEW)

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
            CASE IDC_OK
               MessageBox(0, "Ok", "Message", MB_ICONEXCLAMATION OR MB_TASKMODAL)
            CASE IDC_TEST
               ' Reserved to add code to test ListView functions

         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
