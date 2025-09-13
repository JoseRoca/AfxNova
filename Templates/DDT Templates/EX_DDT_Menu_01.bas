'#TEMPLATE DDT Dialog with a menu
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

CONST ID_OPEN = 401
CONST ID_EXIT = 402
CONST ID_OPTION1 = 403
CONST ID_OPTION2 = 404
CONST ID_HELP = 405
CONST ID_ABOUT = 406

' ========================================================================================
' Build the menu
' ========================================================================================
FUNCTION BuildMenu (BYVAL hDlg AS HWND) AS HMENU
   DIM lResult AS LONG

   ' ** First create a top-level menu:
   DIM hMenu AS HMENU = MenuNewBar

   ' ** Add a top-level menu item with a popup menu:
   DIM hPopup1 AS HMENU = MenuNewPopup
   MenuAddPopup hMenu, "&File", hPopup1, MF_ENABLED
   MenuAddString hPopup1, "&Open", ID_OPEN, MF_ENABLED
   MenuAddString hPopup1, "&Exit", ID_EXIT, MF_ENABLED
   MenuAddString hPopup1, "-", 0, 0

   ' ** Now we can add another item to the menu that will bring up a sub-menu. 
   ' First we obtain a new popup menu handle to distinguish it from the first 
   ' popup menu:
   DIM hPopup2 AS HMENU = MenuNewPopup

   ' ** Now add a new menu item to the first menu. 
   ' This item will bring up the sub-menu when selected:
   MenuAddPopup hPopup1, "&More Options", hPopup2, MF_ENABLED

   ' ** Now we will define the sub menu:
   MenuAddString hPopup2, "Option &1", ID_OPTION1, MF_ENABLED
   MenuAddString hPopup2, "Option &2", ID_OPTION2, MF_ENABLED

   ' ** Finally, we'll add a second top-level menu and popup.
   ' For this popup, we can reuse the first popup variable:
   hPopup1 = MenuNewPopup
   MenuAddPopup hMenu, "&Help", hPopup1, MF_ENABLED
   MenuAddString hPopup1, "&Help", ID_HELP, MF_ENABLED
   MenuAddString hPopup1, "-", 0, 0
   MenuAddString hPopup1, "&About", ID_ABOUT, MF_ENABLED
   
   ' Attach the menu to the dialog
   MenuAttach hMenu, hDlg

   ' // Bold the Exit menu item
   MenuBoldItem(hMenu, ID_EXIT)
   ' // Right justify the Help menu
'   MenuRightJustifyItem(hMenu, 1)

   RETURN hMenu
END FUNCTION
' ========================================================================================

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
   DIM hDlg AS HWND = DialogNew(0, "DDT with FreeBasic - Menu Demo",50, 50, 175, 75, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add controls to the dialog
   ControlAdd "GroupBox", hDlg, 101, "Just a simple question", 5, 5, 160, 55
   ControlAdd "Label", hDlg, 102, "What's your name ?", 10, 20, 80, 10
   ControlAdd "Edit", hDlg, 103, "", 90, 19, 70, 12
   ControlAdd("Button", hDlg, 104, "&Ok", 80, 40, 50, 12, BS_DEFPUSHBUTTON)

   ' // Build and attach a menu
   DIM hMenu AS HMENU = BuildMenu(hDlg)

   ' // Create a keyboard accelerator table
   AddAccelerator hDlg, FVIRTKEY OR FCONTROL, "O", ID_OPEN  ' // Ctrl+O - Open
   AddAccelerator hDlg, FVIRTKEY OR FCONTROL, "H", ID_HELP  ' // Ctrl+H - Help
   AddAccelerator hDlg, FVIRTKEY OR FCONTROL, "E", ID_EXIT  ' // Ctrl+E - Exit
   AddAccelerator hDlg, FVIRTKEY OR FCONTROL, "A", ID_ABOUT ' // Ctrl+A - About
   CreateAccelTable(hDlg)

   ' // Display and activate the dialog as modal
   DialogShowModal(hDlg, @DlgProc)

'   ' // Display and activate the dialog as modeless
'   DialogShowModeless(@DlgProc)

'   ' // Message handler loop
'   DO
'      pDlg.DialogDoEvents
'   LOOP WHILE IsWindow(hDlg)

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
            CASE 104 'IDOK
               DIM dws AS DWSTRING = AfxGetWindowText(GetDlgItem(hDlg, 103))
               IF LEN(dws) THEN MsgBox("Your name is " & dws, MB_ICONINFORMATION OR MB_TASKMODAL, "Answer")
               IF LEN(dws) = 0 THEN MsgBox("Hmm ...", MB_ICONEXCLAMATION OR MB_TASKMODAL, "No answer")
           CASE ID_OPEN TO ID_ABOUT
               MsgBox("WM_COMMAND received from a menu item!", MB_TASKMODAL, "Test menu")
         END SELECT

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
