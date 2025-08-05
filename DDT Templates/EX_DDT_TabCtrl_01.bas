'#TEMPLATE DDT Dialog with a tan control and tab pages
#define UNICODE
#define _WIN32_WINNT &h0602
#include once "AfxNova/DDT.inc"
#include once "AfxNova/CTabPage.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

CONST IDC_TAB       = 1001
CONST IDC_EDIT1     = 1002
CONST IDC_EDIT2     = 1003
CONST IDC_BTNSUBMIT = 1004
CONST IDC_COMBO     = 1005
CONST IDC_LISTBOX   = 1006

' // Forward declarations
DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR
DECLARE FUNCTION TabPage1_WndProc(BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION TabPage2_WndProc(BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION TabPage3_WndProc(BYVAL hWnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

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
   DIM hDlg AS HWND = DialogNewPixels(0, "DDT Dialog - Tab control", 0, 0, 500, 350, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Add a tab control to the dialog
   ControlAddTab, hDlg, IDC_TAB, "", 10, 10, 480, 290
   ' // Anchor the button to the bottom and the right side of the dialog
   ControlAnchor(hDlg, IDC_TAB, AFX_ANCHOR_HEIGHT_WIDTH)

   ' // Creates the first tab page
   DIM hTab AS HWND = ControlHandle(hDlg, IDC_TAB)
   DIM pTabPage1 AS CTabPage PTR = NEW CTabPage
   pTabPage1->InsertPage(hTab, 0, "Tab 1", -1, @TabPage1_WndProc)
   pTabPage1->SetBackColor(RGB_GOLD) 

   ' // Creates the second tab page
   DIM pTabPage2 AS CTabPage PTR = NEW CTabPage
   pTabPage2->InsertPage(hTab, 1, "Tab 2", -1, @TabPage2_WndProc)
   pTabPage2->SetBackColor(RGB_DARKTURQUOISE)

   ' // Creates the third tab page
   DIM pTabPage3 AS CTabPage PTR = NEW CTabPage
   pTabPage3->InsertPage(hTab, 2, "Tab 3", -1, @TabPage3_WndProc)
   pTabPage3->SetBackColor(RGB_CHARTREUSE)

   ' // Displays the first tab page
   ShowWindow pTabPage1->hTabPage, SW_SHOW
   ' // Set the focus to the first tab
   TabCtrl_SetCurFocus(hTab, 0)

   ' // Adds a button
   ControlAddButton, hDlg, IDCANCEL, "&Close", 412, 315, 75, 23
   ' // Anchors the button to the bottom and the right side of the main window
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

      CASE WM_GETMINMAXINFO
         ' Set the pointer to the address of the MINMAXINFO structure
         DIM ptmmi AS MINMAXINFO PTR = CAST(MINMAXINFO PTR, lParam)
         ' Set the minimum and maximum sizes that can be produced by dragging the borders of the window
         ptmmi->ptMinTrackSize.x = 460 * rxRatioForDpi(hDlg)
         ptmmi->ptMinTrackSize.y = 340 * ryRatioForDpi(hDlg)

      CASE WM_SIZE
         ' // Get the tab control handle
         DIM hTab AS HWND = ControlHandle(hDlg, IDC_TAB)
         ' // Get a pointer to the tab page
         DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(hTab, -1)
         ' // Resize the tab pages
         IF pTabPage THEN pTabPage->ResizePages

      CASE WM_PAINT
         DIM hTab AS HWND = ControlHandle(hDlg, IDC_TAB)
         SetWindowPos hTab, NULL, 0, 0, 0, 0, SWP_NOZORDER OR SWP_NOMOVE OR SWP_NOSIZE OR SWP_DRAWFRAME

      CASE WM_NOTIFY
         DIM pTabPage AS CTabPage PTR    ' // Tab page object reference
         DIM ptnmhdr AS NMHDR PTR        ' // Information about a notification message
         ptnmhdr = CAST(NMHDR PTR, lParam)
         SELECT CASE ptnmhdr->idFrom
            CASE IDC_TAB
               SELECT CASE ptnmhdr->code
                  CASE TCN_SELCHANGE
                     ' // Show the selected page
                     pTabPage = AfxCTabPagePtr(ptnmhdr->hwndFrom, -1)
                     IF pTabPage THEN ..ShowWindow pTabPage->hTabPage, SW_SHOW
                  CASE TCN_SELCHANGING
                     ' // Hide the current page
                     pTabPage = AfxCTabPagePtr(ptnmhdr->hwndFrom, -1)
                     IF pTabPage THEN ..ShowWindow pTabPage->hTabPage, SW_HIDE
               END SELECT
         END SELECT

      CASE WM_CLOSE
         ' // Destroy the tab pages
         DIM hTab AS HWND = ControlHandle(hDlg, IDC_TAB)
         AfxDestroyAllTabPages(hTab)
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
' ========================================================================================

' ========================================================================================
' Tab page 1 window procedure
' ========================================================================================
FUNCTION TabPage1_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM hBrush AS HBRUSH, rc AS RECT, tlb AS LOGBRUSH

   SELECT CASE uMsg

      CASE WM_CREATE
         ' // Get a pointer to the TabPage class
         DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(GetParent(hwnd), 0)
         IF pTabPage = NULL THEN EXIT FUNCTION
         ' // Add controls to the first page
         DIM hLabel1 AS HWND = pTabPage->AddControl("Label", hwnd, -1, " First name", 15, 35, 121, 21)
         DIM hLabel2 AS HWND = pTabPage->AddControl("Label", hwnd, -1, " Last name", 15, 70, 121, 21)
         pTabPage->AddControl("Edit", hwnd, IDC_EDIT1, "", 165, 35, 250, 21)
         pTabPage->AddControl("Edit", hwnd, IDC_EDIT2, "", 165, 70, 250, 21)
         pTabPage->AddControl("Button", hwnd, IDC_BTNSUBMIT, "Submit", 340, 185, 76, 25)
         ' // Anchor the controls
         pTabPage->AnchorControl(GetDlgItem(hwnd, IDC_EDIT1), AFX_ANCHOR_WIDTH)
         pTabPage->AnchorControl(GetDlgItem(hwnd, IDC_EDIT2), AFX_ANCHOR_WIDTH)
         pTabPage->AnchorControl(GetDlgItem(hwnd, IDC_BTNSUBMIT), AFX_ANCHOR_BOTTOM_RIGHT)
         ' // Colors
         pTabPage->SetCtlColors(HLabel1, RGB_YELLOW, RGB_RED)
         pTabPage->SetCtlColors(HLabel2, RGB_YELLOW, RGB_RED)
         ' // Set the focus in the edit control
         SetFocus(GetDlgItem(hwnd, IDC_EDIT1))
         EXIT FUNCTION

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDC_BTNSUBMIT
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  MessageBoxW(hWnd, "Submit", "Tab 1", MB_OK)
                  EXIT FUNCTION
               END IF
         END SELECT

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Tab page 2 window procedure
' ========================================================================================
FUNCTION TabPage2_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM hBrush AS HBRUSH, rc AS RECT, tlb AS LOGBRUSH

   SELECT CASE uMsg

      CASE WM_CREATE
         ' // Get a pointer to the TabPage class
         DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(GetParent(hwnd), 1)
         IF pTabPage = NULL THEN EXIT FUNCTION
         ' // Add a combobox to the second page
         DIM hComboBox AS HWND = pTabPage->AddControl("ComboBox", hwnd, IDC_COMBO, "", 20, 20, 191, 105)
         ' // Anchor the control
         pTabPage->AnchorControl(GetDlgItem(hwnd, IDC_COMBO), AFX_ANCHOR_WIDTH)
         ' // Fill the combobox with some data
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = 1 TO 9
            wszText = "Item " & RIGHT("00" & STR(i), 2)
            ComboBox_AddString(hComboBox, @wszText)
         NEXT
         ' // Select the first item in the combo box
         ComboBox_SetCursel(hComboBox, 0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Tab page 3 window procedure
' ========================================================================================
FUNCTION TabPage3_WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   DIM hBrush AS HBRUSH, rc AS RECT, tlb AS LOGBRUSH

   SELECT CASE uMsg

      CASE WM_CREATE
         ' // Get a pointer to the TabPage class
         DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(GetParent(hwnd), 2)
         IF pTabPage = NULL THEN EXIT FUNCTION
         ' // Add a combobox to the third page
         DIM hListBox AS HWND = pTabPage->AddControl("ListBox", hwnd, IDC_LISTBOX)
         pTabPage->SetWindowPos hListBox, NULL, 15, 20, 161, 120, SWP_NOZORDER
         ' // Anchor the control
         pTabPage->AnchorControl(GetDlgItem(hwnd, IDC_LISTBOX), AFX_ANCHOR_WIDTH)
         ' // Fill the combobox with some data
         DIM wszText AS WSTRING * 260
         FOR i AS LONG = 1 TO 9
            wszText = "Item " & RIGHT("00" & STR(i), 2)
            ListBox_AddString(hListBox, @wszText)
         NEXT
         ' // Select the first item in the combo box
         ListBox_SetCursel(hListBox, 0)
         EXIT FUNCTION

   END SELECT

   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
