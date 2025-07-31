# CTabPage Class

The `CTabPage` class extends the `CWindow` class and allows to create generic windows used as a tab pages of a tab control.

### Constructor (CTabPage)

Creates a generic window used as a tab page of a tab control.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [hTabPage](#htabpage) | Returns the window handle of the tab page. |
| [InsertPage](#insertpage) | Adds a tab page and creates a generic window that will be associated with the page. |
| [ResizePage](#resizepage) | Resizes the currently selected tab page. |
| [ResizePages](#resizepages) | Resizes all the tab pages associated with the tab control |

### Procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxCTabPagePtr](#afxctabpageptr) | Returns a pointer to the **CTabPage** class given the handle of the tab control to which the tab page is associated and the zero-based tab index. |
| [AfxDestroyTabPage](#afxdestroytabpage) | Detroys a tab page. |
| [AfxDestroyAllTabPages](#afxdestroyalltabpages) | Detroys all the tab pages. |
| [AfxResizeTabPages](#afxresizetabpages) | Resizes all the tab pages associated with a tab control. |

---

#### Example

```
' ########################################################################################
' Microsoft Windows
' File: CW_TabCtrl_01.bas
' Contents: CWindow - Tab control
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
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
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
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

   ' // Creates the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Tab control", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(500, 350)
   ' // Centers the window
   pWindow.Center

   ' // Adds a tab control
   DIM AS LONG cx, cy
   cx = pWindow.ClientWidth - 20
   cy = pWindow.ClientHeight - 60
   DIM hTab AS HWND = pWindow.AddControl("Tab", hWin, IDC_TAB, "", 10, 10, cx, cy)
   ' // Anchor the control
   pWindow.AnchorControl(hTab, AFX_ANCHOR_HEIGHT_WIDTH)

   ' // Creates the first tab page
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
   cx = pWindow.ClientWidth - 88
   cy = pWindow.ClientHeight - 37
   DIM hButton AS HWND = pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", cx, cy, 75, 23)
   ' // Anchors the button to the bottom and the right side of the main window
   pWindow.AnchorControl(hButton, AFX_ANCHOR_BOTTOM_RIGHT)

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
               RETURN 0
         END SELECT

      CASE WM_GETMINMAXINFO
         ' Set the pointer to the address of the MINMAXINFO structure
         DIM ptmmi AS MINMAXINFO PTR = CAST(MINMAXINFO PTR, lParam)
         ' Set the minimum and maximum sizes that can be produced by dragging the borders of the window
         DIM pWindow AS CWindow PTR = AfxCWindowPtr(hwnd)
         IF pWindow THEN
            ptmmi->ptMinTrackSize.x = 460 * pWindow->rxRatio
            ptmmi->ptMinTrackSize.y = 340 * pWindow->ryRatio
         END IF
         RETURN 0

      CASE WM_SIZE
         ' // Get the tab control handle
         DIM hTab AS HWND = GetDlgItem(hwnd, IDC_TAB)
         ' // Get a pointer to the tab page
         DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(hTab, -1)
         ' // Resize the tab pages
         IF pTabPage THEN pTabPage->ResizePages
         RETURN 0

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

    	CASE WM_DESTROY
         ' // Destroy the tab pages
         DIM hTab AS HWND = GetDlgItem(hwnd, IDC_TAB)
         AfxDestroyAllTabPages(hTab)
         ' // Quit the application
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

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
         pTabPage->AddControl("Edit", hwnd, IDC_EDIT1, "", 165, 35, 186, 21)
         pTabPage->AddControl("Edit", hwnd, IDC_EDIT2, "", 165, 70, 186, 21)
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
```
---

## InsertPage

Adds a tab page and creates a generic window that will be associated with the page.

```
FUNCTION InsertPage (BYVAL hTab AS HWND, BYVAL nPage AS LONG, BYREF wszTitle AS WSTRING = "", _
   BYVAL nImage AS LONG = -1, BYVAL lpfnWndProc AS WNDPROC = NULL, _
   BYVAL dwStyle AS DWORD = WS_CLIPSIBLINGS OR WS_CLIPCHILDREN, BYVAL dwExStyle AS DWORD = 0) AS HWND
```

| Parameter  | Description |
| ---------- | ----------- |
| *hTab* | Handle to the Tab control. |
| *nPage* | The zero based position of the page to be inserted. |
| *wszTitle* | Optional. The text to be displayed of the tab area. |
| *nImage* | Optional. Index in the tab control's image list, or -1 if there is no image for the tab. |
| *lpfnWndProc* | Address of the window callback procedure. |

#### Return value

The handle of the new tab page.

#### Usage example

```
' // Create a tab page
DIM pTabPage AS CTabPage PTR = NEW CTabPage
pTabPage->InsertPage(hTab, 0, "Tab 1", -1, @TabPage_WndProc)
' // Add controls to it
pTabPage->AddControl("Label", pTabPage->hTabPage, -1, "First name", 15, 15, 121, 21)
pTabPage->AddControl("Label", pTabPage->hTabPage, -1, "Last name", 15, 50, 121, 21)
pTabPage->AddControl("Edit", pTabPage->hTabPage, IDC_EDIT1, "", 165, 15, 186, 21)
pTabPage->AddControl("Edit", pTabPage->hTabPage, IDC_EDIT2, "", 165, 50, 186, 21)
pTabPage->AddControl("Button", pTabPage->hTabPage, IDC_BTNSUBMIT, "Submit", 340, 185, 76, 26, BS_DEFPUSHBUTTON)
```
---

## ResizePage

Resizes the currently selected tab page,

```
FUNCTION ResizePage () AS BOOLEAN
```
---

## ResizePages

Resizes the all the pages associated with the tab control

```
FUNCTION ResizePages () AS BOOLEAN
```
---

## AfxCTabPagePtr

Returns a pointer to the `CTabPage` class given the handle of the tab control to which the tab page is associated and the zero-based tab index.

```
FUNCTION AfxCTabPagePtr (BYVAL hTab AS HWND, BYVAL idx AS LONG = -1) AS CTabPage PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hTab* | Handle to the Tab control. |
| *hTab* | Optional. Zero-based index of the tab. If this parameter is ommited, the function will return the pointer of the currently selected tab, if any. |

#### Usage example

```
DIM pTabPage AS CTabPage PTR = AfxCTabPagePtr(hTab, 1)
```
---

## AfxDestroyTabPage

Detroys the specified tab page.

```
FUNCTION AfxDestroyTabPage (BYVAL hTab AS HWND, BYVAL idx AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hTab* | Handle to the Tab control. |
| *hTab* | Zero-based index of the tab. If this parameter is ommited, the function will return the pointer of the currently selected tab, if any. |

#### Return value

TRUE or FALSE.

#### Usage example

```
AfxDestroyTabPage(hTab, 1)
```
---

## AfxDestroyAllTabPages

Detroys all the tab pages.

```
FUNCTION AfxDestroyAllTabPages (BYVAL hTab AS HWND) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hTab* | Handle to the Tab control. |

#### Return value

TRUE or FALSE.

#### Usage example

```
AfxDestroyAllTabPages(hTab)
```
---

## AfxResizeTabPages

Resizes all the tab pages associated with a tab control.

```
FUNCTION AfxResizeTabPages (BYVAL hTab AS HWND) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hTab* | Handle to the Tab control. |

#### Return value

TRUE or FALSE.

#### Usage example

```
AfxResizeTabPages(hTab)
```
---


