' ########################################################################################
' Microsoft Windows
' File: CW_ComboBox_DarkMode.bas
' Contents: CWindow - ComboBox
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

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

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

   ' // Creates the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - ComboBox", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center
   ' // Set the main window background color
   pWindow.SetBackColor(RGB_BLACK)

   ' // Adds a combobox
   DIM hComboBox AS HWND = pWindow.AddControl("Combobox", hWin, IDC_COMBOBOX, "", 80, 30, 230, 100)
   ' // Set dark mode
   SetWindowTheme(hComboBox, "DarkMode_CFD", NULL)
   ' // Get the handle of the edit control of the combobox
   DIM hEditCombo AS HWND = GetWindow(hComboBox, GW_CHILD)
   ' // Set dark mode
   SetWindowTheme(hEditCombo, "DarkMode_Explorer", NULL)
   ' // Colors of the edit text
   pWindow.SetCtlColors(hEditCombo, RGB_WHITE, RGB_BLACK)
   ' // Get the handle of the listbox control of a combobox
   DIM cbi AS COMBOBOXINFO 
   cbi.cbSize = SIZEOF(COMBOBOXINFO)
   GetComboBoxInfo (hComboBox, @cbi)
   ' // Colors of the list text
   pWindow.SetCtlColors(cbi.hwndList, RGB_WHITE, RGB_BLACK)
   ' // Anchors the button to the bottom and the right side of the main window
   pWindow.AnchorControl(IDC_COMBOBOX, AFX_ANCHOR_WIDTH)

   ' // Fill the control with some data
   DIM dwsText AS DWSTRING
   FOR i AS LONG = 1 TO 9
      dwsText = "Item " & RIGHT("00" & WSTR(i), 2)
      ComboBox_AddString(hComboBox, *dwsText)
   NEXT

   ' // Select the first item in the combo box
   ComboBox_SetCursel(hComboBox, 0)

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
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      ' // Theme has changed
      CASE WM_THEMECHANGED
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
            CASE IDC_COMBOBOX
               ' // The selection has changed
               IF CBCTLMSG(wParam, lParam) = LBN_SELCHANGE THEN
                  ' // Handle of the combobox
                  DIM hCombobox AS HWND = GetDlgItem(hwnd, IDC_COMBOBOX)
                  ' // Retrieve the Item selected
                  DIM curSel AS LONG = ComboBox_GetCursel(hCombobox)
                  MessageBoxW hwnd, "You have selected " & _
                     AfxGetComboBoxText(hCombobox, curSel), "ComboBox Test", MB_OK
               END IF
         END SELECT
         RETURN 0

      CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
