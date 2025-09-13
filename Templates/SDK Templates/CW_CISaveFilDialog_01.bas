' ########################################################################################
' Microsoft Windows
' File: CW_CISaveFileDialog_01.bas
' Contents: Save file dialog
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#include once "AfxNova/CWindow.inc"
#include once "AfxNova/CIOpenSaveFile.inc"
USING AfxNova

CONST IDC_RICHEDIT = 1001
CONST IDC_TEST = 1002

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

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

   DIM pWindow AS CWindow
   DIM hWin AS HWND = pWindow.Create(NULL, "ISaveFile Dialog", @WndProc)
   pWindow.SetClientSize(800, 450)
   pWindow.Center

   ' // Add buttons
   pWindow.AddControl("Button", hWin, IDC_TEST, "&Test", pWindow.ClientWidth - 195, pWindow.ClientHeight - 35, 75, 23)
   pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", pWindow.ClientWidth - 95, pWindow.ClientHeight - 35, 75, 23)
  
   ' // Anchor the controls
   pWindow.AnchorControl(IDC_TEST, AFX_ANCHOR_BOTTOM_RIGHT)
   pWindow.AnchorControl(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

   ' // Dispatch Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window callback procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC hFocus AS HWND

   SELECT CASE uMsg

      CASE WM_NCACTIVATE
         ' // Save the handle of the control that has the focus
         IF wParam = 0 THEN hFocus = GetFocus
         ' // Note: Don't use EXIT FUNCTION

      CASE WM_SETFOCUS
         ' // Post a message to set the focus later, since some
         ' // Windows actions can steal it if we set it here
         IF hFocus THEN
            PostMessage hWnd, WM_USER + 999, cast(WPARAM, hFocus), 0
            hFocus = 0
         END IF

      CASE WM_USER + 999
         ' // Set the focus and show the line an column in the status bar
         IF wParam THEN SetFocus(cast(HWND, wParam))
         EXIT FUNCTION

      CASE WM_SYSCOMMAND
         ' // Capture this message and send a WM_CLOSE message
         IF (wParam AND &HFFF0) = SC_CLOSE THEN
            SendMessage hwnd, WM_CLOSE, 0, 0
            EXIT FUNCTION
         END IF

      CASE WM_COMMAND
         
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)

            ' // If ESC key pressed, close the application sending an WM_CLOSE message
            CASE IDCANCEL
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF

            ' // For testing purposes
            CASE IDC_TEST
               SCOPE
               DIM psfd AS CISaveFileDialog
               ' // Set the file types
               psfd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
               psfd.AddFileType("Executable files", "*.exe;*.dll")
               psfd.AddFileType("All files", "*.*")
               psfd.SetFileTypes()
               ' // Optional: Set the title of the dialog
               '   psfd.SetTitle("Save File Dialog")
               ' // Set the folder
               psfd.SetFolder(CURDIR)
               psfd.SetDefaultExtension("bas")
               psfd.SetFileTypeIndex(1)
               ' // Display the dialog
               DIM hr AS HRESULT = psfd.ShowSave(hwnd)
               ' // Folder name
               OutputDebugStringW("Folder name: " & psfd.GetFolder)
               ' // Get the result
               IF hr = S_OK THEN
                  OutputDebugStringW(psfd.GetResultString)
               END IF
               END SCOPE

         END SELECT

    	CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================