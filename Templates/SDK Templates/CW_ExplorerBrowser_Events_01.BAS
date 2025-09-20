' ########################################################################################
' Microsoft Windows
' File: CW_ExplorerBrowser_01.bas
' Contents: CWindow - Explorer Browser with events
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#define _CEXPBROW_DEBUG_ 1
#INCLUDE ONCE "AfxNova/CWindow.inc"
#INCLUDE ONCE "AfxNova/CExplorerBrowser.inc"
USING AfxNova

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

   ' // Creates the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   DIM hWin AS HWND = pWindow.Create(NULL, "Explorer Browser", @WndProc)
   ' // Sizes it by setting the wanted width and height of its client area
   pWindow.SetClientSize(620, 400)
   ' // Centers the window
   pWindow.Center

   ' // Displays the window and dispatches the Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ########################################################################################
' CExplorerBrowserEventsImpl class
' Implementation of the IExplorerBrowserEvents callback interface
' ########################################################################################
TYPE CExplorerBrowserEventsImpl EXTENDS CExplorerBrowserEvents

   DECLARE FUNCTION OnNavigationPending (BYVAL pidlFolder AS PCIDLIST_ABSOLUTE ) AS HRESULT
   DECLARE FUNCTION OnViewCreated (BYVAL psv AS IShellView PTR) AS HRESULT
   DECLARE FUNCTION OnNavigationComplete (BYVAL pidlFolder AS PCIDLIST_ABSOLUTE ) AS HRESULT
   DECLARE FUNCTION OnNavigationFailed (BYVAL pidlFolder AS PCIDLIST_ABSOLUTE ) AS HRESULT

END TYPE

' ========================================================================================
' Notifies clients of a pending Explorer browser navigation to a Shell folder.
' ========================================================================================
FUNCTION CExplorerBrowserEventsImpl.OnNavigationPending (BYVAL pidlFolder AS PCIDLIST_ABSOLUTE ) AS HRESULT
   CEXPBROW_DP("- pidlFolder: " & WSTR(pidlFolder))
   DIM wszName AS WSTRING * MAX_PATH
   SHGetPathFromIDListW(pidlFolder, @wszName)
   CEXPBROW_DP("- " & wszName)
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifica a los clientes que se ha creado la vista del explorador Explorador y que se pueden modificar.
' ========================================================================================
FUNCTION CExplorerBrowserEventsImpl.OnViewCreated (BYVAL psv AS IShellView PTR) AS HRESULT
   CEXPBROW_DP("- pShellView: " & WSTR(psv))
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies clients that the Explorer browser has successfully navigated to a Shell folder.
' ========================================================================================
FUNCTION CExplorerBrowserEventsImpl.OnNavigationComplete (BYVAL pidlFolder AS PCIDLIST_ABSOLUTE ) AS HRESULT
   CEXPBROW_DP("- pidlFolder: " & WSTR(pidlFolder))
   DIM wszName AS WSTRING * MAX_PATH
   SHGetPathFromIDListW(pidlFolder, @wszName)
   CEXPBROW_DP("- " & wszName)
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Notifies clients that the Explorer browser has failed to navigate to a Shell folder.
' ========================================================================================
FUNCTION CExplorerBrowserEventsImpl.OnNavigationFailed (BYVAL pidlFolder AS PCIDLIST_ABSOLUTE ) AS HRESULT
   CEXPBROW_DP("- pidlFolder: " & WSTR(pidlFolder))
   DIM wszName AS WSTRING * MAX_PATH
   SHGetPathFromIDListW(pidlFolder, @wszName)
   CEXPBROW_DP("- " & wszName)
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ########################################################################################


' ########################################################################################
' Main window procedure
' ########################################################################################

FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC peb AS CExplorerBrowser PTR

   SELECT CASE uMsg

      ' // If an application processes this message, it should return zero to continue
      ' // creation of the window. If the application returns –1, the window is destroyed
      ' // and the CreateWindowExW function returns a NULL handle.
      CASE WM_CREATE
         ' // Creates an instance of the Explorer Browser
         peb = NEW CExplorerBrowser
         IF peb = NULL THEN RETURN 0
         peb->SetOptions(EBO_SHOWFRAMES)
         DIM fs AS FOLDERSETTINGS
         fs.ViewMode = FVM_DETAILS
         DIM rc AS RECT
         GetClientRect hwnd, @rc
         peb->Initialize(hwnd, @rc, @fs)
         ' // Set events
         DIM pebe AS ANY PTR = NEW CExplorerBrowserEventsImpl
         peb->SetEvents(pebe)
         ' // Navigate to the Profile folder
         DIM pidlBrowse AS ITEMIDLIST PTR
         IF SUCCEEDED(SHGetFolderLocation(NULL, CSIDL_PROGRAM_FILES, NULL, 0, @pidlBrowse)) THEN
             peb->BrowseToIDList(pidlBrowse, SBSP_ABSOLUTE)
             ILFree(pidlBrowse)
         END IF
         RETURN 0

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
         END SELECT
         RETURN 0

      CASE WM_SIZE
         ' // Resize the explorer browser
         DIM rc AS REct
         GetClientRect hwnd, @rc
         IF peb THEN peb->SetRect(NULL, rc)
         RETURN 0

      CASE WM_DESTROY
         ' // Destroy the CExplorerBrowser class
         IF peb THEN Delete peb
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         RETURN 0

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

