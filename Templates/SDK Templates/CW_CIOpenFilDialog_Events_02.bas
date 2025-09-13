' ########################################################################################
' Microsoft Windows
' File: CW_CIOpenFileDialog_Events_02.bas
' Contents: Open file dialog with dialog and control events
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José© Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#define _COSFD_DEBUG_ 1
#include once "AfxNova/CWindow.inc"
#include once "AfxNova/AfxCOM.inc"
#include once "AfxNova/CIOpenSaveFile.inc"
#include once "AfxNova/CIFileDialogCustomize.inc"
#include once "AfxNova/CIFileDialogEvents.inc"
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
   DIM hWin AS HWND = pWindow.Create(NULL, "DisplayIOpenFile", @WndProc)
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

' ########################################################################################
' CIFileDialogControlEventsImpl class
' Implementation of the IFileDialogControlEvents callback interface
' ########################################################################################
TYPE CIFileDialogControlEventsImpl EXTENDS CIFileDialogControlEvents

   DECLARE FUNCTION OnItemSelected (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   DECLARE FUNCTION OnButtonClicked (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION OnCheckButtonToggled (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL bChecked AS WINBOOL) AS HRESULT
   DECLARE FUNCTION OnControlActivating (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT

END TYPE

' // Override the virtual methods

' ========================================================================================
PRIVATE FUNCTION CIFileDialogControlEventsImpl.OnItemSelected (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   COSFD_DP("")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIFileDialogControlEventsImpl.OnButtonClicked (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
   COSFD_DP("")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIFileDialogControlEventsImpl.OnCheckButtonToggled (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL bChecked AS WINBOOL) AS HRESULT
   COSFD_DP("")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIFileDialogControlEventsImpl.OnControlActivating (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
   COSFD_DP("")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ########################################################################################


' ########################################################################################
' CIFileDialogEventsImpl class
' Implementation of the FileDialogEvents callback interface
' ########################################################################################
TYPE CIFileDialogEventsImpl EXTENDS CIFileDialogEvents

   DECLARE FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   DECLARE FUNCTION OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE FUNCTION OnItemSelected (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT

END TYPE

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogEventsImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   COSFD_DP(AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_IFileOpenDialog) OR _
      IsEqualIID(riid, @IID_IFileSaveDialog) OR _
      IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
      ' // Increment the reference count
      this.AddRef
      RETURN S_OK
   ELSE
      IF IsEqualIID(riid, @IID_IFileDialogControlEvents) THEN
         ' // Create an instance of our implemented IFileDialogControlEvents interface
         DIM pIFileDialogControlEvents AS CIFileDialogControlEventsImpl PTR = NEW CIFileDialogControlEventsImpl
         ' // Increment the reference count
         pIFileDialogControlEvents->AddRef
         ' // Return a pointer to the event's class
         *ppvObj = CAST(ANY PTR, pIFileDialogControlEvents)
         RETURN S_OK
      END IF
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' =====================================================================================

PRIVATE FUNCTION CIFileDialogEventsImpl.OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   ' // Before displaying the dialog get its coordinates
   DIM pOleWindow AS IOleWindow PTR
   ' // Get a pointer to the IOleWindow interface
   pfd->lpvtbl->QueryInterface(pfd, @IID_IOleWindow, @pOleWindow)
   ' // Get the window handle of the dialog
   DIM hOleWindow AS HWND
   IF pOleWindow THEN pOleWindow->lpvtbl->GetWindow(pOleWindow, @hOleWindow)
   COSFD_DP ("hOleWindow: " & WSTR(hOleWindow))
   IF hOleWindow THEN
      ' // Get he bounding rectangle of the parent window
      DIM AS RECT rcDlg, rcOwner
      GetWindowRect(GetParent(hOleWindow), @rcOwner)
      ' // Get he bounding rectangle of the open file dialog
      GetWindowRect(hOleWindow, @rcDlg)
      ' // Maps the open file dialog coordinates relative to its parent window
      MapWindowPoints(GetParent(hOleWindow), hOleWindow, CAST(POINT PTR, @rcDlg), 2)
   END IF
   RETURN S_OK
END FUNCTION
' ########################################################################################

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
               DIM pofd AS CIOpenFileDialog
               ' // Set the file types
               pofd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
               pofd.AddFileType("Executable files", "*.exe;*.dll")
               pofd.AddFileType("All files", "*.*")
               pofd.SetFileTypes()
               ' // Multiple selection (default is single selection)
               DIM options AS FILEOPENDIALOGOPTIONS = pofd.GetOptions
               pofd.SetOptions(options OR FOS_ALLOWMULTISELECT)
               ' // Optional: Set the title of the dialog
               '   pofd.SetTitle("A Single-Selection Dialog")
               ' // Set the folder
               pofd.SetFolder(CURDIR)
               pofd.SetDefaultExtension("bas")
               pofd.SetFileTypeIndex(1)
               ' // Customization
               ' // Get a pointer to the IFileDialog interface
               DIM pfdlg AS IFileDialog PTR = pofd.GetIFileDialogPtr
               IF pfdlg THEN
                  ' // Initialize an instance of the CIFileDialogCustomize class
                  DIM pCust AS CIFileDialogCustomize = pfdlg
                  ' // Relese the IFileDialog interface
                  pfdlg->lpvtbl->Release(pfdlg)
                  ' // Add controls to the dialog
                  IF pCust.GetPtr THEN
                     ' Call methods of the CIFileDialogCustomize class
                     pCust.AddCheckButton(1001, "My check button", TRUE)
                     pCust.AddComboBox(1002)
                     pCust.AddEditBox(1003, "My edit control")
                     pCust.AddMenu(1004, "My menu")
                     pCust.AddPushButton(1005, "My push button")
                     pCust.AddSeparator(1006)
                     pCust.AddText(1007, "My text")
                     pCust.EnableOpenDropDown(1008)
                  END IF
               END IF
               ' // Set events
               DIM pfde AS ANY PTR = NEW CIFileDialogEventsImpl
               pofd.SetEvents(pfde)
               ' // Display the dialog
               DIM hr AS HRESULT = pofd.ShowOpen(hwnd)
               ' // Folder name
               OutputDebugStringW("Folder name: " & pofd.GetFolder)
               ' *** Single selection ***
               ' // Get the result
               'IF hr = S_OK THEN
               '   OutputDebugStringW pofd.GetResult()
               'END IF
               ' *** Multiple selection ***
               IF hr = S_OK THEN
                  DIM dwsRes AS DWSTRING = pofd.GetResultsString
                  FOR i AS LONG = 1 TO pofd.GetResultsCount
                     OutputDebugStringW pofd.ParseResults(dwsRes, i)
                  NEXT
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

