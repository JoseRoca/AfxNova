' ########################################################################################
' Microsoft Windows
' File: ISaveFile.inc
' Contents: Save File dialog
' Compiler: FreeBASIC 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ========================================================================================
' ISaveFileDialog class
' ========================================================================================
TYPE ISaveFileDialog

Private:
   m_IFileSaveDialog AS IFileSaveDialog PTR             ' // Pointer to the IFileSaveDialog interface
   m_hwndOwner AS HWND                                  ' // Handle of the owner of the dialog
   m_xPos AS LONG
   m_yPos AS LONG
   m_dwsTitle AS DWSTRING                               ' // The title to be displayed in the title bar of the dialog box
   m_dwsFolder AS DWSTRING                              ' // The name of the initial file directory to be displayed
   m_rgFilterSpec (ANY) AS COMDLG_FILTERSPEC            ' // Array of COMDLG_FILTERSPEC structures
   m_startFile AS DWSTRING                              ' // File name to be used as the initial file selection
   m_defExt AS DWSTRING                                 ' // Default extension to be appended to the selected file name if the user does not enter it
   m_Flags AS DWORD                                     ' // Flags
   m_pFileDialogEvents AS IFileDialogEvents PTR         ' // Pointer to the IFileDialogEvents interface
   m_dwDlgEvtCookie AS DWORD                            ' // Dialog events cookie
   m_IFileDialogCustomize AS IFileDialogCustomize PTR   ' // Pointer to the IFileDialogCustomize interface

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   DECLARE DESTRUCTOR 
   DECLARE FUNCTION ShowSave (BYVAL hwndOwner AS HWND) AS HRESULT
   DECLARE SUB AddFileType (BYVAL pwszName AS WSTRING PTR, BYVAL pwszSpec AS WSTRING PTR)
   DECLARE FUNCTION SetFileTypes () AS HRESULT
   DECLARE FUNCTION SetFileTypes (BYVAL cFileTypes AS UINT, BYVAL rgFilterSpec AS COMDLG_FILTERSPEC PTR) AS HRESULT
   DECLARE FUNCTION SetFileTypeIndex (BYVAL iFileType AS UINT) AS HRESULT
   DECLARE FUNCTION GetFileTypeIndex () AS UINT
   DECLARE FUNCTION Advise (BYVAL pfde as IFileDialogEvents PTR) AS DWORD
   DECLARE FUNCTION Unadvise (BYVAL pfde as IFileDialogEvents PTR, BYVAL dwCookie AS DWORD) AS HRESULT
   DECLARE FUNCTION SetOptions (BYVAL fos AS FILEOPENDIALOGOPTIONS) AS HRESULT
   DECLARE FUNCTION GetOptions () AS FILEOPENDIALOGOPTIONS
   DECLARE FUNCTION SetDefaultFolder (BYVAL psi AS IShellItem PTR) AS HRESULT
   DECLARE FUNCTION SetFolder (BYVAL pwzFolderName AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION GetFolder (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   DECLARE FUNCTION GetCurrentSelection (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   DECLARE FUNCTION SetFileName (BYVAL pwszName AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION GetFileName () AS DWSTRING
   DECLARE FUNCTION SetTitle (BYVAL pwszTitle AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetOkButtonLabel (BYVAL pwszText AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetFileNameLabel (BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION GetResult (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   DECLARE FUNCTION AddPlace (BYVAL psi AS IShellItem PTR, BYVAL fdap AS FDAP) AS HRESULT
   DECLARE FUNCTION SetDefaultExtension (BYVAL pwszDefaultExtension AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION Close (BYVAL hr AS HRESULT) AS HRESULT
   DECLARE FUNCTION SetClientGuid (BYVAL guid AS GUID PTR) AS HRESULT
   DECLARE FUNCTION ClearClientData () AS HRESULT
   ' // IFileSaveDialog additioal methods
   DECLARE FUNCTION SetSaveAsItem (BYVAL psi AS IShellItem PTR) AS HRESULT
   DECLARE FUNCTION SetProperties (BYVAL pStore AS IPropertyStore PTR) AS HRESULT
   DECLARE FUNCTION SetCollectedProperties (BYVAL pList AS IPropertyDescriptionList PTR, BYVAL fAppendDefault AS WINBOOL) AS HRESULT
   DECLARE FUNCTION GetProperties (BYVAL ppStore AS IPropertyStore PTR PTR) AS HRESULT
   DECLARE FUNCTION ApplyProperties (BYVAL psi AS IShellItem PTR, BYVAL pStore AS IPropertyStore PTR, BYVAL hWnd AS HWND, BYVAL pSink AS IFileOperationProgressSink PTR) AS HRESULT

   ' // IFileDialogCustomize methods
   DECLARE FUNCTION AddCheckButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR, BYVAL bChecked AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION AddComboBox (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION AddControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION AddEditBox (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION AddMenu (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION AddPushButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION AddRadioButtonList (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION AddSeparator (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION AddText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION EnableOpenDropDown (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION EndVisualGroup () AS HRESULT
   DECLARE FUNCTION GetCheckButtonState (BYVAL dwIDCtl AS DWORD) AS BOOLEAN
   DECLARE FUNCTION GetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYREF dwState AS CDCONTROLSTATEF) AS HRESULT
   DECLARE FUNCTION GetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS CDCONTROLSTATEF
   DECLARE FUNCTION GetControlState (BYVAL dwIDCtl AS DWORD) AS CDCONTROLSTATEF
   DECLARE FUNCTION GetEditBoxText (BYVAL dwIDCtl AS DWORD) AS DWSTRING
   DECLARE FUNCTION GetSelectedControlItem (BYVAL dwIDCtl AS DWORD) AS DWSTRING
   DECLARE FUNCTION MakeProminent (BYVAL dwIDCtl AS DWORD) AS HRESULT
'   DECLARE FUNCTION RemoveAllControlItems (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION RemoveControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwItem AS DWORD) AS HRESULT
   DECLARE FUNCTION SetCheckButtonState (BYVAL dwIDCtl AS DWORD, BYVAL bCkecked AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION SetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   DECLARE FUNCTION SetControlItemText (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetControlLabel (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetControlState (BYVAL dwIDCtl AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   DECLARE FUNCTION SetEditBoxText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetSelectedControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   DECLARE FUNCTION StartVisualGroup (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
PRIVATE CONSTRUCTOR ISaveFileDialog
   ' // Initialize the COM library
   CoInitialize NULL
   ' // Create an instance of the FileOpenDialog interface
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_FileSaveDialog, NULL, CLSCTX_INPROC_SERVER, @IID_IFileSaveDialog, @m_IFileSaveDialog)
   OutputDebugStringW("ISaveFileDialog Default Constructor - hr: " & ..WSTR(hr))
   SetLastError(hr)
   ' // Get a pointer to the IFileDialogCustomize interface
   m_IFileSaveDialog->lpvtbl->QueryInterface(m_IFileSaveDialog, @IID_IFileDialogCustomize, @m_IFileDialogCustomize)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR ISaveFileDialog (BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   ' // Initialize the COM library
   CoInitialize NULL
   ' // Create an instance of the FileOpenDialog interface
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_FileSaveDialog, NULL, CLSCTX_INPROC_SERVER, @IID_IFileSaveDialog, @m_IFileSaveDialog)
   OutputDebugStringW("ISaveFileDialog Constructor(xPos, yPos) - hr: " & ..WSTR(hr))
   SetLastError(hr)
   ' // Open file dialog coordinates
   this.ClearClientData
   m_xPos = xPos
   m_yPos = yPos
   ' // Get a pointer to the IFileDialogCustomize interface
   m_IFileSaveDialog->lpvtbl->QueryInterface(m_IFileSaveDialog, @IID_IFileDialogCustomize, @m_IFileDialogCustomize)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR ISaveFileDialog
   OutputDebugStringW("ISaveFileDialog.Destructor")
   ' // Removes an event handler that was attached through the Advise method.
   IF m_dwDlgEvtCookie THEN this.Unadvise(m_pFileDialogEvents, m_dwDlgEvtCookie)
   ' // Release the IFileDialogCustomize interface
   IF m_IFileDialogCustomize THEN m_IFileDialogCustomize->lpvtbl->Release(m_IFileDialogCustomize)
   ' // Release the IFileSaveDialog interface
   m_IFileSaveDialog->lpvtbl->Release(m_IFileSaveDialog)
   ' // Uninitialize the COM library
   CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Display he save file dialog
' Usage example:
' DIM psfd AS ISaveFileDialog = ISaveFileDialog (50, 50)
' // Set the file types
' psfd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
' psfd.AddFileType("Executable files", "*.exe;*.dll")
' psfd.AddFileType("All files", "*.*")
' psfd.SetFileTypes()
' // Optional: Set the title of the dialog
'   psfd.SetTitle("Save File Dialog")
' // Set the folder
' psfd.SetFolder(CURDIR)
' psfd.SetDefaultExtension("bas")
' psfd.SetFileTypeIndex(1)
' psfd.AddCheckButton(1001, "My check button", TRUE)
' psfd.AddComboBox(1002)
' psfd.AddEditBox(1003, "My edit control")
' psfd.AddEditBox(1004, "My menu")
' psfd.AddPushButton(1005, "My push button")
' psfd.AddSeparator(1006)
' psfd.AddText(1007, "My text")
' psfd.EnableOpenDropDown(1008)
' psfd.StartVisualGroup(1009, "Group")
' psfd.AddControlItem(1010, 1009, "Control item")
' psfd.EndVisualGroup()
' print psfd.GetCheckButtonState(1001)
' print psfd.GetControlState(1001)
' // Display the dialog
' DIM hr AS HRESULT = psfd.ShowSave(hwnd)
' // Get the result
'IF hr = S_OK THEN
'   print psfd.GetResult()
'END IF
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.ShowSave (BYVAL hwndOwner AS HWND) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.ShowSave")
   ' // Optional: Create an event handling object, and hook it up to the dialog.
   DIM p AS ANY PTR = NEW CFileDialogEvents(m_xPos, m_yPos)
   m_pFileDialogEvents = p
   m_dwDlgEvtCookie = this.Advise(m_pFileDialogEvents)
   ' // Display the dialog
   RETURN m_IFileSaveDialog->lpvtbl->Show(m_IFileSaveDialog, hwndOwner)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the file type that appears as selected in the dialog.
' This method must be called before the dialog is showing.
' The index is one-based.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetFileTypeIndex (BYVAL iFileType AS UINT) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetFileTypeIndex")
   RETURN m_IFileSaveDialog->lpvtbl->SetFileTypeIndex(m_IFileSaveDialog, iFileType)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetFileTypeIndex () AS UINT
   OutputDebugStringW("ISaveFileDialog.GetFileTypeIndex")
   DIM iFileType AS UINT
   m_IFileSaveDialog->lpvtbl->GetFileTypeIndex(m_IFileSaveDialog, @iFileType)
   RETURN iFileType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an event handler that listens for events coming from the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.Advise (BYVAL pfde as IFileDialogEvents PTR) AS DWORD
   OutputDebugStringW("ISaveFileDialog.Advise")
   DIM dwCookie AS DWORD
   SetLastError(m_IFileSaveDialog->lpvtbl->Advise(m_IFileSaveDialog, pfde, @dwCookie))
   RETURN dwCookie
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an event handler that was attached through the Advise method.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.Unadvise (BYVAL pfde as IFileDialogEvents PTR, BYVAL dwCookie AS DWORD) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.Unadvise")
   RETURN m_IFileSaveDialog->lpvtbl->Unadvise(m_IFileSaveDialog, dwCookie)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets flags to control the behavior of the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetOptions (BYVAL fos AS FILEOPENDIALOGOPTIONS) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetOptions")
   RETURN m_IFileSaveDialog->lpvtbl->SetOptions(m_IFileSaveDialog, fos)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the current flags that are set to control dialog behavior.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetOptions () AS FILEOPENDIALOGOPTIONS
   OutputDebugStringW("ISaveFileDialog.GetOptions")
   DIM fos AS FILEOPENDIALOGOPTIONS
   SetLastError(m_IFileSaveDialog->lpvtbl->GetOptions(m_IFileSaveDialog, @fos))
   RETURN fos
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the folder used as a default if there is not a recently used folder value available.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetDefaultFolder (BYVAL psi AS IShellItem PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetDefaultFolder")
   RETURN m_IFileSaveDialog->lpvtbl->SetDefaultFolder(m_IFileSaveDialog, psi)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets a folder that is always selected when the dialog is opened, regardless of previous user action.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetFolder (BYVAL pwszFolderName AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetFolder")
   DIM pFolder AS IShellItem PTR
   DIM hr AS HRESULT = SHCreateItemFromParsingName(pwszFolderName, NULL, @IID_IShellItem, @pFolder)
   IF hr <> S_OK THEN RETURN hr
   hr = m_IFileSaveDialog->lpVtbl->SetFolder(m_IFileSaveDialog, pFolder)
   IF pFolder THEN pFolder->lpVtbl->Release(pFolder)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets either the folder currently selected in the dialog, or, if the dialog is not currently
' displayed, the folder that is to be selected when the dialog is opened.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetFolder (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   OutputDebugStringW("ISaveFileDialog.GetFolder")
   DIM psi AS IShellItem PTR, pwszName AS WSTRING PTR
   DIM hr AS HRESULT = m_IFileSaveDialog->lpvtbl->GetFolder(m_IFileSaveDialog, @psi)
   IF SUCCEEDED(hr) THEN
      hr = psi->lpVtbl->GetDisplayName(psi, sigdnName, @pwszName)
      psi->lpvtbl->Release(psi)
   END IF
   SetLastError hr
   RETURN pwszName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the folder used as a default if there is not a recently used folder value available.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetCurrentSelection (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   OutputDebugStringW("ISaveFileDialog.GetCurrentSelection")
   DIM psi AS IShellItem PTR, pwszName AS WSTRING PTR
   DIM hr AS HRESULT = m_IFileSaveDialog->lpvtbl->GetCurrentSelection(m_IFileSaveDialog, @psi)
   IF SUCCEEDED(hr) THEN
      hr = psi->lpVtbl->GetDisplayName(psi, sigdnName, @pwszName)
      psi->lpvtbl->Release(psi)
   END IF
   SetLastError hr
   RETURN pwszName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetFileName (BYVAL pwszName AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetFileName")
   RETURN m_IFileSaveDialog->lpvtbl->SetFileName(m_IFileSaveDialog, pwszName)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the text currently entered in the dialog's File name edit box.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetFileName () AS DWSTRING
   OutputDebugStringW("ISaveFileDialog.GetFileName")
   DIM pwszName AS WSTRING PTR
   SetLastError(m_IFileSaveDialog->lpvtbl->GetFileName(m_IFileSaveDialog, @pwszName))
   RETURN pwszName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the title of the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetTitle (BYVAL pwszTitle AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetTitle")
   RETURN m_IFileSaveDialog->lpvtbl->SetTitle(m_IFileSaveDialog, pwszTitle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the text of the Open or Save button.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetOkButtonLabel (BYVAL pwszText AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetOkButtonLabel")
   RETURN m_IFileSaveDialog->lpvtbl->SetOkButtonLabel(m_IFileSaveDialog, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the text of the label next to the file name edit box.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetFileNameLabel (BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetFileNameLabel")
   RETURN m_IFileSaveDialog->lpvtbl->SetFileNameLabel(m_IFileSaveDialog, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the choice that the user made in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetResult (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   OutputDebugStringW("ISaveFileDialog.GetResult")
   DIM psi AS IShellItem PTR
   DIM pwszName AS WSTRING PTR
   DIM hr As HRESULT = m_IFileSaveDialog->lpvtbl->GetResult(m_IFileSaveDialog, @psi)
   IF SUCCEEDED(hr) THEN
      hr = psi->lpVtbl->GetDisplayName(psi, sigdnName, @pwszName)
      psi->lpvtbl->Release(psi)
   END IF
   SetLastError hr
   RETURN pwszName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a folder to the list of places available for the user to open or save items.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddPlace (BYVAL psi AS IShellItem PTR, BYVAL fdap AS FDAP) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.AddPlace")
   RETURN m_IFileSaveDialog->lpvtbl->AddPlace(m_IFileSaveDialog, psi, fdap)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the default extension to be added to file names.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetDefaultExtension (BYVAL pwszDefaultExtension AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetDefaultExtension")
   RETURN m_IFileSaveDialog->lpvtbl->SetDefaultExtension(m_IFileSaveDialog, pwszDefaultExtension)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.Close (BYVAL hr AS HRESULT) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.Close")
   RETURN m_IFileSaveDialog->lpvtbl->Close(m_IFileSaveDialog, hr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables a calling application to associate a GUID with a dialog's persisted state.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetClientGuid (BYVAL guid AS GUID PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetClientGuid")
   RETURN m_IFileSaveDialog->lpvtbl->SetClientGuid(m_IFileSaveDialog, guid)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Instructs the dialog to clear all persisted state information.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.ClearClientData () AS HRESULT
   OutputDebugStringW("ISaveFileDialog.ClearClientData")
   RETURN m_IFileSaveDialog->lpvtbl->ClearClientData(m_IFileSaveDialog)
END FUNCTION
' ========================================================================================

' ########################################################################################
' IFileSaveDialog additioal methods
' ########################################################################################

' ========================================================================================
' Sets an item to be used as the initial entry in a Save As dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetSaveAsItem (BYVAL psi AS IShellItem PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetSaveAsItem")
   RETURN m_IFileSaveDialog->lpvtbl->SetSaveAsItem(m_IFileSaveDialog, psi)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides a property store that defines the default values to be used for the item being saved.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetProperties (BYVAL pStore AS IPropertyStore PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetProperties")
   RETURN m_IFileSaveDialog->lpvtbl->SetProperties(m_IFileSaveDialog, pStore)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Specifies which properties will be collected in the save dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetCollectedProperties (BYVAL pList AS IPropertyDescriptionList PTR, BYVAL fAppendDefault AS WINBOOL) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetCollectedProperties")
   RETURN m_IFileSaveDialog->lpvtbl->SetCollectedProperties(m_IFileSaveDialog, pList, fAppendDefault)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the set of property values for a saved item or an item in the process of being saved.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetProperties (BYVAL ppStore AS IPropertyStore PTR PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.GetProperties")
   RETURN m_IFileSaveDialog->lpvtbl->GetProperties(m_IFileSaveDialog, ppStore)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the set of property values for a saved item or an item in the process of being saved.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.ApplyProperties (BYVAL psi AS IShellItem PTR, BYVAL pStore AS IPropertyStore PTR, BYVAL hWnd AS HWND, BYVAL pSink AS IFileOperationProgressSink PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.ApplyProperties")
   RETURN m_IFileSaveDialog->lpvtbl->ApplyProperties(m_IFileSaveDialog, psi, pStore, hwnd, pSink)
END FUNCTION
' ========================================================================================

' ########################################################################################

' =====================================================================================
' * Adds a file type and pattern to the table.
' Usage example:
'   pofd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
'   pofd.AddFileType("Executable files", "*.exe;*.dll")
'   pofd.AddFileType("All files", "*.*")
'   pofd.SetFileTypes()
' =====================================================================================
PRIVATE SUB ISaveFileDialog.AddFileType (BYVAL pwszName AS WSTRING PTR, BYVAL pwszSpec AS WSTRING PTR)
   OutputDebugStringW("ISaveFileDialog.AddFileType")
   REDIM PRESERVE m_rgFilterSpec(UBOUND(m_rgFilterSpec) + 1) AS COMDLG_FILTERSPEC
   m_rgFilterSpec(UBOUND(m_rgFilterSpec)).pszName = pwszName
   m_rgFilterSpec(UBOUND(m_rgFilterSpec)).pszSpec = pwszSpec
END SUB
' =====================================================================================

' ========================================================================================
' * Sets the file types that the dialog can open or save.
' This overloaded function uses the m_rgFilterSpec array to set the file types.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetFileTypes () AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetFileTypes - individual")
   DIM cFileTypes AS LONG = UBOUND(m_rgFilterSpec) - LBOUND(m_rgFilterSpec) + 1
   IF cFileTYpes = 0 THEN RETURN E_UNEXPECTED
   RETURN m_IFileSaveDialog->lpvtbl->SetFileTypes(m_IFileSaveDialog, cFileTypes, @m_rgFilterSpec(0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the file types that the dialog can open or save.
' Usage example:
'   DIM rgFileTypes(1 TO 3) AS COMDLG_FILTERSPEC
'   rgFileTypes(1).pszName = @WSTR("FB code files")
'   rgFileTypes(2).pszName = @WSTR("Executable files")
'   rgFileTypes(3).pszName = @WSTR("All files")
'   rgFileTypes(1).pszSpec = @WSTR("*.bas;*.inc;*.bi")
'   rgFileTypes(2).pszSpec = @WSTR("*.exe;*.dll")
'   rgFileTypes(3).pszSpec = @WSTR("*.*")
'   pofd.SetFileTypes(3, @rgFileTypes(1))
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetFileTypes (BYVAL cFileTypes AS UINT, BYVAL rgFilterSpec AS COMDLG_FILTERSPEC PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.SetFileTypes - array")
   RETURN m_IFileSaveDialog->lpvtbl->SetFileTypes(m_IFileSaveDialog, cFileTypes, rgFilterSpec)
END FUNCTION
' ========================================================================================


' ########################################################################################
' IFileDialogCustomize methods
' ########################################################################################
' ========================================================================================
' Adds a check button (check box) to the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddCheckButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR, BYVAL bChecked AS BOOLEAN = FAlSE) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddCheckButton(m_IFileDialogCustomize, dwIDCtl, pwszLabel, bChecked)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a combo box to the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddComboBox (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddComboBox(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds an item to a container control in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddControlItem(m_IFileDialogCustomize, dwIDCtl, dwIDItem, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds an edit box control to the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddEditBox (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddEditBox(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a menu to the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddMenu (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddMenu(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a push button to the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddPushButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddPushButton(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a push button to the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddRadioButtonList (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddRadioButtonList(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a separator to the dialog, allowing a visual separation of controls.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddSeparator (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddSeparator(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds text content to the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.AddText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddText(m_IFileDialogCustomize, dwIDCtl, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables a drop-down list on the Open or Save button in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.EnableOpenDropDown (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->EnableOpenDropDown(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Stops the addition of elements to a visual group in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.EndVisualGroup () AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->EndVisualGroup(m_IFileDialogCustomize)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current state of a check button (check box) in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetCheckButtonState (BYVAL dwIDCtl AS DWORD) AS BOOLEAN
   DiM bChecked AS WINBOOL
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetCheckButtonState(m_IFileDialogCustomize, dwIDCtl, @bChecked))
   RETURN bChecked
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current state of an item in a container control found in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS CDCONTROLSTATEF
   DIM dwState AS CDCONTROLSTATEF
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetControlItemState(m_IFileDialogCustomize, dwIDCtl, dwIDItem, @dwState))
   RETURN dwState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current visibility and enabled states of a given control.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetControlState (BYVAL dwIDCtl AS DWORD) AS CDCONTROLSTATEF
   DIM dwState AS CDCONTROLSTATEF
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetControlState(m_IFileDialogCustomize, dwIDCtl, @dwState))
   RETURN dwState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current text in an edit box control.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.GetEditBoxText (BYVAL dwIDCtl AS DWORD) AS DWSTRING
   DIM pwszText AS WSTRING PTR
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetEditBoxText(m_IFileDialogCustomize, dwIDCtl, @pwszText))
   RETURN pwszText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Places a control in the dialog so that it stands out compared to other added controls.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.MakeProminent (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->MakeProminent(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not implemented.
' ========================================================================================
'PRIVATE FUNCTION ISaveFileDialog.RemoveAllControlItems (BYVAL dwIDCtl AS DWORD) AS HRESULT
'   RETURN m_IFileDialogCustomize->lpvtbl->RemoveAllControlItems(m_IFileDialogCustomize, dwIDCtl)
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an item from a container control in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.RemoveControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwItem AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->RemoveControlItem(m_IFileDialogCustomize, dwIDCtl, dwItem)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of a check button (check box) in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetCheckButtonState (BYVAL dwIDCtl AS DWORD, BYVAL bCkecked AS BOOLEAN) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetCheckButtonState(m_IFileDialogCustomize, dwIDCtl, bCkecked)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of a check button (check box) in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlItemState(m_IFileDialogCustomize, dwIDCtl, dwIDItem, dwState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of a control item. For example, the text that accompanies a radio button or an item in a menu.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetControlItemText (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlItemText(m_IFileDialogCustomize, dwIDCtl, dwIDItem, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text associated with a control, such as button text or an edit box label.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetControlLabel (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlLabel(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current visibility and enabled states of a given control.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetControlState (BYVAL dwIDCtl AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlState(m_IFileDialogCustomize, dwIDCtl, dwState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text in an edit box control found in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetEditBoxText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetEditBoxText(m_IFileDialogCustomize, dwIDCtl, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the selected state of a particular item in an option button group or a combo box found in the dialog.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.SetSelectedControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetSelectedControlItem(m_IFileDialogCustomize, dwIDCtl, dwIDItem)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Declares a visual group in the dialog. Subsequent calls to any "add" method add those elements to this group
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.StartVisualGroup (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddText(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================
