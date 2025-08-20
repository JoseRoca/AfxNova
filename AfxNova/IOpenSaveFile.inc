' ########################################################################################
' Microsoft Windows
' File: IOpenSaveFile.inc
' Contents: Open/Save dialogs
' Compiler: FreeBASIC 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "AfxNova/AfxWin.inc"
USING AfxNova

#if _WIN32_WINNT >= &h0600

NAMESPACE AfxNova

' ########################################################################################
' CFileDialogEvents class
' Implementation of the FileDialogEvents callback interface
' ########################################################################################
TYPE CFileDialogEvents EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION OnFileOk (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnFolderChanging (BYVAL pfd AS IFileDialog PTR, BYVAL psiFolder AS IShellItem PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnSelectionChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnShareViolation (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_SHAREVIOLATION_RESPONSE PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnTypeChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnOverwrite (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_OVERWRITE_RESPONSE PTR) AS HRESULT

   DECLARE CONSTRUCTOR (BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount As ULONG
   ' // Open file dialog handle
   m_hOleWindow AS HWND
   ' // Open file dialog coordinates
   m_xPos AS LONG
   m_yPos AS LONG

END TYPE
' ########################################################################################

' ########################################################################################
' Template example to set a IFileDialogEvents innterface.
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CFileDialogEvents (BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   ' // Open file dialog coordinates
   m_xPos = xPos
   m_yPos = yPos
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CFileDialogEvents
END DESTRUCTOR
' =====================================================================================

' IUnknown interface methods
' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CFileDialogEvents.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   OutputDebugStringW("CFileDialogEvents.QueryInterface")
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_IFileOpenDialog) OR _
      IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
      ' // Increment the reference count
      this.AddRef
      RETURN S_OK
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' =====================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
PRIVATE FUNCTION CFileDialogEvents.AddRef () AS ULONG
   refCount += 1
   OutputDebugStringW("CFileDialogEvents.AddRef " & ..WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CFileDialogEvents.Release () AS ULONG
   refCount -= 1
   OutputDebugStringW("CFileDialogEvents.Release " & ..WSTR(refCount))
   IF refCount = 0 THEN
      OutputDebugStringW("CFileDialogEvents - Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' Event Handlers
' =====================================================================================
' Called just before the dialog is about to return with a result.
' When this method is called, the IFileDialog::GetResult and GetResults methods can be called.
' =====================================================================================
PRIVATE FUNCTION CFileDialogEvents.OnFileOk (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   OutputDebugStringW("CFileDialogEvents.OnFileOk")
   RETURN S_OK
END FUNCTION
' =====================================================================================
' =====================================================================================
' Called before OnFolderChange. This allows the implementer to stop navigation to a particular location.
' A pointer to an interface that represents the folder to which the dialog is about to navigate.
' =====================================================================================
PRIVATE FUNCTION CFileDialogEvents.OnFolderChanging (BYVAL pfd AS IFileDialog PTR, BYVAL psiFolder AS IShellItem PTR) AS HRESULT
   OutputDebugStringW("CFileDialogEvents.OnFolderChanging")
   RETURN S_OK
END FUNCTION
' =====================================================================================
' =====================================================================================
' Called when the user navigates to a new folder.
' OnFolderChange is called when the dialog is opened.
' =====================================================================================
PRIVATE FUNCTION CFileDialogEvents.OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   OutputDebugStringW("CFileDialogEvents.OnFolderChange")

'   IF m_hOleWindow = NULL THEN
      ' // Before displaying the dialog get its coordinates
      DIM pOleWindow AS IOleWindow PTR
      ' // Get a pointer to the IOleWindow interface
      pfd->lpvtbl->QueryInterface(pfd, @IID_IOleWindow, @pOleWindow)
      ' // Get the window handle of the dialog
      IF pOleWindow THEN pOleWindow->lpvtbl->GetWindow(pOleWindow, @m_hOleWindow)
      OutputDebugStringW("CFileDialogEvents.OnFolderChange - hOleWindow: " & ..WSTR(m_hOleWindow))
      ' // Get he bounding rectangle of the parent window
      DIM AS RECT rcDlg, rcOwner
      GetWindowRect(GetParent(m_hOleWindow), @rcOwner)
      ' // Get he bounding rectangle of the open file dialog
      GetWindowRect(m_hOleWindow, @rcDlg)
      ' // Maps the open file dialog coordinates relative to its parent window
      MapWindowPoints(GetParent(m_hOleWindow), m_hOleWindow, CAST(POINT PTR, @rcDlg), 2)
'     ' // Moves the position of the open file dialog
'      IF (m_xPos = -1 ANDALSO m_yPos = -1) THEN
'         ' // Calculate the width and height of the parent window
'         DIM nParentWidth AS LONG = rcOwner.Right - rcOwner.Left
'         DIM nParentHeight AS LONG = rcOwner.Bottom - rcOwner.Top
'         ' // Get the width and height of the open file dialog
'         DIM nWidth AS LONG = rcDlg.Right - rcDlg.Left
'         DIM nHeight AS LONG = rcDlg.Bottom - rcDlg.Top
'         ' // Get the height of the non-client area of the dialog
'         DIM rc AS RECT
'         ' // Calculate the new x coordinate
'         DIM pt AS POINT
'         pt.x = rcOwner.Left + ((nParentWidth - nWidth) \ 2)
'         ' // Can't calculate the y coordinate because GetWindowRect returns a higher
'         ' // value for the open file dialog that the visible one
'        pt.y = rcOwner.Top + (nParentHeight \ 2)
''         pt.y = rcOwner.Top + ((nParentHeight - nHeight) \ 2)
'         pt.y = rcOwner.Top
'         SetWindowPos(m_hOleWindow, HWND_TOP, pt.x, pt.y, 0, 0, SWP_NOSIZE)
'      ELSE
'         SetWindowPos m_hOleWindow, 0, rcDlg.Left + m_xPos, rcDlg.Top + m_yPos, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
'      END IF
'   END IF
   SetWindowPos m_hOleWindow, 0, rcDlg.Left + m_xPos, rcDlg.Top + m_yPos, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
   RETURN S_OK
END FUNCTION
' =====================================================================================
' =====================================================================================
' Called when the user changes the selection in the dialog's view.
' =====================================================================================
PRIVATE FUNCTION CFileDialogEvents.OnSelectionChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   OutputDebugStringW("CFileDialogEvents.OnSelectionChange")
   RETURN S_OK
END FUNCTION
' =====================================================================================
' =====================================================================================
' Enables an application to respond to sharing violations that arise from Open or Save operations.
' =====================================================================================
PRIVATE FUNCTION CFileDialogEvents.OnShareViolation (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_SHAREVIOLATION_RESPONSE PTR) AS HRESULT
   OutputDebugStringW("CFileDialogEvents.OnShareViolation")
   RETURN S_OK
END FUNCTION
' =====================================================================================
' =====================================================================================
' Called when the dialog is opened to notify the application of the initial chosen filetype.
' =====================================================================================
PRIVATE FUNCTION CFileDialogEvents.OnTypeChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   OutputDebugStringW("CFileDialogEvents.OnTypeChange")
   RETURN S_OK
END FUNCTION
' =====================================================================================
' =====================================================================================
' Called from the save dialog when the user chooses to overwrite a file.
' =====================================================================================
PRIVATE FUNCTION CFileDialogEvents.OnOverwrite (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_OVERWRITE_RESPONSE PTR) AS HRESULT
   OutputDebugStringW("CFileDialogEvents.OnOverwrite")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' ########################################################################################

' ########################################################################################
' Exposes methods that initialize, show, and get results from the common open file dialog.
' ########################################################################################

' ========================================================================================
' IOpenFileDialog class
' ========================================================================================
TYPE IOpenFileDialog

Private:
   m_IFileOpenDialog AS IFileOpenDialog PTR             ' // Pointer to the IFileOpenDialog interface
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
   DECLARE FUNCTION ShowOpen (BYVAL hwndOwner AS HWND) AS HRESULT
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
   DECLARE FUNCTION GetResults () AS IShellItemArray PTR
   DECLARE FUNCTION GetResultsString (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   DECLARE FUNCTION GetSelectedItems () AS IShellItemArray PTR
   DECLARE FUNCTION GetResultsCount () AS LONG
   DECLARE FUNCTION ParseResults (BYREF wszMainStr AS CONST WSTRING, BYVAL index AS LONG) AS DWSTRING

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
PRIVATE CONSTRUCTOR IOpenFileDialog
   ' // Initialize the COM library
   CoInitialize NULL
   ' // Create an instance of the FileOpenDialog interface
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, @IID_IFileOpenDialog, @m_IFileOpenDialog)
   OutputDebugStringW("IOpenFileDialog Default Constructor - hr: " & ..WSTR(hr))
   SetLastError(hr)
   ' // Get a pointer to the IFileDialogCustomize interface
   m_IFileOpenDialog->lpvtbl->QueryInterface(m_IFileOpenDialog, @IID_IFileDialogCustomize, @m_IFileDialogCustomize)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR IOpenFileDialog (BYVAL xPos AS LONG, BYVAL yPos AS LONG)
   ' // Initialize the COM library
   CoInitialize NULL
   ' // Create an instance of the FileOpenDialog interface
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, @IID_IFileOpenDialog, @m_IFileOpenDialog)
   OutputDebugStringW("IOpenFileDialog Constructor(xPos, yPos) - hr: " & ..WSTR(hr))
   SetLastError(hr)
   ' // Open file dialog coordinates
   this.ClearClientData
   m_xPos = xPos
   m_yPos = yPos
   ' // Get a pointer to the IFileDialogCustomize interface
   m_IFileOpenDialog->lpvtbl->QueryInterface(m_IFileOpenDialog, @IID_IFileDialogCustomize, @m_IFileDialogCustomize)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR IOpenFileDialog
   OutputDebugStringW("IOpenFileDialog.Destructor")
   ' // Removes an event handler that was attached through the Advise method.
   IF m_dwDlgEvtCookie THEN this.Unadvise(m_pFileDialogEvents, m_dwDlgEvtCookie)
   ' // Release the IFileDialogCustomize interface
   IF m_IFileDialogCustomize THEN m_IFileDialogCustomize->lpvtbl->Release(m_IFileDialogCustomize)
   ' // Release the IFileOpenDialog interface
   m_IFileOpenDialog->lpvtbl->Release(m_IFileOpenDialog)
   ' // Uninitialize the COM library
   CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Display he open file dialog
' Usage example:
'   DIM pofd AS IOpenFileDialog
'   ' // Set the file types
'   pofd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
'   pofd.AddFileType("Executable files", "*.exe;*.dll")
'   pofd.AddFileType("All files", "*.*")
'   pofd.SetFileTypes()
'   ' // Multiple selection (default is single selection)
'   DIM options AS FILEOPENDIALOGOPTIONS = pofd.GetOptions
'   pofd.SetOptions(options OR FOS_ALLOWMULTISELECT)
'   ' // Optional: Set the title of the dialog
''   pofd.SetTitle("A Single-Selection Dialog")
'   ' // Set the folder
'   pofd.SetFolder(CURDIR)
'   pofd.SetDefaultExtension("bas")
'   pofd.SetFileTypeIndex(1)
'   ' // Display the dialog
'   DIM hr AS HRESULT = pofd.ShowOpen(hDlg)
'   ' *** Single selection ***
'   ' // Get the result
'   IF hr = S_OK THEN
'      print pofd.GetResult()
'   END IF
'   ' *** Multiple selection ***
'   ' // Get the results
'   print "Folder: "; pofd.GetFolder
'   DIM pItems AS IShellItemArray PTR
'   pItems = pofd.GetResults
'   IF pItems THEN
'      DIM dwItemCount AS LONG
'      pItems->lpVtbl->GetResultsCount(pItems, @dwItemCount)
'      FOR idx AS LONG = 0 TO dwItemCount - 1
'         DIM pItem AS IShellItem PTR
'         pItems->lpVtbl->GetItemAt(pItems, idx, @pItem)
'         IF pItem THEN
'            DIM pwszName AS WSTRING PTR
'            pItem->lpVtbl->GetDisplayName(pItem, SIGDN_NORMALDISPLAY, @pwszName)
'            IF pwszName THEN
'               print *pwszName
'               CoTaskMemFree(pwszName)
'               pwszName = NULL
'            END IF
'            pItem->lpVtbl->Release(pItem)
'            pItem = NULL
'         END IF
'      NEXT
'      pItems->lpVtbl->Release(pItems)
'   END IF
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.ShowOpen (BYVAL hwndOwner AS HWND) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.ShowOpen")
   ' // Optional: Create an event handling object, and hook it up to the dialog.
   DIM p AS ANY PTR = NEW CFileDialogEvents(m_xPos, m_yPos)
   m_pFileDialogEvents = p
   m_dwDlgEvtCookie = this.Advise(m_pFileDialogEvents)
   ' // Display the dialog
   RETURN m_IFileOpenDialog->lpvtbl->Show(m_IFileOpenDialog, hwndOwner)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the file type that appears as selected in the dialog.
' This method must be called before the dialog is showing.
' The index is one-based.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetFileTypeIndex (BYVAL iFileType AS UINT) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetFileTypeIndex")
   RETURN m_IFileOpenDialog->lpvtbl->SetFileTypeIndex(m_IFileOpenDialog, iFileType)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetFileTypeIndex () AS UINT
   OutputDebugStringW("IOpenFileDialog.GetFileTypeIndex")
   DIM iFileType AS UINT
   m_IFileOpenDialog->lpvtbl->GetFileTypeIndex(m_IFileOpenDialog, @iFileType)
   RETURN iFileType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assigns an event handler that listens for events coming from the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.Advise (BYVAL pfde as IFileDialogEvents PTR) AS DWORD
   OutputDebugStringW("IOpenFileDialog.Advise")
   DIM dwCookie AS DWORD
   SetLastError(m_IFileOpenDialog->lpvtbl->Advise(m_IFileOpenDialog, pfde, @dwCookie))
   RETURN dwCookie
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an event handler that was attached through the Advise method.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.Unadvise (BYVAL pfde as IFileDialogEvents PTR, BYVAL dwCookie AS DWORD) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.Unadvise")
   RETURN m_IFileOpenDialog->lpvtbl->Unadvise(m_IFileOpenDialog, dwCookie)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets flags to control the behavior of the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetOptions (BYVAL fos AS FILEOPENDIALOGOPTIONS) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetOptions")
   RETURN m_IFileOpenDialog->lpvtbl->SetOptions(m_IFileOpenDialog, fos)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the current flags that are set to control dialog behavior.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetOptions () AS FILEOPENDIALOGOPTIONS
   OutputDebugStringW("IOpenFileDialog.GetOptions")
   DIM fos AS FILEOPENDIALOGOPTIONS
   SetLastError(m_IFileOpenDialog->lpvtbl->GetOptions(m_IFileOpenDialog, @fos))
   RETURN fos
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the folder used as a default if there is not a recently used folder value available.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetDefaultFolder (BYVAL psi AS IShellItem PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetDefaultFolder")
   RETURN m_IFileOpenDialog->lpvtbl->SetDefaultFolder(m_IFileOpenDialog, psi)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets a folder that is always selected when the dialog is opened, regardless of previous user action.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetFolder (BYVAL pwszFolderName AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetFolder")
   DIM pFolder AS IShellItem PTR
   DIM hr AS HRESULT = SHCreateItemFromParsingName(pwszFolderName, NULL, @IID_IShellItem, @pFolder)
   IF hr <> S_OK THEN RETURN hr
   hr = m_IFileOpenDialog->lpVtbl->SetFolder(m_IFileOpenDialog, pFolder)
   IF pFolder THEN pFolder->lpVtbl->Release(pFolder)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets either the folder currently selected in the dialog, or, if the dialog is not currently
' displayed, the folder that is to be selected when the dialog is opened.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetFolder (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   OutputDebugStringW("IOpenFileDialog.GetFolder")
   DIM psi AS IShellItem PTR, pwszName AS WSTRING PTR
   DIM hr AS HRESULT = m_IFileOpenDialog->lpvtbl->GetFolder(m_IFileOpenDialog, @psi)
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
PRIVATE FUNCTION IOpenFileDialog.GetCurrentSelection (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   OutputDebugStringW("IOpenFileDialog.GetCurrentSelection")
   DIM psi AS IShellItem PTR, pwszName AS WSTRING PTR
   DIM hr AS HRESULT = m_IFileOpenDialog->lpvtbl->GetCurrentSelection(m_IFileOpenDialog, @psi)
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
PRIVATE FUNCTION IOpenFileDialog.SetFileName (BYVAL pwszName AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetFileName")
   RETURN m_IFileOpenDialog->lpvtbl->SetFileName(m_IFileOpenDialog, pwszName)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Retrieves the text currently entered in the dialog's File name edit box.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetFileName () AS DWSTRING
   OutputDebugStringW("IOpenFileDialog.GetFileName")
   DIM pwszName AS WSTRING PTR
   SetLastError(m_IFileOpenDialog->lpvtbl->GetFileName(m_IFileOpenDialog, @pwszName))
   RETURN pwszName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the title of the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetTitle (BYVAL pwszTitle AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetTitle")
   RETURN m_IFileOpenDialog->lpvtbl->SetTitle(m_IFileOpenDialog, pwszTitle)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the text of the Open or Save button.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetOkButtonLabel (BYVAL pwszText AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetOkButtonLabel")
   RETURN m_IFileOpenDialog->lpvtbl->SetOkButtonLabel(m_IFileOpenDialog, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the text of the label next to the file name edit box.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetFileNameLabel (BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetFileNameLabel")
   RETURN m_IFileOpenDialog->lpvtbl->SetFileNameLabel(m_IFileOpenDialog, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the choice that the user made in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetResult (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   OutputDebugStringW("IOpenFileDialog.GetResult")
   DIM psi AS IShellItem PTR
   DIM pwszName AS WSTRING PTR
   DIM hr As HRESULT = m_IFileOpenDialog->lpvtbl->GetResult(m_IFileOpenDialog, @psi)
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
PRIVATE FUNCTION IOpenFileDialog.AddPlace (BYVAL psi AS IShellItem PTR, BYVAL fdap AS FDAP) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.AddPlace")
   RETURN m_IFileOpenDialog->lpvtbl->AddPlace(m_IFileOpenDialog, psi, fdap)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the default extension to be added to file names.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetDefaultExtension (BYVAL pwszDefaultExtension AS WSTRING PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetDefaultExtension")
   RETURN m_IFileOpenDialog->lpvtbl->SetDefaultExtension(m_IFileOpenDialog, pwszDefaultExtension)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.Close (BYVAL hr AS HRESULT) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.Close")
   RETURN m_IFileOpenDialog->lpvtbl->Close(m_IFileOpenDialog, hr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables a calling application to associate a GUID with a dialog's persisted state.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetClientGuid (BYVAL guid AS GUID PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetClientGuid")
   RETURN m_IFileOpenDialog->lpvtbl->SetClientGuid(m_IFileOpenDialog, guid)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Instructs the dialog to clear all persisted state information.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.ClearClientData () AS HRESULT
   OutputDebugStringW("IOpenFileDialog.ClearClientData")
   RETURN m_IFileOpenDialog->lpvtbl->ClearClientData(m_IFileOpenDialog)
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Gets the files selected.
'   DIM pofd AS IOpenFileDialog
'   ' // Set the file types
'   pofd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
'   pofd.AddFileType("Executable files", "*.exe;*.dll")
'   pofd.AddFileType("All files", "*.*")
'   pofd.SetFileTypes()
'   ' // Multiple selection (default is single selection)
'   DIM options AS FILEOPENDIALOGOPTIONS = pofd.GetOptions
'   pofd.SetOptions(options OR FOS_ALLOWMULTISELECT)
'   ' // Optional: Set the title of the dialog
''   pofd.SetTitle("A Single-Selection Dialog")
'   ' // Set the folder
'   pofd.SetFolder(CURDIR)
'   pofd.SetDefaultExtension("bas")
'   pofd.SetFileTypeIndex(1)
'   ' // Display the dialog
'   DIM hr AS HRESULT = pofd.ShowOpen(hDlg)
'   ' // Folder name
'   PRINT "Folder: "; pofd.GetFolder
'   ' *** Single selection ***
'   ' // Get the result
''   IF hr = S_OK THEN
''      print pofd.GetResult()
''   END IF
'   ' *** Multiple selection ***
'   ' // Get the results
'   print "Folder: "; pofd.GetFolder
'   DIM pItems AS IShellItemArray PTR
'   pItems = pofd.GetResults
'   IF pItems THEN
'      DIM dwItemCount AS LONG
'      pItems->lpVtbl->GetResultsCount(pItems, @dwItemCount)
'      FOR idx AS LONG = 0 TO dwItemCount - 1
'         DIM pItem AS IShellItem PTR
'         pItems->lpVtbl->GetItemAt(pItems, idx, @pItem)
'         IF pItem THEN
'            DIM pwszName AS WSTRING PTR
'            pItem->lpVtbl->GetDisplayName(pItem, SIGDN_NORMALDISPLAY, @pwszName)
'            IF pwszName THEN
'               print *pwszName
'               CoTaskMemFree(pwszName)
'               pwszName = NULL
'            END IF
'            pItem->lpVtbl->Release(pItem)
'            pItem = NULL
'         END IF
'      NEXT
'      pItems->lpVtbl->Release(pItems)
'   END IF
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetResults () AS IShellItemArray PTR
   OutputDebugStringW("IOpenFileDialog.GetResults")
   DIM psai AS IShellItemArray PTR
   SetLastError(m_IFileOpenDialog->lpvtbl->GetResults(m_IFileOpenDialog, @psai))
   RETURN psai
END FUNCTION
' ========================================================================================
' ========================================================================================
' * Gets the files selected and returns then in a comma separated string.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetResultsString (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
   OutputDebugStringW("IOpenFileDialog.GetResultsString")
   DIM pItems AS IShellItemArray PTR
   DIM hr AS HRESULT = m_IFileOpenDialog->lpvtbl->GetResults(m_IFileOpenDialog, @pItems)
   IF hr <> S_OK THEN RETURN ""
   DIM dwItemCount AS LONG, dwsResults AS DWSTRING
   IF pItems THEN
      pItems->lpVtbl->GetCount(pItems, @dwItemCount)
      FOR idx AS LONG = 0 TO dwItemCount - 1
         DIM pItem AS IShellItem PTR
         pItems->lpVtbl->GetItemAt(pItems, idx, @pItem)
         IF pItem THEN
            DIM pwszName AS WSTRING PTR
            pItem->lpVtbl->GetDisplayName(pItem, sigdnName, @pwszName)
            IF pwszName THEN
               dwsResults += *pwszName
               IF idx < dwItemCount - 1 THEN dwsResults += ","
               CoTaskMemFree(pwszName)
               pwszName = NULL
            END IF
            pItem->lpVtbl->Release(pItem)
            pItem = NULL
         END IF
      NEXT
      pItems->lpVtbl->Release(pItems)
   END IF
   RETURN dwsResults
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the currently selected items in the dialog. These items may be items selected in
' the view, or text selected in the file name edit box.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetSelectedItems () AS IShellItemArray PTR
   OutputDebugStringW("IOpenFileDialog.GetSelectedItems")
   DIM psai AS IShellItemArray PTR
   SetLastError(m_IFileOpenDialog->lpvtbl->GetSelectedItems(m_IFileOpenDialog, @psai))
   RETURN psai
END FUNCTION
' ========================================================================================

' =====================================================================================
' * Adds a file type and pattern to the table.
' Usage example:
'   pofd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
'   pofd.AddFileType("Executable files", "*.exe;*.dll")
'   pofd.AddFileType("All files", "*.*")
'   pofd.SetFileTypes()
' =====================================================================================
PRIVATE SUB IOpenFileDialog.AddFileType (BYVAL pwszName AS WSTRING PTR, BYVAL pwszSpec AS WSTRING PTR)
   OutputDebugStringW("IOpenFileDialog.AddFileType")
   REDIM PRESERVE m_rgFilterSpec(UBOUND(m_rgFilterSpec) + 1) AS COMDLG_FILTERSPEC
   m_rgFilterSpec(UBOUND(m_rgFilterSpec)).pszName = pwszName
   m_rgFilterSpec(UBOUND(m_rgFilterSpec)).pszSpec = pwszSpec
END SUB
' =====================================================================================

' ========================================================================================
' * Sets the file types that the dialog can open or save.
' This overloaded function uses the m_rgFilterSpec array to set the file types.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetFileTypes () AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetFileTypes - individual")
   DIM cFileTypes AS LONG = UBOUND(m_rgFilterSpec) - LBOUND(m_rgFilterSpec) + 1
   IF cFileTYpes = 0 THEN RETURN E_UNEXPECTED
   RETURN m_IFileOpenDialog->lpvtbl->SetFileTypes(m_IFileOpenDialog, cFileTypes, @m_rgFilterSpec(0))
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
PRIVATE FUNCTION IOpenFileDialog.SetFileTypes (BYVAL cFileTypes AS UINT, BYVAL rgFilterSpec AS COMDLG_FILTERSPEC PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.SetFileTypes - array")
   RETURN m_IFileOpenDialog->lpvtbl->SetFileTypes(m_IFileOpenDialog, cFileTypes, rgFilterSpec)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of files selected
' Usage example:
' DIM pofd AS IOpenFileDialog
' ...
' DIM dwsRes AS DWSTRING = pofd.GetResultsString
' DIM nCount AS LONG = podf.GetResultsCount
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetResultsCount () AS LONG
   OutputDebugStringW("IOpenFileDialog.GetResultsCount")
   DIM pItems AS IShellItemArray PTR
   DIM dwItemCount AS LONG
   SetLastError(m_IFileOpenDialog->lpvtbl->GetResults(m_IFileOpenDialog, @pItems))
   IF pItems THEN
      pItems->lpVtbl->GetCount(pItems, @dwItemCount)
      pItems->lpVtbl->Release(pItems)
   END IF
   RETURN dwItemCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Parses the results.
' Usage example:
' DIM pofd AS IOpenFileDialog
' ...
' DIM dwsRes AS DWSTRING = pofd.GetResultsString
' FOR i AS LONG = 1 TO pofd.GetResultsCount
'    PRINT pofd.ParseResults(dwsRes, i)
' NEXT
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.ParseResults (BYREF wszMainStr AS CONST WSTRING, BYVAL index AS LONG) AS DWSTRING
   DIM dws AS DWSTRING = ""
   IF index < 0 THEN RETURN dws
   DIM nCount AS LONG, nStart AS LONG
   DIM nPosition AS LONG = Index
   DIM nPos AS LONG = 1
   DO
      nStart = nPos
      nPos = INSTR(nPos, wszMainStr, ",")
      IF nPos THEN
         nCount += 1
         nPos += 1
      END IF
   LOOP UNTIL nPos = 0 OR nCount = nPosition
   IF nPos > 0 OR nCount = nPosition - 1 THEN
      IF nPos = 0 THEN
         dws = MID(wszMainStr, nStart)
      ELSE
         dws = MID(wszMainStr, nStart, nPos - 1 - nStart)
      END IF
   END IF
   RETURN dws
END FUNCTION
' ========================================================================================

' ########################################################################################
' IFileDialogCustomize methods
' ########################################################################################
' ========================================================================================
' Adds a check button (check box) to the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.AddCheckButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR, BYVAL bChecked AS BOOLEAN = FAlSE) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddCheckButton(m_IFileDialogCustomize, dwIDCtl, pwszLabel, bChecked)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a combo box to the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.AddComboBox (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddComboBox(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds an item to a container control in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.AddControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddControlItem(m_IFileDialogCustomize, dwIDCtl, dwIDItem, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds an edit box control to the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.AddEditBox (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddEditBox(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a menu to the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.AddMenu (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddMenu(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a push button to the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.AddPushButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddPushButton(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a push button to the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.AddRadioButtonList (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddRadioButtonList(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a separator to the dialog, allowing a visual separation of controls.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.AddSeparator (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddSeparator(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds text content to the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.AddText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddText(m_IFileDialogCustomize, dwIDCtl, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables a drop-down list on the Open or Save button in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.EnableOpenDropDown (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->EnableOpenDropDown(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Stops the addition of elements to a visual group in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.EndVisualGroup () AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->EndVisualGroup(m_IFileDialogCustomize)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current state of a check button (check box) in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetCheckButtonState (BYVAL dwIDCtl AS DWORD) AS BOOLEAN
   DiM bChecked AS WINBOOL
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetCheckButtonState(m_IFileDialogCustomize, dwIDCtl, @bChecked))
   RETURN bChecked
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current state of an item in a container control found in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS CDCONTROLSTATEF
   DIM dwState AS CDCONTROLSTATEF
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetControlItemState(m_IFileDialogCustomize, dwIDCtl, dwIDItem, @dwState))
   RETURN dwState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current visibility and enabled states of a given control.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetControlState (BYVAL dwIDCtl AS DWORD) AS CDCONTROLSTATEF
   DIM dwState AS CDCONTROLSTATEF
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetControlState(m_IFileDialogCustomize, dwIDCtl, @dwState))
   RETURN dwState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current text in an edit box control.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.GetEditBoxText (BYVAL dwIDCtl AS DWORD) AS DWSTRING
   DIM pwszText AS WSTRING PTR
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetEditBoxText(m_IFileDialogCustomize, dwIDCtl, @pwszText))
   RETURN pwszText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Places a control in the dialog so that it stands out compared to other added controls.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.MakeProminent (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->MakeProminent(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not implemented.
' ========================================================================================
'PRIVATE FUNCTION IOpenFileDialog.RemoveAllControlItems (BYVAL dwIDCtl AS DWORD) AS HRESULT
'   RETURN m_IFileDialogCustomize->lpvtbl->RemoveAllControlItems(m_IFileDialogCustomize, dwIDCtl)
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an item from a container control in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.RemoveControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwItem AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->RemoveControlItem(m_IFileDialogCustomize, dwIDCtl, dwItem)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of a check button (check box) in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetCheckButtonState (BYVAL dwIDCtl AS DWORD, BYVAL bCkecked AS BOOLEAN) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetCheckButtonState(m_IFileDialogCustomize, dwIDCtl, bCkecked)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of a check button (check box) in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlItemState(m_IFileDialogCustomize, dwIDCtl, dwIDItem, dwState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of a control item. For example, the text that accompanies a radio button or an item in a menu.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetControlItemText (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlItemText(m_IFileDialogCustomize, dwIDCtl, dwIDItem, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text associated with a control, such as button text or an edit box label.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetControlLabel (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlLabel(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current visibility and enabled states of a given control.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetControlState (BYVAL dwIDCtl AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlState(m_IFileDialogCustomize, dwIDCtl, dwState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text in an edit box control found in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetEditBoxText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetEditBoxText(m_IFileDialogCustomize, dwIDCtl, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the selected state of a particular item in an option button group or a combo box found in the dialog.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.SetSelectedControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetSelectedControlItem(m_IFileDialogCustomize, dwIDCtl, dwIDItem)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Declares a visual group in the dialog. Subsequent calls to any "add" method add those elements to this group
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.StartVisualGroup (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddText(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ########################################################################################

END NAMESPACE

#endif
