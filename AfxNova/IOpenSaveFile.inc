' ########################################################################################
' Platform: Microsoft Windows
' File: CIOpenDaveFile.inc
' Contents: Implementation of IOpenFileDialog and ISaveFileDialog classes.
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2025 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#pragma once
#include once "AfxNova/AfxWin.inc"
USING AfxNova

#if _WIN32_WINNT >= &h0600

NAMESPACE AfxNova

' ========================================================================================
' IOpenFileDialog class
' Exposes methods that initialize, show, and get results from the common open file dialog.
' ========================================================================================
TYPE IOpenFileDialog

Private:
   m_IFileOpenDialog AS IFileOpenDialog PTR             ' // Pointer to the IFileOpenDialog interface
   m_hwndOwner AS HWND                                  ' // Handle of the owner of the dialog
   m_dwsTitle AS DWSTRING                               ' // The title to be displayed in the title bar of the dialog box
   m_dwsFolder AS DWSTRING                              ' // The name of the initial file directory to be displayed
   m_rgFilterSpec (ANY) AS COMDLG_FILTERSPEC            ' // Array of COMDLG_FILTERSPEC structures
   m_startFile AS DWSTRING                              ' // File name to be used as the initial file selection
   m_defExt AS DWSTRING                                 ' // Default extension to be appended to the selected file name if the user does not enter it
   m_Flags AS DWORD                                     ' // Flags
   m_pFileDialogEvents AS IFileDialogEvents PTR         ' // Pointer to the IFileDialogEvents interface
   m_dwDlgEvtCookie AS DWORD                            ' // Dialog events cookie

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR 
   DECLARE FUNCTION ShowOpen (BYVAL hwndOwner AS HWND) AS HRESULT
   DECLARE SUB AddFileType (BYVAL pwszName AS WSTRING PTR, BYVAL pwszSpec AS WSTRING PTR)
   DECLARE FUNCTION SetFileTypes () AS HRESULT
   DECLARE FUNCTION SetFileTypes (BYVAL cFileTypes AS UINT, BYVAL rgFilterSpec AS COMDLG_FILTERSPEC PTR) AS HRESULT
   DECLARE FUNCTION SetFileTypeIndex (BYVAL iFileType AS UINT) AS HRESULT
   DECLARE FUNCTION GetFileTypeIndex () AS UINT
   DECLARE FUNCTION Advise (BYVAL pfde as IFileDialogEvents PTR) AS HRESULT
   DECLARE FUNCTION Unadvise () AS HRESULT
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
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR IOpenFileDialog
   OutputDebugStringW("IOpenFileDialog.Destructor")
   ' // Removes an event handler that was attached through the Advise method.
   IF m_dwDlgEvtCookie THEN m_IFileOpenDialog->lpvtbl->Unadvise(m_IFileOpenDialog, m_dwDlgEvtCookie)
   ' // Release the IFileDialogEvents interface
   IF m_pFileDialogEvents THEN m_pFileDialogEvents->lpvtbl->Release(m_pFileDialogEvents)
   ' // Release the IFileOpenDialog interface
   m_IFileOpenDialog->lpvtbl->Release(m_IFileOpenDialog)
   ' // Uninitialize the COM library
   CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Displays the open file dialog
' Usage example:
' DIM pofd AS IOpenFileDialog = IOpenFileDialog(50, 50)
' // Set the file types
' pofd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
' pofd.AddFileType("Executable files", "*.exe;*.dll")
' pofd.AddFileType("All files", "*.*")
' pofd.SetFileTypes()
' // Multiple selection (default is single selection)
' DIM options AS FILEOPENDIALOGOPTIONS = pofd.GetOptions
' pofd.SetOptions(options OR FOS_ALLOWMULTISELECT)
' // Optional: Set the title of the dialog
'   pofd.SetTitle("A Single-Selection Dialog")
' // Set the folder
' pofd.SetFolder(CURDIR)
' pofd.SetDefaultExtension("bas")
' pofd.SetFileTypeIndex(1)
' // Display the dialog
' DIM hr AS HRESULT = pofd.ShowOpen(hwnd)
' // Folder name
' PRINT "Folder: "; pofd.GetFolder
' *** Single selection ***
' // Get the result
' IF hr = S_OK THEN
'    print pofd.GetResult()
' END IF
' *** Multiple selection *** 
' DIM dwsRes AS DWSTRING = pofd.GetResultsString
' FOR i AS LONG = 1 TO pofd.GetResultsCount
'    PRINT pofd.ParseResults(dwsRes, i)
' NEXT
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.ShowOpen (BYVAL hwndOwner AS HWND) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.ShowOpen")
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
' To set the events use:
' DIM pofd AS IOpenFileDialog
' DIM pfde AS ANY PTR = NEW CIFileDialogEvents
' pofd.Advise(pfde)
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.Advise (BYVAL pfde as IFileDialogEvents PTR) AS HRESULT
   OutputDebugStringW("IOpenFileDialog.Advise")
   RETURN m_IFileOpenDialog->lpvtbl->Advise(m_IFileOpenDialog, pfde, @m_dwDlgEvtCookie)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an event handler that was attached through the Advise method.
' ========================================================================================
PRIVATE FUNCTION IOpenFileDialog.Unadvise () AS HRESULT
   OutputDebugStringW("IOpenFileDialog.Unadvise")
   IF m_dwDlgEvtCookie THEN RETURN m_IFileOpenDialog->lpvtbl->Unadvise(m_IFileOpenDialog, m_dwDlgEvtCookie)
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


' ========================================================================================
' ISaveFileDialog class
' ========================================================================================
TYPE ISaveFileDialog

Private:
   m_IFileSaveDialog AS IFileSaveDialog PTR             ' // Pointer to the IFileSaveDialog interface
   m_hwndOwner AS HWND                                  ' // Handle of the owner of the dialog
   m_dwsTitle AS DWSTRING                               ' // The title to be displayed in the title bar of the dialog box
   m_dwsFolder AS DWSTRING                              ' // The name of the initial file directory to be displayed
   m_rgFilterSpec (ANY) AS COMDLG_FILTERSPEC            ' // Array of COMDLG_FILTERSPEC structures
   m_startFile AS DWSTRING                              ' // File name to be used as the initial file selection
   m_defExt AS DWSTRING                                 ' // Default extension to be appended to the selected file name if the user does not enter it
   m_Flags AS DWORD                                     ' // Flags
   m_pFileDialogEvents AS IFileDialogEvents PTR         ' // Pointer to the IFileDialogEvents interface
   m_dwDlgEvtCookie AS DWORD                            ' // Dialog events cookie

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR 
   DECLARE FUNCTION ShowSave (BYVAL hwndOwner AS HWND) AS HRESULT
   DECLARE SUB AddFileType (BYVAL pwszName AS WSTRING PTR, BYVAL pwszSpec AS WSTRING PTR)
   DECLARE FUNCTION SetFileTypes () AS HRESULT
   DECLARE FUNCTION SetFileTypes (BYVAL cFileTypes AS UINT, BYVAL rgFilterSpec AS COMDLG_FILTERSPEC PTR) AS HRESULT
   DECLARE FUNCTION SetFileTypeIndex (BYVAL iFileType AS UINT) AS HRESULT
   DECLARE FUNCTION GetFileTypeIndex () AS UINT
   DECLARE FUNCTION Advise (BYVAL pfde as IFileDialogEvents PTR) AS HRESULT
   DECLARE FUNCTION Unadvise () AS HRESULT
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
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR ISaveFileDialog
   OutputDebugStringW("ISaveFileDialog.Destructor")
   ' // Removes an event handler that was attached through the Advise method.
   IF m_dwDlgEvtCookie THEN m_IFileSaveDialog->lpvtbl->Unadvise(m_IFileSaveDialog, m_dwDlgEvtCookie)
   ' // Release the IFileDialogEvents interface
   IF m_pFileDialogEvents THEN m_pFileDialogEvents->lpvtbl->Release(m_pFileDialogEvents)
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
' // Display the dialog
' DIM hr AS HRESULT = psfd.ShowSave(hwnd)
' // Get the result
'IF hr = S_OK THEN
'   print psfd.GetResult()
'END IF
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.ShowSave (BYVAL hwndOwner AS HWND) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.ShowSave")
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
' To set the events use:
' DIM psfd AS ISaveFileDialog
' DIM pfde AS ANY PTR = NEW CIFileDialogEvents
' psfd.Advise(pfde)
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.Advise (BYVAL pfde as IFileDialogEvents PTR) AS HRESULT
   OutputDebugStringW("ISaveFileDialog.Advise")
   RETURN m_IFileSaveDialog->lpvtbl->Advise(m_IFileSaveDialog, pfde, @m_dwDlgEvtCookie)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an event handler that was attached through the Advise method.
' ========================================================================================
PRIVATE FUNCTION ISaveFileDialog.Unadvise () AS HRESULT
   OutputDebugStringW("ISaveFileDialog.Unadvise")
   IF m_dwDlgEvtCookie THEN RETURN m_IFileSaveDialog->lpvtbl->Unadvise(m_IFileSaveDialog, m_dwDlgEvtCookie)
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

END NAMESPACE

#endif
