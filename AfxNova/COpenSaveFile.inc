' ########################################################################################
' Platform: Microsoft Windows
' Filename: COpenSaveFile.inc
' Contents: Open/Save file dialogs
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

NAMESPACE AfxNova

' ========================================================================================
' OpenFileDialog class
' ========================================================================================
TYPE OpenFileDialog

Private:
   m_hParent AS HWND                       ' // Handle of the parent window or dialog
   m_xPos AS LONG                          ' // Horizontal position, in pixels, relative to the parent window
   m_yPos AS LONG                          ' // Vertical position, in pixels, relative to the parent window
   m_Flags AS DWORD                        ' // The style attributes of the dialog
   m_FlagsEx AS DWORD                      ' // Extra style attributes of the dialog
   m_Index AS UINT                         ' // File type that appears as selected in the dialog
   m_wszTitle AS WSTRING * MAX_PATH        ' // The title to be displayed in the title bar of the dialog box
   m_wszInitiaDir AS WSTRING * MAX_PATH    ' // The name of the initial file directory to be displayed
   m_wszFileName AS WSTRING * MAX_PATH     ' // The file name to be used as the initial file selection
   m_wszDefaultExt AS WSTRING * MAX_PATH   ' // Default extension to be appended to the selected file name if the user does not enter it
   m_dwsFilter AS DWSTRING                 ' // Filter (pairs of strings separated with "|")
   m_FilterIndex AS DWORD                  ' // The file type that appears as selected in the dialog

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR 
   DECLARE FUNCTION ShowOpen (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING
   DECLARE PROPERTY Title () AS DWSTRING
   DECLARE PROPERTY Title (BYREF wszTitle AS WSTRING)
   DECLARE PROPERTY Flags () AS DWORD
   DECLARE PROPERTY Flags (BYVAL dwFlags AS DWORD)
   DECLARE PROPERTY FlagsEx () AS DWORD
   DECLARE PROPERTY FlagsEx (BYVAL dwFlagsEx AS DWORD)
   DECLARE PROPERTY DefaultExt () AS DWSTRING
   DECLARE PROPERTY DefaultExt (BYREF wszDefaultExt AS WSTRING)
   DECLARE PROPERTY InitialDir () AS DWSTRING
   DECLARE PROPERTY InitialDir (BYREF wszInitialDir AS WSTRING)
   DECLARE PROPERTY FilterIndex () AS DWORD
   DECLARE PROPERTY FilterIndex (BYVAL dwFilterIndex AS DWORD)
   DECLARE PROPERTY FileName () AS DWSTRING
   DECLARE PROPERTY FileName (BYREF wszFileName AS WSTRING)
   DECLARE PROPERTY Filter () AS DWSTRING
   DECLARE PROPERTY Filter (BYREF dwsFilter AS DWSTRING)

Private:
   DECLARE STATIC FUNCTION OFNHookProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS UINT_PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR OpenFileDialog
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR OpenFileDialog
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Open file dialog hook callback procedure
' ========================================================================================
FUNCTION OpenFileDialog.OFNHookProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS UINT_PTR
   DIM pOFN AS OPENFILENAMEW PTR
   DIM AS LONG xPos, yPos
   ' // Process the WM_INITDIALOG message
   IF uMsg = WM_INITDIALOG THEN
      ' // Get the custom data
      pOFN = CAST(OPENFILENAMEW PTR, lParam)
      IF pOFN THEN
         xPos = CSHORT(LOWORD(pOFN->lCustData))
         yPos = CSHORT(HIWORD(pOFN->lCustData))
      END IF
      ' // Note: hwnd is the handle of the child dialog box.
      ' // Use GetParent to get the handle of the Open dialog box.
      DIM hOpenFileDlg AS .HWND = GetParent(hwnd)
      ' // Get the parent of the open file dialog.
      ' // It can be another window or dialog, or the desktop.
      DIM hParent AS .HWND = GetParent(hOpenFileDlg)
      IF hParent = NULL THEN hParent = GetDesktopWindow
      ' // If the user has passed -1, -1 as the position
      IF (xPos = -1 ANDALSO yPos = -1) THEN
         ' // Center the open dialog on the parent window
         DIM AS RECT rc, rcDlg, rcOwner
         GetWindowRect(hParent, @rcOwner)
         GetWindowRect(hOpenFileDlg, @rcDlg)
         CopyRect(@rc, @rcOwner)
         ' // Offset the owner and dialog box rectangles so that right and bottom 
         ' // values represent the width and height, and then offset the owner again 
         ' // to discard space taken up by the dialog box. 
         OffsetRect(@rcDlg, -rcDlg.left, -rcDlg.top)
         OffsetRect(@rc, -rc.left, -rc.top)
         OffsetRect(@rc, -rcDlg.right, -rcDlg.bottom)
         ' // The new position is the sum of half the remaining space and the owner's 
         ' // original position. 
         SetWindowPos(hOpenFileDlg, HWND_TOP, rcOwner.left + (rc.right / 2), _
                      rcOwner.top + (rc.bottom / 2), 0, 0, SWP_NOSIZE)
      ELSE
'         ' // Get the dimensions of the bounding rectangle of the parent window
'         DIM rc AS RECT
'         GetWindowRect hParent, @rc
'         ' // Move the open dialog to the specified location, relative to the parent window
'         SetWindowPos hOpenFileDlg, 0, rc.Left + x, rc.Top + y, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
         ' // If the windows has been shown, GetWindowRect takes into account the drop shadows.
         ' // We must use the DwmGetWindowAttribute function to get the dimensions of the windows without the drop shadows;
         ' // otherwise, the open file dialog will be displaced to the left.
         DIM rc AS RECT
         DIM AS ANY PTR pLib = DyLibLoad("Dwmapi.DLL")
         IF pLib  THEN
            DIM pProc AS FUNCTION (BYVAL hwnd AS .HWND, BYVAL dwAttribute AS DWORD, BYVAL pvAttribute AS ANY PTR, BYVAL cbAttribute AS DWORD) AS HRESULT
            pProc = DyLibSymbol(pLib, "DwmGetWindowAttribute")
            ' DWMWA_EXTENDED_FRAME_BOUNDS = 9
            IF pProc THEN pProc(hParent, 9, @rc, SIZEOF(rc))
            DyLibFree(pLib)
         END IF
         ' // Move the open dialog to the specified location, relative to the parent window
         SetWindowPos hOpenFileDlg, 0, rc.Left + xPos, rc.Top + yPos, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
      END IF   
   END IF
   RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Display the dialog.
' Usage example:
'   DIM pOFN AS OpenFileDialog
'   pOFN.FileName = "*.*"
'   pOFN.InitialDir = CURDIR
'   pOFN.Filter = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
'   pOFN.DefaultExt = "BAS"
'   pOFN.FilterIndex = 1
'   'pOFN.Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY
'   pOFN.Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY OR OFN_ALLOWMULTISELECT
'   DIM dwsFile AS DWSTRING = pOFN.ShowOpen(hWin, 20, 20, TRUE)
'   --or-- DIM dwsFile AS DWSTRING = pOFN.ShowOpen(hWin, -1, -1, TRUE) to center the dialog
'   PRINT dwsFile
' ========================================================================================
PRIVATE FUNCTION OpenFileDialog.ShowOpen (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING

   ' // Filter is a sequence of WSTRINGs with a final (extra) double null terminator
   ' // The "|" characters are replaced with nulls
   DIM wszMarkers AS WSTRING * 4 = "||"
   IF RIGHT(m_dwsFilter, 1) <> "|" THEN wszMarkers += "|"
   m_dwsFilter = m_dwsFilter & wszMarkers
   DIM dwFilterStrSize AS DWORD = LEN(m_dwsFilter)
   ' // Replace markers("|") with nulls
   DIM pchar AS WCHAR PTR = *m_dwsFilter
   FOR i AS LONG = 0 TO LEN(m_dwsFilter) - 1
      IF pchar[i] = ASC("|") THEN pchar[i] = 0
   NEXT

   ' // If the initial directory has not been specified, assume the current directory
   IF LEN(m_wszInitiaDir) = 0 THEN m_wszInitiaDir = CURDIR
   ' // The size of the buffer must be at least MAX_PATH characters
   DIM dwBufLen AS DWORD = MAX_PATH   ' // Make room for at least one path
   IF (m_Flags AND OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT THEN dwBufLen = 32768  ' // 64 Kb buffer
   ' // Allocate the file name and a marker ("|") to be replaced with a null
   DIM dwsFileName AS DWSTRING = m_wszFileName & "|"
   ' // Store the position of the marker
   DIM cbPos AS LONG = LEN(dwsFileName) - 1
   ' // Allocate room for the buffer
   IF LEN(dwsFileName) < dwBufLen THEN dwsFileName += WSPACE(dwBufLen - LEN(dwsFileName))
   DIM dwFileStrSize AS DWORD = LEN(dwsFileName)
   ' // The filename must be null terminated (replace the marker with a null)
   pchar = *dwsFileName
   pchar[cbPos] = 0

   ' // Fill the members of the OPENFILENAMEW structure
   DIM ofn AS OPENFILENAMEW
   ofn.lStructSize = SIZEOF(ofn)
   ofn.hwndOwner = hParent
   ofn.lpstrFilter = *m_dwsFilter
   IF m_FilterIndex < 1 THEN m_FilterIndex = 1
   ofn.nFilterIndex = m_FilterIndex
   ofn.lpstrFile = *dwsFileName
   ofn.nMaxFile = dwFileStrSize
   IF m_wszInitiaDir = "" THEN m_wszInitiaDir = CURDIR
   ofn.lpstrInitialDir = @m_wszInitiaDir
   IF LEN(m_wszTitle) THEN ofn.lpstrTitle = @m_wszTitle
   IF LEN(m_wszDefaultExt) THEN ofn.lpstrDefExt = @m_wszDefaultExt
   ofn.lCustData = MAKELONG(xPos, yPos)
   IF bHook THEN
      ' // Set a hook
      m_Flags = m_Flags OR OFN_EXPLORER OR OFN_ENABLEHOOK
      ofn.Flags = m_Flags
      ofn.lpfnHook = @OpenFileDialog.OFNHookProc
   END IF
   ' // Call the open file dialog
   IF GetOpenFilenameW(@ofn) THEN
      IF (m_Flags AND OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT THEN
         pchar = *dwsFileName
         FOR i AS LONG = 0 TO dwFileStrSize - 1
            ' // If double null, exit
            IF pchar[i] = 0 AND pchar[i + 1] = 0 THEN EXIT FOR
            ' // Replace null with ","
            IF pchar[i] = 0 THEN pchar[i] = ASC(",")
         NEXT
      END IF
      ' // Trim trailing spaces
      dwsFileName = RTRIM(dwsFileName, CHR(32))
      IF RIGHT(dwsFileName, 1) = "," THEN dwsFileName = LEFT(dwsFileName, LEN(dwsFileName) - 1)
   ELSE
      ' // Buffer too small
      IF CommDlgExtendedError = FNERR_BUFFERTOOSMALL THEN
         dwBufLen = ASC(dwsFileName)   ' // needed buffer lenght
      END IF
      dwsFileName = ""
   END IF

   ' // Return the selected file or files
   RETURN dwsFileName

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the title of the dialog.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.Title () AS DWSTRING
   RETURN m_wszTitle
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the title of the dialog.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.Title (BYREF wszTitle AS WSTRING)
   m_wszTitle = wszTitle
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the flags that control the behavior of the dialog.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.Flags () AS DWORD
   RETURN m_Flags
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.FlagsEx () AS DWORD
   RETURN m_FlagsEx
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the flags that control the behavior of the dialog.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.Flags (BYVAL dwFlags AS DWORD)
   m_Flags = dwFlags
END PROPERTY
' ========================================================================================
' ========================================================================================
' Currently, the only option is OFN_EX_NOPLACESBAR
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.FlagsEx (BYVAL dwFlagsEx AS DWORD)
   m_FlagsEx = dwFlagsEx
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.DefaultExt () AS DWSTRING
   RETURN m_wszDefaultExt
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.DefaultExt (BYREF wszDefaultExt AS WSTRING)
   m_wszDefaultExt = wszDefaultExt
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the folder used as a the initial directory
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.InitialDir () AS DWSTRING
   RETURN m_wszInitiaDir
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.InitialDir (BYREF wszInitialDir AS WSTRING)
   m_wszInitiaDir = wszInitialDir
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.FilterIndex () AS DWORD
   RETURN m_FilterIndex
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.FilterIndex (BYVAL dwFilterIndex AS DWORD)
   m_FilterIndex = dwFilterIndex
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.FileName () AS DWSTRING
   RETURN m_wszFileName
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.FileName (BYREF wszFileName AS WSTRING)
   m_wszFileName = wszFileName
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the file types that the dialog can open or save.
' Filter (pairs of strings separated with "|")
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.Filter () AS DWSTRING
   RETURN m_dwsFilter
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file types that the dialog can open or save.
' Filter (pairs of strings separated with "|")
' ========================================================================================
PRIVATE PROPERTY OpenFileDialog.Filter (BYREF dwsFilter AS DWSTRING)
   m_dwsFilter = dwsFilter
END PROPERTY
' ========================================================================================


' ########################################################################################

' ========================================================================================
' SaveFileDialog class
' ========================================================================================
TYPE SaveFileDialog

Private:
   m_hParent AS HWND                       ' // Handle of the parent window or dialog
   m_xPos AS LONG                          ' // Horizontal position, in pixels, relative to the parent window
   m_yPos AS LONG                          ' // Vertical position, in pixels, relative to the parent window
   m_Flags AS DWORD                        ' // The style attributes of the dialog
   m_FlagsEx AS DWORD                      ' // Extra style attributes of the dialog
   m_Index AS UINT                         ' // File type that appears as selected in the dialog
   m_wszTitle AS WSTRING * MAX_PATH        ' // The title to be displayed in the title bar of the dialog box
   m_wszInitiaDir AS WSTRING * MAX_PATH    ' // The name of the initial file directory to be displayed
   m_wszFileName AS WSTRING * MAX_PATH     ' // The file name to be used as the initial file selection
   m_wszDefaultExt AS WSTRING * MAX_PATH   ' // Default extension to be appended to the selected file name if the user does not enter it
   m_dwsFilter AS DWSTRING                 ' // Filter (pairs of strings separated with "|")
   m_FilterIndex AS DWORD                  ' // The file type that appears as selected in the dialog

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR 
   DECLARE FUNCTION ShowSave (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING
   DECLARE PROPERTY Title () AS DWSTRING
   DECLARE PROPERTY Title (BYREF wszTitle AS WSTRING)
   DECLARE PROPERTY Flags () AS DWORD
   DECLARE PROPERTY Flags (BYVAL dwFlags AS DWORD)
   DECLARE PROPERTY FlagsEx () AS DWORD
   DECLARE PROPERTY FlagsEx (BYVAL dwFlagsEx AS DWORD)
   DECLARE PROPERTY DefaultExt () AS DWSTRING
   DECLARE PROPERTY DefaultExt (BYREF wszDefaultExt AS WSTRING)
   DECLARE PROPERTY InitialDir () AS DWSTRING
   DECLARE PROPERTY InitialDir (BYREF wszInitialDir AS WSTRING)
   DECLARE PROPERTY FilterIndex () AS DWORD
   DECLARE PROPERTY FilterIndex (BYVAL dwFilterIndex AS DWORD)
   DECLARE PROPERTY FileName () AS DWSTRING
   DECLARE PROPERTY FileName (BYREF wszFileName AS WSTRING)
   DECLARE PROPERTY Filter () AS DWSTRING
   DECLARE PROPERTY Filter (BYREF dwsFilter AS DWSTRING)

Private:
   DECLARE STATIC FUNCTION OFNHookProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS UINT_PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR SaveFileDialog
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR SaveFileDialog
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Open file dialog hook callback procedure
' ========================================================================================
FUNCTION SaveFileDialog.OFNHookProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS UINT_PTR
   DIM pOFN AS OPENFILENAMEW PTR
   DIM AS LONG xPos, yPos
   ' // Process the WM_INITDIALOG message
   IF uMsg = WM_INITDIALOG THEN
      ' // Get the custom data
      pOFN = CAST(OPENFILENAMEW PTR, lParam)
      IF pOFN THEN
         xPos = CSHORT(LOWORD(pOFN->lCustData))
         yPos = CSHORT(HIWORD(pOFN->lCustData))
      END IF
      ' // Note: hwnd is the handle of the child dialog box.
      ' // Use GetParent to get the handle of the Open dialog box.
      DIM hOpenFileDlg AS .HWND = GetParent(hwnd)
      ' // Get the parent of the open file dialog.
      ' // It can be another window or dialog, or the desktop.
      DIM hParent AS .HWND = GetParent(hOpenFileDlg)
      IF hParent = NULL THEN hParent = GetDesktopWindow
      ' // If the user has passed -1, -1 as the position
      IF (xPos = -1 ANDALSO yPos = -1) THEN
         ' // Center the open dialog on the parent window
         DIM AS RECT rc, rcDlg, rcOwner
         GetWindowRect(hParent, @rcOwner)
         GetWindowRect(hOpenFileDlg, @rcDlg)
         CopyRect(@rc, @rcOwner)
         ' // Offset the owner and dialog box rectangles so that right and bottom 
         ' // values represent the width and height, and then offset the owner again 
         ' // to discard space taken up by the dialog box. 
         OffsetRect(@rcDlg, -rcDlg.left, -rcDlg.top)
         OffsetRect(@rc, -rc.left, -rc.top)
         OffsetRect(@rc, -rcDlg.right, -rcDlg.bottom)
         ' // The new position is the sum of half the remaining space and the owner's 
         ' // original position. 
         SetWindowPos(hOpenFileDlg, HWND_TOP, rcOwner.left + (rc.right / 2), _
                      rcOwner.top + (rc.bottom / 2), 0, 0, SWP_NOSIZE)
      ELSE
'         ' // Get the dimensions of the bounding rectangle of the parent window
'         DIM rc AS RECT
'         GetWindowRect hParent, @rc
'         ' // Move the open dialog to the specified location, relative to the parent window
'         SetWindowPos hOpenFileDlg, 0, rc.Left + x, rc.Top + y, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
         ' // If the windows has been shown, GetWindowRect takes into account the drop shadows.
         ' // We must use the DwmGetWindowAttribute function to get the dimensions of the windows without the drop shadows;
         ' // otherwise, the open file dialog will be displaced to the left.
         DIM rc AS RECT
         DIM AS ANY PTR pLib = DyLibLoad("Dwmapi.DLL")
         IF pLib  THEN
            DIM pProc AS FUNCTION (BYVAL hwnd AS .HWND, BYVAL dwAttribute AS DWORD, BYVAL pvAttribute AS ANY PTR, BYVAL cbAttribute AS DWORD) AS HRESULT
            pProc = DyLibSymbol(pLib, "DwmGetWindowAttribute")
            ' DWMWA_EXTENDED_FRAME_BOUNDS = 9
            IF pProc THEN pProc(hParent, 9, @rc, SIZEOF(rc))
            DyLibFree(pLib)
         END IF
         ' // Move the open dialog to the specified location, relative to the parent window
         SetWindowPos hOpenFileDlg, 0, rc.Left + xPos, rc.Top + yPos, 0, 0, SWP_NOSIZE OR SWP_NOACTIVATE OR SWP_NOZORDER
      END IF   
   END IF
   RETURN 0
END FUNCTION
' ========================================================================================

' ========================================================================================
' Display the dialog.
' Usage example:
'   DIM pSFN AS SaveFileDialog
'   pSFN.FileName = "*.*"
'   pSFN.InitialDir = CURDIR
'   pSFN.Filter = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
'   pSFN.DefaultExt = "BAS"
'   pSFN.FilterIndex = 1
'   pSFN.Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY
'   DIM dwsFile AS DWSTRING = pSFN.ShowSave(hWin, 20, 20, TRUE)
'   --or-- DIM dwsFile AS DWSTRING = pSFN.ShowSave(hWin, -1, -1, TRUE) to center the dialog
'   PRINT dwsFile
' ========================================================================================
PRIVATE FUNCTION SaveFileDialog.ShowSave (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING

   ' // Filter is a sequence of WSTRINGs with a final (extra) double null terminator
   ' // The "|" characters are replaced with nulls
   DIM wszMarkers AS WSTRING * 4 = "||"
   IF RIGHT(m_dwsFilter, 1) <> "|" THEN wszMarkers += "|"
   m_dwsFilter = m_dwsFilter & wszMarkers
   DIM dwFilterStrSize AS DWORD = LEN(m_dwsFilter)
   ' // Replace markers("|") with nulls
   DIM pchar AS WCHAR PTR = *m_dwsFilter
   FOR i AS LONG = 0 TO LEN(m_dwsFilter) - 1
      IF pchar[i] = ASC("|") THEN pchar[i] = 0
   NEXT

   ' // If the initial directory has not been specified, assume the current directory
   IF LEN(m_wszInitiaDir) = 0 THEN m_wszInitiaDir = CURDIR
   ' // Allocate the file name and a marker ("|") to be replaced with a null
   DIM dwsFileName AS DWSTRING = m_wszFileName & "|"
   ' // Store the position of the marker
   DIM cbPos AS LONG = LEN(dwsFileName) - 1
   ' // Allocate room for the buffer
   IF LEN(dwsFileName) < MAX_PATH THEN dwsFileName += WSPACE(MAX_PATH - LEN(dwsFileName))
   DIM dwFileStrSize AS DWORD = LEN(dwsFileName)
   ' // The filename must be null terminated (replace the marker with a null)
   pchar = *dwsFileName
   pchar[cbPos] = 0

   ' // Fill the members of the OPENFILENAMEW structure
   DIM ofn AS OPENFILENAMEW
   ofn.lStructSize = SIZEOF(ofn)
   ofn.hwndOwner = hParent
   ofn.lpstrFilter = *m_dwsFilter
   IF m_FilterIndex < 1 THEN m_FilterIndex = 1
   ofn.nFilterIndex = m_FilterIndex
   ofn.lpstrFile = *dwsFileName
   ofn.nMaxFile = dwFileStrSize
   IF m_wszInitiaDir = "" THEN m_wszInitiaDir = CURDIR
   ofn.lpstrInitialDir = @m_wszInitiaDir
   IF LEN(m_wszTitle) THEN ofn.lpstrTitle = @m_wszTitle
   IF LEN(m_wszDefaultExt) THEN ofn.lpstrDefExt = @m_wszDefaultExt
   ofn.lCustData = MAKELONG(xPos, yPos)
   IF bHook THEN
      ' // Set a hook
      m_Flags = m_Flags OR OFN_EXPLORER OR OFN_ENABLEHOOK
      ofn.Flags = m_Flags
      ofn.lpfnHook = @SaveFileDialog.OFNHookProc
   END IF

   ' // Call the save filename dialog
   IF GetSaveFilenameW(@ofn) = 0 THEN dwsFileName = ""

   ' // Return the selected file or files
   RETURN dwsFileName

END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the title of the dialog.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.Title () AS DWSTRING
   RETURN m_wszTitle
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the title of the dialog.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.Title (BYREF wszTitle AS WSTRING)
   m_wszTitle = wszTitle
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the flags that control the behavior of the dialog.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.Flags () AS DWORD
   RETURN m_Flags
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.FlagsEx () AS DWORD
   RETURN m_FlagsEx
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the flags that control the behavior of the dialog.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.Flags (BYVAL dwFlags AS DWORD)
   m_Flags = dwFlags
END PROPERTY
' ========================================================================================
' ========================================================================================
' Currently, the only option is OFN_EX_NOPLACESBAR
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.FlagsEx (BYVAL dwFlagsEx AS DWORD)
   m_FlagsEx = dwFlagsEx
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.DefaultExt () AS DWSTRING
   RETURN m_wszDefaultExt
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.DefaultExt (BYREF wszDefaultExt AS WSTRING)
   m_wszDefaultExt = wszDefaultExt
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the folder used as a the initial directory
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.InitialDir () AS DWSTRING
   RETURN m_wszInitiaDir
END PROPERTY
' ========================================================================================
' ========================================================================================
' Gets the default extension to be added to file names.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.InitialDir (BYREF wszInitialDir AS WSTRING)
   m_wszInitiaDir = wszInitialDir
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.FilterIndex () AS DWORD
   RETURN m_FilterIndex
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file type that appears as selected in the dialog.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.FilterIndex (BYVAL dwFilterIndex AS DWORD)
   m_FilterIndex = dwFilterIndex
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.FileName () AS DWSTRING
   RETURN m_wszFileName
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file name that appears in the File name edit box when that dialog box is opened.
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.FileName (BYREF wszFileName AS WSTRING)
   m_wszFileName = wszFileName
END PROPERTY
' ========================================================================================

' ========================================================================================
' Sets the file types that the dialog can open or save.
' Filter (pairs of strings separated with "|")
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.Filter () AS DWSTRING
   RETURN m_dwsFilter
END PROPERTY
' ========================================================================================
' ========================================================================================
' Sets the file types that the dialog can open or save.
' Filter (pairs of strings separated with "|")
' ========================================================================================
PRIVATE PROPERTY SaveFileDialog.Filter (BYREF dwsFilter AS DWSTRING)
   m_dwsFilter = dwsFilter
END PROPERTY
' ========================================================================================

END NAMESPACE
