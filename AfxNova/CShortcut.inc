' ########################################################################################
' Platform: Microsoft Windows
' Filename: CShorcut.inc
' Contents: Windows shortcuts
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
#include once "windows.bi"
#include once "AfxNova/AfxWsh.bi"
#include once "AfxNova/AfxCOM.inc"
#include once "AfxNova/DWSTRING.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' CShortcut class
' ========================================================================================
TYPE CShortcut

   Private:
      m_Result AS HRESULT              ' // HRESULT code
      m_bUninitCOM AS BOOLEAN
      m_pLnk AS Afx_IWshShortcut PTR   ' // IWshShortcut interface

   Public:
      DECLARE CONSTRUCTOR (BYREF wszPathName AS WSTRING)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS LONG = -1) AS DWSTRING
      DECLARE PROPERTY FullName () AS DWSTRING
      DECLARE PROPERTY Arguments () AS DWSTRING
      DECLARE PROPERTY Arguments (BYREF wszArguments AS WSTRING)
      DECLARE PROPERTY Description () AS DWSTRING
      DECLARE PROPERTY Description (BYREF wszDescription AS WSTRING)
      DECLARE PROPERTY Hotkey () AS DWSTRING
      DECLARE PROPERTY Hotkey (BYREF wszHotkey AS WSTRING)
      DECLARE PROPERTY IconLocation () AS DWSTRING
      DECLARE PROPERTY IconLocation (BYREF wszIconLocation AS WSTRING)
      DECLARE PROPERTY RelativePath (BYREF wszRelativePath AS WSTRING)
      DECLARE PROPERTY TargetPath () AS DWSTRING
      DECLARE PROPERTY TargetPath (BYREF wszTargetPath AS WSTRING)
      DECLARE PROPERTY WindowStyle () AS LONG
      DECLARE PROPERTY WindowStyle (BYVAL nstyle AS LONG)
      DECLARE PROPERTY WorkingDirectory () AS DWSTRING
      DECLARE PROPERTY WorkingDirectory (BYREF wszWorkingDirectory AS WSTRING)
      DECLARE FUNCTION Save () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CURLShortcut class
' ========================================================================================
TYPE CURLShortcut

   Private:
      m_Result AS HRESULT                 ' // HRESULT code
      m_bUninitCOM AS BOOLEAN
      m_pLnk AS Afx_IWshURLShortcut PTR   ' // IWshURLShortcut interface

   Public:
      DECLARE CONSTRUCTOR (BYREF wszPathName AS WSTRING)
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS LONG = -1) AS DWSTRING
      DECLARE PROPERTY FullName () AS DWSTRING
      DECLARE PROPERTY TargetPath () AS DWSTRING
      DECLARE PROPERTY TargetPath (BYREF wszTargetPath AS WSTRING)
      DECLARE FUNCTION Save () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CShortcut constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CShortcut (BYREF wszPathName AS WSTRING)

   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE

   ' // Create an instance of the IWshShortcut interface
   DIM pWsh AS Afx_IWshShell PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell_ AS GUID = (&hF935DC21, &h1CF0, &h11D0, {&hAD, &hB9, &h00, &hC0, &h4F, &hD5, &h8A, &h0B})
   m_Result = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell_, @pWsh)
   IF pWsh = NULL THEN m_Result = E_NOINTERFACE
   IF pWsh THEN
      DIM bstrPathName AS AFX_BSTR = SysAllocString(wszPathName)
      m_Result = pWsh->CreateShortcut(bstrPathName, cast(Afx_IDIspatch PTR PTR, @m_pLnk))
      SysFreeString bstrPathName
      pWsh->Release
   END IF
   IF pWsh = NULL THEN EXIT CONSTRUCTOR

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CShortcut destructor
' ========================================================================================
PRIVATE DESTRUCTOR CShortcut
   ' // Release the interface
   IF m_pLnk THEN m_pLnk->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CShortcut.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CShortcut.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   IF nError = -1 THEN nError = m_Result
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      dwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   IF nError THEN dwsMsg = "Error &h" + HEX(nError) & CHR(13, 10) + dwsMsg
   RETURN dwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the fully qualified path to the shortcut's target.
' ========================================================================================
PRIVATE PROPERTY CShortcut.FullName () AS DWSTRING
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrFullName AS AFX_BSTR
   m_Result = m_pLnk->get_FullName(@bstrFullName)
   DIM dwsFullName AS DWSTRING = bstrFullName
   SysFreeString bstrFullName
   PROPERTY = dwsFullName
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets/sets the arguments for a shortcut, or identifies a shortcut's arguments.
' ========================================================================================
PRIVATE PROPERTY CShortcut.Arguments () AS DWSTRING
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrArguments AS AFX_BSTR
   m_Result = m_pLnk->get_Arguments(@bstrArguments)
   DIM dwsArguments AS DWSTRING = bstrArguments
   SysFreeString bstrArguments
   PROPERTY = dwsArguments
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.Arguments (BYREF wszArguments AS WSTRING)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrArguments AS AFX_BSTR = SysAllocString(wszArguments)
   m_Result = m_pLnk->put_Arguments(bstrArguments)
   SysFreeString bstrArguments
END PROPERTY
' ========================================================================================

' ========================================================================================
' Gets/sets the shortcut description.
' ========================================================================================
PRIVATE PROPERTY CShortcut.Description () AS DWSTRING
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrDescription AS AFX_BSTR
   m_Result = m_pLnk->get_Description(@bstrDescription)
   DIM dwsDescription AS DWSTRING = bstrDescription
   SysFreeString bstrDescription
   PROPERTY = dwsDescription
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.Description (BYREF wszDescription AS WSTRING)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrDescription AS AFX_BSTR = bstrDescription
   m_Result = m_pLnk->put_Description(bstrDescription)
   SysFreestring bstrDescription
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a key-combination to a shortcut, or identifies the key-combination assigned to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.Hotkey () AS DWSTRING
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrHotkey AS AFX_BSTR
   m_Result = m_pLnk->get_Hotkey(@bstrHotkey)
   DIM dwsHotKey AS DWSTRING = bstrHotkey
   SysFreeString bstrHotkey
   PROPERTY = dwsHotKey
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.Hotkey (BYREF wszHotkey AS WSTRING)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrHotKey AS AFX_BSTR = SysAllocString(wszHotkey)
   m_Result = m_pLnk->put_Hotkey(bstrHotKey)
   SysFreeString bstrHotKey
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.IconLocation () AS DWSTRING
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrIconLocation AS AFX_BSTR
   m_Result = m_pLnk->get_IconLocation(@bstrIconLocation)
   DIM dwsIconLocation AS DWSTRING = bstrIconLocation
   PROPERTY = dwsIconLocation
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.IconLocation (BYREF wszIconLocation AS WSTRING)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrIconLocation AS AFX_BSTR = SysAllocString(wszIconLocation)
   m_Result = m_pLnk->put_IconLocation(bstrIconLocation)
   SysFreestring bstrIconLocation
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a relative path to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.RelativePath (BYREF wszRelativePath AS WSTRING)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrRelativePath AS AFX_BSTR = SysAllocString(wszRelativePath)
   m_Result = m_pLnk->put_RelativePath(bstrRelativePath)
   SysFreeString bstrRelativePath
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.TargetPath () AS DWSTRING
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrTargetPath AS AFX_BSTR
   m_Result = m_pLnk->get_TargetPath(@bstrTargetPath)
   DIM dwsTargetPath AS DWSTRING = bstrTargetPath
   SysFreeString bstrTargetPath
   PROPERTY = dwsTargetPath
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.TargetPath (BYREF wszTargetPath AS WSTRING)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrTargetPath AS AFX_BSTR = SysAllocString(wszTargetPath)
   m_Result = m_pLnk->put_TargetPath(bstrTargetPath)
   SysFreeString bstrTargetPath
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a window style to a shortcut, or identifies the type of window style used by a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.WindowStyle () AS LONG
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM nStyle AS LONG
   m_Result = m_pLnk->get_Windowstyle(@nStyle)
   PROPERTY = nStyle
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.WindowStyle (BYVAL nstyle AS LONG)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   m_Result = m_pLnk->put_Windowstyle(nStyle)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a working directory to a shortcut, or identifies the working directory used by a shortcut.
' ========================================================================================
PRIVATE PROPERTY CShortcut.WorkingDirectory () AS DWSTRING
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrWorkingDirectory AS AFX_BSTR
   m_Result = m_pLnk->get_WorkingDirectory(@bstrWorkingDirectory)
   SysFreeString bstrWorkingDirectory
   DIM dwsWorkingDirectory AS AFX_BSTR = bstrWorkingDirectory
   SysFreeString bstrWorkingDirectory
   PROPERTY = dwsWorkingDirectory
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CShortcut.WorkingDirectory (BYREF wszWorkingDirectory AS WSTRING)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrWorkingDirectory AS AFX_BSTR = SysAllocString(wszWorkingDirectory)
   m_Result = m_pLnk->put_WorkingDirectory(bstrWorkingDirectory)
   SysFreeString bstrWorkingDirectory
END PROPERTY
' ========================================================================================

' ========================================================================================
' Saves a shortcut object to disk.
' ========================================================================================
PRIVATE FUNCTION CShortcut.Save () AS HRESULT
   IF m_pLnk = NULL THEN m_Result = E_POINTER : RETURN m_Result
   m_Result = m_pLnk->Save
END FUNCTION
' ========================================================================================

' ========================================================================================
' CURLShortcut constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CURLShortcut (BYREF wszPathName AS WSTRING)

   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE

   ' // Create an instance of the IWshURLShortcut interface
   DIM pWsh AS Afx_IWshShell PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell_ AS GUID = (&hF935DC21, &h1CF0, &h11D0, {&hAD, &hB9, &h00, &hC0, &h4F, &hD5, &h8A, &h0B})
   m_Result = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell_, @pWsh)
   IF pWsh = NULL THEN m_Result = E_NOINTERFACE
   IF pWsh THEN
      DIM bstrPathName AS AFX_BSTR = SysAllocString(wszPathName)
      m_Result = pWsh->CreateShortcut(bstrPathName, cast(Afx_IDIspatch PTR PTR, @m_pLnk))
      SysFreeString bstrPathName
      pWsh->Release
   END IF
   IF pWsh = NULL THEN EXIT CONSTRUCTOR

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CURLShortcut destructor
' ========================================================================================
PRIVATE DESTRUCTOR CURLShortcut
   ' // Release the interface
   IF m_pLnk THEN m_pLnk->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CURLShortcut.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CURLShortcut.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   IF nError = -1 THEN nError = m_Result
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      dwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   IF nError THEN dwsMsg = "Error &h" + HEX(nError) & CHR(13, 10) + dwsMsg
   RETURN dwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the fully qualified path to the shortcut's target.
' ========================================================================================
PRIVATE PROPERTY CURLShortcut.FullName () AS DWSTRING
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrFullName AS AFX_BSTR
   m_Result = m_pLnk->get_FullName(@bstrFullName)
   DIM dwsFullName AS DWSTRING = bstrFullName
   SysFreeString bstrFullName
   PROPERTY = dwsFullName
END PROPERTY
' ========================================================================================

' ========================================================================================
' Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut.
' ========================================================================================
PRIVATE PROPERTY CURLShortcut.TargetPath () AS DWSTRING
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrTargetPath AS AFX_BSTR
   m_Result = m_pLnk->get_TargetPath(@bstrTargetPath)
   DIM dwsTargetPath AS DWSTRING = bstrTargetPath
   PROPERTY = dwsTargetPath
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CURLShortcut.TargetPath (BYREF wszTargetPath AS WSTRING)
   IF m_pLnk = NULL THEN m_Result = E_POINTER : EXIT PROPERTY
   DIM bstrTargetPath AS AFX_BSTR = SysAllocString(wszTargetPath)
   m_Result = m_pLnk->put_TargetPath(bstrTargetPath)
   SysFreeString bstrTargetPath
END PROPERTY
' ========================================================================================

' ========================================================================================
' Saves a shortcut object to disk.
' ========================================================================================
PRIVATE FUNCTION CURLShortcut.Save () AS HRESULT
   IF m_pLnk = NULL THEN m_Result = E_POINTER : RETURN m_Result
   m_Result = m_pLnk->Save
END FUNCTION
' ========================================================================================

END NAMESPACE
