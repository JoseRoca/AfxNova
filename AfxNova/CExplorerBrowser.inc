' ########################################################################################
' Platform: Microsoft Windows
' Filename: CExplorerBrowser.inc
' Purpose:  Wrapper class for the IExplorerBrowser interface
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
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "win/shobjidl.bi"
#include once "AfxNova/DWSTRING.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' Macro for debug
' To allow debugging, define _CEXPBROW_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CEXPBROW_DEBUG_
   #define _CEXPBROW_DEBUG_ 0
#ENDIF
#ifndef _CEXPBROW_DP_
   #define _CEXPBROW_DP_ 1
   #MACRO CEXPBROW_DP(st)
      #IF (_CEXPBROW_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' CExplorerBrowser class
' ========================================================================================
TYPE CExplorerBrowser

Public:
   m_Result AS HRESULT
   m_pExplorer AS IExplorerBrowser PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION ObjPtr () AS IExplorerBrowser PTR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   DECLARE FUNCTION Initialize (BYVAL hwndParent AS HWND, BYVAL prc AS RECT PTR, BYVAL pfs AS FOLDERSETTINGS PTR) AS HRESULT
   DECLARE FUNCTION Destroy () AS HRESULT
   DECLARE FUNCTION SetRect (BYVAL phdwp AS HDWP PTR, BYVAL rcBrowser AS RECT) AS HRESULT
   DECLARE FUNCTION SetPropertyBag (BYREF wszPropertyBag AS WSTRING) AS HRESULT
   DECLARE FUNCTION SetEmptyText (BYREF wszEmptyText AS WSTRING) AS HRESULT
   DECLARE FUNCTION SetFolderSettings (BYVAL pfs AS FOLDERSETTINGS PTR) AS HRESULT
   DECLARE FUNCTION SetFolderSettings (BYVAL ViewMode AS UINT, BYVAL fFlags AS UINT) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL psbe AS IExplorerBrowserEvents PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   DECLARE FUNCTION Advise (BYVAL psbe AS IExplorerBrowserEvents PTR) AS DWORD
   DECLARE FUNCTION Unadvise (BYVAL dwCookie AS DWORD) AS HRESULT
   DECLARE FUNCTION SetOptions (BYVAL dwFlag AS EXPLORER_BROWSER_OPTIONS) AS HRESULT
   DECLARE FUNCTION GetOptions (BYVAL pdwFlag AS EXPLORER_BROWSER_OPTIONS PTR) AS HRESULT
   DECLARE FUNCTION BrowseToIDList (BYVAL pidl AS LPCITEMIDLIST, BYVAL uFlags AS UINT) AS HRESULT
   DECLARE FUNCTION BrowseToObject (BYVAL punk AS IUnknown PTR, BYVAL uFlags AS UINT) AS HRESULT
   DECLARE FUNCTION FillFromObject (BYVAL punk AS IUnknown PTR, BYVAL dwFlags AS EXPLORER_BROWSER_FILL_FLAGS) AS HRESULT
   DECLARE FUNCTION RemoveAll () AS HRESULT
   DECLARE FUNCTION GetCurrentView (BYVAL riid AS IID PTR, BYVAL ppv AS ANY PTR PTR) AS HRESULT
   DECLARE FUNCTION GetCurrentView (BYVAL riid AS IID PTR) AS ANY PTR
   DECLARE FUNCTION NavigateBack () AS HRESULT
   DECLARE FUNCTION NavigateForward () AS HRESULT
   DECLARE FUNCTION NavigateTo (BYREF wszPath AS WSTRING) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CExplorerBrowser constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CExplorerBrowser
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Dictionary object
   DIM IID_IDictionary_  AS GUID  = (&h42C642C1, &h97E1, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   CoCreateInstance(@CLSID_ExplorerBrowser, NULL, CLSCTX_INPROC_SERVER, @IID_IExplorerBrowser, @m_pExplorer)
   CEXPBROW_DP("- m_pExplorer = " & WSTR(m_pExplorer))
'   IF m_pExplorer = NULL THEN EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CExplorerBrowser destructor
' ========================================================================================
PRIVATE DESTRUCTOR CExplorerBrowser
   CEXPBROW_DP("- m_pExplorer = " & WSTR(m_pExplorer))
   ' // Destroy the instance of the Explorer browser
   IF m_pExplorer THEN
      m_pExplorer->lpvtbl->Destroy(m_pExplorer)
      m_pExplorer->lpvtbl->Release(m_pExplorer)
   END IF
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the underlying IUnknown pointer.
' Don't call IUnknown_Release on it.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.ObjPtr () AS IExplorerBrowser PTR
   RETURN m_pExplorer
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
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
   IF nError THEN dwsMsg = "Error &h" & HEX(nError) & CHR(13, 10) & dwsMsg
   RETURN dwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Prepares the browser to be navigated.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.Initialize (BYVAL hwndParent AS HWND, BYVAL prc AS RECT PTR, BYVAL pfs AS FOLDERSETTINGS PTR) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->Initialize(m_pExplorer, hwndParent, prc, pfs))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Destroys the browser.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.Destroy () AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->Destroy(m_pExplorer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the size and position of the view windows created by the browser.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.SetRect (BYVAL phdwp AS HDWP PTR, BYVAL rcBrowser AS RECT) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->SetRect(m_pExplorer, phdwp, rcBrowser))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the name of the property bag.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.SetPropertyBag (BYREF wszPropertyBag AS WSTRING) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->SetPropertyBag(m_pExplorer, @wszPropertyBag))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the default empty text.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.SetEmptyText (BYREF wszEmptyText AS WSTRING) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->SetEmptyText(m_pExplorer, @wszEmptyText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the default empty text.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.SetFolderSettings (BYVAL pfs AS FOLDERSETTINGS PTR) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->SetFolderSettings(m_pExplorer, pfs))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.SetFolderSettings (BYVAL ViewMode AS UINT, BYVAL fFlags AS UINT) AS HRESULT
   CEXPBROW_DP("")
   DIM fs AS FOLDERSETTINGS = (ViewMode, fFlags)
   RETURN SetResult(m_pExplorer->lpvtbl->SetFolderSettings(m_pExplorer, @fs))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initiates a connection with IExplorerBrowser for event callbacks.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.Advise (BYVAL psbe AS IExplorerBrowserEvents PTR, BYVAL pdwCookie AS DWORD PTR) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->Advise(m_pExplorer, psbe, pdwCookie))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.Advise (BYVAL psbe AS IExplorerBrowserEvents PTR) AS DWORD
   CEXPBROW_DP("")
   DIM dwCookie AS DWORD
   m_Result = m_pExplorer->lpvtbl->Advise(m_pExplorer, psbe, @dwCookie)
   RETURN dwCookie
END FUNCTION
' ========================================================================================

' ========================================================================================
' Terminates an advisory connection.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.Unadvise (BYVAL dwCookie AS DWORD) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->Unadvise(m_pExplorer, dwCookie))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current browser options.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.SetOptions (BYVAL dwFlag AS EXPLORER_BROWSER_OPTIONS) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->SetOptions(m_pExplorer, dwFlag))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current browser options.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.GetOptions (BYVAL pdwFlag AS EXPLORER_BROWSER_OPTIONS PTR) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->GetOptions(m_pExplorer, pdwFlag))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Browses to a pointer to an item identifier list (PIDL)
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.BrowseToIDList (BYVAL pidl AS LPCITEMIDLIST, BYVAL uFlags AS UINT) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->BrowseToIDList(m_pExplorer, pidl, uFlags))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Browses to an object.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.BrowseToObject (BYVAL punk AS IUnknown PTR, BYVAL uFlags AS UINT) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->BrowseToObject(m_pExplorer, punk, uFlags))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a results folder and fills it with items.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.FillFromObject (BYVAL punk AS IUnknown PTR, BYVAL dwFlags AS EXPLORER_BROWSER_FILL_FLAGS) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->FillFromObject(m_pExplorer, punk, dwFlags))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all items from the results folder.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.RemoveAll () AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->RemoveAll(m_pExplorer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets an interface for the current view of the browser.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.GetCurrentView (BYVAL riid AS IID PTR, BYVAL ppv AS ANY PTR PTR) AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->GetCurrentView(m_pExplorer, riid, ppv))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.GetCurrentView (BYVAL riid AS IID PTR) AS ANY PTR
   CEXPBROW_DP("")
   DIM pv AS ANY PTR
   m_Result = m_pExplorer->lpvtbl->GetCurrentView(m_pExplorer, riid, @pv)
   RETURN pv
END FUNCTION
' ========================================================================================

' ========================================================================================
' Navigate back programatically.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.NavigateBack () AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->BrowseToIDList(m_pExplorer, NULL, SBSP_NAVIGATEBACK))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Navigate forward programatically.
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.NavigateForward () AS HRESULT
   CEXPBROW_DP("")
   RETURN SetResult(m_pExplorer->lpvtbl->BrowseToIDList(m_pExplorer, NULL, SBSP_NAVIGATEFORWARD))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Navigate to the specified folder
' ========================================================================================
PRIVATE FUNCTION CExplorerBrowser.NavigateTo (BYREF wszPath AS WSTRING) AS HRESULT
   CEXPBROW_DP("")
   DIM pidlBrowse AS ITEMIDLIST PTR
   m_Result = SHParseDisplayName(@wszPath, NULL, @pidlBrowse, 0, NULL)
   IF m_Result <> S_OK THEN RETURN m_Result
   m_Result = m_pExplorer->lpvtbl->BrowseToIDList(m_pExplorer, pidlBrowse, 0)
   ILFree(pidlBrowse)
   RETURN m_Result
END FUNCTION
' ========================================================================================

END NAMESPACE
