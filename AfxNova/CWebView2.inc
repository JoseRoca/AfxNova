' ########################################################################################
' Microsoft Windows
' File: CWebView2.inc
' Contents: WebView2 class
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "AfxWebView2.bi"
#include once "AfxNova/DWSTRING.inc"
#include once "AfxNova/AfxCOM.inc"
USING AfxNova

NAMESPACE AfxNova

' ########################################################################################
' CWebView2 class
' ########################################################################################
TYPE CWebView2

   m_Result AS HRESULT                                ' // Result code
   m_hWnd AS HWND                                     ' // Handle of the window
   m_pEnv AS Afx_ICoreWebView2Environment  PTR        ' // Pointer to the Afx_ICoreWebView2Environment interface
   m_pController AS Afx_ICoreWebView2Controller PTR   ' // Pointer to the ICoreWebView2Controller interface
   m_pCoreWebView2 AS Afx_ICoreWebView2 PTR           ' // Pointer to the ICoreWebView2 interface

   DECLARE CONSTRUCTOR (BYVAL hWin AS HWND, BYVAL pUserDataFolder AS WSTRING PTR = NULL)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   DECLARE FUNCTION GetWebViewPtr () AS Afx_ICoreWebView2 PTR
   DECLARE FUNCTION GetControllerPtr () AS Afx_ICoreWebView2Controller PTR
   DECLARE FUNCTION GetEnvironmentPtr () AS Afx_ICoreWebView2Environment PTR
   DECLARE FUNCTION GetCoreWebView2 () AS Afx_ICoreWebView2 PTR
   DECLARE FUNCTION IsReady (BYVAL dwMaxWaitMilliseconds AS LONG = -1) AS BOOLEAN
   DECLARE FUNCTION Navigate (BYVAL uri AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION NavigateToString (BYVAL htmlContent AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION Stop () AS HRESULT
   DECLARE FUNCTION Reload () AS HRESULT
   DECLARE FUNCTION GoBack () AS HRESULT
   DECLARE FUNCTION GoForward () AS HRESULT
   DECLARE FUNCTION CanGoBack () AS BOOLEAN
   DECLARE FUNCTION CanGoForward () AS BOOLEAN
   DECLARE FUNCTION GetIsVisible () AS BOOLEAN
   DECLARE FUNCTION SetIsVisible (BYVAL isVisible AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION GetBrowserProcessId () AS UINT32
   DECLARE FUNCTION GetSource () AS DWSTRING
   DECLARE FUNCTION Close () AS HRESULT
   DECLARE FUNCTION GetParentWindow () AS HWND
   DECLARE FUNCTION SetParentWindow (BYVAL parentWindow AS HWND) AS HRESULT
   DECLARE FUNCTION GetZoomFactor () AS DOUBLE
   DECLARE FUNCTION SetZoomFactor (BYVAL zoomFactor AS DOUBLE) AS HRESULT
   DECLARE FUNCTION GetBounds () AS RECT
   DECLARE FUNCTION SetBounds (BYVAL bounds AS RECT) AS HRESULT
   DECLARE FUNCTION MoveFocus (BYVAL reason AS COREWEBVIEW2_MOVE_FOCUS_REASON) AS HRESULT
   DECLARE FUNCTION NotifyParentWindowPositionChanged () AS HRESULT
   DECLARE FUNCTION ExecuteScript (BYVAL javaScript AS LPCWSTR, BYVAL handler AS Afx_ICoreWebView2ExecuteScriptCompletedHandler PTR) AS HRESULT
   DECLARE FUNCTION PostWebMessageAsJson (BYVAL webMessageAsJson AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION PostWebMessageAsString (BYVAL webMessageAsString AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION CapturePreview (BYVAL imageFormat AS COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT, BYVAL imageStream AS IStream PTR, BYVAL handler AS Afx_ICoreWebView2CapturePreviewCompletedHandler PTR) AS HRESULT
   DECLARE FUNCTION OpenDevToolsWindow () AS HRESULT

   ' // Settings
   DECLARE FUNCTION GetSettings () AS Afx_ICoreWebView2Settings PTR
   DECLARE PROPERTY IsScriptEnabled () AS BOOLEAN
   DECLARE PROPERTY IsScriptEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   DECLARE PROPERTY IsWebMessageEnabled () AS BOOLEAN
   DECLARE PROPERTY IsWebMessageEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   DECLARE PROPERTY AreDefaultScriptDialogsEnabled () AS BOOLEAN
   DECLARE PROPERTY AreDefaultScriptDialogsEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   DECLARE PROPERTY IsStatusBarEnabled () AS BOOLEAN
   DECLARE PROPERTY IsStatusBarEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   DECLARE PROPERTY AreDevToolsEnabled () AS BOOLEAN
   DECLARE PROPERTY AreDevToolsEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   DECLARE PROPERTY AreDefaultContextMenusEnabled () AS BOOLEAN
   DECLARE PROPERTY AreDefaultContextMenusEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   DECLARE PROPERTY AreHostObjectsAllowed () AS BOOLEAN
   DECLARE PROPERTY AreHostObjectsAllowed (BYVAL allowed AS BOOLEAN) AS HRESULT
   DECLARE PROPERTY AreRemoteObjectsAllowed () AS BOOLEAN
   DECLARE PROPERTY AreRemoteObjectsAllowed (BYVAL allowed AS BOOLEAN) AS HRESULT
   DECLARE PROPERTY IsZoomControlEnabled () AS BOOLEAN
   DECLARE PROPERTY IsZoomControlEnabled (BYVAL allowed AS BOOLEAN) AS HRESULT
   DECLARE PROPERTY IsBuiltInErrorPageEnabled () AS BOOLEAN
   DECLARE PROPERTY IsBuiltInErrorPageEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT

   ' // Events
   DECLARE FUNCTION AddNavigationCompleted (BYVAL eventHandler AS Afx_ICoreWebView2NavigationCompletedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE FUNCTION RemoveNavigationCompleted (BYVAL token AS EventRegistrationToken) AS HRESULT

END TYPE
' ========================================================================================


' ########################################################################################
' CWebView2EnvironmentCompletedHandlerImpl class
' Implementation of the ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler callback interface.
' ########################################################################################
TYPE CWebView2EnvironmentCompletedHandlerImpl EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG

   ' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL createdEnvironment AS Afx_ICoreWebView2Environment PTR) AS HRESULT

   DECLARE CONSTRUCTOR (BYVAL pWebView2 AS CWebView2 PTR)
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0
   m_pWebView2 AS CWebView2 PTR

END TYPE
' ########################################################################################

' ########################################################################################
' CWebView2ControllerCompletedHandlerImpl class
' Implementation of the ICoreWebView2CreateCoreWebView2ControllerCompletedHandler callback interface.
' ########################################################################################
TYPE CWebView2ControllerCompletedHandlerImpl EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG

   ' ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL createdController AS Afx_ICoreWebView2Controller PTR) AS HRESULT

   DECLARE CONSTRUCTOR (BYVAL pWebView2 AS CWebView2 PTR)
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0
   m_pWebView2 AS CWebView2 PTR

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' Usage 1: DIM pWebView2 AS CWebView2 = hWin
' CreateCoreWebView2Environment, so it relies on the default runtime and default user data folder.
' Usage 2: DIM pWebView2 AS CWebView2 = CWebView2(hWin, wszUserFolder)
' Calls CreateCoreWebView2EnvironmentWithOptions, pointing to any folder you choose.
' This gives you a separate profile: cookies, cache, storage isolated from the default.
' Slightly slower startup, but perfect for multi-tab or multi-user contexts.
' ========================================================================================
PRIVATE CONSTRUCTOR CWebView2 (BYVAL hWin AS HWND, BYVAL pUserDataFolder AS WSTRING PTR = NULL)
   CWV2_DP("hWin: " & WSTR(hWin) & " - this: " & WSTR(@this))
   ' // Initialize the COM library
   CoInitialize NULL
   ' // Store the window handle
   IF hWin = NULL THEN SetResult(E_POINTER): EXIT CONSTRUCTOR
   SetWindowLongPtrW(hWin, GWLP_USERDATA, cast(LONG_PTR, @this))
   m_hwnd = hWin
   ' // Create the environment
   DIM pEnv AS CWebView2EnvironmentCompletedHandlerImpl PTR
   pEnv = NEW CWebView2EnvironmentCompletedHandlerImpl(@this)
   DIM hr AS HRESULT
   IF pUserDataFolder = NULL THEN
      hr = CreateCoreWebView2Environment(cast(ANY PTR, pEnv))
   ELSE
      hr = CreateCoreWebView2EnvironmentWithOptions(NULL, pUserDataFolder, NULL, cast(ANY PTR, pEnv))
   END IF
   SetResult(hr)
   CWV2_DP("pEnv: " & WSTR(pEnv) & " - hr: " & HEX(hr))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CWebView2
   CWV2_DP("release CoreWebView: " & WSTR(m_pCoreWebView2))
   CWV2_DP("release Controller: " & WSTR(m_pController))
   CWV2_DP("release Environment: " & WSTR(m_pEnv))
   IF m_pCoreWebView2 THEN m_pCoreWebView2->Release
   IF m_pController THEN
      m_pController->Close
      m_pController->Release
   END IF
   IF m_pEnv THEN m_pEnv->Release
   CWV2_DP("delete clas - this: " & WSTR(@this))
   ' // Uninitialize the COM library
   CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Get a pointer to the Afx_ICoreWebView2 interface
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetWebViewPtr () AS Afx_ICoreWebView2 PTR
   RETURN m_pCoreWebView2
END FUNCTION
' ========================================================================================
' ========================================================================================
' Get a pointer to the Afx_ICoreWebView2Controller interface
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetControllerPtr () AS Afx_ICoreWebView2Controller PTR
   RETURN m_pController
END FUNCTION
' ========================================================================================
' ========================================================================================
' Get a pointer to the Afx_ICoreWebView2Environment interface
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetEnvironmentPtr () AS Afx_ICoreWebView2Environment PTR
   RETURN m_pEnv
END FUNCTION
' ========================================================================================
' ========================================================================================
' Get an AddRef'ed pointer of the Afx_ICoreWebView2 interface
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetCoreWebView2 () AS Afx_ICoreWebView2 PTR
   DIM CoreWebView2 AS Afx_ICoreWebView2 PTR
   m_Result = m_pController->get_CoreWebView2(@CoreWebView2)
   IF m_Result = S_OK THEN RETURN CoreWebView2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CWebView2.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
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
' Chekcs if WebView2 is ready to be used.
' ========================================================================================
PRIVATE FUNCTION CWebView2.IsReady (BYVAL dwMaxWaitMilliseconds AS LONG = -1) AS BOOLEAN
   CWV2_DP("dwMaxWaitMilliseconds: " & WSTR(dwMaxWaitMilliseconds))
   DIM startTick AS DWORD = GetTickCount
   DO
      ' // Process messages to allow the callbacks to fire
      DIM r AS DWORD = MsgWaitForMultipleObjects(0, NULL, FALSE, 50, QS_ALLINPUT)
      IF r = WAIT_OBJECT_0 THEN
         DIM msg AS MSG
         WHILE PeekMessage(@msg, NULL, 0, 0, PM_REMOVE)
            TranslateMessage(@msg)
            DispatchMessage(@msg)
         WEND
      END IF
      ' Is it ready?
      IF m_pController <> 0 THEN RETURN TRUE
      ' // Maximum time ellapsed?
      IF dwMaxWaitMilliseconds > -1 THEN
         IF (GetTickCount - startTick) >= dwMaxWaitMilliseconds THEN EXIT DO
      END IF
   LOOP
   RETURN FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Cause a navigation of the top-level document to run to the specified URI.
' ========================================================================================
PRIVATE FUNCTION CWebView2.Navigate (BYVAL uri AS LPCWSTR) AS HRESULT
   IF uri THEN
      CWV2_DP(*uri)
   ELSE
      CWV2_DP("")
   END IF
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->Navigate(uri)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Initiates a navigation to htmlContent as source HTML of a new document.
' ========================================================================================
PRIVATE FUNCTION CWebView2.NavigateToString (BYVAL htmlContent AS LPCWSTR) AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->NavigateToString(htmlContent)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Stop all navigations and pending resource fetches. Does not stop scripts.
' ========================================================================================
PRIVATE FUNCTION CWebView2.Stop () AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->Stop
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reload the current page.
' ========================================================================================
PRIVATE FUNCTION CWebView2.Reload () AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->Reload
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Navigates the WebView to the previous page in the navigation history.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GoBack () AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->GoBack
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Navigates the WebView to the next page in the navigation history.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GoForward () AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->GoForward
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Navigates the WebView to the previous page in the navigation history.
' ========================================================================================
PRIVATE FUNCTION CWebView2.CanGoBack () AS BOOLEAN
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS BOOL
   m_Result = m_pCoreWebView2->get_CanGoBack(@bRes)
   RETURN bRes
END FUNCTION
' ========================================================================================

' ========================================================================================
' TRUE if the WebView is able to navigate to a next page in the navigation history.
' ========================================================================================
PRIVATE FUNCTION CWebView2.CanGoForward () AS BOOLEAN
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS BOOL
   m_Result = m_pCoreWebView2->get_CanGoForward(@bRes)
   RETURN bRes
END FUNCTION
' ========================================================================================

' ========================================================================================
' The process ID of the browser process that hosts the WebView.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetBrowserProcessId () AS UINT32
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN 0
   DIM id AS UINT32
   m_Result = m_pCoreWebView2->get_BrowserProcessId(@id)
   RETURN id
END FUNCTION
' ========================================================================================

' ========================================================================================
' The URI of the current top level document.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetSource () AS DWSTRING
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN ""
   DIM uri AS WSTRING PTR
   DIM dwsUri AS DWSTRING
   m_Result = m_pCoreWebView2->get_Source(@uri)
   IF m_Result = S_OK THEN
      dwsUri = *uri
      CoTaskMemFree uri
   END IF
   RETURN dwsUri
END FUNCTION
' ========================================================================================

' ========================================================================================
' Run JavaScript code from the javascript parameter in the current top-level document
' rendered in the WebView.
' ========================================================================================
PRIVATE FUNCTION CWebView2.ExecuteScript (BYVAL javaScript AS LPCWSTR, BYVAL handler AS Afx_ICoreWebView2ExecuteScriptCompletedHandler PTR) AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->ExecuteScript(javaScript, handler)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Post the specified webMessage to the top level document in this WebView.
' ========================================================================================
PRIVATE FUNCTION CWebView2.PostWebMessageAsJson (BYVAL webMessageAsJson AS LPCWSTR) AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->PostWebMessageAsJson(webMessageAsJson)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Posts a message that is a simple string rather than a JSON string representation of a JavaScript object.
' ========================================================================================
PRIVATE FUNCTION CWebView2.PostWebMessageAsString (BYVAL webMessageAsString AS LPCWSTR) AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->PostWebMessageAsString(webMessageAsString)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Capture an image of what WebView is displaying.
' ========================================================================================
PRIVATE FUNCTION CWebView2.CapturePreview (BYVAL imageFormat AS COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT, BYVAL imageStream AS IStream PTR, BYVAL handler AS Afx_ICoreWebView2CapturePreviewCompletedHandler PTR) AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->CapturePreview(imageFormat, imageStream, handler)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens the DevTools window for the current document in the WebView.
' ========================================================================================
PRIVATE FUNCTION CWebView2.OpenDevToolsWindow () AS HRESULT
   CWV2_DP("")
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pCoreWebView2->OpenDevToolsWindow
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes the WebView and cleans up the underlying browser instance.
' ========================================================================================
PRIVATE FUNCTION CWebView2.Close () AS HRESULT
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pController->Close
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The parent window provided by the app that this WebView is using to render content.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetParentWindow () AS HWND
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN NULL
   DIM hWin AS HWND
   m_Result = m_pController->get_ParentWindow(@hWin)
   RETURN hWin
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the parent window for the WebView.
' ========================================================================================
PRIVATE FUNCTION CWebView2.SetParentWindow (BYVAL parentWindow AS HWND) AS HRESULT
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pController->put_ParentWindow(parentWindow)
   IF m_Result = S_OK THEN
      SetWindowLongPtrW(parentWindow, GWLP_USERDATA, cast(LONG_PTR, @this))
      this.m_hwnd = parentWindow
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The IsVisible property determines whether to show or hide the WebView2.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetIsVisible () AS BOOLEAN
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS BOOL
   m_Result = m_pController->get_IsVisible(@bRes)
   RETURN bRes
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the IsVisible property.
' ========================================================================================
PRIVATE FUNCTION CWebView2.SetIsVisible (BYVAL IsVisible AS BOOLEAN) AS HRESULT
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pController->put_IsVisible(IsVisible)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The zoom factor for the WebView.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetZoomFactor () AS DOUBLE
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN 0
   DIM zoom AS DOUBLE
   m_Result = m_pController->get_ZoomFactor(@zoom)
   RETURN zoom
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the ZoomFactor property.
' ========================================================================================
PRIVATE FUNCTION CWebView2.SetZoomFactor (BYVAL zoomFactor AS DOUBLE) AS HRESULT
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pController->put_ZoomFactor(zoomFactor)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the WebView bounds.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetBounds () AS RECT
   CWV2_DP("")
   DIM bounds AS RECT
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN bounds
   m_Result = m_pController->get_Bounds(@bounds)
   RETURN bounds
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the Bounds property.
' ========================================================================================
PRIVATE FUNCTION CWebView2.SetBounds (BYVAL bounds AS RECT) AS HRESULT
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pController->put_Bounds(bounds)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the Bounds property.
' ========================================================================================
PRIVATE FUNCTION CWebView2.MoveFocus (BYVAL reason AS COREWEBVIEW2_MOVE_FOCUS_REASON) AS HRESULT
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pController->MoveFocus(reason)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' This is a notification separate from Bounds that tells WebView that the main WebView
' parent (or any ancestor) HWND moved.
' ========================================================================================
PRIVATE FUNCTION CWebView2.NotifyParentWindowPositionChanged () AS HRESULT
   CWV2_DP("")
   IF m_pController = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = m_pController->NotifyParentWindowPositionChanged
   RETURN m_Result
END FUNCTION
' ========================================================================================

' // *** Settings ************************************************************************

' ========================================================================================
' Gets a pointer to the ICoreWebView2Settings interface.
' The ICoreWebView2Settings object contains various modifiable settings for the running WebView.
' ========================================================================================
PRIVATE FUNCTION CWebView2.GetSettings () AS Afx_ICoreWebView2Settings PTR
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN NULL
   DIM settings AS Afx_ICoreWebView2Settings PTR
   m_Result = m_pCoreWebView2->get_Settings(@settings)
   RETURN settings
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.IsScriptEnabled () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_IsScriptEnabled(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.IsScriptEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_IsScriptEnabled(enabled)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.IsWebMessageEnabled () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_IsWebMessageEnabled(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.IsWebMessageEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_IsWebMessageEnabled(enabled)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.AreDefaultScriptDialogsEnabled () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_AreDefaultScriptDialogsEnabled(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.AreDefaultScriptDialogsEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_AreDefaultScriptDialogsEnabled(enabled)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.IsStatusBarEnabled () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_IsStatusBarEnabled(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.IsStatusBarEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_IsStatusBarEnabled(enabled)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.AreDevToolsEnabled () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_AreDevToolsEnabled(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.AreDevToolsEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_AreDevToolsEnabled(enabled)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.AreDefaultContextMenusEnabled () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_AreDefaultContextMenusEnabled(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.AreDefaultContextMenusEnabled (BYVAL enabled AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_AreDefaultContextMenusEnabled(enabled)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.AreHostObjectsAllowed () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_AreHostObjectsAllowed(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.AreHostObjectsAllowed (BYVAL allowed AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_AreHostObjectsAllowed(allowed)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.AreRemoteObjectsAllowed () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_AreRemoteObjectsAllowed(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.AreRemoteObjectsAllowed (BYVAL allowed AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_AreRemoteObjectsAllowed(allowed)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.IsZoomControlEnabled () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_IsZoomControlEnabled(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.IsZoomControlEnabled (BYVAL allowed AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_IsZoomControlEnabled(allowed)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
PRIVATE PROPERTY CWebView2.IsBuiltInErrorPageEnabled () AS BOOLEAN
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN FALSE
   DIM bRes AS WINBOOL
   m_Result = pSettings->get_IsBuiltInErrorPageEnabled(@bRes)
   pSettings->Release
   RETURN bRes
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CWebView2.IsBuiltInErrorPageEnabled (BYVAL allowed AS BOOLEAN) AS HRESULT
   IF m_pCoreWebView2 = NULL THEN m_Result = E_POINTER : RETURN m_Result
   DIM pSettings AS Afx_ICoreWebView2Settings PTR = this.GetSettings
   IF pSettings = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   m_Result = pSettings->put_IsBuiltInErrorPageEnabled(allowed)
   pSettings->Release
   RETURN m_Result
END PROPERTY
' ========================================================================================

' ========================================================================================
' // Events
' ========================================================================================
PRIVATE FUNCTION CWebView2.AddNavigationCompleted (BYVAL eventHandler AS Afx_ICoreWebView2NavigationCompletedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   m_Result = m_pCoreWebView2->add_NavigationCompleted(eventHandler, token)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CWebView2.RemoveNavigationCompleted (BYVAL token AS EventRegistrationToken) AS HRESULT
   m_Result = m_pCoreWebView2->remove_NavigationCompleted(token)
   RETURN m_Result
END FUNCTION
' ========================================================================================

'++++++++++++

' ########################################################################################
' CWebView2EnvironmentCompletedHandlerImpl class
' Implementation of the ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler callback interface.
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
CONSTRUCTOR CWebView2EnvironmentCompletedHandlerImpl (BYVAL pWebView2 AS CWebView2 PTR)
   CWV2_DP("pWebView2: " & WSTR(pWebView2)  & " - this: " & WSTR(@this))
   m_pWebView2 = pWebView2
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CWebView2EnvironmentCompletedHandlerImpl
   CWV2_DP(WSTR(@this))
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CWebView2EnvironmentCompletedHandlerImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler) OR _
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
FUNCTION CWebView2EnvironmentCompletedHandlerImpl.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CWebView2EnvironmentCompletedHandlerImpl.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
FUNCTION CWebView2EnvironmentCompletedHandlerImpl.Invoke (BYVAL errorCode AS HRESULT, BYVAL createdEnvironment AS Afx_ICoreWebView2Environment PTR) AS HRESULT
   CWV2_DP("")
   IF errorCode = S_OK AND m_pWebView2 <> NULL THEN
      m_pWebView2->m_pEnv = createdEnvironment
      m_pWebView2->m_pEnv->AddRef
      ' // Create controller
      DIM pController AS CWebView2ControllerCompletedHandlerImpl PTR
      pController = NEW CWebView2ControllerCompletedHandlerImpl(m_pWebView2)
      m_pWebView2->m_pEnv->CreateCoreWebView2Controller(m_pWebView2->m_hWnd, cast(ANY PTR, pController))
   END IF
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2ControllerCompletedHandlerImpl class
' Implementation of the ICoreWebView2CreateCoreWebView2ControllerCompletedHandler callback interface.
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
CONSTRUCTOR CWebView2ControllerCompletedHandlerImpl (BYVAL pWebView2 AS CWebView2 PTR)
   CWV2_DP("pWebView2: " & WSTR(pWebView2)  & " - this: " & WSTR(@this))
   m_pWebView2 = pWebView2
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CWebView2ControllerCompletedHandlerImpl
   CWV2_DP(WSTR(@this))
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CWebView2ControllerCompletedHandlerImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler) OR _
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
FUNCTION CWebView2ControllerCompletedHandlerImpl.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CWebView2ControllerCompletedHandlerImpl.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
FUNCTION CWebView2ControllerCompletedHandlerImpl.Invoke (BYVAL errorCode AS HRESULT, BYVAL createdController AS Afx_ICoreWebView2Controller PTR) AS HRESULT
   CWV2_DP(WSTR(errorCode))
   IF errorCode = S_OK AND m_pWebView2 <> NULL THEN
      m_pWebView2->m_pController = createdController
      m_pWebView2->m_pController->AddRef
      CWV2_DP("createdController: " & WSTR(createdController))
      ' // Get a pointer to the ICoreWebView2 inerface
      DIM pCoreWebView2 AS Afx_ICoreWebView2 PTR
      IF createdController->get_CoreWebView2(@pCoreWebView2) = S_OK THEN
         m_pWebView2->m_pCoreWebView2 = pCoreWebView2
         ' Don't call AddRef here; it is already AddRef'ed
         ' // Make the WebView2 control visible
         createdController->put_IsVisible(TRUE)
         ' // adjust the control to the client area of the window
         DIM rc AS RECT
         GetClientRect(m_pWebView2->m_hwnd, @rc)
         createdController->put_Bounds(rc)
      END IF
   END IF
   RETURN S_OK
END FUNCTION
' =====================================================================================

' ########################################################################################
'                                       WRAPPERS
' ########################################################################################
PRIVATE FUNCTION AfxCWebView2Ptr (BYVAL hWin AS HWND) AS CWebView2 PTR
   IF IsWindow(hWin) = FALSE THEN RETURN NULL
   RETURN cast(CWebView2 PTR, GetWindowLongPtrW(hWin, GWLP_USERDATA))
END FUNCTION
' ########################################################################################

END NAMESPACE
