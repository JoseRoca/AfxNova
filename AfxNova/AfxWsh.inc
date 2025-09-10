' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxWsh.inc
' Purpose:  Windows scripting host wrappers
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
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "AfxNova/AfxWsh.bi"
#INCLUDE ONCE "AfxNova/DVARIANT.inc"

' ########################################################################################
'                                   *** REGISTRY ***
' ########################################################################################

' ========================================================================================
' Returns the value of a key or value-name from the registry.
' ========================================================================================
PRIVATE FUNCTION AfxRegRead (BYREF wszKey AS WSTRING) AS DWSTRING
   CoInitialize NULL
      ' // Create an instance of the IWshShell interface
   DIM pWsh AS Afx_IWshShell PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell_ AS GUID = (&hF935DC21, &h1CF0, &h11D0, {&hAD, &hB9, &h00, &hC0, &h4F, &hD5, &h8A, &h0B})
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell_, @pWsh)
   DIM dws AS DWSTRING
   IF pWsh THEN
      DIM bstrKey AS Afx_BSTR = SysAllocString(wszKey)
      DIM vValue AS VARIANT
      hr = pWsh->RegRead(bstrKey, @vValue)
      SysFreeString bstrKey
      IF hr = S_OK THEN
         dws = AfxVarToStr(@vValue)
         VariantClear(@vValue)
      END IF
      pWsh->Release
   END IF
   CoUninitialize
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value of a key or value-name from the registry.
' ========================================================================================
PRIVATE FUNCTION AfxRegWrite (BYREF wszKey AS WSTRING, BYREF dvValue AS DVARIANT, BYREF wszType AS WSTRING) AS HRESULT
   CoInitialize NULL
      ' // Create an instance of the IWshShell interface
   DIM pWsh AS Afx_IWshShell PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell_ AS GUID = (&hF935DC21, &h1CF0, &h11D0, {&hAD, &hB9, &h00, &hC0, &h4F, &hD5, &h8A, &h0B})
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell_, @pWsh)
   DIM dws AS DWSTRING
   IF pWsh THEN
      DIM bstrKey AS Afx_BSTR = SysAllocString(wszKey)
      DIM vType AS VARIANT
      vType.vt = VT_BSTR
      vType.bstrVal = SysAllocString(wszType)
      hr = pWsh->RegWrite(bstrKey, dvValue, @vType)
      SysFreeString bstrKey
      VariantClear @vType
      pWsh->Release
   END IF
   CoUninitialize
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deletes a key or one of its values from the registry.
' ========================================================================================
PRIVATE FUNCTION AfxRegDelete (BYREF wszKey AS WSTRING) AS HRESULT
   CoInitialize NULL
      ' // Create an instance of the IWshShell interface
   DIM pWsh AS Afx_IWshShell PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell_ AS GUID = (&hF935DC21, &h1CF0, &h11D0, {&hAD, &hB9, &h00, &hC0, &h4F, &hD5, &h8A, &h0B})
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell_, @pWsh)
   IF pWsh THEN
      DIM bstrKey AS Afx_BSTR = SysAllocString(wszKey)
      hr = pWsh->RegDelete(bstrKey)
      SysFreeString bstrKey
   END IF
   CoUninitialize
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds an event entry to a log file.
' ========================================================================================
PRIVATE FUNCTION AfxLogEvent (BYVAL nType AS LONG, BYREF wszMessage AS WSTRING) AS HRESULT
   CoInitialize NULL
      ' // Create an instance of the IWshShell interface
   DIM pWsh AS Afx_IWshShell2 PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell2_ AS GUID = (&h24BE5A30, &hEDFE, &h11D2, {&hB9, &h33, &h00, &h10, &h4B, &h36, &h5C, &h9F})
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell2_, @pWsh)
   IF pWsh THEN
      DIM vType AS VARIANT
      vType.vt = VT_I4
      vType.lVal = nType
      DIM bstrMsg AS Afx_BSTR = SysAllocString(wszMessage)
      DIM bstrTarget AS Afx_BSTR = SysALlocString("")
      DIM bSuccess AS VARIANT_BOOL
      hr = pWsh->LogEvent(@vType, bstrMsg, bstrTarget, @bSuccess)
      SysFreeString bstrMsg
      SysFreeString bstrTarget
      pWsh->Release
   END IF
   CoUninitialize
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends keystrokes to the active window.
' Example:
' AfxRun("calc")
' AfxForceSetForegroundWindow(FindWindow("Windows.UI.Core.CoreWindow", NULL))
' ---or-- AfxForceSetForegroundWindow(FindWindow(NULL, "Calc"))
' SLEEP 1000
' AfxSendKeys("1{+}")
' SLEEP 500
' AfxSendKeys("2")
' SLEEP 500
' AfxSendKeys("~")
' SLEEP 500
' AfxSendKeys("*3")
' SLEEP 500
' AfxSendKeys("~")
' SLEEP 2500
' ========================================================================================
PRIVATE FUNCTION AfxSendKeys (BYREF wszKeys AS WSTRING, BYVAL bWait AS BOOLEAN = FALSE) AS HRESULT
   CoInitialize NULL
      ' // Create an instance of the IWshShell interface
   DIM pWsh AS Afx_IWshShell2 PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell2_ AS GUID = (&h24BE5A30, &hEDFE, &h11D2, {&hB9, &h33, &h00, &h10, &h4B, &h36, &h5C, &h9F})
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell2_, @pWsh)
   IF pWsh THEN
      DIM bstrKeys AS Afx_BSTR = SysAllocString(wszKeys)
      DIM vWait AS VARIANT
      vWait.vt = VT_BOOL
      vWait.boolVal = bWait
      hr = pWsh->SendKeys(bstrKeys, @vWait)
      SysFreeString(bstrKeys)
      pWsh->Release
   END IF
   CoUninitialize
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' Runs a program in a new process.
' ========================================================================================
PRIVATE FUNCTION AfxRun (BYREF wszCommand AS WSTRING, BYVAL WindowStyle AS LONG = 0, BYVAL WaitOnReturn AS BOOLEAN = FALSE) AS HRESULT
   CoInitialize NULL
      ' // Create an instance of the IWshShell interface
   DIM pWsh AS Afx_IWshShell PTR
   ' pWsh = AfxNewCom("WScript.Shell")  ' // Does not work well in my system
   DIM CLSID_WshShell_ AS GUID = (&h72C24DD5, &hD70A, &h438B, {&h8A, &h42, &h98, &h42, &h4B, &h88, &hAF, &hB8})
   DIM IID_IWshShell_ AS GUID = (&hF935DC21, &h1CF0, &h11D0, {&hAD, &hB9, &h00, &hC0, &h4F, &hD5, &h8A, &h0B})
   DIM hr AS HRESULT = CoCreateInstance(@CLSID_WshShell_, NULL, CLSCTX_INPROC_SERVER, @IID_IWshShell_, @pWsh)
   IF pWsh THEN
      DIM bstrCommand AS Afx_BSTR = SysAllocString(wszCommand)
      DIM vWindowStyle AS VARIANT
      vWindowStyle.vt = VT_I4
      vWindowStyle.lVal = WindowStyle
      DIM vWaitOnReturn AS VARIANT
      vWaitOnReturn.vt = VT_BOOL
      vWaitOnReturn.boolVal = WaitOnReturn
      DIM out_ExitCode AS INT_
      hr = pWsh->Run(bstrCommand, @vWindowStyle, @vWaitOnReturn, @out_ExitCode)
      SysFreeString(bstrCommand)
      pWsh->Release
   END IF
   CoUninitialize
   RETURN hr
END FUNCTION
' ========================================================================================
