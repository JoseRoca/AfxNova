' ########################################################################################
' Library name: AfxWebView2.bi
' EXTERN_C const IID LIBID_WebView2;
' (C) 2025 José Roca.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/objidl.bi"
#include once "win/oaidl.bi"

' // The definition for BSTR in the FreeBASIC headers was inconveniently changed to WCHAR
#ifndef AFX_BSTR
   #define AFX_BSTR WSTRING PTR
#endif

TYPE EventRegistrationToken
   value AS __int64
END TYPE

' ========================================================================================
' Macro for debug
' To allow debugging, define _CWV2_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CWV2_DEBUG_
   #define _CWV2_DEBUG_ 0
#endif
#ifndef _CWV2_DP_
   #define _CWV2_DP_ 1
   #MACRO CWV2_DP(st)
      #if (_CWV2_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #endif
   #ENDMACRO
#endif
' ========================================================================================

'CONST AFX_ICoreWebView2AcceleratorKeyPressedEventArgs = "{9224476E-D8C3-4EB7-BB65-2FD7792B27CE}"
'CONST AFX_ICoreWebView2AcceleratorKeyPressedEventHandler = "{A7D303F9-503C-4B7E-BC40-5C7CE6CABAAA}"
'CONST AFX_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler = "{7082ABED-0591-428F-A722-60C2F814546B}"
'CONST AFX_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler = "{C20CF895-BA7C-493B-AB2E-8A6E3A3602A2}"
'CONST AFX_ICoreWebView2CapturePreviewCompletedHandler = "{DCED64F8-D9C7-4A3C-B9FD-FBBCA0B43496}"
'CONST AFX_ICoreWebView2 = "{189B8AAF-0426-4748-B9AD-243F537EB46B}"
'CONST AFX_ICoreWebView2Controller = "{7CCC5C7F-8351-4572-9077-9C1C80913835}"
'CONST AFX_ICoreWebView2ContentLoadingEventArgs = "{2A800835-2179-45D6-A745-6657E9A546B9}"
'CONST AFX_ICoreWebView2ContentLoadingEventHandler = "{7AF5CC82-AE19-4964-BD71-B9BC5F03E85D}"
'CONST AFX_ICoreWebView2DocumentTitleChangedEventHandler = "{6423D6B1-5A57-46C5-BA46-DBB3735EE7C9}"
'CONST AFX_ICoreWebView2ContainsFullScreenElementChangedEventHandler = "{120888E3-4CAD-4EC2-B627-B2016D05612D}"
'CONST AFX_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler = "{86EF6808-3C3F-4C6F-975E-8CE0B98F70BA}"
'CONST AFX_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler = "{8B4F98CE-DB0D-4E71-85FD-C4C4EF1F2630}"
'CONST AFX_ICoreWebView2Deferral = "{A7ED8BF0-3EC9-4E39-8427-3D6F157BD285}"
'CONST AFX_ICoreWebView2DevToolsProtocolEventReceivedEventArgs = "{F661B1C2-5FF5-4700-B723-C439034539B4}"
'CONST AFX_ICoreWebView2DevToolsProtocolEventReceivedEventHandler = "{8E1DED79-A40B-4271-8BE6-57640C167F4A}"
'CONST AFX_ICoreWebView2DevToolsProtocolEventReceiver = "{FE59C48C-540C-4A3C-8898-8E1602E0055D}"
'CONST AFX_ICoreWebView2Environment = "{DA66D884-6DA8-410E-9630-8C48F8B3A40E}"
'CONST AFX_ICoreWebView2EnvironmentOptions = "{97E9FBD9-646A-4B75-8682-149B71DACE59}"
'CONST AFX_ICoreWebView2ExecuteScriptCompletedHandler = "{3B717C93-3ED5-4450-9B13-7F56AA367AC7}"
'CONST AFX_ICoreWebView2FocusChangedEventHandler = "{76E67C71-663F-4C17-B71A-9381CCF3B94B}"
'CONST AFX_ICoreWebView2HistoryChangedEventHandler = "{54C9B7D7-D9E9-4158-861F-F97E1C3C6631}"
'CONST AFX_ICoreWebView2HttpHeadersCollectionIterator = "{4212F3A7-0FBC-4C9C-8118-17ED6370C1B3}"
'CONST AFX_ICoreWebView2HttpRequestHeaders = "{2C1F04DF-C90E-49E4-BD25-4A659300337B}"
'CONST AFX_ICoreWebView2HttpResponseHeaders = "{B5F6D4D5-1BFF-4869-85B8-158153017B04}"
'CONST AFX_ICoreWebView2MoveFocusRequestedEventArgs = "{71922903-B180-49D0-AED2-C9F9D10064B1}"
'CONST AFX_ICoreWebView2MoveFocusRequestedEventHandler = "{4B21D6DD-3DE7-47B0-8019-7D3ACE6E3631}"
'CONST AFX_ICoreWebView2NavigationCompletedEventArgs = "{361F5621-EA7F-4C55-95EC-3C5E6992EA4A}"
'CONST AFX_ICoreWebView2NavigationCompletedEventHandler = "{9F921239-20C4-455F-9E3F-6047A50E248B}"
'CONST AFX_ICoreWebView2NavigationStartingEventArgs = "{EE1938CE-D385-4CB0-854B-F498F78C3D88}"
'CONST AFX_ICoreWebView2NavigationStartingEventHandler = "{073337A4-64D2-4C7E-AC9F-987F0F613497}"
'CONST AFX_ICoreWebView2NewBrowserVersionAvailableEventHandler = "{E82E8242-EE39-4A57-A065-E13256D60342}"
'CONST AFX_ICoreWebView2NewWindowRequestedEventArgs = "{9EDC7F5F-C6EA-4F3C-827B-A8880794C0A9}"
'CONST AFX_ICoreWebView2NewWindowRequestedEventHandler = "{ACAA30EF-A40C-47BD-9CB9-D9C2AADC9FCB}"
'CONST AFX_ICoreWebView2PermissionRequestedEventArgs = "{774B5EA1-3FAD-435C-B1FC-A77D1ACD5EAF}"
'CONST AFX_ICoreWebView2PermissionRequestedEventHandler = "{543B4ADE-9B0B-4748-9AB7-D76481B223AA}"
'CONST AFX_ICoreWebView2ProcessFailedEventArgs = "{EA45D1F4-75C0-471F-A6E9-803FBFF8FEF2}"
'CONST AFX_ICoreWebView2ProcessFailedEventHandler = "{7D2183F9-CCA8-40F2-91A9-EAFAD32C8A9B}"
'CONST AFX_ICoreWebView2ScriptDialogOpeningEventArgs = "{B8F6356E-24DC-4D74-90FE-AD071E11CB91}"
'CONST AFX_ICoreWebView2ScriptDialogOpeningEventHandler = "{72D93789-2727-4A9B-A4FC-1B2609CBCBE3}"
'CONST AFX_ICoreWebView2Settings = "{203FBA37-6850-4DCC-A25A-58A351AC625D}"
'CONST AFX_ICoreWebView2SourceChangedEventArgs = "{BD9A4BFB-BE19-40BD-968B-EBCF0D727EF3}"
'CONST AFX_ICoreWebView2SourceChangedEventHandler = "{8FEDD1A7-3A33-416F-AF81-881EEB001433}"
'CONST AFX_ICoreWebView2WebMessageReceivedEventArgs = "{B263B5AE-9C54-4B75-B632-40AE1A0B6912}"
'CONST AFX_ICoreWebView2WebMessageReceivedEventHandler = "{199328C8-9964-4F5F-84E6-E875B1B763D6}"
'CONST AFX_ICoreWebView2WebResourceRequest = "{11B02254-B827-49F6-8974-30F6E6C55AF6}"
'CONST AFX_ICoreWebView2WebResourceRequestedEventArgs = "{2D7B3282-83B1-41CA-8BBF-FF18F6BFE320}"
'CONST AFX_ICoreWebView2WebResourceRequestedEventHandler = "{F6DC79F2-E1FA-4534-8968-4AFF10BBAA32}"
'CONST AFX_ICoreWebView2WebResourceResponse = "{5953D1FC-B08F-46DD-AFD3-66B172419CD0}"
'CONST AFX_ICoreWebView2WindowCloseRequestedEventHandler = "{63C89928-AD32-4421-A0E4-EC99B34AA97E}"
'CONST AFX_ICoreWebView2ZoomFactorChangedEventHandler = "{F1828246-8B98-4274-B708-ECDB6BF3843A}"

DIM SHARED AFX_IID_ICoreWebView2AcceleratorKeyPressedEventArgs AS GUID = TYPE(&h9224476E, &hD8C3, &h4EB7, {&hBB, &h65, &h2F, &hD7, &h79, &h2B, &h27, &hCE})
DIM SHARED AFX_IID_ICoreWebView2AcceleratorKeyPressedEventHandler AS GUID = TYPE(&hA7D303F9, &h503C, &h4B7E, {&hBC, &h40, &h5C, &h7C, &hE6, &hCA, &hBA, &hAA})
DIM SHARED AFX_IID_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler AS GUID = TYPE(&h7082ABED, &h0591, &h428F, {&hA7, &h22, &h60, &hC2, &hF8, &h14, &h54, &h6B})
DIM SHARED AFX_IID_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler AS GUID = TYPE(&hC20CF895, &hBA7C, &h493B, {&hAB, &h2E, &h8A, &h6E, &h3A, &h36, &h02, &hA2})
DIM SHARED AFX_IID_ICoreWebView2CapturePreviewCompletedHandler AS GUID = TYPE(&hDCED64F8, &hD9C7, &h4A3C, {&hB9, &hFD, &hFB, &hBC, &hA0, &hB4, &h34, &h96})
DIM SHARED AFX_IID_ICoreWebView2 AS GUID = TYPE(&h189B8AAF, &h0426, &h4748, {&hB9, &hAD, &h24, &h3F, &h53, &h7E, &hB4, &h6B})
DIM SHARED AFX_IID_ICoreWebView2Controller AS GUID = TYPE(&h7CCC5C7F, &h8351, &h4572, {&h90, &h77, &h9C, &h1C, &h80, &h91, &h38, &h35})
DIM SHARED AFX_IID_ICoreWebView2ContentLoadingEventArgs AS GUID = TYPE(&h2A800835, &h2179, &h45D6, {&hA7, &h45, &h66, &h57, &hE9, &hA5, &h46, &hB9})
DIM SHARED AFX_IID_ICoreWebView2ContentLoadingEventHandler AS GUID = TYPE(&h7AF5CC82, &hAE19, &h4964, {&hBD, &h71, &hB9, &hBC, &h5F, &h03, &hE8, &h5D})
DIM SHARED AFX_IID_ICoreWebView2DocumentTitleChangedEventHandler AS GUID = TYPE(&h6423D6B1, &h5A57, &h46C5, {&hBA, &h46, &hDB, &hB3, &h73, &h5E, &hE7, &hC9})
DIM SHARED AFX_IID_ICoreWebView2ContainsFullScreenElementChangedEventHandler AS GUID = TYPE(&h120888E3, &h4CAD, &h4EC2, {&hB6, &h27, &hB2, &h01, &h6D, &h05, &h61, &h2D})
DIM SHARED AFX_IID_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler AS GUID = TYPE(&h86EF6808, &h3C3F, &h4C6F, {&h97, &h5E, &h8C, &hE0, &hB9, &h8F, &h70, &hBA})
DIM SHARED AFX_IID_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler AS GUID = TYPE(&h8B4F98CE, &hDB0D, &h4E71, {&h85, &hFD, &hC4, &hC4, &hEF, &h1F, &h26, &h30})
DIM SHARED AFX_IID_ICoreWebView2Deferral AS GUID = TYPE(&hA7ED8BF0, &h3EC9, &h4E39, {&h84, &h27, &h3D, &h6F, &h15, &h7B, &hD2, &h85})
DIM SHARED AFX_IID_ICoreWebView2DevToolsProtocolEventReceivedEventArgs AS GUID = TYPE(&hF661B1C2, &h5FF5, &h4700, {&hB7, &h23, &hC4, &h39, &h03, &h45, &h39, &hB4})
DIM SHARED AFX_IID_ICoreWebView2DevToolsProtocolEventReceivedEventHandler AS GUID = TYPE(&h8E1DED79, &hA40B, &h4271, {&h8B, &hE6, &h57, &h64, &h0C, &h16, &h7F, &h4A})
DIM SHARED AFX_IID_ICoreWebView2DevToolsProtocolEventReceiver AS GUID = TYPE(&hFE59C48C, &h540C, &h4A3C, {&h88, &h98, &h8E, &h16, &h02, &hE0, &h05, &h5D})
DIM SHARED AFX_IID_ICoreWebView2Environment AS GUID = TYPE(&hDA66D884, &h6DA8, &h410E, {&h96, &h30, &h8C, &h48, &hF8, &hB3, &hA4, &h0E})
DIM SHARED AFX_IID_ICoreWebView2EnvironmentOptions AS GUID = TYPE(&h97E9FBD9, &h646A, &h4B75, {&h86, &h82, &h14, &h9B, &h71, &hDA, &hCE, &h59})
DIM SHARED AFX_IID_ICoreWebView2ExecuteScriptCompletedHandler AS GUID = TYPE(&h3B717C93, &h3ED5, &h4450, {&h9B, &h13, &h7F, &h56, &hAA, &h36, &h7A, &hC7})
DIM SHARED AFX_IID_ICoreWebView2FocusChangedEventHandler AS GUID = TYPE(&h76E67C71, &h663F, &h4C17, {&hB7, &h1A, &h93, &h81, &hCC, &hF3, &hB9, &h4B})
DIM SHARED AFX_IID_ICoreWebView2HistoryChangedEventHandler AS GUID = TYPE(&h54C9B7D7, &hD9E9, &h4158, {&h86, &h1F, &hF9, &h7E, &h1C, &h3C, &h66, &h31})
DIM SHARED AFX_IID_ICoreWebView2HttpHeadersCollectionIterator AS GUID = TYPE(&h4212F3A7, &h0FBC, &h4C9C, {&h81, &h18, &h17, &hED, &h63, &h70, &hC1, &hB3})
DIM SHARED AFX_IID_ICoreWebView2HttpRequestHeaders AS GUID = TYPE(&h2C1F04DF, &hC90E, &h49E4, {&hBD, &h25, &h4A, &h65, &h93, &h00, &h33, &h7B})
DIM SHARED AFX_IID_ICoreWebView2HttpResponseHeaders AS GUID = TYPE(&hB5F6D4D5, &h1BFF, &h4869, {&h85, &hB8, &h15, &h81, &h53, &h01, &h7B, &h04})
DIM SHARED AFX_IID_ICoreWebView2MoveFocusRequestedEventArgs AS GUID = TYPE(&h71922903, &hB180, &h49D0, {&hAE, &hD2, &hC9, &hF9, &hD1, &h00, &h64, &hB1})
DIM SHARED AFX_IID_ICoreWebView2MoveFocusRequestedEventHandler AS GUID = TYPE(&h4B21D6DD, &h3DE7, &h47B0, {&h80, &h19, &h7D, &h3A, &hCE, &h6E, &h36, &h31})
DIM SHARED AFX_IID_ICoreWebView2NavigationCompletedEventArgs AS GUID = TYPE(&h361F5621, &hEA7F, &h4C55, {&h95, &hEC, &h3C, &h5E, &h69, &h92, &hEA, &h4A})
DIM SHARED AFX_IID_ICoreWebView2NavigationCompletedEventHandler AS GUID = TYPE(&h9F921239, &h20C4, &h455F, {&h9E, &h3F, &h60, &h47, &hA5, &h0E, &h24, &h8B})
DIM SHARED AFX_IID_ICoreWebView2NavigationStartingEventArgs AS GUID = TYPE(&hEE1938CE, &hD385, &h4CB0, {&h85, &h4B, &hF4, &h98, &hF7, &h8C, &h3D, &h88})
DIM SHARED AFX_IID_ICoreWebView2NavigationStartingEventHandler AS GUID = TYPE(&h073337A4, &h64D2, &h4C7E, {&hAC, &h9F, &h98, &h7F, &h0F, &h61, &h34, &h97})
DIM SHARED AFX_IID_ICoreWebView2NewBrowserVersionAvailableEventHandler AS GUID = TYPE(&hE82E8242, &hEE39, &h4A57, {&hA0, &h65, &hE1, &h32, &h56, &hD6, &h03, &h42})
DIM SHARED AFX_IID_ICoreWebView2NewWindowRequestedEventArgs AS GUID = TYPE(&h9EDC7F5F, &hC6EA, &h4F3C, {&h82, &h7B, &hA8, &h88, &h07, &h94, &hC0, &hA9})
DIM SHARED AFX_IID_ICoreWebView2NewWindowRequestedEventHandler AS GUID = TYPE(&hACAA30EF, &hA40C, &h47BD, {&h9C, &hB9, &hD9, &hC2, &hAA, &hDC, &h9F, &hCB})
DIM SHARED AFX_IID_ICoreWebView2PermissionRequestedEventArgs AS GUID = TYPE(&h774B5EA1, &h3FAD, &h435C, {&hB1, &hFC, &hA7, &h7D, &h1A, &hCD, &h5E, &hAF})
DIM SHARED AFX_IID_ICoreWebView2PermissionRequestedEventHandler AS GUID = TYPE(&h543B4ADE, &h9B0B, &h4748, {&h9A, &hB7, &hD7, &h64, &h81, &hB2, &h23, &hAA})
DIM SHARED AFX_IID_ICoreWebView2ProcessFailedEventArgs AS GUID = TYPE(&hEA45D1F4, &h75C0, &h471F, {&hA6, &hE9, &h80, &h3F, &hBF, &hF8, &hFE, &hF2})
DIM SHARED AFX_IID_ICoreWebView2ProcessFailedEventHandler AS GUID = TYPE(&h7D2183F9, &hCCA8, &h40F2, {&h91, &hA9, &hEA, &hFA, &hD3, &h2C, &h8A, &h9B})
DIM SHARED AFX_IID_ICoreWebView2ScriptDialogOpeningEventArgs AS GUID = TYPE(&hB8F6356E, &h24DC, &h4D74, {&h90, &hFE, &hAD, &h07, &h1E, &h11, &hCB, &h91})
DIM SHARED AFX_IID_ICoreWebView2ScriptDialogOpeningEventHandler AS GUID = TYPE(&h72D93789, &h2727, &h4A9B, {&hA4, &hFC, &h1B, &h26, &h09, &hCB, &hCB, &hE3})
DIM SHARED AFX_IID_ICoreWebView2Settings AS GUID = TYPE(&h203FBA37, &h6850, &h4DCC, {&hA2, &h5A, &h58, &hA3, &h51, &hAC, &h62, &h5D})
DIM SHARED AFX_IID_ICoreWebView2SourceChangedEventArgs AS GUID = TYPE(&hBD9A4BFB, &hBE19, &h40BD, {&h96, &h8B, &hEB, &hCF, &h0D, &h72, &h7E, &hF3})
DIM SHARED AFX_IID_ICoreWebView2SourceChangedEventHandler AS GUID = TYPE(&h8FEDD1A7, &h3A33, &h416F, {&hAF, &h81, &h88, &h1E, &hEB, &h00, &h14, &h33})
DIM SHARED AFX_IID_ICoreWebView2WebMessageReceivedEventArgs AS GUID = TYPE(&hB263B5AE, &h9C54, &h4B75, {&hB6, &h32, &h40, &hAE, &h1A, &h0B, &h69, &h12})
DIM SHARED AFX_IID_ICoreWebView2WebMessageReceivedEventHandler AS GUID = TYPE(&h199328C8, &h9964, &h4F5F, {&h84, &hE6, &hE8, &h75, &hB1, &hB7, &h63, &hD6})
DIM SHARED AFX_IID_ICoreWebView2WebResourceRequest AS GUID = TYPE(&h11B02254, &hB827, &h49F6, {&h89, &h74, &h30, &hF6, &hE6, &hC5, &h5A, &hF6})
DIM SHARED AFX_IID_ICoreWebView2WebResourceRequestedEventArgs AS GUID = TYPE(&h2D7B3282, &h83B1, &h41CA, {&h8B, &hBF, &hFF, &h18, &hF6, &hBF, &hE3, &h20})
DIM SHARED AFX_IID_ICoreWebView2WebResourceRequestedEventHandler AS GUID = TYPE(&hF6DC79F2, &hE1FA, &h4534, {&h89, &h68, &h4A, &hFF, &h10, &hBB, &hAA, &h32})
DIM SHARED AFX_IID_ICoreWebView2WebResourceResponse AS GUID = TYPE(&h5953D1FC, &hB08F, &h46DD, {&hAF, &hD3, &h66, &hB1, &h72, &h41, &h9C, &hD0})
DIM SHARED AFX_IID_ICoreWebView2WindowCloseRequestedEventHandler AS GUID = TYPE(&h63C89928, &hAD32, &h4421, {&hA0, &hE4, &hEC, &h99, &hB3, &h4A, &hA9, &h7E})
DIM SHARED AFX_IID_ICoreWebView2ZoomFactorChangedEventHandler AS GUID = TYPE(&hF1828246, &h8B98, &h4274, {&hB7, &h08, &hEC, &hDB, &h6B, &hF3, &h84, &h3A})


' // Interfaces - Forward references

TYPE Afx_ICoreWebView2AcceleratorKeyPressedEventArgs AS Afx_ICoreWebView2AcceleratorKeyPressedEventArgs_
TYPE Afx_ICoreWebView2AcceleratorKeyPressedEventHandler AS Afx_ICoreWebView2AcceleratorKeyPressedEventHandler_
TYPE Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler AS Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_
TYPE Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler AS Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_
TYPE Afx_ICoreWebView2CapturePreviewCompletedHandler AS Afx_ICoreWebView2CapturePreviewCompletedHandler_
TYPE Afx_ICoreWebView2 AS Afx_ICoreWebView2_
TYPE Afx_ICoreWebView2Controller AS Afx_ICoreWebView2Controller_
TYPE Afx_ICoreWebView2ContentLoadingEventArgs AS Afx_ICoreWebView2ContentLoadingEventArgs_
TYPE Afx_ICoreWebView2ContentLoadingEventHandler AS Afx_ICoreWebView2ContentLoadingEventHandler_
TYPE Afx_ICoreWebView2DocumentTitleChangedEventHandler AS Afx_ICoreWebView2DocumentTitleChangedEventHandler_
TYPE Afx_ICoreWebView2ContainsFullScreenElementChangedEventHandler AS Afx_ICoreWebView2ContainsFullScreenElementChangedEventHandler_
TYPE Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler AS Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_
TYPE Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler AS Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_
TYPE Afx_ICoreWebView2Deferral AS Afx_ICoreWebView2Deferral_
TYPE Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs_
TYPE Afx_ICoreWebView2DevToolsProtocolEventReceivedEventHandler AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventHandler_
TYPE Afx_ICoreWebView2DevToolsProtocolEventReceiver AS Afx_ICoreWebView2DevToolsProtocolEventReceiver_
TYPE Afx_ICoreWebView2Environment AS Afx_ICoreWebView2Environment_
TYPE Afx_ICoreWebView2EnvironmentOptions AS Afx_ICoreWebView2EnvironmentOptions_
TYPE Afx_ICoreWebView2ExecuteScriptCompletedHandler AS Afx_ICoreWebView2ExecuteScriptCompletedHandler_
TYPE Afx_ICoreWebView2FocusChangedEventHandler AS Afx_ICoreWebView2FocusChangedEventHandler_
TYPE Afx_ICoreWebView2HistoryChangedEventHandler AS Afx_ICoreWebView2HistoryChangedEventHandler_
TYPE Afx_ICoreWebView2HttpHeadersCollectionIterator AS Afx_ICoreWebView2HttpHeadersCollectionIterator_
TYPE Afx_ICoreWebView2HttpRequestHeaders AS Afx_ICoreWebView2HttpRequestHeaders_
TYPE Afx_ICoreWebView2HttpResponseHeaders AS Afx_ICoreWebView2HttpResponseHeaders_
TYPE Afx_ICoreWebView2MoveFocusRequestedEventArgs AS Afx_ICoreWebView2MoveFocusRequestedEventArgs_
TYPE Afx_ICoreWebView2MoveFocusRequestedEventHandler AS Afx_ICoreWebView2MoveFocusRequestedEventHandler_
TYPE Afx_ICoreWebView2NavigationCompletedEventArgs AS Afx_ICoreWebView2NavigationCompletedEventArgs_
TYPE Afx_ICoreWebView2NavigationCompletedEventHandler AS Afx_ICoreWebView2NavigationCompletedEventHandler_
TYPE Afx_ICoreWebView2NavigationStartingEventArgs AS Afx_ICoreWebView2NavigationStartingEventArgs_
TYPE Afx_ICoreWebView2NavigationStartingEventHandler AS Afx_ICoreWebView2NavigationStartingEventHandler_ 
TYPE Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler AS Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler_
TYPE Afx_ICoreWebView2NewWindowRequestedEventArgs AS Afx_ICoreWebView2NewWindowRequestedEventArgs_
TYPE Afx_ICoreWebView2NewWindowRequestedEventHandler AS Afx_ICoreWebView2NewWindowRequestedEventHandler_
TYPE Afx_ICoreWebView2PermissionRequestedEventArgs AS Afx_ICoreWebView2PermissionRequestedEventArgs_
TYPE Afx_ICoreWebView2PermissionRequestedEventHandler AS Afx_ICoreWebView2PermissionRequestedEventHandler_
TYPE Afx_ICoreWebView2ProcessFailedEventArgs AS Afx_ICoreWebView2ProcessFailedEventArgs_
TYPE Afx_ICoreWebView2ProcessFailedEventHandler AS Afx_ICoreWebView2ProcessFailedEventHandler_
TYPE Afx_ICoreWebView2ScriptDialogOpeningEventArgs AS Afx_ICoreWebView2ScriptDialogOpeningEventArgs_
TYPE Afx_ICoreWebView2ScriptDialogOpeningEventHandler AS Afx_ICoreWebView2ScriptDialogOpeningEventHandler_
TYPE Afx_ICoreWebView2Settings AS Afx_ICoreWebView2Settings_
TYPE Afx_ICoreWebView2SourceChangedEventArgs AS Afx_ICoreWebView2SourceChangedEventArgs_
TYPE Afx_ICoreWebView2SourceChangedEventHandler AS Afx_ICoreWebView2SourceChangedEventHandler_
TYPE Afx_ICoreWebView2WebMessageReceivedEventArgs AS Afx_ICoreWebView2WebMessageReceivedEventArgs_
TYPE Afx_ICoreWebView2WebMessageReceivedEventHandler AS Afx_ICoreWebView2WebMessageReceivedEventHandler_
TYPE Afx_ICoreWebView2WebResourceRequest AS Afx_ICoreWebView2WebResourceRequest_
TYPE Afx_ICoreWebView2WebResourceRequestedEventArgs AS Afx_ICoreWebView2WebResourceRequestedEventArgs_
TYPE Afx_ICoreWebView2WebResourceRequestedEventHandler AS Afx_ICoreWebView2WebResourceRequestedEventHandler_
TYPE Afx_ICoreWebView2WebResourceResponse AS Afx_ICoreWebView2WebResourceResponse_
TYPE Afx_ICoreWebView2WindowCloseRequestedEventHandler AS Afx_ICoreWebView2WindowCloseRequestedEventHandler_
TYPE Afx_ICoreWebView2ZoomFactorChangedEventHandler AS Afx_ICoreWebView2ZoomFactorChangedEventHandler_

' // Enumerations

ENUM COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT
   COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_PNG = 0
   COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_JPEG = COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_PNG + 1
END ENUM

ENUM COREWEBVIEW2_SCRIPT_DIALOG_KIND
   COREWEBVIEW2_SCRIPT_DIALOG_KIND_ALERT = 0
   COREWEBVIEW2_SCRIPT_DIALOG_KIND_CONFIRM = 1 = COREWEBVIEW2_SCRIPT_DIALOG_KIND_ALERT + 1
   COREWEBVIEW2_SCRIPT_DIALOG_KIND_PROMPT = COREWEBVIEW2_SCRIPT_DIALOG_KIND_CONFIRM + 1
   COREWEBVIEW2_SCRIPT_DIALOG_KIND_BEFOREUNLOAD = COREWEBVIEW2_SCRIPT_DIALOG_KIND_PROMPT + 1
END ENUM

ENUM COREWEBVIEW2_PROCESS_FAILED_KIND
   ' // Number of constants: 10
   COREWEBVIEW2_PROCESS_FAILED_KIND_BROWSER_PROCESS_EXITED = 0
   COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_EXITED = COREWEBVIEW2_PROCESS_FAILED_KIND_BROWSER_PROCESS_EXITED + 1
   COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_UNRESPONSIVE = COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_EXITED + 1
END ENUM

ENUM COREWEBVIEW2_PERMISSION_KIND
   COREWEBVIEW2_PERMISSION_KIND_UNKNOWN_PERMISSION = 0
   COREWEBVIEW2_PERMISSION_KIND_MICROPHONE = COREWEBVIEW2_PERMISSION_KIND_UNKNOWN_PERMISSION + 1
   COREWEBVIEW2_PERMISSION_KIND_CAMERA = COREWEBVIEW2_PERMISSION_KIND_MICROPHONE + 1
   COREWEBVIEW2_PERMISSION_KIND_GEOLOCATION = COREWEBVIEW2_PERMISSION_KIND_CAMERA + 1
   COREWEBVIEW2_PERMISSION_KIND_NOTIFICATIONS = COREWEBVIEW2_PERMISSION_KIND_GEOLOCATION + 1
   COREWEBVIEW2_PERMISSION_KIND_OTHER_SENSORS = COREWEBVIEW2_PERMISSION_KIND_NOTIFICATIONS + 1
   COREWEBVIEW2_PERMISSION_KIND_CLIPBOARD_READ = COREWEBVIEW2_PERMISSION_KIND_OTHER_SENSORS + 1
END ENUM

ENUM COREWEBVIEW2_PERMISSION_STATE
   COREWEBVIEW2_PERMISSION_STATE_DEFAULT = 0
   COREWEBVIEW2_PERMISSION_STATE_ALLOW = COREWEBVIEW2_PERMISSION_STATE_DEFAULT + 1
   COREWEBVIEW2_PERMISSION_STATE_DENY = COREWEBVIEW2_PERMISSION_STATE_ALLOW + 1
END ENUM

ENUM COREWEBVIEW2_WEB_ERROR_STATUS
   COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN = 0
   COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_COMMON_NAME_IS_INCORRECT = COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_EXPIRED = COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_COMMON_NAME_IS_INCORRECT + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_CLIENT_CERTIFICATE_CONTAINS_ERRORS = COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_EXPIRED + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_REVOKED = COREWEBVIEW2_WEB_ERROR_STATUS_CLIENT_CERTIFICATE_CONTAINS_ERRORS + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_IS_INVALID = COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_REVOKED + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_SERVER_UNREACHABLE = COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_IS_INVALID + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_TIMEOUT = COREWEBVIEW2_WEB_ERROR_STATUS_SERVER_UNREACHABLE + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_ERROR_HTTP_INVALID_SERVER_RESPONSE = COREWEBVIEW2_WEB_ERROR_STATUS_TIMEOUT + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED = COREWEBVIEW2_WEB_ERROR_STATUS_ERROR_HTTP_INVALID_SERVER_RESPONSE + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET = COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_DISCONNECTED = COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_CANNOT_CONNECT = COREWEBVIEW2_WEB_ERROR_STATUS_DISCONNECTED + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_HOST_NAME_NOT_RESOLVED = COREWEBVIEW2_WEB_ERROR_STATUS_CANNOT_CONNECT + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED = COREWEBVIEW2_WEB_ERROR_STATUS_HOST_NAME_NOT_RESOLVED + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_REDIRECT_FAILED = COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED + 1
   COREWEBVIEW2_WEB_ERROR_STATUS_UNEXPECTED_ERROR = COREWEBVIEW2_WEB_ERROR_STATUS_REDIRECT_FAILED + 1
END ENUM

ENUM COREWEBVIEW2_WEB_RESOURCE_CONTEXT
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL = 0
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_STYLESHEET = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_STYLESHEET + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MEDIA = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FONT = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MEDIA + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SCRIPT = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FONT + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_XML_HTTP_REQUEST = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SCRIPT + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FETCH = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_XML_HTTP_REQUEST + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_TEXT_TRACK = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FETCH + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_EVENT_SOURCE = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_TEXT_TRACK + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_WEBSOCKET = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_EVENT_SOURCE + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MANIFEST = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_WEBSOCKET + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SIGNED_EXCHANGE = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MANIFEST + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_PING = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SIGNED_EXCHANGE + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_CSP_VIOLATION_REPORT = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_PING + 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_OTHER = COREWEBVIEW2_WEB_RESOURCE_CONTEXT_CSP_VIOLATION_REPORT + 1
END ENUM

ENUM COREWEBVIEW2_MOVE_FOCUS_REASON
   COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC = 0
   COREWEBVIEW2_MOVE_FOCUS_REASON_NEXT = COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC + 1
   COREWEBVIEW2_MOVE_FOCUS_REASON_PREVIOUS = COREWEBVIEW2_MOVE_FOCUS_REASON_NEXT + 1
END ENUM

ENUM COREWEBVIEW2_KEY_EVENT_KIND
   COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN = 0
   COREWEBVIEW2_KEY_EVENT_KIND_KEY_UP = COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN + 1
   COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN = COREWEBVIEW2_KEY_EVENT_KIND_KEY_UP + 1
   COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_UP = COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN + 1
END ENUM

TYPE COREWEBVIEW2_PHYSICAL_KEY_STATUS
   RepeatCount AS UINT32
   ScanCode AS UINT32
   IsExtendedKey AS WINBOOL
   IsMenuKeyDown AS WINBOOL
   WasKeyDown AS WINBOOL
   IsKeyReleased AS WINBOOL
END TYPE

' ========================================================================================
' Creates a WebView2 environment with a custom version of WebView2 Runtime, user data folder,
' and with or without additional options.
' ========================================================================================
PRIVATE FUNCTION CreateCoreWebView2EnvironmentWithOptions (BYVAL browserExecutableFolder AS PCWSTR, _
BYVAL userDataFolder AS PCWSTR, BYVAL environmentOptions AS Afx_ICoreWebView2EnvironmentOptions PTR, _
BYVAL environment_created_handler AS Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("WebView2Loader.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pCreateCoreWebView2EnvironmentWithOptions AS FUNCTION (BYVAL browserExecutableFolder AS PCWSTR, _
      BYVAL userDataFolder AS PCWSTR, BYVAL environmentOptions AS Afx_ICoreWebView2EnvironmentOptions PTR, _
      BYVAL environment_created_handler AS Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler PTR) AS HRESULT
   pCreateCoreWebView2EnvironmentWithOptions = DyLibSymbol(pLib, "CreateCoreWebView2EnvironmentWithOptions")
   IF pCreateCoreWebView2EnvironmentWithOptions THEN FUNCTION = pCreateCoreWebView2EnvironmentWithOptions(browserExecutableFolder, _
      userDataFolder, environmentOptions, environment_created_handler)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an environment with a custom version of Edge, user data directory and/or additional
' browser switches.
' This API is going to be removed in next SDK release.
' Please use CreateCoreWebView2EnvironmentWithOptions.
' ========================================================================================
PRIVATE FUNCTION CreateCoreWebView2EnvironmentWithDetails (BYVAL browserExecutableFolder AS PCWSTR, _
BYVAL userDataFolder AS PCWSTR, BYVAL additionalBrowserArguments AS PCWSTR, _
BYVAL environment_created_handler AS Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("WebView2Loader.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pCreateCoreWebView2EnvironmentWithDetails AS FUNCTION (BYVAL browserExecutableFolder AS PCWSTR, _
      BYVAL userDataFolder AS PCWSTR, BYVAL additionalBrowserArguments AS PCWSTR, _
      BYVAL environment_created_handler AS Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler PTR) AS HRESULT
   pCreateCoreWebView2EnvironmentWithDetails = DyLibSymbol(pLib, "CreateCoreWebView2EnvironmentWithDetails")
   IF pCreateCoreWebView2EnvironmentWithDetails THEN FUNCTION = pCreateCoreWebView2EnvironmentWithDetails(browserExecutableFolder, _
      userDataFolder, additionalBrowserArguments, environment_created_handler)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an evergreen WebView2 Environment using the installed Edge version.
' ========================================================================================
PRIVATE FUNCTION CreateCoreWebView2Environment (BYVAL environment_created_handler AS Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("WebView2Loader.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pCreateCoreWebView2Environment AS FUNCTION (BYVAL environment_created_handler AS Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler PTR) AS HRESULT
   pCreateCoreWebView2Environment = DyLibSymbol(pLib, "CreateCoreWebView2Environment")
   IF pCreateCoreWebView2Environment THEN FUNCTION = pCreateCoreWebView2Environment(environment_created_handler)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Get the browser version info including channel name if it is not the WebView2 Runtime.
' ========================================================================================
PRIVATE FUNCTION GetAvailableCoreWebView2BrowserVersionString (BYVAL browserExecutableFolder AS PCWSTR, BYVAL versionInfo AS LPWSTR PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("WebView2Loader.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pGetAvailableCoreWebView2BrowserVersionString AS FUNCTION (BYVAL browserExecutableFolder AS PCWSTR, BYVAL versionInfo AS LPWSTR PTR) AS HRESULT
   pGetAvailableCoreWebView2BrowserVersionString = DyLibSymbol(pLib, "GetAvailableCoreWebView2BrowserVersionString")
   IF pGetAvailableCoreWebView2BrowserVersionString THEN FUNCTION = pGetAvailableCoreWebView2BrowserVersionString(browserExecutableFolder, versionInfo)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Get the browser version info including channel name if it is not the WebView2 Runtime.
' ========================================================================================
PRIVATE FUNCTION CompareBrowserVersions (BYVAL version1 AS PCWSTR, BYVAL version2 AS PCWSTR, BYVAL result AS INT_ PTR) AS HRESULT
   FUNCTION = E_POINTER
   DIM AS ANY PTR pLib = DyLibLoad("WebView2Loader.dll")
   IF pLib = NULL THEN EXIT FUNCTION
   DIM pCompareBrowserVersions AS FUNCTION (BYVAL version1 AS PCWSTR, BYVAL version2 AS PCWSTR, BYVAL result AS INT_ PTR) AS HRESULT
   pCompareBrowserVersions = DyLibSymbol(pLib, "CompareBrowserVersions")
   IF pCompareBrowserVersions THEN FUNCTION = pCompareBrowserVersions(version1, version2, result)
   DyLibFree(pLib)
END FUNCTION
' ========================================================================================

' ########################################################################################
' Interface name: ICoreWebView2AcceleratorKeyPressedEventArgs
' MIDL_INTERFACE("9224476E-D8C3-4EB7-BB65-2FD7792B27CE")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2AcceleratorKeyPressedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2AcceleratorKeyPressedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2AcceleratorKeyPressedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_KeyEventKind (BYVAL KeyEventKind AS COREWEBVIEW2_KEY_EVENT_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_VirtualKey (BYVAL VirtualKey AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_KeyEventLParam (BYVAL lParam AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PhysicalKeyStatus (BYVAL PhysicalKeyStatus AS COREWEBVIEW2_PHYSICAL_KEY_STATUS PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL Handled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL Handled AS WINBOOL) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2AcceleratorKeyPressedEventHandler
' MIDL_INTERFACE("A7D303F9-503C-4B7E-BC40-5C7CE6CABAAA")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2AcceleratorKeyPressedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2AcceleratorKeyPressedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2AcceleratorKeyPressedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS Afx_ICoreWebView2AcceleratorKeyPressedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler
' IID: MIDL_INTERFACE("7082ABED-0591-428F-A722-60C2F814546B")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL id AS LPCWSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2CallDevToolsProtocolMethodCompletedHandler
' MIDL_INTERFACE("C20CF895-BA7C-493B-AB2E-8A6E3A3602A2")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL returnObjectAsJson AS LPCWSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2CapturePreviewCompletedHandler
' MIDL_INTERFACE("DCED64F8-D9C7-4A3C-B9FD-FBBCA0B43496")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CapturePreviewCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CapturePreviewCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CapturePreviewCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2
' MIDL_INTERFACE("189B8AAF-0426-4748-B9AD-243F537EB46B")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Settings (BYVAL Settings AS Afx_ICoreWebView2Settings PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Source (BYVAL uri AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Navigate (BYVAL uri AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION NavigateToString (BYVAL htmlContent AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_NavigationStarting (BYVAL eventHandler AS Afx_ICoreWebView2NavigationStartingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NavigationStarting (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ContentLoading (BYVAL eventHandler AS Afx_ICoreWebView2ContentLoadingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ContentLoading (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_SourceChanged (BYVAL eventHandler AS Afx_ICoreWebView2SourceChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_SourceChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_HistoryChanged (BYVAL eventHandler AS Afx_ICoreWebView2HistoryChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_HistoryChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_NavigationCompleted (BYVAL eventHandler AS Afx_ICoreWebView2NavigationCompletedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NavigationCompleted (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_FrameNavigationStarting (BYVAL eventHandler AS Afx_ICoreWebView2NavigationStartingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_FrameNavigationStarting (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_FrameNavigationCompleted (BYVAL eventHandler AS Afx_ICoreWebView2NavigationCompletedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_FrameNavigationCompleted (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ScriptDialogOpening (BYVAL eventHandler AS Afx_ICoreWebView2ScriptDialogOpeningEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ScriptDialogOpening (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_PermissionRequested (BYVAL eventHandler AS Afx_ICoreWebView2PermissionRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_PermissionRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ProcessFailed (BYVAL eventHandler AS Afx_ICoreWebView2ProcessFailedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ProcessFailed (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddScriptToExecuteOnDocumentCreated (BYVAL javaScript AS LPCWSTR, BYVAL handler AS Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveScriptToExecuteOnDocumentCreated (BYVAL id AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecuteScript (BYVAL javaScript AS LPCWSTR, BYVAL handler AS Afx_ICoreWebView2ExecuteScriptCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CapturePreview (BYVAL imageFormat AS COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT, BYVAL imageStream AS IStream PTR, BYVAL handler AS Afx_ICoreWebView2CapturePreviewCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Reload () AS HRESULT
   DECLARE ABSTRACT FUNCTION PostWebMessageAsJson (BYVAL webMessageAsJson AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION PostWebMessageAsString (BYVAL webMessageAsString AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_WebMessageReceived (BYVAL handler AS Afx_ICoreWebView2WebMessageReceivedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_WebMessageReceived (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION CallDevToolsProtocolMethod (BYVAL methodName AS LPCWSTR, BYVAL parametersAsJson AS LPCWSTR, BYVAL handler AS Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BrowserProcessId (BYVAL value AS UINT32 PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CanGoBack (BYVAL CanGoBack AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CanGoForward (BYVAL CanGoForward AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GoBack () AS HRESULT
   DECLARE ABSTRACT FUNCTION GoForward () AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDevToolsProtocolEventReceiver (BYVAL eventName AS LPCWSTR, BYVAL receiver AS Afx_ICoreWebView2DevToolsProtocolEventReceiver PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Stop () AS HRESULT
   DECLARE ABSTRACT FUNCTION add_NewWindowRequested (BYVAL eventHandler AS Afx_ICoreWebView2NewWindowRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NewWindowRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_DocumentTitleChanged (BYVAL eventHandler AS Afx_ICoreWebView2DocumentTitleChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_DocumentTitleChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DocumentTitle (BYVAL title AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddHostObjectToScript (BYVAL _name AS LPCWSTR, BYVAL _object AS VARIANT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveHostObjectFromScript (BYVAL _name AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OpenDevToolsWindow () AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ContainsFullScreenElementChanged (BYVAL eventHandler AS Afx_ICoreWebView2ContainsFullScreenElementChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ContainsFullScreenElementChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ContainsFullScreenElement (BYVAL ContainsFullScreenElement AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_WebResourceRequested (BYVAL eventHandler AS Afx_ICoreWebView2WebResourceRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_WebResourceRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddWebResourceRequestedFilter (BYVAL uri AS LPCWSTR, BYVAL ResourceContext AS COREWEBVIEW2_WEB_RESOURCE_CONTEXT) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveWebResourceRequestedFilter (BYVAL uri AS LPCWSTR, BYVAL ResourceContext AS COREWEBVIEW2_WEB_RESOURCE_CONTEXT) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_WindowCloseRequested (BYVAL eventHandler AS Afx_ICoreWebView2WindowCloseRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_WindowCloseRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2Controller
' MIDL_INTERFACE("7CCC5C7F-8351-4572-9077-9C1C80913835")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Controller_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Controller_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Controller_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsVisible (BYVAL IsVisible AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsVisible (BYVAL IsVisible AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Bounds (BYVAL Bounds AS RECT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Bounds (BYVAL Bounds AS RECT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ZoomFactor (BYVAL ZoomFactor AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ZoomFactor (BYVAL ZoomFactor AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ZoomFactorChanged (BYVAL eventHandler AS Afx_ICoreWebView2ZoomFactorChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ZoomFactorChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetBoundsAndZoomFactor (BYVAL Bounds AS RECT, BYVAL ZoomFactor AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveFocus (BYVAL reason AS COREWEBVIEW2_MOVE_FOCUS_REASON) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_MoveFocusRequested (BYVAL eventHandler AS Afx_ICoreWebView2MoveFocusRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_MoveFocusRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_GotFocus (BYVAL eventHandler AS Afx_ICoreWebView2FocusChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_GotFocus (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_LostFocus (BYVAL eventHandler AS Afx_ICoreWebView2FocusChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_LostFocus (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_AcceleratorKeyPressed (BYVAL eventHandler AS Afx_ICoreWebView2AcceleratorKeyPressedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_AcceleratorKeyPressed (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ParentWindow (BYVAL topLevelWindow AS HWND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ParentWindow (BYVAL topLevelWindow AS HWND) AS HRESULT
   DECLARE ABSTRACT FUNCTION NotifyParentWindowPositionChanged () AS HRESULT
   DECLARE ABSTRACT FUNCTION Close () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CoreWebView2 (BYVAL CoreWebView2 AS Afx_ICoreWebView2 PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2ContentLoadingEventArgs
' MIDL_INTERFACE("2A800835-2179-45D6-A745-6657E9A546B9")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ContentLoadingEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ContentLoadingEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ContentLoadingEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsErrorPage (BYVAL IsErrorPage AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NavigationId (BYVAL NavigationId AS UINT64 PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2ContentLoadingEventHandler
' MIDL_INTERFACE("7AF5CC82-AE19-4964-BD71-B9BC5F03E85D")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ContentLoadingEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ContentLoadingEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ContentLoadingEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ContentLoadingEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2DocumentTitleChangedEventHandler
' MIDL_INTERFACE("6423D6B1-5A57-46C5-BA46-DBB3735EE7C9")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DocumentTitleChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DocumentTitleChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DocumentTitleChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2ContainsFullScreenElementChangedEventHandler
' MIDL_INTERFACE("120888E3-4CAD-4EC2-B627-B2016D05612D")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ContainsFullScreenElementChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ContainsFullScreenElementChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ContainsFullScreenElementChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
' MIDL_INTERFACE("86EF6808-3C3F-4C6F-975E-8CE0B98F70BA")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL createdController AS Afx_ICoreWebView2Controller PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
' MIDL_INTERFACE("8B4F98CE-DB0D-4E71-85FD-C4C4EF1F2630")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL createdEnvironment AS Afx_ICoreWebView2Environment PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2Deferral
' MIDL_INTERFACE("A7ED8BF0-3EC9-4E39-8427-3D6F157BD285")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Deferral_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Deferral_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Deferral_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Complete () AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2DevToolsProtocolEventReceivedEventArgs
' MIDL_INTERFACE("F661B1C2-5FF5-4700-B723-C439034539B4")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ParameterObjectAsJson (BYVAL ParameterObjectAsJson AS LPWSTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2DevToolsProtocolEventReceivedEventHandler
' MIDL_INTERFACE("8E1DED79-A40B-4271-8BE6-57640C167F4A")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DevToolsProtocolEventReceivedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DevToolsProtocolEventReceivedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DevToolsProtocolEventReceivedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2DevToolsProtocolEventReceiver
' MIDL_INTERFACE("FE59C48C-540C-4A3C-8898-8E1602E0055D")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DevToolsProtocolEventReceiver_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DevToolsProtocolEventReceiver_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DevToolsProtocolEventReceiver_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION add_DevToolsProtocolEventReceived (BYVAL handler AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_DevToolsProtocolEventReceived (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2Environment
' MIDL_INTERFACE("DA66D884-6DA8-410E-9630-8C48F8B3A40E")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION CreateCoreWebView2Controller (BYVAL ParentWindow AS HWND, BYVAL handler AS Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateWebResourceResponse (BYVAL Content AS IStream PTR, BYVAL StatusCode AS INT_, BYVAL ReasonPhrase AS LPCWSTR, BYVAL Headers AS LPCWSTR, BYVAL Response AS Afx_ICoreWebView2WebResourceResponse PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BrowserVersionString (BYVAL versionInfo AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_NewBrowserVersionAvailable (BYVAL eventHandler AS Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NewBrowserVersionAvailable (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2EnvironmentOptions
' MIDL_INTERFACE("97E9FBD9-646A-4B75-8682-149B71DACE59")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EnvironmentOptions_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EnvironmentOptions_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EnvironmentOptions_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_AdditionalBrowserArguments (BYVAL value AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AdditionalBrowserArguments (BYVAL value AS LPWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Language (BYVAL value AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Language (BYVAL value AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TargetCompatibleBrowserVersion (BYVAL value AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TargetCompatibleBrowserVersion (BYVAL value AS LPCWSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2ExecuteScriptCompletedHandler
' MIDL_INTERFACE("3B717C93-3ED5-4450-9B13-7F56AA367AC7")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ExecuteScriptCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ExecuteScriptCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ExecuteScriptCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL resultObjectAsJson AS LPCWSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2FocusChangedEventHandler
' MIDL_INTERFACE("76E67C71-663F-4C17-B71A-9381CCF3B94B")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FocusChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FocusChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FocusChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2HistoryChangedEventHandler
' MIDL_INTERFACE("54C9B7D7-D9E9-4158-861F-F97E1C3C6631")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2HistoryChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2HistoryChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2HistoryChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2HttpHeadersCollectionIterator
' MIDL_INTERFACE("4212F3A7-0FBC-4C9C-8118-17ED6370C1B3")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2HttpHeadersCollectionIterator_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2HttpHeadersCollectionIterator_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2HttpHeadersCollectionIterator_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetCurrentHeader (BYVAL name AS LPWSTR PTR, BYVAL value AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HasCurrentHeader (BYVAL hasCurrent AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveNext (BYVAL hasNext AS WINBOOL PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2HttpRequestHeaders
' MIDL_INTERFACE("2C1F04DF-C90E-49E4-BD25-4A659300337B")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2HttpRequestHeaders_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2HttpRequestHeaders_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2HttpRequestHeaders_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetHeader (BYVAL name AS LPWSTR, BYVAL value AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetHeaders (BYVAL name AS LPCWSTR, BYVAL iterator AS Afx_ICoreWebView2HttpHeadersCollectionIterator PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Contains (BYVAL name AS LPCWSTR, BYVAL Contains AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetHeader (BYVAL name AS LPCWSTR, BYVAL value AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveHeader (BYVAL name AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIterator (BYVAL iterator AS Afx_ICoreWebView2HttpHeadersCollectionIterator PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2HttpResponseHeaders
' MIDL_INTERFACE("B5F6D4D5-1BFF-4869-85B8-158153017B04")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2HttpResponseHeaders_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2HttpResponseHeaders_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2HttpResponseHeaders_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION AppendHeader (BYVAL name AS LPCWSTR, BYVAL value AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Contains (BYVAL name AS LPCWSTR, BYVAL Contains AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetHeader (BYVAL name AS LPWSTR, BYVAL value AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetHeaders (BYVAL name AS LPCWSTR, BYVAL iterator AS Afx_ICoreWebView2HttpHeadersCollectionIterator PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIterator (BYVAL iterator AS Afx_ICoreWebView2HttpHeadersCollectionIterator PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2MoveFocusRequestedEventArgs
' MIDL_INTERFACE("71922903-B180-49D0-AED2-C9F9D10064B1")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2MoveFocusRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2MoveFocusRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2MoveFocusRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_reason (BYVAL reason AS COREWEBVIEW2_MOVE_FOCUS_REASON PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL value AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL value AS WINBOOL) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2MoveFocusRequestedEventHandler
' MIDL_INTERFACE("4B21D6DD-3DE7-47B0-8019-7D3ACE6E3631")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2MoveFocusRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2MoveFocusRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2MoveFocusRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS Afx_ICoreWebView2MoveFocusRequestedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2NavigationCompletedEventArgs
' MIDL_INTERFACE("361F5621-EA7F-4C55-95EC-3C5E6992EA4A")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NavigationCompletedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NavigationCompletedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationCompletedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsSuccess (BYVAL IsSuccess AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_WebErrorStatus (BYVAL WebErrorStatus AS COREWEBVIEW2_WEB_ERROR_STATUS PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NavigationId (BYVAL NavigationId AS UINT64 PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NavigationCompletedEventHandler
' MIDL_INTERFACE("9F921239-20C4-455F-9E3F-6047A50E248B")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx___ICoreWebView2NavigationCompletedEventHandler_INTERFACE_DEFINED__
#define __Afx___ICoreWebView2NavigationCompletedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationCompletedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NavigationCompletedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2NavigationStartingEventArgs
' MIDL_INTERFACE("EE1938CE-D385-4CB0-854B-F498F78C3D88")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NavigationStartingEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NavigationStartingEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationStartingEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsUserInitiated (BYVAL IsUserInitiated AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsRedirected (BYVAL IsRedirected AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RequestHeaders (BYVAL requestHeaders AS Afx_ICoreWebView2HttpRequestHeaders PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Cancel (BYVAL Cancel AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Cancel (BYVAL Cancel AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NavigationId (BYVAL NavigationId AS UINT64 PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2NavigationStartingEventHandler
' MIDL_INTERFACE("073337A4-64D2-4C7E-AC9F-987F0F613497")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NavigationStartingEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NavigationStartingEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationStartingEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NavigationStartingEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2NewBrowserVersionAvailableEventHandler
' MIDL_INTERFACE("E82E8242-EE39-4A57-A065-E13256D60342")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL webviewEnvironment AS Afx_ICoreWebView2Environment PTR, BYVAL args AS IUNknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2NewWindowRequestedEventArgs
' MIDL_INTERFACE("9EDC7F5F-C6EA-4F3C-827B-A8880794C0A9")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NewWindowRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NewWindowRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NewWindowRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_NewWindow (BYVAL newWindow AS Afx_ICoreWebView2 PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NewWindow (BYVAL newWindow AS Afx_ICoreWebView2 PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL Handled AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL Handled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsUserInitiated (BYVAL IsUserInitiated AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2NewWindowRequestedEventHandler
' MIDL_INTERFACE("ACAA30EF-A40C-47BD-9CB9-D9C2AADC9FCB")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NewWindowRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NewWindowRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NewWindowRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NewWindowRequestedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2PermissionRequestedEventArgs
' MIDL_INTERFACE("774B5EA1-3FAD-435C-B1FC-A77D1ACD5EAF")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PermissionRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PermissionRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PermissionRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PermissionKind (BYVAL value AS COREWEBVIEW2_PERMISSION_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsUserInitiated (BYVAL IsUserInitiated AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_State (BYVAL value AS COREWEBVIEW2_PERMISSION_STATE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_State (BYVAL value AS COREWEBVIEW2_PERMISSION_STATE) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2PermissionRequestedEventHandler
' MIDL_INTERFACE("543B4ADE-9B0B-4748-9AB7-D76481B223AA")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PermissionRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PermissionRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PermissionRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2PermissionRequestedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2ProcessFailedEventArgs
' MIDL_INTERFACE("EA45D1F4-75C0-471F-A6E9-803FBFF8FEF2")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessFailedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessFailedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessFailedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ProcessFailedKind (BYVAL ProcessFailedKind AS COREWEBVIEW2_PROCESS_FAILED_KIND PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2ProcessFailedEventHandler
' MIDL_INTERFACE("7D2183F9-CCA8-40F2-91A9-EAFAD32C8A9B")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessFailedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessFailedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessFailedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ProcessFailedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2ScriptDialogOpeningEventArgs
' MIDL_INTERFACE("B8F6356E-24DC-4D74-90FE-AD071E11CB91")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ScriptDialogOpeningEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ScriptDialogOpeningEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ScriptDialogOpeningEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Kind (BYVAL Kind AS COREWEBVIEW2_SCRIPT_DIALOG_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Message (BYVAL Message AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Accept () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DefaultText (BYVAL DefaultText AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResultText (BYVAL resultText AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ResultText (BYVAL resultText AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2ScriptDialogOpeningEventHandler
' MIDL_INTERFACE("72D93789-2727-4A9B-A4FC-1B2609CBCBE3")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ScriptDialogOpeningEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ScriptDialogOpeningEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ScriptDialogOpeningEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ScriptDialogOpeningEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2Settings
' MIDL_INTERFACE("203FBA37-6850-4DCC-A25A-58A351AC625D")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsScriptEnabled (BYVAL IsScriptEnabled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsScriptEnabled (BYVAL IsScriptEnabled AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsWebMessageEnabled (BYVAL IsWebMessageEnabled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsWebMessageEnabled (BYVAL IsWebMessageEnabled AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AreDefaultScriptDialogsEnabled (BYVAL AreDefaultScriptDialogsEnabled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreDefaultScriptDialogsEnabled (BYVAL AreDefaultScriptDialogsEnabled AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsStatusBarEnabled (BYVAL IsStatusBarEnabled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsStatusBarEnabled (BYVAL IsStatusBarEnabled AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AreDevToolsEnabled (BYVAL AreDevToolsEnabled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreDevToolsEnabled (BYVAL AreDevToolsEnabled AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AreDefaultContextMenusEnabled (BYVAL enabled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreDefaultContextMenusEnabled (BYVAL enabled AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AreHostObjectsAllowed (BYVAL allowed AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreHostObjectsAllowed (BYVAL allowed AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AreRemoteObjectsAllowed (BYVAL allowed AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreRemoteObjectsAllowed (BYVAL allowed AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsZoomControlEnabled (BYVAL enabled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsZoomControlEnabled (BYVAL enabled AS WINBOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsBuiltInErrorPageEnabled (BYVAL enabled AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsBuiltInErrorPageEnabled (BYVAL enabled AS WINBOOL) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2SourceChangedEventArgs
' MIDL_INTERFACE("BD9A4BFB-BE19-40BD-968B-EBCF0D727EF3")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2SourceChangedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2SourceChangedEventArgs_INTERFACE_DEFINED__

TYPE ICoreWebView2SourceChangedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsNewDocument (BYVAL isNewDocument AS WINBOOL PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2SourceChangedEventHandler
' MIDL_INTERFACE("8FEDD1A7-3A33-416F-AF81-881EEB001433")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2SourceChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2SourceChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2SourceChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2SourceChangedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2WebMessageReceivedEventArgs
' MIDL_INTERFACE("B263B5AE-9C54-4B75-B632-40AE1A0B6912")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebMessageReceivedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebMessageReceivedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebMessageReceivedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Source (BYVAL source AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_webMessageAsJson (BYVAL webMessageAsJson AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION TryGetWebMessageAsString (BYVAL webMessageAsString AS LPWSTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2WebMessageReceivedEventHandler
' MIDL_INTERFACE("199328C8-9964-4F5F-84E6-E875B1B763D6")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebMessageReceivedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebMessageReceivedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebMessageReceivedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2WebMessageReceivedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2WebResourceRequest
' MIDL_INTERFACE("11B02254-B827-49F6-8974-30F6E6C55AF6")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceRequest_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceRequest_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceRequest_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_uri (BYVAL uri AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Method (BYVAL Method AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Method (BYVAL Method AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Content (BYVAL Content AS IStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Content (BYVAL Content AS IStream PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Headers (BYVAL Headers AS Afx_ICoreWebView2HttpRequestHeaders PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2WebResourceRequestedEventArgs
' MIDL_INTERFACE("2D7B3282-83B1-41CA-8BBF-FF18F6BFE320")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Request (BYVAL request AS Afx_ICoreWebView2WebResourceRequest PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Response (BYVAL response AS Afx_ICoreWebView2WebResourceResponse PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Response (BYVAL response AS Afx_ICoreWebView2WebResourceResponse PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResourceContext (BYVAL context AS COREWEBVIEW2_WEB_RESOURCE_CONTEXT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2WebResourceRequestedEventHandler
' MIDL_INTERFACE("F6DC79F2-E1FA-4534-8968-4AFF10BBAA32")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2WebResourceRequestedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2WebResourceResponse
' MIDL_INTERFACE("5953D1FC-B08F-46DD-AFD3-66B172419CD0")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceResponse_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceResponse_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceResponse_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Content (BYVAL Content AS IStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Content (BYVAL Content AS IStream PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Headers (BYVAL Headers AS Afx_ICoreWebView2HttpResponseHeaders PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_StatusCode (BYVAL statusCode AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_StatusCode (BYVAL statusCode AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ReasonPhrase (BYVAL reasonPhrase AS LPWSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ReasonPhrase (BYVAL reasonPhrase AS LPCWSTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2WindowCloseRequestedEventHandler
' MIDL_INTERFACE("63C89928-AD32-4421-A0E4-EC99B34AA97E")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WindowCloseRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WindowCloseRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WindowCloseRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2ZoomFactorChangedEventHandler
' MIDL_INTERFACE("F1828246-8B98-4274-B708-ECDB6BF3843A")
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ZoomFactorChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ZoomFactorChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ZoomFactorChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

