' ########################################################################################
' Library name: WebView2
' GUID: {26D34152-879F-4065-BEA2-3DAA2CFADFB8}
' © 2025 by José Roca. All rights reserved. Freeware. Use at your own risk.
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

DIM SHARED AFX_IID_ICoreWebView2 AS GUID = (&h76ECEACB, &h0462, &h4D94, {&hAC, &h83, &h42, &h3A, &h67, &h93, &h77, &h5E})
DIM SHARED AFX_IID_ICoreWebView2_2 AS GUID = (&h9E8F0CF8, &hE670, &h4B5E, {&hB2, &hBC, &h73, &hE0, &h61, &hE3, &h18, &h4C})
DIM SHARED AFX_IID_ICoreWebView2_3 AS GUID = (&hA0D6DF20, &h3B92, &h416D, {&hAA, &h0C, &h43, &h7A, &h9C, &h72, &h78, &h57})
DIM SHARED AFX_IID_ICoreWebView2_4 AS GUID = (&h20D02D59, &h6DF2, &h42DC, {&hBD, &h06, &hF9, &h8A, &h69, &h4B, &h13, &h02})
DIM SHARED AFX_IID_ICoreWebView2_5 AS GUID = (&hBEDB11B8, &hD63C, &h11EB, {&hB8, &hBC, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2_6 AS GUID = (&h499AADAC, &hD92C, &h4589, {&h8A, &h75, &h11, &h1B, &hFC, &h16, &h77, &h95})
DIM SHARED AFX_IID_ICoreWebView2_7 AS GUID = (&h79C24D83, &h09A3, &h45AE, {&h94, &h18, &h48, &h7F, &h32, &hA5, &h87, &h40})
DIM SHARED AFX_IID_ICoreWebView2_8 AS GUID = (&hE9632730, &h6E1E, &h43AB, {&hB7, &hB8, &h7B, &h2C, &h9E, &h62, &hE0, &h94})
DIM SHARED AFX_IID_ICoreWebView2_9 AS GUID = (&h4D7B2EAB, &h9FDC, &h468D, {&hB9, &h98, &hA9, &h26, &h0B, &h5E, &hD6, &h51})
DIM SHARED AFX_IID_ICoreWebView2_10 AS GUID = (&hB1690564, &h6F5A, &h4983, {&h8E, &h48, &h31, &hD1, &h14, &h3F, &hEC, &hDB})
DIM SHARED AFX_IID_ICoreWebView2_11 AS GUID = (&h0BE78E56, &hC193, &h4051, {&hB9, &h43, &h23, &hB4, &h60, &hC0, &h8B, &hDB})
DIM SHARED AFX_IID_ICoreWebView2_12 AS GUID = (&h35D69927, &hBCFA, &h4566, {&h93, &h49, &h6B, &h3E, &h0D, &h15, &h4C, &hAC})
DIM SHARED AFX_IID_ICoreWebView2_13 AS GUID = (&hF75F09A8, &h667E, &h4983, {&h88, &hD6, &hC8, &h77, &h3F, &h31, &h5E, &h84})
DIM SHARED AFX_IID_ICoreWebView2_14 AS GUID = (&h6DAA4F10, &h4A90, &h4753, {&h88, &h98, &h77, &hC5, &hDF, &h53, &h41, &h65})
DIM SHARED AFX_IID_ICoreWebView2_15 AS GUID = (&h517B2D1D, &h7DAE, &h4A66, {&hA4, &hF4, &h10, &h35, &h2F, &hFB, &h95, &h18})
DIM SHARED AFX_IID_ICoreWebView2_16 AS GUID = (&h0EB34DC9, &h9F91, &h41E1, {&h86, &h39, &h95, &hCD, &h59, &h43, &h90, &h6B})
DIM SHARED AFX_IID_ICoreWebView2_17 AS GUID = (&h702E75D4, &hFD44, &h434D, {&h9D, &h70, &h1A, &h68, &hA6, &hB1, &h19, &h2A})
DIM SHARED AFX_IID_ICoreWebView2_18 AS GUID = (&h7A626017, &h28BE, &h49B2, {&hB8, &h65, &h3B, &hA2, &hB3, &h52, &h2D, &h90})
DIM SHARED AFX_IID_ICoreWebView2_19 AS GUID = (&h6921F954, &h79B0, &h437F, {&hA9, &h97, &hC8, &h58, &h11, &h89, &h7C, &h68})
DIM SHARED AFX_IID_ICoreWebView2_20 AS GUID = (&hB4BC1926, &h7305, &h11EE, {&hB9, &h62, &h02, &h42, &hAC, &h12, &h00, &h02})
DIM SHARED AFX_IID_ICoreWebView2_21 AS GUID = (&hC4980DEA, &h587B, &h43B9, {&h81, &h43, &h3E, &hF3, &hBF, &h55, &h2D, &h95})
DIM SHARED AFX_IID_ICoreWebView2_22 AS GUID = (&hDB75DFC7, &hA857, &h4632, {&hA3, &h98, &h69, &h69, &hDD, &hE2, &h6C, &h0A})
DIM SHARED AFX_IID_ICoreWebView2_23 AS GUID = (&h508F0DB5, &h90C4, &h5872, {&h90, &hA7, &h26, &h7A, &h91, &h37, &h75, &h02})
DIM SHARED AFX_IID_ICoreWebView2_24 AS GUID = (&h39A7AD55, &h4287, &h5CC1, {&h88, &hA1, &hC6, &hF4, &h58, &h59, &h38, &h24})
DIM SHARED AFX_IID_ICoreWebView2_25 AS GUID = (&hB5A86092, &hDF50, &h5B4F, {&hA1, &h7B, &h6C, &h8F, &h8B, &h40, &hB7, &h71})
DIM SHARED AFX_IID_ICoreWebView2_26 AS GUID = (&h806268B8, &hF897, &h5685, {&h88, &hE5, &hC4, &h5F, &hCA, &h0B, &h1A, &h48})
DIM SHARED AFX_IID_ICoreWebView2_27 AS GUID = (&h00FBE33B, &h8C07, &h517C, {&hAA, &h23, &h0D, &hDD, &h4B, &h5F, &h6F, &hA0})
DIM SHARED AFX_IID_ICoreWebView2_28 AS GUID = (&h62E50381, &h5BF5, &h51A8, {&hAA, &hE0, &hF2, &h0A, &h3A, &h9C, &h8A, &h90})

DIM SHARED AFX_IID_ICoreWebView2Environment AS GUID = (&hB96D755E, &h0319, &h4E92, {&hA2, &h96, &h23, &h43, &h6F, &h46, &hA1, &hFC})
DIM SHARED AFX_IID_ICoreWebView2Environment2 AS GUID = (&h41F3632B, &h5EF4, &h404F, {&hAD, &h82, &h2D, &h60, &h6C, &h5A, &h9A, &h21})
DIM SHARED AFX_IID_ICoreWebView2Environment3 AS GUID = (&h80A22AE3, &hBE7C, &h4CE2, {&hAF, &hE1, &h5A, &h50, &h05, &h6C, &hDE, &hEB})
DIM SHARED AFX_IID_ICoreWebView2Environment4 AS GUID = (&h20944379, &h6DCF, &h41D6, {&hA0, &hA0, &hAB, &hC0, &hFC, &h50, &hDE, &h0D})
DIM SHARED AFX_IID_ICoreWebView2Environment5 AS GUID = (&h319E423D, &hE0D7, &h4B8D, {&h92, &h54, &hAE, &h94, &h75, &hDE, &h9B, &h17})
DIM SHARED AFX_IID_ICoreWebView2Environment6 AS GUID = (&hE59EE362, &hACBD, &h4857, {&h9A, &h8E, &hD3, &h64, &h4D, &h94, &h59, &hA9})
DIM SHARED AFX_IID_ICoreWebView2Environment7 AS GUID = (&h43C22296, &h3BBD, &h43A4, {&h9C, &h00, &h5C, &h0D, &hF6, &hDD, &h29, &hA2})
DIM SHARED AFX_IID_ICoreWebView2Environment8 AS GUID = (&hD6EB91DD, &hC3D2, &h45E5, {&hBD, &h29, &h6D, &hC2, &hBC, &h4D, &hE9, &hCF})
DIM SHARED AFX_IID_ICoreWebView2Environment9 AS GUID = (&hF06F41BF, &h4B5A, &h49D8, {&hB9, &hF6, &hFA, &h16, &hCD, &h29, &hF2, &h74})
DIM SHARED AFX_IID_ICoreWebView2Environment10 AS GUID = (&hEE0EB9DF, &h6F12, &h46CE, {&hB5, &h3F, &h3F, &h47, &hB9, &hC9, &h28, &hE0})
DIM SHARED AFX_IID_ICoreWebView2Environment11 AS GUID = (&hF0913DC6, &hA0EC, &h42EF, {&h98, &h05, &h91, &hDF, &hF3, &hA2, &h96, &h6A})
DIM SHARED AFX_IID_ICoreWebView2Environment12 AS GUID = (&hF503DB9B, &h739F, &h48DD, {&hB1, &h51, &hFD, &hFC, &hF2, &h53, &hF5, &h4E})
DIM SHARED AFX_IID_ICoreWebView2Environment13 AS GUID = (&hAF641F58, &h72B2, &h11EE, {&hB9, &h62, &h02, &h42, &hAC, &h12, &h00, &h02})
DIM SHARED AFX_IID_ICoreWebView2Environment14 AS GUID = (&hA5E9FAD9, &hC875, &h59DA, {&h9B, &hD7, &h47, &h3A, &hA5, &hCA, &h1C, &hEF})
DIM SHARED AFX_IID_ICoreWebView2Environment15 AS GUID = (&h2AC5EBFB, &hE654, &h5961, {&hA6, &h67, &h79, &h71, &h88, &h5C, &h7B, &h27})

DIM SHARED AFX_IID_ICoreWebView2EnvironmentOptions AS GUID = (&h2FDE08A8, &h1E9A, &h4766, {&h8C, &h05, &h95, &hA9, &hCE, &hB9, &hD1, &hC5})
DIM SHARED AFX_IID_ICoreWebView2EnvironmentOptions2 AS GUID = (&hFF85C98A, &h1BA7, &h4A6B, {&h90, &hC8, &h2B, &h75, &h2C, &h89, &hE9, &hE2})
DIM SHARED AFX_IID_ICoreWebView2EnvironmentOptions3 AS GUID = (&h4A5C436E, &hA9E3, &h4A2E, {&h89, &hC3, &h91, &h0D, &h35, &h13, &hF5, &hCC})
DIM SHARED AFX_IID_ICoreWebView2EnvironmentOptions4 AS GUID = (&hAC52D13F, &h0D38, &h475A, {&h9D, &hCA, &h87, &h65, &h80, &hD6, &h79, &h3E})
DIM SHARED AFX_IID_ICoreWebView2EnvironmentOptions5 AS GUID = (&h0AE35D64, &hC47F, &h4464, {&h81, &h4E, &h25, &h9C, &h34, &h5D, &h15, &h01})
DIM SHARED AFX_IID_ICoreWebView2EnvironmentOptions6 AS GUID = (&h57D29CC3, &hC84F, &h42A0, {&hB0, &hE2, &hEF, &hFB, &hD5, &hE1, &h79, &hDE})
DIM SHARED AFX_IID_ICoreWebView2EnvironmentOptions7 AS GUID = (&hC48D539F, &hE39F, &h441C, {&hAE, &h68, &h1F, &h66, &hE5, &h70, &hBD, &hC5})
DIM SHARED AFX_IID_ICoreWebView2EnvironmentOptions8 AS GUID = (&h7C7ECF51, &hE918, &h5CAF, {&h85, &h3C, &hE9, &hA2, &hBC, &hC2, &h77, &h75})

DIM SHARED AFX_IID_ICoreWebView2Controller AS GUID = (&h4D00C0D1, &h9434, &h4EB6, {&h80, &h78, &h86, &h97, &hA5, &h60, &h33, &h4F})
DIM SHARED AFX_IID_ICoreWebView2Controller2 AS GUID = (&hC979903E, &hD4CA, &h4228, {&h92, &hEB, &h47, &hEE, &h3F, &hA9, &h6E, &hAB})
DIM SHARED AFX_IID_ICoreWebView2Controller3 AS GUID = (&hF9614724, &h5D2B, &h41DC, {&hAE, &hF7, &h73, &hD6, &h2B, &h51, &h54, &h3B})
DIM SHARED AFX_IID_ICoreWebView2Controller4 AS GUID = (&h97D418D5, &hA426, &h4E49, {&hA1, &h51, &hE1, &hA1, &h0F, &h32, &h7D, &h9E})
DIM SHARED AFX_IID_ICoreWebView2ControllerOptions AS GUID = (&h12AAE616, &h8CCB, &h44EC, {&hBC, &hB3, &hEB, &h18, &h31, &h88, &h16, &h35})
DIM SHARED AFX_IID_ICoreWebView2ControllerOptions2 AS GUID = (&h06C991D8, &h9E7E, &h11ED, {&hA8, &hFC, &h02, &h42, &hAC, &h12, &h00, &h02})
DIM SHARED AFX_IID_ICoreWebView2ControllerOptions3 AS GUID = (&hB32B191A, &h8998, &h57CA, {&hB7, &hCB, &hE0, &h46, &h17, &hE4, &hCE, &h4A})
DIM SHARED AFX_IID_ICoreWebView2ControllerOptions4 AS GUID = (&h21EB052F, &hAD39, &h555E, {&h82, &h4A, &hC8, &h7B, &h09, &h1D, &h4D, &h36})

DIM SHARED AFX_IID_ICoreWebView2Frame AS GUID = (&hF1131A5E, &h9BA9, &h11EB, {&hA8, &hB3, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2Frame2 AS GUID = (&h7A6A5834, &hD185, &h4DBF, {&hB6, &h3F, &h4A, &h9B, &hC4, &h31, &h07, &hD4})
DIM SHARED AFX_IID_ICoreWebView2Frame3 AS GUID = (&hB50D82CC, &hCC28, &h481D, {&h96, &h14, &hCB, &h04, &h88, &h95, &hE6, &hA0})
DIM SHARED AFX_IID_ICoreWebView2Frame4 AS GUID = (&h188782DC, &h92AA, &h4732, {&hAB, &h3C, &hFC, &hC5, &h9F, &h6F, &h68, &hB9})
DIM SHARED AFX_IID_ICoreWebView2Frame5 AS GUID = (&h99D199C4, &h7305, &h11EE, {&hB9, &h62, &h02, &h42, &hAC, &h12, &h00, &h02})
DIM SHARED AFX_IID_ICoreWebView2Frame6 AS GUID = (&h0DE611FD, &h31E9, &h5DDC, {&h9D, &h71, &h95, &hED, &hA2, &h6E, &hFF, &h32})
DIM SHARED AFX_IID_ICoreWebView2Frame7 AS GUID = (&h3598CFA2, &hD85D, &h5A9F, {&h92, &h28, &h4D, &hDE, &h1F, &h59, &hEC, &h64})

DIM SHARED AFX_IID_ICoreWebView2CompositionController AS GUID = (&h3DF9B733, &hB9AE, &h4A15, {&h86, &hB4, &hEB, &h9E, &hE9, &h82, &h64, &h69})
DIM SHARED AFX_IID_ICoreWebView2CompositionController2 AS GUID = (&h0B6A3D24, &h49CB, &h4806, {&hBA, &h20, &hB5, &hE0, &h73, &h4A, &h7B, &h26})
DIM SHARED AFX_IID_ICoreWebView2CompositionController3 AS GUID = (&h9570570E, &h4D76, &h4361, {&h9E, &hE1, &hF0, &h4D, &h0D, &hBD, &hFB, &h1E})
DIM SHARED AFX_IID_ICoreWebView2CompositionController4 AS GUID = (&h7C367B9B, &h3D2B, &h450F, {&h9E, &h58, &hD6, &h1A, &h20, &hF4, &h86, &hAA})

DIM SHARED AFX_IID_ICoreWebView2NavigationCompletedEventArgs AS GUID = (&h30D68B7D, &h20D9, &h4752, {&hA9, &hCA, &hEC, &h84, &h48, &hFB, &hB5, &hC1})
DIM SHARED AFX_IID_ICoreWebView2NavigationCompletedEventArgs2 AS GUID = (&hFDF8B738, &hEE1E, &h4DB2, {&hA3, &h29, &h8D, &h7D, &h7B, &h74, &hD7, &h92})
DIM SHARED AFX_IID_ICoreWebView2NavigationStartingEventArgs AS GUID = (&h5B495469, &hE119, &h438A, {&h9B, &h18, &h76, &h04, &hF2, &h5F, &h2E, &h49})
DIM SHARED AFX_IID_ICoreWebView2NavigationStartingEventArgs2 AS GUID = (&h9086BE93, &h91AA, &h472D, {&hA7, &hE0, &h57, &h9F, &h2B, &hA0, &h06, &hAD})
DIM SHARED AFX_IID_ICoreWebView2NavigationStartingEventArgs3 AS GUID = (&hDDFFE494, &h4942, &h4BD2, {&hAB, &h73, &h35, &hB8, &hFF, &h40, &hE1, &h9F})
DIM SHARED AFX_IID_ICoreWebView2NewWindowRequestedEventArgs AS GUID = (&h34ACB11C, &hFC37, &h4418, {&h91, &h32, &hF9, &hC2, &h1D, &h1E, &hAF, &hB9})
DIM SHARED AFX_IID_ICoreWebView2NewWindowRequestedEventArgs2 AS GUID = (&hBBC7BAED, &h74C6, &h4C92, {&hB6, &h3A, &h7F, &h5A, &hEA, &hE0, &h3D, &hE3})
DIM SHARED AFX_IID_ICoreWebView2NewWindowRequestedEventArgs3 AS GUID = (&h842BED3C, &h6AD6, &h4DD9, {&hB9, &h38, &h28, &hC9, &h66, &h67, &hAD, &h66})

DIM SHARED AFX_IID_ICoreWebView2PermissionRequestedEventArgs AS GUID = (&h973AE2EF, &hFF18, &h4894, {&h8F, &hB2, &h3C, &h75, &h8F, &h04, &h68, &h10})
DIM SHARED AFX_IID_ICoreWebView2PermissionRequestedEventArgs2 AS GUID = (&h74D7127F, &h9DE6, &h4200, {&h87, &h34, &h42, &hD6, &hFB, &h4F, &hF7, &h41})
DIM SHARED AFX_IID_ICoreWebView2PermissionRequestedEventArgs3 AS GUID = (&hE61670BC, &h3DCE, &h4177, {&h86, &hD2, &hC6, &h29, &hAE, &h3C, &hB6, &hAC})
DIM SHARED AFX_IID_ICoreWebView2PermissionRequestedEventHandler AS GUID = (&h15E1C6A3, &hC72A, &h4DF3, {&h91, &hD7, &hD0, &h97, &hFB, &hEC, &h6B, &hFD})

DIM SHARED AFX_IID_ICoreWebView2ProcessFailedEventArgs AS GUID = (&h8155A9A4, &h1474, &h4A86, {&h8C, &hAE, &h15, &h1B, &h0F, &hA6, &hB8, &hCA})
DIM SHARED AFX_IID_ICoreWebView2ProcessFailedEventArgs2 AS GUID = (&h4DAB9422, &h46FA, &h4C3E, {&hA5, &hD2, &h41, &hD2, &h07, &h1D, &h36, &h80})
DIM SHARED AFX_IID_ICoreWebView2ProcessFailedEventArgs3 AS GUID = (&hAB667428, &h094D, &h5FD1, {&hB4, &h80, &h8B, &h4C, &h0F, &hDB, &hDF, &h2F})
DIM SHARED AFX_IID_ICoreWebView2ProcessFailedEventHandler AS GUID = (&h79E0AEA4, &h990B, &h42D9, {&hAA, &h1D, &h0F, &hCC, &h2E, &h5B, &hC7, &hF1})

DIM SHARED AFX_IID_ICoreWebView2Profile AS GUID = (&h79110AD3, &hCD5D, &h4373, {&h8B, &hC3, &hC6, &h06, &h58, &hF1, &h7A, &h5F})
DIM SHARED AFX_IID_ICoreWebView2Profile2 AS GUID = (&hFA740D4B, &h5EAE, &h4344, {&hA8, &hAD, &h74, &hBE, &h31, &h92, &h53, &h97})
DIM SHARED AFX_IID_ICoreWebView2Profile3 AS GUID = (&hB188E659, &h5685, &h4E05, {&hBD, &hBA, &hFC, &h64, &h0E, &h0F, &h19, &h92})
DIM SHARED AFX_IID_ICoreWebView2Profile4 AS GUID = (&h8F4AE680, &h192E, &h4EC8, {&h83, &h3A, &h21, &hCF, &hAD, &hAE, &hF6, &h28})
DIM SHARED AFX_IID_ICoreWebView2Profile5 AS GUID = (&h2EE5B76E, &h6E80, &h4DF2, {&hBC, &hD3, &hD4, &hEC, &h33, &h40, &hA0, &h1B})
DIM SHARED AFX_IID_ICoreWebView2Profile6 AS GUID = (&hBD82FA6A, &h1D65, &h4C33, {&hB2, &hB4, &h03, &h93, &h02, &h0C, &hC6, &h1B})
DIM SHARED AFX_IID_ICoreWebView2Profile7 AS GUID = (&h7B4C7906, &hA1AA, &h4CB4, {&hB7, &h23, &hDB, &h09, &hF8, &h13, &hD5, &h41})
DIM SHARED AFX_IID_ICoreWebView2Profile8 AS GUID = (&hFBF70C2F, &hEB1F, &h4383, {&h85, &hA0, &h16, &h3E, &h92, &h04, &h40, &h11})

DIM SHARED AFX_IID_ICoreWebView2Settings AS GUID = (&hE562E4F0, &hD7FA, &h43AC, {&h8D, &h71, &hC0, &h51, &h50, &h49, &h9F, &h00})
DIM SHARED AFX_IID_ICoreWebView2Settings2 AS GUID = (&hEE9A0F68, &hF46C, &h4E32, {&hAC, &h23, &hEF, &h8C, &hAC, &h22, &h4D, &h2A})
DIM SHARED AFX_IID_ICoreWebView2Settings3 AS GUID = (&hFDB5AB74, &hAF33, &h4854, {&h84, &hF0, &h0A, &h63, &h1D, &hEB, &h5E, &hBA})
DIM SHARED AFX_IID_ICoreWebView2Settings4 AS GUID = (&hCB56846C, &h4168, &h4D53, {&hB0, &h4F, &h03, &hB6, &hD6, &h79, &h6F, &hF2})
DIM SHARED AFX_IID_ICoreWebView2Settings5 AS GUID = (&h183E7052, &h1D03, &h43A0, {&hAB, &h99, &h98, &hE0, &h43, &hB6, &h6B, &h39})
DIM SHARED AFX_IID_ICoreWebView2Settings6 AS GUID = (&h11CB3ACD, &h9BC8, &h43B8, {&h83, &hBF, &hF4, &h07, &h53, &h71, &h4F, &h87})
DIM SHARED AFX_IID_ICoreWebView2Settings7 AS GUID = (&h488DC902, &h35EF, &h42D2, {&hBC, &h7D, &h94, &hB6, &h5C, &h4B, &hC4, &h9C})
DIM SHARED AFX_IID_ICoreWebView2Settings8 AS GUID = (&h9E6B0E8F, &h86AD, &h4E81, {&h81, &h47, &hA9, &hB5, &hED, &hB6, &h86, &h50})
DIM SHARED AFX_IID_ICoreWebView2Settings9 AS GUID = (&h0528A73B, &hE92D, &h49F4, {&h92, &h7A, &hE5, &h47, &hDD, &hDA, &hA3, &h7D})

DIM SHARED AFX_IID_ICoreWebView2WebResourceRequestedEventArgs AS GUID = (&h453E667F, &h12C7, &h49D4, {&hBE, &h6D, &hDD, &hBE, &h79, &h56, &hF5, &h7A})
DIM SHARED AFX_IID_ICoreWebView2WebResourceRequestedEventArgs2 AS GUID = (&h9C562C24, &hB219, &h4D7F, {&h92, &hF6, &hB1, &h87, &hFB, &hBA, &hDD, &h56})
DIM SHARED AFX_IID_ICoreWebView2WebMessageReceivedEventArgs AS GUID = (&h0F99A40C, &hE962, &h4207, {&h9E, &h92, &hE3, &hD5, &h42, &hEF, &hF8, &h49})
DIM SHARED AFX_IID_ICoreWebView2WebMessageReceivedEventArgs2 AS GUID = (&h06FC7AB7, &hC90C, &h4297, {&h93, &h89, &h33, &hCA, &h01, &hCF, &h6D, &h5E})

DIM SHARED AFX_IID_ICoreWebView2AcceleratorKeyPressedEventArgs AS GUID = (&h9F760F8A, &hFB79, &h42BE, {&h99, &h90, &h7B, &h56, &h90, &h0F, &hA9, &hC7})
DIM SHARED AFX_IID_ICoreWebView2AcceleratorKeyPressedEventArgs2 AS GUID = (&h03B2C8C8, &h7799, &h4E34, {&hBD, &h66, &hED, &h26, &hAA, &h85, &hF2, &hBF})
DIM SHARED AFX_IID_ICoreWebView2AcceleratorKeyPressedEventHandler AS GUID = (&hB29C7E28, &hFA79, &h41A8, {&h8E, &h44, &h65, &h81, &h1C, &h76, &hDC, &hB2})
DIM SHARED AFX_IID_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler AS GUID = (&hB99369F3, &h9B11, &h47B5, {&hBC, &h6F, &h8E, &h78, &h95, &hFC, &hEA, &h17})
DIM SHARED AFX_IID_ICoreWebView2BasicAuthenticationRequestedEventArgs AS GUID = (&hEF05516F, &hD897, &h4F9E, {&hB6, &h72, &hD8, &hE2, &h30, &h7A, &h3F, &hB0})
DIM SHARED AFX_IID_ICoreWebView2BasicAuthenticationRequestedEventHandler AS GUID = (&h58B4D6C2, &h18D4, &h497E, {&hB3, &h9B, &h9A, &h96, &h53, &h3F, &hA2, &h78})
DIM SHARED AFX_IID_ICoreWebView2BasicAuthenticationResponse AS GUID = (&h07023F7D, &h2D77, &h4D67, {&h90, &h40, &h6E, &h7D, &h42, &h8C, &h6A, &h40})
DIM SHARED AFX_IID_ICoreWebView2BrowserExtension AS GUID = (&h7EF7FFA0, &hFAC5, &h462C, {&hB1, &h89, &h3D, &h9E, &hDB, &hE5, &h75, &hDA})
DIM SHARED AFX_IID_ICoreWebView2BrowserExtensionEnableCompletedHandler AS GUID = (&h30C186CE, &h7FAD, &h421F, {&hA3, &hBC, &hA8, &hEA, &hF0, &h71, &hDD, &hB8})
DIM SHARED AFX_IID_ICoreWebView2BrowserExtensionList AS GUID = (&h2EF3D2DC, &hBD5F, &h4F4D, {&h90, &hAF, &hFD, &h67, &h79, &h8F, &h0C, &h2F})
DIM SHARED AFX_IID_ICoreWebView2BrowserExtensionRemoveCompletedHandler AS GUID = (&h8E41909A, &h9B18, &h4BB1, {&h8C, &hDF, &h93, &h0F, &h46, &h7A, &h50, &hBE})
DIM SHARED AFX_IID_ICoreWebView2BrowserProcessExitedEventArgs AS GUID = (&h1F00663F, &hAF8C, &h4782, {&h9C, &hDD, &hDD, &h01, &hC5, &h2E, &h34, &hCB})
DIM SHARED AFX_IID_ICoreWebView2BrowserProcessExitedEventHandler AS GUID = (&hFA504257, &hA216, &h4911, {&hA8, &h60, &hFE, &h88, &h25, &h71, &h28, &h61})
DIM SHARED AFX_IID_ICoreWebView2BytesReceivedChangedEventHandler AS GUID = (&h828E8AB6, &hD94C, &h4264, {&h9C, &hEF, &h52, &h17, &h17, &h0D, &h62, &h51})
DIM SHARED AFX_IID_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler AS GUID = (&h5C4889F0, &h5EF6, &h4C5A, {&h95, &h2C, &hD8, &hF1, &hB9, &h2D, &h05, &h74})
DIM SHARED AFX_IID_ICoreWebView2CapturePreviewCompletedHandler AS GUID = (&h697E05E9, &h3D8F, &h45FA, {&h96, &hF4, &h8F, &hFE, &h1E, &hDE, &hDA, &hF5})
DIM SHARED AFX_IID_ICoreWebView2Certificate AS GUID = (&hC5FB2FCE, &h1CAC, &h4AEE, {&h9C, &h79, &h5E, &hD0, &h36, &h2E, &hAA, &hE0})
DIM SHARED AFX_IID_ICoreWebView2ClearBrowsingDataCompletedHandler AS GUID = (&hE9710A06, &h1D1D, &h49B2, {&h82, &h34, &h22, &h6F, &h35, &h84, &h6A, &hE5})
DIM SHARED AFX_IID_ICoreWebView2ClearServerCertificateErrorActionsCompletedHandler AS GUID = (&h3B40AAC6, &hACFE, &h4FFD, {&h82, &h11, &hF6, &h07, &hB9, &h6E, &h2D, &h5B})
DIM SHARED AFX_IID_ICoreWebView2ClientCertificate AS GUID = (&hE7188076, &hBCC3, &h11EB, {&h85, &h29, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2ClientCertificateCollection AS GUID = (&hEF5674D2, &hBCC3, &h11EB, {&h85, &h29, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2ClientCertificateRequestedEventArgs AS GUID = (&hBC59DB28, &hBCC3, &h11EB, {&h85, &h29, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2ClientCertificateRequestedEventHandler AS GUID = (&hD7175BA2, &hBCC3, &h11EB, {&h85, &h29, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2ContainsFullScreenElementChangedEventHandler AS GUID = (&hE45D98B1, &hAFEF, &h45BE, {&h8B, &hAF, &h6C, &h77, &h28, &h86, &h7F, &h73})
DIM SHARED AFX_IID_ICoreWebView2ContentLoadingEventArgs AS GUID = (&h0C8A1275, &h9B6B, &h4901, {&h87, &hAD, &h70, &hDF, &h25, &hBA, &hFA, &h6E})
DIM SHARED AFX_IID_ICoreWebView2ContentLoadingEventHandler AS GUID = (&h364471E7, &hF2BE, &h4910, {&hBD, &hBA, &hD7, &h20, &h77, &hD5, &h1C, &h4B})
DIM SHARED AFX_IID_ICoreWebView2ContextMenuItem AS GUID = (&h7AED49E3, &hA93F, &h497A, {&h81, &h1C, &h74, &h9C, &h6B, &h6B, &h6C, &h65})
DIM SHARED AFX_IID_ICoreWebView2ContextMenuItemCollection AS GUID = (&hF562A2F5, &hC415, &h45CF, {&hB9, &h09, &hD4, &hB7, &hC1, &hE2, &h76, &hD3})
DIM SHARED AFX_IID_ICoreWebView2ContextMenuRequestedEventArgs AS GUID = (&hA1D309EE, &hC03F, &h11EB, {&h85, &h29, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2ContextMenuRequestedEventHandler AS GUID = (&h04D3FE1D, &hAB87, &h42FB, {&hA8, &h98, &hDA, &h24, &h1D, &h35, &hB6, &h3C})
DIM SHARED AFX_IID_ICoreWebView2ContextMenuTarget AS GUID = (&hB8611D99, &hEED6, &h4F3F, {&h90, &h2C, &hA1, &h98, &h50, &h2A, &hD4, &h72})
DIM SHARED AFX_IID_ICoreWebView2Cookie AS GUID = (&hAD26D6BE, &h1486, &h43E6, {&hBF, &h87, &hA2, &h03, &h40, &h06, &hCA, &h21})
DIM SHARED AFX_IID_ICoreWebView2CookieList AS GUID = (&hF7F6F714, &h5D2A, &h43C6, {&h95, &h03, &h34, &h6E, &hCE, &h02, &hD1, &h86})
DIM SHARED AFX_IID_ICoreWebView2CookieManager AS GUID = (&h177CD9E7, &hB6F5, &h451A, {&h94, &hA0, &h5D, &h7A, &h3A, &h4C, &h41, &h41})
DIM SHARED AFX_IID_ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler AS GUID = (&h02FAB84B, &h1428, &h4FB7, {&hAD, &h45, &h1B, &h2E, &h64, &h73, &h61, &h84})
DIM SHARED AFX_IID_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler AS GUID = (&h6C4819F3, &hC9B7, &h4260, {&h81, &h27, &hC9, &hF5, &hBD, &hE7, &hF6, &h8C})
DIM SHARED AFX_IID_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler AS GUID = (&h4E8A3389, &hC9D8, &h4BD2, {&hB6, &hB5, &h12, &h4F, &hEE, &h6C, &hC1, &h4D})
DIM SHARED AFX_IID_ICoreWebView2CursorChangedEventHandler AS GUID = (&h9DA43CCC, &h26E1, &h4DAD, {&hB5, &h6C, &hD8, &h96, &h1C, &h94, &hC5, &h71})
DIM SHARED AFX_IID_ICoreWebView2CustomItemSelectedEventHandler AS GUID = (&h49E1D0BC, &hFE9E, &h4481, {&hB7, &hC2, &h32, &h32, &h4A, &hA2, &h19, &h98})
DIM SHARED AFX_IID_ICoreWebView2CustomSchemeRegistration AS GUID = (&hD60AC92C, &h37A6, &h4B26, {&hA3, &h9E, &h95, &hCF, &hE5, &h90, &h47, &hBB})
DIM SHARED AFX_IID_ICoreWebView2Deferral AS GUID = (&hC10E7F7B, &hB585, &h46F0, {&hA6, &h23, &h8B, &hEF, &hBF, &h3E, &h4E, &hE0})
DIM SHARED AFX_IID_ICoreWebView2DevToolsProtocolEventReceivedEventArgs AS GUID = (&h653C2959, &hBB3A, &h4377, {&h86, &h32, &hB5, &h8A, &hDA, &h4E, &h66, &hC4})
DIM SHARED AFX_IID_ICoreWebView2DevToolsProtocolEventReceivedEventArgs2 AS GUID = (&h2DC4959D, &h1494, &h4393, {&h95, &hBA, &hBE, &hA4, &hCB, &h9E, &hBD, &h1B})
DIM SHARED AFX_IID_ICoreWebView2DevToolsProtocolEventReceivedEventHandler AS GUID = (&hE2FDA4BE, &h5456, &h406C, {&hA2, &h61, &h3D, &h45, &h21, &h38, &h36, &h2C})
DIM SHARED AFX_IID_ICoreWebView2DevToolsProtocolEventReceiver AS GUID = (&hB32CA51A, &h8371, &h45E9, {&h93, &h17, &hAF, &h02, &h1D, &h08, &h03, &h67})
DIM SHARED AFX_IID_ICoreWebView2DocumentTitleChangedEventHandler AS GUID = (&hF5F2B923, &h953E, &h4042, {&h9F, &h95, &hF3, &hA1, &h18, &hE1, &hAF, &hD4})
DIM SHARED AFX_IID_ICoreWebView2DOMContentLoadedEventArgs AS GUID = (&h16B1E21A, &hC503, &h44F2, {&h84, &hC9, &h70, &hAB, &hA5, &h03, &h12, &h83})
DIM SHARED AFX_IID_ICoreWebView2DOMContentLoadedEventHandler AS GUID = (&h4BAC7E9C, &h199E, &h49ED, {&h87, &hED, &h24, &h93, &h03, &hAC, &hF0, &h19})
DIM SHARED AFX_IID_ICoreWebView2DownloadOperation AS GUID = (&h3D6B6CF2, &hAFE1, &h44C7, {&hA9, &h95, &hC6, &h51, &h17, &h71, &h43, &h36})
DIM SHARED AFX_IID_ICoreWebView2DownloadStartingEventArgs AS GUID = (&hE99BBE21, &h43E9, &h4544, {&hA7, &h32, &h28, &h27, &h64, &hEA, &hFA, &h60})
DIM SHARED AFX_IID_ICoreWebView2DownloadStartingEventHandler AS GUID = (&hEFEDC989, &hC396, &h41CA, {&h83, &hF7, &h07, &hF8, &h45, &hA5, &h57, &h24})
DIM SHARED AFX_IID_ICoreWebView2EstimatedEndTimeChangedEventHandler AS GUID = (&h28F0D425, &h93FE, &h4E63, {&h9F, &h8D, &h2A, &hEE, &hC6, &hD3, &hBA, &h1E})
DIM SHARED AFX_IID_ICoreWebView2ExecuteScriptCompletedHandler AS GUID = (&h49511172, &hCC67, &h4BCA, {&h99, &h23, &h13, &h71, &h12, &hF4, &hC4, &hCC})
DIM SHARED AFX_IID_ICoreWebView2ExecuteScriptResult AS GUID = (&h0CE15963, &h3698, &h4DF7, {&h93, &h99, &h71, &hED, &h6C, &hDD, &h8C, &h9F})
DIM SHARED AFX_IID_ICoreWebView2ExecuteScriptWithResultCompletedHandler AS GUID = (&h1BB5317B, &h8238, &h4C67, {&hA7, &hFF, &hBA, &hF6, &h55, &h8F, &h28, &h9D})
DIM SHARED AFX_IID_ICoreWebView2FaviconChangedEventHandler AS GUID = (&h2913DA94, &h833D, &h4DE0, {&h8D, &hCA, &h90, &h0F, &hC5, &h24, &hA1, &hA4})
DIM SHARED AFX_IID_ICoreWebView2File AS GUID = (&hF2C19559, &h6BC1, &h4583, {&hA7, &h57, &h90, &h02, &h1B, &hE9, &hAF, &hEC})
DIM SHARED AFX_IID_ICoreWebView2FileSystemHandle AS GUID = (&hC65100AC, &h0DE2, &h5551, {&hA3, &h62, &h23, &hD9, &hBD, &h1D, &h0E, &h1F})
DIM SHARED AFX_IID_ICoreWebView2Find AS GUID = (&hA3EC0F5F, &hDDBC, &h54ED, {&h85, &h46, &hAF, &h75, &hA7, &h85, &hB9, &hA6})
DIM SHARED AFX_IID_ICoreWebView2FindActiveMatchIndexChangedEventHandler AS GUID = (&h0054F514, &h9A8E, &h5876, {&hAE, &hD5, &h30, &hB3, &h7F, &h8C, &h86, &hA5})
DIM SHARED AFX_IID_ICoreWebView2FindMatchCountChangedEventHandler AS GUID = (&hDA0D6827, &h4254, &h5B10, {&hA6, &hD9, &h41, &h20, &h76, &hAF, &hC9, &hF3})
DIM SHARED AFX_IID_ICoreWebView2FindOptions AS GUID = (&hE82E3B2B, &hA4AF, &h5BC6, {&h94, &hC6, &h18, &hB4, &h41, &h57, &hA1, &h6C})
DIM SHARED AFX_IID_ICoreWebView2FindStartCompletedHandler AS GUID = (&h6A90ECAF, &h44B0, &h5BD9, {&h8F, &h07, &h19, &h67, &hE1, &h7B, &hE9, &hFB})
DIM SHARED AFX_IID_ICoreWebView2FocusChangedEventHandler AS GUID = (&h05EA24BD, &h6452, &h4926, {&h90, &h14, &h4B, &h82, &hB4, &h98, &h13, &h5D})
DIM SHARED AFX_IID_ICoreWebView2FrameChildFrameCreatedEventHandler AS GUID = (&h569E40E7, &h46B7, &h563D, {&h83, &hAE, &h10, &h73, &h15, &h56, &h64, &hD7})
DIM SHARED AFX_IID_ICoreWebView2FrameContentLoadingEventHandler AS GUID = (&h0D6156F2, &hD332, &h49A7, {&h9E, &h03, &h7D, &h8F, &h2F, &hEE, &hEE, &h54})
DIM SHARED AFX_IID_ICoreWebView2FrameCreatedEventArgs AS GUID = (&h4D6E7B5E, &h9BAA, &h11EB, {&hA8, &hB3, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2FrameCreatedEventHandler AS GUID = (&h38059770, &h9BAA, &h11EB, {&hA8, &hB3, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2FrameDestroyedEventHandler AS GUID = (&h59DD7B4C, &h9BAA, &h11EB, {&hA8, &hB3, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2FrameDOMContentLoadedEventHandler AS GUID = (&h38D9520D, &h340F, &h4D1E, {&hA7, &h75, &h43, &hFC, &hE9, &h75, &h36, &h83})
DIM SHARED AFX_IID_ICoreWebView2FrameInfo AS GUID = (&hDA86B8A1, &hBDF3, &h4F11, {&h99, &h55, &h52, &h8C, &hEF, &hA5, &h97, &h27})
DIM SHARED AFX_IID_ICoreWebView2FrameInfo2 AS GUID = (&h56F85CFA, &h72C4, &h11EE, {&hB9, &h62, &h02, &h42, &hAC, &h12, &h00, &h02})
DIM SHARED AFX_IID_ICoreWebView2FrameInfoCollection AS GUID = (&h8F834154, &hD38E, &h4D90, {&hAF, &hFB, &h68, &h00, &hA7, &h27, &h28, &h39})
DIM SHARED AFX_IID_ICoreWebView2FrameInfoCollectionIterator AS GUID = (&h1BF89E2D, &h1B2B, &h4629, {&hB2, &h8F, &h05, &h09, &h9B, &h41, &hBB, &h03})
DIM SHARED AFX_IID_ICoreWebView2FrameNameChangedEventHandler AS GUID = (&h435C7DC8, &h9BAA, &h11EB, {&hA8, &hB3, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2FrameNavigationCompletedEventHandler AS GUID = (&h609302AD, &h0E36, &h4F9A, {&hA2, &h10, &h6A, &h45, &h27, &h28, &h42, &hA9})
DIM SHARED AFX_IID_ICoreWebView2FrameNavigationStartingEventHandler AS GUID = (&hE79908BF, &h2D5D, &h4968, {&h83, &hDB, &h26, &h3F, &hEA, &h2C, &h1D, &hA3})
DIM SHARED AFX_IID_ICoreWebView2FramePermissionRequestedEventHandler AS GUID = (&h845D0EDD, &h8BD8, &h429B, {&h99, &h15, &h48, &h21, &h78, &h9F, &h23, &hE9})
DIM SHARED AFX_IID_ICoreWebView2FrameScreenCaptureStartingEventHandler AS GUID = (&hA6C1D8AD, &hBB80, &h59C5, {&h89, &h5B, &hFB, &hA1, &h69, &h8B, &h93, &h09})
DIM SHARED AFX_IID_ICoreWebView2FrameWebMessageReceivedEventHandler AS GUID = (&hE371E005, &h6D1D, &h4517, {&h93, &h4B, &hA8, &hF1, &h62, &h9C, &h62, &hA5})
DIM SHARED AFX_IID_ICoreWebView2GetCookiesCompletedHandler AS GUID = (&h5A4F5069, &h5C15, &h47C3, {&h86, &h46, &hF4, &hDE, &h1C, &h11, &h66, &h70})
DIM SHARED AFX_IID_ICoreWebView2GetFaviconCompletedHandler AS GUID = (&hA2508329, &h7DA8, &h49D7, {&h8C, &h05, &hFA, &h12, &h5E, &h4A, &hEE, &h8D})
DIM SHARED AFX_IID_ICoreWebView2GetNonDefaultPermissionSettingsCompletedHandler AS GUID = (&h38274481, &hA15C, &h4563, {&h94, &hCF, &h99, &h0E, &hDC, &h9A, &hEB, &h95})
DIM SHARED AFX_IID_ICoreWebView2GetProcessExtendedInfosCompletedHandler AS GUID = (&hF45E55AA, &h3BC2, &h11EE, {&hBE, &h56, &h02, &h42, &hAC, &h12, &h00, &h02})
DIM SHARED AFX_IID_ICoreWebView2HistoryChangedEventHandler AS GUID = (&hC79A420C, &hEFD9, &h4058, {&h92, &h95, &h3E, &h8B, &h4B, &hCA, &hB6, &h45})
DIM SHARED AFX_IID_ICoreWebView2HttpHeadersCollectionIterator AS GUID = (&h0702FC30, &hF43B, &h47BB, {&hAB, &h52, &hA4, &h2C, &hB5, &h52, &hAD, &h9F})
DIM SHARED AFX_IID_ICoreWebView2HttpRequestHeaders AS GUID = (&hE86CAC0E, &h5523, &h465C, {&hB5, &h36, &h8F, &hB9, &hFC, &h8C, &h8C, &h60})
DIM SHARED AFX_IID_ICoreWebView2HttpResponseHeaders AS GUID = (&h03C5FF5A, &h9B45, &h4A88, {&h88, &h1C, &h89, &hA9, &hF3, &h28, &h61, &h9C})
DIM SHARED AFX_IID_ICoreWebView2IsDefaultDownloadDialogOpenChangedEventHandler AS GUID = (&h3117DA26, &hAE13, &h438D, {&hBD, &h46, &hED, &hBE, &hB2, &hC4, &hCE, &h81})
DIM SHARED AFX_IID_ICoreWebView2IsDocumentPlayingAudioChangedEventHandler AS GUID = (&h5DEF109A, &h2F4B, &h49FA, {&hB7, &hF6, &h11, &hC3, &h9E, &h51, &h33, &h28})
DIM SHARED AFX_IID_ICoreWebView2IsMutedChangedEventHandler AS GUID = (&h57D90347, &hCD0E, &h4952, {&hA4, &hA2, &h74, &h83, &hA2, &h75, &h6F, &h08})
DIM SHARED AFX_IID_ICoreWebView2LaunchingExternalUriSchemeEventArgs AS GUID = (&h07D1A6C3, &h7175, &h4BA1, {&h93, &h06, &hE5, &h93, &hCA, &h07, &hE4, &h6C})
DIM SHARED AFX_IID_ICoreWebView2LaunchingExternalUriSchemeEventHandler AS GUID = (&h74F712E0, &h8165, &h43A9, {&hA1, &h3F, &h0C, &hCE, &h59, &h7E, &h75, &hDF})
DIM SHARED AFX_IID_ICoreWebView2MoveFocusRequestedEventArgs AS GUID = (&h2D6AA13B, &h3839, &h4A15, {&h92, &hFC, &hD8, &h8B, &h3C, &h0D, &h9C, &h9D})
DIM SHARED AFX_IID_ICoreWebView2MoveFocusRequestedEventHandler AS GUID = (&h69035451, &h6DC7, &h4CB8, {&h9B, &hCE, &hB2, &hBD, &h70, &hAD, &h28, &h9F})
DIM SHARED AFX_IID_ICoreWebView2NavigationCompletedEventHandler AS GUID = (&hD33A35BF, &h1C49, &h4F98, {&h93, &hAB, &h00, &h6E, &h05, &h33, &hFE, &h1C})
DIM SHARED AFX_IID_ICoreWebView2NavigationStartingEventHandler AS GUID = (&h9ADBE429, &hF36D, &h432B, {&h9D, &hDC, &hF8, &h88, &h1F, &hBD, &h76, &hE3})
DIM SHARED AFX_IID_ICoreWebView2NewBrowserVersionAvailableEventHandler AS GUID = (&hF9A2976E, &hD34E, &h44FC, {&hAD, &hEE, &h81, &hB6, &hB5, &h7C, &hA9, &h14})
DIM SHARED AFX_IID_ICoreWebView2NewWindowRequestedEventHandler AS GUID = (&hD4C185FE, &hC81C, &h4989, {&h97, &hAF, &h2D, &h3F, &hA7, &hAB, &h56, &h51})
DIM SHARED AFX_IID_ICoreWebView2NonClientRegionChangedEventArgs AS GUID = (&hAB71D500, &h0820, &h4A52, {&h80, &h9C, &h48, &hDB, &h04, &hFF, &h93, &hBF})
DIM SHARED AFX_IID_ICoreWebView2NonClientRegionChangedEventHandler AS GUID = (&h4A794E66, &hAA6C, &h46BD, {&h93, &hA3, &h38, &h21, &h96, &h83, &h76, &h80})
DIM SHARED AFX_IID_ICoreWebView2Notification AS GUID = (&hB7434D98, &h6BC8, &h419D, {&h9D, &hA5, &hFB, &h5A, &h96, &hD4, &hDA, &hCD})
DIM SHARED AFX_IID_ICoreWebView2NotificationCloseRequestedEventHandler AS GUID = (&h47C32D23, &h1E94, &h4733, {&h85, &hF1, &hD9, &hBF, &h4A, &hCD, &h09, &h74})
DIM SHARED AFX_IID_ICoreWebView2NotificationReceivedEventArgs AS GUID = (&h1512DD5B, &h5514, &h4F85, {&h88, &h6E, &h21, &hC3, &hA4, &hC9, &hCF, &hE6})
DIM SHARED AFX_IID_ICoreWebView2NotificationReceivedEventHandler AS GUID = (&h89C5D598, &h8788, &h423B, {&hBE, &h97, &hE6, &hE0, &h1C, &h0F, &h9E, &hE3})
DIM SHARED AFX_IID_ICoreWebView2ObjectCollection AS GUID = (&h5CFEC11C, &h25BD, &h4E8D, {&h9E, &h1A, &h7A, &hCD, &hAE, &hEE, &hC0, &h47})
DIM SHARED AFX_IID_ICoreWebView2ObjectCollectionView AS GUID = (&h0F36FD87, &h4F69, &h4415, {&h98, &hDA, &h88, &h8F, &h89, &hFB, &h9A, &h33})
DIM SHARED AFX_IID_ICoreWebView2PermissionSetting AS GUID = (&h792B6ECA, &h5576, &h421C, {&h91, &h19, &h74, &hEB, &hB3, &hA4, &hFF, &hB3})
DIM SHARED AFX_IID_ICoreWebView2PermissionSettingCollectionView AS GUID = (&hF5596F62, &h3DE5, &h47B1, {&h91, &hE8, &hA4, &h10, &h4B, &h59, &h6B, &h96})
DIM SHARED AFX_IID_ICoreWebView2PointerInfo AS GUID = (&hE6995887, &hD10D, &h4F5D, {&h93, &h59, &h4C, &hE4, &h6E, &h4F, &h96, &hB9})
DIM SHARED AFX_IID_ICoreWebView2PrintCompletedHandler AS GUID = (&h8FD80075, &hED08, &h42DB, {&h85, &h70, &hF5, &hD1, &h49, &h77, &h46, &h1E})
DIM SHARED AFX_IID_ICoreWebView2PrintSettings AS GUID = (&h377F3721, &hC74E, &h48CA, {&h8D, &hB1, &hDF, &h68, &hE5, &h1D, &h60, &hE2})
DIM SHARED AFX_IID_ICoreWebView2PrintSettings2 AS GUID = (&hCA7F0E1F, &h3484, &h41D1, {&h8C, &h1A, &h65, &hCD, &h44, &hA6, &h3F, &h8D})
DIM SHARED AFX_IID_ICoreWebView2PrintToPdfCompletedHandler AS GUID = (&hCCF1EF04, &hFD8E, &h4D5F, {&hB2, &hDE, &h09, &h83, &hE4, &h1B, &h8C, &h36})
DIM SHARED AFX_IID_ICoreWebView2PrintToPdfStreamCompletedHandler AS GUID = (&h4C9F8229, &h8F93, &h444F, {&hA7, &h11, &h2C, &h0D, &hFD, &h63, &h59, &hD5})
DIM SHARED AFX_IID_ICoreWebView2ProcessExtendedInfo AS GUID = (&hAF4C4C2E, &h45DB, &h11EE, {&hBE, &h56, &h02, &h42, &hAC, &h12, &h00, &h02})
DIM SHARED AFX_IID_ICoreWebView2ProcessExtendedInfoCollection AS GUID = (&h32EFA696, &h407A, &h11EE, {&hBE, &h56, &h02, &h42, &hAC, &h12, &h00, &h02})
DIM SHARED AFX_IID_ICoreWebView2ProcessInfo AS GUID = (&h84FA7612, &h3F3D, &h4FBF, {&h88, &h9D, &hFA, &hD0, &h00, &h49, &h2D, &h72})
DIM SHARED AFX_IID_ICoreWebView2ProcessInfoCollection AS GUID = (&h402B99CD, &hA0CC, &h4FA5, {&hB7, &hA5, &h51, &hD8, &h6A, &h1D, &h23, &h39})
DIM SHARED AFX_IID_ICoreWebView2ProcessInfosChangedEventHandler AS GUID = (&hF4AF0C39, &h44B9, &h40E9, {&h8B, &h11, &h04, &h84, &hCF, &hB9, &hE0, &hA1})
DIM SHARED AFX_IID_ICoreWebView2ProfileAddBrowserExtensionCompletedHandler AS GUID = (&hDF1AAB27, &h82B9, &h4AB6, {&hAA, &hE8, &h01, &h7A, &h49, &h39, &h8C, &h14})
DIM SHARED AFX_IID_ICoreWebView2ProfileDeletedEventHandler AS GUID = (&hDF35055D, &h772E, &h4DBE, {&hB7, &h43, &h5F, &hBF, &h74, &hA2, &hB2, &h58})
DIM SHARED AFX_IID_ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler AS GUID = (&hFCE16A1C, &hF107, &h4601, {&h8B, &h75, &hFC, &h49, &h40, &hAE, &h25, &hD0})
DIM SHARED AFX_IID_ICoreWebView2RasterizationScaleChangedEventHandler AS GUID = (&h9C98C8B1, &hAC53, &h427E, {&hA3, &h45, &h30, &h49, &hB5, &h52, &h4B, &hBE})
DIM SHARED AFX_IID_ICoreWebView2RegionRectCollectionView AS GUID = (&h333353B8, &h48BF, &h4449, {&h8F, &hCC, &h22, &h69, &h7F, &hAF, &h57, &h53})
DIM SHARED AFX_IID_ICoreWebView2SaveAsUIShowingEventArgs AS GUID = (&h55902952, &h0E0D, &h5AAA, {&hA7, &hD0, &hE8, &h33, &hCD, &hB3, &h4F, &h62})
DIM SHARED AFX_IID_ICoreWebView2SaveAsUIShowingEventHandler AS GUID = (&h6BAA177E, &h3A2E, &h5CCF, {&h9A, &h13, &hFA, &hD6, &h76, &hCD, &h05, &h22})
DIM SHARED AFX_IID_ICoreWebView2SaveFileSecurityCheckStartingEventArgs AS GUID = (&hCF4FF1D1, &h5A67, &h5660, {&h8D, &h63, &hEF, &h69, &h98, &h81, &hEA, &h65})
DIM SHARED AFX_IID_ICoreWebView2SaveFileSecurityCheckStartingEventHandler AS GUID = (&h7899576C, &h19E3, &h57C8, {&hB7, &hD1, &h55, &h80, &h82, &h92, &hDE, &h57})
DIM SHARED AFX_IID_ICoreWebView2ScreenCaptureStartingEventArgs AS GUID = (&h892C03FD, &hAEE3, &h5EBA, {&hA1, &hFA, &h6F, &hD2, &hF6, &h48, &h4B, &h2B})
DIM SHARED AFX_IID_ICoreWebView2ScreenCaptureStartingEventHandler AS GUID = (&hE24FF05A, &h1DB5, &h59D9, {&h89, &hF3, &h3C, &h86, &h42, &h68, &hDB, &h4A})
DIM SHARED AFX_IID_ICoreWebView2ScriptDialogOpeningEventArgs AS GUID = (&h7390BB70, &hABE0, &h4843, {&h95, &h29, &hF1, &h43, &hB3, &h1B, &h03, &hD6})
DIM SHARED AFX_IID_ICoreWebView2ScriptDialogOpeningEventHandler AS GUID = (&hEF381BF9, &hAFA8, &h4E37, {&h91, &hC4, &h8A, &hC4, &h85, &h24, &hBD, &hFB})
DIM SHARED AFX_IID_ICoreWebView2ScriptException AS GUID = (&h054DAE00, &h84A3, &h49FF, {&hBC, &h17, &h40, &h12, &hA9, &h0B, &hC9, &hFD})
DIM SHARED AFX_IID_ICoreWebView2ServerCertificateErrorDetectedEventArgs AS GUID = (&h012193ED, &h7C13, &h48FF, {&h96, &h9D, &hA8, &h4C, &h1F, &h43, &h2A, &h14})
DIM SHARED AFX_IID_ICoreWebView2ServerCertificateErrorDetectedEventHandler AS GUID = (&h969B3A26, &hD85E, &h4795, {&h81, &h99, &hFE, &hF5, &h73, &h44, &hDA, &h22})
DIM SHARED AFX_IID_ICoreWebView2SetPermissionStateCompletedHandler AS GUID = (&hFC77FB30, &h9C9E, &h4076, {&hB8, &hC7, &h76, &h44, &hA7, &h03, &hCA, &h1B})
DIM SHARED AFX_IID_ICoreWebView2SharedBuffer AS GUID = (&hB747A495, &h0C6F, &h449E, {&h97, &hB8, &h2F, &h81, &hE9, &hD6, &hAB, &h43})
DIM SHARED AFX_IID_ICoreWebView2ShowSaveAsUICompletedHandler AS GUID = (&hE24B07E3, &h8169, &h5C34, {&h99, &h4A, &h7F, &h64, &h78, &h94, &h6A, &h3C})
DIM SHARED AFX_IID_ICoreWebView2SourceChangedEventArgs AS GUID = (&h31E0E545, &h1DBA, &h4266, {&h89, &h14, &hF6, &h38, &h48, &hA1, &hF7, &hD7})
DIM SHARED AFX_IID_ICoreWebView2SourceChangedEventHandler AS GUID = (&h3C067F9F, &h5388, &h4772, {&h8B, &h48, &h79, &hF7, &hEF, &h1A, &hB3, &h7C})
DIM SHARED AFX_IID_ICoreWebView2StateChangedEventHandler AS GUID = (&h81336594, &h7EDE, &h4BA9, {&hBF, &h71, &hAC, &hF0, &hA9, &h5B, &h58, &hDD})
DIM SHARED AFX_IID_ICoreWebView2StatusBarTextChangedEventHandler AS GUID = (&hA5E3B0D0, &h10DF, &h4156, {&hBF, &hAD, &h3B, &h43, &h86, &h7A, &hCA, &hC6})
DIM SHARED AFX_IID_ICoreWebView2StringCollection AS GUID = (&hF41F3F8A, &hBCC3, &h11EB, {&h85, &h29, &h02, &h42, &hAC, &h13, &h00, &h03})
DIM SHARED AFX_IID_ICoreWebView2TrySuspendCompletedHandler AS GUID = (&h00F206A7, &h9D17, &h4605, {&h91, &hF6, &h4E, &h8E, &h4D, &hE1, &h92, &hE3})
DIM SHARED AFX_IID_ICoreWebView2WebMessageReceivedEventHandler AS GUID = (&h57213F19, &h00E6, &h49FA, {&h8E, &h07, &h89, &h8E, &hA0, &h1E, &hCB, &hD2})
DIM SHARED AFX_IID_ICoreWebView2WebResourceRequest AS GUID = (&h97055CD4, &h512C, &h4264, {&h8B, &h5F, &hE3, &hF4, &h46, &hCE, &hA6, &hA5})
DIM SHARED AFX_IID_ICoreWebView2WebResourceRequestedEventHandler AS GUID = (&hAB00B74C, &h15F1, &h4646, {&h80, &hE8, &hE7, &h63, &h41, &hD2, &h5D, &h71})
DIM SHARED AFX_IID_ICoreWebView2WebResourceResponse AS GUID = (&hAAFCC94F, &hFA27, &h48FD, {&h97, &hDF, &h83, &h0E, &hF7, &h5A, &hAE, &hC9})
DIM SHARED AFX_IID_ICoreWebView2WebResourceResponseReceivedEventArgs AS GUID = (&hD1DB483D, &h6796, &h4B8B, {&h80, &hFC, &h13, &h71, &h2B, &hB7, &h16, &hF4})
DIM SHARED AFX_IID_ICoreWebView2WebResourceResponseReceivedEventHandler AS GUID = (&h7DE9898A, &h24F5, &h40C3, {&hA2, &hDE, &hD4, &hF4, &h58, &hE6, &h98, &h28})
DIM SHARED AFX_IID_ICoreWebView2WebResourceResponseView AS GUID = (&h79701053, &h7759, &h4162, {&h8F, &h7D, &hF1, &hB3, &hF0, &h84, &h92, &h8D})
DIM SHARED AFX_IID_ICoreWebView2WebResourceResponseViewGetContentCompletedHandler AS GUID = (&h875738E1, &h9FA2, &h40E3, {&h8B, &h74, &h2E, &h89, &h72, &hDD, &h6F, &hE7})
DIM SHARED AFX_IID_ICoreWebView2WindowCloseRequestedEventHandler AS GUID = (&h5C19E9E0, &h092F, &h486B, {&hAF, &hFA, &hCA, &h82, &h31, &h91, &h30, &h39})
DIM SHARED AFX_IID_ICoreWebView2WindowFeatures AS GUID = (&h5EAF559F, &hB46E, &h4397, {&h88, &h60, &hE4, &h22, &hF2, &h87, &hFF, &h1E})
DIM SHARED AFX_IID_ICoreWebView2ZoomFactorChangedEventHandler AS GUID = (&hB52D71D6, &hC4DF, &h4543, {&hA9, &h0C, &h64, &hA3, &hE6, &h0F, &h38, &hCB})

' // Interfaces - Forward references

TYPE Afx_ICoreWebView2 AS Afx_ICoreWebView2_
TYPE Afx_ICoreWebView2_2 AS Afx_ICoreWebView2_2_
TYPE Afx_ICoreWebView2_3 AS Afx_ICoreWebView2_3_
TYPE Afx_ICoreWebView2_4 AS Afx_ICoreWebView2_4_
TYPE Afx_ICoreWebView2_5 AS Afx_ICoreWebView2_5_
TYPE Afx_ICoreWebView2_6 AS Afx_ICoreWebView2_6_
TYPE Afx_ICoreWebView2_7 AS Afx_ICoreWebView2_7_
TYPE Afx_ICoreWebView2_8 AS Afx_ICoreWebView2_8_
TYPE Afx_ICoreWebView2_9 AS Afx_ICoreWebView2_9_
TYPE Afx_ICoreWebView2_10 AS Afx_ICoreWebView2_10_
TYPE Afx_ICoreWebView2_11 AS Afx_ICoreWebView2_11_
TYPE Afx_ICoreWebView2_12 AS Afx_ICoreWebView2_12_
TYPE Afx_ICoreWebView2_13 AS Afx_ICoreWebView2_13_
TYPE Afx_ICoreWebView2_14 AS Afx_ICoreWebView2_14_
TYPE Afx_ICoreWebView2_15 AS Afx_ICoreWebView2_15_
TYPE Afx_ICoreWebView2_16 AS Afx_ICoreWebView2_16_
TYPE Afx_ICoreWebView2_17 AS Afx_ICoreWebView2_17_
TYPE Afx_ICoreWebView2_18 AS Afx_ICoreWebView2_18_
TYPE Afx_ICoreWebView2_19 AS Afx_ICoreWebView2_19_
TYPE Afx_ICoreWebView2_20 AS Afx_ICoreWebView2_20_
TYPE Afx_ICoreWebView2_21 AS Afx_ICoreWebView2_21_
TYPE Afx_ICoreWebView2_22 AS Afx_ICoreWebView2_22_
TYPE Afx_ICoreWebView2_23 AS Afx_ICoreWebView2_23_
TYPE Afx_ICoreWebView2_24 AS Afx_ICoreWebView2_24_
TYPE Afx_ICoreWebView2_25 AS Afx_ICoreWebView2_25_
TYPE Afx_ICoreWebView2_26 AS Afx_ICoreWebView2_26_
TYPE Afx_ICoreWebView2_27 AS Afx_ICoreWebView2_27_
TYPE Afx_ICoreWebView2_28 AS Afx_ICoreWebView2_28_
TYPE Afx_ICoreWebView2Environment AS Afx_ICoreWebView2Environment_
TYPE Afx_ICoreWebView2Environment2 AS Afx_ICoreWebView2Environment2_
TYPE Afx_ICoreWebView2Environment3 AS Afx_ICoreWebView2Environment3_
TYPE Afx_ICoreWebView2Environment4 AS Afx_ICoreWebView2Environment4_
TYPE Afx_ICoreWebView2Environment5 AS Afx_ICoreWebView2Environment5_
TYPE Afx_ICoreWebView2Environment6 AS Afx_ICoreWebView2Environment6_
TYPE Afx_ICoreWebView2Environment7 AS Afx_ICoreWebView2Environment7_
TYPE Afx_ICoreWebView2Environment8 AS Afx_ICoreWebView2Environment8_
TYPE Afx_ICoreWebView2Environment9 AS Afx_ICoreWebView2Environment9_
TYPE Afx_ICoreWebView2Environment10 AS Afx_ICoreWebView2Environment10_
TYPE Afx_ICoreWebView2Environment11 AS Afx_ICoreWebView2Environment11_
TYPE Afx_ICoreWebView2Environment12 AS Afx_ICoreWebView2Environment12_
TYPE Afx_ICoreWebView2Environment13 AS Afx_ICoreWebView2Environment13_
TYPE Afx_ICoreWebView2Environment14 AS Afx_ICoreWebView2Environment14_
TYPE Afx_ICoreWebView2Environment15 AS Afx_ICoreWebView2Environment15_
TYPE Afx_ICoreWebView2EnvironmentOptions AS Afx_ICoreWebView2EnvironmentOptions_
TYPE Afx_ICoreWebView2EnvironmentOptions2 AS Afx_ICoreWebView2EnvironmentOptions2_
TYPE Afx_ICoreWebView2EnvironmentOptions3 AS Afx_ICoreWebView2EnvironmentOptions3_
TYPE Afx_ICoreWebView2EnvironmentOptions4 AS Afx_ICoreWebView2EnvironmentOptions4_
TYPE Afx_ICoreWebView2EnvironmentOptions5 AS Afx_ICoreWebView2EnvironmentOptions5_
TYPE Afx_ICoreWebView2EnvironmentOptions6 AS Afx_ICoreWebView2EnvironmentOptions6_
TYPE Afx_ICoreWebView2EnvironmentOptions7 AS Afx_ICoreWebView2EnvironmentOptions7_
TYPE Afx_ICoreWebView2EnvironmentOptions8 AS Afx_ICoreWebView2EnvironmentOptions8_
TYPE Afx_ICoreWebView2Frame AS Afx_ICoreWebView2Frame_
TYPE Afx_ICoreWebView2Frame2 AS Afx_ICoreWebView2Frame2_
TYPE Afx_ICoreWebView2Frame3 AS Afx_ICoreWebView2Frame3_
TYPE Afx_ICoreWebView2Frame4 AS Afx_ICoreWebView2Frame4_
TYPE Afx_ICoreWebView2Frame5 AS Afx_ICoreWebView2Frame5_
TYPE Afx_ICoreWebView2Frame6 AS Afx_ICoreWebView2Frame6_
TYPE Afx_ICoreWebView2Frame7 AS Afx_ICoreWebView2Frame7_
TYPE Afx_ICoreWebView2Profile AS Afx_ICoreWebView2Profile_
TYPE Afx_ICoreWebView2Profile2 AS Afx_ICoreWebView2Profile2_
TYPE Afx_ICoreWebView2Profile3 AS Afx_ICoreWebView2Profile3_
TYPE Afx_ICoreWebView2Profile4 AS Afx_ICoreWebView2Profile4_
TYPE Afx_ICoreWebView2Profile5 AS Afx_ICoreWebView2Profile5_
TYPE Afx_ICoreWebView2Profile6 AS Afx_ICoreWebView2Profile6_
TYPE Afx_ICoreWebView2Profile7 AS Afx_ICoreWebView2Profile7_
TYPE Afx_ICoreWebView2Profile8 AS Afx_ICoreWebView2Profile8_
TYPE Afx_ICoreWebView2Settings AS Afx_ICoreWebView2Settings_
TYPE Afx_ICoreWebView2Settings2 AS Afx_ICoreWebView2Settings2_
TYPE Afx_ICoreWebView2Settings3 AS Afx_ICoreWebView2Settings3_
TYPE Afx_ICoreWebView2Settings4 AS Afx_ICoreWebView2Settings4_
TYPE Afx_ICoreWebView2Settings5 AS Afx_ICoreWebView2Settings5_
TYPE Afx_ICoreWebView2Settings6 AS Afx_ICoreWebView2Settings6_
TYPE Afx_ICoreWebView2Settings7 AS Afx_ICoreWebView2Settings7_
TYPE Afx_ICoreWebView2Settings8 AS Afx_ICoreWebView2Settings8_
TYPE Afx_ICoreWebView2Settings9 AS Afx_ICoreWebView2Settings9_
TYPE Afx_ICoreWebView2AcceleratorKeyPressedEventArgs AS Afx_ICoreWebView2AcceleratorKeyPressedEventArgs_
TYPE Afx_ICoreWebView2AcceleratorKeyPressedEventArgs2 AS Afx_ICoreWebView2AcceleratorKeyPressedEventArgs2_
TYPE Afx_ICoreWebView2AcceleratorKeyPressedEventHandler AS Afx_ICoreWebView2AcceleratorKeyPressedEventHandler_
TYPE Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler AS Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_
TYPE Afx_ICoreWebView2BasicAuthenticationRequestedEventArgs AS Afx_ICoreWebView2BasicAuthenticationRequestedEventArgs_
TYPE Afx_ICoreWebView2BasicAuthenticationRequestedEventHandler AS Afx_ICoreWebView2BasicAuthenticationRequestedEventHandler_
TYPE Afx_ICoreWebView2BasicAuthenticationResponse AS Afx_ICoreWebView2BasicAuthenticationResponse_
TYPE Afx_ICoreWebView2BrowserExtension AS Afx_ICoreWebView2BrowserExtension_
TYPE Afx_ICoreWebView2BrowserExtensionEnableCompletedHandler AS Afx_ICoreWebView2BrowserExtensionEnableCompletedHandler_
TYPE Afx_ICoreWebView2BrowserExtensionList AS Afx_ICoreWebView2BrowserExtensionList_
TYPE Afx_ICoreWebView2BrowserExtensionRemoveCompletedHandler AS Afx_ICoreWebView2BrowserExtensionRemoveCompletedHandler_
TYPE Afx_ICoreWebView2BrowserProcessExitedEventArgs AS Afx_ICoreWebView2BrowserProcessExitedEventArgs_
TYPE Afx_ICoreWebView2BrowserProcessExitedEventHandler AS Afx_ICoreWebView2BrowserProcessExitedEventHandler_
TYPE Afx_ICoreWebView2BytesReceivedChangedEventHandler AS Afx_ICoreWebView2BytesReceivedChangedEventHandler_
TYPE Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler AS Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_
TYPE Afx_ICoreWebView2CapturePreviewCompletedHandler AS Afx_ICoreWebView2CapturePreviewCompletedHandler_
TYPE Afx_ICoreWebView2Certificate AS Afx_ICoreWebView2Certificate_
TYPE Afx_ICoreWebView2ClearBrowsingDataCompletedHandler AS Afx_ICoreWebView2ClearBrowsingDataCompletedHandler_
TYPE Afx_ICoreWebView2ClearServerCertificateErrorActionsCompletedHandler AS Afx_ICoreWebView2ClearServerCertificateErrorActionsCompletedHandler_
TYPE Afx_ICoreWebView2ClientCertificate AS Afx_ICoreWebView2ClientCertificate_
TYPE Afx_ICoreWebView2ClientCertificateCollection AS Afx_ICoreWebView2ClientCertificateCollection_
TYPE Afx_ICoreWebView2ClientCertificateRequestedEventArgs AS Afx_ICoreWebView2ClientCertificateRequestedEventArgs_
TYPE Afx_ICoreWebView2ClientCertificateRequestedEventHandler AS Afx_ICoreWebView2ClientCertificateRequestedEventHandler_
TYPE Afx_ICoreWebView2CompositionController AS Afx_ICoreWebView2CompositionController_
TYPE Afx_ICoreWebView2CompositionController2 AS Afx_ICoreWebView2CompositionController2_
TYPE Afx_ICoreWebView2CompositionController3 AS Afx_ICoreWebView2CompositionController3_
TYPE Afx_ICoreWebView2CompositionController4 AS Afx_ICoreWebView2CompositionController4_
TYPE Afx_ICoreWebView2ContainsFullScreenElementChangedEventHandler AS Afx_ICoreWebView2ContainsFullScreenElementChangedEventHandler_
TYPE Afx_ICoreWebView2ContentLoadingEventArgs AS Afx_ICoreWebView2ContentLoadingEventArgs_
TYPE Afx_ICoreWebView2ContentLoadingEventHandler AS Afx_ICoreWebView2ContentLoadingEventHandler_
TYPE Afx_ICoreWebView2ContextMenuItem AS Afx_ICoreWebView2ContextMenuItem_
TYPE Afx_ICoreWebView2ContextMenuItemCollection AS Afx_ICoreWebView2ContextMenuItemCollection_
TYPE Afx_ICoreWebView2ContextMenuRequestedEventArgs AS Afx_ICoreWebView2ContextMenuRequestedEventArgs_
TYPE Afx_ICoreWebView2ContextMenuRequestedEventHandler AS Afx_ICoreWebView2ContextMenuRequestedEventHandler_
TYPE Afx_ICoreWebView2ContextMenuTarget AS Afx_ICoreWebView2ContextMenuTarget_
TYPE Afx_ICoreWebView2Controller AS Afx_ICoreWebView2Controller_
TYPE Afx_ICoreWebView2Controller2 AS Afx_ICoreWebView2Controller2_
TYPE Afx_ICoreWebView2Controller3 AS Afx_ICoreWebView2Controller3_
TYPE Afx_ICoreWebView2Controller4 AS Afx_ICoreWebView2Controller4_
TYPE Afx_ICoreWebView2ControllerOptions AS Afx_ICoreWebView2ControllerOptions_
TYPE Afx_ICoreWebView2ControllerOptions2 AS Afx_ICoreWebView2ControllerOptions2_
TYPE Afx_ICoreWebView2ControllerOptions3 AS Afx_ICoreWebView2ControllerOptions3_
TYPE Afx_ICoreWebView2ControllerOptions4 AS Afx_ICoreWebView2ControllerOptions4_
TYPE Afx_ICoreWebView2Cookie AS Afx_ICoreWebView2Cookie_
TYPE Afx_ICoreWebView2CookieList AS Afx_ICoreWebView2CookieList_
TYPE Afx_ICoreWebView2CookieManager AS Afx_ICoreWebView2CookieManager_
TYPE Afx_ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler AS Afx_ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler_
TYPE Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler AS Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_
TYPE Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler AS Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_
TYPE Afx_ICoreWebView2CursorChangedEventHandler AS Afx_ICoreWebView2CursorChangedEventHandler_
TYPE Afx_ICoreWebView2CustomItemSelectedEventHandler AS Afx_ICoreWebView2CustomItemSelectedEventHandler_
TYPE Afx_ICoreWebView2CustomSchemeRegistration AS Afx_ICoreWebView2CustomSchemeRegistration_
TYPE Afx_ICoreWebView2Deferral AS Afx_ICoreWebView2Deferral_
TYPE Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs_
TYPE Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs2 AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs2_
TYPE Afx_ICoreWebView2DevToolsProtocolEventReceivedEventHandler AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventHandler_
TYPE Afx_ICoreWebView2DevToolsProtocolEventReceiver AS Afx_ICoreWebView2DevToolsProtocolEventReceiver_
TYPE Afx_ICoreWebView2DocumentTitleChangedEventHandler AS Afx_ICoreWebView2DocumentTitleChangedEventHandler_
TYPE Afx_ICoreWebView2DOMContentLoadedEventArgs AS Afx_ICoreWebView2DOMContentLoadedEventArgs_
TYPE Afx_ICoreWebView2DOMContentLoadedEventHandler AS Afx_ICoreWebView2DOMContentLoadedEventHandler_
TYPE Afx_ICoreWebView2DownloadOperation AS Afx_ICoreWebView2DownloadOperation_
TYPE Afx_ICoreWebView2DownloadStartingEventArgs AS Afx_ICoreWebView2DownloadStartingEventArgs_
TYPE Afx_ICoreWebView2DownloadStartingEventHandler AS Afx_ICoreWebView2DownloadStartingEventHandler_
TYPE Afx_ICoreWebView2EstimatedEndTimeChangedEventHandler AS Afx_ICoreWebView2EstimatedEndTimeChangedEventHandler_
TYPE Afx_ICoreWebView2ExecuteScriptCompletedHandler AS Afx_ICoreWebView2ExecuteScriptCompletedHandler_
TYPE Afx_ICoreWebView2ExecuteScriptResult AS Afx_ICoreWebView2ExecuteScriptResult_
TYPE Afx_ICoreWebView2ExecuteScriptWithResultCompletedHandler AS Afx_ICoreWebView2ExecuteScriptWithResultCompletedHandler_
TYPE Afx_ICoreWebView2FaviconChangedEventHandler AS Afx_ICoreWebView2FaviconChangedEventHandler_
TYPE Afx_ICoreWebView2File AS Afx_ICoreWebView2File_
TYPE Afx_ICoreWebView2FileSystemHandle AS Afx_ICoreWebView2FileSystemHandle_
TYPE Afx_ICoreWebView2Find AS Afx_ICoreWebView2Find_
TYPE Afx_ICoreWebView2FindActiveMatchIndexChangedEventHandler AS Afx_ICoreWebView2FindActiveMatchIndexChangedEventHandler_
TYPE Afx_ICoreWebView2FindMatchCountChangedEventHandler AS Afx_ICoreWebView2FindMatchCountChangedEventHandler_
TYPE Afx_ICoreWebView2FindOptions AS Afx_ICoreWebView2FindOptions_
TYPE Afx_ICoreWebView2FindStartCompletedHandler AS Afx_ICoreWebView2FindStartCompletedHandler_
TYPE Afx_ICoreWebView2FocusChangedEventHandler AS Afx_ICoreWebView2FocusChangedEventHandler_
TYPE Afx_ICoreWebView2FrameChildFrameCreatedEventHandler AS Afx_ICoreWebView2FrameChildFrameCreatedEventHandler_
TYPE Afx_ICoreWebView2FrameContentLoadingEventHandler AS Afx_ICoreWebView2FrameContentLoadingEventHandler_
TYPE Afx_ICoreWebView2FrameCreatedEventArgs AS Afx_ICoreWebView2FrameCreatedEventArgs_
TYPE Afx_ICoreWebView2FrameCreatedEventHandler AS Afx_ICoreWebView2FrameCreatedEventHandler_
TYPE Afx_ICoreWebView2FrameDestroyedEventHandler AS Afx_ICoreWebView2FrameDestroyedEventHandler_
TYPE Afx_ICoreWebView2FrameDOMContentLoadedEventHandler AS Afx_ICoreWebView2FrameDOMContentLoadedEventHandler_
TYPE Afx_ICoreWebView2FrameInfo AS Afx_ICoreWebView2FrameInfo_
TYPE Afx_ICoreWebView2FrameInfo2 AS Afx_ICoreWebView2FrameInfo2_
TYPE Afx_ICoreWebView2FrameInfoCollection AS Afx_ICoreWebView2FrameInfoCollection_
TYPE Afx_ICoreWebView2FrameInfoCollectionIterator AS Afx_ICoreWebView2FrameInfoCollectionIterator_
TYPE Afx_ICoreWebView2FrameNameChangedEventHandler AS Afx_ICoreWebView2FrameNameChangedEventHandler_
TYPE Afx_ICoreWebView2FrameNavigationCompletedEventHandler AS Afx_ICoreWebView2FrameNavigationCompletedEventHandler_
TYPE Afx_ICoreWebView2FrameNavigationStartingEventHandler AS Afx_ICoreWebView2FrameNavigationStartingEventHandler_
TYPE Afx_ICoreWebView2FramePermissionRequestedEventHandler AS Afx_ICoreWebView2FramePermissionRequestedEventHandler_
TYPE Afx_ICoreWebView2FrameScreenCaptureStartingEventHandler AS Afx_ICoreWebView2FrameScreenCaptureStartingEventHandler_
TYPE Afx_ICoreWebView2FrameWebMessageReceivedEventHandler AS Afx_ICoreWebView2FrameWebMessageReceivedEventHandler_
TYPE Afx_ICoreWebView2GetCookiesCompletedHandler AS Afx_ICoreWebView2GetCookiesCompletedHandler_
TYPE Afx_ICoreWebView2GetFaviconCompletedHandler AS Afx_ICoreWebView2GetFaviconCompletedHandler_
TYPE Afx_ICoreWebView2GetNonDefaultPermissionSettingsCompletedHandler AS Afx_ICoreWebView2GetNonDefaultPermissionSettingsCompletedHandler_
TYPE Afx_ICoreWebView2GetProcessExtendedInfosCompletedHandler AS Afx_ICoreWebView2GetProcessExtendedInfosCompletedHandler_
TYPE Afx_ICoreWebView2HistoryChangedEventHandler AS Afx_ICoreWebView2HistoryChangedEventHandler_
TYPE Afx_ICoreWebView2HttpHeadersCollectionIterator AS Afx_ICoreWebView2HttpHeadersCollectionIterator_
TYPE Afx_ICoreWebView2HttpRequestHeaders AS Afx_ICoreWebView2HttpRequestHeaders_
TYPE Afx_ICoreWebView2HttpResponseHeaders AS Afx_ICoreWebView2HttpResponseHeaders_
TYPE Afx_ICoreWebView2IsDefaultDownloadDialogOpenChangedEventHandler AS Afx_ICoreWebView2IsDefaultDownloadDialogOpenChangedEventHandler_
TYPE Afx_ICoreWebView2IsDocumentPlayingAudioChangedEventHandler AS Afx_ICoreWebView2IsDocumentPlayingAudioChangedEventHandler_
TYPE Afx_ICoreWebView2IsMutedChangedEventHandler AS Afx_ICoreWebView2IsMutedChangedEventHandler_
TYPE Afx_ICoreWebView2LaunchingExternalUriSchemeEventArgs AS Afx_ICoreWebView2LaunchingExternalUriSchemeEventArgs_
TYPE Afx_ICoreWebView2LaunchingExternalUriSchemeEventHandler AS Afx_ICoreWebView2LaunchingExternalUriSchemeEventHandler_
TYPE Afx_ICoreWebView2MoveFocusRequestedEventArgs AS Afx_ICoreWebView2MoveFocusRequestedEventArgs_
TYPE Afx_ICoreWebView2MoveFocusRequestedEventHandler AS Afx_ICoreWebView2MoveFocusRequestedEventHandler_
TYPE Afx_ICoreWebView2NavigationCompletedEventArgs AS Afx_ICoreWebView2NavigationCompletedEventArgs_
TYPE Afx_ICoreWebView2NavigationCompletedEventArgs2 AS Afx_ICoreWebView2NavigationCompletedEventArgs2_
TYPE Afx_ICoreWebView2NavigationCompletedEventHandler AS Afx_ICoreWebView2NavigationCompletedEventHandler_
TYPE Afx_ICoreWebView2NavigationStartingEventArgs AS Afx_ICoreWebView2NavigationStartingEventArgs_
TYPE Afx_ICoreWebView2NavigationStartingEventArgs2 AS Afx_ICoreWebView2NavigationStartingEventArgs2_
TYPE Afx_ICoreWebView2NavigationStartingEventArgs3 AS Afx_ICoreWebView2NavigationStartingEventArgs3_
TYPE Afx_ICoreWebView2NavigationStartingEventHandler AS Afx_ICoreWebView2NavigationStartingEventHandler_
TYPE Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler AS Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler_
TYPE Afx_ICoreWebView2NewWindowRequestedEventArgs AS Afx_ICoreWebView2NewWindowRequestedEventArgs_
TYPE Afx_ICoreWebView2NewWindowRequestedEventArgs2 AS Afx_ICoreWebView2NewWindowRequestedEventArgs2_
TYPE Afx_ICoreWebView2NewWindowRequestedEventArgs3 AS Afx_ICoreWebView2NewWindowRequestedEventArgs3_
TYPE Afx_ICoreWebView2NewWindowRequestedEventHandler AS Afx_ICoreWebView2NewWindowRequestedEventHandler_
TYPE Afx_ICoreWebView2NonClientRegionChangedEventArgs AS Afx_ICoreWebView2NonClientRegionChangedEventArgs_
TYPE Afx_ICoreWebView2NonClientRegionChangedEventHandler AS Afx_ICoreWebView2NonClientRegionChangedEventHandler_
TYPE Afx_ICoreWebView2Notification AS Afx_ICoreWebView2Notification_
TYPE Afx_ICoreWebView2NotificationCloseRequestedEventHandler AS Afx_ICoreWebView2NotificationCloseRequestedEventHandler_
TYPE Afx_ICoreWebView2NotificationReceivedEventArgs AS Afx_ICoreWebView2NotificationReceivedEventArgs_
TYPE Afx_ICoreWebView2NotificationReceivedEventHandler AS Afx_ICoreWebView2NotificationReceivedEventHandler_
TYPE Afx_ICoreWebView2ObjectCollection AS Afx_ICoreWebView2ObjectCollection_
TYPE Afx_ICoreWebView2ObjectCollectionView AS Afx_ICoreWebView2ObjectCollectionView_
TYPE Afx_ICoreWebView2PermissionRequestedEventArgs AS Afx_ICoreWebView2PermissionRequestedEventArgs_
TYPE Afx_ICoreWebView2PermissionRequestedEventArgs2 AS Afx_ICoreWebView2PermissionRequestedEventArgs2_
TYPE Afx_ICoreWebView2PermissionRequestedEventArgs3 AS Afx_ICoreWebView2PermissionRequestedEventArgs3_
TYPE Afx_ICoreWebView2PermissionRequestedEventHandler AS Afx_ICoreWebView2PermissionRequestedEventHandler_
TYPE Afx_ICoreWebView2PermissionSetting AS Afx_ICoreWebView2PermissionSetting_
TYPE Afx_ICoreWebView2PermissionSettingCollectionView AS Afx_ICoreWebView2PermissionSettingCollectionView_
TYPE Afx_ICoreWebView2PointerInfo AS Afx_ICoreWebView2PointerInfo_
TYPE Afx_ICoreWebView2PrintCompletedHandler AS Afx_ICoreWebView2PrintCompletedHandler_
TYPE Afx_ICoreWebView2PrintSettings AS Afx_ICoreWebView2PrintSettings_
TYPE Afx_ICoreWebView2PrintSettings2 AS Afx_ICoreWebView2PrintSettings2_
TYPE Afx_ICoreWebView2PrintToPdfCompletedHandler AS Afx_ICoreWebView2PrintToPdfCompletedHandler_
TYPE Afx_ICoreWebView2PrintToPdfStreamCompletedHandler AS Afx_ICoreWebView2PrintToPdfStreamCompletedHandler_
TYPE Afx_ICoreWebView2ProcessExtendedInfo AS Afx_ICoreWebView2ProcessExtendedInfo_
TYPE Afx_ICoreWebView2ProcessExtendedInfoCollection AS Afx_ICoreWebView2ProcessExtendedInfoCollection_
TYPE Afx_ICoreWebView2ProcessFailedEventArgs AS Afx_ICoreWebView2ProcessFailedEventArgs_
TYPE Afx_ICoreWebView2ProcessFailedEventArgs2 AS Afx_ICoreWebView2ProcessFailedEventArgs2_
TYPE Afx_ICoreWebView2ProcessFailedEventArgs3 AS Afx_ICoreWebView2ProcessFailedEventArgs3_
TYPE Afx_ICoreWebView2ProcessFailedEventHandler AS Afx_ICoreWebView2ProcessFailedEventHandler_
TYPE Afx_ICoreWebView2ProcessInfo AS Afx_ICoreWebView2ProcessInfo_
TYPE Afx_ICoreWebView2ProcessInfoCollection AS Afx_ICoreWebView2ProcessInfoCollection_
TYPE Afx_ICoreWebView2ProcessInfosChangedEventHandler AS Afx_ICoreWebView2ProcessInfosChangedEventHandler_
TYPE Afx_ICoreWebView2ProfileAddBrowserExtensionCompletedHandler AS Afx_ICoreWebView2ProfileAddBrowserExtensionCompletedHandler_
TYPE Afx_ICoreWebView2ProfileDeletedEventHandler AS Afx_ICoreWebView2ProfileDeletedEventHandler_
TYPE Afx_ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler AS Afx_ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler_
TYPE Afx_ICoreWebView2RasterizationScaleChangedEventHandler AS Afx_ICoreWebView2RasterizationScaleChangedEventHandler_
TYPE Afx_ICoreWebView2RegionRectCollectionView AS Afx_ICoreWebView2RegionRectCollectionView_
TYPE Afx_ICoreWebView2SaveAsUIShowingEventArgs AS Afx_ICoreWebView2SaveAsUIShowingEventArgs_
TYPE Afx_ICoreWebView2SaveAsUIShowingEventHandler AS Afx_ICoreWebView2SaveAsUIShowingEventHandler_
TYPE Afx_ICoreWebView2SaveFileSecurityCheckStartingEventArgs AS Afx_ICoreWebView2SaveFileSecurityCheckStartingEventArgs_
TYPE Afx_ICoreWebView2SaveFileSecurityCheckStartingEventHandler AS Afx_ICoreWebView2SaveFileSecurityCheckStartingEventHandler_
TYPE Afx_ICoreWebView2ScreenCaptureStartingEventArgs AS Afx_ICoreWebView2ScreenCaptureStartingEventArgs_
TYPE Afx_ICoreWebView2ScreenCaptureStartingEventHandler AS Afx_ICoreWebView2ScreenCaptureStartingEventHandler_
TYPE Afx_ICoreWebView2ScriptDialogOpeningEventArgs AS Afx_ICoreWebView2ScriptDialogOpeningEventArgs_
TYPE Afx_ICoreWebView2ScriptDialogOpeningEventHandler AS Afx_ICoreWebView2ScriptDialogOpeningEventHandler_
TYPE Afx_ICoreWebView2ScriptException AS Afx_ICoreWebView2ScriptException_
TYPE Afx_ICoreWebView2ServerCertificateErrorDetectedEventArgs AS Afx_ICoreWebView2ServerCertificateErrorDetectedEventArgs_
TYPE Afx_ICoreWebView2ServerCertificateErrorDetectedEventHandler AS Afx_ICoreWebView2ServerCertificateErrorDetectedEventHandler_
TYPE Afx_ICoreWebView2SetPermissionStateCompletedHandler AS Afx_ICoreWebView2SetPermissionStateCompletedHandler_
TYPE Afx_ICoreWebView2SharedBuffer AS Afx_ICoreWebView2SharedBuffer_
TYPE Afx_ICoreWebView2ShowSaveAsUICompletedHandler AS Afx_ICoreWebView2ShowSaveAsUICompletedHandler_
TYPE Afx_ICoreWebView2SourceChangedEventArgs AS Afx_ICoreWebView2SourceChangedEventArgs_
TYPE Afx_ICoreWebView2SourceChangedEventHandler AS Afx_ICoreWebView2SourceChangedEventHandler_
TYPE Afx_ICoreWebView2StateChangedEventHandler AS Afx_ICoreWebView2StateChangedEventHandler_
TYPE Afx_ICoreWebView2StatusBarTextChangedEventHandler AS Afx_ICoreWebView2StatusBarTextChangedEventHandler_
TYPE Afx_ICoreWebView2StringCollection AS Afx_ICoreWebView2StringCollection_
TYPE Afx_ICoreWebView2TrySuspendCompletedHandler AS Afx_ICoreWebView2TrySuspendCompletedHandler_
TYPE Afx_ICoreWebView2WebMessageReceivedEventArgs AS Afx_ICoreWebView2WebMessageReceivedEventArgs_
TYPE Afx_ICoreWebView2WebMessageReceivedEventArgs2 AS Afx_ICoreWebView2WebMessageReceivedEventArgs2_
TYPE Afx_ICoreWebView2WebMessageReceivedEventHandler AS Afx_ICoreWebView2WebMessageReceivedEventHandler_
TYPE Afx_ICoreWebView2WebResourceRequest AS Afx_ICoreWebView2WebResourceRequest_
TYPE Afx_ICoreWebView2WebResourceRequestedEventArgs AS Afx_ICoreWebView2WebResourceRequestedEventArgs_
TYPE Afx_ICoreWebView2WebResourceRequestedEventArgs2 AS Afx_ICoreWebView2WebResourceRequestedEventArgs2_
TYPE Afx_ICoreWebView2WebResourceRequestedEventHandler AS Afx_ICoreWebView2WebResourceRequestedEventHandler_
TYPE Afx_ICoreWebView2WebResourceResponse AS Afx_ICoreWebView2WebResourceResponse_
TYPE Afx_ICoreWebView2WebResourceResponseReceivedEventArgs AS Afx_ICoreWebView2WebResourceResponseReceivedEventArgs_
TYPE Afx_ICoreWebView2WebResourceResponseReceivedEventHandler AS Afx_ICoreWebView2WebResourceResponseReceivedEventHandler_
TYPE Afx_ICoreWebView2WebResourceResponseView AS Afx_ICoreWebView2WebResourceResponseView_
TYPE Afx_ICoreWebView2WebResourceResponseViewGetContentCompletedHandler AS Afx_ICoreWebView2WebResourceResponseViewGetContentCompletedHandler_
TYPE Afx_ICoreWebView2WindowCloseRequestedEventHandler AS Afx_ICoreWebView2WindowCloseRequestedEventHandler_
TYPE Afx_ICoreWebView2WindowFeatures AS Afx_ICoreWebView2WindowFeatures_
TYPE Afx_ICoreWebView2ZoomFactorChangedEventHandler AS Afx_ICoreWebView2ZoomFactorChangedEventHandler_

' // Enumerations

ENUM COREWEBVIEW2_BOUNDS_MODE
   COREWEBVIEW2_BOUNDS_MODE_USE_RAW_PIXELS = 0
   COREWEBVIEW2_BOUNDS_MODE_USE_RASTERIZATION_SCALE = 1
END ENUM

ENUM COREWEBVIEW2_BROWSER_PROCESS_EXIT_KIND
   COREWEBVIEW2_BROWSER_PROCESS_EXIT_KIND_NORMAL = 0
   COREWEBVIEW2_BROWSER_PROCESS_EXIT_KIND_FAILED = 1
END ENUM

ENUM COREWEBVIEW2_BROWSING_DATA_KINDS
   COREWEBVIEW2_BROWSING_DATA_KINDS_FILE_SYSTEMS = 1
   COREWEBVIEW2_BROWSING_DATA_KINDS_INDEXED_DB = 2
   COREWEBVIEW2_BROWSING_DATA_KINDS_LOCAL_STORAGE = 4
   COREWEBVIEW2_BROWSING_DATA_KINDS_WEB_SQL = 8
   COREWEBVIEW2_BROWSING_DATA_KINDS_CACHE_STORAGE = 16
   COREWEBVIEW2_BROWSING_DATA_KINDS_ALL_DOM_STORAGE = 32
   COREWEBVIEW2_BROWSING_DATA_KINDS_COOKIES = 64
   COREWEBVIEW2_BROWSING_DATA_KINDS_ALL_SITE = 128
   COREWEBVIEW2_BROWSING_DATA_KINDS_DISK_CACHE = 256
   COREWEBVIEW2_BROWSING_DATA_KINDS_DOWNLOAD_HISTORY = 512
   COREWEBVIEW2_BROWSING_DATA_KINDS_GENERAL_AUTOFILL = 1024
   COREWEBVIEW2_BROWSING_DATA_KINDS_PASSWORD_AUTOSAVE = 2048
   COREWEBVIEW2_BROWSING_DATA_KINDS_BROWSING_HISTORY = 4096
   COREWEBVIEW2_BROWSING_DATA_KINDS_SETTINGS = 8192
   COREWEBVIEW2_BROWSING_DATA_KINDS_ALL_PROFILE = 16384
   COREWEBVIEW2_BROWSING_DATA_KINDS_SERVICE_WORKERS = 32768
END ENUM

ENUM COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT
   COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_PNG = 0
   COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_JPEG = COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_PNG + 1
END ENUM

ENUM COREWEBVIEW2_CHANNEL_SEARCH_KIND
   COREWEBVIEW2_CHANNEL_SEARCH_KIND_MOST_STABLE = 0
   COREWEBVIEW2_CHANNEL_SEARCH_KIND_LEAST_STABLE = 1
END ENUM

ENUM COREWEBVIEW2_CLIENT_CERTIFICATE_KIND
   COREWEBVIEW2_CLIENT_CERTIFICATE_KIND_SMART_CARD = 0
   COREWEBVIEW2_CLIENT_CERTIFICATE_KIND_PIN = 1
   COREWEBVIEW2_CLIENT_CERTIFICATE_KIND_OTHER = 2
END ENUM

ENUM COREWEBVIEW2_CONTEXT_MENU_ITEM_KIND
   COREWEBVIEW2_CONTEXT_MENU_ITEM_KIND_COMMAND = 0
   COREWEBVIEW2_CONTEXT_MENU_ITEM_KIND_CHECK_BOX = 1
   COREWEBVIEW2_CONTEXT_MENU_ITEM_KIND_RADIO = 2
   COREWEBVIEW2_CONTEXT_MENU_ITEM_KIND_SEPARATOR = 3
   COREWEBVIEW2_CONTEXT_MENU_ITEM_KIND_SUBMENU = 4
END ENUM

ENUM COREWEBVIEW2_CONTEXT_MENU_TARGET_KIND
   COREWEBVIEW2_CONTEXT_MENU_TARGET_KIND_PAGE = 0
   COREWEBVIEW2_CONTEXT_MENU_TARGET_KIND_IMAGE = 1
   COREWEBVIEW2_CONTEXT_MENU_TARGET_KIND_SELECTED_TEXT = 2
   COREWEBVIEW2_CONTEXT_MENU_TARGET_KIND_AUDIO = 3
   COREWEBVIEW2_CONTEXT_MENU_TARGET_KIND_VIDEO = 4
END ENUM

ENUM COREWEBVIEW2_COOKIE_SAME_SITE_KIND
   COREWEBVIEW2_COOKIE_SAME_SITE_KIND_NONE = 0
   COREWEBVIEW2_COOKIE_SAME_SITE_KIND_LAX = 1
   COREWEBVIEW2_COOKIE_SAME_SITE_KIND_STRICT = 2
END ENUM

ENUM COREWEBVIEW2_DEFAULT_DOWNLOAD_DIALOG_CORNER_ALIGNMENT
   COREWEBVIEW2_DEFAULT_DOWNLOAD_DIALOG_CORNER_ALIGNMENT_TOP_LEFT = 0
   COREWEBVIEW2_DEFAULT_DOWNLOAD_DIALOG_CORNER_ALIGNMENT_TOP_RIGHT = 1
   COREWEBVIEW2_DEFAULT_DOWNLOAD_DIALOG_CORNER_ALIGNMENT_BOTTOM_LEFT = 2
   COREWEBVIEW2_DEFAULT_DOWNLOAD_DIALOG_CORNER_ALIGNMENT_BOTTOM_RIGHT = 3
END ENUM

ENUM COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_NONE = 0
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_FAILED = 1
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_ACCESS_DENIED = 2
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_NO_SPACE = 3
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_NAME_TOO_LONG = 4
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_TOO_LARGE = 5
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_MALICIOUS = 6
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_TRANSIENT_ERROR = 7
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_BLOCKED_BY_POLICY = 8
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_SECURITY_CHECK_FAILED = 9
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_TOO_SHORT = 10
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_FILE_HASH_MISMATCH = 11
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_NETWORK_FAILED = 12
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_NETWORK_TIMEOUT = 13
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_NETWORK_DISCONNECTED = 14
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_NETWORK_SERVER_DOWN = 15
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_NETWORK_INVALID_REQUEST = 16
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_SERVER_FAILED = 17
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_SERVER_NO_RANGE = 18
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_SERVER_BAD_CONTENT = 19
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_SERVER_UNAUTHORIZED = 20
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_SERVER_CERTIFICATE_PROBLEM = 21
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_SERVER_FORBIDDEN = 22
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_SERVER_UNEXPECTED_RESPONSE = 23
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_SERVER_CONTENT_LENGTH_MISMATCH = 24
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_SERVER_CROSS_ORIGIN_REDIRECT = 25
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_USER_CANCELED = 26
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_USER_SHUTDOWN = 27
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_USER_PAUSED = 28
   COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON_DOWNLOAD_PROCESS_CRASHED = 29
END ENUM

ENUM COREWEBVIEW2_DOWNLOAD_STATE
   COREWEBVIEW2_DOWNLOAD_STATE_IN_PROGRESS = 0
   COREWEBVIEW2_DOWNLOAD_STATE_INTERRUPTED = 1
   COREWEBVIEW2_DOWNLOAD_STATE_COMPLETED = 2
END ENUM

ENUM COREWEBVIEW2_FAVICON_IMAGE_FORMAT
   COREWEBVIEW2_FAVICON_IMAGE_FORMAT_PNG = 0
   COREWEBVIEW2_FAVICON_IMAGE_FORMAT_JPEG = 1
END ENUM

ENUM COREWEBVIEW2_FILE_SYSTEM_HANDLE_KIND
   COREWEBVIEW2_FILE_SYSTEM_HANDLE_KIND_FILE = 0
   COREWEBVIEW2_FILE_SYSTEM_HANDLE_KIND_DIRECTORY = 1
END ENUM

ENUM COREWEBVIEW2_FILE_SYSTEM_HANDLE_PERMISSION
   COREWEBVIEW2_FILE_SYSTEM_HANDLE_PERMISSION_READ_ONLY = 0
   COREWEBVIEW2_FILE_SYSTEM_HANDLE_PERMISSION_READ_WRITE = 1
END ENUM

ENUM COREWEBVIEW2_FRAME_KIND
   COREWEBVIEW2_FRAME_KIND_UNKNOWN = 0
   COREWEBVIEW2_FRAME_KIND_MAIN_FRAME = 1
   COREWEBVIEW2_FRAME_KIND_IFRAME = 2
   COREWEBVIEW2_FRAME_KIND_EMBED = 3
   COREWEBVIEW2_FRAME_KIND_OBJECT = 4
END ENUM

ENUM COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND
   COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_DENY = 0
   COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_ALLOW = 1
   COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_DENY_CORS = 2
END ENUM

ENUM COREWEBVIEW2_KEY_EVENT_KIND
   COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN = 0
   COREWEBVIEW2_KEY_EVENT_KIND_KEY_UP = COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN + 1
   COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN = COREWEBVIEW2_KEY_EVENT_KIND_KEY_UP + 1
   COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_UP = COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN + 1
END ENUM

ENUM COREWEBVIEW2_MEMORY_USAGE_TARGET_LEVEL
   COREWEBVIEW2_MEMORY_USAGE_TARGET_LEVEL_NORMAL = 0
   COREWEBVIEW2_MEMORY_USAGE_TARGET_LEVEL_LOW = 1
END ENUM

ENUM COREWEBVIEW2_MOUSE_EVENT_KIND
   COREWEBVIEW2_MOUSE_EVENT_KIND_HORIZONTAL_WHEEL = 526
   COREWEBVIEW2_MOUSE_EVENT_KIND_LEFT_BUTTON_DOUBLE_CLICK = 515
   COREWEBVIEW2_MOUSE_EVENT_KIND_LEFT_BUTTON_DOWN = 513
   COREWEBVIEW2_MOUSE_EVENT_KIND_LEFT_BUTTON_UP = 514
   COREWEBVIEW2_MOUSE_EVENT_KIND_LEAVE = 675
   COREWEBVIEW2_MOUSE_EVENT_KIND_MIDDLE_BUTTON_DOUBLE_CLICK = 521
   COREWEBVIEW2_MOUSE_EVENT_KIND_MIDDLE_BUTTON_DOWN = 519
   COREWEBVIEW2_MOUSE_EVENT_KIND_MIDDLE_BUTTON_UP = 520
   COREWEBVIEW2_MOUSE_EVENT_KIND_MOVE = 512
   COREWEBVIEW2_MOUSE_EVENT_KIND_RIGHT_BUTTON_DOUBLE_CLICK = 518
   COREWEBVIEW2_MOUSE_EVENT_KIND_RIGHT_BUTTON_DOWN = 516
   COREWEBVIEW2_MOUSE_EVENT_KIND_RIGHT_BUTTON_UP = 517
   COREWEBVIEW2_MOUSE_EVENT_KIND_WHEEL = 522
   COREWEBVIEW2_MOUSE_EVENT_KIND_X_BUTTON_DOUBLE_CLICK = 525
   COREWEBVIEW2_MOUSE_EVENT_KIND_X_BUTTON_DOWN = 523
   COREWEBVIEW2_MOUSE_EVENT_KIND_X_BUTTON_UP = 524
   COREWEBVIEW2_MOUSE_EVENT_KIND_NON_CLIENT_RIGHT_BUTTON_DOWN = 164
   COREWEBVIEW2_MOUSE_EVENT_KIND_NON_CLIENT_RIGHT_BUTTON_UP = 165
END ENUM

ENUM COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS
   COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_NONE = 0
   COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_LEFT_BUTTON = 1
   COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_RIGHT_BUTTON = 2
   COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_SHIFT = 4
   COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_CONTROL = 8
   COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_MIDDLE_BUTTON = 16
   COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_X_BUTTON1 = 32
   COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS_X_BUTTON2 = 64
END ENUM

ENUM COREWEBVIEW2_MOVE_FOCUS_REASON
   COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC = 0
   COREWEBVIEW2_MOVE_FOCUS_REASON_NEXT = COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC + 1
   COREWEBVIEW2_MOVE_FOCUS_REASON_PREVIOUS = COREWEBVIEW2_MOVE_FOCUS_REASON_NEXT + 1
END ENUM

ENUM COREWEBVIEW2_NAVIGATION_KIND
   COREWEBVIEW2_NAVIGATION_KIND_RELOAD = 0
   COREWEBVIEW2_NAVIGATION_KIND_BACK_OR_FORWARD = 1
   COREWEBVIEW2_NAVIGATION_KIND_NEW_DOCUMENT = 2
END ENUM

ENUM COREWEBVIEW2_NON_CLIENT_REGION_KIND
   COREWEBVIEW2_NON_CLIENT_REGION_KIND_NOWHERE = 0
   COREWEBVIEW2_NON_CLIENT_REGION_KIND_CLIENT = 1
   COREWEBVIEW2_NON_CLIENT_REGION_KIND_CAPTION = 2
   COREWEBVIEW2_NON_CLIENT_REGION_KIND_MINIMIZE = 8
   COREWEBVIEW2_NON_CLIENT_REGION_KIND_MAXIMIZE = 9
   COREWEBVIEW2_NON_CLIENT_REGION_KIND_CLOSE = 20
END ENUM

ENUM COREWEBVIEW2_PDF_TOOLBAR_ITEMS
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_NONE = 0
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_SAVE = 1
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_PRINT = 2
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_SAVE_AS = 4
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_ZOOM_IN = 8
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_ZOOM_OUT = 16
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_ROTATE = 32
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_FIT_PAGE = 64
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_PAGE_LAYOUT = 128
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_BOOKMARKS = 256
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_PAGE_SELECTOR = 512
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_SEARCH = 1024
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_FULL_SCREEN = 2048
   COREWEBVIEW2_PDF_TOOLBAR_ITEMS_MORE_SETTINGS = 4096
END ENUM

ENUM COREWEBVIEW2_PERMISSION_KIND
   COREWEBVIEW2_PERMISSION_KIND_UNKNOWN_PERMISSION = 0
   COREWEBVIEW2_PERMISSION_KIND_MICROPHONE = 1
   COREWEBVIEW2_PERMISSION_KIND_CAMERA = 2
   COREWEBVIEW2_PERMISSION_KIND_GEOLOCATION = 3
   COREWEBVIEW2_PERMISSION_KIND_NOTIFICATIONS = 4
   COREWEBVIEW2_PERMISSION_KIND_OTHER_SENSORS = 5
   COREWEBVIEW2_PERMISSION_KIND_CLIPBOARD_READ = 6
   COREWEBVIEW2_PERMISSION_KIND_MULTIPLE_AUTOMATIC_DOWNLOADS = 7
   COREWEBVIEW2_PERMISSION_KIND_FILE_READ_WRITE = 8
   COREWEBVIEW2_PERMISSION_KIND_AUTOPLAY = 9
   COREWEBVIEW2_PERMISSION_KIND_LOCAL_FONTS = 10
   COREWEBVIEW2_PERMISSION_KIND_MIDI_SYSTEM_EXCLUSIVE_MESSAGES = 11
   COREWEBVIEW2_PERMISSION_KIND_WINDOW_MANAGEMENT = 12
END ENUM

ENUM COREWEBVIEW2_PERMISSION_STATE
   COREWEBVIEW2_PERMISSION_STATE_DEFAULT = 0
   COREWEBVIEW2_PERMISSION_STATE_ALLOW = COREWEBVIEW2_PERMISSION_STATE_DEFAULT + 1
   COREWEBVIEW2_PERMISSION_STATE_DENY = COREWEBVIEW2_PERMISSION_STATE_ALLOW + 1
END ENUM

ENUM COREWEBVIEW2_POINTER_EVENT_KIND
   COREWEBVIEW2_POINTER_EVENT_KIND_ACTIVATE = 587
   COREWEBVIEW2_POINTER_EVENT_KIND_DOWN = 582
   COREWEBVIEW2_POINTER_EVENT_KIND_ENTER = 585
   COREWEBVIEW2_POINTER_EVENT_KIND_LEAVE = 586
   COREWEBVIEW2_POINTER_EVENT_KIND_UP = 583
   COREWEBVIEW2_POINTER_EVENT_KIND_UPDATE = 581
END ENUM

ENUM COREWEBVIEW2_PREFERRED_COLOR_SCHEME
   COREWEBVIEW2_PREFERRED_COLOR_SCHEME_AUTO = 0
   COREWEBVIEW2_PREFERRED_COLOR_SCHEME_LIGHT = 1
   COREWEBVIEW2_PREFERRED_COLOR_SCHEME_DARK = 2
END ENUM

ENUM COREWEBVIEW2_PRINT_COLLATION
   COREWEBVIEW2_PRINT_COLLATION_DEFAULT = 0
   COREWEBVIEW2_PRINT_COLLATION_COLLATED = 1
   COREWEBVIEW2_PRINT_COLLATION_UNCOLLATED = 2
END ENUM

ENUM COREWEBVIEW2_PRINT_COLOR_MODE
   COREWEBVIEW2_PRINT_COLOR_MODE_DEFAULT = 0
   COREWEBVIEW2_PRINT_COLOR_MODE_COLOR = 1
   COREWEBVIEW2_PRINT_COLOR_MODE_GRAYSCALE = 2
END ENUM

ENUM COREWEBVIEW2_PRINT_DIALOG_KIND
   COREWEBVIEW2_PRINT_DIALOG_KIND_BROWSER = 0
   COREWEBVIEW2_PRINT_DIALOG_KIND_SYSTEM = 1
END ENUM

ENUM COREWEBVIEW2_PRINT_DUPLEX
   COREWEBVIEW2_PRINT_DUPLEX_DEFAULT = 0
   COREWEBVIEW2_PRINT_DUPLEX_ONE_SIDED = 1
   COREWEBVIEW2_PRINT_DUPLEX_TWO_SIDED_LONG_EDGE = 2
   COREWEBVIEW2_PRINT_DUPLEX_TWO_SIDED_SHORT_EDGE = 3
END ENUM

ENUM COREWEBVIEW2_PRINT_MEDIA_SIZE
   COREWEBVIEW2_PRINT_MEDIA_SIZE_DEFAULT = 0
   COREWEBVIEW2_PRINT_MEDIA_SIZE_CUSTOM = 1
END ENUM

ENUM COREWEBVIEW2_PRINT_ORIENTATION
   COREWEBVIEW2_PRINT_ORIENTATION_PORTRAIT = 0
   COREWEBVIEW2_PRINT_ORIENTATION_LANDSCAPE = 1
END ENUM

ENUM COREWEBVIEW2_PRINT_STATUS
   COREWEBVIEW2_PRINT_STATUS_SUCCEEDED = 0
   COREWEBVIEW2_PRINT_STATUS_PRINTER_UNAVAILABLE = 1
   COREWEBVIEW2_PRINT_STATUS_OTHER_ERROR = 2
END ENUM

ENUM COREWEBVIEW2_PROCESS_FAILED_KIND
   COREWEBVIEW2_PROCESS_FAILED_KIND_BROWSER_PROCESS_EXITED = 0
   COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_EXITED = 1
   COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_UNRESPONSIVE = 2
   COREWEBVIEW2_PROCESS_FAILED_KIND_FRAME_RENDER_PROCESS_EXITED = 3
   COREWEBVIEW2_PROCESS_FAILED_KIND_UTILITY_PROCESS_EXITED = 4
   COREWEBVIEW2_PROCESS_FAILED_KIND_SANDBOX_HELPER_PROCESS_EXITED = 5
   COREWEBVIEW2_PROCESS_FAILED_KIND_GPU_PROCESS_EXITED = 6
   COREWEBVIEW2_PROCESS_FAILED_KIND_PPAPI_PLUGIN_PROCESS_EXITED = 7
   COREWEBVIEW2_PROCESS_FAILED_KIND_PPAPI_BROKER_PROCESS_EXITED = 8
   COREWEBVIEW2_PROCESS_FAILED_KIND_UNKNOWN_PROCESS_EXITED = 9
END ENUM

ENUM COREWEBVIEW2_PROCESS_FAILED_REASON
   COREWEBVIEW2_PROCESS_FAILED_REASON_UNEXPECTED = 0
   COREWEBVIEW2_PROCESS_FAILED_REASON_UNRESPONSIVE = 1
   COREWEBVIEW2_PROCESS_FAILED_REASON_TERMINATED = 2
   COREWEBVIEW2_PROCESS_FAILED_REASON_CRASHED = 3
   COREWEBVIEW2_PROCESS_FAILED_REASON_LAUNCH_FAILED = 4
   COREWEBVIEW2_PROCESS_FAILED_REASON_OUT_OF_MEMORY = 5
   COREWEBVIEW2_PROCESS_FAILED_REASON_PROFILE_DELETED = 6
END ENUM

ENUM COREWEBVIEW2_PROCESS_KIND
   COREWEBVIEW2_PROCESS_KIND_BROWSER = 0
   COREWEBVIEW2_PROCESS_KIND_RENDERER = 1
   COREWEBVIEW2_PROCESS_KIND_UTILITY = 2
   COREWEBVIEW2_PROCESS_KIND_SANDBOX_HELPER = 3
   COREWEBVIEW2_PROCESS_KIND_GPU = 4
   COREWEBVIEW2_PROCESS_KIND_PPAPI_PLUGIN = 5
   COREWEBVIEW2_PROCESS_KIND_PPAPI_BROKER = 6
END ENUM

ENUM COREWEBVIEW2_RELEASE_CHANNELS
   COREWEBVIEW2_RELEASE_CHANNELS_NONE = 0
   COREWEBVIEW2_RELEASE_CHANNELS_STABLE = 1
   COREWEBVIEW2_RELEASE_CHANNELS_BETA = 2
   COREWEBVIEW2_RELEASE_CHANNELS_DEV = 4
   COREWEBVIEW2_RELEASE_CHANNELS_CANARY = 8
END ENUM

ENUM COREWEBVIEW2_SAVE_AS_KIND
   COREWEBVIEW2_SAVE_AS_KIND_DEFAULT = 0
   COREWEBVIEW2_SAVE_AS_KIND_HTML_ONLY = 1
   COREWEBVIEW2_SAVE_AS_KIND_SINGLE_FILE = 2
   COREWEBVIEW2_SAVE_AS_KIND_COMPLETE = 3
END ENUM

ENUM COREWEBVIEW2_SAVE_AS_UI_RESULT
   COREWEBVIEW2_SAVE_AS_UI_RESULT_SUCCESS = 0
   COREWEBVIEW2_SAVE_AS_UI_RESULT_INVALID_PATH = 1
   COREWEBVIEW2_SAVE_AS_UI_RESULT_FILE_ALREADY_EXISTS = 2
   COREWEBVIEW2_SAVE_AS_UI_RESULT_KIND_NOT_SUPPORTED = 3
   COREWEBVIEW2_SAVE_AS_UI_RESULT_CANCELLED = 4
END ENUM

ENUM COREWEBVIEW2_SCRIPT_DIALOG_KIND
   COREWEBVIEW2_SCRIPT_DIALOG_KIND_ALERT = 0
   COREWEBVIEW2_SCRIPT_DIALOG_KIND_CONFIRM = 1 = COREWEBVIEW2_SCRIPT_DIALOG_KIND_ALERT + 1
   COREWEBVIEW2_SCRIPT_DIALOG_KIND_PROMPT = COREWEBVIEW2_SCRIPT_DIALOG_KIND_CONFIRM + 1
   COREWEBVIEW2_SCRIPT_DIALOG_KIND_BEFOREUNLOAD = COREWEBVIEW2_SCRIPT_DIALOG_KIND_PROMPT + 1
END ENUM

ENUM COREWEBVIEW2_SCROLLBAR_STYLE
   COREWEBVIEW2_SCROLLBAR_STYLE_DEFAULT = 0
   COREWEBVIEW2_SCROLLBAR_STYLE_FLUENT_OVERLAY = 1
END ENUM

ENUM COREWEBVIEW2_SERVER_CERTIFICATE_ERROR_ACTION
   COREWEBVIEW2_SERVER_CERTIFICATE_ERROR_ACTION_ALWAYS_ALLOW = 0
   COREWEBVIEW2_SERVER_CERTIFICATE_ERROR_ACTION_CANCEL = 1
   COREWEBVIEW2_SERVER_CERTIFICATE_ERROR_ACTION_DEFAULT = 2
END ENUM

ENUM COREWEBVIEW2_SHARED_BUFFER_ACCESS
   COREWEBVIEW2_SHARED_BUFFER_ACCESS_READ_ONLY = 0
   COREWEBVIEW2_SHARED_BUFFER_ACCESS_READ_WRITE = 1
END ENUM

ENUM COREWEBVIEW2_TEXT_DIRECTION_KIND
   COREWEBVIEW2_TEXT_DIRECTION_KIND_DEFAULT = 0
   COREWEBVIEW2_TEXT_DIRECTION_KIND_LEFT_TO_RIGHT = 1
   COREWEBVIEW2_TEXT_DIRECTION_KIND_RIGHT_TO_LEFT = 2
END ENUM

ENUM COREWEBVIEW2_TRACKING_PREVENTION_LEVEL
   COREWEBVIEW2_TRACKING_PREVENTION_LEVEL_NONE = 0
   COREWEBVIEW2_TRACKING_PREVENTION_LEVEL_BASIC = 1
   COREWEBVIEW2_TRACKING_PREVENTION_LEVEL_BALANCED = 2
   COREWEBVIEW2_TRACKING_PREVENTION_LEVEL_STRICT = 3
END ENUM

ENUM COREWEBVIEW2_WEB_ERROR_STATUS
   COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN = 0
   COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_COMMON_NAME_IS_INCORRECT = 1
   COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_EXPIRED = 2
   COREWEBVIEW2_WEB_ERROR_STATUS_CLIENT_CERTIFICATE_CONTAINS_ERRORS = 3
   COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_REVOKED = 4
   COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_IS_INVALID = 5
   COREWEBVIEW2_WEB_ERROR_STATUS_SERVER_UNREACHABLE = 6
   COREWEBVIEW2_WEB_ERROR_STATUS_TIMEOUT = 7
   COREWEBVIEW2_WEB_ERROR_STATUS_ERROR_HTTP_INVALID_SERVER_RESPONSE = 8
   COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_ABORTED = 9
   COREWEBVIEW2_WEB_ERROR_STATUS_CONNECTION_RESET = 10
   COREWEBVIEW2_WEB_ERROR_STATUS_DISCONNECTED = 11
   COREWEBVIEW2_WEB_ERROR_STATUS_CANNOT_CONNECT = 12
   COREWEBVIEW2_WEB_ERROR_STATUS_HOST_NAME_NOT_RESOLVED = 13
   COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED = 14
   COREWEBVIEW2_WEB_ERROR_STATUS_REDIRECT_FAILED = 15
   COREWEBVIEW2_WEB_ERROR_STATUS_UNEXPECTED_ERROR = 16
   COREWEBVIEW2_WEB_ERROR_STATUS_VALID_AUTHENTICATION_CREDENTIALS_REQUIRED = 17
   COREWEBVIEW2_WEB_ERROR_STATUS_VALID_PROXY_AUTHENTICATION_REQUIRED = 18
END ENUM

ENUM COREWEBVIEW2_WEB_RESOURCE_CONTEXT
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL = 0
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT = 1
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_STYLESHEET = 2
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE = 3
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MEDIA = 4
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FONT = 5
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SCRIPT = 6
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_XML_HTTP_REQUEST = 7
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_FETCH = 8
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_TEXT_TRACK = 9
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_EVENT_SOURCE = 10
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_WEBSOCKET = 11
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_MANIFEST = 12
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_SIGNED_EXCHANGE = 13
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_PING = 14
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_CSP_VIOLATION_REPORT = 15
   COREWEBVIEW2_WEB_RESOURCE_CONTEXT_OTHER = 16
END ENUM

ENUM COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS
   COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_NONE = 0
   COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_DOCUMENT = 1
   COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_SHARED_WORKER = 2
   COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_SERVICE_WORKER = 4
   COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_ALL = -1
END ENUM

TYPE COREWEBVIEW2_COLOR
   A AS UBYTE
   R AS UBYTE
   G AS UBYTE
   B AS UBYTE
END TYPE

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
' Interface name: ICoreWebView2
' IID: {76ECEACB-0462-4D94-AC83-423A6793775E}
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
' Interface name: ICoreWebView2_2
' IID: {9E8F0CF8-E670-4B5E-B2BC-73E061E3184C}
' Inherited interface = ICoreWebView2
' ########################################################################################

#ifndef __Afx_ICoreWebView2_2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_2_ EXTENDS Afx_ICoreWebView2
   DECLARE ABSTRACT FUNCTION add_WebResourceResponseReceived (BYVAL eventHandler AS Afx_ICoreWebView2WebResourceResponseReceivedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_WebResourceResponseReceived (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION NavigateWithWebResourceRequest (BYVAL Request AS Afx_ICoreWebView2WebResourceRequest PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_DOMContentLoaded (BYVAL eventHandler AS Afx_ICoreWebView2DOMContentLoadedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_DOMContentLoaded (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CookieManager (BYVAL CookieManager AS Afx_ICoreWebView2CookieManager PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Environment (BYVAL Environment AS Afx_ICoreWebView2Environment PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_3
' IID: {A0D6DF20-3B92-416D-AA0C-437A9C727857}
' Inherited interface = ICoreWebView2_2
' ########################################################################################

#ifndef __Afx_ICoreWebView2_3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_3_ EXTENDS Afx_ICoreWebView2_2
   DECLARE ABSTRACT FUNCTION TrySuspend (BYVAL handler AS Afx_ICoreWebView2TrySuspendCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Resume () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsSuspended (BYVAL IsSuspended AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetVirtualHostNameToFolderMapping (BYVAL hostName AS WSTRING PTR, BYVAL folderPath AS WSTRING PTR, BYVAL accessKind AS COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND) AS HRESULT
   DECLARE ABSTRACT FUNCTION ClearVirtualHostNameToFolderMapping (BYVAL hostName AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_4
' IID: {20D02D59-6DF2-42DC-BD06-F98A694B1302}
' Inherited interface = ICoreWebView2_3
' ########################################################################################

#ifndef __Afx_ICoreWebView2_4_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_4_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_4_ EXTENDS Afx_ICoreWebView2_3
   DECLARE ABSTRACT FUNCTION add_FrameCreated (BYVAL eventHandler AS Afx_ICoreWebView2FrameCreatedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_FrameCreated (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_DownloadStarting (BYVAL eventHandler AS Afx_ICoreWebView2DownloadStartingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_DownloadStarting (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_5
' IID: {BEDB11B8-D63C-11EB-B8BC-0242AC130003}
' Inherited interface = ICoreWebView2_4
' ########################################################################################

#ifndef __Afx_ICoreWebView2_5_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_5_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_5_ EXTENDS Afx_ICoreWebView2_4
   DECLARE ABSTRACT FUNCTION add_ClientCertificateRequested (BYVAL eventHandler AS Afx_ICoreWebView2ClientCertificateRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ClientCertificateRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_6
' IID: {499AADAC-D92C-4589-8A75-111BFC167795}
' Inherited interface = ICoreWebView2_5
' ########################################################################################

#ifndef __Afx_ICoreWebView2_6_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_6_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_6_ EXTENDS Afx_ICoreWebView2_5
   DECLARE ABSTRACT FUNCTION OpenTaskManagerWindow () AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_7
' IID: {79C24D83-09A3-45AE-9418-487F32A58740}
' Inherited interface = ICoreWebView2_6
' ########################################################################################

#ifndef __Afx_ICoreWebView2_7_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_7_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_7_ EXTENDS Afx_ICoreWebView2_6
   DECLARE ABSTRACT FUNCTION PrintToPdf (BYVAL ResultFilePath AS WSTRING PTR, BYVAL printSettings AS Afx_ICoreWebView2PrintSettings PTR, BYVAL handler AS Afx_ICoreWebView2PrintToPdfCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_8
' IID: {E9632730-6E1E-43AB-B7B8-7B2C9E62E094}
' Inherited interface = ICoreWebView2_7
' ########################################################################################

#ifndef __Afx_ICoreWebView2_8_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_8_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_8_ EXTENDS Afx_ICoreWebView2_7
   DECLARE ABSTRACT FUNCTION add_IsMutedChanged (BYVAL eventHandler AS Afx_ICoreWebView2IsMutedChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_IsMutedChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsMuted (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsMuted (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_IsDocumentPlayingAudioChanged (BYVAL eventHandler AS Afx_ICoreWebView2IsDocumentPlayingAudioChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_IsDocumentPlayingAudioChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsDocumentPlayingAudio (BYVAL value AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_9
' IID: {4D7B2EAB-9FDC-468D-B998-A9260B5ED651}
' Inherited interface = ICoreWebView2_8
' ########################################################################################

#ifndef __Afx_ICoreWebView2_9_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_9_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_9_ EXTENDS Afx_ICoreWebView2_8
   DECLARE ABSTRACT FUNCTION add_IsDefaultDownloadDialogOpenChanged (BYVAL handler AS Afx_ICoreWebView2IsDefaultDownloadDialogOpenChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_IsDefaultDownloadDialogOpenChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsDefaultDownloadDialogOpen (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OpenDefaultDownloadDialog () AS HRESULT
   DECLARE ABSTRACT FUNCTION CloseDefaultDownloadDialog () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DefaultDownloadDialogCornerAlignment (BYVAL value AS COREWEBVIEW2_DEFAULT_DOWNLOAD_DIALOG_CORNER_ALIGNMENT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DefaultDownloadDialogCornerAlignment (BYVAL value AS COREWEBVIEW2_DEFAULT_DOWNLOAD_DIALOG_CORNER_ALIGNMENT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DefaultDownloadDialogMargin (BYVAL value AS POINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DefaultDownloadDialogMargin (BYVAL value AS POINT) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_10
' IID: {B1690564-6F5A-4983-8E48-31D1143FECDB}
' Inherited interface = ICoreWebView2_9
' ########################################################################################

#ifndef __Afx_ICoreWebView2_10_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_10_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_10_ EXTENDS Afx_ICoreWebView2_9
   DECLARE ABSTRACT FUNCTION add_BasicAuthenticationRequested (BYVAL eventHandler AS Afx_ICoreWebView2BasicAuthenticationRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_BasicAuthenticationRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_11
' IID: {0BE78E56-C193-4051-B943-23B460C08BDB}
' Inherited interface = ICoreWebView2_10
' ########################################################################################

#ifndef __Afx_ICoreWebView2_11_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_11_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_11_ EXTENDS Afx_ICoreWebView2_10
   DECLARE ABSTRACT FUNCTION CallDevToolsProtocolMethodForSession (BYVAL sessionId AS WSTRING PTR, BYVAL methodName AS WSTRING PTR, BYVAL parametersAsJson AS WSTRING PTR, BYVAL handler AS Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ContextMenuRequested (BYVAL eventHandler AS Afx_ICoreWebView2ContextMenuRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ContextMenuRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_12
' IID: {35D69927-BCFA-4566-9349-6B3E0D154CAC}
' Inherited interface = ICoreWebView2_11
' ########################################################################################

#ifndef __Afx_ICoreWebView2_12_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_12_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_12_ EXTENDS Afx_ICoreWebView2_11
   DECLARE ABSTRACT FUNCTION add_StatusBarTextChanged (BYVAL eventHandler AS Afx_ICoreWebView2StatusBarTextChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_StatusBarTextChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_StatusBarText (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_13
' IID: {F75F09A8-667E-4983-88D6-C8773F315E84}
' Inherited interface = ICoreWebView2_12
' ########################################################################################

#ifndef __Afx_ICoreWebView2_13_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_13_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_13_ EXTENDS Afx_ICoreWebView2_12
   DECLARE ABSTRACT FUNCTION get_Profile (BYVAL value AS Afx_ICoreWebView2Profile PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_14
' IID: {6DAA4F10-4A90-4753-8898-77C5DF534165}
' Inherited interface = ICoreWebView2_13
' ########################################################################################

#ifndef __Afx_ICoreWebView2_14_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_14_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_14_ EXTENDS Afx_ICoreWebView2_13
   DECLARE ABSTRACT FUNCTION add_ServerCertificateErrorDetected (BYVAL eventHandler AS Afx_ICoreWebView2ServerCertificateErrorDetectedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ServerCertificateErrorDetected (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION ClearServerCertificateErrorActions (BYVAL handler AS Afx_ICoreWebView2ClearServerCertificateErrorActionsCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_15
' IID: {517B2D1D-7DAE-4A66-A4F4-10352FFB9518}
' Inherited interface = ICoreWebView2_14
' ########################################################################################

#ifndef __Afx_ICoreWebView2_15_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_15_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_15_ EXTENDS Afx_ICoreWebView2_14
   DECLARE ABSTRACT FUNCTION add_FaviconChanged (BYVAL eventHandler AS Afx_ICoreWebView2FaviconChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_FaviconChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FaviconUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFavicon (BYVAL format AS COREWEBVIEW2_FAVICON_IMAGE_FORMAT, BYVAL completedHandler AS Afx_ICoreWebView2GetFaviconCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_16
' IID: {0EB34DC9-9F91-41E1-8639-95CD5943906B}
' Inherited interface = ICoreWebView2_15
' ########################################################################################

#ifndef __Afx_ICoreWebView2_16_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_16_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_16_ EXTENDS Afx_ICoreWebView2_15
   DECLARE ABSTRACT FUNCTION Print (BYVAL printSettings AS Afx_ICoreWebView2PrintSettings PTR, BYVAL handler AS Afx_ICoreWebView2PrintCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ShowPrintUI (BYVAL printDialogKind AS COREWEBVIEW2_PRINT_DIALOG_KIND) AS HRESULT
   DECLARE ABSTRACT FUNCTION PrintToPdfStream (BYVAL printSettings AS Afx_ICoreWebView2PrintSettings PTR, BYVAL handler AS Afx_ICoreWebView2PrintToPdfStreamCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_17
' IID: {702E75D4-FD44-434D-9D70-1A68A6B1192A}
' Inherited interface = ICoreWebView2_16
' ########################################################################################

#ifndef __Afx_ICoreWebView2_17_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_17_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_17_ EXTENDS Afx_ICoreWebView2_16
   DECLARE ABSTRACT FUNCTION PostSharedBufferToScript (BYVAL sharedBuffer AS Afx_ICoreWebView2SharedBuffer PTR, BYVAL access AS COREWEBVIEW2_SHARED_BUFFER_ACCESS, BYVAL additionalDataAsJson AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_18
' IID: {7A626017-28BE-49B2-B865-3BA2B3522D90}
' Inherited interface = ICoreWebView2_17
' ########################################################################################

#ifndef __Afx_ICoreWebView2_18_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_18_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_18_ EXTENDS Afx_ICoreWebView2_17
   DECLARE ABSTRACT FUNCTION add_LaunchingExternalUriScheme (BYVAL eventHandler AS Afx_ICoreWebView2LaunchingExternalUriSchemeEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_LaunchingExternalUriScheme (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_19
' IID: {6921F954-79B0-437F-A997-C85811897C68}
' Inherited interface = ICoreWebView2_18
' ########################################################################################

#ifndef __Afx_ICoreWebView2_19_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_19_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_19_ EXTENDS Afx_ICoreWebView2_18
   DECLARE ABSTRACT FUNCTION get_MemoryUsageTargetLevel (BYVAL value AS COREWEBVIEW2_MEMORY_USAGE_TARGET_LEVEL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MemoryUsageTargetLevel (BYVAL value AS COREWEBVIEW2_MEMORY_USAGE_TARGET_LEVEL) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_20
' IID: {B4BC1926-7305-11EE-B962-0242AC120002}
' Inherited interface = ICoreWebView2_19
' ########################################################################################

#ifndef __Afx_ICoreWebView2_20_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_20_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_20_ EXTENDS Afx_ICoreWebView2_19
   DECLARE ABSTRACT FUNCTION get_FrameId (BYVAL value AS UINT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_21
' IID: {C4980DEA-587B-43B9-8143-3EF3BF552D95}
' Inherited interface = ICoreWebView2_20
' ########################################################################################

#ifndef __Afx_ICoreWebView2_21_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_21_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_21_ EXTENDS Afx_ICoreWebView2_20
   DECLARE ABSTRACT FUNCTION ExecuteScriptWithResult (BYVAL javaScript AS WSTRING PTR, BYVAL handler AS Afx_ICoreWebView2ExecuteScriptWithResultCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_22
' IID: {DB75DFC7-A857-4632-A398-6969DDE26C0A}
' Inherited interface = ICoreWebView2_21
' ########################################################################################

#ifndef __Afx_ICoreWebView2_22_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_22_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_22_ EXTENDS Afx_ICoreWebView2_21
   DECLARE ABSTRACT FUNCTION AddWebResourceRequestedFilterWithRequestSourceKinds (BYVAL uri AS WSTRING PTR, BYVAL ResourceContext AS COREWEBVIEW2_WEB_RESOURCE_CONTEXT, BYVAL requestSourceKinds AS COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveWebResourceRequestedFilterWithRequestSourceKinds (BYVAL uri AS WSTRING PTR, BYVAL ResourceContext AS COREWEBVIEW2_WEB_RESOURCE_CONTEXT, BYVAL requestSourceKinds AS COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_23
' IID: {508F0DB5-90C4-5872-90A7-267A91377502}
' Inherited interface = ICoreWebView2_22
' ########################################################################################

#ifndef __Afx_ICoreWebView2_23_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_23_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_23_ EXTENDS Afx_ICoreWebView2_22
   DECLARE ABSTRACT FUNCTION PostWebMessageAsJsonWithAdditionalObjects (BYVAL webMessageAsJson AS WSTRING PTR, BYVAL additionalObjects AS Afx_ICoreWebView2ObjectCollectionView PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_24
' IID: {39A7AD55-4287-5CC1-88A1-C6F458593824}
' Inherited interface = ICoreWebView2_23
' ########################################################################################

#ifndef __Afx_ICoreWebView2_24_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_24_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_24_ EXTENDS Afx_ICoreWebView2_23
   DECLARE ABSTRACT FUNCTION add_NotificationReceived (BYVAL eventHandler AS Afx_ICoreWebView2NotificationReceivedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NotificationReceived (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_25
' IID: {B5A86092-DF50-5B4F-A17B-6C8F8B40B771}
' Inherited interface = ICoreWebView2_24
' ########################################################################################

#ifndef __Afx_ICoreWebView2_25_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_25_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_25_ EXTENDS Afx_ICoreWebView2_24
   DECLARE ABSTRACT FUNCTION add_SaveAsUIShowing (BYVAL eventHandler AS Afx_ICoreWebView2SaveAsUIShowingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_SaveAsUIShowing (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION ShowSaveAsUI (BYVAL handler AS Afx_ICoreWebView2ShowSaveAsUICompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_26
' IID: {806268B8-F897-5685-88E5-C45FCA0B1A48}
' Inherited interface = ICoreWebView2_25
' ########################################################################################

#ifndef __Afx_ICoreWebView2_26_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_26_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_26_ EXTENDS Afx_ICoreWebView2_25
   DECLARE ABSTRACT FUNCTION add_SaveFileSecurityCheckStarting (BYVAL eventHandler AS Afx_ICoreWebView2SaveFileSecurityCheckStartingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_SaveFileSecurityCheckStarting (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_27
' IID: {00FBE33B-8C07-517C-AA23-0DDD4B5F6FA0}
' Inherited interface = ICoreWebView2_26
' ########################################################################################

#ifndef __Afx_ICoreWebView2_27_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_27_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_27_ EXTENDS Afx_ICoreWebView2_26
   DECLARE ABSTRACT FUNCTION add_ScreenCaptureStarting (BYVAL eventHandler AS Afx_ICoreWebView2ScreenCaptureStartingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ScreenCaptureStarting (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2_28
' IID: {62E50381-5BF5-51A8-AAE0-F20A3A9C8A90}
' Inherited interface = ICoreWebView2_27
' ########################################################################################

#ifndef __Afx_ICoreWebView2_28_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2_28_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2_28_ EXTENDS Afx_ICoreWebView2_27
   DECLARE ABSTRACT FUNCTION get_Find (BYVAL value AS Afx_ICoreWebView2Find PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2AcceleratorKeyPressedEventArgs
' IID: {9F760F8A-FB79-42BE-9990-7B56900FA9C7}
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
' Interface name: ICoreWebView2AcceleratorKeyPressedEventArgs2
' IID: {03B2C8C8-7799-4E34-BD66-ED26AA85F2BF}
' Inherited interface = ICoreWebView2AcceleratorKeyPressedEventArgs
' ########################################################################################

#ifndef __Afx_ICoreWebView2AcceleratorKeyPressedEventArgs2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2AcceleratorKeyPressedEventArgs2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2AcceleratorKeyPressedEventArgs2_ EXTENDS Afx_ICoreWebView2AcceleratorKeyPressedEventArgs
   DECLARE ABSTRACT FUNCTION get_IsBrowserAcceleratorKeyEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsBrowserAcceleratorKeyEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2AcceleratorKeyPressedEventHandler
' IID: {B29C7E28-FA79-41A8-8E44-65811C76DCB2}
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
' IID: {B99369F3-9B11-47B5-BC6F-8E7895FCEA17}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BasicAuthenticationRequestedEventArgs
' IID: {EF05516F-D897-4F9E-B672-D8E2307A3FB0}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BasicAuthenticationRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BasicAuthenticationRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BasicAuthenticationRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Challenge (BYVAL Challenge AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Response (BYVAL Response AS Afx_ICoreWebView2BasicAuthenticationResponse PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Cancel (BYVAL Cancel AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Cancel (BYVAL Cancel AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BasicAuthenticationRequestedEventHandler
' IID: {58B4D6C2-18D4-497E-B39B-9A96533FA278}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BasicAuthenticationRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BasicAuthenticationRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BasicAuthenticationRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2BasicAuthenticationRequestedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BasicAuthenticationResponse
' IID: {07023F7D-2D77-4D67-9040-6E7D428C6A40}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BasicAuthenticationResponse_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BasicAuthenticationResponse_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BasicAuthenticationResponse_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_UserName (BYVAL UserName AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_UserName (BYVAL UserName AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Password (BYVAL Password AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Password (BYVAL Password AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BrowserExtension
' IID: {7EF7FFA0-FAC5-462C-B189-3D9EDBE575DA}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BrowserExtension_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BrowserExtension_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BrowserExtension_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_id (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_name (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Remove (BYVAL handler AS Afx_ICoreWebView2BrowserExtensionRemoveCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Enable (BYVAL IsEnabled AS LONG, BYVAL handler AS Afx_ICoreWebView2BrowserExtensionEnableCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BrowserExtensionEnableCompletedHandler
' IID: {30C186CE-7FAD-421F-A3BC-A8EAF071DDB8}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BrowserExtensionEnableCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BrowserExtensionEnableCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BrowserExtensionEnableCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BrowserExtensionList
' IID: {2EF3D2DC-BD5F-4F4D-90AF-FD67798F0C2F}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BrowserExtensionList_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BrowserExtensionList_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BrowserExtensionList_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL value AS Afx_ICoreWebView2BrowserExtension PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BrowserExtensionRemoveCompletedHandler
' IID: {8E41909A-9B18-4BB1-8CDF-930F467A50BE}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BrowserExtensionRemoveCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BrowserExtensionRemoveCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BrowserExtensionRemoveCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BrowserProcessExitedEventArgs
' IID: {1F00663F-AF8C-4782-9CDD-DD01C52E34CB}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BrowserProcessExitedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BrowserProcessExitedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BrowserProcessExitedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_BrowserProcessExitKind (BYVAL value AS COREWEBVIEW2_BROWSER_PROCESS_EXIT_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BrowserProcessId (BYVAL value AS UINT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BrowserProcessExitedEventHandler
' IID: {FA504257-A216-4911-A860-FE8825712861}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BrowserProcessExitedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BrowserProcessExitedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BrowserProcessExitedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Environment PTR, BYVAL args AS Afx_ICoreWebView2BrowserProcessExitedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2BytesReceivedChangedEventHandler
' IID: {828E8AB6-D94C-4264-9CEF-5217170D6251}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2BytesReceivedChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2BytesReceivedChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2BytesReceivedChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2DownloadOperation PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CallDevToolsProtocolMethodCompletedHandler
' IID: {5C4889F0-5EF6-4C5A-952C-D8F1B92D0574}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CapturePreviewCompletedHandler
' IID: {697E05E9-3D8F-45FA-96F4-8FFE1EDEDAF5}
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
' Interface name: ICoreWebView2Certificate
' IID: {C5FB2FCE-1CAC-4AEE-9C79-5ED0362EAAE0}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Certificate_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Certificate_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Certificate_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Subject (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Issuer (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ValidFrom (BYVAL value AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ValidTo (BYVAL value AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DerEncodedSerialNumber (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DisplayName (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ToPemEncoding (BYVAL pemEncodedData AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PemEncodedIssuerCertificateChain (BYVAL value AS Afx_ICoreWebView2StringCollection PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ClearBrowsingDataCompletedHandler
' IID: {E9710A06-1D1D-49B2-8234-226F35846AE5}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ClearBrowsingDataCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ClearBrowsingDataCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ClearBrowsingDataCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ClearServerCertificateErrorActionsCompletedHandler
' IID: {3B40AAC6-ACFE-4FFD-8211-F607B96E2D5B}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ClearServerCertificateErrorActionsCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ClearServerCertificateErrorActionsCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ClearServerCertificateErrorActionsCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ClientCertificate
' IID: {E7188076-BCC3-11EB-8529-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ClientCertificate_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ClientCertificate_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ClientCertificate_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Subject (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Issuer (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ValidFrom (BYVAL value AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ValidTo (BYVAL value AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DerEncodedSerialNumber (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DisplayName (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ToPemEncoding (BYVAL pemEncodedData AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PemEncodedIssuerCertificateChain (BYVAL value AS Afx_ICoreWebView2StringCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Kind (BYVAL value AS COREWEBVIEW2_CLIENT_CERTIFICATE_KIND PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ClientCertificateCollection
' IID: {EF5674D2-BCC3-11EB-8529-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ClientCertificateCollection_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ClientCertificateCollection_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ClientCertificateCollection_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL value AS Afx_ICoreWebView2ClientCertificate PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ClientCertificateRequestedEventArgs
' IID: {BC59DB28-BCC3-11EB-8529-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ClientCertificateRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ClientCertificateRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ClientCertificateRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Host (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Port (BYVAL value AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsProxy (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AllowedCertificateAuthorities (BYVAL value AS Afx_ICoreWebView2StringCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MutuallyTrustedCertificates (BYVAL value AS Afx_ICoreWebView2ClientCertificateCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SelectedCertificate (BYVAL value AS Afx_ICoreWebView2ClientCertificate PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SelectedCertificate (BYVAL value AS Afx_ICoreWebView2ClientCertificate PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Cancel (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Cancel (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ClientCertificateRequestedEventHandler
' IID: {D7175BA2-BCC3-11EB-8529-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ClientCertificateRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ClientCertificateRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ClientCertificateRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ClientCertificateRequestedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CompositionController
' IID: {3DF9B733-B9AE-4A15-86B4-EB9EE9826469}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CompositionController_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CompositionController_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CompositionController_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_RootVisualTarget (BYVAL target AS IUnknown PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_RootVisualTarget (BYVAL target AS IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SendMouseInput (BYVAL eventKind AS COREWEBVIEW2_MOUSE_EVENT_KIND, BYVAL virtualKeys AS COREWEBVIEW2_MOUSE_EVENT_VIRTUAL_KEYS, BYVAL mouseData AS UINT, BYVAL point AS POINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION SendPointerInput (BYVAL eventKind AS COREWEBVIEW2_POINTER_EVENT_KIND, BYVAL pointerInfo AS Afx_ICoreWebView2PointerInfo PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Cursor (BYVAL Cursor AS wireHICON PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SystemCursorId (BYVAL SystemCursorId AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_CursorChanged (BYVAL eventHandler AS Afx_ICoreWebView2CursorChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_CursorChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CompositionController2
' IID: {0B6A3D24-49CB-4806-BA20-B5E0734A7B26}
' Inherited interface = ICoreWebView2CompositionController
' ########################################################################################

#ifndef __Afx_ICoreWebView2CompositionController2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CompositionController2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CompositionController2_ EXTENDS Afx_ICoreWebView2CompositionController
   DECLARE ABSTRACT FUNCTION get_AutomationProvider (BYVAL value AS IUnknown PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CompositionController3
' IID: {9570570E-4D76-4361-9EE1-F04D0DBDFB1E}
' Inherited interface = ICoreWebView2CompositionController2
' ########################################################################################

#ifndef __Afx_ICoreWebView2CompositionController3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CompositionController3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CompositionController3_ EXTENDS Afx_ICoreWebView2CompositionController2
   DECLARE ABSTRACT FUNCTION DragEnter (BYVAL dataObject AS IDataObject PTR, BYVAL keyState AS ULONG, BYVAL point AS POINT, BYVAL effect AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DragLeave () AS HRESULT
   DECLARE ABSTRACT FUNCTION DragOver (BYVAL keyState AS ULONG, BYVAL point AS POINT, BYVAL effect AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Drop (BYVAL dataObject AS IDataObject PTR, BYVAL keyState AS ULONG, BYVAL point AS POINT, BYVAL effect AS ULONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CompositionController4
' IID: {7C367B9B-3D2B-450F-9E58-D61A20F486AA}
' Inherited interface = ICoreWebView2CompositionController3
' ########################################################################################

#ifndef __Afx_ICoreWebView2CompositionController4_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CompositionController4_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CompositionController4_ EXTENDS Afx_ICoreWebView2CompositionController3
   DECLARE ABSTRACT FUNCTION GetNonClientRegionAtPoint (BYVAL point AS POINT, BYVAL value AS COREWEBVIEW2_NON_CLIENT_REGION_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION QueryNonClientRegion (BYVAL Kind AS COREWEBVIEW2_NON_CLIENT_REGION_KIND, BYVAL rects AS Afx_ICoreWebView2RegionRectCollectionView PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_NonClientRegionChanged (BYVAL eventHandler AS Afx_ICoreWebView2NonClientRegionChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NonClientRegionChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ContainsFullScreenElementChangedEventHandler
' IID: {E45D98B1-AFEF-45BE-8BAF-6C7728867F73}
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
' Interface name: ICoreWebView2ContentLoadingEventArgs
' IID: {0C8A1275-9B6B-4901-87AD-70DF25BAFA6E}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ContentLoadingEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ContentLoadingEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ContentLoadingEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsErrorPage (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NavigationId (BYVAL value AS ULONGINT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ContentLoadingEventHandler
' IID: {364471E7-F2BE-4910-BDBA-D72077D51C4B}
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
' Interface name: ICoreWebView2ContextMenuItem
' IID: {7AED49E3-A93F-497A-811C-749C6B6B6C65}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ContextMenuItem_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ContextMenuItem_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ContextMenuItem_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_name (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Label (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CommandId (BYVAL value AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShortcutKeyDescription (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Icon (BYVAL value AS IStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Kind (BYVAL value AS COREWEBVIEW2_CONTEXT_MENU_ITEM_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsEnabled (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsChecked (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsChecked (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Children (BYVAL value AS Afx_ICoreWebView2ContextMenuItemCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_CustomItemSelected (BYVAL eventHandler AS Afx_ICoreWebView2CustomItemSelectedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_CustomItemSelected (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ContextMenuItemCollection
' IID: {F562A2F5-C415-45CF-B909-D4B7C1E276D3}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ContextMenuItemCollection_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ContextMenuItemCollection_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ContextMenuItemCollection_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL value AS Afx_ICoreWebView2ContextMenuItem PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveValueAtIndex (BYVAL index AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION InsertValueAtIndex (BYVAL index AS UINT, BYVAL value AS Afx_ICoreWebView2ContextMenuItem PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ContextMenuRequestedEventArgs
' IID: {A1D309EE-C03F-11EB-8529-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ContextMenuRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ContextMenuRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ContextMenuRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_MenuItems (BYVAL value AS Afx_ICoreWebView2ContextMenuItemCollection PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ContextMenuTarget (BYVAL value AS Afx_ICoreWebView2ContextMenuTarget PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Location (BYVAL value AS POINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SelectedCommandId (BYVAL value AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SelectedCommandId (BYVAL value AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ContextMenuRequestedEventHandler
' IID: {04D3FE1D-AB87-42FB-A898-DA241D35B63C}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ContextMenuRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ContextMenuRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ContextMenuRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ContextMenuRequestedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ContextMenuTarget
' IID: {B8611D99-EED6-4F3F-902C-A198502AD472}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ContextMenuTarget_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ContextMenuTarget_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ContextMenuTarget_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Kind (BYVAL value AS COREWEBVIEW2_CONTEXT_MENU_TARGET_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsEditable (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsRequestedForMainFrame (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PageUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FrameUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HasLinkUri (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_LinkUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HasLinkText (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_LinkText (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HasSourceUri (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SourceUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HasSelection (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SelectionText (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Controller
' IID: {4D00C0D1-9434-4EB6-8078-8697A560334F}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Controller_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Controller_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Controller_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsVisible (BYVAL IsVisible AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsVisible (BYVAL IsVisible AS LONG) AS HRESULT
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
   DECLARE ABSTRACT FUNCTION get_ParentWindow (BYVAL ParentWindow AS HWND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ParentWindow (BYVAL ParentWindow AS HWND) AS HRESULT
   DECLARE ABSTRACT FUNCTION NotifyParentWindowPositionChanged () AS HRESULT
   DECLARE ABSTRACT FUNCTION Close () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CoreWebView2 (BYVAL CoreWebView2 AS Afx_ICoreWebView2 PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Controller2
' IID: {C979903E-D4CA-4228-92EB-47EE3FA96EAB}
' Inherited interface = ICoreWebView2Controller
' ########################################################################################

#ifndef __Afx_ICoreWebView2Controller2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Controller2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Controller2_ EXTENDS Afx_ICoreWebView2Controller
   DECLARE ABSTRACT FUNCTION get_DefaultBackgroundColor (BYVAL value AS COREWEBVIEW2_COLOR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DefaultBackgroundColor (BYVAL value AS COREWEBVIEW2_COLOR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Controller3
' IID: {F9614724-5D2B-41DC-AEF7-73D62B51543B}
' Inherited interface = ICoreWebView2Controller2
' ########################################################################################

#ifndef __Afx_ICoreWebView2Controller3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Controller3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Controller3_ EXTENDS Afx_ICoreWebView2Controller2
   DECLARE ABSTRACT FUNCTION get_RasterizationScale (BYVAL scale AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_RasterizationScale (BYVAL scale AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldDetectMonitorScaleChanges (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ShouldDetectMonitorScaleChanges (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_RasterizationScaleChanged (BYVAL eventHandler AS Afx_ICoreWebView2RasterizationScaleChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_RasterizationScaleChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BoundsMode (BYVAL BoundsMode AS COREWEBVIEW2_BOUNDS_MODE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_BoundsMode (BYVAL BoundsMode AS COREWEBVIEW2_BOUNDS_MODE) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Controller4
' IID: {97D418D5-A426-4E49-A151-E1A10F327D9E}
' Inherited interface = ICoreWebView2Controller3
' ########################################################################################

#ifndef __Afx_ICoreWebView2Controller4_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Controller4_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Controller4_ EXTENDS Afx_ICoreWebView2Controller3
   DECLARE ABSTRACT FUNCTION get_AllowExternalDrop (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AllowExternalDrop (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ControllerOptions
' IID: {12AAE616-8CCB-44EC-BCB3-EB1831881635}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ControllerOptions_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ControllerOptions_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ControllerOptions_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ProfileName (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ProfileName (BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsInPrivateModeEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsInPrivateModeEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ControllerOptions2
' IID: {06C991D8-9E7E-11ED-A8FC-0242AC120002}
' Inherited interface = ICoreWebView2ControllerOptions
' ########################################################################################

#ifndef __Afx_ICoreWebView2ControllerOptions2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ControllerOptions2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ControllerOptions2_ EXTENDS Afx_ICoreWebView2ControllerOptions
   DECLARE ABSTRACT FUNCTION get_ScriptLocale (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ScriptLocale (BYVAL value AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ControllerOptions3
' IID: {B32B191A-8998-57CA-B7CB-E04617E4CE4A}
' Inherited interface = ICoreWebView2ControllerOptions2
' ########################################################################################

#ifndef __Afx_ICoreWebView2ControllerOptions3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ControllerOptions3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ControllerOptions3_ EXTENDS Afx_ICoreWebView2ControllerOptions2
   DECLARE ABSTRACT FUNCTION get_DefaultBackgroundColor (BYVAL value AS COREWEBVIEW2_COLOR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DefaultBackgroundColor (BYVAL value AS COREWEBVIEW2_COLOR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ControllerOptions4
' IID: {21EB052F-AD39-555E-824A-C87B091D4D36}
' Inherited interface = ICoreWebView2ControllerOptions3
' ########################################################################################

#ifndef __Afx_ICoreWebView2ControllerOptions4_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ControllerOptions4_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ControllerOptions4_ EXTENDS Afx_ICoreWebView2ControllerOptions3
   DECLARE ABSTRACT FUNCTION get_AllowHostInputProcessing (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AllowHostInputProcessing (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Cookie
' IID: {AD26D6BE-1486-43E6-BF87-A2034006CA21}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Cookie_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Cookie_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Cookie_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_name (BYVAL name AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_value (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_value (BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Domain (BYVAL Domain AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Path (BYVAL Path AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Expires (BYVAL Expires AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Expires (BYVAL Expires AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsHttpOnly (BYVAL IsHttpOnly AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsHttpOnly (BYVAL IsHttpOnly AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SameSite (BYVAL SameSite AS COREWEBVIEW2_COOKIE_SAME_SITE_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SameSite (BYVAL SameSite AS COREWEBVIEW2_COOKIE_SAME_SITE_KIND) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsSecure (BYVAL IsSecure AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsSecure (BYVAL IsSecure AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsSession (BYVAL IsSession AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CookieList
' IID: {F7F6F714-5D2A-43C6-9503-346ECE02D186}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CookieList_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CookieList_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CookieList_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL value AS Afx_ICoreWebView2Cookie PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CookieManager
' IID: {177CD9E7-B6F5-451A-94A0-5D7A3A4C4141}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CookieManager_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CookieManager_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CookieManager_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION CreateCookie (BYVAL name AS WSTRING PTR, BYVAL value AS WSTRING PTR, BYVAL Domain AS WSTRING PTR, BYVAL Path AS WSTRING PTR, BYVAL cookie AS Afx_ICoreWebView2Cookie PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CopyCookie (BYVAL cookieParam AS Afx_ICoreWebView2Cookie PTR, BYVAL cookie AS Afx_ICoreWebView2Cookie PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetCookies (BYVAL uri AS WSTRING PTR, BYVAL handler AS Afx_ICoreWebView2GetCookiesCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddOrUpdateCookie (BYVAL cookie AS Afx_ICoreWebView2Cookie PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteCookie (BYVAL cookie AS Afx_ICoreWebView2Cookie PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteCookies (BYVAL name AS WSTRING PTR, BYVAL uri AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteCookiesWithDomainAndPath (BYVAL name AS WSTRING PTR, BYVAL Domain AS WSTRING PTR, BYVAL Path AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION DeleteAllCookies () AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler
' IID: {02FAB84B-1428-4FB7-AD45-1B2E64736184}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS Afx_ICoreWebView2CompositionController PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
' IID: {6C4819F3-C9B7-4260-8127-C9F5BDE7F68C}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS Afx_ICoreWebView2Controller PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
' IID: {4E8A3389-C9D8-4BD2-B6B5-124FEE6CC14D}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS Afx_ICoreWebView2Environment PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CursorChangedEventHandler
' IID: {9DA43CCC-26E1-4DAD-B56C-D8961C94C571}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CursorChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CursorChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CursorChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2CompositionController PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CustomItemSelectedEventHandler
' IID: {49E1D0BC-FE9E-4481-B7C2-32324AA21998}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CustomItemSelectedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CustomItemSelectedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CustomItemSelectedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2ContextMenuItem PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2CustomSchemeRegistration
' IID: {D60AC92C-37A6-4B26-A39E-95CFE59047BB}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2CustomSchemeRegistration_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2CustomSchemeRegistration_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2CustomSchemeRegistration_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_SchemeName (BYVAL SchemeName AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TreatAsSecure (BYVAL TreatAsSecure AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TreatAsSecure (BYVAL TreatAsSecure AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAllowedOrigins (BYVAL allowedOriginsCount AS UINT PTR, BYVAL allowedOrigins AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetAllowedOrigins (BYVAL allowedOriginsCount AS UINT, BYVAL allowedOrigins AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HasAuthorityComponent (BYVAL HasAuthorityComponent AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_HasAuthorityComponent (BYVAL HasAuthorityComponent AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Deferral
' IID: {C10E7F7B-B585-46F0-A623-8BEFBF3E4EE0}
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
' IID: {653C2959-BB3A-4377-8632-B58ADA4E66C4}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ParameterObjectAsJson (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2DevToolsProtocolEventReceivedEventArgs2
' IID: {2DC4959D-1494-4393-95BA-BEA4CB9EBD1B}
' Inherited interface = ICoreWebView2DevToolsProtocolEventReceivedEventArgs
' ########################################################################################

#ifndef __Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs2_ EXTENDS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs
   DECLARE ABSTRACT FUNCTION get_sessionId (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2DevToolsProtocolEventReceivedEventHandler
' IID: {E2FDA4BE-5456-406C-A261-3D452138362C}
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
' IID: {B32CA51A-8371-45E9-9317-AF021D080367}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DevToolsProtocolEventReceiver_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DevToolsProtocolEventReceiver_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DevToolsProtocolEventReceiver_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION add_DevToolsProtocolEventReceived (BYVAL eventHandler AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_DevToolsProtocolEventReceived (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2DocumentTitleChangedEventHandler
' IID: {F5F2B923-953E-4042-9F95-F3A118E1AFD4}
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
' Interface name: ICoreWebView2DOMContentLoadedEventArgs
' IID: {16B1E21A-C503-44F2-84C9-70ABA5031283}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DOMContentLoadedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DOMContentLoadedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DOMContentLoadedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_NavigationId (BYVAL value AS ULONGINT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2DOMContentLoadedEventHandler
' IID: {4BAC7E9C-199E-49ED-87ED-249303ACF019}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DOMContentLoadedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DOMContentLoadedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DOMContentLoadedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2DOMContentLoadedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2DownloadOperation
' IID: {3D6B6CF2-AFE1-44C7-A995-C65117714336}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DownloadOperation_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DownloadOperation_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DownloadOperation_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION add_BytesReceivedChanged (BYVAL eventHandler AS Afx_ICoreWebView2BytesReceivedChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_BytesReceivedChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_EstimatedEndTimeChanged (BYVAL eventHandler AS Afx_ICoreWebView2EstimatedEndTimeChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_EstimatedEndTimeChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_StateChanged (BYVAL eventHandler AS Afx_ICoreWebView2StateChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_StateChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ContentDisposition (BYVAL ContentDisposition AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MimeType (BYVAL MimeType AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TotalBytesToReceive (BYVAL TotalBytesToReceive AS LONGINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BytesReceived (BYVAL BytesReceived AS LONGINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_EstimatedEndTime (BYVAL EstimatedEndTime AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResultFilePath (BYVAL ResultFilePath AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_State (BYVAL downloadState AS COREWEBVIEW2_DOWNLOAD_STATE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_InterruptReason (BYVAL InterruptReason AS COREWEBVIEW2_DOWNLOAD_INTERRUPT_REASON PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Cancel () AS HRESULT
   DECLARE ABSTRACT FUNCTION Pause () AS HRESULT
   DECLARE ABSTRACT FUNCTION Resume () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_CanResume (BYVAL CanResume AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: ICoreWebView2DownloadStartingEventArgs
' IID: {E99BBE21-43E9-4544-A732-282764EAFA60}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DownloadStartingEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DownloadStartingEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DownloadStartingEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_DownloadOperation (BYVAL DownloadOperation AS Afx_ICoreWebView2DownloadOperation PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Cancel (BYVAL Cancel AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Cancel (BYVAL Cancel AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResultFilePath (BYVAL ResultFilePath AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ResultFilePath (BYVAL ResultFilePath AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL Handled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL Handled AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2DownloadStartingEventHandler
' IID: {EFEDC989-C396-41CA-83F7-07F845A55724}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2DownloadStartingEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2DownloadStartingEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2DownloadStartingEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2DownloadStartingEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment
' IID: {B96D755E-0319-4E92-A296-23436F46A1FC}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION CreateCoreWebView2Controller (BYVAL ParentWindow AS HWND, BYVAL handler AS Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateWebResourceResponse (BYVAL Content AS IStream PTR, BYVAL StatusCode AS INT_, BYVAL ReasonPhrase AS WSTRING PTR, BYVAL Headers AS WSTRING PTR, BYVAL Response AS Afx_ICoreWebView2WebResourceResponse PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BrowserVersionString (BYVAL versionInfo AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_NewBrowserVersionAvailable (BYVAL eventHandler AS Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NewBrowserVersionAvailable (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment2
' IID: {41F3632B-5EF4-404F-AD82-2D606C5A9A21}
' Inherited interface = ICoreWebView2Environment
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment2_ EXTENDS Afx_ICoreWebView2Environment
   DECLARE ABSTRACT FUNCTION CreateWebResourceRequest (BYVAL uri AS WSTRING PTR, BYVAL Method AS WSTRING PTR, BYVAL postData AS IStream PTR, BYVAL Headers AS WSTRING PTR, BYVAL value AS Afx_ICoreWebView2WebResourceRequest PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment3
' IID: {80A22AE3-BE7C-4CE2-AFE1-5A50056CDEEB}
' Inherited interface = ICoreWebView2Environment2
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment3_ EXTENDS Afx_ICoreWebView2Environment2
   DECLARE ABSTRACT FUNCTION CreateCoreWebView2CompositionController (BYVAL ParentWindow AS wireHWND, BYVAL handler AS Afx_ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateCoreWebView2PointerInfo (BYVAL value AS Afx_ICoreWebView2PointerInfo PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment4
' IID: {20944379-6DCF-41D6-A0A0-ABC0FC50DE0D}
' Inherited interface = ICoreWebView2Environment3
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment4_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment4_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment4_ EXTENDS Afx_ICoreWebView2Environment3
   DECLARE ABSTRACT FUNCTION GetAutomationProviderForWindow (BYVAL hwnd AS wireHWND, BYVAL value AS IUnknown PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment5
' IID: {319E423D-E0D7-4B8D-9254-AE9475DE9B17}
' Inherited interface = ICoreWebView2Environment4
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment5_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment5_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment5_ EXTENDS Afx_ICoreWebView2Environment4
   DECLARE ABSTRACT FUNCTION add_BrowserProcessExited (BYVAL eventHandler AS Afx_ICoreWebView2BrowserProcessExitedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_BrowserProcessExited (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment6
' IID: {E59EE362-ACBD-4857-9A8E-D3644D9459A9}
' Inherited interface = ICoreWebView2Environment5
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment6_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment6_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment6_ EXTENDS Afx_ICoreWebView2Environment5
   DECLARE ABSTRACT FUNCTION CreatePrintSettings (BYVAL value AS Afx_ICoreWebView2PrintSettings PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment7
' IID: {43C22296-3BBD-43A4-9C00-5C0DF6DD29A2}
' Inherited interface = ICoreWebView2Environment6
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment7_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment7_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment7_ EXTENDS Afx_ICoreWebView2Environment6
   DECLARE ABSTRACT FUNCTION get_UserDataFolder (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment8
' IID: {D6EB91DD-C3D2-45E5-BD29-6DC2BC4DE9CF}
' Inherited interface = ICoreWebView2Environment7
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment8_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment8_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment8_ EXTENDS Afx_ICoreWebView2Environment7
   DECLARE ABSTRACT FUNCTION add_ProcessInfosChanged (BYVAL eventHandler AS Afx_ICoreWebView2ProcessInfosChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ProcessInfosChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetProcessInfos (BYVAL value AS Afx_ICoreWebView2ProcessInfoCollection PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment9
' IID: {F06F41BF-4B5A-49D8-B9F6-FA16CD29F274}
' Inherited interface = ICoreWebView2Environment8
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment9_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment9_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment9_ EXTENDS Afx_ICoreWebView2Environment8
   DECLARE ABSTRACT FUNCTION CreateContextMenuItem (BYVAL Label AS WSTRING PTR, BYVAL iconStream AS IStream PTR, BYVAL Kind AS COREWEBVIEW2_CONTEXT_MENU_ITEM_KIND, BYVAL value AS Afx_ICoreWebView2ContextMenuItem PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment10
' IID: {EE0EB9DF-6F12-46CE-B53F-3F47B9C928E0}
' Inherited interface = ICoreWebView2Environment9
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment10_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment10_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment10_ EXTENDS Afx_ICoreWebView2Environment9
   DECLARE ABSTRACT FUNCTION CreateCoreWebView2ControllerOptions (BYVAL value AS Afx_ICoreWebView2ControllerOptions PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateCoreWebView2ControllerWithOptions (BYVAL ParentWindow AS wireHWND, BYVAL options AS Afx_ICoreWebView2ControllerOptions PTR, BYVAL handler AS Afx_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateCoreWebView2CompositionControllerWithOptions (BYVAL ParentWindow AS wireHWND, BYVAL options AS Afx_ICoreWebView2ControllerOptions PTR, BYVAL handler AS Afx_ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment11
' IID: {F0913DC6-A0EC-42EF-9805-91DFF3A2966A}
' Inherited interface = ICoreWebView2Environment10
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment11_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment11_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment11_ EXTENDS Afx_ICoreWebView2Environment10
   DECLARE ABSTRACT FUNCTION get_FailureReportFolderPath (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment12
' IID: {F503DB9B-739F-48DD-B151-FDFCF253F54E}
' Inherited interface = ICoreWebView2Environment11
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment12_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment12_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment12_ EXTENDS Afx_ICoreWebView2Environment11
   DECLARE ABSTRACT FUNCTION CreateSharedBuffer (BYVAL Size AS ULONGINT, BYVAL value AS Afx_ICoreWebView2SharedBuffer PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment13
' IID: {AF641F58-72B2-11EE-B962-0242AC120002}
' Inherited interface = ICoreWebView2Environment12
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment13_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment13_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment13_ EXTENDS Afx_ICoreWebView2Environment12
   DECLARE ABSTRACT FUNCTION GetProcessExtendedInfos (BYVAL handler AS Afx_ICoreWebView2GetProcessExtendedInfosCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment14
' IID: {A5E9FAD9-C875-59DA-9BD7-473AA5CA1CEF}
' Inherited interface = ICoreWebView2Environment13
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment14_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment14_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment14_ EXTENDS Afx_ICoreWebView2Environment13
   DECLARE ABSTRACT FUNCTION CreateWebFileSystemFileHandle (BYVAL Path AS WSTRING PTR, BYVAL Permission AS COREWEBVIEW2_FILE_SYSTEM_HANDLE_PERMISSION, BYVAL value AS Afx_ICoreWebView2FileSystemHandle PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateWebFileSystemDirectoryHandle (BYVAL Path AS WSTRING PTR, BYVAL Permission AS COREWEBVIEW2_FILE_SYSTEM_HANDLE_PERMISSION, BYVAL value AS Afx_ICoreWebView2FileSystemHandle PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateObjectCollection (BYVAL length AS UINT, BYVAL items AS IUnknown PTR, BYVAL objectCollection AS Afx_ICoreWebView2ObjectCollection PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Environment15
' IID: {2AC5EBFB-E654-5961-A667-7971885C7B27}
' Inherited interface = ICoreWebView2Environment14
' ########################################################################################

#ifndef __Afx_ICoreWebView2Environment15_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Environment15_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Environment15_ EXTENDS Afx_ICoreWebView2Environment14
   DECLARE ABSTRACT FUNCTION CreateFindOptions (BYVAL value AS Afx_ICoreWebView2FindOptions PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2EnvironmentOptions
' IID: {2FDE08A8-1E9A-4766-8C05-95A9CEB9D1C5}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EnvironmentOptions_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EnvironmentOptions_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EnvironmentOptions_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_AdditionalBrowserArguments (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AdditionalBrowserArguments (BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Language (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Language (BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TargetCompatibleBrowserVersion (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TargetCompatibleBrowserVersion (BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AllowSingleSignOnUsingOSPrimaryAccount (BYVAL allow AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AllowSingleSignOnUsingOSPrimaryAccount (BYVAL allow AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2EnvironmentOptions2
' IID: {FF85C98A-1BA7-4A6B-90C8-2B752C89E9E2}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EnvironmentOptions2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EnvironmentOptions2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EnvironmentOptions2_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ExclusiveUserDataFolderAccess (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ExclusiveUserDataFolderAccess (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2EnvironmentOptions3
' IID: {4A5C436E-A9E3-4A2E-89C3-910D3513F5CC}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EnvironmentOptions3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EnvironmentOptions3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EnvironmentOptions3_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsCustomCrashReportingEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsCustomCrashReportingEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2EnvironmentOptions4
' IID: {AC52D13F-0D38-475A-9DCA-876580D6793E}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EnvironmentOptions4_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EnvironmentOptions4_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EnvironmentOptions4_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetCustomSchemeRegistrations (BYVAL Count AS UINT PTR, BYVAL schemeRegistrations AS Afx_ICoreWebView2CustomSchemeRegistration PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetCustomSchemeRegistrations (BYVAL Count AS UINT, BYVAL schemeRegistrations AS Afx_ICoreWebView2CustomSchemeRegistration PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2EnvironmentOptions5
' IID: {0AE35D64-C47F-4464-814E-259C345D1501}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EnvironmentOptions5_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EnvironmentOptions5_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EnvironmentOptions5_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_EnableTrackingPrevention (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_EnableTrackingPrevention (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2EnvironmentOptions6
' IID: {57D29CC3-C84F-42A0-B0E2-EFFBD5E179DE}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EnvironmentOptions6_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EnvironmentOptions6_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EnvironmentOptions6_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_AreBrowserExtensionsEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreBrowserExtensionsEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2EnvironmentOptions7
' IID: {C48D539F-E39F-441C-AE68-1F66E570BDC5}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EnvironmentOptions7_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EnvironmentOptions7_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EnvironmentOptions7_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ChannelSearchKind (BYVAL value AS COREWEBVIEW2_CHANNEL_SEARCH_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ChannelSearchKind (BYVAL value AS COREWEBVIEW2_CHANNEL_SEARCH_KIND) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ReleaseChannels (BYVAL value AS COREWEBVIEW2_RELEASE_CHANNELS PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ReleaseChannels (BYVAL value AS COREWEBVIEW2_RELEASE_CHANNELS) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2EnvironmentOptions8
' IID: {7C7ECF51-E918-5CAF-853C-E9A2BCC27775}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EnvironmentOptions8_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EnvironmentOptions8_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EnvironmentOptions8_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ScrollBarStyle (BYVAL value AS COREWEBVIEW2_SCROLLBAR_STYLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ScrollBarStyle (BYVAL value AS COREWEBVIEW2_SCROLLBAR_STYLE) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2EstimatedEndTimeChangedEventHandler
' IID: {28F0D425-93FE-4E63-9F8D-2AEEC6D3BA1E}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2EstimatedEndTimeChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2EstimatedEndTimeChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2EstimatedEndTimeChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2DownloadOperation PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ExecuteScriptCompletedHandler
' IID: {49511172-CC67-4BCA-9923-137112F4C4CC}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ExecuteScriptCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ExecuteScriptCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ExecuteScriptCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ExecuteScriptResult
' IID: {0CE15963-3698-4DF7-9399-71ED6CDD8C9F}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ExecuteScriptResult_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ExecuteScriptResult_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ExecuteScriptResult_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Succeeded (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResultAsJson (BYVAL jsonResult AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION TryGetResultAsString (BYVAL stringResult AS WSTRING PTR PTR, BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Exception (BYVAL Exception AS Afx_ICoreWebView2ScriptException PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ExecuteScriptWithResultCompletedHandler
' IID: {1BB5317B-8238-4C67-A7FF-BAF6558F289D}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ExecuteScriptWithResultCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ExecuteScriptWithResultCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ExecuteScriptWithResultCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS Afx_ICoreWebView2ExecuteScriptResult PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FaviconChangedEventHandler
' IID: {2913DA94-833D-4DE0-8DCA-900FC524A1A4}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FaviconChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FaviconChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FaviconChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2File
' IID: {F2C19559-6BC1-4583-A757-90021BE9AFEC}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2File_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2File_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2File_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Path (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FileSystemHandle
' IID: {C65100AC-0DE2-5551-A362-23D9BD1D0E1F}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FileSystemHandle_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FileSystemHandle_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FileSystemHandle_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Kind (BYVAL value AS COREWEBVIEW2_FILE_SYSTEM_HANDLE_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Path (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Permission (BYVAL value AS COREWEBVIEW2_FILE_SYSTEM_HANDLE_PERMISSION PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Find
' IID: {A3EC0F5F-DDBC-54ED-8546-AF75A785B9A6}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Find_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Find_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Find_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ActiveMatchIndex (BYVAL value AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MatchCount (BYVAL value AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ActiveMatchIndexChanged (BYVAL eventHandler AS Afx_ICoreWebView2FindActiveMatchIndexChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ActiveMatchIndexChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_MatchCountChanged (BYVAL eventHandler AS Afx_ICoreWebView2FindMatchCountChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_MatchCountChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION Start (BYVAL options AS Afx_ICoreWebView2FindOptions PTR, BYVAL handler AS Afx_ICoreWebView2FindStartCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION FindNext () AS HRESULT
   DECLARE ABSTRACT FUNCTION FindPrevious () AS HRESULT
   DECLARE ABSTRACT FUNCTION Stop () AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FindActiveMatchIndexChangedEventHandler
' IID: {0054F514-9A8E-5876-AED5-30B37F8C86A5}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FindActiveMatchIndexChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FindActiveMatchIndexChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FindActiveMatchIndexChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Find PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FindMatchCountChangedEventHandler
' IID: {DA0D6827-4254-5B10-A6D9-412076AFC9F3}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FindMatchCountChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FindMatchCountChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FindMatchCountChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Find PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FindOptions
' IID: {E82E3B2B-A4AF-5BC6-94C6-18B44157A16C}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FindOptions_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FindOptions_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FindOptions_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_FindTerm (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_FindTerm (BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsCaseSensitive (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsCaseSensitive (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldHighlightAllMatches (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ShouldHighlightAllMatches (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldMatchWord (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ShouldMatchWord (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SuppressDefaultFindDialog (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SuppressDefaultFindDialog (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FindStartCompletedHandler
' IID: {6A90ECAF-44B0-5BD9-8F07-1967E17BE9FB}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FindStartCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FindStartCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FindStartCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FocusChangedEventHandler
' IID: {05EA24BD-6452-4926-9014-4B82B498135D}
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
' Interface name: ICoreWebView2Frame
' IID: {F1131A5E-9BA9-11EB-A8B3-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Frame_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Frame_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Frame_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_name (BYVAL name AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_NameChanged (BYVAL eventHandler AS Afx_ICoreWebView2FrameNameChangedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NameChanged (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddHostObjectToScriptWithOrigins (BYVAL name AS WSTRING PTR, BYVAL object AS VARIANT PTR, BYVAL originsCount AS UINT, BYVAL origins AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveHostObjectFromScript (BYVAL name AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_Destroyed (BYVAL eventHandler AS Afx_ICoreWebView2FrameDestroyedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_Destroyed (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsDestroyed (BYVAL destroyed AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Frame2
' IID: {7A6A5834-D185-4DBF-B63F-4A9BC43107D4}
' Inherited interface = ICoreWebView2Frame
' ########################################################################################

#ifndef __Afx_ICoreWebView2Frame2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Frame2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Frame2_ EXTENDS Afx_ICoreWebView2Frame
   DECLARE ABSTRACT FUNCTION add_NavigationStarting (BYVAL eventHandler AS Afx_ICoreWebView2FrameNavigationStartingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NavigationStarting (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_ContentLoading (BYVAL eventHandler AS Afx_ICoreWebView2FrameContentLoadingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ContentLoading (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_NavigationCompleted (BYVAL eventHandler AS Afx_ICoreWebView2FrameNavigationCompletedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_NavigationCompleted (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_DOMContentLoaded (BYVAL eventHandler AS Afx_ICoreWebView2FrameDOMContentLoadedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_DOMContentLoaded (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION ExecuteScript (BYVAL javaScript AS WSTRING PTR, BYVAL handler AS Afx_ICoreWebView2ExecuteScriptCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION PostWebMessageAsJson (BYVAL webMessageAsJson AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION PostWebMessageAsString (BYVAL webMessageAsString AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION add_WebMessageReceived (BYVAL handler AS Afx_ICoreWebView2FrameWebMessageReceivedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_WebMessageReceived (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Frame3
' IID: {B50D82CC-CC28-481D-9614-CB048895E6A0}
' Inherited interface = ICoreWebView2Frame2
' ########################################################################################

#ifndef __Afx_ICoreWebView2Frame3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Frame3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Frame3_ EXTENDS Afx_ICoreWebView2Frame2
   DECLARE ABSTRACT FUNCTION add_PermissionRequested (BYVAL eventHandler AS Afx_ICoreWebView2FramePermissionRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_PermissionRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Frame4
' IID: {188782DC-92AA-4732-AB3C-FCC59F6F68B9}
' Inherited interface = ICoreWebView2Frame3
' ########################################################################################

#ifndef __Afx_ICoreWebView2Frame4_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Frame4_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Frame4_ EXTENDS Afx_ICoreWebView2Frame3
   DECLARE ABSTRACT FUNCTION PostSharedBufferToScript (BYVAL sharedBuffer AS Afx_ICoreWebView2SharedBuffer PTR, BYVAL access AS COREWEBVIEW2_SHARED_BUFFER_ACCESS, BYVAL additionalDataAsJson AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Frame5
' IID: {99D199C4-7305-11EE-B962-0242AC120002}
' Inherited interface = ICoreWebView2Frame4
' ########################################################################################

#ifndef __Afx_ICoreWebView2Frame5_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Frame5_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Frame5_ EXTENDS Afx_ICoreWebView2Frame4
   DECLARE ABSTRACT FUNCTION get_FrameId (BYVAL value AS UINT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Frame6
' IID: {0DE611FD-31E9-5DDC-9D71-95EDA26EFF32}
' Inherited interface = ICoreWebView2Frame5
' ########################################################################################

#ifndef __Afx_ICoreWebView2Frame6_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Frame6_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Frame6_ EXTENDS Afx_ICoreWebView2Frame5
   DECLARE ABSTRACT FUNCTION add_ScreenCaptureStarting (BYVAL eventHandler AS Afx_ICoreWebView2FrameScreenCaptureStartingEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_ScreenCaptureStarting (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Frame7
' IID: {3598CFA2-D85D-5A9F-9228-4DDE1F59EC64}
' Inherited interface = ICoreWebView2Frame6
' ########################################################################################

#ifndef __Afx_ICoreWebView2Frame7_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Frame7_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Frame7_ EXTENDS Afx_ICoreWebView2Frame6
   DECLARE ABSTRACT FUNCTION add_FrameCreated (BYVAL eventHandler AS Afx_ICoreWebView2FrameChildFrameCreatedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_FrameCreated (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameChildFrameCreatedEventHandler
' IID: {569E40E7-46B7-563D-83AE-1073155664D7}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameChildFrameCreatedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameChildFrameCreatedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameChildFrameCreatedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS Afx_ICoreWebView2FrameCreatedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameContentLoadingEventHandler
' IID: {0D6156F2-D332-49A7-9E03-7D8F2FEEEE54}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameContentLoadingEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameContentLoadingEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameContentLoadingEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS Afx_ICoreWebView2ContentLoadingEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameCreatedEventArgs
' IID: {4D6E7B5E-9BAA-11EB-A8B3-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameCreatedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameCreatedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameCreatedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Frame (BYVAL value AS Afx_ICoreWebView2Frame PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameCreatedEventHandler
' IID: {38059770-9BAA-11EB-A8B3-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameCreatedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameCreatedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameCreatedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2FrameCreatedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameDestroyedEventHandler
' IID: {59DD7B4C-9BAA-11EB-A8B3-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameDestroyedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameDestroyedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameDestroyedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameDOMContentLoadedEventHandler
' IID: {38D9520D-340F-4D1E-A775-43FCE9753683}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameDOMContentLoadedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameDOMContentLoadedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameDOMContentLoadedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS Afx_ICoreWebView2DOMContentLoadedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameInfo
' IID: {DA86B8A1-BDF3-4F11-9955-528CEFA59727}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameInfo_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameInfo_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameInfo_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_name (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Source (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameInfo2
' IID: {56F85CFA-72C4-11EE-B962-0242AC120002}
' Inherited interface = ICoreWebView2FrameInfo
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameInfo2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameInfo2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameInfo2_ EXTENDS Afx_ICoreWebView2FrameInfo
   DECLARE ABSTRACT FUNCTION get_ParentFrameInfo (BYVAL frameInfo AS Afx_ICoreWebView2FrameInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FrameId (BYVAL id AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FrameKind (BYVAL Kind AS COREWEBVIEW2_FRAME_KIND PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameInfoCollection
' IID: {8F834154-D38E-4D90-AFFB-6800A7272839}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameInfoCollection_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameInfoCollection_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameInfoCollection_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetIterator (BYVAL value AS Afx_ICoreWebView2FrameInfoCollectionIterator PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameInfoCollectionIterator
' IID: {1BF89E2D-1B2B-4629-B28F-05099B41BB03}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameInfoCollectionIterator_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameInfoCollectionIterator_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameInfoCollectionIterator_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_hasCurrent (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetCurrent (BYVAL value AS Afx_ICoreWebView2FrameInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveNext (BYVAL value AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameNameChangedEventHandler
' IID: {435C7DC8-9BAA-11EB-A8B3-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameNameChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameNameChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameNameChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameNavigationCompletedEventHandler
' IID: {609302AD-0E36-4F9A-A210-6A45272842A9}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameNavigationCompletedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameNavigationCompletedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameNavigationCompletedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS Afx_ICoreWebView2NavigationCompletedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameNavigationStartingEventHandler
' IID: {E79908BF-2D5D-4968-83DB-263FEA2C1DA3}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameNavigationStartingEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameNavigationStartingEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameNavigationStartingEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS Afx_ICoreWebView2NavigationStartingEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FramePermissionRequestedEventHandler
' IID: {845D0EDD-8BD8-429B-9915-4821789F23E9}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FramePermissionRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FramePermissionRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FramePermissionRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS Afx_ICoreWebView2PermissionRequestedEventArgs2 PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameScreenCaptureStartingEventHandler
' IID: {A6C1D8AD-BB80-59C5-895B-FBA1698B9309}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameScreenCaptureStartingEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameScreenCaptureStartingEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameScreenCaptureStartingEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS Afx_ICoreWebView2ScreenCaptureStartingEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2FrameWebMessageReceivedEventHandler
' IID: {E371E005-6D1D-4517-934B-A8F1629C62A5}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2FrameWebMessageReceivedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2FrameWebMessageReceivedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2FrameWebMessageReceivedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Frame PTR, BYVAL args AS Afx_ICoreWebView2WebMessageReceivedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2GetCookiesCompletedHandler
' IID: {5A4F5069-5C15-47C3-8646-F4DE1C116670}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2GetCookiesCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2GetCookiesCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2GetCookiesCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS Afx_ICoreWebView2CookieList PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2GetFaviconCompletedHandler
' IID: {A2508329-7DA8-49D7-8C05-FA125E4AEE8D}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2GetFaviconCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2GetFaviconCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2GetFaviconCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS IStream PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2GetNonDefaultPermissionSettingsCompletedHandler
' IID: {38274481-A15C-4563-94CF-990EDC9AEB95}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2GetNonDefaultPermissionSettingsCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2GetNonDefaultPermissionSettingsCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2GetNonDefaultPermissionSettingsCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS Afx_ICoreWebView2PermissionSettingCollectionView PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2GetProcessExtendedInfosCompletedHandler
' IID: {F45E55AA-3BC2-11EE-BE56-0242AC120002}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2GetProcessExtendedInfosCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2GetProcessExtendedInfosCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2GetProcessExtendedInfosCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS Afx_ICoreWebView2ProcessExtendedInfoCollection PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2HistoryChangedEventHandler
' IID: {C79A420C-EFD9-4058-9295-3E8B4BCAB645}
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
' IID: {0702FC30-F43B-47BB-AB52-A42CB552AD9F}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2HttpHeadersCollectionIterator_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2HttpHeadersCollectionIterator_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2HttpHeadersCollectionIterator_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetCurrentHeader (BYVAL name AS WSTRING PTR PTR, BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HasCurrentHeader (BYVAL hasCurrent AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveNext (BYVAL hasNext AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2HttpRequestHeaders
' IID: {E86CAC0E-5523-465C-B536-8FB9FC8C8C60}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2HttpRequestHeaders_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2HttpRequestHeaders_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2HttpRequestHeaders_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetHeader (BYVAL name AS WSTRING PTR, BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetHeaders (BYVAL name AS WSTRING PTR, BYVAL value AS Afx_ICoreWebView2HttpHeadersCollectionIterator PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Contains (BYVAL name AS WSTRING PTR, BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetHeader (BYVAL name AS WSTRING PTR, BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION RemoveHeader (BYVAL name AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIterator (BYVAL value AS Afx_ICoreWebView2HttpHeadersCollectionIterator PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2HttpResponseHeaders
' IID: {03C5FF5A-9B45-4A88-881C-89A9F328619C}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2HttpResponseHeaders_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2HttpResponseHeaders_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2HttpResponseHeaders_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION AppendHeader (BYVAL name AS WSTRING PTR, BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Contains (BYVAL name AS WSTRING PTR, BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetHeader (BYVAL name AS WSTRING PTR, BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetHeaders (BYVAL name AS WSTRING PTR, BYVAL value AS Afx_ICoreWebView2HttpHeadersCollectionIterator PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIterator (BYVAL value AS Afx_ICoreWebView2HttpHeadersCollectionIterator PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2IsDefaultDownloadDialogOpenChangedEventHandler
' IID: {3117DA26-AE13-438D-BD46-EDBEB2C4CE81}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2IsDefaultDownloadDialogOpenChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2IsDefaultDownloadDialogOpenChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2IsDefaultDownloadDialogOpenChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2IsDocumentPlayingAudioChangedEventHandler
' IID: {5DEF109A-2F4B-49FA-B7F6-11C39E513328}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2IsDocumentPlayingAudioChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2IsDocumentPlayingAudioChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2IsDocumentPlayingAudioChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2IsMutedChangedEventHandler
' IID: {57D90347-CD0E-4952-A4A2-7483A2756F08}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2IsMutedChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2IsMutedChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2IsMutedChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2LaunchingExternalUriSchemeEventArgs
' IID: {07D1A6C3-7175-4BA1-9306-E593CA07E46C}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2LaunchingExternalUriSchemeEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2LaunchingExternalUriSchemeEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2LaunchingExternalUriSchemeEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_InitiatingOrigin (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsUserInitiated (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Cancel (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Cancel (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL value AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2LaunchingExternalUriSchemeEventHandler
' IID: {74F712E0-8165-43A9-A13F-0CCE597E75DF}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2LaunchingExternalUriSchemeEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2LaunchingExternalUriSchemeEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2LaunchingExternalUriSchemeEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2LaunchingExternalUriSchemeEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2MoveFocusRequestedEventArgs
' IID: {2D6AA13B-3839-4A15-92FC-D88B3C0D9C9D}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2MoveFocusRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2MoveFocusRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2MoveFocusRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_reason (BYVAL reason AS COREWEBVIEW2_MOVE_FOCUS_REASON PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2MoveFocusRequestedEventHandler
' IID: {69035451-6DC7-4CB8-9BCE-B2BD70AD289F}
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
' IID: {30D68B7D-20D9-4752-A9CA-EC8448FBB5C1}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NavigationCompletedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NavigationCompletedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationCompletedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsSuccess (BYVAL IsSuccess AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_WebErrorStatus (BYVAL WebErrorStatus AS COREWEBVIEW2_WEB_ERROR_STATUS PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NavigationId (BYVAL NavigationId AS ULONGINT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NavigationCompletedEventArgs2
' IID: {FDF8B738-EE1E-4DB2-A329-8D7D7B74D792}
' Inherited interface = ICoreWebView2NavigationCompletedEventArgs
' ########################################################################################

#ifndef __Afx_ICoreWebView2NavigationCompletedEventArgs2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NavigationCompletedEventArgs2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationCompletedEventArgs2_ EXTENDS Afx_ICoreWebView2NavigationCompletedEventArgs
   DECLARE ABSTRACT FUNCTION get_HttpStatusCode (BYVAL value AS INT_ PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NavigationCompletedEventHandler
' IID: {D33A35BF-1C49-4F98-93AB-006E0533FE1C}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NavigationCompletedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NavigationCompletedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationCompletedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NavigationCompletedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NavigationStartingEventArgs
' IID: {5B495469-E119-438A-9B18-7604F25F2E49}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NavigationStartingEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NavigationStartingEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationStartingEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsUserInitiated (BYVAL IsUserInitiated AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsRedirected (BYVAL IsRedirected AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RequestHeaders (BYVAL RequestHeaders AS Afx_ICoreWebView2HttpRequestHeaders PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Cancel (BYVAL Cancel AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Cancel (BYVAL Cancel AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NavigationId (BYVAL NavigationId AS ULONGINT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NavigationStartingEventArgs2
' IID: {9086BE93-91AA-472D-A7E0-579F2BA006AD}
' Inherited interface = ICoreWebView2NavigationStartingEventArgs
' ########################################################################################

#ifndef __Afx_ICoreWebView2NavigationStartingEventArgs2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NavigationStartingEventArgs2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationStartingEventArgs2_ EXTENDS Afx_ICoreWebView2NavigationStartingEventArgs
   DECLARE ABSTRACT FUNCTION get_AdditionalAllowedFrameAncestors (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AdditionalAllowedFrameAncestors (BYVAL value AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NavigationStartingEventArgs3
' IID: {DDFFE494-4942-4BD2-AB73-35B8FF40E19F}
' Inherited interface = ICoreWebView2NavigationStartingEventArgs2
' ########################################################################################

#ifndef __Afx_ICoreWebView2NavigationStartingEventArgs3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NavigationStartingEventArgs3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NavigationStartingEventArgs3_ EXTENDS Afx_ICoreWebView2NavigationStartingEventArgs2
   DECLARE ABSTRACT FUNCTION get_NavigationKind (BYVAL value AS COREWEBVIEW2_NAVIGATION_KIND PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NavigationStartingEventHandler
' IID: {9ADBE429-F36D-432B-9DDC-F8881FBD76E3}
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
' IID: {F9A2976E-D34E-44FC-ADEE-81B6B57CA914}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NewBrowserVersionAvailableEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Environment PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NewWindowRequestedEventArgs
' IID: {34ACB11C-FC37-4418-9132-F9C21D1EAFB9}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NewWindowRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NewWindowRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NewWindowRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_NewWindow (BYVAL NewWindow AS Afx_ICoreWebView2 PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_NewWindow (BYVAL NewWindow AS Afx_ICoreWebView2 PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL Handled AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL Handled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsUserInitiated (BYVAL IsUserInitiated AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_WindowFeatures (BYVAL value AS Afx_ICoreWebView2WindowFeatures PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NewWindowRequestedEventArgs2
' IID: {BBC7BAED-74C6-4C92-B63A-7F5AEAE03DE3}
' Inherited interface = ICoreWebView2NewWindowRequestedEventArgs
' ########################################################################################

#ifndef __Afx_ICoreWebView2NewWindowRequestedEventArgs2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NewWindowRequestedEventArgs2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NewWindowRequestedEventArgs2_ EXTENDS Afx_ICoreWebView2NewWindowRequestedEventArgs
   DECLARE ABSTRACT FUNCTION get_name (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NewWindowRequestedEventArgs3
' IID: {842BED3C-6AD6-4DD9-B938-28C96667AD66}
' Inherited interface = ICoreWebView2NewWindowRequestedEventArgs2
' ########################################################################################

#ifndef __Afx_ICoreWebView2NewWindowRequestedEventArgs3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NewWindowRequestedEventArgs3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NewWindowRequestedEventArgs3_ EXTENDS Afx_ICoreWebView2NewWindowRequestedEventArgs2
   DECLARE ABSTRACT FUNCTION get_OriginalSourceFrameInfo (BYVAL value AS Afx_ICoreWebView2FrameInfo PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NewWindowRequestedEventHandler
' IID: {D4C185FE-C81C-4989-97AF-2D3FA7AB5651}
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
' Interface name: ICoreWebView2NonClientRegionChangedEventArgs
' IID: {AB71D500-0820-4A52-809C-48DB04FF93BF}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NonClientRegionChangedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NonClientRegionChangedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NonClientRegionChangedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_RegionKind (BYVAL value AS COREWEBVIEW2_NON_CLIENT_REGION_KIND PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NonClientRegionChangedEventHandler
' IID: {4A794E66-AA6C-46BD-93A3-382196837680}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NonClientRegionChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NonClientRegionChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NonClientRegionChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2CompositionController PTR, BYVAL args AS Afx_ICoreWebView2NonClientRegionChangedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Notification
' IID: {B7434D98-6BC8-419D-9DA5-FB5A96D4DACD}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Notification_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Notification_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Notification_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION add_CloseRequested (BYVAL eventHandler AS Afx_ICoreWebView2NotificationCloseRequestedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_CloseRequested (BYVAL token AS EventRegistrationToken) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReportShown () AS HRESULT
   DECLARE ABSTRACT FUNCTION ReportClicked () AS HRESULT
   DECLARE ABSTRACT FUNCTION ReportClosed () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Body (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Direction (BYVAL value AS COREWEBVIEW2_TEXT_DIRECTION_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Language (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Tag (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IconUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_title (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BadgeUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_BodyImageUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldRenotify (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RequiresInteraction (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsSilent (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Timestamp (BYVAL value AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetVibrationPattern (BYVAL Count AS UINT PTR, BYVAL vibrationPattern AS ULONGINT PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NotificationCloseRequestedEventHandler
' IID: {47C32D23-1E94-4733-85F1-D9BF4ACD0974}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NotificationCloseRequestedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NotificationCloseRequestedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NotificationCloseRequestedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Notification PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NotificationReceivedEventArgs
' IID: {1512DD5B-5514-4F85-886E-21C3A4C9CFE6}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NotificationReceivedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NotificationReceivedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NotificationReceivedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_SenderOrigin (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Notification (BYVAL value AS Afx_ICoreWebView2Notification PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2NotificationReceivedEventHandler
' IID: {89C5D598-8788-423B-BE97-E6E01C0F9EE3}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2NotificationReceivedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2NotificationReceivedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2NotificationReceivedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NotificationReceivedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ObjectCollectionView
' IID: {0F36FD87-4F69-4415-98DA-888F89FB9A33}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ObjectCollectionView_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ObjectCollectionView_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ObjectCollectionView_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL value AS IUnknown PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ObjectCollection
' IID: {5CFEC11C-25BD-4E8D-9E1A-7ACDAEEEC047}
' Inherited interface = ICoreWebView2ObjectCollectionView
' ########################################################################################

#ifndef __Afx_ICoreWebView2ObjectCollection_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ObjectCollection_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ObjectCollection_ EXTENDS Afx_ICoreWebView2ObjectCollectionView
   DECLARE ABSTRACT FUNCTION RemoveValueAtIndex (BYVAL index AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION InsertValueAtIndex (BYVAL index AS UINT, BYVAL value AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PermissionRequestedEventArgs
' IID: {973AE2EF-FF18-4894-8FB2-3C758F046810}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PermissionRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PermissionRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PermissionRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PermissionKind (BYVAL PermissionKind AS COREWEBVIEW2_PERMISSION_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsUserInitiated (BYVAL IsUserInitiated AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_State (BYVAL State AS COREWEBVIEW2_PERMISSION_STATE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_State (BYVAL State AS COREWEBVIEW2_PERMISSION_STATE) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PermissionRequestedEventArgs2
' IID: {74D7127F-9DE6-4200-8734-42D6FB4FF741}
' Inherited interface = ICoreWebView2PermissionRequestedEventArgs
' ########################################################################################

#ifndef __Afx_ICoreWebView2PermissionRequestedEventArgs2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PermissionRequestedEventArgs2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PermissionRequestedEventArgs2_ EXTENDS Afx_ICoreWebView2PermissionRequestedEventArgs
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PermissionRequestedEventArgs3
' IID: {E61670BC-3DCE-4177-86D2-C629AE3CB6AC}
' Inherited interface = ICoreWebView2PermissionRequestedEventArgs2
' ########################################################################################

#ifndef __Afx_ICoreWebView2PermissionRequestedEventArgs3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PermissionRequestedEventArgs3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PermissionRequestedEventArgs3_ EXTENDS Afx_ICoreWebView2PermissionRequestedEventArgs2
   DECLARE ABSTRACT FUNCTION get_SavesInProfile (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SavesInProfile (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PermissionRequestedEventHandler
' IID: {15E1C6A3-C72A-4DF3-91D7-D097FBEC6BFD}
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
' Interface name: ICoreWebView2PermissionSetting
' IID: {792B6ECA-5576-421C-9119-74EBB3A4FFB3}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PermissionSetting_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PermissionSetting_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PermissionSetting_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_PermissionKind (BYVAL value AS COREWEBVIEW2_PERMISSION_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PermissionOrigin (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PermissionState (BYVAL value AS COREWEBVIEW2_PERMISSION_STATE PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PermissionSettingCollectionView
' IID: {F5596F62-3DE5-47B1-91E8-A4104B596B96}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PermissionSettingCollectionView_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PermissionSettingCollectionView_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PermissionSettingCollectionView_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL permissionSetting AS Afx_ICoreWebView2PermissionSetting PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PointerInfo
' IID: {E6995887-D10D-4F5D-9359-4CE46E4F96B9}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PointerInfo_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PointerInfo_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PointerInfo_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_PointerKind (BYVAL PointerKind AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PointerKind (BYVAL PointerKind AS ULONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PointerId (BYVAL PointerId AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PointerId (BYVAL PointerId AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FrameId (BYVAL FrameId AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_FrameId (BYVAL FrameId AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PointerFlags (BYVAL PointerFlags AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PointerFlags (BYVAL PointerFlags AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PointerDeviceRect (BYVAL PointerDeviceRect AS RECT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PointerDeviceRect (BYVAL PointerDeviceRect AS RECT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DisplayRect (BYVAL DisplayRect AS RECT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DisplayRect (BYVAL DisplayRect AS RECT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PixelLocation (BYVAL PixelLocation AS POINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PixelLocation (BYVAL PixelLocation AS POINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HimetricLocation (BYVAL HimetricLocation AS POINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_HimetricLocation (BYVAL HimetricLocation AS POINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PixelLocationRaw (BYVAL PixelLocationRaw AS POINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PixelLocationRaw (BYVAL PixelLocationRaw AS POINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HimetricLocationRaw (BYVAL HimetricLocationRaw AS POINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_HimetricLocationRaw (BYVAL HimetricLocationRaw AS POINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Time (BYVAL Time AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Time (BYVAL Time AS ULONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HistoryCount (BYVAL HistoryCount AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_HistoryCount (BYVAL HistoryCount AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_InputData (BYVAL InputData AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_InputData (BYVAL InputData AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_KeyStates (BYVAL KeyStates AS ULONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_KeyStates (BYVAL KeyStates AS ULONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PerformanceCount (BYVAL PerformanceCount AS ULONGINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PerformanceCount (BYVAL PerformanceCount AS ULONGINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ButtonChangeKind (BYVAL ButtonChangeKind AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ButtonChangeKind (BYVAL ButtonChangeKind AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PenFlags (BYVAL PenFlags AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PenFlags (BYVAL PenFlags AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PenMask (BYVAL PenMask AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PenMask (BYVAL PenMask AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PenPressure (BYVAL PenPressure AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PenPressure (BYVAL PenPressure AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PenRotation (BYVAL PenRotation AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PenRotation (BYVAL PenRotation AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PenTiltX (BYVAL PenTiltX AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PenTiltX (BYVAL PenTiltX AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PenTiltY (BYVAL PenTiltY AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PenTiltY (BYVAL PenTiltY AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TouchFlags (BYVAL TouchFlags AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TouchFlags (BYVAL TouchFlags AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TouchMask (BYVAL TouchMask AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TouchMask (BYVAL TouchMask AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TouchContact (BYVAL TouchContact AS RECT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TouchContact (BYVAL TouchContact AS RECT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TouchContactRaw (BYVAL TouchContactRaw AS RECT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TouchContactRaw (BYVAL TouchContactRaw AS RECT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TouchOrientation (BYVAL TouchOrientation AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TouchOrientation (BYVAL TouchOrientation AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_TouchPressure (BYVAL TouchPressure AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_TouchPressure (BYVAL TouchPressure AS UINT) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PrintCompletedHandler
' IID: {8FD80075-ED08-42DB-8570-F5D14977461E}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PrintCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PrintCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PrintCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS COREWEBVIEW2_PRINT_STATUS) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PrintSettings
' IID: {377F3721-C74E-48CA-8DB1-DF68E51D60E2}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PrintSettings_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PrintSettings_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PrintSettings_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Orientation (BYVAL Orientation AS COREWEBVIEW2_PRINT_ORIENTATION PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Orientation (BYVAL Orientation AS COREWEBVIEW2_PRINT_ORIENTATION) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ScaleFactor (BYVAL ScaleFactor AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ScaleFactor (BYVAL ScaleFactor AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PageWidth (BYVAL PageWidth AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PageWidth (BYVAL PageWidth AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PageHeight (BYVAL PageHeight AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PageHeight (BYVAL PageHeight AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MarginTop (BYVAL MarginTop AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MarginTop (BYVAL MarginTop AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MarginBottom (BYVAL MarginBottom AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MarginBottom (BYVAL MarginBottom AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MarginLeft (BYVAL MarginLeft AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MarginLeft (BYVAL MarginLeft AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MarginRight (BYVAL MarginRight AS DOUBLE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MarginRight (BYVAL MarginRight AS DOUBLE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldPrintBackgrounds (BYVAL ShouldPrintBackgrounds AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ShouldPrintBackgrounds (BYVAL ShouldPrintBackgrounds AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldPrintSelectionOnly (BYVAL ShouldPrintSelectionOnly AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ShouldPrintSelectionOnly (BYVAL ShouldPrintSelectionOnly AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldPrintHeaderAndFooter (BYVAL ShouldPrintHeaderAndFooter AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ShouldPrintHeaderAndFooter (BYVAL ShouldPrintHeaderAndFooter AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HeaderTitle (BYVAL HeaderTitle AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_HeaderTitle (BYVAL HeaderTitle AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FooterUri (BYVAL FooterUri AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_FooterUri (BYVAL FooterUri AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PrintSettings2
' IID: {CA7F0E1F-3484-41D1-8C1A-65CD44A63F8D}
' Inherited interface = ICoreWebView2PrintSettings
' ########################################################################################

#ifndef __Afx_ICoreWebView2PrintSettings2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PrintSettings2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PrintSettings2_ EXTENDS Afx_ICoreWebView2PrintSettings
   DECLARE ABSTRACT FUNCTION get_PageRanges (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PageRanges (BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PagesPerSide (BYVAL value AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PagesPerSide (BYVAL value AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Copies (BYVAL value AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Copies (BYVAL value AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Collation (BYVAL value AS COREWEBVIEW2_PRINT_COLLATION PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Collation (BYVAL value AS COREWEBVIEW2_PRINT_COLLATION) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ColorMode (BYVAL value AS COREWEBVIEW2_PRINT_COLOR_MODE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ColorMode (BYVAL value AS COREWEBVIEW2_PRINT_COLOR_MODE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Duplex (BYVAL value AS COREWEBVIEW2_PRINT_DUPLEX PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Duplex (BYVAL value AS COREWEBVIEW2_PRINT_DUPLEX) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_MediaSize (BYVAL value AS COREWEBVIEW2_PRINT_MEDIA_SIZE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_MediaSize (BYVAL value AS COREWEBVIEW2_PRINT_MEDIA_SIZE) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PrinterName (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PrinterName (BYVAL value AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PrintToPdfCompletedHandler
' IID: {CCF1EF04-FD8E-4D5F-B2DE-0983E41B8C36}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PrintToPdfCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PrintToPdfCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PrintToPdfCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2PrintToPdfStreamCompletedHandler
' IID: {4C9F8229-8F93-444F-A711-2C0DFD6359D5}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2PrintToPdfStreamCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2PrintToPdfStreamCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2PrintToPdfStreamCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS IStream PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProcessExtendedInfo
' IID: {AF4C4C2E-45DB-11EE-BE56-0242AC120002}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessExtendedInfo_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessExtendedInfo_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessExtendedInfo_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ProcessInfo (BYVAL ProcessInfo AS Afx_ICoreWebView2ProcessInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AssociatedFrameInfos (BYVAL frames AS Afx_ICoreWebView2FrameInfoCollection PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProcessExtendedInfoCollection
' IID: {32EFA696-407A-11EE-BE56-0242AC120002}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessExtendedInfoCollection_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessExtendedInfoCollection_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessExtendedInfoCollection_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL value AS Afx_ICoreWebView2ProcessExtendedInfo PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProcessFailedEventArgs
' IID: {8155A9A4-1474-4A86-8CAE-151B0FA6B8CA}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessFailedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessFailedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessFailedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ProcessFailedKind (BYVAL value AS COREWEBVIEW2_PROCESS_FAILED_KIND PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProcessFailedEventArgs2
' IID: {4DAB9422-46FA-4C3E-A5D2-41D2071D3680}
' Inherited interface = ICoreWebView2ProcessFailedEventArgs
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessFailedEventArgs2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessFailedEventArgs2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessFailedEventArgs2_ EXTENDS Afx_ICoreWebView2ProcessFailedEventArgs
   DECLARE ABSTRACT FUNCTION get_reason (BYVAL reason AS COREWEBVIEW2_PROCESS_FAILED_REASON PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ExitCode (BYVAL ExitCode AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ProcessDescription (BYVAL ProcessDescription AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FrameInfosForFailedProcess (BYVAL frames AS Afx_ICoreWebView2FrameInfoCollection PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProcessFailedEventArgs3
' IID: {AB667428-094D-5FD1-B480-8B4C0FDBDF2F}
' Inherited interface = ICoreWebView2ProcessFailedEventArgs2
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessFailedEventArgs3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessFailedEventArgs3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessFailedEventArgs3_ EXTENDS Afx_ICoreWebView2ProcessFailedEventArgs2
   DECLARE ABSTRACT FUNCTION get_FailureSourceModulePath (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProcessFailedEventHandler
' IID: {79E0AEA4-990B-42D9-AA1D-0FCC2E5BC7F1}
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
' Interface name: ICoreWebView2ProcessInfo
' IID: {84FA7612-3F3D-4FBF-889D-FAD000492D72}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessInfo_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessInfo_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessInfo_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ProcessId (BYVAL value AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Kind (BYVAL Kind AS COREWEBVIEW2_PROCESS_KIND PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProcessInfoCollection
' IID: {402B99CD-A0CC-4FA5-B7A5-51D86A1D2339}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessInfoCollection_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessInfoCollection_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessInfoCollection_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL value AS Afx_ICoreWebView2ProcessInfo PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProcessInfosChangedEventHandler
' IID: {F4AF0C39-44B9-40E9-8B11-0484CFB9E0A1}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProcessInfosChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProcessInfosChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProcessInfosChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Environment PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Profile
' IID: {79110AD3-CD5D-4373-8BC3-C60658F17A5F}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Profile_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Profile_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Profile_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ProfileName (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsInPrivateModeEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ProfilePath (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DefaultDownloadFolderPath (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_DefaultDownloadFolderPath (BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_PreferredColorScheme (BYVAL value AS COREWEBVIEW2_PREFERRED_COLOR_SCHEME PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PreferredColorScheme (BYVAL value AS COREWEBVIEW2_PREFERRED_COLOR_SCHEME) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Profile2
' IID: {FA740D4B-5EAE-4344-A8AD-74BE31925397}
' Inherited interface = ICoreWebView2Profile
' ########################################################################################

#ifndef __Afx_ICoreWebView2Profile2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Profile2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Profile2_ EXTENDS Afx_ICoreWebView2Profile
   DECLARE ABSTRACT FUNCTION ClearBrowsingData (BYVAL dataKinds AS COREWEBVIEW2_BROWSING_DATA_KINDS, BYVAL handler AS Afx_ICoreWebView2ClearBrowsingDataCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ClearBrowsingDataInTimeRange (BYVAL dataKinds AS COREWEBVIEW2_BROWSING_DATA_KINDS, BYVAL startTime AS DOUBLE, BYVAL endTime AS DOUBLE, BYVAL handler AS Afx_ICoreWebView2ClearBrowsingDataCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ClearBrowsingDataAll (BYVAL handler AS Afx_ICoreWebView2ClearBrowsingDataCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Profile3
' IID: {B188E659-5685-4E05-BDBA-FC640E0F1992}
' Inherited interface = ICoreWebView2Profile2
' ########################################################################################

#ifndef __Afx_ICoreWebView2Profile3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Profile3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Profile3_ EXTENDS Afx_ICoreWebView2Profile2
   DECLARE ABSTRACT FUNCTION get_PreferredTrackingPreventionLevel (BYVAL value AS COREWEBVIEW2_TRACKING_PREVENTION_LEVEL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_PreferredTrackingPreventionLevel (BYVAL value AS COREWEBVIEW2_TRACKING_PREVENTION_LEVEL) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Profile4
' IID: {8F4AE680-192E-4EC8-833A-21CFADAEF628}
' Inherited interface = ICoreWebView2Profile3
' ########################################################################################

#ifndef __Afx_ICoreWebView2Profile4_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Profile4_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Profile4_ EXTENDS Afx_ICoreWebView2Profile3
   DECLARE ABSTRACT FUNCTION SetPermissionState (BYVAL PermissionKind AS COREWEBVIEW2_PERMISSION_KIND, BYVAL origin AS WSTRING PTR, BYVAL State AS COREWEBVIEW2_PERMISSION_STATE, BYVAL handler AS Afx_ICoreWebView2SetPermissionStateCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetNonDefaultPermissionSettings (BYVAL handler AS Afx_ICoreWebView2GetNonDefaultPermissionSettingsCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Profile5
' IID: {2EE5B76E-6E80-4DF2-BCD3-D4EC3340A01B}
' Inherited interface = ICoreWebView2Profile4
' ########################################################################################

#ifndef __Afx_ICoreWebView2Profile5_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Profile5_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Profile5_ EXTENDS Afx_ICoreWebView2Profile4
   DECLARE ABSTRACT FUNCTION get_CookieManager (BYVAL value AS Afx_ICoreWebView2CookieManager PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Profile6
' IID: {BD82FA6A-1D65-4C33-B2B4-0393020CC61B}
' Inherited interface = ICoreWebView2Profile5
' ########################################################################################

#ifndef __Afx_ICoreWebView2Profile6_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Profile6_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Profile6_ EXTENDS Afx_ICoreWebView2Profile5
   DECLARE ABSTRACT FUNCTION get_IsPasswordAutosaveEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsPasswordAutosaveEnabled (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsGeneralAutofillEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsGeneralAutofillEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Profile7
' IID: {7B4C7906-A1AA-4CB4-B723-DB09F813D541}
' Inherited interface = ICoreWebView2Profile6
' ########################################################################################

#ifndef __Afx_ICoreWebView2Profile7_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Profile7_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Profile7_ EXTENDS Afx_ICoreWebView2Profile6
   DECLARE ABSTRACT FUNCTION AddBrowserExtension (BYVAL extensionFolderPath AS WSTRING PTR, BYVAL handler AS Afx_ICoreWebView2ProfileAddBrowserExtensionCompletedHandler PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetBrowserExtensions (BYVAL handler AS Afx_ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Profile8
' IID: {FBF70C2F-EB1F-4383-85A0-163E92044011}
' Inherited interface = ICoreWebView2Profile7
' ########################################################################################

#ifndef __Afx_ICoreWebView2Profile8_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Profile8_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Profile8_ EXTENDS Afx_ICoreWebView2Profile7
   DECLARE ABSTRACT FUNCTION Delete_ () AS HRESULT
   DECLARE ABSTRACT FUNCTION add_Deleted (BYVAL eventHandler AS Afx_ICoreWebView2ProfileDeletedEventHandler PTR, BYVAL token AS EventRegistrationToken PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION remove_Deleted (BYVAL token AS EventRegistrationToken) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProfileAddBrowserExtensionCompletedHandler
' IID: {DF1AAB27-82B9-4AB6-AAE8-017A49398C14}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProfileAddBrowserExtensionCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProfileAddBrowserExtensionCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProfileAddBrowserExtensionCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS Afx_ICoreWebView2BrowserExtension PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProfileDeletedEventHandler
' IID: {DF35055D-772E-4DBE-B743-5FBF74A2B258}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProfileDeletedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProfileDeletedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProfileDeletedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Profile PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler
' IID: {FCE16A1C-F107-4601-8B75-FC4940AE25D0}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ProfileGetBrowserExtensionsCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS Afx_ICoreWebView2BrowserExtensionList PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2RasterizationScaleChangedEventHandler
' IID: {9C98C8B1-AC53-427E-A345-3049B5524BBE}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2RasterizationScaleChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2RasterizationScaleChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2RasterizationScaleChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2RegionRectCollectionView
' IID: {333353B8-48BF-4449-8FCC-22697FAF5753}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2RegionRectCollectionView_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2RegionRectCollectionView_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2RegionRectCollectionView_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL value AS RECT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2SaveAsUIShowingEventArgs
' IID: {55902952-0E0D-5AAA-A7D0-E833CDB34F62}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2SaveAsUIShowingEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2SaveAsUIShowingEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2SaveAsUIShowingEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ContentMimeType (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Cancel (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Cancel (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SuppressDefaultDialog (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SuppressDefaultDialog (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL value AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SaveAsFilePath (BYVAL value AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SaveAsFilePath (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AllowReplace (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AllowReplace (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Kind (BYVAL value AS COREWEBVIEW2_SAVE_AS_KIND) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Kind (BYVAL value AS COREWEBVIEW2_SAVE_AS_KIND PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2SaveAsUIShowingEventHandler
' IID: {6BAA177E-3A2E-5CCF-9A13-FAD676CD0522}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2SaveAsUIShowingEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2SaveAsUIShowingEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2SaveAsUIShowingEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2SaveAsUIShowingEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2SaveFileSecurityCheckStartingEventArgs
' IID: {CF4FF1D1-5A67-5660-8D63-EF699881EA65}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2SaveFileSecurityCheckStartingEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2SaveFileSecurityCheckStartingEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2SaveFileSecurityCheckStartingEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_CancelSave (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_CancelSave (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DocumentOriginUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FileExtension (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FilePath (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_SuppressDefaultPolicy (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_SuppressDefaultPolicy (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL value AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2SaveFileSecurityCheckStartingEventHandler
' IID: {7899576C-19E3-57C8-B7D1-55808292DE57}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2SaveFileSecurityCheckStartingEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2SaveFileSecurityCheckStartingEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2SaveFileSecurityCheckStartingEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2SaveFileSecurityCheckStartingEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ScreenCaptureStartingEventArgs
' IID: {892C03FD-AEE3-5EBA-A1FA-6FD2F6484B2B}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ScreenCaptureStartingEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ScreenCaptureStartingEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ScreenCaptureStartingEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Cancel (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Cancel (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Handled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Handled (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_OriginalSourceFrameInfo (BYVAL value AS Afx_ICoreWebView2FrameInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL value AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ScreenCaptureStartingEventHandler
' IID: {E24FF05A-1DB5-59D9-89F3-3C864268DB4A}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ScreenCaptureStartingEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ScreenCaptureStartingEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ScreenCaptureStartingEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ScreenCaptureStartingEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ScriptDialogOpeningEventArgs
' IID: {7390BB70-ABE0-4843-9529-F143B31B03D6}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ScriptDialogOpeningEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ScriptDialogOpeningEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ScriptDialogOpeningEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Kind (BYVAL Kind AS COREWEBVIEW2_SCRIPT_DIALOG_KIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Message (BYVAL Message AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Accept () AS HRESULT
   DECLARE ABSTRACT FUNCTION get_DefaultText (BYVAL DefaultText AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResultText (BYVAL ResultText AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ResultText (BYVAL ResultText AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ScriptDialogOpeningEventHandler
' IID: {EF381BF9-AFA8-4E37-91C4-8AC48524BDFB}
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
' Interface name: ICoreWebView2ScriptException
' IID: {054DAE00-84A3-49FF-BC17-4012A90BC9FD}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ScriptException_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ScriptException_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ScriptException_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_LineNumber (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ColumnNumber (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_name (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Message (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ToJson (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ServerCertificateErrorDetectedEventArgs
' IID: {012193ED-7C13-48FF-969D-A84C1F432A14}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ServerCertificateErrorDetectedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ServerCertificateErrorDetectedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ServerCertificateErrorDetectedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_ErrorStatus (BYVAL value AS COREWEBVIEW2_WEB_ERROR_STATUS PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_RequestUri (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ServerCertificate (BYVAL value AS Afx_ICoreWebView2Certificate PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Action (BYVAL value AS COREWEBVIEW2_SERVER_CERTIFICATE_ERROR_ACTION PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Action (BYVAL value AS COREWEBVIEW2_SERVER_CERTIFICATE_ERROR_ACTION) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ServerCertificateErrorDetectedEventHandler
' IID: {969B3A26-D85E-4795-8199-FEF57344DA22}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ServerCertificateErrorDetectedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ServerCertificateErrorDetectedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ServerCertificateErrorDetectedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ServerCertificateErrorDetectedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2SetPermissionStateCompletedHandler
' IID: {FC77FB30-9C9E-4076-B8C7-7644A703CA1B}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2SetPermissionStateCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2SetPermissionStateCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2SetPermissionStateCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Settings
' IID: {E562E4F0-D7FA-43AC-8D71-C05150499F00}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsScriptEnabled (BYVAL IsScriptEnabled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsScriptEnabled (BYVAL IsScriptEnabled AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsWebMessageEnabled (BYVAL IsWebMessageEnabled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsWebMessageEnabled (BYVAL IsWebMessageEnabled AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AreDefaultScriptDialogsEnabled (BYVAL AreDefaultScriptDialogsEnabled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreDefaultScriptDialogsEnabled (BYVAL AreDefaultScriptDialogsEnabled AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsStatusBarEnabled (BYVAL IsStatusBarEnabled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsStatusBarEnabled (BYVAL IsStatusBarEnabled AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AreDevToolsEnabled (BYVAL AreDevToolsEnabled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreDevToolsEnabled (BYVAL AreDevToolsEnabled AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AreDefaultContextMenusEnabled (BYVAL enabled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreDefaultContextMenusEnabled (BYVAL enabled AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_AreHostObjectsAllowed (BYVAL allowed AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreHostObjectsAllowed (BYVAL allowed AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsZoomControlEnabled (BYVAL enabled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsZoomControlEnabled (BYVAL enabled AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsBuiltInErrorPageEnabled (BYVAL enabled AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsBuiltInErrorPageEnabled (BYVAL enabled AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Settings2
' IID: {EE9A0F68-F46C-4E32-AC23-EF8CAC224D2A}
' Inherited interface = ICoreWebView2Settings
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings2_ EXTENDS Afx_ICoreWebView2Settings
   DECLARE ABSTRACT FUNCTION get_UserAgent (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_UserAgent (BYVAL value AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Settings3
' IID: {FDB5AB74-AF33-4854-84F0-0A631DEB5EBA}
' Inherited interface = ICoreWebView2Settings2
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings3_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings3_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings3_ EXTENDS Afx_ICoreWebView2Settings2
   DECLARE ABSTRACT FUNCTION get_AreBrowserAcceleratorKeysEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_AreBrowserAcceleratorKeysEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Settings4
' IID: {CB56846C-4168-4D53-B04F-03B6D6796FF2}
' Inherited interface = ICoreWebView2Settings3
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings4_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings4_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings4_ EXTENDS Afx_ICoreWebView2Settings3
   DECLARE ABSTRACT FUNCTION get_IsPasswordAutosaveEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsPasswordAutosaveEnabled (BYVAL value AS LONG) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_IsGeneralAutofillEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsGeneralAutofillEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Settings5
' IID: {183E7052-1D03-43A0-AB99-98E043B66B39}
' Inherited interface = ICoreWebView2Settings4
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings5_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings5_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings5_ EXTENDS Afx_ICoreWebView2Settings4
   DECLARE ABSTRACT FUNCTION get_IsPinchZoomEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsPinchZoomEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Settings6
' IID: {11CB3ACD-9BC8-43B8-83BF-F40753714F87}
' Inherited interface = ICoreWebView2Settings5
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings6_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings6_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings6_ EXTENDS Afx_ICoreWebView2Settings5
   DECLARE ABSTRACT FUNCTION get_IsSwipeNavigationEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsSwipeNavigationEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Settings7
' IID: {488DC902-35EF-42D2-BC7D-94B65C4BC49C}
' Inherited interface = ICoreWebView2Settings6
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings7_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings7_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings7_ EXTENDS Afx_ICoreWebView2Settings6
   DECLARE ABSTRACT FUNCTION get_HiddenPdfToolbarItems (BYVAL value AS COREWEBVIEW2_PDF_TOOLBAR_ITEMS PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_HiddenPdfToolbarItems (BYVAL value AS COREWEBVIEW2_PDF_TOOLBAR_ITEMS) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Settings8
' IID: {9E6B0E8F-86AD-4E81-8147-A9B5EDB68650}
' Inherited interface = ICoreWebView2Settings7
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings8_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings8_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings8_ EXTENDS Afx_ICoreWebView2Settings7
   DECLARE ABSTRACT FUNCTION get_IsReputationCheckingRequired (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsReputationCheckingRequired (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2Settings9
' IID: {0528A73B-E92D-49F4-927A-E547DDDAA37D}
' Inherited interface = ICoreWebView2Settings8
' ########################################################################################

#ifndef __Afx_ICoreWebView2Settings9_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2Settings9_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2Settings9_ EXTENDS Afx_ICoreWebView2Settings8
   DECLARE ABSTRACT FUNCTION get_IsNonClientRegionSupportEnabled (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_IsNonClientRegionSupportEnabled (BYVAL value AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2SharedBuffer
' IID: {B747A495-0C6F-449E-97B8-2F81E9D6AB43}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2SharedBuffer_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2SharedBuffer_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2SharedBuffer_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Size (BYVAL value AS ULONGINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Buffer (BYVAL value AS UBYTE PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION OpenStream (BYVAL value AS IStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_FileMappingHandle (BYVAL value AS ANY PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Close () AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ShowSaveAsUICompletedHandler
' IID: {E24B07E3-8169-5C34-994A-7F6478946A3C}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ShowSaveAsUICompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ShowSaveAsUICompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ShowSaveAsUICompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS COREWEBVIEW2_SAVE_AS_UI_RESULT) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2SourceChangedEventArgs
' IID: {31E0E545-1DBA-4266-8914-F63848A1F7D7}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2SourceChangedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2SourceChangedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2SourceChangedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_IsNewDocument (BYVAL value AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2SourceChangedEventHandler
' IID: {3C067F9F-5388-4772-8B48-79F7EF1AB37C}
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
' Interface name: ICoreWebView2StateChangedEventHandler
' IID: {81336594-7EDE-4BA9-BF71-ACF0A95B58DD}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2StateChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2StateChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2StateChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2DownloadOperation PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2StatusBarTextChangedEventHandler
' IID: {A5E3B0D0-10DF-4156-BFAD-3B43867ACAC6}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2StatusBarTextChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2StatusBarTextChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2StatusBarTextChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2StringCollection
' IID: {F41F3F8A-BCC3-11EB-8529-0242AC130003}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2StringCollection_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2StringCollection_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2StringCollection_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Count (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValueAtIndex (BYVAL index AS UINT, BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2TrySuspendCompletedHandler
' IID: {00F206A7-9D17-4605-91F6-4E8E4DE192E3}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2TrySuspendCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2TrySuspendCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2TrySuspendCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS LONG) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebMessageReceivedEventArgs
' IID: {0F99A40C-E962-4207-9E92-E3D542EFF849}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebMessageReceivedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebMessageReceivedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebMessageReceivedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Source (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_webMessageAsJson (BYVAL value AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION TryGetWebMessageAsString (BYVAL value AS WSTRING PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebMessageReceivedEventArgs2
' IID: {06FC7AB7-C90C-4297-9389-33CA01CF6D5E}
' Inherited interface = ICoreWebView2WebMessageReceivedEventArgs
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebMessageReceivedEventArgs2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebMessageReceivedEventArgs2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebMessageReceivedEventArgs2_ EXTENDS Afx_ICoreWebView2WebMessageReceivedEventArgs
   DECLARE ABSTRACT FUNCTION get_additionalObjects (BYVAL value AS Afx_ICoreWebView2ObjectCollectionView PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebMessageReceivedEventHandler
' IID: {57213F19-00E6-49FA-8E07-898EA01ECBD2}
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
' IID: {97055CD4-512C-4264-8B5F-E3F446CEA6A5}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceRequest_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceRequest_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceRequest_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_uri (BYVAL uri AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_uri (BYVAL uri AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Method (BYVAL Method AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Method (BYVAL Method AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Content (BYVAL Content AS IStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Content (BYVAL Content AS IStream PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Headers (BYVAL Headers AS Afx_ICoreWebView2HttpRequestHeaders PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebResourceRequestedEventArgs
' IID: {453E667F-12C7-49D4-BE6D-DDBE7956F57A}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceRequestedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceRequestedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceRequestedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Request (BYVAL Request AS Afx_ICoreWebView2WebResourceRequest PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Response (BYVAL Response AS Afx_ICoreWebView2WebResourceResponse PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Response (BYVAL Response AS Afx_ICoreWebView2WebResourceResponse PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDeferral (BYVAL deferral AS Afx_ICoreWebView2Deferral PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ResourceContext (BYVAL context AS COREWEBVIEW2_WEB_RESOURCE_CONTEXT PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebResourceRequestedEventArgs2
' IID: {9C562C24-B219-4D7F-92F6-B187FBBADD56}
' Inherited interface = ICoreWebView2WebResourceRequestedEventArgs
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceRequestedEventArgs2_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceRequestedEventArgs2_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceRequestedEventArgs2_ EXTENDS Afx_ICoreWebView2WebResourceRequestedEventArgs
   DECLARE ABSTRACT FUNCTION get_RequestedSourceKind (BYVAL value AS COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebResourceRequestedEventHandler
' IID: {AB00B74C-15F1-4646-80E8-E76341D25D71}
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
' IID: {AAFCC94F-FA27-48FD-97DF-830EF75AAEC9}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceResponse_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceResponse_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceResponse_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Content (BYVAL Content AS IStream PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_Content (BYVAL Content AS IStream PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Headers (BYVAL Headers AS Afx_ICoreWebView2HttpResponseHeaders PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_StatusCode (BYVAL StatusCode AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_StatusCode (BYVAL StatusCode AS INT_) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ReasonPhrase (BYVAL ReasonPhrase AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION put_ReasonPhrase (BYVAL ReasonPhrase AS WSTRING PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebResourceResponseReceivedEventArgs
' IID: {D1DB483D-6796-4B8B-80FC-13712BB716F4}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceResponseReceivedEventArgs_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceResponseReceivedEventArgs_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceResponseReceivedEventArgs_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Request (BYVAL value AS Afx_ICoreWebView2WebResourceRequest PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Response (BYVAL value AS Afx_ICoreWebView2WebResourceResponseView PTR PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebResourceResponseReceivedEventHandler
' IID: {7DE9898A-24F5-40C3-A2DE-D4F458E69828}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceResponseReceivedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceResponseReceivedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceResponseReceivedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2WebResourceResponseReceivedEventArgs PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebResourceResponseView
' IID: {79701053-7759-4162-8F7D-F1B3F084928D}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceResponseView_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceResponseView_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceResponseView_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_Headers (BYVAL Headers AS Afx_ICoreWebView2HttpResponseHeaders PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_StatusCode (BYVAL StatusCode AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ReasonPhrase (BYVAL ReasonPhrase AS WSTRING PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetContent (BYVAL handler AS Afx_ICoreWebView2WebResourceResponseViewGetContentCompletedHandler PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WebResourceResponseViewGetContentCompletedHandler
' IID: {875738E1-9FA2-40E3-8B74-2E8972DD6FE7}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WebResourceResponseViewGetContentCompletedHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WebResourceResponseViewGetContentCompletedHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WebResourceResponseViewGetContentCompletedHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL result AS IStream PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2WindowCloseRequestedEventHandler
' IID: {5C19E9E0-092F-486B-AFFA-CA8231913039}
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
' Interface name: ICoreWebView2WindowFeatures
' IID: {5EAF559F-B46E-4397-8860-E422F287FF1E}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2WindowFeatures_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2WindowFeatures_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2WindowFeatures_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION get_HasPosition (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_HasSize (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Left (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Top (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Height (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_Width (BYVAL value AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldDisplayMenuBar (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldDisplayStatus (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldDisplayToolbar (BYVAL value AS LONG PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION get_ShouldDisplayScrollBars (BYVAL value AS LONG PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

' ########################################################################################
' Interface name: ICoreWebView2ZoomFactorChangedEventHandler
' IID: {B52D71D6-C4DF-4543-A90C-64A3E60F38CB}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_ICoreWebView2ZoomFactorChangedEventHandler_INTERFACE_DEFINED__
#define __Afx_ICoreWebView2ZoomFactorChangedEventHandler_INTERFACE_DEFINED__

TYPE Afx_ICoreWebView2ZoomFactorChangedEventHandler_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS IUnknown PTR) AS HRESULT
END TYPE

#endif

' ########################################################################################

