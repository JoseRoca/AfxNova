' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxDwmApi.bi
' Purpose:  TRanslation of DwmApi.h
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
#include once "win/winapifamily.bi"
#include once "win/wtypes.bi"
#include once "win/uxtheme.bi"

extern "Windows"

' // Blur behind data structures
CONST DWM_BB_ENABLE                = &h00000001   ' // fEnable has been specified
CONST DWM_BB_BLURREGION            = &h00000002   ' // hRgnBlur has been specified
CONST DWM_BB_TRANSITIONONMAXIMIZED = &h00000004   ' // fTransitionOnMaximized has been specified

TYPE _DWM_BLURBEHIND
	dwFlags AS DWORD
	fEnable AS WINBOOL
	hRgnBlur AS HRGN
	fTransitionOnMaximized AS WINBOOL
END TYPE

TYPE DWM_BLURBEHIND AS _DWM_BLURBEHIND
TYPE PDWM_BLURBEHIND AS _DWM_BLURBEHIND PTR

' // Window attributes
TYPE DWMWINDOWATTRIBUTE AS LONG
ENUM
	DWMWA_NCRENDERING_ENABLED = 1
	DWMWA_NCRENDERING_POLICY
	DWMWA_TRANSITIONS_FORCEDISABLED
	DWMWA_ALLOW_NCPAINT
	DWMWA_CAPTION_BUTTON_BOUNDS
	DWMWA_NONCLIENT_RTL_LAYOUT
	DWMWA_FORCE_ICONIC_REPRESENTATION
	DWMWA_FLIP3D_POLICY
	DWMWA_EXTENDED_FRAME_BOUNDS
	DWMWA_HAS_ICONIC_BITMAP
	DWMWA_DISALLOW_PEEK
	DWMWA_EXCLUDED_FROM_PEEK
	DWMWA_CLOAK
	DWMWA_CLOAKED
	DWMWA_FREEZE_REPRESENTATION
	DWMWA_PASSIVE_UPDATE_MODE
	DWMWA_USE_HOSTBACKDROPBRUSH
	DWMWA_USE_IMMERSIVE_DARK_MODE = 20
	DWMWA_WINDOW_CORNER_PREFERENCE = 33
	DWMWA_BORDER_COLOR
	DWMWA_CAPTION_COLOR
	DWMWA_TEXT_COLOR
	DWMWA_VISIBLE_FRAME_BORDER_THICKNESS
	DWMWA_SYSTEMBACKDROP_TYPE
	DWMWA_REDIRECTIONBITMAP_ALPHA
	DWMWA_LAST
END ENUM

TYPE DWM_WINDOW_CORNER_PREFERENCE AS LONG
ENUM
	DWMWCP_DEFAULT = 0
	DWMWCP_DONOTROUND = 1
	DWMWCP_ROUND = 2
	DWMWCP_ROUNDSMALL = 3
END ENUM

CONST DWMWA_COLOR_DEFAULT = &hFFFFFFFF
CONST DWMWA_COLOR_NONE = &hFFFFFFFE

TYPE DWM_SYSTEMBACKDROP_TYPE AS LONG
ENUM
	DWMSBT_AUTO
	DWMSBT_NONE
	DWMSBT_MAINWINDOW
	DWMSBT_TRANSIENTWINDOW
	DWMSBT_TABBEDWINDOW
END ENUM

TYPE DWMNCRENDERINGPOLICY AS LONG
ENUM
	DWMNCRP_USEWINDOWSTYLE
	DWMNCRP_DISABLED
	DWMNCRP_ENABLED
	DWMNCRP_LAST
END ENUM

TYPE DWMFLIP3DWINDOWPOLICY AS LONG
ENUM
	DWMFLIP3D_DEFAULT
	DWMFLIP3D_EXCLUDEBELOW
	DWMFLIP3D_EXCLUDEABOVE
	DWMFLIP3D_LAST
END ENUM

CONST DWM_CLOAKED_APP = &h00000001
CONST DWM_CLOAKED_SHELL = &h00000002
CONST DWM_CLOAKED_INHERITED = &h00000004
TYPE HTHUMBNAIL AS HANDLE
TYPE PHTHUMBNAIL AS HTHUMBNAIL PTR
CONST DWM_TNP_RECTDESTINATION = &h00000001
CONST DWM_TNP_RECTSOURCE = &h00000002
CONST DWM_TNP_OPACITY = &h00000004
CONST DWM_TNP_VISIBLE = &h00000008
CONST DWM_TNP_SOURCECLIENTAREAONLY = &h00000010

TYPE _DWM_THUMBNAIL_PROPERTIES
	dwFlags AS DWORD
	rcDestination AS RECT
	rcSource AS RECT
	opacity AS BYTE
	fVisible AS WINBOOL
	fSourceClientAreaOnly AS WINBOOL
END TYPE

TYPE DWM_THUMBNAIL_PROPERTIES AS _DWM_THUMBNAIL_PROPERTIES
TYPE PDWM_THUMBNAIL_PROPERTIES AS _DWM_THUMBNAIL_PROPERTIES PTR
TYPE DWM_FRAME_COUNT AS ULONGLONG
TYPE QPC_TIME AS ULONGLONG

TYPE _UNSIGNED_RATIO
	uiNumerator AS UINT32
	uiDenominator AS UINT32
END TYPE

TYPE UNSIGNED_RATIO AS _UNSIGNED_RATIO

TYPE _DWM_TIMING_INFO
	cbSize AS UINT32
	rateRefresh AS UNSIGNED_RATIO
	qpcRefreshPeriod AS QPC_TIME
	rateCompose AS UNSIGNED_RATIO
	qpcVBlank AS QPC_TIME
	cRefresh AS DWM_FRAME_COUNT
	cDXRefresh AS UINT
	qpcCompose AS QPC_TIME
	cFrame AS DWM_FRAME_COUNT
	cDXPresent AS UINT
	cRefreshFrame AS DWM_FRAME_COUNT
	cFrameSubmitted AS DWM_FRAME_COUNT
	cDXPresentSubmitted AS UINT
	cFrameConfirmed AS DWM_FRAME_COUNT
	cDXPresentConfirmed AS UINT
	cRefreshConfirmed AS DWM_FRAME_COUNT
	cDXRefreshConfirmed AS UINT
	cFramesLate AS DWM_FRAME_COUNT
	cFramesOutstanding AS UINT
	cFrameDisplayed AS DWM_FRAME_COUNT
	qpcFrameDisplayed AS QPC_TIME
	cRefreshFrameDisplayed AS DWM_FRAME_COUNT
	cFrameComplete AS DWM_FRAME_COUNT
	qpcFrameComplete AS QPC_TIME
	cFramePending AS DWM_FRAME_COUNT
	qpcFramePending AS QPC_TIME
	cFramesDisplayed AS DWM_FRAME_COUNT
	cFramesComplete AS DWM_FRAME_COUNT
	cFramesPending AS DWM_FRAME_COUNT
	cFramesAvailable AS DWM_FRAME_COUNT
	cFramesDropped AS DWM_FRAME_COUNT
	cFramesMissed AS DWM_FRAME_COUNT
	cRefreshNextDisplayed AS DWM_FRAME_COUNT
	cRefreshNextPresented AS DWM_FRAME_COUNT
	cRefreshesDisplayed AS DWM_FRAME_COUNT
	cRefreshesPresented AS DWM_FRAME_COUNT
	cRefreshStarted AS DWM_FRAME_COUNT
	cPixelsReceived AS ULONGLONG
	cPixelsDrawn AS ULONGLONG
	cBuffersEmpty AS DWM_FRAME_COUNT
END TYPE

TYPE DWM_TIMING_INFO AS _DWM_TIMING_INFO

TYPE DWM_SOURCE_FRAME_SAMPLING AS LONG
ENUM
	DWM_SOURCE_FRAME_SAMPLING_POINT
	DWM_SOURCE_FRAME_SAMPLING_COVERAGE
	DWM_SOURCE_FRAME_SAMPLING_LAST
END ENUM

'EXTERN_C __declspec(selectany) CONST UINT c_DwmMaxQueuedBuffers = 8;
'EXTERN_C __declspec(selectany) CONST UINT c_DwmMaxMonitors = 16;
'EXTERN_C __declspec(selectany) CONST UINT c_DwmMaxAdapters = 16;

TYPE _DWM_PRESENT_PARAMETERS
	cbSize AS UINT32
	fQueue AS WINBOOL
	cRefreshStart AS DWM_FRAME_COUNT
	cBuffer AS UINT
	fUseSourceRate AS WINBOOL
	rateSource AS UNSIGNED_RATIO
	cRefreshesPerFrame AS UINT
	eSampling AS DWM_SOURCE_FRAME_SAMPLING
END TYPE

TYPE DWM_PRESENT_PARAMETERS AS _DWM_PRESENT_PARAMETERS
CONST DWM_FRAME_DURATION_DEFAULT = -1

DECLARE FUNCTION DwmDefWindowProc LIB "dwmapi.dll" ALIAS "DwmDefWindowProc" ( _
   BYVAL hwnd AS HWND, BYVAL msg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM, BYVAL plResult AS LRESULT PTR) AS WINBOOL
DECLARE FUNCTION DwmEnableBlurBehindWindow LIB "dwmapi.dll" ALIAS "DwmEnableBlurBehindWindow" ( _
   BYVAL hwnd AS HWND, BYVAL msg AS UINT, BYVAL pBlurBehind AS DWM_BLURBEHIND PTR) AS HRESULT


CONST DWM_EC_DISABLECOMPOSITION = 0
CONST DWM_EC_ENABLECOMPOSITION = 1


DECLARE FUNCTION DwmEnableComposition LIB "dwmapi.dll" ALIAS "DwmEnableComposition" (BYVAL uCompositionAction AS UINT) AS HRESULT
DECLARE FUNCTION DwmEnableMMCSS LIB "dwmapi.dll" ALIAS "DwmEnableMMCSS" (BYVAL fEnableMMCSS AS WINBOOL) AS HRESULT
DECLARE FUNCTION DwmExtendFrameIntoClientArea LIB "dwmapi.dll" ALIAS "DwmExtendFrameIntoClientArea" (BYVAL hWnd AS HWND, BYVAL pMarInset AS CONST MARGINS PTR) AS HRESULT
DECLARE FUNCTION DwmGetColorizationColor LIB "dwmapi.dll" ALIAS "DwmGetColorizationColor" (BYVAL pcrColorization AS DWORD PTR, BYVAL pfOpaqueBlend AS WINBOOL PTR) AS HRESULT
DECLARE FUNCTION DwmGetCompositionTimingInfo LIB "dwmapi.dll" ALIAS "DwmGetCompositionTimingInfo" (BYVAL hwnd AS HWND, BYVAL pTimingInfo AS DWM_TIMING_INFO PTR) AS HRESULT
DECLARE FUNCTION DwmGetWindowAttribute LIB "dwmapi.dll" ALIAS "DwmGetWindowAttribute" (BYVAL hwnd AS HWND, BYVAL dwAttribute AS DWORD, BYVAL pvAttribute AS PVOID, BYVAL cbAttribute AS DWORD) AS HRESULT
DECLARE FUNCTION DwmIsCompositionEnabled LIB "dwmapi.dll" ALIAS "DwmIsCompositionEnabled" (BYVAL pfEnabled AS WINBOOL PTR) AS HRESULT
DECLARE FUNCTION DwmModifyPreviousDxFrameDuration LIB "dwmapi.dll" ALIAS "DwmModifyPreviousDxFrameDuration" (BYVAL hwnd AS HWND, BYVAL cRefreshes AS INT_, BYVAL fRelative AS WINBOOL) AS HRESULT
DECLARE FUNCTION DwmQueryThumbnailSourceSize LIB "dwmapi.dll" ALIAS "DwmQueryThumbnailSourceSize" (BYVAL hThumbnail AS HTHUMBNAIL, BYVAL pSize AS PSIZE) AS HRESULT
DECLARE FUNCTION DwmRegisterThumbnail LIB "dwmapi.dll" ALIAS "DwmRegisterThumbnail" (BYVAL hwndDestination AS HWND, BYVAL hwndSource AS HWND, BYVAL phThumbnailId AS PHTHUMBNAIL) AS HRESULT
DECLARE FUNCTION DwmSetDxFrameDuration LIB "dwmapi.dll" ALIAS "DwmSetDxFrameDuration" (BYVAL hwnd AS HWND, BYVAL cRefreshes AS INT_) AS HRESULT
DECLARE FUNCTION DwmSetPresentParameters LIB "dwmapi.dll" ALIAS "DwmSetPresentParameters" (BYVAL hwnd AS HWND, BYVAL pPresentParams AS DWM_PRESENT_PARAMETERS PTR) AS HRESULT
DECLARE FUNCTION DwmSetWindowAttribute LIB "dwmapi.dll" ALIAS "DwmSetWindowAttribute" (BYVAL hwnd AS HWND, BYVAL dwAttribute AS DWORD, BYVAL pvAttribute AS LPCVOID, BYVAL cbAttribute AS DWORD) AS HRESULT
DECLARE FUNCTION DwmUnregisterThumbnail LIB "dwmapi.dll" ALIAS "DwmUnregisterThumbnail" (BYVAL hThumbnailId AS HTHUMBNAIL) AS HRESULT
DECLARE FUNCTION DwmUpdateThumbnailProperties LIB "dwmapi.dll" ALIAS "DwmUpdateThumbnailProperties" (BYVAL hThumbnailId AS HTHUMBNAIL, BYVAL ptnProperties AS DWM_THUMBNAIL_PROPERTIES PTR) AS HRESULT
DECLARE FUNCTION DwmAttachMilContent LIB "dwmapi.dll" ALIAS "DwmAttachMilContent" (BYVAL hwnd AS HWND) AS HRESULT
DECLARE FUNCTION DwmDetachMilContent LIB "dwmapi.dll" ALIAS "DwmDetachMilContent" (BYVAL hwnd AS HWND) AS HRESULT
DECLARE FUNCTION DwmFlush LIB "dwmapi.dll" ALIAS "DwmFlush" () AS HRESULT

TYPE _MilMatrix3x2D
	S_11 AS DOUBLE
	S_12 AS DOUBLE
	S_21 AS DOUBLE
	S_22 AS DOUBLE
	DX AS DOUBLE
	DY AS DOUBLE
END TYPE

TYPE MilMatrix3x2D AS _MilMatrix3x2D
#define _MIL_MATRIX3X2D_DEFINED
TYPE MIL_MATRIX3X2D AS MilMatrix3x2D
#define MILCORE_MIL_MATRIX3X2D_COMPAT_TYPEDEF

DECLARE FUNCTION DwmGetGraphicsStreamTransformHint LIB "dwmapi.dll" ALIAS "DwmGetGraphicsStreamTransformHint" (BYVAL uIndex AS UINT, BYVAL PTRansform AS MilMatrix3x2D PTR) AS HRESULT
DECLARE FUNCTION DwmGetGraphicsStreamClient LIB "dwmapi.dll" ALIAS "DwmGetGraphicsStreamClient" (BYVAL uIndex AS UINT, BYVAL pClientUuid AS UUID PTR) AS HRESULT
DECLARE FUNCTION DwmGetTransportAttributes LIB "dwmapi.dll" ALIAS "DwmGetTransportAttributes" (BYVAL pfIsRemoting AS WINBOOL PTR, BYVAL pfIsConnected AS WINBOOL PTR, BYVAL pDwGeneration AS DWORD PTR) AS HRESULT

TYPE DWMTRANSITION_OWNEDWINDOW_TARGET AS LONG
ENUM
	DWMTRANSITION_OWNEDWINDOW_NULL = -1
	DWMTRANSITION_OWNEDWINDOW_REPOSITION = 0
END ENUM

DECLARE FUNCTION DwmTransitionOwnedWindow LIB "dwmapi.dll" ALIAS "DwmTransitionOwnedWindow" ( _
   BYVAL hwnd AS HWND, BYVAL target AS DWMTRANSITION_OWNEDWINDOW_TARGET) AS HRESULT

TYPE GESTURE_TYPE AS LONG
ENUM
	GT_PEN_TAP = 0
	GT_PEN_DOUBLETAP = 1
	GT_PEN_RIGHTTAP = 2
	GT_PEN_PRESSANDHOLD = 3
	GT_PEN_PRESSANDHOLDABORT = 4
	GT_TOUCH_TAP = 5
	GT_TOUCH_DOUBLETAP = 6
	GT_TOUCH_RIGHTTAP = 7
	GT_TOUCH_PRESSANDHOLD = 8
	GT_TOUCH_PRESSANDHOLDABORT = 9
	GT_TOUCH_PRESSANDTAP = 10
END ENUM

DECLARE FUNCTION DwmRenderGesture LIB "dwmapi.dll" ALIAS "DwmRenderGesture" (BYVAL gt AS GESTURE_TYPE, BYVAL cContacts AS UINT, BYVAL pdwPointerID AS DWORD PTR, BYVAL pPoints AS POINT PTR) AS HRESULT
DECLARE FUNCTION DwmTetherContact LIB "dwmapi.dll" ALIAS "DwmTetherContact" (BYVAL dwPointerID AS DWORD, BYVAL fEnable AS WINBOOL, BYVAL ptTether AS POINT) AS HRESULT

TYPE DWM_SHOWCONTACT AS LONG
ENUM
	DWMSC_DOWN = &h00000001
	DWMSC_UP = &h00000002
	DWMSC_DRAG = &h00000004
	DWMSC_HOLD = &h00000008
	DWMSC_PENBARREL = &h00000010
	DWMSC_NONE = &h00000000
	DWMSC_ALL = &hFFFFFFFF
END ENUM

'DEFINE_ENUM_FLAG_OPERATORS(DWM_SHOWCONTACT);

DECLARE FUNCTION DwmShowContact LIB "dwmapi.dll" ALIAS "DwmShowContact" (BYVAL dwPointerID AS DWORD, BYVAL eShowContact AS DWM_SHOWCONTACT) AS HRESULT

TYPE DWM_TAB_WINDOW_REQUIREMENTS AS LONG
ENUM
	DWMTWR_NONE = &h0000
	DWMTWR_IMPLEMENTED_BY_SYSTEM = &h0001
	DWMTWR_WINDOW_RELATIONSHIP = &h0002
	DWMTWR_WINDOW_STYLES = &h0004
	DWMTWR_WINDOW_REGION = &h0008
	DWMTWR_WINDOW_DWM_ATTRIBUTES = &h0010
	DWMTWR_WINDOW_MARGINS = &h0020
	DWMTWR_TABBING_ENABLED = &h0040
	DWMTWR_USER_POLICY = &h0080
	DWMTWR_GROUP_POLICY = &h0100
	DWMTWR_APP_COMPAT = &h0200
END ENUM

'DEFINE_ENUM_FLAG_OPERATORS(DWM_TAB_WINDOW_REQUIREMENTS);

DECLARE FUNCTION DwmGetUnmetTabRequirements LIB "dwmapi.dll" ALIAS "DwmGetUnmetTabRequirements" ( _
   BYVAL appWindow AS HWND, BYVAL value AS DWM_TAB_WINDOW_REQUIREMENTS PTR) AS HRESULT

end extern


' ========================================================================================
' Allows the window frame for this window to be drawn in dark mode colors.
' ========================================================================================
PRIVATE FUNCTION AfxEnableDarkModeForWindow (BYVAL hwnd AS HWND) AS HRESULT
   DIM value AS BOOL = CTRUE
   RETURN DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, @value, SIZEOF(value))
END FUNCTION
' ========================================================================================
