' ########################################################################################
' Platform: Microsoft Windows
' Filename: CTaskDialog.inc
' Purpose:  Windows common controls - Task dialog
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2026 José Roca
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
#include once "win/uxtheme.bi"
#include once "AfxNova/DWSTRING.inc"
USING AfxNova

#if _WIN32_WINNT >= &h0600

NAMESPACE AfxNova

' ########################################################################################
' CTaskDialog class
' ########################################################################################

TYPE CTaskDialog

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC SUB ClickButton (BYVAL hTaskDlg AS HWND, BYVAL nButton AS DWORD)
   DECLARE STATIC SUB ClickRadioButton (BYVAL hTaskDlg AS HWND, BYVAL nRadioButton AS DWORD)
   DECLARE STATIC SUB ClickVerification (BYVAL hTaskDlg AS HWND, BYVAL fChecked AS BOOLEAN, BYVAL fFocused AS BOOLEAN)
   DECLARE STATIC SUB EnableButton (BYVAL hTaskDlg AS HWND, BYVAL nButton AS LONG)
   DECLARE STATIC SUB DisableButton (BYVAL hTaskDlg AS HWND, BYVAL nButton AS LONG)
   DECLARE STATIC SUB EnableRadioButton (BYVAL hTaskDlg AS HWND, BYVAL nRadioButton AS LONG)
   DECLARE STATIC SUB DisableRadioButton (BYVAL hTaskDlg AS HWND, BYVAL nRadioButton AS LONG)
   DECLARE STATIC SUB NavigatePage (BYVAL hTaskDlg AS HWND, BYVAL pTaskConfig AS TASKDIALOGCONFIG PTR)
   DECLARE STATIC SUB NavigatePage (BYVAL hTaskDlg AS HWND, BYREF TaskConfig AS TASKDIALOGCONFIG)
   DECLARE STATIC SUB SetButtonElevationRequiredState (BYVAL hTaskDlg AS HWND, BYVAL nButton AS LONG, BYVAL fElevation AS BOOLEAN)
   DECLARE STATIC SUB SetElementText (BYVAL hTaskDlg AS HWND, BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   DECLARE STATIC SUB SetMarqueeProgressBar (BYVAL hTaskDlg AS HWND, BYVAL fMarquee AS BOOLEAN)
   DECLARE STATIC SUB SetProgressBarMarquee (BYVAL hTaskDlg AS HWND, BYVAL fMarquee AS LONG, BYVAL nSpeed AS LONG)
   DECLARE STATIC FUNCTION SetProgressBarPos (BYVAL hTaskDlg AS HWND, BYVAL nNewPos AS LONG) AS LONG
   DECLARE STATIC FUNCTION SetProgressBarRange (BYVAL hTaskDlg AS HWND, BYVAL nMinRange AS SHORT, BYVAL nMaxRange AS SHORT) AS DWORD
   DECLARE STATIC FUNCTION SetProgressBarState (BYVAL hTaskDlg AS HWND, BYVAL nNewState AS LONG) AS LONG
   DECLARE STATIC SUB UpdateElementText (BYVAL hTaskDlg AS HWND, BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   DECLARE STATIC SUB UpdateIcon (BYVAL hTaskDlg AS HWND, BYVAL nIconElement AS LONG, BYVAL idResource AS DWORD)

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CTaskDialog
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CTaskDialog
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Simulates the action of a button click in the Task Dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.ClickButton (BYVAL hTaskDlg AS HWND, BYVAL nButton AS DWORD)
   SendMessageW(hTaskDlg, TDM_CLICK_BUTTON, nButton, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Simulates the action of a radio button click in the TaskDialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.ClickRadioButton (BYVAL hTaskDlg AS HWND, BYVAL nRadioButton AS DWORD)
   SendMessageW(hTaskDlg, TDM_CLICK_RADIO_BUTTON, nRadioButton, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Simulates a click on the verification checkbox of the Task Dialog, if it exists.
' ========================================================================================
PRIVATE SUB CTaskDialog.ClickVerification (BYVAL hTaskDlg AS HWND, BYVAL fChecked AS BOOLEAN, BYVAL fFocused AS BOOLEAN)
   SendMessageW(hTaskDlg, TDM_CLICK_VERIFICATION, fChecked, fFocused)
END SUB
' ========================================================================================

' ========================================================================================
' Enables or disables a push button in the TaskDialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.EnableButton (BYVAL hTaskDlg AS HWND, BYVAL nButton AS LONG)
   SendMessageW(hTaskDlg, TDM_ENABLE_BUTTON, nButton, CTRUE)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CTaskDialog.DisableButton (BYVAL hTaskDlg AS HWND, BYVAL nButton AS LONG)
   SendMessageW(hTaskDlg, TDM_ENABLE_BUTTON, nButton, FALsE)
END SUB
' ========================================================================================

' ========================================================================================
' Enables or disables a radio button in the TaskDialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.EnableRadioButton (BYVAL hTaskDlg AS HWND, BYVAL nRadioButton AS LONG)
   SendMessageW(hTaskDlg, TDM_ENABLE_RADIO_BUTTON, nRadioButton, CTRUE)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CTaskDialog.DisableRadioButton (BYVAL hTaskDlg AS HWND, BYVAL nRadioButton AS LONG)
   SendMessageW(hTaskDlg, TDM_ENABLE_RADIO_BUTTON, nRadioButton, FALSE)
END SUB
' ========================================================================================

' ========================================================================================
' Refreshes the contents of the Task Dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.NavigatePage (BYVAL hTaskDlg AS HWND, BYVAL pTaskConfig AS TASKDIALOGCONFIG PTR)
   IF pTaskConfig = NULL THEN EXIT SUB
   pTaskConfig->cbSize = SIZEOF(TASKDIALOGCONFIG)
   SendMessageW(hTaskDlg, TDM_NAVIGATE_PAGE, 0, CAST(LPARAM, pTaskConfig))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CTaskDialog.NavigatePage (BYVAL hTaskDlg AS HWND, BYREF TaskConfig AS TASKDIALOGCONFIG)
   IF VARPTR(TaskConfig) = 0 THEN EXIT SUB
   TaskConfig.cbSize = SIZEOF(TASKDIALOGCONFIG)
   SendMessageW(hTaskDlg, TDM_NAVIGATE_PAGE, 0, CAST(LPARAM, @TaskConfig))
END SUB
' ========================================================================================

' ========================================================================================
' Used by clients to designate whether a given Task Dialog button or command link should
' have a User Account Control (UAC) shield icon (in other words, whether the action
' invoked by the button requires elevation).
' ========================================================================================
PRIVATE SUB CTaskDialog.SetButtonElevationRequiredState (BYVAL hTaskDlg AS HWND, BYVAL nButton AS LONG, BYVAL fElevation AS BOOLEAN)
   SendMessageW(hTaskDlg, TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE, nButton, fElevation)
END SUB
' ========================================================================================

' ========================================================================================
' Used by clients to update a text element on the Task Dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.SetElementText (BYVAL hTaskDlg AS HWND, BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   SendMessageW(hTaskDlg, TDM_SET_ELEMENT_TEXT, eElement, CAST(LPARAM, pszNewText))
END SUB
' ========================================================================================

' ========================================================================================
' Indicates whether the hosted progress bar should be displayed in marquee mode.
' ========================================================================================
PRIVATE SUB CTaskDialog.SetMarqueeProgressBar (BYVAL hTaskDlg AS HWND, BYVAL fMarquee AS BOOLEAN)
   SendMessageW(hTaskDlg, TDM_SET_MARQUEE_PROGRESS_BAR, cast(WPARAM, fMarquee), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Starts and stops the marquee display of the progress bar, and sets the speed of the marquee.
' ========================================================================================
PRIVATE SUB CTaskDialog.SetProgressBarMarquee (BYVAL hTaskDlg AS HWND, BYVAL fMarquee AS LONG, BYVAL nSpeed AS LONG)
   SendMessageW(hTaskDlg, TDM_SET_PROGRESS_BAR_MARQUEE, fMarquee, nSpeed)
END SUB
' ========================================================================================

' ========================================================================================
' The TDM_SET_PROGRESS_BAR_POS message is used set the current position for a progress bar.
' Returns the previous position.
' ========================================================================================
PRIVATE FUNCTION CTaskDialog.SetProgressBarPos (BYVAL hTaskDlg AS HWND, BYVAL nNewPos AS LONG) AS LONG
   FUNCTION = SendMessageW(hTaskDlg, TDM_SET_PROGRESS_BAR_POS, nNewPos, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The TDM_SET_PROGRESS_BAR_RANGE message is used set the minimum and maximum values for
' the hosted progress bar. Returns the previous minimum and maximum values, if successful,
' or zero otherwise. The LOWORD contains the minimum value, and the HIWORD contains the
' maximum value.
' ========================================================================================
PRIVATE FUNCTION CTaskDialog.SetProgressBarRange (BYVAL hTaskDlg AS HWND, BYVAL nMinRange AS SHORT, BYVAL nMaxRange AS SHORT) AS DWORD
   FUNCTION = SendMessageW(hTaskDlg, TDM_SET_PROGRESS_BAR_RANGE, 0, MAKELONG(nMinRange, nMaxRange))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The TDM_SET_PROGRESS_BAR_POS message is used set the current position for a progress bar.
' If the function succeeds, the return value is non zero. If the function fails, the return
' value is zero. To get extended error information call GetLastError.
' ========================================================================================
PRIVATE FUNCTION CTaskDialog.SetProgressBarState (BYVAL hTaskDlg AS HWND, BYVAL nNewState AS LONG) AS LONG
   FUNCTION = SendMessageW(hTaskDlg, TDM_SET_PROGRESS_BAR_STATE, nNewState, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' The TDM_UPDATE_ELEMENT_TEXT message is available for use by clients to update a text
' element on the Task Dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.UpdateElementText (BYVAL hTaskDlg AS HWND, BYVAL eElement AS LONG, BYVAL pszNewText AS WSTRING PTR)
   SendMessageW(hTaskDlg, TDM_UPDATE_ELEMENT_TEXT, eElement, CAST(LPARAM, pszNewText))
END SUB
' ========================================================================================

' ========================================================================================
' Refreshes the icon of a task dialog.
' ========================================================================================
PRIVATE SUB CTaskDialog.UpdateIcon (BYVAL hTaskDlg AS HWND, BYVAL nIconElement AS LONG, BYVAL idResource AS DWORD)
   SendMessageW(hTaskDlg, TDM_UPDATE_ICON, nIconElement, idResource)
END SUB
' ========================================================================================

END NAMESPACE

#endif
