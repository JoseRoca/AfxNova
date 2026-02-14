' ########################################################################################
' Platform: Microsoft Windows
' Filename: CControls.inc
' Purpose:  Windows controls
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
#include once "win/commctrl.bi"
#include once "win/Windowsx.bi"
' // include files for Windows standard controls
#include once "CButton.inc"
#include once "CCombo.inc"
#include once "CEdit.inc"
#include once "CHeader.inc"
#include once "CListBox.inc"
#include once "CListView.inc"
#include once "CMonthCal.inc"
#include once "CProgressBar.inc"
#include once "CRebar.inc"
#include once "CStatic.inc"
#include once "CStatusBar.inc"
#include once "CTab.inc"
#include once "CTaskDialog.inc"
#include once "CToolBar.inc"
#include once "CToolTip.inc"
#include once "CTrackBar.inc"
#include once "CTreeView.inc"
#include once "CUpDown.inc"
USING AfxNova

NAMESPACE AfxNova

' ########################################################################################
'                                    *** ANIMATE ***
' ########################################################################################
TYPE CAnimate

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC FUNCTION Create (BYVAL hParent AS HWND, BYVAL id AS LONG_PTR, BYVAL dwStyle AS DWORD, BYVAL hInstance AS HINSTANCE) AS HWND
   DECLARE STATIC FUNCTION Open (BYVAL hAnimate AS HWND, BYVAL pwzName AS WSTRING PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION OpenEx (BYVAL hAnimate AS HWND, BYVAL hInstance AS HINSTANCE, BYVAL pwzName AS WSTRING PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION Play (BYVAL hAnimate AS HWND, BYVAL from AS UINT, BYVAL to AS UINT, BYVAL rep AS UINT) AS BOOLEAN
   DECLARE STATIC FUNCTION Stop (BYVAL hAnimate AS HWND) AS BOOLEAN
   DECLARE STATIC FUNCTION IsPlaying (BYVAL hAnimate AS HWND) AS BOOLEAN
   DECLARE STATIC FUNCTION Close (BYVAL hAnimate AS HWND) AS BOOLEAN
   DECLARE STATIC FUNCTION Seek (BYVAL hAnimate AS HWND, BYVAL frame AS UINT) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CAnimate
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CAnimate
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates an animation control. Animate_Create calls the CreateWindow function to create
' the animation control.
' ========================================================================================
PRIVATE FUNCTION CAnimate.Create (BYVAL hParent AS HWND, BYVAL id AS LONG_PTR, BYVAL dwStyle AS DWORD, BYVAL hInstance AS HINSTANCE) AS HWND
   RETURN CreateWindowW("SysAnimate32", NULL, dwStyle, 0, 0, 0, 0, hParent, cast(HMENU, id), hInstance, NULL)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Opens an AVI clip and displays its first frame in an animation control. You can use this
' macro or send the ACM_OPEN message explicitly.
' ========================================================================================
PRIVATE FUNCTION CAnimate.Open (BYVAL hAnimate AS HWND, BYVAL pwzName AS WSTRING PTR) AS BOOLEAN
   RETURN SendMessageW(hAnimate, ACM_OPEN, 0, cast(LPARAM, pwzName))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Opens an AVI clip and displays its first frame in an animation control. You can use this
' macro or send the ACM_OPEN message explicitly.
' ========================================================================================
PRIVATE FUNCTION CAnimate.OpenEx (BYVAL hAnimate AS HWND, BYVAL hInstance AS HINSTANCE, BYVAL pwzName AS WSTRING PTR) AS BOOLEAN
   RETURN SendMessageW(hAnimate, ACM_OPEN, cast(WPARAM, hInstance), cast(LPARAM, pwzName))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Plays an AVI clip in an animation control. The control plays the clip in the background
' while the thread continues executing. You can use this macro or send the ACM_PLAY message
' explicitly.
' ========================================================================================
PRIVATE FUNCTION CAnimate.Play (BYVAL hAnimate AS HWND, BYVAL from AS UINT, BYVAL _to AS UINT, BYVAL rep AS UINT) AS BOOLEAN
   RETURN SendMessageW(hAnimate, ACM_PLAY, rep, cast(LPARAM, MAKELONG(from, _to)))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Stops playing an AVI clip in an animation control. You can use this macro or send the
' ACM_STOP message explicitly.
' ========================================================================================
PRIVATE FUNCTION CAnimate.Stop (BYVAL hAnimate AS HWND) AS BOOLEAN
   RETURN SendMessageW(hAnimate, ACM_STOP, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Checks to see if an Audio-Video Interleaved (AVI) clip is playing. You can use this macro
' or send an ACM_ISPLAYING message.
' ========================================================================================
PRIVATE FUNCTION CAnimate.IsPlaying (BYVAL hAnimate AS HWND) AS BOOLEAN
   RETURN SendMessageW(hAnimate, ACM_ISPLAYING, 0, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Closes an AVI clip. You can use this macro or send the ACM_OPEN message explicitly,
' passing in NULL parameters.
' ========================================================================================
PRIVATE FUNCTION CAnimate.Close (BYVAL hAnimate AS HWND) AS BOOLEAN
   RETURN SendMessageW(hAnimate, ACM_OPEN, 0, NULL)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Directs an animation control to display a particular frame of an AVI clip. The control
' displays the clip in the background while the thread continues executing. You can use
' this macro or send the ACM_PLAY message explicitly.
' ========================================================================================
PRIVATE FUNCTION CAnimate.Seek (BYVAL hAnimate AS HWND, BYVAL frame AS UINT) AS BOOLEAN
   RETURN SendMessageW(hAnimate, ACM_PLAY, 1, cast(LPARAM, MAKELONG(frame, frame)))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                    *** HOTKEY ***
' ########################################################################################
TYPE CHotKey

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC FUNCTION GetHotKey (BYVAL hHotKey AS HWND) AS DWORD
   DECLARE STATIC SUB SetHotKey (BYVAL hHotKey AS HWND, BYVAL vkCode AS SHORT, BYVAL fMods AS SHORT)
   DECLARE STATIC SUB SetRules (BYVAL hHotKey AS HWND, BYVAL fwCombInv AS WORD, BYVAL fwModInv AS WORD)

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CHotKey
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CHotKey
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Gets the virtual key code and modifier flags of a hot key from a hot key control.
' Returns the virtual key code and modifier flags. The LOBYTE of the LOWORD is the virtual
' key code of the hot key. The HIBYTE of the LOWORD is the key modifier that specifies the
' keys that define a hot key combination. The modifier flags can be a combination of the
' following values:
'   HOTKEYF_ALT      ALT key
'   HOTKEYF_CONTROL  CONTROL key
'   HOTKEYF_EXT      Extended key
'   HOTKEYF_SHIFT    SHIFT key
' ========================================================================================
PRIVATE FUNCTION CHotKey.GetHotKey (BYVAL hHotKey AS HWND) AS DWORD
   RETURN SendMessageW(hHotKey, HKM_GETHOTKEY, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the hot key combination for a hot key control.
' - vkCode: The virtual key code of the hot key.
' - fMods: The key modifier that indicates the keys that define a hot key combination.
' For a list of modifier flag values, see the description of the HotKey_GetHotKey function.
' ========================================================================================
PRIVATE SUB CHotKey.SetHotKey (BYVAL hHotKey AS HWND, BYVAL vkCode AS SHORT, BYVAL fMods AS SHORT)
   SendMessageW(hHotKey, HKM_SETHOTKEY, MAKELONG(vkCode, fMods), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Defines the invalid combinations and the default modifier combination for a hot key control.
' - fwCombInv: An array of flags that specify invalid key combinations. This parameter can be a combination
'   of the following values:
' Value Meaning
'   HKCOMB_A ALT
'   HKCOMB_C CTRL
'   HKCOMB_CA CTRL+ALT
'   HKCOMB_NONE Unmodified keys
'   HKCOMB_S SHIFT
'   HKCOMB_SA SHIFT+ALT
'   HKCOMB_SC SHIFT+CTRL
'   HKCOMB_SCA SHIFT+CTRL+ALT
' - fwModInv: An array of flags that specify the key combination to use when the user enters
'   an invalid combination. For a list of modifier flag values, see the description of the
'   HotKey_GetHotKey message.
' ========================================================================================
PRIVATE SUB CHotKey.SetRules (BYVAL hHotKey AS HWND, BYVAL fwCombInv AS WORD, BYVAL fwModInv AS WORD)
   SendMessageW(hHotKey, HKM_SETRULES, cast(WPARAM, fwCombInv), cast(LPARAM, MAKELONG(fwModInv, 0)))
END SUB
' ========================================================================================


' ########################################################################################
'                                   *** IP ADDRESS ***
' ########################################################################################
TYPE CIPAddress

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC SUB ClearAddress (BYVAL hIPM AS HWND)
   DECLARE STATIC FUNCTION GetAddress (BYVAL hIPM AS HWND, BYVAL pdwAddr AS DWORD PTR) AS LONG
   DECLARE STATIC SUB IsBlank (BYVAL hIPM AS HWND)
   DECLARE STATIC SUB SetAddress (BYVAL hIPM AS HWND, BYVAL dwAddr AS DWORD)
   DECLARE STATIC SUB SetFocus (BYVAL hIPM AS HWND, BYVAL nField AS LONG)
   DECLARE STATIC FUNCTION SetRange (BYVAL hIPM AS HWND, BYVAL nField AS LONG, BYVAL loRange AS BYTE, BYVAL hiRange AS BYTE) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CIPAddress
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CIPAddress
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Clears the contents of the IP address control.
' ========================================================================================
PRIVATE SUB CIPAddress.ClearAddress (BYVAL hIPM AS HWND)
   SendMessageW(hIPM, IPM_CLEARADDRESS, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the address values for all four fields in the IP address control.
' Returns the number of nonblank fields.
' ========================================================================================
PRIVATE FUNCTION CIPAddress.GetAddress (BYVAL hIPM AS HWND, BYVAL pdwAddr AS DWORD PTR) AS LONG
   RETURN SendMessageW(hIPM, IPM_GETADDRESS, 0, CAST(LPARAM, pdwAddr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if all fields in the IP address control are blank.
' ========================================================================================
PRIVATE SUB CIPAddress.IsBlank (BYVAL hIPM AS HWND)
   SendMessageW(hIPM, IPM_ISBLANK, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the address values for all four fields in the IP address control.
' ========================================================================================
PRIVATE SUB CIPAddress.SetAddress (BYVAL hIPM AS HWND, BYVAL dwAddr AS DWORD)
   SendMessageW(hIPM, IPM_SETADDRESS, 0, cast(LPARAM, dwAddr))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the keyboard focus to the specified field in the IP address control.
' All of the text in that field will be selected.
' ========================================================================================
PRIVATE SUB CIPAddress.SetFocus (BYVAL hIPM AS HWND, BYVAL nField AS LONG)
   SendMessageW(hIPM, IPM_SETFOCUS, cast(WPARAM, nField), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the valid range for the specified field in the IP address control.
' Returns TRUE if successful, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION CIPAddress.SetRange (BYVAL hIPM AS HWND, BYVAL nField AS LONG, BYVAL loRange AS BYTE, BYVAL hiRange AS BYTE) AS BOOLEAN
   RETURN SendMessageW(hIPM, IPM_SETRANGE, cast(WPARAM, nField), cast(LPARAM, MAKEIPRANGE(loRange, hiRange)))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                   *** SYSLINK ***
' ########################################################################################
TYPE CSysLink

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC FUNCTION GetIdealHeight (BYVAL hLM AS HWND) AS LONG
   DECLARE STATIC FUNCTION GetIdealSize (BYVAL hLM AS HWND, BYVAL cxMaxWidth AS LONG, BYVAL psize AS SIZE PTR) AS LONG
   DECLARE STATIC FUNCTION GetIdealSize (BYVAL hLM AS HWND, BYVAL cxMaxWidth AS LONG, BYREF sz AS SIZE) AS LONG
   DECLARE STATIC FUNCTION GetIdealSize (BYVAL hLM AS HWND, BYVAL cxMaxWidth AS LONG) AS SIZE
   DECLARE STATIC FUNCTION GetItem (BYVAL hLM AS HWND, BYVAL pItem AS LITEM PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION GetItem (BYVAL hLM AS HWND, BYREF item AS LITEM) AS BOOLEAN
   DECLARE STATIC FUNCTION HitTest (BYVAL hLM AS HWND, BYVAL phti AS LHITTESTINFO PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION HitTest (BYVAL hLM AS HWND, BYREF hti AS LHITTESTINFO) AS BOOLEAN
   DECLARE STATIC FUNCTION SetItem (BYVAL hLM AS HWND, BYVAL pItem AS LITEM PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION SetItem (BYVAL hLM AS HWND, BYREF item AS LITEM) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CSysLink
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CSysLink
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Retrieves the preferred height of a link for the control's current width.
' ========================================================================================
PRIVATE FUNCTION CSysLink.GetIdealHeight (BYVAL hLM AS HWND) AS LONG
   RETURN SendMessageW(hLM, LM_GETIDEALHEIGHT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the preferred height of a link for the control's current width.
' Note: It is the same message that LM_GETIDEALHEIGHT with sifferent values.
' ========================================================================================
PRIVATE FUNCTION CSysLink.GetIdealSize (BYVAL hLM AS HWND, BYVAL cxMaxWidth AS LONG, BYVAL psize AS SIZE PTR) AS LONG
   RETURN SendMessageW(hLM, LM_GETIDEALHEIGHT, cxMaxWidth, cast(LPARAM, psize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSysLink.GetIdealSize (BYVAL hLM AS HWND, BYVAL cxMaxWidth AS LONG, BYREF sz AS SIZE) AS LONG
   RETURN SendMessageW(hLM, LM_GETIDEALHEIGHT, cxMaxWidth, cast(LPARAM, @sz))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSysLink.GetIdealSize (BYVAL hLM AS HWND, BYVAL cxMaxWidth AS LONG) AS SIZE
   DIM sz AS SIZE
   SendMessageW(hLM, LM_GETIDEALHEIGHT, cxMaxWidth, cast(LPARAM, @sz))
   RETURN sz
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the states and attributes of an item.
' Returns TRUE if the message succeeds in getting the values and attributes specified.
' ========================================================================================
PRIVATE FUNCTION CSyslink.GetItem (BYVAL hLM AS HWND, BYVAL pItem AS LITEM PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hLM, LM_GETITEM, 0, cast(LPARAM, pItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSyslink.GetItem (BYVAL hLM AS HWND, BYREF item AS LITEM) AS BOOLEAN
   FUNCTION = SendMessageW(hLM, LM_GETITEM, 0, cast(LPARAM, @item))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if the user clicked the specified link.
' Returns TRUE if user clicked on a link, otherwise returns FALSE.
' ========================================================================================
PRIVATE FUNCTION CSysLink.HitTest (BYVAL hLM AS HWND, BYVAL phti AS LHITTESTINFO PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hLM, LM_HITTEST, 0, cast(LPARAM, phti))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSysLink.HitTest (BYVAL hLM AS HWND, BYREF hti AS LHITTESTINFO) AS BOOLEAN
   FUNCTION = SendMessageW(hLM, LM_HITTEST, 0, cast(LPARAM, @hti))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the states and attributes of an item.
' Returns TRUE if the message succeeds in setting the values and attributes specified.
' ========================================================================================
PRIVATE FUNCTION CSysLink.SetItem (BYVAL hLM AS HWND, BYVAL pItem AS LITEM PTR) AS BOOLEAN
   FUNCTION = SendMessageW(hLM, LM_SETITEM, 0, cast(LPARAM, pItem))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CSysLink.SetItem (BYVAL hLM AS HWND, BYREF item AS LITEM) AS BOOLEAN
   FUNCTION = SendMessageW(hLM, LM_SETITEM, 0, cast(LPARAM, @item))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                     *** PAGER ***
' ########################################################################################
TYPE CPager

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC SUB ForwardMouse (BYVAL hPager AS HWND, BYVAL bForward AS BOOLEAN)
   DECLARE STATIC FUNCTION GetBkColor (BYVAL hPager AS HWND) AS COLORREF
   DECLARE STATIC FUNCTION GetBorder (BYVAL hPager AS HWND) AS LONG
   DECLARE STATIC FUNCTION GetButtonSize (BYVAL hPager AS HWND) AS LONG
   DECLARE STATIC FUNCTION GetButtonState (BYVAL hPager AS HWND, BYVAL iButton AS LONG) AS DWORD
   DECLARE STATIC SUB GetDropTarget (BYVAL hPager AS HWND, BYVAL ppdt AS IDropTarget PTR PTR)
   DECLARE STATIC FUNCTION GetDropTarget (BYVAL hPager AS HWND) AS IDropTarget PTR
   DECLARE STATIC FUNCTION GetPos (BYVAL hPager AS HWND) AS LONG
   DECLARE STATIC SUB RecalcSize (BYVAL hPager AS HWND)
   DECLARE STATIC FUNCTION SetBkColor (BYVAL hPager AS HWND, BYVAL clr AS COLORREF) AS COLORREF
   DECLARE STATIC FUNCTION SetBorder (BYVAL hPager AS HWND, BYVAL iBorder AS LONG) AS LONG
   DECLARE STATIC FUNCTION SetButtonSize (BYVAL hPager AS HWND, BYVAL iSize AS LONG) AS LONG
   DECLARE STATIC SUB SetChild (BYVAL hPager AS HWND, BYVAL hwndChild AS HWND)
   DECLARE STATIC FUNCTION SetPos (BYVAL hPager AS HWND, BYVAL iPos AS LONG) AS LONG

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CPager
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CPager
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Enables or disables mouse forwarding for the pager control. When mouse forwarding is
' enabled, the pager control forwards WM_MOUSEMOVE messages to the contained window. You
' can use this macro or send the PGM_FORWARDMOUSE message explicitly.
' ========================================================================================
PRIVATE SUB CPager.ForwardMouse (BYVAL hPager AS HWND, BYVAL bForward AS BOOLEAN)
   SendMessageW(hPager, PGM_FORWARDMOUSE, cast(WPARAM, bForWard), 0)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the current background color for the pager control. You can use this macro or
' send the PGM_GETBKCOLOR message explicitly.
' ========================================================================================
PRIVATE FUNCTION CPager.GetBkColor (BYVAL hPager AS HWND) AS COLORREF
   RETURN SendMessageW(hPager, PGM_GETBKCOLOR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current border size for the pager control. You can use this macro or send
' the PGM_GETBORDER message explicitly.
' ========================================================================================
PRIVATE FUNCTION CPager.GetBorder (BYVAL hPager AS HWND) AS LONG
   RETURN SendMessageW(hPager, PGM_GETBORDER, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current button size for the pager control. You can use this macro or send
' the PGM_GETBUTTONSIZE message explicitly.
' ========================================================================================
PRIVATE FUNCTION CPager.GetButtonSize (BYVAL hPager AS HWND) AS LONG
   RETURN SendMessageW(hPager, PGM_GETBUTTONSIZE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the state of the specified button in a pager control. You can use this macro
' or send the PGM_GETBUTTONSTATE message explicitly.
' ========================================================================================
PRIVATE FUNCTION CPager.GetButtonState (BYVAL hPager AS HWND, BYVAL iButton AS LONG) AS DWORD
   RETURN SendMessageW(hPager, PGM_GETBUTTONSTATE, 0, iButton)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a pager control's IDropTarget interface pointer. You can use this macro or send
' the PGM_GETDROPTARGET message explicitly.
' ========================================================================================
PRIVATE SUB CPager.GetDropTarget (BYVAL hPager AS HWND, BYVAL ppdt AS IDropTarget PTR PTR)
   SendMessageW(hPager, PGM_GETDROPTARGET, 0, cast(LPARAM, ppdt))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CPager.GetDropTarget (BYVAL hPager AS HWND) AS IDropTarget PTR
   DIM pdt AS IDropTarget PTR
   SendMessageW(hPager, PGM_GETDROPTARGET, 0, cast(LPARAM, @pdt))
   RETURN pdt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current scroll position of the pager control. You can use this macro or
' send the PGM_GETPOS message explicitly.
' ========================================================================================
PRIVATE FUNCTION CPager.GetPos (BYVAL hPager AS HWND) AS LONG
   RETURN SendMessageW(hPager, PGM_GETPOS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Forces the pager control to recalculate the size of the contained window. Using this
' method will result in a PGN_CALCSIZE notification being sent. You can use this method or
' send the PGM_RECALCSIZE message explicitly.
' ========================================================================================
PRIVATE SUB CPager.RecalcSize (BYVAL hPager AS HWND)
   SendMessageW(hPager, PGM_RECALCSIZE, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the current background color for the pager control. You can use this macro or send
' the PGM_SETBKCOLOR message explicitly.
' ========================================================================================
PRIVATE FUNCTION CPager.SetBkColor (BYVAL hPager AS HWND, BYVAL clr AS COLORREF) AS COLORREF
   RETURN SendMessageW(hPager, PGM_SETBKCOLOR, 0, cast(LPARAM, clr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current border size for the pager control. You can use this macro or send the
' PGM_SETBORDER message explicitly.
' ========================================================================================
PRIVATE FUNCTION CPager.SetBorder (BYVAL hPager AS HWND, BYVAL iBorder AS LONG) AS LONG
   RETURN SendMessageW(hPager, PGM_SETBORDER, 0, cast(LPARAM, iBorder))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current button size for the pager control. You can use this macro or send the
' PGM_SETBUTTONSIZE message explicitly.
' ========================================================================================
PRIVATE FUNCTION CPager.SetButtonSize (BYVAL hPager AS HWND, BYVAL iSize AS LONG) AS LONG
   RETURN SendMessageW(hPager, PGM_SETBUTTONSIZE, 0, cast(LPARAM, iSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the contained window for the pager control. This macro will not change the parent
' of the contained window; it only assigns a window handle to the pager control for scrolling.
' In most cases, the contained window will be a child window. If this is the case, the
' contained window should be a child of the pager control. You can use this macro or send
' the PGM_SETCHILD message explicitly.
' ========================================================================================
PRIVATE SUB CPager.SetChild (BYVAL hPager AS HWND, BYVAL hwndChild AS HWND)
   SendMessageW(hPager, PGM_SETCHILD, 0, cast(LPARAM, hwndChild))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the scroll position for the pager control. You can use this macro or send the
' PGM_SETPOS message explicitly.
' ========================================================================================
PRIVATE FUNCTION CPager.SetPos (BYVAL hPager AS HWND, BYVAL iPos AS LONG) AS LONG
   RETURN SendMessageW(hPager, PGM_SETPOS, 0, cast(LPARAM, iPos))
END FUNCTION
' ========================================================================================


' ########################################################################################
'                                    *** SCROLLBAR ***
' ########################################################################################
TYPE CScrollBar

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC FUNCTION Enable (BYVAL hSB AS HWND, BYVAL flags AS UINT) AS BOOLEAN
   DECLARE STATIC FUNCTION Show (BYVAL hSB AS HWND, BYVAL fShow AS BOOLEAN) AS BOOLEAN
   DECLARE STATIC FUNCTION Show (BYVAL hSB AS HWND) AS BOOLEAN
   DECLARE STATIC FUNCTION Hide (BYVAL hSB AS HWND) AS BOOLEAN
   DECLARE STATIC FUNCTION SetPos (BYVAL hSB AS HWND, BYVAL pos AS LONG, BYVAL fRedraw AS BOOLEAN) AS BOOLEAN
   DECLARE STATIC FUNCTION GetPos (BYVAL hSB AS HWND) AS LONG
   DECLARE STATIC FUNCTION SetRange (BYVAL hSB AS HWND, BYVAL posMin AS LONG, BYVAL posMax AS LONG, BYVAL fedraw AS BOOLEAN) AS BOOLEAN
   DECLARE STATIC FUNCTION GetRange (BYVAL hSB AS HWND, BYVAL posMin AS LONG PTR, BYVAL posMax AS LONG PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION GetRange (BYVAL hSB AS HWND, BYREF posMin AS LONG, BYREF posMax AS LONG) AS BOOLEAN

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CScrollBar
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CScrollBar
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Enables or disables a scroll bar control.
' This method expands to a call to EnableScrollBar with SB_CTL in the wSBFlags parameter.
' ========================================================================================
PRIVATE FUNCTION CScrollBar.Enable (BYVAL hSB AS HWND, BYVAL flags AS UINT) AS BOOLEAN
   RETURN EnableScrollBar(hSB, SB_CTL, flags)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Shows or hides a scroll bar control.
' ========================================================================================
PRIVATE FUNCTION CScrollBar.Show (BYVAL hSB AS HWND, BYVAL fShow AS BOOLEAN) AS BOOLEAN
   RETURN ShowWindow(hSB, iif(fShow, SW_SHOWNORMAL, SW_HIDE))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CScrollBar.Show (BYVAL hSB AS HWND) AS BOOLEAN
   RETURN ShowWindow(hSB, SW_SHOWNORMAL)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CScrollBar.Hide (BYVAL hSB AS HWND) AS BOOLEAN
   RETURN ShowWindow(hSB, SW_HIDE)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the position of the scroll box (thumb) in the specified scroll bar and, if requested,
' redraws the scroll bar to reflect the new position of the scroll box.
' ========================================================================================
PRIVATE FUNCTION CScrollBar.SetPos (BYVAL hSB AS HWND, BYVAL pos AS LONG, BYVAL fRedraw AS BOOLEAN) AS BOOLEAN
   RETURN SetScrollPos(hSB, SB_CTL, pos, fRedraw)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the position of the scroll box (thumb) in the specified scroll bar.
' ========================================================================================
PRIVATE FUNCTION CScrollBar.GetPos (BYVAL hSB AS HWND) AS LONG
   RETURN GetScrollPos(hSB, SB_CTL)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the range of a scroll bar.
' ========================================================================================
PRIVATE FUNCTION CScrollBar.SetRange (BYVAL hSB AS HWND, BYVAL posMin AS LONG, BYVAL posMax AS LONG, BYVAL fRedraw AS BOOLEAN) AS BOOLEAN
   RETURN SetScrollRange(hSB, SB_CTL, posMin, posMax, fRedraw)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the range of a scroll bar.
' ========================================================================================
PRIVATE FUNCTION CScrollBar.GetRange (BYVAL hSB AS HWND, BYVAL posMin AS LONG PTR, BYVAL posMax AS LONG PTR) AS BOOLEAN
   RETURN GetScrollRange(hSB, SB_CTL, posMin, posMax)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CScrollBar.GetRange (BYVAL hSB AS HWND, BYREF posMin AS LONG, BYREF posMax AS LONG) AS BOOLEAN
   RETURN GetScrollRange(hSB, SB_CTL, @posMin, @posMax)
END FUNCTION
' ========================================================================================


#if _WIN32_WINNT = &h0602

' ########################################################################################
'                                    *** PROPSHEET ***
' ########################################################################################
TYPE CPropSheet

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC SUB SetNextText (BYVAL hDlg AS HWND, BYVAL pwszText AS WSTRING PTR)
   DECLARE STATIC SUB ShowWizButtons (BYVAL hDlg AS HWND, BYVAL dwFlag AS DWORD, BYVAL dwButton AS DWORD)
   DECLARE STATIC SUB EnableWizButtons (BYVAL hDlg AS HWND, BYVAL dwState AS DWORD, BYVAL dwMask AS DWORD)
   DECLARE STATIC SUB SetButtonText (BYVAL hDlg AS HWND, BYVAL dwButton AS DWORD, BYVAL pwszText AS WSTRING PTR)

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CPropSheet
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CPropSheet
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' PSM_SETNEXTTEXTW         = (WM_USER + 137)
' PSM_SETNEXTTEXT          = PSM_SETNEXTTEXTW
' ========================================================================================
PRIVATE SUB CPropSheet.SetNextText (BYVAL hDlg AS HWND, BYVAL pwszText AS WSTRING PTR)
   SendMessageW(hDlg, WM_USER + 137, 0, cast(LPARAM, pwszText))
END SUB
' ========================================================================================

' ========================================================================================
' PSWIZB_SHOW                    = 0
' PSWIZB_RESTORE                 = 1
' PSM_SHOWWIZBUTTONS             = (WM_USER + 138)
' ========================================================================================
PRIVATE SUB CPropSheet.ShowWizButtons (BYVAL hDlg AS HWND, BYVAL dwFlag AS DWORD, BYVAL dwButton AS DWORD)
   SendMessageW(hDlg, WM_USER + 138, cast(WPARAM, dwFlag), cast(LPARAM, dwButton))
END SUB
' ========================================================================================

' ========================================================================================
' PSM_ENABLEWIZBUTTONS           = (WM_USER + 139)
' ========================================================================================
PRIVATE SUB CPropSheet.EnableWizButtons (BYVAL hDlg AS HWND, BYVAL dwState AS DWORD, BYVAL dwMask AS DWORD)
   SendMessageW(hDlg, WM_USER + 139, cast(WPARAM, dwState), cast(LPARAM, dwMask))
END SUB
' ========================================================================================

' ========================================================================================
' PSM_SETBUTTONTEXTW         = (WM_USER + 140)
' PSM_SETBUTTONTEXT          = PSM_SETBUTTONTEXTW
' ========================================================================================
PRIVATE SUB CPropSheet.SetButtonText (BYVAL hDlg AS HWND, BYVAL dwButton AS DWORD, BYVAL pwszText AS WSTRING PTR)
   SendMessageW(hDlg, WM_USER + 140, cast(WPARAM, dwButton), cast(LPARAM, pwszText))
END SUB
' ========================================================================================

#endif

' ########################################################################################
'                                    *** DateTime Picker ***
' ########################################################################################
TYPE CDtPicker

Private:
   m_pData AS ANY PTR   ' Not used by the static functions, but we can't have an empty type.

Protected:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

Public:
   DECLARE STATIC FUNCTION GetSystemtime (BYVAL hdp AS HWND, BYVAL pst AS SYSTEMTIME PTR) AS DWORD
   DECLARE STATIC FUNCTION GetSystemtime (BYVAL hdp AS HWND, BYREF st AS SYSTEMTIME) AS DWORD
   DECLARE STATIC FUNCTION GetSystemtime (BYVAL hdp AS HWND) AS SYSTEMTIME
   DECLARE STATIC FUNCTION SetSystemtime (BYVAL hdp AS HWND, BYVAL gd AS DWORD, BYVAL pst AS SYSTEMTIME PTR) AS DWORD
   DECLARE STATIC FUNCTION SetSystemtime (BYVAL hdp AS HWND, BYVAL gd AS DWORD, BYREF st AS SYSTEMTIME) AS DWORD
   DECLARE STATIC FUNCTION GetRange (BYVAL hdp AS HWND, BYVAL rgst AS SYSTEMTIME PTR) AS DWORD
   DECLARE STATIC FUNCTION SetRange (BYVAL hdp AS HWND, BYVAL gd AS DWORD, BYVAL rgst AS SYSTEMTIME PTR) AS DWORD
   DECLARE STATIC FUNCTION SetFormat (BYVAL hdp AS HWND, BYVAL pwsz AS WSTRING PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION SetMonthCalColor (BYVAL hdp AS HWND, BYVAL iColor AS LONG, BYVAL clr AS COLORREF) AS COLORREF
   DECLARE STATIC FUNCTION GetMonthCalColor (BYVAL hdp AS HWND, BYVAL iColor AS LONG) AS COLORREF
   DECLARE STATIC FUNCTION GetMonthCal (BYVAL hdp AS HWND) AS HWND
   DECLARE STATIC FUNCTION GetMonthCalFont (BYVAL hdp AS HWND) AS HFONT
   DECLARE STATIC SUB SetMonthCalFont (BYVAL hdp AS HWND, BYVAL hfont AS HFONT, BYVAL fRedraw AS BOOLEAN)
#if _WIN32_WINNT >= &h0600
   DECLARE STATIC FUNCTION SetMonthCalStyle (BYVAL hdp AS HWND, BYVAL dwStyle AS DWORD) AS LRESULT
   DECLARE STATIC FUNCTION GetMonthCalStyle (BYVAL hdp AS HWND) AS LRESULT
   DECLARE STATIC SUB CloseMonthCal (BYVAL hdp AS HWND)
   DECLARE STATIC SUB GetDateTimePickerInfo (BYVAL hdp AS HWND, BYVAL pdtpi AS DATETIMEPICKERINFO PTR)
   DECLARE STATIC SUB GetDateTimePickerInfo (BYVAL hdp AS HWND, BYREF dtpi AS DATETIMEPICKERINFO)
   DECLARE STATIC FUNCTION GetDateTimePickerInfo (BYVAL hdp AS HWND) AS DATETIMEPICKERINFO
   DECLARE STATIC FUNCTION GetIdealSize (BYVAL hdp AS HWND, BYVAL psize AS SIZE PTR) AS BOOLEAN
   DECLARE STATIC FUNCTION GetIdealSize (BYVAL hdp AS HWND, BYREF sz AS SIZE) AS BOOLEAN
   DECLARE STATIC FUNCTION GetIdealSize (BYVAL hdp AS HWND) AS SIZE
#endif

END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR CDtPicker
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE DESTRUCTOR CDtPicker
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Gets the currently selected time from a date and time picker (DTP) control and places it
' in a specified SYSTEMTIME structure. You can use this macro, or send the DTM_GETSYSTEMTIME
' message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetSystemtime (BYVAL hdp AS HWND, BYVAL pst AS SYSTEMTIME PTR) AS DWORD
   RETURN SendMessageW(hdp, DTM_GETSYSTEMTIME, 0, cast(LPARAM, pst))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetSystemtime (BYVAL hdp AS HWND, BYREF st AS SYSTEMTIME) AS DWORD
   RETURN SendMessageW(hdp, DTM_GETSYSTEMTIME, 0, cast(LPARAM, @st))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetSystemtime (BYVAL hdp AS HWND) AS SYSTEMTIME
   DIM st AS SYSTEMTIME
   SendMessageW(hdp, DTM_GETSYSTEMTIME, 0, cast(LPARAM, @st))
   RETURN st
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets a date and time picker (DTP) control to a given date and time. You can use this
' macro or send the DTM_SETSYSTEMTIME message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.SetSystemtime (BYVAL hdp AS HWND, BYVAL gd AS DWORD, BYVAL pst AS SYSTEMTIME PTR) AS DWORD
   RETURN SendMessageW(hdp, DTM_SETSYSTEMTIME, gd, cast(LPARAM, pst))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDtPicker.SetSystemtime (BYVAL hdp AS HWND, BYVAL gd AS DWORD, BYREF st AS SYSTEMTIME) AS DWORD
   RETURN SendMessageW(hdp, DTM_SETSYSTEMTIME, gd, cast(LPARAM, @st))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current minimum and maximum allowable system times for a date and time picker (DTP)
' control. You can use this macro, or send the DTM_GETRANGE message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetRange (BYVAL hdp AS HWND, BYVAL rgst AS SYSTEMTIME PTR) AS DWORD
   RETURN SendMessageW(hdp, DTM_GETRANGE, 0, cast(LPARAM, rgst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the minimum and maximum allowable system times for a date and time picker (DTP) control.
' You can use this macro or send the DTM_SETRANGE message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.SetRange (BYVAL hdp AS HWND, BYVAL gd AS DWORD, BYVAL rgst AS SYSTEMTIME PTR) AS DWORD
   RETURN SendMessageW(hdp, DTM_SETRANGE, gd, cast(LPARAM, rgst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the display of a date and time picker (DTP) control based on a given format string.
' You can use this macro or send the DTM_SETFORMAT message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.SetFormat (BYVAL hdp AS HWND, BYVAL pwsz AS WSTRING PTR) AS BOOLEAN
   RETURN SendMessageW(hdp, DTM_SETFORMAT, 0, cast(LPARAM, pwsz))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the color for a given portion of the month calendar within a date and time picker (DTP)
' control. You can use this macro or send the DTM_SETMCCOLOR message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.SetMonthCalColor (BYVAL hdp AS HWND, BYVAL iColor AS LONG, BYVAL clr AS COLORREF) AS COLORREF
   RETURN SendMessageW(hdp, DTM_SETMCCOLOR, iColor, clr)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the color for a given portion of the month calendar within a date and time picker (DTP)
' control. You can use this macro or send the DTM_GETMCCOLOR message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetMonthCalColor (BYVAL hdp AS HWND, BYVAL iColor AS LONG) AS COLORREF
   RETURN SendMessageW(hdp, DTM_GETMCCOLOR, iColor, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the handle to a date and time picker's (DTP) child month calendar control. You can
' use this macro or send the DTM_GETMONTHCAL message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetMonthCal (BYVAL hdp AS HWND) AS HWND
   RETURN cast(HWND, SendMessageW(hdp, DTM_GETMONTHCAL, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the font that the date and time picker (DTP) control's child month calendar control
' is currently using. You can use this macro or send the DTM_GETMCFONT message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetMonthCalFont (BYVAL hdp AS HWND) AS HFONT
   RETURN cast(HFONT, SendMessageW(hdp, DTM_GETMCFONT, 0, 0))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font to be used by the date and time picker (DTP) control's child month calendar
' control. You can use this macro or explicitly send the DTM_SETMCFONT message.
' ========================================================================================
PRIVATE SUB CDtPicker.SetMonthCalFont (BYVAL hdp AS HWND, BYVAL hfont AS HFONT, BYVAL fRedraw AS BOOLEAN)
   SendMessageW(hdp, DTM_SETMCFONT, cast(WPARAM, hfont), cast(LPARAM, fRedraw))
END SUB
' ========================================================================================

#if _WIN32_WINNT >= &h0600

' ========================================================================================
' Sets the style for a specified date and time picker (DTP) control. Use this macro or send
' the DTM_SETMCSTYLE message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.SetMonthCalStyle (BYVAL hdp AS HWND, BYVAL dwStyle AS DWORD) AS LRESULT
   RETURN SendMessageW(hdp, DTM_SETMCSTYLE, 0, cast(LPARAM, dwStyle))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the style of a specified date and time picker (DTP) control. Use this macro or send
' the DTM_GETMCSTYLE message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetMonthCalStyle (BYVAL hdp AS HWND) AS LRESULT
   RETURN SendMessageW(hdp, DTM_GETMCSTYLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes the date and time picker (DTP) control. Use this macro or send the DTM_CLOSEMONTHCAL
' message explicitly.
' ========================================================================================
PRIVATE SUB CDtPicker.CloseMonthCal (BYVAL hdp AS HWND)
   SendMessageW(hdp, DTM_CLOSEMONTHCAL, 0, 0)
END SUB
' ========================================================================================

' ========================================================================================
' Gets information for a specified date and time picker (DTP) control.
' ========================================================================================
PRIVATE SUB CDtPicker.GetDateTimePickerInfo (BYVAL hdp AS HWND, BYVAL pdtpi AS DATETIMEPICKERINFO PTR)
   IF pdtpi = NULL THEN EXIT SUB
   pdtpi->cbSize = SIZEOF(DATETIMEPICKERINFO)
   SendMessageW(hdp, DTM_GETDATETIMEPICKERINFO, 0, cast(LPARAM, pdtpi))
END SUB
' ========================================================================================
PRIVATE SUB CDtPicker.GetDateTimePickerInfo (BYVAL hdp AS HWND, BYREF dtpi AS DATETIMEPICKERINFO)
   IF VARPTR(dtpi) = NULL THEN EXIT SUB
   dtpi.cbSize = SIZEOF(DATETIMEPICKERINFO)
   SendMessageW(hdp, DTM_GETDATETIMEPICKERINFO, 0, cast(LPARAM, @dtpi))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetDateTimePickerInfo (BYVAL hdp AS HWND) AS DATETIMEPICKERINFO
   DIM dtpi AS DATETIMEPICKERINFO
   dtpi.cbSize = SIZEOF(DATETIMEPICKERINFO)
   SendMessageW(hdp, DTM_GETDATETIMEPICKERINFO, 0, cast(LPARAM, @dtpi))
   RETURN dtpi
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the size needed to display the control without clipping. Use this macro or send the
' DTM_GETIDEALSIZE message explicitly.
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetIdealSize (BYVAL hdp AS HWND, BYVAL psize AS SIZE PTR) AS BOOLEAN
   RETURN SendMessageW(hdp, DTM_GETIDEALSIZE, 0, cast(LPARAM, psize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetIdealSize (BYVAL hdp AS HWND, BYREF sz AS SIZE) AS BOOLEAN
   RETURN SendMessageW(hdp, DTM_GETIDEALSIZE, 0, cast(LPARAM, @sz))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDtPicker.GetIdealSize (BYVAL hdp AS HWND) AS SIZE
   DIM sz AS SIZE
   SendMessageW(hdp, DTM_GETIDEALSIZE, 0, cast(LPARAM, @sz))
   RETURN sz
END FUNCTION
' ========================================================================================

#endif

END NAMESPACE
