' ########################################################################################
' Platform: Microsoft Windows
' File: CIFileDialogCustomize.inc
' Contents: Wrapper of the FileDialogCustomize interface.
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

#if _WIN32_WINNT >= &h0600

' ########################################################################################
' CIFileDialogCustomize class
' ########################################################################################

TYPE CIFileDialogCustomize

Private:

   m_IFileDialogCustomize AS IFileDialogCustomize PTR

Public:

   DECLARE CONSTRUCTOR (BYVAL pIFileDialog AS IFileDialog PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetPtr () AS IFileDialogCustomize PTR

   ' // IFileDialogCustomize methods
   DECLARE FUNCTION AddCheckButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR, BYVAL bChecked AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION AddComboBox (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION AddControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION AddEditBox (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION AddMenu (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION AddPushButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION AddRadioButtonList (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION AddSeparator (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION AddText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION EnableOpenDropDown (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION EndVisualGroup () AS HRESULT
   DECLARE FUNCTION GetCheckButtonState (BYVAL dwIDCtl AS DWORD) AS BOOLEAN
   DECLARE FUNCTION GetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYREF dwState AS CDCONTROLSTATEF) AS HRESULT
   DECLARE FUNCTION GetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS CDCONTROLSTATEF
   DECLARE FUNCTION GetControlState (BYVAL dwIDCtl AS DWORD) AS CDCONTROLSTATEF
   DECLARE FUNCTION GetEditBoxText (BYVAL dwIDCtl AS DWORD) AS DWSTRING
   DECLARE FUNCTION GetSelectedControlItem (BYVAL dwIDCtl AS DWORD) AS DWSTRING
   DECLARE FUNCTION MakeProminent (BYVAL dwIDCtl AS DWORD) AS HRESULT
'   DECLARE FUNCTION RemoveAllControlItems (BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION RemoveControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwItem AS DWORD) AS HRESULT
   DECLARE FUNCTION SetCheckButtonState (BYVAL dwIDCtl AS DWORD, BYVAL bCkecked AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION SetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   DECLARE FUNCTION SetControlItemText (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetControlLabel (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetControlState (BYVAL dwIDCtl AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   DECLARE FUNCTION SetEditBoxText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION SetSelectedControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   DECLARE FUNCTION StartVisualGroup (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT

END TYPE

' ########################################################################################

' ========================================================================================
' Creates an instance of the IFileDialogCustomize interface.
' ========================================================================================
CONSTRUCTOR CIFileDialogCustomize (BYVAL pIFileDialog AS IFileDialog PTR)
   OutputDebugStringW("CONSTRUCTOR CIFileDialogCustomize")
   CoInitialize NULL
   IF pIFileDialog = NULL THEN EXIT CONSTRUCTOR
   DIM hr AS HRESULT = pIFileDialog->lpvtbl->QueryInterface(pIFileDialog, @IID_IFileDialogCustomize, @m_IFileDialogCustomize)
   IF m_IFileDialogCustomize = NULL THEN EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Releases the IFileDialogCustomize interface
' ========================================================================================
PRIVATE DESTRUCTOR CIFileDialogCustomize
   OutputDebugStringW("DESTRUCTOR CIFileDialogCustomize")
   IF m_IFileDialogCustomize THEN  m_IFileDialogCustomize->lpvtbl->Release(m_IFileDialogCustomize)
   CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying IFileDialogCustomize interface.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.GetPtr () AS IFileDialogCustomize PTR
   RETURN m_IFileDialogCustomize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a check button (check box) to the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.AddCheckButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR, BYVAL bChecked AS BOOLEAN = FAlSE) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddCheckButton(m_IFileDialogCustomize, dwIDCtl, pwszLabel, bChecked)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a combo box to the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.AddComboBox (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddComboBox(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds an item to a container control in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.AddControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddControlItem(m_IFileDialogCustomize, dwIDCtl, dwIDItem, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds an edit box control to the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.AddEditBox (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddEditBox(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a menu to the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.AddMenu (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddMenu(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a push button to the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.AddPushButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddPushButton(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a push button to the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.AddRadioButtonList (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddRadioButtonList(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds a separator to the dialog, allowing a visual separation of controls.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.AddSeparator (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddSeparator(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds text content to the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.AddText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddText(m_IFileDialogCustomize, dwIDCtl, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables a drop-down list on the Open or Save button in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.EnableOpenDropDown (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->EnableOpenDropDown(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Stops the addition of elements to a visual group in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.EndVisualGroup () AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->EndVisualGroup(m_IFileDialogCustomize)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current state of a check button (check box) in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.GetCheckButtonState (BYVAL dwIDCtl AS DWORD) AS BOOLEAN
   DiM bChecked AS WINBOOL
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetCheckButtonState(m_IFileDialogCustomize, dwIDCtl, @bChecked))
   RETURN bChecked
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current state of an item in a container control found in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.GetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS CDCONTROLSTATEF
   DIM dwState AS CDCONTROLSTATEF
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetControlItemState(m_IFileDialogCustomize, dwIDCtl, dwIDItem, @dwState))
   RETURN dwState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current visibility and enabled states of a given control.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.GetControlState (BYVAL dwIDCtl AS DWORD) AS CDCONTROLSTATEF
   DIM dwState AS CDCONTROLSTATEF
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetControlState(m_IFileDialogCustomize, dwIDCtl, @dwState))
   RETURN dwState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current text in an edit box control.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.GetEditBoxText (BYVAL dwIDCtl AS DWORD) AS DWSTRING
   DIM pwszText AS WSTRING PTR
   SetLastError(m_IFileDialogCustomize->lpvtbl->GetEditBoxText(m_IFileDialogCustomize, dwIDCtl, @pwszText))
   RETURN pwszText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Places a control in the dialog so that it stands out compared to other added controls.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.MakeProminent (BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->MakeProminent(m_IFileDialogCustomize, dwIDCtl)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Not implemented.
' ========================================================================================
'PRIVATE FUNCTION CIFileDialogCustomize.RemoveAllControlItems (BYVAL dwIDCtl AS DWORD) AS HRESULT
'   RETURN m_IFileDialogCustomize->lpvtbl->RemoveAllControlItems(m_IFileDialogCustomize, dwIDCtl)
'END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes an item from a container control in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.RemoveControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwItem AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->RemoveControlItem(m_IFileDialogCustomize, dwIDCtl, dwItem)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of a check button (check box) in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.SetCheckButtonState (BYVAL dwIDCtl AS DWORD, BYVAL bCkecked AS BOOLEAN) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetCheckButtonState(m_IFileDialogCustomize, dwIDCtl, bCkecked)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the state of a check button (check box) in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.SetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlItemState(m_IFileDialogCustomize, dwIDCtl, dwIDItem, dwState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of a control item. For example, the text that accompanies a radio button or an item in a menu.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.SetControlItemText (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlItemText(m_IFileDialogCustomize, dwIDCtl, dwIDItem, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text associated with a control, such as button text or an edit box label.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.SetControlLabel (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlLabel(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current visibility and enabled states of a given control.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.SetControlState (BYVAL dwIDCtl AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetControlState(m_IFileDialogCustomize, dwIDCtl, dwState)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text in an edit box control found in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.SetEditBoxText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetEditBoxText(m_IFileDialogCustomize, dwIDCtl, pwszText)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the selected state of a particular item in an option button group or a combo box found in the dialog.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.SetSelectedControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->SetSelectedControlItem(m_IFileDialogCustomize, dwIDCtl, dwIDItem)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Declares a visual group in the dialog. Subsequent calls to any "add" method add those elements to this group
' ========================================================================================
PRIVATE FUNCTION CIFileDialogCustomize.StartVisualGroup (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
   RETURN m_IFileDialogCustomize->lpvtbl->AddText(m_IFileDialogCustomize, dwIDCtl, pwszLabel)
END FUNCTION
' ========================================================================================

' ########################################################################################

#endif
