# CIFileDialogCustomize Class

Wraps the `IFileDialogCustomize`interface, thet exposes methods that allow an application to add controls to a common file dialog.

Include file: AfxNova/CIFileDialogCustomize.inc

---

## Constructor

Creates instances de the `CIFileDialogCustomize` class.

```
CONSTRUCTOR (BYVAL pIFileDialog AS IFileDialog PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pIFileDialog* | A pointer to the **IFileDialog** interface. |

---

## Methods

| Name       | Description |
| ---------- | ----------- |
| [AddCheckButton](#addcheckbutton) | Adds a check button (check box) to the dialog. |
| [AddComboBox](#addcombobox) | Adds a combo box to the dialog. |
| [AddControlItem](#addcontrolitem) | Adds an item to a container control in the dialog. |
| [AddEditBox](#addeditbox) | Adds an edit box control to the dialog. |
| [AddMenu](#addmenu) | Adds a menu to the dialog. |
| [AddPushButton](#addpushbutton) | Adds a button to the dialog. |
| [AddRadioButtonList](#addradiobuttonlist) | Adds an option button (also known as radio button) group to the dialog. |
| [AddSeparator](#addseparatgor) | Adds a separator to the dialog, allowing a visual separation of controls. |
| [AddText](#addtext) | Adds text content to the dialog. |
| [EnableOpenDropDown](#enableopendropdown) | Enables a drop-down list on the Open or Save button in the dialog. |
| [EndVisualGroup](#endvisualgroup) | Stops the addition of elements to a visual group in the dialog. |
| [GetCheckButtonState](#getcheckburronstate) | Gets the current state of a check button (check box) in the dialog. |
| [GetControlState](#getcontrolstate) | Gets the current visibility and enabled states of a given control. |
| [GetEditBoxText](#geteditboxtext) | Gets the current text in an edit box control. |
| [GetSelectedControlItem](#getselectedcontrolitem) | Gets a particular item from specified container controls in the dialog. |
| [MakeProminent](#makeprominent) | Places a control in the dialog so that it stands out compared to other added controls. |
| [RemoveControlItem](#removeconrolitem) | Removes an item from a container control in the dialog. |
| [SetCheckButtonState](#setcheckbuttonstate) | Sets the state of a check button (check box) in the dialog. |
| [SetControlItemState](#secontrolitemstate) | Sets the current state of an item in a container control found in the dialog. |
| [SetControlItemText](#secontrolitemtext) | Sets the text of a control item. For example, the text that accompanies a radio button or an item in a menu. |
| [SetControlLabel](#setcontrollabel) | Sets the text associated with a control, such as button text or an edit box label. |
| [SetControlState](#setcontrolstate) | Sets the current visibility and enabled states of a given control. |
| [SetEditBoxText](#seteditboxtext) | Sets the text in an edit box control found in the dialog. |
| [SetSelectedControlItem](#setselectedcontrolitem) | Sets the selected state of a particular item in an option button group or a combo box found in the dialog. |
| [StartVisualGroup](#startvisualgroup) | Declares a visual group in the dialog. Subsequent calls to any "add" method add those elements to this group. |

---

## AddCheckButton

Adds a check button (check box) to the dialog.

```
FUNCTION AddCheckButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR, BYVAL bChecked AS BOOLEAN = FALSE) AS HRESULT
```
---

## AddComboBox

Adds a combo box to the dialog.

```
FUNCTION AddComboBox (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
---

## AddControlItem

Adds an item to a container control in the dialog.

```
FUNCTION AddControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
---

## AddEditBox

Adds an edit box control to the dialog.

```
FUNCTION AddEditBox (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
---

## AddMenu

Adds a menu to the dialog.

```
FUNCTION AddMenu (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
---

## AddPushButton

Adds a button to the dialog.

```
FUNCTION AddPushButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
---

## AddRadioButtonList

Adds an option button (also known as radio button) group to the dialog.

```
FUNCTION AddRadioButtonList (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
---

## AddSeparator

Adds a separator to the dialog, allowing a visual separation of controls.

```
FUNCTION AddRadioButtonList (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
---

## AddText

Adds text content to the dialog.

```
FUNCTION AddText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
```
---

## EnableOpenDropDown

Enables a drop-down list on the Open or Save button in the dialog.

```
FUNCTION EnableOpenDropDown (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
---

## EndVisualGroup

Stops the addition of elements to a visual group in the dialog.

```
FUNCTION EndVisualGroup () AS HRESULT
```
---

## GetCheckButtonState

Gets the current state of a check button (check box) in the dialog.

```
FUNCTION GetCheckButtonState (BYVAL dwIDCtl AS DWORD) AS BOOLEAN
```
---

## GetControlItemState

Gets the current state of an item in a container control found in the dialog.

```
FUNCTION GetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYREF dwState AS CDCONTROLSTATEF) AS HRESULT
```
---

## GetControlState

Gets the current visibility and enabled states of a given control.

```
FUNCTION GetControlState (BYVAL dwIDCtl AS DWORD) AS CDCONTROLSTATEF
```
---

## GetEditBoxText

Gets the current text in an edit box control.

```
FUNCTION GetEditBoxText (BYVAL dwIDCtl AS DWORD) AS DWSTRING
```
---

## GetSelectedControlItem

Gets a particular item from specified container controls in the dialog.

```
FUNCTION GetSelectedControlItem (BYVAL dwIDCtl AS DWORD) AS DWSTRING
```
---

## MakeProminent

Places a control in the dialog so that it stands out compared to other added controls.

```
FUNCTION MakeProminent (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
---

## RemoveControlItem

Removes an item from a container control in the dialog.

```
FUNCTION RemoveControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwItem AS DWORD) AS HRESULT
```
---

## SetCheckButtonState

Sets the state of a check button (check box) in the dialog.

```
FUNCTION SetCheckButtonState (BYVAL dwIDCtl AS DWORD, BYVAL bCkecked AS BOOLEAN) AS HRESULT
```
---

## SetControlItemState

Sets the current state of an item in a container control found in the dialog.

```
FUNCTION SetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
```
---

## SetControlItemText

Sets the text of a control item. For example, the text that accompanies a radio button or an item in a menu.

```
FUNCTION SetControlItemText (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
---

## SetControlLabel

Sets the text associated with a control, such as button text or an edit box label.

```
FUNCTION SetControlLabel (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
---

## SetControlState

Sets the current visibility and enabled states of a given control.

```
FUNCTION SetControlState (BYVAL dwIDCtl AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
```
---

## SetEditBoxText

Sets the text in an edit box control found in the dialog.

```
FUNCTION SetEditBoxText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
```
---

## SetSelectedControlItem

Sets the selected state of a particular item in an option button group or a combo box found in the dialog.

```
FUNCTION SetSelectedControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
```
---

## StartVisualGroup

Declares a visual group in the dialog. Subsequent calls to any "add" method add those elements to this group.

```
FUNCTION StartVisualGroup (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
---
