# CIFileDialogCustomize Class

Wraps the `IFileDialogCustomize`interface, that exposes methods that allow an application to add controls to a common file dialog.

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

#### Usage example

```
' // Get a pointer to the IFileDialog interface
DIM pfdlg AS IFileDialog PTR = pofd.GetIFileDialogPtr
IF pfdlg THEN
   ' // Initialize an instance of the CIFileDialogCustomize class
   DIM pCust AS CIFileDialogCustomize = pfdlg
   ' // Relese the IFileDialog interface
   pfdlg->lpvtbl->Release(pfdlg)
   ' // Add controls to the dialog
   IF pCust.GetPtr THEN
      ' Call methods of the CIFileDialogCustomize class
      pCust.AddCheckButton(1001, "My check button", TRUE)
      pCust.AddComboBox(1002)
      pCust.AddEditBox(1003, "My edit control")
      pCust.AddEditBox(1004, "My menu")
      pCust.AddPushButton(1005, "My push button")
      pCust.AddSeparator(1006)
      pCust.AddText(1007, "My text")
      pCust.EnableOpenDropDown(1008)
   END IF
END IF
```
See also a full example at the end of this guide.

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
| [SetControlItemState](#setcontrolitemstate) | Sets the current state of an item in a container control found in the dialog. |
| [SetControlItemText](#setcontrolitemtext) | Sets the text of a control item. For example, the text that accompanies a radio button or an item in a menu. |
| [SetControlLabel](#setcontrollabel) | Sets the text associated with a control, such as button text or an edit box label. |
| [SetControlState](#setcontrolstate) | Sets the current visibility and enabled states of a given control. |
| [SetEditBoxText](#seteditboxtext) | Sets the text in an edit box control found in the dialog. |
| [SetSelectedControlItem](#setselectedcontrolitem) | Sets the selected state of a particular item in an option button group or a combo box found in the dialog. |
| [StartVisualGroup](#startvisualgroup) | Declares a visual group in the dialog. Subsequent calls to any "add" method add those elements to this group. |

---

## AddCheckButton

Adds a check button (check box) to the dialog.

```
FUNCTION AddCheckButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR, _
   BYVAL bChecked AS BOOLEAN = FALSE) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the check button to add. |
| *pwszLabel* |A pointer to a buffer that contains the button text as a null-terminated Unicode string. |
| *bChecked* | A BOOLEAN indicating the current state of the check button. TRUE if checked; FALSE otherwise. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state for this control is enabled and visible.

---

## AddComboBox

Adds a combo box to the dialog.

```
FUNCTION AddComboBox (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the vombo box to add. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state for this control is enabled and visible.

---

## AddControlItem

Adds an item to a container control in the dialog.

```
FUNCTION AddControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the container control to which the item is to be added. |
| *dwIDItem* | The ID of the item. |
| *pwszLabel* |A pointer to a buffer that contains the item's text, which can be either a label or, in the case of a drop-down list, the item itself. This text is a null-terminated Unicode string. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state for this item is enabled and visible. Items in control groups cannot be changed after they have been created, with the exception of their enabled and visible states.

Container controls include option button groups, combo boxes, drop-down lists on the **Open** or **Save** button, and menus.

---

## AddEditBox

Adds an edit box control to the dialog.

```
FUNCTION AddEditBox (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the edit box to add. |
| *pwszLabel* | A pointer to a null-terminated Unicode string that provides the default text displayed in the edit box. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state for this control is enabled and visible.

To add a label next to the edit box, place it in a visual group with **StartVisualGroup**.

---

## AddMenu

Adds a menu to the dialog.

```
FUNCTION AddMenu (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the menu to add. |
| *pwszLabel* | A pointer to a buffer that contains the menu name as a null-terminated Unicode string. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state for this control is enabled and visible.

To add items to this control, use **AddControlItem**.

---

## AddPushButton

Adds a button to the dialog.

```
FUNCTION AddPushButton (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the button to add. |
| *pwszLabel* | A pointer to a buffer that contains the button text as a null-terminated Unicode string. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state for this control is enabled and visible.

---

## AddRadioButtonList

Adds an option button (also known as radio button) group to the dialog.

```
FUNCTION AddRadioButtonList (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the option button group to add. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state for this control is enabled and visible.

---

## AddSeparator

Adds a separator to the dialog, allowing a visual separation of controls.

```
AddSeparator (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The control ID of the separator. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state for this control is enabled and visible.

---

## AddText

Adds text content to the dialog.

```
FUNCTION AddText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the text to add. |
| *pwszText* | A pointer to a buffer that contains the text as a null-terminated Unicode string. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state for this control is enabled and visible.

---

## EnableOpenDropDown

Enables a drop-down list on the Open or Save button in the dialog.

```
FUNCTION EnableOpenDropDown (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the drop-down list. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The Open or Save button label takes on the text of the first item in the drop-down. This overrides any label set by IFileDialog::SetOkButtonLabel.

Use **AddControlItem** to add items to the drop-down.

---

## EndVisualGroup

Stops the addition of elements to a visual group in the dialog.

```
FUNCTION EndVisualGroup () AS HRESULT
```
#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

---

## GetCheckButtonState

Gets the current state of a check button (check box) in the dialog.

```
FUNCTION GetCheckButtonState (BYVAL dwIDCtl AS DWORD) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the check box. |

#### Return value

TRUE means checked; FALSE, unchecked.

---

## GetControlItemState

Gets the current state of an item in a container control found in the dialog.

```
FUNCTION GetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYREF dwState AS CDCONTROLSTATEF) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the container control. |
| *dwIDCtl* | The ID of the item. |

#### Return value

A member of the **CDCONTROLSTATE** enumeration that indicates the current state of the control.

| Constant  | Description |
| --------- | ----------- |
| **CDCS_INACTIVE** | The control is inactive and cannot be accessed by the user. |
| **CDCS_ENABLED** | The control is active. |
| **CDCS_VISIBLE** | The control is visible. The absence of this value indicates that the control is hidden. |
| **CDCS_ENABLEDVISIBLE** | Windows 7 and later. The control is visible and enabled. |

#### Remarks

The default state of a control item is enabled and visible. Items in control groups cannot be changed after they have been created, with the exception of their enabled and visible states.

Container controls include option button groups, combo boxes, drop-down lists on the Open or Save button, and menus.

---

## GetControlState

Gets the current visibility and enabled states of a given control.

```
FUNCTION GetControlState (BYVAL dwIDCtl AS DWORD) AS CDCONTROLSTATEF
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the container control. |

#### Return value

A member of the **CDCONTROLSTATE** enumeration that indicates the current state of the control.

| Constant  | Description |
| --------- | ----------- |
| **CDCS_INACTIVE** | The control is inactive and cannot be accessed by the user. |
| **CDCS_ENABLED** | The control is active. |
| **CDCS_VISIBLE** | The control is visible. The absence of this value indicates that the control is hidden. |
| **CDCS_ENABLEDVISIBLE** | Windows 7 and later. The control is visible and enabled. |

---

## GetEditBoxText

Gets the current text in an edit box control.

```
FUNCTION GetEditBoxText (BYVAL dwIDCtl AS DWORD) AS DWSTRING
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the edit box. |

#### Return value

The text of the edit box control.

---

## GetSelectedControlItem

Gets a particular item from specified container controls in the dialog.

```
FUNCTION GetSelectedControlItem (BYVAL dwIDCtl AS DWORD) AS DWSTRING
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the container control. |

#### Return value

The ID of the item that the user selected in the control.

---

## MakeProminent

Places a control in the dialog so that it stands out compared to other added controls.

```
FUNCTION MakeProminent (BYVAL dwIDCtl AS DWORD) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the control. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

This method causes the control to be placed near the Open or Save button instead of being grouped with the rest of the custom controls.

Only check buttons (check boxes), push buttons, combo boxes, and menus—or a visual group that contains only a single item of one of those types—can be made prominent.

Only one control can be marked in this way. If a dialog has only one added control, that control is marked as prominent by default.

---

## RemoveControlItem

Removes an item from a container control in the dialog.

```
FUNCTION RemoveControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwItem AS DWORD) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the container control from which the item is to be removed. |
| *dwItem* | The ID of the item. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

Container controls include option button groups, combo boxes, drop-down lists on the Open or Save button, and menus.

---

## SetCheckButtonState

Sets the state of a check button (check box) in the dialog.

```
FUNCTION SetCheckButtonState (BYVAL dwIDCtl AS DWORD, BYVAL bCkecked AS BOOLEAN) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the check box. |
| *bCkecked* | A BOOEANL value that indicates whether the box is checked. TRUE means checked; FALSE, unchecked. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

---

## SetControlItemState

Sets the current state of an item in a container control found in the dialog.

```
FUNCTION SetControlItemState (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the container control. |
| *dwIDItem* | The ID of the item. |
| *dwState* | One or more values from the **CDCONTROLSTATE** enumeration that indicate the new state of the control. |

| Constant  | Description |
| --------- | ----------- |
| **CDCS_INACTIVE** | The control is inactive and cannot be accessed by the user. |
| **CDCS_ENABLED** | The control is active. |
| **CDCS_VISIBLE** | The control is visible. The absence of this value indicates that the control is hidden. |
| **CDCS_ENABLEDVISIBLE** | Windows 7 and later. The control is visible and enabled. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state of a control item is enabled and visible. Items in control groups cannot be changed after they have been created, with the exception of their enabled and visible states.

Container controls include option button groups, combo boxes, drop-down lists on the Open or Save button, and menus.

---

## SetControlItemText

Sets the text of a control item. For example, the text that accompanies a radio button or an item in a menu.

```
FUNCTION SetControlItemText (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the container control. |
| *dwIDItem* | The ID of the item. |
| *pwszLabel* | A pointer to a null-terminated buffer that contains a Unicode string with the text. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

The default state of a control item is enabled and visible. Items in control groups cannot be changed after they have been created, with the exception of their enabled and visible states.

Container controls include option button groups, combo boxes, drop-down lists on the Open or Save button, and menus.

---

## SetControlLabel

Sets the text associated with a control, such as button text or an edit box label.

```
FUNCTION SetControlLabel (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the control whose text is to be changed. |
| *pwszLabel* | A pointer to a buffer that contains the text as a null-terminated Unicode string. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

Control labels can be changed at any time, including when the dialog is visible.

---

## SetControlState

Sets the current visibility and enabled states of a given control.

```
FUNCTION SetControlState (BYVAL dwIDCtl AS DWORD, BYVAL dwState AS CDCONTROLSTATEF) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the control. |
| *dwState* | One or more values from the **CDCONTROLSTATE** enumeration that indicate the current state of the control. |

| Constant  | Description |
| --------- | ----------- |
| **CDCS_INACTIVE** | The control is inactive and cannot be accessed by the user. |
| **CDCS_ENABLED** | The control is active. |
| **CDCS_VISIBLE** | The control is visible. The absence of this value indicates that the control is hidden. |
| **CDCS_ENABLEDVISIBLE** | Windows 7 and later. The control is visible and enabled. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

When the dialog is shown, controls cannot be added or removed, but the existing controls can be hidden or disabled at any time.

---

## SetEditBoxText

Sets the text in an edit box control found in the dialog.

```
FUNCTION SetEditBoxText (BYVAL dwIDCtl AS DWORD, BYVAL pwszText AS WSTRING PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the edit box. |
| *pwszText* | A pointer to a buffer that contains the text as a null-terminated Unicode string. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

---

## SetSelectedControlItem

Sets the selected state of a particular item in an option button group or a combo box found in the dialog.

```
FUNCTION SetSelectedControlItem (BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the container control. |
| *dwIDItem* | The ID of the item to display as selected in the control. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

---

## StartVisualGroup

Declares a visual group in the dialog. Subsequent calls to any "add" method add those elements to this group.

```
FUNCTION StartVisualGroup (BYVAL dwIDCtl AS DWORD, BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *dwIDCtl* | The ID of the visual group. |
| *pwszLabel* | A pointer to a buffer that contains text, as a null-terminated Unicode string, that appears next to the visual group. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

Controls will continue to be added to this visual group until you call **EndVisualGroup**.

A visual group can be hidden and disabled like any other control, except that doing so affects all of the controls within it. Individual members of the visual group can also be hidden and disabled singly.

---

## Example code

```
#define _WIN32_WINNT &h0602
#include once "AfxNova/CWindow.inc"
#include once "AfxNova/CIOpenSaveFile.inc"
#include once "AfxNova/CIFileDialogEvents.inc"
#include once "AfxNova/CIFileDialogCustomize.inc"
USING AfxNova

DIM pofd AS CIOpenFileDialog
' // Set the file types
pofd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
pofd.AddFileType("Executable files", "*.exe;*.dll")
pofd.AddFileType("All files", "*.*")
pofd.SetFileTypes()
' // Multiple selection (default is single selection)
DIM options AS FILEOPENDIALOGOPTIONS = pofd.GetOptions
pofd.SetOptions(options OR FOS_ALLOWMULTISELECT)
' // Optional: Set the title of the dialog
'   pofd.SetTitle("A Single-Selection Dialog")
' // Set the folder
pofd.SetFolder(CURDIR)
pofd.SetDefaultExtension("bas")
pofd.SetFileTypeIndex(1)

' // Customization --------------------------------------------
DIM pfdlg AS IFileDialog PTR = pofd.GetIFileDialogPtr
IF pfdlg THEN
   DIM pCust AS CIFileDialogCustomize = pfdlg
   pfdlg->lpvtbl->Release(pfdlg)
   IF pCust.GetPtr THEN
      ' Call methods of the CIFileDialogCustomize class
      pCust.AddCheckButton(1001, "My check button", TRUE)
      pCust.AddComboBox(1002)
      pCust.AddEditBox(1003, "My edit control")
      pCust.AddEditBox(1004, "My menu")
      pCust.AddPushButton(1005, "My push button")
      pCust.AddSeparator(1006)
      pCust.AddText(1007, "My text")
      pCust.EnableOpenDropDown(1008)
   END IF
END IF
' // --------------------------------------------------------

' // Set events
DIM pfde AS ANY PTR = NEW CIFileDialogEvents
pofd.SetEvents(pfde)

' // Display the dialog
DIM hr AS HRESULT = pofd.ShowOpen(hwnd)

' // Folder name
print "Folder name: ";  pofd.GetFolder

' *** Single selection ***
' // Get the result
IF hr = S_OK THEN
   print pofd.GetResult()
END IF

' *** Multiple selection ***
DIM dwsRes AS DWSTRING = pofd.GetResultsString
FOR i AS LONG = 1 TO pofd.GetResultsCount
   PRINT pofd.ParseResults(dwsRes, i)
NEXT
```
