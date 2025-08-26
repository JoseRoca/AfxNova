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
