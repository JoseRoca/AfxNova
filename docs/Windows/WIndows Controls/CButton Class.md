# CButton Class

Wrapper class on top of all the Windows `Button` messages and functions.

A button control is a small, rectangular child window that can be clicked on and off. Buttons can be used alone or in groups and can either be labeled or appear without text. A button typically changes appearance when the user clicks it.

Typical buttons are the check box, radio button, and pushbutton. A `CButton` object can become any of these, based on the button style specified at its initialization by the **AddControl** method of the `CWindow`class.

**Include file**: CButton.inc.

| Name       | Description |
| ---------- | ----------- |
| [Check](#check) | Changes the button's state to checked. |
| [CheckDlgButton](#checkdlgbutton) | Changes the check state of a button control. |
| [CheckRadioButton](#checkradiobutton) | Adds a check mark to (checks) a specified radio button in a group and removes a check mark from (clears) all other radio buttons in the group. |
| [Click](#click) | Simulates the user clicking a button. |
| [DeleteBitmap](#deletebitmap) | Deletes a bitmap associated with a button. |
| [DeleteIcon](#deleteicon) | Deletes an icon associated with a button. |
| [DeleteImage](#deleteimage) | Deletes an image (icon or bitmap) associated with a button. |
| [Disable](#disable) | Disables a button. |
| [Enable](#enable) | Enables a button. |
| [GetCheck](#getcheck) | Gets the check state of a radio button or check box.  |
| [GetIdealSize](#getidealsize) | Gets the size of the button that best fits its text and image, if an image list is present. |
| [GetBitmap](#getbitmap) | Retrieves a handle to the bitmap associated with the button. |
| [GetIcon](#geticon) | Retrieves a handle to the icon associated with the button. |
| [GetImage](#getimage) | Retrieves a handle to the image (icon or bitmap) associated with the button. |
| [GetImageList](#getimagelist) | Gets the BUTTON_IMAGELIST structure that describes the image list assigned to a button control.  |
| [GetNote](#getnote) | Gets the text of the note associated with the Command Link button. |
| [GetNoteLength](#getnotelength) | Gets the length of the note text that may be displayed in the description for a command link. |
| [GetSplitInfo](#getsplitinfo) | Gets information for a specified split button control. |
| [GetState](#getstate) | Retrieves the state of a button or check box. |
| [GetStyle](#getstyle) | Retrieves the style of button. |
| [GetText](#gettext) | Retrieves the text in a button control. |
| [GetTextLength](#gettextlength) | Retrieves the length of the text in a button control. |
| [GetTextMargin](#gettextmargin) | Retrieves the margins used to draw text in a button control. |
| [Gray](#gray) | Sets the button state to grayed, indicating an indeterminate state. |
| [IsDlgButtonChecked](#isdlgbuttonchecked) | Determines whether a button control is checked or whether a three-state button control is checked, unchecked, or indeterminate. |
| [SetBitmap](#setbitmap) | Associates a new bitmap with a button. |
| [SetCheck](#setcheck) | Sets the check state of a radio button or check box. |
| [SetDontClick](#setdontclick) | Sets a flag on a radio button that controls the generation of BN_CLICKED messages when the button receives focus. |
| [SetDropdownState](#setdropdownstate) | Sets the drop down state for a specified button with style of BS_SPLITBUTTON. |
| [SetElevationRequiredState](#setelevationrequiredstate) | Sets the elevation required state for a specified button or command link to display an elevated icon. |
| [SetIcon](#seticon) | Associates a new icon with the button. |
| [SetImage](#setimage) | Associates a new image (icon or bitmap) with the button. |
| [SetImageList](#setimagelist) | Assigns an image list to a button control. |
| [SetNote](#setnote) | Sets the text of the note associated with a specified Command Link button. |
| [SetSplitInfo](#setsplitinfo) | Sets information for a specified split button control. |
| [SetState](#setstate) | Sets the highlight state of a button. |
| [SetStyle](#setstyle) | Sets the style of a button. |
| [SetText](#settext) | Sets the text of a button. |
| [SetTextMargin](#settextmargin) | Sets the margins for drawing text in a button control. |
| [Uncheck](#uncheck) | Sets the button state to cleared. |

---

### Check

Changes the button's state to checked.

```
SUB CButton.Check (BYVAL hButton AS HWND)
```
#### Usage example
```
CButton.Check(hButton)
```
---

### Click

Simulates the user clicking a button. This message causes the button to receive the **WM_LBUTTONDOWN** and **WM_LBUTTONUP** messages, and the button's parent window to receive a **BN_CLICKED** notification code.

```
SUb Click (BYVAL hButton AS HWND)
```
#### Remarks

If the button is in a dialog box and the dialog box is not active, the **BM_CLICK** message might fail. To ensure success in this situation, call the **SetActiveWindow** function to activate the dialog box before sending the **BM_CLICK** message to the button.

#### Usage example
```
CButton.Click(hButton)
```
---

### DeleteBitmap

Deletes a bitmap associated with a button. Returns TRUE or FALSE.

```
FUNCTION DeleteBitmap (BYVAL hButton AS HWND) AS BOOLEAN
```

#### Usage example
```
CButton.DeleteBitmap(hButton)
```
---

### DeleteIcon

Deletes an icon associated with a button. Returns TRUE or FALSE.

```
FUNCTION DeleteIcon (BYVAL hButton AS HWND) AS BOOLEAN
```

#### Usage example
```
CButton.DeleteIcon(hButton)
```
---

### DeleteImage

Deletes an image (icon or bitmap) associated with a button. Returns TRUE or FALSE.

```
FUNCTION DeleteImage (BYVAL hButton AS HWND, BYVAL ImageType AS DWORD) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hButton* | The handle of the button control. |
| *ImageType* | The type of image to delete. This parameter can be one of the following values:<br>IMAGE_BITMAP<br>IMAGE_ICON |

#### Usage examples
```
CButton.Deleteimage(hButton, IMAGE_BITMAP)
CButton.Deleteimage(hButton, IMAGE_ICON)
```
---

### Disable

Disables a button. Returns FALSE if the windows was previously disabled; otherwise TRUE.

```
FUNCTION Disable (BYVAL hButton AS HWND) AS BOOLEAN
```

#### Usage example
```
CButton.Disable(hButton)
```
---

### Enable

Enables a button. Returns FALSE if the windows was previously enabled; otherwise TRUE.

```
FUNCTION Enable (BYVAL hButton AS HWND) AS BOOLEAN
```
#### Usage example
```
CButton.Enable(hButton)
```
---

### <a name="getcheck"></a>GetCheck

Gets the check state of a radio button or check box. 

```
FUNCTION GetCheck () AS DWORD
```
#### Return value

The return value from a button created with the **BS_AUTOCHECKBOX**, **BS_AUTORADIOBUTTON**, **BS_AUTO3STATE**, **BS_CHECKBOX**, **BS_RADIOBUTTON**, or **BS_3STATE** style can be one of the following.

| Return code  | Description |
| ------------ | ----------- |
| **BST_CHECKED** | Button is checked. |
| **BST_INDETERMINATE** | Button is grayed, indicating an indeterminate state (applies only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style). |
| **BST_UNCHECKED** | Button is cleared. |

---

### <a name="getidealsize"></a>GetIdealSize

Gets the size of the button that best fits its text and image, if an image list is present.

```
FUNCTION GetIdealSize (BYREF tSize AS SIZE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *tsize* | A pointer to a **SIZE** structure that receives the desired size of the button, including text and image list, if present. The calling application is responsible for allocating this structure. Set the **cx** and **cy** members to zero to have the ideal height and width returned in the **SIZE** structure. To specify a button width, set the **cx** member to the desired button width. The system will calculate the ideal height for this width and return it in the **cy** member. |

#### Return value

If the message succeeds, it returns TRUE. Otherwise it returns FALSE.

#### Remarks

If no special button width is desired, then you must set both members of **SIZE** to zero to calculate and return the ideal height and width. If the value of the **cx** member is greater than zero, then this value is considered the desired button width, and the ideal height for this width is calculated and returned in the **cy** member.

This message is most applicable to PushButtons. When sent to a PushButton, the message retrieves the bounding rectangle required to display the button's text. Additionally, if the PushButton has an image list, the bounding rectangle is also sized to include the button's image.

When sent to a button of any other type, the size of the control's window rectangle is retrieved.

To use this message, you must provide a manifest specifying Comclt32.dll version 6.0.

---

### <a name="getbitmap"></a>GetBitmap

Retrieves a handle to the bitmap associated with the button.

#### Return value

The return value is a handle to the bitmap, if any; otherwise, it is NULL.

---

### <a name="geticon"></a>GetIcon

Retrieves a handle to the icon associated with the button.

#### Return value

The return value is a handle to the icon, if any; otherwise, it is NULL.

---

### <a name="getimage"></a>GetImage

Retrieves a handle to the image (icon or bitmap) associated with the button.

```
FUNCTION GetImage (BYVAL imageType aS LONG) AS HANDLE
```

| Parameter | Description |
| --------- | ----------- |
| *nType* | The type of image to associate with the button. This parameter can be one of the following values. |

| Value | Meaning |
| ------------ | ----------- |
| **IMAGE_BITMAP** | A bitmap. |
| **IMAGE_ICON** | An icon. |

#### Return value

The return value is a handle to the image, if any; otherwise, it is NULL.

---

### <a name="getimagelist"></a>GetImageList

Gets the **BUTTON_IMAGELIST** structure that describes the image list assigned to a button control. 

```
FUNCTION GetImageList () AS BUTTON_IMAGELIST PTR
```

#### Return value

A pointer to a **BUTTON_IMAGELIST** structure that contains image list information.

---

### <a name="getnote"></a>GetNote

Gets the text of the note associated with the Command Link button.

```
FUNCTION GetNote () AS DWSTRING
```

#### Remarks

The **BCM_GETNOTE** message works only with buttons that have the **BS_COMMANDLINK** or **BS_DEFCOMMANDLINK** button style.

**GetLastResul** will contain:

**ERROR_NOT_SUPPORTED**, if the button does not have the **BS_DEFCOMMANDLINK** or **BS_COMMANDLINK** style.

---

### <a name="getnotelength"></a>GetNoteLength

Gets the length of the note text that may be displayed in the description for a command link.

```
FUNCTION GetNoteLength () AS LONG
```

#### Return value

Returns the length of the note text in Unicode characters, not including any terminating NULL, or zero if there is no note text.

---

### <a name="getsplitindo"></a>GetSplitInfo

Gets information for a split button control. 

```
FUNCTION GetSplitInfo (BYREF pInfo AS BUTTON_SPLITINFO) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *pInfo* | A pointer to a **BUTTON_SPLITINFO** structure to receive information on the button. The caller is responsible for allocating the memory for the structure. Set the **mask** member of this structure to determine what information to receive. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

#### Remarks

Use this message only with the **BS_SPLITBUTTON** and **BS_DEFSPLITBUTTON** button styles.

---

### <a name="getstate"></a>GetState

Retrieves the state of a button or check box. 

```
FUNCTION GetState () AS DWORD
```

#### Return value

The return value specifies the current state of the button. It is a combination of the following values.

| Return code  | Description |
| ------------ | ----------- |
| **BST_CHECKED** | The button is checked. |
| **BST_DROPDOWNPUSHED** | The button is in the drop-down state. Applies only if the button has the **TBSTYLE_DROPDOWN** style. |
| **BST_FOCUS** | The button has the keyboard focus. |
| **BST_HOT** | The button is hot; that is, the mouse is hovering over it. |
| **BST_INDETERMINATE** | The state of the button is indeterminate. Applies only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style. |
| **BST_PUSHED** | The button is being shown in the pushed state. |
| **BST_UNCHECKED** | No special state. Equivalent to zero. |

---

### <a name="getstyle"></a>GetStyle

Retrieves the style of button.

```
FUNCTION GetStyle () AS DWORD
```
---

### <a name="gettext"></a>GetText

Retrieves the text in a button control.

```
FUNCTION GetText () AS DWSTRING
```
---

### <a name="gettextlength"></a>GetTextLength

Determines the length, in characters, of the text associated with a button.

```
FUNCTION GetTextLength () AS LONG
```

#### Return value

The return value is the length of the text in characters, not including the terminating null character.

---

### <a name="gettextmargin"></a>GetTextMargin

Retrieves the margins used to draw text in a button control.

```
FUNCTION GetTextMargin (BYREF tMargin AS RECT) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *tMargin* | A pointer to a **RECT** structure that contains the margins to use for drawing text. |

#### Return value

If the message succeeds, it returns TRUE. Otherwise it returns FALSE.

#### Remarks

To use this message, you must provide a manifest specifying Comclt32.dll version 6.0. For more information on manifests, see Enabling Visual Styles.

---

### <a name="gray"></a>Gray

Sets the button state to grayed, indicating an indeterminate state. Use this value only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style.

```
SUB Gray ()
```
---

### <a name="isdlgbuttonchecked"></a>IsDlgButtonChecked

Determines whether a button control is checked or whether a three-state button control is checked, unchecked, or indeterminate.

```
FUNCTION IsDlgButtonChecked (BYVAL cIDButton AS LONG) AS UINT
```
| Parameter | Description |
| --------- | ----------- |
| *cIDButton* | The identifier of the button control. |

#### Return value

The return value from a button created with the **BS_AUTOCHECKBOX**, **BS_AUTORADIOBUTTON**, **BS_AUTO3STATE**, **BS_CHECKBOX**, **BS_RADIOBUTTON**, or **BS_3STATE** styles can be one of the values in the following table. If the button has any other style, the return value is zero.

| Return code | Description |
| ----------- | ----------- |
| **BST_CHECKED** | The button is checked. |
| **BST_INDETERMINATE** | The button is in an indeterminate state (applies only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style). |
| **BST_UNCHECKED** | The button is not checked. |

#### Remarks

The **IsDlgButtonChecked** method sends a **BM_GETCHECK** message to the specified button control.

---

### <a name="setbitmap"></a>SetBitmap

Associates a new bitmap with the button.

```
FUNCTION SetBitmap (BYVAL hbmp AS HBITMAP) AS HBITMAP
```

#### Return value

The return value is a handle to the bitmap previously associated with the button, if any; otherwise, it is NULL.

---

### <a name="setcheck"></a>SetCheck

Sets the check state of a radio button or check box.

```
SUB SetCheck (BYVAL checkState AS LONG)
```
| Value | Meaning |
| ----- | ----------- |
| *checkState* | The check state. This parameter can be one of the following values. |

| Parameter | Description |
| --------- | ----------- |
| **BST_CHECKED** | Sets the button state to checked. |
| **BST_INDETERMINATE** | Sets the button state to grayed, indicating an indeterminate state. Use this value only if the button has the **BS_3STATE** or **BS_AUTO3STATE** style. |
| **BST_UNCHECKED** | Sets the button state to cleared. |

#### Return value

This message always returns zero.

#### Remarks

The **BM_SETCHECK** message has no effect on push buttons.

---

### <a name="setdontclick"></a>SetDontClick

Sets a flag on a radio button that controls the generation of BN_CLICKED messages when the button receives focus.

```
SUB SetDontClick (BYVAL bState AS BOOLEAN)
```

| Parameter | Description |
| --------- | ----------- |
| *bState* | A BOOLEAN that specifies the state. TRUE to set the flag, otherwise FALSE. |

#### Return value

No return value.

---

### <a name="setdropdownstate"></a>SetDropdownState

Sets the drop down state for a button with style **TBSTYLE_DROPDOWN**.
```
FUNCTION SetDropDownState (BYVAL bDropDown AS BOOLEAN) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *bDropDown* | A BOOLEAN that is TRUE for state of **BST_DROPDOWNPUSHED**, or FALSE otherwise. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

### <a name="seticon"></a>SetIcon

Associates a new icon with the button.

```
FUNCTION SetIcon (BYVAL hIcon AS HICON) AS HICON
```

#### Return value

The return value is a handle to the icon previously associated with the button, if any; otherwise, it is NULL.

---

### <a name="setimage"></a>SetImage

Associates a new image (icon or bitmap) with the button. The return value is a handle to the image previously associated with the button, if any; otherwise, it is NULL.

```
FUNCTION SetImage (BYVAL ImageType AS DWORD, BYVAL hImage AS HANDLE) AS HANDLE
```

#### Return value

The return value is a handle to the image (icon or bitmap) previously associated with the button, if any; otherwise, it is NULL.

---

### <a name="setimagelist"></a>SetImageList

Assigns an image list to a button control.

```
FUNCTION SetImageList (BYVAL pbuttonImagelist AS BUTTON_IMAGELIST PTR) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *pbuttonImagelist* | A pointer to a **BUTTON_IMAGELIST** structure that contains image list information. |

```
FUNCTION SetImageList (BYVAL pbuttonImagelist AS BUTTON_IMAGELIST PTR) AS BOOLEAN
FUNCTION SetImageList (BYVAL himl AS HIMAGELIST, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, _
   BYVAL nRight AS LONG, BYVAL nBottom AS LONG, BYVAL uALign AS DWORD = 0) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hImageList* | The handle to the image list. |
| *nLeft* | The x-coordinate of the upper-left corner of the rectangle for the image. |
| *nTop* | The y-coordinate of the upper-left corner of the rectangle for the image. |
| *nRight* | The x-coordinate of the lower-right corner of the rectangle for the image. |
| *nBottom* | The y-coordinate of the lower-right corner of the rectangle for the image. |
| *uAlign* | The alignment to use. This parameter can be one of the following values: |

| Value  | Meaning |
| ------ | ------- |
| **BUTTON_IMAGELIST_ALIGN_LEFT** | Align the image with the left margin. |
| **BUTTON_IMAGELIST_ALIGN_RIGHT** | Align the image with the right margin. |
| **BUTTON_IMAGELIST_ALIGN_TOP** | Align the image with the top margin. |
| **BUTTON_IMAGELIST_ALIGN_BOTTOM** | Align the image with the bottom margin. |
| **BUTTON_IMAGELIST_ALIGN_CENTER** |  Center the image. |

The default value is **BUTTON_IMAGELIST_ALIGN_LEFT**.

#### Return value

If the message succeeds, it returns TRUE. Otherwise it returns FALSE.

#### Remarks

To use this message, you must provide a manifest specifying Comclt32.dll version 6.0. For more information on manifests, see Enabling Visual Styles.

The image list referred to in the himl member of the **BUTTON_IMAGELIST** structure should contain either a single image to be used for all states or individual images for each state. The following states are defined in vssym32.h.
```
enum PUSHBUTTONSTATES {
    PBS_NORMAL = 1,
    PBS_HOT = 2,
    PBS_PRESSED = 3,
    PBS_DISABLED = 4,
    PBS_DEFAULTED = 5,
    PBS_STYLUSHOT = 6,
};
```
Note that **PBS_STYLUSHOT** is used only on tablet computers.

Each value is an index to the appropriate image in the image list. If only one image is present, it is used for all states. If the image list contains more than one image, each index corresponds to one state of the button. If an image is not provided for each state, nothing is drawn for those states without images.

---

### <a name="setelevationrequiredstate"></a>SetElevationRequiredState

Sets the elevation required state for a specified button or command link to display an elevated icon.

```
FUNCTION SetElevationRequiredState (BYVAL bRequired AS BOOLEAN) AS LONG
```
| Parameter | Description |
| --------- | ----------- |
| *bRequired* | A BOOL that is TRUE to draw an elevated icon, or FALSE otherwise. |

#### Return value

Returns 1 if successful, or an error code otherwise.

#### Remarks

An application must be manifested to use comctl32.dll version 6 to gain this functionality.

---

### <a name="setnote"></a>SetNote

Sets the text of the note associated with a specified command link button.

```
FUNCTION SetNote (BYREF pwsz AS WSTRING) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *pwsz* | The handle to the image list. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

#### Remarks

Beginning with comctl32 DLL version 6.01, command link buttons may have a note.

The **BCM_SETNOTE** message works only with the **BS_COMMANDLINK** and **BS_DEFCOMMANDLINK** button styles.

---

### <a name="setnote"></a>SetNote

Sets the text of the note associated with a specified command link button.

```
FUNCTION SetNote (BYREF pwsz AS WSTRING) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *pwsz* | The handle to the image list. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

#### Remarks

Beginning with comctl32 DLL version 6.01, command link buttons may have a note.

The **BCM_SETNOTE** message works only with the **BS_COMMANDLINK** and **BS_DEFCOMMANDLINK** button styles.

---


### <a name="setnote"></a>SetNote

Sets the text of the note associated with a specified command link button.

```
FUNCTION SetNote (BYREF pwsz AS WSTRING) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *pwsz* | The handle to the image list. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

#### Remarks

Beginning with comctl32 DLL version 6.01, command link buttons may have a note.

The **BCM_SETNOTE** message works only with the **BS_COMMANDLINK** and **BS_DEFCOMMANDLINK** button styles.

---

### <a name="setnote"></a>SetNote

Sets the text of the note associated with a specified command link button.

```
FUNCTION SetNote (BYREF pwsz AS WSTRING) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *pwsz* | The handle to the image list. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

#### Remarks

Beginning with comctl32 DLL version 6.01, command link buttons may have a note.

The **BCM_SETNOTE** message works only with the **BS_COMMANDLINK** and **BS_DEFCOMMANDLINK** button styles.

---

### <a name="Setsplitinfo"></a>SetSplitInfo

Sets information for a specified split button control.

```
FUNCTION SetSplitInfo (BYREF pInfo AS BUTTON_SPLITINFO) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *pInfo* | A pointer to a **BUTTON_SPLITINFO** structure containing information about the split button. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

#### Remarks

Use this message only with the **BS_SPLITBUTTON** and **BS_DEFSPLITBUTTON** button styles.

---

### <a name="setstate"></a>SetState

Sets the highlight state of a button. The highlight state indicates whether the button is highlighted as if the user had pushed it. 

```
SUB SetState (BYVAL bstate AS BOOLEAN)
```
| Parameter | Description |
| --------- | ----------- |
| *bstate* | A BOOLEAN that specifies whether the button is highlighted. A value of TRUE highlights the button. A value of FALSE removes any highlighting. |

#### Return value

This message always returns zero.

####Remarks

Highlighting affects only the appearance of a button. It has no effect on the check state of a radio button or check box.

A button is automatically highlighted when the user positions the cursor over it and presses and holds the left mouse button. The highlighting is removed when the user releases the mouse button.

---

### <a name="setstyle"></a>SetStyle

Sets the style of a button.

```
SUB SetStyle (BYVAL dwStyle AS DWORD, BYVAL bRedraw AS BOOLEAN)
```
| Parameter | Description |
| --------- | ----------- |
| *dwStyle* | The button style. This parameter can be a combination of button styles. |
| *bRedraw* | A BOOLEAN that specifies whether the button is to be redrawn. A value of TRUE redraws the button; a value of FALSE does not redraw the button. |

For a table of button styles, see [Button Styles](https://learn.microsoft.com/en-us/windows/win32/controls/button-styles).

---

### <a name="setstext"></a>SetText

Sets the text of a button.

```
FUNCTION SetText (BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *pwszText* | The new text. |

#### Return value

If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.

---

### <a name="settextmargin"></a>SetTextMargin

Sets the margins used to draw text in a button control.

```
FUNCTION SetTextMargin (BYREF tMargin AS RECT) AS BOOLEAN
```
| Parameter | Description |
| --------- | ----------- |
| *tMargin* | A pointer to a **RECT** structure that specifies the margins to use for drawing text. |

#### Return value

If the message succeeds, it returns TRUE. Otherwise it returns FALSE.

#### Remarks

To use this message, you must provide a manifest specifying Comclt32.dll version 6.0. For more information on manifests, see Enabling Visual Styles.

---

### <a name="uncheck"></a>Uncheck

Sets the button state to cleared.

```
SUB CButton.Uncheck ()
```
---


