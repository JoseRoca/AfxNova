# CRebar Class

A Rebar control acts as a container for child windows. It can contain one or more bands, and each band can have any combination of a gripper bar, a bitmap, a text label, and one child window. An application assigns a child window—typically another control— to a rebar control band. As you dynamically reposition a rebar control band, the rebar control manages the size and position of the child window assigned to that band. Also, an application can specify a background bitmap for a band, and the rebar control will display the band's child window over the bitmap.

See: [About Rebar Controls](https://learn.microsoft.com/en-us/windows/win32/controls/rebar-controls)

**Include file**: CRebar.inc

| Name       | Description |
| ---------- | ----------- |
| [BeginDrag](#begindrag) | Puts the rebar control in drag-and-drop mode. |
| [DeleteBand](#deleteband) | Deletes a band from a rebar control. |
| [DragMove](#dragmove) | Updates the drag position in the rebar control after a previous call to **BeginDrag**. |
| [EndDrag](#enddrag) | Terminates the rebar control's drag-and-drop operation. |
| [GetBandBorders](#getbandborders) | Retrieves the borders of a band. |
| [GetBandCount](#getbandcount) | Retrieves the count of bands currently in the rebar control. |
| [GetBandInfo](#getbandinfo) | Retrieves information about a specified band in a rebar control. |
| [GetBandMargins](#getbandmargins) | Retrieves the margins of a band. |
| [GetBarHeight](#getbarheight) | Retrieves the height of the rebar control. |
| [GetBarInfo](#getbarinfo) | Retrieves information about the rebar control and the image list it uses. |
| [GetBkColor](#getbkcolor) | Retrieves a rebar control's default background color. |
| [GetColorScheme](#getcolorscheme) | Retrieves the color scheme information from the rebar control. |
| [GetDropTarget](#getdroptarget) | Retrieves a rebar control's **IDropTarget** interface pointer. |
| [GetPalette](#getpalette) | Retrieves the rebar control's current palette. |
| [GetRect](#getrect) | Retrieves the bounding rectangle for a given band in a rebar control. |
| [GetRowCount](#getrowcount) | Retrieves the number of rows of bands in a rebar control. |
| [GetRowHeight](#getrowheight) | Retrieves the height of a specified row in a rebar control. |
| [GetTextColor](#gettextcolor) | Retrieves a rebar control's default text color. |
| [GetTooltips](#gettooltips) | Retrieves the handle to any ToolTip control associated with the rebar control. |
| [HitTest](#hittest) | Determines which portion of a rebar band is at a given point on the screen, if a rebar band exists at that point. |
| [IdToIndex](#idtoindex) | Converts a band identifier to a band index in a rebar control. |
| [InsertBand](#insertband) | Inserts a new band in a rebar control. |
| [MaximizeBand](#maximizeband) | Resizes a band in a rebar control to either its ideal or largest size. |
| [MinimizeBand](#minimizeband) | Resizes a band in a rebar control to its smallest size. |
| [MoveBand](#moveband) | Moves a band from one index to another. |
| [PushChevron](#pushchevron) | Sent to a rebar control to programmatically push a chevron. |
| [RemoveDarkMode](#removedarkmode) | Removes the rebar dark mode. |
| [SetBandInfo](#setbandinfo) | Sets characteristics of an existing band in a rebar control. |
| [SetBandWidth](#setbandwidth) | Sets the characteristics of a rebar control. |
| [SetBarInfo](#setbabarinfo) | Sets the characteristics of a rebar control. |
| [SetBkColor](#setbkcolor) | Sets a rebar control's default background color. |
| [SetColorScheme](#setcolorscheme) | Sets the color scheme information for the rebar control. |
| [SetDarkMode](#setdarkmode) | Sets the rebar dark mode. |
| [SetPalette](#setpalette) | Sets the rebar control's current palette. |
| [SetParent](#setparent) | Sets a rebar control's parent window. |
| [SetTextColor](#settextcolor) | Sets a rebar control's default text color. |
| [SetTooltips](#settooltips) | Associates a tooltip control with the rebar control. |
| [SetWindowTheme](#setwindowtheme) | Sets the visual style of a rebar control. |
| [ShowBand](#showband) | Shows or hides a given band in a rebar control. |
| [SizeToRect](#sizetorect) | Attempts to find the best layout of the bands for the given rectangle. |

---

### BeginDrag

Puts the rebar control in drag-and-drop mode.

```
SUB BeginDrag (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL dwPos AS DWORD)
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Zero-based index of the band that the drag-and-drop operation will affect. |
| *dwPos* | **DWORD** value that contains the starting mouse coordinates. The horizontal coordinate is contained in the **LOWORD** and the vertical coordinate is contained in the **HIWORD**. If you pass (DWORD)-1, the rebar control will use the position of the mouse the last time the control's thread called **GetMessage** or **PeekMessage**. |

#### Remarks

The **BeginDraw**, **DragMove**, and **EndDrag** methods allow you to implement an **IDropTarget** interface for a rebar control. You call **BeginDrag** in response to **IDropTarget.DragEnter**, **DragMove** in response to **IDropTarget.DragOver**, and **EndDrag** in response to **IDropTarget.Drop** and **IDropTarget.DragLeave**.

---

### DeleteBand

Deletes a band from a rebar control.

```
FUNCTION DeleteBand (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Zero-based index of the band to be deleted. |

#### Return value

Returns TRUE if successful, or FLASE otherwise.

---

### DragMove

Updates the drag position in the rebar control after a previous call to **BeginDrag**.

```
SUB DragMove (BYVAL hRebar AS HWND, BYVAL dwPos AS DWORD)
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *dwPos* | **DWORD** value that contains the new mouse coordinates. The horizontal coordinate is contained in the **LOWORD** and the vertical coordinate is contained in the **HIWORD**. If you pass (DWORD)-1, the rebar control will use the position of the mouse the last time the control's thread called **GetMessage** or **PeekMessage**. |

---

### EndDrag

Terminates the rebar control's drag-and-drop operation.

```
SUB EndDrag (BYVAL hRebar AS HWND)
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

---

### GetBandBorders

Retrieves the borders of a band. The result of this message can be used to calculate the usable area in a band.

```
SUB GetBandBorders (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL prc AS RECT PTR)
SUB GetBandBorders (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYREF rc AS RECT)
FUNCTION GetBandBorders (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD) AS RECT
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Zero-based index of the band for which the borders will be retrieved. |
| *prc* | Pointer to a **RECT** structure that will receive the band borders. If the rebar control has the **RBS_BANDBORDERS** style, each member of this structure will receive the number of pixels, on the corresponding side of the band, that constitute the border. If the rebar control does not have the **RBS_BANDBORDERS** style, only the left member of this structure receives valid information. |

---

### GetBandCount

Retrieves the count of bands currently in the rebar control.

```
FUNCTION GetBandCount (BYVAL hRebar AS HWND) AS DWORD
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

#### Return value

The number of bands assigned to the control.

---

### GetBandInfo

Retrieves information about a specified band in a rebar control.

```
FUNCTION GetBandInfo (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL prbbi AS REBARBANDINFOW PTR) AS BOOLEAN
FUNCTION GetBandInfo (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYREF rbbi AS REBARBANDINFOW) AS BOOLEAN
FUNCTION GetBandInfo (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD) AS REBARBANDINFOW
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Zero-based index of the band for which the information will be retrieved. |
| *prbbi* | Pointer to a **REBARBANDINFO** structure that will receive the requested band information. Before sending this message, you must set the **cbSize** member of this structure to the size of the **REBARBANDINFO** structure and set the fMask member to the items you want to retrieve. Additionally, you must set the cch member of the **REBARBANDINFO** structure to the size of the lpText buffer when **RBBIM_TEXT** is specified. |

#### Return value

Returns nonzero if successful, or zero otherwise.

---

### GetBandMargins

Retrieves the margins of a band.

```
SUB GetBandMargins (BYVAL hRebar AS HWND, BYVAL pMargins AS MARGINS PTR)
SUB GetBandMargins (BYVAL hRebar AS HWND, BYREF _margins AS MARGINS)
FUNCTION GetBandMargins (BYVAL hRebar AS HWND) AS MARGINS
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *pMargins* | Pointer to a **MARGINS** structure that receives the retrieved margins. |

---

### GetBarHeight

Retrieves the height of the rebar control.

```
FUNCTION GetBarHeight (BYVAL hRebar AS HWND) AS DWORD
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

#### Return value

The height, in pixels, of the control.

---

### GetBarInfo

Retrieves information about the rebar control and the image list it uses.

```
FUNCTION GetBarInfo (BYVAL hRebar AS HWND, BYVAL prbi AS REBARINFO PTR) AS BOOLEAN
FUNCTION GetBarInfo (BYVAL hRebar AS HWND, BYREF rbi AS REBARINFO) AS BOOLEAN
FUNCTION GetBarInfo (BYVAL hRebar AS HWND) AS REBARINFO
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *prbi* | Pointer to a **REBARINFO** structure that will receive the rebar control information. You must set the **cbSize** member of this structure to sizeof(REBARINFO) before sending this message. |

#### Return value

Returns nonzero if successful, or zero otherwise.

---

### GetBkColor

Retrieves a rebar control's default background color.

```
FUNCTION GetBkColor (BYVAL hRebar AS HWND) AS COLORREF
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

#### Return value

Returns a **COLORREF** value that represent the current default background color.

---

### GetColorScheme

Retrieves the color scheme information from the rebar control.

```
FUNCTION GetColorScheme (BYVAL hRebar AS HWND, BYVAL pcs AS COLORSCHEME PTR) AS BOOLEAN
FUNCTION GetColorScheme (BYVAL hRebar AS HWND, BYREF cs AS COLORSCHEME) AS BOOLEAN
FUNCTION GetColorScheme (BYVAL hRebar AS HWND) AS COLORSCHEME
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *pcs* | Pointer to a **COLORSCHEME** structure that will receive the color scheme information. You must set the dwSize member of this structure to sizeof(COLORSCHEME) before sending this message. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

---

### GetDropTarget

Retrieves a rebar control's **IDropTarget** interface pointer.

```
FUNCTION GetDropTarget (BYVAL hRebar AS HWND) AS IDropTarget PTR
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

#### Return value

Pointer to an **IDropTarget** interface. It is the caller's responsibility to call **Release** on this pointer when it is no longer needed.

---

### GetPalette

Retrieves the rebar control's current palette.

```
FUNCTION GetPalette (BYVAL hRebar AS HWND) AS HPALETTE
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |


#### Return value

Returns an **HPALETTE** that specifies the rebar control's current palette.

---

### GetRect

Retrieves the bounding rectangle for a given band in a rebar control.

```
FUNCTION GetRect (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL prc AS RECT PTR) AS BOOLEAN
FUNCTION GetRect (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYREF rc AS RECT) AS BOOLEAN
FUNCTION GetRect (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD) AS RECT
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Zero-based index of a band in the rebar control. |
| *prc* | Pointer to a **RECT** structure that will receive the bounds of the rebar band. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

---

### GetRowCount

Retrieves the number of rows of bands in a rebar control.

```
FUNCTION GetRect (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD) AS RECT
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

#### Return value

The number of band rows in the control.

---

### GetRowHeight

Retrieves the height of a specified row in a rebar control.

```
FUNCTION GetRowHeight (BYVAL hRebar AS HWND, BYVAL uRow AS DWORD) AS DWORD
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uRow* | Zero-based index of a band. The height of the row that contains the specified band will be retrieved. |

#### Return value

The row height, in pixels.

#### Remarks

To retrieve the number of rows in a rebar control, use the **GetRowCount** method.

---

### GetTextColor

Retrieves a rebar control's default text color.

```
FUNCTION GetTextColor (BYVAL hRebar AS HWND) AS COLORREF
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

#### Return value

Returns a **COLORREF** value that represent the current default text color.

---

### GetTooltips

Retrieves the handle to any tooltip control associated with the rebar control.

```
FUNCTION GetTooltips (BYVAL hRebar AS HWND) AS HWND
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

#### Return value

Returns the handle to the tooltip control associated with the rebar control, or zero if no tooltip control is associated with the rebar control.

---

### HitTest

Determines which portion of a rebar band is at a given point on the screen, if a rebar band exists at that point.

```
FUNCTION HitTest (BYVAL hRebar AS HWND, BYREF rhbt AS RBHITTESTINFO) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *rhbt* | Pointer to an **RBHITTESTINFO** structure. Before sending the message, the pt member of this structure must be initialized to the point that will be hit tested, in client coordinates. |

#### Return value

Returns the zero-based index of the band at the given point, or -1 if no rebar band was at the point.

---

### IdToIndex

Converts a band identifier to a band index in a rebar control.

```
FUNCTION IdToIndex (BYVAL hRebar AS HWND, BYVAL uBandID AS DWORD) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBandID* | The application-defined identifier of the band in question. This is the value that was passed in the **wID** member of the **REBARBANDINFO** structure when the band was inserted. |

#### Return value

Returns the zero-based band index if successful, or -1 otherwise. If duplicate band identifiers exist, the first one is returned.

---

### InsertBand

Inserts a new band in a rebar control.

```
FUNCTION InsertBand (BYVAL hRebar AS HWND, BYVAL nIndex AS LONG, BYVAL prbbi AS REBARBANDINFOW PTR) AS BOOLEAN
FUNCTION InsertBand (BYVAL hRebar AS HWND, BYVAL nIndex AS LONG, BYREF rbbi AS REBARBANDINFOW) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *nIndex* | Zero-based index of the location where the band will be inserted. If you set this parameter to -1, the control will add the new band at the last location. |
| *prbbi* | Pointer to a **REBARBANDINFO** structure that defines the band to be inserted. You must set the **cbSize** member of this structure to sizeof(REBARBANDINFO) before sending this message. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

---

### MaximizeBand

Resizes a band in a rebar control to either its ideal or largest size.

```
SUB MaximizeBand (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL fIdeal AS BOOLEAN)
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Zero-based index of the band to be maximized. |
| *fIdeal* | Indicates if the ideal width of the band should be used when the band is maximized. If this value is nonzero, the ideal width will be used. If this value is zero, the band will be made as large as possible. |

---

### MinimizeBand

Resizes a band in a rebar control to its smallest size.

```
SUB MinimizeBand (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD)
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Zero-based index of the band to be minimized. |

---


### MoveBand

Moves a band from one index to another.

```
FUNCTION MoveBand (BYVAL hRebar AS HWND, BYVAL nFrom AS DWORD, BYVAL nTo AS DWORD) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *nFrom* | Zero-based index of the band to be moved. |
| *nTo* | Zero-based index of the new band position. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

#### Remarks

This message will most likely change the index of other bands in the rebar control. If a band is moved from index 6 to index 0, all of the bands in between will have their index incremented by one.

*nFrom* must never be greater than the number of bands minus one. The number of bands can be obtained with the *GetBandCount** method.

---

### PushChevron

Sent to a rebar control to programmatically push a chevron.

```
SUB PushChevron (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL iAppValue AS DWORD)
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Zero-based index of the band whose chevron is to be pushed. |
| *iAppValue* | An application defined 32-bit value. It will be passed back to the application as the **lParamNM** member of the **NMREBARCHEVRON** structure that is passed with the **RBN_CHEVRONPUSHED** notification. |

---

### RemoveDarkMode

Removes the rebar dark mode.

```
FUNCTION RemoveDarkMode(BYVAL hRebar AS HWND) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

#### Return value

If this function succeeds, it returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

---

### SetBandInfo

Sets characteristics of an existing band in a rebar control.

```
FUNCTION SetBandInfo (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL prbbi AS REBARBANDINFOW PTR) AS BOOLEAN
FUNCTION SetBandInfo (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYREF rbbi AS REBARBANDINFOW) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Zero-based index of the band to receive the new settings. |
| *prbbi* | Pointer to a **REBARBANDINFO** structure that defines the band to be modified and the new settings. Before sending this message, you must set the **cbSize** member of this structure to the sizeof(REBARBANDINFO) structure. Additionally, you must set the cch member of the **REBARBANDINFO** structure to the size of the **lpText** buffer when **RBBIM_TEXT** is specified. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

---

### SetBandWidth

Sets the width for a docked band.

```
FUNCTION SetBandWidth (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL nWidth AS DWORD) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *uBand* | Index of the band. |
| *nWidth* | New width in pixels. |

#### Return value

Returns TRUE if the value was set and FALSE otherwise.

---

### SetBarInfo

Sets the characteristics of a rebar control.

```
FUNCTION SetBarInfo (BYVAL hRebar AS HWND, BYVAL prbi AS REBARINFO PTR) AS BOOLEAN
FUNCTION SetBarInfo (BYVAL hRebar AS HWND, BYREF rbi AS REBARINFO) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *prbi* | Pointer to a **REBARINFO** structure that contains the information to be set. You must set the cbSize member of this structure to sizeof(REBARINFO) before sending this message. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

---

### SetBkColor

Sets a rebar control's default background color.

```
FUNCTION SetBkColor (BYVAL hRebar AS HWND, BYVAL clrBk AS DWORD) AS DWORD
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *clrBk* | **COLORREF** value that represents the new default background color. |

#### Return value

Returns a **COLORREF** value that represents the previous default background color.

#### Remarks

The rebar control's default background color is used to draw the background in a rebar control and all bands that are added after this message has been sent. The default background color for a particular band can be overridden when a band is added or modified by setting the **RBBIM_COLORS** flag in **fMask** and setting **clrBack** in the **REBARBANDINFO** structure.

When visual styles are enabled, this message has no effect.

---

### SetColorScheme

Sets the color scheme information for the rebar control.

```
FUNCTION SetColorScheme (BYVAL hRebar AS HWND, BYVAL pcs AS COLORSCHEME PTR)
FUNCTION SetColorScheme (BYVAL hRebar AS HWND, BYREF cs AS COLORSCHEME)
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *pcs* | Pointer to a **COLORSCHEME** structure that contains the color scheme information. |

#### Return value

The return value for this message is not used.

#### Remarks

The rebar control uses the color scheme information when drawing the 3-D elements in the control and bands.

---

### SetDarkMode

Sets the color scheme information for the rebar control.

```
FUNCTION SetDarkMode(BYVAL hRebar AS HWND) AS HRESULT
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |

#### Return value

If this function succeeds, it returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

---

### SetPalette

Sets the rebar control's current palette.

```
FUNCTION SetPalette (BYVAL hRebar AS HWND, BYVAL hPal AS HPALETTE) AS HPALETTE
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *hPal* | An **HPALETTE** that specifies the new palette that the rebar control will use. |

#### Return value

Returns an **HPALETTE** that specifies the rebar control's previous palette.

#### Remarks

It is the responsibility of the application sending this message to delete the **HPALETTE** passed in this message (see **DeleteObject**). The rebar control will not delete an HPALETTE set with this message.

---

### SetParent

Sets a rebar control's parent window.

```
FUNCTION SetParent (BYVAL hRebar AS HWND, BYVAL hwndParent AS hWND) AS HWND
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *hwndParent* | Handle to the parent window to be set. |

#### Return value

Returns the handle to the previous parent window, or NULL if there is no previous parent.

### Remarks

The rebar control sends notification messages to the window you specify with this message. This message does not actually change the parent of the rebar control.

---

### SetTextColor

Sets a rebar control's default text color.

```
FUNCTION SetTextColor (BYVAL hRebar AS HWND, BYVAL clrText AS COLORREF) AS COLORREF
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *clrText* | **COLORREF** value that represents the new default text color. |

#### Return value

Returns a **COLORREF** value that represents the previous default text color.

#### Remarks

The rebar control's default text color is used to draw the text in a rebar control and all bands that are added after this message has been sent. The default text color for a particular band can be overridden when a band is added or modified by setting the **RBBIM_COLORS** flag in fMask and setting clrBack in the **REBARBANDINFO** structure.

---

### SetTooltips

Associates a tooltip control with the rebar control.

```
SUB SetTooltips (BYVAL hRebar AS HWND, BYVAL hwndToolTip AS HWND)
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *hwndToolTip* | Handle to the tooltip control to be set. |

---

### SetWindowTheme

Associates a tooltip control with the rebar control.

```
SUB SetWindowTheme (BYVAL hRebar AS HWND, BYVAL pwszTheme AS WSTRING PTR)
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | The handle of the Rebar control. |
| *pwszTheme* | Pointer to a Unicode string that contains the rebar visual style to set. |

---

### ShowBand

Shows or hides a given band in a rebar control.

```
FUNCTION ShowBand (BYVAL hRebar AS HWND, BYVAL uBand AS DWORD, BYVAL fShow AS BOOLEAN) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | Zero-based index of a band in the rebar control. |
| *uBand* | Handle to the tooltip control to be set. |
| *fShow* | Boolean value that indicates if the band should be shown or hidden. If this value is TRUE, the band will be shown. Otherwise, the band will be hidden. |

#### Return value

Returns TRUE if successful, or FALSE otherwise

---

### SizeToRect

Attempts to find the best layout of the bands for the given rectangle.

```
FUNCTION SizeToRect (BYVAL hRebar AS HWND, BYVAL prc AS RECT PTR) AS BOOLEAN
FUNCTION SizeToRect (BYVAL hRebar AS HWND, BYREF rc AS RECT) AS BOOLEAN
FUNCTION SizeToRect (BYVAL hRebar AS HWND) AS RECT
```

| Parameter | Description |
| --------- | ----------- |
| *hRebar* | Zero-based index of a band in the rebar control. |
| *prc* | Pointer to a **RECT** structure that specifies the rectangle to which the rebar control should be sized. |

#### Return value
Returns nonzero if a layout change occurred, or zero otherwise.

#### Remarks

The rebar bands will be arranged and wrapped as necessary to fit the rectangle. Bands that have the **RBBS_VARIABLEHEIGHT** style will be resized as evenly as possible to fit the rectangle. The height of a horizontal rebar or the width of a vertical rebar may change, depending on the new layout.

---

### Example

```
'#RESOURCE "CW_Toolbar_HDPI.rc"
#define UNICODE
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "AfxNova/CWindow.inc"
#INCLUDE ONCE "AfxNova/AfxGdiplus.inc"
#INCLUDE ONCE "AfxNova/CToolbar.inc"
#INCLUDE ONCE "AfxNova/CRebar.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

CONST IDC_TOOLBAR = 1001
CONST IDC_REBAR = 1002
enum
   IDM_LEFTARROW = 28000
   IDM_RIGHTARROW
   IDM_HOME
   IDM_SAVEFILE
end enum

' // Forward reference
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   DIM pWindow AS CWindow
   DIM hWin AS HWND = pWindow.Create(NULL, "CWindow - Toolbar inside a rebar", @WndProc)
   ' // Sets the client size
   pWindow.SetClientSize(400, 220)
   ' // Centers the window
   pWindow.Center
   ' // Set the main window background color
   pWindow.SetBackColor(RGB_OLDLACE)

   ' // Adds a tooolbar
   DIM hToolBar AS HWND = pWindow.AddControl("Toolbar", hWin, IDC_TOOLBAR)
   ' // Allow drop down arrows
   CToolbar.SetExtendedStyle(hToolBar, TBSTYLE_EX_DRAWDDARROWS)

   ' // Calculate the size of the icon according the DPI
   DIM cx AS LONG = 30 * pWindow.DPI \ 96

   ' // Creates an image list for the toolbar
   DIM hNormalImageList AS HIMAGELIST
   hNormalImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 4, 0)
   IF hNormalImageList THEN
      ImageList_ReplaceIcon(hNormalImageList, -1, AfxGdipImageFromRes(hInstance, "IDI_ARROW_LEFT"))
      ImageList_ReplaceIcon(hNormalImageList, -1, AfxGdipImageFromRes(hInstance, "IDI_ARROW_RIGHT"))
      ImageList_ReplaceIcon(hNormalImageList, -1, AfxGdipImageFromRes(hInstance, "IDI_HOME"))
      ImageList_ReplaceIcon(hNormalImageList, -1, AfxGdipImageFromRes(hInstance, "IDI_SAVE"))
      ' // Set the normal image list
      CToolbar.SetImageList(hToolBar, hNormalImageList)
      ' // Set the hot image list with the same images than the normal one
      CToolbar.SetHotImageList(hToolBar, hNormalImageList)
   END IF

   ' // Creates a disabled image list for the toolbar
   DIM hDisabledImageList AS HIMAGELIST
   hDisabledImageList = ImageList_Create(cx, cx, ILC_COLOR32 OR ILC_MASK, 4, 0)
   IF hDisabledImageList THEN
      ImageList_ReplaceIcon(hDisabledImageList, -1, AfxGdipImageFromRes(hInstance, "IDI_ARROW_LEFT", 60, TRUE))
      ImageList_ReplaceIcon(hDisabledImageList, -1, AfxGdipImageFromRes(hInstance, "IDI_ARROW_RIGHT", 60, TRUE))
      ImageList_ReplaceIcon(hDisabledImageList, -1, AfxGdipImageFromRes(hInstance, "IDI_HOME", 60, TRUE))
      ImageList_ReplaceIcon(hDisabledImageList, -1, AfxGdipImageFromRes(hInstance, "IDI_SAVE", 60, TRUE))
      ' // Set the disabled image list
      CToolbar.SetDisabledImageList(hToolBar, hDisabledImageList)
   END IF

   ' // Adds buttons to the toolbar
   CToolbar.AddButton(hToolBar, 0, IDM_LEFTARROW)
   CToolbar.AddButton(hToolBar, 1, IDM_RIGHTARROW, 0, BTNS_DROPDOWN)
   CToolbar.AddButton(hToolBar, 2, IDM_HOME)
   CToolbar.AddButton(hToolBar, 3, IDM_SAVEFILE)

   ' // Disables the save file button
   CToolbar.DisableButton(hToolBar, IDM_SAVEFILE)

   ' // Sizes the toolbar
   CToolbar.AutoSize(hToolBar)
   ' // Anchors the toolbar
   pWindow.AnchorControl(IDC_TOOLBAR, AFX_ANCHOR_WIDTH)

   ' // Create a rebar control
   DIM hRebar AS HWND = pWindow.AddControl("Rebar", , IDC_REBAR)

   ' // Add the band containing the toolbar control to the rebar
   ' // The size of the REBARBANDINFOW is different in Vista/Windows 7
   DIM rc AS RECT, wszText AS WSTRING * 260
   DIM rbbi AS REBARBANDINFOW
   IF AfxWindowsVersion >= 600 AND AfxComCtlVersion >= 600 THEN
      rbbi.cbSize  = REBARBANDINFO_V6_SIZE
   ELSE
      rbbi.cbSize  = REBARBANDINFO_V3_SIZE
   END IF
   ' // Insert the toolbar in the rebar control
   rbbi.fMask      = RBBIM_STYLE OR RBBIM_CHILD OR RBBIM_CHILDSIZE OR _
                     RBBIM_SIZE OR RBBIM_ID OR RBBIM_IDEALSIZE OR RBBIM_TEXT
   rbbi.fStyle     = RBBS_CHILDEDGE
   rbbi.hwndChild  = hToolbar
   rbbi.cxMinChild = 200 * pWindow.rxRatio
   rbbi.cyMinChild = CToolbar.GetButtonHeight(hToolbar)
   rbbi.cx         = 200 * pWindow.rxRatio
   rbbi.cxIdeal    = 200 * pWindow.rxRatio
   wszText = "Toolbar"
   rbbi.lpText = @wszText
   '// Insert band into rebar
   CRebar.InsertBand(hRebar, -1, @rbbi)
   ' // Anchor the rebar
   pWindow.AnchorControl(IDC_REBAR, AFX_ANCHOR_WIDTH)

   ' // Adds a cancel button
   pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", 270, 155, 75, 30)
   ' // Anchors the button to the bottom and the right side of the main window
   pWindow.AnchorControl(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

   ' // Processes event messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_CREATE
         AfxEnableDarkModeForWindow(hwnd)
         EXIT FUNCTION

      ' // Theme has changed
      CASE WM_THEMECHANGED
         AfxEnableDarkModeForWindow(hwnd)
         RETURN 0

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
               END IF
            CASE IDM_LEFTARROW
               MessageBoxW hwnd, "You have clicked the left arrow button", "Toolbar", MB_OK
            CASE IDM_RIGHTARROW
               MessageBoxW hwnd, "You have clicked the right arrow button", "Toolbar", MB_OK
            CASE IDM_HOME
               MessageBoxW hwnd, "You have clicked the home button", "Toolbar", MB_OK
            CASE IDM_SAVEFILE
               MessageBoxW hwnd, "You have clicked the save button", "Toolbar", MB_OK
         END SELECT

      CASE WM_NOTIFY
         ' -------------------------------------------------------
         ' Notification messages are handled here.
         ' The TTN_GETDISPINFO message is sent by a ToolTip control
         ' to retrieve information needed to display a ToolTip window.
         ' ------------------------------------------------------
         DIM ptnmhdr AS NMHDR PTR              ' // Information about a notification message
         DIM ptttdi AS NMTTDISPINFOW PTR       ' // Tooltip notification message information
         DIM wszTooltipText AS WSTRING * 260   ' // Tooltip text
         ptnmhdr = CAST(NMHDR PTR, lParam)
         SELECT CASE ptnmhdr->code
            CASE TTN_GETDISPINFOW
               ptttdi = CAST(NMTTDISPINFOW PTR, lParam)
               ptttdi->hinst = NULL
               wszTooltipText = ""
               SELECT CASE ptttdi->hdr.hwndFrom
                  CASE SendMessageW(GetDlgItem(hwnd, IDC_TOOLBAR), TB_GETTOOLTIPS, 0, 0)
                     SELECT CASE ptttdi->hdr.idFrom
                        CASE IDM_LEFTARROW  : wszTooltipText = "Back"
                        CASE IDM_RIGHTARROW : wszTooltipText = "Forward"
                        CASE IDM_HOME       : wszTooltipText = "Home"
                        CASE IDM_SAVEFILE   : wszTooltipText = "Save"
                     END SELECT
                     IF LEN(wszTooltipText) THEN ptttdi->lpszText = @wszTooltipText
               END SELECT
         CASE TBN_DROPDOWN   ' // dropdown menu
            DIM ptbn AS TBNOTIFY PTR = CAST(TBNOTIFY PTR, lParam)
            SELECT CASE ptbn->iItem
               CASE IDM_RIGHTARROW
                  DIM rc AS RECT
                  SendMessageW(ptbn->hdr.hwndFrom, TB_GETRECT, ptbn->iItem, CAST(LPARAM, @rc))
                  MapWindowPoints(ptbn->hdr.hwndFrom, HWND_DESKTOP, CAST(LPPOINT, @rc), 2)
                  DIM hPopupMenu AS HMENU = CreatePopUpMenu
                  AppendMenuW hPopupMenu, MF_ENABLED, 10001, "Option 1"
                  AppendMenuW hPopupMenu, MF_ENABLED, 10002, "Option 2"
                  AppendMenuW hPopupMenu, MF_ENABLED, 10003, "Option 3"
                  AppendMenuW hPopupMenu, MF_ENABLED, 10004, "Option 4"
                  AppendMenuW hPopupMenu, MF_ENABLED, 10005, "Option 5"
                  DIM flags AS DWORD = TPM_LEFTBUTTON OR TPM_NONOTIFY OR TPM_RETURNCMD OR TPM_NOANIMATION
                  DIM id AS LONG = TrackPopupMenu(hPopupMenu, flags, rc.Left, rc.Bottom, 0, hwnd, NULL)
                  SELECT CASE id
                     CASE 10001
                        MessageBoxW(hwnd, "You clicked item 1", "Message", MB_OK)
                     CASE 10002
                        MessageBoxW(hwnd, "You clicked item 2", "Message", MB_OK)
                     CASE 10003
                        MessageBoxW(hwnd, "You clicked item 3", "Message", MB_OK)
                     CASE 10004
                        MessageBoxW(hwnd, "You clicked item 4", "Message", MB_OK)
                     CASE 10005
                        MessageBoxW(hwnd, "You clicked item 5", "Message", MB_OK)
                  END SELECT
                  DestroyMenu hPopupMenu
            END SELECT
         END SELECT

    	CASE WM_DESTROY
         ' // Destroy the image lists
         ImageList_Destroy CAST(HIMAGELIST, SendMessageW(GetDlgItem(hwnd, IDC_TOOLBAR), TB_SETIMAGELIST, 0, 0))
         ImageList_Destroy CAST(HIMAGELIST, SendMessageW(GetDlgItem(hwnd, IDC_TOOLBAR), TB_SETHOTIMAGELIST, 0, 0))
         ImageList_Destroy CAST(HIMAGELIST, SendMessageW(GetDlgItem(hwnd, IDC_TOOLBAR), TB_SETDISABLEDIMAGELIST, 0, 0))
         ' // Quit the application
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
```
---
