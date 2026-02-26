# CStatic Class

Applications often use static controls to label other controls or to separate a group of controls. Although static controls are child windows, they cannot be selected. Therefore, they cannot receive the keyboard focus and cannot have a keyboard interface. A static control that has the SS_NOTIFY style receives mouse input, notifying the parent window when the user clicks or double clicks the control. Static controls belong to the STATIC window class.

Although static controls can be used in overlapped, pop-up, and child windows, they are designed for use in dialog boxes, where the system standardizes their behavior. By using static controls outside dialog boxes, a developer increases the risk that the application might behave in a nonstandard fashion. Typically, a developer either uses static controls in dialog boxes or uses the SS_OWNERDRAW style to create customized static controls.

#### Static control types

There are four types of static controls. Each type has one or more Static Control Styles.

* Simple Graphics Static Control
* Text Static Control
* Image Static Control
* Owner-Drawn Static Control

#### Simple Graphics Static Control

A simple graphics static control displays a frame or a filled rectangle. A frame can be drawn in a number of styles, included black, gray, or white. In addition, a frame can be drawn with an etched style to give it a three-dimensional appearance. The frame styles include SS_BLACKFRAME, SS_GRAYFRAME, SS_WHITEFRAME, SS_ETCHEDHORZ, SS_ETCHEDVERT, and SS_ETCHEDFRAME.

A rectangle can be filled with color in one of three styles: black, gray, or white. These styles are defined by the constants SS_BLACKRECT, SS_GRAYRECT, and SS_WHITERECT.

The graphics styles cannot be combined.

#### Text Static Control

A text static control displays text in a rectangle in one of five styles:

* left-aligned without word wrap
* left-aligned with word wrap
* centered
* right-aligned
* simple

These styles are defined by the constants SS_LEFTNOWORDWRAP, SS_LEFT, SS_CENTER, SS_RIGHT, and SS_SIMPLE, respectively. The system rearranges the text in these controls in predefined ways, except for "simple" text, which is not rearranged.

An application can change the text in a text static control at any time by using the SetWindowText function or the WM_SETTEXT message.

The system displays as much text as it can in the static control and clips whatever does not fit. To calculate an appropriate size for the control, retrieve the font metrics for the text. For more information about fonts and font metrics, see [Fonts and Text](https://learn.microsoft.com/en-us/windows/win32/gdi/fonts-and-text).

By default, the window text for a static control, as for other controls, can contain an ampersand that defines the following character as the shortcut key for the control (or, in the case of most static controls, for the control that it labels, which is the next control in the tab order). If you wish to display ampersands in the text rather than using them to define shortcuts, include the SS_NOPREFIX style.

#### Image Static Control

An image static control can display bitmaps, icons (including animated icons), or enhanced metafiles. The type of graphic that a particular static control displays depends on the control's style: SS_BITMAP, SS_ICON, or SS_ENHMETAFILE. An application specifies the style when it creates the control and also specifies a handle to the bitmap, icon, or metafile for the control to display. After the control is created, an application can associate a different graphic with the control by sending it an STM_SETIMAGE message, specifying a handle to the new graphic object. An application can retrieve a handle to the graphic object currently associated with a static control by sending it an STM_GETIMAGE message. An application sends messages to a static control by using the SendDlgItemMessage function.

#### Owner-Drawn Static Control

By using the SS_OWNERDRAW style, an application can take responsibility for painting a static control. The parent window of an owner-drawn static control (its owner) receives a WM_DRAWITEM message whenever the static control needs to be painted. The message includes a pointer to a DRAWITEMSTRUCT structure that contains information that the owner window uses when drawing the control.

#### Static Control Default Message Processing

The window procedure for the predefined static control window class performs default processing for all messages that the static control procedure does not process. When the static control returns FALSE for any message, the predefined window procedure checks the messages and carries out the default action described in the following table. In the table, a text static control is a static control with the style SS_LEFTNOWORDWRAP, SS_LEFT, SS_CENTER, SS_RIGHT, or SS_SIMPLE.

| Message | Default action |
| ------- | -------------- |
| **WM_CREATE** | Loads the graphic object and sizes the window to the object's size, for graphic static controls. Takes no action for other static controls. |
| **WM_DESTROY** | Frees and destroys any graphic object, for graphic static controls. Takes no action for other static controls. |
| **WM_ENABLE** | Repaints visible static controls. |
| **WM_ERASEBKGND** | Returns TRUE, indicating the control erases the background. |
| **WM_GETDLGCODE** | Returns DLGC_STATIC. |
| **WM_GETFONT** | Returns a handle to the font for text static controls. |
| **WM_GETTEXT** | Returns the number of characters copied. |
| **WM_GETTEXTLENGTH** | Returns the length, in characters, of the text for a text static control. |
| **WM_LBUTTONDBLCLK** | Sends the parent window an STN_DBLCLK notification code if the control style is SS_NOTIFY. |
| **WM_LBUTTONDOWN** | Sends the parent window an STN_CLICKED notification code if the control style is SS_NOTIFY. |
| **WM_NCLBUTTONDBLCLK** | Sends the parent window an STN_DBLCLK notification code if the control style is SS_NOTIFY. |
| **WM_NCLBUTTONDOWN** | Sends the parent window an STN_CLICKED notification code if the control style is SS_NOTIFY. |
| **WM_NCHITTEST** | Returns HTCLIENT if the control style is SS_NOTIFY; otherwise, returns HTTRANSPARENT. |
| **WM_PAINT** | Repaints the control. |
| **WM_SETFONT** | Sets the font and repaints for text static controls. |
| **WM_SETTEXT** | Sets the text and repaints for text static controls. |

The predefined window procedure passes all other messages to DefWindowProc for default processing.

---

| Name | Description |
| ---- | ----------- |
| [DeleteBitmap](#deletebitmap) | Deletes a bitmap associated with a static control. |
| [DeleteCursor](#deletecursor) | Deletes a cursor associated with a static control. |
| [DeleteEnhancedMetafile](#deleteenhancedmetafile) | Deletes an enhanced metafile associated with a static control. |
| [DeleteIcon](#deleteicon) | Deletes an icon associated with a static control. |
| [DeleteImage](#deleteimage) | Deletes an image associated with a static control. |
| [Disable](#disable) | Disables the control. |
| [Enable](#enable) | Enables the control. |
| [GetIcon](#geticon) | Retrieves a handle to the icon associated with a static control that has the SS_ICON style. |
| [GetImage](#getimage) | Retrieves a handle to the image (icon or bitmap) associated with a static control. |
| [GetText](#getext) | Gets the text of a static control. |
| [GetTextLength](#getextlength) | Gets the length of the text of a static control. |
| [SetBitmap](#setbitmap) | Associates a bitmap with a static control. |
| [SetCursor](#setcursor) | Associates a cursor with a static control. |
| [SetEnhancedMetafile](#setenhancedmetafile) | Associates an enhanced metafile with a static control. |
| [SetIcon](#seticon) | Associates an icon with an static control. |
| [SetImage](#setimage) | Associates a new image with a static control. |
| [SetText](#settext) | Sets the text of a static control. |

---

### DeleteBitmap

Deletes a bitmap associated with a static control. Returns TRUE o FALSE.

```
FUNCTION DeleteBitmap (BYVAL hStatic AS HWND) AS BOOLEAN
```

#### Usage example

```
CStatic.DeleteBitmap(hStatic)
```
---

### DeleteCursor

Deletes a cursor associated with a static control. Returns TRUE or FALSE.

```
FUNCTION DeleteCursor (BYVAL hStatic AS HWND) AS BOOLEAN
```

#### Usage example

```
CStatic.DeleteCursor(hStatic)
```
---

### DeleteEnhancedMetafile

Deletes an enhanced metafile associated with a static control. Returns TRUE or FALSE.

```
FUNCTION DeleteEnhancedMetafile (BYVAL hStatic AS HWND) AS BOOLEAN
```

#### Usage example

```
CStatic.DeleteEnhancedMetafile(hStatic)
```
---

### DeleteIcon

Deletes an icon associated with a static control. Returns TRUE o FALSE.

```
FUNCTION DeleteIcon (BYVAL hStatic AS HWND) AS BOOLEAN
```

#### Usage example

```
CStatic.DeleteIcon(hStatic)
```
---

### DeleteImage

Deletes an image associated with a static control. Returns TRUE or FALSE.

```
FUNCTION DeleteImage (BYVAL hStatic AS HWND) AS BOOLEAN
```

#### Usage example

```
CStatic.DeleteImage(hStatic)
```
---

### Disable

Disables the control. Returns False if the control was previous disabled; otherwise TRUE

```
FUNCTION Disable (BYVAL hStatic AS HWND) AS BOOLEAN
```

#### Usage example

```
CStatic.Disable(hStatic)
```
---

### Enable

Enables the control. Returns False if the control was previous enabled; otherwise TRUE

```
FUNCTION Enable (BYVAL hStatic AS HWND) AS BOOLEAN
```

#### Usage example

```
CStatic.Enable(hStatic)
```
---

### GetIcon

Retrieves a handle to the icon associated with a static control that has the SS_ICON style.

```
FUNCTION GetIcon (BYVAL hStatic AS HWND) AS HICON
```

#### Usage example

```
DIM hIcon AS HICON = CStatic.GetIcon(hStatic)
```
---

### GetImage

Retrieves a handle to the image (icon or bitmap) associated with a static control. The return value is a handle to the image, if any; otherwise, it is NULL.

```
FUNCTION GetImage (BYVAL hStatic AS HWND) AS HICON
```

#### Usage example

```
DIM hImage AS HANDLE = CStatic.GetIcon(hStatic)
```
---

### GetText

Gets the text of a static control.

```
FUNCTION GetText (BYVAL hStatic AS HWND) AS DWSTRING
```

#### Usage example

```
DIM dwsText AS DWSTRING = CStatic.GetText(hStatic)
```
---

### GetTextLength

Gets the length of the text of a static control.

```
FUNCTION GetTextLength (BYVAL hStatic AS HWND) AS LONG
```

#### Usage example

```
DIM nLen AS LONG = CStatic.GetTextLength(hStatic)
```
---

### SetBitmap

Associates a bitmap with a static control. The return value is a handle to the image previously associated with the static control, if any; otherwise, it is NULL.

```
FUNCTION SetBitmap (BYVAL hStatic AS HWND, BYVAL hBmp AS HBITMAP) AS HBITMAP
```

| Parameter | Description |
| ------- | -------------- |
| *hBmp* | Handle of the bitmap to set. |

#### Usage example

```
DIM hOldBmp AS HBITMAP = CStatic.SetBitmap(hStatic, hBmp)
```
---

### SetCursor

Associates a cursor with a static control. The return value is a handle to the cursor previously associated with the static control.

```
FUNCTION SetCursor (BYVAL hStatic AS HWND, BYVAL hCur AS HCURSOR) AS HCURSOR
```

| Parameter | Description |
| ------- | -------------- |
| *hCur* | Handle of the cursor to set. |

#### Usage example

```
DIM hOldCur AS HCURSOR = CStatic.SetCursor(hStatic, hCur)
```
---

### SetEnhancedMetafile

Associates an enhanced metafile with a static control. The return value is a handle to the image previously associated with the static control, if any; otherwise, it is NULL.

```
FUNCTION SetEnhancedMetafile (BYVAL hStatic AS HWND, BYVAL hEMF AS HENHMETAFILE) AS HCURSOR
```

| Parameter | Description |
| ------- | -------------- |
| *hEMF* | Handle of the enhanced metafile to set. |

#### Usage example

```
DIM hOldEMF AS HENHMETAFILE = CStatic.SetCursor(hStatic, hEMF)
```
---

### SetIcon

Associates an icon with an static control. The return value is a handle to the icon previously associated with the static control, if any; otherwise, it is NULL.

```
FUNCTION SetIcon (BYVAL hStatic AS HWND, BYVAL hIcon AS HICON) AS HICON
```

| Parameter | Description |
| ------- | -------------- |
| *hIcon* | Handle of the icon to set. |

#### Usage example

```
DIM hOldIcon AS HICON = CStatic.SetIcon(hStatic, hIcon)
```
---

### SetImage

Associates a new image with a static control. The return value is a handle to the image previously associated with the static control, if any; otherwise, it is NULL.

```
FUNCTION SetIcon (BYVAL hStatic AS HWND, BYVAL nType AS LONG, BYVAL hImage AS HANDLE) AS HANDLE
```

| Parameter | Description |
| ------- | -------------- |
| *hStatic* | Handle of the static control. |
| *nType* | Type of the image to set. |
| *hImage* | Handle of the image to set. |

To associate an image with a static control, the control must have the proper style. The following table shows the style needed for each image type.

| Image type | Static control type |
| ---------- | ------------------- |
| **IMAGE_BITMAP** | SS_BITMAP |
| **IMAGE_CURSOR** | SS_ICON |
| **IMAGE_ENHMETAFILE** | SS_ENHMETAFILE |
| **IMAGE_ICON** | SS_ICON |

#### Usage example

```
DIM hOldBmp AS HICON = CStatic.SetImage(hStatic, IMAGE_BITMAP, hBmp)
DIM hOldIcon AS HICON = CStatic.SetImage(hStatic, IMAGE_ICON, hIcon)
DIM hOldCur AS HCURSOR = CStatic.SetImage(hStatic, IMAGE_CURSOR, hCur)
DIM hOldEMF AS HCURSOR = CStatic.SetImage(hStatic, IMAGE_ENHMETAFILE, hEMF)
```
---

### SetText

Sets the text of a static control.

```
FUNCTION SetText (BYVAL hStatic AS HWND, BYVAL pwszText AS WSTRING PTR) AS BOOLEAN
```

| Parameter | Description |
| ------- | -------------- |
| *hStatic* | Handle of the static control. |
| *pwszText* | The text to set. |

#### Usage example

```
CStatic.SetText(hStatic, "New text")
```
---
