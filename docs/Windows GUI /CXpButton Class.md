# CXpButton Class

Theme aware push button that can also act as a toggle button and display an icon or bitmap.

* Draws a themed button using the functions available in UxTheme.dll or a button with classic appearance using GDI functions. 
* Allows up to three images per button (for normal, hot and disabled states). If you want a button with a centered image and no text, use the XPBI_CENTERCENTER flag in the **ImagePos** method.
* Can be used as a toggle button setting this flag with the **Toggle** method.
* Images and fonts are owned and destroyed by the control.
* If theming is disabled, you can use your own colors for the different elements of the button: background, foreground, text, etc.

### Constructor

Registers the class name and creates an instance of the control.

```
CONSTRUCTOR CXpButton (BYVAL pWindow AS CWindow PTR, BYVAL cID AS LONG_PTR, _
   BYREF wszTitle AS WSTRING = "", BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, _
   BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, BYVAL dwStyle AS DWORD = 0, _
   BYVAL dwExStyle AS DWORD = 0, BYVAL lpParam AS LONG_PTR = 0)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pWindow* | Pointer to the *CWindow* class of the parent window. |
| *cID* | Control identifier. |
| *wszTitle* | Optional. The button caption. |
| *x* | Optional. The x-coordinate of the upper-left corner of the button relative to the upper-left corner of the parent window's client area. |
| *y* | Optional. The initial y-coordinate of the upper-left corner of the button relative to the upper-left corner of the parent window's client area. |
| *nWidth* | Optional. The width of the button. |
| *nHeight* | Optional. The height of the button. |
| *dwStyle* | Optional. The style of the button being created.<br>Default styles: WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER. |
| *dwExStyle* | Optional. The extended window style of the button being created. |
| *lpParam* | Optional. Pointer to custom data. |

#### Return value

A pointer to the new instance of the class.

### Helper Procedure: AfxCXpButtonPtr

Returns a pointer to the CXpButton class given the handle of its associated window.

```
FUNCTION AfxCXpButtonPtr (BYVAL hwnd AS HWND) AS CXpButton PTR
FUNCTION AfxCXpButtonPtr (BYVAL hParent AS HWND, BYVAL cID AS LONG) AS CXpButton PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window associated with the button control. Call the **hWindow** method of the **CXpButton** class to retrieve it. |
| *hParent* | The handle of the parent window of the control. |
| *cID* | The identifier of the control. |

#### Return value

A pointer to the CXpButton class.

### Methods

| Parameter  | Description |
| ---------- | ----------- |
| [DisableTheming](#disabletheming) | Disables theming. |
| [EnableTheming](#enabletheming) | Enables theming. |
| [GetImage](#getimage) | Returns the handle of the image. |
| [hWindow](#hwindow) | Returns the handle of the button. |
| [IsThemed](#isthemed) | Returns CTRUE if themes are enabled or FALSE otherwise. |
| [Redraw](#redraw) | Redraws the button. |
| [SetBitmap](#setbitmap) | Sets the bitmap for the button. |
| [SetBitmapFromFile](#setbitmapfromfile) | Loads an bitmap from file and sets it as the image of the button. |
| [SetIcon](#seticon) | Sets the icon for the button. |
| [SetIconFromFile](#seticonfromfile) | Loads an icon from file and sets it as the image of the button. |
| [SetImage](#setimage) | Sets the image for the button. |
| [SetImageFromFile](#setimagefromfile) | Loads an image from file and sets it as the image of the button. |
| [SetImageFromRes](#setimagefromres) | Loads an image from a resource file and sets it as the image of the button. |

### Properties

| Parameter  | Description |
| ---------- | ----------- |
| [AllowFocusRect](#allowfocusrect) | Gets/sets whether to display the button's focus rectangle. |
| [BkBrush](#bkbrush) | Returns the background color brush. |
| [ButtonBkColor](#buttonbkcolor) | Gets/sets the background color of button. |
| [ButtonBkColorDown](#buttonbkcolordown) | Gets/sets the background color of button when it is down (pressed or toggled). |
| [ButtonBkColorHot](#buttonbkcolorhot) | Gets/sets the background color of button when it is hot (the mouse is over it). |
| [ButtonState](#buttonstate) | Returns the button state. |
| [Cursor](#cursor) | Gets/sets the cursor for the button. |
| [DisabledImageHandle](#disabledimagehandle) | Returns the handle of the disabled image. |
| [Font](#font) | Gets/sets the handle of the font used by the button. |
| [HotImageHandle](#hotimagehandle) | Returns the handle of the hot image. |
| [ImageHeight](#imageheight) | Gets/sets the height of the image. |
| [ImageMargin](#imagemargin) | Gets/sets the image margin. |
| [ImagePos](#imagepos) | Gets/sets the image position. |
| [ImageType](#imagetype) | Returns the image type |
| [ImageWidth](#imagewidth) | Gets/sets the width of the image. |
| [NormalImageHandle](#normalimagehandle) | Returns the handle of the normal image. |
| [TextBkColor](#textbkcolor) | Gets/sets the text background color of the button. |
| [TextBkColorDown](#textbkcolordown) | Gets/sets the text background color of the button when it is down (pressed). |
| [TextForeColor](#textforecolor) | Gets/sets the text foreground color of the button. |
| [TextForeColorDown](#textforecolordown) | Gets/sets the text foreground color of the button when it is down (pressed). |
| [TextForeColorHot](#textforecolorhot) | Gets/sets the text foreground hot color of the button. |
| [TextFormat](#textformat) | Gets/sets the method of formatting the text. |
| [TextMargin](#textmargin) | Gets/sets the margin of the text of the button. |
| [Toggle](#toggle) | Gets/sets button to toggle state (CTRUE) or to pushbutton state (FALSE). |
| [ToggleState](#togglestate) | Gets/sets the toggle state: pushed (CTRUE) or unpushed (FALSE). |

### Example

```
#define UNICODE
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/CXpButton.inc"
USING Afx

CONST IDC_BUTTON1 = 1001
CONST IDC_BUTTON2 = 1002
CONST IDC_BUTTON3 = 1003

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwsszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG

   END wWinMain(GetModuleHandleW(NULL), NULL, WCOMMAND(), SW_NORMAL)

' // Forward declaration
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

   ' // Create the main window
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "XpButton example", @WndProc)
   pWindow.SetClientSize(215, 142)
   pWindow.Center

   ' // Create the first button
   DIM pXpButton1 AS CXpButton = CXpButton(@pWindow, IDC_BUTTON1, "&Hot", 50, 10, 114, 34)
   ' // Load a png icon from file (remember to change the path of the file)
   pXpButton1.SetImageFromFile(ExePath & "/arrow_left_64.png", XPBI_NORMAL)
   ' // Load a Windows predefined icon as the "hot" image
   pXpButton1.SetIcon LoadIcon(NULL, IDI_QUESTION), XPBI_HOT
   ' // Change the cursor shape
   pXpButton1.SetCursor LoadCursor(NULL, IDC_CROSS)
   ' // Set a margin of 10 dpi units (pixels x scaling ratio)
   pXpButton1.SetImageMargin 10

   ' // Create the second button
   DIM pXpButton2 AS CXpButton = CXpButton(@pWindow, IDC_BUTTON2, "&Cancel", 50, 50, 114, 34)
   ' // Load a Windows predefined icon
   pXpButton2.SetIcon LoadIcon(NULL, IDI_ERROR), XPBI_NORMAL
   pXpButton2.SetImagePos XPBI_RIGHT OR XPBI_VCENTER
   pXpButton2.SetTextFormat DT_LEFT OR DT_VCENTER OR DT_SINGLELINE
'   EnableWindow pXpButton2.hWindow, FALSE   ' Disable the button

   ' // Create a coloured button
   DIM pXpButton3 AS CXpButton = CXpButton(@pWindow, IDC_BUTTON3, "&Coloured Button", 50, 90, 114, 34, WS_VISIBLE OR WS_TABSTOP OR BS_PUSHBUTTON OR BS_CENTER OR BS_VCENTER OR BS_FLAT)
   ' // Disable theming
   pXpButton3.DisableTheming
   ' // Set the background color to blue
   pXpButton3.SetButtonBkColor BGR(0, 0, 255)
   ' // Set the background color when pressed to yellow
   pXpButton3.SetButtonBkColorDown BGR(255, 255, 0)
   ' // Set the text foreground color to yellow
   pXpButton3.SetTextForeColor BGR(255, 255, 0)
   ' // Set the text background color to blue
   pXpButton3.SetTextBkColor BGR(0, 0, 255)
   ' // Set the text foreground color when pressed to white
   pXpButton3.SetTextForeColorDown BGR(255, 255, 255)
   ' // Set the text background color when pressed to red
   pXpButton3.SetTextBkColorDown BGR(255, 0, 0)
   ' // Set the text background color when pressed to red
   pXpButton3.SetTextBkColorDown BGR(255, 0, 0)
   ' // Set the text format (centered and single line)
   pXpButton3.TextFormat = DT_CENTER OR DT_VCENTER OR DT_SINGLELINE
' pXpButton3.SetToggle TRUE   ' make it a toggle button

   ' // Set the focus in the first button
   SetFocus pXpButton1.hWindow

   ' // Dispatch Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window callback procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDC_BUTTON1, IDC_BUTTON2, IDC_BUTTON3
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  ' AfxMsg "Button clicked"
                  EXIT FUNCTION
               END IF

            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_DESTROY
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
```

## DisableTheming

Disables theming.

```
SUB DisableTheming
```
---

## EnableTheming

Enables theming.

```
SUB EnableTheming
```
---

## GetImage

Returns the handle of the image.

```
FUNCTION GetImage (BYVAL ImageState AS LONG) AS HANDLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *ImageState* | One of the following values:<br>XPBI_NORMAL = 1, XPBI_HOT = 2, XPBI_DISABLED = 3 |

#### Return value

Returns the handle of the requested image.

---

## <hWindow

Returns the handle of the button.

```
FUNCTION hWindow () AS HWND
```
---

## IsThemed

Returns CTRUE if themes are enabled or FALSE otherwise.

```
FUNCTION IsThemed () AS LONG
```
---

## Redraw

Redraws the button.

```
SUB Redraw
```
---

## SetBitmap

Sets the bitmap for the button.

```
SUB SetBitmap (BYVAL hBitmap AS HBITMAP, _
   BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hBitmap* | Handle to the bitmap. |
| *ImageState* | One of the following values:<br>XPBI_NORMAL = 1, XPBI_HOT = 2, XPBI_DISABLED = 3 |
| *fRedraw* | Optional. CTRUE or FALSE (redraws the button to reflect the changes). |

---

## SetBitmapFromFile

Loads a bitmap from file and sets it as the image of the button.

```
SUB SetBitmapFromFile (BYVAL pwszPath AS WSTRING PTR, _
   BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszPath* | Full path of the bitmap's file. |
| *ImageState* | One of the following values:<br>XPBI_NORMAL = 1, XPBI_HOT = 2, XPBI_DISABLED = 3 |
| *fRedraw* | Optional. CTRUE or FALSE (redraws the button to reflect the changes). |

---

## SetIcon

Sets the icon for the button.

```
SUB SetIcon (BYVAL hIcon AS HICON, _
   BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hIcon* | Handle to the icon. |
| *ImageState* | One of the following values:<br>XPBI_NORMAL = 1, XPBI_HOT = 2, XPBI_DISABLED = 3 |
| *fRedraw* | Optional. CTRUE or FALSE (redraws the button to reflect the changes). |

---

## SetIconFromFile

Loads an icon from file and sets it as the image of the button.

```
SUB SetIconFromFile (BYVAL pwszPath AS WSTRING PTR, _
   BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszPath* | Full path of the icon's file. |
| *ImageState* | One of the following values:<br>XPBI_NORMAL = 1, XPBI_HOT = 2, XPBI_DISABLED = 3 |
| *fRedraw* | Optional. CTRUE or FALSE (redraws the button to reflect the changes). |

---

## SetImage

Sets the image for the button.

```
SUB SetImage (BYVAL hImage AS HANDLE, BYVAL ImageType AS LONG, _
   BYVAL ImageState AS LONG, BYVAL fRedraw AS LONG = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hImage* | Handle to the image. |
| *ImageType* | IMAGE_ICON or IMAGE_BITMAP. |
| *ImageState* | One of the following values:<br>XPBI_NORMAL = 1, XPBI_HOT = 2, XPBI_DISABLED = 3 |
| *fRedraw* | Optional. CTRUE or FALSE (redraws the button to reflect the changes). |

---

## SetImageFromFile

Loads an image from file and sets it as the image of the button.

```
SUB SetImageFromFile (BYREF wszPath AS WSTRING, BYVAL ImageState AS LONG, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL fRedraw AS LONG = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | Full path of the image's file. |
| *ImageState* | One of the following values:<br>XPBI_NORMAL = 1, XPBI_HOT = 2, XPBI_DISABLED = 3 |
| *dimPercent* | Percent of dimming (1-99) |
| *bGrayScale* | CTRUE or FALSE. Convert to gray scale. |
| *fRedraw* | Optional. CTRUE or FALSE (redraws the button to reflect the changes). |

---

## SetImageFromRes

Loads an image from file and sets it as the image of the button.

```
SUB SetImageFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL ImageState AS LONG, BYVAL dimPercent AS LONG = 0, _
   BYVAL bGrayScale AS LONG = FALSE, BYVAL fRedraw AS LONG = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszImageName* | Name of the image in the resource file (.RES). If the image resource uses an integral identifier, wszImage should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff. |
| *ImageState* | One of the following values:<br>XPBI_NORMAL = 1, XPBI_HOT = 2, XPBI_DISABLED = 3 |
| *dimPercent* | Percent of dimming (1-99) |
| *bGrayScale* | CTRUE or FALSE. Convert to gray scale. |
| *fRedraw* | Optional. CTRUE or FALSE (redraws the button to reflect the changes). |

---

## AllowFocusRect

Gets/sets whether to display the button's focus rectangle.

```
PROPERTY AllowFocusRect () AS BOOLEAN
PROPERTY AllowFocusRect (BYVAL bAllowFocusRect AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bAllowFocusRect* | BOOLEAN. TRUE or FALSE. |

#### Return value

TRUE or FALSE.

---

## BkBrush

Returns the background color brush.

```
PROPERTY BkBrush () AS HBRUSH
```
---

## ButtonBkColor

Gets/sets the background color of button.

```
PROPERTY ButtonBkColor () AS COLORREF
PROPERTY ButtonBkColor (BYVAL bkColor AS COLORREF)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bkColor* | A COLORREF color value. Use the FreeBasic BGR function. |

#### Return value

The background color of the button as a COLORREF value.

#### Remarks

Only available if the button is not themed. Call **DisableTheming** to disable theming.

---

## ButtonBkColorDown

Gets/sets the background color of button when it is down (pressed or toggled).

```
PROPERTY ButtonBkColorDown () AS COLORREF
PROPERTY ButtonBkColorDown (BYVAL bkColorDown AS COLORREF)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bkColorDown* | A COLORREF color value. Use the FreeBasic BGR function. |

#### Return value

The background color of button when it is down (pressed or toggled).

#### Remarks

Only available if the button is not themed. Call **DisableTheming** to disable theming.

---

## ButtonBkColorHot

Gets/sets the background color of button when it is hot (the mouse is over it).

```
PROPERTY ButtonBkColorHot () AS COLORREF
PROPERTY ButtonBkColorHot (BYVAL bkColorHot AS COLORREF)
```

| Parameter  | Description |
| ---------- | ----------- |
| *bkColorHot* | A COLORREF color value. Use the FreeBasic BGR function. |

#### Return value

The background color of button when it is hot (the mouse is over it).

#### Remarks

Only available if the button is not themed. Call **DisableTheming** to disable theming.

---

## ButtonState

Returns the button state.

```
PROPERTY ButtonState () AS LONG
```

#### Return value

The button state: Pushed (BST_PUSHED) or unpushed.

---

## Cursor

Gets/sets the cursor for the button.

```
PROPERTY Cursor () AS HCURSOR
PROPERTY Cursor (BYVAL hCursor AS HCURSOR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hCursor* | Handle to the cursor. |

#### Return value

The cursor handle.

---

## NormalImageHandle

Returns the handle of the normal image.

```
PROPERTY NormalImageHandle () AS HANDLE
```
---

## DisabledImageHandle

Returns the handle of the disabled image.

```
PROPERTY DisabledImageHandle () AS HANDLE
```
---

## HotImageHandle

Returns the handle of the hot image.

```
PROPERTY HotImageHandle () AS HANDLE
```
---

## Font

Gets/sets the handle of the font used by the button.

```
PROPERTY Font () AS HFONT
PROPERTY Font (BYVAL hFont AS HFONT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hFont* | Handle to the font. |

#### Return value

The font handle.

---

## ImageHeight

Gets/sets the height of the image.

```
PROPERTY ImageHeight () AS LONG
PROPERTY ImageHeight (BYVAL nHeight AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nHeight* | Height of the image, in pixels. |

#### Return value

The height of the image, in pixels.

---

## ImageMargin

Gets/sets the image margin.

```
PROPERTY ImageMargin () AS LONG
PROPERTY ImageMargin (BYVAL nMargin AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMargin* | Margin (in pixels). |

#### Return value

The margin of the image, in pixels.

---

## ImagePos

Gets/sets the image position.

```
PROPERTY ImagePos () AS LONG
PROPERTY ImagePos (BYVAL nPos AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPos* | The position of the image. Can be one of the following values: |

| Value      | Description |
| ---------- | ----------- |
| XPBI_NONE | No image. |
| XPBI_LEFT | Left side (default). |
| XPBI_RIGHT | Right side. |
| XPBI_CENTER | Center. |
| XPBI_VCENTER | Vertically centered. |
| XPBI_TOP | Top. |
| XPBI_BOTTOM | Bottom. |
| XPBI_ABOVE | Above the text. |
| XPBI_BELOW | Below the text. |
| XPBI_CENTERCENTER | Center-center (no text). |

#### Return value

The position of th image.

---

## ImageType

Returns the image type

```
PROPERTY ImageType () AS LONG
```

#### Return value

The image type: IMAGE_ICON or IMAGE_BITMAP.

---

## ImageWidth

Gets/sets the width of the image.

```
PROPERTY ImageWidth () AS LONG
PROPERTY ImageWidth (BYVAL nWidth AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Margin (in pixels). |

#### Return value

The width of the image, in pixels.

---

## TextBkColor

Gets/sets the text background color of the button.

```
PROPERTY TextBkColor () AS COLORREF
PROPERTY TextBkColor (BYVAL textColor AS COLORREF)
```

| Parameter  | Description |
| ---------- | ----------- |
| *textColor* | A COLORREF color value. Use the FreeBasic BGR function. |

#### Return value

The text background color of the button as a COLORREF value.

---

## TextBkColorDown

Gets/sets the text background color of the button when it is down (pressed).

```
PROPERTY TextBkColorDown () AS COLORREF
PROPERTY TextBkColorDown (BYVAL textColor AS COLORREF)
```

| Parameter  | Description |
| ---------- | ----------- |
| *textColor* | A COLORREF color value. Use the FreeBasic BGR function. |

#### Return value

The text background color of the button when it is down (pressed) as a COLORREF value.

---

## TextForeColor

Gets/sets the text foreground color of the button.

```
PROPERTY TextForeColor () AS COLORREF
PROPERTY TextForeColor (BYVAL textColor AS COLORREF)
```

| Parameter  | Description |
| ---------- | ----------- |
| *textColor* | A COLORREF color value. Use the FreeBasic BGR function. |

#### Return value

The text foreground color of the button as a COLORREF value.

---

## TextForeColorDown

Gets/sets the text foreground color of the button when it is down (pressed).

```
PROPERTY TextForeColorDown () AS COLORREF
PROPERTY TextForeColorDown (BYVAL textColor AS COLORREF)
```

| Parameter  | Description |
| ---------- | ----------- |
| *textColor* | A COLORREF color value. Use the FreeBasic BGR function. |

#### Return value

The text foreground color of the button when it is down (pressed) as a COLORREF value.

---

## TextForeColorHot

Gets/sets the text foreground hot color of the button.

```
PROPERTY TextForeColorHot () AS COLORREF
PROPERTY TextForeColorHot (BYVAL textColor AS COLORREF)
```

| Parameter  | Description |
| ---------- | ----------- |
| *textColor* | A COLORREF color value. Use the FreeBasic BGR function. |

#### Return value

The text foreground hot color of the button.

---

## TextFormat

Gets/sets the method of formatting the text.

```
PROPERTY TextFormat () AS DWORD
PROPERTY TextFormat (BYVAL dwTextFlags AS DWORD)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwTextFlags* | Specifies the method of formatting the text. This parameter can be one or more of the following values: |

| Value      | Meaning |
| ---------- | ----------- |
| DT_BOTTOM | Justifies the text to the bottom of the rectangle. This value is used only with the DT_SINGLELINE value. |
| DT_CALCRECT | Determines the width and height of the rectangle. If there are multiple lines of text, DrawText uses the width of the rectangle pointed to by the lpRect parameter and extends the base of the rectangle to bound the last line of text. If the largest word is wider than the rectangle, the width is expanded. If the text is less than the width of the rectangle, the width is reduced. If there is only one line of text, DrawText modifies the right side of the rectangle so that it bounds the last character in the line. In either case, DrawText returns the height of the formatted text but does not draw the text. |
| DT_CENTER | Centers text horizontally in the rectangle. |
| DT_EDITCONTROL | Duplicates the text-displaying characteristics of a multiline edit control. Specifically, the average character width is calculated in the same manner as for an edit control, and the function does not display a partially visible last line. |
| DT_END_ELLIPSIS | For displayed text, if the end of a string does not fit in the rectangle, it is truncated and ellipses are added. If a word that is not at the end of the string goes beyond the limits of the rectangle, it is truncated without ellipses. The string is not modified unless the DT_MODIFYSTRING flag is specified. Compare with DT_PATH_ELLIPSIS and DT_WORD_ELLIPSIS. |
| DT_EXPANDTABS | Expands tab characters. The default number of characters per tab is eight. The DT_WORD_ELLIPSIS, DT_PATH_ELLIPSIS, and DT_END_ELLIPSIS values cannot be used with the DT_EXPANDTABS value. |
| DT_EXTERNALLEADING | Includes the font external leading in line height. Normally, external leading is not included in the height of a line of text. |
| DT_HIDEPREFIX | Includes the font external leading in line height. Normally, external leading is not included in the height of a line of text.<br>Windows 2000/XP: Ignores the ampersand (&) prefix character in the text. The letter that follows will not be underlined, but other mnemonic-prefix characters are still processed.<br>Example: input string: "A&bc&&d", normal: "Abc&d", DT_HIDEPREFIX: "Abc&d". Compare with DT_NOPREFIX and DT_PREFIXONLY. |
| DT_INTERNAL | Uses the system font to calculate text metrics. |
| DT_LEFT | Aligns text to the left. |
| DT_MODIFYSTRING | Modifies the specified string to match the displayed text. This value has no effect unless DT_END_ELLIPSIS or DT_PATH_ELLIPSIS is specified. |
| DT_NOCLIP | Draws without clipping. **DrawText** is somewhat faster when DT_NOCLIP is used. |
| DT_NOFULLWIDTHCHARBREAK | Windows 98/Me, Windows 2000/XP: Prevents a line break at a DBCS (double-wide character string), so that the line breaking rule is equivalent to SBCS strings. For example, this can be used in Korean windows, for more readability of icon labels. This value has no effect unless DT_WORDBREAK is specified. |
| DT_NOPREFIX | Uses the system font to calculate text metrics.Turns off processing of prefix characters. Normally, DrawText interprets the mnemonic-prefix character & as a directive to underscore the character that follows, and the mnemonic-prefix characters && as a directive to print a single &. By specifying DT_NOPREFIX, this processing is turned off.<br>Example: input string: "A&bc&&d", normal: "Abc&d", DT_NOPREFIX: "A&bc&&d". Compare with DT_HIDEPREFIX and DT_PREFIXONLY. |
| DT_PATH_ELLIPSIS | For displayed text, replaces characters in the middle of the string with ellipses so that the result fits in the specified rectangle. If the string contains backslash (\) characters, DT_PATH_ELLIPSIS preserves as much as possible of the text after the last backslash. The string is not modified unless the DT_MODIFYSTRING flag is specified. Compare with DT_END_ELLIPSIS and DT_WORD_ELLIPSIS. |
| DT_PREFIXONLY | Windows 2000/XP: Draws only an underline at the position of the character following the ampersand (&) prefix character. Does not draw any other characters in the string.<br>Example: input string: "A&bc&&d", normal: "Abc&d", DT_PREFIXONLY: " _ ". Compare with DT_HIDEPREFIX and DT_NOPREFIX. |
| DT_RIGHT | Aligns text to the right. |
| DT_RTLREADING | Layout in right-to-left reading order for bi-directional text when the font selected into the hdc is a Hebrew or Arabic font. The default reading order for all text is left-to-right. |
| DT_SINGLELINE | Displays text on a single line only. Carriage returns and line feeds do not break the line. |
| DT_TABSTOP | Sets tab stops. Bits 15-8 (high-order byte of the low-order word) of the uFormat parameter specify the number of characters for each tab. The default number of characters per tab is eight. The DT_CALCRECT, DT_EXTERNALLEADING, DT_INTERNAL, DT_NOCLIP, and DT_NOPREFIX values cannot be used with the DT_TABSTOP value. |
| DT_TOP | Justifies the text to the top of the rectangle. |
| DT_VCENTER | Centers text vertically. This value is used only with the DT_SINGLELINE value. |
| DT_WORDBREAK | Breaks words. Lines are automatically broken between words if a word would extend past the edge of the rectangle specified by the lpRect parameter. A carriage return-line feed sequence also breaks the line. If this is not specified, output is on one line. |
| DT_WORD_ELLIPSIS | Truncates any word that does not fit in the rectangle and adds ellipses. Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS. |

**Default value**: DT_CENTER OR DT_VCENTER OR DT_SINGLELINE.

#### Return value

The method of formatting the text.

---

## TextMargin

Gets/sets the the margin of the text of the button.

```
PROPERTY TextMargin () AS LONG
PROPERTY TextMargin (BYVAL nMargin AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nMargin* | Margin (in pixels). |

#### Return value

The text margin.

---

## Toggle

Gets/sets button to toggle state (TRUE) or to pushbutton state (FALSE).

```
PROPERTY Toggle () AS LONG
PROPERTY Toggle (BYVAL fToggle AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *fToggle* | Toggled state (CTRUE or FALSE). |

#### Return value

Returns CTRUE if the button is toggled or FALSE otherwise.

---

## ToggleState

Gets/sets the toggle state: pushed (CTRUE) or unpushed (FALSE).

```
PROPERTY ToggleState () AS LONG
PROPERTY ToggleState (BYVAL fState AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *fState* | Toggle state (TRUE or FALSE). |

#### Return value

Returns CTRUE if the button is toggled or FALSE otherwise.

#### Return value

Returns CTRUE if the button is toggled or FALSE otherwise.

---
