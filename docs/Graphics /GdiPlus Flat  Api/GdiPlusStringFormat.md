## StringFormat Functions

The **StringFormat** functions allow to set and get text layout information (such as alignment, orientation, tab stops, and clipping) and display manipulations (such as trimming, font substitution for characters that are not supported by the requested font, and digit substitution for languages that do not use Western European digits). A **StringFormat** object can be passed to the **GdipDrawString** function to format a string.

The **StringFormat** object is the key to unlocking advanced text layout in GDI+. It allows developers to control alignment, trimming, digit substitution, and rendering behavior with precision. Whether you're building multilingual interfaces, custom typography engines, or DPI-aware applications, mastering StringFormat ensures your text looks exactly as intended—on any screen, in any language.

| Name       | Description |
| ---------- | ----------- |
| [GdipCloneStringFormat](#gdipclonestringformat) | Creates a new **StringFormat** object and initializes it with the contents of a **StringFormat** object. |
| [GdipCreateStringFormat](#gdipcreatestringformat) | Creates a **StringFormat** object based on string format flags and a language. |
| [GdipDeleteStringFormat](#gdipdeletestringformat) | Deletes the specified **StringFormat** object. |
| [GdipGetStringFormatAlign](#gdipgetstringformatalign) | Gets an element of the **StringAlignment** enumeration that indicates the character alignment of a **StringFormat** object in relation to the origin of the layout rectangle. A layout rectangle is used to position the displayed string. |
| [GdipGetStringFormatDigitSubstitution](#gdipgetstringformatdigitsubstitution) | Gets an element of the **StringDigitSubstitute** enumeration that indicates the digit substitution method that is used by a **StringFormat** object. |
| [GdipGetStringFormatFlags](#gdipgetstringformatflags) | Gets the string format flags for a **StringFormat** object. |
| [GdipGetStringFormatHotkeyPrefix](#gdipgetstringformathotkeyprefix) | Gets an element of the **HotkeyPrefix** enumeration that indicates the type of processing that is performed on a string when a hot key prefix, an ampersand (&), is encountered. |
| [GdipGetStringFormatLineAlign](#gdipgetstringformatlinealign) | Gets an element of the **StringAlignment** enumeration that indicates the line alignment of a **StringFormat** object in relation to the origin of the layout rectangle. The line alignment setting specifies how to align the string vertically in the layout rectangle. The layout rectangle is used to position the displayed string. |
| [GdipGetStringFormatMeasurableCharacterRangeCount](#gdipgetstringformatmeasurablecharacterrangecount) | Gets the number of measurable character ranges that are currently set. The character ranges that are set can be measured in a string by using the **GdipMeasureCharacterRanges** function. |
| [GdipGetStringFormatTabStopCount](#gdipgetstringformattabstopcount) | Gets the number of tab-stop offsets in a **StringFormat** object. |
| [GdipGetStringFormatTabStops](#gdipgetstringformattabstops) | Gets the offsets of the tab stops in a **StringFormat** object. |
| [GdipGetStringFormatTrimming](#gdipgetstringformattrimming) | Gets an element of the **StringTrimming** enumeration that indicates the trimming style of a **StringFormat** object. The trimming style determines how to trim characters from a string that is too large to fit in the layout rectangle. |
| [GdipSetStringFormatAlign](#gdipsetstringformatalign) | Sets the character alignment of a **StringFormat** object in relation to the origin of the layout rectangle. A layout rectangle is used to position the displayed string. |
| [GdipSetStringFormatDigitSubstitution](#gdipsetstringformatdigitsubstitution) | Sets the digit substitution method and the language that corresponds to the digit substitutes. |
| [GdipSetStringFormatFlags](#gdipsetstringformatflags) | Sets the format flags for a **StringFormat** object. The format flags determine most of the characteristics of a **StringFormat** object. |
| [GdipSetStringFormatHotkeyPrefix](#gdipsetstringformathotkeyprefix) | Sets the type of processing that is performed on a string when the hot key prefix, an ampersand (&), is encountered. The ampersand is called the hot key prefix and can be used to designate certain keys as hot keys. |
| [GdipSetStringFormatLineAlign](#gdipsetstringformatlinealign) | Sets the line alignment of a **StringFormat** object in relation to the origin of the layout rectangle. The line alignment setting specifies how to align the string vertically in the layout rectangle. The layout rectangle is used to position the displayed string. |
| [GdipSetStringFormatMeasurableCharacterRanges](#gdipsetstringformatmeasurablecharacterranges) | Sets a series of character ranges for a **StringFormat** object that, when in a string, can be measured by the **GdipMeasureCharacterRanges** method. |
| [GdipSetStringFormatTabStops](#gdipsetstringformattabstops) | Sets the offsets for tab stops in a **StringFormat** object. |
| [GdipSetStringFormatTrimming](#gdipsetstringformattrimming) | Sets the trimming style for a **StringFormat** object. The trimming style determines how to trim a string so that it fits into the layout rectangle. |
| [GdipStringFormatGetGenericDefault](#gdipstringformatgetgenericdefault) | Creates a generic, default **StringFormat** object. |
| [GdipStringFormatGetGenericTypographic](#gdipstringformatgetgenerictypographic) | Creates a generic, typographic **StringFormat** object. |

---

## GdipCloneStringFormat

Creates a new **StringFormat** object and initializes it with the contents of a **StringFormat** object.

```
FUNCTION GdipCloneStringFormat (BYVAL format AS CONST GpStringFormat PTR, _
   BYVAL newFormat AS GpStringFormat PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *newFormat* | [in] Pointer to a variable that receives a pointer to the new **StringFormat** object. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, clones it, and then uses the clone
' to draw a formatted string.
' ========================================================================================
SUB Example_CloneStringFormat (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(24, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object.
   DIM stringFormat AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @stringFormat)
   hStatus = GdipSetStringFormatAlign(stringFormat, StringAlignmentCenter)

   ' // Clone the StringFormat object.
   DIM clonedStringFormat AS GpStringFormat PTR
   hStatus = GdipCloneStringFormat(stringFormat, @clonedStringFormat)

   ' // Use the generic StringFormat object in a call to DrawString.
   DIM rcf AS GpRectF = (30, 30, 200, 200)
   DIM wszText AS WSTRING * 64 = "This text was formatted by a cloned StringFormat object."
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, clonedStringFormat, solidBrush)

   ' // Draw the rectangle that encloses the text.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF clonedStringFormat THEN GdipDeleteStringFormat(pen)
   IF stringFormat THEN GdipDeleteStringFormat(stringFormat)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateStringFormat

Creates a **StringFormat** object based on string format flags and a language.

```
FUNCTION GdipCreateStringFormat (BYVAL formatAttributes AS INT_, BYVAL language AS LANGID, _
    BYVAL format AS GpStringFormat PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *formatAttributes* | [in] Value that specifies the format flags that control most of the characteristics of the **StringFormat** object. The flags are set by applying a bitwise **OR** to elements of the **StringFormatFlags** enumeration. The default value is 0 (no flags set). |
| *language* | [in] Sixteen-bit value that specifies the language to use. The default value is LANG_NEUTRAL, which is the user's default language. |
| *format* | [out] Pointer to a variable that receives a pointer to the new **StringFormat** object. |

#### Example

```
' ========================================================================================
' The following example uses the specified formatting to draw a string in a layout rectangle.
' ========================================================================================
SUB Example_CreateStringFormat (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(16, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatAlign(format, StringAlignmentCenter)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @solidBrush)

   ' // Draw the string
   DIM wszText AS WSTRING * 64 = "Sample text"
   DIM rcf AS GpRectF = (30, 30, 200, 25)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Create a Pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 3, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Draw a rectangle
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeleteStringFormat

Deletes the specified **StringFormat** object.

```
FUNCTION GdipDeleteStringFormat (BYVAL format AS GpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |

#### Usage example

```
...
...
...
' // Create a StringFormat object
DIM format AS GpStringFormat PTR
hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
...
...
...
' // Delete the string format
GdipDeleteStringFormat(format)
```
---

## GdipGetStringFormatAlign

Gets an element of the **StringAlignment** enumeration that indicates the character alignment of a **StringFormat** object in relation to the origin of the layout rectangle. A layout rectangle is used to position the displayed string.

```
FUNCTION GdipGetStringFormatAlign (BYVAL format AS CONST GpStringFormat PTR, _
    BYVAL align AS StringAlignment PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *align* | [out] Pointer to a variable that receives an element of the **StringAlignment** enumeration. |

#### StringAlignment

The **StringAlignment** enumeration specifies how a string is aligned in reference to the bounding rectangle. A bounding rectangle is used to define the area in which the text displays.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StringAlignmentNear* | 0 | Specifies that alignment is towards the origin of the bounding rectangle. May be used for alignment of characters along the line or for alignment of lines within the rectangle. For a right to left bounding rectangle (**StringFormatFlagsDirectionRightToLeft**), the origin is at the upper right. |
| *StringAlignmentCenter* | 1 | Specifies that alignment is centered between origin and extent (width or height) of the formatting rectangle. |
| *StringAlignmentFar* | 2 | Specifies that alignment is to the far extent (lower right) of the formatting rectangle. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's character alignment,
' gets the alignment value, and stores it in a variable. The code then creates a second
' StringFormat object and uses the stored alignment value to set the character alignment of
' the second StringFormat object. Next, the code uses the second StringFormat object to draw
' a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetStringFormatAlign (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatAlign(format, StringAlignmentCenter)

   ' // Get the alignment setting from the StringFormat object.
   DIM strAlignment AS StringAlignment 
   hStatus = GdipGetStringFormatAlign(format, @strAlignment)

   ' // Create a second StringFormat object with the same alignment.
   DIM format2 AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, format2)
   hStatus = GdipSetStringFormatAlign(format2, strAlignment)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Draw the string
   DIM wszText AS WSTRING * 64 = "This text was formatted by a second StringFormat object."
   DIM rcf AS GpRectF = (30, 30, 150, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format2, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format2 THEN GdipDeleteStringFormat(format2)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetStringFormatDigitSubstitution

Gets an element of the **StringDigitSubstitute** enumeration that indicates the digit substitution method that is used by a **StringFormat** object.

```
FUNCTION GdipGetStringFormatDigitSubstitution (BYVAL format AS CONST GpStringFormat PTR, _
   BYVAL language AS LANGID PTR, BYVAL substitute AS StringDigitSubstitute PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *language* | [out] A 16-bit value that specifies the language to use. |
| *substitute* | [out] Pointer to a variable that receives an element of the **StringDigitSubstitute** enumeration. |

#### Remarks

The digit substitution method replaces, in a string, Western European digits with digits that correspond to a user's locale or language.

#### Example

```
' ========================================================================================
' The following example demonstrates the use GdipSetStringFormatDigitSubstitution to display
' localized digits—such as Arabic-Indic numerals—based on a specified language and substitution
' method. The GdipGetStringFormatDigitSubstitution function returns the language identifier
' and the digit substitute set.
' ========================================================================================
SUB Example_GetStringFormatDigitSubstitution (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Tahoma", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(20, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create StringFormat and set digit substitution
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_ARABIC, @format)
   hStatus = GdipSetStringFormatDigitSubstitution(format, LANG_ARABIC, StringDigitSubstituteNational)

   ' // Get the digit substitute set
   DIM language AS LANGID, substitute AS StringDigitSubstitute
   hStatus = GdipGetStringFormatDigitSubstitution(format, @language, @substitute)
   AfxMsg WSTR("Language: " & WSTR(language) & CHR(13, 10) & "Digit substitute : " & WSTR(substitute))

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' Draw string with digits
   DIM rcf AS GpRectF = (30, 30, 300, 50)
   DIM wszText AS WSTRING * 64 = "Year: 2025"
   GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

---

## GdipGetStringFormatFlags

Gets the string format flags for a **StringFormat** object.

```
FUNCTION GdipGetStringFormatFlags (BYVAL format AS CONST GpStringFormat PTR, _
   BYVAL flags AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *flags* | [out] Pointer to a variable that receives a value that indicates which string format flags are set for this StringFormat object. This value can be any combination (the result of a bitwise **OR** applied to two or more elements) of elements of the **StringFormatFlags** enumeration. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's format flags, and
' then gets the 32-bit value that contains the format flags and stores it in a variable.
' The code then creates another StringFormat object and uses the stored format flags value
' to set the format flags of the second StringFormat object. Next, the code uses the second
' StringFormat object to draw a formatted string. The code also draws the string's layout
' rectangle.
' ========================================================================================
SUB Example_GetStringFormatFlags (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object, and set its format flags.
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatFlags(format, StringFormatFlagsDirectionVertical OR StringFormatFlagsNoFitBlackBox)

   ' // Get the format flags from the StringFormat object.
   DIM flags AS LONG
   hStatus = GdipGetStringFormatFlags(format, @flags)

   ' // Create a second StringFormat object with the same flags.
   DIM format2 AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format2)
   hStatus = GdipSetStringFormatFlags(format2, flags)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Draw the string using the second StringFormat object
   DIM wszText AS WSTRING * 64 = "This text is vertical because of a format flag."
   DIM rcf AS GpRectF = (30, 30, 150, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format2, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format2 THEN GdipDeleteStringFormat(format2)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetStringFormatHotkeyPrefix

Gets an element of the **HotkeyPrefix** enumeration that indicates the type of processing that is performed on a string when a hot key prefix, an ampersand (&), is encountered.

```
FUNCTION GdipGetStringFormatHotkeyPrefix (BYVAL format AS CONST GpStringFormat PTR, _
   BYVAL hotkeyPrefix AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *hotkeyPrefix* | [out] Pointer to a variable that receives an element of the **HotkeyPrefix** enumeration that indicates the type of hot key prefix processing that is performed on a string. |

#### HotKeyPrefix

The **HotkeyPrefix** enumeration specifies how to display hot keys. There are three options: do nothing, display hot keys underlined, and hide the hot key underlines.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *HotkeyPrefixNone* | 0 | Specifies that no hot key processing occurs. |
| *HotkeyPrefixShow* | 1 | Specifies that Unicode text is scanned for ampersands (&), which are interpreted as hot key markers in the same way as menu and dialog resources are processed in the Windows user interface. All pairs of ampersands are replaced by a single ampersand. All single ampersands are removed and the first character that follows the first single ampersand is displayed underlined. |
| *HotkeyPrefixHide* | 2 | Specifies that Unicode text is scanned for ampersands (&), which are substituted and removed as in HotkeyPrefixShow but no underline is shown. This option can be used when accessibility hot key underlines are not required. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the type of hot key prefix
' processing to be performed on the string, and then gets the type of processing and stores
' it in a variable. The code then creates a second StringFormat object and uses the stored
' value to set the type of hot key prefix processing for the second StringFormat object.
' The code uses the second StringFormat object to draw a string that contains the hot key
' prefix character. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetStringFormatHotkeyPrefix (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatHotkeyPrefix(format, HotkeyPrefixShow)

 ' // Get the hot key prefix from the StringFormat object.
   DIM hotkeyPrefix AS LONG
   hStatus = GdipGetStringFormatHotkeyPrefix(format, @hotkeyPrefix)

   ' // Create a second StringFormat object with the same hotkey prefix.
   DIM format2 AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format2)
   hStatus = GdipSetStringFormatHotkeyPrefix(format2, hotkeyPrefix)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Draw the string
   DIM wszText AS WSTRING * 64 = "This &text has some &underlined characters."
   DIM rcf AS GpRectF = (30, 30, 160, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format2, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format2 THEN GdipDeleteStringFormat(format2)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetStringFormatLineAlign

Gets an element of the **StringAlignment** enumeration that indicates the line alignment of a **StringFormat** object in relation to the origin of the layout rectangle. The line alignment setting specifies how to align the string vertically in the layout rectangle. The layout rectangle is used to position the displayed string.

```
FUNCTION GdipGetStringFormatLineAlign (BYVAL format AS CONST GpStringFormat PTR, _
   BYVAL align AS StringAlignment PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *align* | [out] Pointer to a variable that receives an element of the **StringAlignment** enumeration. |


#### StringAlignment

The **StringAlignment** enumeration specifies how a string is aligned in reference to the bounding rectangle. A bounding rectangle is used to define the area in which the text displays.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StringAlignmentNear* | 0 | Specifies that alignment is towards the origin of the bounding rectangle. May be used for alignment of characters along the line or for alignment of lines within the rectangle. For a right to left bounding rectangle (**StringFormatFlagsDirectionRightToLeft**), the origin is at the upper right. |
| *StringAlignmentCenter* | 1 | Specifies that alignment is centered between origin and extent (width or height) of the formatting rectangle. |
| *StringAlignmentFar* | 2 | Specifies that alignment is to the far extent (lower right) of the formatting rectangle. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's line alignment,
' and then gets the line alignment setting and stores it in a variable. The code then
' creates a second StringFormat object and uses the stored alignment value to set the line
' alignment of the second StringFormat object. Next, the code uses the second StringFormat
' object to draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetStringFormatLineAlign (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Create a StringFormat object, and set its trimming style.
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatLineAlign(format, StringAlignmentCenter)

   ' // Get the line alignment from the StringFormat object.
   DIM strAlignment AS LONG
   hStatus = GdipGetStringFormatHotkeyPrefix(format, @strAlignment)

   ' // Create a second StringFormat object with the same line alignment.
   DIM format2 AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format2)
   hStatus = GdipSetStringFormatLineAlign(format2, strAlignment)

   ' // Use the second StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 64 = "This text was formatted by a second StringFormat object."
   DIM rcf AS GpRectF = (30, 30, 150, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format2, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format2 THEN GdipDeleteStringFormat(format2)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetStringFormatMeasurableCharacterRangeCount

Gets the number of measurable character ranges that are currently set. The character ranges that are set can be measured in a string by using the **GdipMeasureCharacterRanges** function.

```
FUNCTION GdipGetStringFormatMeasurableCharacterRangeCount (BYVAL format AS CONST GpStringFormat PTR, _
   BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *count* | [out] Pointer to a variable that receives the number of character ranges that can be measured by **GdipMeasureCharacterRanges**. |

#### Usage example

For a full example see **GdipSetStringFormatMeasurableCharacterRanges**.

```
DIM count AS LONG
hStatus = GdipGetStringFormatMeasurableCharacterRangeCount(stringFormat, @count)
```
---

## GdipGetStringFormatTabStopCount

Gets the number of tab-stop offsets in a **StringFormat** object.

```
FUNCTION GdipGetStringFormatTabStopCount (BYVAL format AS CONST GpStringFormat PTR, _
   BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *count* | [out] Pointer to a variable that receives the number of tab stops in the **StringFormat** object. |

#### Example

For the complete excample see **GdipSetStringFormatTabStops**.

```
' // Get the tab stops.
DIM tabStopCount AS LONG
DIM firstTabOffset AS SINGLE
DIM tabStops(ANY) AS SINGLE

hStatus = GdipGetStringFormatTabStopCount(format, @tabStopCount)
REDIM tabStops(tabStopCount - 1)
hStatus = GdipGetStringFormatTabStops(format, tabStopCount, @firstTabOffset, @tabStops(0))

FOR i AS LONG = 0 TO tabStopCount - 1
   ' // Inspect or use the value in tabStops(i)
   OutputDebugStringW WSTR(tabStops(i))
NEXT
```

---

## GdipGetStringFormatTabStops

Gets the offsets of the tab stops in a **StringFormat** object.

```
FUNCTION GdipGetStringFormatTabStops (BYVAL format AS CONST GpStringFormat PTR, BYVAL count AS INT_, _
         BYVAL firstTabOffset AS REAL PTR, BYVAL tabStops AS REAL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *count* | [out] Integer that specifies the number of tab-stop offsets in the **tabStops** array. |
| *firstTabOffset* | [out] Pointer to a SINGLE variable that receives the initial offset position. This initial offset position is relative to the string's origin and the offset of the first tab stop is relative to the initial offset position. |
| *tabStops* | [out] Pointer to an array of type SINGLE that receives the tab-stop offsets. The offset of the first tab stop is the first value in the array, the offset of the second tab stop, the second value in the array, and so on. |

Remarks

Each tab-stop offset in the *tabStops* array, except the first one, is relative to the previous one. The first tab-stop offset is relative to the initial offset position indicated by *firstTabOffset*. For example, if the initial offset position is 8 and the first tab-stop offset is 50, then the first tab stop is at position 58. If the initial offset position is zero, then the first tab-stop offset is relative to position 0, the string origin.

#### Example

For the complete excample see **GdipSetStringFormatTabStops**.

```
' // Get the tab stops.
DIM tabStopCount AS LONG
DIM firstTabOffset AS SINGLE
DIM tabStops(ANY) AS SINGLE

hStatus = GdipGetStringFormatTabStopCount(format, @tabStopCount)
REDIM tabStops(tabStopCount - 1)
hStatus = GdipGetStringFormatTabStops(format, tabStopCount, @firstTabOffset, @tabStops(0))

FOR i AS LONG = 0 TO tabStopCount - 1
   ' // Inspect or use the value in tabStops(i)
   OutputDebugStringW WSTR(tabStops(i))
NEXT
```
---

## GdipGetStringFormatTrimming

Gets an element of the **StringTrimming** enumeration that indicates the trimming style of a **StringFormat** object. The trimming style determines how to trim characters from a string that is too large to fit in the layout rectangle.

```
FUNCTION GdipGetStringFormatTrimming (BYVAL format AS CONST GpStringFormat PTR, _
    BYVAL trimming AS StringTrimming PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *trimming* | [out] Pointer to a variable that receives an element of the *StringTrimming* enumeration. |

#### StringTrimming

The **StringTrimming** enumeration specifies how to trim characters from a string so that the string fits into a layout rectangle. The layout rectangle is used to position and size the display string.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StringTrimmingNone* | 0 | Specifies that no trimming is done. |
| *StringTrimmingCharacter* | 1 | Specifies that the string is broken at the boundary of the last character that is inside the layout rectangle. This is the default. |
| *StringTrimmingWord* | 2 | Specifies that the string is broken at the boundary of the last word that is inside the layout rectangle. |
| *StringTrimmingEllipsisCharacter* | 3 | Specifies that the string is broken at the boundary of the last character that is inside the layout rectangle and an ellipsis (...) is inserted after the character. |
| *StringTrimmingEllipsisWord* | 4 | Specifies that the string is broken at the boundary of the last word that is inside the layout rectangle and an ellipsis (...) is inserted after the word. |
| *StringTrimmingEllipsisPath* | 5 | Specifies that the center is removed from the string and replaced by an ellipsis. The algorithm keeps as much of the last portion of the string as possible. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the string's trimming style,
' and then gets the trimming style and stores it in a variable. The code then creates a
' second StringFormat object and uses the stored trimming style to set the trimming style
' of the second StringFormat object. Next, the code uses the second StringFormat object to
' draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetStringFormatTrimming (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Create a StringFormat object, and set its trimming style.
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatTrimming(format, StringTrimmingEllipsisWord)

  ' // Get the trimming setting from the StringFormat object.
   DIM strTrimming AS StringTrimming
   hStatus = GdipGetStringFormatTrimming(format, @strTrimming)

   ' // Create a second StringFormat object with the same trimming style.
   DIM format2 AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format2)
   hStatus = GdipSetStringFormatTrimming(format2, strTrimming)

   ' // Use the second StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 64 = "One two three four five six seven eight nine ten"
   DIM rcf AS GpRectF = (30, 30, 160, 60)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format2, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format2 THEN GdipDeleteStringFormat(format2)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetStringFormatDigitSubstitution

Sets the digit substitution method and the language that corresponds to the digit substitutes.

```
FUNCTION GdipSetStringFormatDigitSubstitution (BYVAL format AS GpStringFormat PTR, _
   BYVAL language AS LANGID, BYVAL substitute AS StringDigitSubstitute) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *language* | [in] Sixteen-bit value that forms a national language support (NLS) language identifier. The identifier specifies the language associated with the substitute digits. For example, if this **StringFormat** object uses Arabic substitution digits, then this method will return a value that indicates an Arabic language. |
| *substitute* | [in] Element of the **StringDigitSubstitute** enumeration that specifies the digit substitution method to be used. |

#### Example

```
' ========================================================================================
' The following example demonstrates the use GdipSetStringFormatDigitSubstitution to display
' localized digits—such as Arabic-Indic numerals—based on a specified language and substitution method.
' ========================================================================================
SUB Example_SetStringFormatDigitSubstitution (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Tahoma", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(20, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create StringFormat and set digit substitution
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_ARABIC, @format)
   hStatus = GdipSetStringFormatDigitSubstitution(format, LANG_ARABIC, StringDigitSubstituteNational)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' Draw string with digits
   DIM rcf AS GpRectF = (30, 30, 300, 50)
   DIM wszText AS WSTRING * 64 = "Year: 2025"
   GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetStringFormatFlags

Sets the format flags for a **StringFormat** object. The format flags determine most of the characteristics of a **StringFormat** object.

```
FUNCTION GdipSetStringFormatFlags (BYVAL format AS GpStringFormat PTR, BYVAL flags AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *flags* | [in] Thirty-two bit value that specifies the format flags that control most of the characteristics of the **StringFormat** object. The flags are set by applying a bitwise **OR** to elements of the **StringFormatFlags**. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the format flags, and draws the
' formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_SetStringFormatFlags (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object, and set its format flags.
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatFlags(format, StringFormatFlagsDirectionVertical OR StringFormatFlagsNoFitBlackBox)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Use the StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 64 = "This text is vertical because of a format flag."
   DIM rcf AS GpRectF = (30, 30, 150, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Draw the rectangle that encloses the text.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetStringFormatHotkeyPrefix

Sets the type of processing that is performed on a string when the hot key prefix, an ampersand (&), is encountered. The ampersand is called the hot key prefix and can be used to designate certain keys as hot keys.

```
FUNCTION GdipSetStringFormatHotkeyPrefix (BYVAL format AS GpStringFormat PTR, _
    BYVAL hotkeyPrefix AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *hotkeyPrefix* | [in] Element of the **HotkeyPrefix** enumeration that specifies how to process the hot key prefix. |

#### HotKeyPrefix

The **HotkeyPrefix** enumeration specifies how to display hot keys. There are three options: do nothing, display hot keys underlined, and hide the hot key underlines.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *HotkeyPrefixNone* | 0 | Specifies that no hot key processing occurs. |
| *HotkeyPrefixShow* | 1 | Specifies that Unicode text is scanned for ampersands (&), which are interpreted as hot key markers in the same way as menu and dialog resources are processed in the Windows user interface. All pairs of ampersands are replaced by a single ampersand. All single ampersands are removed and the first character that follows the first single ampersand is displayed underlined. |
| *HotkeyPrefixHide* | 2 | Specifies that Unicode text is scanned for ampersands (&), which are substituted and removed as in HotkeyPrefixShow but no underline is shown. This option can be used when accessibility hot key underlines are not required. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object and sets the type of hot key prefix
' processing to be performed on the string. The code then uses the StringFormat object to
' draw a string that contains the hot key prefix character. The code also draws the string's
' layout rectangle.
' ========================================================================================
SUB Example_SetStringFormatHotkeyPrefix (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object, and set its trimming style.
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatHotkeyPrefix(format, HotkeyPrefixShow)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Use the StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 64 = "This &text has some &underlined characters."
   DIM rcf AS GpRectF = (30, 30, 160, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetStringFormatLineAlign

Sets the character alignment of a **StringFormat** object in relation to the origin of the layout rectangle. A layout rectangle is used to position the displayed string.

```
FUNCTION GdipSetStringFormatAlign (BYVAL format AS GpStringFormat PTR, _
   BYVAL align AS StringAlignment) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *align* | [in] Element of the **StringAlignment** enumeration that specifies how a string is aligned in reference to the bounding rectangle. |

#### StringAlignment

The **StringAlignment** enumeration specifies how a string is aligned in reference to the bounding rectangle. A bounding rectangle is used to define the area in which the text displays.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StringAlignmentNear* | 0 | Specifies that alignment is towards the origin of the bounding rectangle. May be used for alignment of characters along the line or for alignment of lines within the rectangle. For a right to left bounding rectangle (**StringFormatFlagsDirectionRightToLeft**), the origin is at the upper right. |
| *StringAlignmentCenter* | 1 | Specifies that alignment is centered between origin and extent (width or height) of the formatting rectangle. |
| *StringAlignmentFar* | 2 | Specifies that alignment is to the far extent (lower right) of the formatting rectangle. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the line alignment, and draws
' the string using that alignment. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_SetStringFormatLineAlign (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object, and set its trimming style.
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatLineAlign(format, StringAlignmentCenter)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Use the StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 64 = "This text was formatted by a StringFormat object."
   DIM rcf AS GpRectF = (30, 30, 150, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetStringFormatMeasurableCharacterRanges

Sets a series of character ranges for a **StringFormat** object that, when in a string, can be measured by the **GdipMeasureCharacterRanges** method.

```
FUNCTION GdipSetStringFormatMeasurableCharacterRanges (BYVAL format AS GpStringFormat PTR, _
   BYVAL rangeCount AS INT_, BYVAL ranges AS CONST GpCharacterRange PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *rangeCount* | [in] Long integer value that specifies the number of character ranges in the ranges array. |
| *ranges* | [in] Pointer to an array of **GpCharacterRange** structures specify the character ranges to be measured. |

#### Example

```
' ========================================================================================
' The following example defines three ranges of character positions within a string and
' sets those ranges in a StringFormat object. Next, the MeasureCharacterRanges method is
' used to get the three regions of the display that are occupied by the characters that
' are specified by the ranges. This is done for three different layout rectangles to show
' how the regions change according to the layout of the string. Also, on the third repetition
' of this, the string format flags are changed so that the regions measured will include
' trailing spaces.
' ========================================================================================
SUB Example_MeasureCharacterRanges (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create font
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(16, TRUE), FontStyleRegular, UnitPixel, @font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)

   ' // Create brushes and pen
   DIM blueBrush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @blueBrush)
   DIM redBrush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(GDIP_ARGB(100, 255, 0, 0), @redBrush)
   DIM blackPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 1.0, UnitPixel, @blackPen)
   hStatus = GdipScalePenTransform(blackPen, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create string format
   DIM stringFormat As GpStringFormat PTR
   hStatus = GdipStringFormatGetGenericDefault(@stringFormat)

   ' // Define character ranges
   DIM charRanges(2) As GpCharacterRange
   charRanges(0).First = 3  : charRanges(0).Length = 5
   charRanges(1).First = 15 : charRanges(1).Length = 2
   charRanges(2).First = 30 : charRanges(2).Length = 15
   hStatus = GdipSetStringFormatMeasurableCharacterRanges(stringFormat, 3, @charRanges(0))

   ' // Allocate regions
   DIM regions(2) AS GpRegion PTR
   FOR i AS LONG = 0 TO 2
      hStatus = GdipCreateRegion(@regions(i))
   NEXT

   ' // Text to draw
   DIM text AS WSTRING * 128 = "The quick, brown fox easily jumps over the lazy dog."

   ' // Measure and draw for layoutRectA
   DIM layoutRectA AS GpRectF = (20.0, 20.0, 130.0, 130.0)
   hStatus = GdipMeasureCharacterRanges(graphics, @text, -1, font, @layoutRectA, stringFormat, 3, @regions(0))
   hStatus = GdipDrawString(graphics, @text, -1, font, @layoutRectA, stringFormat, blueBrush)
   hStatus = GdipDrawRectangle(graphics, blackPen, layoutRectA.X, layoutRectA.Y, layoutRectA.Width, layoutRectA.Height)
   FOR i AS LONG = 0 TO 2
      hStatus = GdipFillRegion(graphics, redBrush, regions(i))
   NEXT

   ' // Measure and draw for layoutRectB
   DIM layoutRectB AS GpRectF = (160.0, 20.0, 165.0, 130.0)
   hStatus = GdipMeasureCharacterRanges(graphics, @text, -1, font, @layoutRectB, stringFormat, 3, @regions(0))
   hStatus = GdipDrawString(graphics, @text, -1, font, @layoutRectB, stringFormat, blueBrush)
   hStatus = GdipDrawRectangle(graphics, blackPen, layoutRectB.X, layoutRectB.Y, layoutRectB.Width, layoutRectB.Height)
   FOR i AS LONG = 0 TO 2
      hStatus = GdipFillRegion(graphics, redBrush, regions(i))
   NEXT

   ' // Set trailing space flag and draw for layoutRectC
   DIM layoutRectC As GpRectF = (335.0, 20.0, 165.0, 130.0)
   hStatus = GdipSetStringFormatFlags(stringFormat, StringFormatFlagsMeasureTrailingSpaces)
   hStatus = GdipMeasureCharacterRanges(graphics, @text, -1, font, @layoutRectC, stringFormat, 3, @regions(0))
   hStatus = GdipDrawString(graphics, @text, -1, font, @layoutRectC, stringFormat, blueBrush)
   hStatus = GdipDrawRectangle(graphics, blackPen, layoutRectC.X, layoutRectC.Y, layoutRectC.Width, layoutRectC.Height)
   FOR i AS LONG = 0 TO 2
      hStatus = GdipFillRegion(graphics, redBrush, regions(i))
   NEXT

   ' // Cleanup
   FOR i AS LONG = 0 TO 2
      IF regions(i) THEN GdipDeleteRegion(regions(i))
   NEXt
   IF font THEN GdipDeleteFont(font)
   IF stringFormat THEN GdipDeleteStringFormat(stringFormat)
   IF blueBrush THEN GdipDeleteBrush(blueBrush)
   IF redBrush THEN GdipDeleteBrush(redBrush)
   IF blackPen THEN GdipDeletePen(blackPen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetStringFormatTabStops

Sets the offsets for tab stops in a **StringFormat** object.

```
FUNCTION GdipSetStringFormatTabStops (BYVAL format AS GpStringFormat PTR, BYVAL firstTabOffset AS REAL, _
   BYVAL count AS INT_, BYVAL tabStops AS CONST REAL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *firstTabOffset* | [in] Single precision value that specifies the initial offset position. This initial offset position is relative to the string's origin and the offset of the first tab stop is relative to the initial offset position. |
| *count* | [in] Long integer value that specifies the number of tab-stop offsets in the tabStops array. |
| *tabStops* | [in] Pointer to an array of real numbers that specify the tab-stop offsets. The offset of the first tab stop is the first value in the array, the offset of the second tab stop, the second value in the array, and so on. |

#### Remarks

Each tab-stop offset in the *tabStops* array, except the first one, is relative to the previous one. The first tab-stop offset is relative to the initial offset position specified by *firstTabOffset*. For example, if the initial offset position is 8 and the first tab-stop offset is 50, then the first tab stop is at position 58. If the initial offset position is zero, then the first tab-stop offset is relative to position 0, the string origin.

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets tab stops, and uses the
' StringFormat object to draw a string that contains tab characters. The code also draws
' the string's layout rectangle. Then, the code gets the tab stops and proceeds to use or
' inspect the tab stops in some way.
' ========================================================================================
SUB Example_GetStringFormatTabStops (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Create StringFormat and set tab stops
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   DIM tabs(2) AS SINGLE = {150, 100, 100}
   hStatus = GdipSetStringFormatTabStops(format, 0, 3, @tabs(0))
   DIM rcf AS GpRectF = (20, 20, 500, 100)
   rcf.x = 20 : rcf.y = 20 : rcf.Width = 500 : rcf.Height = 100
   DIM wszText AS WSTRING * 128 = "Name" & CHR(9) & "Test 1" & CHR(9) & "Test 2" & CHR(9) & "Test 3"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Draw the rectangle that encloses the text.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Get the tab stops.
   DIM tabStopCount AS LONG
   DIM firstTabOffset AS SINGLE
   DIM tabStops(ANY) AS SINGLE

   hStatus = GdipGetStringFormatTabStopCount(format, @tabStopCount)
   REDIM tabStops(tabStopCount - 1)
   hStatus = GdipGetStringFormatTabStops(format, tabStopCount, @firstTabOffset, @tabStops(0))

   FOR i AS LONG = 0 TO tabStopCount - 1
      ' // Inspect or use the value in tabStops(i)
      OutputDebugStringW WSTR(tabStops(i))
   NEXT

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetStringFormatTrimming

Sets the trimming style for a **StringFormat** object. The trimming style determines how to trim a string so that it fits into the layout rectangle.

```
FUNCTION GdipSetStringFormatTrimming (BYVAL format AS GpStringFormat PTR, _
   BYVAL trimming AS StringTrimming) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [in] Pointer to the **StringFormat** object. |
| *trimming* | [in] Element of the **StringTrimming** enumeration that specifies the style of trimming to be performed on the string. |

#### StringTrimming

The **StringTrimming** enumeration specifies how to trim characters from a string so that the string fits into a layout rectangle. The layout rectangle is used to position and size the display string.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StringTrimmingNone* | 0 | Specifies that no trimming is done. |
| *StringTrimmingCharacter* | 1 | Specifies that the string is broken at the boundary of the last character that is inside the layout rectangle. This is the default. |
| *StringTrimmingWord* | 2 | Specifies that the string is broken at the boundary of the last word that is inside the layout rectangle. |
| *StringTrimmingEllipsisCharacter* | 3 | Specifies that the string is broken at the boundary of the last character that is inside the layout rectangle and an ellipsis (...) is inserted after the character. |
| *StringTrimmingEllipsisWord* | 4 | Specifies that the string is broken at the boundary of the last word that is inside the layout rectangle and an ellipsis (...) is inserted after the word. |
| *StringTrimmingEllipsisPath* | 5 | Specifies that the center is removed from the string and replaced by an ellipsis. The algorithm keeps as much of the last portion of the string as possible. |

#### Example

```
' ========================================================================================
' The following example creates a StringFormat object, sets the trimming style, and uses
' the StringFormat object to draw a string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_SetStringFormatTrimming (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a StringFormat object, and set its trimming style.
   DIM format AS GpStringFormat PTR
   hStatus = GdipCreateStringFormat(0, LANG_NEUTRAL, @format)
   hStatus = GdipSetStringFormatTrimming(format, StringTrimmingEllipsisWord)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Use the StringFormat object in a call to DrawString.
   DIM wszText AS WSTRING * 64 = "One two three four five six seven eight nine ten"
   DIM rcf AS GpRectF = (30, 30, 160, 60)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipStringFormatGetGenericDefault

Creates a generic, default **StringFormat** object.

```
FUNCTION GdipStringFormatGetGenericDefault (BYVAL format AS GpStringFormat PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [out] Pointer to a variable that receives a pointer to the new **StringFormat** object. |
```
' ========================================================================================
' The following example creates a generic, default StringFormat object and then uses it to
' draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetGenericDefault (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a generic StringFormat object.
   DIM format AS GpStringFormat PTR
   hStatus = GdipStringFormatGetGenericDefault(@format)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Draw the string using the second StringFormat object
   DIM wszText AS WSTRING * 64 = "This text was formatted by a generic StringFormat object."
   DIM rcf AS GpRectF = (30, 30, 150, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipStringFormatGetGenericTypographic

Creates a generic, typographic **StringFormat** object.

```
FUNCTION GdipStringFormatGetGenericTypographic (BYVAL format AS GpStringFormat PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *format* | [out] Pointer to a variable that receives a pointer to the new **StringFormat** object. |

#### Example

```
' ========================================================================================
' The following example creates a generic, typographic StringFormat object and then uses it
' to draw a formatted string. The code also draws the string's layout rectangle.
' ========================================================================================
SUB Example_GetGenericTypographic (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   Dim ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Times New Roman", NULL, @fontFamily)
   DIm font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a generic StringFormat object.
   DIM format AS GpStringFormat PTR
   hStatus = GdipStringFormatGetGenericTypographic(@format)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Draw the string using the second StringFormat object
   DIM wszText AS WSTRING * 128 = "This text was formatted by a generic typographic StringFormat object."
   DIM rcf AS GpRectF = (30, 30, 150, 200)
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, format, solidBrush)

   ' // Draw the rectangle that encloses the text
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitPixel, @pen)
   hStatus = GdipScalePenTransform(pen, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF format THEN GdipDeleteStringFormat(format)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
