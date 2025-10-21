# FontFamily Functions

A font family is a group of fonts that have the same typeface but different styles.

| Name       | Description |
| ---------- | ----------- |
| [GdipCloneFontFamily](#gdipclonefontfamily) | Creates a new **FontFamily** object based on another **FontFamily** object. |
| [GdipCreateFontFamilyFromName](#gdipcreatefontfamilyfromname) | Creates a **FontFamily** object based on a specified font family. |
| [GdipDeleteFontFamily](#gdipdeletefontfamily) | Deletes the specified **FontFamily** object. |
| [GdipFontCollectionEnumerable](#gdipfontcollectionenumerable) | Gets the number of fonts found in the collection. |
| [GdipFontCollectionEnumerate](#gdipfontcollectionenumerate) | Gets the number of fonts found in the collection. |
| [GdipGetCellAscent](#gdipgetcellascent) | Gets the cell ascent, in design units, of a font family for the specified style or style combination. |
| [GdipGetCellDescent](#GdipGetCellDescent) | Gets the cell descent, in design units, of a font family for the specified style or style combination. |
| [GdipGetEmHeight](#gdipgetemheight) | Gets the size (commonly called em size or em height), in design units, of a font family. |
| [GdipGetFamilyName](#gdipgetfamilyname) | Gets the name of a a family. |
| [GdipGetGenericFontFamilyMonospace](#gdipgetgenericfontfamilymonospace) | Gets a **FontFamily** object that specifies a generic monospace typeface. |
| [GdipGetGenericFontFamilySansSerif](#gdipgetgenericfontfamilysansserif) | Gets a **FontFamily** object that specifies a generic sans serif typeface. |
| [GdipGetGenericFontFamilySerif](#gdipgetgenericfontfamilyserif) | Gets a **FontFamily** object that specifies a generic serif typeface. |
| [GdipGetLineSpacing](#gdipgetlinespacing) | Gets the line spacing, in design units, of a font family for the specified style or style combination. |
| [GdipIsStyleAvailable](#gdipisstyleavailable) | Determines whether the specified style is available for a font family. |

---

## GdipCloneFontFamily

Creates a new **FontFamily** object based on another **FontFamily** object.

```
FUNCTION GdipCloneFontFamily (BYVAL FontFamily AS GpFontFamily PTR, _
   BYVAL clonedFontFamily AS GpFontFamily PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *FontFamily* | [in] Pointer to the **FontFamily** object to clone. |
| *clonedFontFamily* | [out] Pointer to a **FontFamily** object variable that receives a pointer to the new **FontFamily** object. |

#### Example

```
' ========================================================================================
' The following example clones a Font Family and uses it to draw text.
' ========================================================================================
SUB Example_CloneFontFamily (BYVAL hdc AS HDC)

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
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

   ' // Clone the font family
   DIM clonedFamily AS GpFontFamily PTR
   hStatus = GdipCloneFontFamily(fontFamily, @clonedFamily)

   ' // Create a font from the cloned family
   DIM font AS GpFont PTR
   hStatus = GdipCreateFont(clonedFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleItalic, UnitPixel, @font)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string using the clones font family
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "Text from cloned FontFamily"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF clonedFamily THEN GdipDeleteFontFamily(clonedFamily)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateFontFamilyFromName

Creates a **FontFamily** object based on a specified font family.

```
FUNCTION GdipCreateFontFamilyFromName (BYVAL name AS LPCWSTR, _
   BYVAL fontCollection AS GpFontCollection PTR, BYVAL fontFamily AS GpFontFamily PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *name* | [in] Name of the font family. For example, Arial.ttf is the name of the Arial font family. |
| *fontCollection* | [in] Pointer to a **FontCollection** object that specifies the collection that the font family belongs to. If **FontCollection** is **NULL**, the font family is not part of a collection. The default value is **NULL**. |
| *fontFamily* | [out] Pointer to a **FontFamily** object variable that receives a pointer to the new **FontFamily** object. |

#### Example

```
' ========================================================================================
' The following example creates a Font object from a family name and uses it to draw text.
' ========================================================================================
SUB Example_CreateFontFamilyFromName (BYVAL hdc AS HDC)

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
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   DIM font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "Text drawn using an Arial font"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeleteFontFamily

Deletes the specified **FontFamily** object.

```
FUNCTION GdipDeleteFontFamily (BYVAL fontFamily AS GpFontFamily PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontFamily* | [in] Pointer to the **FontFamily** object to delete. |

#### Example

```
DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create a font family
DIM fontFamily AS GpFontFamily PTR
hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

...
...
...

' // Cleanup
IF fontFamily THEN GdipDeleteFontFamily(fontFamily)

' // Shutdown GDI+
AfxGdipShutdown token

```
---

## GdipFontCollectionEnumerable

Not implemented.

```
FUNCTION GdipFontCollectionEnumerable (BYVAL fontCollection AS GpFontCollection PTR, _
   BYVAL graphics AS GpGraphics PTR, BYVAL numFound AS INT_ PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontCollection* | [in] Pointer to the font collection object. |
| *graphics* | [in] Pointer to the **Graphics** object. |
| *numFound* | [out] Pointer to a long integer variable that receives the number of fonts found in the collection. |

---

## GdipFontCollectionEnumerate

Not implemented.

```
FUNCTION GdipFontCollectionEnumerate (BYVAL fontCollection AS GpFontCollection PTR, _
   BYVAL numSought AS INT_, BYVAL gpfamilies AS GpFontFamily PTR PTR, BYVAL numFound AS INT_ PTR, _
   BYVAL graphics AS GpGraphics PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontCollection* | [in] Pointer to the font collection object. |
| *numSought* | [in] The number of font families in the font collection. |
| *gpfamilies* | [out] Pointer to an array that receives the **FontFamily** objects. |
| *numFound* | [out] Pointer to a long integer variable that receives the number of fonts found in the collection. |
| *graphics* | [in] Pointer to the Graphics object. |

---

## GdipGetCellAscent

Gets the cell ascent, in design units, of a font family for the specified style or style combination.

```
FUNCTION GdipGetCellAscent (BYVAL family AS GpFontFamily PTR, BYVAL style AS INT_, _
   BYVAL CellAscent AS UINT16) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *family* | [in] Pointer to the **FontFamily** object. |
| *style* | [in] Long integer value that specifies the style of the typeface. This value must be an element of the *FontStyle* enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout specifies a combination of the three styles. |
| *CellAscent* | [out] Pointer to a variable that receives the cell ascent, in design units, of a font family for the specified style or style combination. |

#### Example

This example gets the cell ascent of a font family in design units, which is essential for calculating vertical text metrics like baseline alignment and line spacing.

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create the font family
DIM fontFamily AS GpFontFamily PTR
hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

' // Get the cell ascent for regular style
DIM ascent AS UINT16
hStatus = GdipGetCellAscent(fontFamily, FontStyleRegular, @ascent)

' // Display result
PRINT "Cell Ascent (design units): "; ascent

' // Cleanup
IF fontFamily THEN GdipDeleteFontFamily(fontFamily)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---

## GdipGetCellDescent

Gets the cell descent, in design units, of a font family for the specified style or style combination.

```
FUNCTION GdipGetCellDescent (BYVAL family AS GpFontFamily PTR, BYVAL style AS INT_, _
   BYVAL CellDescent AS UINT16 PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *family* | [in] Pointer to the **FontFamily** object. |
| *style* | [in] Long integer value that specifies the style of the typeface. This value must be an element of the *FontStyle* enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout specifies a combination of the three styles. |
| *CellDescent* | [out] Pointer to a variable that receives the cell descent, in design units, of a font family for the specified style or style combination. |

#### Example

```
' ========================================================================================
' This example gets the cell descent of a font family in design units.
' Then combines cell ascent, cell descent, and line spacing to calculate the total line
' height in pixels for a font
' This is crucial for custom layout engines, text editors, or precise UI rendering.
' ascentPixels + descentPixels gives you the actual height of glyphs.
' spacingPixels gives you the recommended line height for consistent vertical spacing.
' ========================================================================================
DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create the font family
DIM fontFamily AS GpFontFamily PTR
hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

' // Font style and size
DIM style AS LONG = FontStyleRegular
DIM fontSize AS SINGLE = 16.0

' // Get metrics in design units
DIM ascent AS UINT16, descent AS UINT16, spacing AS UINT16, emHeight AS UINT16
GdipGetCellAscent(fontFamily, style, @ascent)
GdipGetCellDescent(fontFamily, style, @descent)
GdipGetLineSpacing(fontFamily, style, @spacing)
GdipGetEmHeight(fontFamily, style, @emHeight)

' // Convert to pixels
DIM ascentPixels AS SINGLE = (fontSize * ascent) / emHeight
DIM descentPixels AS SINGLE = (fontSize * descent) / emHeight
DIM spacingPixels AS SINGLE = (fontSize * spacing) / emHeight

' // Display results
PRINT "Ascent (px): "; ascentPixels
PRINT "Descent (px): "; descentPixels
PRINT "Line spacing (px): "; spacingPixels

' // Cleanup
IF fontFamily THEN GdipDeleteFontFamily(fontFamily)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---

## GdipGetEmHeight

Gets the size (commonly called em size or em height), in design units, of a font family.

```
FUNCTION GdipGetEmHeight (BYVAL family AS GpFontFamily PTR, BYVAL style AS INT_, _
   BYVAL EmHeight AS UINT16 PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *family* | [in] Pointer to the **FontFamily** object. |
| *style* | [in] Long integer value that specifies the style of the typeface. This value must be an element of the *FontStyle* enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout specifies a combination of the three styles. |
| *EmHeight* | [out] Pointer to a variable that receives the size, in design units, of a font family. |

#### Example

This example gets the em height of a font family for a given style. This value is expressed in design units, and it's crucial for converting other font metrics (like ascent, descent, and line spacing) into pixel values. For example, if the em height is 2048 and the ascent is 1638, then: ascentPixels = (fontSize * 1638) / 2048

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create the font family
DIM fontFamily AS GpFontFamily PTR
hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

' // Get em height for regular style
DIM emHeight AS UINT16
hStatus = GdipGetEmHeight(fontFamily, FontStyleRegular, @emHeight)

' // Display result
PRINT "Em Height (design units): "; emHeight

' // Cleanup
IF fontFamily THEN GdipDeleteFontFamily(fontFamily)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---

## GdipGetFamilyName

Gets the name of a font family.

```
FUNCTION GdipGetFamilyName (BYVAL family AS GpFontFamily PTR, BYVAL name AS WSTRING PTR, _
   BYVAL language AS LANGID) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *family* | [in] Pointer to the **FontFamily** object. |
| *name* | [out] Name of the font family. |
| *language* | [in] Sixteen-bit value that specifies the language to use. The default value is LANG_NEUTRAL, which is the user's default language. |

#### Example

This example gets the font family name. The name returned is typically the canonical name like "Arial", "Times New Roman", etc. This is handy when listing fonts from a collection or verifying font availability.

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create the font family
DIM fontFamily AS GpFontFamily PTR
hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

' // Get the name of the font family
DIM wszName AS WSTRING * LF_FACESIZE
hStatus = GdipGetFamilyName(fontFamily, @wszName, LANG_NEUTRAL)

' Display result
PRINT "Font family name: "; wszName

' // Cleanup
IF fontFamily THEN GdipDeleteFontFamily(fontFamily)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---

## GdipGetGenericFontFamilyMonospace

Gets a **FontFamily** object that specifies a generic monospace typeface.

```
FUNCTION GdipGetGenericFontFamilyMonospace (BYVAL nativeFamily AS GpFontFamily PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *nativeFamily* | [out] Pointer to a **FontFamily** object variable that receives a pointer to the **FontFamily** object that specifies a generic monospace typeface. |

```
' ========================================================================================
' The following example gets the generic monospace font family and draws a string with it.
' ========================================================================================
SUB Example_GetGenericFontFamilyMonospace (BYVAL hdc AS HDC)

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
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Get the generic monospace font family
   DIM monoFamily AS GpFontFamily PTR
   hStatus = GdipGetGenericFontFamilyMonospace(@monoFamily)

   ' Create a font from the generic monospace family
   DIM font AS GpFont PTR
   hStatus = GdipCreateFont(monoFamily, AfxGdipPointsToPixels(14, TRUE), FontStyleRegular, UnitPixel, @font)

   ' // Create a brush
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' // Draw a string
   DIM text AS WSTRING * 64 = "Fixed-width text example"
   DIM layout AS GpRectF = (20.0, 20.0, 300.0, 100.0)
   hStatus = GdipDrawString(graphics, text, LEN(text), font, @layout, NULL, brush)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF monoFamily THEN GdipDeleteFontFamily(monoFamily)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetGenericFontFamilySansSerif

Gets a **FontFamily** object that specifies a generic sans serif typeface.

```
FUNCTION GdipGetGenericFontFamilySansSerif (BYVAL nativeFamily AS GpFontFamily PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *nativeFamily* | [out] Pointer to a **FontFamily** object variable that receives a pointer to the **FontFamily** object that specifies a generic sans serif typeface. |

#### Example

```
' ========================================================================================
' The following example gets the generic sans serif font family and draws a string with it.
' ========================================================================================
SUB Example_GetGenericFontFamilySansSerif (BYVAL hdc AS HDC)

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
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Get the generic sans serif font family
   DIM monoFamily AS GpFontFamily PTR
   hStatus = GdipGetGenericFontFamilySansSerif(@monoFamily)

   ' Create a font from the generic monospace family
   DIM font AS GpFont PTR
   hStatus = GdipCreateFont(monoFamily, AfxGdipPointsToPixels(16, TRUE), FontStyleRegular, UnitPixel, @font)

   ' // Create a brush
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' // Draw a string
   DIM text AS WSTRING * 64 = "SansSerif text example"
   DIM layout AS GpRectF = (20.0, 20.0, 300.0, 100.0)
   hStatus = GdipDrawString(graphics, text, LEN(text), font, @layout, NULL, brush)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF monoFamily THEN GdipDeleteFontFamily(monoFamily)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetGenericFontFamilySerif

Gets a **FontFamily** object that specifies a generic serif typeface.

```
FUNCTION GdipGetGenericFontFamilySerif (BYVAL nativeFamily AS GpFontFamily PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *nativeFamily* | [out] Pointer to a **FontFamily** object variable that receives a pointer to the **FontFamily** object that specifies a generic serif typeface. |

#### Example

```
' ========================================================================================
' The following example gets the generic serif font family and draws a string with it.
' ========================================================================================
SUB Example_GetGenericFontFamilySerif (BYVAL hdc AS HDC)

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
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Get the generic sans serif font family
   DIM monoFamily AS GpFontFamily PTR
   hStatus = GdipGetGenericFontFamilySerif(@monoFamily)

   ' Create a font from the generic monospace family
   DIM font AS GpFont PTR
   hStatus = GdipCreateFont(monoFamily, AfxGdipPointsToPixels(16, TRUE), FontStyleRegular, UnitPixel, @font)

   ' // Create a brush
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' // Draw a string
   DIM text AS WSTRING * 64 = "Serif text example"
   DIM layout AS GpRectF = (20.0, 20.0, 300.0, 100.0)
   hStatus = GdipDrawString(graphics, text, LEN(text), font, @layout, NULL, brush)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF monoFamily THEN GdipDeleteFontFamily(monoFamily)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLineSpacing

Gets the line spacing, in design units, of a font family for the specified style or style combination. The line spacing is the vertical distance between the base lines of two consecutive lines of text.

```
FUNCTION GdipGetLineSpacing (BYVAL family AS GpFontFamily PTR, BYVAL style AS INT_, _
   BYAL LineSpacing AS UINT16 PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *family* | [in] Pointer to the **FontFamily** object. |
| *style* | [in] Integer that specifies the style of the typeface. This value must be an element of the *FontStyle* enumeration or the result of a bitwise OR applied to two or more of these elements. For example, FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout specifies a combination of the three styles. |
| *LineSpacing* | [out] Pointer to a variable that receives the line spacing of the font family. |

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

' ========================================================================================
' This example retrieves the recommended vertical distancebetween baselines of consecutive
' lines of text for a given font family and style. This value is returned in design units,
' and it's essential for calculating consistent line heights in custom text layouts.
' It includes the cell ascent, cell descent, and any internal leading.
' It defines how much vertical space should be allocated between lines of text.
' It’s not necessarily the sum of ascent and descent—it may include extra spacing for readability.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create a font family
DIM fontFamily AS GpFontFamily PTR
hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

' // Get line spacing for regular style
DIM spacing AS UINT16
hStatus = GdipGetLineSpacing(fontFamily, FontStyleRegular, @spacing)

' // Display result
PRINT "Line Spacing (design units): "; spacing

' // Cleanup
IF fontFamily THEN GdipDeleteFontFamily(fontFamily)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---

## GdipIsStyleAvailable

Determines whether the specified style is available for a font family.

```
FUNCTION GdipIsStyleAvailable (BYVAL family AS GpFontFamily PTR, BYVAL style AS INT_, _
   BYAL IsStyleAvailable AS BOOL PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *family* | [in] Pointer to the **FontFamily** object. |
| *style* | [in] Long integer value that specifies the style of the typeface. This value must be an element of the *FontStyle* enumeration or the result of a bitwise **OR** applied to two or more of these elements. For example, *FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout* specifies a combination of the three styles.
| *IsStyleAvailable* | [out] Pointer to a variable that receives a boolean value indicating if the style or combination of styles is available (TRUE) or not (FALSE). |

#### Remarks

This function returns a misleading result on some third-party fonts. For example, *FontStyleUnderline* may return FALSE because it is really testing for a regular style font that also is an underlined font: (*FontStyleRegular OR FontStyleUnderline*). If the font does not have a regular style, the **GdipIsStyleAvailabl**e function returns FALSE.

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

' ========================================================================================
' This example checks whether a specific font style (like bold, italic, underline, etc.)
' is available for a given GpFontFamily. It’s especially useful when working with third-party
' or custom fonts that may not support all style variants.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create a font family
DIM fontFamily AS GpFontFamily PTR
hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

' // Check if italic style is available
DIM isAvailable AS BOOL
hStatus = GdipIsStyleAvailable(fontFamily, FontStyleItalic, @isAvailable)

IF isAvailable THEN
  PRINT "Italic style is available for Arial."
ELSE
  PRINT "Italic style is NOT available for Arial."
END IF

' // Cleanup
IF fontFamily THEN GdipDeleteFontFamily(fontFamily)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---
