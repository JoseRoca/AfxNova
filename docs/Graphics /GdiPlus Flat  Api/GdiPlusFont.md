# Font Functions

A **Font** object is used when drawing strings.

| Name       | Description |
| ---------- | ----------- |
| [GdipCloneFont](#gdipclonefont) | Creates a new **Font** object based on another **Font** object. |
| [GdipCreateFont](#gdipcreatefont) | Creates a **Font** object based on a **FontFamily** object, a size, a font style, and a unit of measurement. |
| [GdipCreateFontFromDC](#gdipcreatefontfromdc) | Creates a **Font** object based on the Microsoft® Windows® Graphics Device Interface (GDI) font object that is currently selected into a specified device context. |
| [GdipCreateFontFromLogfontA](#gdipcreatefontfromlogfonta) | Creates a **Font** object directly from a Microsoft® Windows® Graphics Device Interface (GDI) logical font. |
| [GdipCreateFontFromLogfontW](#gdipcreatefontfromlogfontw) | Creates a **Font** object directly from a Microsoft® Windows® Graphics Device Interface (GDI) logical font. |
| [GdipDeleteFont](#gdipdeletefont) | Deletes the specified **Font** object. |
| [GdipDeletePrivateFontCollection](#gdipdeleteprivatefontcollection) | Deletes the specified **PrivateFontCollection** object. |
| [GdipGetFamily](#gdipgetfamily) | Gets the font family on which the font is based. |
| [GdipGetFontCollectionFamilyCount](#gdipgetfontcollectionfamilycount) | Gets the number of font families contained in the font collection. |
| [GdipGetFontCollectionFamilyList](#gdipgetfontcollectionfamilylist) | Gets the font families contained in the font collection. |
| [GdipGetFontHeight](#gdipgetfontheight) | Gets the line spacing of the font in the current unit of a specified **Graphics* object. |
| [GdipGetFontHeightGivenDPI](#gdipgetfontheightgivendpi) | Gets the line spacing, in pixels, of the font. |
| [GdipGetFontSize](#gdipgetfontsize) | Returns the font size (commonly called the *em* size) of the Font object. The size is in the units of the **Font** object. |
| [GdipGetFontStyle](#gdipgetfontstyle) | Gets the style of the font's typeface. |
| [GdipGetFontUnit](#gdipgetfontunit) | Returns the unit of measure of tes **Font** object. |
| [GdipGetLogFontA](#gdipgetlogfonta) | Uses a **LOGFONTA** structure to get the attributes of the **Font** object. |
| [GdipGetLogFontW](#gdipgetlogfontw) | Uses a **LOGFONTW** structure to get the attributes of the **Font** object. |
| [GdipNewInstalledFontCollection](#gdipnewinstalledfontcollection) | Creates an **InstalledFontCollection** object. |
| [GdipNewPrivateFontCollection](#gdipnewprivatefontcollection) | Creates an **PrivateFontCollection** object. |
| [GdipPrivateAddFontFile](#gdipprivateaddfontfile) | Adds a font file to the private font collection. |
| [GdipPrivateAddMemoryFont](#gdipprivateaddmemoryfont) | Adds a font that is contained in system memory to a Microsoft® Windows® GDI+ font collection. |

---

# Helper function

| Name       | Description |
| ---------- | ----------- |
| [AfxGdipPointsToPixels](#afxgdippointstopixels) | Convert font point sizes to pixels and scale then for dpi awareness if needed. |

---

## GdipCloneFont

Creates a new Font object based on another Font object.

```
FUNCTION GdipCloneFont (BYVAL font AS GpFont PTR, BYVAL cloneFont AS GpFont PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *font* | [in] Pointer to the **Font** object to clone. |
| *cloneFont* | [out] Pointer to a **GpFont** variable that receives a pointer to the new created **Font** object. |

#### Example

```
' ========================================================================================
' The following example creates a Font object, clones it, and then uses the clone to draw text.
' ========================================================================================
SUB Example_CloneFont (BYVAL hdc AS HDC)

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
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(24, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a clone of the font
   DIM clonedFont AS GpFont PTR
   hStatus = GdipCloneFont(font, @clonedFont)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "This is a cloned Font"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), clonedFont, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF clonedFont THEN GdipDeleteFont(clonedFont)
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateFont

Creates a **Font** object based on a **FontFamily** object, a size, a font style, and a unit of measurement.

```
FUNCTION GdipCreateFont (BYVAL fontFamily AS GpFontFamily PTR, BYVAL emSize AS REAL, _
   BYVAL style AS INT_, BYVAL unit AS GpUnit, BYVAL font AS GpFont PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontFamily* | [in] Pointer to a **FontFamily** object that specifies information such as the string that identifies the font family and the font family's text metrics measured in design units. |
| *emSize* | [in] Single precision value that specifies the em size of the font measured in the units specified in the unit parameter. |
| *style* | [[in] Long integer value that specifies the style of the typeface. This value must be an element of the FontStyle enumeration or the result of a bitwise OR applied to two or more of these elements. For example, %FontStyleBold OR %FontStyleUnderline OR %FontStyleStrikeout sets the style as a combination of the three styles. The default value is %FontStyleRegular |
| *unit* | [in] Element of the *GpUnit* enumeration that specifies the unit of measurement for the font size. The default value is *UnitPoint*. |
| *font* | [out] Pointer to a variable that receives a pointer to the new created **Font** object. |

```
' ========================================================================================
' The following example creates a Font object and then uses it to draw text.
' ========================================================================================
SUB Example_CreateFont (BYVAL hdc AS HDC)

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

#### Example

```
' ========================================================================================
' The following example creates a Font object from a font family and then uses it to draw text.
' ========================================================================================
SUB Example_CreateFontFromFamily (BYVAL hdc AS HDC)

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

   ' // Create a font family
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

   ' // Create a font from the font family
   DIM font AS GpFont PTR
   IF hStatus = StatusOk THEN
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(14, TRUE), FontStyleRegular, UnitPixel, @font)
   END IF

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "This is a Font created from a FontFamily"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateFontFromDC

Creates a **Font** object based on the Microsoft® Windows® Graphics Device Interface (GDI) font object that is currently selected into a specified device context. This constructor is provided for compatibility with GDI.

```
FUNCTION GdipCreateFontFromDC (BYVAL hdc AS HDC, BYVAL font AS GpFont PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | [in] Handle to a Windows device context that has a font selected. A handle is a number that Windows uses internally to reference an object. |
| *font* | [out] Pointer to a variable that receives a pointer to the new created **Font** object. |

#### Remarks

In most cases when you obtain a device context handle by calling the **GGdipGetDC** function, the device context will not have a font selected. If you pass such a handle to this constructor, the constructor will fail.

A device context is a structure that is maintained internally. It is associated with a particular device, such as a video monitor or a printer. There is usually one device context associated with each window displayed on a video monitor. A device context contains some graphics attributes used by GDI+.

---

## GdipCreateFontFromLogfontA

Creates a **Font** object directly from a Microsoft® Windows® Graphics Device Interface (GDI) logical font. The GDI logical font is a **LOGFONTA** structure, which is the one-byte character version of a logical font. This constructor is provided for compatibility with GDI.

```
FUNCTION GdipCreateFontFromLogfontA (BYVAL hdc AS HDC, BYVAL logfont AS LOGFONTA PTR, _
   BYVAL font AS GpFont PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | [in] Handle to a Windows device context. A handle is a number that Windows uses internally to reference an object. |
| *logfont* | [in] Pointer to a **LOGFONTA** structure that contains attributes of the font. The **LOGFONTA** structure is the one-byte character version of the logical font. |
| *font* | [out] Pointer to a variable that receives a pointer to the new created **Font** object. |

#### Remarks

A device context is a structure that is maintained internally. It is associated with a particular device, such as a video monitor or a printer. There is usually one device context associated with each window displayed on a video monitor. A device context contains some graphics attributes used by GDI+.

A **LOGFONTA** structure is a GDI structure. GDI+ uses only some of the attributes contained in this structure.

#### Example

```
' ========================================================================================
' The following example creates a Font object from a LogfontA structure,
' ========================================================================================
SUB Example_GdipCreateFontFromLogfontA (BYVAL hdc AS HDC)

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

   ' // Create a LOGFONTW structure
   DIM lf AS LOGFONTA
   WITH lf
      .lfHeight = -24
      .lfWidth = 0
      .lfEscapement = 0
      .lfOrientation = 0
      .lfWeight = FW_NORMAL
      .lfItalic = FALSE
      .lfUnderline = FALSE
      .lfStrikeOut = FALSE
      .lfCharSet = DEFAULT_CHARSET
      .lfOutPrecision = OUT_DEFAULT_PRECIS
      .lfClipPrecision = CLIP_DEFAULT_PRECIS
      .lfQuality = DEFAULT_QUALITY
      .lfPitchAndFamily = FF_DONTCARE
      .lfFaceName = "Arial"
   END WITH

   ' // Create the font
   DIM font AS GpFont PTR
   hStatus = GdipCreateFontFromLogfontA(hdc, @lf, @font)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "Font from LOGFONTA"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateFontFromLogfontW

Creates a **Font** object directly from a Microsoft® Windows® Graphics Device Interface (GDI) logical font. The GDI logical font is a **LOGFONTW** structure, which is the one-byte character version of a logical font. This constructor is provided for compatibility with GDI.

```
FUNCTION GdipCreateFontFromLogfontA (BYVAL hdc AS HDC, BYVAL logfont AS LOGFONTW PTR, _
   BYVAL font AS GpFont PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *hdc* | [in] Handle to a Windows device context. A handle is a number that Windows uses internally to reference an object. |
| *logfont* | [in] Pointer to a **LOGFONTW** structure that contains attributes of the font. The **LOGFONTW** structure is the one-byte character version of the logical font. |
| *font* | [out] Pointer to a variable that receives a pointer to the new created **Font** object. |

#### Remarks

A device context is a structure that is maintained internally. It is associated with a particular device, such as a video monitor or a printer. There is usually one device context associated with each window displayed on a video monitor. A device context contains some graphics attributes used by GDI+.

A **LOGFONTW** structure is a GDI structure. GDI+ uses only some of the attributes contained in this structure.

#### Example

```
' ========================================================================================
' The following example creates a Font object from a LogfontW structure,
' ========================================================================================
SUB Example_GdipCreateFontFromLogfontW (BYVAL hdc AS HDC)

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

   ' // Create a LOGFONTW structure
   DIM lf AS LOGFONTW
   WITH lf
      .lfHeight = -24
      .lfWidth = 0
      .lfEscapement = 0
      .lfOrientation = 0
      .lfWeight = FW_NORMAL
      .lfItalic = FALSE
      .lfUnderline = FALSE
      .lfStrikeOut = FALSE
      .lfCharSet = DEFAULT_CHARSET
      .lfOutPrecision = OUT_DEFAULT_PRECIS
      .lfClipPrecision = CLIP_DEFAULT_PRECIS
      .lfQuality = DEFAULT_QUALITY
      .lfPitchAndFamily = FF_DONTCARE
      .lfFaceName = "Arial"
   END WITH

   ' // Create the font
   DIM font AS GpFont PTR
   hStatus = GdipCreateFontFromLogfontW(hdc, @lf, @font)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "Font from LOGFONTW"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeleteFont

Deletes the specified **Font** object.

```
FUNCTION GdipDeleteFont (BYVAL font AS GpFont PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *font* | [in] Pointer to the **Font** object to delete. |

```
' ========================================================================================
' The following example creates a Font object, draws a string and deletes the font.
' ========================================================================================
SUB Example_DeleteFont (BYVAL hdc AS HDC)

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
      hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(24, TRUE), FontStyleRegular, UnitPixel, @font)
      GdipDeleteFontFamily(fontFamily)
   END IF

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "This is an Arial Font"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeletePrivateFontCollection

Deletes the specified **PrivateFontCollection** object.

```
FUNCTION GdipDeletePrivateFontCollection (BYVAL fontCollection AS GpFontCollection PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontCollection* | [in] Pointer to the **PrivateFontCollection** object to delete. |

---

## GdipGetFamily

Gets the font family on which a font is based.

```
FUNCTION GdipGetFamily (BYVAL family AS GpFontFamily PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *family* | [out] Pointer to a **FontFamily** object variable that receives a pointer to the font family. |

```
' ========================================================================================
' The following example creates a Font object, retrieves the information about the font
' family on which it is based, and then uses the FontFamily object to create a second Font
' object. The example then uses the second Font object to draw text.
' ========================================================================================
SUB Example_GetFontFamily (BYVAL hdc AS HDC)

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

   ' // Create a font family
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

   ' // Create a font from the font family
   DIM font AS GpFont PTR
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(14, TRUE), FontStyleRegular, UnitPixel, @font)

   ' // Get the font family of the font
   DIM fontFamily2 AS GpFontFamily PTR
   hStatus = GdipGetFamily(font, @fontFamily2)

   ' // Create a second font from the retrieved font family
   DIM familyFont AS GpFont PTR
   hStatus = GdipCreateFont(fontFamily2, AfxGdipPointsToPixels(14, TRUE), FontStyleRegular, UnitPixel, @familyFont)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "This is a Font created from a FontFamily"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), familyFont, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF familyFont THEN GdipDeleteFont(familyFont)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetFontCollectionFamilyCount

Gets the number of font families contained in the font collection.

```
FUNCTION GdipGetFontCollectionFamilyCount (BYVAL fontCollection AS GpFontCollection PTR, _
   BYVAL numFound AS INT_ PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontCollection* | [in] Pointer to the font collection object. |
| *numFound* | [out] Pointer to a long integer variable that receives the number of font families contained in the font collection. |

#### Remarks

A font family consists of a single font type with related styles. An example of a single font type is Arial Regular. An example of a font family is a set of fonts containing Arial Regular, Arial Italic, and Arial Bold style fonts.

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

' ========================================================================================
' Gets the number of font families installed.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create an InstalledFontCollection object
DIM fontCollection AS GpFontCollection PTR
hStatus = GdipNewInstalledFontCollection(@fontCollection)

' // Get the number of font families in the collection
DIM numberFonts AS LONG
hStatus = GdipGetFontCollectionFamilyCount(fontCollection, @numberFonts)

' // Display the result (for example, in console)
PRINT "Number of installed font families: "; numberFonts

' // Buffer for font name
DIM wszName AS WSTRING * 32

' // Allocate array for font family pointers
DIM families(numberFonts - 1) AS GpFontFamily PTR
hStatus = GdipGetFontCollectionFamilyList(fontCollection, numberFonts, @families(0), @numberFonts)
    
' // Print each font family name
FOR i AS LONG = 0 TO numberFonts - 1
   hStatus = GdipGetFamilyName(families(i), @wszName, LANG_NEUTRAL)
   PRINT "Font "; i + 1; ": "; wszName
   GdipDeleteFontFamily(families(i))
NEXT

' // Cleanup
' No delete function exists for installed font collections
' //because it is a shared collection; cleanup is automatic”

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
' ========================================================================================
```
---

## GdipGetFontCollectionFamilyList

Gets the font families contained in the font collection.

```
FUNCTION GdipGetFontCollectionFamilyList (BYVAL fontCollection AS GpFontCollection PTR, _
   BYVAL numSought AS INT_, BYVAL gpfamilies AS GpFontFamily PTR PTR, BYVAL numFound AS INT_ PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontCollection* | [in] Pointer to the font collection object. |
| *numSought* | [in] Long integer value  that specifies the number of font families in the font collection. |
| *gpfamilies* | [out] Pointer to an array that receives the **FontFamily** objects. |
| *numFound* | [out] Pointer to a **LONG** that receives the number of font families found in the collection. This number should be the same as the *numSought* value. |

#### Remarks

A font family consists of a single font type with related styles. An example of a single font type is Arial Regular. An example of a font family is a set of fonts containing Arial Regular, Arial Italic, and Arial Bold style fonts.

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

' ========================================================================================
' Gets the number of font families installed.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create an InstalledFontCollection object
DIM fontCollection AS GpFontCollection PTR
hStatus = GdipNewInstalledFontCollection(@fontCollection)

' // Get the number of font families in the collection
DIM numberFonts AS LONG
hStatus = GdipGetFontCollectionFamilyCount(fontCollection, @numberFonts)

' // Display the result (for example, in console)
PRINT "Number of installed font families: "; numberFonts

' // Buffer for font name
DIM wszName AS WSTRING * 32

' // Allocate array for font family pointers
DIM families(numberFonts - 1) AS GpFontFamily PTR
hStatus = GdipGetFontCollectionFamilyList(fontCollection, numberFonts, @families(0), @numberFonts)
    
' // Print each font family name
FOR i AS LONG = 0 TO numberFonts - 1
   hStatus = GdipGetFamilyName(families(i), @wszName, LANG_NEUTRAL)
   PRINT "Font "; i + 1; ": "; wszName
   GdipDeleteFontFamily(families(i))
NEXT

' // Cleanup
' No delete function exists for installed font collections
' //because it is a shared collection; cleanup is automatic”

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP

' ========================================================================================
```
---

## GdipGetFontHeight

Gets the line spacing of the font in the current unit of a specified **Graphics** object. The line spacing is the vertical distance between the base lines of two consecutive lines of text. Thus, the line spacing includes the blank space between lines along with the height of the character itself.

```
FUNCTION GdipGetFontHeight (BYVAL font AS GpFont PTR, BYVAL graphics AS GpGraphics PTR, _
   BYVAL height AS REAL PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *font* | [in] Pointer to the **Font** object. |
| *graphics* | [[in] Pointer to a **Graphics** object whose unit and vertical resolution are used in the height calculation. |
| *height* | [[out] Pointer to a single variable that receives te line spacing of the font. |

#### Remarks

If the font unit is set to anything other than *UnitPixel*, the height, in pixels, is calculated using the vertical resolution of the specified **Graphics** object. For example, suppose the font unit is inches and the font size is 0.3. Also suppose that for the corresponding font family, the em height is 2048 and the line spacing is 2355. If the unit of the **Graphics** object is UnitPixel and the vertical resolution of the **Graphics** object is 96 dots per inch, the height is calculated as follows:

2355*(0.3/2048)*96 = 33.1171875

Continuing with the same example, suppose the unit of the **Graphics** object is something other than *UnitPixel*, say *UnitMillimeter*. Then (using 1 inch = 25.4 millimeters) the height, in millimeters, is calculated as follows:

2355*(0.3/2048)25.4 = 8.762256

#### Example

```
' ========================================================================================
' The following example creates a Font object, retrieves the height of the Font object, and
' uses the height to position two lines of text, with the second line directly below the first.
' ========================================================================================
SUB Example_GetFontHeight (BYVAL hdc AS HDC)

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

   ' // Create a font family
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

   ' // Create a font from the font family
   DIM font AS GpFont PTR
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)

   ' // Create a black solid brush
   DIM brush1 AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush1)

   ' // Draw first line of text
   DIM text1 AS WSTRING * 64 = "The first line of text"
   DIM rcf1 AS GpRectF = (0, 0, 300, 100)
   hStatus = GdipDrawString(graphics, text1, LEN(text1), font, @rcf1, NULL, brush1)

   ' // Get the font height
   DIM height AS SINGLE
   hStatus = GdipGetFontHeight(font, graphics, @height)

   ' // Create second brush (red)
   DIM brush2 AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @brush2)

   ' // Draw second line of text below the first
   DIM text2 AS WSTRING * 64 = "The second line of text"
   DIM rcf2 AS GpRectF = (0, height, 300, 100)
   hStatus = GdipDrawString(graphics, text2, LEN(text2), font, @rcf2, NULL, brush2)

   ' // Cleanup
   IF brush1 THEN GdipDeleteBrush(brush1)
   IF brush2 THEN GdipDeleteBrush(brush2)
   IF font THEN GdipDeleteFont(font)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetFontHeightGivenDPI

Gets the line spacing, in pixels, of the font. The line spacing is the vertical distance between the base lines of two consecutive lines of text. Thus, the line spacing includes the blank space between lines along with the height of the character itself.

```
FUNCTION GdipGetFontHeightGivenDPI (BYVAL font AS GpFont PTR, BYVAL dpi AS REAL, _
   BYVAL height AS REAL PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *font* | [in] Pointer to the **Font** object. |
| *dpi* | [iin] Single precision value that specifies the vertical resolution, in dots per inch, of the device that displays the font. |
| *height* | [out] Pointer to a single variable that receives the line spacing of the font in pixels. |

#### Remarks

If the font unit is set to anything other than *UnitPixel*, the height, in pixels, is calculated using the specified vertical resolution. For example, suppose the font unit is inches and the font size is 0.3. Also suppose that for the corresponding font family, the em height is 2048 and the line spacing is 2355. If the specified vertical resolution is 96 dots per inch, the height is calculated as follows:

2355*(0.3/2048)*96 = 33.1171875

#### Example

```
' ========================================================================================
' The following example creates a Font object, retrieves the height of the Font object, and
' uses the height to position two lines of text, with the second line directly below the first.
' ========================================================================================
SUB Example_GetFontHeightGivenDpi (BYVAL hdc AS HDC)

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

   ' // Create a font family
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

   ' // Create a font from the font family
   DIM font AS GpFont PTR
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)

   ' // Create a black solid brush
   DIM brush1 AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush1)

   ' // Draw first line of text
   DIM text1 AS WSTRING * 64 = "The first line of text"
   DIM rcf1 AS GpRectF = (0, 0, 300, 100)
   hStatus = GdipDrawString(graphics, text1, LEN(text1), font, @rcf1, NULL, brush1)

   ' // Get the font height
   DIM height AS SINGLE
   hStatus = GdipGetFontHeightGivenDpi(font, dpiY, @height)

   ' // Create second brush (red)
   DIM brush2 AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @brush2)

   ' // Draw second line of text below the first
   DIM text2 AS WSTRING * 64 = "The second line of text"
   DIM rcf2 AS GpRectF = (0, height, 300, 100)
   hStatus = GdipDrawString(graphics, text2, LEN(text2), font, @rcf2, NULL, brush2)

   ' // Cleanup
   IF brush1 THEN GdipDeleteBrush(brush1)
   IF brush2 THEN GdipDeleteBrush(brush2)
   IF font THEN GdipDeleteFont(font)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetFontSize

Returns the font size (commonly called the *em* size) of the **Font** object. The size is in the units of the **Font** object.

```
FUNCTION GdipGetFontSize (BYVAL font AS GpFont PTR, BYVAL size AS REAL PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *font* | [in] Pointer to the **Font** object. |
| *size* | [out] Pointer to a single variable that receives the font size in the units of the Font object. |

```
' ========================================================================================
' The following example creates a Font object, gets the size of the font, and creates a
' second Font object of the same size as the first. The second Font object is then used to
' draw text.
' ========================================================================================
SUB Example_GetFontSize (BYVAL hdc AS HDC)

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
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)

   ' // Get the font size
   DIM nSize AS REAL
   hStatus = GdipGetFontSize(font, @nSize)

   ' // Create a second font object with the same size
   DIM font2 AS GpFont PTR
   hStatus = GdipCreateFont(fontFamily, nSize, FontStyleRegular, UnitPixel, @font2)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "Font with an acquired size"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font2, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF font2 THEN GdipDeleteFont(font2)
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetFontStyle

Gets the style of the font's typeface.

```
FUNCTION GdipGetFontStyle (BYVAL font AS GpFont PTR, BYVAL style AS INT_ PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *font* | [in] Pointer to the **Font** object. |
| *style* | [out] Pointer to a variable that receives the style of the font's typeface. The value is an element of the *FontStyle* enumeration or the result of a bitwise OR applied to two or more elements of the *FontStyle* enumeration. |

#### Example

```
' ========================================================================================
' The following example creates a Font object, gets the FontStyle associated with the font,
' and creates a second Font object by using the same style. The second Font object is then
' used to draw text.
' ========================================================================================
SUB Example_GetFontStyle (BYVAL hdc AS HDC)

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
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleItalic, UnitPixel, @font)

   ' // Get the font size
   DIM style AS LONG
   hStatus = GdipGetFontStyle(font, @style)

   ' // Create a second font object with the same size
   DIM font2 AS GpFont PTR
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), style, UnitPixel, @font2)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "Font with an acquired style"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font2, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF font2 THEN GdipDeleteFont(font2)
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetFontUnit

Returns the unit of measure of the **Font** object.

```
FUNCTION GdipGetFontUnit (BYVAL font AS GpFont PTR, BYVAL unit AS GpUnit PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *font* | [in] Pointer to the **Font** object. |
| *unit* | [out] Pointer to a variable that receives one of the elements of the *GpUnit* enumeration. |

#### Example

```
' ========================================================================================
' The following example creates a Font object, gets the FontStyle associated with the font,
' and creates a second Font object by using the same style. The second Font object is then
' used to draw text.
' ========================================================================================
SUB Example_GetFontUnit (BYVAL hdc AS HDC)

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
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleItalic, UnitPixel, @font)

   ' // Get the unit of measure
   DIM unit AS GpUnit
   hStatus = GdipGetFontUnit(font, @unit)

   ' // Set the Graphics units of graphics to the retrieved unit value.
   hStatus = GdipSetPageUnit(graphics, unit)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "Here is some text"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLogFontA

Uses a **LOGFONTA** structure to get the attributes of the **Font** object.

```
FUNCTION GdipGetLogFontA (BYVAL font AS GpFont PTR, BYVAL graphics AS GpGraphics PTR, _
   BYVAL logFont AS LOGFONTA) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *font* | [in] Pointer to the **Font** object. |
| *graphics* | [in] Pointer to a **Graphics** object that contains attributes of the display device. |
| *logFont* | [out] Pointer to a **LOGFONTA** structure that receives the font attributes. |

#### Example

```
' ========================================================================================
' The following example creates a Font object, gets the font attributes from the Font object,
' and uses these attributes (contained in the LOGFONTA structure) to create a second Font
' object. The second Font object is then used to draw text.
' ========================================================================================
SUB Example_GetLogFontA (BYVAL hdc AS HDC)

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
   ' // Set the scale transform (don't use it in thisexample or the font will be scaled twice)
'   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   DIM font AS GpFont PTR
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, FALSE), FontStyleItalic, UnitPixel, @font)

   ' //Retrieve the LOGFONTA structure
   DIM lf AS LOGFONTA
   hStatus = GdipGetLogFontA(font, graphics, @lf)

   ' // Display some attributes (for example, in console)
   ' PRINT "LOGFONTA FaceName: "; lf.lfFaceName
   ' PRINT "Height: "; lf.lfHeight
   ' PRINT "Weight: "; lf.lfWeight
   ' PRINT "Italic: "; lf.lfItalic

   ' // Create a font from LOGFONTA
   DIM font2 AS GpFont PTR
   hStatus = GdipCreateFontFromLogfontA(hdc, @lf, @font2)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string using the second font
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "Font from LOGFONTA"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font2, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF font2 THEN GdipDeleteFont(font2)
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLogFontW

Uses a **LOGFONTW** structure to get the attributes of the **Font** object.

```
FUNCTION GdipGetLogFontW (BYVAL font AS GpFont PTR, BYVAL graphics AS GpGraphics PTR, _
   BYVAL logFont AS LOGFONTW) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *font* | [in] Pointer to the **Font** object. |
| *graphics* | [in] Pointer to a **Graphics** object that contains attributes of the display device. |
| *logFont* | [out] Pointer to a **LOGFONTW** structure that receives the font attributes. |

#### Example

```
' ========================================================================================
' The following example creates a Font object, gets the font attributes from the Font object,
' and uses these attributes (contained in the LOGFONTW structure) to create a second Font
' object. The second Font object is then used to draw text.
' ========================================================================================
SUB Example_GetLogFontW (BYVAL hdc AS HDC)

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
   ' // Set the scale transform (don't use it in thisexample or the font will be scaled twice)
'   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create the font
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   DIM font AS GpFont PTR
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, FALSE), FontStyleItalic, UnitPixel, @font)

   ' //Retrieve the LOGFONTW structure
   DIM lf AS LOGFONTW
   hStatus = GdipGetLogFontW(font, graphics, @lf)

   ' // Display some attributes (for example, in console)
   ' PRINT "LOGFONTW FaceName: "; lf.lfFaceName
   ' PRINT "Height: "; lf.lfHeight
   ' PRINT "Weight: "; lf.lfWeight
   ' PRINT "Italic: "; lf.lfItalic

   ' // Create a font from LOGFONTW
   DIM font2 AS GpFont PTR
   hStatus = GdipCreateFontFromLogfontW(hdc, @lf, @font2)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush)

   ' // Draw a string using the second font
   DIM rcf AS GpRectF = (30, 30, 0, 0)
   DIM wszText AS WSTRING * 64 = "Font from LOGFONTW"
   hStatus = GdipDrawString(graphics, wszText, LEN(wszText), font2, @rcf, NULL, solidBrush)

   ' // Cleanup
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF font2 THEN GdipDeleteFont(font2)
   IF font THEN GdipDeleteFont(font)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipNewInstalledFontCollection

Creates an **InstalledFontCollection** object.

```
FUNCTION GdipNewInstalledFontCollection (BYVAL fontCollection AS GpFontCollection PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontCollection* | [out] Pointer to a **FontCollection** object variable that receives a pointer to the **InstalledFontCollection** object. |

#### Example

```
' ========================================================================================
' Gets the number of font families installed.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create an InstalledFontCollection object
DIM fontCollection AS GpFontCollection PTR
hStatus = GdipNewInstalledFontCollection(@fontCollection)

' // Get the number of font families in the collection
DIM numberFonts AS LONG
hStatus = GdipGetFontCollectionFamilyCount(fontCollection, @numberFonts)

' // Display the result (for example, in console)
PRINT "Number of installed font families: "; numberFonts

' // Buffer for font name
DIM wszName AS WSTRING * 32

' // Allocate array for font family pointers
DIM families(numberFonts - 1) AS GpFontFamily PTR
hStatus = GdipGetFontCollectionFamilyList(fontCollection, numberFonts, @families(0), @numberFonts)
    
' // Print each font family name
FOR i AS LONG = 0 TO numberFonts - 1
   hStatus = GdipGetFamilyName(families(i), @wszName, LANG_NEUTRAL)
   PRINT "Font "; i + 1; ": "; wszName
   GdipDeleteFontFamily(families(i))
NEXT

' // Cleanup
' No delete function exists for installed font collections
' //because it is a shared collection; cleanup is automatic”

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---

## GdipNewPrivateFontCollection

Creates an **PrivateFontCollection** object.

```
FUNCTION GdipNewPrivateFontCollection (BYVAL fontCollection AS GpFontCollection PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontCollection* | [out] Pointer to a **FontCollection** object variable that receives a pointer to the **PrivateFontCollection** object. |

---

## GdipPrivateAddFontFile

Adds a font file to the private font collection.

```
FUNCTION GdipPrivateAddFontFile (BYVAL fontCollection AS GpFontCollection PTR, _
   BYVAL filename AS LPCWSTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontCollection* | [in] Pointer to the font collection object. |
| *filename* | [in] Pointer to a null-terminated unicode string that specifies the name of a font file. |

---

## GdipPrivateAddMemoryFont

Adds a font that is contained in system memory to a Microsoft® Windows® GDI+ font collection.

```
FUNCTION GdipPrivateAddMemoryFont (BYVAL fontCollection AS GpFontCollection PTR, _
   BYVAL memory AS VOID PTR, BYV AL length AS INT_) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fontCollection* | [in] Pointer to the font collection object. |
| *memory* | [in] Pointer to a font that is contained in memory. |
| *length* | [in] Long integer value that specifies the number of bytes of data in the font. |

---

## AfxGdipPointsToPixels

This function can be used to make fonts DPI aware.

```
' ========================================================================================
' Convert font point sizes to pixels and scale then for dpi awareness if needed.
' IF scale world transform has been set with GdipScaleWorldTransform, it only needs to be
' scaled by multiplying the point size by 1.3333 (one point is one point is approximately
' 1.3333 pixel). If the dpi ratios have been set with GdipScaleWorldTransform, GDI+ will
' scale the font automatically. Otherwise, we will convert and scale the font size using
' the value returned by the GetDeviceCaps API function.
' ========================================================================================
PRIVATE FUNCTION AfxGdipPointsToPixels (BYVAL ptSize AS SINGLE, BYVAL fUsingWorldTransform AS BOOLEAN = FALSE) AS SINGLE
   IF fUsingWorldTransform THEN RETURN ptSize * 1.3333
   ' // Scale according the DPI settings of the computer
   DIM hDC AS HDC = GetDC(HWND_DESKTOP)
   DIM LPX AS LONG = GetDeviceCaps(hDC, LOGPIXELSX)
   ReleaseDC(HWND_DESKTOP, hDC)
   RETURN MulDiv(ptSize, LPX, 72)
END FUNCTION
' ========================================================================================
```

#### Usage example

```
' // Create a font family
DIM fontFamily AS GpFontFamily PTR
hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
' // Create a font from the font family
DIM font AS GpFont PTR
hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(18, TRUE), FontStyleRegular, UnitPixel, @font)
```
---
