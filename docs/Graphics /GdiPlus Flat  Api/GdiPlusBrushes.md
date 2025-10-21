# Brush Functions

A **Brush** object is used to paint the interior of graphics shapes, such as rectangles, ellipses, pies, polygons, and paths.

### Generic Brush Functions

| Name       | Description |
| ---------- | ----------- |
| [GdipCloneBrush](#gdipclonebrush) | Creates a new **Brush** object based on this brush. |
| [GdipDeleteBrush](#gdipdeletebrush) | Deletes the specified **Brush** object. |
| [GdipGetBrushType](#gdipgetbrushtype) | Gets the type of this brush. |

### SolidBrush Functions

| Name       | Description |
| ---------- | ----------- |
| [GdipCreateSolidFill](#gdipcreatesolidfill) | Creates a **SolidBrush** object based on a color. |
| [GdipGetSolidFillColor](#gdipgetsolidfillcolor) | Gets the color of this **SolidBrush** object. |
| [GdipSetSolidFillColor](#gdipsetsolidfillcolor) | Sets the color of this **SolidBrush** object. |

### HatchBrush Functions

An **HatchBrush** defines a rectangular brush with a hatch style, a foreground color, and a background color. There are six hatch styles. The foreground color defines the color of the hatch lines; the background color defines the color over which the hatch lines are drawn.


| Name       | Description |
| ---------- | ----------- |
| [GdipCreateHatchBrush](#gdipcreatehatchbrush) | Creates a **HatchBrush** object based on a hatch style, a foreground color, and a background color. |
| [GdipGetHatchBackgroundColor](#gdipgethatchbackgroundcolor) | Gets the background color of this hatch brush. |
| [GdipGetHatchForegroundColor](#gdipgethatchforegroundcolor) | Gets the foreground color of this hatch brush. |
| [GdipGetHatchStyle](#gdipgethatchstyle) | Gets the hatch style of this hatch brush. |

---

## GdipCloneBrush

Creates a new **Brush** object based on this brush. The cloned brush is a deep copy—you can modify it without affecting the original.

```
FUNCTION GdipCloneBrush (BYVAL brush AS GpBrush PTR, BYVAL cloneBrush AS GpBrush PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **Brush** object created with **GdipCreateSolidFil**, **GdipCreateHatchBrush**, **GdipCloneBrush** or similar function. |
| *cloneBrush* | [out] Pointer to a **Brush** object variable that recives a pointer to the new **Brush** object. |

#### Example

The following example creates a SolidBrush object, clones it, and then uses the clone to fill a rectangle.

```
' ========================================================================================
SUB Example_CloneBrush (BYVAL hdc AS HDC)

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

   ' // Create an ARGB color
   DIM pixelColor AS ARGB
   pixelColor = GDIP_ARGB(255, 255, 0, 0)

   ' // Create a SolidBrush
   DIM brush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(pixelColor, @brush)

   ' // Create a clone of the SolidBrush
   DIM cloneBrush AS GpBrush PTR
   hStatus = GdipCloneBrush(brush, @cloneBrush)

   ' // Use the clone brush to fill a rectagle
   hStatus = GdipFillRectangleI(graphics, cloneBrush, 0, 0, 100, 100)

   ' // Cleanup
   IF cloneBrush THEN GdipDeleteBrush(cloneBrush)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeleteBrush

Deletes the specified **Brush** object.

```
FUNCTION GdipDeleteBrush (BYVAL brush AS GpBrush PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **Brush** object created with **GdipCreateSolidFil**, **GdipCreateHatchBrush**, **GdipCloneBrush** or similar function. |

#### Example

```
DIM hStatus AS LONG

' // Create a graphics object from the device context
DIM graphics AS GpGraphics PTR
hStatus = GdipCreateFromHDC(hdc, @graphics)

' // Create a SolidBrush
DIM brush AS GpSolidFill PTR
hStatus = GdipCreateSolidFill(ARGB_BLUE, @brush)

' ...
' ...
' ...

' // Cleanup
IF brush THEN GdipDeleteBrush(brush)
IF graphics THEN GdipDeleteGraphics(graphics)
```
---

## GdipGetBrushType

Gets the type of this brush.

```
FUNCTION GdipGetBrushType (BYVAL brush AS GpBrush PTR, BYVAL brushType AS GpBrushType PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **Brush** object created with **GdipCreateSolidFil**, **GdipCreateHatchBrush**, **GdipCloneBrush** or similar function. |
| *brushType* | [out] Pointer to a *GpBrushType* variable that receives the type of this brush. The value returned is one of the elements of the *BrushType* enumeration. |

**BrushType enumeration**

| Constant  | Value | Description |
| --------- | ----- | ----------- |
| **BrushTypeSolidColor** | 0 | Indicates a brush of type **SolidBrush**. A solid brush paints a single, constant color that can be opaque or transparent. |
| **BrushTypeHatchFill** | 1 | Indicates a brush of type **HatchBrush**. A hatch brush paints a background and paints, over that ackground, a pattern of lines, dots, dashes, squares, crosshatch, or some variation of these. The hatch brush consists of two olors: one for the background and one for the pattern over the background. The color of the background is called the background olor, and the color of the pattern is called the foreground color. |
| **BrushTypeTextureFill** | 2 | Indicates a brush of type **TextureBrush**. A texture brush paints an image. The image or texture is either a portion of a specified image or a scaled version of a specified image. The type of image (metafile or nonmetafile) determines whether the texture is a ortion of the image or a scaled version of the image. |
| **BrushTypePathGradient** | 3 | Indicates a brush of type **PathGradientBrush**. A path gradient brush paints a color gradient in which the color changes from a enter point outward to a boundary that is defined by a closed curve or path. The color gradient has one color at the center point and ne or multiple colors at the boundary. |
| **BrushTypeLinearGradient** | 4 | Indicates a brush of type **LinearGradientBrush**. A linear gradient brush paints a color gradient in which the color changes evenly rom the starting boundary line of the linear gradient brush to the ending boundary line of the linear gradient brush. The boundary ines of a linear gradient brush are two parallel straight lines. The color gradient is perpendicular to the boundary lines of the inear gradient brush, changing gradually across the stroke from the starting boundary line to the ending boundary line. The color radient has one color at the starting boundary line and another color at the ending boundary line. |

#### Example

The following example creates a SolidBrush, checks the type of the brush, and then, if the type is BrushTypeSolidColor, uses the brush to fill a rectangle.

```
' ========================================================================================
' The following example creates a SolidBrush, checks the type of the brush, and
' then, if the type is BrushTypeSolidColor, uses the brush to fill a rectangle.
' ========================================================================================
SUB Example_GetBrushType (BYVAL hdc AS HDC)

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

   ' // Create a SolidBrush
   DIM brush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @brush)

   ' // Get the type of solidBrush
   DIM nType AS BrushType
   hStatus = GdipGetBrushType(brush, @nType)

   ' // If the type of solidBrush is BrushTypeSolidColor, use it to fill a rectangle
   IF nType = BrushTypeSolidColor THEN
      hStatus = GdipFillRectangle(graphics, brush, 0, 0, 100, 100)
   END IF

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateSolidFill

```
FUNCTION GdipCreateSolidFill (BYVAL color AS ARGB, BYVAL brush AS GpSolidFill PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *color* | [in] ARGB color that specifies the initial color of the solid brush. |
| *brush* | [out] Pointer to a variable that receives a pointer to the new created **SolidBrush** object. |

#### Example

```
' ========================================================================================
' The following example creates a SolidBrush, checks the type of the brush, and
' then, if the type is BrushTypeSolidColor, uses the brush to fill a rectangle.
' ========================================================================================
SUB Example_CreateSolidFill (BYVAL hdc AS HDC)

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

   ' // Create a SolidBrush
   DIM brush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @brush)

   ' // Get the type of solidBrush
   DIM nType AS BrushType
   hStatus = GdipGetBrushType(brush, @nType)

   ' // If the type of solidBrush is BrushTypeSolidColor, use it to fill a rectangle
   IF nType = BrushTypeSolidColor THEN
      hStatus = GdipFillRectangle(graphics, brush, 0, 0, 100, 100)
   END IF

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetSolidFillColor

Gets the color of the **SolidBrush** object.

```
FUNCTION GdipGetSolidFillColor (BYVAL brush AS GpSolidFill PTR, BYVAL color AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **SolidBrush** object. |
| *color* | [out] Pointer to a variable that receives the color of the solid brush. |

#### Example

```
' ========================================================================================
' The following example creates a solid brush and uses it to fill a rectangle. The code
' gets the color of the solid brush and stores it. Then, the code creates a second solid
' brush using the stored color and paints a second rectangle with the second solid brush.
' ========================================================================================
SUB Example_GetSolidFillColor (BYVAL hdc AS HDC)

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

   ' // Create a SolidBrush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush )

   ' // Fill a rectangle with the brush
   GdipFillRectangle(graphics, solidBrush, 10, 10, 185, 100)

   ' // Get the color of the solid brush.
   DIM colour AS ARGB
   GdipGetSolidFillColor(solidBrush, @colour)

   ' // Create a second solid brush with that same color.
   DIM solidBrush2 AS GpSolidFill PTR
   GdipCreateSolidFill(colour, @solidBrush2)

   ' // Paint a second rectangle with the second solid brush.
   GdipFillRectangle(graphics, solidBrush2, 205, 10, 185, 100)

   ' // Cleanup
   IF solidBrush2 THEN GdipDeleteBrush(solidBrush2)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetSolidFillColor

Sets the color of the **SolidBrush** object.

```
FUNCTION GdipSetSolidFillColor (BYVAL brush AS GpSolidFill PTR, BYVAL color AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **SolidBrush** object. |
| *color* | [in] ARGB color that specifies the color to be set in this solid brush. |

#### Example

```
' ========================================================================================
' The following example creates a solid brush and uses it to fill a rectangle. The code
' uses GdipSetSolidFillColor to change the color of the solid brush and then paints a
' second rectangle the new color.
' ========================================================================================
SUB Example_SetSolidFillColor (BYVAL hdc AS HDC)

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

   ' // Create a SolidBrush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @solidBrush )

   ' // Fill a rectangle with the brush
   GdipFillRectangle(graphics, solidBrush, 10, 10, 185, 100)

   ' // Change the color of the brush, and fill another rectangle.
   GdipSetSolidFillColor(solidBrush, ARGB_RED)
   GdipFillRectangle(graphics, solidBrush, 205, 10, 185, 100)

   ' // Cleanup
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateHatchBrush

Creates a **HatchBrush** object based on a hatch style, a foreground color, and a background color.

```
FUNCTION GdipCreateHatchBrush (BYVAL hatchstyle AS GpHatchStyle PTR, BYVAL forecol AS ARGB, _
   BYVAL backcol AS ARGB, BYVAL brush AS GpHatch) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hatchstyle* | [in] Element of the **HatchStyle** enumeration that specifies the pattern of hatch lines that will be used. |
| *forecol* | [in] ARGB color to use for the hatch lines. |
| *backcol* | [in]  ARGB color to use for the background. |
| *brush* | [in] Pointer to a variable that receives a pointer to the new created **HatchBrush**. |

#### Example

```
' ========================================================================================
' The following example draws six of the available hatch styles.
' ========================================================================================
SUB Example_CreateHatchBrush (BYVAL hdc AS HDC)

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

   ' // Set and then draw the first hatch style.
   DIM brush AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleHorizontal, ARGB_BLACK, ARGB_WHITE, @brush)
   hStatus = GdipFillRectangle(graphics, brush, 20, 20, 115, 50)

   ' // Set and then draw the second hatch style.
   DIM brush1 AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleVertical, ARGB_BLACK, ARGB_WHITE, @brush1)
   hStatus = GdipFillRectangle(graphics, brush1, 145, 20, 115, 50)

   ' // Set and then draw the third hatch style.
   DIM brush2 AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleForwardDiagonal, ARGB_BLACK, ARGB_WHITE, @brush2)
   hStatus = GdipFillRectangle(graphics, brush2, 270, 20, 115, 50)

   ' // Set and then draw the fourth hatch style.
   DIM brush3 AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleBackwardDiagonal, ARGB_BLACK, ARGB_WHITE, @brush3)
   hStatus = GdipFillRectangle(graphics, brush3, 20, 100, 115, 50)

   ' // Set and then draw the fifth hatch style.
   DIM brush4 AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleCross, ARGB_BLACK, ARGB_WHITE, @brush4)
   hStatus = GdipFillRectangle(graphics, brush4, 145, 100, 115, 50)

   ' // Set and then draw the sixth hatch style.
   DIM brush5 AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleDiagonalCross, ARGB_BLACK, ARGB_WHITE, @brush5)
   hStatus = GdipFillRectangle(graphics, brush5, 270, 100, 115, 50)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF brush1 THEN GdipDeleteBrush(brush1)
   IF brush2 THEN GdipDeleteBrush(brush2)
   IF brush3 THEN GdipDeleteBrush(brush3)
   IF brush4 THEN GdipDeleteBrush(brush4)
   IF brush5 THEN GdipDeleteBrush(brush5)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetHatchBackgroundColor

Gets the background color of the hatch brush.

```
FUNCTION GdipGetHatchStyle (BYVAL brush AS GpHatch PTR, BYVAL hatchstyle AS GpHatchStyle PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **HatchBrush** object. |
| *hatchstyle* | [out] Pointer to a variable that receives the background color. The background color defines the color over which the hatch lines are drawn. |

#### Example

```
' ========================================================================================
' The following example sets up three colors: black, turquoise, and current (initialized
' to black). A rectangle is painted by using turquoise as the background color and black
' as the foreground color. Then the GdipGetBackgroundColor function is used to get the
' current color of the brush (which at the time is turquoise). The address of the current
' color (initialized to black) is passed as the return point for the call to
' GdipGetBackgroundColor. When the rectangle is painted again, note that the background
' color is again turquoise (not black). This shows that the call to GdipGetBackgroundColor
' was successful.
' ========================================================================================
SUB Example_GetHatchBackgroundColor (BYVAL hdc AS HDC)

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

   ' // Set and then draw the first hatch style.
   DIM brush AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleHorizontal, ARGB_BLACK, ARGB_Turquoise, @brush)
   hStatus = GdipFillRectangle(graphics, brush, 20, 20, 170, 100)

   ' // Get the background color of the current brush.
   DIM colorCurrent AS ARGB
   hStatus = GdipGetHatchBackgroundColor(brush, @colorCurrent)

   ' // Draw the rectangle again using the current color.
   DIM brush2 AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleDiagonalCross, ARGB_BLACK, colorCurrent, @brush2)
   hStatus = GdipFillRectangle(graphics, brush2, 210, 20, 170, 100)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF brush2 THEN GdipDeleteBrush(brush2)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetHatchForegroundColor

Gets the foreground color of the hatch brush.

```
FUNCTION GdipGetHatchForegroundColor (BYVAL brush AS GpHatch PTR, BYVAL forecol AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **HatchBrush** object. |
| *forecol* | [out] Pointer to a variable that receives the foreground color. The foreground color defines the color of the hatch lines. |

#### Example

```
' ========================================================================================
' The following example sets up three colors: blue, turquoise, and current (initialized to
' black). A rectangle is painted by using turquoise as the background color and blue as the
' foreground color. Then the GdipGetBackgroundColor function is used to get the current
' color of the brush (which at the time is blue). The address of the current Color object
' (initialized to black) is passed as the return point for the call to GdipGetBackgroundColor.
' When the rectangle is painted again, notice that the foreground color is again blue (not
' black). This shows that the call to GdipGetForegroundColor was successful.
' ========================================================================================
SUB Example_GetHatchForegroundColor (BYVAL hdc AS HDC)

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

   ' // Set and then draw the first hatch style.
   DIM brush AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleHorizontal, ARGB_BLUE, ARGB_Turquoise, @brush)
   hStatus = GdipFillRectangle(graphics, brush, 20, 20, 170, 100)

   ' // Get the foreground color of the current brush.
   DIM colorCurrent AS ARGB
   hStatus = GdipGetHatchForegroundColor(brush, @colorCurrent)

   ' // Draw the rectangle again using the current color.
   DIM brush2 AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleDiagonalCross, colorCurrent, ARGB_Turquoise, @brush2)
   hStatus = GdipFillRectangle(graphics, brush2, 210, 20, 170, 100)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF brush2 THEN GdipDeleteBrush(brush2)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetHatchStyle

Gets the hatch style of the hatch brush.

```
FUNCTION GdipGetHatchStyle (BYVAL brush AS GpHatch PTR, BYVAL hatchstyle AS GpHatchStyle PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **HatchBrush** object. |
| *hatchstyle* | [out] Pointer to a variable that receives the hatch style, which is one of the elements of the **HatchStyle** enumeration. |

#### Example

```
' ========================================================================================
' The following example sets up three colors: blue, turquoise, and current (initialized to
' black). A rectangle is painted by using turquoise as the background color and blue as the
' foreground color. Then the GdipGetBackgroundColor function is used to get the current
' color of the brush (which at the time is blue). The address of the current Color object
' (initialized to black) is passed as the return point for the call to GdipGetBackgroundColor.
' When the rectangle is painted again, notice that the foreground color is again blue (not
' black). This shows that the call to GdipGetForegroundColor was successful.
' ========================================================================================
SUB Example_GetHatchForegroundColor (BYVAL hdc AS HDC)

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

   ' // Set and then draw the first hatch style.
   DIM brush AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleHorizontal, ARGB_BLUE, ARGB_Turquoise, @brush)
   hStatus = GdipFillRectangle(graphics, brush, 20, 20, 170, 100)

   ' // Get the foreground color of the current brush.
   DIM colorCurrent AS ARGB = HatchStyleDiagonalCross
   hStatus = GdipGetHatchForegroundColor(brush, @colorCurrent)

   ' // Draw the rectangle again using the current color.
   DIM brush2 AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleDiagonalCross, colorCurrent, ARGB_Turquoise, @brush2)
   hStatus = GdipFillRectangle(graphics, brush2, 210, 20, 170, 100)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF brush2 THEN GdipDeleteBrush(brush2)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
