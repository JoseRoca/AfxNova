## Pen Functions

A **Pen** object is a Microsoft Windows GDI+ object used to render lines, curves, and outlines with customizable styles and transformations. Whether you're creating simple strokes or complex compound paths, the Pen API offers fine-grained control over appearance, alignment, and behavior.

**What you can do with a pen**:

* Define stroke color, width, and unit of measure
* Apply brush fills for textured or gradient strokes
* Customize caps, joins, and dash patterns
* Transform pen geometry with scaling, rotation, and translation
* Use compound arrays for multi-line effects
* Inspect and modify pen attributes dynamically

Each function listed below includes a clear description, parameter breakdown, and practical FreeBasic example to help you integrate GDI+ pens into your graphics workflow with confidence.

---

| Name       | Description |
| ---------- | ----------- |
| [GdipClonePen](#gdipclonepen) | Copies a **Pen** object. |
| [GdipCreatePen1](#gdipcreatepen1) | Creates a **Pen** object that uses a specified color and width. |
| [GdipCreatePen2](#gdipcreatepen2) | Creates a **Pen** object that uses the attributes of a brush and a real number to set the width of the **Pen** object. |
| [GdipDeletePen](#gdipdeletepen) | Deletes a **Pen** object. |
| [GdipGetPenBrushFill](#gdipgetpenbrushfill) | Gets the **Brush** object that is currently set for a **Pen** object. |
| [GdipGetPenColor](#gdipgetpencolor) | Gets the color currently set for a **Pen** object. |
| [GdipGetPenCompoundArray](#gdipgetpencompoundarray) | Gets the compound array currently set for a **Pen** object. |
| [GdipGetPenCompoundCount](#gdipgetpencompoundcount) | Gets the number of elements in a compound array. |
| [GdipGetPenCustomEndCap](#gdipgetpencustomendcap) | Gets the custom end cap currently set for a **Pen** object. |
| [GdipGetPenCustomStartCap](#gdipgetpencustomstartcap) | Gets the custom start cap currently set for a **Pen** object. |
| [GdipGetPenDashArray](#gdipgetpendasharray) | Gets an array of custom dashes and spaces currently set for a **Pen** object. |
| [GdipGetPenDashCap197819](#GdipGetPenDashCap197819) | Gets the dash cap style currently set for a **Pen** object. |
| [GdipGetPenDashCount](#gdipgetpendashcount) | Gets the number of elements in a dash pattern array. |
| [GdipGetPenDashOffset](#gdipgetpendashoffset) | Gets the distance from the start of the line to the start of the first space in a dashed line. |
| [GdipGetPenDashOffset](#gdipgetpendashoffset) | Gets the distance from the start of the line to the start of the first space in a dashed line. |
| [GdipGetPenDashStyle](#gdipgetpendashstyle) | Gets the dash style currently set for a **Pen** object. |
| [GdipGetPenEndCap](#gdipgetpenendcap) | Gets the end cap currently set for a **Pen** object. |
| [GdipGetPenFillType](#gdipgetpenfilltype) | Gets the type currently set for a **Pen** object. |
| [GdipGetPenLineJoin](#gdipgetpenlinejoin) | Gets the line join style currently set for a **Pen** object. |
| [GdipGetPenMiterLimit](#gdipgetpenmiterlimit) | Gets the miter length currently set for a **Pen** object. |
| [GdipGetPenMode](#gdipgetpenmode) | Gets the alignment currently set for a **Pen** object. |
| [GdipGetPenStartCap](#gdipgetpenstartcap) | Gets the start cap currently set for a **Pen** object. |
| [GdipGetPenTransform](#gdipgetpentransform) | Gets the world transformation matrix currently set for a **Pen** object. |
| [GdipGetPenUnit](#gdipgetpenunit) | Gets the world transformation matrix currently set for a **Pen** object. |
| [GdipGetPenWidth](#gdipgetpenwidth) | Gets the width currently set for a **Pen** object. |
| [GdipMultiplyPenTransform](#gdipmultiplypentransform) | Updates the world transformation matrix of a **Pen** object with the product of itself and another matrix. |
| [GdipResetPenTransform](#gdipresetpentransform) | Sets the world transformation matrix of a **Pen** object to the identity matrix. |
| [GdipRotatePenTransform](#gdiprotatepentransform) | Updates the world transformation matrix of a **Pen** object with the product of itself and a rotation matrix. |
| [GdipScalePenTransform](#gdipscalepentransform) | Sets the **Pen** object's world transformation matrix equal to the product of itself and a scaling matrix. |
| [GdipSetPenBrushFill](#gdipsetpenbrushfill) | Sets the **Brush** object that a pen uses to fill a line. |
| [GdipSetPenColor](#gdipsetpencolor) | Sets the color for a **Pen** object. |
| [GdipSetPenCompoundArray](#gdipsetpencompoundarray) | Sets the compound array for a **Pen** object. |
| [GdipSetPenCustomEndCap](#gdipsetpencustomendcap) | Sets the custom end cap for a **Pen** object. |
| [GdipSetPenCustomStartCap](#gdipsetpencustomstartcap) | Sets the custom start cap for a **Pen** object. |
| [GdipSetPenDashArray](#gdipsetpendasharray) | Sets an array of custom dashes and spaces for a **Pen** object. |
| [GdipSetPenDashCap197819](#gdipsetpendashcap197819) | Sets the dash cap style for a **Pen** object. |
| [GdipSetPenDashOffset](#gdipsetpendashoffset) | Sets the distance from the start of the line to the start of the first dash in a dashed line. |
| [GdipSetPenDashStyle](#gdipsetpendashstyle) | Sets the dash style for a **Pen** object. |
| [GdipSetPenEndCap](#gdipsetpenendcap) | Sets the end cap for a **Pen** object. |
| [GdipSetPenLineCap197819](#gdipsetpenlinecap197819) | Sets the cap styles for the start, end, and dashes in a line drawn with a pen. |
| [GdipSetPenLineJoin](#gdipsetpenlinejoin) | Sets the line join for a **Pen** object. |
| [GdipSetPenMiterLimit](#gdipsetpenmiterlimit) | Sets the miter limit of a **Pen** object. |
| [GdipSetPenMode](#gdipsetpenmode) | Sets the alignment for a **Pen** object relative to the line. |
| [GdipSetPenStartCap](#gdipsetpenstartcap) | Sets the start cap for a **Pen** object. |
| [GdipSetPenTransform](#gdipsetpentransform) | Sets the world transformation of a **Pen** object. |
| [GdipSetPenUnit](#gdipsetpenunit) | Sets the unit of measure for the pen specified by the pen parameter to the value specified by the unit parameter. |
| [GdipSetPenWidth](#gdipsetpenwidth) | Sets the width for a **Pen** object. |
| [GdipTranslatePenTransform](#gdiptranslatepentransform) | Translates the local geometric transformation by the specified dimensions. This method prepends the translation to the transformation. |

---

## GdipClonePen

Copies a Pen object.

```
FUNCTION GdipClonePen (BYVAL pen AS GpPen PTR, BYVAL clonepen AS GpPen PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *clonepen* | [out] Pointer to a variable that receives a pointer to the new **Pen** object. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object, creates a copy of the Pen object, and then
' draws an ellipse using the cloned Pen object.
' ========================================================================================
SUB Example_ClonePen (BYVAL hdc AS HDC)

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

   ' // Create a Pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 4, UnitWorld, @pen)
   ' // Clone it
   DIM clonedPen AS GpPen PTR
   hStatus = GdipClonePen(pen, @clonedPen)

   ' // Draw a rectangle using the cloned Pen object.
   hStatus = GdipDrawRectangleI(graphics, clonedPen, 10, 10, 100, 50)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF clonedPen THEN GdipDeletePen(clonedPen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreatePen1

Creates a **Pen** object that uses a specified color and width.

```
FUNCTION GdipCreatePen1 (BYVAL color AS ARGB, BYVAL width AS REAL, BYVAL unit AS GpUnit, _
   BYVAL pen AS GpPen PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *color* | [in] ARGB color that specifies the color for a **Pen** object. |
| *width* | [in] Single precision value that specifies the width of a pen's stroke. The default value is 1.0. |
| *unit* | [in] Unit measure. The default unit of measure is UnitPixel, which is an element of the**GpUnit** enumeration. |
| *pen* | [out] Pointer to a variable that receives a pointer to the new created **Pen** object. |

#### Example

```
' ========================================================================================
' This example creates a Pen object using GdipCreatePen1, draws a line, and demonstrates
' how to configure its color, width, and unit.
' ========================================================================================
SUB Example_CreatePen1 (BYVAL hdc AS HDC)

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

   ' // Create a Pen with solid color, width, and unit
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 8, UnitWorld, @pen)

   ' // Draw a line with the pen
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 300, 50)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreatePen2

Creates a **Pen** object that uses the attributes of a brush and a real number to set the width of a **Pen** object.

```
FUNCTION GdipCreatePen2 (BYVAL brush AS GpBrush PTR, BYVAL width AS REAL, BYVAL unit AS GpUnit, _
   BYVAL pen AS GpPen PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to a brush to base a pen on. |
| *width* |[in] Single precision value that specifies the width of a pen's stroke. The default value is 1.0. |
| *unit* | [in] Unit measure. The default unit of measure is UnitPixel, which is an element of the**GpUnit** enumeration. |
| *pen* | [out] Pointer to a variable that receives a pointer to the new created **Pen** object. |

#### Example

```
' ========================================================================================
' This example creates a HatchBrush, uses it to create a Pen with GdipCreatePen2,
' draws a line with that pen, and demonstrates how brush-based pens work.
' ========================================================================================
SUB Example_CreatePen2 (BYVAL hdc AS HDC)

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

   ' // Create a HatchBrush
   DIM hatchBrush AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleDiagonalCross, ARGB_RED, ARGB_YELLOW, @hatchBrush)

   ' // Create a Pen using the HatchBrush
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen2(hatchBrush, 12, UnitWorld, @pen)

   ' // Draw a line with the brush-based pen
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 350, 50)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF hatchBrush THEN GdipDeleteBrush(hatchBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeletePen

Deletes a **Pen** object.

```
FUNCTION GdipDeletePen (BYVAL pen AS GpPen PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |

#### Example

```
' ========================================================================================
' This example creates a Pen, uses it to draw a line, and then deletes it using GdipDeletePen.
' ========================================================================================
SUB Example_DeletePen (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' Create a graphics object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' Create a Pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 5, UnitWorld, @pen)

   ' Draw a line
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 300, 50)

   ' Delete the Pen to free memory
   IF pen THEN GdipDeletePen(pen)

   ' Clean up graphics object
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenBrushFill

Gets the **Brush** object that is currently set for a **Pen** object.

```
FUNCTION GdipGetPenBrushFill (BYVAL pen AS GpPen PTR, BYVAL brush AS GpBrush PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *brush* | [out] Pointer to a variable that receives a pointer to the **Brush** object that is currently used to fill a line. |

#### Example

```
' ========================================================================================
' The following example creates a HachBrush object, creates a Pen object based on the
' HachBrush object, draws a line with the pen, gets the Brush from the pen, and then uses
' the Brush to fill a rectangle.
' ========================================================================================
SUB Example_GetPenBrushFill (BYVAL hdc AS HDC)

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

   ' // Create a HatchBrush object.
   DIM hatchBrush AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleHorizontal, ARGB_RED, ARGB_BLUE, @hatchBrush)

   ' // Create a pen based on a hatch brush, and use that pen to draw a line.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen2(hatchBrush, 15, UnitWorld, @pen)
   hStatus = GdipDrawLine(graphics, pen, 0, 0, 200, 100)

   ' // Get the pen's brush, and use that brush to fill a rectangle.
   DIM brush AS GpBrush PTR
   hStatus = GdipGetPenBrushFill(pen, @brush)
   hStatus = GdipFillRectangleI(graphics, brush, 0, 100, 200, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF brush THEN GdipDeleteBrush(brush)
   IF hatchBrush THEN GdipDeleteBrush(hatchBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenColor

Gets the color currently set for a **Pen** object.

```
FUNCTION GdipGetPenColor (BYVAL pen AS GpPen PTR, BYVAL argb AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *argb* | [out] Pointer to a variable that receives the color of a **Pen** object. |

#### Example

```
' ========================================================================================
' The following example creates a HachBrush object, creates a Pen object based on the
' HachBrush object, draws a line with the pen, gets the Brush from the pen, and then uses
' the Brush to fill a rectangle.
' ========================================================================================
SUB Example_GetPenColor (BYVAL hdc AS HDC)

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

   ' // Create a pen, and use it to draw a line.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(GDIP_ARGB(255, 200, 150, 100), 5, UnitWorld, @pen)
   hStatus = GdipDrawLine(graphics, pen, 0, 0, 200, 100)

   ' // Get the pen's color, and use that color to create a brush.
   DIM argbColor AS ARGB
   hStatus = GdipGetPenColor(pen, @argbColor)
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(argbColor, @solidBrush)

   ' // Use the brush to fill a rectangle.
   hStatus = GdipFillRectangleI(graphics, solidBrush, 0, 100, 200, 100)

   ' // Cleanup
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenCompoundArray

Gets the compound array currently set for a **Pen** object.

```
FUNCTION GdipGetPenCompoundArray (BYVAL pen AS GpPen PTR, BYVAL dash AS REAL PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *dash* | [out] Pointer to an array that receives the compound array. |
| *count* | [in] Long integer value that specifies the number of elements in the *dash* array. |

#### Remarks

Suppose that a compound array contains the values 0.0, 0.2, 0.7, and 1.0 and that the pen has a width of 100. When you use the pen to draw, you get two parallel lines. The first line has a width of 20, the space between the two lines has a width of 50, and the second line has a width of 30.

For a more complex example, suppose that a compound array contains the values 0.0, 0.2, 0.3, 0.6, 0.85, and 1.0 and that the pen has a width of 100. When you use the pen to draw, you get three parallel lines. The widths of the three lines are 20, 30, and 15 respectively. The widths of the two spaces between the lines are 10 and 15 respectively.

#### Example

```
' ========================================================================================
' The following example declares an array, sets the compound array, draws a line, and gets
' the elements in the compound array.
' ========================================================================================
SUB Example_GetPenCompoundArray (BYVAL hdc AS HDC)

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

   ' // Create an array of real numbers and a Pen object.
   DIM compVals(5) AS SINGLE = {0.0, 0.2, 0.5, 0.7, 0.9, 1.0}
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 30, UnitWorld, @pen)

   ' // Set the compound array of the pen.
   hStatus = GdipSetPenCompoundArray(pen, @compVals(0), 6)

   ' // Draw a line with the pen.
   hStatus = GdipDrawLine(graphics, pen, 8, 20, 390, 200)

   ' // Obtain information about the pen.
   DIM count AS LONG
   DIM compValues(ANY) AS SINGLE
   hStatus = GdipGetPenCompoundCount(pen, @count)
   REDIM compValues(count - 1) AS SINGLE
   hStatus = GdipGetPenCompoundArray(pen, @compValues(0), count)

   FOR j AS LONG = 0 TO count - 1
      ' // Inspect or use the value in compValues(j).
      OutputDebugStringW(WSTR(compValues(j)))
   NEXT

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenCompoundCount

Gets the number of elements in a compound array.

```
FUNCTION GdipGetPenCompoundCount (BYVAL pen AS GpPen PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *count* | [out] Pointer to a variable that receives the number of elements in the compound array. |

#### Remarks

Suppose that a compound array contains the values 0.0, 0.2, 0.7, and 1.0 and that the pen has a width of 100. When you use the pen to draw, you get two parallel lines. The first line has a width of 20, the space between the two lines has a width of 50, and the second line has a width of 30.

For a more complex example, suppose that a compound array contains the values 0.0, 0.2, 0.3, 0.6, 0.85, and 1.0 and that the pen has a width of 100. When you use the pen to draw, you get three parallel lines. The widths of the three lines are 20, 30, and 15 respectively. The widths of the two spaces between the lines are 10 and 15 respectively.

#### Example

See example in **GdipGetPenCompoundArray**.

---

## GdipGetPenCustomEndCap

Gets the custom end cap currently set for a **Pen** object.

```
FUNCTION GdipGetPenCustomEndCap (BYVAL pen AS GpPen PTR, BYVAL customCap AS GpCustomLineCap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *customCap* | [out] Pointer to a variable that receives the custom end cap of a **Pen** object. |

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom end cap using the GraphicsPath object, and
' draws a line. Finally, the code gets the custom end cap of the pen and creates another
' Pen object using the same custom end cap. It then draws a second line.
' ========================================================================================
SUB Example_GetPenCustomEndCap (BYVAL hdc AS HDC)

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

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @pStrokePath)
   hStatus = GdipAddPathRectangleI(pStrokePath, -10, -5, 20, 10)

  ' // Create a pen, and set the custom end cap based on the GraphicsPath object.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1, UnitWorld, @pen)
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, pStrokePath, LineCapFlat, 0, @custCap)
   hStatus = GdipSetPenCustomEndCap(pen, custCap)

   ' // Draw a line with the custom end cap.
   hStatus = GdipDrawLine(graphics, pen, 0, 0, 200, 100)

   ' // Obtain the custom end cap for the pen.
   DIM customLineCap AS GpCustomLineCap PTR
   hStatus = GdipGetPenCustomEndCap(pen, @customLineCap)

   ' // Create another pen, and use the same custom end cap.
   DIM pen2 AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 3, UnitWorld, @pen2)
   hStatus = GdipSetPenCustomEndCap(pen2, customLineCap)

   ' // Draw a line using the second pen.
   hStatus = GdipDrawLine(graphics, pen2, 0, 100, 200, 150)

   ' // Cleanup
   IF pen2 THEN GdipDeletePen(pen2)
   IF pen THEN GdipDeletePen(pen)
   IF customLineCap THEN GdipDeleteCustomLineCap(customLineCap)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF pStrokePath THEN GdipDeletePath(pStrokePath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenCustomStartCap

Gets the custom start cap currently set for a **Pen** object.

```
FUNCTION GdipGetPenCustomStartCap (BYVAL pen AS GpPen PTR, BYVAL customCap AS GpCustomLineCap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *customCap* | [out] Pointer to a variable that receives a pointer to the start cap of a **Pen** object. |

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom end cap using the GraphicsPath object, and
' draws a line. Finally, the code gets the custom end cap of the pen and creates another
' Pen object using the same custom end cap. It then draws a second line.
' ========================================================================================
SUB Example_GetPenCustomStartCap (BYVAL hdc AS HDC)

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

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @pStrokePath)
   hStatus = GdipAddPathRectangleI(pStrokePath, -10, -5, 20, 10)

   ' // Create a pen, and set the custom start cap based on the GraphicsPath object.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(GDIP_ARGB(255, 0, 0, 255), 1, UnitWorld, @pen)
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, pStrokePath, LineCapFlat, 0, @custCap)
   hStatus = GdipSetPenCustomStartCap(pen, custCap)

   ' // Draw a line with the custom start cap.
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 200, 100)

   ' // Obtain the custom start cap for the pen.
   DIM customLineCap AS GpCustomLineCap PTR
   hStatus = GdipGetPenCustomStartCap(pen, @customLineCap)

   ' // Create another pen, and use the same custom start cap.
   DIM pen2 AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 3, UnitWorld, @pen2)
   hStatus = GdipSetPenCustomStartCap(pen2, customLineCap)

   ' // Draw a line using the second pen.
   hStatus = GdipDrawLine(graphics, pen2, 50, 100, 200, 150)

   ' // Cleanup
   IF pen2 THEN GdipDeletePen(pen2)
   IF pen THEN GdipDeletePen(pen)
   IF customLineCap THEN GdipDeleteCustomLineCap(customLineCap)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF pStrokePath THEN GdipDeletePath(pStrokePath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenDashArray

Gets an array of custom dashes and spaces currently set for a **Pen** object.

```
FUNCTION GdipGetPenDashArray (BYVAL pen AS GpPen PTR, BYVAL dash AS REAL PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *dash* | [out] Pointer to an array that receives the length of the dashes and spaces in a custom dashed line. |
| *count* | [in] Long integer value that specifies the number of elements in the *dash* array. |

#### Remarks

The elements in the dash array set the length of each dash and space in the dash pattern. The first element sets the length of a dash, the second element sets the length of a space, the third element sets the length of a dash, and so forth.

The length of each dash and space in the dash pattern is the product of each element in the array and the width of the **Pen** object.

#### Example

```
' ========================================================================================
' The following example creates an array of real numbers and a Pen object, sets the dash
' pattern, and draws a custom dashed line. The code then gets the dash pattern currently
' set for the pen.
' ========================================================================================
SUB Example_GetPenDashArray (BYVAL hdc AS HDC)

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

  ' // Create and set an array of real numbers.
   DIM dashVals(3) AS SINGLE = {5.0, 2.0, 15.0, 4.0}

   ' // Create a Pen object.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 5, UnitWorld, @pen)

   ' // Set the dash pattern for the custom dashed line.
   hStatus = GdipSetPenDashArray(pen, @dashVals(0), 4)

   ' // Draw the custom dashed line.
   hStatus = GdipDrawLine(graphics, pen, 10, 20, 390, 200)

   ' // Obtain information about the pen.
   DIM count AS LONG
   DIM dashValues(ANY) AS SINGLE

   hStatus = GdipGetPenDashCount(pen, @count)
   REDIM dashValues(count - 1)
   hStatus = GdipGetPenDashArray(pen, @dashValues(0), count)

   FOR j AS LONG = 0 TO count - 1
      ' // Inspect or use the value in dashValues[j].
      OutputDebugStringW WSTR(dashValues(j))
   NEXT

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenDashCap197819

Gets the dash cap style currently set for a **Pen** object.

```
FUNCTION GdipGetPenDashCap197819 (BYVAL pen AS GpPen PTR, BYVAL dashCap AS GpDashCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *dashCap* | [out] Pointer to a variable that receives an element of the **GpDashCap** enumeration that indicates the dash cap style currently set for a **Pen** object. |

#### Example

```
' ========================================================================================
' This example sets a dash cap style for a Pen object, draws a dashed line, and then
' retrieves the dash cap using GdipGetPenDashCap197819 to verify the setting.
' ========================================================================================
SUB Example_GetPenDashCap (BYVAL hdc AS HDC)

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

   ' // Create a Pen object and set a dash style
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 10, UnitWorld, @pen)
   hStatus = GdipSetPenDashStyle(pen, DashStyleDash)

   ' // Set the dash cap style
   hStatus = GdipSetPenDashCap197819(pen, DashCapRound)

   ' // Draw a dashed line
   hStatus = GdipDrawLine(graphics, pen, 20, 50, 380, 50)

   ' // Retrieve the dash cap style
   DIM dashCap AS GpDashCap
   hStatus = GdipGetPenDashCap197819(pen, @dashCap)
   AfxMsg "Dash cap style: " & WSTR(dashCap)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenDashCount

Gets the number of elements in a dash pattern array.

```
FUNCTION GdipGetPenDashCount (BYVAL pen AS GpPen PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *count* | [out] Pointer to a variable that receives the number of elements in the dash pattern array. |

#### Example

```
' ========================================================================================
' This example sets a custom dash pattern for a Pen object, draws a dashed line, and then
' retrieves the number of elements in the dash array using GdipGetPenDashCount.
' ========================================================================================
SUB Example_GetPenDashCount (BYVAL hdc AS HDC)

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

   ' // Create a Pen object and set a custom dash pattern
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 5, UnitWorld, @pen)
   DIM dashVals(3) AS SINGLE = {5.0, 2.0, 15.0, 4.0}
   hStatus = GdipSetPenDashArray(pen, @dashVals(0), 4)

   ' // Draw a dashed line
   hStatus = GdipDrawLine(graphics, pen, 20, 50, 380, 50)

   ' // Retrieve the dash count
   DIM dashCount AS LONG
   hStatus = GdipGetPenDashCount(pen, @dashCount)
   AfxMsg "Dash count: " & WSTR(dashCount)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenDashOffset

Gets the number of elements in a dash pattern array.

```
FUNCTION GdipGetPenDashOffset (BYVAL pen AS GpPen PTR, BYVAL offset AS REAL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *offset* | [out] Pointer to a variable that receives the number of elements in the dash pattern array. |

#### Example

```
' ========================================================================================
' This example sets a dash offset for a Pen object, draws a dashed line, and then retrieves
' the offset using GdipGetPenDashOffset to verify the setting.
' ========================================================================================
SUB Example_GetPenDashOffset (BYVAL hdc AS HDC)

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

   ' // Create a Pen object and set a dash style
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 10, UnitWorld, @pen)
   hStatus = GdipSetPenDashStyle(pen, DashStyleDash)

   ' // Set a dash offset
   hStatus = GdipSetPenDashOffset(pen, 20)

   ' // Draw a dashed line
   hStatus = GdipDrawLine(graphics, pen, 20, 50, 380, 50)

   ' // Retrieve the dash offset
   DIM offset AS SINGLE
   hStatus = GdipGetPenDashOffset(pen, @offset)
   AfxMsg "Dash offset: " & WSTR(offset)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenDashStyle

Gets the number of elements in a dash pattern array.

```
FUNCTION GdipGetPenDashStyle (BYVAL pen AS GpPen PTR, BYVAL dashstyle AS GpDashStyle PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *dashstyle* | [out] Pointer to a variable that receives the number of elements in the dash pattern array. |

#### Example

```
' ========================================================================================
' This example sets a dash style for a Pen object, draws a line, and then retrieves the
' dash style using GdipGetPenDashStyle to verify the setting.
' ========================================================================================
SUB Example_GetPenDashStyle (BYVAL hdc AS HDC)

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

   ' // Create a Pen object and set a dash style
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 10, UnitWorld, @pen)
   hStatus = GdipSetPenDashStyle(pen, DashStyleDot)

   ' // Draw a dashed line
   hStatus = GdipDrawLine(graphics, pen, 20, 50, 380, 50)

   ' // Retrieve the dash style
   DIM dashStyle AS GpDashStyle
   hStatus = GdipGetPenDashStyle(pen, @dashStyle)
   AfxMsg "Dash style: " & WSTR(dashStyle)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenEndCap

Gets the end cap currently set for a **Pen** object.

```
FUNCTION GdipGetPenEndCap (BYVAL pen AS GpPen PTR, BYVAL endCap AS GpLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *endCap* | [out] Pointer to a variable that receives an element of the **GpLineCap** enumeration that indicates the end cap for a **Pen** object. |

#### Example

```
' ========================================================================================
' This example sets an end cap for a Pen object, draws a line, and then retrieves the
' end cap style using GdipGetPenEndCap to verify the setting.
' ========================================================================================
SUB Example_GetPenEndCap (BYVAL hdc AS HDC)

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

   ' // Create a Pen object and set its end cap
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 10, UnitWorld, @pen)
   hStatus = GdipSetPenEndCap(pen, LineCapArrowAnchor)

   ' // Draw a line with the pen
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 300, 50)

   ' // Retrieve the end cap style
   DIM endCap AS GpLineCap
   hStatus = GdipGetPenEndCap(pen, @endCap)
   OutputDebugStringW(WSTR("End cap style: " & endCap))

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenFillType

Gets the type currently set for a **Pen** object.

```
FUNCTION GdipGetPenFillType (BYVAL pen AS GpPen PTR, BYVAL type AS GpPenType PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *type* | [in] [out] Pointer to a variable that receives an element of the **PenType** enumeration that indicates the style of pen currently set for a **Pen** object. |

#### Example

```
' ========================================================================================
' This example creates a Pen object using a HatchBrush, draws a line, and then retrieves
' the pen's fill type using GdipGetPenFillType to verify the brush-based configuration.
' ========================================================================================
SUB Example_GetPenFillType (BYVAL hdc AS HDC)

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

   ' // Create a HatchBrush
   DIM hatchBrush AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleDiagonalCross, ARGB_RED, ARGB_YELLOW, @hatchBrush)

   ' // Create a Pen using the HatchBrush
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen2(hatchBrush, 10, UnitWorld, @pen)

   ' // Draw a line with the pen
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 300, 50)

   ' // Retrieve the pen's fill type
   DIM fillType AS GpPenType
   hStatus = GdipGetPenFillType(pen, @fillType)
   AfxMsg "Pen fill type: " & WSTR(fillType)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF hatchBrush THEN GdipDeleteBrush(hatchBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenLineJoin

Gets the line join style currently set for a **Pen** object.

```
FUNCTION GdipGetPenLineJoin (BYVAL pen AS GpPen PTR, BYVAL lineJoin AS GpLineJoin PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *lineJoin* | [out] Pointer to a variable that receives an element of the **LineJoin** enumeration that indicates the style used at the point where line segments join. |

---

## GdipGetPenMiterLimit

Gets the miter length currently set for a **Pen** object.

```
FUNCTION GdipGetPenMiterLimit (BYVAL pen AS GpPen PTR, BYVAL miterLimit AS REAL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *miterLimit* | [out] The miter limit of a **Pen** object. |

#### Example

```
' ========================================================================================
' This example sets the miter limit for a Pen object, draws a sharp-angled polygon to
' show the effect, and then retrieves the miter limit to inspect its value.
' ========================================================================================
SUB Example_SetGetMiterLimit (BYVAL hdc AS HDC)

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

   ' // Create a Pen object and set its line join to miter
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 10, UnitWorld, @pen)
   hStatus = GdipSetPenLineJoin(pen, LineJoinMiter)

   ' // Set the miter limit
   hStatus = GdipSetPenMiterLimit(pen, 2.0)

   ' // Draw a polygon with sharp angles
   DIM points(2) AS GpPointF
   points(0).x = 50 : points(0).y = 50
   points(1).x = 100 : points(1).y = 150
   points(2).x = 150 : points(2).y = 50
   hStatus = GdipDrawPolygon(graphics, pen, @points(0), 3)

   ' // Retrieve the miter limit
   DIM miterLimit AS SINGLE
   hStatus = GdipGetPenMiterLimit(pen, @miterLimit)
   OutputDebugStringW("Miter limit: " & WSTR(miterLimit))

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenMode

Gets the alignment currently set for a **Pen** object.

```
FUNCTION GdipGetPenMode (BYVAL pen AS GpPen PTR, BYVAL penMode AS GpPenAlignment PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *penMode* | [out] Pointer to a variable that receives an element of the **GpPenAlignment** enumeration that indicates the current alignment setting for a Pen object. |

#### Example

```
' ========================================================================================
' This example sets a pen mode, draws a line, and then retrieves the mode using
' GdipGetPenMode to verify the setting.
' ========================================================================================
SUB Example_GetPenMode (BYVAL hdc AS HDC)

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

   ' // Create a Pen object
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_GREEN, 10, UnitWorld, @pen)

   ' // Set the pen mode (optional, defaults to PenAlignmentCenter)
   hStatus = GdipSetPenMode(pen, PenAlignmentInset)

   ' // Draw a line with the pen
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 300, 50)

   ' // Retrieve the pen mode
   DIM penMode AS GpPenAlignment
   hStatus = GdipGetPenMode(pen, @penMode)
   AfxMsg "Pen mode: " & WSTR(penMode)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenStartCap

Gets the start cap currently set for a **Pen** object.

```
FUNCTION GdipGetPenStartCap (BYVAL pen AS GpPen PTR, BYVAL startCap AS GpLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *startCap* | [out] Pointer to a variable that receives an element of the **GpLineCap** enumeration that specifies the start cap for a **Pen** object. |

#### Example

```
' ========================================================================================
' This example creates a Pen object, sets a custom start cap, draws a line, and then
' retrieves the start cap style using GdipGetPenStartCap.
' ========================================================================================
SUB Example_GetStartCap (BYVAL hdc AS HDC)

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

   ' // Create a Pen object and set its start cap
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_GREEN, 10, UnitWorld, @pen)
   hStatus = GdipSetPenStartCap(pen, LineCapRoundAnchor)

   ' // Draw a line with the pen
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 300, 50)

   ' // Retrieve the start cap style
   DIM startCap AS GpLineCap
   hStatus = GdipGetPenStartCap(pen, @startCap)
   AfxMsg "Start cap style: " & WSTR(startCap)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenTransform

Gets the world transformation matrix currently set for a **Pen** object.

```
FUNCTION GdipGetPenTransform (BYVAL pen AS GpPen PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *matrix* | [out] Pointer to a variable that receives the transformation matrix. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object, applies a transform to it, retrieves the
' transformation matrix using GdipGetPenTransform, and inspects its elements.
' ========================================================================================
SUB Example_GetPenTransform (BYVAL hdc AS HDC)

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

   ' // Create a pen and apply a transform
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 10, UnitWorld, @pen)
   hStatus = GdipRotatePenTransform(pen, 45, MatrixOrderPrepend)

   ' // Retrieve the pen's transformation matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)
   hStatus = GdipGetPenTransform(pen, matrix)

   ' // Inspect the matrix elements
   DIM elements(5) AS SINGLE
   hStatus = GdipGetMatrixElements(matrix, @elements(0))

   FOR i AS INTEGER = 0 TO 5
      OutputDebugStringW(WSTR(elements(i)))
   NEXT

   ' // Draw a line with the transformed pen
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 300, 50)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenUnit

Gets the world transformation matrix currently set for a **Pen** object.

```
FUNCTION GdipGetPenUnit (BYVAL pen AS GpPen PTR, BYVAL unit AS GpUnit PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *unit* | [out] Pointer to a variable that receives a member of the **GpUnit** enumeration that indicates the unit of measure for the width of the pen. |

#### Example

```
' ========================================================================================
' This example creates a Pen object with a specific unit, draws a line, and then retrieves
' the unit using GdipGetPenUnit to verify the setting.
' ========================================================================================
SUB Example_GetPenUnit (BYVAL hdc AS HDC)

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

   ' // Create a Pen object with UnitWorld
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 10, UnitWorld, @pen)

   ' // Draw a line with the pen
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 300, 50)

   ' // Retrieve the unit of the pen
   DIM unit AS GpUnit
   hStatus = GdipGetPenUnit(pen, @unit)
   AfxMsg "Pen unit: " & WSTR(unit)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPenWidth

Gets the width currently set for a **Pen** object.

```
FUNCTION GdipGetPenWidth (BYVAL pen AS GpPen PTR, BYVAL width AS REAL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *width* | [out] Pointer to a variable that receives the width of a **Pen** object. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object with a specified width and draws a line. The
' code then gets the width of the pen, creates a second pen based on the width of the
' first pen, and draws a second line.
' ========================================================================================
SUB Example_GetPenWidth (BYVAL hdc AS HDC)

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

   ' // Create a pen of width 2, and use it to draw a rectangle.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 15, UnitWorld, @pen)
   hStatus = GdipDrawRectangle(graphics, pen, 20, 20, 200, 100)

   ' // Get the width of the pen.
   DIM nWidth AS REAL
   hStatus = GdipGetPenWidth(pen, @nWidth)

   ' // Create another pen that has the same width.
   DIM pen2 AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 15, UnitWorld, @pen2)

   ' // Draw a second line.
   hStatus = GdipDrawLine(graphics, pen2, 20, 60, 200, 140)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF pen2 THEN GdipDeletePen(pen2)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipMultiplyPenTransform

Updates the world transformation matrix of a **Pen** object with the product of itself and another matrix.

```
FUNCTION GdipMultiplyPenTransform (BYVAL pen AS GpPen PTR, BYVAL matrix AS CONST GpMatrix PTR, _
   BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *matrix* | [in] Pointer to a **Matrix** object that is multiplied with the world transformation matrix of a **Pen** object. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the passed matrix is on the left, and *MatrixOrderAppend* specifies that the passed matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example sets a matrix and creates a Pen object. The code then sets the
' width of the pen, applies a rotation matrix and a stretch matrix to the pen, and then
' draws an ellipse.
' ========================================================================================
SUB Example_MultiplyPenTransform (BYVAL hdc AS HDC)

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

   ' // Create a pen, and use it to draw a rectangle.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2, UnitWorld, @pen)
   hStatus = GdipDrawRectangle(graphics, pen, 10, 50, 150, 100)

   ' // Change the widthofthepen
   hStatus = GdipSetPenWidth(pen, 5)
   ' // Rotate the pen
   hStatus = GdipRotatePenTransform (pen, 30, MatrixOrderPrepend)
   ' // Stretch the pen
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(1, 0, 0, 4, 0, 0, @matrix)   ' // vertical stretch
   hStatus = GdipMultiplyPenTransform(pen, matrix, MatrixOrderPrepend)

   ' // Draw an ellipse with the transformed pen.
   hStatus = GdipDrawEllipse(graphics, pen, 90, 30, 200, 200)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipResetPenTransform

Sets the world transformation matrix of a **Pen** object to the identity matrix.

```
FUNCTION GdipResetPenTransform (BYVAL pen AS GpPen PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |

#### Example

```
' ========================================================================================
' The following example creates a scale matrix and a Pen object, and then draws a rectangle.
' The code then scales the pen by the matrix and draws a second rectangle.
' ========================================================================================
SUB Example_ResetPenTransform (BYVAL hdc AS HDC)

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

   ' // Create a pen, and set its transformation.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2, UnitWorld, @pen)
   hStatus = GdipScalePenTransform(pen, 8, 4, MatrixOrderPrepend)

   ' // Draw a rectangle with the transformed pen.
   hStatus = GdipDrawRectangle(graphics, pen, 30, 50, 150, 100)

   ' // Reset the transfomation
   hStatus = GdipResetPenTransform(pen)

   ' // Draw a rectangle with no pen transformation.
   hStatus = GdipDrawRectangle(graphics, pen, 220, 50, 150, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipRotatePenTransform

Updates the world transformation matrix of a **Pen** object with the product of itself and a rotation matrix.

```
FUNCTION GdipRotatePenTransform (BYVAL pen AS GpPen PTR, BYVAL angle AS REAL, _
   BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *angle* | [in] Real number that specifies the angle of rotation in degrees. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the rotation matrix is on the left, and *MatrixOrderAppend* specifies that the rotation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets a scaling matrix to the pen. The code
' then scales and rotates the pen and draws aen ellipse.
' ========================================================================================
SUB Example_RotatePenTransform (BYVAL hdc AS HDC)

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

   ' // Create a pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 5, UnitWorld, @pen)

   ' // Scale the pen
   hStatus = GdipScalePenTransform(pen, 1, 4, MatrixOrderPrepend)
   ' // Rotate the pen
   hStatus = GdipRotatePenTransform (pen, 30, MatrixOrderPrepend)
   ' // Draw an ellipse with the transformed pen.
   hStatus = GdipDrawEllipse(graphics, pen, 100, 30, 200, 200)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipScalePenTransform

Sets the **Pen** object's world transformation matrix equal to the product of itself and a scaling matrix.

```
FUNCTION GdipScalePenTransform (BYVAL pen AS GpPen PTR, BYVAL sx AS REAL, BYVAL sy AS REAL, _
   BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *sx* | [in] Single precision value that specifies the horizontal scaling factor in the scaling matrix. |
| *sy* | [in] Single precision value that specifies the vertical scaling factor in the scaling matrix. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the scaling matrix is on the left, and *MatrixOrderAppend* specifies that the scaling matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a scale matrix and a Pen object, and then draws a rectangle.
' The code then scales the pen by the matrix and draws a second rectangle.
' ========================================================================================
SUB Example_ScalePenTransform (BYVAL hdc AS HDC)

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

   ' // Create a pen, and use it to draw a rectangle.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2, UnitWorld, @pen)
   hStatus = GdipDrawRectangle(graphics, pen, 30, 50, 150, 100)

   ' // Scale the pen
   hStatus = GdipScalePenTransform(pen, 8, 4, MatrixOrderPrepend)

   '// Draw a rectangle with the transformed pen.
   hStatus = GdipDrawRectangle(graphics, pen, 210, 50, 150, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenBrushFill

Sets the **Brush** object that a pen uses to fill a line.

```
FUNCTION GdipSetPenBrushFill (BYVAL pen AS GpPen PTR, BYVAL brush AS GpBrush PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *brush* | [in] Pointer to a **Brush** object for the pen to use to fill a line. |

#### Example

```
' ========================================================================================
' The following example creates a HatchBrush object and then passes the address of that
' HatchBrush object to a Pen constructor. The code then sets the brush for the pen and
' draws a line and a rectangle.
' ========================================================================================
SUB Example_SetPenBrushFill (BYVAL hdc AS HDC)

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

   ' // Create a HatchBrush object.
   DIM hatchBrush AS GpHatch PTR
   hStatus = GdipCreateHatchBrush(HatchStyleHorizontal, ARGB_RED, ARGB_BLUE, @hatchBrush)

     ' // Create a pen, and set the brush for the pen.
     DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 10, UnitWorld, @pen)
   hStatus = GdipSetPenBrushFill(pen, hatchBrush)

   ' // Draw a line and a rectangle with the pen.
   hStatus = GdipDrawLine(graphics, pen, 0, 0, 200, 100)
   hStatus = GdipDrawRectangle(graphics, pen, 0, 100, 200, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF hatchBrush THEN GdipDeleteBrush(hatchBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenColor

Sets the color for a **Pen** object.

```
FUNCTION GdipSetPenColor (BYVAL pen AS GpPen PTR, BYVAL argb AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *argb* | [in] ARGB color that specifies the color for the **Pen** object. |

#### Example

```
' ========================================================================================
' The following example creates a red pen and draws a line. The code then sets the pen's
' color to blue and draws a second line.
' ========================================================================================
SUB Example_SetPenColor (BYVAL hdc AS HDC)

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

   ' // Create a red pen, and use it to draw a line.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 5, UnitWorld, @pen)
   hStatus = GdipDrawLine(graphics, pen, 0, 0, 200, 100)

   ' // Change the pen's color to blue, and draw a second line.
   hStatus = GdipSetPenColor(pen, ARGB_BLUE)
   hStatus = GdipDrawLineI(graphics, pen, 0, 40, 200, 140)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenCompoundArray

Sets the compound array for a **Pen** object.

```
FUNCTION GdipSetPenCompoundArray (BYVAL pen AS GpPen PTR, BYVAL dash AS CONST REAL PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *dash* | [in] Pointer to an array of real numbers that specifies the compound array. The elements in the array must be in increasing order, not less than 0, and not greater than 1. |
| *count* | [in] Positive even integer that specifies the number of elements in the dash array. The integer must not be greater than the number of elements in the compound array. |

#### Remarks

Suppose you want a pen to draw two parallel lines where the width of the first line is 20 percent of the pen's width, the width of the space that separates the two lines is 50 percent of the pen' s width, and the width of the second line is 30 percent of the pen's width. Start by creating a Pen object and an array of real numbers. You can then set the compound array by passing the array with the values 0.0, 0.2, 0.7, and 1.0 to the GdipSetCompoundArray function.

If you set the alignment of a Pen object to *PenAlignmentInset*, you cannot use that pen to draw compound lines.

#### Example

```
' ========================================================================================
' The following example creates a Pen object and sets the compound array for the pen.
' The code then draws a line using the Pen object.
' ========================================================================================
SUB Example_SetPenCompoundArray (BYVAL hdc AS HDC)

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

   ' // Create an array of real numbers and a Pen object.
   DIM compVals(5) AS SINGLE = {0.0, 0.2, 0.5, 0.7, 0.9, 1.0}
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 30, UnitWorld, @pen)

   ' // Set the compound array of the pen.
   hStatus = GdipSetPenCompoundArray(pen, @compVals(0), 6)

   ' // Draw a line with the pen.
   hStatus = GdipDrawLine(graphics, pen, 8, 20, 390, 200)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenCustomEndCap

Sets the custom end cap for a **Pen** object.

```
FUNCTION GdipSetPenCustomEndCap (BYVAL pen AS GpPen PTR, BYVAL customCap AS GpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *customCap* | [in] Pointer to a **CustomLineCap** object that specifies the custom end cap for the **Pen** object. |

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom end cap based on the GraphicsPath object,
' and draws a line.
' ========================================================================================
SUB Example_SetPenCustomEndCap (BYVAL hdc AS HDC)

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

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @pStrokePath)
   hStatus = GdipAddPathRectangleI(pStrokePath, -10, -5, 20, 10)

   ' // Create a pen, and set the custom end cap based on the GraphicsPath object.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1, UnitWorld, @pen)
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, pStrokePath, LineCapFlat, 0, @custCap)
   hStatus = GdipSetPenCustomEndCap(pen, custCap)

   ' // Draw a line with the custom end cap.
   hStatus = GdipDrawLine(graphics, pen, 0, 0, 200, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF pStrokePath THEN GdipDeletePath(pStrokePath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenCustomStartCap

Sets the custom start cap for a **Pen* object.

```
FUNCTION GdipSetPenCustomStartCap (BYVAL pen AS GpPen PTR, BYVAL customCap AS GpCustomLineCap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *customCap* | [in] Pointer to a **CustomLineCap** object that specifies the custom start cap for the **Pen** object. |

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a rectangle to it. The code
' then creates a Pen object, sets the custom start cap based on the GraphicsPath object,
' and draws a line.
' ========================================================================================
SUB Example_SetPenCustomStartCap (BYVAL hdc AS HDC)

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

   ' // Create a GraphicsPath object, and add a rectangle to it.
   DIM pStrokePath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @pStrokePath)
   hStatus = GdipAddPathRectangleI(pStrokePath, -10, -5, 20, 10)

  ' // Create a pen, and set the custom end cap based on the GraphicsPath object.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1, UnitWorld, @pen)
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, pStrokePath, LineCapFlat, 0, @custCap)
   hStatus = GdipSetPenCustomStartCap(pen, custCap)

   ' // Draw a line with the custom end cap.
   hStatus = GdipDrawLine(graphics, pen, 20, 20, 200, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF pStrokePath THEN GdipDeletePath(pStrokePath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenDashArray

Sets an array of custom dashes and spaces for a **Pen** object.

```
FUNCTION GdipSetPenDashArray (BYVAL pen AS GpPen PTR, BYVAL dash AS CONST REAL PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *dash* | [in] Pointer to an array of real numbers that specifies the length of the custom dashes and spaces. All elements in the array must be positive real numbers. |
| *count* | [in] Long integer that specifies the number of elements in the dash array. The integer must be greater than 0 and not greater than the total number of elements in the array. |

#### Remarks

This method will set the **DashStyle** enumeration for the **Pen** object to *DashStyleCustom*.

The elements in the dash array set the length of each dash and space in the dash pattern. The first element sets the length of a dash, the second element sets the length of a space, the third element sets the length of a dash, and so forth.

The length of each dash and space in the dash pattern is the product of the element value in the array and the width of the **Pen** object.

#### Example

```
' ========================================================================================
' The following example creates an array of real numbers. The code then creates a Pen object,
' sets the dash pattern based on the array, and then draws the custom dashed line.
' ========================================================================================
SUB Example_SetPenDashArray (BYVAL hdc AS HDC)

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

  ' // Create and set an array of real numbers.
   DIM dashVals(3) AS SINGLE = {5.0, 2.0, 15.0, 4.0}

   ' // Create a Pen object.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 5, UnitWorld, @pen)

   ' // Set the dash pattern for the custom dashed line.
   hStatus = GdipSetPenDashArray(pen, @dashVals(0), 4)

   ' // Draw the custom dashed line.
   hStatus = GdipDrawLine(graphics, pen, 10, 20, 390, 200)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenDashCap197819

Sets the dash cap style for a **Pen** object.

```
FUNCTION GdipSetPenDashCap197819 (BYVAL pen AS GpPen PTR, BYVAL dashCap AS GpDashCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *dashCap* | [in] Element of the **GpDashCap** enumeration that specifies the dash cap for a **Pen** object. |

#### Remarks

If you set the alignment of a **Pen** object to *PenAlignmentInset*, you cannot use that pen to draw triangular dash caps.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash cap style, and draws a line.
' The code then gets the dash cap style of the pen and creates a second Pen object with the
' same dash cap style. Finally, the code draws a second line using the second pen.
' ========================================================================================
SUB Example_GetPenDashCap (BYVAL hdc AS HDC)

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

   ' // Create a pen with a dash cap.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 15, UnitWorld, @pen)
   hStatus = GdipSetPenDashStyle(pen, DashStyleDash)
   hStatus = GdipSetPenDashCap197819(pen, DashCapRound)

   hStatus = GdipDrawLine(graphics, pen, 10, 10, 400, 200)

   ' // Obtain the dash cap for the pen.
   DIM dashCap AS DashCap
   hStatus = GdipGetPenDashCap197819(pen, @dashCap)

   ' // Create another pen, and use the same dash cap.
   DIM pen2 AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 15, UnitWorld, @pen2)
   hStatus = GdipSetPenDashStyle(pen2, DashStyleDash)
   hStatus = GdipSetPenDashCap197819(pen2, dashCap)

   ' // Draw a second line.
   hStatus = GdipDrawLine(graphics, pen2, 10, 10, 400, 300)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF pen2 THEN GdipDeletePen(pen2)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenDashOffset

Sets the distance from the start of the line to the start of the first dash in a dashed line.

```
FUNCTION GdipSetPenDashOffset (BYVAL pen AS GpPen PTR, BYVAL offset AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *offset* | [in] Single precision value that specifies the number of times to shift the spaces in a dashed line. Each shift is equal to the length of a space in the dashed line. |

#### Remarks

A positive offset value shifts the first dash forward along the path, and a negative offset value shifts the start of the path forward along the first dash.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash style, and draws a line.
' The code then sets the pen's offset value and draws a second line.
' ========================================================================================
SUB Example_SetPenDashOffset (BYVAL hdc AS HDC)

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

   ' // Create a Pen object, set the dash style, and draw a line.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 15, UnitWorld, @pen)
   hStatus = GdipSetPenDashStyle(pen, DashStyleDash)
   hStatus = GdipDrawLine(graphics, pen, 0, 50, 400, 50)

   ' // Set the dash offset value for the pen, and draw a second line.
   hStatus = GdipSetPenDashOffset(pen, 10)
   hStatus = GdipDrawLine(graphics, pen, 0, 80, 400, 80)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenDashStyle

Sets the dash style for a **Pen** object.

```
FUNCTION GdipSetPenDashStyle (BYVAL pen AS GpPen PTR, BYVAL dashStyle AS GpDashStyle) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *dashStyle* | [in] Value that specifies the number of times to shift the spaces in a dashed line. Each shift is equal to the length of a space in the dashed line. |

#### Remarks

The length of the dashes in a dashed line is dependent on the dash style and the width of the Pen object. The length of the space that separates two dashes in a dashed line is equal to the width of the Pen object.

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the dash style, and draws a line. The code
' then resets the dash style, draws a second line, resets dash style again, and draws a third line.
' ========================================================================================
SUB Example_SetPenDashStyle (BYVAL hdc AS HDC)

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

   ' // Create a Pen object
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 15, UnitWorld, @pen)

   ' // Set the dash style for the pen, and draw a dashed line.
   hStatus = GdipSetPenDashStyle(pen, DashStyleDash)
   hStatus = GdipDrawLine(graphics, pen, 0, 50, 400, 50)

   ' // Reset the dash style for the pen, and draw a second line.
   hStatus = GdipSetPenDashStyle(pen, DashStyleDot)
   hStatus = GdipDrawLine(graphics, pen, 0, 80, 400, 180)

   ' // Reset the dash style for the pen, and draw a third line.
   hStatus = GdipSetPenDashStyle(pen, DashStyleDashDot)
   hStatus = GdipDrawLine(graphics, pen, 0, 110, 400, 210)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenEndCap

Sets the end cap for a **Pen** object.

```
FUNCTION GdipSetPenEndCap (BYVAL pen AS GpPen PTR, BYVAL endCap AS GpLineCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *endCap* | [in] Element of the **GpLineCap** enumeration that specifies the end cap of a line. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the end cap, and draws a line. The code
' then resets the end cap, draws a second line, resets the end cap again, and draws a third line.
' ========================================================================================
SUB Example_SetPenEndCap (BYVAL hdc AS HDC)

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

   ' // Create a Pen object.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 15, UnitWorld, @pen)

   ' // Set the end cap of the pen, and draw a line.
   hStatus = GdipSetPenEndCap(pen, LineCapArrowAnchor)
   hStatus = GdipDrawLine(graphics, pen, 0, 50, 390, 150)

   ' // Reset the end cap, and draw a second line.
   hStatus = GdipSetPenEndCap(pen, LineCapTriangle)
   hStatus = GdipDrawLine(graphics, pen, 0, 80, 390, 180)

   ' // Reset the end cap, and draw a third line.
   hStatus = GdipSetPenEndCap(pen, LineCapRound)
   hStatus = GdipDrawLine(graphics, pen, 0, 110, 390, 210)
   
   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenLineCap197819

Sets the cap styles for the start, end, and dashes in a line drawn with a pen.

```
FUNCTION GdipSetPenLineCap197819 (BYVAL pen AS GpPen PTR, BYVAL startCap AS GpLineCap, _
   BYVAL endCap AS GpLineCap, BYVAL dashCap AS GpDashCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *startCap* | [in] Element of the **GpLineCap** enumeration that specifies the start cap of a line. |
| *endCap* | [in] Element of the **GpLineCap** enumeration that specifies the end cap of a line. |
| *dashCap* | [in] Element of the **GpDashCap** enumeration that specifies the start and end caps of the dashes in a dashed line. |

---

## GdipSetPenLineJoin

Sets the line join for a **Pen** object.

```
FUNCTION GdipSetPenLineJoin (BYVAL pen AS GpPen PTR, BYVAL lineJoin AS GpLineJoin) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *lineJoin* | [in] Element of the **LineJoin** enumeration that specifies the join style used at the end of a line segment that meets another line segment. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the line join style, and draws a rectangle.
' The code then resets the line join style and draws a second rectangle.
' ========================================================================================
SUB Example_SetPenLineJoin (BYVAL hdc AS HDC)

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

   ' // Create a pen.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 15, UnitWorld, @pen)

   ' // Set the join style, and draw a rectangle.
   hStatus = GdipSetPenLineJoin(pen, LineJoinRound)
   hStatus = GdipDrawRectangle(graphics, pen, 20, 20, 160, 100)

   ' // Reset the join style, and draw a second rectangle.
   hStatus = GdipSetPenLineJoin(pen, LineJoinBevel)
   hStatus = GdipDrawRectangle(graphics, pen, 220, 20, 160, 100)
   
   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenMiterLimit

Sets the miter limit of a **Pen** object.

```
FUNCTION GdipSetPenMiterLimit (BYVAL pen AS GpPen PTR, BYVAL miterLimit AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *miterLimit* | [in] Single precision value that specifies the miter limit of a **Pen** object. A real number value that is less than 1.0f will be replaced with 1.0f. |

#### Remarks

The miter length is the distance from the intersection of the line walls on the inside of the join to the intersection of the line walls outside of the join. The miter length can be large when the angle between two lines is small. The miter limit is the maximum allowed ratio of miter length to stroke width. The default value is 10.0f.

If the miter length of the join of the intersection exceeds the limit of the join, then the join will be beveled to keep it within the limit of the join of the intersection.

#### Example

```
' ========================================================================================
' This example sets the miter limit for a Pen object, draws a sharp-angled polygon to
' show the effect, and then retrieves the miter limit to inspect its value.
' ========================================================================================
SUB Example_SetGetMiterLimit (BYVAL hdc AS HDC)

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

   ' // Create a Pen object and set its line join to miter
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 10, UnitWorld, @pen)
   hStatus = GdipSetPenLineJoin(pen, LineJoinMiter)

   ' // Set the miter limit
   hStatus = GdipSetPenMiterLimit(pen, 2.0)

   ' // Draw a polygon with sharp angles
   DIM points(2) AS GpPointF
   points(0).x = 50 : points(0).y = 50
   points(1).x = 100 : points(1).y = 150
   points(2).x = 150 : points(2).y = 50
   hStatus = GdipDrawPolygon(graphics, pen, @points(0), 3)

   ' // Retrieve the miter limit
   DIM miterLimit AS SINGLE
   hStatus = GdipGetPenMiterLimit(pen, @miterLimit)
   OutputDebugStringW("Miter limit: " & WSTR(miterLimit))

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenMode

Sets the alignment for a **Pen** object relative to the line.

```
FUNCTION GdipSetPenMode (BYVAL pen AS GpPen PTR, BYVAL penMode AS GpPenAlignment) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *penMode* | [in] Element of the **GpPenAlignment** enumeration that specifies the alignment setting of the pen relative to the line that is drawn. The default value is *PenAlignmentCenter*. |

#### Example

```
' ========================================================================================
' The following example creates two Pen objects and sets the alignment for one of the pens.
' The code then draws two lines using each of the pens.
' ========================================================================================
SUB Example_SetPenLMode (BYVAL hdc AS HDC)

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

   ' // Create a black and a green pen.
   DIM AS GpPen PTR blackPen, greenPen
   hStatus = GdipCreatePen1(ARGB_BLACK, 1, UnitWorld, @blackPen)
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 15, UnitWorld, @greenPen)

 ' // Set the alignment of the green pen.
   hStatus = GdipSetPenMode(greenPen, PenAlignmentInset)

   ' // Draw two ellipses using each pen.
   hStatus = GdipDrawEllipseI(graphics, greenPen, 150, 20, 100, 200)
   hStatus = GdipDrawEllipseI(graphics, blackPen, 150, 20, 100, 200)
   
   ' // Cleanup
   IF blackPen THEN GdipDeletePen(blackPen)
   IF greenPen THEN GdipDeletePen(greenPen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenStartCap

Sets the start cap for the **Pen** object.

```
FUNCTION GdipSetPenStartCap (BYVAL pen AS GpPen PTR, BYVAL startCap AS GpLineCap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *startCap* | [in] Element of the **GpLineCap** enumeration that specifies the start cap of a line. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object, sets the start cap, and draws a line. The code
' then resets the end cap, draws a second line, resets the start cap again, and draws a third line.
' ========================================================================================
SUB Example_SetPenStartCap (BYVAL hdc AS HDC)

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

   ' // Create a Pen object.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 15, UnitWorld, @pen)

   ' // Set the end cap of the pen, and draw a line.
   hStatus = GdipSetPenStartCap(pen, LineCapArrowAnchor)
   hStatus = GdipDrawLine(graphics, pen, 10, 50, 390, 150)

   ' // Reset the end cap, and draw a second line.
   hStatus = GdipSetPenStartCap(pen, LineCapTriangle)
   hStatus = GdipDrawLine(graphics, pen, 10, 80, 390, 180)

   ' // Reset the end cap, and draw a third line.
   hStatus = GdipSetPenStartCap(pen, LineCapRound)
   hStatus = GdipDrawLine(graphics, pen, 10, 110, 390, 210)
   
   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenTransform

Sets the world transformation of a **Pen** object.

```
FUNCTION GdipSetPenTransform (BYVAL pen AS GpPen PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *matrix* | [in] Pointer to a **Matrix** object that specifies the world transformation. |

#### Remarks

This function ignores the translation portion of the **Matrix** object.

#### Example

```
' ========================================================================================
' The following example creates a scale matrix and a Pen object, and then draws a rectangle.
' The code then scales the pen by the matrix and draws a second rectangle.
' ========================================================================================
SUB Example_SetPenTransform (BYVAL hdc AS HDC)

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

   ' // Create a pen, and use it to draw a rectangle.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2, UnitWorld, @pen)
   hStatus = GdipDrawRectangle(graphics, pen, 10, 50, 150, 100)

   ' // Scale the pen width by a factor of 20 in the horizontal
   ' // direction and a factor of 10 in the vertical direction.
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(20, 0, 0, 10, 0, 0, @matrix)
   hStatus = GdipSetPenTransform(pen, matrix)

   ' // Draw a rectangle with the transformed pen.
   hStatus = GdipDrawRectangle(graphics, pen, 200, 50, 150, 100)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenUnit

Sets the unit of measure for the pen specified by the pen parameter to the value specified by the unit parameter.

```
FUNCTION GdipSetPenUnit (BYVAL pen AS GpPen PTR, BYVAL unit AS GpUnit) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *unit* | [in] The unit parameter is a member of the **GpUnit** enumeration that specifies the unit of measure for the width of the pen. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object with UnitPixel, draws a line, then changes
' the unit to UnitWorld and draws a second line to show the difference in scaling.
' ========================================================================================
SUB Example_SetPenUnit (BYVAL hdc AS HDC)

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

   ' // Create a pen with UnitPixel and draw a line
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 10, UnitPixel, @pen)
   hStatus = GdipDrawLine(graphics, pen, 20, 20, 380, 20)

   ' // Change the pen's unit to UnitWorld
   hStatus = GdipSetPenUnit(pen, UnitWorld)
   hStatus = GdipDrawLine(graphics, pen, 20, 60, 380, 60)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPenWidth

Sets the width for a **Pen** object.

```
FUNCTION GdipSetPenWidth (BYVAL pen AS GpPen PTR, BYVAL width AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *width* | [in] Single precision value that specifies the width of the pen. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object with a specified width and draws a rectangle.
' The code then resets the width of the pen and draws a second rectangle.
' ========================================================================================
SUB Example_SetPenWidth (BYVAL hdc AS HDC)

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

   ' // Create a pen of width 2, and use it to draw a rectangle.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2, UnitWorld, @pen)
   hStatus = GdipDrawRectangle(graphics, pen, 20, 50, 160, 100)

   ' // Reset the pen's width to 15, and draw another rectangle.
   hStatus = GdipSetPenWidth(pen, 15)
   hStatus = GdipDrawRectangle(graphics, pen, 210, 50, 160, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTranslatePenTransform

Translates the local geometric transformation by the specified dimensions. This method prepends the translation to the transformation.

```
FUNCTION GdipTranslatePenTransform (BYVAL pen AS GpPen PTR, BYVAL dx AS REAL, _
   BYVAL dy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pen* | [in] Pointer to the **Pen** object. |
| *dx* | [in] Single precision value that specifies the horizontal component of the translation. |
| *dy* | [in] Single precision value that specifies the vertical component of the translation. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the translation matrix is on the left, and *MatrixOrderAppend* specifies that the translation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a Pen object, applies a translation transform to it,
' and draws two lines—one with the original pen and one with the translated pen.
' ========================================================================================
SUB Example_TranslatePenTransform (BYVAL hdc AS HDC)

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

   ' // Create a Pen object
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 10, UnitWorld, @pen)

   ' // Draw a line with the original pen
   hStatus = GdipDrawLine(graphics, pen, 50, 50, 300, 50)

   ' // Translate the pen's transform
   hStatus = GdipTranslatePenTransform(pen, 10, 10, MatrixOrderPrepend)

   ' // Draw a second line with the translated pen
   hStatus = GdipDrawLine(graphics, pen, 50, 100, 300, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
