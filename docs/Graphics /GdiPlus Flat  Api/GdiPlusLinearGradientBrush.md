## LinearGradientBrush Functions

The **LinearGradientBrush** functions allow to paint a color gradient in which the color changes evenly from the starting boundary line of the linear gradient brush to the ending boundary line of the linear gradient brush. The boundary lines of a linear gradient brush are two parallel straight lines. The color gradient is perpendicular to the boundary lines of the linear gradient brush, changing gradually across the stroke from the starting boundary line to the ending boundary line. The color gradient has one color at the starting boundary line and another color at the ending boundary line.

| Name       | Description |
| ---------- | ----------- |
| [GdipCreateLineBrush](#gdipcreatelinebrush) | Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors. |
| [GdipCreateLineBrushI](#gdipcreatelinebrush) | Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors. |
| [GdipCreateLineBrushFromRect](#gdipcreatelinebrushfromrect) | Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors. |
| [GdipCreateLineBrushFromRectI](#gdipcreatelinebrushfromrect) | Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors. |
| [GdipCreateLineBrushFromRectWithAngle](#gdipcreatelinebrushfromrectwithangle) | Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors. |
| [GdipCreateLineBrushFromRectWithAngleI](#gdipcreatelinebrushfromrectwithangle) | Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors. |
| [GdipGetLineBlend](#gdipgetlineblend) | Gets the blend factors and their corresponding blend positions from a **LinearGradientBrush** object. |
| [GdipGetLineBlendCount](#gdipgetlineblendcount) | Gets the number of blend factors currently set for a **LinearGradientBrush** object. |
| [GdipGetLineColors](#gdipgetlinecolors) | Gets the the starting color and ending color of a linear gradient brush. |
| [GdipGetLineGammaCorrection](#gdipgetlinegammacorrection) | Determines whether gamma correction is enabled for a **LinearGradientBrush** object. |
| [GdipGetLinePresetBlend](#gdipgetlinepresetblend) | Gets the colors currently set to be interpolated for a linear gradient brush and their corresponding blend positions. |
| [GdipGetLinePresetBlendCount](#gdipgetlinepresetblendcount) | Gets the number of colors currently set to be interpolated for a linear gradient brush. |
| [GdipGetLineRect](#gdipgetlinerect) | Gets the rectangle that defines the boundaries of the gradient. |
| [GdipGetLineRectI](#gdipgetlinerect) | Gets the rectangle that defines the boundaries of the gradient. |
| [GdipGetLineTransform](#gdipgetlinetransform) | Gets the transformation matrix of a linear gradient brush. |
| [GdipGetLineWrapMode](#gdipgetlinewrapmode) | Gets the wrap mode for a brush. The wrap mode determines how an area is tiled when it is painted with a brush. |
| [GdipMultiplyLineTransform](#gdipmultiplylinetransform) | Updates a brush's transformation matrix with the product of itself and another matrix. |
| [GdipResetLineTransform](#gdipresetlinetransform) | Resets the transformation matrix of a linear gradient brush to the identity matrix. This means that no transformation takes place. |
| [GdipRotateLineTransform](#gdiprotatelinetransform) | Updates a brush's current transformation matrix with the product of itself and a rotation matrix. |
| [GdipScaleLineTransform](#gdipscalelinetransform) | Updates a brush's current transformation matrix with the product of itself and a scaling matrix. |
| [GdipSetLineBlend](#gdipsetlineblend) | Sets the blend factors and the blend positions of a linear gradient brush to create a custom blend. |
| [GdipSetLineColors](#gdipsetlinecolors) | Sets the starting color and ending color of a linear gradient brush. |
| [GdipSetLineGammaCorrection](#gdipsetlinegammacorrection) | Specifies whether gamma correction is enabled for a linear gradient brush. |
| [GdipSetLineLinearBlend](#gdipsetlinelinearblend) | Sets the blend shape of a linear gradient brush to create a custom blend based on a triangular shape. |
| [GdipSetLinePresetBlend](#gdipsetlinepresetblend) | Sets the colors to be interpolated for a linear gradient brush and their corresponding blend positions. |
| [GdipSetLineSigmaBlend](#gdipsetlinesigmablend) | Sets the blend shape of a linear gradient brush to create a custom blend based on a bell-shaped curve. |
| [GdipSetLineTransform](#gdipsetlinetransform) | Sets the transformation matrix of a linear gradient brush. |
| [GdipSetLineWrapMode](#gdipsetlinewrapmode) | Sets the transformation matrix of a linear gradient brush. |
| [GdipTranslateLineTransform](#gdiptranslatelinetransform) | Updates a brush's current transformation matrix with the product of itself and a translation matrix. |

---

## GdipCreateLineBrush

Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors

```
FUNCTION GdipCreateLineBrush (BYVAL point1 AS CONST GpPointF PTR, BYVAL point2 AS CONST GpPointF PTR, _
   BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL wrapMode AS GpWrapMode, _
   BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
```
```
FUNCTION GdipCreateLineBrushI (BYVAL point1 AS CONST GpPoint PTR, BYVAL point2 AS CONST GpPoint PTR, _
   BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL wrapMode AS GpWrapMode, _
   BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *point1* | [in] Reference to a **GpPointF** object that specifies the starting point of the gradient. The starting boundary line passes through the starting point. |
| *point2* | [in] Reference to a **GpPointF** object that specifies the ending point of the gradient. The ending boundary line passes through the ending point. |
| *color1* | [in] ARGB color that specifies the color at the starting boundary line of a linear gradient brush. |
| *color2* | [in] ARGB color object that specifies the color at the ending boundary line of a linear gradient brush. |
| *wrapMode* | [in] A member of the WrapMode enumeration that specifies how areas filled with the brush are tiled. |
| *lineGradient* | [out] Pointer to a variable that receives a pointer to the new created **LinearGradientBrush** object. |

#### Example

```
' ========================================================================================
' Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
' This function is perfect for diagonal gradients or when you want full control over the
' gradient vector.
' ========================================================================================
SUB Example_CreateLineBrush (BYVAL hdc AS HDC)

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

   ' // Define gradient boundary points
   DIM pt1 AS GpPointF = (20, 10)
   DIM pt2 AS GpPointF = (60, 110)

   ' // Create brush
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrush(@pt1, @pt2, ARGB_BLACK, ARGB_BLUE, WrapModeTile, @brush)

   ' // Fill rectangle using brush
   GdipFillRectangle(graphics, brush, 0, 0, 200, 150)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateLineBrushFromRect

Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors.

```
FUNCTION GdipCreateLineBrushFromRect (BYVAL rect AS const GpRectF PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL mode AS LinearGradientMode, BYVAL wrapMode AS GpWrapMode, _
   BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
```
```
FUNCTION GdipCreateLineBrushFromRectI (BYVAL rect AS const GpRect PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB, BYVAL mode AS LinearGradientMode, BYVAL wrapMode AS GpWrapMode, _
   BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rect* | [in] Reference to a rectangle that specifies the starting and ending points of the gradient. The direction of the gradient, specified by mode, affects how these points are defined. The dimensions of the rectangle affect the direction of the gradient for forward diagonal mode and backward diagonal mode. |
| *color1* | [in] ARGB color that specifies the color at the starting boundary line of a linear gradient brush. |
| *color2* | [in] ARGB color that specifies the color at the ending boundary line of a linear gradient brush. |
| *mode* | [in] Element of the **LinearGradientMode** enumeration that specifies the direction of the gradient. |
| *wrapMode* | [in] A member of the **WrapMode** enumeration that specifies how areas filled with the brush are tiled. |
| *lineGradient* | [out] Pointer to a variable that receives a pointer to the new created **LinearGradientBrush** object. |

#### Example

```
' ========================================================================================
' Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
' This function is ideal for creating gradients aligned to rectangular UI elements or backgrounds.
' ========================================================================================
SUB Example_CreateLineBrushFromRect (BYVAL hdc AS HDC)

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

   ' // Define gradient rectangle
   DIM rcf AS GpRectF = (0, 0, 200, 100)

   ' // Create brush with horizontal gradient
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateLineBrushFromRectWithAngle

Creates a **LinearGradientBrush** object from a set of boundary points and boundary colors.

```
FUNCTION GdipCreateLineBrushFromRectWithAngle (BYVAL rect AS CONST GpRectF PTR, _
   BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS REAL, BYVAL isAngleScalable AS BOOL, _
   BYVAL wrapMode AS GpWrapMode, BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
```
```
FUNCTION GdipCreateLineBrushFromRectWithAngleI (BYVAL rect AS CONST GpRect PTR, _
   BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS REAL, BYVAL isAngleScalable AS BOOL, _
   BYVAL wrapMode AS GpWrapMode, BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rect* | [in] Reference to a rectangle that specifies the starting and ending points of the gradient. The upper-left corner of the rectangle is the starting point. The lower-right corner is the ending point. |
| *color1* | [in] ARGB color that specifies the color at the starting boundary line of a linear gradient brush. |
| *color2* | [in] ARGB color that specifies the color at the ending boundary line of a linear gradient brush. |
| *angle* | [in] Single precision value that, if *isAngleScalable* is TRUE, specifies the base angle from which the angle of the directional line is calculated, or that, if *isAngleScalable* is FALSE, specifies the angle of the directional line. The angle is measured from the top of the rectangle that is specified by rect and must be in degrees. The gradient follows the directional line. |
| *isAngleScalable* | [in] Boolean value that specifies whether the angle is scalable. If *isAngleScalable* is TRUE, the angle of the directional line is scalable; otherwise, the angle is not scalable. |
| *wrapMode* | [in] A member of the **WrapMode** enumeration that specifies how areas filled with the brush are tiled. |
| *lineGradient* | [out] Pointer to a variable that receives a pointer to the new created **LinearGradientBrush** object. |

#### Example

```
' ========================================================================================
' Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
' This function is perfect for diagonal gradients or when you want precise control over the
' gradient direction.
' ========================================================================================
SUB Example_CreateLineBrushFromRectWithAngle (BYVAL hdc AS HDC)

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

   ' // Define gradient rectangle
   DIM rcf AS GpRectF = (0, 0, 200, 100)

   ' // Create brush with angle
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRectWithAngle(@rcf, ARGB_RED, ARGB_BLUE, 45.0, TRUE, WrapModeTile, @brush)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLineBlend

Gets the blend factors and their corresponding blend positions from a **LinearGradientBrush** object.

```
FUNCTION GdipGetLineBlend (BYVAL brush AS GpLineGradient PTR, BYVAL blend AS REAL PTR, _
   BYVAL positions AS REAL PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *blend* | [out] Pointer to an array that receives the blend factors. Each number in the array indicates a percentage of the ending color and is in the range from 0.0 through 1.0. |
| *positions* | [out] Pointer to an array that receives the blend positions. Each number in the array indicates a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. A blend position between 0.0 and 1.0 indicates a line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* | [in] Long integer value that specifies the number of blend factors to retrieve. Before calling the **GdipGetLineBlend** function, call the **GdipGetLineBlendCount** function to determine the current number of blend factors. The number of blend positions retrieved is the same as the number of blend factors retrieved. |

#### Example

```
' ========================================================================================
' This example sets a custom blend, retrieves the count, and then extracts the blend factors
' and positions into arrays. You can use these values for logging, debugging, or applying
' to another brush.
' ========================================================================================
SUB Example_GetLineBlend (BYVAL hdc AS HDC)

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

   ' // Create a linear gradient brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Set custom blend
   DIM setFactors(3) AS SINGLE = {0.0, 0.4, 0.6, 1.0}
   DIM setPositions(3) AS SINGLE = {0.0, 0.2, 0.8, 1.0}
   GdipSetLineBlend(brush, @setFactors(0), @setPositions(0), 4)

   ' // Get blend count
   DIM count AS LONG
   GdipGetLineBlendCount(brush, @count)

   ' // Retrieve blend factors and positions
   DIM getFactors(count - 1) AS SINGLE
   DIM getPositions(count - 1) AS SINGLE
   GdipGetLineBlend(brush, @getFactors(0), @getPositions(0), count)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLineBlendCount

Gets the number of blend factors currently set for a **LinearGradientBrush** object.

```
FUNCTION GdipGetLineBlendCount (BYVAL brush AS GpLineGradient PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *count* | [out] Pointer to a variable that receives the number of blend factors currently set for a **LinearGradientBrush** object. If no custom blend has been set by using **GdipSetLineBlend**, or if invalid positions were passed to **GdipSetLineBlend**, then **GdipGeLinetBlendCount** returns 1. |

#### Example

```
' ========================================================================================
' If no custom blend is set, GdipGetLineBlendCount typically returns 1. You can use this
' count to allocate arrays before calling GdipGetLineBlend.
' ========================================================================================
SUB Example_GetLineBlendCount (BYVAL hdc AS HDC)

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

   ' // Create a linear gradient brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Set custom blend
   DIM factors(3) AS SINGLE = {0.0, 0.4, 0.6, 1.0}
   DIM positions(3) AS SINGLE = {0.0, 0.2, 0.8, 1.0}
   GdipSetLineBlend(brush, @factors(0), @positions(0), 4)

   ' // Retrieve blend count
   DIM count AS LONG
   GdipGetLineBlendCount(brush, @count)

   ' // (Optional) Use count for logic or display
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLineColors

Gets the the starting color and ending color of a linear gradient brush.

```
FUNCTION GdipGetLineColors (BYVAL brush AS GpLineGradient PTR, BYVAL colors AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *colors* | [out] Pointer to an array that receives the starting color and the ending color. The first color in the colors array is the color at the starting boundary line of the gradient; the second color in the colors array is the color at the ending boundary line. |

#### Example

```
' ========================================================================================
' This example shows how to extract the gradient’s endpoints. colors(0) holds the starting
' color, and colors(1) holds the ending color.
' ========================================================================================
SUB Example_GetLineColors (BYVAL hdc AS HDC)

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

   ' // Create a linear gradient brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Retrieve start and end colors
   DIM colors(1) AS ARGB
   GdipGetLineColors(brush, @colors(0))

   ' // (Optional) Use colors(0) and colors(1) for logic or display
   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLineGammaCorrection

Determines whether gamma correction is enabled for a **LinearGradientBrush** object.

```
FUNCTION GdipGetLineGammaCorrection (BYVAL brush AS GpLineGradient PTR, _
   BYVAL useGammaCorrection AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *useGammaCorrection* | [out] Pointer to a variable that receives a boolean value indicating if gamma correction is enabled (TRUE) or not (FALSE). |

#### Example

```
' ========================================================================================
' This example enables gamma correction and then queries its status. You can use the useGamma
' flag to conditionally adjust rendering or log brush settings.
' ========================================================================================
SUB Example_GetLineGammaCorrection (BYVAL hdc AS HDC)

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

   ' // Create a linear gradient brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_GREEN, ARGB_BLACK, LinearGradientModeVertical, WrapModeTile, @brush)

   ' // Enable gamma correction
   GdipSetLineGammaCorrection(brush, TRUE)

   ' // Check if gamma correction is enabled
   DIM useGamma AS BOOL
   GdipGetLineGammaCorrection(brush, @useGamma)

   ' // (Optional) Use useGamma value for logic or display
   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLinePresetBlend

Gets the colors currently set to be interpolated for a linear gradient brush and their corresponding blend positions.

```
FUNCTION GdipGetLinePresetBlend (BYVAL brush AS GpLineGradient PTR, _
   BYVAL blend AS ARGB PTR, BYVAL positions AS REAL PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *blend* | [out] Pointer to an array that receives the colors. A color of a given index in the blend array corresponds to the blend position of that same index in the *positions* array. |
| *positions* | [out] Pointer to an array that receives the blend positions. Each number in the array indicates a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. A blend position between 0.0 and 1.0 indicates a line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* |[in] Long integer value that specifies the number of elements in the blend array. This is the same as the number of elements in the positions array. Before calling the **GdipGetLinePresetBlend** function, call the **GdipGetLinePresetBlendCount** function to determine the current number of colors. The number of blend positions retrieved is the same as the number of colors retrieved. |

#### Example

```
' ========================================================================================
' This example sets a preset blend, retrieves the count, and then extracts the colors and
' positions into arrays. You can use these values for logging, debugging, or applying to
' another brush.
' ========================================================================================
SUB Example_GetLinePresetBlend (BYVAL hdc AS HDC)

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

   ' // Create a linear gradient brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_WHITE, ARGB_BLACK, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Set preset blend
   DIM setColors(2) AS ARGB = {ARGB_RED, ARGB_GREEN, ARGB_BLUE}
   DIM setPositions(2) AS SINGLE = {0.0, 0.5, 1.0}
   GdipSetLinePresetBlend(brush, @setColors(0), @setPositions(0), 3)

   ' // Get blend count
   DIM count AS LONG
   GdipGetLinePresetBlendCount(brush, @count)

   ' // Retrieve blend colors and positions
   DIM getColors(count - 1) AS ARGB
   DIM getPositions(count - 1) AS SINGLE
   GdipGetLinePresetBlend(brush, @getColors(0), @getPositions(0), count)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLinePresetBlendCount

Gets the number of colors currently set to be interpolated for a linear gradient brush.

```
FUNCTION GdipGetLinePresetBlendCount (BYVAL brush AS GpLineGradient PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *count* | [out] Pointer to a variable that receives the number of colors to be interpolated for a linear gradient brush. If no colors have been set by using **GdipSetLinePresetBlend**, or if invalid positions were passed to **GdipSetLinePresetBlend**, then **GdipGetLinePresetBlendCount** returns 0. |

#### Example

```
' ========================================================================================
' This example sets a preset blend and then retrieves how many blend colors are active.
' You can use the count value to allocate arrays for GdipGetLinePresetBlend.
' ========================================================================================
SUB Example_GetLinePresetBlendCount (BYVAL hdc AS HDC)

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

   ' // Create a linear gradient brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_WHITE, ARGB_BLACK, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Set preset blend
   DIM colors(2) AS ARGB = {ARGB_RED, ARGB_GREEN, ARGB_BLUE}
   DIM positions(2) AS SINGLE = {0.0, 0.5, 1.0}
   GdipSetLinePresetBlend(brush, @colors(0), @positions(0), 3)

   ' // Retrieve blend count
   DIM count AS LONG
   GdipGetLinePresetBlendCount(brush, @count)

   ' (Optional) Use count for logic or display
   ' For demonstration, fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLineRect

Gets the rectangle that defines the boundaries of the gradient.

```
FUNCTION GdipGetLineRect (BYVAL brush AS GpLineGradient PTR, BYVAL rcf AS GpRectF PTR) AS GpStatus
```
```
FUNCTION GdipGetLineRectI (BYVAL brush AS GpLineGradient PTR, BYVAL rc AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *rect* | [out] Pointer to a **GpRectF** object that receives the rectangle that defines the boundaries of the gradient. For example, if a linear gradient brush is constructed with a starting point at (20.2, 50.8) and an ending point at (60.5, 110.0), then the defining rectangle has its upper-left point at (20.2, 50.8), a width of 40.3, and a height of 59.2. |

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush. Then the code gets the brush's
' rect and draws a rectangle with these values.
' ========================================================================================
SUB Example_GetLineRect (BYVAL hdc AS HDC)

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

   ' // Create a linear gradient brush
   DIM pt1 AS GpPointF = (20, 10)
   DIM pt2 AS GpPointF = (60, 110)
   DIM brush AS GpLineGradient PTR
   hStatus = GdipCreateLineBrush(@pt1, @pt2, ARGB_BLACK, ARGB_BLUE, WrapModeTile, @brush)
   
   ' // Get the rectangle that defines the boundaries of the gradient.
   DIM rcf AS GpRectF
   hStatus = GdipGetLineRect(brush, @rcf)

   ' // Draw the retrieved rectangle.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 1 * rxRatio, UnitPixel, @pen)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLineTransform

Gets the transformation matrix of a linear gradient brush.

```
FUNCTION GdipGetLineTransform (BYVAL brush AS GpLineGradient PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *matrix* | [in] [out] Pointer to a variable that receives the transformation matrix. |

#### Example

```
' ========================================================================================
' This example shows how to extract the brush’s transformation matrix after applying a rotation.
' You could also inspect the matrix values using GdipGetMatrixElements if needed.
' ========================================================================================
SUB Example_GetLineTransform (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Apply a transformation
   GdipRotateLineTransform(brush, 30.0, MatrixOrderPrepend)

   ' // Retrieve the transformation matrix
   DIM matrix AS GpMatrix PTR
   GdipCreateMatrix(@matrix)
   GdipGetLineTransform(brush, matrix)

   ' (Optional) Use matrix for further operations or inspection

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetLineWrapMode

Gets the wrap mode for a brush. The wrap mode determines how an area is tiled when it is painted with a brush.

```
GdipGetLineWrapMode (BYVAL brush AS GpLineGradient PTR, BYVAL wrapmode AS GpWrapMode PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *wrapmode* | [out] Pointer to a variable that receives one of the following elements of the **WrapMode** enumeration: *WrapModeTile*, *WrapModeTileFlipX*, *WrapModeTileFlipY*, *WrapModeTileFlipXY*. |

#### Example

```
' ========================================================================================
' The wrapMode variable will contain one of the WrapMode enum values like WrapModeTile,
' WrapModeTileFlipX, WrapModeTileFlipY, or WrapModeTileFlipXY. You can use this to conditionally
' adjust rendering or debug brush behavior.
' ========================================================================================
SUB Example_GetLineWrapMode (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 100, 50)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTileFlipXY, @brush)

   ' // Retrieve wrap mode
   DIM wrapMode AS GpWrapMode
   GdipGetLineWrapMode(brush, @wrapMode)

   ' // Optionally use wrapMode value (e.g., display or log)
   ' // For demonstration, fill a rectangle
   GdipFillRectangle(graphics, brush, 0, 0, 300, 150)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipMultiplyLineTransform

Updates a brush's transformation matrix with the product of itself and another matrix.

```
FUNCTION GdipMultiplyLineTransform (BYVAL brush AS GpLineGradient PTR, _
   BYVAL matrix AS CONST GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *matrix* | [in] Pointer to a matrix to be multiplied by the brush's current transformation matrix. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the passed matrix is on the left, and *MatrixOrderAppend* specifies that the passed matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' This example demonstrates how to combine transformations in a modular way. You can swap
' in MatrixOrderAppend to see how the order affects the result.
' ========================================================================================
SUB Example_MultiplyLineTransform (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Create a transformation matrix
   DIM matrix AS GpMatrix PTR
   GdipCreateMatrix(@matrix)
   GdipScaleMatrix(matrix, 1.5, 1.0, MatrixOrderPrepend)
   GdipRotateMatrix(matrix, 30.0, MatrixOrderAppend)

   ' // Multiply the brush's transform by the matrix
   GdipMultiplyLineTransform(brush, matrix, MatrixOrderPrepend)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipResetLineTransform

Resets the transformation matrix of a linear gradient brush to the identity matrix. This means that no transformation takes place.

```
FUNCTION GdipResetLineTransform (BYVAL brush AS GpLineGradient PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |

#### Example

```
' ========================================================================================
SUB Example_ResetLineTransform (BYVAL hdc AS HDC)

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

   ' // Create linear brush
    DIM rcf AS GpRectF = (0, 0, 200, 100)
    DIM brush AS GpLineGradient PTR
    GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Apply a rotation transform
   GdipRotateLineTransform(brush, 45.0, MatrixOrderPrepend)

   ' // Reset the transformation matrix
   GdipResetLineTransform(brush)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipRotateLineTransform

Updates a brush's current transformation matrix with the product of itself and a rotation matrix.

```
FUNCTION GdipRotateLineTransform (BYVAL brush AS GpLineGradient PTR, _
   BYVAL angle AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *angle* | [in] Real number that specifies the angle of rotation in degrees. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the rotation matrix is on the left, and *MatrixOrderAppend* specifies that the rotation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' This example rotates the gradient direction without affecting the brush’s original geometry.
' You can combine it with scaling or translation for more dynamic effects.
' ========================================================================================
SUB Example_RotateLineTransform (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Rotate the brush's gradient by 45 degrees
   GdipRotateLineTransform(brush, 45.0, MatrixOrderPrepend)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipScaleLineTransform

Updates a brush's current transformation matrix with the product of itself and a scaling matrix.

```
FUNCTION GdipScaleLineTransform (BYVAL brush AS GpLineGradient PTR, BYVAL sx AS REAL, _
   BYVAL sy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *sx* | [in] Simple precision number that specifies the amount to scale in the x direction. |
| *sy* | [in] Simple precision number that specifies the amount to scale in the y direction. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the scaling matrix is on the left, and *MatrixOrderAppend* specifies that the scaling matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' This example stretches the gradient horizontally and compresses it vertically. You can
' experiment with different scale factors and matrix orders to achieve unique visual effects.
' ========================================================================================
SUB Example_ScaleLineTransform (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Scale the brush's gradient horizontally by 2x and vertically by 0.5x
   GdipScaleLineTransform(brush, 2.0, 0.5, MatrixOrderPrepend)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetLineBlend

Sets the blend factors and the blend positions of a linear gradient brush to create a custom blend.

```
FUNCTION GdipSetLineBlend (BYVAL brush AS GpLineGradient PTR, BYVAL blend AS CONST REAL PTR, _
   BYVAL positions AS const REAL PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *blend* | [in] Pointer to an array of real numbers that specify blend factors. Each number in the array specifies a percentage of the ending color and should be in the range from 0.0 through 1.0. |
| *positions* | [in] Pointer to an array of real numbers that specify blend positions. Each number in the array indicates a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. There must be at least two positions specified: the first position, which is always 0.0, and the last position, which is always 1.0. Otherwise, the behavior is undefined. A blend position between 0.0 and 1.0 indicates a line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* | [in] Integer that specifies the number of elements in the blend array. This is the same as the number of elements in the positions array. The blend factor at a given array index corresponds to the blend position at that same array index. |

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets a custom blend, and uses the
' brush to fill a rectangle.
' ========================================================================================
SUB Example_SetLineBlend (BYVAL hdc AS HDC)

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

   '// Create a linear gradient brush
   DIM rcf AS GpRectF = (0,0, 200, 100)
   DIM brush AS GpLineGradient PTR
   hStatus = GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

  ' // Sets the blend factors and the blend positions
   DIM factors(3) AS SINGLE = {0.0, 0.4, 0.6, 1.0}
   DIM positions(3) AS SINGLE = {0.0, 0.2, 0.8, 1.0}
   hStatus = GdipSetLineBlend(brush, @factors(0), @positions(0), 4)
   hStatus = GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets a custom blend, and uses the
' brush to fill a rectangle.
' ========================================================================================
SUB Example_SetLineBlend (BYVAL hdc AS HDC)

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

   ' // Create brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Apply triangular blend
   GdipSetLineLinearBlend(brush, 0.5, 1.0)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetLineColors

Sets the starting color and ending color of a linear gradient brush.

```
FUNCTION GdipSetLineColors (BYVAL brush AS GpLineGradient PTR, BYVAL color1 AS ARGB, _
   BYVAL color2 AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *color1* | [in] ARGB color that specifies the color at the starting boundary line of a linear gradient brush. |
| *color2* | [in] ARGB color that specifies the color at the ending boundary line of a linear gradient brush. |

#### Example

```
' ========================================================================================
' This example shows how to override the initial colors of the brush after creation. It’s
' useful when you want to reuse a brush with different color schemes dynamically.
' ========================================================================================
SUB Example_SetLineColors (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_WHITE, ARGB_BLACK, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Change the gradient colors
   GdipSetLineColors(brush, ARGB_GREEN, ARGB_PURPLE)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetLineGammaCorrection

Specifies whether gamma correction is enabled for a linear gradient brush.

```
FUNCTION GdipSetLineGammaCorrection (BYVAL brush AS GpLineGradient PTR, _
   BYVAL useGammaCorrection AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *useGammaCorrection* | [in] Boolean value that specifies whether gamma correction occurs during rendering. TRUE specifies that gamma correction is enabled, and FALSE specifies that gamma correction is not enabled. By default, gamma correction is disabled during construction of a **LinearGradientBrush object**. |

#### Example

```
' ========================================================================================
' Enables gamma correction to adjust how colors blend based on human perception.
' Helps gradients appear more visually consistent across devices.
' ========================================================================================
SUB Example_SetLineGammaCorrection (BYVAL hdc AS HDC)

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

   ' // Create a linear gradient brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_GREEN, ARGB_BLACK, LinearGradientModeVertical, WrapModeTile, @brush)

   ' // Enable gamma correction
   GdipSetLineGammaCorrection(brush, TRUE)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetLineLinearBlend

Sets the blend shape of a linear gradient brush to create a custom blend based on a triangular shape.

```
FUNCTION GdipSetLineLinearBlend (BYVAL brush AS GpLineGradient PTR, _
   BYVAL focus AS REAL, BYVAL scale AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *focus* | [in] Real number that specifies the position of the ending color. This number is a percentage of the distance between the boundary lines and must be in the range from 0.0 through 1.0. |
| *scale* | [in] Simple precision value that specifies the percentage of the gradient's ending color that gets blended, at the focus position, with the gradient's starting color. This number must be in the range from 0.0 through 1.0. The default value is 1.0, which specifies that the ending color is at full intensity. |

#### Example

```
' ========================================================================================
' This example creates a gradient where the blend peaks at the center (50%) and fades evenly
' toward both ends. You can tweak focus and scale to shift the peak or reduce intensity.
' ========================================================================================
SUB Example_SetLineLinearBlend (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Apply triangular blend with focus at 0.5 and full scale
   GdipSetLineLinearBlend(brush, 0.5, 1.0)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets a triangular-shaped blend,
' and uses the brush to fill a rectangle. Twice more, the code sets a triangular-shaped
' blend with different values and, each time, uses the brush to fill a rectangle.
' ========================================================================================
SUB Example_Gdip_SetBlendTriangularShape (BYVAL hdc AS HDC)

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

   ' // Create linear gradient brush
   DIM rcf AS GpRectF = (0, 0, 400, 200)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Add three rectangles filles with bell-shaped blend brushes
   hStatus = GdipSetLineLinearBlend(brush, 0.5, 0.6)
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 400, 50)
   hStatus = GdipSetLineLinearBlend(brush, 0.5, 0.8)
   hStatus = GdipFillRectangle(graphics, brush, 0, 66, 400, 50)
   hStatus = GdipSetLineLinearBlend(brush, 0.5, 1.0)
   hStatus = GdipFillRectangle(graphics, brush, 0, 133, 400, 50)
   hStatus = GdipSetLineLinearBlend(brush, 0.5, 1.8)
   hStatus = GdipFillRectangle(graphics, brush, 0, 198, 400, 50)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetLinePresetBlend

Sets the colors to be interpolated for a linear gradient brush and their corresponding blend positions.

```
FUNCTION GdipSetLinePresetBlend (BYVAL brush AS GpLineGradient PTR, BYVAL blend AS CONST ARGB PTR, _
   BYVAL positions AS CONST REAL PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *blend* | [in] Pointer to an array of ARGB colors that specify the colors to be interpolated for a linear gradient brush. A color of a given index in the blend array corresponds to the blend position of that same index in the *positions* array. |
| *positions* | [in] Pointer to an array of real numbers that specify the blend positions. Each number in the array specifies a percentage of the distance between the starting boundary and the ending boundary and is in the range from 0.0 through 1.0, where 0.0 indicates the starting boundary of the gradient and 1.0 indicates the ending boundary. There must be at least two positions specified: the first position, which is always 0.0, and the last position, which is always 1.0. Otherwise, the behavior is undefined. A blend position between 0.0 and 1.0 indicates the line, parallel to the boundary lines, that is a certain fraction of the distance from the starting boundary to the ending boundary. For example, a blend position of 0.7 indicates the line that is 70 percent of the distance from the starting boundary to the ending boundary. The color is constant on lines that are parallel to the boundary lines. |
| *count* | [in] Long integer value that specifies the number of elements in the blend array. This is the same as the number of elements in the *positions* array. |

#### Example

```
' ========================================================================================
' Sets multiple colors and their positions along the gradient, allowing for smooth transitions
' between more than two colors. Ideal for custom multi-stop gradients.
' ========================================================================================
SUB Example_SetLinePresetBlend (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_WHITE, ARGB_BLACK, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Set preset blend colors and positions
   DIM colors(2) AS ARGB = {ARGB_RED, ARGB_GREEN, ARGB_BLUE}
   DIM positions(2) AS SINGLE = {0.0, 0.5, 1.0}
   GdipSetLinePresetBlend(brush, @colors(0), @positions(0), 3)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetLineSigmaBlend

Sets the blend shape of a linear gradient brush to create a custom blend based on a bell-shaped curve.

```
FUNCTION GdipSetLineSigmaBlend (BYVAL brush AS GpLineGradient PTR, BYVAL focus AS REAL, _
    BYVAL scale AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *focus* | [in] Single precision value that specifies the position of the ending color. This number is a percentage of the distance between the boundary lines and must be in the range from 0.0 through 1.0. |
| *scale* | [in] Single precision value that specifies the percentage of the gradient's ending color that gets blended, at the focus position, with the gradient's starting color. This number must be in the range from 0.0 through 1.0. The default value is 1.0, which specifies that the ending color is at full intensity. |

#### Example

```
' ========================================================================================
' Applies a bell-shaped blend curve to the gradient, concentrating color intensity around
' a focal point. Useful for soft, natural-looking transitions.
' ========================================================================================
SUB Example_SetLineSigmaBlend (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_YELLOW, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Apply sigma blend
   GdipSetLineSigmaBlend(brush, 0.5, 1.0)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates a linear gradient brush, sets a bell-shaped blend, and uses
' the brush to fill a rectangle. Twice more, the code sets a bell-shaped blend with different
' values and, each time, uses the brush to fill a rectangle.
' ========================================================================================
SUB Example_Gdip_SetBlendBellShape (BYVAL hdc AS HDC)

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

   ' // Create linear gradient brush
   DIM rcf AS GpRectF = (0, 0, 400, 200)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Add three rectangles filles with bell-shaped blend brushes
   hStatus = GdipSetLineSigmaBlend(brush, 0.5, 0.6)
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 400, 50)
   hStatus = GdipSetLineSigmaBlend(brush, 0.5, 0.8)
   hStatus = GdipFillRectangle(graphics, brush, 0, 66, 400, 50)
   hStatus = GdipSetLineSigmaBlend(brush, 0.5, 1.0)
   hStatus = GdipFillRectangle(graphics, brush, 0, 133, 400, 50)
   hStatus = GdipSetLineSigmaBlend(brush, 0.5, 1.8)
   hStatus = GdipFillRectangle(graphics, brush, 0, 198, 400, 50)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetLineTransform

Sets the transformation matrix of a linear gradient brush.

```
FUNCTION GdipSetLineTransform (BYVAL brush AS GpLineGradient PTR, _
   BYVAL matrix AS CONST GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *matrix* | [in] Pointer to a **Matrix** object that specifies the transformation matrix. |

#### Example

```
' ========================================================================================
' This example demonstrates how to scale and rotate the gradient using a custom matrix.
' You can swap in GdipTranslateMatrix or GdipShearMatrix for other effects.
' ========================================================================================
SUB Example_SetLineTransform (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 100, 50)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Create a transformation matrix
   DIM matrix AS GpMatrix PTR
   GdipCreateMatrix(@matrix)
   GdipScaleMatrix(matrix, 2.0, 1.0, MatrixOrderPrepend)
   GdipRotateMatrix(matrix, 45.0, MatrixOrderAppend)

   ' // Apply the matrix to the brush
   GdipSetLineTransform(brush, matrix)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, 0, 0, 200, 100)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetLineWrapMode

Sets the transformation matrix of a linear gradient brush.

```
FUNCTION GdipSetLineWrapMode (BYVAL brush AS GpLineGradient PTR, BYVAL wrapmode AS GpWrapMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *wrapmode* | [in] Element of the **WrapMode** enumeration that specifies how areas painted with a linear gradient brush will be tiled. The value of a parameter must be one of the following elements: *WrapModeTile*, *WrapModeTileFlipX*, *WrapModeTileFlipY*. |

#### Example

```
' ========================================================================================
' This example uses WrapModeTileFlipXY to demonstrate mirrored tiling in both directions.
' You could also try WrapModeTileFlipX, WrapModeTileFlipY, or WrapModeClamp for different
' visual effects.
' ========================================================================================
SUB Example_SetLineWrapMode (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 100, 50)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_RED, ARGB_BLUE, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Change wrap mode to FlipXY
   GdipSetLineWrapMode(brush, WrapModeTileFlipXY)

   ' // Fill a larger rectangle to show tiling effect
   GdipFillRectangle(graphics, brush, 0, 0, 300, 150)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTranslateLineTransform

Updates a brush's current transformation matrix with the product of itself and a translation matrix.

```
FUNCTION GdipTranslateLineTransform (BYVAL brush AS GpLineGradient PTR, BYVAL dx AS REAL, _
   BYVAL dy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **LinearGradientBrush** object. |
| *dx* | [in] Simple precision value that specifies the horizontal component of the translation. |
| *dy* | [in] Simple precision value that specifies the vertical component of the translation. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the translation matrix is on the left, and *MatrixOrderAppend* specifies that the translation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' Shifts the gradient’s position by applying a translation matrix. Great for nudging the
' gradient without altering its shape or angle.
' ========================================================================================
SUB Example_TranslateLineTransform (BYVAL hdc AS HDC)

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

   ' // Create linear brush
   DIM rcf AS GpRectF = (0, 0, 200, 100)
   DIM brush AS GpLineGradient PTR
   GdipCreateLineBrushFromRect(@rcf, ARGB_GREEN, ARGB_BLACK, LinearGradientModeHorizontal, WrapModeTile, @brush)

   ' // Translate the brush
   GdipTranslateLineTransform(brush, 20.0, 10.0, MatrixOrderPrepend)

   ' // Fill rectangle
   GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
