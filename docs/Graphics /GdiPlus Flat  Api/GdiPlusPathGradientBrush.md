## PathGradientBrush Functions

The **PathGradientBrush** functions allow to set the attributes of a color gradient that you can use to fill the interior of a path with a gradually changing color. A path gradient brush has a boundary path, a boundary color, a center point, and a center color. When you paint an area with a path gradient brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point.

| Name       | Description |
| ---------- | ----------- |
| [GdipCreatePathGradient](#gdipcreatepathgradient) | Creates a **PathGradientBrush** object based on an array of points. |
| [GdipCreatePathGradientI](#gdipcreatepathgradient) | Creates a **PathGradientBrush** object based on an array of points. |
| [GdipCreatePathGradientFromPath](#gdipcreatepathgradientfrompath) | Creates a **PathGradientBrush** object based on a **GraphicsPath** object. |
| [GdipGetPathGradientBlend](#gdipgetpathgradientblend) | Gets the number of blend factors currently set for this path gradient brush. |
| [GdipGetPathGradientBlendCount](#gdipgetpathgradientblendcount) | Gets the number of blend factors currently set for this path gradient brush. |
| [GdipGetPathGradientCenterColor](#gdipgetpathgradientcentercolor) | Gets the color of the center point of this path gradient brush. |
| [GdipGetPathGradientCenterPoint](#gdipgetpathgradientcenterpoint) | Gets the center point of this path gradient brush. |
| [GdipGetPathGradientCenterPointI](#gdipgetpathgradientcenterpoint) | Gets the center point of this path gradient brush. |
| [GdipGetPathGradientFocusScales](#gdipgetpathgradientfocusscales) | Gets the focus scales of this path gradient brush. |
| [GdipGetPathGradientGammaCorrection](#gdipgetpathgradientgammacorrection) | Determines whether gamma correction is enabled for this path gradient brush. |
| [GdipGetPathGradientPath](#gdipgetpathgradientpath) | Not implemented in Microsoft Windows GDI+ version 1.0. |
| [GdipGetPathGradientPointCount](#gdipgetpathgradientpointcount) | Gets the number of points in the array of points that defines this brush's boundary path. |
| [GdipGetPathGradientPresetBlend](#gdipgetpathgradientpresetblend) | Gets the preset colors and blend positions currently specified for this path gradient brush. |
| [GdipGetPathGradientPresetBlendCount](#gdipgetpathgradientpresetblendcount) | Gets the number of preset colors currently specified for this path gradient brush. |
| [GdipGetPathGradientRect](#gdipgetpathgradientrect) | Gets the smallest rectangle that encloses the boundary path of this path gradient brush. |
| [GdipGetPathGradientRectI](#gdipgetpathgradientrect) | Gets the smallest rectangle that encloses the boundary path of this path gradient brush. |
| [GdipGetPathGradientSurroundColorCount](#gdipgetpathgradientsurroundcolorcount) | Gets the number of colors that have been specified for the boundary path of this path gradient brush. |
| [GdipGetPathGradientSurroundColorsWithCount](#gdipgetpathgradientsurroundcolorswithcount) | Gets the surround colors currently specified for this path gradient brush. |
| [GdipGetPathGradientTransform](#gdipgetpathgradienttransform) | Gets transformation matrix of this path gradient brush. |
| [GdipGetPathGradientWrapMode](#gdipgetpathgradientwrapmode) | Gets the wrap mode currently set for this path gradient brush. |
| [GdipMultiplyPathGradientTransform](#gdipmultiplypathgradienttransform) | Updates the brush's transformation matrix with the product of itself and another matrix. |
| [GdipResetPathGradientTransform](#gdipresetpathgradienttransform) | Resets the transformation matrix of this path gradient brush to the identity matrix. This means that no transformation will take place. |
| [GdipRotatePathGradientTransform](#gdiprotatepathgradienttransform) | Updates this brush's current transformation matrix with the product of itself and a rotation matrix. |
| [GdipScalePathGradientTransform](#gdipscalepathgradienttransform) | Updates this brush's current transformation matrix with the product of itself and a scaling matrix. |
| [GdipSetPathGradientBlend](#gdipsetpathgradientblend) | Sets the blend factors and the blend positions of this path gradient brush. |
| [GdipSetPathGradientCenterColor](#gdipsetpathgradientcentercolor) | Sets the center color of this path gradient brush. The center color is the color that appears at the brush's center point. |
| [GdipSetPathGradientCenterPoint](#gdipsetpathgradientcenterpoint) | Sets the center point of this path gradient brush. By default, the center point is at the centroid of the brush's boundary path, but you can set the center point to any location inside or outside the path. |
| [GdipSetPathGradientCenterPointI](#gdipsetpathgradientcenterpoint) | Sets the center point of this path gradient brush. By default, the center point is at the centroid of the brush's boundary path, but you can set the center point to any location inside or outside the path. |
| [GdipSetPathGradientFocusScales](#gdipsetpathgradientfocusscales) | Sets the focus scales of this path gradient brush. |
| [GdipSetPathGradientGammaCorrection](#gdipsetpathgradientgammacorrection) | Specifies whether gamma correction is enabled for this path gradient brush. |
| [GdipSetPathGradientLinearBlend](#gdipsetpathgradientlinearblend) | Sets the blend shape of this path gradient brush. |
| [GdipSetPathGradientPath](#gdipsetpathgradientpath) | Not implemented in Microsoft Windows GDI+ version 1.0. |
| [GdipSetPathGradientPresetBlend](#gdipsetpathgradientpresetblend) | Sets the preset colors and the blend positions of this path gradient brush. |
| [GdipSetPathGradientSigmaBlend](#gdipsetpathgradientsigmablend) | Sets the blend shape of this path gradient brush. |
| [GdipSetPathGradientSurroundColorsWithCount](#gdipsetpathgradientsurroundcolorswithcount) | Sets the surround colors of this path gradient brush. The surround colors are colors specified for discrete points on the brush's boundary path. |
| [GdipSetPathGradientTransform](#gdipsetpathgradienttransform) | Sets the transformation matrix of this path gradient brush. |
| [GdipSetPathGradientWrapMode](#gdipsetpathgradientwrapmode) | Sets the wrap mode of this path gradient brush. |
| [GdipTranslatePathGradientTransform](#gdiptranslatepathgradienttransform) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

---

## GdipCreatePathGradient

Creates a **PathGradientBrush** object based on an array of points.

```
FUNCTION GdipCreatePathGradient (BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, _
   BYVAL wrapMode AS GpWrapMode, BYVAL polyGradient AS GpPathGradient PTR PTR) AS GpStatus
```
```
FUNCTION GdipCreatePathGradientI (BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_, _
   BYVAL wrapMode AS GpWrapMode, BYVAL polyGradient AS GpPathGradient PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *points* | Pointer to an array of points that specifies the boundary path of the path gradient brush. |
| *count* | [in] Integer that specifies the number of elements in the *points* array. |
| *wrapMode* | [in] Element of the **WrapMode** enumeration that specifies how areas painted with the path gradient brush will be tiled. The default value is *WrapModeClamp*. |
| *polyGradient* | [out] Pointer to a variable that receives a pointer to the new **PathGradientBrush** object. |

#### Usage example

```
' // Define triangle points
DIM points(0 TO 2) AS GpPointF = {(0, 0), (100, 0), (50, 100)}
' // Create a gradient brush
DIM brush AS GpPathGradient PTR
hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)
```
---

## GdipCreatePathGradientFromPath

Creates a **PathGradientBrush** object based on a **GraphicsPath** object.

```
FUNCTION GdipCreatePathGradientFromPath (BYVAL path AS CONST GpPath PTR, _
   BYVAL polyGradient AS GpPathGradient PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to a **GraphicsPath** object that specifies the boundary path of the path gradient brush. |
| *polyGradient* | [out] Pointer to a variable that receives a pointer to the new **PathGradientBrush** object. |

#### Excample

```
' ========================================================================================
' The following example draws a star-shaped GraphicsPath.
' Creates a star-shaped GraphicsPath and fills it using a PathGradientBrush constructed
' from that path. The gradient transitions from a red center to a set of surround colors
' defined at each vertex of the star. DPI scaling is applied globally to ensure consistent
' rendering across high-resolution displays.
' ========================================================================================
SUB Example_CreatePathGradientFromPath (BYVAL hdc AS HDC)

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

   ' // Create a GraphicsPath object and initializes the fill mode.
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Fill the array of points.
   DIM pts(0 TO 9) AS GpPointF
   pts(0).x = 75  : pts(0).y = 0
   pts(1).x = 100 : pts(1).y = 50
   pts(2).x = 150 : pts(2).y = 50
   pts(3).x = 112 : pts(3).y = 75
   pts(4).x = 150 : pts(4).y = 150
   pts(5).x = 75  : pts(5).y = 100
   pts(6).x = 0   : pts(6).y = 150
   pts(7).x = 37  : pts(7).y = 75
   pts(8).x = 0   : pts(8).y = 50
   pts(9).x = 50  : pts(9).y = 50

   ' // Construct the path with the array of points.
   hStatus = GdipAddPathLine2(path, @pts(0), 10)

   ' // Use the path to construct a path gradient brush.
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradientFromPath(path, @brush)

   ' // Set the color at the center of the path to red.
   hStatus = GdipSetPathGradientCenterColor(brush, ARGB_RED)

   ' // Set the colors of the points in the array.
   DIM Colors(9) AS ARGB
   Colors(0) = ARGB_BLACK
   Colors(1) = ARGB_LIGHTGREEN
   Colors(2) = ARGB_BLUE
   Colors(3) = ARGB_WHITE
   Colors(4) = ARGB_BLACK
   Colors(5) = ARGB_LIGHTGREEN
   Colors(6) = ARGB_BLUE
   Colors(7) = ARGB_WHITE
   Colors(8) = ARGB_BLACK
   Colors(9) = ARGB_LIGHTGREEN
   DIM count AS LONG = 10
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @Colors(0), @count)

   ' // Fill the path with the path gradient brush.
   hStatus = GdipFillPath(graphics, brush, path)

   ' // Cleanup
   IF path THEN GdipDeletePath(path)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientBlend

Sets the blend factors and the blend positions of this path gradient brush.

```
FUNCTION GdipGetPathGradientBlend (BYVAL brush AS GpPathGradient PTR, _
  BYVAL blend AS REAL PTR, BYVAL positions AS REAL PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *blend* | [in] Pointer to an array of blend factors. Each number in the array should be in the range 0 through 1. |
| *positions* | [in] Pointer to an array of blend positions. Each number in the array should be in the range 0 through 1. |
| *count* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the rotation matrix is on the left, and *MatrixOrderAppend* specifies that the rotation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Remarks

A **PathGradientBrush** object has a boundary path and a center point. When you fill an area with a path gradient brush, the color changes gradually as you move from the boundary path to the center point. By default, the color is linearly related to the distance, but you can customize the relationship between color and distance by calling the *GdipSetPathGradientBlend* function.

#### Example

```
' ========================================================================================
' The following example demonstrates several PathGradient functions, including
' GdipSetPathGradientBlend, GdipGetPathGradientBlendCount, and GdipGetPathGradientBlend.
' The code creates a PathGradientBrush object and calls the GdipSetPathGradientBlend function
' to establish a set of blend factors and blend positions for the brush. Then the code calls
' the GdipGetPathGradientBlendCount function to retrieve the number of blend factors. After
' the number of blend factors is retrieved, the code allocates two buffers: one to receive
' the array of blend factors and one to receive the array of blend positions. Then the code
' calls the GdipGetPathGradientBlendCount method to retrieve the blend factors and the blend
' positions.
' ========================================================================================
SUB Example_GetPathGradientBlend (BYVAL hdc AS HDC)

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

   ' // Create a path that consists of a single ellipse.
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Use the path to construct a brush.
   DIM brush AS GpPathGradient PTR
   hstatus = GdipCreatePathGradientFromPath(path, @brush)

   ' // Set the color at the center of the path
   hStatus = GdipSetPathGradientCenterColor(brush, ARGB_BLUE)

   ' // Set the color along the entire boundary of the path
   DIM colour(0) AS ARGB = {ARGB_RED}
   DIM count AS LONG = 1
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @Colour(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}
   hStatus = GdipSetPathGradientBlend(brush, @factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   hStatus = GdipFillEllipse(graphics, brush, 100, 70, 200, 100)

   ' // Obtain information about the path gradient brush.
   DIM blendCount AS LONG
   hStatus = GdipGetPathGradientBlendCount(brush,@blendCount)
   DIM blendFactors(blendCount - 1) AS SINGLE
   DIM blendPositions(blendCount - 1) AS SINGLE
   hStatus = GdipGetPathGradientBlend(brush, @blendFactors(0), @blendPositions(0), blendCount)

   FOR i AS LONG = 0 TO blendCount - 1
      ' // Inspect or use the value in rgFactors(j)
      ' // Inspect or use the value in rgPositions(j)
      OutputDebugStringW WSTR(blendFactors(i)) & " " & WSTR(blendPositions(i))
   NEXT

   ' // Cleanup
   IF path THEN GdipDeletePath(path)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientBlendCount

Gets the number of blend factors currently set for this path gradient brush.

```
FUNCTION GdipGetPathGradientBlendCount (BYVAL brush AS GpPathGradient PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *count* | [out] Pointer to a variable that receives the number of blend factors currently set for this path gradient brush. |

####  Remarks

Before you call the **GdipGetPathGradientBlend** function, you must allocate two buffers: one to receive an array of blend factors and one to receive an array of blend positions. To determine the size of the required buffers, call the **GdipGetPathGradientBlendCount** function. The size (in bytes) of each buffer should be the return value of **GdipGetPathGradientBlendCount** multiplied by sizeof(REAL) [4].

#### Example

See example for **GdipGetPathGradientBlend**.

---

## GdipGetPathGradientCenterColor

Gets the color of the center point of this path gradient brush.

```
FUNCTION GdipGetPathGradientCenterColor (BYVAL brush AS GpPathGradient PTR, _
   BYVAL colors AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *colors* | [out] Pointer to a variable that receives the color of the center point. |

Remarks

By default, the center point of a **PathGradientBrush** object is the centroid of the brush's boundary path, but you can set the center point to any location, inside or outside the path, by calling the **GdipSetPathGradientCenterPoint** or **GdipSetPathGradientCenterPointI** functions.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object and uses it to fill an ellipse.
' Then the code calls the GdipGetPathGradientCenterColor function to obtain the center color
' and draws a rectangle fille with it.
' ========================================================================================
SUB Example_GetPathGradientCenterColor (BYVAL hdc AS HDC)

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

   ' // Create a path that consists of a single ellipse.
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Use the path to construct a brush.
   DIM brush AS GpPathGradient PTR
   hstatus = GdipCreatePathGradientFromPath(path, @brush)

   ' // Set the color at the center of the path
   hStatus = GdipSetPathGradientCenterColor(brush, ARGB_BLUE)

   ' // Set the color along the entire boundary of the path
   DIM colour(0) AS ARGB = {ARGB_RED}
   DIM count AS LONG = 1
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @Colour(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}
   hStatus = GdipSetPathGradientBlend(brush, @factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   hStatus = GdipFillEllipse(graphics, brush, 100, 70, 200, 100)

      ' // Obtain information about the path gradient brush.
   DIM centerColor AS ARGB
   hStatus = GdipGetPathGradientCenterColor(brush, @centerColor)

   ' // Fill a rectangle with the retrieved color.
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(centerColor, @solidBrush)
   hStatus = GdipFillRectangle(graphics, solidBrush, 100, 200, 200, 30)

   ' // Cleanup
   IF path THEN GdipDeletePath(path)
   IF brush THEN GdipDeleteBrush(brush)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientCenterPoint

Gets the center point of this path gradient brush.

```
FUNCTION GdipGetPathGradientCenterPoint (BYVAL brush AS GpPathGradient PTR, _
   BYVAL points AS GpPointF PTR) AS GpStatus
```
```
FUNCTION GdipGetPathGradientCenterPointI (BYVAL brush AS GpPathGradient PTR, _
   BYVAL points AS GpPoint PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *points* | [out] Pointer to a **GpPointF** structure that receives the center point. |

#### Remarks

By default, the center point of a **PathGradientBrush** object is at the centroid of the brush's boundary path, but you can set the center point to any location, inside or outside the path, by calling the **GdipSetPathGradientCenterPoint** function.

#### Example

```
' ========================================================================================
' The following example demonstrates several functions of the PathGradientBrush object, including
' GdipGetPathGradientCenterPoint and GdipSetPathGradientCenterColor. The code creates
' a PathGradientBrush object and then sets the brush's center color and boundary color.
' The code calls the GdipGetPathGradientCenterPoint function to determine the center point
' of the path gradient and then draws a line from the origin to that center point.
' ========================================================================================
SUB Example_GetPathGradientCenterPoint (BYVAL hdc AS HDC)

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

   ' // Create a path that consists of a single ellipse.
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Use the path to construct a brush.
   DIM brush AS GpPathGradient PTR
   hstatus = GdipCreatePathGradientFromPath(path, @brush)

   ' // Set the color at the center of the path
   hStatus = GdipSetPathGradientCenterColor(brush, ARGB_BLUE)

   ' // Set the color along the entire boundary of the path
   DIM colour(0) AS ARGB = {ARGB_RED}
   DIM count AS LONG = 1
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @Colour(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}
   hStatus = GdipSetPathGradientBlend(brush, @factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   hStatus = GdipFillEllipse(graphics, brush, 100, 70, 200, 100)

   ' // Get the center point
   DIM centerPoint AS GpPointF
   hStatus = GdipGetPathGradientCenterPoint(brush, @centerPoint)

   ' // Draw a line from the origin to the center of the ellipse.
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 1 * rxRatio, UnitPixel, @pen)
   hStatus = GdipDrawLine(graphics, pen, 0, 0, centerPoint.x, centerpoint.y)

   ' // Cleanup
   IF path THEN GdipDeletePath(path)
   IF brush THEN GdipDeleteBrush(brush)
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientFocusScales

Gets the focus scales of this path gradient brush.

```
FUNCTION GdipGetPathGradientFocusScales (BYVAL brush AS GpPathGradient PTR, _
     BYVAL xScale AS REAL PTR, BYVAL yScale AS REAL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *xScale* | [out] Pointer to a variable that receives the x focus scale value. |
| *yScale* | [out] Pointer to a variable that receives the y focus scale value. |

#### Remarks

By default, the center color of a path gradient is at the center point. By calling **SetFocusScales**, you can specify that the center color should appear along a path that surrounds the center point. For example, suppose the boundary path is a triangle and the center point is at the centroid of that triangle. Also assume that the boundary color is red and the center color is blue. If you set the focus scales to (0.2, 0.2), the color is blue along the boundary of a small triangle that surrounds the center point. That small triangle is the main boundary path scaled by a factor of 0.2 in the x direction and 0.2 in the y direction. When you paint with the path gradient brush, the color will change gradually from red to blue as you move from the boundary of the large triangle to the boundary of the small triangle. The area inside the small triangle will be filled with blue.

#### Example

```
' ========================================================================================
' The focus scales define how far the center color extends from the center point.
' Values range from 0.0 to 1.0:
' Lower values (e.g., 0.2) concentrate the center color tightly.
' Higher values (e.g., 0.8) spread the center color more broadly.
' ========================================================================================
SUB Example_GetPathGradientFocusScales (BYVAL hdc AS HDC)

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

   ' // Define triangle points
   DIM points(0 TO 2) AS GpPointF = {(0, 0), (100, 0), (50, 100)}

   ' // Create a gradient brush
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Set focus scales to shrink the center color region
   hStatus = GdipSetPathGradientFocusScales(brush, 0.3, 0.3)

   ' // Retrieve focus scales
   DIM AS SINGLE xScale, yScale
   hStatus = GdipGetPathGradientFocusScales(brush, @xScale, @yScale)

   ' Output result
   AfxMsg "Focus scale X: " & WSTR(xScale) & CHR(13, 10) & "Focus scale Y: " & WSTR(yScale)

   ' // Fill rectangle to visualize the brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 200, 200)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientGammaCorrection

Determines whether gamma correction is enabled for this path gradient brush.

```
FUNCTION GdipGetPathGradientGammaCorrection (BYVAL brush AS GpPathGradient PTR, _
   BYVAL useGammaCorrection AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *useGammaCorrection* | [out] Pointer to a variable that receives a boolean value that determines whether the gamma correction is enabled (TRUE) or not (FALSE). |

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush from a triangle.
' Enables gamma correction using GdipSetPathGradientGammaCorrection.
' Retrieves the current gamma correction state using GdipGetPathGradientGammaCorrection.
' Displays a message box with the current state.
' ========================================================================================
SUB Example_GetPathGradientGammaCorrection (BYVAL hdc AS HDC)

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

   ' // Define triangle points
   DIM points(0 TO 2) AS GpPointF = {(0, 0), (100, 0), (50, 100)}

   ' // Create PathGradientBrush
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Enable gamma correction
   hStatus = GdipSetPathGradientGammaCorrection(brush, TRUE)

   ' // Check if gamma correction is enabled
   DIM useGamma AS BOOL
   hStatus = GdipGetPathGradientGammaCorrection(brush, @useGamma)

   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 100, 100)

   ' Output result
   IF useGamma THEN
      AfxMsg "Gamma correction is enabled."
   ELSE
      AfxMsg "Gamma correction is disabled."
   END IF

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientPath

Not implemented in Microsoft Windows GDI+ version 1.0.

```
FUNCTION GdipGetPathGradientPath (BYVAL brush AS GpPathGradient PTR, _
   BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *path* | [out] Pointer to the **GraphicsPath** object. |

---

## GdipGetPathGradientPointCount

Gets the number of points in the array of points that defines this brush's boundary path.

```
FUNCTION GdipGetPathGradientPointCount (BYVAL brush AS GpPathGradient PTR, _
    BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *count* | [out] Pointer that receives the number of points in the array of points that defines this brush's boundary path. |

#### Example

```
' ========================================================================================
' Returns the number of points that define the boundary path of a PathGradientBrush.
' ========================================================================================
SUB Example_GetPathGradientPointCount (BYVAL hdc AS HDC)

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

   ' // Define polygon points
   DIM points(0 TO 3) AS GpPointF = {(30, 20), (150, 40), (100, 100), (60, 200)}

   '//  Create PathGradientBrush
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 4, WrapModeClamp, @brush)

   ' // Retrieve the number of boundary points
   DIM pointCount AS LONG
   hStatus = GdipGetPathGradientPointCount(brush, @pointCount)

   ' // Output result
   AfxMsg "Boundary point count: " & WSTR(pointCount)

   ' // Fill rectangle with the brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 200, 200)

  ' // Cleanup
  IF brush THEN GdipDeleteBrush(brush)
  IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientPresetBlend

Gets the preset colors and blend positions currently specified for this path gradient brush.

```
FUNCTION GdipGetPathGradientPresetBlend (BYVAL brush AS GpPathGradient PTR, _
   BYVAL blend AS ARGB PTR, BYVAL positions AS REAL PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *blend* | [out] Pointer to an array that receives the preset colors. A color of a given index in the *blend* array corresponds to the blend position of that same index in the *positions* array. |
| *positions* | [out] Pointer to an array that receives the blend positions. Each blend position is a number from 0 through 1, where 0 indicates the boundary of the gradient and 1 indicates the center point. A blend position between 0 and 1 indicates the set of all points that are a certain fraction of the distance from the boundary to the center point. For example, a blend position of 0.7 indicates the set of all points that are 70 percent of the way from the boundary to the center point. |
| *count* | [in] Long integer that specifies the number of elements in the *blend* colors array. This is the same as the number of elements in the positions array. |

#### Remarks

A simple path gradient brush has two colors: a boundary color and a center color. When you paint with such a brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point. You can create a more complex gradient by specifying an array of preset colors and an array of blend positions.

Before you call the **GdipGetPathGradientPresetBlend** function, you must allocate two buffers: one to hold the array of preset colors and one to hold the array of blend positions. You can call the **GdipGetPathGradientPresetBlendCount** function to determine the required size of those buffers. The size of the color buffer is the return value of **GdipGetPathGradientPresetBlendCount** multiplied by sizeof(Color) [4]. The size of the position buffer is the value of **GdipGetPathGradientPresetBlendCount** multiplied by sizeof(REAL) [4].

#### Example

```
' ========================================================================================
' Sets a simple gradient from red (at boundary) to blue (at center).
' Retrieves the preset blend configuration using GdipGetPathGradientPresetBlend.
' Prints each blend position and its corresponding ARGB color.
' Fills a rectangle to visualize the gradient.
' ========================================================================================
SUB Example_GetPathGradientPresetBlend (BYVAL hdc AS HDC)

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

   ' // Define triangle points
   DIM points(0 TO 2) AS GpPointF = {(0, 0), (100, 0), (50, 100)}

   ' // Create PathGradientBrush
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Set preset blend colors and positions
   DIM blendColors(0 TO 1) AS ARGB = {ARGB_RED, ARGB_BLUE}
   DIM blendPositions(0 TO 1) AS SINGLE = {0.0, 1.0}
   hStatus = GdipSetPathGradientPresetBlend(brush, @blendColors(0), @blendPositions(0), 2)

   ' // Retrieve preset blend count
   DIM blendCount AS LONG
   hStatus = GdipGetPathGradientPresetBlendCount(brush, @blendCount)

   ' // Allocate buffers to receive blend data
   DIM retrievedColors(blendCount - 1) AS ARGB
   DIM retrievedPositions(blendCount - 1) AS SINGLE
   hStatus = GdipGetPathGradientPresetBlend(brush, @retrievedColors(0), @retrievedPositions(0), blendCount)

   ' // Output blend info
   FOR i AS LONG = 0 TO blendCount - 1
      OutputDebugStringW("Blend Position: " & WSTR(retrievedPositions(i)) & "; Color ARGB: " & "h" & HEX(retrievedColors(i)))
   NEXT

   ' // Fill rectangle with the brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 200, 200)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientPresetBlendCount

Gets the number of preset colors currently specified for this path gradient brush.

```
FUNCTION GdipGetPathGradientPresetBlendCount (BYVAL brush AS GpPathGradient PTR, _
    BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *count* | [out] Pointer to a variable that receives the number of preset colors currently specified for this path gradient brush. |

#### Remarks

A simple path gradient brush has two colors: a boundary color and a center color. When you paint with such a brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point. You can create a more complex gradient by specifying an array of preset colors and an array of blend positions.

You can obtain the interpolation colors and interpolation positions currently set for a **PathGradientBrush** object by calling the GdipGetPathGradientPresetBlend function. Before you call the **GdipGetPathGradientPresetBlend** function, you must allocate two buffers: one to hold the array of interpolation colors and one to hold the array of interpolation positions. You can call the **GdipGetPathGradientPointCount** function to determine the required size of those buffers. The size of the color buffer is the return value of **GdipGetPathGradientPointCount** multiplied by sizeof(Color) [4]. The size of the position buffer is the value of **GdipGetPathGradientPointCount** multiplied by sizeof( REAL) [4].

#### Example

```
' ========================================================================================
' Sets a gradient with three preset colors and blend positions.
' Uses GdipGetPathGradientPresetBlendCount to determine how many blend entries are currently
' configured. This count is essential before allocating buffers for GdipGetPathGradientPresetBlend.
' Fills a rectangle to visualize the gradient.
' ========================================================================================
SUB Example_GetPathGradientPresetBlendCount (BYVAL hdc AS HDC)

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

   ' // Define triangle points
   DIM points(0 TO 2) AS GpPointF = {(0, 0), (100, 0), (50, 100)}

   ' // Create a path gradient
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Set preset blend colors and positions
   DIM blendColors(0 TO 2) AS ARGB = {ARGB_RED, ARGB_GREEN, ARGB_BLUE}
   DIM blendPositions(0 TO 2) AS SINGLE = {0.0, 0.5, 1.0}
   hStatus = GdipSetPathGradientPresetBlend(brush, @blendColors(0), @blendPositions(0), 3)

   ' // Retrieve preset blend count
   DIM blendCount AS LONG
   hStatus = GdipGetPathGradientPresetBlendCount(brush, @blendCount)

   ' Output result
   AfxMsg "Preset blend count: " & WsTR(blendCount)

   ' // Fill rectangle with the brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 200, 200)

  ' // Cleanup
  IF brush THEN GdipDeleteBrush(brush)
  IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientRect

Gets the smallest rectangle that encloses the boundary path of this path gradient brush.

```
FUNCTION GdipGetPathGradientRect (BYVAL brush AS GpPathGradient PTR, _
   BYVAL rect AS GpRectF PTR) AS GpStatus
```
```
FUNCTION GdipGetPathGradientRectI (BYVAL brush AS GpPathGradient PTR, _
   BYVAL rect AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *rect* | [out] Pointer to a **GpRectF** structure that receives the bounding rectangle. |

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a polygon that is defined
' by four points. The code calls the GdipGetPathGradientRect function to obtain the smallest
' rectangle that encloses the brush's boundary path. The code calls the GdipFillRectangle
' function passing the address of the PathGradientBrush object and a reference to the brush's
' bounding rectangle. That call fills only the portion of the bounding rectangle that is inside
' the brush's boundary path. Finally the code draws the outline of the bounding rectangle.
' ========================================================================================
SUB Example_GdipGetPathGradientRect (BYVAL hdc AS HDC)

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

   ' // Create a path gradient brush based on an array of points.
   DIM points(0 TO 3) AS GpPointF = {(30, 20), (150, 40), (100, 100), (60, 200)}
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 4, WrapModeClamp, @brush)

   ' // Retrieve the bounding rectangle
   DIM rcf AS GpRectF
   hStatus = GdipGetPathGradientRect(brush, @rcf)

   ' // Fill a rectangle with the gradient brush
   hStatus = GdipFillRectangle(graphics, brush, rcf.x, rcf.y, rcf.width, rcf.height)

   ' Draw the bounding rectangle using a pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 1.0 * rxRatio, UnitPixel, @pen)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.width, rcf.height)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientSurroundColorCount

Gets the number of colors that have been specified for the boundary path of this path gradient brush.

```
FUNCTION GdipGetPathGradientSurroundColorCount (BYVAL brush AS GpPathGradient PTR, _
   BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *count* | [out] Pointer to a variable thatg receives the number of colors that have been specified for the boundary path of this path gradient brush. |

#### Usage example

See full example for **GdipGetPathGradientSurroundColorsWithCount**.

```
' // Get surround color count
DIM colorCount AS LONG
hStatus = GdipGetPathGradientSurroundColorCount(brush, @colorCount)
```
---

## GdipGetPathGradientSurroundColorsWithCount

Gets the surround colors currently specified for this path gradient brush.

```
FUNCTION GdipGetPathGradientSurroundColorsWithCount (BYVAL brush AS GpPathGradient PTR, _
   BYVAL colors AS ARGB PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *colors* | [in] Pointer to an array that receives the surround colors. |
| *count* | [in, out] Pointer to a long integer that, on input, specifies the number of colors requested. If the method succeeds, this parameter, on output, receives the number of colors retrieved. If the method fails, this parameter does not receive a value. |

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path that
' is defined by an array of three points. The code calls the GdipSetPathGradientSurroundColorsWithCount
' function to specify a color for each of the points that define the triangle. The
' GdipGetPathGradientSurroundColorCount function determines the current number of surround
' colors (the colors specified for the brush's boundary path). Next, the code allocates a
' buffer large enough to receive the array of surround colors and calls GdipGetPathGradientSurroundColorsWithCount
' to fill that buffer. Finally the code fills a small square with each of the brush's surround colors.
' ========================================================================================
SUB Example_GetPathGradientSurroundColorsWithCount (BYVAL hdc AS HDC)

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

   ' // Creates a gradient brush based on an array of points
   DIM points(0 TO 2) AS GpPointF = {(100, 0), (200, 200), (0, 200)}
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Sets the surround colors
   DIM  count AS LONG = 3
   DIM colors(0 TO 2) AS ARGB = {ARGB_RED, ARGB_BLUE, ARGB_LIGHTGREEN}
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @colors(0), @count)
   
   ' // Get surround color count
   DIM colorCount AS LONG
   hStatus = GdipGetPathGradientSurroundColorCount(brush, @colorCount)

   ' // Retrieve surround colors
   DIM rgColors(0 TO 2) AS ARGB
   hStatus = GdipGetPathGradientSurroundColorsWithCount(brush, @rgColors(0), @colorCount)

   ' // Create solid brush for rendering
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_WHITE, @solidBrush)

   ' // Draw small rectangles with each surround color
   FOR j AS LONG = 0 TO colorCount - 1
      hStatus = GdipSetSolidFillColor(solidBrush, rgColors(j))
      hStatus = GdipFillRectangle(graphics, solidBrush, 15 * j, 0, 10, 10)
   NEXT

   ' // Cleanup
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientTransform

Gets transformation matrix of this path gradient brush.

```
FUNCTION GdipGetPathGradientTransform (BYVAL brush AS GpPathGradient PTR, _
   BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *matrix* | [out] Pointer to a variable that receives the transformation matrix. |

#### Remarks

A **PathGradientBrush** object maintains a transformation matrix that can store any affine transformation. When you use a path gradient brush to fill an area, GDI+ transforms the brush's boundary path according to the brush's transformation matrix and then fills the interior of the transformed path. The transformed path exists only during rendering; the boundary path stored in **PathGradientBrush** object is not transformed.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an array of three points.
' The GdipScalePathGradientTransform and GdipTranslatePathGradientTransform functions set
' the elements of the brush's transformation matrix so that the matrix represents a composite
' transformation (first scale, then translate). That composite transformation applies to
' the brush's boundary path, so the call to GdipFillRectangle fills the interior of a triangle
' that is the result of scaling and translating the boundary path. The code calls the
' GdipGetPathGradientTransform function to obtain the brush's transformation matrix and
' then calls the GdipGetMatrixElements function to fill an array with the matrix elements.
' ========================================================================================
SUB Example_GetPathGradientTransform (BYVAL hdc AS HDC)

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

   DIM pts(0 TO 2) AS GpPointF = {(0, 0), (100, 0), (100, 100)}
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@pts(0), 3, WrapModeClamp, @brush)
   
   ' // Apply scale and translation transforms to brush
   hStatus = GdipScalePathGradientTransform(brush, 3.0, 1.0, MatrixOrderPrepend)
   hStatus = GdipTranslatePathGradientTransform(brush, 10.0, 30.0, MatrixOrderPrepend)

   ' Fill a rectangle using the transformed brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 200, 200)

   ' Retrieve the transformation matrix from the brush
   DIM matrix AS GpMatrix PTR
   hStatus = GdipGetPathGradientTransform(brush, @matrix)

   ' Extract matrix elements
   DIM elements(0 TO 5) AS SINGLE
   hStatus = GdipGetMatrixElements(matrix, @elements(0))

   FOR j AS LONG = 0 TO 5
      ' Use or inspect elements(j)
      OutputDebugStringW(WSTR(elements(j)))
   NEXT

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathGradientWrapMode

Gets the wrap mode currently set for this path gradient brush.

```
FUNCTION GdipGetPathGradientWrapMode (BYVAL brush AS GpPathGradient PTR, _
   BYVAL wrapMode AS GpWrapMode PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *wrapMode* | [out] Pointer to a variable that receives an element of the **WrapMode** enumeration that indicates the wrap mode currently set for this path gradient brush. |

#### Remarks

The bounding rectangle of a path gradient brush is the smallest rectangle that encloses the brush's boundary path. When you paint the bounding rectangle with the path gradient brush, only the area inside the boundary path gets filled. The area inside the bounding rectangle but outside the boundary path does not get filled.

The default wrap mode for a path gradient brush is *WrapModeClamp*, which indicates that no painting occurs outside of the brush's bounding rectangle. All of the other wrap modes indicate that areas outside the brush's bounding rectangle will be tiled. Each tile is a copy (possibly flipped) of the filled path inside its bounding rectangle.

#### Example

```
' ========================================================================================
' Retrieves the current wrap mode of a PathGradientBrush, which determines how the brush
' behaves when painting outside its boundary path.
' ========================================================================================
SUB Example_GetPathGradientWrapMode (BYVAL hdc AS HDC)

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

   ' // Define triangle points
   DIM points(0 TO 2) AS GpPointF = {(30, 30), (150, 30), (90, 120)}

   ' // Create PathGradientBrush
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeTile, @brush)

   ' // Retrieve wrap mode
   DIM wrapMode AS GpWrapMode
   hStatus = GdipGetPathGradientWrapMode(brush, @wrapMode)

   ' // Fill a rectangle with the brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 400, 250)

   ' // Output result
   SELECT CASE wrapMode
      CASE WrapModeClamp
         AfxMsg "Wrap mode: Clamp (default)"
      CASE WrapModeTile
         AfxMsg "Wrap mode: Tile"
      CASE WrapModeTileFlipX
         AfxMsg "Wrap mode: TileFlipX"
      CASE WrapModeTileFlipY
         AfxMsg "Wrap mode: TileFlipY"
      CASE WrapModeTileFlipXY
         AfxMsg "Wrap mode: TileFlipXY"
      CASE ELSE
         AfxMsg "Wrap mode: Unknown"
   END SELECT

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipMultiplyPathGradientTransform

Updates the brush's transformation matrix with the product of itself and another matrix.

```
FUNCTION GdipMultiplyPathGradientTransform (BYVAL brush AS GpPathGradient PTR, _
    BYVAL matrix AS CONST GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *matrix* | [in] Pointer to the matrix that will be multiplied by the brush's current transformation matrix. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the passed matrix is on the left, and *MatrixOrderAppend* specifies that the passed matrix is on the right. The default value is *MatrixOrderPrepend*. |

####  Remarks

A single 3 3 matrix can store any sequence of affine transformations. If you have several 3 3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3 3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix R represents a rotation and matrix T represents a translation. If matrix M is the product RT, then matrix M represents a composite transformation: first rotate, then translate.

The order of matrix multiplication is important. In general, the matrix product RT is not the same as the matrix product TR. In the example given in the previous paragraph, the composite transformation represented by RT (first rotate, then translate) is not the same as the composite transformation represented by TR (first translate, then rotate).

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' code calls the GdipScalePathGradientTransform function to fill the brush's transformation
' matrix with the elements that represent a horizontal scaling by a factor of 3. Then the
' code calls the GdipMultiplyPathGradientTransform method of that same PathGradientBrush
' object to multiply the brush's existing transformation matrix by a matrix that represents
' a translation (10 right, 30 down). The MatrixOrderAppend argument indicates that the
' multiplication is performed with the translation matrix on the right. After the
' multiplication, the brush's transformation matrix represents a composite transformation:
' first scale, then translate. That composite transformation is applied to the brush's
' boundary path during the call to FillRectangle, so it is the area inside the transformed
' path that gets painted.
' ========================================================================================
SUB Example_MultiplyPathGradientTransform (BYVAL hdc AS HDC)

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

   ' // Define triangle points
   DIM points(0 TO 2) AS GpPointF = {(0, 0), (50, 0), (50, 50)}

   ' // Create PathGradientBrush
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Apply horizontal scale transform to brush
   hStatus = GdipScalePathGradientTransform(brush, 3.0, 1.0, MatrixOrderPrepend)

   ' // Create translation matrix (10 right, 30 down)
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(1.0, 0.0, 0.0, 1.0, 10.0, 30.0, @matrix)

   ' // Multiply brush's transform matrix by translation matrix
   hStatus = GdipMultiplyPathGradientTransform(brush, matrix, MatrixOrderAppend)

   ' // Fill rectangle using transformed brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 200, 200)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipResetPathGradientTransform

Resets the transformation matrix of this path gradient brush to the identity matrix. This means that no transformation will take place.

```
FUNCTION GdipResetPathGradientTransform (BYVAL brush AS GpPathGradient PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' code calls the GdipScalePathGradientTransform function to fill the brush's transformation
' matrix with the elements that represent a horizontal scaling by a factor of 3. The brush
' is transformed using a horizontal scale followed by a translation (100 units right,
' 50 units down), resulting in a composite transformation. The transformed brush is used to
' fill a rectangle, showing the effect of the transformation. Then, the transformation
' matrix of the brush is reset to the identity matrix using GdipResetPathGradientTransform,
' effectively removing all transformations. The same rectangle is filled again, this time
' using the untransformed brush, allowing a visual comparison between the transformed and
' original states. This technique is useful for restoring a brush to its default state after
' applying temporary transformations during rendering.
' ========================================================================================
SUB Example_ResetPathGradientTransform (BYVAL hdc AS HDC)

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

   ' // Define triangle points
   DIM points(0 TO 2) AS GpPointF = {(0, 0), (50, 0), (50, 50)}

   ' // Create PathGradientBrush
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Apply horizontal scale transform to brush
   hStatus = GdipScalePathGradientTransform(brush, 3.0, 1.0, MatrixOrderPrepend)

  ' // Apply translation transform (100 right, 50 down)
   hStatus = GdipTranslatePathGradientTransform(brush, 100.0, 50.0, MatrixOrderAppend)

   ' // Fill rectangle using transformed brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 400, 400)

   ' // Reset the transformation
   hStatus = GdipResetPathGradientTransform(brush)

   ' // Fill the same area with the path gradient brush (no transformation).
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 400, 400)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipRotatePathGradientTransform

Updates this brush's current transformation matrix with the product of itself and a rotation matrix.

```
FUNCTION GdipRotatePathGradientTransform (BYVAL brush AS GpPathGradient PTR, _
   BYVAL angle AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *angle* | [in] Single precision number that specifies the angle of rotation in degrees. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the rotation matrix is on the left, and *MatrixOrderAppend* specifies that the rotation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' calls to the GdipScalePathGradientTransform  and GdipRotatePathGradientTransform functions
' set the elements of the brush's transformation matrix so that it represents a composite
' transformation: first scale, then rotate. The code uses the path gradient brush twice to
' paint a rectangle: once before the transformation is set and once after the transformation
' is set.
' ========================================================================================
SUB Example_RotatePathGradientTransform (BYVAL hdc AS HDC)

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

   ' // Define triangle points
   DIM points(0 TO 2) AS GpPointF = {(0, 0), (50, 0), (50, 50)}

   ' // Create PathGradientBrush
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Fill rectangle with the path gradient brush (no transformation).
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 400, 400)

   ' // First scale
   hStatus = GdipScalePathGradientTransform(brush, 3.0, 1.0, MatrixOrderPrepend)
   ' // then rotate
   hStatus = GdipRotatePathGradientTransform(brush, 60.0, MatrixOrderAppend)

   ' // Fill rectangle using transformed brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 400, 400)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipScalePathGradientTransform

Updates this brush's current transformation matrix with the product of itself and a scaling matrix.

```
FUNCTION GdipScalePathGradientTransform (BYVAL brush AS GpPathGradient PTR, _
   BYVAL sx AS REAL, BYVAL sy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *sx* | [in] Simple precision number that specifies the horizontal scale factor. |
| *sy* | [in] Simple  precision number that specifies the vertical scale factor. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the rotation matrix is on the left, and *MatrixOrderAppend* specifies that the rotation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Remarks

A single 3 3 matrix can store any sequence of affine transformations. If you have several 3 3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3 3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix T represents a translation and matrix S represents a scaling. If matrix M is the product TS, then matrix M represents a composite transformation: first translate, then scale.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' calls to the PathGradientBrush::TranslateTransform and PathGradientBrush.ScaleTransform
' methods of the PathGradientBrush object set the elements of the brush's transformation
' matrix so that it represents a composite transformation: first translate, then scale. The
' code uses the path gradient brush twice to paint a rectangle: once before the transformation
' is set and once after the transformation is set.
' ========================================================================================
SUB Example_ScaleTransform (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(0, 0), GDIP_POINT(50, 0), GDIP_POINT(50, 50)}
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   ' // Fill an area with the path gradient brush (no transformation).
   graphics.FillRectangle(@pthGrBrush, 0, 0, 400, 400)

   pthGrBrush.TranslateTransform(50.0, 40.0)                ' // translate
   pthGrBrush.ScaleTransform(3.0, 2.0, MatrixOrderAppend)   ' // then scale

   ' // Fill the same area with the transformed path gradient brush.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 400, 400)

END SUB
' ========================================================================================
```
---

## GdipSetPathGradientBlend

Sets the blend factors and the blend positions of this path gradient brush.

```
FUNCTION GdipSetPathGradientBlend (BYVAL brush AS GpPathGradient PTR, BYVAL blend AS CONST REAL PTR, _
   BYVAL positions AS CONST REAL PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *blend* | [in] Pointer to an array of blend factors. Each number in the array should be in the range 0 through 1. |
| *positions* | [in] Pointer to an array of blend positions. Each number in the array should be in the range 0 through 1. |
| *count* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the rotation matrix is on the left, and *MatrixOrderAppend* specifies that the rotation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Remarks

A **PathGradientBrush** object has a boundary path and a center point. When you fill an area with a path gradient brush, the color changes gradually as you move from the boundary path to the center point. By default, the color is linearly related to the distance, but you can customize the relationship between color and distance by calling the **GdipSetPathGradientBlend** method.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an ellipse. The code
' calls the GdipSetPathGradientBlend function to establish a set of blend factors and blend
' positions for the brush. Then the code uses the path gradient brush to fill the ellipse.
' ========================================================================================
SUB Example_SetPathGradientBlend (BYVAL hdc AS HDC)

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

   ' // Create a path that consists of a single ellipse.
   DIM path AS GpPath PTR
   hstatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Use the path to construct a brush.
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradientFromPath(path, @brush)
   ' // Set the color at the center of the path to blue.
   hStatus = GdipSetPathGradientCenterColor(brush, ARGB_BLUE)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}
   hStatus = GdipSetPathGradientBlend(brush, @factors(0), @positions(0), 4)

   ' // Fill an area with the path gradient brush
   hStatus = GdipFillEllipse(graphics, brush, 100, 70, 200, 100)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathGradientCenterColor

Sets the center color of this path gradient brush. The center color is the color that appears at the brush's center point.

```
FUNCTION GdipSetPathGradientCenterColor (BYVAL brush AS GpPathGradient PTR, BYVAL colors AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *colors* | [in] An ARGB color that specifies the center color. |

#### Remarks

By default the center point is the centroid of the brush's boundary path, but you can set the center point to any location inside or outside the path.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object and calls the GdipSetPathGradientCenterColor
' function to set the center color to yellow. The GdipSetPathGradientSurroundColorsWithCount
' function sets the color along the entire boundary to red. The GdipFillRectangle function
' uses the path gradient brush to paint a rectangle that contains the ellipse.
' ========================================================================================
SUB Example_SetPathGradientCenterColor (BYVAL hdc AS HDC)

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

   ' // Create a path that consists of a single ellipse.
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Use the path to construct a brush.
   DIM brush AS GpPathGradient PTR
   hstatus = GdipCreatePathGradientFromPath(path, @brush)

   ' // Set the color at the center of the path
   hStatus = GdipSetPathGradientCenterColor(brush, ARGB_YELLOW)

   ' // Set the color along the entire boundary of the path
   DIM colour(0) AS ARGB = {ARGB_RED}
   DIM count AS LONG = 1
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @Colour(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}
   hStatus = GdipSetPathGradientBlend(brush, @factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   hStatus = GdipFillEllipse(graphics, brush, 100, 70, 200, 100)

   ' // Cleanup
   IF path THEN GdipDeletePath(path)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathGradientCenterPoint

Sets the center point of this path gradient brush. By default, the center point is at the centroid of the brush's boundary path, but you can set the center point to any location inside or outside the path.

```
FUNCTION GdipSetPathGradientCenterPoint (BYVAL brush AS GpPathGradient PTR, _
  BYVAL points AS CONST GpPointF PTR) AS GpStatus
```
```
FUNCTION GdipSetPathGradientCenterPointI (BYVAL brush AS GpPathGradient PTR, _
  BYVAL points AS CONST GpPoint PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *points* | [in] Reference to a **GpPointF** structure that specifies the center point. |

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an ellipse. The code
' sets the center color to yellow and sets the color along the boundary to red. By default,
' the center point would be at the center of the ellipse, but the call to the
' GdipSetPathGradientCenterPoint function sets the center point to (190, 100).
' ========================================================================================
SUB Example_SetPathGradientCenterPoint (BYVAL hdc AS HDC)

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

   ' // Create a path that consists of a single ellipse.
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Use the path to construct a brush.
   DIM brush AS GpPathGradient PTR
   hstatus = GdipCreatePathGradientFromPath(path, @brush)

   ' // Set the color at the center of the path
   hStatus = GdipSetPathGradientCenterColor(brush, ARGB_YELLOW)

   ' // Set the center point.
   DIM pt AS GpPointF = (190, 100)
   hstatus = GdipSetPathGradientCenterPoint(brush, @pt)

   ' // Set the color along the entire boundary of the path
   DIM colour(0) AS ARGB = {ARGB_RED}
   DIM count AS LONG = 1
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @Colour(0), @count)

   ' // Set blend factors and positions for the path gradient brush.
   DIM factors(0 TO 3) AS SINGLE = {0.0, 0.4, 0.8, 1.0}
   DIM positions(0 TO 3) AS SINGLE = {0.0, 0.3, 0.7, 1.0}
   hStatus = GdipSetPathGradientBlend(brush, @factors(0), @positions(0), 4)

   ' // Fill the ellipse with the path gradient brush.
   hStatus = GdipFillEllipse(graphics, brush, 100, 70, 200, 100)

   ' // Cleanup
   IF path THEN GdipDeletePath(path)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathGradientFocusScales

Sets the focus scales of this path gradient brush.

```
FUNCTION GdipSetPathGradientFocusScales (BYVAL brush AS GpPathGradient PTR, _
   BYVAL xScale AS REAL, BYVAL yScale AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *xScale* | [in] Single precision number that specifies the x focus scale. |
| *yScale* | [in] Single precision number that specifies the y focus scale. |

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' code calls the GdipSetPathGradientFocusScalesfunction to set the brush's focus scales
' to (0.2, 0.2). Then the code uses the path gradient brush to paint a rectangle that
' includes the triangular path.
' ========================================================================================
SUB Example_SetFocusScales (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   DIM points(0 TO 2) AS GpPoint = {GDIP_POINT(100, 0), GDIP_POINT(200, 200), GDIP_POINT(0, 200)}

   ' // No GraphicsPath object is created. The PathGradientBrush
   ' // object is constructed directly from the array of points.
   DIM pthGrBrush AS CGpPathGradientBrush = CGpPathGradientBrush(@points(0), 3)

   DIM colors(0 TO 1) AS ARGB = {GDIP_ARGB(255, 255, 0, 0), GDIP_ARGB(255, 0, 0, 255)}

   ' // red at the boundary of the outer triangle
   ' // blue at the boundary of the inner triangle
   DIM relativePositions(0 TO 1) AS SINGLE = {0.0, 1.0}
   pthGrBrush.SetInterpolationColors(@colors(0), @relativePositions(0), 2)

   ' // The inner triangle is formed by scaling the outer triangle
   ' // about its centroid. The scaling factor is 0.2 in both the x and y directions.
   pthGrBrush.SetFocusScales(0.2, 0.2)

   ' // Fill a rectangle that is larger than the triangle
   ' // specified in the Point array. The portion of the
   ' // rectangle outside the triangle will not be painted.
   graphics.FillRectangle(@pthGrBrush, 0, 0, 200, 200)

END SUB
' ========================================================================================
```

---

## GdipSetPathGradientGammaCorrection

Specifies whether gamma correction is enabled for this path gradient brush.

```
FUNCTION GdipSetPathGradientGammaCorrection (BYVAL brush AS GpPathGradient PTR, _
   BYVAL useGammaCorrection AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *useGammaCorrection* | [in] Boolean value that specifies whether gamma correction is enabled. TRUE specifies that gamma correction is enabled, and FALSE specifies that gamma correction is not enabled. |

#### Example

```
' ========================================================================================
' The GdipSetPathGradientGammaCorrection function enables or disables gamma correction for
' a PathGradientBrush, which affects how colors blend across the gradient. Enabling it can
' produce smoother and more perceptually accurate transitions, particularly when blending
' bright and dark colors.
' ========================================================================================
SUB Example_SetPathGradientGammaCorrection (BYVAL hdc AS HDC)

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

   ' // Define triangle points
   DIM points(0 TO 2) AS GpPointF = {(0, 0), (200, 0), (100, 100)}

   ' // Create PathGradientBrush
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Set center color
   DIM centerColor AS ARGB = ARGB_WHITE
   hStatus = GdipSetPathGradientCenterColor(brush, centerColor)

   ' // Set surround colors
   DIM surroundColors(0 TO 2) AS ARGB = {ARGB_RED, ARGB_LIGHTGREEN, ARGB_BLUE}

   DIM count AS LONG = 3
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @surroundColors(0), @count)

   ' // Enable gamma correction
   hStatus = GdipSetPathGradientGammaCorrection(brush, TRUE)

   ' // Fill a rectangle with the gamma-corrected brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 300, 300)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathGradientLinearBlend

Sets the blend shape of this path gradient brush.

```
FUNCTION GdipSetPathGradientLinearBlend (BYVAL brush AS GpPathGradient PTR, _
   BYVAL focus AS REAL, BYVAL scale AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *focus* | [in] Single precision number that specifies where the center color will be at its highest intensity. This number must be in the range 0 through 1. |
| *scale* | [in] Single precision number that specifies the maximum intensity of center color that gets blended with the boundary color. This number must be in the range 0 through 1. The default value is 1. |

#### Remarks

By default, as you move from the boundary of a path gradient to the center point, the color changes gradually from the boundary color to the center color. You can customize the positioning and blending of the boundary and center colors by calling the **SetBlendTriangularShape** method.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an ellipse. The code
' calls the GdipSetPathGradientLinearBlend function, passing a focus of 0.2 and a scale
' of 0.7. Then the code uses the path gradient brush to paint a rectangle that contains
' the ellipse.
' ========================================================================================
SUB Example_SetPathGradientLinearBlend (BYVAL hdc AS HDC)

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

   ' // Create a path that consists of a single ellipse.
   DIM path AS GpPath PTR
   hstatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Use the path to construct a brush.
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradientFromPath(path, @brush)

   ' // Set the color at the center of the path to red.
   hStatus = GdipSetPathGradientCenterColor(brush, ARGB_RED)

   ' // Set the surround color
   DIM colors(0) AS ARGB = {ARGB_BLUE}
   DIM count AS LONG = 1
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @colors(0), @count)

   ' // Set the blend shape
   hStatus = GdipSetPathGradientLinearBlend(brush, 0.2, 0.7)

   ' // Fill an area with the path gradient brush
   hStatus = GdipFillEllipse(graphics, brush, 100, 70, 200, 100)

   ' // The color is blue on the boundary and at the center.
   ' // At points that are 20 percent of the way from the boundary to the
   ' // center, the color is 70 percent red and 30 percent blue.

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathGradientPath

Not implemented in Microsoft Windows GDI+ version 1.0.

```
FUNCTION GdipSetPathGradientPath (BYVAL brush AS GpPathGradient PTR, BYVAL path AS CONST GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *path* | [in] Pointer to the **GraphicsPath** object. |

---

## GdipSetPathGradientPresetBlend

Sets the preset colors and the blend positions of this path gradient brush.

```
FUNCTION GdipSetPathGradientPresetBlend (BYVAL brush AS GpPathGradient PTR, _
   BYVAL blend AS CONST ARGB PTR, BYVAL positions AS CONST REAL PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *blend* | [in] Pointer to an array of ARGB colors that specifies the interpolation colors for the gradient. A color of a given index in the *blend* array corresponds to the blend position of that same index in the positions array. |
| *positions* | [in] Pointer to an array that specifies the blend positions. Each blend position is a number from 0 through 1, where 0 indicates the boundary of the gradient and 1 indicates the center point. A blend position between 0 and 1 specifies the set of all points that are a certain fraction of the distance from the boundary to the center point. For example, a blend position of 0.7 specifies the set of all points that are 70 percent of the way from the boundary to the center point. |
| *count* | [in] Integer that specifies the number of Color objects in the blend array. This is the same as the number of elements in the *positions* array. |

#### Remarks

A simple path gradient brush has two colors: a boundary color and a center color. When you paint with such a brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point. You can create a more complex gradient by specifying an array of preset colors and an array of blend positions.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' GdipSetPathGradientPresetBlend function sets the brush's preset colors to red, blue, and
' aqua and sets the blend positions to 0, 0, 4, and 1. The GdipFillRectangle function uses
' the path gradient brush to paint a rectangle that contains the triangular path.
' ========================================================================================
SUB Example_SetPathGradientPresetBlend (BYVAL hdc AS HDC)

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

   ' // Creates a gradient brush based on an array of points
   DIM pts(0 TO 2) AS GpPointF = {(100, 0), (200, 200), (0, 200)}
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@pts(0), 3, WrapModeClamp, @brush)

    ' // Array of interpolation colors
   DIM  count AS LONG = 3
   DIM colors(0 TO 2) AS ARGB = {ARGB_RED, ARGB_BLUE, ARGB_AQUA}
   ' // Array of color positions
   DIM positions(0 TO 2) AS SINGLE = {0.0, 0.4, 1.0}
   ' // Sets the preset colors and the blend positions of the path gradient brush.
   hStatus = GdipSetPathGradientPresetBlend(brush, @colors(0), @positions(0), 3)
   
   ' // Fill an area with the path gradient brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 300, 300)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathGradientSigmaBlend

Sets the blend shape of this path gradient brush.

```
FUNCTION GdipSetPathGradientSigmaBlend (BYVAL brush AS GpPathGradient PTR, _
   BYVAL focus AS REAL, BYVAL scale AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *focus* |[in] Single precision value that specifies where the center color will be at its highest intensity. This number must be in the range 0 through 1. |
| *scale* | [in] Single precision value that specifies the maximum intensity of center color that gets blended with the boundary color. This number must be in the range 0 through 1. The default value is 1. |

#### Remarks

By default, as you move from the boundary of a path gradient to the center point, the color changes gradually from the boundary color to the center color. You can customize the positioning and blending of the boundary and center colors by calling the **GdipSetPathGradientSigmaBlend** method.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an ellipse. The code
' calls the GdipSetPathGradientSigmaBlend function, passing a focus of 0.2 and a scale
' of 0.7. Then the code uses the path gradient brush to paint a rectangle that contains
' the ellipse.
' ========================================================================================
SUB Example_SetPathGradientSigmaBlend (BYVAL hdc AS HDC)

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

   ' // Create a path that consists of a single ellipse.
   DIM path AS GpPath PTR
   hstatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Use the path to construct a brush.
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradientFromPath(path, @brush)

   ' // Set the color at the center of the path to red.
   hStatus = GdipSetPathGradientCenterColor(brush, ARGB_RED)

   ' // Set the surround color
   DIM colors(0) AS ARGB = {ARGB_BLUE}
   DIM count AS LONG = 1
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @colors(0), @count)

   ' // Set the blend shape
   hStatus = GdipSetPathGradientSigmaBlend(brush, 0.2, 0.7)

   ' // Fill an area with the path gradient brush
   hStatus = GdipFillEllipse(graphics, brush, 100, 70, 200, 100)

   ' // The color is blue on the boundary and at the center.
   ' // At points that are 20 percent of the way from the boundary to the
   ' // center, the color is 70 percent red and 30 percent blue.

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

---

## GdipSetPathGradientSurroundColorsWithCount

Sets the surround colors of this path gradient brush. The surround colors are colors specified for discrete points on the brush's boundary path.

```
FUNCTION GdipSetPathGradientSurroundColorsWithCount (BYVAL brush AS GpPathGradient PTR, _
    BYVAL colors AS CONST ARGB PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *colors* | [in] Pointer to an array of ARGB colors that specify the surround colors. |
| *count* | [in, out] Pointer to a long integer that, on input, specifies the number of **Color** objects in the *colors* array. If the method succeeds, this parameter, on output, receives the number of surround colors set. If the method fails, this parameter does not receive a value. |

#### Remarks

A path gradient brush has a boundary path and a center point. The center point is set to a single color, but you can specify different colors for several points on the boundary. For example, suppose you specify red for the center color, and you specify blue, green, and yellow for distinct points on the boundary. Then as you move along the boundary, the color will change gradually from blue to green to yellow and back to blue. As you move along a straight line from any point on the boundary to the center point, the color will change from that boundary point's color to red.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on an array of three points
' that defines a triangular path. The code also initializes an array of three Color objects.
' The call to the GdipSetPathGradientSurroundColorsWithCount function associates each color
' in the color array with the corresponding (same index) point in the point array. After the
' surround colors of the path gradient brush have been set, the GdipFillRectangle function
' uses the path gradient brush to paint a rectangle that includes the triangular path.
' One edge of the rendered triangle changes gradually from red to green. The next edge
' changes gradually from green to black, and the third edge changes gradually from black
' to red. The code does not set the center color, so the center color has the default value
' of black. As you move along a straight line from any point on the boundary path (triangle)
' to the center point, the color changes gradually from that boundary point's color to black.
' ========================================================================================
SUB Example_SetPathGradientSurroundColors (BYVAL hdc AS HDC)

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

   ' // Creates a gradient brush based on an array of points
   DIM pts(0 TO 2) AS GpPointF = {(20, 20), (100, 20), (100, 100)}
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@pts(0), 3, WrapModeClamp, @brush)

    ' // Sets the surround colors
   DIM  count AS LONG = 3
   DIM colors(0 TO 2) AS ARGB = {ARGB_RED, ARGB_BLUE, ARGB_AQUA}
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @colors(0), @count)
   
   ' // Fill an area with the path gradient brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 200, 200)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathGradientTransform

Sets the transformation matrix of this path gradient brush.

```
FUNCTION GdipSetPathGradientTransform (BYVAL brush AS GpPathGradient PTR, _
   BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *matrix* | [in] Pointer to a variable that specifies the transformation matrix. |

#### Remarks

A **PathGradientBrush** object has a **GraphicsPath** object that serves as the boundary path for the brush. When you paint with a path gradient brush, only the area inside the boundary path is filled. If the brush's transformation matrix is set to represent any transformation other than the identity, then the boundary path is transformed according to that matrix during rendering, and only the area inside the transformed path is filled.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' GdipFillRectangle function uses the path gradient brush to paint a rectangle that
' contains the triangular path. Next, the code creates a Matrix object that represents a
' composite transformation (rotate, then translate) and passes the address of that Matrix
' object to the GdipSetPathGradientTransform function. The code calls GdipFillRectangle a
' second time to paint the same rectangle using the transformed path gradient brush.
' ========================================================================================
SUB Example_SetPathGradientTransform (BYVAL hdc AS HDC)

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

   ' // Creates a gradient brush based on an array of points
   DIM pts(0 TO 2) AS GpPointF = {(0, 0), (100, 0), (100, 100)}
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@pts(0), 3, WrapModeClamp, @brush)
   
    ' // Sets the surround colors
   DIM  count AS LONG = 3
   DIM colors(0 TO 2) AS ARGB = {ARGB_RED, ARGB_LIGHTGREEN, ARGB_BLACK}
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @colors(0), @count)
   
   ' // Fill an area with the path gradient brush (no transformation).
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 200, 200)

   ' // Set the transformation for the brush (rotate, then translate).
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(0, 1, -1, 0, 150, 60, @matrix)
   hStatus = GdipSetPathGradientTransform(brush, matrix)

   ' // Fill the same area with the transformed path gradient brush.
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 200, 200)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathGradientWrapMode

Sets the wrap mode of this path gradient brush.

```
FUNCTION GdipSetPathGradientWrapMode (BYVAL brush AS GpPathGradient PTR, _
   BYVAL wrapMode AS GpWrapMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *wrapMode* | [in] Element of the WrapMode enumeration that specifies how areas painted with the path gradient brush will be tiled. The default value is *WrapModeClamp*. |

#### Remarks

The bounding rectangle of a path gradient brush is the smallest rectangle that encloses the brush's boundary path. When you paint the bounding rectangle with the path gradient brush, only the area inside the boundary path gets filled. The area inside the bounding rectangle but outside the boundary path does not get filled.

*WrapModeClamp* (the default wrap mode) indicates that no painting occurs outside of the brush's bounding rectangle. All of the other wrap modes indicate that areas outside the brush's bounding rectangle will be tiled. Each tile is a copy (possibly flipped) of the filled path inside its bounding rectangle.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' code calls the PathGradientBrush::SetWrapMode method of the PathGradientBrush object to
' set the brush's wrap mode to WrapModeTileFlipX. The Graphics::FillRectangle method uses
' the path gradient brush to tile a large area.
' The output of the code is a grid of tiles. As you move from one tile to the next in a
' given row, the image (filled boundary path inside the bounding rectangle) is flipped
' horizontally.
' ========================================================================================
SUB Example_SetPathGradientWrapMode (BYVAL hdc AS HDC)

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

   ' // Creates a gradient brush based on an array of points
   DIM points(0 TO 2) AS GpPointF = {(0, 0), (100, 0), (100, 100)}
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@points(0), 3, WrapModeClamp, @brush)

   ' // Sets the surround colors
   DIM  count AS LONG = 3
   DIM colors(0 TO 2) AS ARGB = {ARGB_RED, ARGB_BLUE, ARGB_LIGHTGREEN}
   hStatus = GdipSetPathGradientSurroundColorsWithCount(brush, @colors(0), @count)
   
   ' // Sets the wrap mode
   hStatus = GdipSetPathGradientWrapMode(brush, WrapModeTileFlipX)

   ' // Fills a rectangle with the brush
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 800, 300)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTranslatePathGradientTransform

Updates this brush's current transformation matrix with the product of itself and a translation matrix.

```
FUNCTION GdipTranslatePathGradientTransform (BYVAL brush AS GpPathGradient PTR, _
   BYVAL dx AS REAL, BYVAL dy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **PathGradientBrush** object. |
| *dx* | [in] Single precision value that specifies the horizontal component of the translation. |
| *dy* | [in] Single precision value that specifies the vertical component of the translation. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the translation matrix is on the left, and *MatrixOrderAppend* specifies that the translation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Remarks

A single 3 3 matrix can store any sequence of affine transformations. If you have several 3 3 matrices, each of which represents an affine transformation, the product of those matrices is a single 3 3 matrix that represents the entire sequence of transformations. The transformation represented by that product is called a composite transformation. For example, suppose matrix S represents a scaling and matrix T represents a translation. If matrix M is the product ST, then matrix M represents a composite transformation: first scale, then translate.

#### Example

```
' ========================================================================================
' The following example creates a PathGradientBrush object based on a triangular path. The
' calls to the GdipScalePathGradientTransform and GdipTranslatePathGradientTransform
' methods of the PathGradientBrush object set the elements of the brush's transformation
' matrix so that it represents a composite transformation: first scale, then translate.
' The code uses the path gradient brush twice to paint a rectangle: once before the
' transformation is set and once after the transformation is set.
' ========================================================================================
SUB Example_TranslatePathGradientTransform (BYVAL hdc AS HDC)

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

   DIM pts(0 TO 2) AS GpPointF = {(0, 0), (50, 0), (50, 50)}
   DIM brush AS GpPathGradient PTR
   hStatus = GdipCreatePathGradient(@pts(0), 3, WrapModeClamp, @brush)

   ' // Fill an area with the path gradient brush (no transformation).
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 500, 500)

   hStatus = GdipScalePathGradientTransform(brush, 3, 1, MatrixOrderPrepend)
   hStatus = GdipTranslatePathGradientTransform(brush, 100, 50, MatrixOrderAppend)

   ' // Fill the same area with the transformed path gradient brush.
   hStatus = GdipFillRectangle(graphics, brush, 0, 0, 500, 500)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
