# GraphicsPath Functions

A **GraphicsPath** object stores a sequence of lines, curves, and shapes. You can draw the entire sequence by calling the **GdipDrawPath** function. You can partition the sequence of lines, curves, and shapes into figures, and with the help of the **PathIterator** functions, you can draw selected figures. You can also place markers in the sequence, so that you can draw selected portions of the path.

| Name       | Description |
| ---------- | ----------- |
| [GdipAddPathArc](#gdipaddpatharc) | Adds an elliptical arc to the current figure of this path. |
| [GdipAddPathArcI](#gdipaddpatharc) | Adds an elliptical arc to the current figure of this path. |
| [GdipAddPathBezier](#gdipaddpathbezier) | Adds a Bézier spline to the current figure of this path. |
| [GdipAddPathBezierI](#gdipaddpathbezier) | Adds a Bézier spline to the current figure of this path. |
| [GdipAddPathBeziers](#gdipaddpathbeziers) | Adds a Bézier spline to the current figure of this path. |
| [GdipAddPathBeziersI](#gdipaddpathbeziers) | Adds a Bézier spline to the current figure of this path. |
| [GdipAddPathClosedCurve](#gdipaddpathclosedcurve) | Adds a closed cardinal spline to this path. |
| [GdipAddPathClosedCurveI](#gdipaddpathclosedcurve) | Adds a closed cardinal spline to this path. |
| [GdipAddPathClosedCurve2](#gdipaddpathclosedcurve2) | Adds a closed cardinal spline to this path. |
| [GdipAddPathClosedCurve2I](#gdipaddpathclosedcurve2) | Adds a closed cardinal spline to this path. |
| [GdipAddPathCurve](#gdipaddpathcurve) | Adds a cardinal spline to the current figure of this path. |
| [GdipAddPathCurveI](#gdipaddpathcurve) | Adds a cardinal spline to the current figure of this path. |
| [GdipAddPathCurve2](#gdipaddpathcurve2) | Adds a cardinal spline to the current figure of this path. |
| [GdipAddPathCurve2I](#gdipaddpathcurve2) | Adds a cardinal spline to the current figure of this path. |
| [GdipAddPathCurve3](#gdipaddpathcurve3) | Adds a cardinal spline to the current figure of this path. |
| [GdipAddPathCurve3I](#gdipaddpathcurve3) | Adds a cardinal spline to the current figure of this path. |
| [GdipAddPathEllipse](#gdipaddpathellipse) | Adds an ellipse to this path. |
| [GdipAddPathEllipseI](#gdipaddpathellipse) | Adds an ellipse to this path. |
| [GdipAddPathLine](#gdipaddpathline) | Adds a line to the current figure of this path. |
| [GdipAddPathLineI](#gdipaddpathline) | Adds a line to the current figure of this path. |
| [GdipAddPathLine2](#gdipaddpathline2) | Adds a line to the current figure of this path. |
| [GdipAddPathLine2I](#gdipaddpathline2) | Adds a line to the current figure of this path. |
| [GdipAddPathPath](#gdipaddpathpath) | Adds a path to this path. |
| [GdipAddPathPie](#gdipaddpathpie) | Adds a pie to this path. |
| [GdipAddPathPieI](#gdipaddpathpie) | Adds a pie to this path. |
| [GdipAddPathPolygon](#gdipaddpathpolygon) | Adds a polygon to this path. |
| [GdipAddPathPolygonI](#gdipaddpathpolygon) | Adds a polygon to this path. |
| [GdipAddPathRectangle](#gdipaddpathrectangle) | Adds a rectangle to this path. |
| [GdipAddPathRectangleI](#gdipaddpathrectangle) | Adds a rectangle to this path. |
| [GdipAddPathRectangles](#gdipaddpathrectangles) | Adds a sequence of rectangles to this path. |
| [GdipAddPathRectanglesI](#gdipaddpathrectangles) | Adds a sequence of rectangles to this path. |
| [GdipAddPathString](#gdipaddpathstring) | Adds the outline of a string to this path. |
| [GdipAddPathStringI](#gdipaddpathstring) | Adds the outline of a string to this path. |
| [GdipClearPathMarkers](#gdipclearpathmarkers) | Clears the markers from this path. |
| [GdipClonePath](#gdipclonepath) | Creates a new **GraphicsPath** object, and initializes it with the contents of this **GraphicsPath** object. |
| [GdipClosePathFigure](#gdipclosepathfigure) | Closes the current figure of this path. |
| [GdipClosePathFigures](#gdipclosepathfigures) | Closes all open figures in this path. |
| [GdipCreatePath](#gdipcreatepath) | Creates a **GraphicsPath** object and initializes the fill mode. |
| [GdipCreatePath2](#gdipcreatepath2) | Creates a **GraphicsPath** object and initializes the fill mode. |
| [GdipCreatePath2I](#gdipcreatepath2) | Creates a **GraphicsPath** object and initializes the fill mode. |
| [GdipDeletePath](#gdipdeletepath) | Deletes the specified **GraphicsPath** object. |
| [GdipFlattenPath](#gdipflattenpath) | Applies a transformation to this path and converts each curve in the path to a sequence of connected lines. |
| [GdipGetPathData](#gdipgetpathdata) | Gets an array of points and an array of point types from this path. Together, these two arrays define the lines, curves, figures, and markers of this path. |
| [GdipGetPathFillMode](#gdipgetpathfillmode) | Sets the fill mode of this path. |
| [GdipGetPathLastPoint](#gdipgetpathlastpoint) | Gets the ending point of the last figure in this path. |
| [GdipGetPathPoints](#gdipgetpathpoints) | Gets this path's array of points. The array contains the endpoints and control points of the lines and Bézier splines that are used to draw the path. |
| [GdipGetPathPointsI](#gdipgetpathpoints) | Gets this path's array of points. The array contains the endpoints and control points of the lines and Bézier splines that are used to draw the path. |
| [GdipGetPathTypes](#gdipgetpathtypes) | Gets this path's array of points. The array contains the endpoints and control points of the lines and Bézier splines that are used to draw the path. |
| [GdipGetPathWorldBounds](#gdipgetpathworldbounds) | Gets a bounding rectangle for this path. |
| [GdipGetPathWorldBoundsI](#gdipgetpathworldbounds) | Gets a bounding rectangle for this path. |
| [GdipGetPointCount](#gdipgetpointcount) | Gets the number of points in this path's array of data points. This is the same as the number of types in the path's array of point types. |
| [GdipIsOutlineVisiblePathPoint](#gdipisoutlinevisiblepathpoint) | Determines whether a specified point touches the outline of this path when the path is drawn by a specified **Graphics** object and a specified pen. |
| [GdipIsOutlineVisiblePathPointI](#gdipisoutlinevisiblepathpoint) | Determines whether a specified point touches the outline of this path when the path is drawn by a specified **Graphics** object and a specified pen. |
| [GdipIsVisiblePathPoint](#gdipisvisiblepathpoint) | Determines whether a specified point lies in the area that is filled when this path is filled by a specified **Graphics** object. |
| [GdipIsVisiblePathPointI](#gdipisvisiblepathpoint) | Determines whether a specified point lies in the area that is filled when this path is filled by a specified **Graphics** object. |
| [GdipResetPath](#gdipresetpath) | Empties the path and sets the fill mode to *FillModeAlternate*. |
| [GdipReversePath](#gdipreversepath) | Reverses the order of the points that define this path's lines and curves. |
| [GdipSetPathFillMode](#gdipsetpathfillmode) | Sets the fill mode of this path. |
| [GdipSetPathMarker](#gdipsetpathmarker) | Designates the last point in this path as a marker point. |
| [GdipStartPathFigure](#gdipsetpathfigure) | Starts a new figure without closing the current figure. Subsequent points added to this path are added to the new figure. |
| [GdipTransformPath](#gdiptransformpath) | Multiplies each of this path's data points by a specified matrix. |
| [GdipWarpPath](#gdipwarppath) | Applies a warp transformation to this path. This function also flattens (converts to a sequence of straight lines) the path. |
| [GdipWidenPath](#gdipwidenpath) | Replaces this path with curves that enclose the area that is filled when this path is drawn by a specified pen. This function also flattens the path. |
| [GdipWindingModeOutline](#gdipwindingmodeoutline) | Transforms and flattens this path, and then converts this path's data points so that they represent only the outline of the path. |

---

## GdipAddPathArc

Adds an elliptical arc to the current figure of this path.

```
FUNCTION GdipAddPathArc (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, _
   BYVAL height AS REAL, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
```
```
FUNCTION GdipAddPathArcI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, _
   BYVAL height AS INT_, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *x* | [in] The x-coordinate of the upper-left corner of the bounding rectangle for the ellipse that contains the arc. |
| *y* | [in] The y-coordinate of the upper-left corner of the bounding rectangle for the ellipse that contains the arc. |
| *width* | [in] The width of the bounding rectangle for the ellipse that contains the arc. |
| *height* | [in] The height of the bounding rectangle for the ellipse that contains the arc. |
| *startAngle* | [in] The clockwise angle, in degrees, between the horizontal axis of the ellipse and the starting point of the arc. |
| *sweepAngle* | [in] The clockwise angle, in degrees, between the starting point (startAngle) and the ending point of the arc. |

#### Example

```
' ========================================================================================
' The following example draws a 90-degree arc.
' ========================================================================================
SUB Example_DrawArc (BYVAL hdc AS HDC)

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
      ' // Adds an elliptical arc to the current figure of this path.
   hStatus = GdipAddPathArc(path, 140, 15, 100, 150, 0, 180)
   ' // Closes the current figure of the path.
   hStatus = GdipClosePathFigure(path)

   ' // Create a red pen
   DIM  pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitWorld, @pen)
   ' // Draw the path.
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathBezier

Adds a Bézier spline to the current figure of this path.

```
FUNCTION GdipAddPathBezier (BYVAL path AS GpPath PTR, BYVAL x1 AS REAL, BYVAL y1 AS REAL, BYVAL x2 AS REAL, _
   BYVAL y2 AS REAL, BYVAL x4 AS REAL, BYVAL y3 AS REAL, BYVAL x4 AS REAL, BYVAL y4 AS REAL) AS GpStatus
```
```
FUNCTION GdipAddPathBezierI (BYVAL path AS GpPath PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, _
   BYVAL y2 AS INT_, BYVAL x4 AS INT_, BYVAL y3 AS INT_, BYVAL x4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *x1* | [in] The x-coordinate of the starting point of the Bézier spline. |
| *y1* | [in] The y-coordinate of the starting point of the Bézier spline. |
| *x2* | [in] The x-coordinate of the first control point of the Bézier spline. |
| *y2* | [in] The y-coordinate of the first control point of the Bézier spline. |
| *x3* | [in] The x-coordinate of the second control point of the Bézier spline. |
| *y3* | [in] The y-coordinate of the second control point of the Bézier spline. |
| *x4* | [in] The x-coordinate of the ending point of the Bézier spline. |
| *y4* | [in] The y-coordinate of the ending point of the Bézier spline. |

#### Example

```
' ========================================================================================
' This example draws a Bézier curve using GdipAddPathBezier.
' ========================================================================================
SUB Example_DrawBezier (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add Bézier curve: start ? control1 ? control2 ? end
   hStatus = GdipAddPathBezier(path, 50, 150, 100, 50, 200, 250, 250, 150)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathBeziers

Adds a Bézier spline to the current figure of this path.

```
FUNCTION GdipAddPathBeziers (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, _
   BYVAL count AS INT_) AS GpStatus
```
```
FUNCTION GdipAddPathBeziersI (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *points* | [in] Pointer to an array of starting points, ending points, and control points for the connected splines. The first spline is constructed from the first point through the fourth point in the array and uses the second and third points as control points. Each subsequent spline in the sequence needs exactly three more points: the ending point of the previous spline is used as the starting point, the next two points in the sequence are control points, and the third point is the ending point. |
| *count* | [in] Long integer value that specifies the number of elements in the points array. |

#### Example

```
' ========================================================================================
' This example draws a sequence of Bézier curves using GdipAddPathBeziers.
' ========================================================================================
SUB Example_DrawBeziers (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Define points for two connected Bézier curves
   DIM pts(0 TO 7) AS GpPointF
   pts(0).x = 30  : pts(0).y = 150   ' Start point
   pts(1).x = 80  : pts(1).y = 50    ' Control 1
   pts(2).x = 130 : pts(2).y = 250   ' Control 2
   pts(3).x = 180 : pts(3).y = 150   ' End of first curve

   pts(4).x = 230 : pts(4).y = 50    ' Control 1 of second
   pts(5).x = 280 : pts(5).y = 250   ' Control 2 of second
   pts(6).x = 330 : pts(6).y = 150   ' End of second curve

   ' // Add connected Bézier curves
   hStatus = GdipAddPathBeziers(path, @pts(0), 7)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_GREEN, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathClosedCurve

Adds a closed cardinal spline to the path.

```
FUNCTION GdipAddPathClosedCurve (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, _
   BYVAL count AS INT_) AS GpStatus
```
```
FUNCTION GdipAddPathClosedCurveI (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *points* | [in] Pointer to an array of points that define the cardinal spline. The cardinal spline is a curve that passes through each point in the array. |
| *count* | [in] Long integer value that specifies the number of elements in the points array. |

#### Remarks

You should keep a copy of the points array if those points will be needed later. The **GraphicsPath** object does not store the points passed to the **GdipAddPathClosedCurve** function; instead, it converts the cardinal spline to a sequence of Bézier splines and stores the points that define those Bézier splines. You cannot retrieve the original array of points from the **GraphicsPath** object.

#### Example

```
' ========================================================================================
' This example draws a closed curve that loops through a set of points.
' ========================================================================================
SUB Example_DrawClosedCurve (BYVAL hdc AS HDC)

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

   ' // Define points for the closed curve
   DIM pts(0 TO 4) AS GpPointF
   pts(0).x = 60  : pts(0).y = 100
   pts(1).x = 120 : pts(1).y = 40
   pts(2).x = 180 : pts(2).y = 100
   pts(3).x = 150 : pts(3).y = 180
   pts(4).x = 90  : pts(4).y = 180

   ' // Add closed curve to path
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 5)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_PURPLE, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)


   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathClosedCurve2

Adds a closed cardinal spline to this path.

```
FUNCTION GdipAddPathClosedCurve2 (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, _
   BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
```
```
FUNCTION GdipAddPathClosedCurve2I (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, _
   BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *points* | [in] Pointer to an array of points that define the cardinal spline. The cardinal spline is a curve that passes through each point in the array. |
| *count* | [in] Long integer value that specifies the number of elements in the *points* array. |
| *tension* | [in] Nonnegative simple precision value that controls the length of the curve and how the curve bends. A value of 0 specifies that the spline is a sequence of straight line segments. As the value increases, the curve becomes fuller. |

#### Remarks

You should keep a copy of the points array if those points will be needed later. The **GraphicsPath** object does not store the points passed to the **GdipAddPathClosedCurve2** function; instead, it converts the cardinal spline to a sequence of Bézier splines and stores the points that define those Bézier splines. You cannot retrieve the original array of points from the **GraphicsPath** object.

#### Example

```
' ========================================================================================
' This example draws a closed cardinal spline with adjustable tension.
' ========================================================================================
SUB Example_DrawClosedCurve2 (BYVAL hdc AS HDC)

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

   ' // Define points for the closed curve
   DIM pts(0 TO 4) AS GpPointF
   pts(0).x = 60  : pts(0).y = 100
   pts(1).x = 120 : pts(1).y = 40
   pts(2).x = 180 : pts(2).y = 100
   pts(3).x = 150 : pts(3).y = 180
   pts(4).x = 90  : pts(4).y = 180

   ' // Add closed curve with tension
   DIM tension AS SINGLE = 0.5
   hStatus = GdipAddPathClosedCurve2(path, @pts(0), 5, tension)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_ORANGE, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathCurve

Adds a cardinal spline to the current figure of this path.

```
FUNCTION GdipAddPathCurve (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, _
   BYVAL count AS INT_) AS GpStatus
```
```
FUNCTION GdipAddPathCurveI (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *points* | [in] Pointer to an array of points that define the cardinal spline. The cardinal spline is a curve that passes through each point in the array. |
| *count* | [in] Long integer value that specifies the number of elements in the points array. |

#### Remarks

You should keep a copy of the points array if those points will be needed later. The **GraphicsPath** object does not store the points passed to the **GdipAddPathCurve** function; instead, it converts the cardinal spline to a sequence of Bézier splines and stores the points that define those Bézier splines. You cannot retrieve the original array of points from the **GraphicsPath** object.

#### Example

```
' ========================================================================================
' This example draws a cardinal spline using GdipAddPathCurve.
' ========================================================================================
SUB Example_DrawCurve (BYVAL hdc AS HDC)

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

   ' // Define points for the curve
   DIM pts(0 TO 4) AS GpPointF
   pts(0).x = 60  : pts(0).y = 100
   pts(1).x = 120 : pts(1).y = 40
   pts(2).x = 180 : pts(2).y = 100
   pts(3).x = 150 : pts(3).y = 180
   pts(4).x = 90  : pts(4).y = 180

   ' // Add cardinal spline to path
   hStatus = GdipAddPathCurve(path, @pts(0), 5)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_TEAL, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathCurve2

Adds a cardinal spline to the current figure of this path.

```
FUNCTION GdipAddPathCurve2 (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, _
   BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
```
```
FUNCTION GdipAddPathCurve2I (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, _
   BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *points* | [in] Pointer to an array of points that define the cardinal spline. The cardinal spline is a curve that passes through each point in the array. |
| *count* | [in] Long integer value that specifies the number of elements in the points array. |
| *tension* | [in] Nonnegative simple precision value that controls the length of the curve and how the curve bends. A value of 0 specifies that the spline is a sequence of straight line segments. As the value increases, the curve becomes fuller. |

#### Remarks

You should keep a copy of the points array if those points will be needed later. The **GraphicsPath** object does not store the points passed to the **GdipAddPathCurve** function; instead, it converts the cardinal spline to a sequence of Bézier splines and stores the points that define those Bézier splines. You cannot retrieve the original array of points from the **GraphicsPath** object.

#### Example

```
' ========================================================================================
' This example draws a cardinal spline with a specified tension using GdipAddPathCurve2.
' ========================================================================================
SUB Example_DrawCurve2 (BYVAL hdc AS HDC)

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

   ' // Define points for the curve
   DIM pts(0 TO 4) AS GpPointF
   pts(0).x = 60  : pts(0).y = 100
   pts(1).x = 120 : pts(1).y = 40
   pts(2).x = 180 : pts(2).y = 100
   pts(3).x = 150 : pts(3).y = 180
   pts(4).x = 90  : pts(4).y = 180

   ' // Add cardinal spline with tension
   DIM tension AS SINGLE = 0.8
   hStatus = GdipAddPathCurve2(path, @pts(0), 5, tension)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_NAVY, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathCurve3

Adds a cardinal spline to the current figure of this path.

```
FUNCTION GdipAddPathCurve3 (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, _
   BYVAL count AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS REAL) AS GpStatus
```
```
FUNCTION GdipAddPathCurve3I (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, _
   BYVAL count AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *points* | [in] Pointer to an array of points that define the cardinal spline. The cardinal spline is a curve that passes through each point in the array. |
| *count* | [in] Long integer value that specifies the number of elements in the points array. |
| *offset* | [in] Long integer value that specifies the index of the array element that is used as the first point of the cardinal spline. |
| *numberOfSegments* | [in] Long integer value that specifies the number of segments in the cardinal spline. Segments are the curves that connect consecutive points in the array. |
| *tension* | [in] Nonnegative simple precision value that controls the length of the curve and how the curve bends. A value of 0 specifies that the spline is a sequence of straight line segments. As the value increases, the curve becomes fuller. |

#### Remarks

You should keep a copy of the points array if those points will be needed later. The **GraphicsPath** object does not store the points passed to the **GdipAddPathCurve** function; instead, it converts the cardinal spline to a sequence of Bézier splines and stores the points that define those Bézier splines. You cannot retrieve the original array of points from the **GraphicsPath** object.

#### Example

```
' ========================================================================================
' This example draws a cardinal spline using GdipAddPathCurve3 with per-point tension.
' ========================================================================================
SUB Example_DrawCurve3 (BYVAL hdc AS HDC)

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

   ' // Define points for the curve
  ' // Define a larger array of points
   DIM pts(0 TO 6) AS GpPointF
   pts(0).x = 40  : pts(0).y = 120
   pts(1).x = 80  : pts(1).y = 60
   pts(2).x = 120 : pts(2).y = 140
   pts(3).x = 160 : pts(3).y = 80
   pts(4).x = 200 : pts(4).y = 160
   pts(5).x = 240 : pts(5).y = 100
   pts(6).x = 280 : pts(6).y = 180

 ' // Add curve starting at offset 2, using 5 points, with tension 0.6
   DIM offset AS LONG = 2
   DIM segmentCount AS LONG = 4
   DIM tension AS SINGLE = 0.6
   hStatus = GdipAddPathCurve3(path, @pts(0), 7, offset, segmentCount, tension)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_OLIVE, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathEllipse

Adds an ellipse to the path.

```
FUNCTION GdipAddPathEllipse (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
   BYVAL width AS REAL, BYVAL height AS REAL) AS GpStatus
```
```
FUNCTION GdipAddPathEllipseI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL width AS INT_, BYVAL height AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *x* | [in] The x-coordinate of the upper-left corner of the bounding rectangle for the ellipse. |
| *y* | [in] The y-coordinate of the upper-left corner of the bounding rectangle for the ellipse. |
| *width* | [in] The width of the bounding rectangle for the ellipse. |
| *height* | [in] The height of the bounding rectangle for the ellipse. |

#### Remarks

A **GraphicsPath** object stores an ellipse as a sequence of four connected Bézier splines. The **GraphicsPath** object does not store the upper-left corner, width, and height of the ellipse's bounding rectangle.

#### Example

```
' ========================================================================================
' This example draws an ellipse using GdipAddPathEllipse.
' ========================================================================================
SUB Example_DrawEllipse (BYVAL hdc AS HDC)

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

   ' // Add ellipse to path
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_MAGENTA, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathLine

Adds a line to the current figure of the path.

```
FUNCTION GdipAddPathLine (BYVAL path AS GpPath PTR, BYVAL x1 AS REAL, BYVAL y1 AS REAL, _
   BYVAL x2 AS REAL, BYVAL y2 AS REAL) AS GpStatus
```
```
FUNCTION GdipAddPathLineI (BYVAL path AS GpPath PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, _
   BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *x1* | [in] The x-coordinate of the starting point of the line. |
| *y1* | [in] The y-coordinate of the starting point of the line. |
| *x2* | [in] The x-coordinate of the ending point of the line. |
| *y2* | [in] The y-coordinate of the ending point of the line. |

#### Example

```
' ========================================================================================
' This example draws a single line using GdipAddPathLine.
' ========================================================================================
SUB Example_DrawLine (BYVAL hdc AS HDC)

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

   ' // Add a line from (50, 50) to (250, 150)
   hStatus = GdipAddPathLine(path, 50, 50, 250, 150)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathLine2

Adds a cardinal spline to the current figure of this path.

```
FUNCTION GdipAddPathLine2 (BYVAL path AS GpPath PTR, BYVAL points AS GpPointF PTR, _
   BYVAL count AS INT_) AS GpStatus
```
```
FUNCTION GdipAddPathLine2I (BYVAL path AS GpPath PTR, BYVAL points AS GpPoint PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *points* | [in] Pointer to an array of points that specify the starting and ending points of the lines. The first point in the array is the starting point of the first line, and the last point in the array is the ending point of the last line. Each of the other points serves as ending point for one line and starting point for the next line. |
| *count* | [in] Long integer value that specifies the number of elements in the *points* array. |

#### Example

```
' ========================================================================================
' This example draws a polyline using GdipAddPathLine2.
' ========================================================================================
SUB Example_DrawLine2 (BYVAL hdc AS HDC)

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

      ' // Define points for the polyline
   DIM pts(0 TO 4) AS GpPointF
   pts(0).x = 60  : pts(0).y = 100
   pts(1).x = 120 : pts(1).y = 40
   pts(2).x = 180 : pts(2).y = 100
   pts(3).x = 150 : pts(3).y = 180
   pts(4).x = 90  : pts(4).y = 180

   ' // Add connected lines to path
   hStatus = GdipAddPathLine2(path, @pts(0), 5)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_DARKGRAY, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathPath

Adds a path to the path. 

```
FUNCTION GdipAddPathPath (BYVAL path AS GpPath PTR, BYVAL addingPath AS GpPath PTR, _
   BYVAL connect AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *addingPath* | [in] Pointer to the path to be added. |
| *connect* | [in] Boolean value that specifies whether the first figure in the added path is part of the last figure in this path. **TRUE**: Specifies that (if possible) the first figure in the added path is part of the last figure in this path. **FALSE**: Specifies that the first figure in the added path is separate from the last figure in this path. |

#### Example

```
' ========================================================================================
' This example combines two GraphicsPath objects using GdipAddPathPath.
' ========================================================================================
SUB Example_DrawPathPath (BYVAL hdc AS HDC)

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

   ' // Create base path
   DIM basePath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @basePath)
   hStatus = GdipAddPathEllipse(basePath, 120, 70, 150, 100)

   ' // Create another path to add
   DIM subPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @subPath)
   hStatus = GdipAddPathRectangle(subPath, 150, 100, 90, 40)

   ' // Merge subPath into basePath
   DIM connect AS BOOL = FALSE  ' FALSE = separate figure
   hStatus = GdipAddPathPath(basePath, subPath, connect)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2, UnitWorld, @pen)

   ' // Draw combined path
   hStatus = GdipDrawPath(graphics, pen, basePath)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF subPath THEN GdipDeletePath(subPath)
   IF basePath THEN GdipDeletePath(basePath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathPie

Adds a pie to this path. An arc is a portion of an ellipse, and a pie is a portion of the area enclosed by an ellipse. A pie is bounded by an arc and two lines (edges) that go from the center of the ellipse to the endpoints of the arc.

```
FUNCTION GdipAddPathPie (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
   BYVAL width AS REAL, BYVAL height AS REAL, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
```
```
FUNCTION GdipAddPathPieI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL width AS INT_, BYVAL height AS INT_, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *x* | [in] The x-coordinate of the upper-left corner of the rectangle that bounds the ellipse that bounds the pie. |
| *y* | [in] The y-coordinate of the upper-left corner of the rectangle that bounds the ellipse that bounds the pie. |
| *width* | [in] The width of the rectangle that bounds the ellipse that bounds the pie. |
| *height* | [in] The height of the rectangle that bounds the ellipse that bounds the pie. |
| *startAngle* | [in] Simple precision value that specifies the clockwise angle, in degrees, between the horizontal axis of the ellipse and the starting point of the arc that defines the pie. |
| *sweepAngle* | [in] Simple precision value that specifies the clockwise angle, in degrees, between the starting and ending points of the arc that defines the pie. |

#### Example

```
' ========================================================================================
' This example draws a pie slice using GdipAddPathPie.
' ========================================================================================
SUB Example_DrawPie (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add pie slice to path
   DIM x AS SINGLE = 100
   DIM y AS SINGLE = 50
   DIM nWidth AS SINGLE = 200
   DIM nHeight AS SINGLE = 150
   DIM startAngle AS SINGLE = 30
   DIM sweepAngle AS SINGLE = 120
   hStatus = GdipAddPathPie(path, x, y, nWidth, nHeight, startAngle, sweepAngle)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_CRIMSON, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathPolygon

Adds a polygon to this path.

```
FUNCTION GdipAddPathPolygon (BYVAL path AS GpPath PTR, BYVAL points AS GpPointF PTR, _
   BYVAL count AS INT_) AS GpStatus
```
```
FUNCTION GdipAddPathPolygonI (BYVAL path AS GpPath PTR, BYVAL points AS GpPoint PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *points* | [in] Pointer to an array of points that specifies the vertices of the polygon. |
| *count* | [in] Long integer value that specifies the number of elements in the *points* array. |

#### Example

```
' ========================================================================================
' This example draws a closed polygon using GdipAddPathPolygon.
' ========================================================================================
SUB Example_DrawPolygon (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Define polygon points
   DIM pts(0 TO 4) AS GpPointF
   pts(0).x = 100 : pts(0).y = 50
   pts(1).x = 200 : pts(1).y = 80
   pts(2).x = 180 : pts(2).y = 160
   pts(3).x = 120 : pts(3).y = 160
   pts(4).x = 80  : pts(4).y = 100

   ' // Add polygon to path
   hStatus = GdipAddPathPolygon(path, @pts(0), 5)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_FORESTGREEN, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathRectangle

Adds a rectangle to the path.

```
FUNCTION GdipAddPathRectangle (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
   BYVAL width AS REAL, BYVAL height AS REAL) AS GpStatus
```
```
FUNCTION GdipAddPathRectangleI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL width AS INT_, BYVAL height AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *x* | [in] The x-coordinate of the upper-left corner of the bounding rectangle. |
| *y* | [in] The y-coordinate of the upper-left corner of the bounding rectangle. |
| *width* | [in] The width of the bounding rectangle. |
| *width* | [in] The height of the bounding rectangle. |

#### Example

```
' ========================================================================================
' This example draws a rectangle using GdipAddPathRectangle.
' ========================================================================================
SUB Example_DrawRectangle (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add rectangle to path
   hStatus = GdipAddPathRectangle(path, 150, 60, 100, 100)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_SADDLEBROWN, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathRectangles

Adds a sequence of rectangles to the path.

```
FUNCTION GdipAddPathRectangles (BYVAL path AS GpPath PTR, BYVAL rects AS GpRectF PTR, _
   BYVAL count AS INT_) AS GpStatus
```
```
FUNCTION GdipAddPathRectanglesI (BYVAL path AS GpPath PTR, BYVAL rects AS GpRect PTR, _
   BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *rects* | [in] Pointer to an array of rectangles that will be added to the path. |
| *count* | [in] Long integer value that specifies the number of elements in the rects array. |

#### Example

```
' ========================================================================================
' This example adds a sequence of rectangles to a GraphicsPath using GdipAddPathRectangles.
' ========================================================================================
SUB Example_DrawRectangles (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Define array of rectangles
   DIM rects(0 TO 2) AS GpRectF
   rects(0).x = 50  : rects(0).y = 60  : rects(0).Width = 80  : rects(0).Height = 40
   rects(1).x = 150 : rects(1).y = 60  : rects(1).Width = 80  : rects(1).Height = 40
   rects(2).x = 250 : rects(2).y = 60  : rects(2).Width = 80  : rects(2).Height = 40

   ' // Add rectangles to path
   hStatus = GdipAddPathRectangles(path, @rects(0), 3)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_INDIGO, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipAddPathString

Adds the outline of a string to this path.

```
FUNCTION GdipAddPathString (BYVAL path AS GpPath PTR, BYVAL string AS CONST WSTRING PTR, _
   BYVAL length AS INT_, BYVAL family AS CONST GpFontFamily PTR, BYVAL style AS INT_, _
   BYVAL emSize AS REAL, BYVAL layoutRect AS CONST GpRectF PTR, _
   BYVAL format AS CONST GpStringFormat PTR) AS GpStatus
```
```
FUNCTION GdipAddPathStringI (BYVAL path AS GpPath PTR, BYVAL string AS CONST WSTRING PTR, _
   BYVAL length AS INT_, BYVAL family AS CONST GpFontFamily PTR, BYVAL style AS INT_, _
   BYVAL emSize AS REAL, BYVAL layoutRect AS CONST GpRect PTR, _
   BYVAL format AS CONST GpStringFormat PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *string* | [in] Pointer to a wide-character string. |
| *length* | [in] Long integer value that specifies the number of characters to display. If the string parameter points to a NULL-terminated string, this parameter can be set to –1. |
| *family* | [in] Pointer to a **GpFontFamily** object that specifies the font family for the string. |
| *style* | [in] Long integer value that specifies the style of the typeface. This value must be an element of the **FontStyle** enumeration or the result of a bitwise OR applied to two or more of these elements. For example, *FontStyleBold OR FontStyleUnderline OR FontStyleStrikeout* sets the style as a combination of the three styles.
| *emSize* | [in] Simple precision value that specifies the em size, in world units, of the string characters. |
| *layoutRect* | [in] Reference to a **GpRect** structure that specifies, in world units, the bounding rectangle for the string. |
| *format* | [in] Pointer to a **StringFormat** object that specifies layout information (alignment, trimming, tab stops, and the like) for the string. |

#### Example

```
' ========================================================================================
' This example adds the outline of a string to a GraphicsPath using GdipAddPathString.
' ========================================================================================
SUB Example_DrawPathString (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Define font family
   DIM fontFamily AS GpFontFamily PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)

   ' // Define layout rectangle
   DIM layout AS GpRectF = (90, 90, 300, 100)

   ' // Add string to path
   DIM text AS WSTRING * 32 = "Hello, GDI+!"
   DIM style AS LONG = FontStyleRegular
   DIM emSize AS SINGLE = 36
   hStatus = GdipAddPathString(path, @text, -1, fontFamily, style, emSize, @layout, NULL)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_DARKRED, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipClearPathMarkers

Clears the markers from the path.

```
FUNCTION GdipClearPathMarkers (BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example sets a marker in a path and then clears it using GdipClearPathMarkers.
' ========================================================================================
SUB Example_ClearMarkers (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add some shapes
   hStatus = GdipAddPathRectangle(path, 50, 50, 100, 60)
   hStatus = GdipSetPathMarker(path)  ' Set a marker after the rectangle
   hStatus = GdipAddPathEllipse(path, 200, 50, 100, 60)

   ' // Clear all markers
   hStatus = GdipClearPathMarkers(path)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_DARKBLUE, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipClonePath

Creates a new **GraphicsPath** object, and initializes it with the contents of this **GraphicsPath** object.

```
FUNCTION GdipClonePath (BYVAL path AS GpPath PTR, BYVAL clonePath AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *clonePath* | [out] Pointer to the new **GraphicsPath** object.

#### Example

```
' ========================================================================================
' This example clones a GraphicsPath using GdipClonePath.
' ========================================================================================
SUB Example_ClonePath (BYVAL hdc AS HDC)

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

   ' // Create original path
   DIM originalPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @originalPath)
   hStatus = GdipAddPathEllipse(originalPath, 120, 70, 150, 100)

   ' // Clone the path
   DIM clonedPath AS GpPath PTR
   hStatus = GdipClonePath(originalPath, @clonedPath)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_DARKORANGE, 2, UnitWorld, @pen)

   ' // Draw cloned path
   hStatus = GdipDrawPath(graphics, pen, clonedPath)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF clonedPath THEN GdipDeletePath(clonedPath)
   IF originalPath THEN GdipDeletePath(originalPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipClosePathFigure

Closes the current figure of a path.

```
FUNCTION GdipClosePathFigure (BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example draws an open arc and closes it using GdipClosePathFigure.
' ========================================================================================
SUB Example_CloseFigure (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add an arc (open by default)
   hStatus = GdipAddPathArc(path, 140, 15, 100, 150, 0, 180)

   ' // Close the figure (connects end point to start point)
   hStatus = GdipClosePathFigure(path)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipClosePathFigures

Closes all open figures in a path.

```
FUNCTION GdipClosePathFigures (BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example draws two open shapes and closes them using GdipClosePathFigures.
' ========================================================================================
SUB Example_CloseFigures (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add first figure (arc)
   hStatus = GdipAddPathArc(path, 100, 50, 100, 100, 0, 180)

   ' // Start new figure
   hStatus = GdipStartPathFigure(path)

   ' // Add second figure (line)
   hStatus = GdipAddPathLine(path, 250, 50, 350, 150)

   ' // Close all open figures
   hStatus = GdipClosePathFigures(path)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_DARKGREEN, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreatePath

Creates a **GraphicsPath** object and initializes the fill mode.

```
FUNCTION GdipCreatePath (BYVAL brushMode AS GpFillMode, BYVAL path AS GpPathPTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brushMode* | [in] Element of the FillMode enumeration that specifies how areas are filled if the path intersects itself. The default value is *FillModeAlternate*. |
| *path* | [out] Pointer to a variable that receives a pointer to the new **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example creates a GraphicsPath and adds a shape to it.
' ========================================================================================
SUB Example_CreatePath (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a shape to the path
   hStatus = GdipAddPathEllipse(path, 120, 80, 150, 100)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_MIDNIGHTBLUE, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreatePath2

Creates a **GraphicsPath** object and initializes the fill mode.

```
FUNCTION GdipCreatePath2 (BYVAL points AS CONST GpPointF PTR, BYVAL types AS CONST BYTE PTR, _
   BYVAL count AS INT_, BYVAL fillMode AS GpFillMode, BYVAL path AS GpPath PTR PTR) AS GpStatus
```
```
FUNCTION GdipCreatePath2I (BYVAL points AS CONST GpPoint PTR, BYVAL types AS CONST BYTE PTR, _
   BYVAL count AS INT_, BYVAL fillMode AS GpFillMode, BYVAL path AS GpPath PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *points* | [in] Pointer to an array of points that specifies the endpoints and control points of the lines and Bezier splines that are used to draw the path. |
| *types* | [in] Pointer to an array of bytes that holds the point type and a set of flags for each point in the points array. The point type is stored in the three least significant bits, and the flags are stored in the four most significant bits. Possible point types and flags are listed in the **PathPointType** enumeration. |
| *count* | [in] Long integer value that specifies the number of elements in the points array. This is the same as the number of elements in the types array. |
| *fillMode* | [in] Element of the **FillMode** enumeration that specifies how areas are filled if the path intersects itself. The default value is *FillModeAlternate*. |
| *path* | [out] Pointer to a variable that receives a pointer to the new **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example creates a GraphicsPath using GdipCreatePath2 with custom points and types.
' ========================================================================================
SUB Example_CreatePath2 (BYVAL hdc AS HDC)

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

   ' // Define points
   DIM pts(0 TO 2) AS GpPointF
   pts(0).x = 100 : pts(0).y = 100  ' Start
   pts(1).x = 200 : pts(1).y = 100  ' Line
   pts(2).x = 200 : pts(2).y = 200  ' Line

   ' // Define types
   DIM types(0 TO 2) AS BYTE
   types(0) = PathPointTypeStart
   types(1) = PathPointTypeLine
   types(2) = PathPointTypeLine OR PathPointTypeCloseSubpath

   ' // Create path from raw data
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath2(@pts(0), @types(0), 3, FillModeAlternate, @path)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_MAROON, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeletePath

Deletes the specified **GraphicsPath** object.

```
FUNCTION GdipDeletePath (BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example creates and deletes a GraphicsPath using GdipDeletePath.
' ========================================================================================
SUB Example_DeletePath (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a shape
   hStatus = GdipAddPathEllipse(path, 100, 80, 150, 100)

   ' // Draw path
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_NAVY, 2, UnitWorld, @pen)
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)  ' <-- Frees the path object
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipFlattenPath

Applies a transformation to this path and converts each curve in the path to a sequence of connected lines.

```
FUNCTION GdipFlattenPath (BYVAL path AS GpPath PTR, BYVAL matrix AS GpMatrix PTR, _
   BYVAL flatness AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *matrix* | [in] Pointer to a Matrix object that specifies the transformation to be applied to the path's data points. The default value is NULL, which specifies that no transformation is to be applied. |
| *flatness* | [in] Simple precision value that specifies the maximum error between the path and its flattened approximation. Reducing the flatness increases the number of line segments in the approximation. The default value is *FlatnessDefault*. |

#### Example

```
' ========================================================================================
' This example flattens a path containing a Bézier curve using GdipFlattenPath.
' ========================================================================================
SUB Example_FlattenPath (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a Bézier curve
   hStatus = GdipAddPathBezier(path, 50, 150, 100, 50, 200, 250, 250, 150)

   ' // Flatten the path (convert curve to lines)
   hStatus = GdipFlattenPath(path, NULL, 0.25)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_ORANGE, 2, UnitWorld, @pen)

   ' // Draw flattened path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathData

Gets an array of points and an array of point types from this path. Together, these two arrays define the lines, curves, figures, and markers of this path.

```
FUNCTION GdipGetPathData (BYVAL path AS GpPath PTR, BYVAL pathData AS GpPathData PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *pathData* | [out] Pointer to a **GpPathData** structure that receives the path data. The **Points** data member of the **GpPathData** structure receives a pointer to an array of **GpPointF** structures that contains the path points. The **Types** data member of the **GpPathData** structure receives a pointer to an array of bytes that contains the point types. The **Count** data member of the **GpPathData** structure receives an integer that indicates the number of elements in the **Points** array. |

#### Example

```
' ========================================================================================
' This example extracts the points and types from a GraphicsPath using GdipGetPathData.
' ========================================================================================
SUB Example_GetPathData (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a shape
   hStatus = GdipAddPathRectangle(path, 100, 80, 150, 100)

   ' // Get point count
   DIM count AS LONG
   GdipGetPointCount(path, @count)

   ' // Allocate memory for points and types
   DIM pts(ANY) AS GpPointF
   DIM types(ANY) AS BYTE
   REDIM pts(count - 1)
   REDIM types(count - 1)

   ' // Fill the GpPathData structure
   DIM pathData AS GpPathData
   pathData.Count = count
   pathData.Points = @pts(0)
   pathData.Types = @types(0)

   ' // Retrieve path data
   hStatus = GdipGetPathData(path, @pathData)
   ' // Print the data
   FOR i AS LONG = 0 TO pathData.Count - 1
      OutputDebugStringW("Point " & WSTR(i) & ": (" & WSTR(pathData.Points[i].x) & ", " & WSTR(pathData.Points[i].y))
      OutputDebugStringW("Type " & WSTR(i) & ": " & WSTR(pathData.Types[i]))
   NEXT

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_ORANGE, 2, UnitWorld, @pen)

   ' // Draw the path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathFillMode

Sets the fill mode of this path.

```
FUNCTION GdipGetPathFillMode (BYVAL path AS GpPath PTR, BYVAL fillMode AS GpFillMode PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *fillMode* | [out] Pointer to a variable that receives an element of the **FillMode** enumeration that specifies how to fill areas when the path intersects itself. |

#### Example

```
' ========================================================================================
' This example retrieves the fill mode of a GraphicsPath using GdipGetPathFillMode.
' ========================================================================================
SUB Example_GetFillMode (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath with alternate fill mode
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a shape
   hStatus = GdipAddPathEllipse(path, 100, 70, 200, 100)

   ' // Retrieve fill mode
   DIM fillMode AS LONG
   hStatus = GdipGetPathFillMode(path, @fillMode)

   ' // Display result
   SELECT CASE fillMode
   CASE FillModeAlternate
       AfxMsg "Fill mode is: Alternate"
   CASE FillModeWinding
       AfxMsg "Fill mode is: Winding"
   CASE ELSE
       AfxMsg "Unknown fill mode: " & WSTR(fillMode)
   END SELECT

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_ORANGE, 2, UnitWorld, @pen)

   ' // Draw the path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathLastPoint

Gets the ending point of the last figure in this path.

```
FUNCTION GdipGetPathLastPoint (BYVAL path AS GpPath PTR, BYVAL lastPoint AS GpPointF PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *lastPoint* | [out] Pointer to a **GpPointF** structure that receives the last point. |

#### Example

```
' ========================================================================================
' This example retrieves the last point from a GraphicsPath using GdipGetPathLastPoint.
' ========================================================================================
SUB Example_GetLastPoint (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a shape
   hStatus = GdipAddPathRectangle(path, 100, 80, 150, 100)

   ' // Retrieve last point
   DIM lastPoint AS GpPointF
   hStatus = GdipGetPathLastPoint(path, @lastPoint)

   ' // Display result
   AfxMsg "Last point: ("& WSTR(lastPoint.x) & ", " & WSTR(lastPoint.y) & ")"

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_ORANGE, 2, UnitWorld, @pen)

   ' // Draw the path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathPoints

Gets this path's array of points. The array contains the endpoints and control points of the lines and Bézier splines that are used to draw the path.

```
FUNCTION GdipGetPathPoints (BYVAL path AS GpPath PTR, BYVAL points AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
```
```
FUNCTION GdipGetPathPointsI (BYVAL path AS GpPath PTR, BYVAL points AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *points* | [out] Pointer to an array of **GpPointF** structures that receives the data points. You must allocate memory for this array. You can call the **GdipGetPointCount** function to determine the required size of the array. The size, in bytes, should be the return value of **GdipGetPointCount** multiplied by *sizeof(GpPointF)*. |
| *count* | [in] Long integer value that specifies the number of elements in the points array. Set this parameter equal to the return value of the *GdipGetPointCount* function.  |

#### Example

```
' ========================================================================================
' This example extracts the points from a GraphicsPath using GdipGetPathPoints.
' ========================================================================================
SUB Example_GetPathPoints (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a shape
   hStatus = GdipAddPathRectangle(path, 100, 80, 150, 100)

   ' // Get number of points
   DIM count AS LONG
   hStatus = GdipGetPointCount(path, @count)

   ' // Allocate array for points
   DIM pts(ANY) AS GpPointF
   REDIM pts(count - 1)

   ' // Retrieve points
   hStatus = GdipGetPathPoints(path, @pts(0), count)

   ' // Display points
   FOR i AS LONG = 0 TO count - 1
       OutputDebugStringW("Point " & WSTR(i) & ": (" & WSTR(pts(i).x) & ", " & WSTR(pts(i).y) & ")")
   NEXT

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_ORANGE, 2, UnitWorld, @pen)

   ' // Draw the path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathTypes

Gets this path's array of points. The array contains the endpoints and control points of the lines and Bézier splines that are used to draw the path.

```
FUNCTION GdipGetPathTypes (BYVAL path AS GpPath PTR, BYVAL types AS BYTE PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *types* | [out] Pointer to an array that receives the point types. You must allocate memory for this array. You can call the **GdipGetPointCount** method to determine the required size of the array. |
| *count* | [in] Long integer values that specifies the number of elements in the **types** array. Set this parameter equal to the return value of the **GdipGetPointCount** function. |

#### Example

```
' ========================================================================================
' This example retrieves the types of each point in a GraphicsPath using GdipGetPathTypes.
' ========================================================================================
SUB Example_GetPathTypes (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a shape
   hStatus = GdipAddPathRectangle(path, 100, 80, 150, 100)

   ' // Get number of points
   DIM count AS LONG
   hStatus = GdipGetPointCount(path, @count)

   ' // Allocate array for types
   DIM types(ANY) AS BYTE
   REDIM types(count - 1)

   ' // Retrieve types
   hStatus = GdipGetPathTypes(path, @types(0), count)

   ' // Display points
   FOR i AS LONG = 0 TO count - 1
       OutputDebugStringW("Type " & WSTR(i) & ": (" & WSTR(types(i)) & ")")
   NEXT

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_ORANGE, 2, UnitWorld, @pen)

   ' // Draw the path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPathWorldBounds

Gets a bounding rectangle for this path.

```
FUNCTION GdipGetPathWorldBounds (BYVAL path AS GpPath PTR, BYVAL bounds AS GpRectF PTR, _
   BYVAL matrix AS CONST GpMatrix PTR, BYVAL pen AS CONST GpPen PTR) AS GpStatus
```
```
FUNCTION GdipGetPathWorldBoundsI (BYVAL path AS GpPath PTR, BYVAL bounds AS GpRect PTR, _
   BYVAL matrix AS CONST GpMatrix PTR, BYVAL pen AS CONST GpPen PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *bounds* | [out] Pointer to a **GpRectF** structure that receives the bounding rectangle. |
| *matrix* | [in] Pointer to a **Matrix** object that specifies a transformation to be applied to this path before the bounding rectangle is calculated. This path is not permanently transformed; the transformation is used only during the process of calculating the bounding rectangle. The default value is NULL. |
| *pen* | [in] Pointer to a **Pen** object that influences the size of the bounding rectangle. The bounding rectangle received in bounds will be large enough to enclose this path when the path is drawn with the pen specified by this parameter. This ensures that the path is enclosed by the bounding rectangle even if the path is drawn with a wide pen. The default value is NULL. |

#### Remarks

The rectangle returned by this method might be larger than necessary to enclose the path as drawn by the specified pen. The rectangle is calculated to allow for the pen's miter limit at sharp corners and to allow for the pen's end caps.

#### Example

```
' ========================================================================================
' This example retrieves the bounding rectangle of a GraphicsPath using GdipGetPathWorldBounds.
' ========================================================================================
SUB Example_GetWorldBounds (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a shape
   hStatus = GdipAddPathEllipse(path, 100, 80, 150, 100)

   ' // Get world bounds
   DIM bounds AS GpRectF
   hStatus = GdipGetPathWorldBounds(path, @bounds, NULL, NULL)

   ' // Display bounds
   OutputDebugStringW("Bounding box:")
   OutputDebugStringW("X = " & WSTR(bounds.x))
   OutputDebugStringW("Y = " & WSTR(bounds.y))
   OutputDebugStringW("Width = " & WSTR(bounds.Width))
   OutputDebugStringW("Height = " & WSTR(bounds.Height))

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_ORANGE, 2, UnitWorld, @pen)

   ' // Draw the path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetPointCount

Gets the number of points in this path's array of data points. This is the same as the number of types in the path's array of point types.

```
FUNCTION GdipGetPointCount (BYVAL path AS GpPath PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *count* | [out] Pointer to a variable thagt receives the number of points in the path's array of data points. |

#### Example

```
' ========================================================================================
' This example uses GdipGetPointCount to determine how many points are in a GraphicsPath.
' ========================================================================================
SUB Example_GetPointCount (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a shape
   hStatus = GdipAddPathRectangle(path, 100, 80, 150, 100)

   ' // Get number of points
   DIM count AS LONG
   hStatus = GdipGetPointCount(path, @count)

   AfxMsg "Number of points in path: " & WSTR(count)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_ORANGE, 2, UnitWorld, @pen)

   ' // Draw the path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipIsOutlineVisiblePathPoint

Determines whether a specified point touches the outline of this path when the path is drawn by a specified **Graphics** object and a specified pen.

```
FUNCTION GdipIsOutlineVisiblePathPoint (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
   BYVAL pen AS GpPen PTR, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
```
```
FUNCTION GdipIsOutlineVisiblePathPointI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL pen AS GpPen PTR, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *x* | [in] The x-coordinate of the point to be tested. |
| *y* | [in] The y-coordinate of the point to be tested. |
| *pen* | [in] Pointer to a **Pen** object. This method determines whether the test point touches the path outline that would be drawn by this pen. More points will touch an outline drawn by a wide pen than will touch an outline drawn by a narrow pen. |
| *graphics* | [in] Pointer to a **Graphics** object that specifies a world-to-device transformation. If the value of this parameter is NULL, the test is done in world coordinates; otherwise, the test is done in device coordinates. The default value is NULL. |
| *result* | [out] Pointer to a variable that receives a boolean value indicating whether test point touches the outline of this path (TRUE) or not (FALSE). |

#### Example

```
' ========================================================================================
' The following example creates an elliptical path and draws that path with a wide yellow pen.
' Then the code tests each point in an array to see whether the point touches the outline
' (as it would be drawn by the wide yellow pen) of the path. Points that touch the outline
' are painted green, and points that don't touch the outline are painted red.
' ========================================================================================
SUB Example_IsOutlineVisible (BYVAL hdc AS HDC)

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

   ' // Create the pen and brush
   DIM yellowPen AS GpPen PTR 
   hStatus = GdipCreatePen1(ARGB_YELLOW, 20, UnitWorld, @yellowPen)
   DIM brush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @brush)

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   ' // Add an ellipse
   hStatus = GdipAddPathEllipse(path, 50, 50, 200, 100)
   ' // Draw the path
   hStatus = GdipDrawPath(graphics, yellowPen, path)

   ' // Create an array of three points, and determine whether each
   ' // point in the array touches the outline of the path.
   ' // If a point touches the outline, paint it in green.
   ' // If a point does not touch the outline, paint it red.
   DIM pts(2) AS GpPointF
   pts(0).x = 230 : pts(0).y = 138
   pts(1).x = 100 : pts(1).y = 120
   pts(2).x = 150 : pts(2).y = 170

   FOR i AS LONG = 0 TO 2
      DIM result AS BOOL
      hStatus = GdipIsOutlineVisiblePathPoint(path, pts(i).x, pts(i).y, yellowPen, NULL, @result)
      IF result THEN
         hStatus = GdipSetSolidFillColor(brush, ARGB_LIGHTGREEN)
      ELSE
         hStatus = GdipSetSolidFillColor(brush, ARGB_RED)
      END IF
      hStatus = GdipFillEllipse(graphics, brush, pts(i).x - 3, pts(i).y - 3, 6, 6)
   NEXT

   ' // Cleanup
   IF yellowPen THEN GdipDeleteBrush(yellowPen)
   IF brush THEN GdipDeleteBrush(brush)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipIsVisiblePathPoint

Determines whether a specified point lies in the area that is filled when this path is filled by a specified **Graphics** object.

```
FUNCTION GdipIsVisiblePathPoint (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
    BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
```
```
FUNCTION GdipIsVisiblePathPointI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
    BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *x* | [in] The x-coordinate of the point to be tested. |
| *y* | [in] The y-coordinate of the point to be tested. |
| *graphics* | [in] Pointer to a **Graphics** object that specifies a world-to-device transformation. If the value of this parameter is NULL, the test is done in world coordinates; otherwise, the test is done in device coordinates. The default value is NULL. |
| *result* | [out] Pointer to a variable that receives a boolean value indicating whether test point lies in the interior of this path (TRUE) or not (FALSE). |

#### Example

```
' ========================================================================================
' This example checks if a point lies inside the filled area of a GraphicsPath.
' ========================================================================================
SUB Example_IsVisiblePathPoint (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathEllipse(path, 100, 80, 150, 100)

   ' // Check if point is on the path
   ' // Returns error 5 (StatusObjectBusy)
   DIM isVisible AS BOOL
   hStatus = GdipIsVisiblePathPoint(path, 150, 120, graphics, @isVisible)
   IF isVisible THEN
      AfxMsg "Point is visible"
   ELSE
      AfxMsg "Point is not visible"
   END IF

   ' // Create a Pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 5, UnitWorld, @pen)

   ' // Draw the path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
#### Example

```
' ========================================================================================
' The following example creates an elliptical path and draws that path with a wide yellow pen.
' Then the code tests each point in an array to see whether the point touches the outline
' (as it would be drawn by the wide yellow pen) of the path. Points that touch the outline
' are painted green, and points that don't touch the outline are painted red.
' ========================================================================================
SUB Example_IsVisiblePathPoint (BYVAL hdc AS HDC)

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

   ' // Create the pen and brush
   DIM yellowPen AS GpPen PTR 
   hStatus = GdipCreatePen1(ARGB_YELLOW, 20, UnitWorld, @yellowPen)
   DIM brush AS GpBrush PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @brush)

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   ' // Add an ellipse
   hStatus = GdipAddPathEllipse(path, 50, 50, 200, 100)
   ' // Draw the path
   hStatus = GdipDrawPath(graphics, yellowPen, path)

   ' // Create an array of four points, and determine whether each
   ' // point in the array touches the outline of the path.
   ' // If a point touches the outline, paint it green.
   ' // If a point does not touch the outline, paint it red.
   DIM pts(3) AS GpPointF
   pts(0).x = 50  : pts(0).y = 100
   pts(1).x = 270 : pts(1).y = 100
   pts(2).x = 150 : pts(2).y = 170
   pts(3).x = 180 : pts(3).y = 60

   FOR i AS LONG = 0 TO 3
      DIM result AS BOOL
      hStatus = GdipIsVisiblePathPoint(path, pts(i).x, pts(i).y, graphics, @result)
      IF result THEN
         hStatus = GdipSetSolidFillColor(brush, ARGB_LIGHTGREEN)
      ELSE
         hStatus = GdipSetSolidFillColor(brush, ARGB_RED)
      END IF
      hStatus = GdipFillEllipse(graphics, brush, pts(i).x - 3, pts(i).y - 3, 6, 6)
   NEXT

   ' // Cleanup
   IF yellowPen THEN GdipDeleteBrush(yellowPen)
   IF brush THEN GdipDeleteBrush(brush)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipResetPath

Empties the path and sets the fill mode to **FillModeAlternate**.

```
FUNCTION FUNCTION GdipResetPath (BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example adds shapes to a path, resets it, and adds a new shape.
' ========================================================================================
SUB Example_ResetPath (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add initial shape
   hStatus = GdipAddPathRectangle(path, 100, 80, 150, 100)

   ' // Reset the path
   hStatus = GdipResetPath(path)

   ' // Add new shape after reset
   hStatus = GdipAddPathEllipse(path, 120, 100, 100, 60)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_DARKSLATEBLUE, 2, UnitWorld, @pen)

   ' // Draw path
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipReversePath

Reverses the order of the points that define this path's lines and curves.

```
FUNCTION GdipReversePath (BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example draws a path before and after reversing it using GdipReversePath.
' ========================================================================================
SUB Example_ReversePath (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a triangle
   hStatus = GdipAddPathLine(path, 100, 100, 200, 100)
   hStatus = GdipAddPathLine(path, 200, 100, 150, 200)
   hStatus = GdipClosePathFigure(path)

   ' // Draw original path in blue
   DIM penBlue AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2, UnitWorld, @penBlue)
   hStatus = GdipDrawPath(graphics, penBlue, path)

   ' // Reverse the path
   hStatus = GdipReversePath(path)

   ' // Offset the path for visibility
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)
   hStatus = GdipTranslateMatrix(matrix, 120, 0, MatrixOrderAppend)
   hStatus = GdipTransformPath(path, matrix)
   GdipDeleteMatrix(matrix)

   ' // Draw reversed path in red
   DIM penRed AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 2, UnitWorld, @penRed)
   hStatus = GdipDrawPath(graphics, penRed, path)

   ' // Cleanup
   IF penBlue THEN GdipDeleteBrush(penBlue)
   IF penRed THEN GdipDeleteBrush(penRed)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathFillMode

Sets the fill mode of this path.

```
FUNCTION GdipSetPathFillMode (BYVAL path AS GpPath PTR, BYVAL fillMode AS GpFillMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *fillMode* | [in] Element of the **FillMode** enumeration that specifies how to fill areas when the path intersects itself. |

#### Example

```
' ========================================================================================
' This example sets the fill mode of a GraphicsPath using GdipSetPathFillMode.
' ========================================================================================
SUB Example_SetPathFillMode (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add overlapping shapes
   hStatus = GdipAddPathEllipse(path, 100, 80, 150, 100)
   hStatus = GdipAddPathEllipse(path, 130, 110, 150, 100)

   ' // Change fill mode to Winding
   hStatus = GdipSetPathFillMode(path, FillModeWinding)

   ' // Fill the path
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_LIGHTSKYBLUE, @brush)
   hStatus = GdipFillPath(graphics, brush, path)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetPathMarker

Designates the last point in this path as a marker point.

```
FUNCTION GdipSetPathMarker (BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example sets a marker after adding a rectangle, then adds an ellipse.
' ========================================================================================
SUB Example_SetPathMarker (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add first shape
   hStatus = GdipAddPathRectangle(path, 50, 50, 100, 60)

   ' // Set marker after rectangle
   hStatus = GdipSetPathMarker(path)

   ' // Add second shape
   hStatus = GdipAddPathEllipse(path, 200, 50, 100, 60)

   ' // Draw entire path
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_DARKBLUE, 2, UnitWorld, @pen)
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipStartPathFigure

Starts a new figure without closing the current figure. Subsequent points added to this path are added to the new figure.

```
FUNCTION GdipStartPathFigure (BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |

#### Example

```
' ========================================================================================
' This example creates two separate figures in a path using GdipStartPathFigure.
' ========================================================================================
SUB Example_StartPathFigure (BYVAL hdc AS HDC)

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

   ' // Create GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add first figure (rectangle)
   hStatus = GdipAddPathRectangle(path, 50, 50, 100, 60)

   ' // Start a new figure
   hStatus = GdipStartPathFigure(path)

   ' // Add second figure (ellipse)
   hStatus = GdipAddPathEllipse(path, 200, 50, 100, 60)

   ' // Draw path
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_DARKMAGENTA, 2, UnitWorld, @pen)
   hStatus = GdipDrawPath(graphics, pen, path)

   ' // Cleanup
   IF pen THEN GdipDeleteBrush(pen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTransformPath

Multiplies each of this path's data points by a specified matrix.

```
FUNCTION GdipTransformPath (BYVAL path AS GpPath PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *matrix* | [in] Pointer to a Matrix object that specifies the transformation. |

#### Example

```
' ========================================================================================
' This example moves and rotates a GraphicsPath using GdipTransformPath.
' ========================================================================================
SUB Example_TransformPath (BYVAL hdc AS HDC)

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

   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   hStatus = GdipAddPathRectangle(path, 100, 10, 200, 50)

   ' // Draw the path in blue before applying a transformation.
   DIM bluePen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1, UnitWorld, bluePen)
   hStatus = GdipDrawPath(graphics, bluePen, path)

   ' // Transform the path.
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)
   hStatus = GdipRotateMatrix(matrix, 30, MatrixOrderPrepend)
   hStatus = GdipTransformPath(path, matrix)
   DIM redPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitWorld, @redPen)
   hStatus = GdipDrawPath(graphics, redPen, path)

   ' // Cleanup
   IF bluePen THEN GdipDeletePen(bluePen)
   IF redPen THEN GdipDeletePen(redPen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipWarpPath

Applies a warp transformation to this path. This functtion also flattens (converts to a sequence of straight lines) the path.

```
FUNCTION GdipWarpPath (BYVAL path AS GpPath PTR, BYVAL matrix AS GpMatrix PTR, _
   BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL srcx AS REAL, BYVAL srcy AS REAL, _
   BYVAL srcwidth AS REAL, BYVAL srcheight AS REAL, BYVAL warpMode AS WarpMode, _
   BYVAL flatness AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *matrix* | [in] Pointer to a **Matrix** object that represents a transformation to be applied along with the warp. If this parameter is NULL, no transformation is applied. The default value is NULL. |
| *points* | [in] Pointer to an array of points that, along with the **srcRect** parameter, defines the warp transformation. |
| *count* | [in] Long integer value that specifies the number of points in the **destPoints** array. The value of this parameter must be 3 or 4. |
| *srcx* | [in] Simple precision value that specifies the x-coordinate of the upper-left corner of the  rectangle that, along with the destPoints parameter, defines the warp transformation. |
| *srcy* | [in] Simple precision value that specifies the y-coordinate of the upper-left corner of the rectangle that, along with the destPoints parameter, defines the warp transformation. |
| *srcwidth* | [in] Simple precision value that specifies the width of the rectangle that, along with the **destPoints** parameter, defines the warp transformation. |
| *srcheight* | [in] Real number that specifies the height of the rectangle that, along with the **destPoints** parameter, defines the warp transformation. |
| *warpMode* | [in] Element of the WarpMode enumeration that specifies the kind of warp to be applied. The default value is *WarpModePerspective*. |
| *flatness* | [in] Simple precision value that influences the number of line segments that are used to approximate the original path. Small values specify that many line segments are used, and large values specify that few line segments are used. The default value is *FlatnessDefault*. |

#### Remarks

A **GraphicsPath** object stores a collection of data points that represent lines and curves. The **GdipWarpPath** function converts those data points so that they represent only lines. The flatness parameter influences the number of lines that are stored. The original data points that represented curves are lost.

If the *count* parameter has a value of 4, the warp transformation is defined as shown in the following table.

| Source point  | Destination point |
| ------------- | ----------------- |
| Upper-left corner of *rcRect* | *destPoints[0]* |
| Upper-right corner of *rcRect* | *destPoints[1]* |
| Lower-left corner of *rcRect* | *destPoints[2]* |
| Lower-right corner of *rcRect* | *destPoints[3]* |

A transformation specified by a source rectangle and four destination points is capable of mapping a rectangle to an arbitrary quadrilateral that is not necessarily a parallelogram.

If the count parameter has a value of 3, the warp transformation is defined as shown in the following table.

| Source point  | Destination point |
| ------------- | ----------------- |
| Upper-left corner of *rcRect* | *destPoints[0]* |
| Upper-right corner of *rcRect* | *destPoints[1]* |
| Lower-left corner of *rcRect* | *destPoints[2]* |

A transformation specified by a source rectangle and three destination points maps rectangles to parallelograms.

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a closed figure to the path.
' The code defines a warp transformation by specifying a source rectangle and an array of
' four destination points. The source rectangle and destination points are passed to the
' GdipWarpPath function. The code draws the path twice: once before it has been warped and
' once after it has been warped.
' ========================================================================================
SUB Example_WarpPath (BYVAL hdc AS HDC)

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

   ' // Create a path
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   DIM pts(7) AS GpPointF
   pts(0).x = 20 : pts(0).y = 60
   pts(1).x = 30 : pts(1).y = 90
   pts(2).x = 15 : pts(2).y = 110
   pts(3).x = 15 : pts(3).y = 145
   pts(4).x = 55 : pts(4).y = 145
   pts(5).x = 55 : pts(5).y = 110
   pts(6).x = 40 : pts(6).y = 90
   pts(7).x = 50 : pts(7).y = 60

   hStatus = GdipAddPathLine2(path, @pts(0), 8)
   hStatus = GdipClosePathFigure(path)

   ' // Draw the path before applying a warp transformation.
   DIM bluePen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1, UnitWorld, @bluePen)
   hStatus = GdipDrawPath(graphics, bluePen, path)

   ' // Define a warp transformation, and warp the path.
   DIM srcRect AS GpRectF = (10, 50, 50, 100)

   ' // Destinations points
   DIM destPts(3) AS GpPointF
   destPts(0).x = 220 : destPts(0).y = 10
   destPts(1).x = 280 : destPts(1).y = 10
   destPts(2).x = 100 : destPts(2).y = 150
   destPts(3).x = 400 : destPts(3).y = 150

   ' // Warp the path
   hStatus = GdipWarpPath(path, NULL, @destPts(0), 4, srcRect.x, srcRect.y, srcRect.Width, srcRect.Height, WarpModePerspective, FlatnessDefault)

   ' // Draw the warped path.
   hStatus = GdipDrawPath(graphics, bluePen, path)

   ' // Draw the source rectangle and the destination polygon.
   DIM blackPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 1, UnitWorld, @blackPen)
   hStatus = GdipDrawRectangle(graphics, blackPen, srcRect.x, srcRect.y, srcRect.Width, srcRect.Height)
   hStatus = GdipDrawLine(graphics, blackPen, destpts(0).x, destPts(0).y, destPts(1).x, destPts(1).y)
   hStatus = GdipDrawLine(graphics, blackPen, destpts(0).x, destPts(0).y, destPts(2).x, destPts(2).y)
   hStatus = GdipDrawLine(graphics, blackPen, destpts(1).x, destPts(1).y, destPts(3).x, destPts(3).y)
   hStatus = GdipDrawLine(graphics, blackPen, destpts(2).x, destPts(2).y, destPts(3).x, destPts(3).y)

   ' // Cleanup
   IF bluePen THEN GdipDeletePen(bluePen)
   IF blackPen THEN GdipDeletePen(blackPen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipWidenPath

Replaces this path with curves that enclose the area that is filled when this path is drawn by a specified pen. The Widen method also flattens the path.

```
FUNCTION GdipWidenPath (BYVAL path AS GpPath PTR, BYVAL pen AS GpPen PTR, _
   BYVAL matrix AS GpMatrix PTR, BYVAL flatness AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *pen* | [in] Pointer to a **Pen** object. The path is made as wide as it would be when drawn by this pen. |
| *matrix* | [in] Pointer to a **Matrix** object that specifies a transformation to be applied along with the widening. If this parameter is NULL, no transformation is applied. The default value is NULL. |
| *flatness* | [in] Simple precision value that specifies the maximum error between the path and its flattened approximation. Reducing the flatness increases the number of line segments in the approximation. The default value is *FlatnessDefault*. |

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and adds a closed curve to the path.
' The code creates a green pen that has a width of 10, changes its style to DashStyleDash,
' and passes the address of that pen to the GdipWidenPath function. Then the code draws the
' path with a blue pen of width 1.
' ========================================================================================
SUB Example_WidenPath (BYVAL hdc AS HDC)

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

   ' // Create a path
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Create the pens.
   DIM bluePen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitWorld, @bluePen)
   DIM greenPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 10, UnitWorld, @greenPen)

   ' // Add a closed curve.
   DIM pts(3) AS GpPointF
   pts(0).x = 20  : pts(0).y = 20
   pts(1).x = 160 : pts(1).y = 100
   pts(2).x = 140 : pts(2).y = 60
   pts(3).x = 60  : pts(3).y = 100
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 4)

   ' // The GdipWidenPath function respects the dash style of the specified pen
   ' // (the green pen in this case). The following instruction changes it.
   hStatus = GdipSetPenDashStyle(greenPen, DashStyleDash)

   ' // Widen the path.
   hStatus = GdipWidenPath(path, greenPen, NULL, FlatnessDefault)

   ' // Draw the path.
   hStatus = GdipDrawPath(graphics, bluePen, path)

   ' // Cleanup
   IF bluePen THEN GdipDeletePen(bluePen)
   IF greenPen THEN GdipDeletePen(greenPen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipWindingModeOutline

Transforms and flattens this path, and then converts this path's data points so that they represent only the outline of the path.

```
FUNCTION GdipWindingModeOutline (BYVAL path AS GpPath PTR, BYVAL matrix AS GpMatrix PTR, _
    BYVAL flatness AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to the **GraphicsPath** object. |
| *matrix* | [in] Pointer to a **Matrix** object that specifies the transformation. If this parameter is NULL, no transformation is applied. The default value is NULL. |
| *flatness* | [in] Simple prevision value that specifies the maximum error between the path and its flattened approximation. Reducing the flatness increases the number of line segments in the approximation. The default value is *FlatnessDefault*. |

#### Example

```
' ========================================================================================
' The following example creates a GraphicsPath object and calls the GdiAddPathClosedCurve
' method to add a closed cardinal spline to the path. The code calls the GdipWidenPath function
' to widen the path and then draws the path. Next, the code calls the GdipWindingModeOutline
' function The code calls the GdipTranslateWorldTransform function so that the outlined path
' drawn by the subsequent call to GdipDrawPath sits to the right of the first path.
' ========================================================================================
SUB Example_WindingModeOutline (BYVAL hdc AS HDC)

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

   ' // Create a path
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Create the pens.
   DIM bluePen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 1, UnitWorld, @bluePen)
   DIM greenPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 10, UnitWorld, @greenPen)

   ' // Add a closed curve
   DIM pts(3) AS GpPointF
   pts(0).x = 20  : pts(0).y = 20
   pts(1).x = 160 : pts(1).y = 100
   pts(2).x = 140 : pts(2).y = 60
   pts(3).x = 60  : pts(3).y = 100
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 4)

   hStatus = GdipWidenPath(path, greenPen, NULL, FlatnessDefault)
   hStatus = GdipDrawPath(graphics, bluePen, path)
   hStatus = GdipWindingModeOutline(path, NULL, FlatnessDefault)
   hStatus = GdipTranslateWorldTransform(graphics, 180, 0, MatrixOrderPrepend)
   hStatus = GdipDrawPath(graphics, bluePen, path)

   ' // Cleanup
   IF bluePen THEN GdipDeletePen(bluePen)
   IF greenPen THEN GdipDeletePen(greenPen)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
