## PathIterator Functions

The **PathIterator** functions allow to isolate selected subsets of the path stored in a **GraphicsPath** object. A path consists of one or more figures. You can use a **GraphicsPathIterator** to isolate one or more of those figures. A path can also have markers that divide the path into sections. You can use a **GraphicsPathIterator** object to isolate one or more of those sections.

| Name       | Description |
| ---------- | ----------- |
| [GdipCreatePathIter](#gdipcreatepathiter) | Creates a new **GraphicsPathIterator** object and associates it with a **GraphicsPath** object. |
| [GdipDeletePathIter](#gdipdeletepathiter) | Deletes the specified **GraphicsPathIterator** object. |
| [GdipPathIterCopyData](#gdippathitercopydata) | Copies a subset of the path's data points to a **PointF** array and copies a subset of the path's point types to a **BYTE** array. |
| [GdipPathIterEnumerate](#gdippathiterenumerate) | Copies the path's data points to a **PointF** array and copies the path's point types to a **BYTE** array. |
| [GdipPathIterGetCount](#gdippathitergetcount) | Returns the number of data points in the path. |
| [GdipPathIterGetSubpathCount](#gdippathitergetsubpathcount) | Returns the number of subpaths (also called figures) in the path. |
| [GdipPathIterHasCurve](#gdippathiterhascurve) | Determines whether the path has any curves. |
| [GdipPathIterIsValid](#gdippathiterisvalid) | Checks whether a **GraphicsPathIterator** is valid. |
| [GdipPathIterNextMarker](#gdippathiternextmarker) | Gets the starting index and the ending index of the next marker-delimited section in this iterator's associated path. |
| [GdipPathIterNextMarkerPath](#gdippathiternextmarkerpath) | Gets the next marker-delimited section of this iterator's associated path. |
| [GdipPathIterNextPathType](#gdippathiternextpathtype) | Gets the starting index and the ending index of the next group of data points that all have the same type. |
| [GdipPathIterNextSubpath](#gdippathiternextsubpath) | Gets the starting index and the ending index of the next subpath (figure) in this iterator's associated path. |
| [GdipPathIterNextSubpathPath](#gdippathiternextsubpathpath) | Gets the the next figure (subpath) from this iterator's associated path. |
| [GdipPathIterRewind](#gdippathiterrewind) | Rewinds this iterator to the beginning of its associated path. |

---

## GdipCreatePathIter

Creates a new **GraphicsPathIterator** object and associates it with a **GraphicsPath** object.

```
FUNCTION GdipCreatePathIter (BYVAL iterator AS GpPathIterator PTR PTR, BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [out] Pointer to the new **GraphicsPathIterator** object. |
| *path* | [in] Pointer to a **GraphicsPath** object that will be associated with this **GraphicsPathIterator** object. |

#### Example

```
' ========================================================================================
' This example demonstrates
' How to build a path with multiple figures.
' How to use PathIterator to inspect each figure’s bounds and closure status.
' ========================================================================================
SUB Example_CreatePathIter (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with two figures
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' First figure: triangle (closed)
   hStatus = GdipAddPathLine(path, 10, 10, 100, 10)
   hStatus = GdipAddPathLine(path, 100, 10, 55, 80)
   hStatus = GdipClosePathFigure(path)

   ' Second figure: open zigzag
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 120, 10, 180, 50)
   hStatus = GdipAddPathLine(path, 180, 50, 120, 90)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Get subpath count
   DIM subpathCount AS LONG
   hStatus = GdipPathIterGetSubpathCount(iterator, @subpathCount)

   ' Create font and brush for drawing text
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' Iterate through subpaths and draw info
   DIM resultCount AS LONG, startIdx AS LONG, endIdx AS LONG
   DIM isClosed AS BOOL
   DIM yOffset AS SINGLE = 10.0

   FOR i AS LONG = 1 TO subpathCount
      hStatus = GdipPathIterNextSubpath(iterator, @resultCount, @startIdx, @endIdx, @isClosed)

      DIM info AS STRING
      info = "Subpath " & i & ": Start=" & startIdx & ", End=" & endIdx & ", Closed=" & IIF(isClosed, "True", "False")

      DIM layout AS GpRectF = (10 * rxRatio, yOffset * rxRatio, 300 * rxRatio, 20 * rxRatio)
      hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

      yOffset += 20.0
   NEXT

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

---

## GdipDeletePathIter

Deletes the specified **GraphicsPathIterator** object.

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |

#### Usage example

```
IF iterator THEN GdipDeletePathIter(iterator)
```
---

## GdipPathIterCopyData

Copies a subset of the path's data points to a **GpPointF** array and copies a subset of the path's point types to a **BYTE** array.

```
FUNCTION GdipPathIterCopyData (BYVAL iterator AS GpPathIterator PTR, _
   BYVAL resultCount AS INT_ PTR, BYVAL points AS GpPointF PTR, BYVAL types AS BYTE PTR, _
   BYVAL startIndex AS INT_, BYVAL endIndex AS INT_) AS GpStatus
```
                 
| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *resultCount* | [out] Number of points copied. This is the same as the number of types copied. |
| *points* | [out] Pointer to an array that receives a subset of the path's data points. |
| *types* | [out] Pointer to an array that receives a subset of the path's point types. |
| *startIndex* | [in] Long integer value that specifies the starting index of the points and types to be copied. |
| *endIndex* | [in] Long integer value that specifies the ending index of the points and types to be copied. |

#### Example

```
' ========================================================================================
' Creates a path with two figures.
' Uses GdipPathIterCopyData to extract the first three points and their types.
' Draws the extracted segment and displays point info in the window.
' ========================================================================================
SUB Example_PathIterCopyData (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with multiple segments
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' Add a triangle
   hStatus = GdipAddPathLine(path, 20, 20, 120, 20)
   hStatus = GdipAddPathLine(path, 120, 20, 70, 100)
   hStatus = GdipClosePathFigure(path)

   ' Add a zigzag line
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 150, 30, 200, 80)
   hStatus = GdipAddPathLine(path, 200, 80, 150, 130)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Copy a subset of the path (e.g., indices 0 to 2)
   DIM resultCount AS LONG
   DIM points(0 TO 2) AS GpPointF
   DIM types(0 TO 2) AS BYTE
   hStatus = GdipPathIterCopyData(iterator, @resultCount, @points(0), @types(0), 0, 2)

   ' Draw the copied segment using a pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2.0, UnitPixel, @pen)
   hStatus = GdipDrawLines(graphics, pen, @points(0), resultCount)

   ' Display info about copied points
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   DIM yOffset AS SINGLE = 140
   FOR i AS LONG = 0 TO resultCount - 1
      DIM info AS STRING
      info = "Point " & i & ": (" & points(i).x & ", " & points(i).y & ") Type=" & types(i)
      DIM layout AS GpRectF = (10.0, yOffset, 300.0, 20.0)
      hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)
      yOffset += 20.0
   NEXT

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF pen THEN GdipDeletePen(pen)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterEnumerate

Copies the path's data points to a **GpPointF** array and copies the path's point types to a **BYTE** array.

```
FUNCTION GdipPathIterEnumerate (BYVAL iterator AS GpPathIterator PTR,_
   BYVAL resultCount AS INT_ PTR, BYVAL points AS GpPointF PTR, _
   BYVAL types AS BYTE PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *resultCount* | [out] Number of points copied. This is the same as the number of types copied. |
| *points* | [out] Pointer to an array that receives the path's data points. |
| *types* | [out] Pointer to an array that receives the path's point types. |
| *count* | [in] Long integer value that specifies the number of elements in the points array. This is the same as the number of elements in the types array. |

#### Example

```
' ========================================================================================
' Using GdipPathIterEnumerate to Extract All Path Data.
' Builds a path with two figures.
' Uses GdipPathIterEnumerate to extract all points and types.
' Draws the path and overlays each point’s coordinates and type.
' ========================================================================================
SUB Example_PathIterEnumerate (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with mixed figures
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' Add a triangle
   hStatus = GdipAddPathLine(path, 20, 20, 120, 20)
   hStatus = GdipAddPathLine(path, 120, 20, 70, 100)
   hStatus = GdipClosePathFigure(path)

   ' Add a zigzag line
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 150, 30, 200, 80)
   hStatus = GdipAddPathLine(path, 200, 80, 150, 130)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Get total point count
   DIM totalCount AS LONG
   hStatus = GdipPathIterGetCount(iterator, @totalCount)

   ' Allocate arrays
   DIM points(0 TO totalCount - 1) AS GpPointF
   DIM types(0 TO totalCount - 1) AS BYTE

   ' Enumerate all path data
   DIM resultCount AS LONG
   hStatus = GdipPathIterEnumerate(iterator, @resultCount, @points(0), @types(0), totalCount)

   ' Draw the path
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2.0, UnitPixel, @pen)
   hStatus = GdipDrawPath(graphics, pen, path)

   ' Display point info
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(9, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   DIM yOffset AS SINGLE = 150
   FOR i AS LONG = 0 TO resultCount - 1
      DIM info AS STRING
      info = "Point " & i & ": (" & points(i).x & ", " & points(i).y & ") Type=" & types(i)
      DIM layout AS GpRectF = (10.0, yOffset, 300.0, 20.0)
      hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)
      yOffset += 20.0
   NEXT

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF pen THEN GdipDeletePen(pen)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterGetCount

Returns the number of data points in the path.

```
FUNCTION GdipPathIterGetCount (BYVAL iterator AS GpPathIterator PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *ount* | [out] Pointer to a variable that receives the number of data points in the path. |

#### Remarks

The **GraphicsPathIterator** object is associated with a **GraphicsPath object**. That **GraphicsPath** object has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points. Possible point types and flags are listed in the **PathPointType** enumeration.

#### Example

```
' ========================================================================================
' Using GdipPathIterGetCount to Inspect Point Count.
' It reflects the total number of points across all figures in the path.
' ========================================================================================
SUB Example_PathIterGetCount (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with multiple segments
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' Add a triangle
   hStatus = GdipAddPathLine(path, 20, 20, 120, 20)
   hStatus = GdipAddPathLine(path, 120, 20, 70, 100)
   hStatus = GdipClosePathFigure(path)

   ' Add a zigzag line
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 150, 30, 200, 80)
   hStatus = GdipAddPathLine(path, 200, 80, 150, 130)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Get total point count
   DIM totalCount AS LONG
   hStatus = GdipPathIterGetCount(iterator, @totalCount)

   ' Display the count in the window
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   DIM info AS STRING = "Total path points: " & totalCount
   DIM layout AS GpRectF = (10.0, 10.0, 300.0, 20.0)
   hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterGetSubpathCount

Returns the number of subpaths (also called figures) in the path.

```
FUNCTION GdipPathIterGetSubpathCount (BYVAL iterator AS GpPathIterator PTR, BYVAL count AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *ount* | [out] Pointer to a variable that receives the number of subpaths (also called figures) in the path. |

####  Example

```
' ========================================================================================
' Using GdipPathIterGetSubpathCount to Count Figures.
' Each figure in a path can be open or closed, and may represent a distinct shape.
' Knowing the count lets you iterate with GdipPathIterNextSubpath or extract each figure
' with GdipPathIterNextSubpathPath.Each figure in a path can be open or closed, and may
' represent a distinct shape.
' ========================================================================================
SUB Example_PathIterGetSubpathCount (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with two figures
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' First figure: triangle
   hStatus = GdipAddPathLine(path, 30, 30, 130, 30)
   hStatus = GdipAddPathLine(path, 130, 30, 80, 100)
   hStatus = GdipClosePathFigure(path)

   ' Second figure: open zigzag
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 160, 40, 210, 90)
   hStatus = GdipAddPathLine(path, 210, 90, 160, 140)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Get subpath count
   DIM subpathCount AS LONG
   hStatus = GdipPathIterGetSubpathCount(iterator, @subpathCount)

   ' Display the count in the window
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   DIM info AS STRING = "Subpath (figure) count: " & subpathCount
   DIM layout AS GpRectF = (10.0, 10.0, 300.0, 20.0)
   hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterHasCurve

Determines whether the path has any curves.

```
FUNCTION GdipPathIterHasCurve (BYVAL iterator AS GpPathIterator PTR, BYVAL hasCurve AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *hasCurve* | [out] Pointer to a variable that receives the number of data points in the path. |

#### Remarks

All curves in a path are stored as sequences of Bézier splines. For example, when you add an ellipse to a path, you specify the upper-left corner, the width, and the height of the ellipse's bounding rectangle. Those numbers (upper-left corner, width, and height) are not stored in the path; instead; the ellipse is converted to a sequence of four Bzier splines. The path stores the endpoints and control points of those Bézier splines.

A path stores an array of data points, each of which belongs to a line or a Bézier spline. If some of the points in the array belong to Bézier splines, then **GdipPathHasCurve returns** TRUE. If all points in the array belong to lines, then  **GdipPathHasCurve** returns FALSE.

Certain methods flatten a path, which means that all the curves in the path are converted to sequences of lines. After a path has been flattened, **GdipPathHasCurve** will always return FALSE. Flattening happens when you call the **GdipFlattenPath**, **GdipWidenPath**, or **GdipWarpPath** functions.

#### Example

```
' ========================================================================================
' Example: Using GdipPathIterHasCurve to Detect Bézier Segments.
' Paths with curves are stored as Bézier splines.
' Some operations (like flattening or widening) convert curves to straight lines.
' This check helps you decide whether to preserve or simplify the path.
' ========================================================================================
SUB Example_PathIterHasCurve (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' Add a line segment
   hStatus = GdipAddPathLine(path, 20, 20, 120, 20)

   ' Add a Bézier curve
   hStatus = GdipAddPathBezier(path, 130, 30, 160, 10, 190, 50, 220, 30)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Check for curves
   DIM hasCurve AS BOOL
   hStatus = GdipPathIterHasCurve(iterator, @hasCurve)

   ' Display result in window
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   DIM info AS STRING = "Path contains curves: " & IIF(hasCurve, "True", "False")
   DIM layout AS GpRectF = (10.0, 10.0, 300.0, 20.0)
   hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterIsValid

Checks whether a **GraphicsPathIterator** is valid, meaning it’s properly associated with a **GraphicsPath** and ready for use.

```
FUNCTION GdipPathIterIsValid (BYVAL iterator AS GpPathIterator PTR, BYVAL valid AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *valid* | [out] Pointer to a variable that receives a boolean value that indicates whether the path iterator specified by the iterator parameter is valid. |

#### Example

```
' ========================================================================================
' Example: Using GdipPathIterIsValid to Verify Iterator State.
' Checks whether a GraphicsPathIterator is valid, meaning it’s properly associated with a
' GraphicsPath and ready for use
' ========================================================================================
SUB Example_PathIterIsValid (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' Add a simple figure
   hStatus = GdipAddPathLine(path, 50, 50, 150, 50)
   hStatus = GdipAddPathLine(path, 150, 50, 100, 120)
   hStatus = GdipClosePathFigure(path)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Check if iterator is valid
   DIM isValid AS BOOL
   hStatus = GdipPathIterIsValid(iterator, @isValid)

   ' Display result in window
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   DIM info AS STRING = "PathIterator is valid: " & IIF(isValid, "True", "False")
   DIM layout AS GpRectF = (10.0, 10.0, 300.0, 20.0)
   hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterNextMarker

Gets the starting index and the ending index of the next marker-delimited section in this iterator's associated path.

```
FUNCTION GdipPathIterNextMarker (BYVAL iterator AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, _
   BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *resultCount* | [out] Number of data points in the retrieved section. If there are no more marker-delimited sections to retrieve, this method returns 0. |
| *startIndex* | [out] Pointer to an LONG that receives the starting index. |
| *endIndex* | [out] Pointer to an LONG that receives the ending index. |

#### Example

```
' ========================================================================================
' Using GdipPathIterNextMarker to Identify Marker Sections.
' Markers let you segment a path into meaningful chunks.
' GdipPathIterNextMarker helps you identify and isolate those chunks for rendering or analysis.
' ========================================================================================
SUB Example_PathIterNextMarker (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with two marker sections
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' First section: triangle
   hStatus = GdipAddPathLine(path, 30, 30, 130, 30)
   hStatus = GdipAddPathLine(path, 130, 30, 80, 100)
   hStatus = GdipClosePathFigure(path)
   hStatus = GdipSetPathMarker(path)  ' Set marker after first figure

   ' Second section: zigzag
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 160, 40, 210, 90)
   hStatus = GdipAddPathLine(path, 210, 90, 160, 140)
   hStatus = GdipSetPathMarker(path)  ' Set marker after second figure

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Create font and brush for drawing text
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' Iterate through markers
   DIM resultCount AS LONG, startIdx AS LONG, endIdx AS LONG
   DIM yOffset AS SINGLE = 10.0
   DIM markerIndex AS LONG = 1

   DO
      hStatus = GdipPathIterNextMarker(iterator, @resultCount, @startIdx, @endIdx)
      IF resultCount = 0 THEN EXIT DO

      DIM info AS STRING
      info = "Marker " & markerIndex & ": Start=" & startIdx & ", End=" & endIdx
      DIM layout AS GpRectF = (10.0, yOffset, 300.0, 20.0)
      hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

      yOffset += 20.0
      markerIndex += 1
   LOOP

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterNextMarkerPath

Gets the next marker-delimited section of this iterator's associated path.

```
FUNCTION GdipPathIterNextMarkerPath (BYVAL iterator AS GpPathIterator PTR, _
   BYVAL resultCount AS INT_ PTR, BYVAL path AS GpPath PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *resultCount* | [out] Number of data points in the retrieved section. If there are no more marker-delimited sections to retrieve, this method returns 0. |
| *path* | [out] Pointer to a **GraphicsPath** object. This method sets the data points of this **GraphicsPath** object to match the data points of the retrieved section. |

#### Example

```
' ========================================================================================
' Example: Using GdipPathIterNextMarkerPath to Extract and Draw Marker Sections.
' Lets you isolate and render each marker-defined section independently.
' Useful for animations, selective editing, or exporting segments.
' ========================================================================================
SUB Example_PathIterNextMarkerPath (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with two marker sections
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' First section: triangle
   hStatus = GdipAddPathLine(path, 30, 30, 130, 30)
   hStatus = GdipAddPathLine(path, 130, 30, 80, 100)
   hStatus = GdipClosePathFigure(path)
   hStatus = GdipSetPathMarker(path)

   ' Second section: zigzag
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 160, 40, 210, 90)
   hStatus = GdipAddPathLine(path, 210, 90, 160, 140)
   hStatus = GdipSetPathMarker(path)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Create pen for drawing
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2.0, UnitPixel, @pen)

   ' Create font and brush for text
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' Extract and draw each marker section
   DIM resultCount AS LONG
   DIM markerPath AS GpPath PTR
   DIM yOffset AS SINGLE = 170
   DIM markerIndex AS LONG = 1

   DO
      hStatus = GdipCreatePath(FillModeAlternate, @markerPath)
      hStatus = GdipPathIterNextMarkerPath(iterator, @resultCount, markerPath)
      IF resultCount = 0 THEN EXIT DO

      ' Draw the marker path
      hStatus = GdipDrawPath(graphics, pen, markerPath)

      ' Display marker info
      DIM info AS STRING = "Marker " & markerIndex & ": Points=" & resultCount
      DIM layout AS GpRectF = (10.0, yOffset, 300.0, 20.0)
      hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

      yOffset += 20.0
      markerIndex += 1

      IF markerPath THEN GdipDeletePath(markerPath)
   LOOP

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF pen THEN GdipDeletePen(pen)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterNextPathType

Gets the starting index and the ending index of the next group of data points that all have the same type.

**Note**: Apparently is not implemented and *resultCount* always returns 0.

```
FUNCTION GdipPathIterNextPathType (BYVAL iterator AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, _
   BYVAL pathType AS BYTE PTR, BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *resultCount* | [out] Number of data points in the retrieved section. If there are no more marker-delimited sections to retrieve, this method returns 0. |
| *pathType* | [out] Pointer to a BYTE that receives the point type shared by all points in the group. Possible values are *PathPointTypeLine* and *PathPointTypeBezier*, which are elements of the **PathPointType** enumeration. |
| *startIndex* | [out] Pointer to an LONG that receives the starting index of the group of points. |
| *endIndex* | [out] Pointer to an LONG that receives the ending index of the group of points. |

#### Example

```
' ========================================================================================
' Example: Using GdipPathIterNextPathType to Group Points by Type.
' Helps you understand how a path is constructed—lines vs. curves.
' Enables selective rendering, editing, or flattening based on segment type.
' ========================================================================================
SUB Example_PathIterNextPathType (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with grouped types
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' Add 3 lines (same type group)
   hStatus = GdipAddPathLine(path, 20, 20, 120, 20)
   hStatus = GdipAddPathLine(path, 120, 20, 70, 100)
   hStatus = GdipAddPathLine(path, 70, 100, 20, 20)

   ' Add Bézier curve (new type group)
   hStatus = GdipAddPathBezier(path, 150, 30, 180, 10, 210, 50, 240, 30)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Create font and brush
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' Iterate through path types
   DIM resultCount AS LONG, pathType AS BYTE
   DIM startIdx AS LONG, endIdx AS LONG
   DIM yOffset AS SINGLE = 10.0
   DIM segmentIndex AS LONG = 1

   DO
      hStatus = GdipPathIterNextPathType(iterator, @resultCount, @pathType, @startIdx, @endIdx)
      IF resultCount = 0 THEN EXIT DO

      DIM typeName AS STRING
      SELECT CASE pathType AND &H07
         CASE PathPointTypeStart
            typeName = "Start"
         CASE PathPointTypeLine
            typeName = "Line"
         CASE PathPointTypeBezier
            typeName = "Bezier"
         CASE ELSE
            typeName = "Unknown"
      END SELECT

      DIM info AS STRING
      info = "Segment " & segmentIndex & ": Type=" & typeName & ", Start=" & startIdx & ", End=" & endIdx
      DIM layout AS GpRectF = (10.0, yOffset, 400.0, 20.0)
      hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

      yOffset += 20.0
      segmentIndex += 1
   LOOP

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterNextSubpath

Gets the starting index and the ending index of the next subpath (figure) in this iterator's associated path.

```
FUNCTION GdipPathIterNextSubpath (BYVAL iterator AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, _
   BYVAL statIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR, BYVAL isClosed AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *resultCount* | [out] Number of data points in the next figure. If there are no more figures in the path, this method returns 0. |
| *statIndex* | [out] Pointer to an LONG that receives the starting index. |
| *endIndex* | [out] Pointer to an LONG that receives the ending index. |
| *isClosed* | [out] Pointer to a Boolean that receives a value that indicates whether the obtained figure is closed. If the figure is closed, the received value is TRUE; otherwise, the received value is FALSE. |

#### Example

```
' ========================================================================================
' Example: Using GdipPathIterNextSubpath to Inspect Figures.
' Lets you inspect each figure in a path individually.
' Tells you whether each figure is closed (e.g., polygon) or open (e.g., polyline).
' ========================================================================================
SUB Example_PathIterNextSubpath (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with two figures
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' First figure: triangle (closed)
   hStatus = GdipAddPathLine(path, 30, 30, 130, 30)
   hStatus = GdipAddPathLine(path, 130, 30, 80, 100)
   hStatus = GdipClosePathFigure(path)

   ' Second figure: zigzag (open)
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 160, 40, 210, 90)
   hStatus = GdipAddPathLine(path, 210, 90, 160, 140)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Create font and brush for drawing text
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' Iterate through subpaths
   DIM resultCount AS LONG, startIdx AS LONG, endIdx AS LONG
   DIM isClosed AS BOOL
   DIM yOffset AS SINGLE = 10.0
   DIM subpathIndex AS LONG = 1

   DO
      hStatus = GdipPathIterNextSubpath(iterator, @resultCount, @startIdx, @endIdx, @isClosed)
      IF resultCount = 0 THEN EXIT DO

      DIM info AS STRING
      info = "Subpath " & subpathIndex & ": Start=" & startIdx & ", End=" & endIdx & ", Closed=" & IIF(isClosed, "True", "False")
      DIM layout AS GpRectF = (10.0, yOffset, 400.0, 20.0)
      hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

      yOffset += 20.0
      subpathIndex += 1
   LOOP

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterNextSubpathPath

Gets the the next figure (subpath) from this iterator's associated path.

```
FUNCTION GdipPathIterNextSubpathPath (BYVAL iteratir AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, _
   BYVAL path AS GpPath PTR, BYVAL isClosed AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |
| *resultCount* | [out] Number of data points in the retrieved figure. If there are no more figures to retrieve, this method returns 0. |
| *path* | [out] Pointer to a **GraphicsPath** object. This method sets the data points of this **GraphicsPath** object to match the data points of the retrieved figure. |
| *isClosed* | [out] Pointer to a Boolean that receives a value that indicates whether the obtained figure is closed. If the figure is closed, the received value is TRUE; otherwise, the received value is FALSE. |

#### Example

```
' ========================================================================================
' Example: Using GdipPathIterNextSubpathPath to Extract and Draw Figures.
' Lets you isolate and render each figure independently.
' Tells you whether each figure is closed, which is useful for fill operations or hit testing.
' ========================================================================================
SUB Example_PathIterNextSubpathPath (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with two figures
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' First figure: triangle (closed)
   hStatus = GdipAddPathLine(path, 30, 30, 130, 30)
   hStatus = GdipAddPathLine(path, 130, 30, 80, 100)
   hStatus = GdipClosePathFigure(path)

   ' Second figure: zigzag (open)
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 160, 40, 210, 90)
   hStatus = GdipAddPathLine(path, 210, 90, 160, 140)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' Create pen for drawing
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2.0, UnitPixel, @pen)

   ' Create font and brush for text
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' Extract and draw each subpath
   DIM resultCount AS LONG
   DIM subpath AS GpPath PTR
   DIM isClosed AS BOOL
   DIM yOffset AS SINGLE = 170
   DIM subpathIndex AS LONG = 1

   DO
      hStatus = GdipCreatePath(FillModeAlternate, @subpath)
      hStatus = GdipPathIterNextSubpathPath(iterator, @resultCount, subpath, @isClosed)
      IF resultCount = 0 THEN EXIT DO

      ' Draw the subpath
      hStatus = GdipDrawPath(graphics, pen, subpath)

      ' Display subpath info
      DIM info AS STRING = "Subpath " & subpathIndex & ": Points=" & resultCount & ", Closed=" & IIF(isClosed, "True", "False")
      DIM layout AS GpRectF = (10.0, yOffset, 400.0, 20.0)
      hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

      yOffset += 20.0
      subpathIndex += 1

      IF subpath THEN GdipDeletePath(subpath)
   LOOP

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF pen THEN GdipDeletePen(pen)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipPathIterRewind

Rewinds this iterator to the beginning of its associated path.

```
FUNCTION GdipPathIterRewind (BYVAL iterator AS GpPathIterator PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *iterator* | [in] Pointer to the **GraphicsPathIterator** object. |

#### Example

```
' ========================================================================================
' Example: Using GdipPathIterRewind to Restart Iteration.
' Lets you reuse the same iterator for multiple passes.
' Essential when combining different iterator functions (e.g., markers first, then subpaths).
' Saves memory and avoids recreating the iterator object.
' ========================================================================================
SUB Example_PathIterRewind (BYVAL hdc AS HDC)

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

   ' Create a GraphicsPath with two figures
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' First figure: triangle
   hStatus = GdipAddPathLine(path, 30, 30, 130, 30)
   hStatus = GdipAddPathLine(path, 130, 30, 80, 100)
   hStatus = GdipClosePathFigure(path)

   ' Second figure: zigzag
   hStatus = GdipStartPathFigure(path)
   hStatus = GdipAddPathLine(path, 160, 40, 210, 90)
   hStatus = GdipAddPathLine(path, 210, 90, 160, 140)

   ' Create PathIterator
   DIM iterator AS GpPathIterator PTR
   hStatus = GdipCreatePathIter(@iterator, path)

   ' First pass: count subpaths
   DIM subpathCount AS LONG
   hStatus = GdipPathIterGetSubpathCount(iterator, @subpathCount)

   ' Display count
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   hStatus = GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   hStatus = GdipCreateSolidFill(ARGB_BLACK, @brush)

   DIM info AS STRING = "Subpath count (first pass): " & subpathCount
   DIM layout AS GpRectF = (10.0, 10.0, 300.0, 20.0)
   hStatus = GdipDrawString(graphics, info, -1, font, @layout, NULL, brush)

   ' Rewind iterator
   hStatus = GdipPathIterRewind(iterator)

   ' Second pass: iterate and display each subpath
   DIM resultCount AS LONG, startIdx AS LONG, endIdx AS LONG
   DIM isClosed AS BOOL
   DIM yOffset AS SINGLE = 40.0
   DIM subpathIndex AS LONG = 1

   DO
      hStatus = GdipPathIterNextSubpath(iterator, @resultCount, @startIdx, @endIdx, @isClosed)
      IF resultCount = 0 THEN EXIT DO

      DIM lin AS STRING
      lin = "Subpath " & subpathIndex & ": Start=" & startIdx & ", End=" & endIdx & ", Closed=" & IIF(isClosed, "True", "False")
      DIM layout AS GpRectF = (10.0, yOffset, 400.0, 20.0)
      hStatus = GdipDrawString(graphics, lin, -1, font, @layout, NULL, brush)

      yOffset += 20.0
      subpathIndex += 1
   LOOP

   ' Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF font THEN GdipDeleteFont(font)
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
   IF iterator THEN GdipDeletePathIter(iterator)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
