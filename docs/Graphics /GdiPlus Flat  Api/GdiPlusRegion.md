## Region Functions

A **Region** describes an area of the display surface. The area can be any shape. In other words, the boundary of the area can be a combination of curved and straight lines. Regions can also be created from the interiors of rectangles, paths, or a combination of these. Regions are used in clipping and hit-testing operations.

| Name       | Description |
| ---------- | ----------- |
| [GdipCloneRegion](#gdipcloneregion) | Makes a copy of this **Region** object and returns the address of the new **Region** object. |
| [GdipCombineRegionPath](#gdipcombineregionpath) | Updates this region to the portion of itself that intersects the specified path's interior. |
| [GdipCombineRegionRect](#gdipcombineregionrect) | Updates this region to the portion of itself that intersects the specified rectangle's interior. |
| [GdipCombineRegionRectI](#gdipcombineregionrect) | Updates this region to the portion of itself that intersects the specified rectangle's interior. |
| [GdipCombineRegionRegion](#gdipcombineregionregion) | Updates this region to the portion of itself that intersects another region. |
| [GdipCreateRegion](#gdipcreateregion) | Creates a region that is infinite. This is the default constructor. |
| [GdipCreateRegionHrgn](#gdipcreateregionhrgn) | Creates a region that is identical to the region that is specified by a handle to a Microsoft Windows Graphics Device Interface (GDI) region. |
| [GdipCreateRegionPath](#gdipcreateregionpath) | Creates a region that is defined by a path (a **GraphicsPath** object) and has a fill mode that is contained in the **GraphicsPath** object. |
| [GdipCreateRegionRect](#gdipcreateregionrect) | Creates a region that is defined by a rectangle. |
| [GdipCreateRegionRectI](#gdipcreateregionrect) | Creates a region that is defined by a rectangle. |
| [GdipCreateRegionRgnData](#gdipcreateregionrgndata) | Creates a region that is defined by data obtained from another region. |
| [GdipDeleteRegion](#gdipdeleteregion) | Deletes a **Region** object. |
| [GdipGetRegionBounds](#gdipgetregionbounds) | Gets a rectangle that encloses this region. |
| [GdipGetRegionBoundsI](#gdipgetregionbounds) | Gets a rectangle that encloses this region. |
| [GdipGetRegionData](#gdipgetregiondata) | Gets data that describes this region. |
| [GdipGetRegionDataSize](#gdipgetregiondatasize) | Gets the number of bytes of data that describes this region. |
| [GdipGetRegionHRgn](#gdipgetregionhrgn) | Creates a Microsoft Windows Graphics Device Interface (GDI) region from this region. |
| [GdipGetRegionScans](#gdipgetregionscans) | Gets an array of rectangles that approximate this region. The region is transformed by a specified matrix before the rectangles are calculated. |
| [GdipGetRegionScansI](#gdipgetregionscans) | Gets an array of rectangles that approximate this region. The region is transformed by a specified matrix before the rectangles are calculated. |
| [GdipGetRegionScansCount](#gdipgetregionscanscount) | Gets the number of rectangles that approximate this region. The region is transformed by a specified matrix before the rectangles are calculated. |
| [GdipIsEmptyRegion](#gdipisemptyregion) | Determines whether this region is empty. |
| [GdipIsEqualRegion](#gdipisequalregion) | Determines whether this region is equal to a specified region. |
| [GdipIsInfiniteRegion](#gdipisinfiniteregion) | Determines whether this region is infinite. |
| [GdipIsVisibleRegionPoint](#gdipisvisibleregionpoint) | Determines whether a point is inside this region. |
| [GdipIsVisibleRegionPointI](#gdipisvisibleregionpoint) | Determines whether a point is inside this region. |
| [GdipIsVisibleRegionRect](#gdipisvisibleregionrect) | Determines whether a rectangle intersects this region. |
| [GdipIsVisibleRegionRectI](#gdipisvisibleregionrect) | Determines whether a rectangle intersects this region. |
| [GdipSetEmpty](#gdipsetempty) | Updates this region to an empty region. In other words, the region occupies no space on the display device. |
| [GdipSetInfinite](#gdipsetinfinite) | Updates this region to an infinite region. |
| [GdipTransformRegion](#gdiptransformregion) | Transforms this region by multiplying each of its data points by a specified matrix. |
| [GdipTranslateRegion](#gdiptranslateregion) | Offsets this region by specified amounts in the horizontal and vertical directions. |
| [GdipTranslateRegionI](#gdiptranslateregion) | Offsets this region by specified amounts in the horizontal and vertical directions. |

---

## GdipCloneRegion

Makes a copy of a **Region** object and returns the address of the new **Region** object.

```
FUNCTION GdipCloneRegion (BYVAL region AS GpRegion PTR, BYVAL cloneRegion AS GpRegion PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *cloneRegion* | [ou] Pointer to a variable that receives a pointer to the new **Region** object. |

#### Example

```
' ========================================================================================
' The following example creates two regions, one from a rectangle and the other from a path.
' Next, the code clones the region that was created from a path and uses a solid brush to
' fill the cloned region. Then, it forms the union of the two regions and fills it.
' ========================================================================================
SUB Example_CloneRegion (BYVAL hdc AS HDC)

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

   ' // Create a Path
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a closed curve to the Path
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)

   ' // Create a Region object from a rectangle.
   DIM rcf AS GpRectF = (65, 15, 70, 45)
   DIM rectRegion AS GpRegion PTR
   hStatus = GdipCreateRegionRect(@rcf, @rectRegion)

   ' // Create a region from a curved path.
   Dim pathRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @pathRegion)

   ' // Make a copy (clone) of the curved region.
   DIM clonedRegion AS GpRegion PTR
   hStatus = GdipCloneRegion(pathRegion, @clonedRegion)

   ' // Fill the cloned region with a red brush.
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)
   hStatus = GdipFillRegion(graphics, solidBrush, clonedRegion)

   ' // Form the union of the cloned region and the rectangular region.
   hStatus = GdipCombineRegionRect(clonedRegion, @rcf, CombineModeUnion)

   ' // Fill the union of the two regions with a semitransparent blue brush.
   ' // Create brushes
   DIM alphaBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(GDIP_ARGB(128, 0, 0, 255), @alphaBrush)
   hStatus = GdipFillRegion(graphics, alphaBrush, clonedRegion)

   ' // Cleanup
   IF rectRegion THEN GdipDeleteRegion(rectRegion)
   IF pathRegion THEN GdipDeleteRegion(pathRegion)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF alphaBrush THEN GdipDeleteBrush(alphaBrush)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCombineRegionPath

Updates this region to the portion of itself that intersects the specified path's interior.

```
FUNCTION GdipCombineRegionPath (BYVAL region AS GpRegion PTR, BYVAL path AS GpPath PTR, _
    BYVAL combineMode AS CombineMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *path* | [in] Pointer to a **GraphicsPath** object that specifies the path to use to update this region. |
| *combineMode* | [in] Member of the **CombineMode** enumeration that specifies how the region and path are combined. |

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle and then uses a path to update
' the region. The updated region is the union of the rectangular region and the path.
' ========================================================================================
SUB Example_CombineRegionPath (BYVAL hdc AS HDC)

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

   ' // Create a Path
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a closed curve to the Path
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)

   ' // Create a Region object from a rectangle.
   DIM rcf AS GpRectF = (65, 15, 70, 45)
   DIM rectRegion AS GpRegion PTR
   hStatus = GdipCreateRegionRect(@rcf, @rectRegion)

   ' // Form the union of the region and the interior of a path.
   hStatus = GdipCombineRegionPath(rectRegion, path, CombineModeUnion)

   ' // Fill the union of the two regions with a red brush.
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)
   hStatus = GdipFillRegion(graphics, solidBrush, rectRegion)

   ' // Cleanup
   IF rectRegion THEN GdipDeleteRegion(rectRegion)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCombineRegionRect

Updates this region to the portion of itself that intersects the specified rectangle's interior.

```
FUNCTION GdipCombineRegionRect (BYVAL region AS GpRegion PTR, BYVAL rect AS CONST GpRectF PTR, _
   BYVAL combineMode AS CombineMode) AS GpStatus
```
```
FUNCTION GdipCombineRegionRectI (BYVAL region AS GpRegion PTR, BYVAL rect AS CONST GpRect PTR, _
   BYVAL combineMode AS CombineMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *rect* | [in] Reference to a rectangle to use to update the **Region** object. |
| *combineMode* | [in] Member of the **CombineMode** enumeration that specifies how the region and rect are combined. |

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update
' the region. The updated region is the union of the path region and a rectangle.
' ========================================================================================
SUB Example_CombineRegionRect (BYVAL hdc AS HDC)

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

   ' // Create a Path
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a closed curve to the Path
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)

   ' // Create a region from a path.
   DIM pathRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @pathRegion)

   ' // Form the union of the region and a rectangle.
   DIM rcf AS GpRectF = (65, 15, 70, 45)
   hStatus = GdipCombineRegionRect(pathRegion, @rcf, CombineModeUnion)

   ' // Fill the union of the two regions with a red brush.
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)
   hStatus = GdipFillRegion(graphics, solidBrush, pathRegion)

   ' // Cleanup
   IF pathRegion THEN GdipDeleteRegion(pathRegion)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCombineRegionRegion

Updates this region to the portion of itself that intersects another region.

```
FUNCTION GdipCombineRegionRegion (BYVAL region AS GpRegion PTR, BYVAL region2 AS GpRegion PTR, _
   BYVAL combineMode AS CombineMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *region2* | [in] Pointer to a **Region** object to use to update an existing **Region** object. |
| *combineMode* | [in] Member of the **CombineMode** enumeration that specifies how the regions are combined. |

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then uses a rectangle to update
' the region.
' ========================================================================================
SUB Example_CombineRegionRegion (BYVAL hdc AS HDC)

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

   ' // Create a Path
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add a closed curve to the Path
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)

   ' // Create a region from a path.
   DIM pathRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @pathRegion)

   ' // Create a Region from a rectangle.
   DIM rcf AS GpRectF = (65, 15, 70, 45)
   DIM rectRegion AS GpRegion PTR
   hStatus = GdipCreateRegionRect(@rcf, @rectRegion)

   ' // Exclude a rectangle region from the path region.
   hStatus = GdipCombineRegionRegion(pathRegion, rectRegion, CombineModeIntersect)

   ' // Fill the union of the two regions with a red brush.
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)
   hStatus = GdipFillRegion(graphics, solidBrush, pathRegion)

   ' // Cleanup
   IF rectRegion THEN GdipDeleteRegion(rectRegion)
   IF pathRegion THEN GdipDeleteRegion(pathRegion)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateRegion

Creates a region that is infinite.

```
FUNCTION GdipCreateRegion (BYVAL region AS GpRegion PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [out] Pointer to a variable that receives a pointer to the new created **Region** object. |

#### Usage example

```
DIM region AS GpRegion PTR
GdipCreateRegion(@region)
```
---

## GdipCreateRegionHrgn

Creates a region that is identical to the region that is specified by a handle to a Microsoft Windows Graphics Device Interface (GDI) region.

```
FUNCTION GdipCreateRegionHrgn (BYVAL hRgn AS HRGN, BYVAL region AS GpRegion PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hRgn* | [in] Handle to an existing GDI region. |
| *region* | [out] Pointer to a variable that receives a pointer to the new created **Region** object. |

#### Example

```
' ========================================================================================
' Creates a region that is identical to the region that is specified by a handle to a
' Microsoft Windows Graphics Device Interface (GDI) region.
' ========================================================================================
SUB Example_CreateRegionHrgn (BYVAL hdc AS HDC)

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

   ' // Create a simple GDI region (an ellipse)
   DIM hRgn AS HRGN
   hRgn = CreateEllipticRgn(50, 50, 200, 150)

   ' // Convert the GDI region to a GDI+ region
   DIM gdipRegion AS GpRegion PTR
   hStatus = GdipCreateRegionHrgn(hRgn, @gdipRegion)

   ' // Fill the GDI+ region with a brush
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_GREEN, @brush)
   hStatus = GdipFillRegion(graphics, brush, gdipRegion)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF gdipRegion THEN GdipDeleteRegion(gdipRegion)
   IF graphics THEN GdipDeleteGraphics(graphics)
   IF hRgn THEN DeleteObject(hRgn)

END SUB
' ========================================================================================
```
---

## GdipCreateRegionPath

Creates a region that is defined by a path (a **GraphicsPath** object) and has a fill mode that is contained in the **GraphicsPath** object.

```
FUNCTION GdipCreateRegionPath (BYVAL path AS GpPath PTR, BYVAL region AS GpRegion PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *path* | [in] Pointer to a **GraphicsPath** object that specifies the path. |
| *region* | [out] Pointer to a variable that receives a pointer to the new created **Region** object. |

#### Usage example

```
   ' // Create a path
   DIM path AS GpPath PTR
   GdipCreatePath(FillModeAlternate, @path)

   ' // Create a region from a path.
   DIM pathRegion AS GpRegion PTR
   GdipCreateRegionPath(path, @pathRegion)
```

---

## GdipCreateRegionRect

Creates a region that is defined by a rectangle.

```
FUNCTION GdipCreateRegionRect (BYVAL rect AS CONST GpRectF PTR, BYVAL region AS GpRegion PTR PTR) AS GpStatus
```
```
FUNCTION GdipCreateRegionRect (BYVAL rect AS CONST GpRect PTR, BYVAL region AS GpRegion PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rect* | [in] Reference to a rectangle. |
| *region* | [out] Pointer to a variable that receives a pointer to the new created **Region** object. |

#### Usage example

```
   ' // Create a Region object from a rectangle.
   DIM rcf AS GpRectF = (65, 15, 70, 45)
   DIM rectRegion AS GpRegion PTR
   GdipCreateRegionRect(@rcf, @rectRegion)
```

---

## GdipCreateRegionRgnData

Creates a region that is defined by data obtained from another region.

```
FUNCTION GdipCreateRegionRgnData (BYVAL regionData AS CONST BYTE PTR, _
   BYVAL size AS INT_, BYVAL region AS GpRegion PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *regionData* | [in] Pointer to an array of bytes that specifies a region. The data contained in the bytes is obtained from another region by using the **GdipGetRegionData** function. |
| *size* | [in] Long integer value that specifies the number of bytes in the *regionData* array. |
| *region* | [out] Pointer to a variable that receives a pointer to the new created **Region** object. |

#### Example

```
' ========================================================================================
' The following example recreates a region from the data of that region.
' ========================================================================================
SUB Example_CreateRegionRgnData (BYVAL hdc AS HDC)

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

   ' // Add a closed curve to the path
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)

   ' // Create a region from a path.
   DIM originalRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @originalRegion)

   ' // Get the region data size
   DIM dataSize AS UINT
   hStatus = GdipGetRegionDataSize(originalRegion, @dataSize)

   ' // Allocate buffer and get region data
   DIM regionBuffer AS STRING = SPACE(dataSize)
   DIM sizeFilled AS UINT
   hStatus = GdipGetRegionData(originalRegion, STRPTR(regionBuffer), dataSize, @sizeFilled)

   ' // Create a new region from the region data
   DIM recreatedRegion AS GpRegion PTR
   hStatus = GdipCreateRegionRgnData(STRPTR(regionBuffer), sizeFilled, @recreatedRegion)

   ' // Fill the recreated region with a brush
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @brush)
   hStatus = GdipFillRegion(graphics, brush, recreatedRegion)

   ' // Cleanup
   IF brush THEN GdipDeleteBrush(brush)
   IF recreatedRegion THEN GdipDeleteRegion(recreatedRegion)
   IF originalRegion THEN GdipDeleteRegion(originalRegion)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeleteRegion

Deletes a Region object.

```
FUNCTION GdipDeleteRegion (BYVAL region AS GpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |

#### Usage example

```
   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   GdipCreateFromHDC(hdc, @graphics)

   ' // Create a rectangular region.
   DIM rectRegion AS GpRegion PTR
   DIM rcf AS GpRectF = (65, 15, 70, 45)
   GdipCreateRegionRect(@rcf, @rectRegion)

   ' ...   
   ' ...   
   ' ...   
   
   ' // Cleanup
   IF rectRegion THEN GdipDeleteRegion(rectRegion)
   IF graphics THEN GdipDeleteGraphics(graphics)
```
---

## GdipGetRegionBounds

Gets a rectangle that encloses this region.

```
FUNCTION GdipGetRegionBounds (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, _
   BYVAL rect AS GpRectF PTR) AS GpStatus
```
```
FUNCTION GdipGetRegionBoundsI (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, _
   BYVAL rect AS GpRect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *graphics* | [in] Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region and the rectangle. |
| *rect* | [out] Pointer to a **GpRectF** structure that receives the enclosing rectangle. |

#### Remarks

The current world and page transformations of the graphics object are used to calculate the region and the rectangle as they are drawn on the display device. The rectangle returned by **GdipGetRegionBounds** is not always the smallest possible rectangle.

#### Example

```
' ========================================================================================
' The following example creates a region from a path, gets the region's enclosing rectangle,
' and then displays both.
' ========================================================================================
SUB Example_GetRegionBounds (BYVAL hdc AS HDC)

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

   ' // Add a closed curve to the path
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)

   ' // Create a region from a path.
   DIM pathRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @pathRegion)

   ' // Get the region's enclosing rectangle.
   DIM rcf AS GpRectF
   hStatus = GdipGetRegionBounds(pathRegion, graphics, @rcf)

   ' // Create a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)

   ' // Create a pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 1.0, UnitWorld, @pen)

  ' // Show the region and the enclosing rectangle.
   hStatus = GdipFillRegion(graphics, solidBrush, pathRegion)
   hStatus = GdipDrawRectangle(graphics, pen, rcf.x, rcf.y, rcf.Width, rcf.Height)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pathRegion THEN GdipDeleteRegion(pathRegion)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetRegionData

Gets data that describes this region.

```
FUNCTION GdipGetRegionData (BYVAL region AS GpRegion PTR, BYVAL buffer AS BYTE PTR, _
    BYVAL bufferSize AS UINT, BYVAL sizeFilled AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *buffer* | [out] Pointer to an array of BYTE values that receives the region data. |
| *bufferSize* | [in] Integer that specifies the size, in bytes, of the buffer array. The size of the *buffer* array can be greater than or equal to the number of bytes required to store the region data. The exact number of bytes required can be determined by calling the **GdipGetRegionDataSize** function. |
| *sizeFilled* | [out] Pointer to an UINT that receives the number of bytes of data actually received by the *buffer* array. The default value is NULL. |

#### Remarks

The **GdipGetRegionDataSize** function can be used before the **GdipGetRegionData** function to determine the number of bytes needed to store the region data. Then, you can allocate a buffer that is the correct size to store the region data and set the buffer parameter to point to the buffer.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and then gets the data that describes
' the region.
' ========================================================================================
SUB Example_GetRegionData (BYVAL hdc AS HDC)

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

   ' // Add a closed curve to the path
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)

   ' // Create a region from a path.
   DIM pathRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @pathRegion)

   ' // Get the region data.
   DIM bufferSize AS UINT
   hStatus = GdipGetRegionDataSize(pathRegion, @bufferSize)
   DIM sizeFilled AS UINT
   DIM strBuffer AS STRING = SPACE(bufferSize)
   hStatus = GdipGetRegionData(pathRegion, STRPTR(strBuffer), bufferSize, @sizeFilled)
   'PRINT strBuffer

   ' // Fill the region with a semitransparent blue brush.
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)
   hStatus = GdipFillRegion(graphics, solidBrush, pathRegion)

   ' // Cleanup
   IF pathRegion THEN GdipDeleteRegion(pathRegion)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetRegionDataSize

Gets the number of bytes of data that describes this region.

```
FUNCTION GdipGetRegionDataSize (BYVAL region AS GpRegion PTR, BYVAL bufferSize AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *bufferSize* | [out] Pointer to a variable that receives the number of bytes of region data. |

#### Remarks

The **GdipGetRegionDataSize** function can be used before the ^^GdipGetRegionData** function to determine the number of bytes needed to store the region data. Then, you can allocate a buffer that is the correct size to store the region data that is obtained by the **GdipGetRegionData**.

#### Usage exanple

See full example in **GdipGetRegionData**.

```
   ' // Get the region data.
   DIM bufferSize AS UINT
   hStatus = GdipGetRegionDataSize(pathRegion, @bufferSize)
   DIM sizeFilled AS UINT
   DIM strBuffer AS STRING = SPACE(bufferSize)
   hStatus = GdipGetRegionData(pathRegion, STRPTR(strBuffer), bufferSize, @sizeFilled)
```
---

## GdipGetRegionHRgn

Creates a Microsoft Windows Graphics Device Interface (GDI) region from this region.

```
FUNCTION GdipGetRegionHRgn (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, _
    BYVAL hRgn AS HRGN PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *graphics* | [in] Pointer to a Graphics object that contains the world and page transformations required to calculate the device coordinates of this region. |
| *hRgn* | [out] Pointer to a variable that receives a Windows handle to a GDI region that is created from this region. |

#### Example

```
' ========================================================================================
' The following example creates a Windows GDI+ region from a path and then uses the GDI+
' region to create a GDI region. The code then uses a GDI function to display the GDI region.
' ========================================================================================
SUB Example_GetRegionHRgn (BYVAL hdc AS HDC)

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

   ' // Create a path.
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)
   
   ' // Add a closed curve to a path.
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)


   ' // Create a region from a path.
   DIM pathRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @pathRegion)

   ' // Get a handle to a GDI region.
   DIM hRegion AS HRGN
   hStatus = GdipGetRegionHRgn(pathRegion, graphics, @hRegion)

   ' // Create a GDI red brush
   DIM hBrush AS HBRUSH = CreateSolidBrush(&h0000FF)

   ' // Use GDI to display the region.
   FillRgn(hdc, hRegion, hBrush)

   ' // Cleanup
   IF hBrush THEN DeleteObject(hBrush)
   IF hRegion THEN DeleteObject(hRegion)
   IF pathRegion THEN GdipDeleteRegion(pathRegion)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetRegionScans

Gets an array of rectangles that approximate this region. The region is transformed by a specified matrix before the rectangles are calculated.

```
FUNCTION GdipGetRegionScans (BYVAL region AS GpRegion PTR, BYVAL rects AS GpRectF PTR, _
   BYVAL count AS INT_ PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
```
```
FUNCTION GdipGetRegionScansI (BYVAL region AS GpRegion PTR, BYVAL rects AS GpRect PTR, _
   BYVAL count AS INT_ PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *rects* | [out] Pointer to an array of **GpRectF** structures that receives the rectangles. |
| *count* | [out] Pointer to a LONG that receives a value that indicates the number of rectangles that approximate this region. The value is valid even if rects is a NULL pointer. |
| *region* | [in] Pointer to a **Matrix** object that is used to transform the region. |

#### Remarks

The **GdipGetRegionScansCount** function can be used first to determine the number of rectangles. Then, you can allocate a buffer that is the correct size and set the rects parameter to point to the buffer.

#### Example

```
' ========================================================================================
' The following example creates a region from a path and gets a set of rectangles that
' approximate the region. The code then draws each of the rectangles.
' ========================================================================================
SUB Example_GetRegionScans (BYVAL hdc AS HDC)

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

   ' // Create a Path
   DIM path AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @path)

   ' // Add an ellipse to the path
   hStatus = GdipAddPathEllipse(path, 30, 10, 50, 230)

   ' // Create a region from a path.
   DIM pathRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @pathRegion)

   ' // Fill the region with a solid brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)
   hStatus = GdipFillRegion(graphics, solidBrush, pathRegion)

   ' // Get current world transform matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)
   hStatus = GdipGetWorldTransform(graphics, @matrix)

   ' // Get the rectangles.
   DIM count AS UINT
   hStatus = GdipGetRegionScansCount(pathRegion, @count, matrix)
   DIM rects(count - 1) AS GpRectF
   hStatus = GdipGetRegionScans(pathRegion, @rects(0), @count, matrix)

   ' // Create a Pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 1.0, UnitWorld, @pen)

   ' // Draw the rectangles.
   FOR j AS LONG = 0 TO count - 1
      hStatus = GdipDrawRectangle(graphics, pen, rects(j).x, rects(j).y, rects(j).width, rects(j).height)
   NEXT

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF pen THEN GdipDeletePen(pen)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pathRegion THEN GdipDeleteRegion(pathRegion)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetRegionScansCount

Gets the number of rectangles that approximate this region. The region is transformed by a specified matrix before the rectangles are calculated.

```
FUNCTION GdipGetRegionScansCount (BYVAL region AS GpRegion PTR, BYVAL count AS UINT PTR, _
   BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *count* | [out] Pointer to an UINT that receives a value that indicates the number of rectangles that approximate this region. |
| *matrix* | [in] Pointer to a **Matrix** object that is used to transform the region. |

#### Remarks

The **GdipGetRegionScansCount** function can be used before the **GdipGetRegionScans** function to determine the number of rectangles. Then, you can allocate a buffer that is the correct size to store the rectangles that are obtained with the **GdipGetRegionScans** function.

#### Usage example

See full example in**GdipGetRegionScans**.

```
   ' // Get the rectangles.
   DIM count AS UINT
   hStatus = GdipGetRegionScansCount(pathRegion, @count, matrix)
   DIM rects(count - 1) AS GpRectF
   hStatus = GdipGetRegionScans(pathRegion, @rects(0), @count, matrix)
```
---

## GdipIsEmptyRegion

Determines whether this region is empty.

```
FUNCTION GdipIsEmptyRegion (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, _
   BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *graphics* | [in] Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region. |
| *result* | [out] Pointer to a variable that receives a boolean value that determines whether this region is empty (TRUE) or not (FALSE). |

### Usage example

```
   ' // Make the region empty
   hStatus = GdipSetEmpty(region)
   DIM isEmpty AS BOOL
   hStatus = GdipIsEmptyRegion(region, graphics, @isEmpty)
   IF isEmpty THEN AfxMsg "The region is empty"
```
---

## GdipIsEqualRegion

Determines whether this region is equal to a specified region.

```
FUNCTION GdipIsEqualRegion (BYVAL region AS GpRegion PTR, BYVAL region2 AS GpRegion PTR, _
   BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *region2* | [in] Pointer to a **Region** object to test against this **Region** object. |
| *graphics* | [in] Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region and the region that is specified by the region parameter. |
| *result* | [out] Pointer to a variable that receives a boolean value that determines whether this region equal to a specified region (TRUE) or not (FALSE). |

#### Example

```
   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   GdipCreateFromHDC(hdc, @graphics)

   ' // Create a rectangular region.
   DIM rectRegion AS GpRegion PTR
   DIM rcf AS GpRectF = (65, 15, 70, 45)
   GdipCreateRegionRect(@rcf, @rectRegion)

   ' // Create another rectangular region.
   DIM rectRegion2 AS GpRegion PTR
   GdipCreateRegionRect(@rcf, @rectRegion2)

   DIM result AS BOOL
   GdipIsEqualRegion (rectRegion, rectRegion2, graphics, @result)
   IF result THEN AfxMsg("Equal") ELSE AfxMsg("Not equal")
   
   ' // Cleanup
   IF rectRegion THEN GdipDeleteRegion(rectRegion)
   IF rectRegion2 THEN GdipDeleteRegion(rectRegion2)
   IF graphics THEN GdipDeleteGraphics(graphics)
```
---

## GdipIsInfiniteRegion

Determines whether this region is infinite.

```
FUNCTION GdipIsInfiniteRegion (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, _
   BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *graphics* | [in] Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region. |
| *result* | [out] Pointer to a variable that receives a boolean value that determines whether this region is infinite (TRUE) or not (FALSE). |

### Usage example

```
  ' // Make the region infinite
   hStatus = GdipSetInfinite(region)
   DIM isInfinite AS BOOL
   hStatus = GdipIsInfiniteRegion(region, graphics, @isInfinite)
   IF isInfinite THEN AfxMsg "The region is infinite"
```
---

## GdipIsVisibleRegionPoint

Determines whether a point is inside this region.

```
FUNCTION GdipIsVisibleRegionPoint (BYVAL region AS GpRegion PTR, BYVAL x AS REAL, _
   BYVAL y AS REAL, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
```
```
FUNCTION GdipIsVisibleRegionPointI (BYVAL region AS GpRegion PTR, BYVAL x AS INT_, _
   BYVAL y AS INT_, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *x* | [in] The x-coordinate of the point to test. |
| *y* | [in] The y-coordinate of the point to test. |
| *graphics* | [in] Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region and the point. The default value is NULL. |
| *result* | [out] Pointer to a variable that receives a boolean value that determines whether a point is inside this region (TRUE) or not (FALSE). |

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle and then checks if a point is
' inside this region.
' ========================================================================================
SUB Example_IsVisibleRegionPoint (BYVAL hdc AS HDC)

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

   ' // Create a region from a rectangle
   DIM rcf AS GpRectF = (50, 50, 100, 100)
   DIM region AS GpRegion PTR
   hStatus = GdipCreateRegionRect(@rcf, @region)

   ' // Fill the the region with a red brush.
   DIM redBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @redBrush)
   hStatus = GdipFillRegion(graphics, redBrush, region)

   ' // Test if a point is inside the region
   DIM result AS BOOL
   hStatus = GdipIsVisibleRegionPoint(region, 75.0, 75.0, graphics, @result)
   IF result THEN
      MessageBox(NULL, "Point is inside the region", "GDI+", MB_OK)
   ELSE
      MessageBox(NULL, "Point is outside the region", "GDI+", MB_OK)
   END IF

   ' // Cleanup
   IF region THEN GdipDeleteRegion(region)
   IF redBrush THEN GdipDeleteBrush(redBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipIsVisibleRegionRect

Determines whether a rectangle intersects this region.

```
FUNCTION GdipIsVisibleRegionRect (BYVAL region AS GpRegion PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
   BYVAL width AS REAL, BYVAL height AS REAL, BYVAL graphics AS GpGraphics PTR, _
   BYVAL result AS BOOL PTR) AS GpStatus
```
```
FUNCTION GdipIsVisibleRegionRectI (BYVAL region AS GpRegion PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL width AS INT_, BYVAL height AS INT_, BYVAL graphics AS GpGraphics PTR, _
   BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *x* | [in] The x-coordinate of the upper-left corner of the rectangle to test. |
| *y* | [in] The y-coordinate of the upper-left corner of the rectangle to test. |
| *width* | [in] The width of the rectangle to test. |
| *height* | [in] The height of the rectangle to test. |
| *graphics* | [in] Pointer to a **Graphics** object that contains the world and page transformations required to calculate the device coordinates of this region and the rectangle. The default value is NULL. |
| *result* | [out] Pointer to a variable that receives a boolean value that determines whether a rectangle intersects this region (TRUE) or not (FALSE). |

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle and then tests if another
' rectangle intersects the region.
' ========================================================================================
SUB Example_IsVisibleRegionRect (BYVAL hdc AS HDC)

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

   ' // Create a region from a rectangle
   DIM rcf AS GpRectF = (50, 50, 100, 100)
   DIM region AS GpRegion PTR
   hStatus = GdipCreateRegionRect(@rcf, @region)

   ' // Fill the the region with a red brush.
   DIM redBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @redBrush)
   hStatus = GdipFillRegion(graphics, redBrush, region)

   ' // Test if another rectangle intersects the region
   DIM result AS BOOL
   hStatus = GdipIsVisibleRegionRect(region, 120.0, 120.0, 30.0, 30.0, graphics, @result)
   IF result THEN
      MessageBox(NULL, "Rectangle intersects the region", "GDI+", MB_OK)
   ELSE
      MessageBox(NULL, "Rectangle does not intersect the region", "GDI+", MB_OK)
   END IF

   ' // Cleanup
   IF region THEN GdipDeleteRegion(region)
   IF redBrush THEN GdipDeleteBrush(redBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetEmpty

Updates this region to an empty region. In other words, the region occupies no space on the display device.

```
FUNCTION GdipSetEmpty (BYVAL region AS GpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle and paints it with a brush. Then
' the code makes the region empty and paints the region with a different color to show that
' the region occupies no space on the display device.
' ========================================================================================
SUB Example_SetEmpty (BYVAL hdc AS HDC)

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

   ' // Create a rectangular region.
   DIM rectRegion AS GpRegion PTR
   DIM rcf AS GpRectF = (50, 60, 300, 120)
   hStatus = GdipCreateRegionRect(@rcf, @rectRegion)
   
   ' // Fill the region with a red brush
   DIM redBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @redBrush)
   hStatus = GdipFillRegion(graphics, redBrush, rectRegion)

   ' // Make the region empty, and then fill it with a blue brush.
   ' // The color won't change because it is empty and can't be filled with FillRegion.
   hStatus = GdipSetEmpty(rectRegion)
   DIM blueBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @blueBrush)
   hStatus = GdipFillRegion(graphics, blueBrush, rectRegion)

      ' // Cleanup
   IF redBrush THEN GdipDeleteBrush(redBrush)
   IF blueBrush THEN GdipDeleteBrush(blueBrush)
   IF rectRegion THEN GdipDeleteRegion(rectRegion)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetInfinite

Updates this region to an infinite region.

```
FUNCTION GdipSetInfinite (BYVAL region AS GpRegion PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |

#### Example

```
' ========================================================================================
' The following example creates a region from a rectangle, makes the region infinite, and
' fills the region to show the infinite region.
' ========================================================================================
SUB Example_SetInfinite (BYVAL hdc AS HDC)

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

   ' // Create a rectangular region.
   DIM rectRegion AS GpRegion PTR
   DIM rcf AS GpRectF = (65, 15, 70, 45)
   hStatus = GdipCreateRegionRect(@rcf, @rectRegion)

  ' // Make the region infinite, and then fill it with a blue brush.
   hStatus = GdipSetInfinite(rectRegion)
   
   ' // Fill the region with a blue brush
   DIM blueBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @blueBrush)
   hStatus = GdipFillRegion(graphics, blueBrush, rectRegion)

   ' // Cleanup
   IF blueBrush THEN GdipDeleteBrush(blueBrush)
   IF rectRegion THEN GdipDeleteRegion(rectRegion)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTransformRegion

Transforms this region by multiplying each of its data points by a specified matrix.

```
FUNCTION GdipTransformRegion (BYVAL region AS GpRegion PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *matrix* | [in] Pointer to a matrix that specifies the transformation. |

#### Example

```
' ========================================================================================
' The following example creates a region from a path and fills it. The code then transforms
' the region and fills the transformed region to show the transformation.
' ========================================================================================
SUB Example_GetTransformRegion (BYVAL hdc AS HDC)

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

   ' // Add a closed curve to the path
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)

   ' // Create a region from a path.
   DIM pathRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @pathRegion)

   ' // Transform the region.
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(1, 0, 0, 3, 100, 0, @matrix)   ' // vertical stretch, shift right
   hStatus = GdipTransformRegion(pathRegion, matrix)

   ' // Fill the region with a brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)
   hStatus = GdipFillRegion(graphics, solidBrush, pathRegion)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pathRegion THEN GdipDeleteRegion(pathRegion)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTranslateRegion

Offsets this region by specified amounts in the horizontal and vertical directions.

```
FUNCTION GdipTranslateRegion (BYVAL region AS GpRegion PTR, BYVAL dx AS REAL, BYVAL dy AS REAL) AS GpStatus
```
```
FUNCTION GdipTranslateRegionI (BYVAL region AS GpRegion PTR, BYVAL dx AS INT_, BYVAL dy AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *region* | [in] Pointer to the **Region** object. |
| *dx* | The amount to shift the region in the x direction. |
| *dy* | The amount to shift the region in the y direction. |

#### Example

```
' ========================================================================================
' The following example creates a region from a path and fills it. The code then translates
' the region and fills the translated region to show how the region has shifted.
' ========================================================================================
SUB Example_GetTranslateRegion (BYVAL hdc AS HDC)

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

   ' // Add a closed curve to the path
   DIM pts(5) AS GpPointF
   pts(0).x = 110 : pts(0).y = 20
   pts(1).x = 120 : pts(1).y = 30
   pts(2).x = 100 : pts(2).y = 60
   pts(3).x = 120 : pts(3).y = 70
   pts(4).x = 150 : pts(4).y = 60
   pts(5).x = 140 : pts(5).y = 10
   hStatus = GdipAddPathClosedCurve(path, @pts(0), 6)

   ' // Create a region from a path.
   DIM pathRegion AS GpRegion PTR
   hStatus = GdipCreateRegionPath(path, @pathRegion)

   ' // Translate the region.
   hStatus = GdipTranslateRegion(pathRegion, 100, 60)

   ' // Fill the region with a brush
   DIM solidBrush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_RED, @solidBrush)
   hStatus = GdipFillRegion(graphics, solidBrush, pathRegion)

   ' // Cleanup
   IF solidBrush THEN GdipDeleteBrush(solidBrush)
   IF pathRegion THEN GdipDeleteRegion(pathRegion)
   IF path THEN GdipDeletePath(path)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
