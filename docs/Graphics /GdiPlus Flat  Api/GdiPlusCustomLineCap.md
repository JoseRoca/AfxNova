# CustomLineCap Functions

A line cap defines the style of graphic used to draw the ends of a line. It can be various shapes, such as a square, circle, or diamond. A custom line cap is defined by the path that draws it. The path is drawn by using a **Pen** object to draw the outline of a shape or by using a **Brush** object to fill the interior. The cap can be used on either or both ends of the line. Spacing can be adjusted between the end caps and the line.

| Name       | Description |
| ---------- | ----------- |
| [GdipCloneCustomLineCap](#gdipclonecustomlinecap) | Copies the contents of the existing object into a new **CustomLineCap** object. |
| [GdipCreateCustomLineCap](#gdipcreatecustomlinecap) | Creates a **CustomLineCap** object. |
| [GdipDeleteCustomLineCap](#gdipdeletecustomlinecap) | Deletes a **CustomLineCap** object. |
| [GdipGetCustomLineCapBaseCap](#gdipgetcustomlinecapbasecap) | Gets the style of the base cap. The base cap is a **LineCap** object used as a cap at the end of a line along with this **CustomLineCap** object. |
| [GdipGetCustomLineCapBaseInset](#gdipgetcustomlinecapbaseinset) | Gets the base inset value of this custom line cap. This is the distance between the end of a line and the base cap. |
| [GdipGetCustomLineCapStrokeCaps](#gdipgetcustomlinecapstrokecaps) | Gets the end cap styles for both the start line cap and the end line cap. Line caps are **LineCap** objects that end the individual lines within a path. |
| [GdipGetCustomLineCapStrokeJoin](#gdipgetcustomlinecapstrokejoin) | Returns the style of **LineJoin** used to join multiple lines in the same **GraphicsPath** object. |
| [GdipGetCustomLineCapType](#gdipgetcustomlinecaptype) | Gets the type of the **CustomLineCap** specified by *customCap*. |
| [GdipGetCustomLineCapWidthScale](#gdipgetcustomlinecapwidthscale) | Gets the value of the scale width. |
| [GdipSetCustomLineCapBaseCap](#gdipsetcustomlinecapbasecap) | Sets the **LineCap** that appears as part of this **CustomLineCap** at the end of a line. |
| [GdipSetCustomLineCapBaseInset](#gdipsetcustomlinecapbaseinset) | Sets the base inset value of this custom line cap. This is the distance between the end of a line and the base cap. |
| [GdipSetCustomLineCapStrokeCaps](#gdipsetcustomlinecapstrokecaps) | Sets the **LineCap** objects used to start and end lines within the **GraphicsPath** object that defines this **CustomLineCap** object. |
| [GdipSetCustomLineCapStrokeJoin](#gdipsetcustomlinecapstrokejoin) | Sets the style of line join for the stroke. |
| [GdipSetCustomLineCapWidthScale](#gdipsetcustomlinecapwidthscale) | Sets the value of the scale width. |

---

## GdipCloneCustomLineCap

Copies the contents of the existing object into a new **CustomLineCap** object.

```
FUNCTION GdipCloneCustomLineCap (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL clonedCap AS GpCustomLineCap PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *clonedCap* | [out] Pointer to a **CustomLineCap** object variable that receives a pointer to the new **CustomLineCap** object. |

#### Example

The following example creates a CustomLineCap object with a stroke path, creates a second CustomLineCap object by cloning the first, and then assigns the cloned cap to a Pen object. It then draws a line by using the Pen object.

```
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the window device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get the DPI scaling ratios
   DIM dpiX AS SINGLE, dpiY AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM ryRatio AS SINGLE = dpiY / 96

   ' // Set the scale transform
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

   ' // Create the first CustomLineCap object
   DIM firstCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapFlat, 0, @firstCap)

   ' // Clone the first CustomLineCap into secondCap
   DIM secondCap AS GpCustomLineCap PTR
   hStatus = GdipCloneCustomLineCap(firstCap, @secondCap)

   ' // Create a violet Pen with width 1
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_VIOLET, 1, UnitWorld, @pen)

   ' // Assign the cloned cap as the end cap of the pen
   hStatus = GdipSetPenCustomEndCap(pen, secondCap)

   ' // Draw a line using the pen
   hStatus = GdipDrawLine(graphics, pen, 0, 0, 100, 100)

   ' // Clean up resources
   IF pen THEN GdipDeletePen(pen)
   IF secondCap THEN GdipDeleteCustomLineCap(secondCap)
   IF firstCap THEN GdipDeleteCustomLineCap(firstCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateCustomLineCap

Creates a **CustomLineCap** object.

```
FUNCTION GdipCreateCustomLineCap (BYVAL fillPath AS GpPath PTR, BYVAL strokePath AS GpPath PTR, _
   BYVAL baseCap AS GpLineCap, BYVAL baseInset AS REAL, BYVAL customCap AS GpCustomLineCap PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *fillPath* | [in] Pointer to a **Path** object. |
| *strokePath* | [in] Pointer to a **Path** object. |
| *baseCap* | [in] Element of the *LineCap* enumeration that specifies the line cap that will be used. The default value is *LineCapFlat*. |
| *baseInset* | [in] The default value is 0. |
| *customCap* | [out] Pointer to a variable that receives a pointer to the new **CustomLineCap** object. |

#### Remarks

The *fillPath* and *strokePath* parameters cannot be used at the same time. You should pass **NULL** to one of those two parameters. If you pass nonnull values to both parameters, then *fillPath* is ignored.

```
' ========================================================================================
SUB Example_CreateCustomLineCap (BYVAL hdc AS HDC)

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

   ' // Create an empty Path object for the custom cap
   DIM capPath AS GpPath PTR
   GdipCreatePath(FillModeAlternate, @capPath)

   ' // Create a CustomLineCap using the empty path
   DIM customCap AS GpCustomLineCap PTR
   GdipCreateCustomLineCap(NULL, capPath, LineCapFlat, 0, @customCap)

   ' // Create a green Pen with width 10
   DIM pen AS GpPen PTR
   GdipCreatePen1(ARGB_RED, 10, UnitWorld, @pen)

   ' // Assign the custom cap as the end cap of the pen
   GdipSetPenCustomEndCap(pen, customCap)

   ' // Draw a line using the pen and graphics object
   GdipDrawLine(graphics, pen, 0, 0, 100, 100)

   ' // Clean up resources
   IF pen THEN GdipDeletePen(pen)
   IF customCap THEN GdipDeleteCustomLineCap(customCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeleteCustomLineCap

Deletes a **CustomLineCap** object.

```
FUNCTION GdipDeleteCustomLineCap (BYVAL customCap AS GpCustomLineCap PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |

---

## GdipGetCustomLineCapBaseCap

Gets the style of the base cap. The base cap is a **LineCap** object used as a cap at the end of a line along with this **CustomLineCap** object.
```
FUNCTION GdipGetCustomLineCapBaseCap (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL baseCap AS GpLineCap PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *baseCap* | [out] Pointer to a *GpLineCap* variable that receives the value of the line cap used at the base of the line. |

The following example creates a CustomLineCap object, sets its base cap, and then gets the base cap and assigns it to a Pen object. It then uses the Pen object to draw a line.

```
' ========================================================================================
SUB Example_GetBaseCap (BYVAL hdc AS HDC)

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

   ' // Create an empty Path object for the custom cap
   DIM capPath AS GpPath PTR
   GdipCreatePath(FillModeAlternate, @capPath)

   ' // Create a CustomLineCap using the empty path
   DIM customCap AS GpCustomLineCap PTR
   GdipCreateCustomLineCap(NULL, capPath, LineCapFlat, 0, @customCap)

   ' // Set the base cap style to LineCapRound
   GdipSetCustomLineCapBaseCap(customCap, LineCapRound)

   ' // Get the base cap of custCap
   DIM baseCap AS LineCap
   GdipGetCustomLineCapBaseCap(customCap, @baseCap)

   ' // Create a green Pen with width 10
   DIM pen AS GpPen PTR
   GdipCreatePen1(ARGB_LIGHTGREEN, 10, UnitWorld, @pen)

   ' // Assign the custom cap as the end cap of the pen
   GdipSetPenCustomEndCap(pen, customCap)
   GdipSetPenEndCap(pen, baseCap)

   ' // Draw a line using the pen and graphics object
   GdipDrawLine(graphics, pen, 0, 0, 100, 100)

   ' // Clean up resources
   IF pen THEN GdipDeletePen(pen)
   IF customCap THEN GdipDeleteCustomLineCap(customCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetCustomLineCapBaseInset

Gets the base inset value of this custom line cap. This is the distance between the end of a line and the base cap.

```
FUNCTION GdipGetCustomLineCapBaseInset (BYVAL customCap AS GpCustomLineCap PTR, BYVAL inset AS REAL PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *inset* | [in] Single precision value that specifies the distance, in units, from the base cap to the start of the line. |

#### Example

The following example creates a CustomLineCap object, gets the base inset of the cap, and then creates a second CustomLineCap object that uses the same base inset.

```
' ========================================================================================
SUB Example_GetBaseInset (BYVAL hdc AS HDC)

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

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

   ' // Create a CustomLineCap with a base cap set to LineCapRound and baseInset = 5
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapRound, 5, @custCap)

   ' // Get the base inset from custCap
   DIM baseInset AS SINGLE
   hStatus = GdipGetCustomLineCapBaseInset(custCap, @baseInset)

   ' // Create a second CustomLineCap using the same baseInset
   DIM insetCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapRound, baseInset, @insetCap)

   ' // Create a blue Pen with width 5
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 5, UnitWorld, @pen)

   ' // Assign insetCap as the custom end cap
   hStatus = GdipSetPenCustomEndCap(pen, insetCap)

   ' // Draw a line using the pen
   hStatus = GdipDrawLine(graphics, pen, 10, 10, 200, 200)

   ' // Clean up resources
   IF pen THEN GdipDeletePen(pen)
   IF insetCap THEN GdipDeleteCustomLineCap(insetCap)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetCustomLineCapStrokeCaps

Gets the end cap styles for both the start line cap and the end line cap. Line caps are **LineCap** objects that end the individual lines within a path.

```
FUNCTION GdipGetCustomLineCapStrokeCaps (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL startCap AS GpLineCap PTR, BYVAL endCap AS GpLineCap PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *startCap* | [out] Pointer to a variable that receives an element of the **LineCap** enumeration that indicates the line cap used at the start of the line to be drawn. |
| *endCap* | [out] Pointer to a variable that receives an element of the **LineCap** enumeration that indicates the line cap used at the end of the line to be drawn. |

---

## GdipGetCustomLineCapStrokeJoin

Returns the style of **LineJoin** used to join multiple lines in the same **GraphicsPath** object.

```
FUNCTION GdipGetCustomLineCapStrokeJoin (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL lineJoin AS GpLineJoin PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *lineJoin* | [out] The style of line join. |

#### Remarks

The **CustomLineCap** object uses a path and a stroke to define the end cap. The stroke is contained in a **GraphicsPath** object, which can contain more than one figure. If there is more than one figure in the **GraphicsPath** object, the stroke join determines how their joint is graphically displayed.

#### Example

The following example creates a CustomLineCap object with a stroke join. It then gets the stroke join and assigns it as the line join of a Pen object that it then uses to draw a line.

```
' ========================================================================================
SUB Example_GetStrokeJoin (BYVAL hdc AS HDC)

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

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

   ' // Create a CustomLineCap object
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapRound, 0, @custCap)

   ' // Set the stroke join style
   hStatus = GdipSetCustomLineCapStrokeJoin(custCap, LineJoinBevel)

   ' // Get the stroke join from custCap
   DIM strokeJoin AS LineJoin
   hStatus = GdipGetCustomLineCapStrokeJoin(custCap, @strokeJoin)

   ' // Create a red Pen with width 15.1
   DIM strokeJoinPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 15.1, UnitWorld, @strokeJoinPen)

   ' // Set the line join style on the pen
   hStatus = GdipSetPenLineJoin(strokeJoinPen, strokeJoin)

   ' // Create a path with two joined lines
   DIM joinPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @joinPath)
   hStatus = GdipAddPathLine(joinPath, 10, 10, 10, 200)
   hStatus = GdipAddPathLine(joinPath, 10, 200, 200, 200)

   ' // Draw the path using the pen
   hStatus = GdipDrawPath(graphics, strokeJoinPen, joinPath)

   ' // Clean up resources
   IF strokeJoinPen THEN GdipDeletePen(strokeJoinPen)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF joinPath THEN GdipDeletePath(joinPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetCustomLineCapType

Gets the type of the CustomLineCap specified by *customCap*.

```
FUNCTION GdipGetCustomLineCapType (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL capType AS CustomLineCapType PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *capType* | [out] Pointer to a variable that receives the type of the **CustomLineCap** specified by *customCap*.<br> The *CustomLineCapType* enumeration has two elements: CustomLineCapTypeDefault = 0 and CustomLineCapTypeAdjustableArrow = 1. |

```
' ========================================================================================
' The following example creates two AdjustableArrowCap objects, arrowCapStart and
' arrowCapEnd, and sets the fill mode to TRUE. The code then creates a Pen object and
' assigns arrowCapStart as the starting line cap for this Pen object and arrowCapEnd as
' the ending line cap. Next, draws a line.
' ========================================================================================
SUB Example_LineCapType (BYVAL hdc AS HDC)

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

   ' // Create an AdjustableArrowCap that is filled.
   DIM arrowCapStart AS GpAdjustableArrowCap PTR
   hStatus = GdipCreateAdjustableArrowCap(10, 10, TRUE, @arrowCapStart)
   ' // Adjust to DPI by setting the scale width
   hStatus = GdipSetCustomLineCapWidthScale(arrowCapStart, rxRatio)

   ' // Get the type of CustomLineCap
   ' // It will return 1 (CustomLineCapTypeAdjustableArrow)
   DIM nType AS CustomLineCapType
   hStatus = GdipGetCustomLineCapType(arrowCapStart, @nType)
   AfxMsg "arrowCapStart type = " & WSTR(nType)

   ' // Create an AdjustableArrowCap that is not filled.
   DIM arrowCapEnd AS GpAdjustableArrowCap PTR
   hStatus = GdipCreateAdjustableArrowCap(15, 15, FALSE, @arrowCapEnd)
   ' // Adjust to DPI by setting the scale width
   hStatus = GdipSetCustomLineCapWidthScale(arrowCapEnd, rxRatio)

   ' // Get the type of CustomLineCap
   ' // It will return 1 (CustomLineCapTypeAdjustableArrow)
   hStatus = GdipGetCustomLineCapType(arrowCapEnd, @nType)
   AfxMsg "arrowCapEnd type = " & WSTR(nType)

   ' // Create a Pen
   DIM arrowPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_Violet, 1, UnitWorld, @arrowPen)

   ' Assign arrowCapStart as the start cap.
   hStatus = GdipSetPenCustomStartCap(arrowPen, arrowCapStart)
   ' Assign arrowCapEnd as the end cap.
   hStatus = GdipSetPenCustomEndCap(arrowPen, arrowCapEnd)

   ' // Draw a line using arrowPen.
   hStatus = GdipDrawLine(graphics, arrowPen, 0, 0, 100, 100)

   ' // Cleanup
   IF arrowCapStart THEN GdipDeleteCustomLineCap(arrowCapStart)
   IF arrowCapEnd THEN GdipDeleteCustomLineCap(arrowCapEnd)
   IF arrowPen THEN GdipDeletePen(arrowPen)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetCustomLineCapWidthScale

Gets the value of the scale width. This is the amount to scale the custom line cap relative to the width of the **Pen** object used to draw a line. The default value of 1.0 does not scale the line cap.

```
FUNCTION GdipGetCustomLineCapWidthScale (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL widthScale AS REAL PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *widthScale* | [in] Pointer to a variable that receives the value of the width-scaling factor. |

#### Example

The following example creates a CustomLineCap object and sets the width scale. It then gets the width scale, assigns the custom cap to a Pen object, and draws a line.

```
' ========================================================================================
SUB Example_GetWidthScale (BYVAL hdc AS HDC)

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

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

  ' // Create a CustomLineCap
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapRound, 0, @custCap)

   ' // // Set the width scale for custCap.
   hStatus = GdipSetCustomLineCapWidthScale(custCap, 3)

   ' // Get the width scale from custCap
   DIM widthScale AS SINGLE
   hStatus = GdipGetCustomLineCapWidthScale(custCap, @widthScale)

   ' // If the width scale is 3, assign custCap as the end cap of a Pen object and draw a line
   IF widthScale = 3.0 THEN
      DIM widthScalePen AS GpPen PTR
      hStatus = GdipCreatePen1(ARGB_LIGHTGREEN, 1, UnitWorld, @widthScalePen)
      hStatus = GdipSetPenCustomEndCap(widthScalePen, custCap)
      hStatus = GdipDrawLine(graphics, widthScalePen, 0, 0, 200, 200)
      IF widthScalePen THEN GdipDeletePen(widthScalePen)
   END IF

   ' // Clean up resources
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetCustomLineCapBaseCap

Sets the **LineCap** that appears as part of this **CustomLineCap** at the end of a line.

```
FUNCTION GdipSetCustomLineCapBaseCap (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL baseCap AS GpLineCap) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *baseCap* | [in] Element of the **LineCap** enumeration that specifies the line cap used on the ends of the line to be drawn. |

#### Example

The following example creates a CustomLineCap object and sets its base cap. It then assigns the custom cap as the end cap of a Pen object and draws a line.

```
' ========================================================================================
SUB Example_SetBaseCap (BYVAL hdc AS HDC)

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

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)

   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

   ' // Create a CustomLineCap object
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapFlat, 0, @custCap)

   ' // Set the base cap style to LineCapRound
   hStatus = GdipSetCustomLineCapBaseCap(custCap, LineCapRound)

   ' // Create a blue Pen with width 5.3
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 5.3, UnitWorld, @pen)

   ' // Assign the custom cap as the end cap of the pen
   hStatus = GdipSetPenCustomEndCap(pen, custCap)

   ' // Draw a line using the pen
   hStatus = GdipDrawLine(graphics, pen, 10, 10, 200, 200)

   ' // Clean up resources
   IF pen THEN GdipDeletePen(pen)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetCustomLineCapBaseInset

Sets the base inset value of this custom line cap. This is the distance between the end of a line and the base cap.

```
FUNCTION GdipSetCustomLineCapBaseInset (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL inset AS REAL) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *inset* | [in] Single precision value that specifies the distance, in units, from the base cap to the start of the line. If this value is greater than the length of the line, the behavior of this method is undefined. |

#### Example

The following example creates a CustomLineCap object and sets the base inset of the cap. It then assigns the custom cap to a Pen object and draws a line.

```
' ========================================================================================
SUB Example_SetBaseInset (BYVAL hdc AS HDC)

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

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

   ' // Create a CustomLineCap object with base cap LineCapRound
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapRound, 0.0, @custCap)

   ' // Set the base inset value
   hStatus = GdipSetCustomLineCapBaseInset(custCap, 5.0)

   ' // Create a black Pen with width 5.1
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 5.1, UnitWorld, @pen)

   ' // Assign the custom cap as the end cap of the pen
   hStatus = GdipSetPenCustomEndCap(pen, custCap)

   ' // Draw a line using the pen
   hStatus = GdipDrawLine(graphics, pen, 10, 10, 200, 200)

   ' // Clean up resources
   IF pen THEN GdipDeletePen(pen)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetCustomLineCapStrokeCaps

Sets the **LineCap** objects used to start and end lines within the **GraphicsPath** object that defines this **CustomLineCap** object.

```
FUNCTION GdipSetCustomLineCapStrokeCaps (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL startCap AS GpLineCap, BYVAL endCap AS GpLineCap) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *startCap* | [in] Element of the **LineCap** enumeration that specifies the line cap that will be used for the start of the line to be drawn. |
| *endCap* | [[in] Element of the **LineCap** enumeration that specifies the line cap that will be used for the end of the line to be drawn. |

#### Example

The following example creates a CustomLineCap object and sets its start and end stroke caps. It then assigns the custom cap to a Pen object and draws a line.

```
' ========================================================================================
SUB Example_SetStrokeCaps (BYVAL hdc AS HDC)

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

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

   ' // Create a CustomLineCap object
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapRound, 0.0, @custCap)

   ' // Set the start and end stroke caps
   hStatus = GdipSetCustomLineCapStrokeCaps(custCap, LineCapTriangle, LineCapRound)

   ' // Create a magenta Pen with width 5.0
   DIM strokeCapPen AS GpPen PTR
   hStatus = GdipCreatePen1(GDIP_ARGB(255, 255, 0, 255), 5.0, UnitWorld, @strokeCapPen)

   ' // Assign the custom cap as the end cap of the pen
   hStatus = GdipSetPenCustomEndCap(strokeCapPen, custCap)

   ' // Draw a line using the pen
   hStatus = GdipDrawLine(graphics, strokeCapPen, 100, 100, 300, 100)

   ' // Clean up resources
   IF strokeCapPen THEN GdipDeletePen(strokeCapPen)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetCustomLineCapStrokeJoin

Sets the style of line join for the stroke. The line join specifies how two lines that intersect within the **GraphicsPath** object that makes up the custom line cap are joined.

```
FUNCTION GdipSetCustomLineCapStrokeJoin (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL lineJoin AS GpLineJoin) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *lineJoin* | [in] Element of the *LineJoin* enumeration that specifies the line join that will be used for two lines that are drawn by the same pen and that have overlapping ends. |

#### Example

The following example creates a CustomLineCap object with a stroke join. It then gets the stroke join and assigns it as the line join of a Pen object that it then uses to draw a line.
```
' ========================================================================================
SUB Example_SetStrokeJoin (BYVAL hdc AS HDC)

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

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

   ' // Create a CustomLineCap object with base cap LineCapRound
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapRound, 0.0, @custCap)

   ' // Set the stroke join style to LineJoinBevel
   hStatus = GdipSetCustomLineCapStrokeJoin(custCap, LineJoinBevel)

   ' // Create a red Pen with width 15.1
   DIM strokeJoinPen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 15.1, UnitWorld, @strokeJoinPen)

   ' // Assign the custom cap as the end cap of the pen
   hStatus = GdipSetPenCustomEndCap(strokeJoinPen, custCap)

   ' // Draw a line using the pen
   hStatus = GdipDrawLine(graphics, strokeJoinPen, 0, 0, 200, 200)

   ' // Clean up resources
   IF strokeJoinPen THEN GdipDeletePen(strokeJoinPen)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

#### Example

The following example creates a CustomLineCap object with a stroke join. It then sets the stroke join and assigns the custom cap to a Pen object, which is then used to draw a line.
```
' ========================================================================================
SUB Example_SetStrokeJoin (BYVAL hdc AS HDC)

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

   ' // Create a path and add two lines to it
   DIM capPath AS GpPath PTR
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   DIM pts(0 TO 2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipAddPathLine2(capPath, @pts(0), 3)

   ' // Create a CustomLineCap object with base cap LineCapRound
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapRound, 0.0, @custCap)

   ' // Set the stroke join style to LineJoinBevel
   hStatus = GdipSetCustomLineCapStrokeJoin(custCap, LineJoinBevel)

   ' // Create an orange Pen with width 5.0
   DIM strokeJoinPen AS GpPen PTR
   hStatus = GdipCreatePen1(GDIP_ARGB(255, 200, 150, 0), 5.0, UnitWorld, @strokeJoinPen)

   ' // Assign the custom cap as the end cap of the pen
   hStatus = GdipSetPenCustomEndCap(strokeJoinPen, custCap)

   ' // Draw a line using the pen
   hStatus = GdipDrawLine(graphics, strokeJoinPen, 0, 0, 200, 200)

   ' // Clean up resources
   IF strokeJoinPen THEN GdipDeletePen(strokeJoinPen)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================

```
---

## GdipSetCustomLineCapWidthScale

Sets the value of the scale width. This is the amount to scale the custom line cap relative to the width of the **Pen** used to draw lines. The default value of 1.0 does not scale the line cap.

```
FUNCTION GdipSetCustomLineCapWidthScale (BYVAL customCap AS GpCustomLineCap PTR, _
   BYVAL widthScale AS REAL) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *customCap* | [in] Pointer to a **CustomLineCap** object. |
| *widthScale* | [in] Single precision value that specifies the scaling factor that will be used to scale the line width. |

#### Example

The following example creates a CustomLineCap and sets the width scale. It then assigns the custom cap to a Pen, and draws a line.

```
' ========================================================================================
SUB Example_SetWidthScale (BYVAL hdc AS HDC)

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

   ' // Create a path and add two lines to it.
   DIM capPath AS GpPath PTR
   DIM pointsf(2) AS GpPointF = {(-15, -15), (0, 0), (15, -15)}
   hStatus = GdipCreatePath(FillModeAlternate, @capPath)
   hStatus = GdipAddPathLine2(capPath, @PointsF(0), 3)

  ' // Create a CustomLineCap
   DIM custCap AS GpCustomLineCap PTR
   hStatus = GdipCreateCustomLineCap(NULL, capPath, LineCapFlat, 0, @custCap)

   ' // // Set the width scale for custCap.
   hStatus = GdipSetCustomLineCapWidthScale(custCap, 3)

   ' // Assign custCap to a Pen object and draw a line.
   DIM widthScalePen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_VIOLET, 1.7, UnitWorld, @widthScalePen)
   hStatus = GdipSetPenCustomEndCap(widthScalePen, custCap)
   hStatus = GdipDrawLine(graphics, widthScalePen, 0, 0, 200, 200)

   ' // Cleanup
   IF widthScalePen THEN GdipDeletePen(widthScalePen)
   IF custCap THEN GdipDeleteCustomLineCap(custCap)
   IF capPath THEN GdipDeletePath(capPath)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
