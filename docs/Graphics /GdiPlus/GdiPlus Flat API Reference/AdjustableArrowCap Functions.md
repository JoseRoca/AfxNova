The **AdjustableArrowCap** functions build a line cap that looks like an arrow.

---

| Name       | Description |
| ---------- | ----------- |
| [GdipCreateAdjustableArrowCap](#gdipcreateadjustablearrowcap) | Creates an adjustable arrow line cap with the specified height and width. |
| [GdipGetAdjustableArrowCapHeight](#gdipgetadjustablearrowcapHeight) | Gets the height of the arrow cap. |

## GdipCreateAdjustableArrowCap

Creates an adjustable arrow line cap with the specified height and width. The arrow line cap can be filled or nonfilled. The middle inset defaults to zero.

C++ Flat function
```
GpStatus GdipCreateAdjustableArrowCap (REAL height, REAL width, BOOL isFilled, GpAdjustableArrowCap **cap);
```
FreeBASIC flat function
```
FUNCTION GdipCreateAdjustableArrowCap (BYVAL height AS SINGLE, BYVAL width AS SINGLE, _
   BYVAL isFilled AS BOOL, BYVAL pCap AS GpAdjustableArrowCap PTR PTR) AS GpStatus
```
PowerBASIC function
```
FUNCTION GdipCreateAdjustableArrowCap (BYVAL height AS SINGLE, BYVAL width AS SINGLE, _
    BYVAL isFilled AS LONG, BYREF cap AS DWORD) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *height* | [in] Single precision number that specifies the length, in units, of the arrow from its base to its point. |
| *width* | [in] Single precision number that specifies the distance, in units, between the corners of the base of the arrow. |
| *isFilled* | [in] Boolean value (TRUE or FALSE) that specifies whether the arrow is filled. |
| *cap* | [out] Pointer to a variable that receives a pointer to the new created adjustable arrow line cap. |

#### Remarks

The middle inset is the number of units that the midpoint of the base shifts towards the vertex. A middle inset of zero results in no shift — the base is a straight line, giving the arrow a triangular shape. A positive (greater than zero) middle inset results in a shift the specified number of units toward the vertex — the base is an arrow shape that points toward the vertex, giving the arrow cap a V-shape. A negative (less than zero) middle inset results in a shift the specified number of units away from the vertex — the base becomes an arrow shape that points away from the vertex, giving the arrow either a diamond shape (if the absolute value of the middle inset is equal to the height) or distorted diamond shape. If the middle inset is equal to or greater than the height of the arrow cap, the cap does not appear at all. The value of the middle inset affects the arrow cap only if the arrow cap is filled. The middle inset defaults to zero when an **AdjustableArrowCap** object is constructed.

#### PowerBasic Example

The following example creates two **AdjustableArrowCap objects**, *pArrowCapStart* and *pArrowCapEnd*, and sets the fill mode to TRUE. The code then creates a **Pen** object and assigns *pArrowCapStart* as the starting line cap for this **Pen** object and *pArrowCapEnd* as the ending line cap. Next, draws a line.
```

SUB GDIP_CreateAdjustableArrowCap (BYVAL hdc AS DWORD)

   LOCAL hStatus AS LONG
   LOCAL pGraphics AS DWORD
   LOCAL pArrowPen AS DWORD
   LOCAL pArrowCapStart AS DWORD
   LOCAL pArrowCapEnd AS DWORD
   LOCAL fillstate AS LONG

   hStatus = GdipCreateFromHDC(hdc, pGraphics)

   ' // Create an AdjustableArrowCap that is filled.
   hStatus = GdipCreateAdjustableArrowCap(10, 10, %TRUE, pArrowCapStart)
   ' // Create an AdjustableArrowCap that is not filled.
   hStatus = GdipCreateAdjustableArrowCap(15, 15, %FALSE, pArrowCapEnd)

   ' // Create a Pen.
   hStatus = GdipCreatePen1(GDIP_ARGB(255, 0, 0, 0), 1.0!, %UnitWorld, pArrowPen)

   ' Assign pArrowStart as the start cap.
   hStatus = GdipSetPenCustomStartCap(pArrowPen, pArrowCapStart)
   ' Assign pArrowEnd as the end cap.
   hStatus = GdipSetPenCustomEndCap(pArrowPen, pArrowCapEnd)

   ' // Draw a line using pArrowPen.
   hStatus = GdipDrawLineI(pGraphics, pArrowPen, 0, 0, 100, 100)

   ' // Cleanup
   IF pArrowCapStart THEN GdipDeleteCustomLineCap(pArrowCapStart)
   IF pArrowCapEnd THEN GdipDeleteCustomLineCap(pArrowCapEnd)
   IF pArrowPen THEN GdipDeletePen(pArrowPen)
   IF pGraphics THEN GdipDeleteGraphics(pGraphics)

END SUB
```
---

## GdipGetAdjustableArrowCapHeight

Gets the height of the arrow cap. The height is the distance from the base of the arrow to its vertex.

C++ Flat function
```
GpStatus GdipGetAdjustableArrowCapHeight(GpAdjustableArrowCap* cap, REAL* height);
```
FreeBASIC flat function
```
FUNCTION GdipGetAdjustableArrowCapHeight (BYVAL cap AS GpAdjustableArrowCap PTR, _
   BYVAL height AS SINGLE PTR) AS GpStatus
```
PowerBASIC flat Function
```
FUNCTION GdipGetAdjustableArrowCapHeight (BYVAL cap AS DWORD, BYREF height AS SINGLE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *cap* | [in] Pointer to the arrow cap. |
| *height* | [out] Pointer to a single precision variable that receives a value that indicates the height, in units, of the arrow cap. |

#### C++ Example
```

VOID Example_GetHeight(HDC hdc)
{
   Graphics graphics(hdc);

   // Create an AdjustableArrowCap with a height of 10 pixels.
   AdjustableArrowCap myArrow(10, 10, true);

   // Create a Pen, and assign myArrow as the end cap.
   Pen arrowPen(Color(255, 0, 0, 0));
   arrowPen.SetCustomEndCap(&myArrow);
   
   // Draw a line using arrowPen.
   graphics.DrawLine(&arrowPen, Point(0, 20), Point(100, 20));

   // Create a second arrow cap using the height of the first one.
   AdjustableArrowCap otherArrow(myArrow.GetHeight(), 20, true);

   // Assign the new arrow cap as the end cap for arrowPen.
   arrowPen.SetCustomEndCap(&otherArrow);

   // Draw a line using arrowPen.
   graphics.DrawLine(&arrowPen, Point(0, 55), Point(100, 55));
}
```
#### PowerBasic Example

The following example creates an AdjustableArrowCap, pMyArrowCap, and sets the height of the cap. The code then creates a Pen, assigns pMyArrowCap as the ending line cap for this Pen, and draws a capped line. Next, the code gets the height of the arrow cap, creates a new arrow cap with height equal to the height of pMyArrowCap, assigns the new arrow cap as the ending line cap for the Pen, and draws another capped line.

```
SUB GDIP_GetHeight(BYVAL hdc AS DWORD)

   LOCAL hStatus AS LONG
   LOCAL pGraphics AS DWORD
   LOCAL pArrowPen AS DWORD
   LOCAL pMyArrowCap AS DWORD
   LOCAL pOtherArrowCap AS DWORD
   LOCAL nHeight AS SINGLE

   hStatus = GdipCreateFromHDC(hdc, pGraphics)

   ' // Create an AdjustableArrowCap with a height of 10 pixels.
   hStatus = GdipCreateAdjustableArrowCap(10, 10, %TRUE, pMyArrowCap)

   ' // Create a Pen, and assign pMyArrowCap as the end cap.
   hStatus = GdipCreatePen1(GDIP_ARGB(255, 0, 0, 0), 1.0!, %UnitWorld, pArrowPen)
   hStatus = GdipSetPenCustomEndCap(pArrowPen, pMyArrowCap)

   ' // Draw a line using pArrowPen.
   hStatus = GdipDrawLineI(pGraphics, pArrowPen, 0, 20, 100, 20)

   ' // Get the height of the arrow cap.
   hStatus = GdipGetAdjustableArrowCapHeight(pMyArrowCap, nHeight)

   ' // Create a second arrow cap using the height of the first one.
   hStatus = GdipCreateAdjustableArrowCap(nHeight, 20, %TRUE, pOtherArrowCap)

   ' // Assign the new arrow cap as the end cap for pArrowPen.
   hStatus = GdipSetPenCustomEndCap(pArrowPen, pOtherArrowCap)

   ' // Draw a line using pArrowPen.
   hStatus = GdipDrawLineI(pGraphics, pArrowPen, 0, 55, 100, 55)

   ' // Cleanup
   IF pOtherArrowCap THEN GdipDeleteCustomLineCap(pOtherArrowCap)
   IF pMyArrowCap THEN GdipDeleteCustomLineCap(pMyArrowCap)
   IF pArrowPen THEN GdipDeletePen(pArrowPen)
   IF pGraphics THEN GdipDeleteGraphics(pGraphics)

END SUB
```
---





Flat function

GpStatus WINGDIPAPI GdipSetAdjustableArrowCapHeight (GpAdjustableArrowCap* cap, REAL height)

C++ Wrapper method

Status SetHeight(IN REAL height)

FB wrapper method

FUNCTION SetHeight (BYVAL nHeight AS SINGLE) AS GpStatus

---

Flat function

GpStatus WINGDIPAPI GdipGetAdjustableArrowCapWidth (GpAdjustableArrowCap* cap, REAL* width)

C++ Wrapper method

REAL GetWidth() const

FB wrapper method

FUNCTION GetWidth () AS SINGLE

---

Flat function

GpStatus WINGDIPAPI GdipSetAdjustableArrowCapWidth (GpAdjustableArrowCap* cap, REAL width)

C++ Wrapper method

Status SetWidth(IN REAL width)

FB wrapper method

FUNCTION SetWidth (BYVAL nWidth AS SINGLE) AS GpStatus

---

Flat function

GpStatus WINGDIPAPI GdipGetAdjustableArrowCapMiddleInset (GpAdjustableArrowCap* cap, REAL* middleInset)

C++ Wrapper method

REAL GetMiddleInset() const

FB Wrapper method

FUNCTION GetMiddleInset () AS SINGLE

---

Flat function

GpStatus WINGDIPAPI GdipSetAdjustableArrowCapMiddleInset (GpAdjustableArrowCap* cap, REAL middleInset)

C++ Wrapper method

Status SetMiddleInset(IN REAL middleInset)

FB Wrapper method

FUNCTION SetMiddleInset (BYVAL middleInset AS SINGLE) AS GpStatus

---

Flat function

GpStatus WINGDIPAPI GdipSetAdjustableArrowCapFillState (GpAdjustableArrowCap* cap, BOOL fillState)

C++ Wrapper method

Status SetFillState(IN BOOL isFilled)

PB wrapper method

FUNCTION SetFillState (BYVAL bIsFilled AS BOOL) AS GpStatus

---

Flat function

GpStatus WINGDIPAPI GdipGetAdjustableArrowCapFillState (GpAdjustableArrowCap* cap, BOOL* fillState)

C++ Wrapper method

BOOL IsFilled() const

PB wrapper method

FUNCTION IsFilled () AS BOOLEAN

---
