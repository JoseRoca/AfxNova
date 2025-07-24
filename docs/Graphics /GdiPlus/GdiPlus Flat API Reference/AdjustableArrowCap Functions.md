The **AdjustableArrowCap** functions build a line cap that looks like an arrow.

Remarks

The middle inset is the number of units that the midpoint of the base shifts towards the vertex. A middle inset of zero results in no shift —the base is a straight line, giving the arrow a triangular shape. A positive (greater than zero) middle inset results in a shift the specified number of units toward the vertex — the base is an arrow shape that points toward the vertex, giving the arrow cap a V-shape. A negative (less than zero) middle inset results in a shift the specified number of units away from the vertex — the base becomes an arrow shape that points away from the vertex, giving the arrow either a diamond shape (if the absolute value of the middle inset is equal to the height) or distorted diamond shape. If the middle inset is equal to or greater than the height of the arrow cap, the cap does not appear at all. The value of the middle inset affects the arrow cap only if the arrow cap is filled.

---

| Name       | Description |
| ---------- | ----------- |
| [GdipCreateAdjustableArrowCap](#gdipcreateadjustablearrowcap) | Creates an adjustable arrow line cap with the specified height and width. |

## GdipCreateAdjustableArrowCap

Creates an adjustable arrow line cap with the specified height and width. The arrow line cap can be filled or nonfilled. The middle inset defaults to zero.

C++ Flat function
```
GpStatus GdipCreateAdjustableArrowCap (REAL height, REAL width, BOOL isFilled, GpAdjustableArrowCap **cap);
```
FB flat function
```
function GdipCreateAdjustableArrowCap (byval as REAL, byval as REAL, byval as BOOL, _
   byval as GpAdjustableArrowCap ptr ptr) as GpStatus
```
C++ wrapper method
```
AdjustableArrowCap(IN REAL height, IN REAL width, IN BOOL isFilled = TRUE)
```
FB wrapper method
```
CONSTRUCTOR CGpAdjustableArrowCap (BYVAL nHeight AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL bIsFilled AS BOOL = CTRUE)
```


---

Flat function

GpStatus WINGDIPAPI GdipGetAdjustableArrowCapHeight (GpAdjustableArrowCap* cap, REAL* height)

c++ Wrapper method

REAL GetHeight() const

FB wrapper method

FUNCTION GetHeight () AS SINGLE

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
