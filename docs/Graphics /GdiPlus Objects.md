# GdiPlus Objects

In Microsoft documentation, GDI+ objects—such as GpGraphics, GpBitmap, GpBrush, GpPen, etc.—are described as abstract "objects" manipulated via Flat API functions. However, these are not true objects in the object-oriented sense. They are actually opaque pointers to undocumented C-style structures.

Each GDI+ object is a pointer to a structure with a well-defined memory layout. These structures contain fields for internal state, handles, flags, and configuration data. The Flat API functions (e.g. GdipSetPenWidth, GdipGetImageBounds) are essentially accessors and mutators—they read or write values to these internal fields.

For example, the GpCustomLineCap "object", is a pointer to the following structure:

```
typedef struct {
    CustomLineCapType type;
    GpPathData pathdata;
    BOOL fill;
    GpLineCap cap;
    REAL inset;
    GpLineJoin join;
    REAL scale;
} GpCustomLineCap;
```

So when you call:

```
GdipSetCustomLineCapWidthScale(cap, 1.5)
```

…it’s just setting cap->scale = 1.5.

Big APIs like the Windows GDI+ flat API—tend to expose enormous functionality through low‑level procedural calls. They are powerful, but also verbose, error‑prone, and difficult to use safely. Every object must be created manually, destroyed manually, passed as a pointer, and handled with strict lifetime discipline. A single missing cleanup call or an incorrect pointer cast can destabilize an entire application.

`AfxNova` takes a different approach. Instead of building a heavy framework with deep inheritance trees, virtual methods, and opaque abstractions, `AfxNova` wraps these large APIs using tiny, lightweight classes—each one focused on a single GDI+ object.

The philosophy is simple:

1. Wrap the object, not the API

Each GDI+ object—GpGraphics, GpBitmap, GpPen, GpBrush, etc.—is represented by a small class whose only job is to manage:
* creation
* destruction
* ownership
* pointer access

The class does not attempt to replace the API. It simply makes the API safe and pleasant to use.

2. Automatic lifetime management

Every class uses RAII semantics:
* the constructor creates the underlying GDI+ object
* the destructor frees it automatically
* This eliminates leaks, double frees, and forgotten cleanup calls—problems that plague raw GDI+ code.

3. Full transparency and zero lock‑in

`AfxNova` never hides the underlying pointer. Every class exposes:
* *object → the raw GDI+ pointer
* @object → the address of the pointer (for functions that require GpObject**)

This means the developer can always call the flat API directly whenever needed. Nothing is trapped behind an abstraction.

4. No inheritance, no frameworks, no ceremony

Each class is self‑contained. There is no base class, no virtual destructor, no hierarchy, no “framework rules” to follow.

This keeps the system:
* predictable
* easy to read
* easy to maintain
* easy to extend

And it avoids the complexity that often appears in Microsoft’s own wrappers.

5. A consistent, minimal interface

All classes follow the same pattern:
* a constructor that initializes the object
* a destructor that frees it
* an operator to expose the pointer
* optional helper methods for convenience

This uniformity makes the entire API intuitive. Once you learn one class, you understand them all.

6. 100% compatibility with the flat API

Because the classes are thin wrappers, they remain fully compatible with:
* all GDI+ functions
* all Win32 functions
* mixed GDI/GDI+ rendering
+ custom drawing pipelines
+ user‑defined message loops
+ any external library that expects raw pointers

This is not a framework—it is a safety layer.

The result is a system that preserves the full power of the underlying API while eliminating its friction. Developers get:
* the safety of RAII
* the clarity of object‑oriented syntax
* the transparency of raw pointers
* the flexibility of procedural APIs
* the simplicity of tiny, focused classes

`AfxNova` does not try to redesign GDI+. It simply makes GDI+ usable.

The file `AfxGdipObjects.inc`, contains the complete set of these tiny classes for all major GDI+ objects. Each class is small, predictable, and easy to understand—yet together they form a robust, safe, and modern interface to a large and complex graphics subsystem.

---

### GdiPlusObjects
```
' ========================================================================================
' Constructor - Initializes GDI+
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusObjects
   DIM StartupInput AS GdiplusStartupInput
   StartupInput.GdiplusVersion = 1   ' Version must be 1
   GdiplusStartup(@m_token, @StartupInput, NULL)
   GDIP_DP("token: " & WSTR(m_token))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Dstructor - Shutdowns GDI+
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusObjects
   GDIP_DP("token: " & WSTR(m_token))
   IF m_token THEN GdiplusShutdown(m_token)
END DESTRUCTOR
' ========================================================================================
```
---

### GdiPlusCustomLLineCap

```
' ========================================================================================
TYPE GdiPlusCustomLineCap
Public:
   m_CustomLineCap AS GpCustomLineCap PTR
   DECLARE CONSTRUCTOR (BYVAL customCap AS GpPath PTR)
   DECLARE CONSTRUCTOR (BYVAL fillPath AS GpPath PTR, BYVAL strokePath AS GpPath PTR, _
           BYVAL baseCap AS GpLineCap = LinecapFlat, BYVAL baseInset AS SINGLE = 0.0)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpCustomLineCap PTR PTR
   DECLARE OPERATOR CAST () AS GpCustomLineCap PTR
END TYPE
' ========================================================================================

' ========================================================================================
' Creates a CustomLineCap
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusCustomLineCap (BYVAL customCap AS GpPath PTR)
   IF customCap THEN SetLastError(GdipCloneCustomLineCap(customCap, @m_CustomLineCap))
   GDIP_DP("Clone - " & WSTR(m_CustomLineCap))
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Creates a CustomLineCap object from a fill path and a stroke path.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusCustomLineCap (BYVAL fillPath AS GpPath PTR, BYVAL strokePath AS GpPath PTR, _
BYVAL baseCap AS GpLineCap = LinecapFlat, BYVAL baseInset AS SINGLE = 0.0)
   SetLastError(GdipCreateCustomLineCap(fillPath, strokePath, baseCap, baseinset, @m_CustomLineCap))
   GDIP_DP("m_pCustomLineCap: " & WSTR(m_CustomLineCap))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusCustomLineCap
   GDIP_DP("m_pCustomLineCap: " & WSTR(m_CustomLineCap))
   IF m_CustomLineCap THEN SetLastError(GdipDeleteCustomLineCap(m_CustomLineCap))
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpCustomLineCap object
' ========================================================================================
PRIVATE OPERATOR * (BYREF cap AS GdiPlusCustomLineCap) AS GpCustomLineCap PTR
   RETURN cap.m_CustomLineCap
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpCustomLineCap object
' ========================================================================================
PRIVATE OPERATOR GdiPlusCustomLineCap.@ () AS GpCustomLineCap PTR PTR
   RETURN @m_CustomLineCap
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusCustomLineCap.CAST () AS GpCustomLineCap PTR
   RETURN m_CustomLineCap
END OPERATOR
' ========================================================================================
```
---

### GdiPlusAdjustableArrowCap

```
' ========================================================================================
TYPE GdiPlusAdjustableArrowCap
Public:
   m_AdjustableArrowCap AS GpAdjustableArrowCap PTR
   DECLARE CONSTRUCTOR (BYVAL cap AS GpAdjustableArrowCap PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL nHeight AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL bIsFilled AS BOOL = CTRUE)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpAdjustableArrowCap PTR PTR
   DECLARE OPERATOR CAST () AS GpAdjustableArrowCap PTR
END TYPE
' ========================================================================================

' ========================================================================================
' Creates a AdjustableArrowCap
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusAdjustableArrowCap (BYVAL cap AS GpAdjustableArrowCap PTR = NULL)
   GDIP_DP(WSTR(cap))
   IF cap THEN SetLastError(GdipCloneCustomLineCap(cap, @m_AdjustableArrowCap))
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Creates an adjustable arrow line cap with the specified height and width. The arrow
' line cap can be filled or nonfilled. The middle inset defaults to zero.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusAdjustableArrowCap (BYVAL nHeight AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL bIsFilled AS BOOL = CTRUE)
   SetLastError(GdipCreateAdjustableArrowCap(nHeight, nWidth, bIsFilled, @m_AdjustableArrowCap))
   GDIP_DP("m_AdjustableArrowCap: " & WSTR(m_AdjustableArrowCap))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusAdjustableArrowCap
   GDIP_DP("m_AdjustableArrowCap: " & WSTR(m_AdjustableArrowCap))
   IF m_AdjustableArrowCap THEN SetLastError(GdipDeleteCustomLineCap(m_AdjustableArrowCap))
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpAdjustableArrowCap object
' ========================================================================================
PRIVATE OPERATOR * (BYREF cap AS GdiPlusAdjustableArrowCap) AS GpAdjustableArrowCap PTR
   RETURN cap.m_AdjustableArrowCap
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpAdjustableArrowCap object
' ========================================================================================
PRIVATE OPERATOR GdiPlusAdjustableArrowCap.@ () AS GpAdjustableArrowCap PTR PTR
   RETURN @m_AdjustableArrowCap
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusAdjustableArrowCap.CAST () AS GpAdjustableArrowCap PTR
   RETURN m_AdjustableArrowCap
END OPERATOR
' ========================================================================================
```
---

### GdiPlusBrush

```
' ========================================================================================
TYPE GdiPlusBrush
Public:
   m_Brush AS GpBrush PTR
   DECLARE CONSTRUCTOR (BYVAL brush AS GpBrush PTR = NULL)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpBrush PTR PTR
   DECLARE OPERATOR CAST () AS GpBrush PTR
END TYPE
' ========================================================================================

' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBrush (BYVAL brush AS GpBrush PTR = NULL)
   GDIP_DP(WSTR(brush))
   IF brush THEN SetLastError(GdipCloneBrush(brush, @m_brush))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusBrush
   GDIP_DP("m_Brush: " & WSTR(m_Brush))
   IF m_Brush THEN GdipDeleteBrush(m_Brush)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpBrush object
' ========================================================================================
PRIVATE OPERATOR * (BYREF brush AS GdiPlusBrush) AS GpBrush PTR
   RETURN brush.m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpBrush object
' ========================================================================================
PRIVATE OPERATOR GdiPlusBrush.@ () AS GpBrush PTR PTR
   RETURN @m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusBrush.CAST () AS GpBrush PTR
   RETURN m_Brush
END OPERATOR
' ========================================================================================
```
