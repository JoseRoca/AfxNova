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
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Dstructor - Shutdowns GDI+
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusObjects
   IF m_token THEN GdiplusShutdown(m_token)
END DESTRUCTOR
' ========================================================================================
```
---

### GdiPlusCustomLLineCap

A line cap defines the style of graphic used to draw the ends of a line. It can be various shapes, such as a square, circle, or diamond. A custom line cap is defined by the path that draws it. The path is drawn by using a Pen object to draw the outline of a shape or by using a Brush object to fill the interior. The cap can be used on either or both ends of the line. Spacing can be adjusted between the end caps and the line.

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
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Creates a CustomLineCap object from a fill path and a stroke path.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusCustomLineCap (BYVAL fillPath AS GpPath PTR, BYVAL strokePath AS GpPath PTR, _
BYVAL baseCap AS GpLineCap = LinecapFlat, BYVAL baseInset AS SINGLE = 0.0)
   SetLastError(GdipCreateCustomLineCap(fillPath, strokePath, baseCap, baseinset, @m_CustomLineCap))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusCustomLineCap
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

Builds a line cap that looks like an arrow.

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
   IF cap THEN SetLastError(GdipCloneCustomLineCap(cap, @m_AdjustableArrowCap))
END CONSTRUCTOR
' ========================================================================================

' =====================================================================================
' Creates an adjustable arrow line cap with the specified height and width. The arrow
' line cap can be filled or nonfilled. The middle inset defaults to zero.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusAdjustableArrowCap (BYVAL nHeight AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL bIsFilled AS BOOL = CTRUE)
   SetLastError(GdipCreateAdjustableArrowCap(nHeight, nWidth, bIsFilled, @m_AdjustableArrowCap))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusAdjustableArrowCap
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

A Brush object is used to paint the interior of graphics shapes, such as rectangles, ellipses, pies, polygons, and paths.

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
   IF brush THEN SetLastError(GdipCloneBrush(brush, @m_brush))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusBrush
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
---

### GdiPlusSolidBrush

Defines a solid color Brush object. A Brush object is used to fill in shapes similar to the way a paint brush can paint the inside of a shape.

```
' ========================================================================================
TYPE GdiPlusSolidBrush
Public:
   m_Brush AS GpSolidBrush PTR
   DECLARE CONSTRUCTOR (BYVAL brush AS GpSolidBrush PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL colour AS ARGB = &hFF000000)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpSolidBrush PTR PTR
   DECLARE OPERATOR CAST () AS GpSolidBrush PTR
END TYPE
' ========================================================================================

' ========================================================================================
' Creates a brush from another brush.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusSolidBrush (BYVAL brush AS GpSolidBrush PTR = NULL)
   IF brush THEN SetLastError(GdipCloneBrush(brush, @m_brush))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a solid brush filled with the specified color.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusSolidBrush (BYVAL colour AS ARGB = &hFF000000)
   SetLastError(GdipCreateSolidFill(colour, @m_Brush))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusSolidBrush
   IF m_Brush THEN GdipDeleteBrush(m_Brush)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpSolidBrush object
' ========================================================================================
PRIVATE OPERATOR * (BYREF brush AS GdiPlusSolidBrush) AS GpSolidBrush PTR
   RETURN brush.m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpBrush object
' ========================================================================================
PRIVATE OPERATOR GdiPlusSolidBrush.@ () AS GpSolidBrush PTR PTR
   RETURN @m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusSolidBrush.CAST () AS GpSolidBrush PTR
   RETURN m_Brush
END OPERATOR
' ========================================================================================
```
---

### GdiPlusHatchBrush

An HatchBrush defines a rectangular brush with a hatch style, a foreground color, and a background color. There are six hatch styles. The foreground color defines the color of the hatch lines; the background color defines the color over which the hatch lines are drawn

```
' ========================================================================================
TYPE GdiPlusHatchBrush
Public:
   m_Brush AS GpHatchBrush PTR
   DECLARE CONSTRUCTOR (BYVAL brush AS GpHatchBrush PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL nHatchStyle AS HatchStyle, BYVAL foreColor AS ARGB, BYVAL backColor AS ARGB = &HFF000000)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpHatchBrush PTR PTR
   DECLARE OPERATOR CAST () AS GpHatchBrush PTR
END TYPE
' ========================================================================================

' ========================================================================================
' Creates a brush from another brush.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusHatchBrush (BYVAL brush AS GpHatchBrush PTR = NULL)
   IF brush THEN SetLastError(GdipCloneBrush(brush, @m_brush))
END CONSTRUCTOR
' ========================================================================================
' =====================================================================================
' Creates a HatchBrush object based on a hatch style, a foreground color, and a background color.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusHatchBrush (BYVAL nHatchStyle AS HatchStyle, BYVAL foreColor AS ARGB, BYVAL backColor AS ARGB = &HFF000000)
   SetLastError(GdipCreateHatchBrush(nHatchStyle, foreColor, backColor, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusHatchBrush
   IF m_Brush THEN GdipDeleteBrush(m_Brush)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpHatchBrush object
' ========================================================================================
PRIVATE OPERATOR * (BYREF brush AS GdiPlusHatchBrush) AS GpHatchBrush PTR
   RETURN brush.m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpHatchBrush object
' ========================================================================================
PRIVATE OPERATOR GdiPlusHatchBrush.@ () AS GpHatchBrush PTR PTR
   RETURN @m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusHatchBrush.CAST () AS GpHatchBrush PTR
   RETURN m_Brush
END OPERATOR
' ========================================================================================
```
---

### GdiPlusLinearGradientBrush

The LinearGradientBrush functions allow to paint a color gradient in which the color changes evenly from the starting boundary line of the linear gradient brush to the ending boundary line of the linear gradient brush. The boundary lines of a linear gradient brush are two parallel straight lines. The color gradient is perpendicular to the boundary lines of the linear gradient brush, changing gradually across the stroke from the starting boundary line to the ending boundary line. The color gradient has one color at the starting boundary line and another color at the ending boundary line.

```
' ========================================================================================
TYPE GdiPlusLinearGradientBrush
Public:
   m_Brush AS GpLinearGradientBrush PTR
   DECLARE CONSTRUCTOR (BYVAL point1 AS GpPointF PTR, BYVAL point2 AS GpPointF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL wrap AS WrapMode = WrapModeTile)
   DECLARE CONSTRUCTOR (BYVAL point1 AS GpPoint PTR, BYVAL point2 AS GpPoint PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL wrap AS WrapMode = WrapModeTile)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL nMode AS LinearGradientMode, BYVAL wrap AS WrapMode = WrapModeTile)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL nMode AS LinearGradientMode, BYVAL wrap AS WrapMode = WrapModeTile)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL, BYVAL wrap AS WrapMode = WrapModeTile)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL, BYVAL wrap AS WrapMode = WrapModeTile)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpLinearGradientBrush PTR PTR
   DECLARE OPERATOR CAST () AS GpLinearGradientBrush PTR
END TYPE
' ========================================================================================

' =====================================================================================
' Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusLinearGradientBrush (BYVAL point1 AS GpPointF PTR, BYVAL point2 AS GpPointF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL wrap AS WrapMode = WrapModeTile)
   SetLastError(GdipCreateLineBrush(point1, point2, color1, color2, wrap, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusLinearGradientBrush (BYVAL point1 AS GpPoint PTR, BYVAL point2 AS GpPoint PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL wrap AS WrapMode = WrapModeTile)
   SetLastError(GdipCreateLineBrushI(point1, point2, color1, color2, wrap, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusLinearGradientBrush (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL nMode AS LinearGradientMode, BYVAL wrap AS WrapMode = WrapModeTile)
   SetLastError(GdipCreateLineBrushFromRect(rc, color1, color2, nMode, wrap, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusLinearGradientBrush (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL nMode AS LinearGradientMode, BYVAL wrap AS WrapMode = WrapModeTile)
   SetLastError(GdipCreateLineBrushFromRectI(rc, color1, color2, nMode, wrap, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a LinearGradientBrush object from a set of boundary points and boundary colors.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusLinearGradientBrush (BYVAL rc AS GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL, BYVAL wrap AS WrapMode = WrapModeTile)
   SetLastError(GdipCreateLineBrushFromRectWithAngle(rc, color1, color2, angle, isAngleScalable, wrap, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusLinearGradientBrush (BYVAL rc AS GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL angle AS SINGLE, BYVAL isAngleScalable AS BOOL, BYVAL wrap AS WrapMode = WrapModeTile)
   SetLastError(GdipCreateLineBrushFromRectWithAngleI(rc, color1, color2, angle, isAngleScalable, wrap, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusLinearGradientBrush
   IF m_Brush THEN GdipDeleteBrush(m_Brush)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpLinearGradientBrush object
' ========================================================================================
PRIVATE OPERATOR * (BYREF brush AS GdiPlusLinearGradientBrush) AS GpLinearGradientBrush PTR
   RETURN brush.m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpLinearGradientBrush object
' ========================================================================================
PRIVATE OPERATOR GdiPlusLinearGradientBrush.@ () AS GpLinearGradientBrush PTR PTR
   RETURN @m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusLinearGradientBrush.CAST () AS GpLinearGradientBrush PTR
   RETURN m_Brush
END OPERATOR
' ========================================================================================
```
---

### GdiPlusPathGradientBrush

The PathGradientBrush functions allow to set the attributes of a color gradient that you can use to fill the interior of a path with a gradually changing color. A path gradient brush has a boundary path, a boundary color, a center point, and a center color. When you paint an area with a path gradient brush, the color changes gradually from the boundary color to the center color as you move from the boundary path to the center point.

```
' ========================================================================================
TYPE GdiPlusPathGradientBrush
Public:
   m_Brush AS GpPathGradientBrush PTR
   DECLARE CONSTRUCTOR (BYVAL path AS GpPath PTR)
   DECLARE CONSTRUCTOR (BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL wrapMode AS GpWrapMode = WrapModeClamp)
   DECLARE CONSTRUCTOR (BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL wrapMode AS GpWrapMode = WrapModeClamp)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpPathGradientBrush PTR PTR
   DECLARE OPERATOR CAST () AS GpPathGradientBrush PTR
END TYPE
' ========================================================================================

' =====================================================================================
' Creates a path gradient brush from a path
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusPathGradientBrush (BYVAL path AS GpPath PTR)
   SetLastError(GdipCreatePathGradientFromPath(path, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a path gradient brush
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusPathGradientBrush (BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL wrapMode AS GpWrapMode = WrapModeClamp)
   SetLastError(GdipCreatePathGradient(points, count, wrapMode, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusPathGradientBrush (BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL wrapMode AS GpWrapMode = WrapModeClamp)
   SetLastError(GdipCreatePathGradientI(points, count, wrapMode, @m_brush))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusPathGradientBrush
   IF m_Brush THEN GdipDeleteBrush(m_Brush)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpPathGradientBrush object
' ========================================================================================
PRIVATE OPERATOR * (BYREF brush AS GdiPlusPathGradientBrush) AS GpPathGradientBrush PTR
   RETURN brush.m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpPathGradientBrush object
' ========================================================================================
PRIVATE OPERATOR GdiPlusPathGradientBrush.@ () AS GpPathGradientBrush PTR PTR
   RETURN @m_Brush
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusPathGradientBrush.CAST () AS GpPathGradientBrush PTR
   RETURN m_Brush
END OPERATOR
' ========================================================================================
```
---

### GdiPlusTextureBrush

The TextureBrush object is a Brush object that contains an Image object that is used for the fill. The fill image can be transformed by using the local Matrix object contained in the Brush object.

```
' ========================================================================================
TYPE GdiPlusTextureBrush
Public:
   m_Texture AS GpTextureBrush PTR
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode = WrapModeTile)
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR, BYVAL dstRect AS GpRectF PTR, BYVAL imageAttributes AS GpImageAttributes PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR, BYVAL dstRect AS GpRect PTR, BYVAL imageAttributes AS GpImageAttributes PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR, BYVAL dstX AS SINGLE, BYVAL dstY AS SINGLE, BYVAL dstWidth AS SINGLE, BYVAL dstHeight AS SINGLE, BYVAL imageAttributes AS GpImageAttributes PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR, BYVAL dstX AS INT_, BYVAL dstY AS INT_, BYVAL dstWidth AS INT_, BYVAL dstHeight AS INT_, BYVAL imageAttributes AS GpImageAttributes PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstRect AS GpRect PTR)
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstRect AS GpRectF PTR)
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstX AS SINGLE, BYVAL dstY AS SINGLE, BYVAL dstWidth AS SINGLE, BYVAL dstHeight AS SINGLE)
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstX AS INT_, BYVAL dstY AS INT_, BYVAL dstWidth AS INT_, BYVAL dstHeight AS INT_)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpTextureBrush PTR PTR
   DECLARE OPERATOR CAST () AS GpTextureBrush PTR
END TYPE
' ========================================================================================

' =====================================================================================
' Creates a TextureBrush object based on an image and a wrap mode. The size of the brush
' defaults to the size of the image, so the entire image is used by the brush.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusTextureBrush (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode = WrapModeTile)
   SetLastError(GdipCreateTexture(image, nWrapmode, @m_Texture))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a TextureBrush object based on an image, a defining set of coordinates, and
' a set of image properties.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusTextureBrush (BYVAL image AS GpImage PTR, BYVAL dstRect AS GpRectF PTR, BYVAL imageAttributes AS GpImageAttributes PTR = NULL)
   SetLastError(GdipCreateTextureIA(image, imageAttributes, dstRect->x, dstRect->y, dstRect->Width, dstRect->Height, @m_Texture))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusTextureBrush (BYVAL image AS GpImage PTR, BYVAL dstRect AS GpRect PTR, BYVAL imageAttributes AS GpImageAttributes PTR = NULL)
   SetLastError(GdipCreateTextureIAI(image, imageAttributes, dstRect->x, dstRect->y, dstRect->Width, dstRect->Height, @m_Texture))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusTextureBrush (BYVAL image AS GpImage PTR, BYVAL dstX AS SINGLE, BYVAL dstY AS SINGLE, BYVAL dstWidth AS SINGLE, BYVAL dstHeight AS SINGLE, BYVAL imageAttributes AS GpImageAttributes PTR = NULL)
   SetLastError(GdipCreateTextureIA(image, imageAttributes, dstX, dstY, dstWidth, dstHeight, @m_Texture))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusTextureBrush (BYVAL image AS GpImage PTR, BYVAL dstX AS INT_, BYVAL dstY AS INT_, BYVAL dstWidth AS INT_, BYVAL dstHeight AS INT_, BYVAL imageAttributes AS GpImageAttributes PTR = NULL)
   SetLastError(GdipCreateTextureIA(image, imageAttributes, dstX, dstY, dstWidth, dstHeight, @m_Texture))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a TextureBrush object based on an image, a wrap mode, and a defining set of coordinates.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusTextureBrush (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstRect AS GpRect PTR)
   SetLastError(GdipCreateTexture2I(image, nWrapMode, dstRect->x, dstRect->y, dstRect->Width, dstRect->Height, @m_Texture))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusTextureBrush (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstRect AS GpRectF PTR)
   SetLastError(GdipCreateTexture2(image, nWrapMode, dstRect->x, dstRect->y, dstRect->Width, dstRect->Height, @m_Texture))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusTextureBrush (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstX AS SINGLE, BYVAL dstY AS SINGLE, BYVAL dstWidth AS SINGLE, BYVAL dstHeight AS SINGLE)
   SetLastError(GdipCreateTexture2(image, nWrapMode, dstX, dstY, dstWidth, dstHeight, @m_Texture))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusTextureBrush (BYVAL image AS GpImage PTR, BYVAL nWrapMode AS WrapMode, BYVAL dstX AS INT_, BYVAL dstY AS INT_, BYVAL dstWidth AS INT_, BYVAL dstHeight AS INT_)
   SetLastError(GdipCreateTexture2I(image, nWrapMode, dstX, dstY, dstWidth, dstHeight, @m_Texture))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusTextureBrush
   IF m_Texture THEN GdipDeleteBrush(m_Texture)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpTextureBrush object
' ========================================================================================
PRIVATE OPERATOR * (BYREF texture AS GdiPlusTextureBrush) AS GpTextureBrush PTR
   RETURN texture.m_Texture
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpTextureBrush object
' ========================================================================================
PRIVATE OPERATOR GdiPlusTextureBrush.@ () AS GpTextureBrush PTR PTR
   RETURN @m_Texture
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusTextureBrush.CAST () AS GpTextureBrush PTR
   RETURN m_Texture
END OPERATOR
' ========================================================================================
```
---

### GdiPlusImage

The Image functions allow loading and saving raster images (bitmaps) and vector images (metafiles). An Image object encapsulates a bitmap or a metafile and stores attributes that you can retrieve by calling various Get functions. You can construct Image objects from a variety of file types including BMP, ICON, GIF, JPEG, Exif, PNG, TIFF, WMF, and EMF.

```
' ========================================================================================
TYPE GdiPlusImage
Public:
   m_Image AS GpImage PTR
   DECLARE CONSTRUCTOR (BYVAL image AS GpImage PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL stream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL hInst AS HINSTANCE, BYREF wszImageName AS WSTRING)
   DECLARE OPERATOR CAST () AS GpImage PTR
   DECLARE OPERATOR @ () AS GpImage PTR PTR
   DECLARE DESTRUCTOR
   DECLARE SUB SetResolution (BYVAL graphics AS GpGraphics PTR, BYVAL dpiX AS SINGLE = 0, BYVAL dpiY AS SINGLE = 0)
END TYPE
' ========================================================================================

' ========================================================================================
' Creates and initializes an Image object from another Image object.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusImage (BYVAL image AS GpImage PTR)
   IF image THEN SetLastError(GdipCloneImage(image, @m_Image))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates an Image object based on a file.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusImage (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   IF useEmbeddedColorManagement THEN
      SetLastError(GdipLoadImageFromFileICM(pwszFileName, @m_Image))
   ELSE
      SetLastError(GdipLoadImageFromFile(pwszFileName, @m_Image))
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates an Image object based on a stream.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusImage (BYVAL stream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   IF useEmbeddedColorManagement THEN
      SetLastError(GdipLoadImageFromStreamICM(stream, @m_Image))
   ELSE
      SetLastError(GdipLoadImageFromStream(stream, @m_Image))
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates an Image object from a resource file.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusImage (BYVAL hInst AS HINSTANCE, BYREF wszImageName AS WSTRING)
   DIM stream AS IStream PTR = AfxImageStreamFromRes(hInst, wszImageName)
   IF stream = NULL THEN SetLastError(InvalidParameter): EXIT CONSTRUCTOR
   SetLastError(GdipLoadImageFromStream(stream, @m_Image))
   stream->lpVtbl->Release(stream)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusImage
   IF m_Image THEN SetLastError(GdipDisposeImage(m_Image))
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpImage object
' ========================================================================================
PRIVATE OPERATOR * (BYREF img AS GdiPlusImage) AS GpImage PTR
   RETURN img.m_Image
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusImage.CAST () AS GpImage PTR
   RETURN m_Image
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpImage object
' ========================================================================================
PRIVATE OPERATOR GdiPlusImage.@ () AS GpImage PTR PTR
   RETURN @m_Image
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE SUB GdiPlusImage.SetResolution (BYVAL graphics AS GpGraphics PTR, BYVAL dpiX AS SINGLE = 0, BYVAL dpiY AS SINGLE = 0)
   IF dpiX = 0 AND dpiY = 0 THEN
      GdipGetDpiX(graphics, @dpiX)
      GdipGetDpiY(graphics, @dpiY)
   END IF
   SetLastError(GdipBitmapSetResolution(m_Image, dpiX, dpiY))
END SUB
' ========================================================================================
```
---

### GdiPlusBitmap

The Bitmap functions expands on the capabilities of the Image functions by providing additional methods for creating and manipulating raster images.

```
' ========================================================================================
TYPE GdiPlusBitmap
Public:
   m_Bitmap AS GpBitmap PTR
   DECLARE CONSTRUCTOR (BYVAL bmp AS GpBitmap PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL x AS REAL, BYVAL y AS REAL, BYVAL nWidth AS REAL, BYVAL nHeight AS REAL, _
           BYVAL pxFormat AS PixelFormat, BYVAL srcBitmap AS GpBitmap PTR)
   DECLARE CONSTRUCTOR (BYVAL rcf AS GpRectF PTR, BYVAL pxFormat AS PixelFormat, BYVAL srcBitmap AS GpBitmap PTR)
   DECLARE CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL stream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL stride AS INT_, BYVAL pxFormat AS PixelFormat, BYVAL scan0 AS UBYTE PTR)
   DECLARE CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pxFormat AS PixelFormat)
   DECLARE CONSTRUCTOR (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL pTarget AS GpGraphics PTR)
   #ifdef IDirectDrawSurface7
   DECLARE CONSTRUCTOR (BYVAL surface AS IDirectDrawSurface7 PTR)
   #endif
   DECLARE CONSTRUCTOR (BYVAL gdiBitmapInfo AS BITMAPINFO PTR, BYVAL gdiBitmapData AS ANY PTR)
   DECLARE CONSTRUCTOR (BYVAL hbm AS HBITMAP, BYVAL hPal AS HPALETTE)
   DECLARE CONSTRUCTOR (BYVAL hicon AS HICON)
   DECLARE CONSTRUCTOR (BYVAL hInstance AS HINSTANCE, BYVAL pwszBitmapName AS WSTRING PTR)
   DECLARE CONSTRUCTOR (BYVAL inputBitmap AS GpBitmap PTR, BYVAL effect AS GpEffect PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpBitmap PTR PTR
   DECLARE OPERATOR CAST () AS GpBitmap PTR
   DECLARE SUB SetResolution (BYVAL graphics AS GpGraphics PTR, BYVAL dpiX AS SINGLE = 0, BYVAL dpiY AS SINGLE = 0)
END TYPE
' ========================================================================================

' ========================================================================================
' Creates and initializes a Bitmap object from another Bitmap object.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL bmp AS GpBitmap PTR = NULL)
   IF bmp THEN SetLastError(GdipCloneImage(bmp, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' =====================================================================================
' Creates a new Bitmap object by copying a portion of the specified bitmap.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL x AS REAL, BYVAL y AS REAL, BYVAL nWidth AS REAL, BYVAL nHeight AS REAL, _
BYVAL pxFormat AS PixelFormat, BYVAL srcBitmap AS GpBitmap PTR)
   SetLastError(GdipCloneBitmapArea(x, y, nWidth, nHeight, pxFormat, srcBitmap, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' =====================================================================================
' Creates a new Bitmap object by copying a portion of the specified bitmap.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL rcf AS GpRectF PTR, BYVAL pxFormat AS PixelFormat, BYVAL srcBitmap AS GpBitmap PTR)
   SetLastError(GdipCloneBitmapArea(rcf->x, rcf->y, rcf->Width, rcf->Height, pxFormat, srcBitmap, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates an Bitmap object based on a file.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   IF useEmbeddedColorManagement THEN
      SetLastError(GdipCreateBitmapFromFileICM(pwszFileName, @m_Bitmap))
   ELSE
      SetLastError(GdipCreateBitmapFromFile(pwszFileName, @m_Bitmap))
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates an Image object based on a stream.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL stream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
   IF useEmbeddedColorManagement THEN
      SetLastError(GdipCreateBitmapFromStream(stream, @m_Bitmap))
   ELSE
      SetLastError(GdipCreateBitmapFromStream(stream, @m_Bitmap))
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Bitmap object based on an array of bytes along with size and format information.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL stride AS INT_, BYVAL pxFormat AS PixelFormat, BYVAL scan0 AS UBYTE PTR)
   SetLastError(GdipCreateBitmapFromScan0(nWidth, nHeight, stride, pxFormat, scan0, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Bitmap object based on size and format information.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pxFormat AS PixelFormat)
   SetLastError(GdipCreateBitmapFromScan0(nWidth, nHeight, 0, pxFormat, NULL, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Bitmap object based on a Graphics object, a width, and a height.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG, BYVAL pTarget AS GpGraphics PTR)
   SetLastError(GdipCreateBitmapFromGraphics(nWidth, nHeight, pTarget, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Bitmap object based on a DirectDraw surface. The Bitmap object maintains a
' reference to the DirectDraw surface until the Bitmap object is deleted or goes out of scope.
' ========================================================================================
#ifdef IDirectDrawSurface7
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL surface AS IDirectDrawSurface7 PTR)
   SetLastError(GdipCreateBitmapFromDirectDrawSurface(surface, @m_Bitmap))
END CONSTRUCTOR
#endif
' ========================================================================================
' ========================================================================================
' Creates a Bitmap object based on a BITMAPINFO structure and an array of pixel data.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL gdiBitmapInfo AS BITMAPINFO PTR, BYVAL gdiBitmapData AS ANY PTR)
   SetLastError(GdipCreateBitmapFromGdiDib(gdiBitmapInfo, gdiBitmapData, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Bitmap object based on a handle to a Windows Windows Graphics Device
' Interface (GDI) bitmap and a handle to a GDI palette.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL hbm AS HBITMAP, BYVAL hPal AS HPALETTE)
   SetLastError(GdipCreateBitmapFromHBITMAP(hbm, hpal, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Bitmap object based on an icon.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL hicon AS HICON)
   SetLastError(GdipCreateBitmapFromHICON(hicon, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Bitmap object based on an application or DLL instance handle and the name of a bitmap.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL hInstance AS HINSTANCE, BYVAL pwszBitmapName AS WSTRING PTR)
   SetLastError(GdipCreateBitmapFromResource(hInstance, pwszBitmapName, @m_Bitmap))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a new Bitmap object by applying a specified effect to an existing Bitmap object.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusBitmap (BYVAL inputBitmap AS GpBitmap PTR, BYVAL effect AS GpEffect PTR)
   SetLastError(GdipBitmapCreateApplyEffect(@inputBitmap, 1, effect, NULL, NULL, @m_Bitmap, FALSE, NULL, NULL))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusBitmap
   IF m_Bitmap THEN SetLastError(GdipDisposeImage(m_Bitmap))
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpBitmap object
' ========================================================================================
PRIVATE OPERATOR * (BYREF bmp AS GdiPlusBitmap) AS GpBitmap PTR
   RETURN bmp.m_Bitmap
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpBitmap object
' ========================================================================================
PRIVATE OPERATOR GdiPlusBitmap.@ () AS GpBitmap PTR PTR
   RETURN @m_Bitmap
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusBitmap.CAST () AS GpBitmap PTR
   RETURN m_Bitmap
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE SUB GdiPlusBitmap.SetResolution (BYVAL graphics AS GpGraphics PTR, BYVAL dpiX AS SINGLE = 0, BYVAL dpiY AS SINGLE = 0)
   IF dpiX = 0 AND dpiY = 0 THEN
      GdipGetDpiX(graphics, @dpiX)
      GdipGetDpiY(graphics, @dpiY)
   END IF
   SetLastError(GdipBitmapSetResolution(m_Bitmap, dpiX, dpiY))
END SUB
' ========================================================================================
```
---

### GdiPlusImageAttributes

An ImageAttributes object contains information about how bitmap and metafile colors are manipulated during rendering. An ImageAttributes object maintains several color-adjustment settings, including color-adjustment matrices, grayscale-adjustment matrices, gamma-correction values, color-map tables, and color-threshold values.

```
' ========================================================================================
TYPE GdiPlusImageAttributes
Public:
   m_ImgAttr AS GpImageAttributes PTR
   DECLARE CONSTRUCTOR (BYVAL imgAttr AS GpImageAttributes PTR = NULL)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpImageAttributes PTR PTR
   DECLARE OPERATOR CAST () AS GpImageAttributes PTR
END TYPE
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusImageAttributes (BYVAL imgAttr AS GpImageAttributes PTR = NULL)
   IF imgAttr THEN
      SetLastError(GdipCloneImageAttributes(imgAttr, @m_ImgAttr))
   ELSE
      SetLastError(GdipCreateImageAttributes(@m_ImgAttr))
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusImageAttributes
   IF m_ImgAttr THEN GdipDisposeImageAttributes(m_ImgAttr)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpImageAttributes object
' ========================================================================================
PRIVATE OPERATOR * (BYREF imgAttr AS GdiPlusImageAttributes) AS GpImageAttributes PTR
   RETURN imgAttr.m_ImgAttr
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpImageAttributes object
' ========================================================================================
PRIVATE OPERATOR GdiPlusImageAttributes.@ () AS GpImageAttributes PTR PTR
   RETURN @m_ImgAttr
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusImageAttributes.CAST () AS GpImageAttributes PTR
   RETURN m_ImgAttr
END OPERATOR
' ========================================================================================
```
---

### GdiPlusEffect

Allows to apply effects and adjustments to bitmaps.

```
' ========================================================================================
TYPE GdiPlusEffect
Public:
   m_Effect AS GpEffect PTR
   m_auxDataSize AS INT_
   m_auxData AS ANY PTR
   m_useAuxData AS BOOLEAN
   DECLARE CONSTRUCTOR (BYVAL gd AS GUID)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpEffect PTR PTR
   DECLARE OPERATOR CAST () AS GpEffect PTR
END TYPE
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusEffect (BYVAL gd AS GUID)
   SetLastError GdipCreateEffect(gd, @m_Effect)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusEffect
   ' // auxData is allocated by ApplyEffect.
   IF m_auxData THEN GdipFree(m_auxData)
   ' // Delete the Effect object
   IF m_Effect THEN GdipDeleteEffect(m_Effect)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the CGpEffect object
' ========================================================================================
PRIVATE OPERATOR * (BYREF effect AS GdiPlusEffect) AS GpEffect PTR
   RETURN effect.m_Effect
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpEffect object
' ========================================================================================
PRIVATE OPERATOR GdiPlusEffect.@ () AS GpEffect PTR PTR
   RETURN @m_Effect
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusEffect.CAST () AS GpEffect PTR
   RETURN m_Effect
END OPERATOR
' ========================================================================================
```
---

### GdiPlusGraphics

The Graphics functions provide methods for drawing lines, curves, figures, images, and text. A Graphics object stores attributes of the display device and attributes of the items to be drawn.

```
' ========================================================================================
TYPE GdiPlusGraphics
Public:
   m_Graphics AS GpGraphics PTR
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC)
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpGraphics PTR PTR
   DECLARE OPERATOR CAST () AS GpGraphics PTR
   DECLARE FUNCTION DpiX () AS SINGLE
   DECLARE FUNCTION DpiY () AS SINGLE
   DECLARE FUNCTION DpiRatio () AS SINGLE
   DECLARE FUNCTION DpiRatioX () AS SINGLE
   DECLARE FUNCTION DpiRatioY () AS SINGLE
   DECLARE FUNCTION ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION ScaleTransform (BYVAL dpiRatio_ AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DECLARE FUNCTION ScaleTransform (BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
END TYPE
' ========================================================================================

' ========================================================================================
' * Creates a Graphics object that is associated with a specified device context.
' When you use these constructors to create a Graphics object, make sure that the Graphics
' object is deleted or goes out of scope before the device context is released.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusGraphics (BYVAL hdc AS HDC)
   SetLastError(GdipCreateFromHDC(hdc, @m_Graphics))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusGraphics (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE)
   SetLastError(GdipCreateFromHDC2(hdc, hDevice, @m_Graphics))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusGraphics
   IF m_Graphics THEN SetLastError(GdipDeleteGraphics(m_Graphics))
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpGraphics object
' ========================================================================================
PRIVATE OPERATOR * (BYREF graphics AS GdiPlusGraphics) AS GpGraphics PTR
   RETURN graphics.m_Graphics
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpGraphics object
' ========================================================================================
PRIVATE OPERATOR GdiPlusGraphics.@ () AS GpGraphics PTR PTR
   RETURN @m_Graphics
END OPERATOR
' ========================================================================================

' ========================================================================================
PRIVATE OPERATOR GdiPlusGraphics.CAST () AS GpGraphics PTR
   RETURN m_Graphics
END OPERATOR
' ========================================================================================

' ========================================================================================
' Get the DPI scaling
' ========================================================================================
PRIVATE FUNCTION GdiPlusGraphics.DpiX () AS SINGLE
   DIM _dpiX AS SINGLE
   SetLastError(GdipGetDpiX(m_Graphics, @_dpiX))
   RETURN _dpiX
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GdiPlusGraphics.DpiY () AS SINGLE
   DIM _dpiY AS SINGLE
   SetLastError(GdipGetDpiX(m_Graphics, @_dpiY))
   RETURN _dpiY
END FUNCTION
' ========================================================================================
' ========================================================================================
' Get the DPI scaling ratio
' ========================================================================================
PRIVATE FUNCTION GdiPlusGraphics.DpiRatio () AS SINGLE
   DIM _dpiX AS SINGLE
   SetLastError(GdipGetDpiX(m_Graphics, @_dpiX))
   DIM rxRatio AS SINGLE = _dpiX / 96
   RETURN rxRatio
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GdiPlusGraphics.DpiRatioX () AS SINGLE
   DIM _dpiX AS SINGLE
   SetLastError(GdipGetDpiX(m_Graphics, @_dpiX))
   DIM rxRatio AS SINGLE = _dpiX / 96
   RETURN rxRatio
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GdiPlusGraphics.DpiRatioY () AS SINGLE
   DIM _dpiY AS SINGLE
   SetLastError(GdipGetDpiY(m_Graphics, @_dpiY))
   DIM ryRatio AS SINGLE = _dpiY / 96
   RETURN ryRatio
END FUNCTION
' ========================================================================================
' ========================================================================================
' Updates this Graphic object's current transformation matrix with the product of itself
' and a scaling matrix.
' ========================================================================================
PRIVATE FUNCTION GdiPlusGraphics.ScaleTransform (BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN GdipScaleWorldTransform(m_Graphics, sx, sy, order)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GdiPlusGraphics.ScaleTransform (BYVAL dpiRatio_ AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   RETURN GdipScaleWorldTransform(m_Graphics, dpiRatio, dpiRatio_, order)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GdiPlusGraphics.ScaleTransform (BYVAL order AS MatrixOrder = MatrixOrderPrepend) AS GpStatus
   DIM _dpiX AS SINGLE
   GdipGetDpiX(m_Graphics, @_dpiX)
   DIM rxRatio AS SINGLE = _dpiX / 96
   DIM _dpiY AS SINGLE
   GdipGetDpiY(m_Graphics, @_dpiY)
   DIM ryRatio AS SINGLE = _dpiY / 96
   RETURN GdipScaleWorldTransform(m_Graphics, rxRatio, ryRatio, order)
END FUNCTION
' ========================================================================================
```
---

### GdiPlusFont

A Font object is used when drawing strings.

```
' ========================================================================================
TYPE GdiPlusFont
Public:
   m_Font AS GpFont PTR
   DECLARE CONSTRUCTOR (BYVAL font AS GpFont PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC)
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTA PTR)
   DECLARE CONSTRUCTOR (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTW PTR)
   DECLARE CONSTRUCTOR (BYVAL fontFamily AS GpFontFamily PTR, BYVAL emSize AS SINGLE, BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0)
   DECLARE CONSTRUCTOR (BYVAL pwszFamilyName AS WSTRING PTR, BYVAL emSize AS SINGLE, BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0, BYVAL fontCollection AS GpFontCollection PTR = NULL)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpFont PTR PTR
   DECLARE OPERATOR CAST () AS GpFont PTR
END TYPE
' ========================================================================================

' ========================================================================================
' Creates and initializes a Font object from another Font object.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusFont (BYVAL font AS GpFont PTR = NULL)
   IF font THEN GdipCloneFont(font, @m_Font)
END CONSTRUCTOR
' =====================================================================================

' ========================================================================================
' Creates a Font object based on the GDI font object that is currently selected into a
' specified device context. This constructor is provided for compatibility with GDI.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusFont (BYVAL hdc AS HDC)
   SetLastError(GdipCreateFontFromDC(hdc, @m_Font))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Font object directly from a Windows Graphics Device Interface (GDI) logical
' font. The GDI logical font is a LOGFONTA structure, which is the one-byte character
' version of a logical font. This constructor is provided for compatibility with GDI.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusFont (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTA PTR)
   IF plogfont <> NULL THEN
      SetLastError(GdipCreateFontFromLogfontA(hdc, plogfont, @m_Font))
   ELSE
      SetLastError(GdipCreateFontFromDC(hdc, @m_Font))
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Font object directly from a Windows Graphics Device Interface (GDI) logical
' font. The GDI logical font is a LOGFONTW structure, which is the one-byte character
' version of a logical font. This constructor is provided for compatibility with GDI.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusFont (BYVAL hdc AS HDC, BYVAL plogfont AS LOGFONTW PTR)
   IF plogfont <> NULL THEN
      SetLastError(GdipCreateFontFromLogfontW(hdc, plogfont, @m_Font))
   ELSE
      SetLastError(GdipCreateFontFromDC(hdc, @m_Font))
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Font object based on a FontFamily object, a size, a font style, and a unit of measurement.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusFont (BYVAL fontFamily AS GpFontFamily PTR, BYVAL emSize AS SINGLE, BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0)
   SetLastError(GdipCreateFont(fontFamily, emSize, nStyle, nUnit, @m_Font))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates a Font object based on a font family, a size, a font style, a unit of
' measurement, and a FontCollection object.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusFont (BYVAL pwszFamilyName AS WSTRING PTR, BYVAL emSize AS SINGLE, BYVAL nStyle AS INT_ = 0, BYVAL nUnit AS GpUnit = 0, BYVAL fontCollection AS GpFontCollection PTR = NULL)
   DIM fontFamily AS GpFontFamily PTR
   SetLastError(GdipCreateFontFamilyFromName(pwszFamilyName, fontCollection, @fontFamily))
   IF fontFamily THEN
      SetLastError GdipCreateFont(fontFamily, emSize, nStyle, nUnit, @m_Font)
   ELSE
      SetLastError(GdipGetGenericFontFamilySansSerif(@fontFamily))
      SetLastError(GdipCreateFont(fontFamily, emSize, nStyle, nUnit, @m_Font))
   END IF
   IF fontFamily THEN GdipDeleteFontFamily(fontFamily)
END CONSTRUCTOR
' ========================================================================================
' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR GdiPlusFont
   IF m_Font THEN GdipDeleteFont(m_Font)
END DESTRUCTOR
' ===========================================================================================
' ========================================================================================
' Returns a pointer to the GpFont object
' ========================================================================================
PRIVATE OPERATOR * (BYREF font AS GdiPlusFont) AS GpFont PTR
   RETURN font.m_Font
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpFont object
' ========================================================================================
PRIVATE OPERATOR GdiPlusFont.@ () AS GpFont PTR PTR
   RETURN @m_Font
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusFont.CAST () AS GpFont PTR
   RETURN m_Font
END OPERATOR
' ========================================================================================
```
---

### GdiPlusFontFamily

A font family is a group of fonts that have the same typeface but different styles.

```
' ========================================================================================
TYPE GdiPlusFontFamily
Public:
   m_FontFamily AS GpFontFamily PTR
   DECLARE CONSTRUCTOR (BYVAL fontFamily AS GpFontFamily PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL pwszName AS WSTRING PTR, BYVAL fontCollection AS GpFontCollection PTR = NULL)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpFontFamily PTR PTR
   DECLARE OPERATOR CAST () AS GpFontFamily PTR
END TYPE
' ========================================================================================

' ========================================================================================
' Creates and initializes a FontFamily object from another FontFamily object.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusFontFamily (BYVAL fontFamily AS GpFontFamily PTR = NULL)
   IF fontFamily THEN SetLastError(GdipCloneFontFamily(fontFamily, @m_FontFamily))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' Creates a FontFamily object based on a specified font collection.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusFontFamily (BYVAL pwszName AS WSTRING PTR, BYVAL fontCollection AS GpFontCollection PTR = NULL)
   SetLastError(GdipCreateFontFamilyFromName(pwszName, fontCollection, @m_FontFamily))
END CONSTRUCTOR
' ========================================================================================
' =====================================================================================
' Cleanup
' =====================================================================================
PRIVATE DESTRUCTOR GdiPlusFontFamily
   IF m_FontFamily THEN GdipDeleteFontFamily(m_FontFamily)
END DESTRUCTOR
' =====================================================================================
' ========================================================================================
' Returns a pointer to the GpFontFamily object
' ========================================================================================
PRIVATE OPERATOR * (BYREF fontFamily AS GdiPlusFontFamily) AS GpFontFamily PTR
   RETURN fontFamily.m_FontFamily
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpFontFamily object
' ========================================================================================
PRIVATE OPERATOR GdiPlusFontFamily.@ () AS GpFontFamily PTR PTR
   RETURN @m_FontFamily
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusFontFamily.CAST () AS GpFontFamily PTR
   RETURN m_FontFamily
END OPERATOR
' ========================================================================================
```
---

### GdiPlusGraphicsPath

A GraphicsPath object stores a sequence of lines, curves, and shapes. You can draw the entire sequence by calling the GdipDrawPath function. You can partition the sequence of lines, curves, and shapes into figures, and with the help of the PathIterator functions, you can draw selected figures. You can also place markers in the sequence, so that you can draw selected portions of the path.

```
' ========================================================================================
TYPE GdiPlusGraphicsPath
Public:
   m_Path AS GpPath PTR
   DECLARE CONSTRUCTOR (BYVAL path AS GpPath PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL nFillMode AS FillMode)
   DECLARE CONSTRUCTOR (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
   DECLARE CONSTRUCTOR (BYVAL pts AS GpPoint PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpPath PTR PTR
   DECLARE OPERATOR CAST () AS GpPath PTR
END TYPE
' ========================================================================================

' =====================================================================================
' Constructors
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusGraphicsPath (BYVAL path AS GpPath PTR = NULL)
   IF path THEN SetLastError(GdipClonePath(path, @m_Path))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusGraphicsPath (BYVAL nFillMode AS FillMode)
   SetLastError(GdipCreatePath(nFillMode, @m_Path))
END CONSTRUCTOR
' ========================================================================================
' =====================================================================================
' Creates a GraphicsPath object based on an array of points, an array of types, and a fill mode.
' Default value for fillMode: FillModeAlternate(0).
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusGraphicsPath (BYVAL pts AS GpPointF PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
   SetLastError(GdipCreatePath2(pts, types, nCount, nFillmode, @m_Path))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusGraphicsPath (BYVAL pts AS GpPoint PTR, BYVAL types AS BYTE PTR, BYVAL nCount AS LONG, BYVAL nFillMode AS FillMode = FillModeAlternate)
   SetLastError(GdipCreatePath2I(pts, types, nCount, nFillmode, @m_Path))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Cleanup
' =====================================================================================
PRIVATE DESTRUCTOR GdiPlusGraphicsPath
   IF m_Path THEN GdipDeletePath(m_Path)
END DESTRUCTOR
' =====================================================================================
' ========================================================================================
' Returns a pointer to the GpGraphicsPath object
' ========================================================================================
PRIVATE OPERATOR * (BYREF path AS GdiPlusGraphicsPath) AS GpPath PTR
   RETURN path.m_Path
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpGraphicsPath object
' ========================================================================================
PRIVATE OPERATOR GdiPlusGraphicsPath.@ () AS GpPath PTR PTR
   RETURN @m_Path
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusGraphicsPath.CAST () AS GpPath PTR
   RETURN m_Path
END OPERATOR
' ========================================================================================
```
---

### GdiPlusMatrix

A Matrix object represents a 3 ×3 matrix that, in turn, represents an affine transformation. A Matrix object stores only six of the 9 numbers in a 3 ×3 matrix because all 3 ×3 matrices that represent affine transformations have the same third column (0, 0, 1).

```
' ========================================================================================
TYPE GdiPlusMatrix
Public:
   m_Matrix AS GpMatrix PTR
   DECLARE CONSTRUCTOR (BYVAL matrix AS GpMatrix PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL m11 AS SINGLE, BYVAL m12 AS SINGLE, BYVAL m21 AS SINGLE, BYVAL m22 AS SINGLE, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE)
   DECLARE CONSTRUCTOR (BYVAL rcf AS GpRectF PTR, BYVAL dstplg AS GpPointF PTR)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRect PTR, BYVAL dstplg AS GpPoint PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS GpMatrix PTR
   DECLARE OPERATOR @ () AS GpMatrix PTR PTR
END TYPE
' ========================================================================================

' ========================================================================================
' Constructors
' ========================================================================================
' ========================================================================================
' Creates and initializes a Matrix object from another Matrix object
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusMatrix (BYVAL matrix AS GpMatrix PTR = NULL)
   IF matrix THEN
      SetLastError(GdipCloneMatrix(matrix, @m_Matrix))
   ELSE
      SetLastError(GdipCreateMatrix(@m_Matrix))
   END IF
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Creates and initializes a Matrix object based on six numbers that define an affine transformation.
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusMatrix (BYVAL m11 AS SINGLE, BYVAL m12 AS SINGLE, BYVAL m21 AS SINGLE, BYVAL m22 AS SINGLE, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE)
   SetLastError(GdipCreateMatrix2(m11, m12, m21, m22, dx, dy, @m_Matrix))
END CONSTRUCTOR
' ========================================================================================
' =====================================================================================
' Creates a Matrix object based on a rectangle and a point.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusMatrix (BYVAL rcf AS GpRectF PTR, BYVAL dstplg AS GpPointF PTR)
   SetLastError(GdipCreateMatrix3(rcf, dstplg, @m_Matrix))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a Matrix object based on a rectangle and a point.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusMatrix (BYVAL rc AS GpRect PTR, BYVAL dstplg AS GpPoint PTR)
   SetLastError(GdipCreateMatrix3I(rc, dstplg, @m_Matrix))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusMatrix
   IF m_Matrix THEN GdipDeleteMatrix(m_Matrix)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpMatrix object
' ========================================================================================
PRIVATE OPERATOR * (BYREF matrix AS GdiPlusMatrix) AS GpMatrix PTR
   RETURN matrix.m_Matrix
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusMatrix.CAST () AS GpMatrix PTR
   RETURN m_Matrix
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpMatrix object
' ========================================================================================
PRIVATE OPERATOR GdiPlusMatrix.@ () AS GpMatrix PTR PTR
   RETURN @m_Matrix
END OPERATOR
' ========================================================================================
```
---

### GdiPlusPathIterator

The PathIterator functions allow to isolate selected subsets of the path stored in a GraphicsPath object. A path consists of one or more figures. You can use a GraphicsPathIterator to isolate one or more of those figures. A path can also have markers that divide the path into sections. You can use a GraphicsPathIterator object to isolate one or more of those sections.

```
' ========================================================================================
TYPE GdiPlusPathIterator
Public:
   m_PathIretator AS GpPathIterator PTR
   DECLARE CONSTRUCTOR (BYVAL path AS GpPath PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpPathIterator PTR PTR
   DECLARE OPERATOR CAST () AS GpPathIterator PTR
END TYPE
' ========================================================================================

' =====================================================================================
' Creates a new CGpGraphicsPathIterator object and associates it with a CGpGraphicsPath object.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusPathIterator (BYVAL path AS GpPath PTR)
   SetLasTError(GdipCreatePathIter(@m_pathIretator, path))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusPathIterator
   IF m_PathIretator THEN GdipDeletePathIter(m_PathIretator)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpPathIterator object
' ========================================================================================
PRIVATE OPERATOR * (BYREF pathIterator AS GdiPlusPathIterator) AS GpPathIterator PTR
   RETURN pathIterator.m_PathIretator
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpPathIterator object
' ========================================================================================
PRIVATE OPERATOR GdiPlusPathIterator.@ () AS GpPathIterator PTR PTR
   RETURN @m_PathIretator
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusPathIterator.CAST () AS GpPathIterator PTR
   RETURN m_PathIretator
END OPERATOR
' ========================================================================================
```
---

### GdiPlusPen

A Pen object is a Microsoft Windows GDI+ object used to draw lines and curves.

```
' ========================================================================================
TYPE GdiPlusPen
Public:
   m_Pen AS GpPen PTR
   DECLARE CONSTRUCTOR (BYVAL pen AS GpPen PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL colour AS ARGB, BYVAL nWidth AS SINGLE = 1.0, BYVAL unit AS GpUnit = UnitWorld)
   DECLARE CONSTRUCTOR (BYVAL brush AS GpBrush PTR, BYVAL nWidth AS SINGLE, BYVAL unit AS GpUnit = UnitWorld)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpPen PTR PTR
   DECLARE OPERATOR CAST () AS GpPen PTR
END TYPE
' ========================================================================================

' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusPen (BYVAL pen AS GpPen PTR = NULL)
   IF pen THEN SetLastError(GdipClonePen(pen, @m_Pen))
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' * Creates a Pen object that uses a specified color and width.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusPen (BYVAL colour AS ARGB, BYVAL nWidth AS SINGLE = 1.0, BYVAL unit AS GpUnit = UnitWorld)
   SetLastError(GdipCreatePen1(colour, nWidth, unit, @m_Pen))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a Pen object that uses the attributes of a brush and a real number to set the
' width of this Pen object.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusPen (BYVAL brush AS GpBrush PTR, BYVAL nWidth AS SINGLE, BYVAL unit AS GpUnit = UnitWorld)
   SetLastError(GdipCreatePen2(brush, nWidth, unit, @m_Pen))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusPen
   IF m_Pen THEN SetLastError(GdipDeletePen(m_Pen))
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpPen object
' ========================================================================================
PRIVATE OPERATOR * (BYREF pen AS GdiPlusPen) AS GpPen PTR
   RETURN pen.m_Pen
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpPen object
' ========================================================================================
PRIVATE OPERATOR GdiPlusPen.@ () AS GpPen PTR PTR
   RETURN @m_Pen
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusPen.CAST () AS GpPen PTR
   RETURN m_Pen
END OPERATOR
' ========================================================================================
```
---

### GdiPlusRegion

A Region describes an area of the display surface. The area can be any shape. In other words, the boundary of the area can be a combination of curved and straight lines. Regions can also be created from the interiors of rectangles, paths, or a combination of these. Regions are used in clipping and hit-testing operations.

```
' ========================================================================================
TYPE GdiPlusRegion
Public:
   m_Region AS GpRegion PTR
   DECLARE CONSTRUCTOR (BYVAL region AS GpRegion PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL rcf AS GpRectF PTR)
   DECLARE CONSTRUCTOR (BYVAL rc AS GpRect PTR)
   DECLARE CONSTRUCTOR (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE)
   DECLARE CONSTRUCTOR (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_)
   DECLARE CONSTRUCTOR (BYVAL regionData AS BYTE PTR, BYVAL nSize AS LONG)
   DECLARE CONSTRUCTOR (BYVAL hRgn AS HRGN)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpRegion PTR PTR
   DECLARE OPERATOR CAST () AS GpRegion PTR
   DECLARE SUB FromPath (BYVAL pah AS GpPath PTR)
END TYPE
' ========================================================================================

' =====================================================================================
' Default constructor
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusRegion (BYVAL region AS GpRegion PTR = NULL)
   IF region THEN
      SetLastError(GdipCloneRegion(region, @m_Region))
   ELSE
      SetLastError(GdipCreateRegion(@m_Region))
   END IF
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusRegion (BYVAL rcf AS GpRectF PTR)
   SetLastError(GdipCreateRegionRect(rcf, @m_Region))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusRegion (BYVAL rc AS GpRect PTR)
   SetLastError(GdipCreateRegionRectI(rc, @m_Region))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusRegion (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, BYVAL nHeight AS SINGLE)
   DIM rcf AS GpRectF : rcf.x = x : rcf.y = y : rcf.Width = nwidth : rcf.Height = nHeight
   SetLastError(GdipCreateRegionRect(@rcf, @m_Region))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR GdiPlusRegion (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_)
   DIM rc AS GpRect : rc.x = x : rc.y = y : rc.Width = nwidth : rc.Height = nHeight
   SetLastError(GdipCreateRegionRectI(@rc, @m_Region))
END CONSTRUCTOR
' ========================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusRegion (BYVAL regionData AS BYTE PTR, BYVAL nSize AS LONG)
   SetLastError(GdipCreateRegionRgnData(regionData, nSize, @m_Region))
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusRegion (BYVAL hRgn AS HRGN)
   SetLastError(GdipCreateRegionHrgn(hRgn, @m_Region))
END CONSTRUCTOR
' =====================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusRegion
   IF m_Region THEN GdipDeleteRegion(m_Region)
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpRegion object
' ========================================================================================
PRIVATE OPERATOR * (BYREF region AS GdiPlusRegion) AS GpRegion PTR
   RETURN region.m_Region
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpRegion object
' ========================================================================================
PRIVATE OPERATOR GdiPlusRegion.@ () AS GpRegion PTR PTR
   RETURN @m_Region
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusRegion.CAST () AS GpRegion PTR
   RETURN m_Region
END OPERATOR
' ========================================================================================
' =====================================================================================
' Creates a region from a path
' A separate function is needed because using a constructor conflicts with the default one.
' =====================================================================================
PRIVATE SUB GdiPlusRegion.FromPath (BYVAL path AS GpPath PTR)
   SetLastError(GdipCreateRegionPath(path, @m_Region))
END SUB
' =====================================================================================
```
---

### GdiPlusStringFormat

The StringFormat functions allow to set and get text layout information (such as alignment, orientation, tab stops, and clipping) and display manipulations (such as trimming, font substitution for characters that are not supported by the requested font, and digit substitution for languages that do not use Western European digits). A StringFormat object can be passed to the GdipDrawString function to format a string.

```
' ========================================================================================
TYPE GdiPlusStringFormat
Public:
   m_StringFormat AS GpStringFormat PTR
   DECLARE CONSTRUCTOR (BYVAL stringFormat AS GpStringFormat PTR = NULL)
   DECLARE CONSTRUCTOR (BYVAL formatAttributes AS INT_, BYVAL language AS LANGID)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS GpStringFormat PTR PTR
   DECLARE OPERATOR CAST () AS GpStringFormat PTR
END TYPE
' ========================================================================================

' =====================================================================================
' Creates and initializes a default StringFormat object or clones it from another StringFormat object.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusStringFormat (BYVAL stringFormat AS GpStringFormat PTR = NULL)
   IF stringFormat THEN
      SetLastError(GdipCloneStringFormat(stringFormat, @m_StringFormat))
   ELSE
      SetLastError(GdipCreateStringFormat(0, LANG_NEUTRAL, @m_StringFormat))
   END IF
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Creates a StringFormat object based on string format flags and a language.
' =====================================================================================
PRIVATE CONSTRUCTOR GdiPlusStringFormat (BYVAL formatFlags AS INT_ = 0, BYVAL language AS LANGID = LANG_NEUTRAL)
   SetLastError(GdipCreateStringFormat(formatFlags, language, @m_StringFormat))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR GdiPlusStringFormat
   IF m_StringFormat THEN SetLastError(GdipDeleteStringFormat(m_StringFormat))
END DESTRUCTOR
' ========================================================================================
' ========================================================================================
' Returns a pointer to the GpStringFormat object
' ========================================================================================
PRIVATE OPERATOR * (BYREF fmt AS GdiPlusStringFormat) AS GpStringFormat PTR
   RETURN fmt.m_StringFormat
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of a pointer to the GpStringFormat object
' ========================================================================================
PRIVATE OPERATOR GdiPlusStringFormat.@ () AS GpStringFormat PTR PTR
   RETURN @m_StringFormat
END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE OPERATOR GdiPlusStringFormat.CAST () AS GpStringFormat PTR
   RETURN m_StringFormat
END OPERATOR
' ========================================================================================
```
---
