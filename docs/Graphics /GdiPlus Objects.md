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

