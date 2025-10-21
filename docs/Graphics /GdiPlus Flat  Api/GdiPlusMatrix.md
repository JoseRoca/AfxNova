## Matrix Functions

A **Matrix** object represents a 3 ×3 matrix that, in turn, represents an affine transformation. A **Matrix** object stores only six of the 9 numbers in a 3 ×3 matrix because all 3 ×3 matrices that represent affine transformations have the same third column (0, 0, 1).

| Name       | Description |
| ---------- | ----------- |
| [GdipCloneMatrix](#gdipclonemnatrix) | Creates a new **Matrix** object that is a copy of a **Matrix** object. |
| [GdipCreateMatrix](#gdipcreatematrix) | Creates and initializes a **Matrix** object that represents the identity matrix. |
| [GdipCreateMatrix2](#gdipcreatematrix2) | Creates and initializes a **Matrix** object based on six numbers that define an affine transformation. |
| [GdipCreateMatrix3](#gdipcreatematrix3) | Creates a **Matrix** object based on a rectangle and a point. |
| [GdipCreateMatrix3I](#gdipcreatematrix3) | Creates a **Matrix** object based on a rectangle and a point. |
| [GdipDeleteMatrix](#gdipdeletematrix) | Deletes the specified **Matrix** object. |
| [GdipGetMatrixElements](#gdipgetmatrixelements) | Gets the elements of a matrix. The elements are placed in an array in the order m11, m12, m21, m22, m31, m32, where mij denotes the element in row i, column j. |
| [GdipInvertMatrix](#gdipinvertmatrix) | If the matrix is invertible, this function replaces its elements  with the elements of its inverse. |
| [GdipIsMatrixEqual](#gdipismatrixequal) | Determines whether the elements of a matrix are equal to the elements of another matrix. |
| [GdipIsMatrixIdentity](#gdipismatrixidentity) | Determines whether the matrix is the identity matrix. |
| [GdipIsMatrixInvertible](#gdipismatrixinvertible) | Determines whether the matrix is invertible. |
| [GdipMultiplyMatrix](#gdipmultiplymatrix) | Updates a matrix with the product of itself and another matrix. |
| [GdipRotateMatrix](#gdiprotatematrix) | Updates a matrix with the product of itself and a rotation matrix. |
| [GdipScaleMatrix](#gdipscalematrix) | Updates a matrix with the product of itself and a scaling matrix. |
| [GdipSetMatrixElements](#gdipsetmatrixelements) | Sets the elements of a matrix. |
| [GdipShearMatrix](#gdipshearmatrix) | Updates a matrix with the product of itself and a shearing matrix. |
| [GdipTransformMatrixPoints](#gdiptransformmatrixpoints) | Multiplies each point in an array by a matrix. Each point is treated as a row matrix. The multiplication is performed with the row matrix on the left and a matrix on the right. |
| [GdipTransformMatrixPointsI](#gdiptransformmatrixpoints) | Multiplies each point in an array by a matrix. Each point is treated as a row matrix. The multiplication is performed with the row matrix on the left and a matrix on the right. |
| [GdipTranslateMatrix](#gdiptranslatematrix) | Updates a matrix with the product of itself and a translation matrix. |
| [GdipVectorTransformMatrixPoints](#gdipvectortransformmatrixpoints) | Multiplies each vector in an array by a matrix. The translation elements of a matrix (third row) are ignored. Each vector is treated as a row matrix. The multiplication is performed with the row matrix on the left and a matrix on the right. |
| [GdipVectorTransformMatrixPointsI](#gdipvectortransformmatrixpoints) | Multiplies each vector in an array by a matrix. The translation elements of a matrix (third row) are ignored. Each vector is treated as a row matrix. The multiplication is performed with the row matrix on the left and a matrix on the right. |

---

## GdipCloneMatrix

Creates a new **Matrix** object that is a copy of a **Matrix** object.

```
FUNCTION GdipCloneMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL cloneMatrix AS GpMatrix PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *cloneMatrix* | [out] Pointer to a variable that receives a pointer to the new **Matrix** object. |

#### Example

```
' ========================================================================================
' This example demonstrates matrix cloning and transformation using the GDI+ Flat API.
' It draws two rectangles: one scaled horizontally, and one scaled and translated.
' Note: DPI-Aware Composite Transform Strategy Apply logical transforms via matrix, then
' prepend DPI scaling via GdipScaleWorldTransform. Use UnitPixel for pens to maintain
' consistent stroke width.
' ========================================================================================
SUB Example_CloneMatrix (BYVAL hdc AS HDC)

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

   ' // Create a matrix for horizontal scaling
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(3, 0, 0, 1, 0, 0, @matrix)

   ' // Clone the matrix
   DIM clonedMatrix AS GpMatrix PTR
   hStatus = GdipCloneMatrix(matrix, @clonedMatrix)

   ' // Apply translation to the cloned matrix
   hStatus = GdipTranslateMatrix(clonedMatrix, 40 * rxRatio, 25 * ryRatio, MatrixOrderAppend)

   ' // Create a blue pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1 * rxRatio, UnitPixel, @pen)

   ' // Apply first transform and draw scaled rectangle
   hStatus = GdipSetWorldTransform(graphics, matrix)
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, 0, 0, 100, 100)

   ' // Apply second transform and draw scaled + translated rectangle
   hStatus = GdipSetWorldTransform(graphics, clonedMatrix)
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, 0, 0, 100, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF clonedMatrix THEN GdipDeleteMatrix(clonedMatrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateMatrix

Creates and initializes a Matrix object that represents the identity matrix.

```
FUNCTION GdipCreateMatrix (BYVAL matrix AS GpMatrix PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [out] Pointer to a variable that receives a pointer to the new created **Matrix** object. |

#### Usage example

```
DIM matrix AS GpMatrix PTR
GdipCreateMatrix(@matrix)
```
---

## GdipCreateMatrix2

Creates and initializes a **Matrix** object based on six numbers that define an affine transformation.

```
FUNCTION GdipCreateMatrix2 (BYVAL m11 AS REAL, BYVAL m12 AS REAL, BYVAL m21 AS REAL, BYVAL m22 AS REAL, _
         BYVAL dx AS REAL, BYVAL dy AS REAL, BYVAL matrix AS GpMatrix PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *m11* | [in] Simple precision value that specifies the element in the first row, first column. |
| *m12* | [in] Simple precision value that specifies the element in the first row, second column. |
| *m21* | [in] Simple precision value that specifies the element in the second row, first column. |
| *m22* | [in] Simple precision value that specifies the element in the second row, second column. |
| *dx* | [in] Simple precision value that specifies the element in the third row, first column. |
| *dy* | [in] Simple precision value that specifies the element in the third row, second column. |
| *matrix* | [out] Pointer to a variable that receives a pointer to the new created **Matrix** object. |

#### Usage example

```
DIM matrix AS GpMatrix PTR
GdipCreateMatrix2(3, 0, 0, 1, 0, 0, @matrix)
```
---

## GdipCreateMatrix3

Creates a **Matrix** object based on a rectangle and a point.

```
FUNCTION GdipCreateMatrix3 (BYVAL rect AS CONST GpRectF PTR, BYVAL dstplg AS CONST GpPointF PTR, _
   BYVAL matrix AS GpMatrix PTR PTR) AS GpStatus
```
```
FUNCTION GdipCreateMatrix3I (BYVAL rect AS CONST GpRect PTR, BYVAL dstplg AS CONST GpPoint PTR, _
   BYVAL matrix AS GpMatrix PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rect* | [in] Reference to a **GpRect** structure. The *x* data member of the rectangle specifies the matrix element in row 1, column 1. The *y* data member of the rectangle specifies the matrix element in row 1, column 2. The *width* data member of the rectangle specifies the matrix element in row 2, column 1. The *height* data member of the rectangle specifies the matrix element in row 2, column 2. |
| *dstplg* | [in] Pointer to a **GpPoint** structure. The *x* data member of the point specifies the matrix element in row 3, column 1. The *y* data member of the point specifies the matrix element in row 3, column 2. |
| *matrix* | [out] Pointer to a variable that receives a pointer to the new created **Matrix** object. |

#### Example

```
' ========================================================================================
' Creates a transformation matrix based on a source rectangle and a destination parallelogram.
' This is especially useful for mapping one shape onto another—like skewing, rotating, or
' stretching a rectangle into a custom shape.
' ========================================================================================
SUB Example_CreateMatrix3 (BYVAL hdc AS HDC)

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

   ' // Define source rectangle
   DIM srcRect AS GpRectF = (0, 0, 100, 60)

   ' // Define destination parallelogram points
   DIM destPoints(0 TO 2) AS GpPointF
   destPoints(0).x = 50  : destPoints(0).y = 50      ' Top-left
   destPoints(1).x = 160 : destPoints(1).y = 40      ' Top-right
   destPoints(2).x = 60  : destPoints(2).y = 120     ' Bottom-left

   ' // Create matrix to map srcRect to dstPoints
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix3(@srcRect, @destPoints(0), @matrix)

   ' // Apply matrix to graphics
   hStatus = GdipSetWorldTransform(graphics, matrix)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_RED, 2 * rxRatio, UnitPixel, @pen)

   ' // Draw transformed rectangle
   GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, 0, 0, 100, 60)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeleteMatrix

Deletes the specified **Matrix** object.

```
FUNCTION GdipDeleteMatrix (BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |

#### Usage example

```
DIM matrix AS GpMatrix PTR
hStatus = GdipCreateMatrix2(3, 0, 0, 1, 0, 0, @matrix)
...
...
...
IF matrix THEN GdipDeleteMatrix(matrix)
```
---

## GdipGetMatrixElements

Gets the elements of a matrix. The elements are placed in an array in the order m11, m12, m21, m22, m31, m32, where mij denotes the element in row i, column j.

```
FUNCTION GdipGetMatrixElements (BYVAL matrix AS CONST GpMatrix PTR, BYVAL matrixOut AS REAL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *matrixOut* | [out] Pointer to an array that receives the matrix elements. The size of the array should be 6 x SIZEOF(SINGLE), i.e. 24. |

#### Example

```
' ========================================================================================
' Retrieves the six elements of a transformation matrix. This is useful for inspecting or
' debugging the current transformation state.
' ========================================================================================
SUB Example_GetMatrixElements (BYVAL hdc AS HDC)

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
   GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create identity matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)

   ' // First translate
   hStatus = GdipTranslateMatrix(matrix, 40 * rxRatio, 30 * ryRatio, MatrixOrderAppend)
   ' // Then rotate
   hStatus = GdipRotateMatrix(matrix, 45, MatrixOrderAppend)

   ' // Retrieve matrix elements
   DIM elements(0 TO 5) AS SINGLE
   hStatus = GdipGetMatrixElements(matrix, @elements(0))

   ' // Create font and brush
   DIM fontFamily AS GpFontFamily PTR
   DIM font AS GpFont PTR
   DIM brush AS GpSolidFill PTR
   GdipCreateFontFamilyFromName("Arial", NULL, @fontFamily)
   GdipCreateFont(fontFamily, AfxGdipPointsToPixels(12, TRUE), FontStyleRegular, UnitPixel, @font)
   GdipCreateSolidFill(ARGB_BLACK, @brush)

   ' // Display matrix values
   DIM text AS STRING
   text = "Matrix Elements:" & !"\n"
   text &= "m11 = " & STR(elements(0)) & !"\n"
   text &= "m12 = " & STR(elements(1)) & !"\n"
   text &= "m21 = " & STR(elements(2)) & !"\n"
   text &= "m22 = " & STR(elements(3)) & !"\n"
   text &= "dx  = " & STR(elements(4)) & !"\n"
   text &= "dy  = " & STR(elements(5))
   DIM layout AS GpRectF = (10.0, 10.0, 300.0, 200.0)
   GdipDrawString(graphics, text, LEN(text), font, @layout, NULL, brush)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipInvertMatrix

If the matrix is invertible, this function replaces its elements  with the elements of its inverse.

```
FUNCTION GdipInvertMatrix (BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |

#### Example

```
' ========================================================================================
' The following example passes the address of a Matrix object to the SetTransform method
' of a Graphics object and then draws a rectangle. The rectangle is translated 30 units
' right and 20 units down by the world transformation of the Graphics object. The code
' calls the Matrix.Invert method of the Matrix object and sets the world transformation
' of the Graphics object to the inverted matrix. The code draws a second rectangle that is
' translated 30 units left and 20 units up.
' ========================================================================================
SUB Example_MatrixInvert (BYVAL hdc AS HDC)

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

   ' // Create a matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(1, 0, 0, 1, 30 *rxRatio, 20 * ryRatio, @matrix)

   ' // Create a pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1 * rxRatio, UnitPixel, @pen)

   hStatus = GdipSetWorldTransform(graphics, matrix)
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, 0, 0, 200, 100)

   hStatus = GdipInvertMatrix(matrix)
   hStatus = GdipSetWorldTransform(graphics, matrix)
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, 0, 0, 200, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipIsMatrixEqual

Determines whether the elements of a matrix are equal to the elements of another matrix.

```
FUNCTION GdipIsMatrixEqual (BYVAL matrix AS CONST GpMatrix PTR, _
  BYVAL matrix2 AS CONST GpMatrix PTR, BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *matrix2* | [in] Pointer to a **Matrix** object that is compared with another **Matrix** object. |
| *result* | [out] Pointer to a variable that receives a boolean value indicating whether the elements of the two matrices are the same (TRUE) or not (FALSE). |

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

' ========================================================================================
' Determines whether the elements of a matrix are equal to the elements of another matrix.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create the matrices
DIM matrix AS GpMatrix PTR
hStatus = GdipCreateMatrix2(1, 2, 1, 1, 2, 2, @matrix)

DIM matrix2 AS GpMatrix PTR
hStatus = GdipCreateMatrix2(1, 2, 1, 1, 2, 2, @matrix2)

DIM result AS BOOL
hStatus = GdipIsMatrixEqual(matrix, matrix2, @result)
IF result THEN PRINT "They are the same"  ELSE PRINT "They are the different"

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---

## GdipIsMatrixIdentity

Determines whether a matrix is the identity matrix.

```
FUNCTION GdipIsMatrixIdentity (BYVAL matrix AS CONST GpMatrix PTR, BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *result* | [out] Pointer to a variable that receives a boolean value indicating whether a matrix is the identity matrix (TRUE) or not (FALSE). |

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

' ========================================================================================
' Determines whether a matrix is the identity matrix.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create the matrix
DIM matrix AS GpMatrix PTR
hStatus = GdipCreateMatrix2(1.0, 0.0, 0.0, 1.0, 0.0, 0.0, @matrix)

DIM result AS BOOL
hStatus = GdipIsMatrixIdentity(matrix, @result)
IF result THEN PRINT "The matrix is the identity matrix"  ELSE PRINT "The matrix is not the identity matrix"

' // Cleanup
IF matrix THEN GdipDeleteMatrix(matrix)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---

## GdipIsMatrixInvertible

Determines whether a matrix is invertible.

```
FUNCTION GdipIsMatrixInvertible (BYVAL matrix AS CONST GpMatrix PTR, BYVAL result AS BOOL PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *result* | [out] Pointer to a variable that receives a boolean value indicating whether a matrix is invertible (TRUE) or not (FALSE). |

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

' ========================================================================================
' Determines whether a matrix is invertible.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Create the matrix
DIM matrix AS GpMatrix PTR
hStatus = GdipCreateMatrix2(3.0, 0.0, 0.0, 2.0, 20.0, 10.0, @matrix)

DIM result AS BOOL
hStatus = GdipIsMatrixInvertible(matrix, @result)
IF result THEN PRINT "The matrix is invertible"  ELSE PRINT "The matrix is not invertible"

' // Cleanup
IF matrix THEN GdipDeleteMatrix(matrix)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP
```
---

## GdipMultiplyMatrix

Updates a matrix with the product of itself and another matrix.

```
FUNCTION GdipMultiplyMatrix (BYVAL matrix AS GpMatrix PTR, _
   BYVAL matrix2 AS GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *matrix2* | [in] Pointer to a matrix that will be multiplied by a matrix. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the passed matrix is on the left, and *MatrixOrderAppend* specifies that the passed matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates two Matrix objects: matrix1 and matrix2. The call to the
' Matrix.Multiply method updates matrix1 with the product ( matrix1)( matrix2). At that
' point, matrix1 represents a composite transformation: first scale, then translate. The
' code uses matrix1 to set the world transformation of a Graphics object and then draws a
' rectangle that is transformed according to that world transformation.
' ========================================================================================
SUB Example_MultiplyMatrix (BYVAL hdc AS HDC)

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

   ' // Create a matrix for horizontal scaling
   DIM matrix1 AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(3, 0, 0, 1, 0, 0, @matrix1)

   ' // Translation
   DIM matrix2 AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(1, 0, 0, 1, 20 * rxRatio, 40 * ryRatio, @matrix2)

   ' // Create a pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1 * rxRatio, UnitPixel, @pen)

   hStatus = GdipMultiplyMatrix(matrix1, matrix2, MatrixOrderAppend)
   hStatus = GdipSetWorldTransform(graphics, matrix1)
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, 0, 0, 100, 100)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix1 THEN GdipDeleteMatrix(matrix1)
   IF matrix2 THEN GdipDeleteMatrix(matrix2)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipRotateMatrix

Updates a matrix with the product of itself and a rotation matrix.

```
FUNCTION GdipRotateMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL angle AS REAL, _
   BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *angle* | [in] Simple precision value that specifies the angle of rotation in degrees. Positive values specify clockwise rotation. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the rotation matrix is on the left, and *MatrixOrderAppend* specifies that the rotation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a Matrix object and calls the GdipTranslateMatrix function
' to set the elements of that matrix to a translation. Then the code calls the GdipRotateMatrix
' function to update the matrix with the product of itself and a rotation matrix. At that
' point, the matrix represents a composite transformation: first translate, then rotate.
' The code uses the matrix to set the world transformation of a Graphics object and then
' draws an ellipse that is transformed according to the composite transformation.
' ========================================================================================
SUB Example_RotateMatrix (BYVAL hdc AS HDC)

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

   ' // Create a pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1 * rxRatio, UnitPixel, @pen)

   ' Create identity matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)

   ' // First a translation
   hStatus = GdipTranslateMatrix(matrix, 40 * rxRatio, 0, MatrixOrderAppend)
   ' Rotate 30 degrees clockwise
   hStatus = GdipRotateMatrix(matrix, 30, MatrixOrderAppend)

   ' // Apply matrix to graphics
   hStatus = GdipSetWorldTransform(graphics, matrix)

   ' // Draw rotated rectangle
   GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawEllipse(graphics, pen, 0, 0, 100, 50)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipScaleMatrix

Updates a matrix with the product of itself and a scaling matrix.

```
FUNCTION GdipScaleMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL scaleX AS REAL, _
   BYVAL scaleY AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *scaleX* | [in] Simple precision value that specifies the horizontal scale factor. |
| *scaleY* | [in] Simple precision value that specifies the vertical scale factor. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the scaling matrix is on the left, and *MatrixOrderAppend* specifies that the scaling matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a Matrix object and calls the Matrix.Rotate method to set
' the elements of that matrix to a rotation. Then the code calls the Matrix.Scale method
' to update the matrix with the product of itself and a scaling matrix. At that point, the
' matrix represents a composite transformation: first rotate, then scale. The code uses
' the matrix to set the world transformation of a Graphics object and then draws an ellipse
' that is transformed according to the composite transformation.
' ========================================================================================
SUB Example_ScaleMatrix (BYVAL hdc AS HDC)

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

   ' // Create identity matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)

   ' // Rotate 30 degrees clockwise
   hStatus = GdipRotateMatrix(matrix, 30, MatrixOrderAppend)

   ' // Apply horizontal and vertical scaling
   hStatus = GdipScaleMatrix(matrix, 3.0, 2.0, MatrixOrderAppend)

   ' // Apply matrix to graphics
   hStatus = GdipSetWorldTransform(graphics, matrix)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2 * rxRatio, UnitPixel, @pen)

   ' // Draw scaled ellipse
   GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawEllipse(graphics, pen, 0, 0, 80, 40)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' Applies a scaling transformation to a matrix by multiplying it with a scaling matrix.
' This adjusts the size of graphics or coordinates along the X and Y axes.
' ========================================================================================
SUB Example_ScaleMatrix (BYVAL hdc AS HDC)

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

   ' // Create identity matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)

   ' // Apply horizontal and vertical scaling
   hStatus = GdipScaleMatrix(matrix, 2.0, 1.5, MatrixOrderAppend)

   ' // Apply matrix to graphics
   hStatus = GdipSetWorldTransform(graphics, matrix)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 2 * rxRatio, UnitPixel, @pen)

   ' // Draw scaled rectangle
   GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, 50, 50, 100, 60)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetMatrixElements

Sets the elements of the matrix.

```
FUNCTION GdipSetMatrixElements (BYVAL matrix AS GpMatrix PTR, BYVAL m11 AS REAL, BYVAL m12 AS REAL, _
   BYVAL m21 AS REAL, BYVAL m22 AS REAL, BYVAL dx AS REAL, BYVAL dy AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *m11* | [in] Simple precision value that specifies the element in the first row, first column. |
| *m12* | [in] Simple precision value that specifies the element in the first row, second column. |
| *m21* | [in] Simple precision value that specifies the element in the second row, first column. |
| *m22* | [in] Simple precision value that specifies the element in the second row, second column. |
| *dx* | [in] Simple precision value that specifies the element in the third row, first column. |
| *dy* | [in] Simple precision valuethat specifies the element in the third row, second column. |

#### Example

```
' ========================================================================================
' The following example calls the GdipSetMatrixElements function. The code uses the matrix
' to set the world transformation of a Graphics object and then draws a rectangle that is
' transformed by that world transformation.
' ========================================================================================
SUB Example_SetMatrixElements (BYVAL hdc AS HDC)

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

   ' // Create identity matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)
   hStatus = GdipSetMatrixElements(matrix, 1, 0, 0, 1, 30 * rxRatio, 50 *ryRatio)

   ' // Apply matrix to graphics
   hStatus = GdipSetWorldTransform(graphics, matrix)

   ' // Scale graphics according the dpi ratios
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2 * rxRatio, UnitPixel, @pen)

   ' // Draw a rectangle
   hStatus = GdipDrawRectangle(graphics, pen, 0, 0, 80, 40)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipShearMatrix

Updates a matrix with the product of itself and a shearing matrix.

```
FUNCTION GdipShearMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL shearX AS REAL, _
   BYVAL shearY AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *shearX* | [in] Simple precision value that specifies the horizontal shear factor. |
| *shearY* | [in] Simple precision value that specifies the vertical shear factor. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the shearing matrix is on the left, and *MatrixOrderAppend* specifies that the shearing matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a Matrix object and calls the GdipScaleMatrix function to set
' the elements of that matrix to a scaling transformation. Then the code calls the GdipShearMatrix
' function to update the matrix with the product of itself and a shearing matrix. At that point,
' the matrix represents a composite transformation: first scale, then shear. The code uses
' the matrix to set the world transformation of a Graphics object and then draws a rectangle
' that is transformed according to the composite transformation.
' In the call to GdipShearMatrix, the shearX parameter is 3 and the shearY parameter is 0.
' That particular shearing transformation slides the bottom edge of the rectangle to the
' right. The distance that the bottom edge slides is shearX multiplied by the height of
' the rectangle after it is stretched by the scaling transformation.
' ========================================================================================
SUB Example_ShearMatrix (BYVAL hdc AS HDC)

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

   ' // Create identity matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)

   ' // First a scaling
   hStatus = GdipScaleMatrix(matrix, 2.0, 2.0, MatrixOrderAppend)

   ' // Then a shear
   hStatus = GdipShearMatrix(matrix, 3.0, 0.0, MatrixOrderAppend)

   ' // Apply matrix to graphics
   hStatus = GdipSetWorldTransform(graphics, matrix)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLACK, 2 * rxRatio, UnitPixel, @pen)

   ' // Draw scaled rectangle
   GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)
   hStatus = GdipDrawRectangle(graphics, pen, 0, 0, 100, 50)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTransformMatrixPoints

Multiplies each point in an array by a matrix. Each point is treated as a row matrix. The multiplication is performed with the row matrix on the left and a matrix on the right.

```
FUNCTION GdipTransformMatrixPoints (BYVAL matrix AS GpMatrix PTR, _
   BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
```
```
FUNCTION GdipTransformMatrixPointsI (BYVAL matrix AS GpMatrix PTR, _
   BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *matrix* | [in, out] Pointer to an array of **GpPointF** structures that, on input, contains the points to be transformed and, on output, receives the transformed points. Each point in the array is transformed (multiplied by a matrix) and updated with the result of the transformation. |
| *count* | [in] Long integer value that specifies the number of points to be transformed. The default value is 1. |

#### Example

```
' ========================================================================================
' The following example creates an array of five points and then transforms those points
' by calling the GdipTransformMatrixPoints function. The code passes the array of points
' to the GdipDrawCurve function before the transformation and again after the transformation.
' ========================================================================================
SUB Example_TransformMatrixPoints (BYVAL hdc AS HDC)

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
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create identity matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(1, 0, 0, 2, 0, 0, @matrix)

   ' // Array of points
   DIM rgPoints(0 TO 4) AS GpPointF = {(50, 100), (100, 50), _
       (160, 125), (200, 100), (250, 150)}

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1 * rxRatio, UnitPixel, @pen)

   ' // Draw the curve
   hStatus = GdipDrawCurve(graphics, pen, @rgPoints(0), 5)
   ' // Transform the points
   hStatus = GdipTransformMatrixPoints(matrix, @rgPoints(0), 5)
   ' // Draw the curve after the transformation
   hStatus = GdipDrawCurve(graphics, pen, @rgPoints(0), 5)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTranslateMatrix

Updates a matrix with the product of itself and a translation matrix.

```
FUNCTION GdipTranslateMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL offsetX AS REAL, _
   BYVAL offsetY AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *offsetX* | [in] Single precision value that specifies the horizontal component of the translation. |
| *offsetY* | [in] Single precision value that specifies the vertical component of the translation. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the translation matrix is on the left, and *MatrixOrderAppend* specifies that the translation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a Matrix object and calls the GdipRotateMatrix function to
' set the elements of that matrix to a rotation. Then the code calls the GdipTranslateMatrix
' function to update the matrix with the product of itself and a translation matrix. At that
' point, the matrix represents a composite transformation: first rotate, then translate.
' The code uses the matrix to set the world transformation of a Graphics object and then
' draws an ellipse that is transformed according to the composite transformation.
' ========================================================================================
SUB Example_TranslateMatrix (BYVAL hdc AS HDC)

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

   ' // Create identity matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)

   ' // First a rotation
   hStatus = GdipRotateMatrix(matrix, 30, MatrixOrderAppend)
   ' // Then a translation
   hStatus = GdipTranslateMatrix(matrix, 150 * rxRatio, 100 * ryRatio, MatrixOrderAppend)

   ' // Apply matrix to graphics
   hStatus = GdipSetWorldTransform(graphics, matrix)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 1 * rxRatio, UnitPixel, @pen)

   ' // Scale graphics according the dpi ratios
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Draw a tramsformed ellipse. The composite transformation
   ' // is rotate 30 degrees, then translate 150 right and 100 down.
   hstatus = GdipDrawEllipse(graphics, pen, -40, -20, 80, 40)

   ' // Draw rotated axes with the origin at the center of the ellipse.
   hStatus = GdipDrawLine(graphics, pen, -50, 0, 50, 0)
   hStatus = GdipDrawLine(graphics, pen, 0, -50, 0, 50)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipVectorTransformMatrixPoints

Multiplies each vector in an array by a matrix. The translation elements of a matrix (third row) are ignored. Each vector is treated as a row matrix. The multiplication is performed with the row matrix on the left and a matrix on the right.

```
FUNCTION GdipVectorTransformMatrixPoints (BYVAL matrix AS GpMatrix PTR, _
   BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
```
```
FUNCTION GdipVectorTransformMatrixPointsI (BYVAL matrix AS GpMatrix PTR, _
   BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | [in] Pointer to the **Matrix** object. |
| *pts* | [in, out] Pointer to an array of **GpPointF** objects that, on input, contains the vectors to be transformed and, on output, receives the transformed vectors. Each vector in the array is transformed (multiplied by a matrix) and updated with the result of the transformation. |
| *count* | [in] Long integer value that specifies the number of points to be transformed. The default value is 1. |

#### Example

```
' ========================================================================================
' The following example creates a vector and a point. The tip of the vector and the point
' are at the same location: (100, 50). The code creates a Matrix object and initializes its
' elements so that it represents a clockwise rotation followed by a translation 100 units
' to the right. The code calls the GdipTransformMatrixPoints functioon of the matrix
' to transform the point and calls the GdipVectorTransformMatrixPoints function to transform
' the vector. The entire transformation (rotation followed by translation) is performed on
' the point, but only the rotation part of the transformation is performed on the vector.
' The elements of the matrix that represent translation are ignored by the
' GdipVectorTransformMatrixPoints function
' ========================================================================================
SUB Example_VectorTransformMatrixPoints (BYVAL hdc AS HDC)

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
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create pen
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_BLUE, 2 * rxRatio, UnitPixel, @pen)
   hStatus = GdipSetPenEndCap(pen, LineCapArrowAnchor)
   
   ' // Create a solid brush
   DIM brush AS GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_BLUE, @brush)
   
   ' // A point and a vector, same representation but different behavior
   DIM pt AS GpPointF = (100, 50)
   DIM vector AS GpPointF = (100, 50)

   ' // Draw the original point and vector in blue
   hStatus = GdipFillEllipse(graphics, brush, pt.x - 5.0, pt.y - 5.0, 10.0, 10.0)
   hStatus = GdipDrawLine(graphics, pen, 0, 0, vector.x, vector.y)

   ' // Transform
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(0.8, 0.6, -0.6, 0.8, 100.0, 0.0,@matrix)
   hStatus = GdipTransformMatrixPoints(matrix, @pt, 1)
   hStatus = GdipVectorTransformMatrixPoints(matrix, @vector, 1)

   ' // Draw the transformed point and vector in red.
   hStatus = GdipSetPenColor(pen, ARGB_RED)
   hStatus = GdipSetSolidFillColor(brush, ARGB_RED)
   hStatus = GdipFillEllipse(graphics, brush, pt.X - 5.0, pt.Y - 5.0, 10.0, 10.0)
   hStatus = GdipDrawLine(graphics, pen, 0, 0, vector.x, vector.y)

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF brush THEN GdipDeleteBrush(brush)
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
