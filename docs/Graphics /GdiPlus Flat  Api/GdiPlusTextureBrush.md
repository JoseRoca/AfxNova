## Texture Brush Functions

The **TextureBrush** object is a **Brush** object that contains an **Image** object that is used for the fill. The fill image can be transformed by using the local **Matrix** object contained in the **Brush** object.

| Name       | Description |
| ---------- | ----------- |
| [GdipCreateTexture](#gdipcreatetexture) | Creates a **TextureBrush** object based on an image and a wrap mode. The size of the brush defaults to the size of the image, so the entire image is used by the brush. |
| [GdipCreateTexture2](#gdipcreatetexture2) | Creates a **TextureBrush** object based on an image, a wrap mode, and a defining set of coordinates. |
| [GdipCreateTexture2I](#gdipcreatetexture2) | Creates a **TextureBrush** object based on an image, a wrap mode, and a defining set of coordinates. |
| [GdipCreateTextureIA](#gdipcreatetextureia) | Creates a **TextureBrush** object based on an image, a defining set of coordinates, and a set of image properties. |
| [GdipCreateTextureIAI](#gdipcreatetextureia) | Creates a **TextureBrush** object based on an image, a defining set of coordinates, and a set of image properties. |
| [GdipGetTextureImage](#gdipgettextureimage) | Gets a pointer to the **Image** object that is defined by this texture brush. |
| [GdipGetTextureTransform](#gdipgettexturetransform) | Gets the transformation matrix of this texture brush. |
| [GdipGetTextureWrapMode](#gdipgettexturewrapmode) | Gets the wrap mode currently set for this texture brush. |
| [GdipMultiplyTextureTransform](#gdipmultiplytexturetransform) | Updates this brush's transformation matrix with the product of itself and another matrix. |
| [GdipResetTextureTransform](#gdipresettexturetransform) | Resets the transformation matrix of this texture brush to the identity matrix. This means that no transformation takes place. |
| [GdipRotateTextureTransform](#gdiprotatetexturetransform) | Resets the transformation matrix of this texture brush to the identity matrix. This means that no transformation takes place. |
| [GdipScaleTextureTransform](#gdipscaletexturetransform) | Updates this texture brush's current transformation matrix with the product of itself and a scaling matrix. |
| [GdipSetTextureTransform](#gdipsettexturetransform) | Sets the transformation matrix of this texture brush. |
| [GdipSetTextureWrapMode](#gdipsettexturewrapmode) | Sets the transformation matrix of this texture brush. |
| [GdipTranslateTextureTransform](#gdiptranslatetexturetransform) | Updates this brush's current transformation matrix with the product of itself and a translation matrix. |

---

## GdipCreateTexture

Creates a **TextureBrush** object based on an image and a wrap mode. The size of the brush defaults to the size of the image, so the entire image is used by the brush.

```
FUNCTION GdipCreateTexture (BYVAL image AS GpImage PTR, BYVAL wrapMode AS GpWrapMode, _
   BYVAL texture AS GpTexture PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *image* | [in] Pointer to the **Image** object. |
| *wrapMode* | [in] Element of the **WrapMode** enumeration that specifies how repeated copies of an image are used to tile an area when it is painted with this texture brush. The default value is *WrapModeTile*. |
| *texture* | [out] Pointer to a variable that receives a pointer to the new created **TextureBrush** object. |

The **WrapMode** enumeration specifies how repeated copies of an image are used to tile an area.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *WrapModeTile* | 0 | Specifies tiling without flipping. |
| *WrapModeTileFlipX* | 1 | Specifies that tiles are flipped horizontally as you move from one tile to the next in a row. |
| *WrapModeTileFlipY* | 2 | Specifies that tiles are flipped vertically as you move from one tile to the next in a column. |
| *WrapModeTileFlipXY* | 3 | Specifies that tiles are flipped horizontally as you move along a row and flipped vertically as you move along a column. |
| *WrapModeClamp* | 4 | Specifies that no tiling takes place. |

#### Remarks

An area that extends beyond the boundaries of the brush is tiled with repeated copies of the brush. A texture brush may have alternate tiles flipped in a certain direction, as specified by the wrap mode. Flipping has the effect of reversing the brush's image. For example, if the wrap mode is specified as *WrapModeTileFlipX*, the brush is flipped about a line that is parallel to the y-axis.

The texture brush is always oriented at (0, 0). If the wrap mode is specified as *WrapModeClamp*, no area outside of the brush is tiled. For example, suppose you create a texture brush, specifying *WrapModeClamp* as the wrap mode:
```
hStatus = GdipCreateTextureBrush(someImage, WrapModeClamp, @texture)
```
Then, you paint an area with the brush. If the size of the brush has a height of 50 and the painted area is a rectangle with its upper-left corner at (0, 50), you will see no repeated copies of the brush (no tiling).

The default wrap mode for a texture brush is *WrapModeTile*, which specifies no flipping of the tile and no clamping.

#### Example

```
' ========================================================================================
' The following example creates a texture brush fills a rectangle.
' ========================================================================================
SUB Example_CreateTexture (BYVAL hdc AS HDC)

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

   ' // Load the image and set the dpi resolution
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("HouseAndTree.gif", @Image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)

   ' // Create a texture brush, and set its wrap mode.
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   ' // Fill a rectangle with the texture brush
   hStatus = GdipFillRectangle(graphics, textureBrush, 20, 20, 360, 210)

   ' // Cleanup
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateTexture2

Creates a **TextureBrush** object based on an image, a wrap mode, and a defining set of coordinates.

```
FUNCTION GdipCreateTexture2 (BYVAL image AS GpImage PTR, BYVAL wrapMode AS GpWrapMode, _
   BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL, _
   BYVAL texture AS GpTexture PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *image* | [in] Pointer to the **Image** object. |
| *wrapMode* | [in] Element of the **WrapMode** enumeration that specifies how repeated copies of an image are used to tile an area when it is painted with this texture brush. |
| *x* | [in] Leftmost coordinate of the image portion to be used by this brush. |
| *y* | [in] Uppermost coordinate of the image portion to be used by this brush. |
| *width* | [in] Width of the brush and width of the image portion to be used by the brush. |
| *height* | [in] Height of the brush and height of the image portion to be used by the brush. |
| *texture* | [out] Pointer to a variable that receives a pointer to the new created TextureBrush object. |

The **WrapMode** enumeration specifies how repeated copies of an image are used to tile an area.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *WrapModeTile* | 0 | Specifies tiling without flipping. |
| *WrapModeTileFlipX* | 1 | Specifies that tiles are flipped horizontally as you move from one tile to the next in a row. |
| *WrapModeTileFlipY* | 2 | Specifies that tiles are flipped vertically as you move from one tile to the next in a column. |
| *WrapModeTileFlipXY* | 3 | Specifies that tiles are flipped horizontally as you move along a row and flipped vertically as you move along a column. |
| *WrapModeClamp* | 4 | Specifies that no tiling takes place. |

#### Remarks

The *x*, *y*, *width*, and *height* parameters specify a rectangle. The size of the brush is defined by *width* and *height*. The *x* and *y* parameters have no affect on the brush's size or position the brush is always oriented at (0, 0). The *x*, *y*, *width*, and *height* parameters define the portion of the image to be used by the brush.

For example, suppose you have an image that is stored in an **Image** object and is 256 x 512 (width, height) pixels. Then you create a **TextureBrush** object based on this image as follows:
```
hStatus = GdipCreateTexture2(someImage, WrapModeTile, 12, 50, 100, 150, @texture).
```
The brush will have a width of 100 units and a height of 150 units. The brush will use a rectangular portion of the image. This portion begins at the pixel having coordinates (12, 50). The width and height of the portion are 100 and 150, respectively, measured from the starting pixel.

Now suppose you create another **TextureBrush** object based on the same image and specify a different rectangle:
```
hStatus = GdipCreateTexture2(someImage, WrapModeTile, 0, 0, 256, 512, @texture).
```
The brush will have width and height equal to 256 and 512, respectively. The brush will use the entire image instead of a portion of it because the rectangle specifies a starting pixel at coordinates (0, 0) and dimensions identical to those of the image.

#### Example

```
' ========================================================================================
' The following example creates a texture brush fills a rectangle.
' ========================================================================================
SUB Example_CreateTexture2 (BYVAL hdc AS HDC)

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

   ' // Load the image and set the dpi resolution
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("HouseAndTree.gif", @Image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)

   ' // Create a texture brush, and set its wrap mode.
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture2(image, WrapModeTile, 5, 5, 55, 55, @textureBrush)
   ' // Fill a rectangle with the etxture brush
   hStatus = GdipFillRectangle(graphics, textureBrush, 20, 20, 360, 210)

   ' // Cleanup
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateTextureIA

Creates a **TextureBrush** object based on an image, a defining set of coordinates, and a set of image properties.

```
FUNCTION GdipCreateTextureIA (BYVAL image AS GpImage PTR, _
   BYVAL imageAttributes AS CONST GpImageAttributes PTR, _
   BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL, _
   BYVAL texture AS GpTexture PTR PTR) AS GpStatus
```
```
FUNCTION GdipCreateTextureIAI (BYVAL image AS GpImage PTR, _
   BYVAL imageAttributes AS CONST GpImageAttributes PTR, _
   BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_, _
   BYVAL texture AS GpTexture PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *image* | [in] Pointer to the **Image** object. |
| *imageAttributes* | [in] Pointer to an **ImageAttributes** object that contains properties of the image. The default value is NULL. |
| *x* | [in] Leftmost coordinate of the image portion to be used by this brush. |
| *y* | [in] Uppermost coordinate of the image portion to be used by this brush. |
| *width* | [in] Width of the brush and width of the image portion to be used by the brush. |
| *height* | [in] Height of the brush and height of the image portion to be used by the brush. |
| *texture* | [out] Pointer to a variable that receives a pointer to the new created **TextureBrush** object. |

#### Example

```
' ========================================================================================
' This example creates a texture brush based on a defined rectangular area of an image
' modified by image attributes to reduce brightness, scale the texture brush to tile across
' the fill area and fills a rectangle with the texture brush.
' ========================================================================================
SUB Example_GetTextureIA (BYVAL hdc AS HDC)

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

   DIM image AS GpImage PTR
   DIM wszFileName AS WSTRING * 64 = "HouseAndTree.gif"
   hStatus = GdipLoadImageFromFile(wszFileName, @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)

   ' // Create image attributes with a color matrix (e.g., reduce brightness)
   DIM attr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@attr)
   DIM clrMatrix(4, 4) AS SINGLE = {{0.5, 0.0, 0.0, 0.0, 0.0}, {0.0, 0.5, 0.0, 0.0, 0.0}, _
       {0.0, 0.0, 0.5, 0.0, 0.0}, {0.0, 0.0, 0.0, 1.0, 0.0}, {0.0, 0.0, 0.0, 0.0, 1.0}}
   hStatus = GdipSetImageAttributesColorMatrix(attr, ColorAdjustTypeDefault, 1, cast(ColorMatrix PTR, @clrMatrix(0, 0)), 0, 0)

   ' // Create texture brush using a portion of the image and the attributes
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTextureIA(image, attr, 10, 10, 55, 55, @textureBrush)

   ' // Scale the texture brush to tile across the fill area
   GdipSetTextureWrapMode(textureBrush, WrapModeTile)

   ' // Fill a rectangle with the texture brush
   hStatus = GdipFillRectangle(graphics, textureBrush, 20, 20, 360, 210)

   ' // Cleanup
   IF attr THEN GdipDisposeImageAttributes(attr)
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetTextureImage

Gets a pointer to the **Image** object that is defined by this texture brush.

```
FUNCTION GdipGetTextureImage (BYVAL brush AS GpTexture PTR, BYVAL image AS GpImage PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |
| *image* | [out] Pointer to a variable that receives a pointer to the **Image** object that is defined by this texture brush. |

#### Example

```
' ========================================================================================
' The following example creates a texture brush and uses it to fill an ellipse. The code
' then gets the brush's image and draws it.
' ========================================================================================
SUB Example_GetTextureImage (BYVAL hdc AS HDC)

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

   DIM image AS GpImage PTR
   DIM wszFileName AS WSTRING * 64 = "HouseAndTree.gif"
   hStatus = GdipLoadImageFromFile(wszFileName, @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   hStatus = GdipFillEllipse(graphics, textureBrush, 0, 0, 200, 150)

   ' // Get the brush's image, and draw that image.
   DIM image2 AS GpImage PTR
   hStatus = GdipGetTextureImage(textureBrush, @image2)
   hStatus = GdipDrawImage(graphics, image2, 10, 160)

   ' // Cleanup
   IF image THEN GdipDisposeImage(image)
   IF image2 THEN GdipDisposeImage(image2)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetTextureTransform

Gets the transformation matrix of this texture brush.

```
FUNCTION GdipGetTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |
| *matrix* | [out] Pointer to a **Matrix** object that receives the transformation matrix. |

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then gets the brush's transformation matrix and proceeds to inspect or use the
' elements.
' ========================================================================================
SUB Example_GdipGetTextureTransform (BYVAL hdc AS HDC)

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

   DIM image AS GpImage PTR
   DIM wszFileName AS WSTRING * 64 = "HouseAndTree.gif"
   hStatus = GdipLoadImageFromFile(wszFileName, @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   hStatus = GdipFillEllipse(graphics, textureBrush, 0, 0, 400, 250)

   ' // Obtain information about the texture brush.
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix(@matrix)
   DIM elements(23) AS SINGLE   ' 6 * SIZEOF(SINGLE), i.e. 24 elements

   hStatus = GdipGetTextureTransform(textureBrush, @matrix)
   hStatus = GdipGetMatrixElements(matrix, @elements(0))

   DIM j AS LONG
   FOR j = 0 TO 23
      ' // Inspect or use the value in elements[j].
      OutputDebugStringW WSTR(elements(j))
   NEXT

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetTextureWrapMode

Gets the wrap mode currently set for this texture brush.

```
FUNCTION GdipGetTextureWrapMode (BYVAL brush AS GpTexture PTR, BYVAL wrapMode AS GpWrapMode PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |
| *wrapMode* | [out] Pointer to a variable that receives the wrap mode currently set for this texture brush. The value returned is one of the elements of the **GpWrapMode** enumeration. |

#### Example

```
' ========================================================================================
' The following example creates a texture brush, sets the wrap mode of the brush, and uses
' the brush to fill a rectangle. Next, the code gets the wrap mode of the brush and stores
' the value. The code creates a second texture brush and uses the stored wrap mode to set
' the wrap mode of the second texture brush. Then, the code uses the second texture brush
' to fill a rectangle.
' ========================================================================================
SUB Example_GetTextureWrapMode (BYVAL hdc AS HDC)

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

   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("HouseAndTree.gif", @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   hStatus = GdipFillRectangleI(graphics, textureBrush, 0, 0, 200, 200)

   ' // Get the brush's wrap mode.
   DIM wrapMode AS LONG
   hStatus = GdipGetTextureWrapMode(textureBrush, @wrapMode)

   ' // Create a second texture brush with the same wrap mode.
   DIM image2 AS GpImage PTR
   hStatus = GdipLoadImageFromFile("MyTexture.png", @image2)
   hStatus = GdipBitmapSetResolution(image2, dpiX, dpiY)
   DIM textureBrush2 AS GpTexture PTR
   hStatus = GdipCreateTexture(image2, wrapMode, @textureBrush2)
   hStatus = GdipFillRectangle(graphics, textureBrush2, 200, 0, 200, 200)

   ' // Cleanup
   IF image2 THEN GdipDisposeImage(image2)
   IF textureBrush2 THEN GdipDeleteBrush(textureBrush2)
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipMultiplyTextureTransform

Updates this brush's transformation matrix with the product of itself and another matrix.

```
FUNCTION GdipMultiplyTextureTransform (BYVAL brush AS GpTexture PTR, _
   BYVAL matrix AS CONST GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |
| *matrix* | [in] Pointer to the matrix to be multiplied by this brush's current transformation matrix. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of multiplication. *MatrixOrderPrepend* specifies that the passed matrix is on the left, and *MatrixOrderAppend* specifies that the passed matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill a rectangle.
' ========================================================================================
SUB Example_MultiplyTextureTransform (BYVAL hdc AS HDC)

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

   ' // Load the image and set the dpi resolution
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("HouseAndTree.gif", @Image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)

   ' // Create a texture brush, and set its wrap mode.
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   ' // Rotate the image
   hStatus = GdipRotateTextureTransform(textureBrush, 30, MatrixOrderPrepend)
   ' // Horizontal stretch
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(3, 0, 0, 1, 0, 0, @matrix)
   ' // Stretch the image
   hStatus = GdipMultiplyTextureTransform(textureBrush, matrix, MatrixOrderAppend)
   ' // Fill a rectangle with the etxture brush
   hStatus = GdipFillRectangle(graphics, textureBrush, 20, 20, 360, 210)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipResetTextureTransform

Resets the transformation matrix of this texture brush to the identity matrix. This means that no transformation takes place.

```
FUNCTION GdipResetTextureTransform (BYVAL brush AS GpTexture PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill a rectangle.
' ========================================================================================
SUB Example_ResetTextureTransform (BYVAL hdc AS HDC)

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

   DIM image AS GpImage PTR
   DIM wszFileName AS WSTRING * 64 = "HouseAndTree.gif"
   hStatus = GdipLoadImageFromFile(wszFileName, @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   hStatus = GdipRotateTextureTransform(textureBrush, 30, MatrixOrderPrepend)

   ' // Fill a rectangle with the transformed texture brush.
   hStatus = GdipFillEllipse(graphics, textureBrush, 100, 20, 200, 100)

   ' // Reset the transformation matrix of this texture brush to the identity matrix.
   hStatus = GdipResetTextureTransform(textureBrush)

   ' // Fill a rectangle with the texture brush (no transformation).
   hStatus = GdipFillEllipse(graphics, textureBrush, 100, 140, 200, 100)


   ' // Cleanup
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipRotateTextureTransform

Resets the transformation matrix of this texture brush to the identity matrix. This means that no transformation takes place.

```
FUNCTION GdipRotateTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL angle AS REAL, _
   BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |
| *angle* | [in] Single precision value that specifies the angle, in degrees, of rotation. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the rotation matrix is on the left, and *MatrixOrderAppend* specifies that the rotation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill a rectangle.
' ========================================================================================
SUB Example_RotateTextureTransform (BYVAL hdc AS HDC)

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

   DIM image AS GpImage PTR
   DIM wszFileName AS WSTRING * 64 = "HouseAndTree.gif"
   hStatus = GdipLoadImageFromFile(wszFileName, @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   hStatus = GdipScaleTextureTransform(textureBrush, 3, 1, MatrixOrderPrepend)
   hStatus = GdipRotateTextureTransform(textureBrush, 30, MatrixOrderAppend)
   hStatus = GdipFillEllipse(graphics, textureBrush, 0, 0, 400, 250)

   ' // Cleanup
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipScaleTextureTransform

Updates this texture brush's current transformation matrix with the product of itself and a scaling matrix.

```
FUNCTION GdipScaleTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL sx AS REAL, _
   BYVAL sy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |
| *sx* | [in] Single precision value that specifies the amount to scale in the x direction. |
| *sy* | [in] Single precision value that specifies the amount to scale in the y direction. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the rotation matrix is on the left, and *MatrixOrderAppend* specifies that the rotation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation and the scale
' of the brush. The code then uses the transformed brush to fill a rectangle.
' ========================================================================================
SUB Example_ScaleTextureTransform (BYVAL hdc AS HDC)

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

   DIM image AS GpImage PTR
   DIM wszFileName AS WSTRING * 64 = "HouseAndTree.gif"
   hStatus = GdipLoadImageFromFile(wszFileName, @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   hStatus = GdipScaleTextureTransform(textureBrush, 3, 1, MatrixOrderPrepend)
   hStatus = GdipTranslateTextureTransform(textureBrush, 30, 0, MatrixOrderAppend)
   hStatus = GdipFillEllipse(graphics, textureBrush, 0, 0, 400, 250)

   ' // Cleanup
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetTextureTransform

Sets the transformation matrix of this texture brush.

```
FUNCTION GdipSetTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL matrix AS CONST GpMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |
| *matrix* | [in] Pointer to a Matrix object that specifies the transformation matrix to use. |

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill an ellipse.
' ========================================================================================
SUB Example_ScaleTextureTransform (BYVAL hdc AS HDC)

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

   ' // Create a matrix
   DIM matrix AS GpMatrix PTR
   hStatus = GdipCreateMatrix2(2, 0, 0, 1, 0, 0, matrix)   ' // Horizontal stretch

   DIM image AS GpImage PTR
   DIM wszFileName AS WSTRING * 64 = "HouseAndTree.gif"
   hStatus = GdipLoadImageFromFile(wszFileName, @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   hStatus = GdipSetTextureTransform(textureBrush, @matrix)
   hStatus = GdipFillEllipse(graphics, textureBrush, 0, 0, 400, 250)

   ' // Cleanup
   IF matrix THEN GdipDeleteMatrix(matrix)
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetTextureWrapMode

Sets the transformation matrix of this texture brush.

```
FUNCTION GdipSetTextureWrapMode (BYVAL brush AS GpTexture PTR, BYVAL wrapMode AS GpWrapMode) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |
| *wrapMode* | [in] Element of the **WrapMode** enumeration that specifies how repeated copies of an image are used to tile an area when it is painted with this texture brush. |

#### Remarks

An area that extends beyond the boundaries of the brush is tiled with repeated copies of the brush. A texture brush may have alternate tiles flipped in a certain direction, as specified by the wrap mode. Flipping has the effect of reversing the brush's image. For example, if the wrap mode is specified as WrapModeTileFlipX, the brush is flipped about a line that is parallel to the y-axis.

The texture brush is always oriented at (0, 0). If the wrap mode is specified as *WrapModeClamp*, no area outside of the brush is tiled. For example, suppose you create a texture brush, specifying *WrapModeClamp* as the wrap mode:
```
hStatus = GdipCreateTexture(pImage, %WrapModeClamp, pTextureBrush)
```
Then you paint an area with the brush. If the size of the brush has a height of 50 and the painted area is a rectangle with its upper-left corner at (0, 50), you will see no repeated copies of the brush (no tiling).

The default wrap mode for a texture brush is *WrapModeTile*, which specifies no flipping of the tile and no clamping.

#### Example

```
' ========================================================================================
' The following example creates a texture brush, sets the wrap mode of the brush, and uses
' the brush to fill an ellipse.
' ========================================================================================
SUB Example_SetTextureWrapMode (BYVAL hdc AS HDC)

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

   DIM image AS GpImage PTR
   DIM wszFileName AS WSTRING * 64 = "HouseAndTree.gif"
   hStatus = GdipLoadImageFromFile(wszFileName, @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   hStatus = GdipSetTextureWrapMode(textureBrush, WrapModeTileFlipX)
   hStatus = GdipFillEllipse(graphics, textureBrush, 0, 0, 400, 250)

   ' // Cleanup
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTranslateTextureTransform

Updates this brush's current transformation matrix with the product of itself and a translation matrix.

```
FUNCTION GdipTranslateTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL dx AS REAL, _
    BYVAL dy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *brush* | [in] Pointer to the **TextureBrush** object. |
| *dx* | [in] Single precision value that specifies the horizontal component of the translation. |
| *dy* | [in] Single precision value that specifies the vertical component of the translation. |
| *order* | [in] Element of the **MatrixOrder** enumeration that specifies the order of the multiplication. *MatrixOrderPrepend* specifies that the translation matrix is on the left, and *MatrixOrderAppend* specifies that the translation matrix is on the right. The default value is *MatrixOrderPrepend*. |

#### Example

```
' ========================================================================================
' The following example creates a texture brush and sets the transformation of the brush.
' The code then uses the transformed brush to fill an ellipse
' ========================================================================================
SUB Example_TranslateTextureTransform (BYVAL hdc AS HDC)

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

   DIM image AS GpImage PTR
   DIM wszFileName AS WSTRING * 64 = "HouseAndTree.gif"
   hStatus = GdipLoadImageFromFile(wszFileName, @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   DIM textureBrush AS GpTexture PTR
   hStatus = GdipCreateTexture(image, WrapModeTile, @textureBrush)
   hStatus = GdipTranslateTextureTransform(textureBrush, 30, 0, MatrixOrderAppend)
   hStatus = GdipFillEllipse(graphics, textureBrush, 0, 0, 400, 250)

   ' // Cleanup
   IF image THEN GdipDisposeImage(image)
   IF textureBrush THEN GdipDeleteBrush(textureBrush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
