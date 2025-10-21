# ImageAttributes Functions

An **ImageAttributes** object contains information about how bitmap and metafile colors are manipulated during rendering. An **ImageAttributes** object maintains several color-adjustment settings, including color-adjustment matrices, grayscale-adjustment matrices, gamma-correction values, color-map tables, and color-threshold values.

| Name       | Description |
| ---------- | ----------- |
| [GdipCloneImageAttributes](#gdipcloneimageattributes) | Makes a copy of an **ImageAttributes** object. |
| [GdipCreateImageAttributes](#gdipcreateimageattributes) | Creates an **ImageAttributes** object. |
| [GdipDisposeImageAttributes](#gdipdisposeimageattributes) | Disposes the specified **ImageAttributes** object. |
| [GdipGetImageAttributesAdjustedPalette](#gdipgetimageattributesadjustedpalette) | Adjusts the colors in a palette according to the adjustment settings of a specified category. |
| [GdipResetImageAttributes](#gdipresetimageattributes) | Adjusts the colors in a palette according to the adjustment settings of a specified category. |
| [GdipSetImageAttributesCachedBackground](#gdipsetimageattributescachedbackground) | Sets or clears the *CachedBackground* member of a specified **GpImageAttributes** object. Microsoft Windows GDI+ does not use the *CachedBackground* member, so calling this function has no effect. |
| [GdipSetImageAttributesColorKeys](#gdipsetimageattributescolorkeys) | Sets the color key (transparency range) for a specified category. |
| [GdipSetImageAttributesColorMatrix](#gdipsetimageattributescolormatrix) | Sets the color-adjustment matrix for a specified category. |
| [GdipSetImageAttributesGamma](#gdipsetimageattributesgamma) | Sets the gamma value for a specified category. |
| [GdipSetImageAttributesICMMode](#gdipsetimageattributesicmmode) | Sets an internal state variable to the value specified by the on parameter. |
| [GdipSetImageAttributesNoOp](#gdipsetimageattributesnoop) | Turns off color adjustment for a specified category. |
| [GdipSetImageAttributesOutputChannel](#gdipsetimageattributesoutputchannel) | Sets the CMYK output channel for a specified category. |
| [GdipSetImageAttributesOutputChannelColorProfile](#gdipsetimageattributesoutputchannelcolorprofile) | Sets the CMYK output channel for a specified category. |
| [GdipSetImageAttributesRemapTable](#gdipsetimageattributesremaptable) | Sets the color-remap table for a specified category. |
| [GdipSetImageAttributesThreshold](#gdipsetimageattributesthreshold) | Sets the threshold (transparency range) for a specified category. |
| [GdipSetImageAttributesToIdentity](#gdipsetimageattributestoidentity) | Sets the color-adjustment matrix of a specified category to identity matrix. |
| [GdipSetImageAttributesWrapMode](#gdipsetimageattributeswrapmode) | Sets the wrap mode of the **ImageAttributes** object. |

---

## GdipCloneImageAttributes

Makes a copy of an **ImageAttributes** object.

```
FUNCTION GdipCloneImageAttributes (BYVAL imageattr AS CONST GpImageAttributes PTR, _
    BYVAL cloneImageattr AS GpImageAttributes PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *cloneImageattr* | out] Pointer to a variable that receives a pointer to the new **ImageAttributes** object. |

#### Example

```
' ========================================================================================
' This example loads an image, creates an ImageAttributes object with a wrap mode,
' clones it using GdipCloneImageAttributes, and draws the same image twice using
' the original and cloned attributes.
' ========================================================================================
SUB Example_CloneImageAttributes (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("climber.jpg", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Set the wrap mode to WrapModeTile
   GdipSetImageAttributesWrapMode(imgAttr, WrapModeTile, ARGB_RED, FALSE)

   ' // Clone the ImageAttributes
   DIM cloneAttr AS GpImageAttributes PTR
   hStatus = GdipCloneImageAttributes(imgAttr, @cloneAttr)

   ' // Draw the image
   hStatus = GdipDrawImageRectRect(graphics, image, _
             10, 10, nWidth, nHeight, _         ' dest rect
             0, 0, 2 * nWidth, 2 * nHeight, _   ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Draw the image using cloned attributes
   hStatus = GdipDrawImageRectRect(graphics, image, _
             200, 10, nWidth, nHeight, _        ' dest rect
             0, 0, 2 * nWidth, 2 * nHeight, _   ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF cloneAttr THEN GdipDisposeImageAttributes(cloneAttr)
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateImageAttributes

Creates an **ImageAttributes** object.

```
FUNCTION GdipCreateImageAttributes (BYVAL imageattr AS GpImageAttributes PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [out] Pointer to a variable that receives a pointer to the new created **ImageAttributes** object. |

#### Example

```
' ========================================================================================
' The following example creates an Image object loading an image from file, creates an
' ImageAttributes object, sets the wrap mode to WrapModeTile and draws the specified
' source rectangle to the destination rectangle, which is four times as large as the
' image itself.
' ========================================================================================
SUB Example_CreateImageAttributes (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("climber.jpg", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Set the wrap mode to WrapModeTile
   GdipSetImageAttributesWrapMode(imgAttr, WrapModeTile, ARGB_BLUE, FALSE)

   ' // Draw the image
   hStatus = GdipDrawImageRectRect(graphics, image, _
             10, 10, nWidth, nHeight, _         ' dest rect
             0, 0, 2 * nWidth, 2 * nHeight, _   ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDisposeImageAttributes

Disposes the specified **ImageAttributes** object.

```
FUNCTION GdipDisposeImageAttributes (BYVAL imageattr AS GpImageAttributes PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object to dispose. |

#### Usage example

```
DIM imgAttr AS GpImageAttributes PTR
DIM hStatus AS LONG

' // Create the ImageAttributes object
hStatus = GdipCreateImageAttributes(@imgAttr)

' // ... apply adjustments, use in rendering ...

' // Dispose when done
IF imgAttr THEN
   hStatus = GdipDisposeImageAttributes(imgAttr)
END IF
```
---

## GdipGetImageAttributesAdjustedPalette

Adjusts the colors in a palette according to the adjustment settings of a specified category.

```
FUNCTION GdipGetImageAttributesAdjustedPalette (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL colorPalette AS ColorPalette PTR, BYVAL colorAdjustType AS ColorAdjustType) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *colorPalette* | [in, out] Pointer to a **ColorPalette** structure that on input, contains the palette to be adjusted and, on output, receives the adjusted palette. |
| *colorAdjustType* | [in] Element of the **ColorAdjustType** enumeration that specifies the category whose adjustment settings will be applied to the palette. |

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color-remap table for the default category, a different color-remap table for the bitmap category, and still a different color-remap table for the pen category.

When you call **GdipGetImageAttributesAdjustedPalette**, you can specify the adjustment category that is used to adjust the palette colors. For example, if you pass **ColorAdjustTypeBitmap** to the **GdipGetImageAttributesAdjustedPalette** function, then the adjustment settings of the bitmap category are used to adjust the palette colors.

#### Example

```
' ========================================================================================
' The following example initializes a ColorPalette structure with four colors: aqua, black,
' red, and green. The code also creates an ImageAttributes object and sets its bitmap remap
' table so that green will be converted to blue. Then the code adjusts the palette colors by
' passing the address of the palette to the GdipGetAdjustedPalette function. The code displays
' the four palette colors twice: once before the adjustment and once after the adjustment.
' ========================================================================================
SUB GDIP_GetAdjustedPalette  (BYVAL hdc AS HDC)

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

   ' // Create a palette that has four entries.
   DIM clrPalette AS COLORPALETTE
   clrPalette.Flags = 0
   clrPalette.Count = 4

   clrPalette.Entries(0) = ARGB_Aqua
   clrPalette.Entries(1) = ARGB_Black
   clrPalette.Entries(2) = ARGB_Red
   clrPalette.Entries(3) = ARGB_Green

   ' // Create a SolidBrush
   DIM brush As GpSolidFill PTR
   hStatus = GdipCreateSolidFill(ARGB_Black, @brush)
   
   ' // Display the four palette colors with no adjustment.
   FOR j AS LONG = 0 TO 3
      hStatus = GdipSetSolidFillColor(brush, clrPalette.Entries(j))
      hStatus = GdipFillRectangle (graphics, brush, 30*j, 0, 20, 20)
   NEXT
   
   ' // Create a remap table that converts green to blue.
   DIM map AS GpColorMap
   map.oldColor = ARGB_Green
   map.newColor = ARGB_Blue

   ' // Create an ImageAttributes object, and set its bitmap remap table.
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)
   hStatus = GdipSetImageAttributesRemapTable(imgAttr, _
             ColorAdjustTypeBitmap, TRUE, 1, @map)

   ' // Adjust the palette.
   hStatus = GdipGetImageAttributesAdjustedPalette(imgAttr, @clrPalette, ColorAdjustTypeBitmap)

   ' // Display the four palette colors after the adjustment.
   FOR j AS LONG = 0 TO 3
      hStatus = GdipSetSolidFillColor(brush, clrPalette.Entries(j))
      hStatus = GdipFillRectangle (graphics, brush, 30 * j, 30, 20, 20)
   NEXT

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF brush THEN GdipDeleteBrush(brush)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipResetImageAttributes

Adjusts the colors in a palette according to the adjustment settings of a specified category.

```
FUNCTION GdipResetImageAttributes (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which color adjustment is reset. The default value is *ColorAdjustTypeDefault*. |

#### Example

```
' ========================================================================================
' This example sets a color remap table, draws an adjusted image, resets the attributes,
' and then draws the original image again to show the effect of GdipResetImageAttributes.
' ========================================================================================
SUB Example_ResetImageAttributes (BYVAL hdc AS HDC)

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

   ' // Load an image
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("RedGreenStripes.bmp", @image)
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)

   ' // Get image dimensions
   DIM nWidth AS LONG, nHeight AS LONG
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Set a remap table: red to blue
   DIM cMap AS GpColorMap
   cMap.oldColor = ARGB_RED
   cMap.newColor = ARGB_BLUE
   hStatus = GdipSetImageAttributesRemapTable(imgAttr, ColorAdjustTypeDefault, TRUE, 1, @cMap)

   ' // Draw the image with color remap applied
   hStatus = GdipDrawImageRectRect(graphics, image, _
      10, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Reset the ImageAttributes for the default category
   hStatus = GdipResetImageAttributes(imgAttr, ColorAdjustTypeDefault)

   ' // Draw the image again with reset attributes (should appear unaltered)
   hStatus = GdipDrawImageRectRect(graphics, image, _
      120, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesCachedBackground

Sets or clears the **CachedBackground** member of a specified **GpImageAttributes** object. Microsoft Windows GDI+ does not use the **CachedBackground** member, so calling this function has no effect. The imageattr parameter specifies the **GpImageAttributes** object. The *enableFlag* parameter specifies whether the **CachedBackground** member is set (TRUE) or cleared (FALSE).

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *enableFlag* | [in] Boolean value that specifies whether the **CachedBackground** member is set (TRUE) or cleared (FALSE). |

---

## GdipSetImageAttributesColorKeys

Sets the color key (transparency range) for a specified category.

```
FUNCTION GdipResetImageAttributes (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType, BYVAL enableFlag AS BOOL, BYVAL colorLow AS ARGB, _
   BYVAL colorHigh AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is set. The default value is *ColorAdjustTypeDefault*. |
| *enableFlag* | [in] Boolean value that specifies whether a separate transparency range is enabled for the category specified by the type parameter. |
| *colorLow* | [in] ARGB color that specifies the low color-key value. |
| *colorHigh* | [in] ARGB color that specifies the high color-key value. |

#### Example

```
' ========================================================================================
' This example sets a color key to make green pixels transparent, then draws the image
' with and without the transparency effect.
' ========================================================================================
SUB Example_SetColorKeys (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("RedGreenStripes.bmp", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Set color key: make green transparent
   hStatus = GdipSetImageAttributesColorKeys(imgAttr, ColorAdjustTypeDefault, TRUE, GDIP_ARGB(255,0,128,0), GDIP_ARGB(255,0,255,0))

   ' // Draw original image (no transparency)
   hStatus = GdipDrawImageRect(graphics, image, 10, 10, nWidth, nHeight)

   ' // Draw image with transparency applied
   hStatus = GdipDrawImageRectRect(graphics, image, _
      120, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesColorMatrix

Sets the color-adjustment matrix for a specified category.

```
FUNCTION GdipSetImageAttributesColorMatrix (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType, BYVAL enableFlag AS BOOL, BYVAL colorMatrix AS ColorMatrix PTR, _
   BYVAL grayMatrix AS ColorMatrix PTR, BYVAL flags AS ColorMatrixFlags) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which the color-adjustment matrix is set. The default value is *ColorAdjustTypeDefault*. |
| *enableFlag* | [in] Boolean value that specifies whether a separate color adjustment is enabled for the category specified by the type parameter. |
| *colorMatrix* | [in] Pointer to a 5x5 color-adjustment matrix. |
| *grayMatrix* | [in] Specifies a matrix to be used for adjusting gray shades when the value of the flags parameter is *ColorMatrixFlagsAltGray*. |
| *flags* | [in] Element of the **ColorMatrixFlags** enumeration that specifies the type of image and color that will be affected by the color-adjustment matrix. |

#### Example

```
' ========================================================================================
' This example applies a grayscale color matrix to an image using GdipSetImageAttributesColorMatrix.
' ========================================================================================
SUB Example_SetColorMatrix (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("climber.jpg", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Define grayscale color matrix
   DIM grayMatrix AS ColorMatrix
   grayMatrix.m(0, 0) = 0.3 : grayMatrix.m(0, 1) = 0.3 : grayMatrix.m(0, 2) = 0.3 : grayMatrix.m(0, 3) = 0 : grayMatrix.m(0, 4) = 0
   grayMatrix.m(1, 0) = 0.59: grayMatrix.m(1, 1) = 0.59: grayMatrix.m(1, 2) = 0.59: grayMatrix.m(1, 3) = 0 : grayMatrix.m(1, 4) = 0
   grayMatrix.m(2, 0) = 0.11: grayMatrix.m(2, 1) = 0.11: grayMatrix.m(2, 2) = 0.11: grayMatrix.m(2, 3) = 0 : grayMatrix.m(2, 4) = 0
   grayMatrix.m(3, 0) = 0   : grayMatrix.m(3, 1) = 0   : grayMatrix.m(3, 2) = 0   : grayMatrix.m(3, 3) = 1 : grayMatrix.m(3, 4) = 0
   grayMatrix.m(4, 0) = 0   : grayMatrix.m(4, 1) = 0   : grayMatrix.m(4, 2) = 0   : grayMatrix.m(4, 3) = 0 : grayMatrix.m(4, 4) = 1

   ' // Apply the color matrix
   hStatus = GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeDefault, TRUE, @grayMatrix, NULL, ColorMatrixFlagsDefault)

   ' // Draw original image
   hStatus = GdipDrawImageRect(graphics, image, 10, 10, nWidth, nHeight)

   ' // Draw grayscale image
   hStatus = GdipDrawImageRectRect(graphics, image, _
      200, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesGamma

Sets the gamma value for a specified category.

```
FUNCTION GdipSetImageAttributesGamma (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType, BYVAL enableFlag AS BOOL, BYVAL gamma AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which the gamma value is set. The default value is *ColorAdjustTypeDefault*. |
| *enableFlag* | [in] Boolean value that specifies whether a separate gamma is enabled for the category specified by the type parameter. |
| *gamma* | [in] Single precision value number that specifies the gamma value. |

#### Example

```
' ========================================================================================
' This example applies gamma correction to an image using GdipSetImageAttributesGamma.
' ========================================================================================
SUB Example_SetGamma (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("climber.jpg", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Apply gamma correction (e.g., 2.2)
   hStatus = GdipSetImageAttributesGamma(imgAttr, ColorAdjustTypeDefault, TRUE, 2.2)

   ' // Draw original image
   hStatus = GdipDrawImageRect(graphics, image, 10, 10, nWidth, nHeight)

   ' // Draw gamma-adjusted image
   hStatus = GdipDrawImageRectRect(graphics, image, _
      200, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesICMMode

Unsupported. This function sets an internal state variable to the value specified by the on parameter.

```
FUNCTION GdipSetImageAttributesICMMode (BYVAL imageattr AS GpImageAttributes PTR, BYVAL on AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *on* | [in] If this value is true, Image Color Management (ICM) is used for all color adjustment. If the value is false, ICM is not used. |

---

## GdipSetImageAttributesNoOp

Turns off color adjustment for a specified category.

```
FUNCTION GdipSetImageAttributesNoOp (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType, BYVAL enableFlag AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which color correction is turned off. The default value is *ColorAdjustTypeDefault*. |
| *enableFlag* | [in] Boolean value that specifies whether a color adjustment is enabled for the category specified by the type parameter. |

#### Example

```
' ========================================================================================
' This example sets a color remap, disables it with NoOp, and shows the difference.
' ========================================================================================
SUB Example_SetNoOp (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("RedGreenStripes.bmp", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Set color remap: red to blue
   DIM cMap AS GpColorMap
   cMap.oldColor = ARGB_RED
   cMap.newColor = ARGB_BLUE
   hStatus = GdipSetImageAttributesRemapTable(imgAttr, ColorAdjustTypeDefault, TRUE, 1, @cMap)

   ' // Draw image with remap applied
   hStatus = GdipDrawImageRectRect(graphics, image, _
      10, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Disable color adjustment using NoOp
   hStatus = GdipSetImageAttributesNoOp(imgAttr, ColorAdjustTypeDefault, TRUE)

   ' // Draw image again (should appear unaltered)
   hStatus = GdipDrawImageRectRect(graphics, image, _
      100, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesOutputChannel

Sets the CMYK output channel for a specified category.

```
FUNCTION GdipSetImageAttributesOutputChannel (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType, BYVAL enableFlag AS BOOL, BYVAL channelFlags AS ColorChannelFlags) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which the output channel is set. The default value is *ColorAdjustTypeDefault*. |
| *enableFlag* | [in]  Boolean value that specifies whether a separate output channel is enabled for the category specified by the type parameter. |
| *channelFlags* | [in] Element of the **ColorChannelFlags** enumeration that specifies the output channel. |

#### Example

```
' ========================================================================================
' The following example creates an Image object and calls the GdipDrawImageRect function
' to draw the image. Then the code creates an ImageAttributes object and sets its bitmap
' output channel to cyan (ColorChannelFlagsC). The code calls GdipDrawImageRect a second
' time, passing the address of the Image object and the address of the ImageAttributes object.
' The cyan channel of each pixel is calculated, and the rendered image shows the intensities
' of the cyan channel as shades of gray. The code calls GdipDrawImageRect three more times
' to show the intensities of the magenta, yellow, and black channels.
' ========================================================================================
SUB Example_SetOutputChannel (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("Mosaic2.bmp", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Draw the image unaltered.
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)
   hStatus = GdipDrawImageRect(graphics, image, 10, 10, nWidth, nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Draw the image, showing the intensity of the cyan channel.
   hStatus = GdipSetImageAttributesOutputChannel(imgAttr, ColorAdjustTypeBitmap, TRUE, ColorChannelFlagsC)
   hStatus = GdipDrawImageRectRect(graphics, image, _
             110, 10, nWidth, nHeight, _   ' dest rect
             0, 0, nWidth, nHeight, _      ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Draw the image, showing the intensity of the magenta channel.
   hStatus = GdipSetImageAttributesOutputChannel(imgAttr, ColorAdjustTypeBitmap, TRUE, ColorChannelFlagsM)
   hStatus = GdipDrawImageRectRect(graphics, image, _
             210, 10, nWidth, nHeight, _   ' dest rect
             0, 0, nWidth, nHeight, _      ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Draw the image, showing the intensity of the yellow channel.
   hStatus = GdipSetImageAttributesOutputChannel(imgAttr, ColorAdjustTypeBitmap, TRUE, ColorChannelFlagsY)
   hStatus = GdipDrawImageRectRect(graphics, image, _
             10, 110, nWidth, nHeight, _   ' dest rect
             0, 0, nWidth, nHeight, _      ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Draw the image, showing the intensity of the black channel.
   hStatus = GdipSetImageAttributesOutputChannel(imgAttr, ColorAdjustTypeBitmap, TRUE, ColorChannelFlagsK)
   hStatus = GdipDrawImageRectRect(graphics, image, _
             110, 110, nWidth, nHeight, _  ' dest rect
             0, 0, nWidth, nHeight, _      ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesOutputChannelColorProfile

Sets the output channel color-profile file for a specified category.

```
FUNCTION GdipSetImageAttributesOutputChannelColorProfile (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType, BYVAL enableFlag AS BOOL, BYVAL colorProfileFilename AS CONST WSTRING PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which the output channel color-profile file is set. The default value is *ColorAdjustTypeDefault*. |
| *enableFlag* | [in]  Boolean value that specifies whether a separate output channel color profile is enabled for the category specified by the type parameter. |
| *colorProfileName* | [in] Path name of a color-profile file. If the color-profile file is in the %SystemRoot%\System32\Spool\Drivers\Color directory, then this parameter can be the file name. Otherwise, this parameter must be the fully-qualified path name. |

#### Example

```
' ========================================================================================
' This example sets a CMYK output channel color profile using GdipSetImageAttributesOutputChannelColorProfile.
' ========================================================================================
SUB Example_SetOutputChannelColorProfile (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("climber.jpg", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Draw the image unaltered.
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)
   hStatus = GdipDrawImageRect(graphics, image, 10, 10, nWidth, nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Set CMYK color profile (must be a valid ICC file path)
   DIM colorProfileFilename AS WSTRING *MAX_PATH = "C:\Program Files\Bullzip\PDF Printer\icc\sRGB_IEC61966-2-1_no_black_scaling.icc"
   hStatus = GdipSetImageAttributesOutputChannelColorProfile(imgAttr, _
      ColorAdjustTypeDefault, TRUE, colorProfileFilename)

   ' // Draw image with color profile applied
   hStatus = GdipDrawImageRectRect(graphics, image, _
      10, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesRemapTable

Sets the color-remap table for a specified category.

```
FUNCTION GdipSetImageAttributesRemapTable (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType, BYVAL enableFlag AS BOOL, BYVAL mapSize AS UINT, _
   BYVAL map AS GpColorMap) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which the color-remap table is set. The default value is *ColorAdjustTypeDefault*. |
| *enableFlag* | [in]  Boolean value that specifies whether a separate color remap table is enabled for the category specified by the type parameter. |
| *mapSize* | [in] Long integer value that specifies the number of elements in the map array. |
| *map* | [in] Pointer to an array of **GpColorMap** structures that defines the color map. |

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a .bmp file and then draws the image.
' The code creates an ImageAttributes object and sets its default remap table so that red
' is converted to blue. Then the code draws the image again using the color adjustment
' specified by the remap table.
' ========================================================================================
SUB Example_SetRemapTable (BYVAL hdc AS HDC)

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

   ' // Create an Image object based on a .bmp file.
   ' // The image has one red stripe and one green stripe.
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("RedGreenStripes.bmp", @image)

   ' // Create an ImageAttributes object and set its remap table.
   DIM cMap AS GpColorMap
   cMap.oldColor = ARGB_RED
   cMap.newColor = ARGB_BLUE
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)
   hStatus = GdipSetImageAttributesRemapTable(imgAttr, ColorAdjustTypeDefault, TRUE, 1, @cMap)

   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)
   
   ' // Draw the image with no color adjustment
   hStatus = GdipDrawImageRect(graphics, image, 10, 10, nWidth, nHeight)

   ' // Draw the image with red converted to blue.
   hStatus = GdipDrawImageRectRect(graphics, image, _
             100, 10, nWidth, nHeight, _       ' // dest rect
             0, 0, nWidth, nHeight, _          ' // source rect
             UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesThreshold

Sets the threshold (transparency range) for a specified category.

```
FUNCTION GdipSetImageAttributesThreshold (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType, BYVAL enableFlag AS BOOL, BYVAL threshold AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which the color threshold is set. The default value is *ColorAdjustTypeDefault*. |
| *enableFlag* | [in]  Boolean value that specifies whether a separate threshold is enabled for the category specified by the type parameter. |
| *threshold* | [in] Single precision value number that specifies the threshold value. |

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a .bmp file. The code also creates
' an ImageAttributes object and sets its bitmap threshold value to 0.6. Then the code draws
' the image twice: once with no color adjustment and once with the adjustment specified by
' the threshold.
' ========================================================================================
SUB Example_SetThreshold (BYVAL hdc AS HDC)

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

   ' // Create an Image object based on a .bmp file.
   ' // The image has one stripe with RGB components (160, 0, 0)
   ' // and one stripe with RGB components (0, 140, 0).
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("RedGreenThreshold.bmp", @image)

   ' // Create an ImageAttributes object, and set its bitmap threshold to 0.6.
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)
   hStatus = GdipSetImageAttributesThreshold(imgAttr, ColorAdjustTypeBitmap, TRUE, 0.6)

   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Draw the image with no color adjustment.
   hStatus = GdipDrawImageRect(graphics, image, 10, 10, nWidth, nHeight)

   ' // Draw the image with the threshold applied.
   ' // 160 > 0.6*255, so the red stripe will be changed to full intensity.
   ' // 140 < 0.6*255, so the green stripe will be changed to zero intensity.
   hStatus = GdipDrawImageRectRect(graphics, image, _
             100, 10, nWidth, nHeight, _   ' dest rect
             0, 0, nWidth, nHeight, _      ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesToIdentity

Sets the color-adjustment matrix of a specified category to identity matrix.

```
FUNCTION GdipSetImageAttributesToIdentity (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL type AS ColorAdjustType) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *type* | [in] Element of the **ColorAdjustType** enumeration that specifies the category for which the color-adjustment matrix is set to identity. The default value is *ColorAdjustTypeDefault*. |

#### Example

```
' ========================================================================================
' This example applies a grayscale matrix, then resets it using GdipSetImageAttributesToIdentity.
' ========================================================================================
SUB Example_SetToIdentity (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("climber.jpg", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Define grayscale matrix
   DIM grayMatrix AS ColorMatrix
   grayMatrix.m(0, 0) = 0.3 : grayMatrix.m(0, 1) = 0.3 : grayMatrix.m(0, 2) = 0.3 : grayMatrix.m(0, 3) = 0 : grayMatrix.m(0, 4) = 0
   grayMatrix.m(1, 0) = 0.59: grayMatrix.m(1, 1) = 0.59: grayMatrix.m(1, 2) = 0.59: grayMatrix.m(1, 3) = 0 : grayMatrix.m(1, 4) = 0
   grayMatrix.m(2, 0) = 0.11: grayMatrix.m(2, 1) = 0.11: grayMatrix.m(2, 2) = 0.11: grayMatrix.m(2, 3) = 0 : grayMatrix.m(2, 4) = 0
   grayMatrix.m(3, 0) = 0   : grayMatrix.m(3, 1) = 0   : grayMatrix.m(3, 2) = 0   : grayMatrix.m(3, 3) = 1 : grayMatrix.m(3, 4) = 0
   grayMatrix.m(4, 0) = 0   : grayMatrix.m(4, 1) = 0   : grayMatrix.m(4, 2) = 0   : grayMatrix.m(4, 3) = 0 : grayMatrix.m(4, 4) = 1

   ' // Apply grayscale matrix
   hStatus = GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeDefault, TRUE, @grayMatrix, NULL, ColorMatrixFlagsDefault)

   ' // Draw grayscale image
   hStatus = GdipDrawImageRectRect(graphics, image, _
      10, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Reset to identity matrix (remove grayscale)
   hStatus = GdipSetImageAttributesToIdentity(imgAttr, ColorAdjustTypeDefault)

   ' // Draw original image (no color adjustment)
   hStatus = GdipDrawImageRectRect(graphics, image, _
      200, 10, nWidth, nHeight, _
      0, 0, nWidth, nHeight, _
      UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipSetImageAttributesWrapMode

Sets the wrap mode of the **ImageAttributes** object.

```
FUNCTION GdipSetImageAttributesWrapMode (BYVAL imageattr AS GpImageAttributes PTR, _
   BYVAL wrap AS WrapMode, BYVAL argb AS ARGB, BYVAL clamp AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *imageattr* | [in] Pointer to the **ImageAttributes** object. |
| *wrap* | [in] Element of the **WrapMode** enumeration that specifies how repeated copies of an image are used to tile an area. |
| *argb* | [in] ARGB color that specifies the color of pixels outside of a rendered image. This color is visible if the wrap parameter is set to *WrapModeClamp* and the source rectangle passed to **GdipDrawImage** is larger than the image itself. |

#### Example

```
' ========================================================================================
' The following example calls DrawImage twice. In both calls, the specified source rectangle
' is four times as large as the image itself. In the first call to DrawImage, the
' ImageAttributes object has its wrap mode set to WrapModeClamp, so the portion of the
' source rectangle outside the image is filled with red according to the second argument
' passed to ImageAttributes.SetWrapMode. In the second call to DrawImage, the ImageAttributes
' object has its wrap mode set to WrapModeTile, so the portion of the source rectangle
' outside the image is filled with repeated copies of the image.
' ========================================================================================
SUB Example_SetWrapMode (BYVAL hdc AS HDC)

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

   ' // Load an image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("climber.jpg", @image)
   ' // Set the resolution of this image object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(image, dpiX, dpiY)
   
   ' // Get the width and height of the image
   DIM AS LONG nWidth, nHeight
   hStatus = GdipGetImageWidth(image, @nWidth)
   hStatus = GdipGetImageHeight(image, @nHeight)

   ' // Create an ImageAttributes object
   DIM imgAttr AS GpImageAttributes PTR
   hStatus = GdipCreateImageAttributes(@imgAttr)

   ' // Set the wrap mode to WrapModeClamp (no tile)
   GdipSetImageAttributesWrapMode(imgAttr, WrapModeClamp, ARGB_RED, FALSE)

   ' // Draw the image
   hStatus = GdipDrawImageRectRect(graphics, image, _
             10, 10, nWidth, nHeight, _         ' dest rect
             0, 0, 2 * nWidth, 2 * nHeight, _   ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Change the wrap mode to WrapModeTile
   GdipSetImageAttributesWrapMode(imgAttr, WrapModeTile, ARGB_RED, FALSE)

   ' // Draw the image with the new attribute
   hStatus = GdipDrawImageRectRect(graphics, image, _
             200, 10, nWidth, nHeight, _        ' dest rect
             0, 0, 2 * nWidth, 2 * nHeight, _   ' source dest
             UnitPixel, imgAttr, NULL, NULL)

   ' // Cleanup
   IF imgAttr THEN GdipDisposeImageAttributes(imgAttr)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---
