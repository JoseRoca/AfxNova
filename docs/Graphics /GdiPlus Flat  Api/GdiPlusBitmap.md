# Bitmap Functions

The **Bitmap** functions expands on the capabilities of the Image functions by providing additional methods for creating and manipulating raster images.

| Name       | Description |
| ---------- | ----------- |
| [GdipBitmapApplyEffect](#gdipbitmapapplyeffect) | Alters a **Bitmap** object by applying a specified effect. |
| [GdipBitmapConvertFormat](#gdipbitmapconvertformat) | Converts a bitmap to a specified pixel format. The original pixel data in the bitmap is replaced by the new pixel data. |
| [GdipBitmapCreateApplyEffect](#gdipbitmapcreateapplyeffect) | Creates a new **Bitmap** object by applying a specified effect to an existing **Bitmap** object. |
| [GdipBitmapGetHistogram](#gdipbitmapgethistogram) | Returns one or more histograms for specified color channels of the specified **Bitmap** object. |
| [GdipBitmapGetHistogramSize](#gdipbitmapgethistogramsize) | Returns the number of elements (in an array of **DWORD**s) that you must allocate before you call the **GdipBitmapGetHistogram** function. |
| [GdipBitmapGetPixel](#gdipbitmapgetpixel) | Gets the color of a specified pixel in this bitmap. |
| [GdipBitmapLockBits](#gdipbitmaplockbits) | Locks a rectangular portion of this bitmap and provides a temporary buffer that you can use to read or write pixel data in a specified format. |
| [GdipBitmapSetPixel](#gdipbitmapsetpixel) | Sets the color of a specified pixel in this bitmap. |
| [GdipBitmapSetResolution](#gdipbitmapsetresolution) | Sets the resolution of this **Bitmap** object. |
| [GdipBitmapUnlockBits](#gdipbitmapunlockbits) | Unlocks a portion of this bitmap that was previously locked by a call to **GdipBitmapLockBits**. |
| [GdipCloneBitmapArea](#gdipclonebitmaparea) | Creates a new **Bitmap** object by copying a portion of this bitmap. |
| [GdipCloneBitmapAreaI](#gdipclonebitmapareaI) | Creates a new **Bitmap** object by copying a portion of this bitmap. |
| [GdipCreateBitmapFromDirectDrawSurface](#gdipcreatebitmapfromdirectdrawsurface) | Creates a **Bitmap** object based on a **DirectDraw** surface. |
| [GdipCreateBitmapFromFile](#gdipcreatebitmapfromfile) | Creates a **Bitmap** object based on an image file. |
| [GdipCreateBitmapFromFileICM](#gdipcreatebitmapfromfileicm) | Creates a **Bitmap** object based on an image file. This function uses ICM. |
| [GdipCreateBitmapFromGdiDib](#gdipcreatebitmapfromgdidib) | Creates a **Bitmap** object based on a **BITMAPINFO** structure and an array of pixel data. |
| [GdipCreateBitmapFromGraphics](#gdipcreatebitmapfromgraphics) | Creates a **Bitmap** object based on a **Graphics** object, a width, and a height. |
| [GdipCreateBitmapFromHBITMAP](#gdipcreatebitmapfromHbitmap) | Creates a **Bitmap** object based on a handle to a Windows Microsoft® Windows® Graphics Device Interface (GDI) bitmap and a handle to a GDI palette. |
| [GdipCreateBitmapFromHICON](#gdipcreatebitmapfromhicon) | Creates a **Bitmap** object based on an icon. |
| [GdipCreateBitmapFromResource](#gdipcreatebitmapfromresource) | Creates a **Bitmap** object based on an application or dynamic-link library (DLL) instance handle and the name of a bitmap resource. |
| [GdipCreateBitmapFromScan0](#gdipcreatebitmapfromscan0) | Creates a **Bitmap** object based on an array of bytes along with size and format information. |
| [GdipCreateBitmapFromStream](#gdipcreatebitmapfromstream) | Creates a **Bitmap** object based on a stream. |
| [GdipCreateBitmapFromStreamICM](#gdipcreatebitmapfromstreamicm) | Creates a **Bitmap** object based on a stream. This function uses ICM. |
| [GdipCreateEffect](#gdipcreateeffect) | Creates an **Effect** object of the type specified by the guid parameter. |
| [GdipCreateHBITMAPFromBitmap](#gdipcreatehbitmapfrombitmap) | Creates a Microsoft Windows Graphics Device Interface (GDI) bitmap from a **Bitmap** object. |
| [GdipCreateHICONFromBitmap](#gdipcreatehiconfrombitmap) | Creates an icon from a **Bitmap** object. |
| [GdipDeleteEffect](#gdipdeleteeffect) | Deletes the specified **Effect** object. |
| [GdipGetEffectParameters](#gdipgeteffectparameters) | Returns the parameters for the specified **Effect** object. |
| [GdipGetEffectParameterSize](#gdipgeteffectparametersize) | Gets the total size, in bytes, of the parameters currently set for the specified effect. |
| [GdipInitializePalette](#gdipinitializepalette) | Initializes a standard, optimal, or custom color palette. |
| [GdipSetEffectParameters](#gdipseteffectparameters) | Sets the parameters for the specified **Effect** object. |
| [GdipTestControl](#gdiptestcontrol) | Used for internal testing of Microsoft® Windows® GDI+. |

---

# CachedBitmap Functions

A **CachedBitmap** object stores a bitmap in a format that is optimized for display on a particular device. To display a cached bitmap, call the **GdipDrawCachedBitmap** function.

**CachedBitmap** is optimized for repeated rendering on a specific display device. It does not support scaling, rotation, or DPI-aware rendering. Use only when performance is critical and the image remains static and untransformed.

| Name       | Description |
| ---------- | ----------- |
| [GdipCreateCachedBitmap](#gdipcreatecachedbitmap) | Creates a **CachedBitmap** object based on a **Bitmap** object and a **Graphics** object. |
| [GdipDeleteCachedBitmap](#gdipdeletecachedbitmap) | Deletes a **CachedBitmap** object. |
| [GdipDrawCachedBitmap](#gdipdrawcachedbitmap) | Draws the image stored in a **CachedBitmap** object. |

---

## GdipBitmapApplyEffect

Alters a **Bitmap** object by applying a specified effect.

```
FUNCTION GdipBitmapApplyEffect (BYVAL bmp AS GpBitmap PTR, BYVAL effect AS CGpEffect PTR, _
   BYVAL roi AS RECT PTR, BYVAL useAuxData AS BOOL, BYVAL auxData AS VOID PTR PTR, _
   BYVAL auxDataSize AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Address of a pointer to a **Bitmap** object to which the effect is applied. |
| *effect* | [in] The effect to be applied. |
| *roi* | in] Pointer to a **RECT** structure that specifies the portion of the input bitmap to which the effect is applied. Pass **NULL** to specify that the effect applies to the entire input bitmap. |
| *useAuxData* | [in] Flag that specifies whether the function should return a pointer to the auxiliary data that it creates. |
| *auxData* | [out] Pointer to a set of lookup tables. |
| *auxDataSize* | [out] Size, in bytes, of the auxiliary data. |

#### Example

```
' ========================================================================================
' This example loads an image from disk and applies a brightness and contrast effect using GDI+ 1.1.
' Adjusts the image’s brightness and contrast.
' ========================================================================================
SUB Example_BrightnessEffect (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create a brightness and contrast effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(BrightnessContrastEffectGuid, @effect)

   ' // Set parameters: Brightness = 50, contrast = 50
   DIM bcParams AS BrightnessContrastParams
   bcParams.brightnessLevel = 50
   bcParams.contrastLevel = 20
   hStatus = GdipSetEffectParameters(effect, @bcParams, SIZEOF(BrightnessContrastParams))

   ' // Apply effects to the whole image
   hStatus = GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

---

## GdipBitmapConvertFormat

Converts a bitmap to a specified pixel format. The original pixel data in the bitmap is replaced by the new pixel data.

```
FUNCTION GdipBitmapConvertFormat (BYVAL bmp AS GpBitmap PTR, BYVAL pixelformat AS PixelFormat, _
   BYVAL dithertype AS DitherType, BYVAL palettetype AS PaletteType, BYVAL palette AS ColorPalette PTR, _
   BYVAL alphaThresholdPercent AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object to be converted. |
| *pixelformat* | [in] Pixel format constant that specifies the new pixel format. |
| *dithertype* | [in] Element of the *DitherType* enumeration that specifies the dithering algorithm. In cases where the conversion does not reduce the bit depth of the pixel data, pass *DitherTypeNone*. |
| *palettetype* | [in] Element of the *PaletteType* enumeration that specifies a standard palette to be used for dithering. If you are converting to a non-indexed format, this parameter is ignored. In that case, pass any element of the *PaletteType* enumeration, say *PaletteTypeCustom*. |
| *palette* | [in] Pointer to a **ColorPalette** structure that specifies the palette whose indexes are stored in the pixel data of the converted bitmap. This palette (called the actual palette) does not have to have the type specified by the *palettetype* parameter. The *palettetype* parameter specifies a standard palette that can be used by any of the ordered or spiral dithering algorithms. If the actual palette has a type other than that specified by the *palettetype* parameter, then the **GdipBitmapCpnvert** function performs a nearest-color conversion from the standard palette to the actual palette. |
| *alphaThresholdPercent* | [in] Real number in the range 0 through 100 that specifies which pixels in the source bitmap will map to the transparent color in the converted bitmap. A value of 0 specifies that none of the source pixels map to the transparent color. A value of 100 specifies that any pixel that is not fully opaque will map to the transparent color. A value of t specifies that any source pixel less than t percent of fully opaque will map to the transparent color. Note that for the alpha threshold to be effective, the palette must have a transparent color. If the palette does not have a transparent color, pixels with alpha values below the threshold will map to color that most closely matches (0, 0, 0, 0), usually black.

#### Example

```
' ========================================================================================
' GdipBitmapConvertFormat is a GDI+ 1.1 function that converts the pixel format of an existing
' bitmap to another format—such as from 32bpp ARGB to 8bpp indexed. It replaces the original
' pixel data with new data that matches the target format.
' ========================================================================================
SUB Example_ConvertBitmapFormat (BYVAL hdc AS HDC)

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

   ' Create original bitmap from buffer
   DIM nWidth AS LONG = 100
   DIM nHeight AS LONG = 100
   DIM bytesPerPixel AS LONG = 4
   DIM stride AS LONG = nWidth * bytesPerPixel
   DIM buffer(stride * nHeight - 1) AS UBYTE

   FOR y AS LONG = 0 TO nHeight - 1
      FOR x AS LONG = 0 TO nWidth - 1
         DIM offset_ AS LONGINT = y * stride + x * bytesPerPixel
         buffer(offset_ + 0) = x * 255 \ nWidth    ' Blue
         buffer(offset_ + 1) = y * 255 \ nHeight   ' Green
         buffer(offset_ + 2) = 128                 ' Red
         buffer(offset_ + 3) = 255                 ' Alpha
      NEXT
   NEXT

   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromScan0(nWidth, nHeight, stride, PixelFormat32bppARGB, @buffer(0), @bmp)

   ' Convert to 8bpp indexed format
   hStatus = GdipBitmapConvertFormat(bmp, PixelFormat8bppIndexed, DitherTypeNone, PaletteTypeFixedHalftone216, NULL, 0)

   ' Draw converted bitmap
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipBitmapCreateApplyEffect

Creates a new **Bitmap** object by applying a specified effect to an existing **Bitmap** object.

```
FUNCTION GdipBitmapCreateApplyEffect (BYVAL inputBitmaps AS GpBitmap PTR PTR, BYVAL numInputs AS INT_, _
   BYVAL effect AS CGpEffect PTR, BYVAL roi AS RECT PTR, BYVAL outputRect AS RECT P=TR, _
   BYVAL outputBitmap AS GpBitmap PTR PTR, BYVAL useAuxData AS BOOL, BYVAL auxData AS VOID PTR PTR, _
   BYVAL auxDataSize AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *inputBitmaps* | [in] Address of a pointer to a **Bitmap** object to which the effect is applied. |
| *numInputs* | [in] Integer that specifies the number of input bitmaps. This parameter must be set to 1. |
| *effect* | [in] The effect to be applied. |
| *roi* | [in] Pointer to a **RECT** structure that specifies the portion of the input bitmap that is used. |
| *outputRect* | [out] Pointer to a **RECT** structure that receives the portion of the input bitmap that was used. If the rectangle specified by ROI lies entirely within the input bitmap, the rectangle returned in outputRect is the same as roi. If part of rectangle specified by roi lies outside the input bitmap, then the rectangle returned in outputRect is the portion of roi that lies within the input bitmap. Pass **NULL** if you do not want to receive the output rectangle. |
| *outputBitmap* | [out] Address of a variable that receives a pointer to the new **Bitmap** object. |
| *useAuxData* | [in] Flag that specifies whether the function should return a pointer to the auxiliary data that it creates. |
| *auxData* | [out] Pointer to a set of loo kup tables. |
| *auxDataSize* | [out] Size, in bytes, of the auxiliary data. |

Example:

```
' ========================================================================================
' Creates a new Bitmap object by applying a specified effect to an existing Bitmap object.
' Demonstrates GdipBitmapCreateApplyEffect to generate a new image from an existing one.
' ========================================================================================
SUB Example_CreateBitmapWithApplyEffect (BYVAL hdc AS HDC)

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

   ' // Load the original bitmap
   DIM srcBitmap AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @srcBitmap)
   hStatus = GdipBitmapSetResolution(srcBitmap, dpiX, dpiY)

   ' // Create the tint effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(TintEffectGuid, @effect)

   ' // Set parameters: Hue = 120 (green), Amount = 60
   DIM tintParams AS TintParams
   tintParams.hue = 120
   tintParams.amount = 60
   hStatus = GdipSetEffectParameters(effect, @tintParams, SIZEOF(TintParams))

   ' // Prepare input array
   DIM inputBitmaps(0 TO 0) AS GpBitmap PTR
   inputBitmaps(0) = srcBitmap

   ' // Create output bitmap
   DIM outputBitmap AS GpBitmap PTR
   hStatus = GdipBitmapCreateApplyEffect(@inputBitmaps(0), 1, effect, NULL, NULL, @outputBitmap, FALSE, NULL, NULL)

   ' // Draw the new bitmap
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, outputBitmap), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF srcBitmap THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, srcBitmap))
   IF outputBitmap THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, outputBitmap))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipBitmapGetHistogram

Returns one or more histograms for specified color channels of the specified **Bitmap** object.

```
FUNCTION GdipBitmapGetHistogram (BYVAL bmp AS GpBitmap PTR, BYVAL fmt AS HistogramFormat, _
   BYVAL NumberOfEntries AS UINT, BYVAL channel0 AS UINT PTR, BYVAL channel1 AS UINT PTR, _
   BYVAL channel2 AS UINT PTR, BYVAL channel3 AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object. |
| *fmt* | [in] Element of the **HistogramFormat** enumeration that specifies the channels for which histograms will be created. |
| *NumberOfEntries* | [in] Integer that specifies the number of elements (of type UINT) in each of the arrays pointed to by *channel0*, *channel1*, *channel2*, and *channel3*. You must allocate memory for those arrays before you call **GdipBitmapGetHistogram**. To determine the required number of elements, call **GdipBitmapGetHistogramSize**. |
| *channel0* | [out] Pointer to an array of UINTs that receives the first histogram. |
| *channel1* | [out] Pointer to an array of UINTs that receives the second histogram if there is a second histogram. Pass **NULL** if there is no second histogram. |
| *channel2* | [out] Pointer to an array of UINTs that receives the third histogram if there is a third histogram. Pass **NULL** if there is no third histogram. |
| *channel3* | [out] Pointer to an array of UINTs that receives the fourth histogram if there is a fourth histogram. Pass **NULL** if there is no fourth histogram. |

#### Example

```
' ========================================================================================
' This example loads an image from disk and gets the histogram.
' ========================================================================================
SUB Example_GetHistogram (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' Get histogram size
   DIM entries AS UINT
   GdipBitmapGetHistogramSize(HistogramFormatARGB, @entries)

   ' Allocate histogram arrays
   DIM red(entries - 1) AS UINT
   DIM green(entries - 1) AS UINT
   DIM blue(entries - 1) AS UINT
   DIM alpha(entries - 1) AS UINT

   ' Extract histograms
   GdipBitmapGetHistogram(bmp, HistogramFormatARGB, entries, @red(0), @green(0), @blue(0), @alpha(0))

   ' Example: print peak red value
   DIM maxRed AS UINT = 0
   FOR i AS UINT = 0 TO entries - 1
      IF red(i) > maxRed THEN maxRed = red(i)
   NEXT
   AfxMsg ("Max red intensity count: " & WSTR(maxRed))

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipBitmapGetHistogramSize

Returns the number of elements (in an array of UINTs) that you must allocate before you call the **GdipBitmapGetHistogram** function.

```
FUNCTION GdipBitmapGetHistogramSize (BYVAL fmt AS HistogramFormat, BYVAL NumberOfEntries AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *fmt* | [in] Element of the *HistogramFormat* enumeration that specifies the pixel format of the bitmap. |
| *NumberOfEntries* | [out] Pointer to a UINT variable that receives the number of entries. |

#### Example

```
' ========================================================================================
' This example loads an image from disk and gets the histogram size.
' ========================================================================================
SUB Example_GetHistogramSize (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' Get histogram size
   DIM entries AS UINT
   GdipBitmapGetHistogramSize(HistogramFormatARGB, @entries)
   AfxMsg ("Histogram size - entries: " & WSTR(entries))

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipBitmapGetPixel

Gets the color of a specified pixel in this bitmap.

```
FUNCTION GdipBitmapGetPixel (BYVAL bmp AS GpBitmap PTR, BYVAL x AS INT_, BYVAL y aS INT_, _
   BYVAL clr AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object. |
| *x* | [in] Long value that specifies the x-coordinate (column) of the pixel. |
| *y* | [in] Long value that specifies the y-coordinate (row) of the pixel. |
| *clr* | [out] Pointer to a ARGB variable that receives the color of the specified pixel. |

#### Remarks

Depending on the format of the bitmap, **GdipBitmapGetPixel** might not return the same value as was set by **GdipBitmapSetPixel**. For example, if you call **GdipBitmapSetPixel** on a **Bitmap** object whose pixel format is 32bppPARGB, the pixel's RGB components are premultiplied. A subsequent call to **GdipBitmapGetPixel** might return a different value because of rounding. Also, if you call **GdipBitmapSetPixel** on a **Bitmap** object whose color depth is 16 bits per pixel, information could be lost during the conversion from 32 to 16 bits, and a subsequent call to **GdipBitmapGetPixel** might return a different value.

## Example

```
' ========================================================================================
' This example demonstrates how to create a GpBitmap object from a manually filled pixel
' buffer using GdipCreateBitmapFromScan0. It generates a 100×100 pixel image with a vertical
' gradient and renders it to the screen using GDI+.
' The example is DPI-aware thanks to GdipScaleWorldTransform, which scales the graphics
' context according to the user's display DPI. This ensures consistent rendering across
' different screen resolutions.
' ========================================================================================
SUB Example_CreateBitmapFromScan0 (BYVAL hdc AS HDC)

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

   ' // Define dimensions and pixel format
   DIM nWidth AS LONG = 100
   DIM nHeight AS LONG = 100
   DIM bytesPerPixel AS LONG = 4
   DIM stride AS LONG = nWidth * bytesPerPixel

   ' // Allocate pixel buffer
   DIM buffer(stride * nHeight - 1) AS UBYTE
   ' // Fill buffer with a gradient
   FOR y AS LONG = 0 TO nHeight - 1
      FOR x AS LONG = 0 TO nWidth - 1
         DIM _offset AS LONGINT = y * stride + x * bytesPerPixel
         buffer(_offset + 0) = x * 255 \ nWidth    ' Blue
         buffer(_offset + 1) = y * 255 \ nHeight   ' Green
         buffer(_offset + 2) = 128                 ' Red
         buffer(_offset + 3) = 255                 ' Alpha
      NEXT
   NEXT

   ' // Create bitmap from buffer
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromScan0(nWidth, nHeight, stride, PixelFormat32bppARGB, @buffer(0), @bmp)

   ' // Draw the bitmap
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipBitmapLockBits

Locks a rectangular portion of this bitmap and provides a temporary buffer that you can use to read or write pixel data in a specified format. Any pixel data that you write to the buffer is copied to the **Bitmap** object when you call **GdipBitmapUnlockBits**.

```
FUNCTION GdipBitmapLockBits (BYVAL bmp AS GpBitmap PTR, BYVAL prc AS GpRect PTR, _
   BYVAL flags AS UINT, BYVAL fmt AS PixelFormat, BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object. |
| *prc* | [in] Pointer to a rectangle that specifies the portion of the bitmap to be locked. |
| *flags* | [in] Set of flags that specify whether the locked portion of the bitmap is available for reading or for writing and whether the caller has already allocated a buffer. Individual flags are defined in the *ImageLockMode* enumeration. |
| *fmt* | [in] Long integer value that specifies the format of the pixel data in the temporary buffer. The pixel format of the temporary buffer does not have to be the same as the pixel format of this **Bitmap** object. Microsoft® Windows® GDI+ version 1.0 does not support processing of 16-bits-per-channel images, so you should not set this parameter equal to *PixelFormat48bppRGB*, *PixelFormat64bppARGB*, or *PixelFormat64bppPARGB*. |
| *lockedBitmapData* | [in, out] Pointer to a **BitmapData** object. If the *ImageLockModeUserInputBuf* flag of the flags parameter is cleared, then *lockedBitmapData* serves only as an output parameter. In that case, the *Scan0* data member of the **BitmapData** object receives a pointer to a temporary buffer, which is filled with the values of the requested pixels. The other data members of the **BitmapData** object receive attributes (width, height, format, and stride) of the pixel data in the temporary buffer. If the pixel data is stored bottom-up, the *Stride* data member is negative. If the pixel data is stored top-down, the *Stride* data member is positive. If the *ImageLockModeUserInputBuf* flag of the flags parameter is set, then *lockedBitmapData* serves as an input parameter (and possibly as an output parameter). In that case, the caller must allocate a buffer for the pixel data that will be read or written. The caller also must create a **BitmapData** object, set the *Scan0* data member of that **BitmapData** object to the address of the buffer, and set the other data members of the **BitmapData** object to specify the attributes (width, height, format, and stride) of the buffer. |

#### Remarks

**GdipBitmapLockBits** and **GdipBitmapUnlockBits** must be used as a pair. A call to **GdipBitmapLockBits** establishes a temporary buffer that you can use to read or write pixel data in a specified format. After you write to the temporary buffer, a call to **GdipBitmapUnlockBits** copies the pixel data in the buffer to the **Bitmap** object. If the pixel format of the temporary buffer is different from the pixel format of the **Bitmap** object, the pixel data is converted appropriately.

#### Example

```
' ========================================================================================
' Demonstrates the use of GdipBitmapLockBits/GdipBitmapUnlockBits.
' This example inverts the colors of a Bitmap.
' ========================================================================================
SUB Example_LockBits (BYVAL hdc AS HDC)

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

   ' // Load bitmap
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Get image dimensions
   DIM nWidth AS UINT, nHeight AS UINT
   hStatus = GdipGetImageWidth(bmp, @nWidth)
   hStatus = GdipGetImageHeight(bmp, @nHeight)

   ' // Define lock rectangle
   DIM rect AS GpRect = (0, 0, nWidth, nHeight)

   ' // Lock bitmap bits
   DIM bmpData AS BitmapData
   hStatus = GdipBitmapLockBits(bmp, @rect, ImageLockModeRead OR ImageLockModeWrite, PixelFormat32bppARGB, @bmpData)

   ' // Invert colors
   DIM pPixels AS UBYTE PTR = bmpData.Scan0
   FOR y AS LONG = 0 TO nHeight - 1
      FOR x AS LONG = 0 TO nWidth - 1
         DIM offset AS LONG = y * bmpData.Stride + x * 4
         pPixels[offset + 0] = 255 - pPixels[offset + 0] ' Blue
         pPixels[offset + 1] = 255 - pPixels[offset + 1] ' Green
         pPixels[offset + 2] = 255 - pPixels[offset + 2] ' Red
         ' Alpha remains unchanged
      NEXT
   NEXT

   ' // Unlock bitmap bits
   hStatus = GdipBitmapUnlockBits(bmp, @bmpData)

   ' // Draw the modified bitmap
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipBitmapSetPixel

Sets the color of a specified pixel in this bitmap.

```
FUNCTION GdipBitmapSetPixel (BYVAL bmp AS GpBitmap PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
   BYVAL clr AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object. |
| *x* | [in] Long integer value that specifies the x-coordinate (column) of the pixel. |
| *y* | [in] Long integer value that specifies the y-coordinate (row) of the pixel. |
| *clr* | [in] ARGB color value that specifies the color to set. |

#### Remarks

Depending on the format of the bitmap, **GdipBitmapGetPixel** might not return the same value as was set by **GdipBitmapSetPixel**. For example, if you call **GdipBitmapSetPixel** on a **Bitmap** object whose pixel format is 32bppPARGB, the pixel's RGB components are premultiplied. A subsequent call to **GdipBitmapGetPixel** might return a different value because of rounding. Also, if you call **GdipBitmapSetPixel** on a **Bitmap** object whose color depth is 16 bits per pixel, information could be lost during the conversion from 32 to 16 bits, and a subsequent call to **GdipBitmapGetPixel** might return a different value.

#### Example

```
' ========================================================================================
' The following example creates a Bitmap object based on a JPEG file. The code draws the
' bitmap once unaltered. Then the code calls the SetPixel method to create a
' checkered pattern of black pixels in the bitmap and draws the altered bitmap.
' ========================================================================================
SUB Example_SetPixel (BYVAL hdc AS HDC)

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

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @myBitmap)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(myBitmap, dpiX, dpiY)

   ' // Draw the bitmap.
   hStatus = GdipDrawImage(graphics, cast(GpImage PTR, myBitmap), 0, 0)

   ' // Get the width and height of the bitmap
   DIM AS UINT nWidth, nHeight
   hStatus = GdipGetImageWidth(myBitmap, @nWidth)
   hStatus = GdipGetImageHeight(myBitmap, @nHeight)

   ' // Make an ARGB color
   DIM pixelColor AS ARGB = ARGB_Black 'GDIP_ARGB(255, 0, 0, 0)

   ' // Create a checkered pattern with black pixels.
   DIM AS LONG row, col
   FOR row = 0 TO nWidth - 1 STEP 2
      FOR col = 0 TO nHeight STEP 2
         hStatus = GdipBitmapSetPixel(myBitmap, row, col, pixelColor)
      NEXT
   NEXT

  ' // Draw the altered bitmap.
   hStatus = GdipDrawImage(graphics, myBitmap, 200, 0)

   ' // Cleanup
   IF myBitmap THEN hStatus = GdipDisposeImage(cast(GpImage PTR, myBitmap))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipBitmapSetResolution

Sets the resolution of this **Bitmap** object.

```
FUNCTION GdipBitmapSetResolution (BYVAL bmp AS GpBitmap PTR, BYVAL xdpi AS REAL, BYVAL ydpi AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object. |
| *xdpi* | [in] Single precision value that specifies the horizontal resolution in dots per inch. |
| *ydpi* | [in] Single precision value that specifies the vertical resolution in dots per inch. |

#### Example

```
' ========================================================================================
' This example loads a .jpg image from file and sets the resolution to the DPI used by the;
' computer. Otherwise, as Each image can use a different internal DPI, an image with a
' DPI of 72 will be rendered bigger and an image with a DPI of 300 will be rendered smaller.
' ========================================================================================
SUB Example_SetResolution (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Draw image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

---

## GdipBitmapUnlockBits

Unlocks a portion of this bitmap that was previously locked by a call to **GdipBitmapLockBits**.

```
FUNCTION GdipBitmapUnlockBits (BYVAL bmp AS GpBitmap PTR, BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object. |
| *lockedBitmapData* | [in] Pointer to a **BitmapData** structure that was previously passed to **GdipBitmapLockBits**. |

#### Remarks

**GdipBitmapLockBits** and **GdipBitmapUnlockBits** must be used as a pair. A call to **GdipBitmapLockBits** establishes a temporary buffer that you can use to read or write pixel data in a specified format. After you write to the temporary buffer, a call to **GdipBitmapUnlockBits** copies the pixel data in the buffer to the **Bitmap** object. If the pixel format of the temporary buffer is different from the pixel format of the **Bitmap** object, the pixel data is converted appropriately.

#### Example

```
' ========================================================================================
' Demonstrates the use of GdipBitmapLockBits/GdipBitmapUnlockBits.
' This example inverts the colors of a Bitmap.
' ========================================================================================
SUB Example_LockBits (BYVAL hdc AS HDC)

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

   ' // Load bitmap
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Get image dimensions
   DIM nWidth AS UINT, nHeight AS UINT
   hStatus = GdipGetImageWidth(bmp, @nWidth)
   hStatus = GdipGetImageHeight(bmp, @nHeight)

   ' // Define lock rectangle
   DIM rect AS GpRect = (0, 0, nWidth, nHeight)

   ' // Lock bitmap bits
   DIM bmpData AS BitmapData
   hStatus = GdipBitmapLockBits(bmp, @rect, ImageLockModeRead OR ImageLockModeWrite, PixelFormat32bppARGB, @bmpData)

   ' // Invert colors
   DIM pPixels AS UBYTE PTR = bmpData.Scan0
   FOR y AS LONG = 0 TO nHeight - 1
      FOR x AS LONG = 0 TO nWidth - 1
         DIM offset AS LONG = y * bmpData.Stride + x * 4
         pPixels[offset + 0] = 255 - pPixels[offset + 0] ' Blue
         pPixels[offset + 1] = 255 - pPixels[offset + 1] ' Green
         pPixels[offset + 2] = 255 - pPixels[offset + 2] ' Red
         ' Alpha remains unchanged
      NEXT
   NEXT

   ' // Unlock bitmap bits
   hStatus = GdipBitmapUnlockBits(bmp, @bmpData)

   ' // Draw the modified bitmap
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCloneBitmapArea

Creates a new **Bitmap** object by copying a portion of this bitmap.

```
FUNCTION GdipCloneBitmapArea (BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, _
   BYVAL height AS REAL, BYVAL fmt AS PixelFormat, BYVAL srcBitmap AS GpBitmap PTR, _
   BYVAL dstBitmap AS GpBitmap PTR PTR) AS GpStatus
```
```
FUNCTION GdipCloneBitmapAreaI (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, _
   BYVAL height AS INT_, BYVAL fmt AS PixelFormat, BYVAL srcBitmap AS GpBitmap PTR, _
   BYVAL dstBitmap AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object. |
| *x* | [in] A value that specifies the x-coordinate of the upper-left corner of the rectangle that specifies the portion of this bitmap to copy. |
| *y* | [in] A value that specifies the y-coordinate of the upper-left corner of the rectangle that specifies the portion of this bitmap to copy. |
| *width* | [in] A value that specifies the width of the rectangle that specifies the portion of this bitmap to copy. |
| *height* | [in] A value that specifies the height of the rectangle that specifies the portion of this image to copy. |
| *fmt* | [in] Long integer value that specifies the pixel format of the new bitmap. |
| *srcBitmap* | [in] Pointer to the source **Bitmap** object. |
| *destBitmap* | [out] Pointer to a **Bitmap** object that receives a pointer to the cloned **Bitmap** object. |

#### Example

```
' ========================================================================================
' The following example creates a Bitmap from an image file, clones the upper-left
' portion of the image, and then draws the cloned image.
' ========================================================================================
SUB Example_CloneArea (BYVAL hdc AS HDC)

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

   DIM myBitmap AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @myBitmap)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(myBitmap, dpiX, dpiY)

   ' // Clone a portion of the bitmap.
   DIM cloneBitmap AS GpBitmap PTR
   hstatus = GdipCloneBitmapArea(0, 0, 100, 100, PixelFormatDontCare, myBitmap, @cloneBitmap)

   ' // Draw the clone.
   hStatus = GdipDrawImage(graphics, cast(GpImage PTR, cloneBitmap), 0, 0)

   ' // Cleanup
   IF myBitmap THEN hStatus = GdipDisposeImage(cast(GpImage PTR, myBitmap))
   IF cloneBitmap THEN hStatus = GdipDisposeImage(cast(GpImage PTR, cloneBitmap))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateBitmapFromDirectDrawSurface

Creates a **Bitmap** object based on a **DirectDraw** surface. The **Bitmap** object maintains a reference to the **DirectDraw** surface until the **Bitmap** object is deleted or goes out of scope.

```
FUNCTION GdipCreateBitmapFromDirectDrawSurface (BYVAL surface AS IDirectDrawSurface7 PTR, _
   BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *surface* | [in] Pointer to an **IDrectDrawSurface7** interface. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

---

## GdipCreateBitmapFromFile

Creates a **Bitmap** object based on an image file.

```
FUNCTION GdipCreateBitmapFromFile (BYVAL filename AS CONST WSTRING PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *filename* | [in] Pointer to a null-terminated unicode string that specifies the path name of the image file. The graphics file formats supported by GDI+ are BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

#### Example

```
' ========================================================================================
' The following example creates a Bitmap object based on a JPEG file and draws it.
' ========================================================================================
SUB Example_CreateBitmap (BYVAL hdc AS HDC)

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

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @myBitmap)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(myBitmap, dpiX, dpiY)

   ' // Draw the bitmap.
   hStatus = GdipDrawImage(graphics, cast(GpImage PTR, myBitmap), 0, 0)

   ' // Cleanup
   IF myBitmap THEN hStatus = GdipDisposeImage(cast(GpImage PTR, myBitmap))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateBitmapFromFileICM

Creates a **Bitmap** object based on an image file. This function uses ICM.

```
FUNCTION GdipCreateBitmapFromFileICM (BYVAL filename AS CONST WSTRING PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *filename* | [in] Pointer to a null-terminated unicode string that specifies the path name of the image file. The graphics file formats supported by GDI+ are BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

#### Example

```
' ========================================================================================
' The following example creates a Bitmap object based on a JPEG file and draws it.
' This excample uses ICM.
' ========================================================================================
SUB Example_CreateBitmapFromFileICM (BYVAL hdc AS HDC)

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

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFileICM("climber.jpg", @myBitmap)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(myBitmap, dpiX, dpiY)

   ' // Draw the bitmap.
   hStatus = GdipDrawImage(graphics, cast(GpImage PTR, myBitmap), 0, 0)

   ' // Cleanup
   IF myBitmap THEN hStatus = GdipDisposeImage(cast(GpImage PTR, myBitmap))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateBitmapFromGdiDib

Creates a **Bitmap** object based on a **BITMAPINFO** structure and an array of pixel data.

```
FUNCTION GdipCreateBitmapFromGdiDib (BYVAL gdiBitmapInfo AS BITMAPINFO PTR, _
   BYVAL gdiBitmapData AS VOID PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *gdiBitmapInfo* | [in] Pointer to a Microsoft® Windows® Graphics Device Interface (GDI)**BITMAPINFO** structure. |
| *gdiBitmapData* | [in] Pointer to an array of bytes that contains the pixel data. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

#### Example

This example creates a GpBitmap from a manually constructed BITMAPINFO and pixel buffer.

```
' ========================================================================================
SUB Example_CreateBitmapFromGdiDib (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create graphics object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get DPI scaling
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM ryRatio AS SINGLE = dpiY / 96

   ' // Apply DPI scaling
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Define bitmap dimensions
   DIM nWidth AS LONG = 100
   DIM nHeight AS LONG = 100
   DIM bytesPerPixel AS LONG = 4
   DIM stride AS LONG = nWidth * bytesPerPixel

   ' // Allocate pixel buffer
   DIM buffer(stride * nHeight - 1) AS UBYTE

   ' // Fill buffer with gradient
   FOR y AS LONG = 0 TO nHeight - 1
      FOR x AS LONG = 0 TO nWidth - 1
         DIM offset AS LONG = y * stride + x * bytesPerPixel
         buffer(offset + 0) = x * 255 \ nWidth   ' Blue
         buffer(offset + 1) = y * 255 \ nHeight  ' Green
         buffer(offset + 2) = 128                ' Red
         buffer(offset + 3) = 255                ' Alpha
      NEXT
   NEXT

   ' // Define BITMAPINFO header
   DIM bmi AS BITMAPINFO
   WITH bmi.bmiHeader
      .biSize = SIZEOF(BITMAPINFOHEADER)
      .biWidth = nWidth
      .biHeight = -nHeight ' Top-down DIB
      .biPlanes = 1
      .biBitCount = 32
      .biCompression = BI_RGB
      .biSizeImage = stride * nHeight
      .biXPelsPerMeter = dpiX * 39.3701
      .biYPelsPerMeter = dpiY * 39.3701
   END WITH

   ' // Create bitmap from DIB
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromGdiDib(@bmi, @buffer(0), @bmp)

   ' // Draw bitmap
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateBitmapFromGraphics

Creates a **Bitmap** object based on a **Graphics** object, a width, and a height.

```
FUNCTION GdipCreateBitmapFromGraphics (BYVAL width AS INT_, BYVAL height AS INT_, _
   BYVAL target AS GpGraphics PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *width* | [in] Long integer value that specifies the width, in pixels, of the bitmap. |
| *height* | [in] Long integer value that specifies the height, in pixels, of the bitmap. |
| *target* | [in] Pointer to a **Graphics** object that contains information used to initialize certain properties (for example, dots per inch) of the new **Bitmap** object. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

#### Example

This example creates a GpBitmap from a graphics object, inheriting DPI and rendering settings.

```
' ========================================================================================
SUB Example_CreateBitmapFromGraphics (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create graphics object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get DPI scaling
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Apply DPI scaling
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Define bitmap dimensions
   DIM nWidth AS LONG = 200
   DIM nHeight AS LONG = 100

   ' // Create bitmap from graphics context
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromGraphics(nWidth, nHeight, graphics, @bmp)

   ' // Create graphics object from bitmap
   DIM bmpGraphics AS GpGraphics PTR
   hStatus = GdipGetImageGraphicsContext(CAST(GpImage PTR, bmp), @bmpGraphics)

   ' // Fill bitmap with color
   hStatus = GdipGraphicsClear(bmpGraphics, ARGB_Red)

   ' // Draw bitmap to screen
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmpGraphics THEN hStatus = GdipDeleteGraphics(bmpGraphics)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateBitmapFromHBITMAP

Creates a **Bitmap** object based on a handle to a Windows Microsoft® Windows® Graphics Device Interface (GDI) bitmap and a handle to a GDI palette.

```
FUNCTION GdipCreateBitmapFromHBITMAP (BYVAL hbm AS HBITMAP, BYVAL hpal AS HPALETTE, _
   BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hbm* | [in] Handle to a GDI bitmap. |
| *hpal* | [in] Handle to a GDI palette used to define the bitmap colors if hbm is not a device-independent bitmap (DIB). |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

#### Remarks

You are responsible for deleting the GDI bitmap and the GDI palette.

Do not pass to the **GdipCreateBitmapFromHBITMAP** function a GDI bitmap or a GDI palette that is currently (or was previously) selected into a device context.

#### Example

This example creates a GpBitmap from a legacy HBITMAP and renders it using GDI+.

```
' ========================================================================================
SUB Example_CreateBitmapFromHBITMAP (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create graphics object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get DPI scaling
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM ryRatio AS SINGLE = dpiY / 96

   ' // Apply DPI scaling
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a legacy HBITMAP using GDI
   DIM hbm AS HBITMAP
   DIM memDC AS HDC = CreateCompatibleDC(hdc)
   DIM bmpInfo AS BITMAPINFOHEADER
   bmpInfo.biSize = SIZEOF(BITMAPINFOHEADER)
   bmpInfo.biWidth = 100
   bmpInfo.biHeight = -100 ' Top-down
   bmpInfo.biPlanes = 1
   bmpInfo.biBitCount = 32
   bmpInfo.biCompression = BI_RGB
   bmpInfo.biSizeImage = 0

   DIM pBits AS ANY PTR
   hbm = CreateDIBSection(memDC, CAST(BITMAPINFO PTR, @bmpInfo), DIB_RGB_COLORS, @pBits, NULL, 0)

   ' // Fill with red pixels
   FOR i AS LONG = 0 TO 100 * 100 - 1
      CAST(ULONG PTR, pBits)[i] = &HFFFF0000 ' ARGB Red
   NEXT

   ' // Create GDI+ Bitmap from HBITMAP
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromHBITMAP(hbm, NULL, @bmp)

   ' // Set resolution
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Draw the bitmap
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)
   IF hbm THEN DeleteObject(hbm)
   IF memDC THEN DeleteDC(memDC)

END SUB
' ========================================================================================
```

#### Helper function

```
' ========================================================================================
' Prints a Windows bitmap in the default printer.
' Parameters:
' - hbmp = Handle to the bitmap.
' - bStretch = Stretch the image.
' - nStretchMode = Stretching mode. Default value = InterpolationModeHighQualityBicubic.
'   InterpolationModeLowQuality = 1
'   InterpolationModeHighQuality = 2
'   InterpolationModeBilinear = 3
'   InterpolationModeBicubic = 4
'   InterpolationModeNearestNeighbor = 5
'   InterpolationModeHighQualityBilinear = 6
'   InterpolationModeHighQualityBicubic = 7
' Return value: Returns TRUE if the bitmap has been printed successfully, or FALSE otherwise.
' ========================================================================================
PRIVATE FUNCTION AfxGdipPrintHBITMAP (BYVAL hbmp AS HBITMAP, BYVAL bStretch AS BOOLEAN = FALSE, BYVAL nStretchMode AS LONG = InterpolationModeHighQualityBicubic) AS BOOLEAN

   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT FUNCTION
   DIM hPrinter AS HANDLE
   DIM pbufferDoc AS UBYTE PTR
   DIM pBitmap AS GpBitmap PTR
   DIM pGraphics AS GpGraphics PTR

   ' // Note: Using a DO... LOOP, with an EXIT DO before LOOP, to avoid the use of GOTO.
   DO
      ' // Get the name of the default printer
      DIM wszPrinterName AS WSTRING * MAX_PATH
      GetProfileStringW("WINDOWS", "DEVICE", "", wszPrinterName, SIZEOF(wszPrinterName))
      DIM nPos AS LONG = INSTR(wszPrinterName, ",")
      IF nPos THEN wszPrinterName = LEFT(wszPrinterName, nPos - 1)
      IF wszPrinterName = "" THEN EXIT DO
      ' // Open the printer
      IF OpenPrinterW(wszPrinterName, @hPrinter, NULL) = 0 THEN EXIT DO
      ' // Allocate a buffer of the correct size
      DIM dwNeeded AS DWORD = DocumentPropertiesW(NULL, hPrinter, wszPrinterName, NULL, NULL, 0)
      pbufferDoc = CAllocate(1, dwNeeded)
      ' // Retrieve the printer configuration data
      DIM nRet AS LONG = DocumentPropertiesW(NULL, hPrinter, wszPrinterName, CAST(PDEVMODEW, pbufferDoc), NULL, DM_OUT_BUFFER)
      IF nRet <> IDOK THEN EXIT DO
      ' // Cast it to a DEVMODEW structure
      DIM pDevMode AS DEVMODEW PTR = CAST(DEVMODEW PTR, pbufferDoc)
      ' // Create a device context for the printer
      DIM hDC AS .HDC = CreateDCW(wszPrinterName, wszPrinterName, NULL, pDevMode)
      IF hDC = NULL THEN EXIT DO
      ' // Create a graphics object from the printer DC
      GdipCreateFromHDC(hDC, @pGraphics)
      IF pGraphics = NULL THEN EXIT DO
      ' // Create a Bitmap object from an HBITMAP
      GdipCreateBitmapFromHBITMAP(hbmp, NULL, @pBitmap)
      IF pBitmap = NULL THEN EXIT DO
      ' // Stretching
      DIM cx AS SINGLE, cy AS SINGLE
      IF bStretch THEN GdipSetInterpolationMode(pGraphics, nStretchMode)
      ' // Get the DPIs of the printer
      DIM dpiX AS SINGLE, dpiY AS SINGLE
      GdipGetDpiX(pGraphics, @dpiX)
      GdipGetDpiY(pGraphics, @dpiY)
      ' // Calculate the width and height according to the DPIs of the printer
      cx = GetDeviceCaps(hdc, HORZRES) / (dpiX / 100)
      cy = GetDeviceCaps(hdc, VERTRES) / (dpiY / 100)
      ' // Print the bitmap
      DIM di AS DOCINFOW
      DIM wszDocName AS WSTRING * 260 = "Printing bitmap"
      di.cbSize = SIZEOF(DOCINFOW)
      di.lpszDocName = @wszDocName
      DIM hr AS LONG = StartDocW(hDC, @di)
      IF hr <= 0 THEN EXIT DO
      DIM i AS LONG
      FOR i = 1 TO pDevMode->dmCopies
         IF StartPage(hDC) THEN
            ' // Draw the image
            IF bStretch THEN
               GdipDrawImageRect(pGraphics, CAST(GpImage PTR, pBitmap), 0, 0, cx, cy)
            ELSE
               GdipDrawImage(pGraphics, CAST(GpImage PTR, pBitmap), 0, 0)
            END IF
            EndPage(hDC)
         END IF
      NEXT
      EndDoc(hDC)
      FUNCTION = TRUE
      EXIT DO
   LOOP

   IF pBitmap THEN GdipDisposeImage(cast(GpImage PTR, pBitmap))
   IF pGraphics THEN GdipDeleteGraphics(pGraphics)
   ' // Finished with the printer
   IF hPrinter THEN ClosePrinter(hPrinter)
   ' // Cleanup
   IF pbufferDoc THEN Delete(pbufferDoc)
   ' // Shutdown Gdiplus
   GdiplusShutdown token

END FUNCTION
' ========================================================================================
```
---

## GdipCreateBitmapFromHICON

Creates a **Bitmap** object based on an icon.

```
FUNCTION GdipCreateBitmapFromHICON (BYVAL hicon AS HICON, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hIcon* | [in] Handle to a Microsoft® Windows® Graphics Device Interface (GDI) icon. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

#### Example

```
' ========================================================================================
SUB Example_CreateBitmapFromHICON  (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create graphics object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get DPI scaling
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Apply DPI scaling
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Load a system icon (e.g., information icon)
   DIM hIcon AS HICON = LoadIcon(NULL, IDI_INFORMATION)

   ' // Create a GDI+ Bitmap from the HICON
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromHICON(hIcon, @bmp)

   ' // Set the bitmap resolution to match the display DPI
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Draw the bitmap
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateBitmapFromResource

Creates a **Bitmap** object based on an application or dynamic-link library (DLL) instance handle and the name of a bitmap resource.

```
FUNCTION GdipCreateBitmapFromResource (BYVAL hInst AS HINSTANCE, BYVAL lpBitmapName AS WSTRING PTR, _
   BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInst* | [in] Handle to an instance of a module whose executable file contains a bitmap resource. |
| *lpBitmapName* | [in] Pointer to a null-terminated unicode string that specifies the path name of the bitmap resource to be loaded. Alternatively, this parameter can consist of the resource identifier in the low-order word and zero in the high-order word. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

#### Helper functions

```
' ========================================================================================
' Loads an image from a resource, converts it to an icon or bitmap and returns the handle.
' Parameters:
' - hInstance     = [in] A handle to the module whose portable executable file or an accompanying
'                   MUI file contains the resource. If this parameter is NULL, the function searches
'                   the module used to create the current process.
' - wszImageName  = [in] Name of the image in the resource file (.RES). If the image resource uses
'                   an integral identifier, wszImage should begin with a number symbol (#)
'                   followed by the identifier in an ASCII format, e.g., "#998". Otherwise,
'                   use the text identifier name for the image. Only images embedded as raw data
'                   (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff.
' - dimPercent    = Percent of dimming (1-99)
' - bGrayScale    = TRUE or FALSE. Convert to gray scale.
' - imageType     = IMAGE_ICON or IMAGE_BITMAP.
' - clrBackground = [in] The background color. This parameter is ignored if the image type
'                   is IMAGE_ICON or the bitmap is totally opaque.
' Return Value:
'   If the function succeeds, the return value is the handle of the created icon or bitmap.
'   If the function fails, the return value is NULL.
' ========================================================================================
PRIVATE FUNCTION AfxGdipImageFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL imageType AS LONG = IMAGE_ICON, BYVAL clrBackground AS ARGB = 0) AS HANDLE

   DIM pImage AS GpImage PTR, hImage AS HANDLE
   DIM hRes AS HRSRC, pResData AS HRSRC, wID AS WORD, dwID AS DWORD, imageSize AS DWORD
   DIM pImageStream AS IStream PTR, hGlobal AS HGLOBAL, pGlobalBuffer AS LPVOID
   DIM ImageWidth AS LONG, ImageHeight AS LONG, x AS LONG, y AS LONG
   DIM pixColor AS GDIP_BGRA, iColor AS LONG, rFactor AS SINGLE

   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT FUNCTION

   ' // Find the resource
   IF LEFT(wszImageName, 1) = "#" THEN
      wID = VAL(MID(wszImageName, 2))
      dwID = MAKELONG(wID, 0)
      hRes = FindResourceW(hInstance, MAKEINTRESOURCEW(dwID), CAST(LPCWSTR, RT_RCDATA))
   ELSE
      hRes = FindResourceW(hInstance, wszImageName, CAST(LPCWSTR, RT_RCDATA))
   END IF
   IF hRes THEN
      ' // Retrieve the size of the image
      imageSize = SizeofResource(hInstance, hRes)
      IF imageSize THEN
         ' // Load the resource and get a pointer to the resource data.
         ' // Note: LockResource does not actually lock memory; it is just used
         ' // to obtain a pointer to the memory containing the resource data.
         pResData = LockResource(LoadResource(hInstance, hRes))
         IF pResData THEN
            ' // Allocate memory to hold the image
            hGlobal = GlobalAlloc(GMEM_MOVEABLE, imageSize)
            IF hGlobal THEN
               ' // Lock the memory
               pGlobalBuffer = GlobalLock(hGlobal)
               IF pGlobalBuffer THEN
                  ' // Copy the image from the binary string file to global memory
                  CopyMemory(pGlobalBuffer, pResData, imageSize)
                  ' // Create an stream in global memory
                  IF CreateStreamOnHGlobal(hGlobal, FALSE, @pImageStream) = S_OK THEN
                     IF pImageStream THEN
                        ' // Create a bitmap from the data contained in the stream
                        GdipCreateBitmapFromStream(pImageStream, CAST(GpBitmap PTR PTR, @pImage))
                        ' // Get the image width and height
                        GdipGetImageWidth(pImage, @ImageWidth)
                        GdipGetImageHeight(pImage, @ImageHeight)
                        ' // Dim or/and gray the image
                        IF dimPercent > 0 AND dimPercent < 100 THEN rFactor = dimPercent / 100
                        IF rFactor <> 0 OR bGrayScale <> 0 THEN
                           FOR y = 0 TO ImageHeight - 1
                              FOR x = 0 TO ImageWidth - 1
                                 ' // Get the pixel color
                                 GdipBitmapGetPixel(CAST(GpBitmap PTR, pImage), x, y, @pixColor.color)
                                 IF dimPercent > 0 THEN
                                    pixColor.red   = (255 - pixColor.red) * rFactor + pixColor.red
                                    pixColor.green = (255 - pixColor.green) * rFactor + pixColor.green
                                    pixColor.blue  = (255 - pixColor.blue) * rFactor + pixColor.blue
                                 END IF
                                 IF bGrayScale THEN
                                    ' Note: The sum of the percentages for the three colors should add tp up 1
                                    iColor = 0.299 * pixColor.red + 0.587 * pixColor.green + 0.114 * pixColor.blue
                                    pixColor.Color = GDIP_BGRA (iColor, iColor, iColor, pixColor.alpha)
                                 ELSE
                                    pixColor.color = GDIP_ARGB(pixColor.alpha, pixColor.red, pixColor.green, pixColor.Blue)
                                 END IF
                                 ' // Set the modified pixel color
                                 GdipBitmapSetPixel(CAST(GpBitmap PTR, pImage), x, y, pixColor.color)
                              NEXT
                           NEXT
                        END IF
                        ' // Create icon from image
                        IF imageType = IMAGE_ICON THEN
                           GdipCreateHICONFromBitmap(CAST(GpBitmap PTR, pImage), @hImage)
                        ELSE
                           GdipCreateHBITMAPFromBitmap(CAST(GpBitmap PTR, pImage), @hImage, clrBackground)
                        END IF
                        ' // Free the image
                        IF pImage THEN GdipDisposeImage pImage
                        pImageStream->lpVtbl->Release(pImageStream)
                     END IF
                  END IF
                  ' // Unlock the memory
                  GlobalUnlock pGlobalBuffer
               END IF
               ' // Free the memory
               GlobalFree hGlobal
            END IF
         END IF
      END IF
   END IF

   ' // Shutdown Gdiplus
   GdiplusShutdown token
   ' // Return the handle of the icon
   FUNCTION = hImage

END FUNCTION
' ========================================================================================
' ========================================================================================
' Loads an image from a resource file, converts it to an icon and returns the handle.
' ========================================================================================
PRIVATE FUNCTION AfxGdipIconFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS HANDLE
   FUNCTION = AfxGdipImageFromRes(hInstance, wszImageName, dimPercent, bGrayScale, IMAGE_ICON, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
' Loads an image from a resource file, converts it to a bitmap and returns the handle.
' ========================================================================================
PRIVATE FUNCTION AfxGdipBitmapFromRes (BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBackground AS ARGB = 0) AS HANDLE
   FUNCTION = AfxGdipImageFromRes(hInstance, wszImageName, dimPercent, bGrayScale, IMAGE_BITMAP, clrBackground)
END FUNCTION
' ========================================================================================
```
---

## GdipCreateBitmapFromScan0

Creates a **Bitmap** object based on an array of bytes along with size and format information.
```
FUNCTION GdipCreateBitmapFromScan0 (BYVAL width AS INT_, BYVAL height AS INT_, BYVAL stride AS INT_, _
   BYV AL fmt AS PixelFormat, BYV AL scan0 AS BYTE PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *width* | [in] Long integer value that specifies the width, in pixels, of the bitmap. |
| *height* | [in] Long integer value that specifies the height, in pixels, of the bitmap. |
| *stride* | [in] Long integer value that specifies the byte offset between the beginning of one scan line and the next. This is usually (but not necessarily) the number of bytes in the pixel format (for example, 2 for 16 bits per pixel) multiplied by the width of the bitmap. The value passed to this parameter must be a multiple of four. |
| *fmt* | [in] Long integer value that specifies the pixel format of the bitmap. |
| *scan0* | [in] Pointer to an array of bytes that contains the pixel data. The caller is responsible for allocating and freeing the block of memory pointed to by this parameter. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

Example:

```
' ========================================================================================
' This example demonstrates how to create a GpBitmap object from a manually filled pixel
' buffer using GdipCreateBitmapFromScan0. It generates a 100×100 pixel image with a vertical
' gradient and renders it to the screen using GDI+.
' The example is DPI-aware thanks to GdipScaleWorldTransform, which scales the graphics
' context according to the user's display DPI. This ensures consistent rendering across
' different screen resolutions.
' ========================================================================================
SUB Example_CreateBitmapFromScan0 (BYVAL hdc AS HDC)

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

   ' // Define dimensions and pixel format
   DIM nWidth AS LONG = 100
   DIM nHeight AS LONG = 100
   DIM bytesPerPixel AS LONG = 4
   DIM stride AS LONG = nWidth * bytesPerPixel

   ' // Allocate pixel buffer
   DIM buffer(stride * nHeight - 1) AS UBYTE
   ' // Fill buffer with a gradient
   FOR y AS LONG = 0 TO nHeight - 1
      FOR x AS LONG = 0 TO nWidth - 1
         DIM _offset AS LONGINT = y * stride + x * bytesPerPixel
         buffer(_offset + 0) = x * 255 \ nWidth    ' Blue
         buffer(_offset + 1) = y * 255 \ nHeight   ' Green
         buffer(_offset + 2) = 128                 ' Red
         buffer(_offset + 3) = 255                 ' Alpha
      NEXT
   NEXT

   ' // Create bitmap from buffer
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromScan0(nWidth, nHeight, stride, PixelFormat32bppARGB, @buffer(0), @bmp)

   ' // Draw the bitmap
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateBitmapFromStream

Creates a **Bitmap** object based on a stream.

```
FUNCTION GdipCreateBitmapFromStream (BYVAL stream AS IStream PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *stream* | [in] Pointer to an **IStream** interface. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

#### Helper function

```
' ========================================================================================
' Converts an image stored in a buffer into an icon or bitmap and returns the handle.
' Parameters:
' - pBuffer       = [in] Pointer to the buffer
' - bufferSize    = Size of the buffer
' - dimPercent    = Percent of dimming (1-99)
' - bGrayScale    = TRUE or FALSE. Convert to gray scale.
' - imageType     = IMAGE_ICON or IMAGE_BITMAP.
' - clrBackground = [in] The background color. This parameter is ignored if the image type
'                   is IMAGE_ICON or the bitmap is totally opaque.
' Return Value:
'   If the function succeeds, the return value is the handle of the created icon or bitmap.
'   If the function fails, the return value is NULL.
' Usage example:
'   DIM wszFileName AS WSTRING * MAX_PATH
'   wszFileName = ExePath & "\arrow_left_256.png"
'   DIM bufferSize AS SIZE_T_
'   DIM nFile AS LONG
'   nFile = FREEFILE
'   OPEN wszFileName FOR BINARY AS nFile
'   IF ERR THEN EXIT FUNCTION
'   bufferSize = LOF(nFile)
'   DIM pBuffer AS UBYTE PTR
'   pBuffer = CAllocate(1, bufferSize)
'   GET #nFile, , *pBuffer, bufferSize
'   CLOSE nFile
'   IF pBuffer THEN
'      ImageList_ReplaceIcon(hImageList, -1, AfxGdipIconFromBuffer(pBuffer, ImageSize))
'      DeAllocate(pBuffer)
'   END IF
' ========================================================================================
PRIVATE FUNCTION AfxGdipImageFromBuffer (BYVAL pBuffer AS ANY PTR, BYVAL bufferSize AS SIZE_T_, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, _
   BYVAL imageType AS LONG = IMAGE_ICON, BYVAL clrBackground AS ARGB = 0) AS HANDLE

   DIM pImage AS GpImage PTR, hImage AS HANDLE
   DIM pImageStream AS IStream PTR, hGlobal AS HGLOBAL, pGlobalBuffer AS LPVOID
   DIM ImageWidth AS LONG, ImageHeight AS LONG, x AS LONG, y AS LONG
   DIM pixColor AS GDIP_BGRA, iColor AS LONG, rFactor AS SINGLE

   ' // Initialize Gdiplus
   DIM token AS ULONG_PTR = AfxGdipInit
   IF token = NULL THEN EXIT FUNCTION
   ' // Allocate memory to hold the image
   hGlobal = GlobalAlloc(GMEM_MOVEABLE, bufferSize)
   IF hGlobal THEN
      ' // Lock the memory
      pGlobalBuffer = GlobalLock(hGlobal)
      IF pGlobalBuffer THEN
         ' // Copy the image from the binary string file to global memory
         CopyMemory(pGlobalBuffer, pBuffer, bufferSize)
         ' // Create an stream in global memory
         IF CreateStreamOnHGlobal(hGlobal, FALSE, @pImageStream) = S_OK THEN
            IF pImageStream THEN
               ' // Create a bitmap from the data contained in the stream
               GdipCreateBitmapFromStream(pImageStream, CAST(GpBitmap PTR PTR, @pImage))
               ' // Get the image width and height
               GdipGetImageWidth(pImage, @ImageWidth)
               GdipGetImageHeight(pImage, @ImageHeight)
               ' // Dim or/and gray the image
               IF dimPercent > 0 AND dimPercent < 100 THEN rFactor = dimPercent / 100
               IF rFactor <> 0 OR bGrayScale <> 0 THEN
                  FOR y = 0 TO ImageHeight - 1
                     FOR x = 0 TO ImageWidth - 1
                        ' // Get the pixel color
                        GdipBitmapGetPixel(CAST(GpBitmap PTR, pImage), x, y, @pixColor.color)
                        IF dimPercent > 0 THEN
                           pixColor.red   = (255 - pixColor.red) * rFactor + pixColor.red
                           pixColor.green = (255 - pixColor.green) * rFactor + pixColor.green
                           pixColor.blue  = (255 - pixColor.blue) * rFactor + pixColor.blue
                        END IF
                        IF bGrayScale THEN
                           ' Note: The sum of the percentages for the three colors should add to up 1
                           iColor = 0.299 * pixColor.red + 0.587 * pixColor.green + 0.114 * pixColor.blue
                           pixColor.Color = GDIP_BGRA (iColor, iColor, iColor, pixColor.alpha)
                        ELSE
                           pixColor.color = GDIP_ARGB(pixColor.alpha, pixColor.red, pixColor.green, pixColor.Blue)
                        END IF
                        ' // Set the modified pixel color
                        GdipBitmapSetPixel(CAST(GpBitmap PTR, pImage), x, y, pixColor.color)
                     NEXT
                  NEXT
               END IF
               ' // Create icon from image
               IF imageType = IMAGE_ICON THEN
                  GdipCreateHICONFromBitmap(CAST(GpBitmap PTR, pImage), @hImage)
               ELSE
                  GdipCreateHBITMAPFromBitmap(CAST(GpBitmap PTR, pImage), @hImage, clrBackground)
               END IF
               ' // Free the image
               IF pImage THEN GdipDisposeImage pImage
               pImageStream->lpVtbl->Release(pImageStream)
            END IF
         END IF
         ' // Unlock the memory
         GlobalUnlock pGlobalBuffer
      END IF
      ' // Free the memory
      GlobalFree hGlobal
   END IF

   ' // Shutdown Gdiplus
   GdiplusShutdown token
   ' // Return the handle of the icon
   FUNCTION = hImage

END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGdipIconFromBuffer (BYVAL pBuffer AS ANY PTR, BYVAL bufferSize AS SIZE_T_, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE) AS HANDLE
   FUNCTION = AfxGdipImageFromBuffer(pBuffer, bufferSize, dimPercent, bGrayScale, IMAGE_ICON, 0)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION AfxGdipBitmapFromBuffer (BYVAL pBuffer AS ANY PTR, BYVAL bufferSize AS SIZE_T_, _
   BYVAL dimPercent AS LONG = 0, BYVAL bGrayScale AS LONG = FALSE, BYVAL clrBackground AS ARGB = 0) AS HANDLE
   FUNCTION = AfxGdipImageFromBuffer(pBuffer, bufferSize, dimPercent, bGrayScale, IMAGE_BITMAP, clrBackground)
END FUNCTION
' ========================================================================================
```
---

## GdipCreateBitmapFromStreamICM

Creates a **Bitmap** object based on a stream. This function uses ICM.

```
FUNCTION GdipCreateBitmapFromStreamICM (BYVAL stream AS IStream PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *stream* | [in] Pointer to an **IStream** interface. |
| *bmp* | [out] Pointer to a **Bitmap** object that receives a pointer to a **Bitmap** object. |

---

## GdipCreateEffect

Creates an **Effect** object of the type specified by the guid parameter.

```
FUNCTION GdipCreateEffect (BYVAL gd AS GUID, BYVAL effect AS CGpEffect PTR PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *gd* | [in] Guid that specifies the kind of effect to create. |
| *effect* | [out] Pointer to a **Effect** object that receives a pointer to a **Effect** object. |

#### Example

```
' ========================================================================================
' This example loads an image from disk and applies a blur effect using GDI+ 1.1.
' ========================================================================================
SUB Example_BlurEffect (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create a blur effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(BlurEffectGuid, @effect)

   ' // Set parameters: radius = 6.0, expandEdge = FALSE
   DIM blurParams(0 TO 1) AS SINGLE = {6.0, FALSE}
   DIM array_size AS UNIT = (UBOUND(blurParams) - LBOUND(blurParams) + 1) * SIZEOF(SINGLE)
   hStatus = GdipSetEffectParameters(effect, @blurParams(0), array_size)

   ' // Apply effects to the whole image
   hStatus = GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
This example converts the image to grayscale by averaging RGB channels.

```
SUB Example_GrayScale (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create grayscale matrix
   DIM grayMatrix(0 TO 4, 0 TO 4) AS SINGLE = {{0.3, 0.3, 0.3, 0.0, 0.0}, _
      {0.59, 0.59, 0.59, 0.0, 0.0}, {0.11, 0.11, 0.11, 0.0, 0.0}, _
      {0.0, 0.0, 0.0, 1.0, 0.0}, {0.0, 0.0, 0.0, 0.0, 1.0}}
   DIM array_size AS UINT = (UBOUND(grayMatrix, 1) - LBOUND(grayMatrix, 1) + 1) * _
                            (UBOUND(grayMatrix, 2) - LBOUND(grayMatrix, 2) + 1) * SIZEOF(SINGLE)

   ' // Create effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(ColorMatrixEffectGuid, @effect)
   ' // Set parameters
   hStatus = GdipSetEffectParameters(effect, @grayMatrix(0, 0), array_size)
   ' // Apply effect to the whole image
   GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)
   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
#### Example

The HueSaturationLightness effecty enables you to change the hue, saturation, and lightness of a bitmap. To specify the magnitudes of the changes in hue, saturation, and lightness, pass a HueSaturationLightnessParams structure to the GdipSetEffectParameters function.

```
' ========================================================================================
SUB Example_GdipBitmapHueSaturationLightnessEffect (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create ColorLUT effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(HueSaturationLightnessEffectGuid, @effect)

   ' // Set the parameters
   DIM hslParams AS HueSaturationLightnessParams
   hslParams.hueLevel = 45            ' Rotate hue slightly
   hslParams.saturationLevel = 30     ' Boost saturation
   hslParams.lightnessLevel = -10     ' Slightly darken
   hStatus = GdipSetEffectParameters(effect, @hslParams, SIZEOF(hslParams))

   ' // Apply effect
   hStatus = GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)

   ' // Draw image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

#### Example

The RedEyeCorrection effect enables you to correct the red eyes that sometimes occur in flash photographs. To specify areas of the bitmap that have red eyes, pass a RedEyeCorrectionParams structure to the GdipSetEffectParameters function.

```
' ========================================================================================
SUB Example_Gdip_RedEyeCorrectionEffect (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("RedEyes.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create a levels effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(RedEyeCorrectionEffectGuid, @effect)

   ' // We need to specify one or more rectangles that enclose the red-eye areas.
   ' // These are passed as an array of RECT structures.

   ' // Define two rectangles around the eyes
   ' // Change the values according the coordinates of your image
   DIM eyeRects(0 TO 1) AS RECT
   eyeRects(0).left   = 1 : eyeRects(0).top    = 1
   eyeRects(0).right  = 400 : eyeRects(0).bottom = 250
   eyeRects(1).left   = 1 : eyeRects(1).top    = 1
   eyeRects(1).right  = 400 : eyeRects(1).bottom = 250

   ' // Fill a RedEyeCorrectionParams structure
   DIM redeyeParams AS RedEyeCorrectionParams
   redeyeParams.numberOfAreas = 2
   redeyeParams.areas = @eyeRects(0)
   ' // Calculate the size of the paameters
   DIM paramsSize AS UINT = SIZEOF(RedEyeCorrectionParams) + redeyeParams.numberOfAreas * SIZEOF(RECT)
   ' // Set the parameters
   hStatus = GdipSetEffectParameters(effect, @redeyeParams, paramsSize)

   ' // Apply effect
   hStatus = GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

#### Example

The Tint effect enables you to apply a tint to a bitmap. To specify the nature of the tint, pass the address of a TintParams structure to the GdipSetEffectParameters function.

```
' ========================================================================================
' The Tint effect enables you to apply a tint to a bitmap. To specify the nature of the tint,
' pass the address of a TintParams structure to the GdipSetEffectParameters function.
' ========================================================================================
SUB Example_TintEffect (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create a tint effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(TintEffectGuid, @effect)

   ' // Set parameters: Hue = +120 (green), Amount = 60 (moderate tint)
   DIM tintPrms AS TintParams
   tintPrms.hue = 120
   tintPrms.amount = 60
   hStatus = GdipSetEffectParameters(effect, @tintPrms, SIZEOF(TintParams))

   ' // Apply effect to the whole image
   hStatus = GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateHBITMAPFromBitmap

Creates a Microsoft Windows Graphics Device Interface (GDI) bitmap from a Bitmap object.

```
FUNCTION GdipCreateHBITMAPFromBitmap (BYVAL bmp AS GpBitmap PTR, BYVAL hbmReturn AS HBITMAP PTR, _
   BYVAL background AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object. |
| *hbmReturn* | [out] Pointer to an **HBITMAP** that receives a handle to the GDI bitmap. |
| *background* | [in] ARGB color that specifies the background color. This parameter is ignored if the bitmap is totally opaque. |

#### Example

This example converts a GDI+ Bitmap into a legacy HBITMAP and draws it using GDI.

```
' ========================================================================================
SUB Example_CreateHBITMAPFromBitmap  (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create graphics object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get DPI scaling
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Apply DPI scaling
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a GDI+ Bitmap from file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Convert to HBITMAP
   DIM hbm AS HBITMAP
   hStatus = GdipCreateHBITMAPFromBitmap(bmp, @hbm, ARGB_White)

   ' // Draw using classic GDI
   DIM memDC AS HDC = CreateCompatibleDC(hdc)
   DIM oldBmp AS HBITMAP = SelectObject(memDC, hbm)
   ' // Retrieve bitmap dimensions from HBITMAP
   DIM bmpInfo AS BITMAP
   GetObject(hbm, SIZEOF(BITMAP), @bmpInfo)
   ' // Apply DPI scaling
   DIM scaledWidth AS LONG = bmpInfo.bmWidth * rxRatio
   DIM scaledHeight AS LONG = bmpInfo.bmHeight * ryRatio
   StretchBlt(hdc, 0, 0, scaledWidth, scaledHeight, memDC, 0, 0, bmpInfo.bmWidth, bmpInfo.bmHeight, SRCCOPY)

   ' // Cleanup
   IF oldBmp THEN SelectObject(memDC, oldBmp)
   IF memDC THEN DeleteDC(memDC)
   IF hbm THEN DeleteObject(hbm)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipCreateHICONFromBitmap

Creates an icon from a **Bitmap** object.

```
FUNCTION GdipCreateHICONFromBitmap (BYVAL bmp AS GpBitmap PTR, BYVAL hbmReturn AS HICON PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to the **Bitmap** object. |
| *hbmReturn* | [out] Pointer to an **HICON** that receives a handle to the GDI icon. |

#### Example

This example converts a GDI+ Bitmap into a Windows HICON and displays it.

```
SUB Example_CreateHICONFromBitmap (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create graphics object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get DPI scaling
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Apply DPI scaling
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Create a GDI+ Bitmap from file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Convert to HICON
   DIM hIcon AS HICON
   hStatus = GdipCreateHICONFromBitmap(bmp, @hIcon)

   ' // Get width and heig ht of the icon
   DIM iconInfo AS ICONINFO
   DIM iconWidth AS LONG
   DIM iconHeight AS LONG
   IF GetIconInfo(hIcon, @iconInfo) THEN
      DIM bmp AS BITMAP
      IF iconInfo.hbmColor THEN
         ' Icon has a color bitmap
         GetObject(iconInfo.hbmColor, SIZEOF(BITMAP), @bmp)
         iconWidth = bmp.bmWidth
         iconHeight = bmp.bmHeight
         DeleteObject(iconInfo.hbmColor)
      ELSEIF iconInfo.hbmMask THEN
         ' Monochrome icon: height is double, so divide by 2
         GetObject(iconInfo.hbmMask, SIZEOF(BITMAP), @bmp)
         iconWidth = bmp.bmWidth
         iconHeight = bmp.bmHeight \ 2
      END IF
      IF iconInfo.hbmMask THEN DeleteObject(iconInfo.hbmMask)
   END IF

   ' // Scale it according to the DPI settings
   DIM scaledWidth AS LONG = iconWidth * rxRatio
   DIM scaledHeight AS LONG = iconHeight * ryRatio

   ' // Draw the icon
   DrawIconEx(hdc, 0, 0, hIcon, scaledWidth, scaledHeight, 0, NULL, DI_NORMAL)

   ' // Cleanup
   IF hIcon THEN DestroyIcon(hIcon)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeleteEffect

Deletes the specified **Effect** object.

```
FUNCTION GdipDeleteEffect (BYVAL effect AS CGpEffect PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *effect* | [in] Pointer to the **Effect** object created with **GdipCreateEffect**. |

#### Usage example

```
IF effect THEN GdipDeleteEffect(effect)
```
---

## GdipGetEffectParameters

Returns the parameters for the specified **Effect** object.

```
FUNCTION GdipGetEffectParameters (BYVAL effect AS CGpEffect PTR, BYVAL size AS UINT PTR, _
   BYVAL params AS VOID PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *effect* | [in] Pointer to the **Effect** object created with **GdipCreateEffect**. |
| *size* | [in/out] Pointer to a UINT variable containing the total size, in bytes, of the parameters currently set for this **Effect**. |
| *params* | [in] Variable that receives a pointer to the parameters. |

```
' ========================================================================================
SUB Example_GetEffectParameters (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create a tint effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(TintEffectGuid, @effect)

   ' // Set parameters: Hue = +120 (green), Amount = 60 (moderate tint)
   DIM tintPrms AS TintParams
   tintPrms.hue = 120
   tintPrms.amount = 60
   hStatus = GdipSetEffectParameters(effect, @tintPrms, SIZEOF(TintParams))

   ' // Retrieve parameters size
   DIM paramsSize AS UINT
   hStatus = GdipGetEffectParameterSize(effect, @paramsSize)
   AfxMsg WSTR("Parameters size: " & WSTR(paramsSize))
   ' // Allocate buffer and retrieve parameters
   DIM buffer AS TintParams
   hStatus = GdipGetEffectParameters(effect, @paramsSize, @buffer)
   ' // Display retrieved values
   AfxMsg("Tint hue: " & WSTR(buffer.hue) & ", amount: " & WSTR(buffer.amount))

   ' // Apply effect to the whole image
   hStatus = GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipGetEffectParameterSize

Gets the total size, in bytes, of the parameters currently set for the specified effect.

```
FUNCTION GdipGetEffectParameterSize (BYVAL effect AS CGpEffect PTR, BYVAL size AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *effect* | [in] Pointer to the **Effect** object created with **GdipCreateEffect**. |
| *size* | [out] Pointer to a UINT variable that receives the total size, in bytes, of the parameters. |

```
' ========================================================================================
SUB Example_GetEffectParameters (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create a tint effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(TintEffectGuid, @effect)

   ' // Set parameters: Hue = +120 (green), Amount = 60 (moderate tint)
   DIM tintPrms AS TintParams
   tintPrms.hue = 120
   tintPrms.amount = 60
   hStatus = GdipSetEffectParameters(effect, @tintPrms, SIZEOF(TintParams))

   ' // Retrieve parameters size
   DIM paramsSize AS UINT
   hStatus = GdipGetEffectParameterSize(effect, @paramsSize)
   AfxMsg WSTR("Parameters size: " & WSTR(paramsSize))
   ' // Allocate buffer and retrieve parameters
   DIM buffer AS TintParams
   hStatus = GdipGetEffectParameters(effect, @paramsSize, @buffer)
   ' // Display retrieved values
   AfxMsg("Tint hue: " & WSTR(buffer.hue) & ", amount: " & WSTR(buffer.amount))

   ' // Apply effect to the whole image
   hStatus = GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipInitializePalette

Initializes a standard, optimal, or custom color palette.

```
FUNCTION GdipInitializePalette (BYVAL palette AS ColorPalette PTR, BYVAL palettetype AS PaletteType, _
   BYVAL optimalColors AS INT_, BYVAL useTransparentColor AS BOOL, BYVAL bmp AS GpBitmap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *palette* | [out] Pointer to a buffer that contains a **ColorPalette** structure followed by an array of ARGB values. The **Entries** member of a **ColorPalette** structure is an array of one ARGB value. You must allocate memory for the **ColorPalette** structure and for the additional *ARGB* values in the palette. |
| *palettetype* | [in] Element of the *PaletteType* enumeration that specifies the palette type. The palette can have one of several standard types, or it can be a custom palette that you define. Also, the **GdipInitializePalette** method can create an optimal palette based on a specified bitmap. |
| *optimalColors* | [in] Integer that specifies the number of colors you want to have in an optimal palette based on a specified bitmap. If this parameter is greater than 0, the *palettetype* parameter must be set to *PaletteTypeOptimal*, and the bitmap parameter must point to a **Bitmap** object. If you are creating a standard or custom palette rather than an optimal palette, set this parameter to 0. |
| *useTransparentColor* | [in] Boolean value that specifies whether to include the transparent color in the palette. Set to TRUE to include the transparent color; otherwise FALSE. |
| *bmp* | [in] Pointer to a **Bitmap** object for which an optimal palette will be created. If *palettetype* is set to *PaletteTypeOptimal* and *optimalColors* is set to a positive integer, set this parameter to the address of a **Bitmap** object. Otherwise, set this parameter to **NULL**. |

#### Example

This example demonstrates how to initialize an optimal palette based on a bitmap.
Uses GdipInitializePalette to fill a ColorPalette structure with the best-fit colors.
It fails in my computer with hStatus = 1 (Generic error)

```
SUB Example_InitializeOptimalPalette (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create graphics object
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Get DPI scaling
   DIM dpiX AS SINGLE
   hStatus = GdipGetDpiX(graphics, @dpiX)
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   hStatus = GdipGetDpiY(graphics, @dpiY)
   DIM ryRatio AS SINGLE = dpiY / 96
   ' // Apply DPI scaling
   hStatus = GdipScaleWorldTransform(graphics, rxRatio, ryRatio, MatrixOrderPrepend)

   ' // Define dimensions and pixel format
   DIM nWidth AS LONG = 100
   DIM nHeight AS LONG = 100
   DIM bytesPerPixel AS LONG = 4
   DIM stride AS LONG = nWidth * bytesPerPixel

   ' // Allocate pixel buffer
   DIM buffer(stride * nHeight - 1) AS UBYTE
   ' // Fill buffer with a gradient
   FOR y AS LONG = 0 TO nHeight - 1
      FOR x AS LONG = 0 TO nWidth - 1
         DIM _offset AS LONGINT = y * stride + x * bytesPerPixel
         buffer(_offset + 0) = x * 255 \ nWidth    ' Blue
         buffer(_offset + 1) = y * 255 \ nHeight   ' Green
         buffer(_offset + 2) = 128                 ' Red
         buffer(_offset + 3) = 255                 ' Alpha
      NEXT
   NEXT

   ' // Create bitmap from buffer
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromScan0(nWidth, nHeight, stride, PixelFormat32bppARGB, @buffer(0), @bmp)

   ' // Allocate ColorPalette buffer
   CONST maxColors AS LONG = 256
   DIM paletteSize AS LONG = SIZEOF(ColorPalette) + (maxColors - 1) * SIZEOF(ARGB)
   DIM paletteBuffer AS UBYTE PTR = CALLocate(paletteSize)
   DIM paletteptr AS ColorPalette PTR = CAST(ColorPalette PTR, paletteBuffer)

   ' // Initialize optimal palette
   hStatus = GdipInitializePalette(paletteptr, PaletteTypeOptimal, maxColors, FALSE, bmp)

   ' // Example: print first color
   IF paletteptr->Count > 0 THEN
      DIM clr AS ARGB = paletteptr->Entries(0)
      AfxMsg("First palette color: ARGB = " & HEX(clr))
   END IF

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)
   IF paletteBuffer THEN DEALLOCATE paletteBuffer

END SUB
' ========================================================================================
```

---

## GdipSetEffectParameters

Sets the parameters for the specified **Effect** object.

```
FUNCTION GdipSetEffectParameters (BYVAL effect AS Gpeffect PTR, BYVAL params AS VOID PTR, _
   BYVAL size AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *effect* | [in] Pointer to the **Effect** object created with **GdipCreateEffect**. |
| *params* | [in] The parameters. |
| *size* | [in] The total size, in bytes, of the parameters currently set for this **Effect**. |

#### Example

```
' ========================================================================================
' This example loads an image from disk and applies a sharpen effect using GDI+ 1.1.
' Enhances edges and fine details by increasing contrast around transitions.
' Useful for restoring clarity in slightly blurred images.
' ========================================================================================
SUB Example_BlurEffect (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create a sharpen effect
   DIM sharpenEffect AS CGpEffect PTR
   hStatus = GdipCreateEffect(SharpenEffectGuid, @sharpenEffect)

   ' // Set parameters: radius = 3.0, expandEdge = FALSE
   DIM sharpenParams(0 TO 1) AS SINGLE = {3.0, FALSE}
   DIM array_size AS UNIT = (UBOUND(sharpenParams) - LBOUND(sharpenParams) + 1) * SIZEOF(SINGLE)
   hStatus = GdipSetEffectParameters(sharpenEffect, @sharpenParams(0), array_size)

   ' // Apply effects to the whole image
   hStatus = GdipBitmapApplyEffect(bmp, sharpenEffect, NULL, FALSE, NULL, NULL)

   ' // Draw the image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF sharpenEffect THEN hStatus = GdipDeleteEffect(sharpenEffect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
#### Example

This example applies the color LUT effect to the whole image.

```
' ========================================================================================
' The ColorLUT effect lets you define lookup tables for each channel—Alpha, Red, Green,
' and Blue—with 256 entries each. For every pixel in the image:
' The channel value (0–255) is used as an index.
' The corresponding value in the LUT replaces the original.
' This means you can:
' Invert colors
' Apply posterization
' Simulate film-like grading
' Create custom stylizations
' ========================================================================================
SUB Example_ColorLUTEffect (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create ColorBalance effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(ColorLUTEffectGuid, @effect)

   ' // Create LUTs: identity (no change)
   DIM lut AS ColorLUTParams
   FOR i AS LONG = 0 TO 255
       lut.lutB(i) = i
       lut.lutG(i)   = i
       lut.lutR(i) = i
       lut.lutA(i)  = i
   NEXT

   ' // Example: invert red channel
   FOR i AS LONG = 0 TO 255
       lut.lutR(i) = 255 - i
   NEXT

   ' // Set parameters
   hStatus = GdipSetEffectParameters(effect, @lut, SIZEOF(lut))

   ' // Apply effect
   hStatus = GdipBitmapApplyEffect(bmp, effect, NULL, FALSE, NULL, NULL)

   ' // Draw image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```

#### Example

This example applies the effect to the specified region only.

```
' ========================================================================================
SUB Example_ColorLUTEffectRegion (BYVAL hdc AS HDC)

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

   ' // Create a bitmap object from a .jpg file
   DIM bmp AS GpBitmap PTR
   hStatus = GdipCreateBitmapFromFile("climber.jpg", @bmp)
   ' // Set the resolution of this Bitmap object to the user's DPI settings
   hStatus = GdipBitmapSetResolution(bmp, dpiX, dpiY)

   ' // Create ColorLUT effect
   DIM effect AS CGpEffect PTR
   hStatus = GdipCreateEffect(ColorLUTEffectGuid, @effect)

   ' // Create LUTs: identity (no change)
   DIM lut AS ColorLUTParams
   FOR i AS LONG = 0 TO 255
       lut.lutB(i) = 255 - i   ' Invert red
       lut.lutG(i)   = i       ' Identity
       lut.lutR(i) = i         ' Identity
       lut.lutA(i)  = i        ' Identity
   NEXT

   ' // Example: invert red channel
   FOR i AS LONG = 0 TO 255
       lut.lutR(i) = 255 - i
   NEXT

   ' // Set parameters
   hStatus = GdipSetEffectParameters(effect, @lut, SIZEOF(lut))

   ' // Define the region of interest (ROI)
   DIM roi AS RECT
   roi.left   = 20
   roi.top    = 20
   roi.right  = 150
   roi.bottom = 100

   ' // Apply effect to the specified region only
   hStatus = GdipBitmapApplyEffect(bmp, effect, @roi, FALSE, NULL, NULL)

   ' // Draw image
   hStatus = GdipDrawImage(graphics, CAST(GpImage PTR, bmp), 0, 0)

   ' // Cleanup
   IF effect THEN hStatus = GdipDeleteEffect(effect)
   IF bmp THEN hStatus = GdipDisposeImage(CAST(GpImage PTR, bmp))
   IF graphics THEN hStatus = GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipTestControl

```
FUNCTION GdipTestControl (BYVAL control AS GpTestControlEnum, BYVAL param AS VOID PTR) AS GpStatus
```

#### Remarks

Used for internal testing of Microsoft® Windows® GDI+. Specifies changes in image-processing behavior. The meaning of the param parameter varies depending on the value passed in the control parameter.

If the control parameter is set to *TestControlForceBilinear*, *param* points to a boolean value. If that value is TRUE and the current interpolation mode is anything other than *InterpolationModeNearestNeighbor*, then the *InerpolationModeBilinear* algorithm is used. If the value pointed to by param is FALSE or if the current interpolation mode is *InterpolationModeNearestNeighbor*, then the behavior of GDI+ is unchanged.

If the control parameter is set to *TestControlNoICM*, param points to a boolean value. If that value is TRUE, then ICM profiles are ignored when images are loaded into memory. If the value pointed to by param is FALSE, then the behavior of GDI+ is unchanged.

If the control parameter is set to *TestControlGetBuildNumber*, param is an output parameter that points to a 32-bit integer variable. That variable receives the GDI+ build number.

**GdipTestControl** is not thread-safe; it assumes that there is only one thread using GDI+. If you call **GdipTestControl** in a multithreaded environment, the results are unpredictable.

The *GpTestControlEnum* enumeration has three elements: TestControlForceBilinear = 0, TestControlNoICM = 1, and TestControlGetBuildNumber = 2.

#### Example

Retrieves the internal build number of GDI+.

```
DIM buildNumber AS LONG
GdipTestControl(TestControlGetBuildNumber, @buildNumber)
print buildNumber
```
---

## GdipCreateCachedBitmap

Creates a **CachedBitmap** object based on a **Bitmap** object and a **Graphics** object. The cached bitmap takes the pixel data from the **Bitmap** object and stores it in a format that is optimized for the display device associated with the **Graphics** object.

```
FUNCTION GdipCreateCachedBitmap (BYVAL bmp AS GpBitmap PTR, BYVAL graphics AS GpGraphics PTR, _
   BYVAL cachedBitmap AS GpCachedBitmap PTR PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *bmp* | [in] Pointer to a **Bitmap** object that contains the pixel data to be optimized. |
| *graphics* | [in] Pointer to a **Graphics** object that is associated with a display device for which the image will be optimized. |
| *cachedBitmap* | [out] Pointer to a **CachedBitmap** variable that receives a pointer to the new created **CachedBitmap** object. |

#### Remarks

You can display a cached bitmap by passing the address of a **CachedBitmap** object to the **GdipDrawCachedBitmap** function.

#### Example

```
' ========================================================================================
' This example loads a bitmap, creates a CachedBitmap from it, and draws it twice.
' ========================================================================================
SUB Example_CachedBitmap (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create a graphics object from the device context
   DIM graphics AS GpGraphics PTR
   hStatus = GdipCreateFromHDC(hdc, @graphics)

   ' // Load the original image from file
   DIM image AS GpImage PTR
   hStatus = GdipLoadImageFromFile("climber.jpg", @image)

   ' // Create CachedBitmap optimized for this graphics device
   DIM cachedBmp AS GpCachedBitmap PTR
   hStatus = GdipCreateCachedBitmap(image, graphics, @cachedBmp)

   ' // Draw the cached bitmap twice
   hStatus = GdipDrawCachedBitmap(graphics, cachedBmp, 50, 50)
   hStatus = GdipDrawCachedBitmap(graphics, cachedBmp, 300, 50)

   ' // Cleanup
   IF cachedBmp THEN GdipDeleteCachedBitmap(cachedBmp)
   IF image THEN GdipDisposeImage(image)
   IF graphics THEN GdipDeleteGraphics(graphics)

END SUB
' ========================================================================================
```
---

## GdipDeleteCachedBitmap

Deletes a **CachedBitmap** object.

```
FUNCTION GdipDeleteCachedBitmap (BYVAL cachedBitmap AS GpCachedBitmap PTR) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *cachedBitmap* | [in] Pointer to a **CachedBitmap** object. |

#### Example

See example for **GdipCreateCachedBitmap**.

---

## GdipDrawCachedBitmap

Draws the image stored in a **CachedBitmap** object.

```
FUNCTION GdipDrawCachedBitmap (BYVAL graphics AS GpGraphics PTR, BYVAL cachedBitmap AS GpCachedBitmap PTR, _
   BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
```
| Parameter  | Description |
| ---------- | ----------- |
| *graphics* | [in] Pointer to a **Graphics** object. |
| *cachedBitmap* | [in] Pointer to a **CachedBitmap** object that contains the image to be drawn. |
| *x* | [in] Long integer value that specifies the x-coordinate of the upper-left corner of the image. |
| *y* | [in] Long integer value that specifies the y-coordinate of the upper-left corner of the image. |

#### Remarks

A **CachedBitmap** object stores an image in a format that is optimized for a particular display screen. You cannot draw a cached bitmap to a printer or to a metafile.

Cached bitmaps will not work with any transformations other than translation.

When you construct a **CachedBitmap** object, you must pass the address of a **Graphics** object to the **GdipCreateCachedBitmap** function. If the screen associated with that Graphics object has its bit depth changed after the cached bitmap is constructed, then the **GdipDrawCachedBitmap** function will fail, and you should reconstruct the cached bitmap. Alternatively, you can hook the display change notification message and reconstruct the cached bitmap at that time.

#### Example

See example for **GdipCreateCachedBitmap**.

---
