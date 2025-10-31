# CGpImage Class

`CGpImage`is an abstract base class that provides functionality for the `CGpBitmap` and `CGpMetafile` descended classes.

`CGpImage` provides methods for loading and saving raster images (bitmaps) and vector images (metafiles). An **Image** object encapsulates a bitmap or a metafile and stores attributes that you can retrieve by calling various methods. You can construct **Image** objects from a variety of file types including BMP, ICON, GIF, JPEG, Exif, PNG, TIFF, WMF, and EMF.

While it can be used directly, it is not recommended for most scenarios because it lacks essential functionality like resolution control, pixel access, and image effects. Use `CGpBitmap` for regular image files (e.g., JPEG, PNG, BMP, TIFF) and `CGpMetafile` for vector-based images (e.g., EMF, WMF). 

**DPI Awareness**: When working with raster images, predictable rendering size depends on correctly setting the image’s resolution. The **SetResolution** method, available only in `CGpBitmap`, allows you to define the horizontal and vertical DPI (dots per inch) of the image. Without explicitly setting the resolution, GDI+ assumes a default of 96 DPI, which can lead to incorrect scaling on high-DPI displays. This means the image may appear smaller or larger than intended depending on its internal DPI metadata. To ensure consistent and accurate rendering across different screen resolutions and printing contexts, it is essential to use `CGpBitmap` instead  of `CGpImage` and call **SetResolution** or **SetResolutionForDpi** after loading the image. If the image is a metafile we don't need to call **SetResolution**. Metafiles (EMF, WMF) are vector-based, meaning they store drawing commands rather than pixel data. Their rendering is resolution-independent — they scale smoothly across devices and DPI settings without distortion. GDI+ uses logical units (like points or inches) to interpret metafile dimensions, so embedded DPI doesn’t affect their size or clarity.

**Inherits from**: CGpBase.<br>
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructorsimage) | Creates an Image object based on a file or stream. |
| [Clone](#cloneimage) | Copies the contents of the existing Image object into a new Image object. |
| [FindFirstItem](#findfirstitem) | Retrieves the description and the data size of the first metadata item in this Image object. |
| [FindNextItem](#findnextitem) | Retrieves the description and the data size of the next metadata item in this Image object. |
| [GetAllPropertyItems](#getallpropertyitems) | Gets all the property items (metadata) stored in this Image object. |
| [GetBounds](#getbounds) | Gets the bounding rectangle for this image. |
| [GetEncoderParameterList](#getencoderparameterlist) | Gets a list of the parameters supported by a specified image encoder. |
| [GetEncoderParameterListSize](#getencoderparameterlistsize) | Gets the size, in bytes, of the parameter list for a specified image encoder. |
| [GetFlags](#getflags) | Gets a set of flags that indicate certain attributes of this Image object. |
| [GetFrameCount](#getframecount) | Gets the number of frames in a specified dimension of this Image object. |
| [GetFrameDimensionsCount](#getframedimensionscount) | Gets the number of frame dimensions in this Image object. |
| [GetFrameDimensionsList](#getframedimensionslist) | Gets the identifiers for the frame dimensions of this Image object. |
| [GetHeight](#getheight) | Gets the image height, in pixels, of this image. |
| [GetHorizontalResolution](#GetHorizontalResolution) | Gets the horizontal resolution, in dots per inch, of this image. |
| [GetItemData](#getitemdata) | Gets one piece of metadata from this Image object. |
| [GetPalette](#getpalette) | Gets the **ColorPalette** of this Image object. |
| [GetPaletteSize](#getpalettesize) | Gets the size, in bytes, of the color palette of this Image object. |
| [GetPhysicalDimension](#getphysicaldimension) | Gets the width and height of this image. |
| [GetPixelFormat](#getpixelformat) | Gets the pixel format of this Image object. |
| [GetPropertyCount](#getpropertycount) | Gets the number of properties (pieces of metadata) stored in this Image object. |
| [GetPropertyIdList](#getpropertyidlist) | Gets a list of the property identifiers used in the metadata of this Image object. |
| [GetPropertyItem](#getpropertyitem) | Gets a specified property item (piece of metadata) from this Image object. |
| [GetPropertyItemSize](#getpropertyitemsize) | Gets the size, in bytes, of a specified property item of this Image object. |
| [GetPropertySize](#getpropertysize) | Gets the total size, in bytes, of all the property items stored in this Image object. |
| [GetRawFormat](#getrawformat) | Gets a globally unique identifier ( GUID) that identifies the format of this Image object. |
| [GetThumbnailImage](#getthumbnailimage) | Gets a thumbnail image from this Image object. |
| [GetType](#gettype) | Gets the type (bitmap or metafile) of this Image object. |
| [GetVerticalResolution](#getverticalresolution) | Gets the vertical resolution, in dots per inch, of this image. |
| [GetWidth](#getwidth) | Gets the width, in pixels, of this image. |
| [RemovePropertyItem](#removepropertyitem) | Removes a property item (piece of metadata) from this Image object. |
| [RotateFlip](#rotateflip) | Rotates and flips this image. |
| [Save](#save) | Saves this image to a file. |
| [SaveAdd](#saveadd) | Adds a frame to a file or stream specified in a previous call to the Save method. |
| [SelectActiveFrame](#selectactiveframe) | Selects the frame in this Image object specified by a dimension and an index. |
| [SetPalette](#setpalette) | Sets the color palette of this Image object. |
| [SetPropertyItem](#setpropertyitem) | Sets a property item (piece of metadata) for this Image object. |

---

# CGpBitmap Class

The **CGpBitmap** class inherits from the **CGpImage** class. It encapsulates a GDI+ bitmap, which consists of the pixel data for a graphics image and its attributes. A **Bitmap** is an object used to work with images defined by pixel data.

**Inherits from**: CGpImage.<br>
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructorsbitmap) | Creates a **Bitmap** object. |
| [ApplyEffect](#applyeffect) | Alters a **Bitmap** object by applying a specified effect. |
| [Clone](#clonebitmap) | Creates a new **Bitmap** object by copying a portion of this bitmap. |
| [ConvertFormat](#convertformat) | Converts a bitmap to a specified pixel format. |
| [GetHBITMAP](#gethbitmap) | Creates a Windows Graphics Device Interface (GDI) bitmap from this **Bitmap** object. |
| [GetHICON](#gethicon) | Creates an icon from this **Bitmap** object. |
| [GetHistogram](#gethistogram) | Returns one or more histograms for specified color channels of this **Bitmap** object. |
| [GetHistogramSize](#gethistogramsize) | Returns the number of elements (in an array of DWORDs) that you must allocate before you call the **GetHistogram** method of a **Bitmap** object. |
| [GetPixel](#getpixel) | Gets the color of a specified pixel in this **Bitmap** object. |
| [InitializePalette](#initializepalette) | Initializes a standard, optimal, or custom color palette. |
| [LockBits](#lockbits) | Locks a rectangular portion of this bitmap and provides a temporary buffer that you can use to read or write pixel data in a specified format. |
| [SetPixel](#setpixel) | Sets the color of a specified pixel in this **Bitmap** object. |
| [SetResolution](#setresolution) | Sets the resolution of this **Bitmap** object. |
| [SetResolutionForDpi](#setresolutionfordpi) | Sets the resolution of this **Bitmap** object. This function uses the current DPI settings. |
| [UnlockBits](#unlockbits) | Unlocks a portion of this bitmap that was previously locked by a call to **LockBits**. |

---

# CGpMetafile Class

Extends the **CGpImage** class. The **Metafile** object defines a graphic metafile. A metafile contains records that describe a sequence of graphics API calls. Metafiles can be recorded (constructed) and played back (displayed).

**Inherits from**: CGpImage.<br>
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructorsmetafile) | Creates a Windows GDI+ Metafile. |
| [EmfToWmfBits](#emftowmfbits) | Converts an enhanced-format metafile to a Windows Metafile Format (WMF) metafile and stores the converted records in a specified buffer. |
| [GetDownLevelRasterizationLimit](#getdownlevelrasterizationlimit) | Gets the rasterization limit currently set for this metafile. |
| [GetHENHMETAFILE](#gethenhmetafile) | Gets a Windows handle to an Enhanced Metafile (EMF) file. |
| [GetMetafileHeader](#getmetafileheader) | Gets the metafile header of this metafile. |
| [PlayRecord](#playrecord) | Plays a metafile record. |
| [SetDownLevelRasterizationLimit](#setdownlevelrasterizationlimit) | Sets the resolution for certain brush bitmaps that are stored in this metafile. |

---

# CGpCachedBitmap Class

Creates a **CachedBitmap** object based on a **Bitmap** object and a **Graphics** object. The cached bitmap takes the pixel data from the **Bitmap** object and stores it in a format that is optimized for the display device associated with the **Graphics** object.

**Inherits from**: CGpBase.<br>
**Include file**: CGpBitmap.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructors](#constructorcachedbitmap) | Creates a **CachedBitmap** object based on a **Bitmap** object and a **Graphics** object. |

---

# CGpImageAttributes Class

Creates an **ImageAttributes** object that contains information about how bitmap and metafile colors are manipulated during rendering. An **ImageAttributes** object maintains several color-adjustment settings, including color-adjustment matrices, grayscale-adjustment matrices, gamma-correction values, color-map tables, and color-threshold values.

**Inherits from**: CGpBase.<br>
**Include file**: CGpImageAttributes.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorimageattributes) | Creates an **ImageAttributes** object. |
| [ClearBrushRemapTable](#clearbrushremaptable) | Clears the brush color-remap table of this **ImageAttributes** object. |
| [ClearColorKey](#clearcolorkey) | Clears the color key (transparency range) for a specified category. |
| [ClearColorMatrices](#clearcolormatrices) | Clears the color-adjustment matrix and the grayscale-adjustment matrix for a specified category. |
| [ClearColorMatrix](#clearcolormatrix) | Clears the color-adjustment matrix for a specified category. |
| [ClearGamma](#cleargamma) | Disables gamma correction for a specified category. |
| [ClearNoOp](#clearnoop) | Clears the **NoOp** setting for a specified category. |
| [ClearOutputChannel](#clearoutputchannel) | Clears the cyan-magenta-yellow-black (CMYK) output channel setting for a specified category. |
| [ClearOutputChannelColorProfile](#clearoutputchannelcolorprofile) | Clears the output channel color profile setting for a specified category. |
| [ClearRemapTable](#clearremaptable) | Clears the color-remap table for a specified category. |
| [ClearThreshold](#clearthreshold) | Clears the threshold value for a specified category. |
| [Clone](#clone) | Copies the contents of the existing **ImageAttributes** object into a new **ImageAttributes** object. |
| [GetAdjustedPalette](#getadjustedpalette) | Adjusts the colors in a palette according to the adjustment settings of a specified category. |
| [Reset](#reset) | Clears all color- and grayscale-adjustment settings for a specified category. |
| [SetBrushRemapTable](#setbrushremaptable) | Sets the color remap table for the brush category. |
| [SetColorKey](#setcolorkey) | Sets the color key (transparency range) for a specified category. |
| [SetColorMatrices](#setcolormatrices) | Sets the color-adjustment matrix and the grayscale-adjustment matrix for a specified category. |
| [SetColorMatrix](#setcolormatrix) | Sets the color-adjustment matrix for a specified category. |
| [SetGamma](#setgamma) | Sets the gamma value for a specified category. |
| [SetNoOp](#setnoop) | Turns off color adjustment for a specified category. |
| [SetOutputChannel](#setoutputchannel) | Sets the CMYK output channel for a specified category. |
| [SetOutputChannelColorProfile](#setoutputchannelcolorprofile) | Sets the output channel color-profile file for a specified category. |
| [SetRemapTable](#setremaptable) | Sets the color-remap table for a specified category. |
| [SetThreshold](#setthreshold) | Sets the threshold (transparency range) for a specified category. |
| [SetToIdentity](#settoidentity) | Sets the color-adjustment matrix of a specified category to identity matrix. |
| [SetWrapMode](#setwrapmode) | Sets the the wrap mode of the **ImageAttributes** object. |

---

# <a name="constructorsimage"></a>Constructors (CGpImage)

```
CONSTRUCTOR CGpImage (BYVAL pImage AS CGpImage PTR)
CONSTRUCTOR CGpImage (BYVAL pwszFileName AS WSTRING PTR,  BYVAL useicm AS BOOLEAN = FALSE)
CONSTRUCTOR CGpImage (BYVAL pStream AS IStream PTR, BYVAL useicm AS BOOLEAN = FALSE)
CONSTRUCTOR CGpImage (BYVAL hInst AS HINSTANCE, BYREF wszImageName AS WSTRING)
```

Creates an **Image** object from another **Image** object.

```
CONSTRUCTOR CGpImage (BYVAL pImage AS CGpImage PTR)
```

Creates an **Image** object based on a file.

```
CONSTRUCTOR CGpImage (BYVAL pwszFileName AS WSTRING PTR,  BYVAL useicm AS BOOLEAN = FALSE)
```

Create a **Image** object based on an **IStream** interface.

```
CONSTRUCTOR CGpImage (BYVAL pStream AS IStream PTR, BYVAL useicm AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileName* | Pointer to a null-terminated string that specifies the path name of the image file. The graphics file formats supported by GDI+ are BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF. |
| *pStream* | Pointer to an **IStream** interface. |
| *useicm* | Optional. Boolean value that specifies whether the new Image object applies color correction according to color management information that is embedded in the image file. Embedded information can include International Color Consortium (ICC) profiles, gamma values, and chromaticity information. TRUE specifies that color correction is enabled, and FALSE specifies that color correction is not enabled. The default value is FALSE. |

```
CONSTRUCTOR CGpImage (BYVAL hInst AS HINSTANCE, BYREF wszImageName AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hInst* | A handle to the module whose portable executable file or an accompanying MUI file contains the resource. If this parameter is NULL, the function searches the module used to create the current process. |
| *wszImageName* | Name of the image in the resource file (.RES). If the image resource uses an integral identifier, wszImage should begin with a number symbol (#) followed by the identifier in an ASCII format, e.g., "#998". Otherwise, use the text identifier name for the image. Only images embedded as raw data (type RCDATA) are valid. These must be icons in format .png, .jpg, .gif, .tiff. |

#### Example

This example is not DPI aware. Therefore, unless it is run virtualized, the coordinates won't be scaled, and the image will be displayed smaller or bigger if its resolution is not 96.

```
' ========================================================================================
' The following example creates an Image object based on a JPEG file. Then the code calls
' the DrawImage method to draw the image.
' ========================================================================================
SUB Example_Constructor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc

   ' // Create an Image object, and then clone it.
   DIM image AS CGpImage = "climber.jpg"

   ' // Draw the original image and the cloned image.
   graphics.DrawImage(@image, 10, 10)

END SUB
' ========================================================================================
```

This version is DPI aware. The coordinates will be scaled and the image will be displayed with the same relative size regardless of the DPI of the image. As the **SetResolution** and **SetResolutionForDpi** methods aren't available in the `CGpImage` class, we must use the `CGpBitmap` class.

**Tip**: Always normalize image resolution when displaying multiple images. It’s the only way to guarantee consistent rendering size and layout across devices and formats.

```
' ========================================================================================
' The following example creates an Image object based on a JPEG file. Then the code calls
' the DrawImage method to draw the image.
' ========================================================================================
SUB Example_Constructor_HDPI (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create an Image object
   DIM image AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios od the device
   image.SetResolutionForDpi

   ' // Draw the original image and the cloned image.
   graphics.DrawImage(@image, 10, 10)

END SUB
' ========================================================================================
```
#### Example

The following example draws part of an image.

```
SUB Example_DrawImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set scaling
   graphics.SetPageUnit(UnitPixel)
   graphics.SetPageScale(rxRatio)

   ' // Create an Image object from a resource
   DIM image AS CGpImage = CGpImage(GetModuleHandle(NULL), "#998")   ' // load from resource by ordinal
'   DIM image AS CGpImage = CGpImage(GetModuleHandle(NULL), "IDI_CLIMBER")   ' // load from resource by name

   ' // Draw the original source image.
   graphics.DrawImage(@image, 10, 10)

   ' // Part of the source image to draw
   DIM rcsrc AS GpRectF = (80, 30, 80, 80)
   ' // Destination rectangle
   DIM rcdest AS GpRectF = (200, 10, 80, 80)

   ' // Draw the scaled image.
   graphics.DrawImage(@image, @rcdest, @rcsrc, UnitPixel)

END SUB
```
---

## <a name="cloneimage"></a>Clone (CGpImage)

Copies the contents of the existing **Image** object into a new **Image** object.

```
FUNCTION Clone (BYVAL pCloneImage AS CGpImage PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneImage* | Pointer to the **Image** object where to copy the contents of the existing object. |

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a JPEG file. The code creates a
' second Image object by cloning the first. Then the code calls the DrawImage method twice
' to draw the two images.
' ========================================================================================
SUB Example_Clone (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create an Image object, and then clone it.
   DIM image1 AS CGpImage = "climber.jpg"
   DIM pImage2 AS CGpImage
   image1.Clone(@pImage2)
   ' // You can also use:
   ' DIM pImage2 AS CGpImage = @image1

   ' // Draw the original image and the cloned image.
   graphics.DrawImage(@image1, 20 * rxRatio, 20 * ryRatio)
   graphics.DrawImage(@pImage2, 230 * rxRatio, 20 * ryRatio)

END SUB
' ========================================================================================
```
---

## <a name="findfirstitem"></a>FindFirstItem (CGpImage)

Retrieves the description and the data size of the first metadata item in this **Image** object.

```
FUNCTION FindFirstItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | In, Out. Pointer to an **ImageItemData** structure. On input, the **Desc** member points to a buffer (allocated by the caller) large enough to hold the metadata description (1 byte for JPEG, 4 bytes for PNG, 11 bytes for GIF), and the **DescSize** member specifies the size (1, 4, or 6) of the buffer pointed to by **Desc**. On output, the buffer pointed to by **Desc** receives the metadata description, and the **DataSize** member receives the size, in bytes, of the metadata itself. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

Use **FindFirstItem** along with **FindNextItem** to enumerate the metadata items, including custom metadata, stored in an image. **FindFirstItem** and **FindNextItem** do not enumerate the metadata items stored by the **SetPropertyItem** method.

---

## <a name="findnextitem"></a>FindNextItem (CGpImage)

The **FindNextItem** method is used along with the **FindFirstItem** method to enumerate the metadata items stored in this **Image** object. The **FindNextItem** method retrieves the description and the data size of the next metadata item in this **Image** object. 

```
FUNCTION FindNextItem (BYVAL pitem AS ImageItemData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | In, Out. Pointer to an **ImageItemData** structure. On input, the **Desc** member points to a buffer (allocated by the caller) large enough to hold the metadata description (1 byte for JPEG, 4 bytes for PNG, 11 bytes for GIF), and the **DescSize** member specifies the size (1, 4, or 6) of the buffer pointed to by **Desc**. On output, the buffer pointed to by **Desc** receives the metadata description, and the **DataSize** member receives the size, in bytes, of the metadata itself. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

**FindFirstItem** and **FindNextItem** do not enumerate the metadata items stored by the **SetPropertyItem** method.

---

# <a name="getallpropertyitems"></a>GetAllPropertyItems (CGpImage)

Gets all the property items (metadata) stored in this **Image** object.

```
FUNCTION GetAllPropertyItems (BYVAL totalBufferSize AS UINT, BYVAL numProperties AS UINT, _
   BYVAL allItems AS PropertyItem PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *totalBufferSize* | Integer that specifies the size, in bytes, of the allItems buffer. Call the **GetPropertySize** method to obtain the required size.  |
| *numProperties* | Integer that specifies the number of properties in the image. Call the **GetPropertySize** method to obtain this number. |
| *allItems* | Pointer to an array of **PropertyItem** structures that receives the property items. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

Some image files contain metadata that you can read to determine features of the image. For example, a digital photograph might contain metadata that you can read to determine the make and model of the camera used to capture the image.

GDI+ stores an individual piece of metadata in a PropertyItem object. The **GetAllPropertyItems** method returns an array of PropertyItem objects. Before you call **GetAllPropertyItems**, you must allocate a buffer large enough to receive that array. You can call the **GetPropertySize** method of an **Image** object to get the size, in bytes, of the required buffer. The **GetPropertySize** method also gives you the number of properties (pieces of metadata) in the image.

Several enumerations and constants related to image metadata are defined in Gdiplusimaging.inc.

---

## <a name="getbounds"></a>GetBounds (CGpImage)

Gets the bounding rectangle for this image.

```
FUNCTION GetBounds (BYVAL srcRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *srcRect* | Pointer to a **GpRectF** object that receives the bounding rectangle. |
| *srcUnit* | Pointer to a variable that receives an element of the **GpUnit** enumeration that indicates the unit of measure for the bounding rectangle. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

The bounding rectangle for a metafile does not necessarily have (0, 0) as its upper-left corner. The coordinates of the upper-left corner can be negative or positive, depending on the drawing commands that were issued during the recording of the metafile. For example, suppose a metafile consists of a single ellipse that was recorded with the following statement:

```
DrawEllipse(pen, 200, 100, 80, 40)
```

Then the bounding rectangle for the metafile will enclose that one ellipse. The upper-left corner of the bounding rectangle will not be (0, 0); rather, it will be a point near (200, 100).

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a file and then draws the image.
' Next, the code calls the Image.GetBounds method to get the bounding rectangle for the image
' and redraws a 75 per cent of the image.
' ========================================================================================
SUB Example_GetBounds (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   DIM image AS CGpBitmap = "climber.jpg"
   image.SetResolutionForDpi
   graphics.DrawImage(@image, 10, 10)

   ' // Get the bounding rectangle for the image
   DIM rcf AS GpRectF, unit AS GpUnit
   image.GetBounds(@rcf, @unit)

   ' // Draw 75 percent of the image.
   graphics.DrawImage(@image, 200, 10, rcf.X, rcf.Y, 0.75 * rcf.Width, rcf.Height, unit)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The following example creates an Image object based on a metafile and then draws the image.
' Next, the code calls the Image.GetBounds method to get the bounding rectangle for the image
' and redraws a 75 per cent of the image.
' ========================================================================================
SUB Example_GetBounds (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   DIM image AS CGpBitmap = "SampleMetafile.emf"
   graphics.DrawImage(@image, 10, 10)

   ' // Get the bounding rectangle for the image
   DIM rcf AS GpRectF, unit AS GpUnit
   image.GetBounds(@rcf, @unit)

   ' // Draw 75 percent of the image.
   graphics.DrawImage(@image, 200, 10, rcf.X, rcf.Y, 0.75 * rcf.Width, rcf.Height, unit)

END SUB
' ========================================================================================
```
---

## <a name="getencoderparameterlist"></a>GetEncoderParameterList (CGpImage)

Gets a list of the parameters supported by a specified image encoder.

```
FUNCTION GetEncoderParameterList (BYVAL clsidEncoder AS GUID PTR, BYVAL uSize AS UINT, _
   BYVAL buffer AS EncoderParameters PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *clsidEncoder* | Pointer to a **CLSID** that specifies the encoder. |
| *uSize* | Integer that specifies the size, in bytes, of the buffer array. Call the **GetEncoderParameterListSize** method to obtain the required size. |
| *buffer* | Pointer to an **EncoderParameters** object that receives the list of supported parameters. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

The **GetEncoderParameterList** method returns an array of **EncoderParameter** objects. Before you call **GetEncoderParameterList**, you must allocate a buffer large enough to receive that array, which is part of an **EncoderParameters** object. You can call the **GetEncoderParameterListSize** method to get the size, in bytes, of the required **EncoderParameters** object.

---

## <a name="getencoderparameterlistsize"></a>GetEncoderParameterListSize (CGpImage)

Gets the size, in bytes, of the parameter list for a specified image encoder.

```
FUNCTION GetEncoderParameterListSize (BYVAL clsidEncoder AS GUID PTR) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *clsidEncoder* | Pointer to a **CLSID** that specifies the encoder. |

---

## <a name="getflags"></a>GetFlags (CGpImage)

Gets a set of flags that indicate certain attributes of this **Image** object.

```
FUNCTION GetFlags () AS UINT
```

#### Return value

This method returns a value that holds a set of single-bit flags. The flags are defined in the **ImageFlags** enumeration.

**ImageFlags enumeration**

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *ImageFlagsNone* | 0 | Specifies no format information. |
| *ImageFlagsScalable* | &h0001 | Specifies that the image can be scaled. |
| *ImageFlagsHasAlpha* | &h0002 | Specifies that the pixel data contains alpha values. |
| *ImageFlagsHasTranslucent* | &h0004 | Specifies that the pixel data has alpha values other than 0 (transparent) and 255 (opaque). |
| *ImageFlagsPartiallyScalable* | &h0008 | Specifies that the pixel data is partially scalable with some limitations. |
| *ImageFlagsColorSpaceRGB* | &h0010 | Specifies that the image is stored using an RGB color space. |
| *ImageFlagsColorSpaceCMYK* | &h0020 | Specifies that the image is stored using a CMYK color space. |
| *ImageFlagsColorSpaceGRAY* | &h0040 | Specifies that the image is a grayscale image. |
| *ImageFlagsColorSpaceYCBCR* | &h0080 | Specifies that the image is stored using a YCBCR color space. |
| *ImageFlagsColorSpaceYCCK* | &h0100 | Specifies that the image is stored using a YCCK color space. |
| *ImageFlagsHasRealDPI* | &h1000 | Specifies that dots per inch information is stored in the image. |
| *ImageFlagsHasRealPixelSize* | &h2000 | Specifies that the pixel size is stored in the image. |
| *ImageFlagsReadOnly* | &h00010000 | Specifies that the pixel data is read-only. |
| *ImageFlagsCaching* | &h00020000 | Specifies that the pixel data can be cached for faster access. |

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/CGdiPlus.inc"
USING AfxNova

' ========================================================================================
' The following example retrieves the flags of an image.
' ========================================================================================

' // Create a Bitmap object from a JPEG file.
DIM image AS CGpImage = "climber.jpg"

' // Get the image flags
DIM flags AS UINT = image.GetFlags
PRINT "Image flags &h" & HEX(flags)

' // Interpret the flags
IF (flags AND ImageFlagsHasAlpha) THEN PRINT "Has Alpha Channel"
IF (flags AND ImageFlagsHasRealDPI) THEN PRINT "Has real DPI"
' ... etc.

PRINT
PRINT "Press any key"
SLEEP
```
---

## <a name="getframecount"></a>GetFrameCount (CGpImage)

Gets the number of frames in a specified dimension of this **Image** object.

```
FUNCTION GetFrameCount (BYVAL dimensionID AS GUID PTR) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dimensionID* | Pointer to a GUID that specifies the dimension. GUIDs that identify various dimensions are defined in Gdiplusimaging.inc. |

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/CGdiPlus.inc"
#INCLUDE ONCE "AfxNova/AfxCOM.inc"
USING AfxNova

' ========================================================================================
' Gets the number of frames in a specified dimension of this Image object.
' ========================================================================================

' // Create a Bitmap object from a JPEG file.
DIM image AS CGpImage = "MultiFrame.tiff"

' // How many frame dimensions does the Image object have?
DIM nCount AS UINT = image.GetFrameDimensionsCount
PRINT "The number of dimensions is " & STR(nCount)

' // Get the list of frame dimensions from the Image object.
DIM rgDimensionIDs(nCount - 1) AS GUID
image.GetFrameDimensionsList(@rgDimensionIDs(0), nCount)

' // Display the GUID of the first (and only) frame dimension.
DIM strGuid AS STRING = AfxGuidText(rgDimensionIDs(0))
'   strGuid = AfxGuidText(rgDimensionIDs(0))
PRINT "The first (and only) dimension ID is " & strGuid

' // Get the number of frames in the first dimension.
DIM frameCount AS UINT = image.GetFrameCount(@rgDimensionIDs(0))
PRINT "The number of frames in that dimension is " & STR(frameCount)

PRINT
PRINT "Press any key"
SLEEP
```
---

## <a name="getframedimensionscount"></a>GetFrameDimensionsCount (CGpImage)

Gets the number of frame dimensions in this **Image** object.

```
FUNCTION GetFrameDimensionsCount () AS UINT
```

#### Remarks

This method returns information about multiple-frame images, which come in two styles: multiple page and multiple resolution.

A multiple-page image is an image that contains more than one image. Each page contains a single image (or frame). These pages (or images, or frames) are typically displayed in succession to produce an animated sequence, such as in an animated GIF file.

A multiple-resolution image is an image that contains more than one copy of an image at different resolutions.

Windows GDI+ can support an arbitrary number of pages (or images, or frames), as well as an arbitrary number of resolutions.

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/CGdiPlus.inc"
#INCLUDE ONCE "AfxNova/AfxCOM.inc"
USING AfxNova

' ========================================================================================
' Gets the number of frame dimensions in the image object and the identifier of the frame.
' ========================================================================================

' // Create a Bitmap object from a JPEG file.
DIM image AS CGpImage = "MultiFrame.tiff"

' // How many frame dimensions does the Image object have?
DIM nCount AS UINT = image.GetFrameDimensionsCount
PRINT "The number of dimensions is " & STR(nCount)

' // Get the list of frame dimensions from the Image object.
DIM rgDimensionIDs(nCount - 1) AS GUID
image.GetFrameDimensionsList(@rgDimensionIDs(0), nCount)

' // Display the GUID of the first (and only) frame dimension.
DIM strGuid AS STRING = AfxGuidText(rgDimensionIDs(0))
'   strGuid = AfxGuidText(rgDimensionIDs(0))
PRINT "The first (and only) dimension ID is " & strGuid

' // Get the number of frames in the first dimension.
DIM frameCount AS UINT = image.GetFrameCount(@rgDimensionIDs(0))
PRINT "The number of frames in that dimension is " & STR(frameCount)

PRINT
PRINT "Press any key"
SLEEP
```
---

## <a name="getframedimensionslist"></a>GetFrameDimensionsList (CGpImage)

Gets the identifiers for the frame dimensions of this Image object.

```
FUNCTION GetFrameDimensionsList (BYVAL dimensionIDs AS GUID PTR, BYVAL count AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dimensionIDs* | Pointer to an array that receives the identifiers. GUIDs that identify various dimensions are defined in Gdiplusimaging.inc. |
| *count* | Integer that specifies the number of elements in the dimensionIDs array. Call the **GetFrameDimensionsCount** method to determine this number. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

This method returns information about multiple-frame images, which come in two styles: multiple page and multiple resolution.

A multiple-page image is an image that contains more than one image. Each page contains a single image (or frame). These pages (or images, or frames) are typically displayed in succession to produce an animated sequence, such as in an animated GIF file.

A multiple-resolution image is an image that contains more than one copy of an image at different resolutions.

Windows GDI+ can support an arbitrary number of pages (or images, or frames), as well as an arbitrary number of resolutions.

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/CGdiPlus.inc"
#INCLUDE ONCE "AfxNova/AfxCOM.inc"
USING AfxNova

' ========================================================================================
' Get the list of frame dimensions from the Image object.
' ========================================================================================

' // Create a Bitmap object from a JPEG file.
DIM image AS CGpImage = "MultiFrame.tiff"

' // How many frame dimensions does the Image object have?
DIM nCount AS UINT = image.GetFrameDimensionsCount
PRINT "The number of dimensions is " & STR(nCount)

' // Get the list of frame dimensions from the Image object.
DIM rgDimensionIDs(nCount - 1) AS GUID
image.GetFrameDimensionsList(@rgDimensionIDs(0), nCount)

' // Display the GUID of the first (and only) frame dimension.
DIM strGuid AS STRING = AfxGuidText(rgDimensionIDs(0))
'   strGuid = AfxGuidText(rgDimensionIDs(0))
PRINT "The first (and only) dimension ID is " & strGuid

' // Get the number of frames in the first dimension.
DIM frameCount AS UINT = image.GetFrameCount(@rgDimensionIDs(0))
PRINT "The number of frames in that dimension is " & STR(frameCount)

PRINT
PRINT "Press any key"
SLEEP
```
---

## <a name="getheight"></a>GetHeight (CGpImage)

Gets the image height, in pixels, of this image.

```
FUNCTION GetHeight () AS UINT
```

#### Example

```
'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/CGdiPlus.inc"
USING AfxNova

' ========================================================================================
' The following example retrieves the height of an image.
' ========================================================================================

' // Create a Bitmap object from a JPEG file.
DIM image AS CGpImage = "climber.jpg"

' // Get the image flags
DIM flags AS UINT = image.GetFlags
PRINT "Image flags &h" & HEX(flags)

' // Get the height of the image
DIM nHeight AS UINT = image.GetHeight
PRINT "Image height: " & WSTR(nHeight)

PRINT
PRINT "Press any key"
SLEEP
```
---

## <a name="gethorizontalresolution"></a>GetHorizontalResolution (CGpImage)

Gets the horizontal resolution, in dots per inch, of this image.

```
FUNCTION GetHorizontalResolution () AS SINGLE
```
---

## <a name="getitemdata"></a>GetItemData (CGpImage)

Gets one piece of metadata from this **Image** object.

```
FUNCTION GetItemData (BYVAL pitem AS ImageItemData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | Pointer to an **ImageItemData** object that specifies the item to be retrieved. The **Data** member of the **ImageItemData** object points to a buffer that receives the custom metadata. If the **Data** member is set to NULL, this method returns the size of the required buffer in the **DataSize** member of the **ImageItemData** object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="getpalette"></a>GetPalette (CGpImage)

Gets the **ColorPalette** of this **Image** object.

```
FUNCTION GetPalette (BYVAL pal AS ColorPalette PTR, BYVAL nSize AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pal* | Pointer to a **ColorPalette** structure that receives the palette. |
| *nSize* | Integer that specifies the size, in bytes, of the palette. Call the **GetPaletteSize** method to determine the size. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="getpalettesize"></a>GetPaletteSize (CGpImage)

Gets the size, in bytes, of the color palette of this **Image** object.

```
FUNCTION GetPaletteSize () AS INT_
```
---

## <a name="getphysicaldimension"></a>GetPhysicalDimension (CGpImage)

Gets the width and height of this **Image** object.

```
FUNCTION GetPhysicalDimension (BYVAL psize AS SizeF PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *psize* | Pointer to a **GpSizeF** object that receives the width and height of this image. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="getpixelformat"></a>GetPixelFormat (CGpImage)

Gets the pixel format of this **Image** object.

```
FUNCTION GetPixelFormat () AS PixelFormat
```

#### Return value

This method returns an integer that indicates the pixel format of this **Image** object. The **PixelFormat** data type and constants that represent various pixel formats are defined in Gdipluspixelformats.h.

---

# <a name="getpropertycount"></a>GetPropertyCount (CGpImage)

Gets the number of properties (pieces of metadata) stored in this **Image** object.

```
FUNCTION GetPropertyCount () AS UINT
```
---

# <a name="getpropertyiidlist"></a>GetPropertyIdList (CGpImage)

Gets a list of the property identifiers used in the metadata of this **Image** object.

```
FUNCTION GetPropertyIdList (BYVAL numOfProperty AS UINT, BYVAL list AS PROPID PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *numOfProperty* | Integer that specifies the number of elements in the list array. Call the **GetPropertyCount** method to determine this number. |
| *list* | Pointer to an array that receives the property identifiers. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

The **GetPropertyIdList** method returns an array of **PROPIDs**. Before you call **GetPropertyIdList**, you must allocate a buffer large enough to receive that array. You can call the **GetPropertyCount** method of an **Image** object to determine the size of the required buffer. The size of the buffer should be the return value of **GetPropertyCount** multiplied by **SIZEOF(PROPID)**.

---

# <a name="getpropertyitem"></a>GetPropertyItem (CGpImage)

Gets a specified property item (piece of metadata) from this **Image** object.

```
FUNCTION GetPropertyItem (BYVAL propId AS PROPID, BYVAL propSize AS UINT, _
   BYVAL buffer AS PropertyItem PTR= AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *propId* | Integer that identifies the property item to be retrieved. |
| *propSize* | Integer that specifies the size, in bytes, of the property item to be retrieved. Call the **GetPropertyItemSize** method to determine the size. |
| *buffer* | Pointer to a **PropertyItem** structure that receives the property item. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="getpropertyitemsize"></a>GetPropertyItemSize (CGpImage)

Gets the size, in bytes, of a specified property item of this **Image** object.

```
FUNCTION GetPropertyItemSize (BYVAL propId AS PROPID) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *propId* | Integer that identifies the property item. |

---

## <a name="getpropertysize"></a>GetPropertySize (CGpImage)

Gets the total size, in bytes, of all the property items stored in this **Image** object. The **GetPropertySize** method also gets the number of property items stored in this Image object.

```
FUNCTION GetPropertySize (BYVAL totalBufferSize AS UINT PTR, BYVAL numProperties AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *totalBufferSize* | Pointer to a UINT that receives the total size, in bytes, of all the property items. |
| *numProperties* | Pointer to a UINT that receives the number of property items. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

Windows GDI+ stores an individual piece of metadata in a **PropertyItem** structure. The **GetAllPropertyItems** method returns an array of **PropertyItem** structures. Before you call **GetAllPropertyItems**, you must allocate a buffer large enough to receive that array. You can call the **GetPropertySize** method of an **Image** object to get the size, in bytes, of the required buffer. The **GetPropertySize** method also gives you the number of properties (pieces of metadata) in the image.

---

## <a name="getrawformat"></a>GetRawFormat (CGpImage)

Gets a globally unique identifier ( GUID) that identifies the format of this **Image** object.

```
FUNCTION GetRawFormat (BYVAL guidformat AS GUID PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *guidformat* | Pointer to a GUID that receives the format identifier. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="getthumbnailimage"></a>GetThumbnailImage (CGpImage)

Gets a thumbnail image from this **Image** object.

```
FUNCTION GetThumbnailImage (BYVAL thumbWidth AS UINT, BYVAL thumbHeight AS UINT, _
   BYVAL pThumbnail AS CGpImage PTR, BYVAL pCallback AS ANY PTR = NULL, _
   BYVAL callbackData AS ANY PTR = NULL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *thumbWidth* | Width, in pixels, of the requested thumbnail image. |
| *thumbHeight* | Height, in pixels, of the requested thumbnail image. |
| *pThumbnail* | Pointer to an **Image** object that will receive the requested thumbnail image. |
| *pCallback* | Optional. Callback function that you provide. During the process of creating or retrieving the thumbnail image, GDI+ calls this function to give you the opportunity to abort the process. The default value is NULL. |
| *callbackData* | Optional. Pointer to a block of memory that contains data to be used by the callback function. The default value is NULL. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

A thumbnail image is a small copy of an image. Some image files have a thumbnail image embedded in the file. In such cases, this method retrieves the embedded thumbnail image. If there is no embedded thumbnail image, this method creates a thumbnail image by scaling the main image to the size specified in the *thumbWidth* and *thumbHeight* parameters. If both of those parameters are 0, a system-defined size is used.

#### Example

```
' ========================================================================================
' The following example creates a Bitmap object based on a JPEG file. The code calls the
' Bitmao.GetThumbnailImage method of that Bitmap object and then displays the thumbnail image
' along with the main image.
' ========================================================================================
SUB Example_GetThumbnailImage (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Get a thumbnail of the image of the specified width and height
   DIM thumbnail AS CGpBitmap
   bmp.GetThumbnailImage(70, 50, @thumbnail)

   ' // Draw the original image
   graphics.DrawImage(@bmp, 10, 10)
   ' // Draw the thumbnail image
   graphics.DrawImage(@thumbnail, 220, 10)

END SUB
' ========================================================================================
```
---

## <a name="gettype"></a>GetType (CGpImage)

Gets the type (bitmap or metafile) of this **Image** object.

```
FUNCTION GetType () AS ImageType
```

#### Return value

This method returns an element of the ImageType enumeration that indicates the image type.

---

## <a name="getverticalresolution"></a>GetVerticalResolution (CGpImage)

Gets the vertical resolution, in dots per inch, of this image.

```
FUNCTION GetVerticalResolution () AS SINGLE
```
---

## <a name="getwidth"></a>GetWidth (CGpImage)

Gets the width, in pixels, of this image

```
FUNCTION GetWidth () AS UINT
```
---

# <a name="removepropertyitem"></a>RemovePropertyItem (CGpImage)

Removes a property item (piece of metadata) from this **Image** object.

```
FUNCTION RemovePropertyItem (BYVAL propId AS PROPID) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *propId* | Integer that identifies the property item. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

The **RemovePropertyItem** method removes a specified property from an **Image** object, but that property item is not removed from the file or stream that was used to construct the Image object. To save the image (with the property item removed) to a new JPEG file or stream, call the **Save** method of the **Image** object.

---

## <a name="rotateflip"></a>RotateFlip (CGpImage)

Rotates and flips this image.

```
FUNCTION RotateFlip (BYVAL rotateFlipType AS RotateFlipType) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rotateFlipType* | Element of the **RotateFlipType** enumeration that specifies the type of rotation and the type of flip. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Example

```
' ========================================================================================
' The following example creates a Bitmap object based on a JPEG file. The code calls the
' RotateFlip method to rotate and flip the image 90 degrees.
' ========================================================================================
SUB Example_RotateFlip (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM image AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   image.SetResolutionForDpi
   ' // Draw the bitmap
   graphics.DrawImage(@image, 10, 10)

   ' // Rotate and flip the birmap 90 degrees
   image.RotateFlip(Rotate90FlipY)
   ' // Draw the rotated bitmap
   graphics.DrawImage(@image, 220, 10)

END SUB
' ========================================================================================
```

Because metafiles are a list of drawing commands, not a raster representation, **RotateFlip** can't be used with them. A simulation can be doneusing the **TranslateTransform** and **RotateTransform** methods.

```
' ========================================================================================
' The following example creates an Image object based on a metafile. The code calls the
' TranslateTransform and RotateTransform methods to rotate and flip the image 90 degrees.
' ========================================================================================
SUB Example_RotateFlipMetafile (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Metafile object from file.
   DIM image AS CGpBitmap = "SampleMetafile.emf"

   ' // Translate to rotation origin
   graphics.TranslateTransform(250, 10) ' Adjust as needed

   ' // Rotate 90 degrees clockwise
   graphics.RotateTransform(90)

   graphics.DrawImage(@image, 0, 0)

END SUB
' ========================================================================================
```
---

## <a name="save"></a>Save (CGpImage)

Saves this image to a file.

```
FUNCTION Save (BYVAL pwszFileName AS WSTRING PTR, BYVAL clsidEncoder AS GUID PTR, _
   BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
FUNCTION Save (BYVAL pStream AS IStream PTR, BYVAL clsidEncoder AS GUID PTR, _
   BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileName* | Pointer to a null-terminated string that specifies the path name for the saved image. |
| *pStream* | Pointer to an **IStream** interface. The implementation of **IStream** must include the **Seek**, **Read**, **Write**, and **Stat** methods. |
| *clsidEncoder* | Pointer to a **CLSID** that specifies the encoder to use to save the image. |
| *encoderParams* | Optional. Pointer to an **EncoderParameters** structure that holds parameters used by the encoder. The default value is NULL. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

GDI+ does not allow you to save an image to the same file that you used to construct the image. The following code creates an **Image** object by passing the file name MyImage.jpg to an **Image** constructor. That same file name is passed to the **Save** method of the **Image** object, so the **Save** method fails.

---

## <a name="saveadd"></a>SaveAdd (CGpImage)

Adds a frame to a file or stream specified in a previous call to the **Save** method.

```
FUNCTION SaveAdd (BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
FUNCTION SaveAdd (BYVAL pNewImage AS CGpImage PTR, BYVAL encoderParams AS EncoderParameters PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *encoderParams* | Pointer to an **Image** object that holds the frame to be added. |
| *pNewImage* | Pointer to an **EncoderParameters** structure that holds parameters required by the image encoder used by the save-add operation.  |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="selectactiveframe"></a>SelectActiveFrame (CGpImage)

Selects the frame in this **Image** object specified by a dimension and an index.

```
FUNCTION SelectActiveFrame (BYVAL dimensionID AS GUID PTR, BYVAL frameIndex AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *dimensionID* | Pointer to a GUID that specifies the frame dimension. GUIDs that identify various frame dimensions are defined in Gdiplusimaging.inc. |
| *frameIndex* | Integer that specifies the index of the frame within the specified frame dimension. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

When you call the **SelectActiveFrame** method, all changes that you made to the previously active frame are discarded. If you want to retain changes that you make to a frame, call the Save method before you switch to a different frame.

Among all the image formats currently supported by GDI+, the only formats that support multiple-frame images are GIF and TIFF. When you call the **SelectActiveFrame** method on a GIF image, you should use **FrameDimensionTime**. When you call the SelectActiveFrame method on a TIFF image, you should use **FrameDimensionPage**.

---

## <a name="setpalette"></a>SetPalette (CGpImage)

Sets the color palette of this **Image** object.

```
FUNCTION SetPalette (BYVAL pal AS ColorPalette PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *palette* | Pointer to a **ColorPalette** structure that specifies the palette. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setpropertyitem"></a>SetPropertyItem (CGpImage)

Sets a property item (piece of metadata) for this Image object. If the item already exists, then its contents are updated; otherwise, a new item is added.

```
FUNCTION SetPropertyItem (BYVAL pitem AS PropertyItem PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pitem* | Pointer to a **PropertyItem** object that specifies the property item to be set. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="constructorsbitmap"></a>Constructors (CGpBitmap)

```
CONSTRUCTOR
CONSTRUCTOR (BYVAL pBitmap AS CGpBitmap PTR)
CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
CONSTRUCTOR (BYVAL pStream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL stride AS INT_, _
            BYVAL pxFormat AS PixelFormat, BYVAL scan0 AS UBYTE PTR)
CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pxFormat AS PixelFormat = PixelFormat32bppPARGB)
CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pTarget AS CGpGraphics_ PTR)
CONSTRUCTOR (BYVAL gdiBitmapInfo AS BITMAPINFO PTR, BYVAL gdiBitmapData AS ANY PTR)
CONSTRUCTOR (BYVAL hbm AS HBITMAP, BYVAL hPal AS HPALETTE)
CONSTRUCTOR (BYVAL hicon AS HICON)
CONSTRUCTOR (BYVAL hInstance AS HINSTANCE, BYVAL pwszbitmapName AS WSTRING PTR)
CONSTRUCTOR (BYVAL surface AS IDirectDrawSurface7 PTR)
```
---

Creates a **Bitmap** object based on another **Bitmap** obejct (clones it).

Flat API function: **GdipCloneImage**

```
CONSTRUCTOR CGpBitmap (BYVAL pBitmap AS CGpBitmap PTR)
```

#### Example

```
' ========================================================================================
' The following example creates a Bitmap object from another image object.
' ========================================================================================
SUB Example_CloneConstructor (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   myBitmap.SetResolutionForDpi

   ' // Clone a portion of the bitmap.
   DIM cloneBitmap AS CGpBitmap = @myBitmap

   ' // Draw the clone.
   graphics.DrawImage(@cloneBitmap, 10, 10)

END SUB
' ========================================================================================
```
---

Creates a **Bitmap** object based on an image file.
```
CONSTRUCTOR (BYVAL pwszFileName AS WSTRING PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
```
| Parameter  | Description |
| ---------- | ----------- |
| *pwszFileName* | Pointer to a null-terminated string that specifies the path name of the image file. The graphics file formats supported by GDI+ are BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF. |
| *useEmbeddedColorManagement* | Optional. Boolean value that specifies whether the new **Bitmap** object applies color correction according to color management information that is embedded in the image file. Embedded information can include International Color Consortium (ICC) profiles, gamma values, and chromaticity information. **TRUE** specifies that color correction is enabled, and **FALSE** specifies that color correction is not enabled. The default value is **FALSE**. |

#### Usage example

```
   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   myBitmap.SetResolutionForDpi
```
---

Creates a **Bitmap** object based on an **IStream** COM interface..
```
CONSTRUCTOR (BYVAL pStream AS IStream PTR, BYVAL useEmbeddedColorManagement AS BOOLEAN = FALSE)
```
| Parameter  | Description |
| ---------- | ----------- |
| *pStream* | Pointer to an **IStream** COM interface. |
| *useEmbeddedColorManagement* | Optional. Boolean value that specifies whether the new **Bitmap** object applies color correction according to color management information that is embedded in the image file. Embedded information can include International Color Consortium (ICC) profiles, gamma values, and chromaticity information. **TRUE** specifies that color correction is enabled, and **FALSE** specifies that color correction is not enabled. The default value is **FALSE**. |

#### Usage example

```
   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS CGpBitmap = pstream
   ' // Set the resolution of the image using the DPI ratios
   myBitmap.SetResolutionForDpi
```
---

Creates a **Bitmap** object based on a **Graphics** object, a width, and a height.
```
CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pTarget AS CGpGraphics_ PTR)
```
| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Integer that specifies the width, in pixels, of the bitmap. |
| *nHeight* | Integer that specifies the height, in pixels, of the bitmap. |
| *pTarget* | Pointer to a **Graphics** object that contains information used to initialize certain properties (for example, dots per inch) of the new **Bitmap** object. |

#### Example

```
' ========================================================================================
' This example creates a GpBitmap from a graphics object, inheriting DPI and rendering settings.
' ========================================================================================
SUB Example_CreateFromGraphics (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create bitmap from graphics context
   DIM bmp AS CGpBitmap = CGpBitmap(200, 100, @graphics)

   ' // Create graphics object from bitmap
   DIM bmpGraphics AS CGpGraphics = @bmp

   ' // Fill bitmap with color
   bmpGraphics.Clear(ARGB_RED)

   ' // Draw the bitmap
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```
---

Creates a **Bitmap** object based on an array of bytes along with size and format information.
```
CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL stride AS INT_, _
             BYVAL pxFormat AS PixelFormat, BYVAL scan0 AS UBYTE PTR)
```
| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Integer that specifies the width, in pixels, of the bitmap. |
| *nHeight* | Integer that specifies the height, in pixels, of the bitmap. |
| *stride* | Integer that specifies the byte offset between the beginning of one scan line and the next. This is usually (but not necessarily) the number of bytes in the pixel format (for example, 2 for 16 bits per pixel) multiplied by the width of the bitmap. The value passed to this parameter must be a multiple of four. |
| *pxFormat* | Integer that specifies the pixel format of the bitmap. For more information about pixel format constants, see [Image Pixel Format Constants](https://learn.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-image-pixel-format-constants). |
| *scan0* | Pointer to an array of bytes that contains the pixel data. The caller is responsible for allocating and freeing the block of memory pointed to by this parameter. |

#### Example

```
' ========================================================================================
' This example demonstrates how to create a GpBitmap object from a manually filled pixel
' buffer. It generates a 100×100 pixel image with a vertical gradient and renders it to
' the screen using GDI+.
' ========================================================================================
SUB Example_CreateFromScan0 (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

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
   DIM bmp AS CGpBitmap = CGpBitmap(nWidth, nHeight, stride, PixelFormat32bppARGB, @buffer(0))

   ' // Draw the bitmap
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```
---

Creates a **Bitmap** object of a specified size and pixel format. The pixel data must be provided after the **Bitmap** object is constructed.
```
CONSTRUCTOR (BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, BYVAL pxFormat AS PixelFormat = PixelFormat32bppPARGB)
```
| Parameter  | Description |
| ---------- | ----------- |
| *nWidth* | Integer that specifies the width, in pixels, of the bitmap. |
| *nHeight* | Integer that specifies the height, in pixels, of the bitmap. |
| *pxFormat* | Integer that specifies the pixel format of the bitmap. The default value is **PixelFormat32bppARGB**. For more information about pixel format constants, see [Image Pixel Format Constants](https://learn.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-image-pixel-format-constants). |

---

Creates a **Bitmap** object based on a handle to a Windows Windows Graphics Device Interface (GDI) bitmap and a handle to a GDI palette.
```
CONSTRUCTOR (BYVAL hbm AS HBITMAP, BYVAL hPal AS HPALETTE = FALSE)
```
| Parameter  | Description |
| ---------- | ----------- |
| *hbm* | Handle to a GDI bitmap. |
| *hPal* | Optional. Handle to a GDI palette used to define the bitmap colors if *hbm* is not a device-independent bitmap (DIB). |

Flat API function: **GdipCreateBitmapFromHBITMAP**

#### Example

```
' ========================================================================================
' This example creates a GpBitmap from a legacy HBITMAP and renders it using GDI+.
' ========================================================================================
SUB Example_CreateFromHBITMAP (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

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
      CAST(ULONG PTR, pBits)[i] = ARGB_RED
   NEXT

   ' // Create GDI+ Bitmap from HBITMAP
   DIM bmp AS CGpBitmap = CGpBitmap(hbm)

   ' // Draw the bitmap
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```
---

Creates a **Bitmap** object based on a **BITMAPINFO** structure and an array of pixel data.
```
CONSTRUCTOR (BYVAL gdiBitmapInfo AS BITMAPINFO PTR, BYVAL gdiBitmapData AS ANY PTR)
```
| Parameter  | Description |
| ---------- | ----------- |
| *gdiBitmapInfo* | Pointer to a GDI **BITMAPINFO** structure. This structure defines several bitmap attributes, such as size and pixel format. |
| *gdiBitmapData* | Pointer to an array of bytes that contains the pixel data. |

Flat API function: **GdipCreateBitmapFromGdiDib**

#### Example

```
' ========================================================================================
' This example creates a GpBitmap from a manually constructed BITMAPINFO and pixel buffer.
' ========================================================================================
SUB Example_CreateFromGdiDib (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

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
      .biXPelsPerMeter = 39.3701 * graphics.GetDpiX
      .biYPelsPerMeter = 39.3701 * graphics.GetDpiY
   END WITH

   ' // Create bitmap from DIB
   DIM bmp AS CGpBitmap = CGpBitmap(@bmi, @buffer(0))

   ' // Draw the modified bitmap.
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```
---

Creates a **Bitmap** object based on an icon.
```
CONSTRUCTOR CGpBitmap (BYVAL hicon AS HICON)
```
| Parameter  | Description |
| ---------- | ----------- |
| *hIcon* | Handle to a GDI icon. |

Flat API function: **GdipCreateBitmapFromHICON**

#### Example

```
' ========================================================================================
' This example creates a GpBitmap from an icon and renders it using GDI+.
' ========================================================================================
SUB Example_CreateFromHICON (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Load a system icon (e.g., information icon)
   DIM hIcon AS HICON = LoadIcon(NULL, IDI_INFORMATION)

   ' // Create a GDI+ Bitmap from the HICON
   DIM bmp AS CGpBitmap = CGpBitmap(hIcon)

   ' // Draw the bitmap
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```
---

Creates a **Bitmap** object based on an application or DLL instance handle and the name of a bitmap resource.
```
CONSTRUCTOR CGpBitmap (BYVAL hInstance AS HINSTANCE, BYVAL pwszBitmapName AS WSTRING PTR)
```
| Parameter  | Description |
| ---------- | ----------- |
| *hInstance* | Handle to an instance of a module whose executable file contains a bitmap resource. |
| *pwszBitmapName* | Pointer to a null-terminated string that specifies the path name of the bitmap resource to be loaded. Alternatively, this parameter can consist of the resource identifier in the low-order word and zero in the high-order word. You can use the **MAKEINTRESOURCE** macro to create this value. |

Creates a **Bitmap** object based on a **DirectDraw surface**. The **Bitmap** object maintains a reference to the **DirectDraw** surface until the **Bitmap** object is deleted or goes out of scope.
```
CONSTRUCTOR (BYVAL surface AS IDirectDrawSurface7 PTR)
```
| Parameter  | Description |
| ---------- | ----------- |
| *surface* | Pointer to an **IDrectDrawSurface7** COM interface. |

Flat API function: **GdipCreateBitmapFromDirectDrawSurface**

---

## <a name="applyeffect"></a>ApplyEffect (CGpBitmap)

Alters a **Bitmap** object by applying a specified effect.

```
FUNCTION ApplyEffect (BYVAL pEffect AS CGpEffect PTR, BYVAL roi AS RECT PTR = NULL) AS GpStatus
FUNCTION ApplyEffect (BYVAL pEffect AS CGpEffect PTR, BYVAL roi AS RECT PTR, _
   BYVAL useAuxData AS BOOLEAN, BYVAL auxData AS ANY PTR PTR, BYVAL auxDataSize AS INT_ PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *effect* | [in] The effect to be applied. |
| *roi* | in] Pointer to a **RECT** structure that specifies the portion of the input bitmap to which the effect is applied. Pass **NULL** to specify that the effect applies to the entire input bitmap. |
| *useAuxData* | [in] Flag that specifies whether the function should return a pointer to the auxiliary data that it creates. |
| *auxData* | [out] Pointer to a set of lookup tables. |
| *auxDataSize* | [out] Size, in bytes, of the auxiliary data. |

Flat API function: **GdipBitmapApplyEffect**

#### Example

```
' ========================================================================================
' This example loads an image from disk and applies a blur effect using GDI+ 1.1.
' ========================================================================================
SUB Example_BlurEffect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create a blur effect
   DIM blurEffect AS CGpBlur
   ' // Set parameters: radius = 6.0, expandEdge = FALSE
   DIM blurPrms AS BlurParams = (6.0, FALSE)
   blurEffect.SetParameters(@blurPrms)
   ' // Apply effect to the whole image
   bmp.ApplyEffect(@blurEffect)

   ' // Draw the image
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' This example loads an image from disk and applies a brightness effect using GDI+ 1.1.
' ========================================================================================
SUB Example_Brightnessffect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create a brightness and contrast effect
   DIM brightnessEffect AS CGpBrightnessContrast
   ' // Set parameters: brightness = 50, contrast = 50
   DIM bcParams AS BrightnessContrastParams = (50, 20)
   brightnessEffect.SetParameters(@bcParams)
   ' // Apply effect to the whole image
   bmp.ApplyEffect(@brightnessEffect)

   ' // Draw the image
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' This example loads an image from disk and applies a sharpen effect using GDI+ 1.1.
' ========================================================================================
SUB Example_SharpenEffect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create a sharpen effect
   DIM sharpenEffect AS CGpSharpen
   ' // Set parameters: radius = 3.0, expandEdge = FALSE
   DIM sharpenParams AS SharpenParams = (3.0, FALSE)
   sharpenEffect.SetParameters(@sharpenParams)
   ' // Apply effect to the whole image
   bmp.ApplyEffect(@sharpenEffect)

   ' // Draw the image
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' This example loads an image from disk and applies a red eyes correction effect using GDI+ 1.1.
' ========================================================================================
SUB Example_BitmapRedEyeCorrectionEffect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "RedEyes.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create a red eyes correction effect
   DIM redEyeEffect AS CGpRedEyeCorrection

   ' // We need to specify one or more rectangles that enclose the red-eye areas.
   ' // These are passed as an array of RECT structures.

   ' // Define two rectangles around the eyes
   ' // Change the values according the coordinates of your image
   DIM eyeRects(0 TO 1) AS RECT
   eyeRects(0).left   =   1 : eyeRects(0).top    =   1
   eyeRects(0).right  = 400 : eyeRects(0).bottom = 250
   eyeRects(1).left   =   1 : eyeRects(1).top    =   1
   eyeRects(1).right  = 400 : eyeRects(1).bottom = 250

   ' // Fill a RedEyeCorrectionParams structure
   DIM redeyeParams AS RedEyeCorrectionParams
   redeyeParams.numberOfAreas = 2
   redeyeParams.areas = @eyeRects(0)
   ' // Calculate the size of the paameters
   DIM paramsSize AS UINT = SIZEOF(RedEyeCorrectionParams) + redeyeParams.numberOfAreas * SIZEOF(RECT)
   ' // Set the parameters
   redEyeEffect.SetParameters(@redeyeParams)

   ' // Apply effect to the whole image
   bmp.ApplyEffect(@redEyeEffect)

   ' // Draw the image
   graphics.DrawImage(@bmp, 5, 5)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The HueSaturationLightness effecty enables you to change the hue, saturation, and lightness
' of a bitmap. To specify the magnitudes of the changes in hue, saturation, and lightness,
' pass a HueSaturationLightnessParams structure to the SetParameters method.
' ========================================================================================
SUB Example_BitmapHueSaturationLightnessEffect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create a HueSaturationLightness effect
   DIM hueEffect AS CGpHueSaturationLightness
   ' // Set the parameters
   DIM hslParams AS HueSaturationLightnessParams
   hslParams.hueLevel = 45            ' Rotate hue slightly
   hslParams.saturationLevel = 30     ' Boost saturation
   hslParams.lightnessLevel = -10     ' Slightly darken
   hueEffect.SetParameters(@hslParams)

   ' // Apply effect to the whole image
   bmp.ApplyEffect(@hueEffect)

   ' // Draw the image
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The Levels effect encompasses three bitmap adjustments: highlight, midtone, and shadow.
' You can apply one or more of those adjustments to a bitmap by passing calling the
' ApplyEffect method. To specify the intensities of the adjustments, pass the
' address of a LevelsParams structure to the SetParameters method.
' ========================================================================================
SUB Example_LevelsEffect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create a levels effect
   DIM levelsEffect AS CGpLevels
   ' // Set the parameters
   DIM levelsParams AS LevelsParams
   levelsParams.highlight =  20    ' Boost highlights
   levelsParams.midtone   =  10    ' Slightly brighten midtones
   levelsParams.shadow    = -15    ' Deepen shadows
   levelsEffect.SetParameters(@levelsParams)

   ' // Apply effect to the whole image
   bmp.ApplyEffect(@levelsEffect)

   ' // Draw the image
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' The Tint effect enables you to apply a tint to a bitmap. To specify the nature of the tint,
' pass the address of a TintParams structure to the SetParameters method.
' ========================================================================================
SUB Example_TintEffect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create a tint effect
   DIM tintEffect AS CGpTint
   ' // Set parameters: Hue = +120 (green), Amount = 60 (moderate tint)
   DIM tintPrms AS TintParams = (120, 60)
   tintEffect.SetParameters(@tintPrms)
   ' // Apply effect to the whole image
   bmp.ApplyEffect(@tintEffect)

   ' // Draw the image
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```

#### Example

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

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create a ColorLUT effect
   DIM colorLUTEffect AS CGpColorLUT

   ' // Create LUTs: identity (no change)
   DIM lut AS ColorLUTParams
   FOR i AS LONG = 0 TO 255
      lut.lutB(i) = i
      lut.lutG(i) = i
      lut.lutR(i) = i
      lut.lutA(i) = i
   NEXT

   ' // Example: invert red channel
   FOR i AS LONG = 0 TO 255
      lut.lutR(i) = 255 - i
   NEXT

   ' // Set parameters
   colorLUTEffect.SetParameters(@lut)

   ' // Apply effect to the whole image
   bmp.ApplyEffect(@colorLUTEffect)

   ' // Draw the image
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```

#### Example

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
' The following example applies the effect to the specified region only.
' ========================================================================================
SUB Example_ColorLUTEffectRegion (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create a ColorLut effect
   DIM colorLUTEffect AS CGpColorLUT

   ' // Create LUTs: identity (no change)
   DIM lut AS ColorLUTParams
   FOR i AS LONG = 0 TO 255
      lut.lutB(i) = i
      lut.lutG(i) = i
      lut.lutR(i) = i
      lut.lutA(i) = i
   NEXT

   ' // Example: invert red channel
   FOR i AS LONG = 0 TO 255
      lut.lutR(i) = 255 - i
   NEXT

   ' // Set parameters
   colorLUTEffect.SetParameters(@lut)

   ' // Define the region of interest (ROI)
   DIM roi AS RECT
   roi.left   = 20
   roi.top    = 20
   roi.right  = 150
   roi.bottom = 100

   ' // Apply effect to the whole image
   bmp.ApplyEffect(@colorLUTEffect, @roi)

   ' // Draw the image
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```

#### Example

```
' ========================================================================================
' Converts the image to grayscale by averaging RGB channels.
' ========================================================================================
SUB Example_GrayScaleEffect (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Create grayscale matrix
   DIM grayMatrix(0 TO 4, 0 TO 4) AS SINGLE = {{0.3, 0.3, 0.3, 0.0, 0.0}, _
      {0.59, 0.59, 0.59, 0.0, 0.0}, {0.11, 0.11, 0.11, 0.0, 0.0}, _
      {0.0, 0.0, 0.0, 1.0, 0.0}, {0.0, 0.0, 0.0, 0.0, 1.0}}

   ' // Create a gray matrix effect
   DIM grayEffect AS CGpColorMatrix
   ' // Set parameters
   grayEffect.SetParameters(cast(ColorMatrix PTR, @grayMatrix(0, 0)))
   ' // Apply effect to the whole image
   bmp.ApplyEffect(@grayEffect)

   ' // Draw the image
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```
---

## <a name="clonebitmap"></a>Clone (CGpBitmap)

Creates a new **Bitmap** object from another **Bitmap** object (clones it).

```
FUNCTION Clone (BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
```

Creates a new **Bitmap** object by copying a portion of this bitmap.

```
FUNCTION Clone (BYVAL x AS SINGLE, BYVAL y AS SINGLE, BYVAL nWidth AS SINGLE, _
   BYVAL nHeight AS SINGLE, BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
FUNCTION Clone (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL nWidth AS INT_, BYVAL nHeight AS INT_, _
   BYVAL pxFormat AS PixelFormat, BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
FUNCTION Clone (BYVAL rc AS GpRectF PTR, BYVAL pxFormat AS PixelFormat, _
   BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
FUNCTION Clone (BYVAL rc AS GpRect PTR, BYVAL pxFormat AS PixelFormat, _
   BYVAL pCloneBitmap AS CGpBitmap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | The x-coordinate of the upper-left corner of the rectangle that specifies the portion of this bitmap to copy. |
| *y* | The width of the rectangle that specifies the portion of this bitmap to copy. |
| *nWidth* | The y-coordinate of the upper-left corner of the rectangle that specifies the portion of this bitmap to copy. |
| *nHeight* | The height of the rectangle that specifies the portion of this image to copy. |
| *pxFormat* | Integer that specifies the pixel format of the new bitmap. The **PixelFormat** data type and constants that represent various pixel formats are defined in Gdipluspixelformats.h. |
| *pCloneBitmap* | Pointer to the **Bitmap** object where to copy the contents of the existing object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API functions: **GdipCloneImage** and **GdipCloneBitmapArea**

#### Example

```
' ========================================================================================
' The following example creates a Bitmap from an image file, clones the upper-left
' portion of the image, and then draws the cloned image.
' ========================================================================================
SUB Example_CloneArea (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratio
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS CGpBitmap = "climber.jpg"

   ' // Clone a portion of the bitmap.
   DIM cloneBitmap AS CGpBitmap
   myBitmap.Clone(0, 0, 100, 100, PixelFormatDontCare, @cloneBitmap)

   ' // Draw the clone.
   graphics.DrawImage(@cloneBitmap, 0, 0)

END SUB
' ========================================================================================
```
---

## <a name="convertformat"></a>ConvertFormat (CGpBitmap)

Converts a bitmap to a specified pixel format. The original pixel data in the bitmap is replaced by the new pixel data.

```
FUNCTION ConvertFormat (BYVAL pxFormat AS PixelFormat, BYVAL nDitherType AS DitherType, _
   BYVAL nPaletteType AS PaletteType, BYVAL colourPalette AS ColorPalette PTR, _
   BYVAL alphaThresholdPercent AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pxFormat* | Pixel format constant that specifies the new pixel format. |
| *dithertype* | Element of the **DitherType** enumeration that specifies the dithering algorithm. In cases where the conversion does not reduce the bit depth of the pixel data, pass **DitherTypeNone**. |
| *nPaletteType* | Element of the **PaletteType** enumeration that specifies a standard palette to be used for dithering. If you are converting to a non-indexed format, this parameter is ignored. In that case, pass any element of the **PaletteType** enumeration, say **PaletteTypeCustom**. |
| *colourPalette* | Pointer to a ColorPalette structure that specifies the palette whose indexes are stored in the pixel data of the converted bitmap. This palette (called the actual palette) does not have to have the type specified by the *palettetype* parameter. The *palettetype* parameter specifies a standard palette that can be used by any of the ordered or spiral dithering algorithms. If the actual palette has a type other than that specified by the palettetype parameter, then the **ConvertFormat** method performs a nearest-color conversion from the standard palette to the actual palette. |
| *alphaThresholdPercent* | Single precision number in the range 0 through 100 that specifies which pixels in the source bitmap will map to the transparent color in the converted bitmap. A value of 0 specifies that none of the source pixels map to the transparent color. A value of 100 specifies that any pixel that is not fully opaque will map to the transparent color. A value of t specifies that any source pixel less than t percent of fully opaque will map to the transparent color. Note that for the alpha threshold to be effective, the palette must have a transparent color. If the palette does not have a transparent color, pixels with alpha values below the threshold will map to color that most closely matches (0, 0, 0, 0), usually black. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipBitmapConvertFormat**


#### Example

```
' ========================================================================================
' The CGpBitmap ConvertFormat method wraps the GdipBitmapConvertFormat, which is a GDI+ 1.1
' function that converts the pixel format of an existing bitmap to another format—such as
' from 32bpp ARGB to 8bpp indexed. It replaces the original pixel data with new data that
' matches the target format.
' ========================================================================================
SUB Example_ConvertFormat (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create original bitmap from buffer
   DIM nWidth AS LONG = 100
   DIM nHeight AS LONG = 100
   DIM bytesPerPixel AS LONG = 4
   DIM stride AS LONG = nWidth * bytesPerPixel
   DIM buffer(stride * nHeight - 1) AS UBYTE

   FOR y AS LONG = 0 TO nHeight - 1
      FOR x AS LONG = 0 TO nWidth - 1
         DIM offset AS LONGINT = y * stride + x * bytesPerPixel
         buffer(offset + 0) = x * 255 \ nWidth    ' Blue
         buffer(offset + 1) = y * 255 \ nHeight   ' Green
         buffer(offset + 2) = 128                 ' Red
         buffer(offset + 3) = 255                 ' Alpha
      NEXT
   NEXT

   DIM bmp AS CGpBitmap = CGpBitmap(nWidth, nHeight, stride, PixelFormat32bppARGB, @buffer(0))

   ' // Convert to 8bpp indexed format
   bmp.ConvertFormat(PixelFormat8bppIndexed, DitherTypeNone, PaletteTypeFixedHalftone216, NULL, 0)

   ' // Draw the converted bitmap.
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```
---

## <a name="gethbitmap"></a>GetHBITMAP (CGpBitmap)

Creates a Windows Graphics Device Interface (GDI) bitmap from this **Bitmap** object.

```
FUNCTION GetHBITMAP (BYVAL colorBackground AS ARGB, BYVAL hbmReturn AS HBITMAP PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colorBackground* | ARGB color that specifies the background color. This parameter is ignored if the bitmap is totally opaque. |
| *hbmReturn* | Pointer to an HBITMAP that receives a handle to the GDI bitmap. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipCreateHBITMAPFromBitmap**

#### Example

```
' ========================================================================================
' This example converts a GDI+ Bitmap into a legacy HBITMAP and draws it using GDI.
' ========================================================================================
SUB Example_GetHBITMAP (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Convert to HBITMAP
   DIM hbm AS HBITMAP = bmp.GetHBITMAP(ARGB_WHITE)

   ' // Draw using classic GDI
   DIM memDC AS HDC = CreateCompatibleDC(hdc)
   DIM oldBmp AS HBITMAP = SelectObject(memDC, hbm)

   ' // Retrieve bitmap dimensions from HBITMAP
   DIM bmpInfo AS BITMAP
   GetObject(hbm, SIZEOF(BITMAP), @bmpInfo)

   ' // Apply DPI scaling
   DIM scaledWidth AS LONG = bmpInfo.bmWidth * graphics.GetDpiX / 96
   DIM scaledHeight AS LONG = bmpInfo.bmHeight * graphics.GetDpiY / 96
   StretchBlt(hdc, 0, 0, scaledWidth, scaledHeight, memDC, 0, 0, bmpInfo.bmWidth, bmpInfo.bmHeight, SRCCOPY)

END SUB
' ========================================================================================
```
---

## <a name="gethicon"></a>GetHICON (CGpBitmap)

Creates an icon from this **Bitmap** object.

```
FUNCTION GetHICON (BYVAL hiconReturn AS HICON PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hIconReturn* | Pointer to an HICON that receives a handle to the icon. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipCreateHICONFromBitmap**

#### Example

```
' ========================================================================================
' This example converts a GDI+ Bitmap into a Windows HICON and displays it.
' ========================================================================================
SUB Example_GetHICON (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Convert to HBITMAP
   DIM hIcon AS HICON = bmp.GetHICON

   ' // Get width and height of the icon
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
   DIM scaledWidth AS LONG = iconWidth * graphics.GetDpiX / 96
   DIM scaledHeight AS LONG = iconHeight * graphics.GetDpiY / 96

   ' // Draw the icon
   DrawIconEx(hdc, 0, 0, hIcon, scaledWidth, scaledHeight, 0, NULL, DI_NORMAL)

END SUB
' ========================================================================================
```
---

## <a name="gethistogram"></a>GetHistogram (CGpBitmap)

Returns one or more histograms for specified color channels of this **Bitmap** object.

```
FUNCTION GetHistogram (BYVAL nFormat AS HistogramFormat, BYVAL NumberOfEntries AS UINT, _
   BYVAL channel0 AS UINT PTR, BYVAL channel1 AS UINT PTR, BYVAL channel2 AS UINT PTR, _
   BYVAL channel3 AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nFormat* | Element of the **HistogramFormat** enumeration that specifies the channels for which histograms will be created. |
| *NumberOfEntries* | Integer that specifies the number of elements (of type UINT) in each of the arrays pointed to by *channel0*, *channel1*, *channel2*, and *channel3*. You must allocate memory for those arrays before you call **GetHistogram**. To determine the required number of elements, call **GetHistogramSize**. |
| *channel0* | Pointer to an array of DWORDs that receives the first histogram. |
| *channel1* | Pointer to an array of DWORDs that receives the second histogram if there is a second histogram. Pass NULL if there is no second histogram. |
| *channel2* | Pointer to an array of DWORDs that receives the third histogram if there is a third histogram. Pass NULL if there is no third histogram. |
| *channel3* | Pointer to an array of DWORDs that receives the fourth histogram if there is a fourth histogram. Pass NULL if there is no fourth histogram. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

The number of histograms returned depends on the HistogramFormat enumeration element passed to the format parameter. For example, if format is equal to **HistogramFormatRGB**, then three histograms are returned: one each for the red, green, and blue channels. In that case, *channel0* points to the array that receives the red-channel histogram, *channel1* points to the array that receives the green-channel histogram, and *channel2* points to the array that receives the blue-channel histogram. For **HistogramFormatRGB**, *channel3* must be set to NULL because there is no fourth histogram. For more details, see the **HistogramFormat** enumeration.

Flat API function: **GdipBitmapGetHistogram**

#### Example

```
' ========================================================================================
' This example loads an image from disk and gets the histogram.
' ========================================================================================
SUB Example_GetHistogram (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scale transform to the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Get histogram size
   DIM entries AS UINT = bmp.GetHistogramSize(HistogramFormatARGB)

   ' // Allocate histogram arrays
   DIM red(entries - 1) AS UINT
   DIM green(entries - 1) AS UINT
   DIM blue(entries - 1) AS UINT
   DIM alpha(entries - 1) AS UINT

   ' // Extract histograms
   bmp.GetHistogram(HistogramFormatARGB, entries, @red(0), @green(0), @blue(0), @alpha(0))

   ' // Example: print peak red value
   DIM maxRed AS UINT = 0
   FOR i AS UINT = 0 TO entries - 1
      IF red(i) > maxRed THEN maxRed = red(i)
   NEXT
   AfxMsg ("Max red intensity count: " & WSTR(maxRed))

   ' // Draw the image
   graphics.DrawImage(@bmp, 0, 0)

END SUB
' ========================================================================================
```
---

## <a name="gethistogramsize"></a>GetHistogramSize (CGpBitmap)

Returns the number of elements (in an array of DWORDs) that you must allocate before you call the **GetHistogram** method of a **Bitmap** object.

```
FUNCTION GetHistogramSize (BYVAL nFormat AS HistogramFormat, BYVAL NumberOfEntries AS UINT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nFormat* | Element of the **HistogramFormat** enumeration that specifies the pixel format of the bitmap. |
| *NumberOfEntries* | Pointer to a DWORD that receives the number of entries. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipBitmapGetHistogramSize**

#### Usage example

```
' ========================================================================================
' This example loads an image from disk and gets the histogram.
' ========================================================================================
SUB Example_GetHistogram (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scale transform to the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Get histogram size
   DIM entries AS UINT = bmp.GetHistogramSize(HistogramFormatARGB)

   ' // Allocate histogram arrays
   DIM red(entries - 1) AS UINT
   DIM green(entries - 1) AS UINT
   DIM blue(entries - 1) AS UINT
   DIM alpha(entries - 1) AS UINT

   ' // Extract histograms
   bmp.GetHistogram(HistogramFormatARGB, entries, @red(0), @green(0), @blue(0), @alpha(0))

   ' // Example: print peak red value
   DIM maxRed AS UINT = 0
   FOR i AS UINT = 0 TO entries - 1
      IF red(i) > maxRed THEN maxRed = red(i)
   NEXT
   AfxMsg ("Max red intensity count: " & WSTR(maxRed))

   ' // Draw the image
   graphics.DrawImage(@bmp, 0, 0)

END SUB
' ========================================================================================
```
---

## <a name="getpixel"></a>GetPixel (CGpBitmap)

Gets the color of a specified pixel in this bitmap.

```
FUNCTION GetPixel (BYVAL x AS LONG, BYVAL y AS LONG) AS ARGB
FUNCTION GetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL colour AS ARGB PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Integer that specifies the x-coordinate (column) of the pixel. |
| *y* | Integer that specifies the y-coordinate (row) of the pixel. |
| *colour* | Pointer to a DWORD that receives the color of the specified pixel. |

#### Return value

First overloaded function: Returns the color value.

Second overloaded function: If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration. If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

Depending on the format of the bitmap, **GetPixel** might not return the same value as was set by **SetPixel**. For example, if you call **SetPixel** on a **Bitmap** object whose pixel format is 32bppPARGB, the pixel's RGB components are premultiplied. A subsequent call to **GetPixel** might return a different value because of rounding. Also, if you call **SetPixel** on a **Bitmap** object whose color depth is 16 bits per pixel, information could be lost during the conversion from 32 to 16 bits, and a subsequent call to **GetPixel** might return a different value.

Flat API function: **GdipBitmapGetPixel**

#### Example

```
' ========================================================================================
' The following example creates a Bitmap object based on a JPEG file. The code calls the
' GetPixel method to obtain the color of a pixel in the bitmap and then fills a rectangle
' with the retrieved color.
' ========================================================================================
SUB Example_GetPixel (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS CGpBitmap = "climber.jpg"

   ' // Fill a rectangle with the pixel color.
   DIM brush AS CGpSolidBrush = myBitmap.GetPixel(50, 50)
   graphics.FillRectangle(@brush, 10, 10, 100, 100)

END SUB
' ========================================================================================
```
---

## <a name="initializepalette"></a>InitializePalette (CGpBitmap)

Initializes a standard, optimal, or custom color palette.

```
FUNCTION InitializePalette (BYVAL colourPalette AS ColorPalette PTR, BYVAL nPaletteType AS PaletteType, _
   BYVAL optimalColors AS INT_, BYVAL useTransparentColor AS BOOL, BYVAL pBitmap AS CGpBitmap PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colourPalette* | Pointer to a buffer that contains a **ColorPalette** structure followed by an array of ARGB values. The Entries member of a **ColorPalette** structure is an array of one ARGB value. You must allocate memory for the **ColorPalette** structure and for the additional ARGB values in the palette. For example, if the palette has 36 ARGB values, allocate a buffer as follows: HeapAlloc(SIZEOF(ColorPalette) + 35\*4). |
| *palettetype* | Element of the **PaletteType** enumeration that specifies the palette type. The palette can have one of several standard types, or it can be a custom palette that you define. Also, the **InitializePalette** method can create an optimal palette based on a specified bitmap. |
| *optimalColors* | Integer that specifies the number of colors you want to have in an optimal palette based on a specified bitmap. If this parameter is greater than 0, the *palettetype* parameter must be set to **PaletteTypeOptimal**, and the bitmap parameter must point to a **Bitmap** object. If you are creating a standard or custom palette rather than an optimal palette, set this parameter to 0. |
| *useTransparentColor* | Boolean value that specifies whether to include the transparent color in the palette. Set to TRUE to include the transparent color; otherwise FALSE. |
| *pBitmap* | Pointer to a **Bitmap** object for which an optimal palette will be created. If palettetype is set to **PaletteTypeOptimal** and optimalColors is set to a positive integer, set this parameter to the address of a **Bitmap** object. Otherwise, set this parameter to NULL. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API format: **GdipInitializePalette**

---

## <a name="lockbits"></a>LockBits (CGpBitmap)

Locks a rectangular portion of this bitmap and provides a temporary buffer that you can use to read or write pixel data in a specified format. Any pixel data that you write to the buffer is copied to the **Bitmap** object when you call **UnlockBits**.

```
FUNCTION LockBits (BYVAL rc AS GpRect PTR, BYVAL flags AS UINT, BYVAL pxFormat AS PixelFormat, _
   BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *rc* | Pointer to a rectangle that specifies the portion of the bitmap to be locked. |
| *flags* | Set of flags that specify whether the locked portion of the bitmap is available for reading or for writing and whether the caller has already allocated a buffer. Individual flags are defined in the **ImageLockMode** enumeration. |
| *pxFormat* |  Integer that specifies the format of the pixel data in the temporary buffer. The pixel format of the temporary buffer does not have to be the same as the pixel format of this **Bitmap** object. The **PixelFormat** data type and constants that represent various pixel formats are defined in Gdipluspixelformats.h. GDI+ version 1.0 does not support processing of 16-bits-per-channel images, so you should not set this parameter equal to PixelFormat48bppRGB, PixelFormat64bppARGB, or PixelFormat64bppPARGB. |
| *lockedBitmapData* | Pointer to a **BitmapData** object. If the **ImageLockModeUserInputBuf** flag of the flags parameter is cleared, then *lockedBitmapData* serves only as an output parameter. In that case, the *Scan0* data member of the **BitmapData** object receives a pointer to a temporary buffer, which is filled with the values of the requested pixels. The other data members of the **BitmapData** object receive attributes (width, height, format, and stride) of the pixel data in the temporary buffer. If the pixel data is stored bottom-up, the *Stride* data member is negative. If the pixel data is stored top-down, the *Stride* data member is positive. If the **ImageLockModeUserInputBuf** flag of the *flags* parameter is set, then *lockedBitmapData* serves as an input parameter (and possibly as an output parameter). In that case, the caller must allocate a buffer for the pixel data that will be read or written. The caller also must create a **BitmapData** object, set the *Scan0* data member of that **BitmapData** object to the address of the buffer, and set the other data members of the **BitmapData** object to specify the attributes (width, height, format, and stride) of the buffer. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipBitmapLockBits**

#### Example

```
' ========================================================================================
' Demonstrates the use of GdipBitmapLockBits/GdipBitmapUnlockBits.
' This example inverts the colors of a Bitmap.
' ========================================================================================
SUB Example_LockBits (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Get the width and height of the bitmap
   DIM nWidth AS DWORD = bmp.GetWidth
   DIM nHeight AS DWORD = bmp.GetHeight

   ' // Define lock rectangle
   DIM rc AS GpRect = (0, 0, nWidth, nHeight)

   ' // Lock bitmap bits
   DIM bmpData AS BitmapData
   bmp.LockBits(@rc, ImageLockModeRead OR ImageLockModeWrite, PixelFormat32bppARGB, @bmpData)

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
   bmp.UnlockBits(@bmpData)

   ' // Draw the modified bitmap.
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```
---

## <a name="setpixel"></a>SetPixel (CGpBitmap)

Sets the color of a specified pixel in this bitmap.

```
FUNCTION SetPixel (BYVAL x AS LONG, BYVAL y AS LONG, BYVAL colour AS ARGB) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *x* | Integer that specifies the x-coordinate (column) of the pixel. |
| *y* | Integer that specifies the y-coordinate (row) of the pixel. |
| *colour* | Pointer to a variable that receives the color to set.|

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

Depending on the format of the bitmap, **GetPixel** might not return the same value as was set by **SetPixel**. For example, if you call SetPixel on a Bitmap object whose pixel format is 32bppPARGB, the RGB components are premultiplied. A subsequent call to **GetPixel** might return a different value because of rounding. Also, if you call **SetPixel** on a **Bitmap** whose color depth is 16 bits per pixel, information could be lost in the conversion from 32 to 16 bits, and a subsequent call to **GetPixel** return a different value.

Flat API function: **GdipBitmapSetPixel**

#### Example

```
' ========================================================================================
' The following example creates a Bitmap object based on a JPEG file. The code draws the
' bitmap once unaltered. Then the code calls the SetPixel method to create a
' checkered pattern of black pixels in the bitmap and draws the altered bitmap.
' ========================================================================================
SUB Example_SetPixel (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM myBitmap AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   myBitmap.SetResolutionForDpi

   '// Draw the bitmap
   graphics.DrawImage(@myBitmap, 10, 10)

   ' // Get the width and height of the bitmap
   DIM nWidth AS DWORD = myBitmap.GetWidth
   DIM nHeight AS DWORD = myBitmap.GetHeight

   ' // Make a checkered pattern of black pixels
   FOR row AS LONG = 0 TO nWidth - 1 STEP 2
      FOR col AS LONG = 0 TO nHeight STEP 2
         myBitmap.SetPixel(row, col, ARGB_BLACK)
      NEXT
   NEXT

   ' // Draw the altered bitmap.
   graphics.DrawImage(@myBitmap, 200, 10)

END SUB
' ========================================================================================
```
---

## <a name="setresolution"></a>SetResolution (CGpBitmap)

Sets the resolution of this **Bitmap** object.

```
FUNCTION SetResolution (BYVAL xdpi AS SINGLE, BYVAL ydpi AS SINGLE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *xdpi* | The horizontal resolution in dots per inch. |
| *ydpi* | The vertical resolution in dots per inch. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipBitmapSetResolution**

#### Usage example

```
   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolution(rxRatio, ryRatio)
```
---

## <a name="setresolutionfordpi"></a>SetResolutionForDpi (CGpBitmap)

Sets the resolution of this **Bitmap** object. Uses the current DPI seetings to set the horizontal and vertical resolution in dots per inch.

```
FUNCTION SetResolutionForDpi () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipBitmapSetResolution**

#### Usage example

```
   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scale transform using the DPI settings
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI settings
   bmp.SetResolutionForDpi
```

---

## <a name="unlockbits"></a>UnlockBits (CGpBitmap)

Unlocks a portion of this bitmap that was previously locked by a call to **LockBits**.

```
FUNCTION UnlockBits (BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *lockedBitmapData* | Pointer to a **BitmapData** object that was previously passed to **LockBits**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

**LockBits** and **UnlockBits** must be used as a pair. A call to the **LockBits** method of a **Bitmap** object establishes a temporary buffer that you can use to read or write pixel data in a specified format. After you write to the temporary buffer, a call to **UnlockBits** copies the pixel data in the buffer to the **Bitmap** object. If the pixel format of the temporary buffer is different from the pixel format of the **Bitmap** object, the pixel data is converted appropriately.

Flat API function: **GdipBitmapUnlockBits**

#### Example

```
' ========================================================================================
' Demonstrates the use of GdipBitmapLockBits/GdipBitmapUnlockBits.
' This example inverts the colors of a Bitmap.
' ========================================================================================
SUB Example_LockBits (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Set the scaling factors using the DPI ratios
   graphics.ScaleTransformForDpi

   ' // Create a Bitmap object from a JPEG file.
   DIM bmp AS CGpBitmap = "climber.jpg"
   ' // Set the resolution of the image using the DPI ratios
   bmp.SetResolutionForDpi

   ' // Get the width and height of the bitmap
   DIM nWidth AS DWORD = bmp.GetWidth
   DIM nHeight AS DWORD = bmp.GetHeight

   ' // Define lock rectangle
   DIM rc AS GpRect = (0, 0, nWidth, nHeight)

   ' // Lock bitmap bits
   DIM bmpData AS BitmapData
   bmp.LockBits(@rc, ImageLockModeRead OR ImageLockModeWrite, PixelFormat32bppARGB, @bmpData)

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
   bmp.UnlockBits(@bmpData)

   ' // Draw the modified bitmap.
   graphics.DrawImage(@bmp, 10, 10)

END SUB
' ========================================================================================
```
---

## <a name="constructorsmetafile"></a>Constructors (CGpMetafile)

Creates a Windows GDI+ Metafile object for playback based on a Windows Graphics Device Interface (GDI) Enhanced Metafile (EMF) file.

```
CONSTRUCTOR CGpMetafile (BYVAL hEmf AS HENHMETAFILE, BYVAL deleteEmf AS BOOL = FALSE)
```

Creates a **Metafile** object from an **IStream** interface for playback.

```
CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR)
```

Creates a **Metafile** object for playback.

```
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR)
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR)
CONSTRUCTOR CGpMetafile (BYVAL hWmf AS HMETAFILE, BYVAL wmfPFH AS WmfPlaceableFileHeader PTR, _
   BYVAL deleteWmf AS BOOL = FALSE)
```

Creates a **Metafile** object for recording.

```
CONSTRUCTOR CGpMetafile (BYVAL referenceHdc AS HDC, BYVAL nType AS EmfType = EmfTypeEmfPlusDual, _
   BYVAL description AS WSTRING PTR = NULL)
CONSTRUCTOR CGpMetafile (BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRectF PTR, _
   BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
CONSTRUCTOR CGpMetafile (BYVAL referenceHdc AS HDC, BYVAL frmRect AS GpRect PTR, _
   BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
```

Creates a **Metafile** object for recording.

```
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, _
   BYVAL frmRect AS GpRectF PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
CONSTRUCTOR CGpMetafile (BYVAL pwszFileName AS WSTRING PTR, BYVAL referenceHdc AS HDC, _
   BYVAL frmRect AS GpRect PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType = EmfTypeEmfPlusDual, BYVAL description AS WSTRING PTR = NULL)
```

Creates a **Metafile** object from an **IStream** interface for recording.

```
CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, _
   BYVAL nType AS EmfType, BYVAL description AS WSTRING PTR)
CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, _
   BYVAL frmRect AS GpRectF PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType, BYVAL description AS WSTRING PTR)
CONSTRUCTOR CGpMetafile (BYVAL pStream AS IStream PTR, BYVAL referenceHdc AS HDC, _
   BYVAL frmRect AS GpRect PTR, BYVAL frameUnit AS MetafileFrameUnit = MetafileFrameUnitGdi, _
   BYVAL nType AS EmfType, BYVAL description AS WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hEmf* | Windows handle to an enhanced-format metafile. |
| *deleteEmf* | Optional. Boolean value that specifies whether the Windows handle to a metafile is deleted when the **Metafile** object is deleted. TRUE specifies that the *hEmf* Windows handle is deleted, and FALSE specifies that the *hEmf* Windows handle is not | *pStream* | Pointer to a **IStream** interface that points to a data stream in a file. If the **Metafile** has been created for recording, when the commands are recorded, they will be saved to this stream. |
| *pwszFileName* | Pointer to a wide-character string that specifies the name of an existing disk file used to create the **Metafile** object for playback.  |
| *hWmf* | Windows handle to an windows metafile format. |
| *wmfPFH* | Pointer to a **WmfPlaceableFileHeader** structure that specifies a preheader preceding the metafile header. |
| *deleteWmf* | Optional. Boolean value that specifies whether the Windows handle to a metafile is deleted when the **Metafile** object is deleted. TRUE specifies that the *hWmf* Windows handle is deleted, and FALSE specifies that the *hWmf* Windows handle is not | *referenceHdc* | Windows handle to a metafile. |
| *frameRect* | Reference to a rectangle that bounds the metafile display. |
| *frameUnit* | Optional. Element of the **MetafileFrameUnit** enumeration that specifies the unit of measure for **frameRect**. The default value is **MetafileFrameUnitGdi**. |
| *nType* | Optional. Element of the **EmfType** enumeration that specifies the type of metafile that will be recorded. The default value is **EmfTypeEmfPlusDual**. |
| *description* | Optional. Pointer to a wide-character string that specifies the descriptive name of the metafile. The default value is NULL. |

#### Remarks

When recording to a file, the file must be writable, and Windows GDI+ must be able to obtain an exclusive lock on the file.

---

## <a name="emftowmfbits"></a>EmfToWmfBits (CGpMetafile)

Converts an enhanced-format metafile to a Windows Metafile Format (WMF) metafile and stores the converted records in a specified buffer.

```
FUNCTION EmfToWmfBits (BYVAL hEmf AS HENHMETAFILE, BYVAL cbData16 AS UINT, BYVAL pData16 AS BYTE PTR, _
   BYVAL iMapMode AS INT_ = MM_ANISOTROPIC, BYVAL eFlags AS INT_ = EmfToWmfBitsFlagsDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *hEmf* | Windows handle to an enhanced-format metafile. |
| *cbData16* | Unsigned integer that specifies the number of bytes in the buffer pointed to by the *pData16* parameter. |
| *pData16* | Pointer to a buffer that receives the converted records. If *pData16* is NULL, **EmfToWmfBits** returns the number of bytes required to store the converted metafile records. |
| *iMapMode* | Optional. Specifies the mapping mode to use in the converted metafile. For a list of possible mapping modes, see **SetMapMode**. The default value is MM_ANISOTROPIC. |
| *eFlags* | Optional. Element of the **EmfToWmfBitsFlags** enumeration that specifies options for the conversion. The default value is **EmfToWmfBitsFlagsDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

This method replaces the records originally in the Metafile object with the converted records. To retain a copy of the original **Metafile** object, call the **Clone** method.

If you set the *emfType* parameter to **EmfTypeEmfPlusDual**, the converted metafile contains an Enhanced Metafile (EMF) representation and an EMF+ representation. The EMF representation is the original set of EMF records rather than EMF records converted back from the newly created EMF+ records.

It is possible for the return value to be **Ok** and the value returned in *conversionSuccess* to be FALSE. Sometimes the overall conversion is considered to be successful even if a few individual records failed to convert with complete accuracy. For example, the original metafile might have records or operations that are not supported by Windows GDI+ (or EMF+), in which case those records or operations are emulated.

---

## <a name="getdownlevelrasterizationlimit"></a>GetDownLevelRasterizationLimit (CGpMetafile)

Gets the rasterization limit currently set for this metafile. The rasterization limit is the resolution used for certain brush bitmaps that are stored in the metafile. For a detailed explanation of the rasterization limit, see **SetDownLevelRasterizationLimit**.

```
FUNCTION GetDownLevelRasterizationLimit () AS UINT
```
#### Return value

This method returns the rasterization limit in dpi.

---

## <a name="gethenhmetafile"></a>GetHENHMETAFILE (CGpMetafile)

Gets a Windows handle to an Enhanced Metafile (EMF) file.

```
FUNCTION GetHENHMETAFILE () AS HENHMETAFILE
```
#### Return value

This method sets the **Metafile** object to an invalid state. The user is responsible for calling **DeleteEnhMetafile**, to delete the Windows handle.

---

## <a name="getmetafileheader"></a>GetMetafileHeader (CGpMetafile)

Gets the metafile header of this metafile.

```
FUNCTION GetMetafileHeader (BYVAL mh AS MetafileHeader PTR) AS GpStatus
FUNCTION GetMetafileHeader (BYVAL hEmf AS HENHMETAFILE, BYVAL mh AS MetafileHeader PTR) AS GpStatus
FUNCTION GetMetafileHeader (BYVAL hWmf AS HMETAFILE, BYVAL mh AS MetafileHeader PTR) AS GpStatus
FUNCTION GetMetafileHeader (BYVAL pwsFileName AS WSTRING PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
FUNCTION GetMetafileHeader (BYVAL pStream AS IStream PTR, BYVAL mh AS MetafileHeader PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *mf* | Pointer to a MetafileHeader structure that receives the header, which contains the attributes of the metafile. |
| *hEmf* | Windows handle to an enhanced-format metafile. |
| *hWmf* | Windows handle to an windows metafile format. |
| *wmfPFH* | Pointer to a placeable metafile header. |
| *pwszFileName* | Pointer to a wide-character string that specifies the name of an existing **Metafile** object that contains the header. |
| *pStream* | Pointer to a **IStream** interface that points to a data stream in a file. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="playrecord"></a>PlayRecord (CGpMetafile)

Plays a metafile record.

```
FUNCTION PlayRecord (BYVAL recordType AS EmfPlusRecordType, BYVAL flags AS UINT, _
   BYVAL dataSize AS UINT, BYVAL pData AS BYTE PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *recordType* | Element of the **EmfPlusRecordType** enumeration that specifies the type of metafile record to be played. |
| *flags* | Set of flags that specify attributes of the record to be played. |
| *dataSize* | Integer that specifies the number of bytes contained in the record data. |
| *pData* | Pointer to an array of bytes that contains the record data. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setdownlevelrasterizationlimit"></a>SetDownLevelRasterizationLimit (CGpMetafile)

Sets the resolution for certain brush bitmaps that are stored in this metafile.

```
FUNCTION SetDownLevelRasterizationLimit (BYVAL metafileRasterizationLimitDpi AS UINT) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *metafileRasterizationLimitDpi* | Non-negative integer that specifies the resolution in dpi. If you set this parameter equal to 0, the resolution is set to match the resolution of the device context handle that was passed to the **Metafile** constructor. If you set this parameter to a value greater than 0 but less than 10, the resolution is left unchanged. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

The purpose of this method is to prevent metafiles from becoming too large as a result of texture and gradient brushes being stored at high resolution. Suppose you construct a **Metafile** object (for recording an **EmfTypeEmfOnly** metafile) based on the device context of a printer that has a resolution of 600 dpi. Also suppose you create a path gradient brush or a texture brush based on a **Bitmap** object that has a resolution of 96 dpi. If the bitmap that represents that brush is stored in the metafile with a resolution of 96 dpi, it will require much less space than if it is stored with a resolution of 600 dpi.

The default rasterization limit for metafiles is 96 dpi. So if you do not call this method at all, path gradient brush and texture brush bitmaps are stored with a resolution of 96 dpi.

The rasterization limit has an effect on metafiles of type **EmfTypeEmfOnly** and **EmfTypeEmfPlusDual**, but it has no effect on metafiles of type **EmfTypeEmfPlusOnly**.

---

## <a name="constructorcachedbitmap"></a>Constructor (CGpCachedBitmap)

Creates a **CachedBitmap** object based on a **Bitmap** object and a **Graphics** object. The cached bitmap takes the pixel data from the **Bitmap** object and stores it in a format that is optimized for the display device associated with the **Graphics** object.

```
CONSTRUCTOR CGpCachedBitmap (BYVAL pBitmap AS CGpBitmap PTR, BYVAL pGraphics AS CGpGraphics PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pBitmap* | Pointer to a **Bitmap** object that contains the pixel data to be optimized.  |
| *pGraphics* | Pointer to a **Graphics** object that is associated with a display device for which the image will be optimized. |

#### Remarks

You can display a cached bitmap by passing the address of a **CachedBitmap** object to the **DrawCachedBitmap** method of a **Graphics** object. Use the **Graphics** object that was passed to the **CachedBitmap** constructor or another **Graphics** object that represents the same device.

#### Example

```
' ========================================================================================
' The following example creates a CachedBitmap object based on a Bitmap object and a
' Graphics object. The code calls the DrawCachedBitmap method of that Graphics object
' to display the cached bitmap.
' ========================================================================================
SUB Example_CachedBitmap (BYVAL hdc AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc

   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' ScaleTransform can't be used with cached bitmaps, that can't be resized

   ' // Load the image
   DIM myBitmap AS CGpBitmap = "Climber.jpg"

   ' // Create a cached bitmap
   DIM cachedBitmap AS CGpCachedBitmap = CGpCachedBitmap(@myBitmap, @graphics)

   ' // Draw the cached bitmap
   graphics.DrawCachedBitmap(@cachedBitmap, 10 * rxRatio, 10 * ryRatio)

END SUB
' ========================================================================================
```
---

## <a name="constructorimageattributes"></a>Constructor (CGpImageAttributes)

Creates a new **ImageAttributes** object. This constructor is the default constructor and has no parameters.

```
CONSTRUCTOR CGpImageAttributes
```

Creates and initializes an **ImageAttributes** object from another **ImageAttributes** object.

Flat API function: **GdipCreateImageAttributes**.

```
CONSTRUCTOR CGpImageAttributes (BYVAL pImgAttr AS CGpImageAttributes PTR)
```
Flat API function: **GdipCloneImageAttributes**.

---

## <a name="clearbrushremaptable"></a>ClearBrushRemapTable (CGpImageAttributes)

Clears the brush color-remap table of this **ImageAttributes** object.

```
FUNCTION ClearBrushRemapTable () AS GpStatus
```

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipSetImageAttributesRemapTable**.

---

## <a name="clearcolorkey"></a>ClearColorKey (CGpImageAttributes)

Clears the color key (transparency range) for a specified category.

```
FUNCTION ClearColorKey (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify one color key for the default category, a different color key for the bitmap category, and still a different color key for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a default color key that makes any color with a red component from 200 through 255 transparent and you specify a default gamma value of 1.8. If you set the color key of the pen category by calling **SetColorKey**, then the default color key and the default gamma value will not apply to pens. If you later clear the pen color key by calling **ClearColorKey**, the pen category will not revert to the default color key; rather, the pen category will have no color key. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipSetImageAttributesColorKeys**.

---

## <a name="clearcolormatrices"></a>ClearColorMatrices (CGpImageAttributes)

Clears the color-adjustment matrix and the grayscale-adjustment matrix for a specified category.

```
FUNCTION ClearColorMatrices (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a pair (color and grayscale) of adjustment matrices for the default category, a different pair of adjustment matrices for the bitmap category, and still a different pair of adjustment matrices for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a pair (color and grayscale) of adjustment matrices and a gamma value for the default category. If you set a pair of adjustment matrices for the pen category by calling **SetColorMatrices**, then the default adjustment matrices will not apply to pens. If you later clear the pen adjustment matrices by calling **ClearColorMatrices**, the pen category will not revert to the default adjustment matrices; rather, the pen category will have no adjustment matrices. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

Flat API function: **GdipSetImageAttributesColorMatrix**.

---

## <a name="clearcolormatrix"></a>ClearColorMatrix (CGpImageAttributes)

Clears the color-adjustment matrix for a specified category.

```
FUNCTION ClearColorMatrix (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color-adjustment matrix for the default category, a different color-adjustment matrix for the bitmap category, and still a different color-adjustment matrix for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a color-adjustment matrix and a gamma value for the default category. If you set a color-adjustment matrix for the pen category by calling **SetColorMatrix**, then the default color-adjustment matrix will not apply to pens. If you later clear the pen color-adjustment matrix by calling **ClearColorMatrix**, the pen category will not revert to the default adjustment matrix; rather, the pen category will have no adjustment matrix. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

Flat API function: **GdipSetImageAttributesColorMatrix**.

---

## <a name="cleargamma"></a>ClearGamma (CGpImageAttributes)

Disables gamma correction for a specified category.

```
FUNCTION ClearGamma (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a gamma value for the default category, a different gamma value for the bitmap category, and still a different gamma value for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a gamma value and a color-adjustment matrix for the default category. If you set the gamma value for the pen category by calling **SetGamma**, then the default gamma value will not apply to pens. If you later clear the pen gamma value by calling **ClearGamma**, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value. Similarly, the pen category will not revert to the default color-adjustment matrix; rather, the pen category will have no color-adjustment matrix.

Flat API function: **GdipSetImageAttributesGamma**.

---

## <a name="clearnoop"></a>ClearNoOp (CGpImageAttributes)

Clears the **NoOp** setting for a specified category.

```
FUNCTION ClearNoOp (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

You can disable color adjustment for a certain object type by calling the **SetNoOp** method. Later, you can reinstate color adjustment for that object type by calling the **ClearNoOp** method. For example, the following statement disables color adjustment for brushes:

```
myImageAttributes.SetNoOp(ColorAdjustTypeBrush)
```

The following statement reinstates the brush color adjustment that was in place before the call to **SetNoOp**:

```
myImageAttributes.ClearNoOp(ColorAdjustTypeBrush)
```

Flat API function: **GdipSetImageAttributesNoOp**.

---

## <a name="clearoutputchannel"></a>ClearOutputChannel (CGpImageAttributes)

Clears the cyan-magenta-yellow-black (CMYK) output channel setting for a specified category.

```
FUNCTION ClearOutputChannel (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify an output channel for the default category and a different output channel for the bitmap category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the bitmap category, then the default settings apply to the bitmap category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify an output channel and an adjustment matrix for the default category. If you set the output channel for the bitmap category by calling **SetOutputChannel**, then the default output channel will not apply to bitmaps. If you later clear the bitmap output channel by calling **ClearOutputChannel**, the bitmap category will not revert to the default output channel; rather, the bitmap category will have no output channel setting. Similarly, the bitmap category will not revert to the default color-adjustment matrix; rather, the bitmap category will have no color-adjustment matrix.


Flat API function: **GdipSetImageAttributesOutputChannel**.

---

## <a name="clearoutputchannelcolorprofile"></a>ClearOutputChannelColorProfile (CGpImageAttributes)

Clears the output channel color profile setting for a specified category.

```
FUNCTION ClearOutputChannelColorProfile (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify an output channel profile for the default category and a different output channel profile for the bitmap category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the bitmap category, then the default settings apply to the bitmap category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify an output channel profile and an adjustment matrix for the default category. If you set the output channel profile for the bitmap category by calling **SetOutputChannelColorProfile**, then the default output channel profile will not apply to bitmaps. If you later clear the bitmap output channel profile by calling **ClearOutputChannelColorProfile**, the bitmap category will not revert to the default output channel profile; rather, the bitmap category will have no output channel profile setting. Similarly, the bitmap category will not revert to the default color-adjustment matrix; rather, the bitmap category will have no color-adjustment matrix.

Flat API function: **GdipSetImageAttributesOutputChannelColorProfile**.

---

## <a name="clearremaptable"></a>ClearRemapTable (CGpImageAttributes)

Clears the color-remap table for a specified category.

```
FUNCTION ClearRemapTable (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a remap table for the default category, a different remap table for the bitmap category, and still a different remap table for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a remap table and a gamma value for the default category. If you set the remap table for the pen category by calling **SetRemapTable**, then the default remap table will not apply to pens. If you later clear the pen remap table by calling **ClearRemapTable**, the pen category will not revert to the default remap table; rather, the pen category will have no remap table. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

Flat API function: **GdipSetImageAttributesRemapTable**.

---

## <a name="clearthreshold"></a>ClearThreshold (CGpImageAttributes)

Clears the threshold value for a specified category.

```
FUNCTION ClearThreshold (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

The threshold is a value from 0 through 1 that specifies a cutoff point for each color component. For example, suppose the threshold is set to 0.7, and suppose you are rendering a color whose red, green, and blue components are 230, 50, and 220. The red component, 230, is greater than 0.7×255, so the red component will be changed to 255 (full intensity). The green component, 50, is less than 0.7×255, so the green component will be changed to 0. The blue component, 220, is greater than 0.7×255, so the blue component will be changed to 255.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a threshold for the default category, a different threshold for the bitmap category, and still a different threshold for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a threshold and a gamma value for the default category. If you set the threshold for the pen category by calling **SetThreshold**, then the default threshold will not apply to pens. If you later clear the pen threshold by calling **ClearThreshold**, the pen category will not revert to the default threshold; rather, the pen category will have no threshold. Similarly, the pen category will not revert to the default gamma value; rather, the pen category will have no gamma value.

Flat API function: **GdipSetImageAttributesThreshold**.

---

## <a name="clone"></a>Clone (CGpImageAttributes)

Copies the contents of the existing ImageAttributes object into a new **ImageAttributes** object.

```
FUNCTION Clone (BYVAL pCloneImgAttr AS CGpImageAttributes PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pCloneImgAttr* | Pointer to the **ImageAttributes** object where to copy the contents of the existing object. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipCloneImageAttributes**.

---

## <a name="getadjustedpalette"></a>GetAdjustedPalette (CGpImageAttributes)

Adjusts the colors in a palette according to the adjustment settings of a specified category.

```
FUNCTION GetAdjustedPalette (BYVAL pColorPalette AS ColorPalette PTR, BYVAL colorAdjustType AS LONG) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pColorPalette* | Pointer to a **ColorPalette** structure that on input, contains the palette to be adjusted and, on output, receives the adjusted palette. |
| *colorAdjustType* | Element of the **ColorAdjustType** enumeration that specifies the category whose adjustment settings will be applied to the palette. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color-remap table for the default category, a different color-remap table for the bitmap category, and still a different color-remap table for the pen category.

When you call **GetAdjustedPalette**, you can specify the adjustment category that is used to adjust the palette colors. For example, if you pass **ColorAdjustTypeBitmap** to the **GetAdjustedPalette** method, then the adjustment settings of the bitmap category are used to adjust the palette colors.

Flat API function: **GdipGetImageAttributesAdjustedPalette**.

---

## <a name="reset"></a>Reset (CGpImageAttributes)

Clears all color- and grayscale-adjustment settings for a specified category.

```
FUNCTION Reset (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a gamma value for the default category, a different gamma value for the bitmap category, and still a different gamma value for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. You can reinstate the default settings for that category by calling **Reset**. For example, suppose you specify a gamma value for the default category. If you set the gamma value for the pen category by calling **SetGamma**, then the default gamma value will not apply to pens. If you later pass **ColorAdjustTypePen** to the **Reset** method, the pen category will revert to the default gamma value.

Flat API function: **GdipResetImageAttributes**.

---

## <a name="setbrushremaptable"></a>SetBrushRemapTable (CGpImageAttributes)

Sets the color remap table for the brush category.

```
FUNCTION SetBrushRemapTable (BYVAL mapSize AS UINT, BYVAL map AS ANY PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *mapSize* | The number of elements in the map array. |
| *map* | Pointer to an array of **ColorMap** structures. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

A color-remap table is an array of **ColorMap** structures. Each **ColorMap** structure has two colors: one that specifies an old color and one that specifies a corresponding new color. During rendering, any color that matches one of the old colors in the remap table is changed to the corresponding new color.

Calling the **SetBrushRemapTable** method has the same effect as passing **ColorAdjustTypeBrush** to the **SetRemapTable**. The specified remap table applies to items in metafiles that are filled with a brush.

Flat API function: **GdipSetImageAttributesRemapTable**.

---

## <a name="setcolorkey"></a>SetColorKey (CGpImageAttributes)

Sets the color key (transparency range) for a specified category.

```
FUNCTION SetColorKey (BYVAL colorLow AS ARGB, BYVAL colorHigh AS ARGB, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *colorLow* | ARGB color that specifies the low color-key value. |
| *colorHigh* | ARGB color that specifies the high color-key value. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

This method sets the high and low color-key values so that arange of colors can be made transparent. Any color that has each of its three components (red, green, blue) between the corresponding components of the high and low color keys is made transparent.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color key for the default category, a different color key for the bitmap category, and still a different color key for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the color key for the pen category by passing **ColorAdjustTypePen** to the **SetColorKey** method, then none of the default adjustment settings will apply to pens.

Flat API function: **GdipSetImageAttributesColorKeys**.

---

## <a name="setcolormatrices"></a>SetColorMatrices (CGpImageAttributes)

Sets the color-adjustment matrix and the grayscale-adjustment matrix for a specified category.

```
FUNCTION SetColorMatrices (BYVAL pColorMatrix AS ColorMatrix PTR, _
   BYVAL pGrayMatrix AS ColorMatrix PTR, _
   BYVAL nMode AS ColorAdjustType = ColorMatrixFlagsDefault, _
   BYVAL nType AS LONG = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pColorMatrix* | Pointer to a 5×5 color-adjustment matrix. |
| *pGrayMatrix* | Pointer to a 5×5 grayscale-adjustment matrix. |
| *nMode* | Optional. Element of the **ColorMatrixFlags** enumeration that specifies the type of image and color that will be affected by the color-adjustment and grayscale-adjustment matrices. The default value is **ColorMatrixFlagsDefault**. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify adjustment matrices for the default category, different adjustment matrices for the bitmap category, and still different adjustment matrices for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the color-adjustment and grayscale-adjustment matrices for the pen category by passing **ColorAdjustTypePen** to the **SetColorMatrices** method, then none of the default adjustment settings will apply to pens.

Flat API function: **GdipSetImageAttributesColorMatrix**.

---

## <a name="setcolormatrix"></a>SetColorMatrix (CGpImageAttributes)

Sets the color-adjustment matrix for a specified category.

```
FUNCTION SetColorMatrix (BYVAL pColorMatrix AS ColorMatrix PTR, _
   BYVAL nMode AS ColorMatrixFlags = ColorMatrixFlagsDefault, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pColorMatrix* | Pointer to a 5×5 color-adjustment matrix. |
| *nMode* | Optional. Element of the **ColorMatrixFlags** enumeration that specifies the type of image and color that will be affected by the color-adjustment and grayscale-adjustment matrices. The default value is **ColorMatrixFlagsDefault**. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color-adjustment matrix for the default category, a different color-adjustment matrix for the bitmap category, and still a different color-adjustment matrix for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the color-adjustment matrix for the pen category by passing **ColorAdjustTypePen** to the **SetColorMatrix** method, then none of the default adjustment settings will apply to pens.

Flat API function: **GdipSetImageAttributesColorMatrix**.

---

## <a name="setgamma"></a>SetGamma (CGpImageAttributes)

Sets the gamma value for a specified category.

```
FUNCTION SetGamma (BYVAL gamma AS SINGLE, BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *gamma* | The gamma value. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a gamma value for the default category, a different gamma value for the bitmap category, and still a different gamma value for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the gamma value for the pen category by passing **ColorAdjustTypePen** to the **SetGamma** method, then none of the default adjustment settings will apply to pens.

Flat API function: **GdipSetImageAttributesGamma**.

---

## <a name="setnoop"></a>SetNoOp (CGpImageAttributes)

Turns off color adjustment for a specified category. You can call **ClearNoOp** to reinstate the color-adjustment settings that were in place before the call to **SetNoOp**.

```
FUNCTION SetNoOp (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *gamma* | The gamma value. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipSetImageAttributesNoOp**.

---

## <a name="setoutputchannel"></a>SetOutputChannel (CGpImageAttributes)

Sets the CMYK output channel for a specified category.

```
FUNCTION SetOutputChannel (BYVAL channelFlags AS LONG, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *channelFlags* | Element of the **ColorChannelFlags** enumeration that specifies the output channel. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

You can use the **SetOutputChannel** method to convert an image to a cyan-magenta-yellow-black (CMYK) color space and examine the intensities of one of the CMYK color channels. For example, suppose you create an **ImageAttributes** object and set its bitmap output channel to **ColorChannelFlagsC**. If you pass the address of that **ImageAttributes** object to the **DrawImage** method, the cyan component of each pixel is calculated, and each pixel in the rendered image is a shade of gray that indicates the intensity of its cyan channel. Similarly, you can render images that indicate the intensities of the magenta, yellow, and black channels.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify an output channel for the default category and a different output channel for the bitmap category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the bitmap category, then the default settings apply to the bitmap category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the output channel for the bitmap category by passing **ColorAdjustTypeBitmap** to the **SetOutputChannel** method, then none of the default adjustment settings will apply to bitmaps.

Flat API function: **GdipSetImageAttributesOutputChannel**.

---

## <a name="setoutputchannelcolorprofile"></a>SetOutputChannelColorProfile (CGpImageAttributes)

Sets the output channel color-profile file for a specified category.

```
FUNCTION SetOutputChannelColorProfile ( BYVAL pwszColorProfileFilename AS WSTRING PTR, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszColorProfileFilename* | Path name of a color-profile file. If the color-profile file is in the %SystemRoot%\System32\Spool\Drivers\Color directory, then this parameter can be the file name. Otherwise, this parameter must be the fully-qualified path name. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

You can use the **SetOutputChannel** and **SetOutputChannelColorProfile** methods to convert an image to a cyan-magenta-yellow-black (CMYK) color space and examine the intensities of one of the CMYK color channels. For example, suppose you write code that performs the following steps:

* Create an **Image** object.
* Create an **ImageAttributes** object.
* Pass **ColorChannelFlagsC** to the **SetOutputChannel** method of the **ImageAttributes** object.
* Pass the path name of a color profile file to the **SetOutputChannelColorProfile** method of the **ImageAttributes** object.
* Pass the addresses of the Image and **ImageAttributes** objects to the **DrawImage** method.

Windows GDI+ will use the color-profile file to calculate the cyan component of each pixel in the image, and each pixel in the rendered image will be a shade of gray that indicates the intensity of its cyan channel.

An ImageAttributes object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify an output channel color-profile file for the default category and a different output channel color-profile file for the bitmap category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the bitmap category, then the default settings apply to the bitmap category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the output channel color-profile file for the bitmap category by passing **ColorAdjustTypeBitmap** to the **SetOutputChannelColorProfile** method, then none of the default adjustment settings will apply to bitmaps.

Flat API function: **GdipSetImageAttributesOutputChannelColorProfile**.

---

## <a name="setremaptable"></a>SetRemapTable (CGpImageAttributes)

Sets the color-remap table for a specified category.

```
FUNCTION SetRemapTable (BYVAL mapSize AS UINT, BYVAL map AS ANY PTR, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *mapSize* | The number of elements in the map array. |
| *map* | Pointer to an array of **ColorMap** structures that defines the color map. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

A color-remap table is an array of **ColorMap** structures. Each **ColorMap** structure has two colors: one that specifies an old color and one that specifies a corresponding new color. During rendering, any color that matches one of the old colors in the remap table is changed to the corresponding new color.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a color remap for the default category, a color-remap table for the bitmap category, and still a different color-remap table for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the color-remap table for the pen category by passing **ColorAdjustTypePen** to the **SetRemapTable** method, then none of the default adjustment settings will apply to pens.

Flat API function: **GdipSetImageAttributesRemapTable**.

---

## <a name="setthreshold"></a>SetThreshold (CGpImageAttributes)

Sets the threshold (transparency range) for a specified category.

```
FUNCTION SetThreshold (BYVAL threshold AS SINGLE, _
   BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *threshold* | The threshold value. |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

The threshold is a value from 0 through 1 that specifies a cutoff point for each color component. For example, suppose the threshold is set to 0.7, and suppose you are rendering a color whose red, green, and blue components are 230, 50, and 220. The red component, 230, is greater than 0.7×255, so the red component will be changed to 255 (full intensity). The green component, 50, is less than 0.7×255, so the green component will be changed to 0. The blue component, 220, is greater than 0.7×255, so the blue component will be changed to 255.

An **ImageAttributes** object maintains color and grayscale settings for five adjustment categories: default, bitmap, brush, pen, and text. For example, you can specify a threshold for the default category, a threshold for the bitmap category, and still a different threshold for the pen category.

The default color- and grayscale-adjustment settings apply to all categories that don't have adjustment settings of their own. For example, if you never specify any adjustment settings for the pen category, then the default settings apply to the pen category.

As soon as you specify a color- or grayscale-adjustment setting for a certain category, the default adjustment settings no longer apply to that category. For example, suppose you specify a collection of adjustment settings for the default category. If you set the threshold for the pen category by passing **ColorAdjustTypePen** to the **SetThreshold** method, then none of the default adjustment settings will apply to pens.

Flat API function: **GdipSetImageAttributesThreshold**.

---

# <a name="settoidentity"></a>SetToIdentity (CGpImageAttributes)

Sets the color-adjustment matrix of a specified category to identity matrix.

```
FUNCTION SetToIdentity (BYVAL nType AS ColorAdjustType = ColorAdjustTypeDefault) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Optional. Element of the **ColorAdjustType** enumeration that specifies the category for which the color key is cleared. The default value is **ColorAdjustTypeDefault**. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipSetImageAttributesToIdentity**.

---

## <a name="setwrapmode"></a>SetWrapMode (CGpImageAttributes)

Sets the wrap mode and color used to decide how to tile a texture across a shape, or at shape boundaries. A texture is tiled across a shape to fill it in when the texture is smaller than the shape it is filling.

```
FUNCTION SetWrapMode (BYVAL nWrap AS WrapMode, BYVAL colour AS ARGB = &HFF000000, _
   BYVAL clamp AS BOOL = FALSE) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWrap* | Element of the **WrapMode** enumeration that specifies how repeated copies of an image are used to tile an area. |
| *colour* | Optional. ARGB color that specifies the color of pixels outside of a rendered image. This color is visible if the wrap parameter is set to **WrapModeClamp** and the source rectangle passed to **DrawImage** is larger than the image itself. The default color is black. |
| *clamp* | Optional. This parameter has no effect in GDI+ version 1.0. Set this parameter to FALSE. The default value is FALSE. |

#### Return value

If the function succeeds, it returns **Ok**, which is an element of the **GpStatus** enumeration.

If the function fails, it returns one of the other elements of the **GpStatus** enumeration.

Flat API function: **GdipSetImageAttributesWrapMode**.

---
