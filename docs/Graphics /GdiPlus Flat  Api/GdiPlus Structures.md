# GDI+ Structures

Structures used by the GDI+ Flat API functions. Some of them have been prefixed with "Gp" to avoid name conflicts with other structures with the same name.

| Name       | Description |
| ---------- | ----------- |
| [BitmapData](#bitmapdata) | Stores attributes of a **Bbitmap** object. |
| [BlurParams](#blurparams) | Contains members that specify the nature of a Gaussian blur. |
| [BrightnessContrastParams](#brightnesscontrastparams) | Contains members that specify a the nature of a brightness or contrast adjustment. |
| [ColorBalanceParams](#colorbalanceparams) | Contains members that specify the nature of a color balance adjustment. |
| [ColorLUTParams](#colorlutparams) | Contains members (color lookup tables) that specify color adjustments to a bitmap. |
| [ColorMatrix](#colormatrix) | Contains a 5x5 matrix of real numbers. |
| [ColorPalette](#colorpalette) | Defines an array of colors that make up a color palette. |
| [EncoderParameter](#encoderparameter) | Holds a parameter that can be passed to an image encoder. |
| [EncoderParameters](#encoderparameters) | An array of **EncoderParameter** structures along with a data member that specifies the number of **EncoderParameter** structures in the array. |
| [ENHMETAHEADER](#enhmetaheader3 ) | Contains enhanced-metafile data including the dimensions of the metafile image, the number of records in the metafile, and the resolution of the device on which the metafile was created. |
| [GdiplusStartupInput](#gdiplusstartupinput) | Holds a block of arguments that are required by the GdiplusStartup function. |
| [GdiplusStartupOutput](#gdiplusstartupoutput) | Used by the **GdiplusStartup** function to return (in its output parameter) a pointer to a hook function and a pointer to an unhook function.  |
| [GpCharacterRange](#gpcharacterrange) | Specifies a range of character positions within a string. |
| [GpColorMap](#gpcolormap) | Contains two COLORREF values. |
| [GpPathData](#gppathdata) | Used to get or set the data points (and their types) of a path. |
| [GpPointF](#gppointf) | Encapsulates a point in a 2-D coordinate system. |
| [GpRectF](#gprectf) | Defines a rectangle. |
| [GpSizeF](#gpsizef) | Encapsulates a Width and Height dimension in a 2-D coordinate system. |
| [HueSaturationLightnessParams](#huesaturationlightnessparams) | Contains members that specify hue, saturation and lightness adjustments to a bitmap. |
| [ImageCodecInfo](#imagecodecinfo) | Stores information about an image codec (encoder/decoder). |
| [LevelsParams](#levelparams) | Contains members that specify adjustments to the light, midtone, or dark areas of a bitmap. |
| [MetafileHeader](#metafileheader) | Stores properties of an associated metafile. |
| [PropertyItem](#propertyitem) | Holds one piece of image metadata. |
| [PWMFRect16](#pwmfrect16) | Defines a rectangle that bounds a placeable metafile. |
| [RedEyeCorrectionParams](#redeyecorrectionparams) | Contains members that specify the areas of a bitmap to which a red-eye correction is applied. |
| [SharpenParams](#sharpenparams) | Contains members that specify the nature of a sharpening adjustment to a bitmap. |
| [TintParams](#tintparams) | Contains members that specify the nature of a tint adjustment to a bitmap. |
| [WmfPlaceableFileHeader](#wmfplaceablefilemeader) | Defines the fields of a placeable metafile header. |

---

## BitmapData

Stores attributes of a **Bitmap** objec. The **BitmapData** structure is used by the **GdpiBitmapLockBits** and **GdpiBitmapUnlockBits** functions.

| Member | Description |
| ------ | ----------- |
| **Width** | UINT. Number of pixels in one scan line of the bitmap. |
| **Height** | UINT. Number of scan lines in the bitmap. |
| **Stride** | INT.Offset, in bytes, between consecutive scan lines of the bitmap. If the stride is positive, the bitmap is top-down. If the stride is negative, the bitmap is bottom-up. |
| **PixelFormat** | INT. The pixel format of the bitmap. |
| **Scan0** | ANY PTR. Pointer to the first (index 0) scan line of the bitmap. |
| **Reserved** | UINT_PTR. Reserved for future use. |

---

## BlurParams

A **BlurParams** structure contains members that specify the nature of a Gaussian blur.

You can apply a Gaussian blur effect to a bitmap by following these steps.

1. Create and initialize a **BlurParams** structure.
2. Pass the address of the **BlurParams** structure to the **GdipSetEffectParameters** function.

| Member | Description |
| ------ | ----------- |
| **radius** | REAL. Real number that specifies the blur radius (the radius of the Gaussian convolution kernel) in pixels. The radius must be in the range 0 through 255. As the radius increases, the resulting bitmap becomes more blurry. |
| **expandEdge** | BOOL. Boolean value that specifies whether the bitmap expands by an amount equal to the blur radius. If TRUE, the bitmap expands by an amount equal to the radius so that it can have soft edges. If FALSE, the bitmap remains the same size and the soft edges are clipped. |

**Remarks**: Minimum availability: GDI+ 1.1

---

## BrightnessContrastParams

A **BrightnessContrastParams** structure contains members that specify a the nature of a brightness or contrast adjustment.

You can change the brightness or contrast (or both) of a bitmap by following these steps.

1. Create and initialize a **BrightnessContrastParams** structure.
2. Pass the address of the **BrightnessContrastParams** structure to the **GdipSetEffectParameters** function.

| Member | Description |
| ------ | ----------- |
| **brightnessLevel** | INT. Integer in the range -255 through 255 that specifies the brightness level. If the value is 0, the brightness remains the same. As the value moves from 0 to 255, the brightness of the image increases. As the value moves from 0 to -255, the brightness of the image decreases. |
| **contrastLevel** | INT. Integer in the range -100 through 100 that specifies the contrast level. If the value is 0, the contrast remains the same. As the value moves from 0 to 100, the contrast of the image increases. As the value moves from 0 to -100, the contrast of the image decreases. |

**Remarks**: Minimum availability: GDI+ 1.1

---

## ColorBalanceParams

A **ColorBalanceParams** structure contains members that specify the nature of a color balance adjustment.

You can change the color balance of a bitmap by following these steps.

1. Create and initialize a **ColorBalanceParams** structure.
2. Pass the address of the **ColorBalanceParams** structure to the **GdipSetEffectParameters** function.

| Member | Description |
| ------ | ----------- |
| **cyanRed** | INT. Integer in the range -100 through 100 that specifies a change in the amount of red in the image. If the value is 0, there is no change. As the value moves from 0 to 100, the amount of red in the image increases and the amount of cyan decreases. As the value moves from 0 to -100, the amount of red in the image decreases and the amount of cyan increases. |
| **magentaGreen** | INT. Integer in the range -100 through 100 that specifies a change in the amount of green in the image. If the value is 0, there is no change. As the value moves from 0 to 100, the amount of green in the image increases and the amount of magenta decreases. As the value moves from 0 to -100, the amount of green in the image decreases and the amount of magenta increases. |
| **yellowBlue** | INT. Integer in the range -100 through 100 that specifies a change in the amount of blue in the image. If the value is 0, there is no change. As the value moves from 0 to 100, the amount of blue in the image increases and the amount of yellow decreases. As the value moves from 0 to -100, the amount of blue in the image decreases and the amount of yellow increases. |

---

## ColorLUTParams

A **ColorLUTParams** structure contains members (color lookup tables) that specify color adjustments to a bitmap.

You can apply a custom adjustment to a bitmap by following these steps.

1. Create a ColorLUTParams structure.
2. Each member of the ColorLUTParams structure is a color lookup table (array of 256 bytes) for a particular color channel, alpha, red, green, or blue. Assign values of your choice to the four lookup tables.
3. Pass the address of the ColorLUTParams structure to the ColorLUT::SetParameters method of a ColorLUT object.
4. Pass the address of the ColorLUT object to the Graphics::DrawImage method or to the Bitmap::ApplyEffect method.

| Member | Description |
| ------ | ----------- |
| **lutB** (0 TO 255) | Array of 256 bytes that specifies the adjustment for the blue channel. |
| **lutG** (0 TO 255) | Array of 256 bytes that specifies the adjustment for the green channel. |
| **lutR** (0 TO 255) | Array of 256 bytes that specifies the adjustment for the red channel. |
| **lutA** (0 TO 255) | Array of 256 bytes that specifies the adjustment for the alpha channel. |

#### Remarks

A lookup table specifies how existing color channel values should be replaced by new values. A color channel value of j is replaced by the jth entry in the lookup table for that channel. For example, an existing blue channel value of 25 would be replaced by the value of lutB(25).

---

##  ColorMatrix

The **ColorMatrix** structure contains a 5x5 matrix of real numbers. Several of the **ImageAttributes** functions adjust image colors by using a color matrix. Alias: **GpColorMatrix**.

```
m(0 to 4, 0 to 4) as REAL
```

Use an array (0 TO 4, 0 TO 4) AS SINGLE, and pass @array(0, 0)

#### Remarks

A 5x5 color matrix is a homogeneous matrix for a 4-space transformation. The element in the fifth row and fifth column of a 5x5 homogeneous matrix must be 1, and all of the other elements in the fifth column must be 0. Color matrices are used to transform color vectors. The first four components of a color vector hold the red, green, blue, and alpha components (in that order) of a color. The fifth component of a color vector is always 1.

---

## ColorPalette

The **ColorPalette** structure defines an array of colors that make up a color palette. The colors are 32-bit ARGB colors.

| Member | Description |
| ------ | ----------- |
| **Flags** | UINT. Combination of flags from the **PaletteFlags** enumeration. |
| **Count** | UINT. Number of elements in the **Entries** array. |
| **Entries** | UINT. Variable length array of ARGB colors. |

To ease the use of this structure, **GpColorPalette** defines **Entries** as an array (2 TO 255) of ARGB colors.

---

## EncoderParameter

An **EncoderParameter** structure holds a parameter that can be passed to an image encoder. An **EncoderParameter** structure can also be used to receive a list of possible values supported by a particular parameter of a particular image encoder.

| Member | Description |
| ------ | ----------- |
| **Guid** | GUID. Identifies the parameter category. GUIDs that represent various parameter categories (*EncoderCompression*, *EncoderColorDepth*, and the like) are defined in "GdiPlus constamts". |
| **NumberOfValues** | ULONG. Number of values in the array pointed to by the Value data member. |
| **Type** | ULONG. Identifies the data type of the parameter. The **EncoderParameterValueType** enumeration defines several possible value types. |
| **Value** | ANY PTR. Pointer to an array of values. Each value has the type specified by the **Type** data member.. |

```
' // Encoder parameter GUIDs
Dim Shared EncoderChrominanceTable AS GUID = Type(&hF2E455DC, &h09B3, &h4316, {&h82, &h60, &h67, &h6A, &hDA, &h32, &h48, &h1C})
Dim Shared EncoderColorDepth       AS GUID = Type(&h66087055, &hAD66, &h4C7C, {&h9A, &h18, &h38, &hA2, &h31, &h0B, &h83, &h37})
Dim Shared EncoderColorSpace       AS GUID = Type(&hAE7A62A0, &hEE2C, &h49D8, {&h9D, &h07, &h16, &hAF, &h4E, &h4D, &hBD, &h1E})
Dim Shared EncoderCompression      AS GUID = Type(&hECCBFEFF, &h4F77, &h4D61, {&h8A, &hB9, &h6A, &h51, &hF6, &hD3, &h8E, &h8B})
Dim Shared EncoderImageItems       AS GUID = Type(&h63875E13, &h1F1D, &h45AB, {&h91, &h95, &hCF, &h0F, &h5A, &h0C, &hBE, &h1B})
Dim Shared EncoderLuminanceTable   AS GUID = Type(&hEDB33BCE, &h0266, &h4A77, {&hB9, &h04, &h27, &h21, &h60, &h99, &hE7, &h17})
Dim Shared EncoderQuality          AS GUID = Type(&h1D5BE4B5, &hFA4A, &h452D, {&h9C, &hDD, &h5D, &hB3, &h51, &h05, &hE7, &hEB})
Dim Shared EncoderRenderMethod     AS GUID = Type(&h6DB9FCAA, &h0E84, &h485E, {&h9A, &h8D, &h1E, &h8B, &h12, &h1F, &h0C, &hC1})
Dim Shared EncoderSaveAsCMYK       AS GUID = Type(&hC617F8E1, &h1E5F, &h4B1F, {&hA6, &hD3, &h5E, &h1D, &h73, &h6C, &h6A, &h1D})
Dim Shared EncoderSaveFlag         AS GUID = Type(&h292266FC, &hAC40, &h47BF, {&h8C, &hFC, &hA8, &h5B, &h89, &hA6, &h55, &hDE})
Dim Shared EncoderScanMethod       AS GUID = Type(&h3A4E2661, &h3109, &h4E56, {&h85, &h1E, &h9F, &h7E, &h42, &h9F, &h1A, &h22})
Dim Shared EncoderTransformation   AS GUID = Type(&h8D0EB2D1, &h6A48, &h4A94, {&hAC, &h61, &h52, &h95, &hC7, &h2E, &h8E, &h33})
Dim Shared EncoderVersion          AS GUID = Type(&h24D18C76, &h814A, &h41A4, {&hBF, &h53, &h1C, &h21, &h9C, &hCC, &hF7, &h97})
```
---

## EncoderParameters

An **EncoderParameters** structure is an array of **EncoderParameter** structures along with a data member that specifies the number of **EncoderParameter** structures in the array.

| Member | Description |
| ------ | ----------- |
| **Count** | UINT. Number of **EncoderParameter** structures in the array. |
| **Parameter** | Array of **EncoderParameter** structures. |

---

## ENHMETAHEADER

The **ENHMETAHEADER** structure contains enhanced-metafile data such as the dimensions of the picture stored in the enhanced metafile, the count of records in the enhanced metafile, the resolution of the device on which the picture was created, and so on.

This structure is always the first record in an enhanced metafile.

[ENHMETAHEADER structure] (https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-enhmetaheader)

---

## GdiplusStartupInput

The **GdiplusStartupInput** structure holds a block of arguments that are required by the **GdiplusStartup** function.

| Member | Description |
| ------ | ----------- |
| **GdiplusVersion** | UINT32. Specifies the version of Microsoft Windows GDI+. Must be 1. |
| **DebugEventProc** | Pointer to a callback function that GDI+ can call, on debug builds, for assertions and warnings. The default value is NULL. |
| **SuppressBackgroundThread** | BOOL. Boolean value that specifies whether to suppress the GDI+ background thread. If you set this member to TRUE, **GdiplusStartup** returns (in its output parameter) a pointer to a hook function and a pointer to an unhook function. You must call those functions appropriately to replace the background thread. If you do not want to be responsible for calling the hook and unhook functions, set this member to FALSE. The default value is FALSE. |
| **SuppressExternalCodecs** | BOOL. Boolean value that specifies whether you want GDI+ to suppress external image codecs. GDI+ version 1.0 does not support external image codecs, so this parameter is ignored. |

#### Remarks

The **GdiplusStartupInput** structure provides a constructor that sets the **GdiplusVersion** member to 1 and allows you to specify values for the other three members. All of the constructor parameters are optional, so you can declare a variable of type **GdiplusStartupInput** without passing any arguments to the constructor, and all of the members will be initialized with appropriate default values.

If you set the **SuppressBackgroundThread** member of the **GdiplusStartupinput** parameter to TRUE, you must call the hook and unhook functions returned in the output parameter. Call those functions before and after the application's main message loop; that is, a message loop that is active for the lifetime of GDI+. Call the hook function before the loop starts, and call the unhook function after the loop ends.

---

## GdiplusStartupOutput

The **GdiplusStartup** function uses the **GdiplusStartupOutput** structure to return (in its output parameter) a pointer to a hook function and a pointer to an unhook function. If you set the **SuppressBackgroundThread** member of the input parameter to TRUE, then you are responsible for calling those functions to replace the Microsoft Windows GDI+ background thread.

Call the hook and unhook functions before and after the application's main message loop; that is, a message loop that is active for the lifetime of GDI+. Call the hook function before the loop starts, and call the unhook function after the loop ends. The token parameter of the hook function receives an identifier that you should later pass to the unhook function. If you do not pass the proper identifier (the one returned by the hook function) to the unhook function, there will be resource leaks that won't be cleaned up until the process exits.

If you do not want to be responsible for calling the hook and unhook functions, set the **SuppressBackgroundThread** member of the input parameter (passed to **GdiplusStartup**) to FALSE.

| Member | Description |
| ------ | ----------- |
| **NotificationHook** | Receives a pointer to a hook function. |
| **NotificationUnhook** | Receives a pointer to a unhook function. |

---

## GpCharacterRange

A **CharacterRange** structure specifies a range of character positions within a string.

| Member | Description |
| ------ | ----------- |
| **First** | INT. Specifies the first position of this range. |
| **Length** | INT. Specifies the number of positions in this range. |

#### Remarks

A character range is a range of character positions within a string of text. The area of the display that is occupied by a group of characters that are specified by the character range is the bounding region. A character range is set by **GdipSetStringFormatMeasurableCharacterRanges**. The number of ranges that are currently set can be determined by calling **GdipGetStringFormatMeasurableCharacterRangeCount**. This number is also the number of regions expected to be obtained by the **GdipMeasureCharacterRanges** function.

---

## GpColorMap

The **ColorMap** structure contains two **COLORREF** values. Several of the **ImageAttributes** functions adjust image colors by using a color remap table, which is an array of **ColorMap** structures.

Renamed as **GpColorMp** because there is a GDI **ColorMap** structure with different member names.

| Member | Description |
| ------ | ----------- |
| **oldColor** | COLORREF. The original color. |
| **newColor** | COLORREF. The new color. |

---

## GpPathData

The **PathData** structure is used by the **GdipGetPathData** function. It has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points. You can use a **PathData** structure to get or set the data points (and their types) of a path.

| Member | Description |
| ------ | ----------- |
| **Count** | INT. Number of elements in the *Points* array. |
| **Points** | GpPointF PTR. Pointer to an array of **GpPointF** objects that contains the path points. |
| **Types** | BYTE PTR. pointer to an array of bytes that contains the point types. |

---

## GpPointF

Encapsulates a point in a 2-D coordinate system.

| Member | Description |
| ------ | ----------- |
| **x** | REAL. Horizontal coordinate. |
| **y** | REAL. Vertical coordinate. |

---

## GpRectF

Stores the upper-left corner, width, and height of a rectangle.

| Member | Description |
| ------ | ----------- |
| **x** | REAL. Horizontal coordinate. |
| **y** | REAL. Vertical coordinate. |
| **Width** | REAL. The width of the rectangle. |
| **Heighjt* | REAL. The height of the rectangle. |

---

## GpSizeF

The **GpSizeF** structure encapsulates a **Width** and **Height** dimension in a 2-D coordinate system.

| Member | Description |
| ------ | ----------- |
| **Width** | REAL. The width measurement. |
| **Height** | REAL. The height measurement. |

---

## HueSaturationLightnessParams

The **HueSaturationLightnessParams** structure contains members that specify hue, saturation and lightness adjustments to a bitmap.

You can adjust the hue, saturation, and lightness of a bitmap by following these steps.

1. Create and initialize a **HueSaturationLightnessParams** structure.
2. Pass the address of the **HueSaturationLightnessParams** structure to the **GdipSetEffectParameters** function.

| Member | Description |
| ------ | ----------- |
| **hueLevel** | INT. Integer in the range -180 through 180 that specifies the change in hue. A value of 0 specifies no change. Positive values specify counterclockwise rotation on the color wheel. Negative values specify clockwise rotation on the color wheel. |
| **saturationLevel** | INT. Integer in the range -100 through 100 that specifies the change in saturation. A value of 0 specifies no change. Positive values specify increased saturation and negative values specify decreased saturation. |
| **lightnessLevel** | INT. Integer in the range -100 through 100 that specifies the change in lightness. A value of 0 specifies no change. Positive values specify increased lightness and negative values specify decreased lightness. |

**Remarks**: Minimum availability: GDI+ 1.1

---

## ImageCodecInfo

An **ImageCodecInfo** structure stores information about an image codec (encoder/decoder). GDI+ provides several built-in image codecs. You can obtain information about those codecs by calling the **GetImageEncoders** function and the **GetImageDecoders** function. Each of those functions returns an array of **ImageCodecInfo** structures, one for each available encoder or decoder.

| Member | Description |
| ------ | ----------- |
| **Clsid** | CLSID. Codec identifier. |
| **FormatID** | GUID. File format identifier. GUIDs that identify various file formats (*ImageFormatBMP*, *ImageFormatEMF*, and the like) are defined in "GdiPlus constamts". |
| **CodecName** | WSTRING PTR. Pointer to a null-terminated wide string that contains the codec name. |
| **DllName** | WSTRING PTR. Pointer to a null-terminated wide string that contains the path name of the DLL in which the codec resides. If the codec is not in a DLL, this pointer is NULL. |
| **FormatDescription** | WSTRING PTR. Pointer to a null-terminated wide string that contains the name of the file format used by the codec. |
| **FilenameExtension** | WSTRING PTR. Pointer to a null-terminated wide string that contains all file-name extensions associated with the codec. The extensions are separated by semicolons. |
| **MimeType** | WSTRING PTR. Pointer to a null-terminated wide string that contains the mime type of the codec. |
| **Flags** | DWORD. Combination of flags from the **ImageCodecFlags** enumeration. |
| **Version** | DWORD. Integer that indicates the version of the codec. |
| **SigCount** | DWORD. Integer that indicates the number of signatures used by the file format associated with the codec. |
| **SigSize** | DWORD. Integer that indicates the number of bytes in each signature. |
| **SigPattern** | BYTE PTR. Pointer to an array of bytes that contains the pattern for each signature. |
| **SigMask** | BYTE PTR. Pointer to an array of bytes that contains the mask for each signature. |

---

## LevelParams

The LevelsParams structure contains members that specify adjustments to the light, midtone, or dark areas of a bitmap.

You can adjust the light, midtone, or dark areas of a bitmap by following these steps.

1. Create and initialize a **LevelsParams** structure.
2. Pass the address of the **LevelsParams** structure to the **GdipSetEffectParameters** function.

| Member | Description |
| ------ | ----------- |
| **highlight** | INT. Integer in the range 0 through 100 that specifies which pixels should be lightened. You can use this adjustment to lighten pixels that are already lighter than a certain threshold. Setting highlight to 100 specifies no change. Setting **highlight** to t specifies that a color channel value is increased if it is already greater than t percent of full intensity. For example, setting **highlight** to 90 specifies that all color channel values greater than 90 percent of full intensity are increased. |
| **midtone** | INT. Integer in the range -100 through 100 that specifies how much to lighten or darken an image. Color channel values in the middle of the intensity range are altered more than color channel values near the minimum or maximum intensity. You can use this adjustment to lighten (or darken) an image without loosing the contrast between the darkest and lightest portions of the image. A value of 0 specifies no change. Positive values specify that the midtones are made lighter, and negative values specify that the midtones are made darker. |
| **shadow** | INT. Integer in the range 0 through 100 that specifies which pixels should be darkened. You can use this adjustment to darken pixels that are already darker than a certain threshold. Setting **shadow** to 0 specifies no change. Setting **shadow** to t specifies that a color channel value is decreased if it is already less than t percent of full intensity. For example, setting **shadow** to 10 specifies that all color channel values less than 10 percent of full intensity are decreased. |

**Remarks**: Minimum availability: GDI+ 1.1

---

## MetafileHeader

A **MetafileHeader** object stores properties of an associated metafile.

| Member | Description |
| ------ | ----------- |
| **Type** | UINT. The type of metafile (e.g., WMF, EMF, EMF+). |
| **Size** | UINT. The total size of the metafile in bytes. |
| **Version** | UINT. The metafile's version. |
| **EmfPlusFlags** | UINT. An unsigned integer that contains information about how this metafile was recorded. |
| **DpiX**/**DpiY** | UINT. The resolution of the metafile in pixels per inch, for a rasterized image.  |
| **DpiY** | UINT. The vertical resolution of the metafile in pixels per inch, for a rasterized image  |
| **X**/**Y**/**Width**/**Height** | INT. The bounding box of the metafile in device units (pixels). |
| **X**/**Y**/**Width**/**Height** | INT. The bounding box of the metafile in device units (pixels). |
| **EmfPlusHeaderSize** | INT. The size of the EMF+ header, if present. |
| **WmfHeader**/**EmfHeader** | Umion of **METAHEADER** (for WMF files) and **ENHMETAHEADER3** (for EMF files) structures |
| **LogicalDpiX**/**LogicalDpiY** | INT. The logical DPI of the reference device context, which can be used to convert between device and logical units. |

#### Metafile type

| Constant | Description |
| -------- | ----------- |
| **MetafileTypeInvalid** | Specifies a metafile format that is not recognized in GDI+. |
| **MetafileTypeWmf** |Specifies a WMF file. Such a file contains only GDI records. |
| **MetafileTypeWmfPlaceable** |Specifies a WMF file that has a placeable metafile header in front of it. |
| **MetafileTypeEmf** | Specifies an EMF file. Such a file contains only GDI records. |
| **MetafileTypeEmfPlusOnly** | Specifies an EMF+ file. Such a file contains only GDI+ records and must be displayed by using GDI+. Displaying the records using GDI may cause unpredictable results. |
| **MetafileTypeEmfPlusDual** | Specifies an EMF+ Dual file. Such a file contains GDI+ records along with alternative GDI records and can be displayed by using either GDI or GDI+. Displaying the records using GDI may cause some quality degradation. |

#### Remarks

Metafiles provide a device-independent and application-independent way to share pictures. They contain records that describe a sequence of graphics application programming interfaces (APIs) to invoke in a particular order with their associated graphics data. Metafiles can be recorded by an application and later played back by that application or by another one to reproduce a particular picture. They can also be used to send content to a print spooler. Enhanced metafiles support the ability to provide both Microsoft Windows GDI+ and Windows Graphics Device Interface (GDI) descriptions of the same picture so that both GDI+ and down-level GDI applications can render it.

---

## PropertyItem

A PropertyItem structure holds one piece of image metadata.

| Member | Description |
| ------ | ----------- |
| **id** | PROPID. Integer that identifies the kind of metadata stored in this **PropertyItem** object. Constants that identify various kinds of metadata (*PropertyTagEquipMake*, *PropertyTagEquipModel*, and the like) are defined in "GdiPlus constamts". |
| **length** | ULONG. Size, in bytes, of the value array. |
| **type** | WORD. Integer that identifies the data type of the values in the value array. Constants that identify various data types (*PropertyTagTypeByte*, *PropertyTagTypeASCII*, and the like) are defined in "GdiPlus constamts". |
| **value** | ANY PTR. Pointer to an array of values. Each value in the array has the data type specified by the type data member. |

---

## PWMFRect16

The **PWMFRect16** structure defines a rectangle that bounds a placeable metafile. The rectangle defines the size and position of the metafile.

| Member | Description |
| ------ | ----------- |
| **Left** | INT16. Left coordinate of the rectangle that bounds the metafile. |
| **Top** | INT16. Top coordinate of the rectangle that bounds the metafile. |
| **Right** | INT16. Right coordinate of the rectangle that bounds the metafile. |
| **Bottom** | INT16. Bottom coordinate of the rectangle that bounds the metafile. |

---

## RedEyeCorrectionParams

A **RedEyeCorrectionParams** structure contains members that specify the areas of a bitmap to which a red-eye correction is applied.

You can can correct red eyes in a bitmap by following these steps.

1. Create and initialize a **RedEyeCorrectionParams** structure.
2. Pass the address of the **RedEyeCorrectionParams** structure to the **GdipSetEffectParameters** function.

| Member | Description |
| ------ | ----------- |
| **numberOfAreas** | UINT. Integer that specifies the number of **RECT** structures in the areas array. |
| **areas** | RECT PTR. Pointer to an array of **RECT** structures, each of which specifies an area of the bitmap to which red eye correction should be applied. |

**Remarks**: Minimum availability: GDI+ 1.1

---

## SharpenParams

The **SharpenParams** structure contains members that specify the nature of a sharpening adjustment to a bitmap.

You can adjust the sharpness of a bitmap by following these steps.

1. Create and initialize a **SharpenParams** structure.
2. Pass the address of the **SharpenParams** structure to the **GdipSetEffectParameters** function.

| Member | Description |
| ------ | ----------- |
| **radius** | REAL. Real number that specifies the sharpening radius (the radius of the convolution kernel) in pixels. The radius must be in the range 0 through 255. As the radius increases, more surrounding pixels are involved in calculating the new value of a given pixel. |
| **amount** | REAL. Real number in the range 0 through 100 that specifies the **amount** of sharpening to be applied. A value of 0 specifies no sharpening. As the value of amount increases, the sharpness increases. |

**Remarks**: Minimum availability: GDI+ 1.1

---

## TintParams

A **TintParams** structure contains members that specify the nature of a tint adjustment to a bitmap.

You can adjust the tint of a bitmap by following these steps.

1. Create and initialize a **TintParams** structure.

2. Pass the address of the **TintParams** structure to the **GdipSetEffectParameters** function.

| Member | Description |
| ------ | ----------- |
| **hue** | INT. Integer in the range -180 through 180 that specifies the hue to be strengthened or weakened. A value of 0 specifies blue. A positive value specifies a clockwise angle on the color wheel. For example, positive 60 specifies cyan and positive 120 specifies green. A negative value specifies a counter-clockwise angle on the color wheel. For example, negative 60 specifies magenta and negative 120 specifies red. |
| **amount** | INT. Integer in the range -100 through 100 that specifies how much the hue (given by the hue parameter) is strengthened or weakened. A value of 0 specifies no change. Positive values specify that the **hue** is strengthened and negative values specify that the hue is weakened. |

**Remarks**: Minimum availability: GDI+ 1.1

---

## WmfPlaceableFileHeader

The **WmfPlaceableFileHeader** structure defines the fields of a placeable metafile header. Placeable metafiles were created as a way of specifying how a metafile is mapped and scaled on a display device.

| Member | Description |
| ------ | ----------- |
| **Key** | UINT32. Identification value that indicates the presence of a placeable metafile header. This value is always &h9AC6CDD7. |
| **Hmf** | INT16. Handle to the metafile in memory. When written to disk, this field is not used and will always contains the value 0. |
| **BoundingBox** | PWMFRect16. Destination rectangle, measured in twips, for displaying the metafile. |
| **Inch** | INT16. Number of twips per inch used to represent the image. Normally, there are 1440 twips per inch; however, this number can be changed to scale the image.<br>A value of 720 specifies that the image is twice its normal size.<br>A value of 360 specifies that the image is four times its normal size.<br>A value of 2880 specifies that the image is half its normal size. |
| **Reserved** | UINT32. Not used and is always set to 0. |
| **Checksum** | INT16. Checksum for the previous 10 WORDs in the header. This value can be used to determine whether the metafile has become corrupted. |

#### Remarks

Although placeable metafiles are quite common, they are not directly supported by the Microsoft Windows application programming interface (API). To display a placeable metafile using the Windows API, you must first strip the placeable metafile header from the file. This is typically performed by copying the metafile to a temporary file starting at file offset 22 (0x16). This is because each placeable metafile begins with a 22-byte header that is followed by a standard metafile.

---
