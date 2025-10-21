| Name       | Description |
| ---------- | ----------- |
| [BrushType](#brushtype) | Indicates the type of brush. |
| [Color](#color) | Predefined ARGB colors. Represent an ARGB (alpha, red, green, blue) color. |
| [ColorAdjustType](#coloradjusttype) | Specifies which GDI+ objects use color-adjustment information. You can adjust the colors in a rendered image using the **ImageAttributes** functions. There are five adjustment categories: default, bitmap, brush, pen, and text. |
| [ColorChannelFlags](#colorchannelflags) | Specifies individual channels in the CMYK (cyan, magenta, yellow, black) color space. This enumeration is used by the **GdipSetImageAttributesOutputChannel** function. |
| [ColorMatrixFlags](#colormatrixflags) | Specifies the types of images and colors that will be affected by the color and grayscale adjustment settings. It is used by the **GdipSetImageAttributesColorMatrix** function. |
| [ColorMode](#colormode) | Specifies two modes for color component values. |
| [CombineMode](#combinemode) | Specifies how a new region is combined with an existing region. |
| [CompositingMode](#compositingmode) | Specifies how rendered colors are combined with background colors. This enumeration is used by the **GdipGetCompositingMode** and **GdipSetCompositingMode** functions. |
| [CompositingQuality](#compositingquality) | Specifies whether gamma correction is applied when colors are blended with background colors. This enumeration is used by the **GdipGetCompositingQuality** and **GdipSetCompositingQuality** functions. |
| [ConvertToEmfPlusFlags](#converttoemfplusflags) | Conversion of Emf To Emf+ bits flags. |
| [CoordinateSpace](#coordinatespace) | Specifies coordinate spaces. This enumeration is used by the **GdipTransformPoints** and **GdipTransformPointsI** functions, which converts points from one coordinate space to another. |
| [CurveAdjustments](#curveadjustments) | Bitmap adjustments. |
| [CurveChannel](#curvechannel) | Specifies which color channels are affected by a **ColorCurve** bitmap adjustment. |
| [CustomLineCapType](#customlinecaptype) | Specifies the type of custom line cap. |
| [DashCap](#dashcap) | Specifies the type of graphic shape to use on both ends of each dash in a dashed line. |
| [DashStyle](#dashstyle) | Specifies the line style of a line drawn with a Microsoft Windows GDI+ pen. The line can be drawn by using one of several predefined styles or a custom style. |
| [DitherType](#dithertype) | Identifies the available algorithms for dithering when a bitmap is converted. |
| [DriverStringOptions](#driverstringoptions) | Specifies the spacing, orientation, and quality of the rendering for driver strings. |
| [EmfPlusRecordType](#emfplusrecordtype) | Identifies metafile record types used in Microsoft Windows Metafile Format (WMF), Enhanced Metafile (EMF), and EMF+ files. |
| [EmfToWmfBitsFlags](#emftowmfbitsflags) | Specifies options for the **GdipEmfToWmfBits** method, which converts an Enhanced Metafile (EMF) metafile to a Microsoft Windows Metafile Format (WMF) metafile. |
| [EmfType](#EmfType) | Specifies the nature of the records that are placed in an Enhanced Metafile (EMF) file. |
| [EncoderParameterValueType](#encoderparametervaluetype) | Specifies data types for image codec (encoder/decoder) parameters. An element of this enumeration is assigned to the Type data member of an **EncoderParameter** structure. |
| [EncoderValue](#encodervalue) | Specifies values that can be passed as arguments to image encoders. |
| [FillMode](#fillmode) | Specifies how to fill areas that are formed when a path or curve intersects itself. |
| [FlushIntention](#flushintention) | Specifies when to flush the queue of graphics operations. |
| [FontStyle](#fontstyle) | Specifies the style of the typeface of a font. Styles can be combined. |
| [GenericFontFamily](#genericfontfamily) | Specifies the type of generic font family.  |
| [HatchStyle](#hatchstyle) | Specifies the hatch pattern used by a brush of type **HatchBrush**. The hatch pattern consists of a solid background color and lines drawn over the background. |
| [HistogramFormat](#histogramformat) | Specifies  the number and type of histograms that represent the color channels of a bitmap. This enumeration is used with the **GdipBitmapGetHistogram** function. |
| [HotKeyPrefix](#hotkeyprefix) | Specifies how to display hot keys. There are three options: do nothing, display hot keys underlined, and hide the hot key underlines. |
| [ImageCodecFlags](#imagecodecflags) | Indicates attributes of an image codec. |
| [ImageFlags](#imageflags) | Specifies the attributes of the pixel data contained in an image. The **GdipGetImageFlags** function returns an element of this enumeration. |
| [ImageLockMode](#imagelockmode) | Specifies flags that are passed to the flags parameter of the **GdipBitmapLockBits** function. The GdipBitmapLockBits function locks a portion of an image so that you can read or write the pixel data. |
| [ImageType](#imagetype) | Indicates whether an image is a bitmap or a metafile. The **GdipGetImageType** function returns an element of this enumeration. |
| [InterpolationMode](#interpolationmode) | Specifies the algorithm that is used when images are scaled or rotated. This enumeration is used by the **GdipGetInterpolationMode** and **GdipSetInterpolationMode** functions of the **Graphics** functions. |
| [ItemDataPosition](#itemdataposition) | Used to specify the location of custom metadata in an image file. |
| [LinearGradientMode](#lineargradientmode) | Specifies the direction in which the change of color occurs for a linear gradient brush. |
| [LineCap](#linecap) | Specifies the type of graphic shape to use on the end of a line drawn with a Microsoft Windows GDI+ pen. The cap can be a square, circle, triangle, arrowhead, custom, or masked (hidden). End caps can also "anchor" the line by centering the cap at the end of the line. |
| [LineJoin](#linejoin) | Specifies how to join two lines that are drawn by the same pen and whose ends meet. At the intersection of the two line ends, a line join makes the join look more continuous. |
| [MatrixOrder](#matrixorder) | Specifies the order of multiplication when a new matrix is multiplied by an existing matrix. |
| [MetafileFrameUnit](#metafileframeunit) | Specifies the order of multiplication when a new matrix is multiplied by an existing matrix. |
| [MetafileType](#metafiletype) | Specifies types of metafiles. |
| [ObjectType](#objecttype) | Indicates the object type value of an EMF+ record. |
| [PaletteFlags](#paletteflags) | Indicates attributes of the color data in a palette. |
| [PaletteType](#palettetypr) | Used by the **GdipInitializePalette** and **GdipBitmapConvertFormat** functions. The members of the enumeration identify several standard color palette formats. |
| [PathPointType](#pathpointtype) | Indicates  point types and flags for the data points in a path. Bits 0 through 2 indicate the type of a point, and bits 3 through 7 hold a set of flags that specify attributes of a point. |
| [PenAlignment](#penalignment) | Specifies the alignment of a pen relative to the stroke that is being drawn. |
| [PenType](#pentype) | Indicates the type of pattern, texture, or gradient that a pen draws. |
| [PixelOffsetMode](#pixeloffsetmode) | Specifies the pixel offset mode. This enumeration is used by the **GdipGetPixelOffsetMode** and **GdipSetPixelOffsetMode** methods of the **Graphics** class. |
| [QualityMode](#qualitymode) | Specifies the quality mode. |
| [RotateFlipType](#rotatefliptype) | Specifies the direction of an image's rotation and the axis used to flip the image. |
| [SmoothingMode](#smoothingmode) | Specifies the type of smoothing (antialiasing) that is applied to lines and curves. This enumeration is used by the **GdipGetSmoothingMode** and **GdipSetSmoothingMode** functions. |
| [Status](#status) | Indicates the result of a Microsoft Windows GDI+ method call. |
| [StringAlignment](#stringalignment) | Specifies how a string is aligned in reference to the bounding rectangle. A bounding rectangle is used to define the area in which the text displays. |
| [StringDigitSubstitute](#stringdigitsubstitute) | Specifies how to substitute digits in a string according to a user's locale or language. |
| [StringFormatFlags](#stringformatflags) | Specifies text layout information (such as orientation and clipping) and display manipulations (such as ellipsis insertion, digit substitution, and representation of characters that are not supported by a font). |
| [StringTrimming](#stringtrimming) | Specifies how to trim characters from a string so that the string fits into a layout rectangle. The layout rectangle is used to position and size the display string. |
| [TestControlEnum](#testcontrolenum) | Used for internal testing of GDI+. Specifies changes in image-processing behavior. |
| [TextRenderingHint](#Textrenderinghint) | Specifies the process used to render text. The process affects the quality of the text. |
| [Unit](#unit) | Specifies the unit of measure for a given data type. |
| [WarpMode](#warpmode) | Specifies warp modes that can be used to transform images. |
| [WrapMode](#wrapmode) | Specifies how repeated copies of an image are used to tile an area. |

---

## BrushType

The **BrushType** enumeration indicates the type of brush. There are five types of brushes.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *BrushTypeSolidColor* | 0 | Indicates a brush of type SolidBrush. A solid brush paints a single, constant color that can be opaque or transparent. |
| *BrushTypeHatchFill* | 1 | Indicates a brush of type **HatchBrush**. A hatch brush paints a background and paints, over that background, a pattern of lines, dots, dashes, squares, crosshatch, or some variation of these. The hatch brush consists of two colors: one for the background and one for the pattern over the background. The color of the background is called the background color, and the color of the pattern is called the foreground color. |
| *BrushTypeTextureFill* | 2 | Indicates a brush of type **TextureBrush**. A texture brush paints an image. The image or texture is either a portion of a specified image or a scaled version of a specified image. The type of image (metafile or nonmetafile) determines whether the texture is a portion of the image or a scaled version of the image. |
| *BrushTypePathGradient* | 3 | Indicates a brush of type PathGradientBrush. A path gradient brush paints a color gradient in which the color changes from a center point outward to a boundary that is defined by a closed curve or path. The color gradient has one color at the center point and one or multiple colors at the boundary. |
| *BrushTypeLinearGradient* | 4 | Indicates a brush of type LinearGradientBrush. A linear gradient brush paints a color gradient in which the color changes evenly from the starting boundary line of the linear gradient brush to the ending boundary line of the linear gradient brush. The boundary lines of a linear gradient brush are two parallel straight lines. The color gradient is perpendicular to the boundary lines of the linear gradient brush, changing gradually across the stroke from the starting boundary line to the ending boundary line. The color gradient has one color at the starting boundary line and another color at the ending boundary line. |

---

## Color

Predefined ARGB colors. Represent an ARGB (alpha, red, green, blue) color.

| Name  | Value |
| ----- | ----- |
| ARGB_AliceBlue  | &hFFF0F8FF |
| ARGB_AntiqueWhite  | &hFFFAEBD7 |
| ARGB_Aqua  | &hFF00FFFF |
| ARGB_Aquamarine  | &hFF7FFFD4 |
| ARGB_Azure  | &hFFF0FFFF |
| ARGB_Beige  | &hFFF5F5DC |
| ARGB_Bisque  | &hFFFFE4C4 |
| ARGB_Black  | &hFF000000 |
| ARGB_BlanchedAlmond  | &hFFFFEBCD |
| ARGB_Blue  | &hFF0000FF |
| ARGB_BlueViolet  | &hFF8A2BE2 |
| ARGB_Brown  | &hFFA52A2A |
| ARGB_BurlyWood  | &hFFDEB887 |
| ARGB_CadetBlue  | &hFF5F9EA0 |
| ARGB_Chartreuse  | &hFF7FFF00 |
| ARGB_Chocolate  | &hFFD2691E |
| ARGB_Coral  | &hFFFF7F50 |
| ARGB_CornflowerBlue  | &hFF6495ED |
| ARGB_Cornsilk  | &hFFFFF8DC |
| ARGB_Crimson  | &hFFDC143C |
| ARGB_Cyan  | &hFF00FFFF |
| ARGB_DarkBlue  | &hFF00008B |
| ARGB_DarkCyan  | &hFF008B8B |
| ARGB_DarkGoldenrod  | &hFFB8860B |
| ARGB_DarkGray  | &hFFA9A9A9 |
| ARGB_DarkGreen  | &hFF006400 |
| ARGB_DarkKhaki  | &hFFBDB76B |
| ARGB_DarkMagenta  | &hFF8B008B |
| ARGB_DarkOliveGreen  | &hFF556B2F |
| ARGB_DarkOrange  | &hFFFF8C00 |
| ARGB_DarkOrchid  | &hFF9932CC |
| ARGB_DarkRed  | &hFF8B0000 |
| ARGB_DarkSalmon  | &hFFE9967A |
| ARGB_DarkSeaGreen  | &hFF8FBC8B |
| ARGB_DarkSlateBlue  | &hFF483D8B |
| ARGB_DarkSlateGray  | &hFF2F4F4F |
| ARGB_DarkTurquoise  | &hFF00CED1 |
| ARGB_DarkViolet  | &hFF9400D3 |
| ARGB_DeepPink  | &hFFFF1493 |
| ARGB_DeepSkyBlue  | &hFF00BFFF |
| ARGB_DimGray  | &hFF696969 |
| ARGB_DodgerBlue  | &hFF1E90FF |
| ARGB_Firebrick  | &hFFB22222 |
| ARGB_FloralWhite  | &hFFFFFAF0 |
| ARGB_ForestGreen  | &hFF228B22 |
| ARGB_Fuchsia  | &hFFFF00FF |
| ARGB_Gainsboro  | &hFFDCDCDC |
| ARGB_GhostWhite  | &hFFF8F8FF |
| ARGB_Gold  | &hFFFFD700 |
| ARGB_Goldenrod  | &hFFDAA520 |
| ARGB_Gray  | &hFF808080 |
| ARGB_Green  | &hFF008000 |
| ARGB_GreenYellow  | &hFFADFF2F |
| ARGB_Honeydew  | &hFFF0FFF0 |
| ARGB_HotPink  | &hFFFF69B4 |
| ARGB_IndianRed  | &hFFCD5C5C |
| ARGB_Indigo  | &hFF4B0082 |
| ARGB_Ivory  | &hFFFFFFF0 |
| ARGB_Khaki  | &hFFF0E68C |
| ARGB_Lavender  | &hFFE6E6FA |
| ARGB_LavenderBlush  | &hFFFFF0F5 |
| ARGB_LawnGreen  | &hFF7CFC00 |
| ARGB_LemonChiffon  | &hFFFFFACD |
| ARGB_LightBlue  | &hFFADD8E6 |
| ARGB_LightCoral  | &hFFF08080 |
| ARGB_LightCyan  | &hFFE0FFFF |
| ARGB_LightGoldenrodYellow  | &hFFFAFAD2 |
| ARGB_LightGray  | &hFFD3D3D3 |
| ARGB_LightGreen  | &hFF90EE90 |
| ARGB_LightPink  | &hFFFFB6C1 |
| ARGB_LightSalmon  | &hFFFFA07A |
| ARGB_LightSeaGreen  | &hFF20B2AA |
| ARGB_LightSkyBlue  | &hFF87CEFA |
| ARGB_LightSlateGray  | &hFF778899 |
| ARGB_LightSteelBlue  | &hFFB0C4DE |
| ARGB_LightYellow  | &hFFFFFFE0 |
| ARGB_Lime  | &hFF00FF00 |
| ARGB_LimeGreen  | &hFF32CD32 |
| ARGB_Linen  | &hFFFAF0E6 |
| ARGB_Magenta  | &hFFFF00FF |
| ARGB_Maroon  | &hFF800000 |
| ARGB_MediumAquamarine  | &hFF66CDAA |
| ARGB_MediumBlue  | &hFF0000CD |
| ARGB_MediumOrchid  | &hFFBA55D3 |
| ARGB_MediumPurple  | &hFF9370DB |
| ARGB_MediumSeaGreen  | &hFF3CB371 |
| ARGB_MediumSlateBlue  | &hFF7B68EE |
| ARGB_MediumSpringGreen  | &hFF00FA9A |
| ARGB_MediumTurquoise  | &hFF48D1CC |
| ARGB_MediumVioletRed  | &hFFC71585 |
| ARGB_MidnightBlue  | &hFF191970 |
| ARGB_MintCream  | &hFFF5FFFA |
| ARGB_MistyRose  | &hFFFFE4E1 |
| ARGB_Moccasin  | &hFFFFE4B5 |
| ARGB_NavajoWhite  | &hFFFFDEAD |
| ARGB_Navy  | &hFF000080 |
| ARGB_OldLace  | &hFFFDF5E6 |
| ARGB_Olive  | &hFF808000 |
| ARGB_OliveDrab  | &hFF6B8E23 |
| ARGB_Orange  | &hFFFFA500 |
| ARGB_OrangeRed  | &hFFFF4500 |
| ARGB_Orchid  | &hFFDA70D6 |
| ARGB_PaleGoldenrod  | &hFFEEE8AA |
| ARGB_PaleGreen  | &hFF98FB98 |
| ARGB_PaleTurquoise  | &hFFAFEEEE |
| ARGB_PaleVioletRed  | &hFFDB7093 |
| ARGB_PapayaWhip  | &hFFFFEFD5 |
| ARGB_PeachPuff  | &hFFFFDAB9 |
| ARGB_Peru  | &hFFCD853F |
| ARGB_Pink  | &hFFFFC0CB |
| ARGB_Plum  | &hFFDDA0DD |
| ARGB_PowderBlue  | &hFFB0E0E6 |
| ARGB_Purple  | &hFF800080 |
| ARGB_Red  | &hFFFF0000 |
| ARGB_RosyBrown  | &hFFBC8F8F |
| ARGB_RoyalBlue  | &hFF4169E1 |
| ARGB_SaddleBrown  | &hFF8B4513 |
| ARGB_Salmon  | &hFFFA8072 |
| ARGB_SandyBrown  | &hFFF4A460 |
| ARGB_SeaGreen  | &hFF2E8B57 |
| ARGB_SeaShell  | &hFFFFF5EE |
| ARGB_Sienna  | &hFFA0522D |
| ARGB_Silver  | &hFFC0C0C0 |
| ARGB_SkyBlue  | &hFF87CEEB |
| ARGB_SlateBlue  | &hFF6A5ACD |
| ARGB_SlateGray  | &hFF708090 |
| ARGB_Snow  | &hFFFFFAFA |
| ARGB_SpringGreen  | &hFF00FF7F |
| ARGB_SteelBlue  | &hFF4682B4 |
| ARGB_Tan_  | &hFFD2B48C |
| ARGB_Teal  | &hFF008080 |
| ARGB_Thistle  | &hFFD8BFD8 |
| ARGB_Tomato  | &hFFFF6347 |
| ARGB_Transparent_  | &h00FFFFFF |
| ARGB_Turquoise  | &hFF40E0D0 |
| ARGB_Violet  | &hFFEE82EE |
| ARGB_Wheat  | &hFFF5DEB3 |
| ARGB_White  | &hFFFFFFFF |
| ARGB_WhiteSmoke  | &hFFF5F5F5 |
| ARGB_Yellow  | &hFFFFFF00 |
| ARGB_YellowGreen  | &hFF9ACD32 |

---

## ColorAdjustType

The **ColorAdjustType** enumeration specifies which GDI+ objects use color-adjustment information. You can adjust the colors in a rendered image using the ImageAttributes functions. There are five adjustment categories: default, bitmap, brush, pen, and text.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *ColorAdjustTypeDefault* | 0 | Specifies that color or grayscale adjustment applies to all categories that do not have adjustment settings of their own. |
| *ColorAdjustTypeBitmap* | 1 | Specifies that color or grayscale adjustment applies to bitmapped images. |
| *ColorAdjustTypeBrush* | 2 | Specifies that color or grayscale adjustment applies to brush operations in metafiles. |
| *ColorAdjustTypePen* | 3 | Specifies that color or grayscale adjustment applies to pen operations in metafiles. |
| *ColorAdjustTypeText* | 4 | Specifies that color or grayscale adjustment applies to text drawn in metafiles. |
| *ColorAdjustTypeCount* | 5 | Used internally to record the number of color adjustment types. |
| *ColorAdjustTypeAny* | 6 | Reserved. |

---

## ColorChannelFlags

Thee *ColorChannelFlags* enumeration specifies individual channels in the CMYK (cyan, magenta, yellow, black) color space. This enumeration is used by the **GdipSetImageAttributesOutputChannel** function.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *ColorChannelFlagsC* | 0 | Specifies the cyan color channel. |
| *ColorChannelFlagsM* | 1 | Specifies the magenta color channel. |
| *ColorChannelFlagsY* | 2 | Specifies the yellow color channel. |
| *ColorChannelFlagsK* | 3 | Specifies the black color channel. |
| *ColorChannelFlagsLast* | 4 |  |

---

## ColorMatrixFlags

The **ColorMatrixFlags** enumeration specifies the types of images and colors that will be affected by the color and grayscale adjustment settings. It is used by the **GdipSetImageAttributesColorMatrix** function.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *ColorMatrixFlagsDefault* | 0 | Specifies that all color values (including grays) are adjusted by the same color-adjustment matrix. |
| *ColorMatrixFlagsSkipGrays* | 1 | Specifies that colors are adjusted but gray shades are not adjusted. A gray shade is any color that has the same value for its red, green, and blue components. |
| *ColorMatrixFlagsAltGray* | 2 | Specifies that colors are adjusted by one matrix and gray shades are adjusted by another matrix. |

---

## ColorMode

Specifies two modes for color component values.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *ColorModeARGB32* | 0 | The integer values supplied are 32-bit values. |
| *ColorModeARGB64* | 0 | The integer values supplied are 64-bit values. |

#### Remarks

For computers with 32-bit integers, use Argb32Mode. For computers with 64-bit integers, use Argb64Mode.

---

## CombineMode

The **CombineMode** enumeration specifies how a new region is combined with an existing region.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *CombineModeReplace* | 0 | Specifies that the existing region is replaced by the new region. |
| *CombineModeIntersect* | 1 | Specifies that the existing region is replaced by the intersection of itself and the new region. |
| *CombineModeUnion* | 2 | Specifies that the existing region is replaced by the union of itself and the new region. |
| *CombineModeXor* | 3 | Specifies that the existing region is replaced by the result of performing an XOR on the two regions. A point is in the XOR of two regions if it is in one region or the other but not in both regions. |
| *CombineModeExclude* | 4 | Specifies that the existing region is replaced by the portion of itself that is outside of the new region. |
| *CombineModeComplement* | 5 | Specifies that the existing region is replaced by the portion of the new region that is outside of the existing region. |

---

## CompositingMode

The **CompositingMode** enumeration specifies how rendered colors are combined with background colors. This enumeration is used by the **GdipGetCompositingMode** and **GdipSetCompositingMode** functions.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *CompositingModeSourceOver* | 0 | Specifies that when a color is rendered, it is blended with the background color. The blend is determined by the alpha component of the color being rendered. |
| *CompositingModeSourceCopy* | 0 | Specifies that when a color is rendered, it overwrites the background color. This mode cannot be used along with *TextRenderingHintClearTypeGridFit*. |

---

## CompositingQuality

The **CompositingQuality** enumeration specifies whether gamma correction is applied when colors are blended with background colors. This enumeration is used by the **GdipGetCompositingQuality** and **GdipSetCompositingQuality** functions.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *CompositingQualityDefault* | 0 | Specifies that gamma correction is not applied. |
| *CompositingQualityHighSpeed* | 1 | Specifies that gamma correction is not applied. |
| *CompositingQualityHighQuality* | 2 | Specifies that gamma correction is applied. |
| *CompositingQualityGammaCorrected* | 3 | Specifies that gamma correction is applied. |
| *CompositingQualityAssumeLinear* | 4 | Specifies that gamma correction is not applied. |

---

## ConvertToEmfPlusFlags

Conversion of Emf To Emf+ bits flags.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *ConvertToEmfPlusFlagsDefault* | &h00000000 | Specifies the default conversion. |
| *ConvertToEmfPlusFlagsRopUsed* | &H00000001 | Specifies that the source EMF metafile is embedded as a comment in the resulting WMF metafile. |
| *ConvertToEmfPlusFlagsText* | &H00000002 | Specifies that the resulting WMF metafile is in the placeable metafile format; that is, it has the additional 22-byte header required by a placeable metafile. |
| *ConvertToEmfPlusFlagsInvalidRecord* | &H00000004 | Specifies that the clipping region is stored in the metafile in the traditional way. If you do not set this flag, the **GdipEmfToWmfBits** function applies an optimization that stores the clipping region as a path and simulates clipping by using the XOR operator. |

---

## CoordinateSpace

The **CoordinateSpace** enumeration specifies coordinate spaces. This enumeration is used by the **GdipTransformPoints** function, which converts points from one coordinate space to another. For more information about coordinate spaces, see [Types of Coordinate Systems](https://learn.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-types-of-coordinate-systems-about).

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *CoordinateSpaceWorld* | 0 | Specifies the world coordinate space. |
| *CoordinateSpacePage* | 1 | Specifies the page coordinate space. |
| *CoordinateSpaceDevice* | 2 | Specifies the device coordinate space. |

---

## CurveAdjustments

Bitmap adjustments-

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *AdjustExposure* | 0 | Simulates increasing or decreasing the exposure of a photograph. When you set the adjustment member of a **ColorCurveParams** object to **AdjustExposure**, you should set the **adjustValue** member to an integer in the range -255 through 255. A value of 0 specifies no change in exposure. Positive values specify increased exposure and negative values specify decreased exposure. |
| *AdjustDensity* | 1 | Simulates increasing or decreasing the film density of a photograph. When you set the adjustment member of a **ColorCurveParams** object to **AdjustDensity**, you should set the **adjustValue** member to an integer in the range -255 through 255. A value of 0 specifies no change in density. Positive values specify increased density (lighter picture) and negative values specify decreased density (darker picture). |
| *AdjustContrast* | 2 | Increases or decreases the contrast of a bitmap. When you set the adjustment member of a **ColorCurveParams** object to **AdjustContrast**, you should set the **adjustValue** member to an integer in the range -100 through 100. A value of 0 specifies no change in contrast. Positive values specify increased contrast and negative values specify decreased contrast. |
| *AdjustHighlight* | 3 | Increases or decreases the value of a color channel if that channel already has a value that is above half intensity. You can use this adjustment to get more definition in the light areas of an image without affecting the dark areas. When you set the adjustment member of a **ColorCurveParams** object to **AdjustHighlight**, you should set the **adjustValue** member to an integer in the range -100 through 100. A value of 0 specifies no change. Positive values specify that the light areas are made lighter, and negative values specify that the light areas are made darker. |
| *AdjustShadow* | 4 | Increases or decreases the value of a color channel if that channel already has a value that is below half intensity. You can use this adjustment to get more definition in the dark areas of an image without affecting the light areas. When you set the adjustment member of a **ColorCurveParams** object to **AdjustShadow**, you should set the adjustValue member to an integer in the range -100 through 100. A value of 0 specifies no change. Positive values specify that the dark areas are made lighter, and negative values specify that the dark areas are made darker. |
| *AdjustMidtone* | 5 | Lightens or darkens an image. Color channel values in the middle of the intensity range are altered more than color channel values near the minimum or maximum intensity. You can use this adjustment to lighten (or darken) an image without loosing the contrast between the darkest and lightest portions of the image. When you set the adjustment member of a **ColorCurveParams** object to **AdjustMidtone**, you should set the **adjustValue** member to an integer in the range -100 through 100. A value of 0 specifies no change. Positive values specify that the midtones are made lighter, and negative values specify that the midtones are made darker. |
| *AdjustWhiteSaturation* | 6 | When you set the adjustment member of a **ColorCurveParams** object to **AdjustWhiteSaturation**, you should set the **adjustValue** member to an integer in the range 0 through 255. A value of t specifies that the interval [0, t] is mapped linearly to the interval [0, 255]. For example, if **adjustValue** is equal to 240, then color channel values in the interval [0, 240] are adjusted so that they spread out over the interval [0, 255]. Color channel values greater than 240 are set to 255. |
| *AdjustBlackSaturation* | 7 | When you set the adjustment member of a **ColorCurveParams** object to **AdjustBlackSaturation**, you should set the **adjustValue** member to an integer in the range 0 through 255. A value of t specifies that the interval [t, 255] is mapped linearly to the interval [0, 255]. For example, if **adjustValue** is equal to 15, then color channel values in the interval [15, 255] are adjusted so that they spread out over the interval [0, 255]. Color channel values less than 15 are set to 0. |

To apply one of the eight adjustments to a bitmap, follow these steps.

1. Create a *ColorCurveParams* structure, and set its adjustment member to one of the elements of the **CurveAdjustments** enumeration.
2. Set the other two members (adjustValue and channel) of the **ColorCurveParams** structure.
3. Pass the address of the **ColorCurveParams** structure to the **GdipSetEffectParameters** function.
4. Call the **GdipBitmapApplyEffect** function.

---

## CurveChannel

The **CurveChannel** enumeration specifies which color channels are affected by a **ColorCurve** bitmap adjustment.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *CurveChannelAll* | 0 | Specifies that the color adjustment applies to all channels. |
| *CurveChannelRed* | 1 | Specifies that the color adjustment applies only to the red channel. |
| *CurveChannelGreen* | 2 | Specifies that the color adjustment applies only to the green channel. |
| *CurveChannelBlue* | 2 | Specifies that the color adjustment applies only to the blue channel. |

---

## CustomLineCapType

Specifies the type of custom line cap. You would typically use this value when creating a new **AdjustableArrowCap** object or when setting a custom line cap for a **Pen** object in GDI+

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *CustomLineCapTypeDefault* | 0 | Default custom line cap. |
| *CustomLineCapTypeAdjustableArrow* | 1 | Indicates that the line cap being defined is an adjustable arrow. |

---

## DashCap

Specifies the type of graphic shape to use on both ends of each dash in a dashed line.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *DashCapFlat* | 0 | Specifies a square cap that squares off both ends of each dash. |
| *DashCapRound* | 2 | Specifies a circular cap that rounds off both ends of each dash. |
| *DashCapTriangle* | 3 | Specifies a triangular cap that points both ends of each dash. |

---

## DashStyle

The **DashStyle** enumeration specifies the line style of a line drawn with a Windows GDI+ pen. The line can be drawn by using one of several predefined styles or a custom style.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *DashStyleSolid* | 0 | Specifies a solid line. |
| *DashStyleDash* | 1 | Specifies a dashed line. |
| *DashStyleDot* | 2 | Specifies a dotted line. |
| *DashStyleDashDot* | 3 | Specifies an alternating dash-dot line. |
| *DashStyleDashDotDot* | 4 | Specifies an alternating dash-dot-dot line. |
| *DashStyleCustom* | 5 | Specifies a user-defined, custom dashed line. |

#### Remarks

A custom dashed line is created by calling the **GdipSetPenDashArray** function, which takes an array of values for the dash lengths and the space lengths.

---

## DitherType

Identifies the available algorithms for dithering when a bitmap is converted.

Calling the **GdipBitmapConvertFormat** function of a **Bitmap** object changes the pixel format of that **Bitmap**. If the conversion results in a reduced bit depth (fewer bits per pixel), then certain colors in the original bitmap will be simulated by a dither (checkerboard) pattern made up of colors that are available in the new pixel format. The members of the **DitherType** enumeration identify the algorithms available for performing this dithering.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *DitherTypeNone* | 0 | No dithering is performed. Pixels in the source bitmap are mapped to the nearest color in the palette specified by the palette parameter of the **GdipBitmapConvertFormat** function. This algorithm can be used with any palette. If the palette specified by the palette parameter does not have one of the standard fixed formats listed in the PaletteType enumeration, pass PaletteTypeCustom to the palettetype parameter. |
| *DitherTypeSolid* | 1 | No dithering is performed. Pixels in the source bitmap are mapped to the nearest color in the palette specified by the palette parameter of the **GdipBitmapConvertFormat** function. This algorithm can be used with any palette. If the palette specified by the palette parameter does not have one of the standard fixed formats listed in the PaletteType enumeration, pass *PaletteTypeCustom* to the palettetype parameter. |
| *DitherTypeOrdered4x4* | 2 | You can use this algorithm to perform dithering based on the colors in one of the standard fixed palettes. You can also use this algorithm to convert a bitmap to a 16-bits-per-pixel format that has no palette. |
| *DitherTypeOrdered8x8* | 3 | Dithering is performed using the colors in one of the standard fixed palettes. |
| *DitherTypeOrdered16x16* | 4 | Dithering is performed using the colors in one of the standard fixed palettes. |
| *DitherTypeSpiral4x4* | 5 | Dithering is performed using the colors in one of the standard fixed palettes. |
| *DitherTypeSpiral8x8* | 6 | Dithering is performed using the colors in one of the standard fixed palettes. |
| *DitherTypeDualSpiral4x4* | 7 | Dithering is performed using the colors in one of the standard fixed palettes. |
| *DitherTypeDualSpiral8x8* | 8 | Dithering is performed using the colors in one of the standard fixed palettes. |
| *DitherTypeErrorDiffusion* | 9 | Dithering is performed based on the palette specified by the palette parameter of the **GdipBitmapConvertFormat** function. This algorithm can be used with any palette. If the palette specified by the palette parameter does not have one of the standard fixed formats listed in the **PaletteType** enumeration, pass *PaletteTypeCustom* to the palettetype parameter. |
| *DitherTypeMax* | 10 |  |

---

## DriverStringOptions

The **DriverStringOptions** enumeration specifies the spacing, orientation, and quality of the rendering for driver strings.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *DriverStringOptionsCmapLookup* | 1 | Specifies that the string array contains Unicode character values. If this flag is not set, each value in array is interpreted as an index to a font glyph that defines a character to be displayed. |
| *DriverStringOptionsVertical* | 2 | Specifies that the string is displayed vertically. |
| *DriverStringOptionsRealizedAdvance* | 4 | Specifies that the glyph positions are calculated from the position of the first glyph. If this flag is not set, the glyph positions are obtained from an array of coordinates. |
| *DriverStringOptionsLimitSubpixel* | 8 | Specifies that less memory should be used for cache of antialiased glyphs. This also produces lower quality. If this flag isn't set, more memory is used, but the quality is higher. |

---

## EmfPlusRecordType

The **EmfPlusRecordType** enumeration identifies metafile record types used in Windows Metafile Format (WMF), Enhanced Metafile (EMF), and EMF+ files. The elements of the **EmfPlusRecordType** enumeration come in three groups.

* Elements in the first group have the prefix WmfRecordType and identity WMF records.
* Elements in the second group have the prefix EmfRecordType and identify EMF records.
* Elements in the third group have the prefix EmfPlusRecordType and identify EMF+ records.

WMF and EMF records can be displayed by Windows GDI+ and by Windows Graphics Device Interface (GDI). EMF+ records can be displayed by GDI+ but not by GDI.

Elements that have the WmfRecordType prefix are analogous to constants (defined in Wingdi.h) that have the prefix META_. For example, the element **WmfRecordTypeSetBkColor** is analogous to the constant META_SETBKCOLOR. For more information about WMF files, see Windows-Format Metafiles.

Elements that have the EmfRecordType prefix are analogous to constants (defined in Wingdi.h) that have the prefix EMR_. For example, the element EmfRecordTypePolygon is analogous to the constant EMR_POLYGON. For more information about EMR constants, see EMR.

Elements that have the EmfPlusRecordType prefix are specific to GDI+. Most of those elements correspond to methods of the Graphics class. The remaining elements identify the header, the end of the file, and other sections of the metafile. The Constants section of this topic shows the correspondence between the EMF+ record types and the methods of the Graphics class.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *WmfRecordTypeSetBkColor* | 66049 |  |
| *WmfRecordTypeSetBkMode* | 65794 |  |
| *WmfRecordTypeSetMapMode* | 65795 |  |
| *WmfRecordTypeSetROP2* | 65796 |  |
| *WmfRecordTypeSetRelAbs* | 65797 |  |
| *WmfRecordTypeSetPolyFillMode* | 65798 |  |
| *WmfRecordTypeSetStretchBltMode* | 65799 |  |
| *WmfRecordTypeSetTextCharExtra* | 65800 |  |
| *WmfRecordTypeSetTextColor* | 66057 |  |
| *WmfRecordTypeSetTextJustification* | 66058 |  |
| *WmfRecordTypeSetWindowOrg* | 66059 |  |
| *WmfRecordTypeSetWindowExt* | 66060 |  |
| *WmfRecordTypeSetViewportOrg* | 66061 |  |
| *WmfRecordTypeSetViewportExt* | 66062 |  |
| *WmfRecordTypeOffsetWindowOrg* | 66063 |  |
| *WmfRecordTypeScaleWindowExt* | 66576 |  |
| *WmfRecordTypeOffsetViewportOrg* | 66065 |  |
| *WmfRecordTypeScaleViewportExt* | 66578 |  |
| *WmfRecordTypeLineTo* | 66067 |  |
| *WmfRecordTypeMoveTo* | 66068 |  |
| *WmfRecordTypeExcludeClipRect* | 66581 |  |
| *WmfRecordTypeIntersectClipRect* | 66582 |  |
| *WmfRecordTypeArc* | 67607 |  |
| *WmfRecordTypeEllipse* | 66584 |  |
| *WmfRecordTypeFloodFill* | 66585 |  |
| *WmfRecordTypePie* | 67610 |  |
| *WmfRecordTypeRectangle* | 66587 |  |
| *WmfRecordTypeRoundRect* | 67100 |  |
| *WmfRecordTypePatBlt* | 67101 |  |
| *WmfRecordTypeSaveDC* | 65566 |  |
| *WmfRecordTypeSetPixel* | 66591 |  |
| *WmfRecordTypeOffsetClipRgn* | 66080 |  |
| *WmfRecordTypeTextOut* | 66849 |  |
| *WmfRecordTypeBitBlt* | 67874 |  |
| *WmfRecordTypeStretchBlt* | 68387 |  |
| *WmfRecordTypePolygon* | 66340 |  |
| *WmfRecordTypePolyline* | 66341 |  |
| *WmfRecordTypeEscape* | 67110 |  |
| *WmfRecordTypeRestoreDC* | 65831 |  |
| *WmfRecordTypeFillRegion* | 66088 |  |
| *WmfRecordTypeFrameRegion* | 66601 |  |
| *WmfRecordTypeInvertRegion* | 65834 |  |
| *WmfRecordTypePaintRegion* | 65835 |  |
| *WmfRecordTypeSelectClipRegion* | 65836 |  |
| *WmfRecordTypeSelectObject* | 65837 |  |
| *WmfRecordTypeSetTextAlign* | 65838 |  |
| *WmfRecordTypeDrawText* | 67119 |  |
| *WmfRecordTypeChord* | 67632 |  |
| *WmfRecordTypeSetMapperFlags* | 66097 |  |
| *WmfRecordTypeExtTextOut* | 68146 |  |
| *WmfRecordTypeSetDIBToDev* | 68915 |  |
| *WmfRecordTypeSelectPalette* | 66100 |  |
| *WmfRecordTypeRealizePalette* | 65589 |  |
| *WmfRecordTypeAnimatePalette* | 66614 |  |
| *WmfRecordTypeSetPalEntries* | 65591 |  |
| *WmfRecordTypePolyPolygon* | 66872 |  |
| *WmfRecordTypeResizePalette* | 65849 |  |
| *WmfRecordTypeDIBBitBlt* | 67904 |  |
| *WmfRecordTypeDIBStretchBlt* | 68417 |  |
| *WmfRecordTypeDIBCreatePatternBrush* | 65858 |  |
| *WmfRecordTypeStretchDIB* | 69443 |  |
| *WmfRecordTypeExtFloodFill* | 66888 |  |
| *WmfRecordTypeSetLayout* | 65865 |  |
| *WmfRecordTypeResetDC* | 65868 |  |
| *WmfRecordTypeStartDoc* | 65869 |  |
| *WmfRecordTypeStartPage* | 65615 |  |
| *WmfRecordTypeEndPage* | 65616 |  |
| *WmfRecordTypeAbortDoc* | 65618 |  |
| *WmfRecordTypeEndDoc* | 65630 |  |
| *WmfRecordTypeDeleteObject* | 66032 |  |
| *WmfRecordTypeCreatePalette* | 65783 |  |
| *WmfRecordTypeCreateBrush* | 65784 |  |
| *WmfRecordTypeCreatePatternBrush* | 66041 |  |
| *WmfRecordTypeCreatePenIndirect* | 66298 |  |
| *WmfRecordTypeCreateFontIndirect* | 66299 |  |
| *WmfRecordTypeCreateBrushIndirect* | 66300 |  |
| *WmfRecordTypeCreateBitmapIndirect* | 66301 |  |
| *WmfRecordTypeCreateRegion* | 67327 |  |
| *EmfRecordTypeHeader* | 1 |  |
| *EmfRecordTypePolyBezier* | 2 |  |
| *EmfRecordTypePolygon* | 3 |  |
| *EmfRecordTypePolyline* | 4 |  |
| *EmfRecordTypePolyBezierTo* | 5 |  |
| *EmfRecordTypePolyLineTo* | 6 |  |
| *EmfRecordTypePolyPolyline* | 7 |  |
| *EmfRecordTypePolyPolygon* | 8 |  |
| *EmfRecordTypeSetWindowExtEx* | 9 |  |
| *EmfRecordTypeSetWindowOrgEx* | 10 |  |
| *EmfRecordTypeSetViewportExtEx* | 11 |  |
| *EmfRecordTypeSetViewportOrgEx* | 12 |  |
| *EmfRecordTypeSetBrushOrgEx* | 13 |  |
| *EmfRecordTypeEOF* | 14 |  |
| *EmfRecordTypeSetPixelV* | 15 |  |
| *EmfRecordTypeSetMapperFlags* | 16 |  |
| *EmfRecordTypeSetMapMode* | 17 |  |
| *EmfRecordTypeSetBkMode* | 18 |  |
| *EmfRecordTypeSetPolyFillMode* | 19 |  |
| *EmfRecordTypeSetROP2 | 20 |  |
| *EmfRecordTypeSetStretchBltMode* | 21 |  |
| *EmfRecordTypeSetTextAlign* | 22 |  |
| *EmfRecordTypeSetColorAdjustment* | 23 |  |
| *EmfRecordTypeSetTextColor* | 24 |  |
| *EmfRecordTypeSetBkColor* | 25 |  |
| *EmfRecordTypeOffsetClipRgn* | 26 |  |
| *EmfRecordTypeMoveToEx* | 27 |  |
| *EmfRecordTypeSetMetaRgn* | 28 |  |
| *EmfRecordTypeExcludeClipRect* | 29 |  |
| *EmfRecordTypeIntersectClipRect* | 30 |  |
| *EmfRecordTypeScaleViewportExtEx* | 31 |  |
| *EmfRecordTypeScaleWindowExtEx* | 32 |  |
| *EmfRecordTypeSaveDC* | 33 |  |
| *EmfRecordTypeRestoreDC* | 34 |  |
| *EmfRecordTypeSetWorldTransform* | 35 |  |
| *EmfRecordTypeModifyWorldTransform* | 36 |  |
| *EmfRecordTypeSelectObject* | 37 |  |
| *EmfRecordTypeCreatePen* | 38 |  |
| *EmfRecordTypeCreateBrushIndirect* | 39 |  |
| *EmfRecordTypeDeleteObject* | 40 |  |
| *EmfRecordTypeAngleArc* | 41 |  |
| *EmfRecordTypeEllipse* | 42 |  |
| *EmfRecordTypeRectangle* | 43 |  |
| *EmfRecordTypeRoundRect* | 44 |  |
| *EmfRecordTypeArc* | 45 |  |
| *EmfRecordTypeChord* | 46 |  |
| *EmfRecordTypePie* | 47 |  |
| *EmfRecordTypeSelectPalette* | 48 |  |
| *EmfRecordTypeCreatePalette* | 49 |  |
| *EmfRecordTypeSetPaletteEntries* | 50 |  |
| *EmfRecordTypeResizePalette* | 51 |  |
| *EmfRecordTypeRealizePalette* | 52 |  |
| *EmfRecordTypeExtFloodFill* | 53 |  |
| *EmfRecordTypeLineTo* | 54 |  |
| *EmfRecordTypeArcTo* | 55 |  |
| *EmfRecordTypePolyDraw* | 56 |  |
| *EmfRecordTypeSetArcDirection* | 57 |  |
| *EmfRecordTypeSetMiterLimit* | 58 |  |
| *EmfRecordTypeBeginPath* | 59 |  |
| *EmfRecordTypeEndPath* | 60 |  |
| *EmfRecordTypeCloseFigure* | 61 |  |
| *EmfRecordTypeFillPath* | 62 |  |
| *EmfRecordTypeStrokeAndFillPath* | 63 |  |
| *EmfRecordTypeStrokePath* | 64 |  |
| *EmfRecordTypeFlattenPath* | 65 |  |
| *EmfRecordTypeWidenPath* | 66 |  |
| *EmfRecordTypeSelectClipPath* | 67 |  |
| *EmfRecordTypeAbortPath* | 68 |  |
| *EmfRecordTypeReserved_069 | 69 |  |
| *EmfRecordTypeGdiComment* | 70 |  |
| *EmfRecordTypeFillRgn* | 71 |  |
| *EmfRecordTypeFrameRgn* | 72 |  |
| *EmfRecordTypeInvertRgn* | 73 |  |
| *EmfRecordTypePaintRgn* | 74 |  |
| *EmfRecordTypeExtSelectClipRgn* | 75 |  |
| *EmfRecordTypeBitBlt* | 76 |  |
| *EmfRecordTypeStretchBlt* | 77 |  |
| *EmfRecordTypeMaskBlt* | 78 |  |
| *EmfRecordTypePlgBlt* | 79 |  |
| *EmfRecordTypeSetDIBitsToDevice* | 80 |  |
| *EmfRecordTypeStretchDIBits* | 81 |  |
| *EmfRecordTypeExtCreateFontIndirect* | 82 |  |
| *EmfRecordTypeExtTextOutA* | 83 |  |
| *EmfRecordTypeExtTextOutW* | 84 |  |
| *EmfRecordTypePolyBezier16* | 85 |  |
| *EmfRecordTypePolygon16* | 86 |  |
| *EmfRecordTypePolyline16* | 87 |  |
| *EmfRecordTypePolyBezierTo16* | 88 |  |
| *EmfRecordTypePolylineTo16* | 89 |  |
| *EmfRecordTypePolyPolyline16* | 90 |  |
| *EmfRecordTypePolyPolygon16* | 91 |  |
| *EmfRecordTypePolyDraw16* | 92 |  |
| *EmfRecordTypeCreateMonoBrush* | 93 |  |
| *EmfRecordTypeCreateDIBPatternBrushPt* | 94 |  |
| *EmfRecordTypeExtCreatePen* | 95 |  |
| *EmfRecordTypePolyTextOutA* | 96 |  |
| *EmfRecordTypePolyTextOutW* | 97 |  |
| *EmfRecordTypeSetICMMode* | 98 |  |
| *EmfRecordTypeCreateColorSpace* | 99 |  |
| *EmfRecordTypeSetColorSpace* | 100 |  |
| *EmfRecordTypeDeleteColorSpace* | 101 |  |
| *EmfRecordTypeGLSRecord* | 102 |  |
| *EmfRecordTypeGLSBoundedRecord* | 103 |  |
| *EmfRecordTypePixelFormat* | 104 |  |
| *EmfRecordTypeDrawEscape* | 105 |  |
| *EmfRecordTypeExtEscape* | 106 |  |
| *EmfRecordTypeStartDoc* | 107 |  |
| *EmfRecordTypeSmallTextOut* | 108 |  |
| *EmfRecordTypeForceUFIMapping* | 109 |  |
| *EmfRecordTypeNamedEscape* | 110 |  |
| *EmfRecordTypeColorCorrectPalette* | 111 |  |
| *EmfRecordTypeSetICMProfileA* | 112 |  |
| *EmfRecordTypeSetICMProfileW* | 113 |  |
| *EmfRecordTypeAlphaBlend* | 114 |  |
| *EmfRecordTypeSetLayout* | 115 |  |
| *EmfRecordTypeTransparentBlt* | 116 |  |
| *EmfRecordTypeReserved_117 | 117 |  |
| *EmfRecordTypeGradientFill* | 118 |  |
| *EmfRecordTypeSetLinkedUFIs* | 119 |  |
| *EmfRecordTypeSetTextJustification* | 120 |  |
| *EmfRecordTypeColorMatchToTargetW* | 121 |  |
| *EmfRecordTypeCreateColorSpaceW* | 122 |  |
| *EmfRecordTypeMax* | 122 |  |
| *EmfRecordTypeMin* | 1 |  |
| *EmfPlusRecordTypeInvalid* | 16384 |  |
| *EmfPlusRecordTypeHeader* | 16385 | Identifies a record that is the EMF+ header. |
| *EmfPlusRecordTypeEndOfFile* | 16386 | Identifies a record that marks the last EMF+ record of a metafile. |
| *EmfPlusRecordTypeComment* | 16387 | Graphics::AddMetafileComment |
| *EmfPlusRecordTypeGetDC* | 16388 | Graphics::GetHDC |
| *EmfPlusRecordTypeMultiFormatStart* | 16389 | Identifies the start of a multiple-format block. |
| *EmfPlusRecordTypeMultiFormatSection* | 16390 | Identifies a section in a multiple-format block. Multiple-format records allow the same set of records to be represented in several formats. |
| *EmfPlusRecordTypeMultiFormatEnd* | 16391 | Identifies the end of a multiple-format block. |
| *EmfPlusRecordTypeObject* | 16392 |  |
| *EmfPlusRecordTypeClear* | 16393 | Graphics::Clear |
| *EmfPlusRecordTypeFillRects* | 16394 | FillRectangles Methods |
| *EmfPlusRecordTypeDrawRects* | 16395 | DrawRectangles Methods |
| *EmfPlusRecordTypeFillPolygon* | 16396 | FillPolygon Methods |
| *EmfPlusRecordTypeDrawLines* | 16397 | DrawLines Methods |
| *EmfPlusRecordTypeFillEllipse* | 16398 | FillEllipse Methods |
| *EmfPlusRecordTypeDrawEllipse* | 16399 | DrawEllipse Methods |
| *EmfPlusRecordTypeFillPie* | 16400 | FillPie Methods |
| *EmfPlusRecordTypeDrawPie* | 16401 | DrawPie Methods |
| *EmfPlusRecordTypeDrawArc* | 16402 | DrawArc Methods |
| *EmfPlusRecordTypeFillRegion* | 16403 | Graphics::FillRegion |
| *EmfPlusRecordTypeFillPath* | 16404 | Graphics::FillPath |
| *EmfPlusRecordTypeDrawPath* | 16405 | Graphics::DrawPath |
| *EmfPlusRecordTypeFillClosedCurve* | 16406 | FillClosedCurve Methods |
| *EmfPlusRecordTypeDrawClosedCurve* | 16407 | DrawClosedCurve Methods |
| *EmfPlusRecordTypeDrawCurve* | 16408 | DrawCurve Methods |
| *EmfPlusRecordTypeDrawBeziers* | 16409 | DrawBeziers Methods |
| *EmfPlusRecordTypeDrawImage* | 16410 | DrawImage Methods (all methods that do not receive an array of destination points) |
| *EmfPlusRecordTypeDrawImagePoints* | 16411 | DrawImage Methods (all methods that receive an array of destination points) |
| *EmfPlusRecordTypeDrawString* | 16412 | DrawString Methods |
| *EmfPlusRecordTypeSetRenderingOrigin* | 16413 | Graphics::SetRenderingOrigin |
| *EmfPlusRecordTypeSetAntiAliasMode* | 16414 | Graphics::SetSmoothingMode |
| *EmfPlusRecordTypeSetTextRenderingHint* | 16415 | Graphics::SetTextRenderingHint |
| *EmfPlusRecordTypeSetTextContrast* | 16416 | Graphics::SetTextContrast |
| *EmfPlusRecordTypeSetInterpolationMode* | 16417 | Graphics::SetInterpolationMode |
| *EmfPlusRecordTypeSetPixelOffsetMode* | 16418 | Graphics::SetPixelOffsetMode |
| *EmfPlusRecordTypeSetCompositingMode* | 16419 | Graphics::SetCompositingMode |
| *EmfPlusRecordTypeSetCompositingQuality* | 16420 | Graphics::SetCompositingQuality |
| *EmfPlusRecordTypeSave* | 16421 | Graphics::Save |
| *EmfPlusRecordTypeRestore* | 16422 | Graphics::Restore |
| *EmfPlusRecordTypeBeginContainer* | 16423 | Graphics::BeginContainer |
| *EmfPlusRecordTypeBeginContainerNoParams* | 16424 | Graphics::BeginContainer |
| *EmfPlusRecordTypeEndContainer* | 16425 | Graphics::EndContainer |
| *EmfPlusRecordTypeSetWorldTransform* | 16426 | Graphics::SetTransform |
| *EmfPlusRecordTypeResetWorldTransform* | 16427 | Graphics::ResetTransform |
| *EmfPlusRecordTypeMultiplyWorldTransform* | 16428 | Graphics::MultiplyTransform |
| *EmfPlusRecordTypeTranslateWorldTransform* | 16429 | Graphics::TranslateTransform |
| *EmfPlusRecordTypeScaleWorldTransform* | 16430 | Graphics::ScaleTransform |
| *EmfPlusRecordTypeRotateWorldTransform* | 16431 | Graphics::RotateTransform |
| *EmfPlusRecordTypeSetPageTransform* | 16432 | Graphics::SetPageScale and Graphics::SetPageUnit |
| *EmfPlusRecordTypeResetClip* | 16433 | Graphics::ResetClip |
| *EmfPlusRecordTypeSetClipRect* | 16434 | Graphics::SetClip and Graphics::SetClip |
| *EmfPlusRecordTypeSetClipPath* | 16435 | Graphics::SetClip |
| *EmfPlusRecordTypeSetClipRegion* | 16436 | Graphics::SetClip |
| *EmfPlusRecordTypeOffsetClip* | 16437 | TranslateClip Methods |
| *EmfPlusRecordTypeDrawDriverString* | 16438 | Graphics::DrawDriverString |
| *EmfPlusRecordTypeStrokeFillPath* | 16439 |  |
| *EmfPlusRecordTypeSerializableObject* | 16440 |  |
| *EmfPlusRecordTypeSetTSGraphics* | 16441 |  |
| *EmfPlusRecordTypeSetTSClip* | 16442 |  |

---

## EmfToWmfBitsFlags

Specifies options for the **GdipEmfToWmfBits** function, which converts an Enhanced Metafile (EMF) metafile to a Windows Metafile Format (WMF) metafile.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *EmfToWmfBitsFlagsDefault* | &H00000000 | Specifies the default conversion. |
| *EmfToWmfBitsFlagsEmbedEmf* | &H00000001 | Specifies that the source EMF metafile is embedded as a comment in the resulting WMF metafile. |
| *EmfToWmfBitsFlagsIncludePlaceable* | &H00000002 | Specifies that the resulting WMF metafile is in the placeable metafile format; that is, it has the additional 22-byte header required by a placeable metafile. |
| *EmfToWmfBitsFlagsNoXORClip* | &H00000004 | Specifies that the clipping region is stored in the metafile in the traditional way. If you do not set this flag, the **GdipEmfToWmfBits** function applies an optimization that stores the clipping region as a path and simulates clipping by using the XOR operator. |

---

## EmfType

The **EmfType** enumeration specifies the nature of the records that are placed in an Enhanced Metafile (EMF) file.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *EmfTypeEmfOnly* | 3 | Specifies that all of the records in the metafile are EMF records, which can be displayed by GDI or GDI+. |
| *EmfTypeEmfPlusOnly* | 4 | Specifies that all of the records in the metafile are EMF+ records, which can be displayed by GDI+ but not by GDI. |
| *EmfTypeEmfPlusDual* | 5 | Specifies that all EMF+ records in the metafile are associated with an alternate EMF record. Metafiles of type **EmfTypeEmfPlusDual** can be displayed by GDI or by GDI+. |

---

## EncoderParameterValueType

Specifies data types for image codec (encoder/decoder) parameters. An element of this enumeration is assigned to the Type data member of an **EncoderParameter** structure

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *EncoderParameterValueTypeByte* | 1 | Specifies that the parameter is an 8-bit unsigned integer. |
| *EncoderParameterValueTypeASCII* | 2 | Specifies that the parameter is a null-terminated character string. |
| *EncoderParameterValueTypeShort* | 3 | Specifies that the parameter is a 16-bit unsigned integer. |
| *EncoderParameterValueTypeShort* | 4 | Specifies that the parameter is a 32-bit unsigned integer. |
| *EncoderParameterValueTypeRational* | 5 | Specifies that the parameter is an array of two, 32-bit unsigned integers. The pair of integers represents a fraction. The first integer in the pair is the numerator, and the second integer in the pair is the denominator. |
| *EncoderParameterValueTypeLongRange* | 6 | Specifies that the parameter is an array of two, 32-bit unsigned integers. The pair of integers represents a range of numbers. The first integer is the smallest number in the range, and the second integer is the largest number in the range. |
| *EncoderParameterValueTypeUndefined* | 7 | Specifies that the parameter is an array of bytes that can hold values of any type. |
| *EncoderParameterValueTypeRationalRange* | 8 | Specifies that the parameter is an array of four, 32-bit unsigned integers. The first two integers represent one fraction, and the second two integers represent a second fraction. The two fractions represent a range of rational numbers. The first fraction is the smallest rational number in the range, and the second fraction is the largest rational number in the range. |
| *EncoderParameterValueTypePointer* | 9 | Specifies that the parameter is a pointer to a block of custom metadata. |

---

## EncoderValue

Specifies values that can be passed as arguments to image encoders.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *EncoderValueColorTypeCMYK* | 0 | Not used in GDI+ version 1.0. |
| *EncoderValueColorTypeYCCK* | 1 | Not used in GDI+ version 1.0. |
| *EncoderValueCompressionLZW* | 2 | For a TIFF image, specifies the LZW compression method. |
| *EncoderValueCompressionCCITT3* | 3 | For a TIFF image, specifies the CCITT3 compression method. |
| *EncoderValueCompressionCCITT4* | 4 | For a TIFF image, specifies the CCITT4 compression method. |
| *EncoderValueCompressionRle* | 5 | For a TIFF image, specifies the RLE compression method. |
| *EncoderValueCompressionNone* | 6 | For a TIFF image, specifies no compression. |
| *EncoderValueScanMethodInterlaced* | 7 | Not used in GDI+ version 1.0. |
| *EncoderValueScanMethodNonInterlaced* | 8 | Not used in GDI+ version 1.0. |
| *EncoderValueVersionGif87* | 9 | Not used in GDI+ version 1.0. |
| *EncoderValueVersionGif89* | 10 | Not used in GDI+ version 1.0. |
| *EncoderValueRenderProgressive* | 11 | Not used in GDI+ version 1.0. |
| *EncoderValueRenderNonProgressive* | 12 | Not used in GDI+ version 1.0. |
| *EncoderValueTransformRotate90* | 13 | For a JPEG image, specifies lossless 90-degree clockwise rotation. |
| *EncoderValueTransformRotate180* | 14 | For a JPEG image, specifies lossless 180-degree rotation. |
| *EncoderValueTransformRotate270* | 15 | For a JPEG image, specifies lossless 270-degree clockwise rotation. |
| *EncoderValueTransformFlipHorizontal* | 16 | For a JPEG image, specifies a lossless horizontal flip. |
| *EncoderValueTransformFlipVertical* | 17 | For a JPEG image, specifies a lossless vertical flip. |
| *EncoderValueMultiFrame* | 18 | Specifies multiple-frame encoding. |
| *EncoderValueLastFrame* | 19 | Specifies the last frame of a multiple-frame image. |
| *EncoderValueFlush* | 20 | Specifies that the encoder object is to be closed. |
| *EncoderValueFrameDimensionTime* | 21 | Not used in GDI+ version 1.0. |
| *EncoderValueFrameDimensionResolution* | 22 | Not used in GDI+ version 1.0. |
| *EncoderValueFrameDimensionPage* | 23 | For a TIFF image, specifies the page frame dimension |
| *EncoderValueColorTypeGray* | 24 |  |
| *EncoderValueColorTypeRGB* | 26 |  |

---

## FillMode

The **FillMode** enumeration specifies how to fill areas that are formed when a path or curve intersects itself.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *FillModeAlternate* | 0 | Specifies that areas are filled according to the even-odd parity rule. According to this rule, you can determine whether a test point is inside or outside a closed curve as follows: Draw a line from the test point to a point that is distant from the curve. If that line crosses the curve an odd number of times, the test point is inside the curve; otherwise, the test point is outside the curve. |
| *FillModeWinding* | 1 | Specifies that areas are filled according to the nonzero winding rule. According to this rule, you can determine whether a test point is inside or outside a closed curve as follows: Draw a line from a test point to a point that is distant from the curve. Count the number of times the curve crosses the test line from left to right, and count the number of times the curve crosses the test line from right to left. If those two numbers are the same, the test point is outside the curve; otherwise, the test point is inside the curve. |

---

## FlushIntention

The **FlushIntention** enumeration specifies when to flush the queue of graphics operations.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *FlushIntentionFlush* | 0 | When passed to the **GdipFlush** function, specifies that pending rendering operations are executed as soon as possible. The **GdipFlush** function is not synchronized with the completion of the rendering operations and might return before the rendering operations are completed. |
| *FlushIntentionSync* | 1 | When passed to the **GdipFlush** function, specifies that pending rendering operations are executed as soon as possible. The **GdipFlush** function is synchronized with the completion of the rendering operations; that is, it will not return until after the rendering operations are completed. |

---

## FontStyle

The **FontStyle** enumeration specifies the style of the typeface of a font. Styles can be combined.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *FontStyleRegular* | 0 | Specifies normal weight or thickness of the typeface. |
| *FontStyleBold* | 1 | Specifies bold typeface. Bold is a heavier weight or thickness. |
| *FontStyleItalic* | 2 | Specifies italic typeface, which produces a noticeable slant to the vertical stems of the characters. |
| *FontStyleBoldItalic* | 3 | Specifies the typeface as both bold and italic. |
| *FontStyleUnderline* | 4 | Specifies underline, which displays a line underneath the baseline of the characters. |
| *FontStyleStrikeout* | 8 | Specifies strikeout, which displays a horizontal line drawn through the middle of the characters. |

---

## GenericFontFamily

Specifies the type of generic font family.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *GenericFontFamilySerif* | 0 | Specifies a generic serif typeface. |
| *GenericFontFamilySansSerif* | 1 | Specifies a generic sans serif typeface. |
| *GenericFontFamilyMonospace* | 2 | Specifies a generic monospace typeface. |

---

## HatchStyle

The **HatchStyle** enumeration specifies the hatch pattern used by a brush of type **HatchBrush**. The hatch pattern consists of a solid background color and lines drawn over the background.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *HatchStyleHorizontal* | 0 | Specifies horizontal lines. |
| *HatchStyleVertical* | 1 | Specifies vertical lines. |
| *HatchStyleForwardDiagonal* | 2 | Specifies diagonal lines that slant to the right from top points to bottom points. The lines are antialiased. |
| *HatchStyleBackwardDiagonal* | 3 | Specifies diagonal lines that slant to the left from top points to bottom points. The lines are antialiased. |
| *HatchStyleCross* | 4 | Specifies horizontal and vertical lines that cross at 90-degree angles. |
| *HatchStyleLargeGrid* | 4 | Specifies horizontal and vertical lines that cross at 90-degree angles. |
| *HatchStyleDiagonalCross* | 5 | Specifies forward diagonal and backward diagonal lines that cross at 90-degree angles. The lines are antialiased. |
| *HatchStyle05Percent* | 6 | Specifies a 5-percent hatch. The ratio of foreground color to background color is 5:100. |
| *HatchStyle10Percent* | 7 | Specifies a 10-percent hatch. The ratio of foreground color to background color is 10:100. |
| *HatchStyle20Percent* | 8 | Specifies a 20-percent hatch. The ratio of foreground color to background color is 20:100. |
| *HatchStyle25Percent* | 9 | Specifies a 25-percent hatch. The ratio of foreground color to background color is 25:100. |
| *HatchStyle30Percent* | 10 | Specifies a 30-percent hatch. The ratio of foreground color to background color is 30:100. |
| *HatchStyle40Percent* | 11 | Specifies a 40-percent hatch. The ratio of foreground color to background color is 40:100. |
| *HatchStyle50Percent* | 12 | Specifies a 50-percent hatch. The ratio of foreground color to background color is 50:100. |
| *HatchStyle60Percent* | 13 | Specifies a 60-percent hatch. The ratio of foreground color to background color is 60:100. |
| *HatchStyle70Percent* | 14 | Specifies a 70-percent hatch. The ratio of foreground color to background color is 70:100. |
| *HatchStyle75Percent* | 15 | Specifies a 75-percent hatch. The ratio of foreground color to background color is 75:100. |
| *HatchStyle80Percent* | 16 | Specifies an 80-percent hatch. The ratio of foreground color to background color is 80:100. |
| *HatchStyle90Percent* | 17 | Specifies a 90-percent hatch. The ratio of foreground color to background color is 90:100. |
| *HatchStyleLightDownwardDiagonal* | 18 | Specifies diagonal lines that slant to the right from top points to bottom points and are spaced 50 percent closer together than *HatchStyleForwardDiagonal* but are not antialiased. |
| *HatchStyleLightUpwardDiagonal* | 19 | Specifies diagonal lines that slant to the left from top points to bottom points and are spaced 50 percent closer together than HatchStyleBackwardDiagonal but are not antialiased. |
| *HatchStyleDarkDownwardDiagonal* | 20 | Specifies diagonal lines that slant to the right from top points to bottom points, are spaced 50 percent closer together than HatchStyleForwardDiagonal, and are twice the width of HatchStyleForwardDiagonal but are not antialiased. |
| *HatchStyleDarkUpwardDiagonal* | 21 | Specifies diagonal lines that slant to the left from top points to bottom points, are spaced 50 percent closer together than HatchStyleBackwardDiagonal, and are twice the width of HatchStyleBackwardDiagonal but are not antialiased. |
| *HatchStyleWideDownwardDiagonal* | 22 | Specifies diagonal lines that slant to the right from top points to bottom points, have the same spacing as HatchStyleForwardDiagonal, and are triple the width of HatchStyleForwardDiagonal but are not antialiased. |
| *HatchStyleWideUpwardDiagonal* | 23 | Specifies diagonal lines that slant to the left from top points to bottom points, have the same spacing as HatchStyleBackwardDiagonal, and are triple the width of HatchStyleBackwardDiagonal but are not antialiased. |
| *HatchStyleLightVertical* | 24 | Specifies vertical lines that are spaced 50 percent closer together than HatchStyleVertical. |
| *HatchStyleLightHorizontal* | 25 | Specifies horizontal lines that are spaced 50 percent closer together than HatchStyleHorizontal. |
| *HatchStyleNarrowVertical* | 26 | Specifies vertical lines that are spaced 75 percent closer together than*HatchStyleVertical* (or 25 percent closer together than *HatchStyleLightVertical*). |
| *HatchStyleNarrowHorizontal* | 27 | Specifies horizontal lines that are spaced 75 percent closer together than *HatchStyleHorizontal* ( or 25 percent closer together than *HatchStyleLightHorizontal*). |
| *HatchStyleDarkVertical* | 28 | Specifies vertical lines that are spaced 50 percent closer together than *HatchStyleVerical* and are twice the width of *HatchStyleVertical*. |
| *HatchStyleDarkHorizontal* | 29 | Specifies horizontal lines that are spaced 50 percent closer together than *HatchStyleHorizontal* and are twice the width of *HatchStyleHorizontal*. |
| *HatchStyleDashedDownwardDiagonal* | 30 | Specifies horizontal lines that are composed of forward diagonals. |
| *HatchStyleDashedUpwardDiagonal* | 31 | Specifies horizontal lines that are composed of backward diagonals. |
| *HatchStyleDashedHorizontal* | 32 | Specifies horizontal dashed lines. |
| *HatchStyleDashedVertical* | 33 | Specifies vertical dashed lines. |
| *HatchStyleSmallConfetti* | 34 | Specifies a hatch that has the appearance of confetti. |
| *HatchStyleLargeConfetti* | 35 | Specifies a hatch that has the appearance of confetti composed of larger pieces than *HatchStyleSmallConfetti*. |
| *HatchStyleZigZag* | 36 | Specifies horizontal lines of zigzags. |
| *HatchStyleWave* | 37 | Specifies horizontal lines of tildes. |
| *HatchStyleDiagonalBrick* | 38 | Specifies a hatch that has the appearance of a wall of bricks laid in a backward diagonal direction. |
| *HatchStyleHorizontalBrick* | 39 | Specifies a hatch that has the appearance of a wall of bricks laid horizontally. |
| *HatchStyleWeave* | 40 | Specifies a hatch that has the appearance of a woven material. |
| *HatchStylePlaid* | 41 | Specifies a hatch that has the appearance of a plaid material. |
| *HatchStyleDivot* | 42 | Specifies a hatch that has the appearance of divots. |
| *HatchStyleDottedGrid* | 43 | Specifies horizontal and vertical dotted lines that cross at 90-degree angles. |
| *HatchStyleDottedDiamond* | 44 | Specifies forward diagonal and backward diagonal dotted lines that cross at 90-degree angles. |
| *HatchStyleShingle* | 45 | Specifies a hatch that has the appearance of shingles laid in a forward diagonal direction. |
| *HatchStyleTrellis* | 46 | Specifies a hatch that has the appearance of a trellis. |
| *HatchStyleSphere* | 47 | Specifies a hatch that has the appearance of a checkerboard of spheres. |
| *HatchStyleSmallGrid* | 48 | Specifies horizontal and vertical lines that cross at 90-degree angles and are spaced 50 percent closer together than *HatchStyleCross*. |
| *HatchStyleSmallCheckerBoard* | 49 | Specifies a hatch that has the appearance of a checkerboard. |
| *HatchStyleLargeCheckerBoard* | 50 | Specifies a hatch that has the appearance of a checkerboard with squares that are twice the size of *HatchStyleSmallCheckerBoard*. |
| *HatchStyleOutlinedDiamond* | 51 | Specifies forward diagonal and backward diagonal lines that cross at 90-degree angles but are not antialiased. |
| *HatchStyleSolidDiamond* | 52 | Specifies a hatch that has the appearance of a checkerboard placed diagonally. |
| *HatchStyleTotal* | 53 | Specifies no hatch thereby allowing the brush to be transparent. |
| *HatchStyleMin* | HatchStyleHorizontal | Specifies *HatchStyleHorizonal*. |
| *HatchStyleMax* | HatchStyleTotal - 1 | Specifies *HatchStyleSolidDiamond*. |

---

## HistogramFormat

The **HistogramFormat** enumeration specifies the number and type of histograms that represent the color channels of a bitmap. This enumeration is used with the **GdipBitmapGetHistogram** function.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *HistogramFormatARGB* | 0 | Specifies that the **GdipBitmapGetHistogram** function returns four histograms: one each for the alpha, red, green, and blue channels. The alpha-channel histogram is written to the buffer pointed to by the *channel0* parameter of the **GdipBitmapGetHistogram** function. The red-channel histogram is written to the buffer pointed to by the *channel1* parameter. The green-channel histogram is written to the buffer pointed to by the *channel2* parameter. The blue-channel histogram is written to the buffer pointed to by the *channel3* parameter. |
| *HistogramFormatPARGB* | 1 | Specifies that the **GdipBitmapGetHistogram** function returns four histograms: one each for the alpha, red, green, and blue channels. The red, green, and blue channels are each multiplied by the alpha channel before the histograms are created. The bitmap is not permanently altered when the color channels are multiplied by the alpha channel; that multiplication is only for the purpose of creating the histograms. The alpha-channel histogram is written to the buffer pointed to by the *channel0* parameter of the **GdipBitmapGetHistogram** function. The red-channel histogram is written to the buffer pointed to by the *channel1* parameter. The green-channel histogram is written to the buffer pointed to by the *channel2* parameter. The blue-channel histogram is written to the buffer pointed to by the *channel3* parameter. |
| *HistogramFormatRGB* | 2 | Specifies that the **GdipBitmapGetHistogram** function returns three histograms: one each for the red, green, and blue channels. The red-channel histogram is written to the buffer pointed to by the *channel0* parameter of the **GdipBitmapGetHistogram** function. The green-channel histogram is written to the buffer pointed to by the *channel1* parameter. The blue-channel histogram is written to the buffer pointed to by the *channel2* parameter. The *channel3* parameter must be set to NULL. |
| *HistogramFormatGray* | 3 | Specifies that each pixel is converted to a grayscale value in the range 0 through 255, and then one histogram, based on those grayscale value, is returned. The bitmap is not permanently altered by the conversion to grayscale values; those values are calculated only for the purpose of creating the histogram. The grayscale histogram is written to the buffer pointed to by the *channel0* parameter of the **GdipBitmapGetHistogram** function. The *channel1*, *channel2*, and *channel3* parameters must be set to NULL. |
| *HistogramFormatB* | 4 | Specifies that the **GdipBitmapGetHistogram** function returns a histogram for the blue channel. The blue-channel histogram is written to the buffer pointed to by the *channel0* parameter of the **GdipBitmapGetHistogram** fdunction. The *channel1*, *channel2*, and *channel3* parameters must be set to NULL. |
| *HistogramFormatG* | 5 | Specifies that the **GdipBitmapGetHistogram** function returns a histogram for the green channel. The green-channel histogram is written to the buffer pointed to by the *channel0* parameter of the **GdipBitmapGetHistogram** function. The *channel1*, *channel2*, and *channel3* parameters must be set to NULL. |
| *HistogramFormatR* | 6 | Specifies that the **GdipBitmapGetHistogram** function returns a histogram for the red channel. The red-channel histogram is written to the buffer pointed to by the *channel0* parameter of the **GdipBitmapGetHistogram** function. The *channel1*, *channel2*, and *channel3* parameters must be set to NULL. |
| *HistogramFormatA* | 7 | Specifies that the **GdipBitmapGetHistogram** function returns a histogram for the alpha channel. The alpha-channel histogram is written to the buffer pointed to by the *channel0* parameter of the **GdipBitmapGetHistogram** function. The *channel1*, *channel2*, and *channel3* parameters must be set to NULL. |

---

## HotKeyPrefix

The **HotkeyPrefix** enumeration specifies how to display hot keys. There are three options: do nothing, display hot keys underlined, and hide the hot key underlines.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *HotkeyPrefixNone* | 0 | Specifies that no hot key processing occurs. |
| *HotkeyPrefixShow* | 1 | Specifies that Unicode text is scanned for ampersands (&), which are interpreted as hot key markers in the same way as menu and dialog resources are processed in the Windows user interface. All pairs of ampersands are replaced by a single ampersand. All single ampersands are removed and the first character that follows the first single ampersand is displayed underlined. |
| *HotkeyPrefixHide* | 2 | Specifies that Unicode text is scanned for ampersands (&), which are substituted and removed as in HotkeyPrefixShow but no underline is shown. This option can be used when accessibility hot key underlines are not required. |

---

## ImageCodecFlags

The **ImageCodecFlags** enumeration indicates attributes of an image codec.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *ImageCodecFlagsEncoder* | &h00000001 | Indicates that the codec supports encoding (saving). |
| *ImageCodecFlagsDecoder* | &h00000002 | Indicates that the codec supports decoding (reading). |
| *ImageCodecFlagsSupportBitmap* | &h00000004 | Indicates that the codec supports raster images (bitmaps). |
| *ImageCodecFlagsSupportVector* | &h00000008 | Indicates that the codec supports vector images (metafiles). |
| *ImageCodecFlagsSeekableEncode* | &h00000010 | Indicates that an encoder requires a seekable output stream. |
| *ImageCodecFlagsBlockingDecode* | &h00000020 | Indicates that a decoder has blocking behavior during the decoding process. |
| *ImageCodecFlagsBuiltin* | &h00010000 | Indicates that the codec is built in to GDI+. |
| *ImageCodecFlagsSystem* | &h00020000 | Not used in GDI+ version 1.0. |
| *ImageCodecFlagsUser* | &h00040000 | Not used in GDI+ version 1.0. |

---

## ImageFlags

Specifies the attributes of the pixel data contained in an image. The **GdipGetImageFlags** function returns an element of this enumeration.

---

## ImageLockMode

Specifies flags that are passed to the flags parameter of the **GdipBitmapLockBits** function. The GdipBitmapLockBits function locks a portion of an image so that you can read or write the pixel data.

---

## ImageType

The **ImageType** enumeration indicates whether an image is a bitmap or a metafile. The **GdipGetImageType** method returns an element of this enumeration.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *ImageTypeUnknown* | 0 | Indicates that the image type is not known. |
| *ImageTypeBitmap* | 1 | Indicates a bitmap image. |
| *ImageTypeMetafile* | 2 | Indicates a metafile image. |

---

## InterpolationMode

The **InterpolationMode** enumeration specifies the algorithm that is used when images are scaled or rotated. This enumeration is used by the **GdipGetInterpolationMode** and **GdipSetInterpolationMode** functions of the **Graphics** functions.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *InterpolationModeInvalid* | -1 | Used internally |
| *InterpolationModeDefault* | 0 | Specifies the default interpolation mode. |
| *InterpolationModeLowQuality* | 1 | Specifies a low-quality mode. |
| *InterpolationModeHighQuality* | 2 | Specifies a high-quality mode. |
| *InterpolationModeBilinear* | 3 | Specifies bilinear interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 50 percent of its original size. |
| *InterpolationModeBicubic* | 4 | Specifies bicubic interpolation. No prefiltering is done. This mode is not suitable for shrinking an image below 25 percent of its original size. |
| *InterpolationModeNearestNeighbor* | 5 | Specifies nearest-neighbor interpolation. |
| *InterpolationModeHighQualityBilinear* | 6 | Specifies high-quality, bilinear interpolation. Prefiltering is performed to ensure high-quality shrinking. |
| *InterpolationModeHighQualityBicubic* | 7 | Specifies high-quality, bicubic interpolation. Prefiltering is performed to ensure high-quality shrinking. This mode produces the highest quality transformed images. |

---

## ItemDataPosition

Used to specify the location of custom metadata in an image file.

---

## LinearGradientMode

The **LinearGradientMode** enumeration specifies the direction in which the change of color occurs for a linear gradient brush.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *LinearGradientModeHorizontal* | 0 | Specifies the color to change in a horizontal direction from the left of the display to the right of the display. |
| *LinearGradientModeVertical* | 1 | Specifies the color to change in a vertical direction from the top of the display to the bottom of the display. |
| *LinearGradientModeForwardDiagonal* | 2 | Specifies the color to change in a forward diagonal direction from the upper-left corner to the lower-right corner of the display. |
| *LinearGradientModeBackwardDiagonal* | 3 | Specifies the color to change in a backward diagonal direction from the upper-right corner to the lower-left corner of the display. |

---

## LineCap

The **LineCap** enumeration specifies the type of graphic shape to use on the end of a line drawn with a Windows GDI+ pen. The cap can be a square, circle, triangle, arrowhead, custom, or masked (hidden). End caps can also "anchor" the line by centering the cap at the end of the line.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *LineCapFlat* | 0 | Specifies that the line ends at the last point. The end is squared off. |
| *LineCapSquare* | 1 | Specifies a square cap. The center of the square is the last point in the line. The height and width of the square are the line width. |
| *LineCapRound* | 2 | Specifies a circular cap. The center of the circle is the last point in the line. The diameter of the circle is the line width. |
| *LineCapTriangle* | 3 | Specifies a triangular cap. The base of the triangle is the last point in the line. The base of the triangle is the line width. |
| *LineCapNoAnchor* | &H10 | Specifies that the line ends are not anchored. |
| *LineCapSquareAnchor* | &H11 | Specifies that the line ends are anchored with a square. The center of the square is the last point in the line. The height and width of the square are the line width. |
| *LineCapRoundAnchor* | &H12 | Specifies that the line ends are anchored with a circle. The center of the circle is at the last point in the line. The circle is wider than the line. |
| *LineCapDiamondAnchor* | &H13 | Specifies that the line ends are anchored with a diamond (a square turned at 45 degrees). The center of the diamond is at the last point in the line. The diamond is wider than the line. |
| *LineCapArrowAnchor* | &H14 | Specifies that the line ends are anchored with arrowheads. The arrowhead point is located at the last point in the line. The arrowhead is wider than the line. |
| *LineCapCustom* | &HFF | Specifies that the line ends are made from a **CustomLineCap**. |
| *LineCapAnchorMask* | &HF0 |  |

---

## LineJoin

The **LineJoin** enumeration specifies how to join two lines that are drawn by the same pen and whose ends meet. At the intersection of the two line ends, a line join makes the join look more continuous.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *LineJoinMiter* | 0 | Specifies a mitered join. This produces a sharp corner or a clipped corner, depending on whether the length of the miter exceeds the miter limit. |
| *LineJoinBevel* | 1 | Specifies a beveled join. This produces a diagonal corner. |
| *LineJoinRound* | 2 | Specifies a circular join. This produces a smooth, circular arc between the lines. |
| *LineJoinMiterClipped* | 3 | Specifies a mitered join. This produces a sharp corner or a beveled corner, depending on whether the length of the miter exceeds the miter limit. |

---

## MatrixOrder

The **MatrixOrder** enumeration specifies the order of multiplication when a new matrix is multiplied by an existing matrix.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *MatrixOrderPrepend* | 0 | Specifies that the new matrix is on the left and the existing matrix is on the right. |
| *MatrixOrderAppend* | 1 | Specifies that the existing matrix is on the left and the new matrix is on the right. |

---

## MetafileFrameUnit

The **MetafileFrameUnit** enumeration specifies the unit of measure for a metafile frame rectangle.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *MetafileFrameUnitPixel* | 2 | Specifies that a unit is 1 pixel. |
| *MetafileFrameUnitPoint* | 3 | Specifies that a unit is 1 point. |
| *MetafileFrameUnitInch* | 4 | Specifies that a unit is 1 inch. |
| *MetafileFrameUnitDocument* | 5 | Specifies that a unit is 1/300 inch. |
| *MetafileFrameUnitMillimeter* | 6 | Specifies that a unit is 1 millimeter. |
| *MetafileFrameUnitGdi* | 7 | Specifies that a unit is 0.01 millimeter. This element is provided for compatibility with GDI. |

---

## MetafileType

The MetafileType enumeration specifies types of metafiles.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *MetafileTypeInvalid* | 0 | Specifies a metafile format that is not recognized in GDI+. |
| *MetafileTypeWmf* | 1 | Specifies a WMF file. Such a file contains only GDI records. |
| *MetafileTypeWmfPlaceable* | 2 | Specifies a WMF file that has a placeable metafile header in front of it. |
| *MetafileTypeEmf* | 3 | Specifies an EMF file. Such a file contains only GDI records. |
| *MetafileTypeEmfPlusOnly* | 4 | Specifies an EMF+ file. Such a file contains only GDI+ records and must be displayed by using GDI+. Displaying the records using GDI may cause unpredictable results. |
| *MetafileTypeEmfPlusDual* | 5 | Specifies an EMF+ Dual file. Such a file contains GDI+ records along with alternative GDI records and can be displayed by using either GDI or GDI+. Displaying the records using GDI may cause some quality degradation. |

---

## ObjectType

The **ObjectType** enumeration indicates the object type value of an EMF+ record.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *ObjectTypeInvalid* | 0 | Object type is invalid. |
| *ObjectTypeBrush* | 1 | Object type is a brush. |
| *ObjectTypePen* | 2 | Object type is a pen. |
| *ObjectTypePath* | 3 | Object type is a path. |
| *ObjectTypeRegion* | 4 | Object type is a region. |
| *ObjectTypeFont* | 5 | Object type is a font. |
| *ObjectTypeStringFormat* | 6 | Object type is a string format. |
| *ObjectTypeImageAttributes* | 7 | Object type is an image attribute. |
| *ObjectTypeCustomLineCap* | 8 | Object type is a custom line cap. |
| *ObjectTypeGraphics* | 9 | Object type is graphics. |
| *ObjectTypeMin* | ObjectTypeBrush | Minimum enumeration value. Currently, it is ObjectTypeBrush. |
| *ObjectTypeMax* | ObjectTypeGraphics | Maximum enumeration value. Currently, it is ObjectTypeGraphics. |

#### Remarks

To determine whether the object type value of an EMF+ record is valid, call **ObjectTypeIsValid**.

---

## PaletteFlags

The **PaletteFlags** enumeration indicates attributes of the color data in a palette.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *PaletteFlagsHasAlpha* | &h0001 | Indicates that one or more of the palette entries contains alpha (transparency) information. |
| *PaletteFlagsGrayScale* | &h0002 | Indicates that the palette contains only grayscale entries. |
| *PaletteFlagsHalftone* | &h0004 | Indicates that the palette is the Windows halftone palette. |
  
---

## PaletteType

Specifies color palette types. Used by the **GdipInitializePalette** and **GdipBitmapConvertFormat** functions. The members of the enumeration identify several standard color palette formats.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *PaletteTypeCustom* | 0 | An arbitrary custom palette specified by the caller. |
| *PaletteTypeOptimal* | 1 | A palette of colors that are optimal for a particular bitmap. To create an optimal palette, pass PaletteTypeOptimal, the number of colors you want in the palette, and the address of a **Bitmap** object to the **GdipInitializePalette** function. |
| *PaletteTypeFixedBW* | 2 | A palette that has two colors. This palette type is suitable for bitmaps that store 1 bit per pixel. |
| *PaletteTypeFixedHalftone8* | 3 | A palette based on two intensities each (off or full) for the red, green, and blue channels. Also contains the 16 colors of the system palette. Because all eight of the on/off combinations of red, green, and blue are already in the system palette, this palette is the same as the system palette. This palette type is suitable for bitmaps that store 4 bits per pixel. |
| *PaletteTypeFixedHalftone27* | 4 | A palette based on three intensities each for the red, green, and blue channels. Also contains the 16 colors of the system palette. Eight of the 16 system palette colors are among the 27 three-intensity combinations of red, green, and blue, so the total number of colors in the palette is 35. If the palette also includes the transparent color, the total number of colors is 36. |
| *PaletteTypeFixedHalftone64* | 5 | A palette based on four intensities each for the red, green, and blue channels. Also contains the 16 colors of the system palette. Eight of the 16 system palette colors are among the 64 four-intensity combinations of red, green, and blue, so the total number of colors in the palette is 72. If the palette also includes the transparent color, the total number of colors is 73. |
| *PaletteTypeFixedHalftone125* | 6 | A palette based on five intensities each for the red, green, and blue channels. Also contains the 16 colors of the system palette. Eight of the 16 system palette colors are among the 125 five-intensity combinations of red, green, and blue, so the total number of colors in the palette is 133. If the palette also includes the transparent color, the total number of colors is 134. |
| *PaletteTypeFixedHalftone216* | 7 | A palette based on six intensities each for the red, green, and blue channels. Also contains the 16 colors of the system palette. Eight of the 16 system palette colors are among the 216 six-intensity combinations of red, green, and blue, so the total number of colors in the palette is 224. If the palette also includes the transparent color, the total number of colors is 225. This palette is sometimes called the Windows halftone palette or the Web palette. |
| *PaletteTypeFixedHalftone252* | 8 | A palette based on 6 intensities of red, 7 intensities of green, and 6 intensities of blue. The system palette is not included. The total number of colors is 252. If the palette also includes the transparent color, the total number of colors is 253. |
| *PaletteTypeFixedHalftone256* | 9 | A palette based on 8 intensities of red, 8 intensities of green, and 4 intensities of blue. The system palette is not included. The total number of colors is 256. If the transparent color is included in this palette, it must replace one of the RGB combinations. |

---

## PathPointType

Indicates  point types and flags for the data points in a path. Bits 0 through 2 indicate the type of a point, and bits 3 through 7 hold a set of flags that specify attributes of a point.

The **PathPointType** enumeration indicates point types and flags for the data points in a path. Bits 0 through 2 indicate the type of a point, and bits 3 through 7 hold a set of flags that specify attributes of a point. This enumeration is used by the GraphicsPath and PathIterator functions.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *PathPointTypeStart* | &h00 | Indicates that the point is the start of a figure. |
| *PathPointTypeLine* | &h01 | Indicates that the point is one of the two endpoints of a line. |
| *PathPointTypeBezier* | &h03 | Indicates that the point is an endpoint or control point of a cubic Bézier spline. |
| *PathPointTypeBezier3* | &h03 | Indicates that the point is an endpoint or control point of a cubic Bézier spline. |
| *PathPointTypePathTypeMask* | &h07 | Masks all bits except for the three low-order bits, which indicate the point type. |
| *PathPointTypePathDashMode* | &h10 | Not used. |
| *PathPointTypePathMarker* | &h20 | Specifies that the point is a marker. |
| *PathPointTypeCloseSubpath* | &h80 | Specifies that the point is the last point in a closed subpath (figure). |

#### Remarks

A **GraphicsPath** object has an array of points and an array of types. Each element in the array of types is a byte that specifies the point type and a set of flags for the corresponding element in the array of points.

---

## PenAlignment

The **PenAlignment** enumeration specifies the alignment of a pen relative to the stroke that is being drawn.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *PenAlignmentCenter* | 0 | Specifies that the pen is aligned on the center of the line that is drawn. |
| *PenAlignmentInset* | 1 | Specifies, when drawing a polygon, that the pen is aligned on the inside of the edge of the polygon. |

#### Remarks

If you set the alignment of a **Pen** object to **PenAlignmentInset**, you cannot use that pen to draw compound lines or triangular dash caps.

---

## PenType

The **PenType** enumeration indicates the type of pattern, texture, or gradient that a pen draws.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *PenTypeUnknown* | -1 | Indicates that the pen type is unknown. |
| *PenTypeSolidColor* | 0 | Indicates that the pen draws with a solid color. |
| *PenTypeHatchFill* | 1 | Indicates that the pen draws with a hatch pattern that is specified by a **HatchBrush** object. |
| *PenTypeTextureFill* | 2 | Indicates that the pen draws with a texture that is specified by a **TextureBrush** object. |
| *PenTypePathGradient* | 3 | Indicates that the pen draws with a color gradient that is specified by a **PathGradientBrush** object. |
| *PenTypeLinearGradient* | 4 | Indicates that the pen draws with a color gradient that is specified by a **LinearGradientBrush** object. |

#### Remarks

A pen's type is determined when the pen is constructed. For example, if you pass a **HatchBrush** object to a Pen constructor, then the pen that is constructed has a pen type of *PenTypeHatchFill*. If you pass a **Color** object or a **SolidBrush** object to a **Pen** constructor, then the pen that is constructed has a pen type of *PenTypeSolidColor*.

---

## PixelOffsetMode

The **PixelOffsetMode** enumeration specifies the pixel offset mode. This enumeration is used by the **GdipGetPixelOffsetMode** and **GdipSetPixelOffsetMode** methods of the **Graphics** class.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *PixelOffsetModeInvalid* | -1 | Used internally. |
| *PixelOffsetModeDefault* | 0 | Equivalent to *PixelOffsetModeNone*. |
| *PixelOffsetModeHighSpeed* | 1 | Equivalent to *PixelOffsetModeNone*. |
| *PixelOffsetModeHighQuality* | 2 | Equivalent to *PixelOffsetModeHalf*. |
| *PixelOffsetModeNone* | 3 | Indicates that pixel centers have integer coordinates. |
| *PixelOffsetModeHalf* | 4 | Indicates that pixel centers have coordinates that are half way between integer values. |

#### Remarks

Consider the pixel in the upper-left corner of an image with address (0, 0). With *PixelOffsetModeNone*, the pixel covers the area between –0.5 and 0.5 in both the x and y directions; that is, the pixel center is at (0, 0). With *PixelOffsetModeHalf*, the pixel covers the area between 0 and 1 in both the x and y directions; that is, the pixel center is at (0.5, 0.5).

---

## QualityMode

Specifies the quality mode.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *QualityModeInvalid* | -1 | Specifies an invalid mode. |
| *QualityModeDefault* | 0 | Specifies the default mode. |
| *QualityModeLow* | 1 | Specifies low quality, high speed rendering. |
| *QualityModeHigh* | 2 | Specifies high quality, low speed rendering. |

#### Remarks

The quality level is inversely proportional to the amount of time spent rendering. This mode does not affect text. To set the text rendering quality, use the **TextRenderingHint** enumeration.

---

## RotateFlipType

The **RotateFlipType** enumeration specifies the direction of an image's rotation and the axis used to flip the image.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *RotateNoneFlipNone* | 0 | Specifies no rotation and no flipping. |
| *Rotate90FlipNone* | 1 | Specifies a 90-degree rotation without flipping. |
| *Rotate180FlipNone* | 2 | Specifies a 180-degree rotation without flipping. |
| *Rotate270FlipNone* | 3 | Specifies a 270-degree rotation without flipping. |
| *RotateNoneFlipX* | 4 | Specifies no rotation and a horizontal flip. |
| *Rotate90FlipX* | 5 | Specifies a 90-degree rotation followed by a horizontal flip. |
| *Rotate180FlipX* | 6 | Specifies a 180-degree rotation followed by a horizontal flip. |
| *Rotate270FlipX* | 7 | Specifies a 270-degree rotation followed by a horizontal flip. |
| *RotateNoneFlipY* | Rotate180FlipX | Specifies no rotation and a vertical flip. |
| *Rotate90FlipY* | Rotate270FlipX | Specifies a 90-degree rotation followed by a vertical flip. |
| *Rotate180FlipY* | RotateNoneFlipX | Specifies a 180-degree rotation followed by a vertical flip. |
| *Rotate270FlipY* | Rotate90FlipX | Specifies a 270-degree rotation followed by a vertical flip. |
| *RotateNoneFlipXY* | Rotate180FlipNone | Specifies no rotation, a horizontal flip, and then a vertical flip. |
| *Rotate90FlipXY* | Rotate270FlipNone | Specifies a 90-degree rotation followed by a horizontal flip and then a vertical flip. |
| *Rotate180FlipXY* | RotateNoneFlipNone | Specifies a 180-degree rotation followed by a horizontal flip and then a vertical flip. |
| *Rotate270FlipXY* | Rotate90FlipNone | Specifies a 270-degree rotation followed by a horizontal flip and then a vertical flip. |

---

## SmoothingMode

The **SmoothingMode** enumeration specifies the type of smoothing (antialiasing) that is applied to lines and curves. This enumeration is used by the **GdipGetSmoothingMode** and **GdipSetSmoothingMode** functions.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *SmoothingModeInvalid* | QualityModeInvalid | Reserved. |
| *SmoothingModeDefault* | 0 | Specifies that smoothing is not applied. |
| *SmoothingModeHighSpeed* | 1 | Specifies that smoothing is not applied. |
| *SmoothingModeHighQuality* | 2 | Specifies that smoothing is applied using an 8 X 4 box filter. |
| *SmoothingModeNone* | 3 | Specifies that smoothing is not applied. |
| *SmoothingModeAntiAlias8x4* | 4 | Specifies that smoothing is applied using an 8 X 4 box filter. |
| *SmoothingModeAntiAlias* | 4 | Specifies that smoothing is applied using an 8 X 4 box filter. |
| *SmoothingModeAntiAlias8x8* | 5 | Specifies that smoothing is applied using an 8 X 8 box filter. |

#### Remarks

Smoothing performed by an 8 X 4 box filter gives better results for nearly vertical lines than it does for nearly horizontal lines. Smoothing performed by an 8 X 8 box filter gives equally good results for nearly vertical and nearly horizontal lines. The 8x8 algorithm produces higher quality smoothing but is slower than the 8 X 4 algorithm.

---

## Status

The **Status** enumeration indicates the result of a Microsoft Windows GDI+ method call. Renamed as **GpStatus** to avoid naming conflicts.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StatusOk* | 0 | Indicates that the method call was successful. |
| *StatusGenericError* | 1 | Indicates that there was an error on the method call, which is identified as something other than those defined by the other elements of this enumeration. |
| *StatusInvalidParameter* | 2 | Indicates that one of the arguments passed to the method was not valid. |
| *StatusOutOfMemory* | 3 | Indicates that the operating system is out of memory and could not allocate memory to process the method call. For an explanation of how constructors use the *OutOfMemory* status, see the Remarks section at the end of this topic. |
| *StatusObjectBusy* | 4 | Indicates that one of the arguments specified in the API call is already in use in another thread. |
| *StatusInsufficientBuffer* | 5 | Indicates that a buffer specified as an argument in the API call is not large enough to hold the data to be received. |
| *StatusNotImplemented* | 6 | Indicates that the method is not implemented. |
| *StatusWin32Error* | 7 | Indicates that the method generated a Win32 error. |
| *StatusWrongState* | 8 | Indicates that the object is in an invalid state to satisfy the API call. For example, calling
**GdipGetPenColor** from a pen that is not a single, solid color results in a *WrongState* status. |
| *StatusAborted* | 9 | Indicates that the method was aborted. |
| *StatusFileNotFound* | 10 | Indicates that the specified image file or metafile cannot be found. |
| *StatusValueOverflow* | 11 | Indicates that the method performed an arithmetic operation that produced a numeric overflow. |
| *StatusAccessDenied* | 12 | Indicates that a write operation is not allowed on the specified file. |
| *StatusUnknownImageFormat* | 13 | Indicates that the specified image file format is not known. |
| *StatusFontFamilyNotFound* | 14 | Indicates that the specified font family cannot be found. Either the font family name is incorrect or the font family is not installed. |
| *StatusFontStyleNotFound* | 15 | Indicates that the specified style is not available for the specified font family. |
| *StatusNotTrueTypeFont* | 16 | Indicates that the font retrieved from an **HDC** or **LOGFONT** is not a TrueType font and cannot be used with GDI+. |
| *StatusUnsupportedGdiplusVersion* | 17 | Indicates that the version of GDI+ that is installed on the system is incompatible with the version with which the application was compiled. |
| *StatusGdiplusNotInitialized* | 18 | Indicates that the GDI+API is not in an initialized state. To function, all GDI+ objects require that GDI+ be in an initialized state. Initialize GDI+ by calling **GdiplusStartup**. |
| *StatusPropertyNotFound* | 19 | Indicates that the specified property does not exist in the image. |
| *StatusPropertyNotSupported* | 20 | Indicates that the specified property is not supported by the format of the image and, therefore, cannot be set. |
| *ProfileNotFound* | 21 | Indicates that the color profile required to save an image in CMYK format was not found. |
 
---

## StringAlignment

The **StringAlignment** enumeration specifies how a string is aligned in reference to the bounding rectangle. A bounding rectangle is used to define the area in which the text displays.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StringAlignmentNear* | 0 | Specifies that alignment is towards the origin of the bounding rectangle. May be used for alignment of characters along the line or for alignment of lines within the rectangle. For a right to left bounding rectangle (**StringFormatFlagsDirectionRightToLeft**), the origin is at the upper right. |
| *StringAlignmentCenter* | 1 | Specifies that alignment is centered between origin and extent (width or height) of the formatting rectangle. |
| *StringAlignmentFar* | 2 | Specifies that alignment is to the far extent (lower right) of the formatting rectangle. |

---

## StringDigitSubstitute

The **StringDigitSubstitute** enumeration specifies how to substitute digits in a string according to a user's locale or language.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StringDigitSubstituteUser* | 0 | Specifies a user-defined substitution scheme. |
| *StringDigitSubstituteNone* | 1 | Specifies to disable substitutions. |
| *StringDigitSubstituteNational* | 2 | Specifies substitution digits that correspond with the official national language of the user's locale. |
| *StringDigitSubstituteTraditional* | 3 | Specifies substitution digits that correspond with the user's native script or language, which may be different from the official national language of the user's locale. |

---

## StringFormatFlags

The **StringFormatFlags** enumeration specifies text layout information (such as orientation and clipping) and display manipulations (such as ellipsis insertion, digit substitution, and representation of characters that are not supported by a font).

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StringFormatFlagsDirectionRightToLeft* | &h00000001 | Specifies that reading order is right to left. For horizontal text, characters are read from right to left. For vertical text, columns are read from right to left. By default, horizontal or vertical text is read from left to right. |
| *StringFormatFlagsDirectionVertical* | &h00000002 | Specifies that individual lines of text are drawn vertically on the display device. By default, lines of text are horizontal, each new line below the previous line. |
| *StringFormatFlagsNoFitBlackBox* | &h00000004 | Specifies that parts of characters are allowed to overhang the string's layout rectangle. By default, characters are first aligned inside the rectangle's boundaries, then any characters which still overhang the boundaries are repositioned to avoid any overhang and thereby avoid affecting pixels outside the layout rectangle. An italic, lowercase letter F (f) is an example of a character that may have overhanging parts. Setting this flag ensures that the character aligns visually with the lines above and below but may cause parts of characters, which lie outside the layout rectangle, to be clipped or painted. |
| *StringFormatFlagsDisplayFormatControl* | &h00000020 | Specifies that Unicode layout control characters are displayed with a representative character. |
| *StringFormatFlagsNoFontFallback* | &h00000400 | Specifies that an alternate font is used for characters that are not supported in the requested font. By default, any missing characters are displayed with the "fonts missing" character, usually an open square. |
| *StringFormatFlagsMeasureTrailingSpaces* | &h00000800 | Specifies that the space at the end of each line is included in a string measurement. By default, the boundary rectangle returned by the **GdipMeasureString** method excludes the space at the end of each line. Set this flag to include that space in the measurement. |
| *StringFormatFlagsNoWrap* | &h00001000 | Specifies that the wrapping of text to the next line is disabled. NoWrap is implied when a point of origin is used instead of a layout rectangle. When drawing text within a rectangle, by default, text is broken at the last word boundary that is inside the rectangle's boundary and wrapped to the next line. |
| *StringFormatFlagsLineLimit* | &h00002000 | Specifies that only entire lines are laid out in the layout rectangle. By default, layout continues until the end of the text or until no more lines are visible as a result of clipping, whichever comes first. The default settings allow the last line to be partially obscured by a layout rectangle that is not a whole multiple of the line height. To ensure that only whole lines are seen, set this flag and be careful to provide a layout rectangle at least as tall as the height of one line. |
| *StringFormatFlagsNoClip* | &h00004000 | Specifies that characters overhanging the layout rectangle and text extending outside the layout rectangle are allowed to show. By default, all overhanging characters and text that extends outside the layout rectangle are clipped. Any trailing spaces (spaces that are at the end of a line) that extend outside the layout rectangle are clipped. Therefore, the setting of this flag will have an effect on a string measurement if trailing spaces are being included in the measurement. If clipping is enabled, trailing spaces that extend outside the layout rectangle are not included in the measurement. If clipping is disabled, all trailing spaces are included in the measurement, regardless of whether they are outside the layout rectangle. |
| *StringFormatFlagsBypassGDI* | &h80000000 |  |

#### Remarks

Multiple flags set can produce combined effects:

* When both **StringFormatFlagsDirectionVertical** and **StringFormatFlagsDirectionRightToLeft** are set, individual lines of text are drawn vertically. The first line starts at the right edge of the layout rectangle; the second line of text is to the left of the first line, and so on.
* When **StringFormatFlagsDirectionVertical** is set and **StringFormatFlagsDirectionRightToLeft** is not set, individual lines of text are drawn vertically. The first line starts at the left edge of the layout rectangle; the second line of text is to the right of the first line.
* When **StringFormatFlagsDirectionRightToLeft** is set and **StringFormatFlagsDirectionVertical** is not set, the individual lines of text are horizontal and the reading order is from right to left. This setting does not change the order in which characters are displayed, it simply specifies the order in which characters can be read.

The **StringFormatFlagsDirectionVertical** and **StringFormatFlagsDirectionRightToLeft** flags can affect string alignment.

---

## StringTrimming

The **StringTrimming** enumeration specifies how to trim characters from a string so that the string fits into a layout rectangle. The layout rectangle is used to position and size the display string.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *StringTrimmingNone* | 0 | Specifies that no trimming is done. |
| *StringTrimmingCharacter* | 1 | Specifies that the string is broken at the boundary of the last character that is inside the layout rectangle. This is the default. |
| *StringTrimmingWord* | 2 | Specifies that the string is broken at the boundary of the last word that is inside the layout rectangle. |
| *StringTrimmingEllipsisCharacter* | 3 | Specifies that the string is broken at the boundary of the last character that is inside the layout rectangle and an ellipsis (...) is inserted after the character. |
| *StringTrimmingEllipsisWord* | 4 | Specifies that the string is broken at the boundary of the last word that is inside the layout rectangle and an ellipsis (...) is inserted after the word. |
| *StringTrimmingEllipsisPath* | 5 | Specifies that the center is removed from the string and replaced by an ellipsis. The algorithm keeps as much of the last portion of the string as possible. |

#### Remarks

Trimming affects only the last visible or partly visible (due to clipping) line of text.

---

## TestControlEnum

Used for internal testing of GDI+. Specifies changes in image-processing behavior.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *TestControlForceBilinear* | 0 | Allows the use of the the *InerpolationModeBilinear* algorithm is used. |
| *TestControlForceNoICM* | 1 | Allows to ignore the ICM profiles when images are loaded into memory. |
| *TestControlGetBuildNumber* | 2 | Gets the GDI+ build number. |

---

## TextRenderingHint

The **TextRenderingHint** enumeration specifies the process used to render text. The process affects the quality of the text.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *TextRenderingHintSystemDefault* | 0 | Specifies that a character is drawn using the currently selected system font smoothing mode (also called a rendering hint). |
| *TextRenderingHintSingleBitPerPixelGridFit* | 1 | Specifies that a character is drawn using its glyph bitmap and hinting to improve character appearance on stems and curvature. |
| *TextRenderingHintSingleBitPerPixel* | 2 | Specifies that a character is drawn using its glyph bitmap and no hinting. This results in better performance at the expense of quality. |
| *TextRenderingHintAntiAliasGridFit* | 3 | Specifies that a character is drawn using its antialiased glyph bitmap and hinting. This results in much better quality due to antialiasing at a higher performance cost. |
| *TextRenderingHintAntiAlias* | 4 | Specifies that a character is drawn using its antialiased glyph bitmap and no hinting. Stem width differences may be noticeable because hinting is turned off. |
| *TextRenderingHintClearTypeGridFit* | 5 | Specifies that a character is drawn using its glyph ClearType bitmap and hinting. This type of text rendering cannot be used along with **CompositingModeSourceCopy**. Windows XP and Windows Server 2003 and later versions of Windows only: ClearType rendering is supported only on Windows XP and Windows Server 2003 and later versions of Windows. Therefore, **TextRenderingHintClearTypeGridFit** is ignored on other operating systems even though GDI+ is supported on those operating systems. |

#### Remarks

The quality associated with each process varies according to the circumstances. **TextRenderingHintClearTypeGridFit** provides the best quality for most LCD monitors and relatively small font sizes. **TextRenderingHintAntiAlias** provides the best quality for rotated text. Generally, a process that produces higher quality text is slower than a process that produces lower quality text.

---

## Unit

The **Unit** enumeration specifies the unit of measure for a given data type.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *UnitWorld* | 0 | Specifies world coordinates, a nonphysical unit. |
| *UnitDisplay* | 1 | Specifies display units. For example, if the display device is a monitor, then the unit is 1 pixel. |
| *UnitPixel* | 2 | Specifies that a unit is 1 pixel. |
| *UnitPoint* | 3 | Specifies that a unit is 1 point or 1/72 inch. |
| *UnitInch* | 4 | Specifies that a unit is 1 inch. |
| *UnitDocument* | 5 | Specifies that a unit is 1/300 inch. |
| *UnitMillimeter* | 6 | Specifies that a unit is 1 millimeter. |

---

## WarpMode

The **WarpMode** enumeration specifies warp modes that can be used to transform images.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *WarpModePerspective* | 0 | Specifies the perspective warp mode. |
| *WarpModeBilinear* | 1 | Specifies the bilinear warp mode. |

---

## WrapMode

The **WrapMode** enumeration specifies how repeated copies of an image are used to tile an area.

| Name  | Value | Description |
| ----- | ----- | ----------- |
| *WrapModeTile* | 0 | Specifies tiling without flipping. |
| *WrapModeTileFlipX* | 1 | Specifies that tiles are flipped horizontally as you move from one tile to the next in a row. |
| *WrapModeTileFlipY* | 2 | Specifies that tiles are flipped vertically as you move from one tile to the next in a column. |
| *WrapModeTileFlipXY* | 3 | Specifies that tiles are flipped horizontally as you move along a row and flipped vertically as you move along a column. |
| *WrapModeClamp* | 4 | Specifies that no tiling takes place. |

---
