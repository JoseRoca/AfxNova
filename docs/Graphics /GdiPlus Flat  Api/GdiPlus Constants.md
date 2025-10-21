# Image File Format Constants

The **GdipGetImageRawFormat** function returns a globally unique identifier (GUID) that indicates the file format of an image. The following constants represent the GUIDs that identify those file formats.

#### ImageFormatBMP 
B96B3CAB-0728-11D3-9D7B-0000F81EF32E
Indicates the Microsoft Windowsbitmap (BMP) format.

#### ImageFormatEMF 
B96B3CAC-0728-11D3-9D7B-0000F81EF32E
Indicates the Enhanced Metafile (EMF) format.

#### ImageFormatEXIF 
B96B3CB2-0728-11D3-9D7B-0000F81EF32E
Indicates the Exif (Exchangeable Image File) format.

#### ImageFormatGIF 
B96B3CB0-0728-11D3-9D7B-0000F81EF32E
Indicates the Graphics Interchange Format (GIF) format.

#### ImageFormatIcon 
B96B3CB5-0728-11D3-9D7B-0000F81EF32E
Indicates the Icon format.

#### ImageFormatJPEG 
B96B3CAE-0728-11D3-9D7B-0000F81EF32E
Indicates the JPEG format.

#### ImageFormatMemoryBMP 
B96B3CAB-0728-11D3-9D7B-0000F81EF32E
Indicates that the image was constructed from a memory bitmap.

#### ImageFormatPNG 
B96B3CAF-0728-11D3-9D7B-0000F81EF32E
Indicates the Portable Network Graphics (PNG) format.

#### ImageFormatTIFF 
B96B3CB1-0728-11D3-9D7B-0000F81EF32E
Indicates the Tagged Image File Format (TIFF) format.

#### ImageFormatUndefined 
B96B3CA9-0728-11D3-9D7B-0000F81EF32E
Indicates that Windows GDI+ is unable to determine the format.

#### ImageFormatWMF 
B96B3CAD-0728-11D3-9D7B-0000F81EF32E
Indicates the Windows Metafile Format (WMF) format.

---

# Image Frame Dimension Constants

The Graphics Interchange Format (GIF) and Tagged Image File Format (TIFF) image file formats enable you to store multiple frames in a single image file. Multiple frames in a GIF file are used for animation, so the frames are said to be in the time dimension. Multiple frames in a TIFF file are typically used as separate pages, so the frames are said to be in the page dimension.

The following constants represent globally unique identifiers (GUIDs) that identify the page and time dimensions.

#### FrameDimensionPage
7462DC86-6180-4C7E-8E3F-EE7333A7A483

#### FrameDimensionResolution
84236F7B-3BD3-428F-8DAB-4EA1439CA315

#### FrameDimensionTime
6AEDBD6D-3FB5-418A-83A6-7F45229DC872

---

#  Image Encoder Constants

The **GdpiSaveImageToFile**, **GdpiSaveImageToStream**, **GdpiSaveAdd** and **GdpiSaveAddImage** functions receive an **EncoderParameters** structure that contains an array of **EncoderParameter** objects. Each **EncoderParameter** object has a globally unique identifier (GUID) data member that specifies the parameter category. The following constants represent GUIDs that specify the various parameter categories.

#### EncoderChrominanceTable
F2E455DC-09B3-4316-8260-676ADA32481C

#### EncoderColorDepth
66087055-AD66-4C7C-9A18-38A2310B8337

#### EncoderColorSpace
AE7A62A0-EE2C-49D8-9D07-1BA8A927596E

#### EncoderCompression
E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58

#### EncoderLuminanceTable
EDB33BCE-0266-4A77-B904-27216099E717

#### EncoderQuality
1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB

#### EncoderRenderMethod
6D42C53A-229A-4825-8BB7-5C99E2B9A8B8

#### EncoderSaveFlag
292266FC-AC40-47BF-8CFC-A85B89A655DE

#### EncoderScanMethod
3A4E2661-3109-4E56-8536-42C156E7DCFA

#### EncoderTransformation
8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9

#### EncoderVersion
24D18C76-814A-41A4-BF53-1C219CCCF797

#### EncoderImageItems
63875E13-1F1D-45AB-9195-A29B6066A650

#### EncoderSaveAsCMYK
A219BBC9-0A9D-4005-A3EE-3A421B8BB06C

---

# Image Pixel Format Constants

The following constants specify various pixel formats used in bitmaps.

#### PixelFormat1bppIndexed
196865
Specifies that the format is 1 bit per pixel, indexed.

#### PixelFormat4bppIndexed 
197634
Specifies that the format is 4 bits per pixel, indexed.

#### PixelFormat8bppIndexed 
198659
Specifies that the format is 8 bits per pixel, indexed.

#### PixelFormat16bppARGB1555 
397319
Specifies that the format is 16 bits per pixel; 1 bit is used for the alpha component, and 5 bits each are used for the red, green, and blue components.

#### PixelFormat16bppGrayScale 
1052676
Specifies that the format is 16 bits per pixel, grayscale.

#### PixelFormat16bppRGB555 
135173
Specifies that the format is 16 bits per pixel; 5 bits each are used for the red, green, and blue components. The remaining bit is not used.

#### PixelFormat16bppRGB565 
135174
Specifies that the format is 16 bits per pixel; 5 bits are used for the red component, 6 bits are used for the green component, and 5 bits are used for the blue component.

#### PixelFormat24bppRGB 
137224
Specifies that the format is 24 bits per pixel; 8 bits each are used for the red, green, and blue components.

#### PixelFormat32bppARGB 
2498570
Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components.

#### PixelFormat32bppPARGB 
925707
Specifies that the format is 32 bits per pixel; 8 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component.

#### PixelFormat32bppRGB 
139273
Specifies that the format is 32 bits per pixel; 8 bits each are used for the red, green, and blue components. The remaining 8 bits are not used.

#### PixelFormat48bppRGB 
1060876
Specifies that the format is 48 bits per pixel; 16 bits each are used for the red, green, and blue components.

#### PixelFormat64bppARGB 
3424269
Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components.

#### PixelFormat64bppPARGB 
1851406
Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components. The red, green, and blue components are premultiplied according to the alpha component.

#### Remarks

**PixelFormat48bppRGB**, **PixelFormat64bppARGB**, and **PixelFormat64bppPARGB** use 16 bits per color component (channel). Microsoft Windows GDI+ version 1.0 can read 16-bits-per-channel images, but such images are converted to an 8-bits-per-channel format for processing, displaying, and saving.

---

# Image Property Tag Type Constants

You can store and retrieve image metadata with the help of a **PropertyItem** structure. The type data member of a **PropertyItem**  structure specifies the data type of the values stored in the value data member of that same **PropertyItem** structure.

The following constants can be assigned to the type data member of a **PropertyItem** structure.

PixelFormat4bppIndexed 

197634

Specifies that the format is 4 bits per pixel, indexed.

 

PropertyTagTypeASCII 

2

Specifies that the value data member is a null-terminated ASCII string. If you set the type data member of a PropertyItem object to PropertyTagTypeASCII, you should set the length data member to the length of the string including the NULL terminator. For example, the string HELLO would have a length of 6.

#### PropertyTagTypeByte 
1
Specifies that the value data member is an array of bytes.

#### PropertyTagTypeLong 
4
Specifies that the value data member is an array of unsigned long (32-bit) integers.

#### PropertyTagTypeRational 
4
Specifies that the value data member is an array of pairs of unsigned long integers. Each pair represents a fraction; the first integer is the numerator and the second integer is the denominator.

#### PropertyTagTypeShort 
3
Specifies that the value data member is an array of unsigned short (16-bit) integers.

#### PropertyTagTypeSLONG 
9
Specifies that the value data member is an array of signed long (32-bit) integers.

#### PropertyTagTypeSRational 
10
Specifies that the value data member is an array of pairs of signed long integers. Each pair represents a fraction; the first integer is the numerator and the second integer is the denominator.

#### PropertyTagTypeUndefined 
7
Specifies that the value data member is an array of bytes that can hold values of any data type.

---

# Image Property Tag Constants

Several image file formats enable you to store metadata along with an image. Metadata is information about an image, for example, title, width, camera model, and artist. Microsoft Windows GDI+ provides a uniform way of storing and retrieving metadata from image files in the Tagged Image File Format (TIFF), JPEG, Portable Network Graphics (PNG), and Exchangeable Image File (EXIF) formats.

In GDI+, a piece of metadata is called a property item, and an individual property item is identified by a numerical constant called a property tag. You can store and retrieve metadata by calling the **GdipSetPropertyItem** and **GdipGetPropertyItem** functions, and you don't have to be concerned with the details of how a particular file format stores that metadata.

The following topics list and describe the property items supported by GDI+. The descriptions are brief and general so that they apply to a variety of file formats. For a detailed description of how a property item is used in a particular file format, see the specification for that file format. For links to several file format specifications, see [Image File Format Specifications](https://learn.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-image-file-format-specifications).For more information about metadata, see [Reading and Writing Metadata](https://learn.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-reading-and-writing-metadata-use) and [Image Property Tag Type Constants](https://learn.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-image-property-tag-type-constants).

The following table lists the Microsoft Windows GDI+ image property tags in numerical order.

| Property Tag | Value | 
| ------------ | ----- |
| PropertyTagGpsVer | &h0000 |
| PropertyTagGpsLatitudeRef | &h0001 |
| PropertyTagGpsLatitude | &h0002 |
| PropertyTagGpsLongitudeRef | &h0003 |
| PropertyTagGpsLongitude | &h0004 |
| PropertyTagGpsAltitudeRef | &h0005 |
| PropertyTagGpsAltitude | &h0006 |
| PropertyTagGpsGpsTime | &h0007 |
| PropertyTagGpsGpsSatellites | &h0008 |
| PropertyTagGpsGpsStatus | &h0009 |
| PropertyTagGpsGpsMeasureMode | &h000A |
| PropertyTagGpsGpsDop | &h000B |
| PropertyTagGpsSpeedRef | &h000C |
| PropertyTagGpsSpeed | &h000D |
| PropertyTagGpsTrackRef | &h000E |
| PropertyTagGpsTrack | &h000F |
| PropertyTagGpsImgDirRef | &h0010 |
| PropertyTagGpsImgDir | &h0011 |
| PropertyTagGpsMapDatum | &h0012 |
| PropertyTagGpsDestLatRef | &h0013 |
| PropertyTagGpsDestLat | &h0014 |
| PropertyTagGpsDestLongRef | &h0015 |
| PropertyTagGpsDestLong | &h0016 |
| PropertyTagGpsDestBearRef | &h0017 |
| PropertyTagGpsDestBear | &h0018 |
| PropertyTagGpsDestDistRef | &h0019 |
| PropertyTagGpsDestDist | &h001A |
| PropertyTagNewSubfileType | &h00FE |
| PropertyTagSubfileType | &h00FF |
| PropertyTagImageWidth | &h0100 |
| PropertyTagImageHeight | &h0101 |
| PropertyTagBitsPerSample | &h0102 |
| PropertyTagCompression | &h0103 |
| PropertyTagPhotometricInterp | &h0106 |
| PropertyTagThreshHolding | &h0107 |
| PropertyTagCellWidth | &h0108 |
| PropertyTagCellHeight | &h0109 |
| PropertyTagFillOrder | &h010A |
| PropertyTagDocumentName | &h010D |
| PropertyTagImageDescription | &h010E |
| PropertyTagEquipMake | &h010F |
| PropertyTagEquipModel | &h0110 |
| PropertyTagStripOffsets | &h0111 |
| PropertyTagOrientation | &h0112 |
| PropertyTagSamplesPerPixel | &h0115 |
| PropertyTagRowsPerStrip | &h0116 |
| PropertyTagStripBytesCount | &h0117 |
| PropertyTagMinSampleValue | &h0118 |
| PropertyTagMaxSampleValue | &h0119 |
| PropertyTagXResolution | &h011A |
| PropertyTagYResolution | &h011B |
| PropertyTagPlanarConfig | &h011C |
| PropertyTagPageName | &h011D |
| PropertyTagXPosition | &h011E |
| PropertyTagYPosition | &h011F |
| PropertyTagFreeOffset | &h0120 |
| PropertyTagFreeByteCounts | &h0121 |
| PropertyTagGrayResponseUnit | &h0122 |
| PropertyTagGrayResponseCurve | &h0123 |
| PropertyTagT4Option | &h0124 |
| PropertyTagT6Option | &h0125 |
| PropertyTagResolutionUnit | &h0128 |
| PropertyTagPageNumber | &h0129 |
| PropertyTagTransferFunction | &h012D |
| PropertyTagSoftwareUsed | &h0131 |
| PropertyTagDateTime | &h0132 |
| PropertyTagArtist | &h013B |
| PropertyTagHostComputer | &h013C |
| PropertyTagPredictor | &h013D |
| PropertyTagWhitePoint | &h013E |
| PropertyTagPrimaryChromaticities | &h013F |
| PropertyTagColorMap | &h0140 |
| PropertyTagHalftoneHints | &h0141 |
| PropertyTagTileWidth | &h0142 |
| PropertyTagTileLength | &h0143 |
| PropertyTagTileOffset | &h0144 |
| PropertyTagTileByteCounts | &h0145 |
| PropertyTagInkSet | &h014C |
| PropertyTagInkNames | &h014D |
| PropertyTagNumberOfInks | &h014E |
| PropertyTagDotRange | &h0150 |
| PropertyTagTargetPrinter | &h0151 |
| PropertyTagExtraSamples | &h0152 |
| PropertyTagSampleFormat | &h0153 |
| PropertyTagSMinSampleValue | &h0154 |
| PropertyTagSMaxSampleValue | &h0155 |
| PropertyTagTransferRange | &h0156 |
| PropertyTagJPEGProc | &h0200 |
| PropertyTagJPEGInterFormat | &h0201 |
| PropertyTagJPEGInterLength | &h0202 |
| PropertyTagJPEGRestartInterval | &h0203 |
| PropertyTagJPEGLosslessPredictors | &h0205 |
| PropertyTagJPEGPointTransforms | &h0206 |
| PropertyTagJPEGQTables | &h0207 |
| PropertyTagJPEGDCTables | &h0208 |
| PropertyTagJPEGACTables | &h0209 |
| PropertyTagYCbCrCoefficients | &h0211 |
| PropertyTagYCbCrSubsampling | &h0212 |
| PropertyTagYCbCrPositioning | &h0213 |
| PropertyTagREFBlackWhite | &h0214 |
| PropertyTagGamma | &h0301 |
| PropertyTagICCProfileDescriptor | &h0302 |
| PropertyTagSRGBRenderingIntent | &h0303 |
| PropertyTagImageTitle | &h0320 |
| PropertyTagResolutionXUnit | &h5001 |
| PropertyTagResolutionYUnit | &h5002 |
| PropertyTagResolutionXLengthUnit | &h5003 |
| PropertyTagResolutionYLengthUnit | &h5004 |
| PropertyTagPrintFlags | &h5005 |
| PropertyTagPrintFlagsVersion | &h5006 |
| PropertyTagPrintFlagsCrop | &h5007 |
| PropertyTagPrintFlagsBleedWidth | &h5008 |
| PropertyTagPrintFlagsBleedWidthScale | &h5009 |
| PropertyTagHalftoneLPI | &h500A |
| PropertyTagHalftoneLPIUnit | &h500B |
| PropertyTagHalftoneDegree | &h500C |
| PropertyTagHalftoneShape | &h500D |
| PropertyTagHalftoneMisc | &h500E |
| PropertyTagHalftoneScreen | &h500F |
| PropertyTagJPEGQuality | &h5010 |
| PropertyTagGridSize | &h5011 |
| PropertyTagThumbnailFormat | &h5012 |
| PropertyTagThumbnailWidth | &h5013 |
| PropertyTagThumbnailHeight | &h5014 |
| PropertyTagThumbnailColorDepth | &h5015 |
| PropertyTagThumbnailPlanes | &h5016 |
| PropertyTagThumbnailRawBytes | &h5017 |
| PropertyTagThumbnailSize | &h5018 |
| PropertyTagThumbnailCompressedSize | &h5019 |
| PropertyTagColorTransferFunction | &h501A |
| PropertyTagThumbnailData | &h501B |
| PropertyTagThumbnailImageWidth | &h5020 |
| PropertyTagThumbnailImageHeight | &h5021 |
| PropertyTagThumbnailBitsPerSample | &h5022 |
| PropertyTagThumbnailCompression | &h5023 |
| PropertyTagThumbnailPhotometricInterp | &h5024 |
| PropertyTagThumbnailImageDescription | &h5025 |
| PropertyTagThumbnailEquipMake | &h5026 |
| PropertyTagThumbnailEquipModel | &h5027 |
| PropertyTagThumbnailStripOffsets | &h5028 |
| PropertyTagThumbnailOrientation | &h5029 |
| PropertyTagThumbnailSamplesPerPixel | &h502A |
| PropertyTagThumbnailRowsPerStrip | &h502B |
| PropertyTagThumbnailStripBytesCount | &h502C |
| PropertyTagThumbnailResolutionX | &h502D |
| PropertyTagThumbnailResolutionY | &h502E |
| PropertyTagThumbnailPlanarConfig | &h502F |
| PropertyTagThumbnailResolutionUnit | &h5030 |
| PropertyTagThumbnailTransferFunction | &h5031 |
| PropertyTagThumbnailSoftwareUsed | &h5032 |
| PropertyTagThumbnailDateTime | &h5033 |
| PropertyTagThumbnailArtist | &h5034 |
| PropertyTagThumbnailWhitePoint | &h5035 |
| PropertyTagThumbnailPrimaryChromaticities | &h5036 |
| PropertyTagThumbnailYCbCrCoefficients | &h5037 |
| PropertyTagThumbnailYCbCrSubsampling | &h5038 |
| PropertyTagThumbnailYCbCrPositioning | &h5039 |
| PropertyTagThumbnailRefBlackWhite | &h503A |
| PropertyTagThumbnailCopyRight | &h503B |
| PropertyTagLuminanceTable | &h5090 |
| PropertyTagChrominanceTable | &h5091 |
| PropertyTagFrameDelay | &h5100 |
| PropertyTagLoopCount | &h5101 |
| PropertyTagGlobalPalette | &h5102 |
| PropertyTagIndexBackground | &h5103 |
| PropertyTagIndexTransparent | &h5104 |
| PropertyTagPixelUnit | &h5110 |
| PropertyTagPixelPerUnitX | &h5111 |
| PropertyTagPixelPerUnitY | &h5112 |
| PropertyTagPaletteHistogram | &h5113 |
| PropertyTagCopyright | &h8298 |
| PropertyTagExifExposureTime | &h829A |
| PropertyTagExifFNumber | &h829D |
| PropertyTagExifIFD | &h8769 |
| PropertyTagICCProfile | &h8773 |
| PropertyTagExifExposureProg | &h8822 |
| PropertyTagExifSpectralSense | &h8824 |
| PropertyTagGpsIFD | &h8825 |
| PropertyTagExifISOSpeed | &h8827 |
| PropertyTagExifOECF | &h8828 |
| PropertyTagExifVer | &h9000 |
| PropertyTagExifDTOrig | &h9003 |
| PropertyTagExifDTDigitized | &h9004 |
| PropertyTagExifCompConfig | &h9101 |
| PropertyTagExifCompBPP | &h9102 |
| PropertyTagExifShutterSpeed | &h9201 |
| PropertyTagExifAperture | &h9202 |
| PropertyTagExifBrightness | &h9203 |
| PropertyTagExifExposureBias | &h9204 |
| PropertyTagExifMaxAperture | &h9205 |
| PropertyTagExifSubjectDist | &h9206 |
| PropertyTagExifMeteringMode | &h9207 |
| PropertyTagExifLightSource | &h9208 |
| PropertyTagExifFlash | &h9209 |
| PropertyTagExifFocalLength | &h920A |
| PropertyTagExifMakerNote | &h927C |
| PropertyTagExifUserComment | &h9286 |
| PropertyTagExifDTSubsec | &h9290 |
| PropertyTagExifDTOrigSS | &h9291 |
| PropertyTagExifDTDigSS | &h9292 |
| PropertyTagExifFPXVer | &hA000 |
| PropertyTagExifColorSpace | &hA001 |
| PropertyTagExifPixXDim | &hA002 |
| PropertyTagExifPixYDim | &hA003 |
| PropertyTagExifRelatedWav | &hA004 |
| PropertyTagExifInterop | &hA005 |
| PropertyTagExifFlashEnergy | &hA20B |
| PropertyTagExifSpatialFR | &hA20C |
| PropertyTagExifFocalXRes | &hA20E |
| PropertyTagExifFocalYRes | &hA20F |
| PropertyTagExifFocalResUnit | &hA210 |
| PropertyTagExifSubjectLoc | &hA214 |
| PropertyTagExifExposureIndex | &hA215 |
| PropertyTagExifSensingMethod | &hA217 |
| PropertyTagExifFileSource | &hA300 |
| PropertyTagExifSceneType | &hA301 |
| PropertyTagExifCfaPattern | &hA302 |

---

# Property Item Descrptions

Descriptions of the property items supported by Windows GDI+.

[Property item descriptions](https://learn.microsoft.com/en-us/windows/win32/gdiplus/-gdiplus-constant-property-item-descriptions)

---

# Image File Format Specifications

The following table lists several documents that give detailed descriptions of how metadata is stored in image files.

| Document | Location |
| -------- | -------- |
| Tagged Image File Format (TIFF) Specification | https://www.loc.gov/preservation/digital/formats/fdd/fdd000022.shtml |
| Exif Specification | 	https://www.loc.gov/preservation/digital/formats/fdd/fdd000146.shtml |
| Portable Network Graphics (PNG) Specification | https://www.w3.org/TR/2003/REC-PNG-20031110/ |
| ISO/IEC 10918-1 (JPEG) |	https://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=18902 |
| ICC.1:1998-09 | https://www.color.org |

---
