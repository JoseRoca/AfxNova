' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxGdiPlus.bi
' Purpose:  GDI+ Flat API header
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2025 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#pragma once

#include once "crt/stddef.bi"
#include once "crt/math.bi"
#include once "win/windef.bi"
#include once "win/wingdi.bi"
#include once "win/basetyps.bi"
#include once "win/ddraw.bi"

extern "Windows" LIB "gdiplus.dll"

' #define __GDIPLUS_H
TYPE REAL AS SINGLE
TYPE INT16 AS SHORT
TYPE UINT16 AS WORD

'#define __GDIPLUS_ENUMS_H

type BrushType as long
enum
	BrushTypeSolidColor = 0
	BrushTypeHatchFill = 1
	BrushTypeTextureFill = 2
	BrushTypePathGradient = 3
	BrushTypeLinearGradient = 4
end enum

type CombineMode as long
enum
	CombineModeReplace = 0
	CombineModeIntersect = 1
	CombineModeUnion = 2
	CombineModeXor = 3
	CombineModeExclude = 4
	CombineModeComplement = 5
end enum

type CompositingMode as long
enum
	CompositingModeSourceOver = 0
	CompositingModeSourceCopy = 1
end enum

type CompositingQuality as long
enum
	CompositingQualityDefault = 0
	CompositingQualityHighSpeed = 1
	CompositingQualityHighQuality = 2
	CompositingQualityGammaCorrected = 3
	CompositingQualityAssumeLinear = 4
end enum

type CoordinateSpace as long
enum
	CoordinateSpaceWorld = 0
	CoordinateSpacePage = 1
	CoordinateSpaceDevice = 2
end enum

type CustomLineCapType as long
enum
	CustomLineCapTypeDefault = 0
	CustomLineCapTypeAdjustableArrow = 1
end enum

type DashCap as long
enum
	DashCapFlat = 0
	DashCapRound = 2
	DashCapTriangle = 3
end enum

type DashStyle as long
enum
	DashStyleSolid = 0
	DashStyleDash = 1
	DashStyleDot = 2
	DashStyleDashDot = 3
	DashStyleDashDotDot = 4
	DashStyleCustom = 5
end enum

type DitherType as long
enum
	DitherTypeNone = 0
	DitherTypeSolid = 1
	DitherTypeOrdered4x4 = 2
	DitherTypeOrdered8x8 = 3
	DitherTypeOrdered16x16 = 4
	DitherTypeOrdered91x91 = 5
	DitherTypeSpiral4x4 = 6
	DitherTypeSpiral8x8 = 7
	DitherTypeDualSpiral4x4 = 8
	DitherTypeDualSpiral8x8 = 9
	DitherTypeErrorDiffusion = 10
end enum

type DriverStringOptions as long
enum
	DriverStringOptionsCmapLookup = 1
	DriverStringOptionsVertical = 2
	DriverStringOptionsRealizedAdvance = 4
	DriverStringOptionsLimitSubpixel = 8
end enum

#define GDIP_WMF_RECORD_TO_EMFPLUS(meta) ((meta) or &h10000)
const GDIP_EMFPLUS_RECORD_BASE = &h4000

type EmfPlusRecordType as long
enum
	WmfRecordTypeSetBkColor = &h0201 or &h10000
	WmfRecordTypeSetBkMode = &h0102 or &h10000
	WmfRecordTypeSetMapMode = &h0103 or &h10000
	WmfRecordTypeSetROP2 = &h0104 or &h10000
	WmfRecordTypeSetRelAbs = &h0105 or &h10000
	WmfRecordTypeSetPolyFillMode = &h0106 or &h10000
	WmfRecordTypeSetStretchBltMode = &h0107 or &h10000
	WmfRecordTypeSetTextCharExtra = &h0108 or &h10000
	WmfRecordTypeSetTextColor = &h0209 or &h10000
	WmfRecordTypeSetTextJustification = &h020A or &h10000
	WmfRecordTypeSetWindowOrg = &h020B or &h10000
	WmfRecordTypeSetWindowExt = &h020C or &h10000
	WmfRecordTypeSetViewportOrg = &h020D or &h10000
	WmfRecordTypeSetViewportExt = &h020E or &h10000
	WmfRecordTypeOffsetWindowOrg = &h020F or &h10000
	WmfRecordTypeScaleWindowExt = &h0410 or &h10000
	WmfRecordTypeOffsetViewportOrg = &h0211 or &h10000
	WmfRecordTypeScaleViewportExt = &h0412 or &h10000
	WmfRecordTypeLineTo = &h0213 or &h10000
	WmfRecordTypeMoveTo = &h0214 or &h10000
	WmfRecordTypeExcludeClipRect = &h0415 or &h10000
	WmfRecordTypeIntersectClipRect = &h0416 or &h10000
	WmfRecordTypeArc = &h0817 or &h10000
	WmfRecordTypeEllipse = &h0418 or &h10000
	WmfRecordTypeFloodFill = &h0419 or &h10000
	WmfRecordTypePie = &h081A or &h10000
	WmfRecordTypeRectangle = &h041B or &h10000
	WmfRecordTypeRoundRect = &h061C or &h10000
	WmfRecordTypePatBlt = &h061D or &h10000
	WmfRecordTypeSaveDC = &h001E or &h10000
	WmfRecordTypeSetPixel = &h041F or &h10000
	WmfRecordTypeOffsetClipRgn = &h0220 or &h10000
	WmfRecordTypeTextOut = &h0521 or &h10000
	WmfRecordTypeBitBlt = &h0922 or &h10000
	WmfRecordTypeStretchBlt = &h0B23 or &h10000
	WmfRecordTypePolygon = &h0324 or &h10000
	WmfRecordTypePolyline = &h0325 or &h10000
	WmfRecordTypeEscape = &h0626 or &h10000
	WmfRecordTypeRestoreDC = &h0127 or &h10000
	WmfRecordTypeFillRegion = &h0228 or &h10000
	WmfRecordTypeFrameRegion = &h0429 or &h10000
	WmfRecordTypeInvertRegion = &h012A or &h10000
	WmfRecordTypePaintRegion = &h012B or &h10000
	WmfRecordTypeSelectClipRegion = &h012C or &h10000
	WmfRecordTypeSelectObject = &h012D or &h10000
	WmfRecordTypeSetTextAlign = &h012E or &h10000
	WmfRecordTypeDrawText = &h062F or &h10000
	WmfRecordTypeChord = &h0830 or &h10000
	WmfRecordTypeSetMapperFlags = &h0231 or &h10000
	WmfRecordTypeExtTextOut = &h0a32 or &h10000
	WmfRecordTypeSetDIBToDev = &h0d33 or &h10000
	WmfRecordTypeSelectPalette = &h0234 or &h10000
	WmfRecordTypeRealizePalette = &h0035 or &h10000
	WmfRecordTypeAnimatePalette = &h0436 or &h10000
	WmfRecordTypeSetPalEntries = &h0037 or &h10000
	WmfRecordTypePolyPolygon = &h0538 or &h10000
	WmfRecordTypeResizePalette = &h0139 or &h10000
	WmfRecordTypeDIBBitBlt = &h0940 or &h10000
	WmfRecordTypeDIBStretchBlt = &h0b41 or &h10000
	WmfRecordTypeDIBCreatePatternBrush = &h0142 or &h10000
	WmfRecordTypeStretchDIB = &h0f43 or &h10000
	WmfRecordTypeExtFloodFill = &h0548 or &h10000
	WmfRecordTypeSetLayout = &h0149 or &h10000
	WmfRecordTypeResetDC = &h014C or &h10000
	WmfRecordTypeStartDoc = &h014D or &h10000
	WmfRecordTypeStartPage = &h004F or &h10000
	WmfRecordTypeEndPage = &h0050 or &h10000
	WmfRecordTypeAbortDoc = &h0052 or &h10000
	WmfRecordTypeEndDoc = &h005E or &h10000
	WmfRecordTypeDeleteObject = &h01f0 or &h10000
	WmfRecordTypeCreatePalette = &h00f7 or &h10000
	WmfRecordTypeCreateBrush = &h00F8 or &h10000
	WmfRecordTypeCreatePatternBrush = &h01F9 or &h10000
	WmfRecordTypeCreatePenIndirect = &h02FA or &h10000
	WmfRecordTypeCreateFontIndirect = &h02FB or &h10000
	WmfRecordTypeCreateBrushIndirect = &h02FC or &h10000
	WmfRecordTypeCreateBitmapIndirect = &h02FD or &h10000
	WmfRecordTypeCreateBitmap = &h06FE or &h10000
	WmfRecordTypeCreateRegion = &h06FF or &h10000
	EmfRecordTypeHeader = 1
	EmfRecordTypePolyBezier = 2
	EmfRecordTypePolygon = 3
	EmfRecordTypePolyline = 4
	EmfRecordTypePolyBezierTo = 5
	EmfRecordTypePolyLineTo = 6
	EmfRecordTypePolyPolyline = 7
	EmfRecordTypePolyPolygon = 8
	EmfRecordTypeSetWindowExtEx = 9
	EmfRecordTypeSetWindowOrgEx = 10
	EmfRecordTypeSetViewportExtEx = 11
	EmfRecordTypeSetViewportOrgEx = 12
	EmfRecordTypeSetBrushOrgEx = 13
	EmfRecordTypeEOF = 14
	EmfRecordTypeSetPixelV = 15
	EmfRecordTypeSetMapperFlags = 16
	EmfRecordTypeSetMapMode = 17
	EmfRecordTypeSetBkMode = 18
	EmfRecordTypeSetPolyFillMode = 19
	EmfRecordTypeSetROP2 = 20
	EmfRecordTypeSetStretchBltMode = 21
	EmfRecordTypeSetTextAlign = 22
	EmfRecordTypeSetColorAdjustment = 23
	EmfRecordTypeSetTextColor = 24
	EmfRecordTypeSetBkColor = 25
	EmfRecordTypeOffsetClipRgn = 26
	EmfRecordTypeMoveToEx = 27
	EmfRecordTypeSetMetaRgn = 28
	EmfRecordTypeExcludeClipRect = 29
	EmfRecordTypeIntersectClipRect = 30
	EmfRecordTypeScaleViewportExtEx = 31
	EmfRecordTypeScaleWindowExtEx = 32
	EmfRecordTypeSaveDC = 33
	EmfRecordTypeRestoreDC = 34
	EmfRecordTypeSetWorldTransform = 35
	EmfRecordTypeModifyWorldTransform = 36
	EmfRecordTypeSelectObject = 37
	EmfRecordTypeCreatePen = 38
	EmfRecordTypeCreateBrushIndirect = 39
	EmfRecordTypeDeleteObject = 40
	EmfRecordTypeAngleArc = 41
	EmfRecordTypeEllipse = 42
	EmfRecordTypeRectangle = 43
	EmfRecordTypeRoundRect = 44
	EmfRecordTypeArc = 45
	EmfRecordTypeChord = 46
	EmfRecordTypePie = 47
	EmfRecordTypeSelectPalette = 48
	EmfRecordTypeCreatePalette = 49
	EmfRecordTypeSetPaletteEntries = 50
	EmfRecordTypeResizePalette = 51
	EmfRecordTypeRealizePalette = 52
	EmfRecordTypeExtFloodFill = 53
	EmfRecordTypeLineTo = 54
	EmfRecordTypeArcTo = 55
	EmfRecordTypePolyDraw = 56
	EmfRecordTypeSetArcDirection = 57
	EmfRecordTypeSetMiterLimit = 58
	EmfRecordTypeBeginPath = 59
	EmfRecordTypeEndPath = 60
	EmfRecordTypeCloseFigure = 61
	EmfRecordTypeFillPath = 62
	EmfRecordTypeStrokeAndFillPath = 63
	EmfRecordTypeStrokePath = 64
	EmfRecordTypeFlattenPath = 65
	EmfRecordTypeWidenPath = 66
	EmfRecordTypeSelectClipPath = 67
	EmfRecordTypeAbortPath = 68
	EmfRecordTypeReserved_069 = 69
	EmfRecordTypeGdiComment = 70
	EmfRecordTypeFillRgn = 71
	EmfRecordTypeFrameRgn = 72
	EmfRecordTypeInvertRgn = 73
	EmfRecordTypePaintRgn = 74
	EmfRecordTypeExtSelectClipRgn = 75
	EmfRecordTypeBitBlt = 76
	EmfRecordTypeStretchBlt = 77
	EmfRecordTypeMaskBlt = 78
	EmfRecordTypePlgBlt = 79
	EmfRecordTypeSetDIBitsToDevice = 80
	EmfRecordTypeStretchDIBits = 81
	EmfRecordTypeExtCreateFontIndirect = 82
	EmfRecordTypeExtTextOutA = 83
	EmfRecordTypeExtTextOutW = 84
	EmfRecordTypePolyBezier16 = 85
	EmfRecordTypePolygon16 = 86
	EmfRecordTypePolyline16 = 87
	EmfRecordTypePolyBezierTo16 = 88
	EmfRecordTypePolylineTo16 = 89
	EmfRecordTypePolyPolyline16 = 90
	EmfRecordTypePolyPolygon16 = 91
	EmfRecordTypePolyDraw16 = 92
	EmfRecordTypeCreateMonoBrush = 93
	EmfRecordTypeCreateDIBPatternBrushPt = 94
	EmfRecordTypeExtCreatePen = 95
	EmfRecordTypePolyTextOutA = 96
	EmfRecordTypePolyTextOutW = 97
	EmfRecordTypeSetICMMode = 98
	EmfRecordTypeCreateColorSpace = 99
	EmfRecordTypeSetColorSpace = 100
	EmfRecordTypeDeleteColorSpace = 101
	EmfRecordTypeGLSRecord = 102
	EmfRecordTypeGLSBoundedRecord = 103
	EmfRecordTypePixelFormat = 104
	EmfRecordTypeDrawEscape = 105
	EmfRecordTypeExtEscape = 106
	EmfRecordTypeStartDoc = 107
	EmfRecordTypeSmallTextOut = 108
	EmfRecordTypeForceUFIMapping = 109
	EmfRecordTypeNamedEscape = 110
	EmfRecordTypeColorCorrectPalette = 111
	EmfRecordTypeSetICMProfileA = 112
	EmfRecordTypeSetICMProfileW = 113
	EmfRecordTypeAlphaBlend = 114
	EmfRecordTypeSetLayout = 115
	EmfRecordTypeTransparentBlt = 116
	EmfRecordTypeReserved_117 = 117
	EmfRecordTypeGradientFill = 118
	EmfRecordTypeSetLinkedUFIs = 119
	EmfRecordTypeSetTextJustification = 120
	EmfRecordTypeColorMatchToTargetW = 121
	EmfRecordTypeCreateColorSpaceW = 122
	EmfRecordTypeMax = 122
	EmfRecordTypeMin = 1
	EmfPlusRecordTypeInvalid = &h4000
	EmfPlusRecordTypeHeader
	EmfPlusRecordTypeEndOfFile
	EmfPlusRecordTypeComment
	EmfPlusRecordTypeGetDC
	EmfPlusRecordTypeMultiFormatStart
	EmfPlusRecordTypeMultiFormatSection
	EmfPlusRecordTypeMultiFormatEnd
	EmfPlusRecordTypeObject
	EmfPlusRecordTypeClear
	EmfPlusRecordTypeFillRects
	EmfPlusRecordTypeDrawRects
	EmfPlusRecordTypeFillPolygon
	EmfPlusRecordTypeDrawLines
	EmfPlusRecordTypeFillEllipse
	EmfPlusRecordTypeDrawEllipse
	EmfPlusRecordTypeFillPie
	EmfPlusRecordTypeDrawPie
	EmfPlusRecordTypeDrawArc
	EmfPlusRecordTypeFillRegion
	EmfPlusRecordTypeFillPath
	EmfPlusRecordTypeDrawPath
	EmfPlusRecordTypeFillClosedCurve
	EmfPlusRecordTypeDrawClosedCurve
	EmfPlusRecordTypeDrawCurve
	EmfPlusRecordTypeDrawBeziers
	EmfPlusRecordTypeDrawImage
	EmfPlusRecordTypeDrawImagePoints
	EmfPlusRecordTypeDrawString
	EmfPlusRecordTypeSetRenderingOrigin
	EmfPlusRecordTypeSetAntiAliasMode
	EmfPlusRecordTypeSetTextRenderingHint
	EmfPlusRecordTypeSetTextContrast
	EmfPlusRecordTypeSetGammaValue
	EmfPlusRecordTypeSetInterpolationMode
	EmfPlusRecordTypeSetPixelOffsetMode
	EmfPlusRecordTypeSetCompositingMode
	EmfPlusRecordTypeSetCompositingQuality
	EmfPlusRecordTypeSave
	EmfPlusRecordTypeRestore
	EmfPlusRecordTypeBeginContainer
	EmfPlusRecordTypeBeginContainerNoParams
	EmfPlusRecordTypeEndContainer
	EmfPlusRecordTypeSetWorldTransform
	EmfPlusRecordTypeResetWorldTransform
	EmfPlusRecordTypeMultiplyWorldTransform
	EmfPlusRecordTypeTranslateWorldTransform
	EmfPlusRecordTypeScaleWorldTransform
	EmfPlusRecordTypeRotateWorldTransform
	EmfPlusRecordTypeSetPageTransform
	EmfPlusRecordTypeResetClip
	EmfPlusRecordTypeSetClipRect
	EmfPlusRecordTypeSetClipPath
	EmfPlusRecordTypeSetClipRegion
	EmfPlusRecordTypeOffsetClip
	EmfPlusRecordTypeDrawDriverString
	EmfPlusRecordTypeStrokeFillPath
	EmfPlusRecordTypeSerializableObject
	EmfPlusRecordTypeSetTSGraphics
	EmfPlusRecordTypeSetTSClip
	EmfPlusRecordTotal
	EmfPlusRecordTypeMax = EmfPlusRecordTotal - 1
	EmfPlusRecordTypeMin = EmfPlusRecordTypeHeader
end enum

type EmfToWmfBitsFlags as long
enum
	EmfToWmfBitsFlagsDefault = 0
	EmfToWmfBitsFlagsEmbedEmf = 1
	EmfToWmfBitsFlagsIncludePlaceable = 2
	EmfToWmfBitsFlagsNoXORClip = 4
end enum

type EmfType as long
enum
	EmfTypeEmfOnly = 3
	EmfTypeEmfPlusOnly = 4
	EmfTypeEmfPlusDual = 5
end enum

type EncoderParameterValueType as long
enum
	EncoderParameterValueTypeByte = 1
	EncoderParameterValueTypeASCII = 2
	EncoderParameterValueTypeShort = 3
	EncoderParameterValueTypeLong = 4
	EncoderParameterValueTypeRational = 5
	EncoderParameterValueTypeLongRange = 6
	EncoderParameterValueTypeUndefined = 7
	EncoderParameterValueTypeRationalRange = 8
	EncoderParameterValueTypePointer = 9
end enum

type EncoderValue as long
enum
	EncoderValueColorTypeCMYK = 0
	EncoderValueColorTypeYCCK = 1
	EncoderValueCompressionLZW = 2
	EncoderValueCompressionCCITT3 = 3
	EncoderValueCompressionCCITT4 = 4
	EncoderValueCompressionRle = 5
	EncoderValueCompressionNone = 6
	EncoderValueScanMethodInterlaced = 7
	EncoderValueScanMethodNonInterlaced = 8
	EncoderValueVersionGif87 = 9
	EncoderValueVersionGif89 = 10
	EncoderValueRenderProgressive = 11
	EncoderValueRenderNonProgressive = 12
	EncoderValueTransformRotate90 = 13
	EncoderValueTransformRotate180 = 14
	EncoderValueTransformRotate270 = 15
	EncoderValueTransformFlipHorizontal = 16
	EncoderValueTransformFlipVertical = 17
	EncoderValueMultiFrame = 18
	EncoderValueLastFrame = 19
	EncoderValueFlush = 20
	EncoderValueFrameDimensionTime = 21
	EncoderValueFrameDimensionResolution = 22
	EncoderValueFrameDimensionPage = 23
   EncoderValueColorTypeGray = 24
   EncoderValueColorTypeRGB = 25
end enum

type FillMode as long
enum
	FillModeAlternate = 0
	FillModeWinding = 1
end enum

type FlushIntention as long
enum
	FlushIntentionFlush = 0
	FlushIntentionSync = 1
end enum

type FontStyle as long
enum
	FontStyleRegular = 0
	FontStyleBold = 1
	FontStyleItalic = 2
	FontStyleBoldItalic = 3
	FontStyleUnderline = 4
	FontStyleStrikeout = 8
end enum

type HatchStyle as long
enum
	HatchStyleHorizontal = 0
	HatchStyleVertical = 1
	HatchStyleForwardDiagonal = 2
	HatchStyleBackwardDiagonal = 3
	HatchStyleCross = 4
	HatchStyleLargeGrid = 4
	HatchStyleDiagonalCross = 5
	HatchStyle05Percent = 6
	HatchStyle10Percent = 7
	HatchStyle20Percent = 8
	HatchStyle25Percent = 9
	HatchStyle30Percent = 10
	HatchStyle40Percent = 11
	HatchStyle50Percent = 12
	HatchStyle60Percent = 13
	HatchStyle70Percent = 14
	HatchStyle75Percent = 15
	HatchStyle80Percent = 16
	HatchStyle90Percent = 17
	HatchStyleLightDownwardDiagonal = 18
	HatchStyleLightUpwardDiagonal = 19
	HatchStyleDarkDownwardDiagonal = 20
	HatchStyleDarkUpwardDiagonal = 21
	HatchStyleWideDownwardDiagonal = 22
	HatchStyleWideUpwardDiagonal = 23
	HatchStyleLightVertical = 24
	HatchStyleLightHorizontal = 25
	HatchStyleNarrowVertical = 26
	HatchStyleNarrowHorizontal = 27
	HatchStyleDarkVertical = 28
	HatchStyleDarkHorizontal = 29
	HatchStyleDashedDownwardDiagonal = 30
	HatchStyleDashedUpwardDiagonal = 31
	HatchStyleDashedHorizontal = 32
	HatchStyleDashedVertical = 33
	HatchStyleSmallConfetti = 34
	HatchStyleLargeConfetti = 35
	HatchStyleZigZag = 36
	HatchStyleWave = 37
	HatchStyleDiagonalBrick = 38
	HatchStyleHorizontalBrick = 39
	HatchStyleWeave = 40
	HatchStylePlaid = 41
	HatchStyleDivot = 42
	HatchStyleDottedGrid = 43
	HatchStyleDottedDiamond = 44
	HatchStyleShingle = 45
	HatchStyleTrellis = 46
	HatchStyleSphere = 47
	HatchStyleSmallGrid = 48
	HatchStyleSmallCheckerBoard = 49
	HatchStyleLargeCheckerBoard = 50
	HatchStyleOutlinedDiamond = 51
	HatchStyleSolidDiamond = 52
	HatchStyleTotal = 53
	HatchStyleMin = HatchStyleHorizontal
	HatchStyleMax = HatchStyleTotal - 1
end enum

type HotkeyPrefix as long
enum
	HotkeyPrefixNone = 0
	HotkeyPrefixShow = 1
	HotkeyPrefixHide = 2
end enum

type ImageType as long
enum
	ImageTypeUnknown = 0
	ImageTypeBitmap = 1
	ImageTypeMetafile = 2
end enum

type InterpolationMode as long
enum
	InterpolationModeInvalid = -1
	InterpolationModeDefault = 0
	InterpolationModeLowQuality = 1
	InterpolationModeHighQuality = 2
	InterpolationModeBilinear = 3
	InterpolationModeBicubic = 4
	InterpolationModeNearestNeighbor = 5
	InterpolationModeHighQualityBilinear = 6
	InterpolationModeHighQualityBicubic = 7
end enum

type LinearGradientMode as long
enum
	LinearGradientModeHorizontal = 0
	LinearGradientModeVertical = 1
	LinearGradientModeForwardDiagonal = 2
	LinearGradientModeBackwardDiagonal = 3
end enum

type LineCap as long
enum
	LineCapFlat = 0
	LineCapSquare = 1
	LineCapRound = 2
	LineCapTriangle = 3
	LineCapNoAnchor = &H10
	LineCapSquareAnchor = &H11
	LineCapRoundAnchor = &H12
	LineCapDiamondAnchor = &H13
	LineCapArrowAnchor = &H14
	LineCapCustom = &HFF
   LineCapAnchorMask = &HF0
end enum

type LineJoin as long
enum
	LineJoinMiter = 0
	LineJoinBevel = 1
	LineJoinRound = 2
	LineJoinMiterClipped = 3
end enum

type MatrixOrder as long
enum
	MatrixOrderPrepend = 0
	MatrixOrderAppend = 1
end enum

type MetafileFrameUnit as long
enum
	MetafileFrameUnitPixel = 2
	MetafileFrameUnitPoint = 3
	MetafileFrameUnitInch = 4
	MetafileFrameUnitDocument = 5
	MetafileFrameUnitMillimeter = 6
	MetafileFrameUnitGdi = 7
end enum

type MetafileType as long
enum
	MetafileTypeInvalid = 0
	MetafileTypeWmf = 1
	MetafileTypeWmfPlaceable = 2
	MetafileTypeEmf = 3
	MetafileTypeEmfPlusOnly = 4
	MetafileTypeEmfPlusDual = 5
end enum

type ObjectType as long
enum
	ObjectTypeInvalid = 0
	ObjectTypeBrush = 1
	ObjectTypePen = 2
	ObjectTypePath = 3
	ObjectTypeRegion = 4
	ObjectTypeFont = 5
	ObjectTypeStringFormat = 6
	ObjectTypeImageAttributes = 7
	ObjectTypeCustomLineCap = 8
	ObjectTypeGraphics = 9
	ObjectTypeMin = ObjectTypeBrush
	ObjectTypeMax = ObjectTypeGraphics
end enum

type PathPointType as long
enum
	PathPointTypeStart = &h00
	PathPointTypeLine = &h01
	PathPointTypeBezier = &h03
	PathPointTypeBezier3 = &h03
	PathPointTypePathTypeMask = &h07
	PathPointTypePathDashMode = &h10
	PathPointTypePathMarker = &h20
	PathPointTypeCloseSubpath = &h80
end enum

type PenAlignment as long
enum
	PenAlignmentCenter = 0
	PenAlignmentInset = 1
end enum

type PenType as long
enum
	PenTypeUnknown = -1
	PenTypeSolidColor = 0
	PenTypeHatchFill = 1
	PenTypeTextureFill = 2
	PenTypePathGradient = 3
	PenTypeLinearGradient = 4
end enum

type PixelOffsetMode as long
enum
	PixelOffsetModeInvalid = -1
	PixelOffsetModeDefault = 0
	PixelOffsetModeHighSpeed = 1
	PixelOffsetModeHighQuality = 2
	PixelOffsetModeNone = 3
	PixelOffsetModeHalf = 4
end enum

type QualityMode as long
enum
	QualityModeInvalid = -1
	QualityModeDefault = 0
	QualityModeLow = 1
	QualityModeHigh = 2
end enum

type SmoothingMode as long
enum
	SmoothingModeInvalid = QualityModeInvalid
	SmoothingModeDefault = 0
	SmoothingModeHighSpeed = 1
	SmoothingModeHighQuality = 2
	SmoothingModeNone = 3
	SmoothingModeAntiAlias8x4 = 4
	SmoothingModeAntiAlias = 4
	SmoothingModeAntiAlias8x8 = 5
end enum

type StringAlignment as long
enum
	StringAlignmentNear = 0
	StringAlignmentCenter = 1
	StringAlignmentFar = 2
end enum

type StringDigitSubstitute as long
enum
	StringDigitSubstituteUser = 0
	StringDigitSubstituteNone = 1
	StringDigitSubstituteNational = 2
	StringDigitSubstituteTraditional = 3
end enum

type StringFormatFlags as long
enum
	StringFormatFlagsDirectionRightToLeft = &h00000001
	StringFormatFlagsDirectionVertical = &h00000002
	StringFormatFlagsNoFitBlackBox = &h00000004
	StringFormatFlagsDisplayFormatControl = &h00000020
	StringFormatFlagsNoFontFallback = &h00000400
	StringFormatFlagsMeasureTrailingSpaces = &h00000800
	StringFormatFlagsNoWrap = &h00001000
	StringFormatFlagsLineLimit = &h00002000
	StringFormatFlagsNoClip = &h00004000
   StringFormatFlagsBypassGDI = &h80000000
end enum

type StringTrimming as long
enum
	StringTrimmingNone = 0
	StringTrimmingCharacter = 1
	StringTrimmingWord = 2
	StringTrimmingEllipsisCharacter = 3
	StringTrimmingEllipsisWord = 4
	StringTrimmingEllipsisPath = 5
end enum

type TextRenderingHint as long
enum
	TextRenderingHintSystemDefault = 0
	TextRenderingHintSingleBitPerPixelGridFit = 1
	TextRenderingHintSingleBitPerPixel = 2
	TextRenderingHintAntiAliasGridFit = 3
	TextRenderingHintAntiAlias = 4
	TextRenderingHintClearTypeGridFit = 5
end enum

type Unit as long
enum
	UnitWorld = 0
	UnitDisplay = 1
	UnitPixel = 2
	UnitPoint = 3
	UnitInch = 4
	UnitDocument = 5
	UnitMillimeter = 6
end enum

type WarpMode as long
enum
	WarpModePerspective = 0
	WarpModeBilinear = 1
end enum

type WrapMode as long
enum
	WrapModeTile = 0
	WrapModeTileFlipX = 1
	WrapModeTileFlipY = 2
	WrapModeTileFlipXY = 3
	WrapModeClamp = 4
end enum

type GpTestControlEnum as long
enum
	TestControlForceBilinear = 0
	TestControlForceNoICM = 1
	TestControlGetBuildNumber = 2
end enum

TYPE GraphicsContainer AS DWORD
TYPE GraphicsState AS DWORD
CONST FlatnessDefault = CAST(REAL, 0.25f)

PRIVATE FUNCTION ObjectTypeIsValid (BYVAL type_ as ObjectType) AS BOOLEAN
	FUNCTION = -((type_ >= ObjectTypeMin) AND (type_ <= ObjectTypeMax))
END FUNCTION

' #define __GDIPLUS_TYPES_H
' #define WINGDIPAPI __stdcall

type Status_ as long
enum
	Ok = 0
	GenericError = 1
	InvalidParameter = 2
	OutOfMemory = 3
	ObjectBusy = 4
	InsufficientBuffer = 5
	NotImplemented = 6
	Win32Error = 7
	WrongState = 8
	Aborted = 9
	FileNotFound = 10
	ValueOverflow = 11
	AccessDenied = 12
	UnknownImageFormat = 13
	FontFamilyNotFound = 14
	FontStyleNotFound = 15
	NotTrueTypeFont = 16
	UnsupportedGdiplusVersion = 17
	GdiplusNotInitialized = 18
	PropertyNotFound = 19
	PropertyNotSupported = 20
	ProfileNotFound = 21
end enum

type GpStatus as long
enum
	StatusOk = 0
	StatusGenericError = 1
	StatusInvalidParameter = 2
	StatusOutOfMemory = 3
	StatusObjectBusy = 4
	StatusInsufficientBuffer = 5
	StatusNotImplemented = 6
	StatusWin32Error = 7
	StatusWrongState = 8
	StatusAborted = 9
	StatusFileNotFound = 10
	StatusValueOverflow = 11
	StatusAccessDenied = 12
	StatusUnknownImageFormat = 13
	StatusFontFamilyNotFound = 14
	StatusFontStyleNotFound = 15
	StatusNotTrueTypeFont = 16
	StatusUnsupportedGdiplusVersion = 17
	StatusGdiplusNotInitialized = 18
	StatusPropertyNotFound = 19
	StatusPropertyNotSupported = 20
	StatusProfileNotFound = 21
end enum

type GpSize
	Width as INT_
	Height as INT_
end type

type GpSizeF
	Width as REAL
	Height as REAL
end type

type GpPoint
	X as INT_
	Y as INT_
end type

type GpPointF
	X as REAL
	Y as REAL
end type

type GpRect
	X as INT_
	Y as INT_
	Width as INT_
	Height as INT_
end type

type GpRectF
	X as REAL
	Y as REAL
	Width as REAL
	Height as REAL
end type

type GpCharacterRange
	First as INT_
	Length as INT_
end type

type GpPathData
	Count as INT_
	Points as GpPointF ptr
	Types as BYTE ptr
end type

type DebugEventProc as any ptr
type EnumerateMetafileProc as function (BYVAL AS EmfPlusRecordType, BYVAL AS UINT, BYVAL AS UINT, BYVAL AS const UBYTE PTR, BYVAL AS any PTR) as BOOL
type DrawImageAbort as any ptr
type GetThumbnailImageAbort as any ptr

'#define __GDIPLUS_GPSTUBS_H
type GpBrushType as BrushType
type GpCombineMode as CombineMode
type GpCompositingMode as CompositingMode
type GpCompositingQuality as CompositingQuality
type GpCoordinateSpace as CoordinateSpace
type GpCustomLineCapType as CustomLineCapType
type GpDashCap as DashCap
type GpDashStyle as DashStyle
type GpDitherType as DitherType
type GpDriverStringOptions as DriverStringOptions
type GpEmfPlusRecordType as EmfPlusRecordType
type GpEmfToWmfBitsFlags as EmfToWmfBitsFlags
type GpEmfType as EmfType
type GpEncoderParameterValueType as EncoderParameterValueType
type GpEncoderValue as EncoderValue
type GpFillMode as FillMode
type GpFlushIntention as FlushIntention
type GpFontStyle as FontStyle
type GpHatchStyle as HatchStyle
type GpHotkeyPrefix as HotkeyPrefix
type GpImageType as ImageType
type GpInterpolationMode as InterpolationMode
type GpLinearGradientMode as LinearGradientMode
type GpLineCap as LineCap
type GpLineJoin as LineJoin
type GpMatrixOrder as MatrixOrder
type GpMetafileFrameUnit as MetafileFrameUnit
type GpMetafileType as MetafileType
type GpObjectType as ObjectType
type GpPathPointType as PathPointType
type GpPenAlignment as PenAlignment
type GpPenType as PenType
type GpPixelOffsetMode as PixelOffsetMode
type GpQualityMode as QualityMode
type GpSmoothingMode as SmoothingMode
type GpStringAlignment as StringAlignment
type GpStringDigitSubstitute as StringDigitSubstitute
type GpStringFormatFlags as StringFormatFlags
type GpStringTrimming as StringTrimming
type GpTextRenderingHint as TextRenderingHint
type GpUnit as Unit
type GpWarpMode as WarpMode
type GpWrapMode as WrapMode
type CGpEffect as any
type GpAdjustableArrowCap as any
type GpBitmap as any
type GpBrush as any
type GpCachedBitmap as any
type GpCustomLineCap as any
type GpFont as any
type GpFontFamily as any
type GpFontCollection as any
type GpGraphics as any
type GpHatch as any
type GpImage as any
type GpImageAttributes as any
type GpLineGradient as any
type GpMatrix as any
type GpMetafile as any
type GpPath as any
type GpPathGradient as any
type GpPathIterator as any
type GpPen as any
type GpRegion as any
type GpSolidFill as any
type GpStringFormat as any
type GpTexture as any

' #define __GDIPLUS_IMAGING_H

type ImageCodecFlags as long
enum
	ImageCodecFlagsEncoder = &h00000001
	ImageCodecFlagsDecoder = &h00000002
	ImageCodecFlagsSupportBitmap = &h00000004
	ImageCodecFlagsSupportVector = &h00000008
	ImageCodecFlagsSeekableEncode = &h00000010
	ImageCodecFlagsBlockingDecode = &h00000020
	ImageCodecFlagsBuiltin = &h00010000
	ImageCodecFlagsSystem = &h00020000
	ImageCodecFlagsUser = &h00040000
end enum

type ImageFlags as long
enum
	ImageFlagsNone = 0
	ImageFlagsScalable = &h00000001
	ImageFlagsHasAlpha = &h00000002
	ImageFlagsHasTranslucent = &h00000004
	ImageFlagsPartiallyScalable = &h00000008
	ImageFlagsColorSpaceRGB = &h00000010
	ImageFlagsColorSpaceCMYK = &h00000020
	ImageFlagsColorSpaceGRAY = &h00000040
	ImageFlagsColorSpaceYCBCR = &h00000080
	ImageFlagsColorSpaceYCCK = &h00000100
	ImageFlagsHasRealDPI = &h00001000
	ImageFlagsHasRealPixelSize = &h00002000
	ImageFlagsReadOnly = &h00010000
	ImageFlagsCaching = &h00020000
end enum

type ImageLockMode as long
enum
	ImageLockModeRead = 1
	ImageLockModeWrite = 2
	ImageLockModeUserInputBuf = 4
end enum

type ItemDataPosition as long
enum
	ItemDataPositionAfterHeader = 0
	ItemDataPositionAfterPalette = 1
	ItemDataPositionAfterBits = 2
end enum

type RotateFlipType as long
enum
	RotateNoneFlipNone = 0
	Rotate90FlipNone = 1
	Rotate180FlipNone = 2
	Rotate270FlipNone = 3
	RotateNoneFlipX = 4
	Rotate90FlipX = 5
	Rotate180FlipX = 6
	Rotate270FlipX = 7
	Rotate180FlipXY = 0
	Rotate270FlipXY = 1
	RotateNoneFlipXY = 2
	Rotate90FlipXY = 3
	Rotate180FlipY = 4
	Rotate270FlipY = 5
	RotateNoneFlipY = 6
	Rotate90FlipY = 7
end enum

type BitmapData
	Width as UINT
	Height as UINT
	Stride as INT_
	PixelFormat as INT_
	Scan0 as any ptr
	Reserved as UINT_PTR
end type

type EncoderParameter
	Guid as GUID
	NumberOfValues as ULONG
	as ULONG Type
	Value as any ptr
end type

type EncoderParameters
	Count as UINT
	Parameter(0 to 0) as EncoderParameter
end type

type ImageCodecInfo
	Clsid as CLSID
	FormatID as GUID
	CodecName as wstring ptr
	DllName as wstring ptr
	FormatDescription as wstring ptr
	FilenameExtension as wstring ptr
	MimeType as wstring ptr
	Flags as DWORD
	Version as DWORD
	SigCount as DWORD
	SigSize as DWORD
	SigPattern as UBYTE ptr
	SigMask as UBYTE ptr
end type

type ImageItemData
	Size as UINT
	Position as UINT
	Desc as any ptr
	DescSize as UINT
	Data as UINT ptr
	DataSize as UINT
	Cookie as UINT
end type

type PropertyItem
	id as PROPID
	length as ULONG
	as WORD type
	value as any ptr
end type

const PropertyTagGpsVer = cast(PROPID, &h0000)
const PropertyTagGpsLatitudeRef = cast(PROPID, &h0001)
const PropertyTagGpsLatitude = cast(PROPID, &h0002)
const PropertyTagGpsLongitudeRef = cast(PROPID, &h0003)
const PropertyTagGpsLongitude = cast(PROPID, &h0004)
const PropertyTagGpsAltitudeRef = cast(PROPID, &h0005)
const PropertyTagGpsAltitude = cast(PROPID, &h0006)
const PropertyTagGpsGpsTime = cast(PROPID, &h0007)
const PropertyTagGpsGpsSatellites = cast(PROPID, &h0008)
const PropertyTagGpsGpsStatus = cast(PROPID, &h0009)
const PropertyTagGpsGpsMeasureMode = cast(PROPID, &h000A)
const PropertyTagGpsGpsDop = cast(PROPID, &h000B)
const PropertyTagGpsSpeedRef = cast(PROPID, &h000C)
const PropertyTagGpsSpeed = cast(PROPID, &h000D)
const PropertyTagGpsTrackRef = cast(PROPID, &h000E)
const PropertyTagGpsTrack = cast(PROPID, &h000F)
const PropertyTagGpsImgDirRef = cast(PROPID, &h0010)
const PropertyTagGpsImgDir = cast(PROPID, &h0011)
const PropertyTagGpsMapDatum = cast(PROPID, &h0012)
const PropertyTagGpsDestLatRef = cast(PROPID, &h0013)
const PropertyTagGpsDestLat = cast(PROPID, &h0014)
const PropertyTagGpsDestLongRef = cast(PROPID, &h0015)
const PropertyTagGpsDestLong = cast(PROPID, &h0016)
const PropertyTagGpsDestBearRef = cast(PROPID, &h0017)
const PropertyTagGpsDestBear = cast(PROPID, &h0018)
const PropertyTagGpsDestDistRef = cast(PROPID, &h0019)
const PropertyTagGpsDestDist = cast(PROPID, &h001A)
const PropertyTagNewSubfileType = cast(PROPID, &h00FE)
const PropertyTagSubfileType = cast(PROPID, &h00FF)
const PropertyTagImageWidth = cast(PROPID, &h0100)
const PropertyTagImageHeight = cast(PROPID, &h0101)
const PropertyTagBitsPerSample = cast(PROPID, &h0102)
const PropertyTagCompression = cast(PROPID, &h0103)
const PropertyTagPhotometricInterp = cast(PROPID, &h0106)
const PropertyTagThreshHolding = cast(PROPID, &h0107)
const PropertyTagCellWidth = cast(PROPID, &h0108)
const PropertyTagCellHeight = cast(PROPID, &h0109)
const PropertyTagFillOrder = cast(PROPID, &h010A)
const PropertyTagDocumentName = cast(PROPID, &h010D)
const PropertyTagImageDescription = cast(PROPID, &h010E)
const PropertyTagEquipMake = cast(PROPID, &h010F)
const PropertyTagEquipModel = cast(PROPID, &h0110)
const PropertyTagStripOffsets = cast(PROPID, &h0111)
const PropertyTagOrientation = cast(PROPID, &h0112)
const PropertyTagSamplesPerPixel = cast(PROPID, &h0115)
const PropertyTagRowsPerStrip = cast(PROPID, &h0116)
const PropertyTagStripBytesCount = cast(PROPID, &h0117)
const PropertyTagMinSampleValue = cast(PROPID, &h0118)
const PropertyTagMaxSampleValue = cast(PROPID, &h0119)
const PropertyTagXResolution = cast(PROPID, &h011A)
const PropertyTagYResolution = cast(PROPID, &h011B)
const PropertyTagPlanarConfig = cast(PROPID, &h011C)
const PropertyTagPageName = cast(PROPID, &h011D)
const PropertyTagXPosition = cast(PROPID, &h011E)
const PropertyTagYPosition = cast(PROPID, &h011F)
const PropertyTagFreeOffset = cast(PROPID, &h0120)
const PropertyTagFreeByteCounts = cast(PROPID, &h0121)
const PropertyTagGrayResponseUnit = cast(PROPID, &h0122)
const PropertyTagGrayResponseCurve = cast(PROPID, &h0123)
const PropertyTagT4Option = cast(PROPID, &h0124)
const PropertyTagT6Option = cast(PROPID, &h0125)
const PropertyTagResolutionUnit = cast(PROPID, &h0128)
const PropertyTagPageNumber = cast(PROPID, &h0129)
const PropertyTagTransferFunction = cast(PROPID, &h012D)
const PropertyTagSoftwareUsed = cast(PROPID, &h0131)
const PropertyTagDateTime = cast(PROPID, &h0132)
const PropertyTagArtist = cast(PROPID, &h013B)
const PropertyTagHostComputer = cast(PROPID, &h013C)
const PropertyTagPredictor = cast(PROPID, &h013D)
const PropertyTagWhitePoint = cast(PROPID, &h013E)
const PropertyTagPrimaryChromaticities = cast(PROPID, &h013F)
const PropertyTagColorMap = cast(PROPID, &h0140)
const PropertyTagHalftoneHints = cast(PROPID, &h0141)
const PropertyTagTileWidth = cast(PROPID, &h0142)
const PropertyTagTileLength = cast(PROPID, &h0143)
const PropertyTagTileOffset = cast(PROPID, &h0144)
const PropertyTagTileByteCounts = cast(PROPID, &h0145)
const PropertyTagInkSet = cast(PROPID, &h014C)
const PropertyTagInkNames = cast(PROPID, &h014D)
const PropertyTagNumberOfInks = cast(PROPID, &h014E)
const PropertyTagDotRange = cast(PROPID, &h0150)
const PropertyTagTargetPrinter = cast(PROPID, &h0151)
const PropertyTagExtraSamples = cast(PROPID, &h0152)
const PropertyTagSampleFormat = cast(PROPID, &h0153)
const PropertyTagSMinSampleValue = cast(PROPID, &h0154)
const PropertyTagSMaxSampleValue = cast(PROPID, &h0155)
const PropertyTagTransferRange = cast(PROPID, &h0156)
const PropertyTagJPEGProc = cast(PROPID, &h0200)
const PropertyTagJPEGInterFormat = cast(PROPID, &h0201)
const PropertyTagJPEGInterLength = cast(PROPID, &h0202)
const PropertyTagJPEGRestartInterval = cast(PROPID, &h0203)
const PropertyTagJPEGLosslessPredictors = cast(PROPID, &h0205)
const PropertyTagJPEGPointTransforms = cast(PROPID, &h0206)
const PropertyTagJPEGQTables = cast(PROPID, &h0207)
const PropertyTagJPEGDCTables = cast(PROPID, &h0208)
const PropertyTagJPEGACTables = cast(PROPID, &h0209)
const PropertyTagYCbCrCoefficients = cast(PROPID, &h0211)
const PropertyTagYCbCrSubsampling = cast(PROPID, &h0212)
const PropertyTagYCbCrPositioning = cast(PROPID, &h0213)
const PropertyTagREFBlackWhite = cast(PROPID, &h0214)
const PropertyTagGamma = cast(PROPID, &h0301)
const PropertyTagICCProfileDescriptor = cast(PROPID, &h0302)
const PropertyTagSRGBRenderingIntent = cast(PROPID, &h0303)
const PropertyTagImageTitle = cast(PROPID, &h0320)
const PropertyTagResolutionXUnit = cast(PROPID, &h5001)
const PropertyTagResolutionYUnit = cast(PROPID, &h5002)
const PropertyTagResolutionXLengthUnit = cast(PROPID, &h5003)
const PropertyTagResolutionYLengthUnit = cast(PROPID, &h5004)
const PropertyTagPrintFlags = cast(PROPID, &h5005)
const PropertyTagPrintFlagsVersion = cast(PROPID, &h5006)
const PropertyTagPrintFlagsCrop = cast(PROPID, &h5007)
const PropertyTagPrintFlagsBleedWidth = cast(PROPID, &h5008)
const PropertyTagPrintFlagsBleedWidthScale = cast(PROPID, &h5009)
const PropertyTagHalftoneLPI = cast(PROPID, &h500A)
const PropertyTagHalftoneLPIUnit = cast(PROPID, &h500B)
const PropertyTagHalftoneDegree = cast(PROPID, &h500C)
const PropertyTagHalftoneShape = cast(PROPID, &h500D)
const PropertyTagHalftoneMisc = cast(PROPID, &h500E)
const PropertyTagHalftoneScreen = cast(PROPID, &h500F)
const PropertyTagJPEGQuality = cast(PROPID, &h5010)
const PropertyTagGridSize = cast(PROPID, &h5011)
const PropertyTagThumbnailFormat = cast(PROPID, &h5012)
const PropertyTagThumbnailWidth = cast(PROPID, &h5013)
const PropertyTagThumbnailHeight = cast(PROPID, &h5014)
const PropertyTagThumbnailColorDepth = cast(PROPID, &h5015)
const PropertyTagThumbnailPlanes = cast(PROPID, &h5016)
const PropertyTagThumbnailRawBytes = cast(PROPID, &h5017)
const PropertyTagThumbnailSize = cast(PROPID, &h5018)
const PropertyTagThumbnailCompressedSize = cast(PROPID, &h5019)
const PropertyTagColorTransferFunction = cast(PROPID, &h501A)
const PropertyTagThumbnailData = cast(PROPID, &h501B)
const PropertyTagThumbnailImageWidth = cast(PROPID, &h5020)
const PropertyTagThumbnailImageHeight = cast(PROPID, &h5021)
const PropertyTagThumbnailBitsPerSample = cast(PROPID, &h5022)
const PropertyTagThumbnailCompression = cast(PROPID, &h5023)
const PropertyTagThumbnailPhotometricInterp = cast(PROPID, &h5024)
const PropertyTagThumbnailImageDescription = cast(PROPID, &h5025)
const PropertyTagThumbnailEquipMake = cast(PROPID, &h5026)
const PropertyTagThumbnailEquipModel = cast(PROPID, &h5027)
const PropertyTagThumbnailStripOffsets = cast(PROPID, &h5028)
const PropertyTagThumbnailOrientation = cast(PROPID, &h5029)
const PropertyTagThumbnailSamplesPerPixel = cast(PROPID, &h502A)
const PropertyTagThumbnailRowsPerStrip = cast(PROPID, &h502B)
const PropertyTagThumbnailStripBytesCount = cast(PROPID, &h502C)
const PropertyTagThumbnailResolutionX = cast(PROPID, &h502D)
const PropertyTagThumbnailResolutionY = cast(PROPID, &h502E)
const PropertyTagThumbnailPlanarConfig = cast(PROPID, &h502F)
const PropertyTagThumbnailResolutionUnit = cast(PROPID, &h5030)
const PropertyTagThumbnailTransferFunction = cast(PROPID, &h5031)
const PropertyTagThumbnailSoftwareUsed = cast(PROPID, &h5032)
const PropertyTagThumbnailDateTime = cast(PROPID, &h5033)
const PropertyTagThumbnailArtist = cast(PROPID, &h5034)
const PropertyTagThumbnailWhitePoint = cast(PROPID, &h5035)
const PropertyTagThumbnailPrimaryChromaticities = cast(PROPID, &h5036)
const PropertyTagThumbnailYCbCrCoefficients = cast(PROPID, &h5037)
const PropertyTagThumbnailYCbCrSubsampling = cast(PROPID, &h5038)
const PropertyTagThumbnailYCbCrPositioning = cast(PROPID, &h5039)
const PropertyTagThumbnailRefBlackWhite = cast(PROPID, &h503A)
const PropertyTagThumbnailCopyRight = cast(PROPID, &h503B)
const PropertyTagLuminanceTable = cast(PROPID, &h5090)
const PropertyTagChrominanceTable = cast(PROPID, &h5091)
const PropertyTagFrameDelay = cast(PROPID, &h5100)
const PropertyTagLoopCount = cast(PROPID, &h5101)
const PropertyTagGlobalPalette = cast(PROPID, &h5102)
const PropertyTagIndexBackground = cast(PROPID, &h5103)
const PropertyTagIndexTransparent = cast(PROPID, &h5104)
const PropertyTagPixelUnit = cast(PROPID, &h5110)
const PropertyTagPixelPerUnitX = cast(PROPID, &h5111)
const PropertyTagPixelPerUnitY = cast(PROPID, &h5112)
const PropertyTagPaletteHistogram = cast(PROPID, &h5113)
const PropertyTagCopyright = cast(PROPID, &h8298)
const PropertyTagExifExposureTime = cast(PROPID, &h829A)
const PropertyTagExifFNumber = cast(PROPID, &h829D)
const PropertyTagExifIFD = cast(PROPID, &h8769)
const PropertyTagICCProfile = cast(PROPID, &h8773)
const PropertyTagExifExposureProg = cast(PROPID, &h8822)
const PropertyTagExifSpectralSense = cast(PROPID, &h8824)
const PropertyTagGpsIFD = cast(PROPID, &h8825)
const PropertyTagExifISOSpeed = cast(PROPID, &h8827)
const PropertyTagExifOECF = cast(PROPID, &h8828)
const PropertyTagExifVer = cast(PROPID, &h9000)
const PropertyTagExifDTOrig = cast(PROPID, &h9003)
const PropertyTagExifDTDigitized = cast(PROPID, &h9004)
const PropertyTagExifCompConfig = cast(PROPID, &h9101)
const PropertyTagExifCompBPP = cast(PROPID, &h9102)
const PropertyTagExifShutterSpeed = cast(PROPID, &h9201)
const PropertyTagExifAperture = cast(PROPID, &h9202)
const PropertyTagExifBrightness = cast(PROPID, &h9203)
const PropertyTagExifExposureBias = cast(PROPID, &h9204)
const PropertyTagExifMaxAperture = cast(PROPID, &h9205)
const PropertyTagExifSubjectDist = cast(PROPID, &h9206)
const PropertyTagExifMeteringMode = cast(PROPID, &h9207)
const PropertyTagExifLightSource = cast(PROPID, &h9208)
const PropertyTagExifFlash = cast(PROPID, &h9209)
const PropertyTagExifFocalLength = cast(PROPID, &h920A)
const PropertyTagExifMakerNote = cast(PROPID, &h927C)
const PropertyTagExifUserComment = cast(PROPID, &h9286)
const PropertyTagExifDTSubsec = cast(PROPID, &h9290)
const PropertyTagExifDTOrigSS = cast(PROPID, &h9291)
const PropertyTagExifDTDigSS = cast(PROPID, &h9292)
const PropertyTagExifFPXVer = cast(PROPID, &hA000)
const PropertyTagExifColorSpace = cast(PROPID, &hA001)
const PropertyTagExifPixXDim = cast(PROPID, &hA002)
const PropertyTagExifPixYDim = cast(PROPID, &hA003)
const PropertyTagExifRelatedWav = cast(PROPID, &hA004)
const PropertyTagExifInterop = cast(PROPID, &hA005)
const PropertyTagExifFlashEnergy = cast(PROPID, &hA20B)
const PropertyTagExifSpatialFR = cast(PROPID, &hA20C)
const PropertyTagExifFocalXRes = cast(PROPID, &hA20E)
const PropertyTagExifFocalYRes = cast(PROPID, &hA20F)
const PropertyTagExifFocalResUnit = cast(PROPID, &hA210)
const PropertyTagExifSubjectLoc = cast(PROPID, &hA214)
const PropertyTagExifExposureIndex = cast(PROPID, &hA215)
const PropertyTagExifSensingMethod = cast(PROPID, &hA217)
const PropertyTagExifFileSource = cast(PROPID, &hA300)
const PropertyTagExifSceneType = cast(PROPID, &hA301)
const PropertyTagExifCfaPattern = cast(PROPID, &hA302)
const PropertyTagTypeByte = cast(WORD, 1)
const PropertyTagTypeASCII = cast(WORD, 2)
const PropertyTagTypeShort = cast(WORD, 3)
const PropertyTagTypeLong = cast(WORD, 4)
const PropertyTagTypeRational = cast(WORD, 5)
const PropertyTagTypeUndefined = cast(WORD, 7)
const PropertyTagTypeSLONG = cast(WORD, 9)
const PropertyTagTypeSRational = cast(WORD, 10)

'extern EncoderChrominanceTable as const GUID
'extern EncoderColorDepth as const GUID
'extern EncoderColorSpace as const GUID
'extern EncoderCompression as const GUID
'extern EncoderImageItems as const GUID
'extern EncoderLuminanceTable as const GUID
'extern EncoderQuality as const GUID
'extern EncoderRenderMethod as const GUID
'extern EncoderSaveAsCMYK as const GUID
'extern EncoderSaveFlag as const GUID
'extern EncoderScanMethod as const GUID
'extern EncoderTransformation as const GUID
'extern EncoderVersion as const GUID
'extern ImageFormatBMP as const GUID
'extern ImageFormatEMF as const GUID
'extern ImageFormatEXIF as const GUID
'extern ImageFormatGIF as const GUID
'extern ImageFormatIcon as const GUID
'extern ImageFormatJPEG as const GUID
'extern ImageFormatMemoryBMP as const GUID
'extern ImageFormatPNG as const GUID
'extern ImageFormatTIFF as const GUID
'extern ImageFormatUndefined as const GUID
'extern ImageFormatWMF as const GUID
'extern FrameDimensionPage as const GUID
'extern FrameDimensionResolution as const GUID
'extern FrameDimensionTime as const GUID

' // Encoder parameter GUIDs
Dim Shared EncoderChrominanceTable   As GUID = Type(&hF2E455DC, &h09B3, &h4316, {&h82, &h60, &h67, &h6A, &hDA, &h32, &h48, &h1C})
Dim Shared EncoderColorDepth         As GUID = Type(&h66087055, &hAD66, &h4C7C, {&h9A, &h18, &h38, &hA2, &h31, &h0B, &h83, &h37})
Dim Shared EncoderColorSpace         As GUID = Type(&hAE7A62A0, &hEE2C, &h49D8, {&h9D, &h07, &h16, &hAF, &h4E, &h4D, &hBD, &h1E})
Dim Shared EncoderCompression        As GUID = Type(&hECCBFEFF, &h4F77, &h4D61, {&h8A, &hB9, &h6A, &h51, &hF6, &hD3, &h8E, &h8B})
Dim Shared EncoderImageItems         As GUID = Type(&h63875E13, &h1F1D, &h45AB, {&h91, &h95, &hCF, &h0F, &h5A, &h0C, &hBE, &h1B})
Dim Shared EncoderLuminanceTable     As GUID = Type(&hEDB33BCE, &h0266, &h4A77, {&hB9, &h04, &h27, &h21, &h60, &h99, &hE7, &h17})
Dim Shared EncoderQuality            As GUID = Type(&h1D5BE4B5, &hFA4A, &h452D, {&h9C, &hDD, &h5D, &hB3, &h51, &h05, &hE7, &hEB})
Dim Shared EncoderRenderMethod       As GUID = Type(&h6DB9FCAA, &h0E84, &h485E, {&h9A, &h8D, &h1E, &h8B, &h12, &h1F, &h0C, &hC1})
Dim Shared EncoderSaveAsCMYK         As GUID = Type(&hC617F8E1, &h1E5F, &h4B1F, {&hA6, &hD3, &h5E, &h1D, &h73, &h6C, &h6A, &h1D})
Dim Shared EncoderSaveFlag           As GUID = Type(&h292266FC, &hAC40, &h47BF, {&h8C, &hFC, &hA8, &h5B, &h89, &hA6, &h55, &hDE})
Dim Shared EncoderScanMethod         As GUID = Type(&h3A4E2661, &h3109, &h4E56, {&h85, &h1E, &h9F, &h7E, &h42, &h9F, &h1A, &h22})
Dim Shared EncoderTransformation     As GUID = Type(&h8D0EB2D1, &h6A48, &h4A94, {&hAC, &h61, &h52, &h95, &hC7, &h2E, &h8E, &h33})
Dim Shared EncoderVersion            As GUID = Type(&h24D18C76, &h814A, &h41A4, {&hBF, &h53, &h1C, &h21, &h9C, &hCC, &hF7, &h97})

' // Image Format GUIDs
Dim Shared ImageFormatBMP           As GUID = Type(&hB96B3CAB, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatEMF           As GUID = Type(&hB96B3CAC, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatEXIF          As GUID = Type(&hB96B3CB2, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatGIF           As GUID = Type(&hB96B3CB0, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatIcon          As GUID = Type(&hB96B3CB5, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatJPEG          As GUID = Type(&hB96B3CAE, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatMemoryBMP     As GUID = Type(&hB96B3CB4, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatPNG           As GUID = Type(&hB96B3CAF, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatTIFF          As GUID = Type(&hB96B3CB1, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatUndefined     As GUID = Type(&hB96B3CAA, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})
Dim Shared ImageFormatWMF           As GUID = Type(&hB96B3CAD, &h0728, &h11D3, {&h9D, &h7B, &h00, &h00, &hF8, &h1E, &hF3, &h2E})

' // Frame Dimension GUIDs
Dim Shared FrameDimensionPage       As GUID = Type(&h7462DC86, &h6180, &h4C7E, {&h8E, &h3F, &hEE, &h73, &h33, &hA7, &hA4, &h83})
Dim Shared FrameDimensionResolution As GUID = Type(&h84236F7B, &h3BD3, &h4281, {&h8D, &h62, &h1A, &h65, &h47, &hE0, &hEA, &hD1})
Dim Shared FrameDimensionTime       As GUID = Type(&h6AEDBD6D, &h3FB5, &h418A, {&h83, &hA6, &h7F, &h45, &h22, &h9D, &hC8, &h72})

' #define __GDIPLUS_INIT_H

type GdiplusStartupInput
	GdiplusVersion as UINT32
	DebugEventCallback as DebugEventProc
	SuppressBackgroundThread as BOOL
	SuppressExternalCodecs as BOOL
end type

type NotificationHookProc as function (byval token as ULONG_PTR PTR) AS GpStatus
type NotificationUnhookProc as sub (byval token as ULONG_PTR)

type GdiplusStartupOutput
	NotificationHook as NotificationHookProc
	NotificationUnhook as NotificationUnhookProc
end type

DECLARE FUNCTION GdiplusStartup (BYVAL token AS ULONG_PTR PTR, BYVAL input AS CONST GdiplusStartupInput PTR, BYVAL output AS GdiplusStartupOutput PTR) AS GpStatus
declare sub GdiplusShutdown (BYVAL token  AS ULONG_PTR)
DECLARE FUNCTION GdiplusNotificationHook (BYVAL token AS ULONG_PTR PTR) AS GpStatus
DECLARE SUB GdiplusNotificationUnhook (BYVAL token AS ULONG_PTR)

'#define __GDIPLUS_MEM_H
DECLARE FUNCTION GdipAlloc (BYVAL size AS size_t) AS ANY PTR
DECLARE SUB GdipFree (BYVAL AS ANY PTR)

'#define __GDIPLUS_METAHEADER_H
const GDIP_EMFPLUSFLAGS_DISPLAY = cast(UINT, 1)

type tagENHMETAHEADER3
	iType as DWORD
	nSize as DWORD
	rclBounds as RECTL
	rclFrame as RECTL
	dSignature as DWORD
	nVersion as DWORD
	nBytes as DWORD
	nRecords as DWORD
	nHandles as WORD
	sReserved as WORD
	nDescription as DWORD
	offDescription as DWORD
	nPalEntries as DWORD
	szlDevice as SIZEL
	szlMillimeters as SIZEL
end type

type ENHMETAHEADER3 as tagENHMETAHEADER3
type LPENHMETAHEADER3 as tagENHMETAHEADER3 ptr

type PWMFRect16
	Left as INT16
	Top as INT16
	Right as INT16
	Bottom as INT16
end type

type WmfPlaceableFileHeader
	Key as UINT32
	Hmf as INT16
	BoundingBox as PWMFRect16
	Inch as INT16
	Reserved as UINT32
	Checksum as INT16
end type

type MetafileHeader
	as MetafileType Type
	Size as UINT
	Version as UINT
	EmfPlusFlags as UINT
	DpiX as REAL
	DpiY as REAL
	X as INT_
	Y as INT_
	Width as INT_
	Height as INT_

	union
		WmfHeader as METAHEADER
		EmfHeader as ENHMETAHEADER3
	end union

	EmfPlusHeaderSize as INT_
	LogicalDpiX as INT_
	LogicalDpiY as INT_
end type

' #define __GDIPLUS_PIXELFORMATS_H
type ARGB as DWORD
type PixelFormat as INT_
#define PixelFormatIndexed cast(INT_, &h00010000)
#define PixelFormatGDI cast(INT_, &h00020000)
#define PixelFormatAlpha cast(INT_, &h00040000)
#define PixelFormatPAlpha cast(INT_, &h00080000)
#define PixelFormatExtended cast(INT_, &h00100000)
#define PixelFormatCanonical cast(INT_, &h00200000)
#define PixelFormatUndefined cast(INT_, 0)
#define PixelFormatDontCare cast(INT_, 0)
#define PixelFormat1bppIndexed cast(INT_, ((1 or (1 shl 8)) or PixelFormatIndexed) or PixelFormatGDI)
#define PixelFormat4bppIndexed cast(INT_, ((2 or (4 shl 8)) or PixelFormatIndexed) or PixelFormatGDI)
#define PixelFormat8bppIndexed cast(INT_, ((3 or (8 shl 8)) or PixelFormatIndexed) or PixelFormatGDI)
#define PixelFormat16bppGrayScale cast(INT_, (4 or (16 shl 8)) or PixelFormatExtended)
#define PixelFormat16bppRGB555 cast(INT_, (5 or (16 shl 8)) or PixelFormatGDI)
#define PixelFormat16bppRGB565 cast(INT_, (6 or (16 shl 8)) or PixelFormatGDI)
#define PixelFormat16bppARGB1555 cast(INT_, ((7 or (16 shl 8)) or PixelFormatAlpha) or PixelFormatGDI)
#define PixelFormat24bppRGB cast(INT_, (8 or (24 shl 8)) or PixelFormatGDI)
#define PixelFormat32bppRGB cast(INT_, (9 or (32 shl 8)) or PixelFormatGDI)
#define PixelFormat32bppARGB cast(INT_, (((10 or (32 shl 8)) or PixelFormatAlpha) or PixelFormatGDI) or PixelFormatCanonical)
#define PixelFormat32bppPARGB cast(INT_, (((11 or (32 shl 8)) or PixelFormatAlpha) or PixelFormatPAlpha) or PixelFormatGDI)
#define PixelFormat48bppRGB cast(INT_, (12 or (48 shl 8)) or PixelFormatExtended)
#define PixelFormat64bppARGB cast(INT_, (((13 or (64 shl 8)) or PixelFormatAlpha) or PixelFormatCanonical) or PixelFormatExtended)
#define PixelFormat64bppPARGB cast(INT_, (((14 or (64 shl 8)) or PixelFormatAlpha) or PixelFormatPAlpha) or PixelFormatExtended)
#define PixelFormatMax cast(INT_, 15)

type PaletteFlags as long
enum
	PaletteFlagsHasAlpha = 1
	PaletteFlagsGrayScale = 2
	PaletteFlagsHalftone = 4
end enum

type PaletteType as long
enum
	PaletteTypeCustom = 0
	PaletteTypeOptimal = 1
	PaletteTypeFixedBW = 2
	PaletteTypeFixedHalftone8 = 3
	PaletteTypeFixedHalftone27 = 4
	PaletteTypeFixedHalftone64 = 5
	PaletteTypeFixedHalftone125 = 6
	PaletteTypeFixedHalftone216 = 7
	PaletteTypeFixedHalftone252 = 8
	PaletteTypeFixedHalftone256 = 9
end enum

type ColorPalette
	Flags as UINT
	Count as UINT
	Entries(0 to 0) as ARGB
end type

#define GetPixelFormatSize(pixfmt) cast(UINT, (cast(UINT, (pixfmt)) and &hff00u) shr 8)
#define IsAlphaPixelFormat(pixfmt) cast(BOOL, -(((pixfmt) and cast(INT_, &h00040000)) <> 0))
#define IsCanonicalPixelFormat(pixfmt) cast(BOOL, -(((pixfmt) and cast(INT_, &h00200000)) <> 0))
#define IsExtendedPixelFormat(pixfmt) cast(BOOL, -(((pixfmt) and cast(INT_, &h00100000)) <> 0))
#define IsIndexedPixelFormat(pixfmt) cast(BOOL, -(((pixfmt) and cast(INT_, &h00010000)) <> 0))
#define __GDIPLUS_COLOR_H

type ColorChannelFlags as long
enum
	ColorChannelFlagsC = 0
	ColorChannelFlagsM = 1
	ColorChannelFlagsY = 2
	ColorChannelFlagsK = 3
	ColorChannelFlagsLast = 4
end enum

type Color
	Value as ARGB
end type

' #define __GDIPLUS_COLORMATRIX_H

type ColorAdjustType as long
enum
	ColorAdjustTypeDefault = 0
	ColorAdjustTypeBitmap = 1
	ColorAdjustTypeBrush = 2
	ColorAdjustTypePen = 3
	ColorAdjustTypeText = 4
	ColorAdjustTypeCount = 5
	ColorAdjustTypeAny = 6
end enum

type ColorMatrixFlags as long
enum
	ColorMatrixFlagsDefault = 0
	ColorMatrixFlagsSkipGrays = 1
	ColorMatrixFlagsAltGray = 2
end enum

type HistogramFormat as long
enum
	HistogramFormatARGB = 0
	HistogramFormatPARGB = 1
	HistogramFormatRGB = 2
	HistogramFormatGray = 3
	HistogramFormatB = 4
	HistogramFormatG = 5
	HistogramFormatR = 6
	HistogramFormatA = 7
end enum

type ColorMap_
	oldColor as Color
	newColor as Color
end type
TYPE GpColorMap AS ColorMap_

type ColorMatrix
	m(0 to 4, 0 to 4) as REAL
end type
type GpColorMatrix AS ColorMatrix
' // Use an array (0 TO 4, 0 TO 4) AS SINGLE, and pass @array(0, 0)

' #define __GDIPLUS_FLAT_H

' // AdjustableArroCap APIs
DECLARE FUNCTION GdipCreateAdjustableArrowCap (BYVAL height AS REAL, BYVAL Width AS REAL, BYVAL isFilled AS BOOL, BYVAL cap AS GpAdjustableArrowCap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipSetAdjustableArrowCapHeight (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL height AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetAdjustableArrowCapHeight (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL height AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipSetAdjustableArrowCapWidth (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL width AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetAdjustableArrowCapWidth (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL width AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipSetAdjustableArrowCapMiddleInset (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL middleInset AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetAdjustableArrowCapMiddleInset (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL middleInset AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipSetAdjustableArrowCapFillState (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL fillState AS BOOL) AS GpStatus
DECLARE FUNCTION GdipGetAdjustableArrowCapFillState (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL fillState AS BOOL PTR) AS GpStatus

' // Bitmap APIs
DECLARE FUNCTION GdipCreateBitmapFromStream (BYVAL stream AS IStream PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromFile (BYVAL filename AS CONST WSTRING PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromStreamICM (BYVAL stream AS IStream PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromFileICM (BYVAL filename AS CONST WSTRING PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromScan0 (BYVAL width AS INT_, BYVAL height AS INT_, BYVAL stride AS INT_, _
                 BYVAL format AS PixelFormat, BYVAL scan0 AS BYTE PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromGraphics (BYVAL width AS INT_, BYVAL height AS INT_, _
                 BYVAL target AS GpGraphics PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromDirectDrawSurface (BYVAL surface AS IDirectDrawSurface7 PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromGdiDib (BYVAL gdiBitmapInfo AS CONST BITMAPINFO PTR, _
                 BYVAL gdiBitmapData AS ANY PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromHBITMAP (BYVAL hbm AS HBITMAP, BYVAL hpal AS HPALETTE, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateHBITMAPFromBitmap (BYVAL bmp AS GpBitmap PTR, BYVAL hbmReturn AS HBITMAP PTR, BYVAL background AS ARGB) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromHICON (BYVAL hicon AS HICON, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateHICONFromBitmap (BYVAL bmp AS GpBitmap PTR, BYVAL hbmReturn AS HICON PTR) AS GpStatus
DECLARE FUNCTION GdipCreateBitmapFromResource (BYVAL hInst AS HINSTANCE, BYVAL lpBitmapName AS CONST WSTRING PTR, BYVAL bmp AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCloneBitmapArea (BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL, _
                 BYVAL format AS PixelFormat, BYVAL srcBitmap AS GpBitmap PTR, BYVAL dstBitmap AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCloneBitmapAreaI (BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_, _
                 BYVAL format AS PixelFormat, BYVAL srcBitmap AS GpBitmap PTR, BYVAL dstBitmap AS GpBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipBitmapLockBits (BYVAL bmp AS GpBitmap PTR, BYVAL rc AS CONST GpRect PTR, BYVAL flags AS UINT, _
                 BYVAL format AS PixelFormat, BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
DECLARE FUNCTION GdipBitmapUnlockBits (BYVAL bmp AS GpBitmap PTR, BYVAL lockedBitmapData AS BitmapData PTR) AS GpStatus
DECLARE FUNCTION GdipBitmapGetPixel (BYVAL bmp AS GpBitmap PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL color AS ARGB PTR) AS GpStatus
DECLARE FUNCTION GdipBitmapSetPixel (BYVAL bmp AS GpBitmap PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL color AS ARGB) AS GpStatus
'#if (GDIPVER >= 0x0110)
type GdiplusAbort as GdiplusAbort_
DECLARE FUNCTION GdipImageSetAbort (BYVAL pImage AS GpImage PTR, BYVAL pIAbort AS GdiplusAbort PTR) AS GpStatus
DECLARE FUNCTION GdipGraphicsSetAbort (BYVAL pGraphics AS GpGraphics PTR, BYVAL pIAbort AS GdiplusAbort PTR) AS GpStatus
DECLARE FUNCTION GdipBitmapConvertFormat (BYVAL pInputBitmap AS GpBitmap PTR, BYVAL format AS PixelFormat, _
                 BYVAL dithertype AS DitherType, BYVAL palettetype AS PaletteType, BYVAL palette AS ColorPalette PTR, _
                 BYVAL alphaThresholdPercent AS REAL) AS GpStatus
DECLARE FUNCTION GdipInitializePalette (BYVAL palette AS ColorPalette PTR, BYVAL palettetype AS PaletteType, _
                 BYVAL optimalColors AS INT_, BYVAL useTransparentColor AS BOOL, BYVAL bmp AS GpBitmap PTR) AS GpStatus
DECLARE FUNCTION GdipBitmapApplyEffect (BYVAL bmp AS GpBitmap PTR, BYVAL effect AS CGpEffect PTR, BYVAL roi AS RECT PTR, _
                 BYVAL useAuxData AS BOOL, BYVAL auxData AS ANY PTR PTR, BYVAL auxDataSize AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipBitmapCreateApplyEffect (BYVAL inputBitmaps AS GpBitmap PTR PTR, BYVAL numInputs AS INT_, _
                 BYVAL effect AS CGpEffect PTR, BYVAL roi AS RECT PTR, BYVAL outputRect AS RECT PTR, _
                 BYVAL outputBitmap AS GpBitmap PTR PTR, BYVAL useAuxData AS BOOL, BYVAL auxData AS ANY PTR PTR, _
                 BYVAL auxDataSize AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipBitmapGetHistogram (BYVAL bmp AS GpBitmap PTR, BYVAL format AS HistogramFormat, BYVAL NumberOfEntries AS UINT, _
                 BYVAL channel0 AS UINT PTR, BYVAL channel1 AS UINT PTR, BYVAL channel2 AS UINT PTR, BYVAL channel3 AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipBitmapGetHistogramSize (BYVAL format AS HistogramFormat, BYVAL NumberOfEntries AS UINT PTR) AS GpStatus
'#endif
DECLARE FUNCTION GdipBitmapSetResolution (BYVAL bmp AS GpBitmap PTR, BYVAL xdpi AS REAL, BYVAL ydpi AS REAL) AS GpStatus

' // Brush APIs
DECLARE FUNCTION GdipCloneBrush (BYVAL brush AS GpBrush PTR, BYVAL cloneBrush AS GpBrush PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteBrush (BYVAL brush AS GpBrush PTR) AS GpStatus
DECLARE FUNCTION GdipGetBrushType (BYVAL brush AS GpBrush PTR, BYVAL brushType AS GpBrushType PTR) AS GpStatus

' // Cached Bitmap APIs
DECLARE FUNCTION GdipCreateCachedBitmap (BYVAL bmp AS GpBitmap PTR, BYVAL graphics AS GpGraphics PTR, BYVAL cachedBitmap AS GpCachedBitmap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteCachedBitmap (BYVAL cachedBitmap AS GpCachedBitmap PTR) AS GpStatus
DECLARE FUNCTION GdipDrawCachedBitmap (BYVAL graphics AS GpGraphics PTR, BYVAL cachedBitmap AS GpCachedBitmap PTR, _
                 BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus

' // CustomLineCap APIs
DECLARE FUNCTION GdipCreateCustomLineCap (BYVAL fillPath AS GpPath PTR, BYVAL strokePath AS GpPath PTR, _
                 BYVAL baseCap AS GpLineCap, BYVAL baseInset AS REAL, BYVAL customCap AS GpCustomLineCap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteCustomLineCap (BYVAL customCap AS GpCustomLineCap PTR) AS GpStatus
DECLARE FUNCTION GdipCloneCustomLineCap (BYVAL customCap AS GpCustomLineCap PTR, BYVAL clonedCap AS GpCustomLineCap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetCustomLineCapType (BYVAL customCap AS GpCustomLineCap PTR, BYVAL capType AS CustomLineCapType PTR) AS GpStatus
DECLARE FUNCTION GdipSetCustomLineCapStrokeCaps (BYVAL customCap AS GpCustomLineCap PTR, BYVAL startCap AS GpLineCap, BYVAL endCap AS GpLineCap) AS GpStatus
DECLARE FUNCTION GdipGetCustomLineCapStrokeCaps (BYVAL customCap AS GpCustomLineCap PTR, BYVAL startCap AS GpLineCap PTR, BYVAL endCap AS GpLineCap PTR) AS GpStatus
DECLARE FUNCTION GdipSetCustomLineCapStrokeJoin (BYVAL customCap AS GpCustomLineCap PTR, BYVAL lineJoin AS GpLineJoin) AS GpStatus
DECLARE FUNCTION GdipGetCustomLineCapStrokeJoin (BYVAL customCap AS GpCustomLineCap PTR, BYVAL lineJoin AS GpLineJoin PTR) AS GpStatus
DECLARE FUNCTION GdipSetCustomLineCapBaseCap (BYVAL customCap AS GpCustomLineCap PTR, BYVAL baseCap As GpLineCap) AS GpStatus
DECLARE FUNCTION GdipGetCustomLineCapBaseCap (BYVAL customCap AS GpCustomLineCap PTR, BYVAL baseCap AS GpLineCap PTR) AS GpStatus
DECLARE FUNCTION GdipSetCustomLineCapBaseInset (BYVAL customCap AS GpCustomLineCap PTR, BYVAL inset AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetCustomLineCapBaseInset (BYVAL customCap AS GpCustomLineCap PTR, BYVAL inset AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipSetCustomLineCapWidthScale (BYVAL customCap AS GpCustomLineCap PTR, BYVAL widthScale AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetCustomLineCapWidthScale (BYVAL customCap AS GpCustomLineCap PTR, BYVAL widthScale As REAL PTR) AS GpStatus

' // Effect APIs
DECLARE FUNCTION GdipCreateEffect (BYVAL guid AS CONST GUID, BYVAL effect AS CGpEffect PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteEffect (BYVAL effect AS CGpEffect PTR) AS GpStatus
DECLARE FUNCTION GdipGetEffectParameterSize (BYVAL effect AS CGpEffect PTR, BYVAL size AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipSetEffectParameters (BYVAL effect AS CGpEffect PTR, BYVAL params AS CONST ANY PTR, BYVAL size AS CONST UINT) AS GpStatus
DECLARE FUNCTION GdipGetEffectParameters (BYVAL effect AS CGpEffect PTR, BYVAL size AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

' // Font APIs
DECLARE FUNCTION GdipCreateFontFromDC (BYVAL hdc AS HDC, BYVAL font AS GpFont PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateFontFromLogfontA (BYVAL hdc AS HDC, BYVAL logfont AS CONST LOGFONTA PTR, BYVAL font AS GpFont PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateFontFromLogfontW (BYVAL hdc AS HDC, BYVAL logfont AS CONST LOGFONTW PTR, BYVAL font AS GpFont PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateFont (BYVAL fontFamily AS CONST GpFontFamily PTR, BYVAL emSize AS REAL, BYVAL style AS INT_, _
                 BYVAL unit AS GpUnit, BYVAL font AS GpFont PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCloneFont (BYVAL font AS GpFont PTR, BYVAL cloneFont AS GpFont PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteFont (BYVAL font AS GpFont PTR) AS GpStatus
DECLARE FUNCTION GdipGetFamily (BYVAL font AS GpFont PTR, BYVAL family AS GpFontFamily PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetFontStyle (BYVAL font AS GpFont PTR, BYVAL style AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetFontSize (BYVAL font AS GpFont PTR, BYVAL size AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipGetFontUnit (BYVAL font AS GpFont PTR, BYVAL unit AS GpUnit PTR) AS GpStatus
DECLARE FUNCTION GdipGetFontHeight (BYVAL font AS CONST GpFont PTR, BYVAL graphics AS CONST GpGraphics PTR, BYVAL height AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipGetFontHeightGivenDPI (BYVAL font AS CONST GpFont PTR, BYVAL dpi AS REAL, BYVAL height AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipGetLogFontA (BYVAL font AS GpFont PTR, BYVAL graphcis AS GpGraphics PTR, BYVAL logFontA AS LOGFONTA PTR) AS GpStatus
DECLARE FUNCTION GdipGetLogFontW (BYVAL font AS GpFont PTR, BYVAL graphics AS GpGraphics PTR, BYVAL logFontW AS LOGFONTW PTR) AS GpStatus
DECLARE FUNCTION GdipNewInstalledFontCollection (BYVAL fontCollection AS GpFontCollection PTR PTR) AS GpStatus
DECLARE FUNCTION GdipNewPrivateFontCollection (BYVAL fontCollection AS GpFontCollection PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeletePrivateFontCollection (BYVAL fontCollection AS GpFontCollection PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetFontCollectionFamilyCount (BYVAL fontCollection AS GpFontCollection PTR, BYVAL numFound AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetFontCollectionFamilyList (BYVAL fontCollection AS GpFontCollection PTR, BYVAL numSought AS INT_, _
                 BYVAL gpfamilies AS GpFontFamily PTR PTR, BYVAL numFound AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipPrivateAddFontFile (BYVAL fontCollection AS GpFontCollection PTR, BYVAL filename AS CONST WSTRING PTR) AS GpStatus
DECLARE FUNCTION GdipPrivateAddMemoryFont (BYVAL fontCollection AS GpFontCollection PTR, BYVAL memory AS CONST ANY PTR, BYVAL length AS INT_) AS GpStatus

'  // FontFamily APIs
DECLARE FUNCTION GdipCreateFontFamilyFromName (BYVAL name AS CONST WSTRING PTR, BYVAL fontCollection AS GpFontCollection PTR, BYVAL fontFamily AS GpFontFamily PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteFontFamily (BYVAL fontFamily AS GpFontFamily PTR) AS GpStatus
DECLARE FUNCTION GdipCloneFontFamily (BYVAL fontFamily AS GpFontFamily PTR, BYVAL clonedFontFamily AS GpFontFamily PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetGenericFontFamilySansSerif (BYVAL nativeFamily AS GpFontFamily PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetGenericFontFamilySerif (BYVAL nativeFamily AS GpFontFamily PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetGenericFontFamilyMonospace (BYVAL nativeFamily AS GpFontFamily PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetFamilyName (BYVAL family AS CONST GpFontFamily PTR, BYVAL name AS WSTRING PTR, BYVAL language AS LANGID) AS GpStatus
DECLARE FUNCTION GdipIsStyleAvailable (BYVAL family AS CONST GpFontFamily PTR, BYVAL style AS INT_, BYVAL IsStyleAvailable AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipFontCollectionEnumerable (BYVAL fontCollection AS GpFontCollection PTR, BYVAL graphics AS GpGraphics PTR, BYVAL numFound AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipFontCollectionEnumerate (BYVAL fontCollection AS GpFontCollection PTR, BYVAL numSought AS INT_, _
                 BYVAL gpfamilies AS GpFontFamily PTR PTR, BYVAL numFound AS INT_ PTR, BYVAL graphics AS GpGraphics PTR) AS GpStatus
DECLARE FUNCTION GdipGetEmHeight (BYVAL family AS CONST GpFontFamily PTR, BYVAL style AS INT_, BYVAL EmHeight AS UINT16 PTR) AS GpStatus
DECLARE FUNCTION GdipGetCellAscent (BYVAL family AS CONST GpFontFamily PTR, BYVAL style AS INT_, BYVAL CellAscent AS UINT16 PTR) AS GpStatus
DECLARE FUNCTION GdipGetCellDescent (BYVAL family AS CONST GpFontFamily PTR, BYVAL style AS INT_, BYVAL CellDescent AS UINT16 PTR) AS GpStatus
DECLARE FUNCTION GdipGetLineSpacing (BYVAL family AS CONST GpFontFamily PTR, BYVAL style AS INT_, BYVAL LineSpacing AS UINT16 PTR) AS GpStatus

' // Graphics APIs
DECLARE FUNCTION GdipFlush (BYVAL graphics AS GpGraphics PTR, BYVAL intention AS GpFlushIntention) AS GpStatus
DECLARE FUNCTION GdipCreateFromHDC (BYVAL hdc AS HDC, BYVAL graphics AS GpGraphics PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateFromHDC2 (BYVAL hdc AS HDC, BYVAL hDevice AS HANDLE, BYVAL graphics AS GpGraphics PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateFromHWND (BYVAL hwnd AS HWND, BYVAL graphics AS GpGraphics PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateFromHWNDICM (BYVAL hwnd AS HWND, BYVAL graphics AS GpGraphics PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteGraphics (BYVAL graphics AS GpGraphics PTR) AS GpStatus
DECLARE FUNCTION GdipGetDC (BYVAL graphics AS GpGraphics PTR, BYVAL hdc AS HDC PTR) AS GpStatus
DECLARE FUNCTION GdipReleaseDC (BYVAL graphics AS GpGraphics PTR, BYVAL hdc AS HDC) AS GpStatus
DECLARE FUNCTION GdipSetCompositingMode (BYVAL graphics AS GpGraphics PTR, BYVAL compositingMode AS CompositingMode) AS GpStatus
DECLARE FUNCTION GdipGetCompositingMode (BYVAL graphics AS GpGraphics PTR, BYVAL compositingMode AS CompositingMode PTR) AS GpStatus
DECLARE FUNCTION GdipSetRenderingOrigin (BYVAL graphics AS GpGraphics PTR, BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetRenderingOrigin (BYVAL graphics AS GpGraphics PTR, BYVAL x AS INT_ PTR, BYVAL y AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipSetCompositingQuality (BYVAL graphics AS GpGraphics PTR, BYVAL compositingQuality AS CompositingQuality) AS GpStatus
DECLARE FUNCTION GdipGetCompositingQuality (BYVAL graphics AS GpGraphics PTR, BYVAL compositingQuality AS CompositingQuality PTR) AS GpStatus
DECLARE FUNCTION GdipSetSmoothingMode (BYVAL graphics AS GpGraphics PTR, BYVAL smoothingMode AS SmoothingMode) AS GpStatus
DECLARE FUNCTION GdipGetSmoothingMode (BYVAL graphics AS GpGraphics PTR, BYVAL smoothingMode AS SmoothingMode PTR) AS GpStatus
DECLARE FUNCTION GdipSetPixelOffsetMode (BYVAL graphics AS GpGraphics PTR, BYVAL pixelOffsetMode AS PixelOffsetMode) AS GpStatus
DECLARE FUNCTION GdipGetPixelOffsetMode (BYVAL graphics AS GpGraphics PTR, BYVAL pixelOffsetMode AS PixelOffsetMode PTR) AS GpStatus
DECLARE FUNCTION GdipSetTextRenderingHint (BYVAL graphics AS GpGraphics PTR, BYVAL mode AS TextRenderingHint) AS GpStatus
DECLARE FUNCTION GdipGetTextRenderingHint (BYVAL graphics AS GpGraphics PTR, BYVAL mode AS TextRenderingHint PTR) AS GpStatus
DECLARE FUNCTION GdipSetTextContrast (BYVAL graphics AS GpGraphics PTR, BYVAL contrast AS UINT) AS GpStatus
DECLARE FUNCTION GdipGetTextContrast (BYVAL graphics AS GpGraphics PTR, BYVAL contrast AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipSetInterpolationMode (BYVAL graphics AS GpGraphics PTR, BYVAL interpolationMode AS InterpolationMode) AS GpStatus
DECLARE FUNCTION GdipGetInterpolationMode (BYVAL graphics AS GpGraphics PTR, BYVAL interpolationMode AS InterpolationMode PTR) AS GpStatus
DECLARE FUNCTION GdipSetWorldTransform (BYVAL graphics AS GpGraphics PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipResetWorldTransform (BYVAL graphics AS GpGraphics PTR) AS GpStatus
DECLARE FUNCTION GdipMultiplyWorldTransform (BYVAL graphics AS GpGraphics PTR, BYVAL matrix AS CONST GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipTranslateWorldTransform (BYVAL graphics AS GpGraphics PTR, BYVAL dx AS REAL, BYVAL dy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipScaleWorldTransform (BYVAL graphics AS GpGraphics PTR, BYVAL sx AS REAL, BYVAL sy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipRotateWorldTransform (BYVAL graphics AS GpGraphics PTR, BYVAL angle AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipGetWorldTransform (BYVAL graphics AS GpGraphics PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipResetPageTransform (BYVAL graphics AS GpGraphics PTR) AS GpStatus
DECLARE FUNCTION GdipGetPageUnit (BYVAL graphics AS GpGraphics PTR, BYVAL unit AS GpUnit PTR) AS GpStatus
DECLARE FUNCTION GdipGetPageScale (BYVAL graphics AS GpGraphics PTR, BYVAL scale AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipSetPageUnit (BYVAL graphics AS GpGraphics PTR, BYVAL unit AS GpUnit) AS GpStatus
DECLARE FUNCTION GdipSetPageScale (BYVAL graphics AS GpGraphics PTR, BYVAL scale AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetDpiX (BYVAL graphics AS GpGraphics PTR, BYVAL dpi AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipGetDpiY (BYVAL graphics AS GpGraphics PTR, BYVAL dpi AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipTransformPoints (BYVAL graphics AS GpGraphics PTR, BYVAL destSpace AS GpCoordinateSpace, _
                 BYVAL srcSpace AS GpCoordinateSpace, BYVAL points AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipTransformPointsI (BYVAL graphics AS GpGraphics PTR, BYVAL destSpace AS GpCoordinateSpace, _
                 BYVAL srcSpace AS GpCoordinateSpace, BYVAL points AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetNearestColor (BYVAL graphics AS GpGraphics PTR, BYVAL argb AS ARGB PTR) AS GpStatus
DECLARE FUNCTION GdipCreateHalftonePalette () AS HPALETTE
DECLARE FUNCTION GdipDrawLine (BYVAL graphics AS GpGraphics PTR, BYVAL AS GpPen PTR, BYVAL x1 AS REAL, BYVAL y1 AS REAL, BYVAL x2 AS REAL, BYVAL y2 AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawLineI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawLines (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS const GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawLinesI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS const GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawArc (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
                 BYVAL width AS REAL, BYVAL height AS REAL, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawArcI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
                 BYVAL width AS INT_, BYVAL height AS INT_, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawBezier (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x1 AS REAL, BYVAL y1 AS REAL, _
                 BYVAL x2 AS REAL, BYVAL y2 AS REAL, BYVAL x3 AS REAL, BYVAL y3 AS REAL, BYVAL x4 AS REAL, BYVAL y4 AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawBezierI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, _
                 BYVAL x2 AS INT_, BYVAL y2 AS INT_, BYVAL x3 AS INT_, BYVAL y3 AS INT_, BYVAL c4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawBeziers (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawBeziersI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS const GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawRectangle (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawRectangleI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawRectangles (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL rects AS CONST GpRectF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawRectanglesI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL rects AS CONST GpRect PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawEllipse (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawEllipseI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawPie (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
                 BYVAL width AS REAL, BYVAL height AS REAL, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawPieI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
                 BYVAL width AS INT_, BYVAL height AS INT_, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawPolygon (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawPolygonI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawPath (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipDrawCurve (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawCurveI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawCurve2 (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawCurve2I (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawCurve3 (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL AS CONST GpPointF PTR, _
                 BYVAL count AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawCurve3I (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL AS CONST GpPoint PTR, _
                 BYVAL count AS INT_, BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawClosedCurve (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawClosedCurveI (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawClosedCurve2 (BYVAL graphics AS GpGraphics PTR, BYVAL oen AS GpPen PTR, BYVAL AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawClosedCurve2I (BYVAL graphics AS GpGraphics PTR, BYVAL oen AS GpPen PTR, BYVAL AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipGraphicsClear (BYVAL graphics AS GpGraphics PTR, BYVAL color AS ARGB) AS GpStatus
DECLARE FUNCTION GdipFillRectangle (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL) AS GpStatus
DECLARE FUNCTION GdipFillRectangleI (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_) AS GpStatus
DECLARE FUNCTION GdipFillRectangles (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL rects AS CONST GpRectF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipFillRectanglesI (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL rects AS CONST GpRect PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipFillPolygon (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL fillMode AS GpFillMode) AS GpStatus
DECLARE FUNCTION GdipFillPolygonI (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL fillMode AS GpFillMode) AS GpStatus
DECLARE FUNCTION GdipFillPolygon2 (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipFillPolygon2I (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipFillEllipse (BYVAL graphics AS GpGraphics PTR, BYVAL AS GpBrush PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL) AS GpStatus
DECLARE FUNCTION GdipFillEllipseI (BYVAL graphics AS GpGraphics PTR, BYVAL AS GpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_) AS GpStatus
DECLARE FUNCTION GdipFillPie (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
                 BYVAL width AS REAL, BYVAL height AS REAL, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipFillPieI (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
                 BYVAL width AS INT_, BYVAL height AS INT_, BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipFillPath (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipFillClosedCurve (BYVAL AS GpGraphics PTR, BYVAL AS GpBrush PTR, BYVAL AS const GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipFillClosedCurveI (BYVAL AS GpGraphics PTR, BYVAL AS GpBrush PTR, BYVAL AS const GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipFillClosedCurve2 (BYVAL AS GpGraphics PTR, BYVAL AS GpBrush PTR, BYVAL points AS CONST GpPointF PTR, _
                 BYVAL count AS INT_, BYVAL tension AS REAL, BYVAL fillmode AS GpFillMode) AS GpStatus
DECLARE FUNCTION GdipFillClosedCurve2I (BYVAL AS GpGraphics PTR, BYVAL AS GpBrush PTR, BYVAL points AS CONST GpPoint PTR, _
                 BYVAL count AS INT_, BYVAL tension AS REAL, BYVAL fillmode AS GpFillMode) AS GpStatus
DECLARE FUNCTION GdipFillRegion (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL region AS GpRegion PTR) AS GpStatus
'#if (GDIPVER >= 0x0110)
DECLARE FUNCTION GdipDrawImageFX (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL source AS GpRectF PTR, _
                 BYVAL xForm AS GpMatrix PTR, BYVAL effect AS CGpEffect PTR, BYVAL imageAttributes AS GpImageAttributes PTR, BYVAL srcUnit AS GpUnit) AS GpStatus
'#endif
DECLARE FUNCTION GdipDrawImage (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL x AS REAL, BYVAL y AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawImageI (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawImageRect (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL) AS GpStatus
DECLARE FUNCTION GdipDrawImageRectI (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawImagePoints (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawImagePointsI (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL AS CONST GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipDrawImagePointRect (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL x AS REAL, BYVAL y AS REAL, _
                 BYVAL srcx AS REAL, BYVAL srcy AS REAL, BYVAL srcwidth AS REAL, BYVAL srcheight AS REAL, BYVAL srcUnit AS GpUnit) AS GpStatus
DECLARE FUNCTION GdipDrawImagePointRectI (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL x AS INT_, BYVAL y AS INT_, _
                 BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit) AS GpStatus
DECLARE FUNCTION GdipDrawImageRectRect (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL dstx AS REAL, BYVAL dsty AS REAL, _
                 BYVAL dstwidth AS REAL, BYVAL dstheight AS REAL, BYVAL srcx AS REAL, BYVAL srcy AS REAL, BYVAL srcwidth AS REAL, BYVAL srcheight AS REAL, _
                 BYVAL srcUnit AS GpUnit, BYVAL imageAttributes AS CONST GpImageAttributes PTR, BYVAL callback AS DrawImageAbort, BYVAL callbackData AS ANY PTR) AS GpStatus
DECLARE FUNCTION GdipDrawImageRectRectI (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL dstx AS INT_, BYVAL dsty AS INT_, _
                 BYVAL dstwidth AS INT_, BYVAL dstheight AS INT_, BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, _
                 BYVAL srcUnit AS GpUnit, BYVAL imageAttributes AS CONST GpImageAttributes PTR, BYVAL callback AS DrawImageAbort, BYVAL callbackData AS ANY PTR) AS GpStatus
DECLARE FUNCTION GdipDrawImagePointsRect (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL points AS CONST GpPointF PTR, _
                 BYVAL count AS INT_, BYVAL srcx AS REAL, BYVAL srcy AS REAL, BYVAL srcwidth AS REAL, BYVAL srcheight AS REAL, BYVAL srcUnit AS GpUnit, _
                 BYVAL imageAttributes AS CONST GpImageAttributes PTR, BYVAL callback AS DrawImageAbort, BYVAL callbackData AS ANY PTR) AS GpStatus
DECLARE FUNCTION GdipDrawImagePointsRectI (BYVAL graphics AS GpGraphics PTR, BYVAL image AS GpImage PTR, BYVAL points AS CONST GpPoint PTR, _
                 BYVAL count AS INT_, BYVAL srcx AS INT_, BYVAL srcy AS INT_, BYVAL srcwidth AS INT_, BYVAL srcheight AS INT_, BYVAL srcUnit AS GpUnit, _
                 BYVAL imageAttributes AS CONST GpImageAttributes PTR, BYVAL callback AS DrawImageAbort, BYVAL callbackData AS ANY PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileDestPoint (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destPoint AS CONST GpPointF PTR, BYVAL callback AS EnumerateMetafileProc, BYVAL callbackData AS ANY PTR, _
                 BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileDestPointI (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destPoint AS CONST GpPoint PTR, BYVAL callback AS EnumerateMetafileProc, BYVAL callbackData AS ANY PTR, _
                 BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileDestRect (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destRect AS CONST GpRectF PTR, BYVAL callback AS EnumerateMetafileProc, BYVAL callbackData AS ANY PTR, _
                 BYVAL imageAttributes AS GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileDestRectI (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destRect AS CONST GpRect PTR, BYVAL callback AS EnumerateMetafileProc, BYVAL callbackData AS ANY PTR, _
                 BYVAL imageAttributes AS GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileDestPoints (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destPoints AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL AS EnumerateMetafileProc, _
                 BYVAL callbackData AS ANY PTR, BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileDestPointsI (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destPoints AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL AS EnumerateMetafileProc, _
                 BYVAL callbackData AS ANY PTR, BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileSrcRectDestPoint (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destPoint AS CONST GpPointF PTR, BYVAL srcRect AS CONST GpRectF PTR, BYVAL srcUnit AS GpUnit, _
                 BYVAL callback AS EnumerateMetafileProc, BYVAL callbackData AS ANY PTR, BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileSrcRectDestPointI (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destPoint AS CONST GpPoint PTR, BYVAL srcRect AS CONST GpRect PTR, BYVAL srcUnit AS GpUnit, _
                 BYVAL callback AS EnumerateMetafileProc, BYVAL callbackData AS ANY PTR, BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileSrcRectDestRect (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destRect AS CONST GpRectF PTR, BYVAL srcRect AS CONST GpRectF PTR, BYVAL srcUnit AS GpUnit, BYVAL callback AS EnumerateMetafileProc, _
                 BYVAL callbackData AS ANY PTR, BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileSrcRectDestRectI (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destRect AS CONST GpRect PTR, BYVAL srcRect AS CONST GpRect PTR, BYVAL srcUnit AS GpUnit, BYVAL callback AS EnumerateMetafileProc, _
                 BYVAL callbackData AS ANY PTR, BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileSrcRectDestPoints (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destPoints AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL srcRect AS CONST GpRectF PTR, _
                 BYVAL srcUnit AS GpUnit, BYVAL callback AS EnumerateMetafileProc, BYVAL callbackData AS ANY PTR, _
                 BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipEnumerateMetafileSrcRectDestPointsI (BYVAL graphics AS GpGraphics PTR, BYVAL metafile AS CONST GpMetafile PTR, _
                 BYVAL destPoints AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL srcRect AS CONST GpRect PTR, _
                 BYVAL srcUnit AS GpUnit, BYVAL callback AS EnumerateMetafileProc, BYVAL callbackData AS ANY PTR, _
                 BYVAL imageAttributes AS CONST GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipPlayMetafileRecord (BYVAL metafile AS CONST GpMetafile PTR, BYVAL recordType AS EmfPlusRecordType, _
                 BYVAL flags AS UINT, BYVAL dataSize AS UINT, BYVAL data AS CONST BYTE PTR) AS GpStatus
DECLARE FUNCTION GdipSetClipGraphics (BYVAL graphics AS GpGraphics PTR, BYVAL srcgraphics AS GpGraphics PTR, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipSetClipRect (BYVAL graphics AS GpGraphics PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, _
                 BYVAL height AS REAL, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipSetClipRectI (BYVAL graphics AS GpGraphics PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, _
                 BYVAL height AS INT_, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipSetClipPath (BYVAL graphics AS GpGraphics PTR, BYVAL path AS GpPath PTR, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipSetClipRegion (BYVAL graphics AS GpGraphics PTR, BYVAL region AS GpRegion PTR, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipSetClipHrgn (BYVAL graphics AS GpGraphics PTR, BYVAL hRgn AS HRGN, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipResetClip (BYVAL graphics AS GpGraphics PTR) AS GpStatus
DECLARE FUNCTION GdipTranslateClip (BYVAL graphics AS GpGraphics PTR, BYVAL dx AS REAL, BYVAL dy AS REAL) AS GpStatus
DECLARE FUNCTION GdipTranslateClipI (BYVAL graphics AS GpGraphics PTR, BYVAL dx AS INT_, BYVAL dy AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetClip (BYVAL graphics AS GpGraphics PTR, BYVAL region AS GpRegion PTR) AS GpStatus
DECLARE FUNCTION GdipGetClipBounds (BYVAL graphics AS GpGraphics PTR, BYVAL rcf AS GpRectF PTR) AS GpStatus
DECLARE FUNCTION GdipGetClipBoundsI (BYVAL graphics AS GpGraphics PTR, BYVAL rc AS GpRect PTR) AS GpStatus
DECLARE FUNCTION GdipIsClipEmpty (BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipGetVisibleClipBounds (BYVAL graphics AS GpGraphics PTR, BYVAL rcf AS GpRectF PTR) AS GpStatus
DECLARE FUNCTION GdipGetVisibleClipBoundsI (BYVAL graphics AS GpGraphics PTR, BYVAL rc AS GpRect PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisibleClipEmpty (BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisiblePoint (BYVAL graphics AS GpGraphics PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisiblePointI (BYVAL graphics AS GpGraphics PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisibleRect (BYVAL graphics AS GpGraphics PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisibleRectI (BYVAL graphics AS GpGraphics PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipSaveGraphics (BYVAL graphics AS GpGraphics PTR, BYVAL state AS GraphicsState PTR) AS GpStatus
DECLARE FUNCTION GdipRestoreGraphics (BYVAL graphics AS GpGraphics PTR, BYVAL AS GraphicsState) AS GpStatus
DECLARE FUNCTION GdipBeginContainer (BYVAL graphics AS GpGraphics PTR, BYVAL dstrect AS CONST GpRectF PTR, _
                 BYVAL srcrect AS CONST GpRectF PTR, BYVAL unit AS GpUnit, BYVAL state AS GraphicsContainer PTR) AS GpStatus
DECLARE FUNCTION GdipBeginContainerI (BYVAL graphics AS GpGraphics PTR, BYVAL dstrect AS CONST GpRect PTR, _
                 BYVAL srcrect AS CONST GpRect PTR, BYVAL unit AS GpUnit, BYVAL state AS GraphicsContainer PTR) AS GpStatus
DECLARE FUNCTION GdipBeginContainer2 (BYVAL graphics AS GpGraphics PTR, BYVAL state AS GraphicsContainer PTR) AS GpStatus
DECLARE FUNCTION GdipEndContainer (BYVAL graphics AS GpGraphics PTR, BYVAL state AS GraphicsContainer) AS GpStatus
DECLARE FUNCTION GdipGetMetafileHeaderFromWmf (BYVAL hWmf AS HMETAFILE, BYVAL wmfPlaceableFileHeader AS WmfPlaceableFileHeader PTR, BYVAL header AS MetafileHeader PTR) AS GpStatus
DECLARE FUNCTION GdipGetMetafileHeaderFromEmf (BYVAL hEmf AS HENHMETAFILE, BYVAL header AS MetafileHeader PTR) AS GpStatus
DECLARE FUNCTION GdipGetMetafileHeaderFromFile (BYVAL filename AS CONST WSTRING PTR, BYVAL header AS MetafileHeader PTR) AS GpStatus
DECLARE FUNCTION GdipGetMetafileHeaderFromStream (BYVAL stream AS IStream PTR, BYVAL header AS MetafileHeader PTR) AS GpStatus
DECLARE FUNCTION GdipGetMetafileHeaderFromMetafile (BYVAL metafile AS GpMetafile PTR, BYVAL header AS MetafileHeader PTR) AS GpStatus
DECLARE FUNCTION GdipGetHemfFromMetafile (BYVAL metafile AS GpMetafile PTR, BYVAL hEmf AS HENHMETAFILE PTR) AS GpStatus
DECLARE FUNCTION GdipCreateStreamOnFile (BYVAL filename AS CONST WSTRING PTR, BYVAL access AS UINT, BYVAL stream AS IStream PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateMetafileFromWmf (BYVAL hWmf AS HMETAFILE, BYVAL deleteWmf AS BOOL, BYVAL wmfPlaceableFileHeader AS CONST WmfPlaceableFileHeader PTR, BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateMetafileFromEmf (BYVAL hEmf AS HENHMETAFILE, BYVAL deleteEmf AS BOOL, BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateMetafileFromFile (BYVAL filename AS CONST WSTRING PTR, BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateMetafileFromWmfFile (BYVAL filename AS CONST WSTRING PTR, BYVAL wmfPlaceableFileHeader AS CONST WmfPlaceableFileHeader PTR, BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateMetafileFromStream (BYVAL stream AS IStream PTR, BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipRecordMetafile (BYVAL referenceHdc AS HDC, BYVAL type AS EmfType, BYVAL frameRect AS CONST GpRectF PTR, _
                 BYVAL frameUnit AS MetafileFrameUnit, BYVAL description AS CONST WSTRING PTR, BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipRecordMetafileI (BYVAL referenceHdc AS HDC, BYVAL type AS EmfType, BYVAL frameRect AS CONST GpRect PTR, _
                 BYVAL frameUnit AS MetafileFrameUnit, BYVAL description AS CONST WSTRING PTR, BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipRecordMetafileFileName (BYVAL fileName AS CONST WSTRING PTR, BYVAL referenceHdc AS HDC, BYVAL type AS EmfType, _
                 BYVAL frameRect AS const GpRectF PTR, BYVAL frameUnit AS MetafileFrameUnit, BYVAL description AS CONST WSTRING PTR, _
                 BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipRecordMetafileFileNameI (BYVAL fileName AS CONST WSTRING PTR, BYVAL referenceHdc AS HDC, BYVAL type AS EmfType, _
                 BYVAL frameRect AS const GpRect PTR, BYVAL frameUnit AS MetafileFrameUnit, BYVAL description AS CONST WSTRING PTR, _
                 BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipRecordMetafileStream (BYVAL stream AS IStream PTR, BYVAL hdc AS HDC, BYVAL type AS EmfType, _
                 BYVAL frameRect AS CONST GpRectF PTR, BYVAL frameUnit AS MetafileFrameUnit, BYVAL description AS CONST WSTRING PTR, _
                 BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipRecordMetafileStreamI (BYVAL stream AS IStream PTR, BYVAL hdc AS HDC, BYVAL type AS EmfType, _
                 BYVAL frameRect AS CONST GpRect PTR, BYVAL frameUnit AS MetafileFrameUnit, BYVAL description AS CONST WSTRING PTR, _
                 BYVAL metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipSetMetafileDownLevelRasterizationLimit (BYVAL metafile AS GpMetafile PTR, BYVAL metafileRasterizationLimitDpi AS UINT) AS GpStatus
DECLARE FUNCTION GdipGetMetafileDownLevelRasterizationLimit (BYVAL metafile AS GpMetafile PTR, BYVAL metafileRasterizationLimitDpi AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageDecodersSize (BYVAL numDecoders AS UINT PTR, BYVAL size AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageDecoders (BYVAL numDecoders AS UINT, BYVAL size AS UINT, BYVAL decoders AS ImageCodecInfo PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageEncodersSize (BYVAL numEncoders AS UINT PTR, BYVAL  size AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageEncoders (BYVAL numEncoders AS UINT, BYVAL size AS UINT, BYVAL encoders AS ImageCodecInfo PTR) AS GpStatus
DECLARE FUNCTION GdipComment (BYVAL graphics AS GpGraphics PTR, BYVAL sizeData AS UINT, BYVAL data AS CONST BYTE PTR) AS GpStatus

' // GraphicsPath APIs
DECLARE FUNCTION GdipCreatePath (BYVAL brushMode AS GpFillMode, BYVAL path AS GpPath PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreatePath2 (BYVAL points AS CONST GpPointF PTR, BYVAL types AS CONST BYTE PTR, _
                 BYVAL count AS INT_, BYVAL fillMode AS GpFillMode, BYVAL path AS GpPath PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreatePath2I (BYVAL points AS CONST GpPoint PTR, BYVAL types AS CONST BYTE PTR, _
                 BYVAL count AS INT_, BYVAL fillMode AS GpFillMode, BYVAL path AS GpPath PTR PTR) AS GpStatus
DECLARE FUNCTION GdipClonePath (BYVAL path AS GpPath PTR, BYVAL clonePath AS GpPath PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeletePath (BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipResetPath (BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipGetPointCount (BYVAL path AS GpPath PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathTypes (BYVAL path AS GpPath PTR, BYVAL types AS BYTE PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetPathPoints (BYVAL path AS GpPath PTR, BYVAL points AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetPathPointsI (BYVAL path AS GpPath PTR, BYVAL points AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetPathFillMode (BYVAL path AS GpPath PTR, BYVAL fillMode AS GpFillMode PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathFillMode (BYVAL path AS GpPath PTR, BYVAL fillMode AS GpFillMode) AS GpStatus
DECLARE FUNCTION GdipGetPathData (BYVAL path AS GpPath PTR, BYVAL pathData AS GpPathData PTR) AS GpStatus
DECLARE FUNCTION GdipStartPathFigure (BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipClosePathFigure (BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipClosePathFigures (BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathMarker (BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipClearPathMarkers (BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipReversePath (BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathLastPoint (BYVAL path AS GpPath PTR, BYVAL lastPoint AS GpPointF PTR) AS GpStatus
DECLARE FUNCTION GdipAddPathLine (BYVAL path AS GpPath PTR, BYVAL x1 AS REAL, BYVAL y1 AS REAL, BYVAL x2 AS REAL, BYVAL y2 AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathLine2 (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, BYVAL AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathArc (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL, _
                 BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathBezier (BYVAL path AS GpPath PTR, BYVAL x1 AS REAL, BYVAL y1 AS REAL, BYVAL x2 AS REAL, BYVAL y2 AS REAL, _
                 BYVAL x3 AS REAL, BYVAL y3 AS REAL, BYVAL x4 AS REAL, BYVAL y4 AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathBeziers (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathCurve (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathCurve2 (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathCurve3 (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, _
                 BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathClosedCurve (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathClosedCurve2 (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathRectangle (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathRectangles (BYVAL path AS GpPath PTR, BYVAL rects AS CONST GpRectF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathEllipse (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathPie (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL, _
                 BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathPolygon (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathPath (BYVAL path AS GpPath PTR, BYVAL addingPath AS CONST GpPath PTR, BYVAL connect AS BOOL) AS GpStatus
DECLARE FUNCTION GdipAddPathString (BYVAL path AS GpPath PTR, BYVAL string AS CONST WSTRING PTR, BYVAL length AS INT_, _
                 BYVAL family AS CONST GpFontFamily PTR, BYVAL style AS INT_, BYVAL emSize AS REAL, BYVAL layoutRect AS CONST GpRectF PTR, _
                 BYVAL format AS CONST GpStringFormat PTR) AS GpStatus
DECLARE FUNCTION GdipAddPathStringI (BYVAL path AS GpPath PTR, BYVAL string AS CONST WSTRING PTR, BYVAL length AS INT_, _
                 BYVAL family AS CONST GpFontFamily PTR, BYVAL style AS INT_, BYVAL emSize AS REAL, BYVAL layoutRect AS CONST GpRect PTR, _
                 BYVAL format AS CONST GpStringFormat PTR) AS GpStatus
DECLARE FUNCTION GdipAddPathLineI (BYVAL path AS GpPath PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathLine2I (BYVAL path AS GpPath PTR, BYVAL points AS const GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathArcI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_, _
                 BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathBezierI (BYVAL path AS GpPath PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_, _
                 BYVAL x3 AS INT_, BYVAL y3 AS INT_, BYVAL x4 AS INT_, BYVAL y4 AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathBeziersI (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathCurveI (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathCurve2I (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathCurve3I (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_, _
                 BYVAL offset AS INT_, BYVAL numberOfSegments AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathClosedCurveI (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathClosedCurve2I (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL tension AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathRectangleI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL mHeight AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathRectanglesI (BYVAL path AS GpPath PTR, BYVAL rects AS CONST GpRect PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathEllipseI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_) AS GpStatus
DECLARE FUNCTION GdipAddPathPieI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_, _
                 BYVAL startAngle AS REAL, BYVAL sweepAngle AS REAL) AS GpStatus
DECLARE FUNCTION GdipAddPathPolygonI (BYVAL path AS GpPath PTR, BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipFlattenPath (BYVAL path AS GpPath PTR, BYVAL matrix AS GpMatrix PTR, BYVAL flatness AS REAL) AS GpStatus
DECLARE FUNCTION GdipWindingModeOutline (BYVAL path AS GpPath PTR, BYVAL matrix AS GpMatrix PTR, BYVAL flatness AS REAL) AS GpStatus
DECLARE FUNCTION GdipWidenPath (BYVAL path AS GpPath PTR, BYVAL pen AS GpPen PTR, BYVAL matrix AS GpMatrix PTR, BYVAL flatness AS REAL) AS GpStatus
DECLARE FUNCTION GdipWarpPath (BYVAL path AS GpPath PTR, BYVAL matrix AS GpMatrix PTR, BYVAL points AS CONST GpPointF PTR, _
                 BYVAL count AS INT_, BYVAL srcx AS REAL, BYVAL srcy AS REAL, BYVAL srcwidth AS REAL, BYVAL srcheight AS REAL, _
                 BYVAL warpMode AS WarpMode, BYVAL flatness AS REAL) AS GpStatus
DECLARE FUNCTION GdipTransformPath (BYVAL path AS GpPath PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathWorldBounds (BYVAL path AS GpPath PTR, BYVAL bounds AS GpRectF PTR, BYVAL matrix AS CONST GpMatrix PTR, BYVAL pen AS CONST GpPen PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathWorldBoundsI (BYVAL path AS GpPath PTR, BYVAL bounds AS GpRect PTR, BYVAL matrix AS CONST GpMatrix PTR, BYVAL pen AS CONST GpPen PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisiblePathPoint (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisiblePathPointI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsOutlineVisiblePathPoint (BYVAL path AS GpPath PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL pen AS GpPen PTR, _
                 BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsOutlineVisiblePathPointI (BYVAL path AS GpPath PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL pen AS GpPen PTR, _
                 BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus


' // HatchBrush APIs
DECLARE FUNCTION GdipCreateHatchBrush (BYVAL hatchstyle AS GpHatchStyle, BYVAL forecol AS ARGB, BYVAL backcol AS ARGB, BYVAL brush AS GpHatch PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetHatchStyle (BYVAL brush AS GpHatch PTR, BYVAL hatchstyle AS GpHatchStyle PTR) AS GpStatus
DECLARE FUNCTION GdipGetHatchForegroundColor (BYVAL brush AS GpHatch PTR, BYVAL forecol AS ARGB PTR) AS GpStatus
DECLARE FUNCTION GdipGetHatchBackgroundColor (BYVAL brush AS GpHatch PTR, BYVAL backcol AS ARGB PTR) AS GpStatus


' // Image APIs
DECLARE FUNCTION GdipLoadImageFromStream (BYVAL stream AS IStream PTR, BYVAL image AS GpImage PTR PTR) AS GpStatus
DECLARE FUNCTION GdipLoadImageFromFile (BYVAL filename AS CONST WSTRING PTR, BYVAL image AS GpImage PTR PTR) AS GpStatus
DECLARE FUNCTION GdipLoadImageFromStreamICM (BYVAL stream AS IStream PTR, BYVAL image AS GpImage PTR PTR) AS GpStatus
DECLARE FUNCTION GdipLoadImageFromFileICM (BYVAL filename AS CONST WSTRING PTR, BYVAL image AS GpImage PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCloneImage (BYVAL image AS GpImage PTR, BYVAL cloneImage AS GpImage PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDisposeImage (BYVAL image AS GpImage PTR) AS GpStatus
DECLARE FUNCTION GdipSaveImageToFile (BYVAL image AS GpImage PTR, BYVAL filename AS CONST WSTRING PTR, BYVAL clsidEncoder AS CONST CLSID PTR, _
                 BYVAL encoderParams AS CONST EncoderParameters PTR) AS GpStatus
DECLARE FUNCTION GdipSaveImageToStream (BYVAL image AS GpImage PTR, BYVAL stream AS IStream PTR, BYVAL clsidEncoder AS CONST CLSID PTR, _
                 BYVAL encoderParams AS const EncoderParameters PTR) AS GpStatus
DECLARE FUNCTION GdipSaveAdd (BYVAL image AS GpImage PTR, BYVAL encoderParams AS CONST EncoderParameters PTR) AS GpStatus
DECLARE FUNCTION GdipSaveAddImage (BYVAL image AS GpImage PTR, BYVAL newImage AS GpImage PTR, BYVAL encoderParams AS CONST EncoderParameters PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageGraphicsContext (BYVAL image AS GpImage PTR, BYVAL graphics AS GpGraphics PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageBounds (BYVAL image AS GpImage PTR, BYVAL srcRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageDimension (BYVAL image AS GpImage PTR, BYVAL width AS REAL PTR, BYVAL height AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageType (BYVAL image AS GpImage PTR, BYVAL type AS ImageType PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageWidth (BYVAL image AS GpImage PTR, BYVAL width AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageHeight (BYVAL image AS GpImage PTR, BYVAL height AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageHorizontalResolution (BYVAL image AS GpImage PTR, BYVAL resolution AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageVerticalResolution (BYVAL IMAGE AS GpImage PTR, BYVAL resolution AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageFlags (BYVAL image AS GpImage PTR, BYVAL flags AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageRawFormat (BYVAL image AS GpImage PTR, BYVAL format AS GUID PTR) AS GpStatus
DECLARE FUNCTION GdipGetImagePixelFormat (BYVAL image AS GpImage PTR, BYVAL format AS PixelFormat PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageThumbnail (BYVAL image AS GpImage PTR, BYVAL thumbWidth AS UINT, BYVAL thumbHeight AS UINT, _
                 BYVAL thumbImage AS GpImage PTR PTR, BYVAL callback AS GetThumbnailImageAbort, BYVAL callbackData AS ANY PTR) AS GpStatus
DECLARE FUNCTION GdipGetEncoderParameterListSize (BYVAL image AS GpImage PTR, BYVAL clsidEncoder AS CONST CLSID PTR, BYVAL size AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetEncoderParameterList (BYVAL image AS GpImage PTR, BYVAL clsidEncoder AS CONST CLSID PTR, _
                 BYVAL size AS UINT, BYVAL buffer AS EncoderParameters PTR) AS GpStatus
DECLARE FUNCTION GdipImageGetFrameDimensionsCount (BYVAL image AS GpImage PTR, BYVAL count AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipImageGetFrameDimensionsList (BYVAL image AS GpImage PTR, BYVAL dimensionIDs AS GUID PTR, BYVAL count AS UINT) AS GpStatus
DECLARE FUNCTION GdipImageGetFrameCount (BYVAL image AS GpImage PTR, BYVAL dimensionID AS CONST GUID PTR, BYVAL count AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipImageSelectActiveFrame (BYVAL image AS GpImage PTR, BYVAL dimensionID AS CONST GUID PTR, BYVAL frameIndex AS UINT) AS GpStatus
DECLARE FUNCTION GdipImageRotateFlip (BYVAL image AS GpImage PTR, BYVAL rfType AS RotateFlipType) AS GpStatus
DECLARE FUNCTION GdipGetImagePalette (BYVAL image AS GpImage PTR, BYVAL palette AS ColorPalette PTR, BYVAL size AS INT_) AS GpStatus
DECLARE FUNCTION GdipSetImagePalette (BYVAL image AS GpImage PTR, BYVAL palette AS CONST ColorPalette PTR) AS GpStatus
DECLARE FUNCTION GdipGetImagePaletteSize (BYVAL image AS GpImage PTR, BYVAL size AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetPropertyCount (BYVAL image AS GpImage PTR, BYVAL numOfProperty AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetPropertyIdList (BYVAL image AS GpImage PTR, BYVAL numOfProperty AS UINT, BYVAL list AS PROPID PTR) AS GpStatus
DECLARE FUNCTION GdipGetPropertyItemSize (BYVAL image AS GpImage PTR, BYVAL propiId AS PROPID, BYVAL size AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetPropertyItem (BYVAL image AS GpImage PTR, BYVAL propId AS PROPID, BYVAL propSize AS UINT, BYVAL buffer AS PropertyItem PTR) AS GpStatus
DECLARE FUNCTION GdipGetPropertySize (BYVAL image AS GpImage PTR, BYVAL totalBufferSize AS UINT PTR, BYVAL numProperties AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetAllPropertyItems (BYVAL image AS GpImage PTR, BYVAL totalBufferSize AS UINT, BYVAL numProperties AS UINT, BYVAL allItems AS PropertyItem PTR) AS GpStatus
DECLARE FUNCTION GdipRemovePropertyItem (BYVAL image AS GpImage PTR, BYVAL propId AS PROPID) AS GpStatus
DECLARE FUNCTION GdipSetPropertyItem (BYVAL image AS GpImage PTR, BYVAL item AS CONST PropertyItem PTR) AS GpStatus
' #if (GDIPVER >= 0x0110)
DECLARE FUNCTION GdipFindFirstImageItem (BYVAL image AS GpImage PTR, BYVAL item AS ImageItemData PTR) AS GpStatus
DECLARE FUNCTION GdipFindNextImageItem (BYVAL image AS GpImage PTR, BYVAL item AS ImageItemData PTR) AS GpStatus
DECLARE FUNCTION GdipGetImageItemData (BYVAL image AS GpImage PTR, BYVAL item AS ImageItemData PTR) AS GpStatus
' #endif //(GDIPVER >= 0x0110)
DECLARE FUNCTION GdipImageForceValidation (BYVAL image AS GpImage PTR) AS GpStatus


' // ImageAttributes APIs
DECLARE FUNCTION GdipCreateImageAttributes (BYVAL imageattr AS GpImageAttributes PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCloneImageAttributes (BYVAL imageattr AS CONST GpImageAttributes PTR, BYVAL cloneImageattr AS GpImageAttributes PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDisposeImageAttributes (BYVAL imageattr AS GpImageAttributes PTR) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesToIdentity (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType) AS GpStatus
DECLARE FUNCTION GdipResetImageAttributes (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesColorMatrix (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType, _
                 BYVAL enableFlag AS BOOL, BYVAL colorMatrix AS CONST ColorMatrix PTR, BYVAL grayMatrix AS CONST ColorMatrix PTR, _
                 BYVAL flags AS ColorMatrixFlags) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesThreshold (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType, _
                 BYVAL enableFlag AS BOOL, BYVAL threshold AS REAL) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesGamma (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType, _
                 BYVAL enableFlag AS BOOL, BYVAL gamma AS REAL) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesNoOp (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType, BYVAL enableFlag AS BOOL) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesColorKeys (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType, _
                 BYVAL enableFlag AS BOOL, BYVAL colorLow AS ARGB, BYVAL colorHigh AS ARGB) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesOutputChannel (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType, _
                 BYVAL enableFlag AS BOOL, BYVAL channelFlags AS ColorChannelFlags) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesOutputChannelColorProfile (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType, _
                 BYVAL enableFlag AS BOOL, BYVAL colorProfileFilename AS CONST WSTRING PTR) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesRemapTable (BYVAL imageattr AS GpImageAttributes PTR, BYVAL type AS ColorAdjustType, _
                 BYVAL enableFlag AS BOOL, BYVAL mapSize AS UINT, BYVAL map AS CONST GpColorMap PTR) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesWrapMode (BYVAL imageAttr AS GpImageAttributes PTR, BYVAL wrap AS WrapMode, _
                 BYVAL argb AS ARGB, BYVAL clamp AS BOOL) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesICMMode (BYVAL imageAttr AS GpImageAttributes PTR, BYVAL on AS BOOL) AS GpStatus
DECLARE FUNCTION GdipGetImageAttributesAdjustedPalette (BYVAL imageAttr AS GpImageAttributes PTR, BYVAL colorPalette AS ColorPalette PTR, _
                 BYVAL colorAdjustType AS ColorAdjustType) AS GpStatus
DECLARE FUNCTION GdipSetImageAttributesCachedBackground (BYVAL imageattr AS GpImageAttributes PTR, BYVAL enableFlag AS BOOL) AS GpStatus


' // LineBrush APIs
DECLARE FUNCTION GdipCreateLineBrush (BYVAL point1 AS CONST GpPointF PTR, BYVAL point2 AS CONST GpPointF PTR, _
                 BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL wrapMode AS GpWrapMode, BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateLineBrushI (BYVAL point1 AS CONST GpPoint PTR, BYVAL point2 AS CONST GpPoint PTR, _
                 BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, BYVAL wrapMode AS GpWrapMode, BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateLineBrushFromRect (BYVAL rc AS const GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, _
                 BYVAL mode AS LinearGradientMode, BYVAL wrapMode AS GpWrapMode, BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateLineBrushFromRectI (BYVAL rc AS CONST GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, _
                 BYVAL mode AS LinearGradientMode, BYVAL wrapMode AS GpWrapMode, BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateLineBrushFromRectWithAngle (BYVAL rc AS CONST GpRectF PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, _
                 BYVAL angle AS REAL, BYVAL isAngleScalable AS BOOL, BYVAL wrapMode AS GpWrapMode, BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateLineBrushFromRectWithAngleO (BYVAL rc AS CONST GpRect PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB, _
                 BYVAL angle AS REAL, BYVAL isAngleScalable AS BOOL, BYVAL wrapMode AS GpWrapMode, BYVAL lineGradient AS GpLineGradient PTR PTR) AS GpStatus
DECLARE FUNCTION GdipSetLineColors (BYVAL brush AS GpLineGradient PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB) AS GpStatus
DECLARE FUNCTION GdipGetLineColors (BYVAL brush AS GpLineGradient PTR, BYVAL colors AS ARGB PTR) AS GpStatus
DECLARE FUNCTION GdipGetLineRect (BYVAL brush AS GpLineGradient PTR, BYVAL rc AS GpRectF PTR) AS GpStatus
DECLARE FUNCTION GdipGetLineRectI (BYVAL brush AS GpLineGradient PTR, BYVAL rc AS GpRect PTR) AS GpStatus
DECLARE FUNCTION GdipSetLineGammaCorrection (BYVAL brush AS GpLineGradient PTR, BYVAL useGammaCorrection AS BOOL) AS GpStatus
DECLARE FUNCTION GdipGetLineGammaCorrection (BYVAL brush AS GpLineGradient PTR, BYVAL useGammaCorrection AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipGetLineBlendCount (BYVAL brush AS GpLineGradient PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetLineBlend (BYVAL brush AS GpLineGradient PTR, BYVAL blend AS REAL PTR, BYVAL positions AS REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipSetLineBlend (BYVAL brush AS GpLineGradient PTR, BYVAL blend AS CONST REAL PTR, BYVAL positions AS const REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetLinePresetBlendCount (BYVAL brush AS GpLineGradient PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetLinePresetBlend (BYVAL brush AS GpLineGradient PTR, BYVAL blend AS ARGB PTR, BYVAL positions AS REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipSetLinePresetBlend (BYVAL brush AS GpLineGradient PTR, BYVAL blend AS CONST ARGB PTR, BYVAL positions AS CONST REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipSetLineSigmaBlend (BYVAL brush AS GpLineGradient PTR, BYVAL focus AS REAL, BYVAL scale AS REAL) AS GpStatus
DECLARE FUNCTION GdipSetLineLinearBlend (BYVAL brush AS GpLineGradient PTR, BYVAL focus AS REAL, BYVAL scale AS REAL) AS GpStatus
DECLARE FUNCTION GdipSetLineWrapMode (BYVAL brush AS GpLineGradient PTR, BYVAL wrapmode AS GpWrapMode) AS GpStatus
DECLARE FUNCTION GdipGetLineWrapMode (BYVAL brush AS GpLineGradient PTR, BYVAL wrapmode AS GpWrapMode PTR) AS GpStatus
DECLARE FUNCTION GdipGetLineTransform (BYVAL brush AS GpLineGradient PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipSetLineTransform (BYVAL brush AS GpLineGradient PTR, BYVAL matrix AS CONST GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipResetLineTransform (BYVAL brush AS GpLineGradient PTR) AS GpStatus
DECLARE FUNCTION GdipMultiplyLineTransform (BYVAL brush AS GpLineGradient PTR, BYVAL matrix AS CONST GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipTranslateLineTransform (BYVAL brush AS GpLineGradient PTR, BYVAL dx AS REAL, BYVAL dy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipScaleLineTransform (BYVAL brush AS GpLineGradient PTR, BYVAL sx AS REAL, BYVAL sy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipRotateLineTransform (BYVAL brush AS GpLineGradient PTR, BYVAL angle AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus


' // Matrix APIs
DECLARE FUNCTION GdipCreateMatrix (BYVAL matrix AS GpMatrix PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateMatrix2 (BYVAL m11 AS REAL, BYVAL m12 AS REAL, BYVAL m21 AS REAL, BYVAL m22 AS REAL, _
                 BYVAL dx AS REAL, BYVAL dy AS REAL, BYVAL matrix AS GpMatrix PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateMatrix3 (BYVAL rcf AS CONST GpRectF PTR, BYVAL dstplg AS CONST GpPointF PTR, BYVAL matrix AS GpMatrix PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateMatrix3I (BYVAL rc AS CONST GpRect PTR, BYVAL dstplg AS CONST GpPoint PTR, BYVAL matrix AS GpMatrix PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCloneMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL cloneMatrix AS GpMatrix PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteMatrix (BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipSetMatrixElements (BYVAL matrix AS GpMatrix PTR, BYVAL m11 AS REAL, BYVAL m12 AS REAL, _
                 BYVAL m21 AS REAL, BYVAL m22 AS REAL, BYVAL dx AS REAL, BYVAL dy AS REAL) AS GpStatus
DECLARE FUNCTION GdipMultiplyMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL matrix2 AS GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipTranslateMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL offsetX AS REAL, BYVAL offsetY AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipScaleMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL scaleX AS REAL, BYVAL scaleY AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipRotateMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL angle AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipShearMatrix (BYVAL matrix AS GpMatrix PTR, BYVAL shearX AS REAL, BYVAL shearY AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipInvertMatrix (BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipTransformMatrixPoints (BYVAL matrix AS GpMatrix PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipTransformMatrixPointsI (BYVAL matrix AS GpMatrix PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipVectorTransformMatrixPoints (BYVAL matrix AS GpMatrix PTR, BYVAL pts AS GpPointF PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipVectorTransformMatrixPointsI (BYVAL matrix AS GpMatrix PTR, BYVAL pts AS GpPoint PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetMatrixElements (BYVAL matrix AS CONST GpMatrix PTR, BYVAL matrixOut AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipIsMatrixInvertible (BYVAL matrix AS CONST GpMatrix PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsMatrixIdentity (BYVAL matrix AS CONST GpMatrix PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsMatrixEqual (BYVAL matrix AS CONST GpMatrix PTR, BYVAL matrix2 AS CONST GpMatrix PTR, BYVAL result AS BOOL PTR) AS GpStatus

' #if (GDIPVER >= 0x0110)
DECLARE FUNCTION GdipConvertToEmfPlus (BYVAL refGraphics AS CONST GpGraphics PTR, BYVAL metafile AS GpMetafile PTR, _
                 BYVAL conversionFailureFlag AS INT_ PTR, BYVAL emfType AS EmfType, BYVAL description AS CONST WSTRING PTR, _
                 BYVAL out_metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipConvertToEmfPlusToFile (BYVAL refGraphics AS CONST GpGraphics PTR, BYVAL metafile AS GpMetafile PTR, _
                 BYVAL conversionFailureFlag AS INT_ PTR, BYVAL filename AS CONST WSTRING PTR, BYVAL AS EmfType, _
                 BYVAL description AS CONST WSTRING PTR, BYVAL out_metafile AS GpMetafile PTR PTR) AS GpStatus
DECLARE FUNCTION GdipConvertToEmfPlusToStream (BYVAL refGraphics AS CONST GpGraphics PTR, BYVAL metafile AS GpMetafile PTR, _
                 BYVAL conversionFailureFlag AS INT_ PTR, BYVAL stream AS IStream PTR, BYVAL emfType AS EmfType, _
                 BYVAL description AS CONST WSTRING PTR, BYVAL out_metafile AS GpMetafile PTR PTR) AS GpStatus
' #endif //(GDIPVER >= 0x0110)

DECLARE FUNCTION GdipEmfToWmfBits (BYVAL hemf AS HENHMETAFILE, BYVAL cbData16 AS UINT, BYVAL pData16 AS LPBYTE, _
                 BYVAL iMapMode AS INT_, BYVAL eFlags AS INT_) as UINT


' // PathGradientBrush APIs
DECLARE FUNCTION GdipCreatePathGradient (BYVAL points AS CONST GpPointF PTR, BYVAL count AS INT_, BYVAL wrapMode AS GpWrapMode, BYVAL polyGradient AS GpPathGradient PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreatePathGradientI (BYVAL points AS CONST GpPoint PTR, BYVAL count AS INT_, BYVAL wrapMode AS GpWrapMode, BYVAL polyGradient AS GpPathGradient PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreatePathGradientFromPath (BYVAL path AS CONST GpPath PTR, BYVAL polyGradient AS GpPathGradient PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientCenterColor (BYVAL brush AS GpPathGradient PTR, BYVAL colors AS ARGB PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientCenterColor (BYVAL brush AS GpPathGradient PTR, BYVAL colors AS ARGB) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientSurroundColorsWithCount (BYVAL brush AS GpPathGradient PTR, BYVAL colors AS ARGB PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientSurroundColorsWithCount (BYVAL brush AS GpPathGradient PTR, BYVAL colors AS CONST ARGB PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientPath (BYVAL brush AS GpPathGradient PTR, BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientPath (BYVAL brush AS GpPathGradient PTR, BYVAL path AS CONST GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientCenterPoint (BYVAL brush AS GpPathGradient PTR, BYVAL points AS GpPointF PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientCenterPointI (BYVAL brush AS GpPathGradient PTR, BYVAL points AS GpPoint PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientCenterPoint (BYVAL brush AS GpPathGradient PTR, BYVAL points AS CONST GpPointF PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientCenterPointI (BYVAL brush AS GpPathGradient PTR, BYVAL points AS CONST GpPoint PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientRect (BYVAL brush AS GpPathGradient PTR, BYVAL rcf AS GpRectF PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientRectI (BYVAL brush AS GpPathGradient PTR, BYVAL rc AS GpRect PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientPointCount (BYVAL brush AS GpPathGradient PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientSurroundColorCount (BYVAL brush AS GpPathGradient PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientGammaCorrection (BYVAL brush AS GpPathGradient PTR, BYVAL useGammaCorrection AS BOOL) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientGammaCorrection (BYVAL brush AS GpPathGradient PTR, BYVAL useGammaCorrection AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientBlendCount (BYVAL brush AS GpPathGradient PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientBlend (BYVAL brush AS GpPathGradient PTR, BYVAL blend AS REAL PTR, BYVAL positions AS REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientBlend (BYVAL brush AS GpPathGradient PTR, BYVAL blend AS CONST REAL PTR, BYVAL positions AS CONST REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientPresetBlendCount (BYVAL brush AS GpPathGradient PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientPresetBlend (BYVAL brush AS GpPathGradient PTR, BYVAL blend AS ARGB PTR, BYVAL positions AS REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientPresetBlend (BYVAL brush AS GpPathGradient PTR, BYVAL blend AS CONST ARGB PTR, BYVAL positions AS CONST REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientSigmaBlend (BYVAL brush AS GpPathGradient PTR, BYVAL focus AS REAL, BYVAL scale AS REAL) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientLinearBlend (BYVAL brush AS GpPathGradient PTR, BYVAL focus AS REAL, BYVAL scale AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientWrapMode (BYVAL brush AS GpPathGradient PTR, BYVAL wrapMode AS GpWrapMode PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientWrapMode (BYVAL brush AS GpPathGradient PTR, BYVAL wrapMode AS GpWrapMode) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientTransform (BYVAL brush AS GpPathGradient PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientTransform (BYVAL brush AS GpPathGradient PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipResetPathGradientTransform (BYVAL brush AS GpPathGradient PTR) AS GpStatus
DECLARE FUNCTION GdipMultiplyPathGradientTransform (BYVAL brush AS GpPathGradient PTR, BYVAL matrix AS CONST GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipTranslatePathGradientTransform (BYVAL brush AS GpPathGradient PTR, BYVAL dx AS REAL, BYVAL dy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipScalePathGradientTransform (BYVAL brush AS GpPathGradient PTR, BYVAL sx AS REAL, BYVAL sy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipRotatePathGradientTransform (BYVAL brush AS GpPathGradient PTR, BYVAL angle AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipGetPathGradientFocusScales (BYVAL brush AS GpPathGradient PTR, BYVAL xScale AS REAL PTR, BYVAL yScale AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipSetPathGradientFocusScales (BYVAL brush AS GpPathGradient PTR, BYVAL xScale AS REAL, BYVAL yScale AS REAL) AS GpStatus


' // PathIterator APIs
DECLARE FUNCTION GdipCreatePathIter (BYVAL iterator AS GpPathIterator PTR PTR, BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipDeletePathIter (BYVAL iterator AS GpPathIterator PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterNextSubpath (BYVAL iterator AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, _
                 BYVAL statIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR, BYVAL isClosed AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterNextSubpathPath (BYVAL iteratir AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, _
                 BYVAL path AS GpPath PTR, BYVAL isClosed AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterNextPathType (BYVAL iterator AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, _
                 BYVAL pathType AS BYTE PTR, BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterNextMarker (BYVAL iterator AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, _
                 BYVAL startIndex AS INT_ PTR, BYVAL endIndex AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterNextMarkerPath (BYVAL iterator AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, BYVAL path AS GpPath PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterGetCount (BYVAL iterator AS GpPathIterator PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterGetSubpathCount (BYVAL iterator AS GpPathIterator PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterIsValid (BYVAL iterator AS GpPathIterator PTR, BYVAL valid AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterHasCurve (BYVAL iterator AS GpPathIterator PTR, BYVAL hasCurve AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterRewind (BYVAL iterator AS GpPathIterator PTR) AS GpStatus
DECLARE FUNCTION GdipPathIterEnumerate (BYVAL iterator AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, BYVAL points AS GpPointF PTR, _
                 BYVAL types AS BYTE PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipPathIterCopyData (BYVAL iterator AS GpPathIterator PTR, BYVAL resultCount AS INT_ PTR, BYVAL points AS GpPointF PTR, _
                 BYVAL types AS BYTE PTR, BYVAL startIndex AS INT_, BYVAL endIndex AS INT_) AS GpStatus


' // Pen APIs
DECLARE FUNCTION GdipCreatePen1 (BYVAL color AS ARGB, BYVAL width AS REAL, BYVAL unit AS GpUnit, BYVAL pen AS GpPen PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreatePen2 (BYVAL brush AS GpBrush PTR, BYVAL width AS REAL, BYVAL unit AS GpUnit, BYVAL pen AS GpPen PTR PTR) AS GpStatus
DECLARE FUNCTION GdipClonePen (BYVAL pen AS GpPen PTR, BYVAL clonepen AS GpPen PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeletePen (BYVAL pen AS GpPen PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenWidth (BYVAL pen AS GpPen PTR, BYVAL width AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetPenWidth (BYVAL pen AS GpPen PTR, BYVAL width AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenUnit (BYVAL pen AS GpPen PTR, BYVAL unit AS GpUnit) AS GpStatus
DECLARE FUNCTION GdipGetPenUnit (BYVAL pen AS GpPen PTR, BYVAL unit AS GpUnit PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenLineCap197819 (BYVAL pen AS GpPen PTR, BYVAL startCap AS GpLineCap, BYVAL endCap AS GpLineCap, BYVAL dashCap AS GpDashCap) AS GpStatus
DECLARE FUNCTION GdipSetPenStartCap (BYVAL pen AS GpPen PTR, BYVAL startCap AS GpLineCap) AS GpStatus
DECLARE FUNCTION GdipSetPenEndCap (BYVAL pen AS GpPen PTR, BYVAL endCap AS GpLineCap) AS GpStatus
DECLARE FUNCTION GdipSetPenDashCap197819 (BYVAL pen AS GpPen PTR, BYVAL dashCap AS GpDashCap) AS GpStatus
DECLARE FUNCTION GdipGetPenStartCap (BYVAL pen AS GpPen PTR, BYVAL startCap AS GpLineCap PTR) AS GpStatus
DECLARE FUNCTION GdipGetPenEndCap (BYVAL pen AS GpPen PTR, BYVAL endCap AS GpLineCap PTR) AS GpStatus
DECLARE FUNCTION GdipGetPenDashCap197819 (BYVAL pen AS GpPen PTR, BYVAL dashCap AS GpDashCap PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenLineJoin (BYVAL pen AS GpPen PTR, BYVAL lineJoin AS GpLineJoin) AS GpStatus
DECLARE FUNCTION GdipGetPenLineJoin (BYVAL pen AS GpPen PTR, BYVAL lineJoin AS GpLineJoin PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenCustomStartCap (BYVAL pen AS GpPen PTR, BYVAL customCap AS GpCustomLineCap PTR) AS GpStatus
DECLARE FUNCTION GdipGetPenCustomStartCap (BYVAL pen AS GpPen PTR, BYVAL customCap AS GpCustomLineCap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenCustomEndCap (BYVAL pen AS GpPen PTR, BYVAL customCap AS GpCustomLineCap PTR) AS GpStatus
DECLARE FUNCTION GdipGetPenCustomEndCap (BYVAL pen AS GpPen PTR, BYVAL customCap AS GpCustomLineCap PTR PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenMiterLimit (BYVAL pen AS GpPen PTR, BYVAL miterLimit AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetPenMiterLimit (BYVAL pen AS GpPen PTR, BYVAL miterLimit AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenMode (BYVAL pen AS GpPen PTR, BYVAL penMode AS GpPenAlignment) AS GpStatus
DECLARE FUNCTION GdipGetPenMode (BYVAL pen AS GpPen PTR, BYVAL penMode AS GpPenAlignment PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenTransform (BYVAL pen AS GpPen PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipGetPenTransform (BYVAL pen AS GpPen PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipResetPenTransform (BYVAL pen AS GpPen PTR) AS GpStatus
DECLARE FUNCTION GdipMultiplyPenTransform (BYVAL pen AS GpPen PTR, BYVAL matrix AS CONST GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipTranslatePenTransform (BYVAL pen AS GpPen PTR, BYVAL dx AS REAL, BYVAL dy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipScalePenTransform (BYVAL pen AS GpPen PTR, BYVAL sx AS REAL, BYVAL sy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipRotatePenTransform (BYVAL pen AS GpPen PTR, BYVAL angle AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipSetPenColor (BYVAL pen AS GpPen PTR, BYVAL argb AS ARGB) AS GpStatus
DECLARE FUNCTION GdipGetPenColor (BYVAL pen AS GpPen PTR, BYVAL argb AS ARGB PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenBrushFill (BYVAL pen AS GpPen PTR, BYVAL brush AS GpBrush PTR) AS GpStatus
DECLARE FUNCTION GdipGetPenBrushFill (BYVAL pen AS GpPen PTR, BYVAL brush AS GpBrush PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetPenFillType (BYVAL pen AS GpPen PTR, BYVAL type AS GpPenType PTR) AS GpStatus
DECLARE FUNCTION GdipGetPenDashStyle (BYVAL pen AS GpPen PTR, BYVAL dashstyle AS GpDashStyle PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenDashStyle (BYVAL pen AS GpPen PTR, BYVAL dashStyle AS GpDashStyle) AS GpStatus
DECLARE FUNCTION GdipGetPenDashOffset (BYVAL pen AS GpPen PTR, BYVAL offset AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenDashOffset (BYVAL pen AS GpPen PTR, BYVAL offset AS REAL) AS GpStatus
DECLARE FUNCTION GdipGetPenDashCount (BYVAL pen AS GpPen PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenDashArray (BYVAL pen AS GpPen PTR, BYVAL dash AS CONST REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetPenDashArray (BYVAL pen AS GpPen PTR, BYVAL dash AS REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetPenCompoundCount (BYVAL pen AS GpPen PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipSetPenCompoundArray (BYVAL pen AS GpPen PTR, BYVAL dash AS CONST REAL PTR, BYVAL count AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetPenCompoundArray (BYVAL pen AS GpPen PTR, BYVAL dash AS REAL PTR, BYVAL count AS INT_) AS GpStatus


' // Region APIs
DECLARE FUNCTION GdipCreateRegion (BYVAL region AS GpRegion PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateRegionRect (BYVAL rcf AS CONST GpRectF PTR, BYVAL region AS GpRegion PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateRegionRectI (BYVAL rc AS CONST GpRect PTR, BYVAL region AS GpRegion PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateRegionPath (BYVAL path AS GpPath PTR, BYVAL region AS GpRegion PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateRegionRgnData (BYVAL regionData AS CONST BYTE PTR, BYVAL size AS INT_, BYVAL region AS GpRegion PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateRegionHrgn (BYVAL hRgn AS HRGN, BYVAL region AS GpRegion PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCloneRegion (BYVAL region AS GpRegion PTR, BYVAL cloneRegion AS GpRegion PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteRegion (BYVAL region AS GpRegion PTR) AS GpStatus
DECLARE FUNCTION GdipSetInfinite (BYVAL region AS GpRegion PTR) AS GpStatus
DECLARE FUNCTION GdipSetEmpty (BYVAL region AS GpRegion PTR) AS GpStatus
DECLARE FUNCTION GdipCombineRegionRect (BYVAL region AS GpRegion PTR, BYVAL rcf AS CONST GpRectF PTR, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipCombineRegionRectI (BYVAL region AS GpRegion PTR, BYVAL rc AS CONST GpRect PTR, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipCombineRegionPath (BYVAL region AS GpRegion PTR, BYVAL path AS GpPath PTR, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipCombineRegionRegion (BYVAL region AS GpRegion PTR, BYVAL region2 AS GpRegion PTR, BYVAL combineMode AS CombineMode) AS GpStatus
DECLARE FUNCTION GdipTranslateRegion (BYVAL region AS GpRegion PTR, BYVAL dx AS REAL, BYVAL dy AS REAL) AS GpStatus
DECLARE FUNCTION GdipTranslateRegionI (BYVAL region AS GpRegion PTR, BYVAL dx AS INT_, BYVAL dy AS INT_) AS GpStatus
DECLARE FUNCTION GdipTransformRegion (BYVAL region AS GpRegion PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipGetRegionBounds (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, BYVAL rcf AS GpRectF PTR) AS GpStatus
DECLARE FUNCTION GdipGetRegionBoundsI (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, BYVAL rc AS GpRect PTR) AS GpStatus
DECLARE FUNCTION GdipGetRegionHRgn (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, BYVAL hRgn AS HRGN PTR) AS GpStatus
DECLARE FUNCTION GdipIsEmptyRegion (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsInfiniteRegion (BYVAL region AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsEqualRegion (BYVAL region AS GpRegion PTR, BYVAL region2 AS GpRegion PTR, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipGetRegionDataSize (BYVAL region AS GpRegion PTR, BYVAL bufferSize AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipGetRegionData (BYVAL region AS GpRegion PTR, BYVAL buffer AS BYTE PTR, BYVAL bufferSize AS UINT, BYVAL sizeFilled AS UINT PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisibleRegionPoint (BYVAL region AS GpRegion PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisibleRegionPointI (BYVAL region AS GpRegion PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisibleRegionRect (BYVAL region AS GpRegion PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, _
                 BYVAL height AS REAL, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipIsVisibleRegionRectI (BYVAL region AS GpRegion PTR, BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, _
                 BYVAL height AS INT_, BYVAL graphics AS GpGraphics PTR, BYVAL result AS BOOL PTR) AS GpStatus
DECLARE FUNCTION GdipGetRegionScansCount (BYVAL region AS GpRegion PTR, BYVAL count AS UINT PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipGetRegionScans (BYVAL region AS GpRegion PTR, BYVAL rects AS GpRectF PTR, BYVAL count AS INT_ PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipGetRegionScansI (BYVAL region AS GpRegion PTR, BYVAL rects AS GpRect PTR, BYVAL count AS INT_ PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus


' // SolidBrush APIs
DECLARE FUNCTION GdipCreateSolidFill (BYVAL color AS ARGB, BYVAL brush AS GpSolidFill PTR PTR) AS GpStatus
DECLARE FUNCTION GdipSetSolidFillColor (BYVAL brush AS GpSolidFill PTR, BYVAL color AS ARGB) AS GpStatus
DECLARE FUNCTION GdipGetSolidFillColor (BYVAL brush AS GpSolidFill PTR, BYVAL color AS ARGB PTR) AS GpStatus


' // String format APIs
DECLARE FUNCTION GdipCreateStringFormat (BYVAL formatAttributes AS INT_, BYVAL language AS LANGID, BYVAL format AS GpStringFormat PTR PTR) AS GpStatus
DECLARE FUNCTION GdipStringFormatGetGenericDefault (BYVAL format AS GpStringFormat PTR PTR) AS GpStatus
DECLARE FUNCTION GdipStringFormatGetGenericTypographic (BYVAL format AS GpStringFormat PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDeleteStringFormat (BYVAL format AS GpStringFormat PTR) AS GpStatus
DECLARE FUNCTION GdipCloneStringFormat (BYVAL format AS CONST GpStringFormat PTR, BYVAL newFormat AS GpStringFormat PTR PTR) AS GpStatus
DECLARE FUNCTION GdipSetStringFormatFlags (BYVAL format AS GpStringFormat PTR, BYVAL flags AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetStringFormatFlags (BYVAL format AS CONST GpStringFormat PTR, BYVAL flags AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipSetStringFormatAlign (BYVAL format AS GpStringFormat PTR, BYVAL align AS StringAlignment) AS GpStatus
DECLARE FUNCTION GdipGetStringFormatAlign (BYVAL format AS CONST GpStringFormat PTR, BYVAL align AS StringAlignment PTR) AS GpStatus
DECLARE FUNCTION GdipSetStringFormatLineAlign (BYVAL format AS GpStringFormat PTR, BYVAL align AS StringAlignment) AS GpStatus
DECLARE FUNCTION GdipGetStringFormatLineAlign (BYVAL format AS CONST GpStringFormat PTR, BYVAL align AS StringAlignment PTR) AS GpStatus
DECLARE FUNCTION GdipSetStringFormatTrimming (BYVAL format AS GpStringFormat PTR, BYVAL trimming AS StringTrimming) AS GpStatus
DECLARE FUNCTION GdipGetStringFormatTrimming (BYVAL format AS CONST GpStringFormat PTR, BYVAL trimming AS StringTrimming PTR) AS GpStatus
DECLARE FUNCTION GdipSetStringFormatHotkeyPrefix (BYVAL format AS GpStringFormat PTR, BYVAL hotkeyPrefix AS INT_) AS GpStatus
DECLARE FUNCTION GdipGetStringFormatHotkeyPrefix (BYVAL format AS CONST GpStringFormat PTR, BYVAL hotkeyPrefix AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipSetStringFormatTabStops (BYVAL format AS GpStringFormat PTR, BYVAL firstTabOffset AS REAL, _
                 BYVAL count AS INT_, BYVAL tabStops AS CONST REAL PTR) AS GpStatus
DECLARE FUNCTION GdipGetStringFormatTabStops (BYVAL format AS CONST GpStringFormat PTR, BYVAL count AS INT_, _
                 BYVAL firstTabOffset AS REAL PTR, BYVAL tabStops AS REAL PTR) AS GpStatus
DECLARE FUNCTION GdipGetStringFormatTabStopCount (BYVAL format AS CONST GpStringFormat PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipSetStringFormatDigitSubstitution (BYVAL format AS GpStringFormat PTR, BYVAL language AS LANGID, BYVAL substitute AS StringDigitSubstitute) AS GpStatus
DECLARE FUNCTION GdipGetStringFormatDigitSubstitution (BYVAL format AS CONST GpStringFormat PTR, BYVAL language AS LANGID PTR, BYVAL substitute AS StringDigitSubstitute PTR) AS GpStatus
DECLARE FUNCTION GdipGetStringFormatMeasurableCharacterRangeCount (BYVAL format AS CONST GpStringFormat PTR, BYVAL count AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipSetStringFormatMeasurableCharacterRanges (BYVAL format AS GpStringFormat PTR, BYVAL rangeCount AS INT_, BYVAL ranges AS CONST GpCharacterRange PTR) AS GpStatus


' // Text APIs
DECLARE FUNCTION GdipDrawString (BYVAL graphics AS GpGraphics PTR, BYVAL string AS CONST WSTRING PTR, BYVAL length AS INT_, _
                 BYVAL font AS CONST GpFont PTR, BYVAL layoutRect AS CONST GpRectF PTR, BYVAL stringFormat AS CONST GpStringFormat PTR, BYVAL brush AS CONST GpBrush PTR) AS GpStatus
DECLARE FUNCTION GdipMeasureString (BYVAL graphics AS GpGraphics PTR, BYVAL string AS CONST WSTRING PTR, BYVAL length AS INT_, _
                 BYVAL font AS CONST GpFont PTR, BYVAL layoutRect AS CONST GpRectF PTR, BYVAL stringFormat AS CONST GpStringFormat PTR, _
                 BYVAL boundingBox AS GpRectF PTR, BYVAL codepointsFitted AS INT_ PTR, BYVAL linesFilled AS INT_ PTR) AS GpStatus
DECLARE FUNCTION GdipMeasureCharacterRanges (BYVAL graphics AS GpGraphics PTR, BYVAL string AS CONST WSTRING PTR, BYVAL length AS INT_, _
                 BYVAL font AS GpFont PTR, BYVAL layoutRect AS GpRectF PTR, BYVAL stringFormat AS GpStringFormat PTR, _
                 BYVAL regionCount AS INT_, BYVAL regions AS GpRegion PTR PTR) AS GpStatus
DECLARE FUNCTION GdipDrawDriverString (BYVAL graphics AS GpGraphics PTR, BYVAL text AS CONST UINT16 PTR, BYVAL length AS INT_, _
                 BYVAL font AS CONST GpFont PTR, BYVAL brush AS CONST GpBrush PTR, BYVAL positions AS CONST GpPointF PTR, _
                 BYVAL flags AS INT_, BYVAL matrix AS CONST GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipMeasureDriverString (BYVAL graphics AS GpGraphics PTR, BYVAL text AS CONST UINT16 PTR, BYVAL length AS INT_, _
                 BYVAL font AS CONST GpFont PTR, BYVAL positions AS CONST GpPointF PTR, BYVAL flags AS INT_, BYVAL matrix AS CONST GpMatrix PTR, _
                 BYVAL boundingBox AS GpRectF PTR) AS GpStatus


' // TextureBrush APIs
DECLARE FUNCTION GdipCreateTexture (BYVAL image AS GpImage PTR, BYVAL wrapMode AS GpWrapMode, BYVAL texture AS GpTexture PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateTexture2 (BYVAL image AS GpImage PTR, BYVAL wrapMode AS GpWrapMode, BYVAL x AS REAL, BYVAL y AS REAL, _
                 BYVAL width AS REAL, BYVAL height AS REAL, BYVAL texture AS GpTexture PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateTexture2I (BYVAL image AS GpImage PTR, BYVAL wrapMode AS GpWrapMode, BYVAL x AS INT_, BYVAL y AS INT_, _
                 BYVAL width AS INT_, BYVAL height AS INT_, BYVAL texture AS GpTexture PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateTextureIA (BYVAL image AS GpImage PTR, BYVAL imageAttributes AS CONST GpImageAttributes PTR, _
                 BYVAL x AS REAL, BYVAL y AS REAL, BYVAL width AS REAL, BYVAL height AS REAL, BYVAL texture AS GpTexture PTR PTR) AS GpStatus
DECLARE FUNCTION GdipCreateTextureIAI (BYVAL image AS GpImage PTR, BYVAL imageAttributes AS CONST GpImageAttributes PTR, _
                 BYVAL x AS INT_, BYVAL y AS INT_, BYVAL width AS INT_, BYVAL height AS INT_, BYVAL texture AS GpTexture PTR PTR) AS GpStatus
DECLARE FUNCTION GdipGetTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL matrix AS GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipSetTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL matrix AS CONST GpMatrix PTR) AS GpStatus
DECLARE FUNCTION GdipResetTextureTransform (BYVAL brush AS GpTexture PTR) AS GpStatus
DECLARE FUNCTION GdipMultiplyTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL matrix AS CONST GpMatrix PTR, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipTranslateTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL dx AS REAL, BYVAL dy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipScaleTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL sx AS REAL, BYVAL sy AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipRotateTextureTransform (BYVAL brush AS GpTexture PTR, BYVAL angle AS REAL, BYVAL order AS GpMatrixOrder) AS GpStatus
DECLARE FUNCTION GdipSetTextureWrapMode (BYVAL brush AS GpTexture PTR, BYVAL wrapMode AS GpWrapMode) AS GpStatus
DECLARE FUNCTION GdipGetTextureWrapMode (BYVAL brush AS GpTexture PTR, BYVAL wrapMode AS GpWrapMode PTR) AS GpStatus
DECLARE FUNCTION GdipGetTextureImage (BYVAL brush AS GpTexture PTR, BYVAL image AS GpImage PTR PTR) AS GpStatus

DECLARE FUNCTION GdipTestControl (BYVAL AS GpTestControlEnum, BYVAL AS any PTR) AS GpStatus

' #define __GDIPLUS_EFFECTS_H

type CurveAdjustments as long
enum
	AdjustExposure = 0
	AdjustDensity = 1
	AdjustContrast = 2
	AdjustHighlight = 3
	AdjustShadow = 4
	AdjustMidtone = 5
	AdjustWhiteSaturation = 6
	AdjustBlackSaturation = 7
end enum

type CurveChannel as long
enum
	CurveChannelAll = 0
	CurveChannelRed = 1
	CurveChannelGreen = 2
	CurveChannelBlue = 3
end enum

type BlurParams
	radius as REAL
	expandEdge as BOOL
end type

type BrightnessContrastParams
	brightnessLevel as INT_
	contrastLevel as INT_
end type

type ColorBalanceParams
	cyanRed as INT_
	magentaGreen as INT_
	yellowBlue as INT_
end type

type ColorCurveParams
	adjustment as CurveAdjustments
	channel as CurveChannel
	adjustValue as INT_
end type

type ColorLUTParams
	lutB(0 to 255) as UBYTE
	lutG(0 to 255) as UBYTE
	lutR(0 to 255) as UBYTE
	lutA(0 to 255) as UBYTE
end type

type HueSaturationLightnessParams
	hueLevel as INT_
	saturationLevel as INT_
	lightnessLevel as INT_
end type

type LevelsParams
	highlight as INT_
	midtone as INT_
	shadow as INT_
end type

type RedEyeCorrectionParams
	numberOfAreas as UINT
	areas as RECT ptr
end type

type SharpenParams
	radius as REAL
	amount as REAL
end type

type TintParams
	hue as INT_
	amount as INT_
end type

'extern BlurEffectGuid as const GUID
'extern BrightnessContrastEffectGuid as const GUID
'extern ColorBalanceEffectGuid as const GUID
'extern ColorCurveEffectGuid as const GUID
'extern ColorLUTEffectGuid as const GUID
'extern ColorMatrixEffectGuid as const GUID
'extern HueSaturationLightnessEffectGuid as const GUID
'extern LevelsEffectGuid as const GUID
'extern RedEyeCorrectionEffectGuid as const GUID
'extern SharpenEffectGuid as const GUID
'extern TintEffectGuid as const GUID

' BlurEffect - {633C80A4-1843-482b-9EF2-BE2834C5FDD4}
Dim Shared BlurEffectGuid As GUID = Type(&h633C80A4, &h1843, &h482B, {&h9E, &hF2, &hBE, &h28, &h34, &hC5, &hFD, &hD4})
' BrightnessContrastEffect - {D3A1DBE1-8EC4-4c17-9F4C-EA97AD1C343D}
Dim Shared BrightnessContrastEffectGuid As GUID = Type(&hD3A1DBE1, &h8EC4, &h4c17, {&h9F, &h4C, &hEA, &h97, &hAD, &h1C, &h34, &h3D})
' ColorBalanceEffect - {537E597D-251E-48da-9664-29CA496B70F8}
Dim Shared ColorBalanceEffectGuid As GUID = Type(&h537E597D, &h251E, &h48DA, {&h96, &h64, &h29, &hCA, &h49, &h68, &h70, &hF8})
' ColorCurveEffect - {DD6A0022-58E4-4a67-9D9B-D48EB881A53D}
Dim Shared ColorCurveEffectGuid As GUID = Type(&hDD6A0022, &h58E4, &h4a67, {&h90, &h98, &hD4, &h8E, &hB8, &h81, &hA5, &h3D})
' ColorLUTEffect - {A7CE72A9-0F7F-40d7-B3CC-D0C02D5C3212}
Dim Shared ColorLUTEffectGuid As GUID = Type(&hA7CE72A9, &h0F7F, &h40d7, {&hB3, &hCC, &hD0, &hC0, &h2D, &h5C, &h32, &h12})
' ColorMatrixEffect - {718F2615-7933-40e3-A511-5F68FE14DD74}
Dim Shared ColorMatrixEffectGuid As GUID = Type(&h718F2615, &h7933, &h40e3, {&hA5, &h11, &h5F, &h68, &hFE, &h14, &hDD, &h74})
' HueSaturationLightnessEffect - {8B2DD6C3-EB07-4d87-A5F0-7108E26A9C5F}
Dim Shared HueSaturationLightnessEffectGuid As GUID = Type(&h8B2DD6C3, &hEB07, &h4d87, {&hA5, &hF0, &h71, &h08, &hE2, &h6A, &h9C, &h5F})
' LevelsEffect - {99C354EC-2A31-4f3a-8C34-17A803B33A25}
Dim Shared LevelsEffectGuid As GUID = Type(&h99C354EC, &h2A31, &h4F3A, {&h8C, &h34, &h17, &hA8, &h03, &hB3, &h3A, &h25})
' RedEyeCorrectionEffect - {74D29D05-69A4-4266-9549-3CC52836B632}
Dim Shared RedEyeCorrectionEffectGuid As GUID = Type(&h74D29D05, &h69A4, &h4266, {&h95, &h49, &h3C, &hC5, &h28, &h36, &hB6, &h32})
' SharpenEffect - {63CBF3EE-C526-402c-8F71-62C540BF5142}
Dim Shared SharpenEffectGuid As GUID = Type(&h63CBF3EE, &hC526, &h402c, {&h8F, &h71, &h62, &hC5, &h40, &hBF, &h51, &h42})
' TintEffect - {1077AF00-2848-4441-9489-44AD4C2D7A2C}
Dim Shared TintEffectGuid As GUID = Type(&h1077AF00, &h2848, &h4441, {&h94, &h89, &h44, &hAD, &h4C, &h2D, &h7A, &h2C})

' #define __GDIPLUS_IMAGECODEC_H
#define GetImageDecoders(numDecoders, size, decoders) cast(GpStatus, GdipGetImageDecoders((numDecoders), (size), (decoders)))
#define GetImageDecodersSize(numDecoders, size) cast(GpStatus, GdipGetImageDecodersSize((numDecoders), (size)))
#define GetImageEncoders(numEncoders, size, encoders) cast(GpStatus, GdipGetImageEncoders((numEncoders), (size), (encoders)))
#define GetImageEncodersSize(numEncoders, size) cast(GpStatus, GdipGetImageEncodersSize((numEncoders), (size)))

end extern
