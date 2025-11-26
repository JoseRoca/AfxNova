' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxGdiPlusObjects.inc
' Purpose:  GDI+ Flat Api procedures
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
#include once "AfxNova/AfxGdipObjects.inc"
USING AfxNova

NAMESPACE AfxNova

' ########################################################################################
' GpCustomLineCap
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION GpCustomLineCap_Clone (BYVAL customCap AS GpCustomLineCap PTR) AS GpCustomLineCap PTR
   DIM clonedCap AS GpCustomLineCap PTR
   SetLastError(GdipCloneCustomLineCap(customCap, @clonedCap))
   RETURN clonedCap
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpCustomLineCap_GetType (BYVAL customCap AS GpCustomLineCap PTR) AS CustomLineCapType
   DIM capType AS CustomLineCapType
   SetLastError(GdipGetCustomLineCapType(customCap, @capType))
   RETURN capType
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpCustomLineCap_GetStrokeCaps (BYVAL customCap AS GpCustomLineCap PTR, BYVAL startCap AS GpLineCap PTR, BYVAL endCap AS GpLineCap PTR)
   SetLastError(GdipGetCustomLineCapStrokeCaps(customCap, startCap, endCap))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpCustomLineCap_SetStrokeCaps (BYVAL customCap AS GpCustomLineCap PTR, BYVAL startCap AS GpLineCap, BYVAL endCap AS GpLineCap)
   SetLastError(GdipSetCustomLineCapStrokeCaps(customCap, startCap, endCap))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpCustomLineCap_GetStrokeStartCap (BYVAL customCap AS GpCustomLineCap PTR) AS GpLineCap
   DIM AS GpLineCap startCap, endCap
   SetLastError(GdipGetCustomLineCapStrokeCaps(customCap, @startCap, @endCap))
   RETURN startCap
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpCustomLineCap_GetStrokeEndCaps (BYVAL customCap AS GpCustomLineCap PTR) AS GpLineCap
   DIM AS GpLineCap startCap, endCap
   SetLastError(GdipGetCustomLineCapStrokeCaps(customCap, @startCap, @endCap))
   RETURN endCap
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpCustomLineCap_GetStrokeJoin (BYVAL customCap AS GpCustomLineCap PTR) AS GpLineJoin
   DIM lineJoin AS GpLineJoin
   SetLastError(GdipGetCustomLineCapStrokeJoin(customCap, @lineJoin))
   RETURN lineJoin
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpCustomLineCap_SetStrokeJoin (BYVAL customCap AS GpCustomLineCap PTR, BYVAL lineJoin AS GpLineJoin)
   SetLastError(GdipSetCustomLineCapStrokeJoin(customCap, lineJoin))
END SUB
' ========================================================================================
' =====================================================================================
PRIVATE FUNCTION GpCustomLineCap_GetBaseCap (BYVAL customCap AS GpCustomLineCap PTR) AS GpLineCap
   DIM baseCap AS LineCap
   SetLastError(GdipGetCustomLineCapBaseCap(customCap, @baseCap))
   RETURN baseCap
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpCustomLineCap_SetBaseCap (BYVAL customCap AS GpCustomLineCap PTR, BYVAL baseCap AS GpLineCap)
   SetLastError(GdipSetCustomLineCapBaseCap(customCap, baseCap))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpCustomLineCap_GetBaseInset (BYVAL customCap AS GpCustomLineCap PTR) AS REAL
   DIM baseInset AS REAL
   SetLastError(GdipGetCustomLineCapBaseInset(customCap, @baseInset))
   RETURN baseInset
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpCustomLineCap_SetBaseInset (BYVAL customCap AS GpCustomLineCap PTR, BYVAL baseInset AS REAL)
   SetLastError(GdipSetCustomLineCapBaseInset(customCap, baseInset))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpCustomLineCap_GetWidthScale (BYVAL customCap AS GpCustomLineCap PTR) AS REAL
   DIM widthScale AS REAL
   SetLastError(GdipGetCustomLineCapWidthScale(customCap, @widthScale))
   RETURN widthScale
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpCustomLineCap_SetWidthScale (BYVAL customCap AS GpCustomLineCap PTR, BYVAL widthScale AS REAL)
   SetLastError(GdipSetCustomLineCapWidthScale(customCap, widthScale))
END SUB
' =====================================================================================

' ########################################################################################
' GpAdjustableArrowCap
' As AdjustableArrowCap inherits from CustomLineCap, the GpCustomLineCap functions can be
' called using a pointer to GpAdjustableArrowCap.
' ########################################################################################

#define GpAdjustableArrowCap_Clone GpCustomLineCap_Clone
#define GpAdjustableArrowCap_GetType GpCustomLineCap_GetType
#define GpAdjustableArrowCap_GetStrokeCaps GpCustomLineCap_GetStrokeCaps
#define GpAdjustableArrowCap_SetStrokeCaps GpCustomLineCap_SetStrokeCaps
#define GpAdjustableArrowCap_GetStrokeStartCap GpCustomLineCap_GetStrokeStartCap
#define GpAdjustableArrowCap_GetStrokeEndCaps GpCustomLineCap_GetStrokeEndCaps
#define GpAdjustableArrowCap_GetStrokeJoin GpCustomLineCap_GetStrokeJoin
#define GpAdjustableArrowCap_SetStrokeJoin GpCustomLineCap_SetStrokeJoin
#define GpAdjustableArrowCap_GetBaseCap GpCustomLineCap_GetBaseCap
#define GpAdjustableArrowCap_SetBaseCap GpCustomLineCap_SetBaseCap
#define GpAdjustableArrowCap_GetBaseInset GpCustomLineCap_GetBaseInset
#define GpAdjustableArrowCap_SetBaseInset GpCustomLineCap_SetBaseInset
#define GpAdjustableArrowCap_GetWidthScale GpCustomLineCap_GetWidthScale
#define GpAdjustableArrowCap_SetWidthScale GpCustomLineCap_SetWidthScale

' =====================================================================================
PRIVATE FUNCTION GpAdjustableArrowCap_GetHeight (BYVAL cap AS GpAdjustableArrowCap PTR) AS REAL
   DIM nHeight AS REAL
   SetLastError(GdipGetAdjustableArrowCapHeight(cap, @nHeight))
   RETURN nHeight
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpAdjustableArrowCap_SetHeight (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL nHeight AS REAL)
   SetLastError(GdipSetAdjustableArrowCapHeight(cap, nHeight))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpAdjustableArrowCap_GetWidth (BYVAL cap AS GpAdjustableArrowCap PTR) AS REAL
   DIM nWidth AS REAL
   SetLastError(GdipGetAdjustableArrowCapWidth(cap, @nWidth))
   RETURN nWidth
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpAdjustableArrowCap_SetWidth (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL nWidth AS REAL)
   SetLastError(GdipSetAdjustableArrowCapWidth(cap, nWidth))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpAdjustableArrowCap_GetMiddleInset (BYVAL cap AS GpAdjustableArrowCap PTR) AS REAL
   DIM middleInset AS REAL
   SetLastError(GdipGetAdjustableArrowCapMiddleInset(cap, @middleInset))
   RETURN middleInset
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpAdjustableArrowCap_SetMiddleInset (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL middleInset AS REAL)
   SetLastError(GdipSetAdjustableArrowCapMiddleInset(cap, middleInset))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpAdjustableArrowCap_GetFillState (BYVAL cap AS GpAdjustableArrowCap PTR) AS BOOL
   DIM fillState AS BOOL
   SetLastError(GdipGetAdjustableArrowCapFillState(cap, @fillState))
   RETURN fillState
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpAdjustableArrowCap_SetFillState (BYVAL cap AS GpAdjustableArrowCap PTR, BYVAL fillState AS BOOL)
   SetLastError(GdipSetAdjustableArrowCapFillState(cap, fillState))
END SUB
' =====================================================================================

' ########################################################################################
' GdiPlusBrush
' ########################################################################################

' =====================================================================================
PRIVATE FUNCTION GpBrush_GetType (BYVAL brush AS GpBrush PTR) AS GpBrushType
   DIM nType AS GpBrushType
   SetLastError(GdipGetBrushType(brush, @nType))
   RETURN nType
END FUNCTION
' =====================================================================================

' ########################################################################################
' GdiPlusSolidBrush
' Extends GpBrush.
' ########################################################################################

#define GpSolidBrush_GetType GpBrush_GetType

' =====================================================================================
PRIVATE FUNCTION GpSolidBrush_GetColor (BYVAL brush AS GpSolidBrush PTR) AS ARGB
   DIM colour AS ARGB
   SetLastError(GdipGetSolidFillColor(brush, @colour))
   RETURN colour
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB CGpSolidBrush_SetColor (BYVAL brush AS GpSolidBrush PTR, BYVAL colour AS ARGB)
   SetLastError(GdipSetSolidFillColor(brush, colour))
END SUB
' =====================================================================================


' ########################################################################################
' GdiPlusHatchBrush
' Extends GpBrush.
' ########################################################################################

#define GpHatchBrush_GetType GpBrush_GetType

' =====================================================================================
' Gets the hatch style of the specified hatch brush.
' =====================================================================================
PRIVATE FUNCTION GpHatchBrush_GetHatchStyle (BYVAL brush AS GpHatchBrush PTR) AS HatchStyle
   DIM nHatchStyle AS HatchStyle
   SetLastError(GdipGetHatchStyle(brush, @nHatchStyle))
   RETURN nHatchStyle
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpHatchBrush_GetForegroundColor (BYVAL brush AS GpHatchBrush PTR) AS ARGB
   DIM colour AS ARGB
   SetLastError(GdipGetHatchForegroundColor(brush, @colour))
   RETURN colour
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpHatchBrush_GetBackgroundColor (BYVAL brush AS GpHatchBrush PTR) AS ARGB
   DIM colour AS ARGB
   SetLastError(GdipGetHatchBackgroundColor(brush, @colour))
   RETURN colour
END FUNCTION
' =====================================================================================


' #####################################################################################
' GpLinearGradientBrush interface.
' Extends GpBrush.
' #####################################################################################

#define GpLinearGradientBrush_GetType GpBrush_GetType

' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_GetBlend (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG)
   SetLastError(GdipGetLineBlend(brush, blendFactors, blendPositions, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpLinearGradientBrush_GetBlendCount (BYVAL brush AS GpLinearGradientBrush PTR) AS INT_
   DIM nCount AS INT_
   SetLastError(GdipGetLineBlendCount(brush, @nCount))
   RETURN nCount
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpLinearGradientBrush_GetGammaCorrection (BYVAL brush AS GpLinearGradientBrush PTR) AS BOOLEAN
   DIM useGammaCorrection AS BOOL
   SetLastError(GdipGetLineGammaCorrection(brush, @useGammaCorrection))
   RETURN useGammaCorrection
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_GetInterpolationColors (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG)
   SetLastError(GdipGetLinePresetBlend(brush, presetColors, blendPositions, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpLinearGradientBrush_GetInterpolationColorCount (BYVAL brush AS GpLinearGradientBrush PTR) AS INT_
   DIM nCount AS INT_
   SetLastError(GdipGetLinePresetBlendCount(brush, @nCount))
   RETURN nCount
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_GetLinearColors (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL colors AS ARGB PTR)
   SetLastError(GdipGetLineColors(brush, colors))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpLinearGradientBrush_GetRectangle (BYVAL brush AS GpLinearGradientBrush PTR) AS GpRectF
   DIM rcf AS GpRectF
   SetLastError(GdipGetLineRect(brush, @rcf))
   RETURN rcf
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_GetTransform (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL matrix AS GpMatrix PTR)
   SetLastError(GdipGetLineTransform(brush, matrix))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpLinearGradientBrush_GetWrapMode (BYVAL brush AS GpLinearGradientBrush PTR) AS GpWrapMode
   DIM nWrapMode AS GpWrapMode
   SetLastError(GdipGetLineWrapMode(brush, @nWrapMode))
   RETURN nWrapMode
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_MultiplyTransform (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL matrix AS GpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipMultiplyLineTransform(brush, matrix, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_ResetTransform (BYVAL brush AS GpLinearGradientBrush PTR)
   SetLastError(GdipResetLineTransform(brush))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_RotateTransform (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipRotateLineTransform(brush, angle, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_ScaleTransform (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipScaleLineTransform(brush, sx, sy, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_SetBlend (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG)
   SetLastError(GdipSetLineBlend(brush, blendFactors, blendPositions, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_SetBlendBellShape (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0)
   SetLastError(GdipSetLineSigmaBlend(brush, focus, scale))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_SetBlendTriangularShape (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0)
   SetLastError(GdipSetLineLinearBlend(brush, focus, scale))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_SetGammaCorrection (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL useGammaCorrection AS BOOL)
   SetLastError(GdipSetLineGammaCorrection(brush, useGammaCorrection))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_SetInterpolationColors (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS LONG)
   SetLastError(GdipSetLinePresetBlend(brush, presetColors, blendPositions, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_SetLinearColors (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL color1 AS ARGB, BYVAL color2 AS ARGB)
   SetLastError(GdipSetLineColors(brush, color1, color2))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_SetTransform (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL matrix AS GpMatrix PTR)
   SetLastError(GdipSetLineTransform(brush, matrix))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_SetWrapMode (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL nWrapMode AS WrapMode)
   SetLasTError(GdipSetLineWrapMode(brush, nWrapMode))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpLinearGradientBrush_TranslateTransform (BYVAL brush AS GpLinearGradientBrush PTR, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipTranslateLineTransform(brush, dx, dy, order))
END SUB
' =====================================================================================


' #####################################################################################
' CGpPathGradientBrush interface.
' Extends CGpBrush.
' #####################################################################################

' =====================================================================================
PRIVATE SUB GpPathGradientBrush_GetBlend (BYVAL brush AS GpPathGradient PTR, BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_)
   SetLastError(GdipGetPathGradientBlend(brush, blendFactors, blendPositions, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetBlendCount (BYVAL brush AS GpPathGradient PTR) AS INT_
   DIM nCount AS INT_
   SetLastError(GdipGetPathGradientBlendCount(brush, @nCount))
   RETURN nCount
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetCenterColor (BYVAL brush AS GpPathGradient PTR) AS ARGB
   DIM colour AS ARGB
   SetLastError(GdipGetPathGradientCenterColor(brush, @colour))
   RETURN colour
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetCenterPoint (BYVAL brush AS GpPathGradient PTR) AS GpPointF
   DIM pt AS GpPointF
   SetLastError(GdipGetPathGradientCenterPoint(brush, @pt))
   RETURN pt
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_GetFocusScales (BYVAL brush AS GpPathGradient PTR, BYVAL xScale AS SINGLE PTR, BYVAL yScale AS SINGLE PTR)
   SetLastError(GdipGetPathGradientFocusScales(brush, xScale, yScale))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetGammaCorrection (BYVAL brush AS GpPathGradient PTR) AS BOOLEAN
   DIM useGammaCorrection AS BOOL
   SetLastError(GdipGetPathGradientGammaCorrection(brush, @useGammaCorrection))
   RETURN useGammaCorrection
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetGraphicsPath (BYVAL brush AS GpPathGradient PTR) AS GpGraphicsPath PTR
   DIM path AS GpPath PTR
   SetLastError(GdipGetPathGradientPath(brush, @path))
   RETURN path
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetInterpolationColorCount (BYVAL brush AS GpPathGradient PTR) AS INT_
   DIM nCount AS INT_
   SetLastError(GdipGetPathGradientPresetBlendCount(brush, @nCount))
   RETURN nCount
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB CGpPathGradientBrush_GetInterpolationColors (BYVAL brush AS GpPathGradient PTR, BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_)
   SetLastError(GdipGetPathGradientPresetBlend(brush, presetColors, blendPositions, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetPointCount (BYVAL brush AS GpPathGradient PTR) AS INT_
   DIM nCount AS INT_
   SetLastError(GdipGetPathGradientPointCount(brush, @nCount))
   RETURN nCount
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetRectangle (BYVAL brush AS GpPathGradient PTR) AS GpRectF
   DIM rc AS GpRectF
   SetLasTError(GdipGetPathGradientRect(brush, @rc))
   RETURN rc
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetSurroundColorCount (BYVAL brush AS GpPathGradient PTR) AS INT_
   DIM nCount AS INT_
   SetLastError(GdipGetPathGradientSurroundColorCount(brush, @nCount))
   RETURN nCount
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_GetSurroundColors (BYVAL brush AS GpPathGradient PTR, BYVAL colors AS ARGB PTR, BYVAL count AS INT_ PTR)
   SetLastError(GdipGetPathGradientSurroundColorsWithCount(brush, colors, count))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_GetTransform (BYVAL brush AS GpPathGradient PTR, BYVAL matrix AS GpMatrix PTR)
   SetLastError(GdipGetPathGradientTransform(brush, matrix))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPathGradientBrush_GetWrapMode (BYVAL brush AS GpPathGradient PTR) AS WrapMode
   DIM nWrapMode AS WrapMode
   SetLastError(GdipGetPathGradientWrapMode(brush, @nWrapMode))
   FUNCTION = nWrapMode
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_MultiplyTransform (BYVAL brush AS GpPathGradient PTR, BYVAL matrix AS GpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipMultiplyPathGradientTransform(brush, matrix, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_ResetTransform (BYVAL brush AS GpPathGradient PTR)
   SetLastError(GdipResetPathGradientTransform(brush))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_RotateTransform (BYVAL brush AS GpPathGradient PTR, BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipRotatePathGradientTransform(brush, angle, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_ScaleTransform (BYVAL brush AS GpPathGradient PTR, BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipScalePathGradientTransform(brush, sx, sy, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetBlend (BYVAL brush AS GpPathGradient PTR, BYVAL blendFactors AS SINGLE PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_)
   SetLastError(GdipSetPathGradientBlend(brush, blendFactors, blendPositions, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetBlendBellShape (BYVAL brush AS GpPathGradient PTR, BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0)
   SetLastError(GdipSetPathGradientSigmaBlend(brush, focus, scale))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetBlendTriangularShape (BYVAL brush AS GpPathGradient PTR, BYVAL focus AS SINGLE, BYVAL scale AS SINGLE = 1.0)
   SetLastError(GdipSetPathGradientLinearBlend(brush, focus, scale))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetCenterColor (BYVAL brush AS GpPathGradient PTR, BYVAL colour AS ARGB)
   SetLastError(GdipSetPathGradientCenterColor(brush, colour))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetCenterPoint (BYVAL brush AS GpPathGradient PTR, BYVAL pt AS GpPointF PTR)
   SetLastError(GdipSetPathGradientCenterPoint(brush, pt))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetFocusScales (BYVAL brush AS GpPathGradient PTR, BYVAL xScale AS SINGLE, BYVAL yScale AS SINGLE)
   SetLastError(GdipSetPathGradientFocusScales(brush, xScale, yScale))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetGammaCorrection (BYVAL brush AS GpPathGradient PTR, BYVAL useGammaCorrection AS BOOL)
   SetLastError(GdipSetPathGradientGammaCorrection(brush, useGammaCorrection))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetGraphicsPath (BYVAL brush AS GpPathGradient PTR, BYVAL path AS GpGraphicsPath PTR)
   SetLastError(GdipGetPathGradientPath(brush, path))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetInterpolationColors (BYVAL brush AS GpPathGradient PTR, BYVAL presetColors AS ARGB PTR, BYVAL blendPositions AS SINGLE PTR, BYVAL nCount AS INT_)
   SetLastError(GdipSetPathGradientPresetBlend(brush, presetColors, blendPositions, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetSurroundColors (BYVAL brush AS GpPathGradient PTR, BYVAL colors AS ARGB PTR, BYVAL count AS INT_ PTR)
   SetLastError(GdipSetPathGradientSurroundColorsWithCount(brush, colors, count))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetTransform (BYVAL brush AS GpPathGradient PTR, BYVAL matrix AS GpMatrix PTR)
   SetLastError(GdipSetPathGradientTransform(brush, matrix))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_TranslateTransform (BYVAL brush AS GpPathGradient PTR, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipTranslatePathGradientTransform(brush, dx, dy, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPathGradientBrush_SetWrapMode (BYVAL brush AS GpPathGradient PTR, BYVAL nWrapMode AS WrapMode)
   SetLastError(GdipSetPathGradientWrapMode(brush, nWrapMode))
END SUB
' =====================================================================================


' #####################################################################################
' CGpTextureBrush interface.
' Extends CGpBrush.
' #####################################################################################

' =====================================================================================
PRIVATE FUNCTION GpTextureBrush_GetImage (BYVAL brush AS GpTextureBrush PTR) AS GpImage PTR
   DIM image AS GpImage PTR
   SetLastError(GdipGetTextureImage(brush, @image))
   RETURN image
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpTextureBrush_GetTransform (BYVAL brush AS GpTextureBrush PTR, BYVAL matrix AS GpMatrix PTR)
   SetLastError(GdipGetTextureTransform(brush, matrix))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpTextureBrush_GetWrapMode (BYVAL brush AS GpTextureBrush PTR) AS WrapMode
   DIM nWrapMode AS WrapMode
   SetLastError(GdipGetTextureWrapMode(brush, @nWrapMode))
   RETURN nWrapMode
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpTextureBrush_MultiplyTransform (BYVAL brush AS GpTextureBrush PTR, BYVAL matrix AS GpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipMultiplyTextureTransform(brush, matrix, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpTextureBrush_ResetTransform (BYVAL brush AS GpTextureBrush PTR)
   SetLastError(GdipResetTextureTransform(brush))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpTextureBrush_RotateTransform (BYVAL brush AS GpTextureBrush PTR, BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipRotateTextureTransform(brush, angle, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpTextureBrush_ScaleTransform (BYVAL brush AS GpTextureBrush PTR, BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipScaleTextureTransform(brush, sx, sy, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpTextureBrush_SetTransform (BYVAL brush AS GpTextureBrush PTR, BYVAL matrix AS GpMatrix PTR)
   SetLastError(GdipSetTextureTransform(brush, matrix))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpTextureBrush_SetWrapMode (BYVAL brush AS GpTextureBrush PTR, BYVAL nWrapMode AS WrapMode)
   SetLastError(GdipSetTextureWrapMode(brush, nWrapMode))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpTextureBrush_TranslateTransform (BYVAL brush AS GpTextureBrush PTR, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipTranslateTextureTransform(brush, dx, dy, order))
END SUB
' =====================================================================================


' ########################################################################################
' GdiPlusImage object
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION GpImage_GetEncoderParameterListSize (BYVAL image AS GpImage PTR, BYVAL clsidEncoder AS GUID PTR) AS UINT
   DIM uSize AS UINT
   SetLastError(GdipGetEncoderParameterListSize(image, clsidEncoder, @uSize))
   RETURN uSize
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetEncoderParameterList (BYVAL image AS GpImage PTR, BYVAL clsidEncoder AS GUID PTR, BYVAL uSize AS UINT, BYVAL buffer AS EncoderParameters PTR)
   SetLastError(GdipGetEncoderParameterList(image, clsidEncoder, uSize, buffer))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_Save OVERLOAD (BYVAL image AS GpImage PTR, BYVAL pwszFileName AS WSTRING PTR, BYVAL clsidEncoder AS GUID PTR, BYVAL encoderParams AS EncoderParameters PTR)
   SetLastError(GdipSaveImageToFile(image, pwszFileName, clsidEncoder, encoderParams))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_Save OVERLOAD (BYVAL image AS GpImage PTR, BYVAL stream AS IStream PTR, BYVAL clsidEncoder AS GUID PTR, BYVAL encoderParams AS EncoderParameters PTR)
   SetLastError(GdipSaveImageToStream(image, stream, clsidEncoder, encoderParams))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_SaveAdd OVERLOAD (BYVAL image AS GpImage PTR, BYVAL encoderParams AS EncoderParameters PTR)
   SetLastError(GdipSaveAdd(image, encoderParams))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_SaveAdd OVERLOAD (BYVAL image AS GpImage PTR, BYVAL newImage AS GpImage PTR, BYVAL encoderParams AS EncoderParameters PTR)
   SetLastError(GdipSaveAddImage(image, newImage, encoderParams))
END SUB
' ========================================================================================
' ========================================================================================
' Gets the type (bitmap or metafile) of this Image object.
' ========================================================================================
PRIVATE FUNCTION GpImage_GetType (BYVAL image AS GpImage PTR) AS GpImageType
   DIM nType AS GpImageType
   SetLastError(GdipGetImageType(image, @nType))
   RETURN nType
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetPhysicalDimension (BYVAL image AS GpImage PTR) AS GpSizeF
   DIM AS REAL nWidth, nHeight
   SetLastError(GdipGetImageDimension(image, @nWidth, @nHeight))
   DIM sizeF AS GpSizeF = (nWidth, nHeight)
   RETURN sizeF
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetBounds (BYVAL image AS GpImage PTR, BYVAL srcRect AS GpRectF PTR, BYVAL srcUnit AS GpUnit PTR)
   SetLastError(GdipGetImageBounds(image, srcRect, srcUnit))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetWidth (BYVAL image AS GpImage PTR) AS UINT
   DIM uWidth AS UINT
   SetLastError(GdipGetImageWidth(image, @uWidth))
   RETURN uWidth
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetHeight (BYVAL image AS GpImage PTR) AS UINT
   DIM uHeight AS UINT
   SetLastError(GdipGetImageHeight(image, @uHeight))
   RETURN uHeight
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetHorizontalResolution (BYVAL image AS GpImage PTR) AS SINGLE
   DIM resolution AS SINGLE
   SetLastError(GdipGetImageHorizontalResolution(image, @resolution))
   RETURN resolution
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetVerticalResolution (BYVAL image AS GpImage PTR) AS SINGLE
   DIM resolution AS SINGLE
   SetLastError(GdipGetImageVerticalResolution(image, @resolution))
   RETURN resolution
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetFlags (BYVAL image AS GpImage PTR) AS UINT
   DIM flags AS UINT
   SetLastError(GdipGetImageFlags(image, @flags))
   RETURN flags
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetRawFormat (BYVAL image AS GpImage PTR, BYVAL guidformat AS GUID PTR)
   SetLastError(GdipGetImageRawFormat(image, guidformat))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetPixelFormat (BYVAL image AS GpImage PTR) AS PixelFormat
   DIM pxformat AS PixelFormat
   SetLastError(GdipGetImagePixelFormat(image, @pxformat))
   RETURN pxformat
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetPaletteSize (BYVAL image AS GpImage PTR) AS INT_
   DIM nSize AS INT_
   SetLastError(GdipGetImagePaletteSize(image, @nSize))
   RETURN nSize
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetPalette (BYVAL image AS GpImage PTR, BYVAL pal AS ColorPalette PTR, BYVAL nSize AS INT_)
   SetLastError(GdipGetImagePalette(image, pal, nSize))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_SetPalette (BYVAL image AS GpImage PTR, BYVAL pal AS ColorPalette PTR)
   SetLastError(GdipSetImagePalette(image, pal))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetFrameDimensionsCount (BYVAL image AS GpImage PTR) AS UINT
   DIM nCount AS UINT
   SetLastError(GdipImageGetFrameDimensionsCount(image, @nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetFrameDimensionsList (BYVAL image AS GpImage PTR, BYVAL dimensionIDs AS GUID PTR, BYVAL nCount AS UINT)
   SetLastError(GdipImageGetFrameDimensionsList(image, dimensionIDs, nCount))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetFrameCount (BYVAL image AS GpImage PTR, BYVAL dimensionID AS GUID PTR) AS UINT
   DIM nCount AS UINT
   SetLastError(GdipImageGetFrameCount(image, dimensionID, @nCount))
   RETURN nCount
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_SelectActiveFrame (BYVAL image AS GpImage PTR, BYVAL dimensionID AS GUID PTR, BYVAL frameIndex AS UINT)
   SetLastError(GdipImageSelectActiveFrame(image, dimensionID, frameIndex))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_RotateFlip (BYVAL image AS GpImage PTR, BYVAL rotateFlipType AS RotateFlipType)
   SetLastError(GdipImageRotateFlip(image, rotateFlipType))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetPropertyCount (BYVAL image AS GpImage PTR) AS UINT
   DIM numProperty AS UINT
   SetLastError(GdipGetPropertyCount(image, @numProperty))
   RETURN numProperty
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetPropertyIdList (BYVAL image AS GpImage PTR, BYVAL numOfProperty AS UINT, BYVAL list AS PROPID PTR)
   SetLastError(GdipGetPropertyIdList(image, numOfProperty, list))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpImage_GetPropertyItemSize (BYVAL image AS GpImage PTR, BYVAL propId AS PROPID) AS UINT
   DIM nSize AS UINT
   SetLastError(GdipGetPropertyItemSize(image, propId, @nSize))
   RETURN nSize
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetPropertyItem (BYVAL image AS GpImage PTR, BYVAL propId AS PROPID, BYVAL propSize AS UINT, BYVAL buffer AS PropertyItem PTR)
   SetLastError(GdipGetPropertyItem(image, propId, propSize, buffer))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetPropertySize (BYVAL image AS GpImage PTR, BYVAL totalBufferSize AS UINT PTR, BYVAL numProperties AS UINT PTR)
   SetLastError(GdipGetPropertySize(image, totalBufferSize, numProperties))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetAllPropertyItems (BYVAL image AS GpImage PTR, BYVAL totalBufferSize AS UINT, BYVAL numProperties AS UINT, BYVAL allItems AS PropertyItem PTR)
   SetLastError(GdipGetAllPropertyItems(image, totalBufferSize, numProperties, allItems))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_RemovePropertyItem (BYVAL image AS GpImage PTR, BYVAL propId AS PROPID)
   SetLastError(GdipRemovePropertyItem(image, propId))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_SetPropertyItem (BYVAL image AS GpImage PTR, BYVAL item AS PropertyItem PTR)
   SetLastError(GdipSetPropertyItem(image, item))
END SUB
' ========================================================================================

'#if (GDIPVER >= &h0110)
' ========================================================================================
PRIVATE SUB GpImage_FindFirstItem (BYVAL image AS GpImage PTR, BYVAL item AS ImageItemData PTR)
   SetLastError(GdipFindFirstImageItem(image, item))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_FindNextItem (BYVAL image AS GpImage PTR, BYVAL item AS ImageItemData PTR)
   SetLastError(GdipFindNextImageItem(image, item))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpImage_GetItemData (BYVAL image AS GpImage PTR, BYVAL item AS ImageItemData PTR)
   SetLastError(GdipGetImageItemData(image, item))
END SUB
' ========================================================================================
'#endif   ' //(GDIPVER >= 0x0110)

' ========================================================================================
PRIVATE SUB GpImage_GetThumbnailImage (BYVAL image AS GpImage PTR, BYVAL thumbWidth AS UINT, BYVAL thumbHeight AS UINT, BYVAL thumbimage AS GpImage PTR PTR, BYVAL pCallback AS ANY PTR = NULL, BYVAL callbackData AS ANY PTR = NULL)
   SetLastError(GdipGetImageThumbnail(image, thumbWidth, thumbHeight, thumbimage, pCallback, callbackData))
END SUB
' ========================================================================================


' ########################################################################################
' GdiPlusBitmap object
' ########################################################################################

' =====================================================================================
PRIVATE FUNCTION GpBitmap_GetHBITMAP (BYVAL bmp AS GpBitmap PTR, BYVAL colorBackground AS ARGB) AS HBITMAP
   DIM hbm AS HBITMAP
   SetLastError(GdipCreateHBITMAPFromBitmap(bmp, @hbm, colorBackground))
   RETURN hbm
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpBitmap_GetHICON (BYVAL bmp AS GpBitmap PTR) AS HICON
   DIM icon AS HICON
   SetLastError(GdipCreateHICONFromBitmap(bmp, @icon))
   RETURN icon
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpBitmap_LockBits (BYVAL bmp AS GpBitmap PTR, BYVAL rc AS GpRect PTR, BYVAL flags AS UINT, BYVAL pxFormat AS PixelFormat, BYVAL lockedBitmapData AS BitmapData PTR)
   SetLastError(GdipBitmapLockBits(bmp, rc, flags, pxFormat, lockedBitmapData))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpBitmap_UnlockBits (BYVAL bmp AS GpBitmap PTR, BYVAL lockedBitmapData AS BitmapData PTR)
   SetLastError(GdipBitmapUnlockBits(bmp, lockedBitmapData))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpBitmap_GetPixel (BYVAL bmp AS GpBitmap PTR, BYVAL x AS LONG, BYVAL y AS LONG) AS ARGB
   DIM colour As ARGB
   SetLastError(GdipBitmapGetPixel(bmp, x, y, @colour))
   RETURN colour
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpBitmap_SetPixel (BYVAL bmp AS GpBitmap PTR, BYVAL x AS LONG, BYVAL y AS LONG, BYVAL colour AS ARGB)
   SetLastError(GdipBitmapSetPixel(bmp, x, y, colour))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpBitmap_SetResolution (BYVAL bmp AS GpBitmap PTR, BYVAL xdpi AS SINGLE, BYVAL ydpi AS SINGLE)
   SetLastError(GdipBitmapSetResolution(bmp, xdpi, ydpi))
END SUB
' =====================================================================================
'#if (GDIPVER >= &H0110)
' =====================================================================================
PRIVATE SUB GpBitmap_InitializePalette (BYVAL colourPalette AS ColorPalette PTR, BYVAL nPaletteType AS PaletteType, BYVAL optimalColors AS INT_, BYVAL useTransparentColor AS BOOL, BYVAL bmp AS GpBitmap PTR)
   SetLastError(GdipInitializePalette(colourPalette, nPaletteType, optimalColors, useTransparentColor, bmp))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpBitmap_ConvertFormat (BYVAL bmp AS GpBitmap PTR, BYVAL pxFormat AS PixelFormat, BYVAL nDitherType AS DitherType, BYVAL nPaletteType AS PaletteType, BYVAL colourPalette AS ColorPalette PTR, BYVAL alphaThresholdPercent AS SINGLE)
   SetLastError(GdipBitmapConvertFormat(bmp, pxFormat, nDitherType, nPaletteType, colourPalette, alphaThresholdPercent))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpBitmap_GetHistogram (BYVAL bmp AS GpBitmap PTR, BYVAL nFormat AS HistogramFormat, BYVAL NumberOfEntries AS UINT, BYVAL channel0 AS UINT PTR, BYVAL channel1 AS UINT PTR, BYVAL channel2 AS UINT PTR, BYVAL channel3 AS UINT PTR)
   SetLastError(GdipBitmapGetHistogram(bmp, nFormat, NumberOfEntries, channel0, channel1, channel2, channel3))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpBitmap_GetHistogramSize (BYVAL bmp AS GpBitmap PTR, BYVAL nFormat AS HistogramFormat) AS UINT
   DIM NumberOfEntries AS UINT
   SetLastError(GdipBitmapGetHistogramSize(nFormat, @NumberOfEntries))
   RETURN NumberOfEntries
END FUNCTION
' =====================================================================================
'#endif ' (GDIPVER >= &H0110)

' // Effects
' =====================================================================================
PRIVATE SUB GpBitmap_ApplyEffect (BYVAL bmp AS GpBitmap PTR, BYVAL effect AS GpEffect PTR, BYVAL roi AS RECT PTR, _
BYVAL useAuxData AS BOOLEAN, BYVAL auxData AS ANY PTR PTR, BYVAL auxDataSize AS INT_ PTR)
   SetLastError(GdipBitmapApplyEffect(bmp, effect, roi, useAuxData, auxData, auxDataSize))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GdipBitmap_CreateApplyEffect (BYVAL inputBitmaps AS GpBitmap PTR PTR, BYVAL numInputs AS INT_, BYVAL effect AS GpEffect PTR, _
BYVAL roi AS RECT PTR, BYVAL outputRect AS RECT PTR, BYVAL outputBitmap AS GpBitmap PTR PTR, BYVAL useAuxData AS BOOL, _
BYVAL auxData AS ANY PTR PTR, BYVAL auxDataSize AS INT_ PTR)
   SetLastError(GdipBitmapCreateApplyEffect(inputBitmaps, numInputs, effect, roi, outputRect, outputBitmap, useAuxData, auxData, auxDataSize))
END SUB
' =====================================================================================


' ########################################################################################
' GdiPlusFontFamily object
' ########################################################################################

' =====================================================================================
PRIVATE SUB GpFontFamily_GetFamilyName (BYVAL fontFamily AS GpFontFamily PTR, BYVAL pwszName AS WSTRING PTR, BYVAL language AS LANGID = LANG_NEUTRAL)
   SetLastError(GdipGetFamilyName(fontFamily, pwszName, language))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpFontFamily_IsStyleAvailable (BYVAL fontFamily AS GpFontFamily PTR, BYVAL nStyle AS INT_) AS BOOLEAN
   DIM bStyleAvailable AS BOOL
   SetLastError(GdipIsStyleAvailable(fontFamily, nStyle, @bStyleAvailable))
   RETURN bStyleAvailable
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpFontFamily_GetEmHeight (BYVAL fontFamily AS GpFontFamily PTR, BYVAL nStyle AS INT_) AS UINT16
   DIM EmHeight AS UINT16
   SetLastError(GdipGetEmHeight(fontFamily, nStyle, @EmHeight))
   RETURN EmHeight
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpFontFamily_GetCellAscent (BYVAL fontFamily AS GpFontFamily PTR, BYVAL nStyle AS INT_) AS UINT16
   DIM CellAscent AS UINT16
   SetLastError(GdipGetCellAscent(fontFamily, nStyle, @CellAscent))
   RETURN CellAscent
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpFontFamily_GetCellDescent (BYVAL fontFamily AS GpFontFamily PTR, BYVAL nStyle AS INT_) AS UINT16
   DIM CellDescent AS UINT16
   SetLastError(GdipGetCellDescent(fontFamily, nStyle, @CellDescent))
   RETURN CellDescent
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpFontFamily_GetLineSpacing (BYVAL fontFamily AS GpFontFamily PTR, BYVAL nStyle AS INT_) AS UINT16
   DIM LineSpacing AS UINT16
   SetLastError(GdipGetLineSpacing(fontFamily, nStyle, @LineSpacing))
   RETURN LineSpacing
END FUNCTION
' =====================================================================================


' ########################################################################################
' GdiPlusFont object
' ########################################################################################

' ========================================================================================
PRIVATE FUNCTION GpFont_GetLogFontA (BYVAL font AS GpFont PTR, BYVAL graphics AS GpGraphics PTR) AS LOGFONTA
   DIM logFontA AS LOGFONTA
   SetLastError(GdipGetLogFontA(font, graphics, @logFontA))
   RETURN logFontA
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpFont_GetLogFontW (BYVAL font AS GpFont PTR, BYVAL graphics AS GpGraphics PTR) AS LOGFONTW
   DIM logFontW AS LOGFONTW
   SetLastError(GdipGetLogFontW(font, graphics, @logFontW))
   RETURN logFontW
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpFont_GetFamily (BYVAL font AS GpFont PTR) AS GpFontFamily PTR
   DIM fontFamily AS GpFontFamily PTR
   SetLastError(GdipGetFamily(font, @fontFamily))
   RETURN fontFamily
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpFont_GetStyle (BYVAL font AS GpFont PTR) AS INT_
   DIM nStyle AS INT_
   SetLastError(GdipGetFontStyle(font, @nStyle))
   RETURN nStyle
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpFont_GetSize (BYVAL font AS GpFont PTR) AS SINGLE
   DIM nSize AS SINGLE
   SetLastError(GdipGetFontSize(font, @nSize))
   RETURN nSize
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpFont_GetUnit (BYVAL font AS GpFont PTR) AS GpUnit
   DIM nUnit AS GpUnit
   SetLastError(GdipGetFontUnit(font, @nUnit))
   RETURN nUnit
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpFont_GetHeight OVERLOAD (BYVAL font AS GpFont PTR, BYVAL graphics AS GpGraphics PTR) AS SINGLE
   DIM nHeight AS SINGLE
   SetLastError(GdipGetFontHeight(font, graphics, @nHeight))
   RETURN nHeight
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpFont_GetHeight OVERLOAD (BYVAL font AS GpFont PTR, BYVAL dpi AS SINGLE) AS SINGLE
   DIM nHeight AS SINGLE
   SetLastError(GdipGetFontHeightGivenDPI(font, dpi, @nHeight))
   RETURN nHeight
END FUNCTION
' ========================================================================================


' ########################################################################################
' GdiPlusPen object
' ########################################################################################

' =====================================================================================
PRIVATE FUNCTION GpPen_GetPenType (BYVAL pen AS GpPen PTR) AS GpPenType
   DIM nPenType AS GpPenType
   SetLastError(GdipGetPenFillType(pen, @nPenType))
   RETURN nPenType
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetAlignment (BYVAL pen AS GpPen PTR) AS GpPenAlignment
   DIM nPenAlignment AS GpPenAlignment
   SetLastError(GdipGetPenMode(pen, @nPenAlignment))
   RETURN nPenAlignment
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetBrush (BYVAL pen AS GpPen PTR) AS GpBrush PTR
   DIM brush AS GpBrush PTR
   SetLastError(GdipGetPenBrushFill(pen, @brush))
   RETURN brush
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetColor (BYVAL pen AS GpPen PTR) AS ARGB
   IF GpPen_GetPenType(pen) <> PenTypeSolidColor THEN
      SetLastError(WrongState)
      RETURN 0
   END IF
   DIM colour AS ARGB
   SetLastError(GdipGetPenColor(pen, @colour))
   RETURN colour
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_GetCompoundArray (BYVAL pen AS GpPen PTR, BYVAL compoundArray AS SINGLE PTR, BYVAL nCount AS INT_)
   SetLastError(GdipGetPenCompoundArray(pen, compoundArray, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetCompoundArrayCount (BYVAL pen AS GpPen PTR) AS INT_
   DIM nCount AS INT_
   SetLastError(GdipGetPenCompoundCount(pen, @nCount))
   RETURN nCount
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetCustomEndCap (BYVAL pen AS GpPen PTR) AS GpCustomLineCap PTR
   DIM customCap AS GpCustomLineCap PTR
   SetLastError(GdipGetPenCustomEndCap(pen, @customCap))
   RETURN customCap
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetCustomStartCap (BYVAL pen AS GpPen PTR) AS GpCustomLineCap PTR
   DIM customCap AS GpCustomLineCap PTR
   SetLastError(GdipGetPenCustomStartCap(pen, @customCap))
   RETURN customCap
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_GetDashPattern (BYVAL pen AS GpPen PTR, BYVAL dashArray AS SINGLE PTR, BYVAL nCount AS INT_)
   SetLastError(GdipGetPenDashArray(pen, dashArray, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetDashPatternCount (BYVAL pen AS GpPen PTR) AS INT_
   DIM nCount AS INT_
   SetLastError(GdipGetPenDashCount(pen, @nCount))
   RETURN nCount
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetMiterLimit (BYVAL pen AS GpPen PTR) AS SINGLE
   DIM miterLimit AS SINGLE
   SetLastError(GdipGetPenMiterLimit(pen, @miterLimit))
   RETURN miterLimit
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpPen_GetDashCap (BYVAL pen AS GpPen PTR) AS DashCap
   DIM nDashCap AS DashCap
   SetLastError(GdipGetPenDashCap197819(pen, @nDashCap))
   RETURN nDashCap
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetDashOffset (BYVAL pen AS GpPen PTR) AS SINGLE
   DIM dashOffset AS SINGLE
   SetLastError(GdipGetPenDashOffset(pen, @dashOffset))
   RETURN dashOffset
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetDashStyle (BYVAL pen AS GpPen PTR) AS GpDashStyle
   DIM nDashStyle AS GpDashStyle
   SetLastError(GdipGetPenDashStyle(pen, @nDashStyle))
   RETURN nDashStyle
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetEndCap (BYVAL pen AS GpPen PTR) AS GpLineCap
   DIM endCap AS GpLineCap
   SetLastError(GdipGetPenEndCap(pen, @endCap))
   RETURN endCap
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetLineJoin (BYVAL pen AS GpPen PTR) AS GpLineJoin
   DIM nLineJoin AS GpLineJoin
   SetLastError(GdipGetPenLineJoin(pen, @nLineJoin))
   RETURN nLineJoin
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION GpPen_GetStartCap (BYVAL pen AS GpPen PTR) AS GpLineCap
   DIM startCap AS GpLineCap
   SetLastError( GdipGetPenStartCap(pen, @startCap))
   RETURN startCap
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_GetTransform (BYVAL pen AS GpPen PTR, BYVAL matrix AS GpMatrix PTR)
   SetLastError(GdipGetPenTransform(pen, matrix))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE FUNCTION CGpPen_GetWidth (BYVAL pen AS GpPen PTR) AS SINGLE
   DIM nWidth AS SINGLE
   SetLasTError(GdipGetPenWidth(pen, @nWidth))
   RETURN nWidth
END FUNCTION
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_MultiplyTransform (BYVAL pen AS GpPen PTR, BYVAL matrix AS GpMatrix PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipMultiplyPenTransform(pen, matrix, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_ResetTransform (BYVAL pen AS GpPen PTR)
   SetLastError(GdipResetPenTransform(pen))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_RotateTransform (BYVAL pen AS GpPen PTR, BYVAL angle AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLasTError(GdipRotatePenTransform(pen, angle, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_ScaleTransform (BYVAL pen AS GpPen PTR, BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipScalePenTransform(pen, sx, sy, order))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetAlignment (BYVAL pen AS GpPen PTR, BYVAL nPenAlignment AS GpPenAlignment)
   SetLastError(GdipSetPenMode(pen, nPenAlignment))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetBrush (BYVAL pen AS GpPen PTR, BYVAL brush AS GpBrush PTR)
   SetLastError(GdipSetPenBrushFill(pen, brush))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetColor (BYVAL pen AS GpPen PTR, BYVAL colour AS ARGB)
   SetLastError(GdipSetPenColor(pen, colour))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetCompoundArray (BYVAL pen AS GpPen PTR, BYVAL compoundArray AS SINGLE PTR, BYVAL nCount AS INT_)
   SetLastError(GdipSetPenCompoundArray(pen, compoundArray, nCount))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetCustomEndCap (BYVAL pen AS GpPen PTR, BYVAL customLineCap AS GpCustomLineCap PTR)
   SetLastError(GdipSetPenCustomEndCap(pen, customLineCap))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetCustomStartCap (BYVAL pen AS GpPen PTR, BYVAL customLineCap AS GpCustomLineCap PTR)
   SetLastError(GdipSetPenCustomStartCap(pen, customLineCap))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetMiterLimit (BYVAL pen AS GpPen PTR, BYVAL miterLimit AS SINGLE)
   SetLastError(GdipSetPenMiterLimit(pen, miterLimit))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetDashCap (BYVAL pen AS GpPen PTR, BYVAL nDashCap AS DashCap)
   SetLastError(GdipSetPenDashCap197819(pen, nDashCap))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetDashOffset (BYVAL pen AS GpPen PTR, BYVAL dashOffset AS SINGLE)
   SetLastError(GdipSetPenDashOffset(pen, dashOffset))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetDashStyle (BYVAL pen AS GpPen PTR, BYVAL nDashStyle AS GpDashStyle)
   SetLastError(GdipSetPenDashStyle(pen, nDashStyle))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetEndCap (BYVAL pen AS GpPen PTR, BYVAL endCap AS GpLineCap)
   SetLastError(GdipSetPenEndCap(pen, endCap))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetLineCap (BYVAL pen AS GpPen PTR, BYVAL startCap AS GpLineCap, BYVAL endCap AS GpLineCap, BYVAL nDashCap AS DashCap)
   SetLastError(GdipSetPenLineCap197819(pen, startCap, endCap, nDashCap))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetLineJoin (BYVAL pen AS GpPen PTR, BYVAL nLineJoin AS GpLineJoin)
   SetLastError(GdipSetPenLineJoin(pen, nLineJoin))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetStartCap (BYVAL pen AS GpPen PTR, BYVAL startCap AS LineCap)
   SetLastError(GdipSetPenStartCap(pen, startCap))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetTransform (BYVAL pen AS GpPen PTR, BYVAL matrix AS GpMatrix PTR)
   SetLastError(GdipSetPenTransform(pen, matrix))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_SetWidth (BYVAL pen AS GpPen PTR, BYVAL nWidth AS SINGLE)
   SetLastError(GdipSetPenWidth(pen, nWidth))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpPen_TranslateTransform (BYVAL pen AS GpPen PTR, BYVAL dx AS SINGLE, BYVAL dy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipTranslatePenTransform(pen, dx, dy, order))
END SUB
' =====================================================================================


' ########################################################################################
' GdiPlusRegion object
' ########################################################################################

' =====================================================================================
PRIVATE SUB GpRegion_MakeInfinite (BYVAL region AS GpRegion PTR)
   SetLastError(GdipSetInfinite(region))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpRegion_MakeEmpty (BYVAL region AS GpRegion PTR)
   SetLastError(GdipSetEmpty(region))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpRegion_IntersectRect OVERLOAD (BYVAL region AS GpRegion PTR, BYVAL rcf AS GpRectF PTR)
   SetLastError(GdipCombineRegionRect(region, rcf, CombineModeIntersect))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpRegion_IntersectRect OVERLOAD (BYVAL region AS GpRegion PTR, BYVAL rc AS GpRect PTR)
   SetLastError(GdipCombineRegionRectI(region, rc, CombineModeIntersect))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpRegion_IntersectPath (BYVAL region AS GpRegion PTR, BYVAL path AS GpGraphicsPath PTR)
   SetLastError(GdipCombineRegionPath(region, path, CombineModeIntersect))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpRegion_IntersectRegion (BYVAL region AS GpRegion PTR, BYVAL region2 AS GpRegion PTR)
   SetLastError(GdipCombineRegionRegion(region, region2, CombineModeIntersect))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpRegion_UnionRect OVERLOAD (BYVAL region AS GpRegion PTR, BYVAL rcf AS GpRectF PTR)
   SetLastError(GdipCombineRegionRect(region, rcf, CombineModeUnion))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpRegion_UnionRect OVERLOAD (BYVAL region AS GpRegion PTR, BYVAL rc AS GpRect PTR)
   SetLastError(GdipCombineRegionRectI(region, rc, CombineModeUnion))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpRegion_UnionPath (BYVAL region AS GpRegion PTR, BYVAL path AS GpGraphicsPath PTR)
   SetLastError(GdipCombineRegionPath(region, path, CombineModeUnion))
END SUB
' =====================================================================================
' =====================================================================================
PRIVATE SUB GpRegion_UnionRegion (BYVAL region AS GpRegion PTR, BYVAL region2 AS GpRegion PTR)
   SetLastError(GdipCombineRegionRegion(region, region2, CombineModeUnion))
END SUB
' =====================================================================================

'++++++++++++++++++++

' ========================================================================================
' Get the DPI scaling ratio
' ========================================================================================
PRIVATE FUNCTION GpGraphics_DpiRatio (BYVAL graphics AS GpGraphics PTR) AS SINGLE
   DIM dpiX AS SINGLE
   SetLastError(GdipGetDpiX(graphics, @dpiX))
   DIM rxRatio AS SINGLE = dpiX / 96
   RETURN rxRatio
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpGraphics_DpiRatioX (BYVAL graphics AS GpGraphics PTR) AS SINGLE
   DIM dpiX AS SINGLE
   SetLastError(GdipGetDpiX(graphics, @dpiX))
   DIM rxRatio AS SINGLE = dpiX / 96
   RETURN rxRatio
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION GpGraphics_DpiRatioY (BYVAL graphics AS GpGraphics PTR) AS SINGLE
   DIM dpiY AS SINGLE
   SetLastError(GdipGetDpiY(graphics, @dpiY))
   DIM ryRatio AS SINGLE = dpiY / 96
   RETURN ryRatio
END FUNCTION
' ========================================================================================

' ========================================================================================
' Updates this Graphic object's current transformation matrix with the product of itself
' and a scaling matrix.
' ========================================================================================
PRIVATE SUB GpGraphics_ScaleTransform OVERLOAD (BYVAL graphics AS GpGraphics PTR, BYVAL sx AS SINGLE, BYVAL sy AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipScaleWorldTransform(graphics, sx, sy, order))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpGraphics_ScaleTransform OVERLOAD (BYVAL graphics AS GpGraphics PTR, BYVAL dpiRatio AS SINGLE, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   SetLastError(GdipScaleWorldTransform(graphics, dpiRatio, dpiRatio, order))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpGraphics_ScaleTransform OVERLOAD (BYVAL graphics AS GpGraphics PTR, BYVAL order AS MatrixOrder = MatrixOrderPrepend)
   DIM dpiX AS SINGLE
   SetLastError(GdipGetDpiX(graphics, @dpiX))
   DIM rxRatio AS SINGLE = dpiX / 96
   DIM dpiY AS SINGLE
   SetLastError(GdipGetDpiY(graphics, @dpiY))
   DIM ryRatio AS SINGLE = dpiY / 96
   SetLastError(GdipScaleWorldTransform(graphics, rxRatio, ryRatio, order))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the transformation matrix of this Graphics object.
' ========================================================================================
PRIVATE SUB GpGraphics_SetTransform (BYVAL graphics AS GpGraphics PTR, BYVAL matrix AS GpMatrix PTR)
   SetLastError(GdipSetWorldTransform(graphics, matrix))
END SUB
' ========================================================================================

' ========================================================================================
' Draws a line that connects two points.
' ========================================================================================
PRIVATE SUB GpGraphics_DrawLine OVERLOAD (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x1 AS SINGLE, BYVAL y1 AS SINGLE, BYVAL x2 AS SINGLE, BYVAL y2 AS SINGLE)
   SetLastError(GdipDrawLine(graphics, pen, x1, y1, x2, y2))
END SUB
' ========================================================================================
' ========================================================================================
' Draws a line that connects two points.
' ========================================================================================
PRIVATE SUB GpGraphics_DrawLine OVERLOAD (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL x1 AS INT_, BYVAL y1 AS INT_, BYVAL x2 AS INT_, BYVAL y2 AS INT_)
   SetLastError(GdipDrawLineI(graphics, pen, x1, y1, x2, y2))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpGraphics_DrawLine OVERLOAD (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL pt1 AS GpPointF PTR, BYVAL pt2 AS GpPointF PTR)
   SetLastError(GdipDrawLine(graphics, pen, pt1->X, pt1->Y, pt2->X, pt2->Y))
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB GpGraphics_DrawLine OVERLOAD (BYVAL graphics AS GpGraphics PTR, BYVAL pen AS GpPen PTR, BYVAL pt1 AS GpPoint PTR, BYVAL pt2 AS GpPoint PTR)
   SetLastError(GdipDrawLineI(graphics, pen, pt1->X, pt1->Y, pt2->X, pt2->Y))
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB GpGraphics_FillRectangle (BYVAL graphics AS GpGraphics PTR, BYVAL brush AS GpBrush PTR, BYVAL x AS REAL, BYVAL y AS REAL, BYVAL nWidth AS REAL, BYVAL nHeight AS REAL)
   SetLastError(GdipFillRectangle(graphics, brush, x, y, nWidth, nHeight))
END SUB
' ========================================================================================

END NAMESPACE
