' ########################################################################################
' Microsoft Windows
' File: CGpEffect.inc
' Contents: GDI+ Effects class.
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once

' ########################################################################################
' CGpEffect Class
' ########################################################################################

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpEffect
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' * Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CGpEffect
   ' // auxData is allocated by ApplyEffect.
   IF m_auxData THEN GdipFree(m_auxData)
   ' // Delete the Effect object
   IF m_pEffect THEN GdipDeleteEffect(m_pEffect)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' * Returns a pointer to the GpEffect class
' ========================================================================================
PRIVATE FUNCTION CGpEffect.GetEffectPtr () AS CGpEffect PTR
   RETURN m_pEffect
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the size of the auxiliary data
' ========================================================================================
PRIVATE FUNCTION CGpEffect.GetAuxDataSize () AS INT_
   RETURN m_auxDataSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a pointer to the auxiliary data
' ========================================================================================
PRIVATE FUNCTION CGpEffect.GetAuxData () AS ANY PTR
   RETURN m_auxData
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpEffect.UseAuxData method in GDI+ is a boolean flag for the Effect class that, when
' set to TRUE, instructs the CGpBitmap.ApplyEffect method to return a pointer to its auxiliary data.
' ========================================================================================
PRIVATE FUNCTION CGpEffect.UseAuxData () AS BOOLEAN
   RETURN m_useAuxData
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the total size, in bytes, of the parameters currently set for the effect.
' ========================================================================================
PRIVATE FUNCTION CGpEFfect.GetParameterSize () AS UINT
   DIM nSize AS UINT
   this.SetStatus(GdipGetEffectParameterSize(m_pEffect, @nSize))
   RETURN nSize
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpBlur Class
' The Blur class enables you to apply a Gaussian blur effect to a bitmap and specify the
' nature of the blur. Pass the address of a Blur object to the CGpGraphics.DrawImage method
' or to the CGpBitmap.ApplyEffect method. To specify the nature of the blur, pass a BlurParam
' structure to the Blur.SetParameters method of a Blur object.
' ########################################################################################

TYPE CGpBlur EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS BlurParams PTR) AS GpStatus
   DECLARE FUNCTION SetParameters (BYVAL radius AS REAL, BYVAL expandEdge AS BOOL) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBlur
   SetStatus(GdipCreateEffect(BlurEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' The CGpBlur.SetParameters method sets the parameters of this Blur object.
' ========================================================================================
PRIVATE FUNCTION CGpBlur.SetParameters (BYVAL params AS BlurParams PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(BlurParams)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpBlur.SetParameters (BYVAL radius AS REAL, BYVAL expandEdge AS BOOL) AS GpStatus
   DIM params AS BlurParams = TYPE(radius, expandEdge)
   DIM uSize AS UINT = SIZEOF(params)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, @params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpBlur.GetParameters method gets the current values of the parameters of this Blur object.
' ========================================================================================
PRIVATE FUNCTION CGpBlur.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpSharpen Class
' The CGpSharpen class enables you to adjust the sharpness of a bitmap. Pass the address of
' a CGpSharpen object to the Graphics.DrawImage method or to the CGpBitmap.ApplyEffect method.
' To specify the nature of the sharpening adjustment, pass the address of a SharpenParams
' structure to the CGpSharpen.SetParameters method of a Sharpen object.
' ########################################################################################

TYPE CGpSharpen EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS SharpenParams PTR) AS GpStatus
   DECLARE FUNCTION SetParameters (BYVAL radius AS REAL, BYVAL amount AS REAL) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpSharpen
   SetStatus(GdipCreateEffect(SharpenEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' The CGpSharpen.SetParameters method sets the parameters of this Sharpen object.
' ========================================================================================
PRIVATE FUNCTION CGpSharpen.SetParameters (BYVAL params AS SharpenParams PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(SharpenParams)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
' The CGpSharpen.SetParameters method sets the parameters of this Sharpen object.
' ========================================================================================
PRIVATE FUNCTION CGpSharpen.SetParameters (BYVAL radius AS REAL, BYVAL amount AS REAL) AS GpStatus
   DIM params AS SharpenParams = TYPE(radius, amount)
   DIM uSize AS UINT = SIZEOF(params)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, @params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpSharpen.GetParameters method gets the current values of the parameters of this Sharpen object.
' ========================================================================================
PRIVATE FUNCTION CGpSharpen.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpRedEyeCorrection Class
' The RedEyeCorrection class enables you to correct the red eyes that sometimes occur in
' flash photographs. Pass the address of a RedEyeCorrection object to the CGpGraphics.DrawImage
' method or to the CGpBitmap.ApplyEffect method. To specify areas of the bitmap that have red
' eyes, pass a RedEyeCorrectionParams structure to the CGpRedEyeCorrection.SetParameters method
' of a RedEyeCorrection object.
' ########################################################################################

TYPE CGpRedEyeCorrection EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS RedEyeCorrectionParams PTR) AS GpStatus
   DECLARE FUNCTION SetParameters (BYVAL numberOfAreas AS UINT, BYVAL areas AS ..RECT PTR) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpRedEyeCorrection
   SetStatus(GdipCreateEffect(RedEyeCorrectionEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
'The CGpRedEyeCorrection.SetParameters method sets the parameters of this RedEyeCorrection object.
' ========================================================================================
PRIVATE FUNCTION CGpRedEyeCorrection.SetParameters (BYVAL params AS RedEyeCorrectionParams PTR) AS GpStatus
   DIM status AS GpStatus = 2   ' Invalid parameter
   IF params THEN
      DIM inputParam AS RedEyeCorrectionParams PTR = params
      DIM uSize AS UINT = SIZEOF(RedEyeCorrectionParams) + inputParam->numberOfAreas * SIZEOF(..RECT)
      RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
   END IF
   RETURN status
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpRedEyeCorrection.SetParameters (BYVAL numberOfAreas AS UINT, BYVAL areas AS ..RECT PTR) AS GpStatus
   DIM status AS GpStatus = 2   ' Invalid parameter
   DIM params AS RedEyeCorrectionParams = TYPE(numberOfAreas, areas)
   DIM uSize AS UINT = SIZEOF(params) + numberOfAreas * SIZEOF(..RECT)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, @params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpRedEyeCorrection.GetParameters method gets the current values of the parameters of
' this RedEyeCorrection object.
' ========================================================================================
PRIVATE FUNCTION CGpRedEyeCorrection.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpBrightnessContrast Class
' The BrightnessContrast class enables you to change the brightness and contrast of a bitmap.
' Pass the address of a BrightnessContrast object to the CGpGraphics.DrawImage method or to
' the CGpBitmap.ApplyEffect method. To specify the brightness and contrast levels, pass a
' BrightnessContrastParams structure to the CGpBrightnessContrast.SetParameters method of a
' BrightnessContrast object.
' ########################################################################################

TYPE CGpBrightnessContrast EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS BrightnessContrastParams PTR) AS GpStatus
   DECLARE FUNCTION SetParameters (BYVAL brightnessLevel AS INT_, BYVAL contrastLevel AS INT_) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpBrightnessContrast
   SetStatus(GdipCreateEffect(BrightnessContrastEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
'The CGpBrightnessContrast.SetParameters method sets the parameters of this BrightnessContrast object.
' ========================================================================================
PRIVATE FUNCTION CGpBrightnessContrast.SetParameters (BYVAL params AS BrightnessContrastParams PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(BrightnessContrastParams)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpBrightnessContrast.SetParameters (BYVAL brightnessLevel AS INT_, BYVAL contrastLevel AS INT_) AS GpStatus
   DIM params AS BrightnessContrastParams = TYPE(brightnessLevel, contrastLevel)
   DIM uSize AS UINT = SIZEOF(params)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, @params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpBrightnessContrast.GetParameters method gets the current values of the parameters of
' this BrightnessContrast object.
' ========================================================================================
PRIVATE FUNCTION CGpBrightnessContrast.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpHueSaturationLightness Class
' The HueSaturationLightness class enables you to change the hue, saturation, and lightness
' of a bitmap. Pass the address of a HueSaturationLightness object to the CGpGraphics.DrawImage
' method or to the CGpBitmap.ApplyEffect method. To specify the magnitudes of the changes in
' hue, saturation, and lightness, pass a HueSaturationLightnessParams structure to the
' CGpHueSaturationLightness.SetParameters method of a HueSaturationLightness object.
' ########################################################################################

TYPE CGpHueSaturationLightness EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS HueSaturationLightnessParams PTR) AS GpStatus
   DECLARE FUNCTION SetParameters (BYVAL hueLevel AS INT_, BYVAL saturationLevel AS INT_, BYVAL lightnessLevel AS INT_) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpHueSaturationLightness
   SetStatus(GdipCreateEffect(HueSaturationLightnessEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
'The CGpHueSaturationLightness.SetParameters method sets the parameters of this HueSaturationLightness object.
' ========================================================================================
PRIVATE FUNCTION CGpHueSaturationLightness.SetParameters (BYVAL params AS HueSaturationLightnessParams PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(HueSaturationLightnessParams)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpHueSaturationLightness.SetParameters (BYVAL hueLevel AS INT_, BYVAL saturationLevel AS INT_, BYVAL lightnessLevel AS INT_) AS GpStatus
   DIM params AS HueSaturationLightnessParams = TYPE(hueLevel, saturationLevel, lightnessLevel)
   DIM uSize AS UINT = SIZEOF(params)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, @params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpHueSaturationLightness.GetParameters method gets the current values of the parameters
' of this HueSaturationLightness object.
' ========================================================================================
PRIVATE FUNCTION CGpHueSaturationLightness.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpLevels Class
' The Levels class encompasses three bitmap adjustments: highlight, midtone, and shadow.
' You can apply one or more of those adjustments to a bitmap by passing the address of a
' Levels object to the Graphics::DrawImage method or to the CGpBitmap.ApplyEffect method.
' To specify the intensities of the adjustments, pass the address of a LevelsParams
' structure to the CGpLevels.SetParameters method of a Levels object.
' ########################################################################################

TYPE CGpLevels EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS LevelsParams PTR) AS GpStatus
   DECLARE FUNCTION SetParameters (BYVAL highlight AS INT_, BYVAL midtone AS INT_, BYVAL shadow AS INT_) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpLevels
   SetStatus(GdipCreateEffect(LevelsEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
'The CGpLevels.SetParameters method sets the parameters of this Levels object.
' ========================================================================================
PRIVATE FUNCTION CGpLevels.SetParameters (BYVAL params AS LevelsParams PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(LevelsParams)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpLevels.SetParameters (BYVAL highlight AS INT_, BYVAL midtone AS INT_, BYVAL shadow AS INT_) AS GpStatus
   DIM params AS LevelsParams = TYPE(highlight, midtone, shadow)
   DIM uSize AS UINT = SIZEOF(params)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, @params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpLevels.GetParameters method gets the current values of the parameters of
' this Levels object.
' ========================================================================================
PRIVATE FUNCTION CGpLevels.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpTint Class
' The Tint class enables you to apply a tint to a bitmap. Pass the address of a Tint object
' to the CGpGraphics.DrawImage method or to the CGpBitmap.ApplyEffect method. To specify the nature
' of the tint, pass the address of a TintParams structure to the CGpTint.SetParameters method
' of a Tint object.
' ########################################################################################

TYPE CGpTint EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS TintParams PTR) AS GpStatus
   DECLARE FUNCTION SetParameters (BYVAL hue AS INT_, BYVAL amount AS INT_) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpTint
   SetStatus(GdipCreateEffect(TintEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
'The CGpTint.SetParameters method sets the parameters of this Tint object.
' ========================================================================================
PRIVATE FUNCTION CGpTint.SetParameters (BYVAL params AS TintParams PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(TintParams)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpTint.SetParameters (BYVAL hue AS INT_, BYVAL amount AS INT_) AS GpStatus
   DIM params AS TintParams = TYPE(hue, amount)
   DIM uSize AS UINT = SIZEOF(params)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, @params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpTint.GetParameters method gets the current values of the parameters of
' this Tint object.
' ========================================================================================
PRIVATE FUNCTION CGpTint.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpColorBalance Class
' The ColorBalance class enables you to change the color balance (relative amounts of red,
' green, and blue) of a bitmap. Pass the address of a ColorBalance object to the
' CGpGraphics.DrawImage method or to the CGpBitmap.ApplyEffect method. To specify the nature of
' the change, pass the address of a ColorBalanceParams structure to the CGpColorBalance.SetParameters
' method of a ColorBalance object.
' ########################################################################################

TYPE CGpColorBalance EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS ColorBalanceParams PTR) AS GpStatus
   DECLARE FUNCTION SetParameters (BYVAL cyanRed AS INT_, BYVAL magentaGreen AS INT_, BYVAL yellowBlue AS INT_) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpColorBalance
   SetStatus(GdipCreateEffect(ColorBalanceEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
'The CGpColorBalance.SetParameters method sets the parameters of this ColorBalance object.
' ========================================================================================
PRIVATE FUNCTION CGpColorBalance.SetParameters (BYVAL params AS ColorBalanceParams PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(ColorBalanceParams)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpColorBalance.SetParameters (BYVAL cyanRed AS INT_, BYVAL magentaGreen AS INT_, BYVAL yellowBlue AS INT_) AS GpStatus
   DIM params AS ColorBalanceParams = TYPE(cyanRed, magentaGreen, yellowBlue)
   DIM uSize AS UINT = SIZEOF(params)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, @params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpColorBalance.GetParameters method gets the current values of the parameters of
' this ColorBalance object.
' ========================================================================================
PRIVATE FUNCTION CGpColorBalance.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpColorMatrix Class
' The ColorMatrixEffect class enables you to apply an affine transformation to a bitmap.
' Pass the address of a ColorMatrixEffect object to the CGpGraphics.DrawImage method or to the
' CGpBitmap.ApplyEffect method. To specify the transformation, set the elements of a ColorMatrix
' structure, and pass the address of that structure to the CGpColorMatrixEffect.SetParameters
' method of a ColorMatrixEffect object.
' ########################################################################################

TYPE CGpColorMatrix EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL matrix AS ColorMatrix PTR) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL matrix AS ColorMatrix PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpColorMatrix
   SetStatus(GdipCreateEffect(ColorMatrixEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
'The CGpColorBalance.SetParameters method sets the parameters of this ColorBalance object.
' ========================================================================================
PRIVATE FUNCTION CGpColorMatrix.SetParameters (BYVAL matrix AS ColorMatrix PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(ColorMatrix)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, matrix, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpColorBalance.GetParameters method gets the current values of the parameters of
' this ColorBalance object.
' ========================================================================================
PRIVATE FUNCTION CGpColorMatrix.GetParameters (BYVAL uSize AS UINT PTR, BYVAL matrix AS ColorMatrix PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, matrix))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpColorLUT Class
' A ColorLUTParams structure has four members, each being a lookup table for a particular
' color channel: alpha, red, green, or blue. The lookup tables can be used to make custom
' color adjustments to bitmaps. Each lookup table is an array of 256 bytes that you can
' set to values of your choice. After you have initialized a ColorLUTParams structure, pass
' its address to the CGpColorLUT.SetParameters method of a CGpColorLUT object. Then pass the
' address of that CGpColorLUT object to the CGpGraphics.DrawImage method or to the
' CGpBitmap.ApplyEffect method.
' ########################################################################################

TYPE CGpColorLUT EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS ColorLUTParams PTR) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpColorLUT
   SetStatus(GdipCreateEffect(ColorLUTEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
'The CGpColorLUT.SetParameters method sets the parameters of this ColorLUT object.
' ========================================================================================
PRIVATE FUNCTION CGpColorLUT.SetParameters (BYVAL params AS ColorLUTParams PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(ColorLUTParams)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpColorLUT.GetParameters method gets the current values of the parameters of
' this ColorLUT object.
' ========================================================================================
PRIVATE FUNCTION CGpColorLUT.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================


' ########################################################################################
' CGpColorCurve Class
' The CGpColorCurve class encompasses eight separate adjustments: exposure, density, contrast,
' highlight, shadow, midtone, white saturation, and black saturation. You can apply one of
' those adjustments to a bitmap by passing the address of a CGpColorCurve object to the
' CGpGraphics.DrawImage method or to the CGpBitmap.ApplyEffect method. To specify the
' adjustment, the intensity of the adjustment, and the color channel to which the adjustment
' applies, pass the address of a ColorCurveParams structure to the CGpColorCurve.SetParameters
' method of a ColorCurve object.
' ########################################################################################

TYPE CGpColorCurve EXTENDS CGpEffect

Public:
   DECLARE CONSTRUCTOR
   DECLARE FUNCTION SetParameters (BYVAL params AS ColorCurveParams PTR) AS GpStatus
   DECLARE FUNCTION SetParameters (BYVAL adjustment AS CurveAdjustments, BYVAL channel AS CurveChannel, BYVAL adjustValue AS INT_) AS GpStatus
   DECLARE FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus

END TYPE

' ========================================================================================
' * Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CGpColorCurve
   SetStatus(GdipCreateEffect(ColorCurveEffectGuid, @m_pEffect))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
'The CGpColorLUT.SetParameters method sets the parameters of this ColorLUT object.
' ========================================================================================
PRIVATE FUNCTION CGpColorCurve.SetParameters (BYVAL params AS ColorCurveParams PTR) AS GpStatus
   DIM uSize AS UINT = SIZEOF(ColorCurveParams)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, params, uSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CGpColorCurve.SetParameters (BYVAL adjustment AS CurveAdjustments, BYVAL channel AS CurveChannel, BYVAL adjustValue AS INT_) AS GpStatus
   DIM params AS ColorCurveParams = TYPE(adjustment, channel, adjustValue)
   DIM uSize AS UINT = SIZEOF(params)
   RETURN SetStatus(GdipSetEffectParameters(m_pEffect, @params, uSize))
END FUNCTION
' ========================================================================================

' ========================================================================================
' The CGpColorLUT.GetParameters method gets the current values of the parameters of
' this ColorLUT object.
' ========================================================================================
PRIVATE FUNCTION CGpColorCurve.GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
   RETURN SetStatus(GdipGetEffectParameters(m_pEffect, uSize, params))
END FUNCTION
' ========================================================================================
