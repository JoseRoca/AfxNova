# CGpEffect Class

The `GpEffect` class serves as a base class for eleven classes that you can use to apply effects and adjustments to bitmaps. The following classes descend from Effect.

* CGpBlur
* CGpSharpen
* CGpTint
* CGpRedEyeCorrection
* CGpColorMatrixEffect
* CGpColorLUT
* CGpBrightnessContrast
* CGpHueSaturationLightness
* CGpColorBalance
* CGpLevels
* CGpColorCurve

To apply and effect to a bitmap, create an instance of one of the descendants of the `CGpEffect` class, and pass the address of that descendant to the **DrawImage** method of the CGGraphics class or to the **ApplyEffect** method of the **CGpBitmal** class.

**Inherits from**: CGpBase.<br>
**Include file**: CGpEffect.inc.

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructoreffect) | Creates an instance of the **CGpEffect** class. |
| [GetAuxData](#getauxdata) | Gets a pointer to a set of lookup tables created by a previous call to the **ApplyEffect** method of the `CGpBitmap` class. |
| [GetAuxDataSize](#getauxdatasize) | Gets the size, in bytes, of the auxiliary data created by a previous call to the **ApplyEffect** method of the `CGpBitmap` class. |
| [GetParameterSize](#getparametersize) | Gets the total size, in bytes, of the parameters currently set for this Effect. The **GetParameterSize** method is usually called on an object that is an instance of a descendant of the `CGpEffect` class. |
| [UseAuxDara](#useauxdata) | Sets or clears a flag that specifies whether the **ApplyEffect** method of the `CGpBitmap`class should return a pointer to the auxiliary data that it creates. |

---

## <a name="constructoreffect"></a>Constructor (CGpEffect)

Creates an instance of the `CGpEffect`class.

```
CONSTRUCTOR
```
---

## GetAuxData

```
FUNCTION GetAuxData () AS ANY PTR
```

Gets a pointer to a set of lookup tables created by a previous call to the **ApplyEffect** method of the `CGpBitmap` class.

#### Return value

This method returns a pointer to a set of lookup tables created by a previous call to **ApplyEffect** method of the `CGpBitmap` class. If no lookup tables are available, the return value is NULL.

#### Remarks
You can apply an effect to a bitmap by creating an instance of one of the descendants of the `CGpEffect` class and passing the address of that descendant to the **ApplyEffect** method of the `CGpBitmap` class. For certain descendants of `CGpEffect`, **ApplyEffect** creates lookup tables and returns the address of those tables to the descendant object. For example, you can retrieve the lookup tables for a **BrightnessContrast** object as follows:

* Create a **BrightnessContrast** object and call its **SetParameters** method.
* Pass TRUE to the **UseAuxData** method of the **BrightnessContrast** object.
* Pass the address of the **BrightnessContrast** object to the **ApplyEffect** method.
* Call the **GetAuxData** method of the **BrightnessContrast** object to obtain a pointer to the lookup tables created by **ApplyEffect**. The buffer for the lookup tables is allocated by **ApplyEffect**; you are not responsible for freeing the buffer.
**ApplyEffect** can return the address of lookup tables for the following descendants of the `CGpEffect` class.

* BrightnessContrast
* HueSaturationLightness
* Levels
* ColorBalance
* ColorCurve

For the classes in the preceding list, **ApplyEffect** creates four lookup tables: one each for the blue, green, red, and alpha channels. Each lookup table is an array of 256 bytes so the size of the entire set of tables is 1024 bytes. The tables are stored in the order blue, green, red, alpha.

---

## GetAuxDataSize

Gets the size, in bytes, of the auxiliary data created by a previous call to the **ApplyEffect** method of the `CGpBitmap` class.

```
FUNCTION GetAuxDataSize () AS INT_
```

#### Return value

This method returns the size, in bytes, of the auxiliary data.

---

#### GetParameterSize

Gets the total size, in bytes, of the parameters currently set for the effect. The **GetParameterSize** method is usually called on an object that is an instance of a descendant of the `CGpEffect` class.

```
FUNCTION GetParameterSize () AS UINT
```

#### Return value

The total size, in bytes, of the parameters currently set for the effect.

---

## UseAuxData

Sets or clears a flag that specifies whether the **ApplyEffect** method of the `CGpBitmap`class should return a pointer to the auxiliary data that it creates.

```
SUB UseAuxData (BYVAL useAuxDataFlag AS BOOL)
```

| Parameter  | Description |
| ---------- | ----------- |
| *useAuxDataFlag* | Set to TRUE to specify that **ApplyEffect** should return a pointer to its auxiliary data; FALSE otherwise. |

---

# CGpBlur Class

The `CGpBlur` class enables you to apply a Gaussian blur effect to a bitmap and specify the nature of the blur. Pass the address of a **Blur** object to the **DrawImage** method of the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object. To specify the nature of the blur, pass a **BlurParams** structure to the **SetParameters** method of a **Blur** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorblureffect) | Creates an instance of the **CGpBlur** class. |
| [GetParameters](#getparametersblur) | Gets the current values of the parameters of this **Blur** object. |
| [SetParameters](#setparametersblur) | Sets the parameters of this **Blur** object. |

---

## <a name="constructorblureffect"></a>Constructor (CGpBlur)

Creates an instance of the `CGpBlur`class.

```
CONSTRUCTOR
```
---

## <a name="getparametersblur"></a>GetParameters (CGpBlur)

Gets the current values of the parameters of this **Blur** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **BlurParams** structure. |
| *params* | Pointer to a **BlurParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus* enumeration.

---

## <a name="setparametersblur"></a>SetParameters (CGpBlur)

Sets the current values of the parameters of this **Blur** object.

```
FUNCTION SetParameters (BYVAL params AS BlurParams PTR) AS GpStatus
FUNCTION SetParameters (BYVAL radius AS REAL, BYVAL expandEdge AS BOOL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *params* | Pointer to a **BlurParams** structure that specifies the parameters. |

| Parameter  | Description |
| ---------- | ----------- |
| *radius* | Real number that specifies the blur radius (the radius of the Gaussian convolution kernel) in pixels. The radius must be in the range 0 through 255. As the radius increases, the resulting bitmap becomes more blurry. |
| *expandEdge* | Boolean value that specifies whether the bitmap expands by an amount equal to the blur radius. If TRUE, the bitmap expands by an amount equal to the radius so that it can have soft edges. If FALSE, the bitmap remains the same size and the soft edges are clipped. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Remarks

One of the two **ApplyEffect** methods of the `CGpBitmap`class blurs a bitmap in place. That particular **ApplyEffect** method ignores the *expandEdge* parameter.

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
---

# CGpBrightnessContrast Class

The `CGpBrightnessContrast` class enables you to change the brightness and contrast of a bitmap. Pass the address of a **BrightnessContrast** object to the **DrawImage** methodof the **Graphics** or to the **ApplyEffect** methodof the **Bitmap** object. To specify the brightness and contrast levels, pass a **BrightnessContrastParams** structure to the **SetParameters** method of a **BrightnessContrast** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorbrightnesseffect) | Creates an instance of the **CGpBrightnessContrast** class. |
| [GetParameters](#getparametersbrightness) | Gets the current values of the parameters of this **BrightnessContrast** object. |
| [SetParameters](#setparametersbrightness) | Sets the parameters of this **BrightnessContrast** object. |

---

## <a name="constructorbrightnesseffect"></a>Constructor (CGpBrightnessContrast)

Creates an instance of the `CGpBrightnessContrast`class.

```
CONSTRUCTOR
```
---

## <a name="getparametersbrightness"></a>GetParameters (CGpBrightnessContrast)

Gets the current values of the parameters of this **BrightnessContrast** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **BrightnessContrastParams** structure. |
| *params* | Pointer to a **BrightnessContrastParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparametersbrightness"></a>SetParameters (CGpBrightnessContrast)

Sets the current values of the parameters of this **BrightnessContrast** object.

```
FUNCTION SetParameters (BYVAL params AS BrightnessContrastParams PTR) AS GpStatus
FUNCTION SetParameters (BYVAL brightnessLevel AS INT_, BYVAL contrastLevel AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *params* | Pointer to a **BrightnessContrastParams** structure that specifies the parameters. |

| Parameter  | Description |
| ---------- | ----------- |
| *brightnessLevel* | Integer in the range -255 through 255 that specifies the brightness level. If the value is 0, the brightness remains the same. As the value moves from 0 to 255, the brightness of the image increases. As the value moves from 0 to -255, the brightness of the image decreases. |
| *contrastLevel* | Integer in the range -100 through 100 that specifies the contrast level. If the value is 0, the contrast remains the same. As the value moves from 0 to 100, the contrast of the image increases. As the value moves from 0 to -100, the contrast of the image decreases. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

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
---

# CGpColorBalance Class

The `CGpColorBalance` class enables you to change the color balance (relative amounts of red, green, and blue) of a bitmap. Pass the address of a **ColorBalance** object to the **DrawImage** methodof the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object. To specify the nature of the change, pass the address of a **ColorBalanceParams** structure to the **SetParameters** method of a **ColorBalance** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorcolorbalanceeffect) | Creates an instance of the **CGpColorBalance** class. |
| [GetParameters](#getparameterscolorbalance) | Gets the current values of the parameters of this **ColorBalance** object. |
| [SetParameters](#setparameterscolorbalance) | Sets the parameters of this **ColorBalance** object. |

---

## <a name="constructorcolorbalanceeffect"></a>Constructor (CGpColorBalance)

Creates an instance of the `CGpColorBalance`class.

```
CONSTRUCTOR
```
---

## <a name="getparameterscolorbalance"></a>GetParameters (CGpColorBalance)

Gets the current values of the parameters of this **ColorBalance** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **ColorBalanceParams** structure. |
| *params* | Pointer to a **ColorBalanceParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparameterscolorbalance"></a>SetParameters (CGpColorBalance)

Sets the current values of the parameters of this **ColorBalance** object.

```
FUNCTION SetParameters (BYVAL params AS ColorBalanceParams PTR) AS GpStatus
FUNCTION SetParameters (SetParameters (BYVAL cyanRed AS INT_, BYVAL magentaGreen AS INT_, _
   BYVAL yellowBlue AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *params* | Pointer to a **ColorBalanceParams** structure that specifies the parameters. |

| Parameter  | Description |
| ---------- | ----------- |
| *cyanRed* | Integer in the range -100 through 100 that specifies a change in the amount of red in the image. If the value is 0, there is no change. As the value moves from 0 to 100, the amount of red in the image increases and the amount of cyan decreases. As the value moves from 0 to -100, the amount of red in the image decreases and the amount of cyan increases. |
| *magentaGreen* | Integer in the range -100 through 100 that specifies a change in the amount of green in the image. If the value is 0, there is no change. As the value moves from 0 to 100, the amount of green in the image increases and the amount of magenta decreases. As the value moves from 0 to -100, the amount of green in the image decreases and the amount of magenta increases. |
| *yellowBlue* | Integer in the range -100 through 100 that specifies a change in the amount of blue in the image. If the value is 0, there is no change. As the value moves from 0 to 100, the amount of blue in the image increases and the amount of yellow decreases. As the value moves from 0 to -100, the amount of blue in the image decreases and the amount of yellow increases. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

# CGpColorCurve Class

The `CGpColorCurve` class encompasses eight separate adjustments: exposure, density, contrast, highlight, shadow, midtone, white saturation, and black saturation. You can apply one of those adjustments to a bitmap by passing the address of a **ColorCurve** object to the **DrawImage** method of the **Graphics** or to the **ApplyEffect** method of the **Bitmap** object. To specify the adjustment, the intensity of the adjustment, and the color channel to which the adjustment applies, pass the address of a **ColorCurveParams** structure to the **SetParameters** method of a **ColorCurve** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorcolorcurveeffect) | Creates an instance of the **CGpColorCurve** class. |
| [GetParameters](#getparameterscolorcurve) | Gets the current values of the parameters of this **ColorCurve** object. |
| [SetParameters](#setparameterscolorcurve) | Sets the parameters of this **ColorCurve** object. |

---

## <a name="constructorcolorcurveeffect"></a>Constructor (CGpColorCurve)

Creates an instance of the `CGpColorCurve`class.

```
CONSTRUCTOR
```
---

## <a name="getparameterscolorcurve"></a>GetParameters (CGpColorCurve)

Gets the current values of the parameters of this **ColorBalance** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **ColorCurveParams** structure. |
| *params* | Pointer to a **ColorCurveParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparameterscolorcurve"></a>SetParameters (CGpColorCurve)

Sets the current values of the parameters of this **ColorCurve** object.

```
FUNCTION SetParameters (BYVAL params AS ColorBalanceParams PTR) AS GpStatus
FUNCTION SetParameters (BYVAL adjustment AS CurveAdjustments, BYVAL channel AS CurveChannel, _
   BYVAL adjustValue AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *params* | Pointer to a **ColorCurveParams** structure that specifies the parameters. |

| Parameter  | Description |
| ---------- | ----------- |
| *adjustment* | Element of the **CurveAdjustments** enumeration that specifies the adjustment to be applied. |
| *channel* | Element of the **CurveChannel** enumeration that specifies the color channel to which the adjustment applies. |
| *adjustValue* | Integer that specifies the intensity of the adjustment. The range of acceptable values depends on which adjustment is being applied. To see the range of acceptable values for a particular adjustment, see the **CurveAdjustments** enumeration. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

# CGpColorLUT Class

A **ColorLUTParams** structure has four members, each being a lookup table for a particular color channel: alpha, red, green, or blue. The lookup tables can be used to make custom color adjustments to bitmaps. Each lookup table is an array of 256 bytes that you can set to values of your choice. After you have initialized a **ColorLUTParams** structure, pass its address to the **SetParameters** method of a **ColorLUT** object. Then pass the address of that **ColorLUT** object to the **DrawImage** methodof the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorcolorluteffect) | Creates an instance of the **CGpColorLUT** class. |
| [GetParameters](#getparameterscolorlut) | Gets the current values of the parameters of this **ColorLUT** object. |
| [SetParameters](#setparameterscolorlut) | Sets the parameters of this **ColorLUT** object. |

---

## <a name="constructorcolorluteffect"></a>Constructor (CGpColorLUT)

Creates an instance of the `CGpColorLUT`class.

```
CONSTRUCTOR
```
---

## <a name="getparameterscolorlut"></a>GetParameters (CGpColorLUT)

Gets the current values of the parameters of this **ColorLUT** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **ColorLUTParams** structure. |
| *params* | Pointer to a **ColorLUTParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparameterscolorlut"></a>SetParameters (CGpColorLUT)

Sets the current values of the parameters of this **ColorLUT** object.

```
FUNCTION SetParameters (BYVAL params AS ColorLUTParams PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *params* | Pointer to a **ColorLUTParams** structure that specifies the parameters. |

**ColorLUTParams** strucruee
```
type ColorLUTParams
	lutB(0 TO 255) AS BYTE
	lutG(0 TO 255) AS BYTE
	lutR(0 TO 255) AS BYTE
	lutA(0 TO 255) AS BYTE
end type
```

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

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
---

# CGpColorMatrixEffect Class

The `CGpColorMatrixEffect` class enables you to apply an affine transformation to a bitmap. Pass the address of a **ColorMatrixEffect** object to the **DrawImage** methodof the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object. To specify the transformation, set the elements of a **ColorMatrix** structure, and pass the address of that structure to the **SetParameters** method of a **ColorMatrixEffect** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorcolormatrixeffect) | Creates an instance of the **CGpColorMatrixEffect** class. |
| [GetParameters](#getparameterscolormatrix) | Gets the current values of the parameters of this **ColorMatrixEffect** object. |
| [SetParameters](#setparameterscolormatrix) | Sets the parameters of this **ColorMatrixEffect** object. |

---

## <a name="constructorcolormatrixeffect"></a>Constructor (CGpColorMatrixEffect)

Creates an instance of the `CGpColorMatrixEffect`class.

```
CONSTRUCTOR
```
---

## <a name="getparameterscolormatrix"></a>GetParameters (CGpColorMatrixEffect)

Gets the current values of the parameters of this **ColorMatrixEffect** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL matrix AS ColorMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **ColorMatrix** structure. |
| *matrix* | Pointer to a **ColorMatrix** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparameterscolormatrix"></a>SetParameters (CGpColorMatrixEffect)

Sets the current values of the parameters of this **ColorMatrixEffect** object.

```
FUNCTION SetParameters (BYVAL matrix AS ColorMatrix PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *matrix* | Pointer to a **ColorMatrix** structure that specifies the parameters. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

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
   DIM grayEffect AS CGpColorMatrixEffect
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

# CGpHueSaturationLightness Class

The `CGpHueSaturationLightness` class enables you to change the hue, saturation, and lightness of a bitmap. Pass the address of a **HueSaturationLightness** object to the **DrawImage** methodof the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object. To specify the magnitudes of the changes in hue, saturation, and lightness, pass a **HueSaturationLightnessParams** structure to the **SetParameters** method of a **HueSaturationLightness** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorhueeffect) | Creates an instance of the **CGpHueSaturationLightness** class. |
| [GetParameters](#getparametershue) | Gets the current values of the parameters of this **HueSaturationLightness** object. |
| [SetParameters](#setparametershue) | Sets the parameters of this **HueSaturationLightness** object. |

---

## <a name="constructorhueeffect"></a>Constructor (CGpHueSaturationLightness)

Creates an instance of the `CGpHueSaturationLightness`class.

```
CONSTRUCTOR
```
---

## <a name="getparametershue"></a>GetParameters (CGpHueSaturationLightness)

Gets the current values of the parameters of this **HueSaturationLightness** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS HueSaturationLightnessParams PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **HueSaturationLightnessParams** structure. |
| *params* | Pointer to a **HueSaturationLightnessParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparametershue"></a>SetParameters (CGpHueSaturationLightness)

Sets the current values of the parameters of this **HueSaturationLightness** object.

```
FUNCTION SetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
FUNCTION SetParametes (BYVAL hueLevel AS INT_, BYVAL saturationLevel AS INT_, _
   BYVAL lightnessLevel AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a **HueSaturationLightnessParams** structure that specifies the parameters. |
| *params* | Pointer to a **HueSaturationLightnessParams** structure that specifies the parameters. |


| Parameter  | Description |
| ---------- | ----------- |
| *hueLevel* | Integer in the range -180 through 180 that specifies the change in hue. A value of 0 specifies no change. Positive values specify counterclockwise rotation on the color wheel. Negative values specify clockwise rotation on the color wheel. |
| *saturationLevel* | Integer in the range -100 through 100 that specifies the change in saturation. A value of 0 specifies no change. Positive values specify increased saturation and negative values specify decreased saturation. |
| *lightnessLevel* | Integer in the range -100 through 100 that specifies the change in lightness. A value of 0 specifies no change. Positive values specify increased lightness and negative values specify decreased lightness. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

#### Example

```
' ========================================================================================
' The HueSaturationLightness effecty enables you to change the hue, saturation, and lightness
' of a bitmap. To specify the magnitudes of the changes in hue, saturation, and lightness,
' pass a HueSaturationLightnessParams structure to the SettParameters function.
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
---
