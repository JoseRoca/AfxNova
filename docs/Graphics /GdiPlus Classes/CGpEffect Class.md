# CGpEffect Class

The `GpEffect` class serves as a base class for eleven classes that you can use to apply effects and adjustments to bitmaps.

To apply and effect to a bitmap, create an instance of one of the descendants of the `CGpEffect` class, and pass the address of that descendant to the **DrawImage** method of the CGGraphics class or to the **ApplyEffect** method of the **CGpBitmap** class.

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

The following classes descend from **CGpEffect**.

| Name       | Description |
| ---------- | ----------- |
| [CGpBlur](#cgpblur) | Enables to apply a Gaussian blur effect to a bitmap and specify the nature of the blur. |
| [CGpBrightnessContrast](#cgpbrightnesscontrast) | Enables to change the brightness and contrast of a bitmap. |
| [CGpColorBalance](#cgpcolorbalance) | Enables to o change the color balance (relative amounts of red, green, and blue) of a bitmap. |
| [CGpColorCurve](#cgpcolorcurve) | Encompasses eight separate adjustments: exposure, density, contrast, highlight, shadow, midtone, white saturation, and black saturation. |
| [CGpColorLUT](#cgpcolorlut) | Allows to make custom color adjustments to bitmaps. |
| [CGpColorMatrixEffect](#cgpcolormatrixeffect) | Enables to apply an affine transformation to a bitmap. |
| [CGpHueSaturationLightness](#cgphuesaturationlightness) | Enables to change the hue, saturation, and lightness of a bitmap. |
| [CGpLevels](#cgplevels) | Encompasses three bitmap adjustments: highlight, midtone, and shadow. |
| [CGpRedEyeCorrection](#cgpredeyecorrection) | Enables to correct the red eyes that sometimes occur in flash photographs. |
| [CGpSharpen](#cgpsharpen) | Enables to adjust the sharpness of a bitmap. |
| [CGpTint](#cgptint) | Enables to apply a tint to a bitmap. |

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

# <a name="cgpblur"></a>CGpBlur Class

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

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

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

# <a name="cgpbrightnesscontrast"></a>CGpBrightnessContrast Class

The `CGpBrightnessContrast` class enables you to change the brightness and contrast of a bitmap. Pass the address of a **BrightnessContrast** object to the **DrawImage** method of the **Graphics** or to the **ApplyEffect** method of the **Bitmap** object. To specify the brightness and contrast levels, pass a **BrightnessContrastParams** structure to the **SetParameters** method of a **BrightnessContrast** object.

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

## <a name="cgpcolorbalance"></a>CGpColorBalance Class

The `CGpColorBalance` class enables you to change the color balance (relative amounts of red, green, and blue) of a bitmap. Pass the address of a **ColorBalance** object to the **DrawImage** method of the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object. To specify the nature of the change, pass the address of a **ColorBalanceParams** structure to the **SetParameters** method of a **ColorBalance** object.

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

# <a name="cgpcolorcurve"></a>CGpColorCurve Class

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

# <a name="cgpcolorlut"></a>GpColorLUT Class

A **ColorLUTParams** structure has four members, each being a lookup table for a particular color channel: alpha, red, green, or blue. The lookup tables can be used to make custom color adjustments to bitmaps. Each lookup table is an array of 256 bytes that you can set to values of your choice. After you have initialized a **ColorLUTParams** structure, pass its address to the **SetParameters** method of a **ColorLUT** object. Then pass the address of that **ColorLUT** object to the **DrawImage** method of the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object.

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

# <a name="cgpcolormatrixeffect"></a>CGpColorMatrixEffect Class

The `CGpColorMatrixEffect` class enables you to apply an affine transformation to a bitmap. Pass the address of a **ColorMatrixEffect** object to the **DrawImage** method of the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object. To specify the transformation, set the elements of a **ColorMatrix** structure, and pass the address of that structure to the **SetParameters** method of a **ColorMatrixEffect** object.

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

# <a name="cgphuesaturationlightness"></a>CGpHueSaturationLightness Class

The `CGpHueSaturationLightness` class enables you to change the hue, saturation, and lightness of a bitmap. Pass the address of a **HueSaturationLightness** object to the **DrawImage** method of the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object. To specify the magnitudes of the changes in hue, saturation, and lightness, pass a **HueSaturationLightnessParams** structure to the **SetParameters** method of a **HueSaturationLightness** object.

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
FUNCTION SetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS HueSaturationLightnessParams PTR) AS GpStatus
FUNCTION SetParametes (BYVAL hueLevel AS INT_, BYVAL saturationLevel AS INT_, BYVAL lightnessLevel AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **HueSaturationLightnessParams** structure. |
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

# <a name="cgplevels"></a>CGpLevels Class

The `CGpLevels` class encompasses three bitmap adjustments: highlight, midtone, and shadow. You can apply one or more of those adjustments to a bitmap by passing the address of a Levels object to the **DrawImage** method of the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object. To specify the intensities of the adjustments, pass the address of a **LevelsParams** structure to the **SetParameters** method of a **Levels** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorlevelseffect) | Creates an instance of the **CGpLevels** class. |
| [GetParameters](#getparameterslevels) | Gets the current values of the parameters of this **Levels** object. |
| [SetParameters](#setparameterslevels) | Sets the parameters of this **Levels** object. |

---

## <a name="constructorlevelseffect"></a>Constructor (CGpLevels)

Creates an instance of the `CGpLevels`class.

```
CONSTRUCTOR
```
---

## <a name="getparameterslevels"></a>GetParameters (CGpLevels)

Gets the current values of the parameters of this **Leveks** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **LevelsParams** structure. |
| *params* | Pointer to a **LevelsParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparameterslevels"></a>SetParameters (CGpLevels)

Sets the current values of the parameters of this **Levels** object.

```
FUNCTION SetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS LevelsParams PTR) AS GpStatus
FUNCTION SetParameters (BYVAL highlight AS INT_, BYVAL midtone AS INT_, BYVAL shadow AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **LevelsParams** structure. |
| *params* | Pointer to a **LevelsParams** structure that specifies the parameters. |


| Parameter  | Description |
| ---------- | ----------- |
| *highlight* | Integer in the range 0 through 100 that specifies which pixels should be lightened. You can use this adjustment to lighten pixels that are already lighter than a certain threshold. Setting highlight to 100 specifies no change. Setting highlight to t specifies that a color channel value is increased if it is already greater than t percent of full intensity. For example, setting highlight to 90 specifies that all color channel values greater than 90 percent of full intensity are increased. |
| *midtone* | Integer in the range -100 through 100 that specifies how much to lighten or darken an image. Color channel values in the middle of the intensity range are altered more than color channel values near the minimum or maximum intensity. You can use this adjustment to lighten (or darken) an image without loosing the contrast between the darkest and lightest portions of the image. A value of 0 specifies no change. Positive values specify that the midtones are made lighter, and negative values specify that the midtones are made darker. |
| *shadow* | Integer in the range 0 through 100 that specifies which pixels should be darkened. You can use this adjustment to darken pixels that are already darker than a certain threshold. Setting shadow to 0 specifies no change. Setting shadow to t specifies that a color channel value is decreased if it is already less than t percent of full intensity. For example, setting shadow to 10 specifies that all color channel values less than 10 percent of full intensity are decreased. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

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
---

# <a name="cgpredeyecorrection"></a>CGpRedEyeCorrection Class

The `CGpRedEyeCorrection` class enables you to correct the red eyes that sometimes occur in flash photographs. Pass the address of a **RedEyeCorrection** object to the **DrawImage** method of the **Graphics** object or to the **ApplyEffect** method of the **Bitmap** object. To specify areas of the bitmap that have red eyes, pass a **RedEyeCorrectionParams** structure to the **SetParameters** method of a **RedEyeCorrection** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorlevelseffect) | Creates an instance of the **CGpRedEyeCorrection** class. |
| [GetParameters](#getparameterslevels) | Gets the current values of the parameters of this **RedEyeCorrection** object. |
| [SetParameters](#setparameterslevels) | Sets the parameters of this **RedEyeCorrection** object. |

---

## <a name="constructorlevelseffect"></a>Constructor (CGpRedEyeCorrection)

Creates an instance of the `CGpRedEyeCorrection`class.

```
CONSTRUCTOR
```
---

## <a name="getparameterslevels"></a>GetParameters (CGpRedEyeCorrection)

Gets the current values of the parameters of this **RedEyeCorrection** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS ANY PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **RedEyeCorrectionParams** structure. |
| *params* | Pointer to a **RedEyeCorrectionParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparameterslevels"></a>SetParameters (CGpRedEyeCorrection)

Sets the current values of the parameters of this **RedEyeCorrection** object.

```
FUNCTION SetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS LevelsParams PTR) AS GpStatus
FUNCTION SetParameters (BYVAL numberOfAreas AS UINT, BYVAL areas AS RECT PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **RedEyeCorrectionParams** structure. |
| *params* | Pointer to a **RedEyeCorrectionParams** structure that specifies the parameters. |


| Parameter  | Description |
| ---------- | ----------- |
| *numberOfAreas* | Integer that specifies the number of **RECT** structures in the **areas** array. |
| *areas* | Pointer to an array of **RECT** structures, each of which specifies an area of the bitmap to which red eye correction should be applied. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

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
---

# <a name="cgpsharpen"></a>CGpSharpen Class

The `CGpSharpen` class enables you to adjust the sharpness of a bitmap. Pass the address of a **Sharpen** object to the **DrawImage** method of the **Graphics** oject or to the **ApplyEffect** method of the **Bitmap**. To specify the nature of the sharpening adjustment, pass the address of a **SharpenParams** structure to the **SetParameters** method of a **Sharpen** object.

**Inherits from**: CGpEffect.<br>
**Include file**: CGpEffect.inc.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructorlevelseffect) | Creates an instance of the **CGpSharpen** class. |
| [GetParameters](#getparameterslevels) | Gets the current values of the parameters of this **Sharpen** object. |
| [SetParameters](#setparameterslevels) | Sets the parameters of this **Sharpen** object. |

---

## <a name="constructorlevelseffect"></a>Constructor (CGpSharpen)

Creates an instance of the `CGpSharpen`class.

```
CONSTRUCTOR
```
---

## <a name="getparameterslevels"></a>GetParameters (CGpSharpen)

Gets the current values of the parameters of this **Sharpen** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS SharpenParams PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **SharpenParams** structure. |
| *params* | Pointer to a **SharpenParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparameterslevels"></a>SetParameters (CGpSharpen)

Sets the current values of the parameters of this **Sharpen** object.

```
FUNCTION SetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS SharpenParams PTR) AS GpStatus
FUNCTION SetParameters (BYVAL radius AS REAL, BYVAL amount AS REAL) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **SharpenParams** structure. |
| *params* | Pointer to a **SharpenParams** structure that specifies the parameters. |


| Parameter  | Description |
| ---------- | ----------- |
| *radius* | Real number that specifies the sharpening radius (the radius of the convolution kernel) in pixels. The radius must be in the range 0 through 255. As the radius increases, more surrounding pixels are involved in calculating the new value of a given pixel. |
| *amount* | Real number in the range 0 through 100 that specifies the amount of sharpening to be applied. A value of 0 specifies no sharpening. As the value of amount increases, the sharpness increases. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

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
---

# <a name="cgptint"></a>CGpTint Class

The `CGpTint` class enables you to apply a tint to a bitmap. Pass the address of a **Tint** object to the **DrawImage** method of the **Graphics** objecxt or to the **ApplyEffect** method of the **Bitmap** object. To specify the nature of the tint, pass the address of a **TintParams** structure to the **SetParameters** method of a **Tint** object.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructortintffect) | Creates an instance of the **CGpTint** class. |
| [GetParameters](#getparameterstint) | Gets the current values of the parameters of this **Tint** object. |
| [SetParameters](#setparameterstint) | Sets the parameters of this **Tint** object. |

---

## <a name="constructortintffect"></a>Constructor (CGpTint)

Creates an instance of the `CGpTint`class.

```
CONSTRUCTOR
```
---

## <a name="getparameterstint"></a>GetParameters (CGpTint)

Gets the current values of the parameters of this **Tint** object.

```
FUNCTION GetParameters (BYVAL uSize AS UINT PTR, BYVAL params AS TintParams PTR) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **TintParams** structure. |
| *params* | Pointer to a **TintParams** structure that receives the parameter values. |


#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

---

## <a name="setparameterstint"></a>SetParameters (CGpTint)

Sets the current values of the parameters of this **Tint** object.

```
FUNCTION SetParameters (BYVAL params AS TintParams PTR) AS GpStatus
DECLARE FUNCTION SetParameters (BYVAL hue AS INT_, BYVAL amount AS INT_) AS GpStatus
```

| Parameter  | Description |
| ---------- | ----------- |
| *uSize* | Pointer to a UINT that specifies the size, in bytes, of a **TintParams** structure. |
| *params* | Pointer to a **TintParams** structure that specifies the parameters. |


| Parameter  | Description |
| ---------- | ----------- |
| *hue* | Integer in the range -180 through 180 that specifies the hue to be strengthened or weakened. A value of 0 specifies blue. A positive value specifies a clockwise angle on the color wheel. For example, positive 60 specifies cyan and positive 120 specifies green. A negative value specifies a counter-clockwise angle on the color wheel. For example, negative 60 specifies magenta and negative 120 specifies red. |
| *amount* | Integer in the range -100 through 100 that specifies how much the hue (given by the hue parameter) is strengthened or weakened. A value of 0 specifies no change. Positive values specify that the hue is strengthened and negative values specify that the hue is weakened. |

#### Return value

If the method succeeds, it returns **StatusOk**, which is an element of the **GpStatus** enumeration.

If the method fails, it returns one of the other elements of the **GpStatus** enumeration.

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
---
