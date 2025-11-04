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
| [Constructor](#constructorbrightnessffect) | Creates an instance of the **CGpBrightnessContrast** class. |
| [GetParameters](#getparametersbrightness) | Gets the current values of the parameters of this **BrightnessContrast** object. |
| [SetParameters](#setparametersbrightness) | Sets the parameters of this **BrightnessContrast** object. |

## <a name="constructorbrightnessffect"></a>Constructor (CGpBrightnessContrast)

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
