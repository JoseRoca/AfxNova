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

## <a name="constructoreffecty"></a>Constructor (CGpEffect)

Creates an instance of the `CgpEffect`class.

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
