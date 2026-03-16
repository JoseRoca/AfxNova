# CUpDown control

An up-down control is a pair of arrow buttons that the user can click to increment or decrement a value, such as a scroll position or a number displayed in a companion control (called a buddy window).

See: [About Up-Down Controls](https://learn.microsoft.com/en-us/windows/win32/controls/up-down-controls)

**Include file**: CUpDown.inc


| Name       | Description |
| ---------- | ----------- |
| [GetAccel](#getaccel) | Retrieves acceleration information for an up-down control. |
| [GetBase](#getbase) | Retrieves the current radix base (that is, either base 10 or 16) for an up-down control. |
| [GetBuddy](#getbuddy) | Retrieves the handle to the current buddy window. |
| [GetPos](#getpos) | Retrieves the current position of an up-down control with 16-bit precision. |
| [GetPos32](#getpos32) | Returns the 32-bit position of an up-down control. |
| [GetRange](#getrange) | Retrieves the minimum and maximum positions (range) for an up-down control. |
| [GetRange32](#getrange32) | Retrieves the 32-bit range of an up-down control. |
| [SetAccel](#setaccel) | Sets the acceleration for an up-down control. |
| [SetBase](#setbase) | Sets the radix base for an up-down control. The base value determines whether the buddy window displays numbers in decimal or hexadecimal digits. Hexadecimal numbers are always unsigned, and decimal numbers are signed. |
| [SetAccel](#setaccel) | Sets the acceleration for an up-down control. |
| [SetBuddy](#setbuddy) | Sets the buddy window for an up-down control. |
| [SetPos](#setpos) | Sets the current position for an up-down control with 16-bit precision. |
| [SetPos32](#setpos32) | Sets the position of an up-down control with 32-bit precision. |
| [SetRange](#setRange) | Sets the minimum and maximum positions (range) for an up-down control. |
| [SetRange32](#setRange32) | Sets the 32-bit range of an up-down control. |

---

### GetAccel

Retrieves acceleration information for an up-down control.

```
FUNCTION GetAccel (BYVAL hUpDown AS HWND, BYVAL cAccels AS DWORD, BYVAL pAccels AS UDACCEL PTR) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *cAccels* | Number of elements in the array specified by *paccels*. |
| *pAccels* | Pointer to an array of **UDACCEL** structures that receive acceleration information. |

#### Return value

The return value is the number of accelerators currently set for the control.

---

### GetBase

Retrieves the current radix base (that is, either base 10 or 16) for an up-down control.

```
FUNCTION GetBase (BYVAL hUpDown AS HWND) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |

#### Return value

The return value is the current base value.

---

### GetBuddy

Retrieves the handle to the current buddy window.

```
FUNCTION GetBuddy (BYVAL hUpDown AS HWND) AS HWND
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |

#### Return value

The return value is the handle to the current buddy window.

---

### GetPos

Retrieves the current position of an up-down control with 16-bit precision.

```
FUNCTION GetPos (BYVAL hUpDown AS HWND) AS HWND
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |

#### Return value

If successful, the **HIWORD** is set to zero and the **LOWORD** is set to the control's current position. If an error occurs, the **HIWORD** is set to a nonzero value.

#### Remarks

When processing this message, the up-down control updates its current position based on the caption of the buddy window. The up-down control returns an error if there is no buddy window or if the caption specifies an invalid or out-of-range value. Also, an error flag is set in the **HIWORD** of the return if the control is created without the **UDS_SETBUDDYINT** style, even though it returns a valid position value in the **LOWORD** of the return.

If 32-bit values have been enabled for an up-down control with **SetRange32**, this message returns only the lower 16 bits of the position. To retrieve the full 32-bit position, use **GetPos32**.

---

### GetPos32

Returns the 32-bit position of an up-down control.

```
FUNCTION GetPos32 (BYVAL hUpDown AS HWND, BYVAL pfError AS LONG PTR = NULL) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *pfError*| Pointer to a LONG variable that that is set to zero if the value is successfully retrieved or nonzero if an error occurs. If this parameter is set to NULL, errors are not reported. If **SetPos32** is used in a cross-process situation, this parameter must be NULL. |

#### Return value

Returns the position of an up-down control with 32-bit precision. Applications must check the lParam value to determine whether the return value is valid.

#### Remarks

When it processes this message, the up-down control updates its current position based on the caption of the buddy window. It returns an error if there is no buddy window or if the caption specifies an invalid or out-of-range value.

---

### GetRange

Retrieves the minimum and maximum positions (range) for an up-down control.

```
FUNCTION GetRange (BYVAL hUpDown AS HWND) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |

#### Return value

The return value is a 32-bit value that contains the minimum and maximum positions. The **LOWORD** is the maximum position for the control, and the **HIWORD** is the minimum position.

---

### GetRange32

Retrieves the 32-bit range of an up-down control.

```
SUB GetRange32 (BYVAL hUpDown AS HWND, BYVAL pLow AS LONG PTR, BYVAL pHigh AS LONG PTR)
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *pLow* | Pointer to a signed integer that receives the lower limit of the up-down control range. This parameter may be NULL. |
| *pHigh* | Pointer to a signed integer that receives the upper limit of the up-down control range. This parameter may be NULL. |

---

### SetAccel

Sets the acceleration for an up-down control.

```
FUNCTION SetAccel (BYVAL hUpDown AS HWND, BYVAL cAccels AS DWORD, BYVAL pAccels AS UDACCEL PTR) AS BOOLEAN
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *cAccels* | Number of **UDACCEL** structures specified by aAccels.. |
| *pAccels* | Pointer to an array of **UDACCEL** structures that contain acceleration information. Elements should be sorted in ascending order based on the nSec member. |

#### Return value

Returns TRUE if successful, or FALSE otherwise.

---

### SetBase

Sets the radix base for an up-down control. The base value determines whether the buddy window displays numbers in decimal or hexadecimal digits. Hexadecimal numbers are always unsigned, and decimal numbers are signed.

```
FUNCTION SetBase (BYVAL hUpDown AS HWND, BYVAL nBase AS LONG) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *nBase* | New base value for the control. This parameter can be 10 for decimal or 16 for hexadecimal. |

#### Return value

The return value is the previous base value. If an invalid base is given, the return value is zero.

---

### SetBuddy

Sets the buddy window for an up-down control.

```
FUNCTION SetBuddy (BYVAL hUpDown AS HWND, BYVAL hwndBuddy AS HWND) AS HWND
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *hwndBuddy* | Handle to the new buddy window. |

#### Return value

The return value is the handle to the previous buddy window.

---

### SetPos

Sets the current position for an up-down control with 16-bit precision.

```
FUNCTION SetPos (BYVAL hUpDown AS HWND, BYVAL nPos AS SHORT) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *nPos* | New position for the up-down control. If the parameter is outside the control's specified range, lParam will be set to the nearest valid value. |

#### Return value

The return value is the previous position.

#### Remarks

This message only supports 16-bit positions. If 32-bit values have been enabled for an up-down control with **SetRange32**, use **SepPos32**.

---

### SetPos32

Sets the position of an up-down control with 32-bit precision.

```
FUNCTION SetPos32 (BYVAL hUpDown AS HWND, BYVAL nPos AS LONG) AS LONG
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *nPos* | Variable of type integer that specifies the new position for the up-down control. If the parameter is outside the control's specified range, *nPos* is set to the nearest valid value. |

#### Return value

Returns the previous position.

---

### SetRange

Sets the minimum and maximum positions (range) for an up-down control.

```
SUB SetRange (BYVAL hUpDown AS HWND, BYVAL nUpper AS SHORT, BYVAL nLower AS SHORT)
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *nUpper* | The maximum position for the up-down control. |
| *nLower* | The minimum position for the up-down control. |

#### Remarks

Neither position can be greater than the **UD_MAXVAL** value or less than the **UD_MINVAL** value. In addition, the difference between the two positions cannot exceed **UD_MAXVAL**.

The maximum position can be less than the minimum position. Clicking the up arrow button moves the current position closer to the maximum position, and clicking the down arrow button moves toward the minimum position.

---

### SetRange32

Sets the 32-bit range of an up-down control.

```
SUB SetRange32 (BYVAL hUpDown AS HWND, BYVAL iLow AS LONG, BYVAL iHigh AS LONG)
```

| Parameter | Description |
| --------- | ----------- |
| *hUpDown* | Handle to the up-down control. |
| *iLow* | Signed integer value that represents the new lower limit of the up-down control range. |
| *iHigh* | Signed integer value that represents the new upper limit of the up-down control range. |

---
