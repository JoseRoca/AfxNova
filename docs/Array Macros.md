# Array Macros

Macros to manipulate one-dimensional dynamic arrays.

**Include file**: AfxNova/AfxArrays.inc.

| Name       | Description |
| ---------- | ----------- |
| [AppendElementToArray](#appendelementtoarray) | Appends a new element to the array. |
| [AppendArrayToArray](#appendarraytoarray) | Appends one array to another array. |
| [InsertElementIntoArray](#insertelementintoarray) | Inserts a new element into an array. |
| [InsertArrayIntoArray](#insertarrayintoarray) | Inserts an array into another array. |
| [RemoveElementFromArray](#removeelementfromarray) | Removes the specified element from an array. |
| [RemoveFirstElementFromArray](#removefirstelementfromarray) | Removes the first element from an array. |
| [RemoveLastElementFromArray](#removelastelementfromarray) | Removes the last element from an array. |
| [RemoveElementsFromArray](#removeelementsfromarray) | Removes the specified elements from an array. |

---

# Sorting Macros

**Include file**: AfxSort2.inc.

| Name       | Description |
| ---------- | ----------- |
| [AfxSortNumericArray](#afxsortnumericarray) | Sorts one-dimensional numeric arrays of any type. |
| [AfxSortStringArray](#afxsortstringarray) | Sorts one-dimensional string arrays of any type. |

---

### <a name="appendelementtoarray"></a>AppendElementToArray

Appends a new element to a dynamic one-dimensional array.
```
#macro AppendElementToArray(rg, elem)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *elem* | The element to append. |

#### Remarks

The array and the element to append can be of any type, but they have to be of the same type between them.

If the array has a fixed length, **GetLastError** will return an E_ACCESSDENIED error.

#### Usage example

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i))
NEXT
' Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="appendarraytoarray"></a>AppendArrayToArray

Appends a one-dimensional array to another dynamic one-dimensional array.
```
#macro AppendArrayToArray(rg, rg2)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *rg2* | The array to append. |

#### Remarks

The array and the array to append can be of any type, but they have to be of the same type between them.

If the array has a fixed length, **GetLastError** will return an E_ACCESSDENIED error.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM rg2(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i))
NEXT
' Fill the array to append
xStr = "String 2 - "
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, xStr & WSTR(i))
NEXT
' // Append the array to the first array
AppendArrayToArray(rg, rg2)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
DIM rg2(ANY) AS LONG
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, nLong)
   nLong += 1
NEXT
' Fill the array to insert
nLong = 12345
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, nLong)
   nLong += 1
NEXT
AppendArrayToArray(rg, rg2)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="insertelementintoarray"></a>InsertElementIntoArray

Inserts a new element before the specified position into a dynamic one-dimensional array.
```
#macro InsertElementIntoArray(rg, pos, elem)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *pos* | The position in the array where the new element will be added. This position is relative to the lower bound of the array. |
| *elem* | The element to insert. |

#### Remarks

The array and the array to append can be of any type, but they have to be of the same type between them.

If the array has a fixed length, **GetLastError** will return an E_ACCESSDENIED error.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i))
NEXT
InsertElementIntoArray(rg, 5, xStr & WSTR(11))
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
DIM rg2(ANY) AS LONG
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, nLong)
   nLong += 1
NEXT
' Fill the array to insert
nLong = 12345
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, nLong)
   nLong += 1
NEXT
InsertElementIntoArray(rg, 5, nLong)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="insertarrayintoarray"></a>InsertArrayIntoArray

Inserts a one-dimensional array before the specified position in another dynamic one-dimensional array.
```
#macro InsertArrayIntoArray(rg, pos, rg2)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *pos* | The position in the array where the new element will be added. This position is relative to the lower bound of the array. |
| *rg2* | The array to insert. |

#### Remarks

The array and the array to insert can be of any type, but they have to be of the same type between them.

If the array has a fixed length, **GetLastError** will return an E_ACCESSDENIED error.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM rg2(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i))
NEXT
' Fill the array to insert
xStr = "String 2 - "
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, xStr & WSTR(i))
NEXT
' // Insert the array to the first array
InsertArrayIntoArray(rg, 5, rg2)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
DIM rg2(ANY) AS LONG
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, nLong)
   nLong += 1
NEXT
' Fill the array to insert
nLong = 12345
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rg2, nLong)
   nLong += 1
NEXT
' // Insert the array to the first array
InsertArrayIntoArray(rg, 5, rg2)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="removeelementfromarray"></a>RemoveElementFromArray

Removes the specified element of a dynamic one-dimensional array.
```
#macro RemoveElementFromArray(rg, pos)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *pos* | The position in the array of the element to remove. This position is relative to the lower bound of the array. |

#### Remarks

The array can be of any type.

If the array has a fixed length, **GetLastError** will return an E_ACCESSDENIED error.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i))
NEXT
' // Remove the fifth element
RemoveElementFromArray(rg, 5)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   REDIM PRESERVE rg(UBOUND(rg) + 1)
   rg(i - 1) = nLong
   nLong += 1
NEXT
' // Remove the specified element
RemoveElementFromArray(rg, 5)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="removefirstelementfromarray"></a>RemoveFirstElementFromArray

Removes the first element of a dynamic one-dimensional array.
```
#macro RemoveFirstElementFromArray(rg)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |

#### Remarks

The array can be of any type.

If the array has a fixed length, **GetLastError** will return an E_ACCESSDENIED error.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i))
NEXT
' // Remove the frst element
RemoveFirstElementFromArray(rg)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG
' // Fill the array
DIM nLong AS LONG = 12345
FOR i AS LONG = 1 TO 10
   REDIM PRESERVE rg(UBOUND(rg) + 1)
   rg(i - 1) = nLong
   nLong += 1
NEXT
' // Remove the first element
RemoveFirstElementFromArray(rg)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="removelastelementfromarray"></a>RemoveLastElementFromArray

Removes the last element of a dynamic one-dimensional array.
```
#macro RemoveLastElementFromArray(rg)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |

#### Remarks

The array can be of any type.

If the array has a fixed length, **GetLastError** will return an E_ACCESSDENIED error.

#### Usage examples

```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rg(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, xStr & WSTR(i))
NEXT
' // Remove the last element
RemoveLastElementFromArray(rg)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
#### Can also be used with numbers:
```
DIM rg(ANY) AS LONG   ' // or DWORD, SINGLE, DOUBLE, etc.
' // Fill the array
DIM nLong AS LONG = 12345
FOR i AS LONG = 1 TO 10
   REDIM PRESERVE rg(UBOUND(rg) + 1)
   rg(i - 1) = nLong
   nLong += 1
NEXT
' // Remove the last element
RemoveLastElementFromArray(rg)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```

### <a name="Removeelementsfromarray"></a>RemoveElementsFromArray

Removes the specified elements of a dynamic one-dimensional array.
```
#macro RemoveElementsFromArray(rg, rgElements)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array. |
| *rgElements* | Array of long values indicating the 0-based positions of the elements to remove. |

#### Remarks

The array can be of any type.

If the array has a fixed length, **GetLastError** will return an E_ACCESSDENIED error.

#### Usage example

```
DIM rg(ANY) AS LONG
' // Fill the array
DIM nLong AS LONG = 1
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rg, nLong)
   nLong += 1
NEXT
' Fill the array of elements to delete
DIM rgElements(0 TO 3) AS LONG => {2, 4, 6, 8}
RemoveElementsFromArray(rg, rgElements)
' // Display the array
FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
   print rg(i)
NEXT
```
---

### <a name="afxsortnumericarray"></a>AfxSortNumericArray

Sorts one-dimensional numeric array of any type. The macro detects the type of array and calls the appropriate sorting procedures.
```
#macro AfxSortNumericArray(rg, ascending)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array to sort. |
| *ascending* | If true, sorts the array in ascending order. I false, sorts the array in descending order. |

#### Return code

It does not return a result code.

#### Usage example
```
#define XNUMBER DOUBLE ' // or SHORT, INTEGER, LONG, LONGINT, etc.
DIM rgNum(ANY) AS XNUMBER
DIM rgNum2(ANY) AS XNUMBER
DIM xNum AS XNUMBER = 1234567.89
' // Fill the array
FOR i AS LONG = 1 TO 10
   xNum += 0.01
   AppendElementToArray(rgNum, xNum)
NEXT
' Fill the array to insert
xNum = 111.01
FOR i AS LONG = 1 TO 5
   xNum += 1
   AppendElementToArray(rgNum, xNum)
NEXT
' // Insert the array to the first array
InsertArrayIntoArray(rgNum, 5, rgNum2)
' // Display the array
FOR i AS LONG = LBOUND(rgNum) TO UBOUND(rgNum)
   print rgNum(i)
NEXT
' // Sort the numeric array
AfxSortNumericArray(rgNum, true)
' // Display the sorted array
FOR i AS LONG = LBOUND(rgNum) TO UBOUND(rgNum)
   print rgNum(i)
NEXT
```
---

### <a name="afxsortstringarray"></a>AfxSortStringArray

Sorts one-dimensional string array of any type. The macro detects the type of array and calls the appropriate sorting procedures.
```
#macro AfxSortStringArray(rg, ascending)
```
| Parameter  | Description |
| ---------- | ----------- |
| *rg* | The array to sort. |
| *ascending* | If true, sorts the array in ascending order. I false, sorts the array in descending order. |

#### Return code

It does not return a result code.

#### Usage example
```
#define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
DIM rgstr(ANY) AS XSTRING
DIM rgstr2(ANY) AS XSTRING
DIM xStr AS XSTRING = "String - "
' // Fill the array
FOR i AS LONG = 1 TO 10
   AppendElementToArray(rgstr, xStr & WSTR(i))
NEXT
' Fill the array to append
xStr = "String 2 - "
FOR i AS LONG = 1 TO 5
   AppendElementToArray(rgstr2, xStr & WSTR(i))
NEXT
' // Append the array to the first array
AppendArrayToArray(rgstr, rgstr2)
' // Display the array
FOR i AS LONG = LBOUND(rgstr) TO UBOUND(rgstr)
   print rgstr(i)
NEXT
' // Sort the array
AfxSortStringArray(rgstr, true)
' // Display the sorted array
FOR i AS LONG = LBOUND(rgstr) TO UBOUND(rgstr)
   print rgstr(i)
NEXT
```
---
