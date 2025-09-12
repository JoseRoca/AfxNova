# CDicObj Class

`CDicObj` is an associative array of variants. Each item is associated with a unique key. The key is used to retrieve an individual item.

#### Example

```
' // Creates an instance of the CDicObj class
DIM pDic AS CDicObj

' // Adds some key, value pairs
pDic.Add "a", "Athens"
pDic.Add "b", "Belgrade"
pDic.Add "c", "Cairo"

print "Count: "; pDic.Count
print pDic.Exists("a")
print

' // Retrieve an item and display it
print pDic.Item("b")
print

' // Change key "b" to "m" and "Belgrade" to "México"
pDic.Key("b") = "m"
pDic.Item("m") = "México"
print pDic.Item("m")
print

' // Get all the items and display them
DIM dvItems AS DVARIANT = pDic.Items
FOR i AS LONG = dvItems.GetLBound TO dvItems.GetUBound
  print dvItems.GetVariantElement(i)
NEXT
print

' // Get all the keys and display them
DIM dvKeys AS DVARIANT = pDic.Keys
FOR i AS LONG = dvKeys.GetLBound TO dvKeys.GetUBound
  print dvKeys.GetVariantElement(i)
NEXT
print

' // Remove key "m"
pDic.Remove "m"
IF pDic.Exists("m") THEN PRINT "Key m exists" ELSE PRINT "Key m doesn't exist"

' // Remove all keys
pDic.RemoveAll
print "All the keys must have been deleted"
print "Count: "; pDic.Count
print
```

**Include file**: CDicObj.inc

### Methods and Properties

| Name       | Description |
| ---------- | ----------- |
| [Add](#add) | Adds a key and item pair to the associative array. |
| [Count](#count) | Returns the number of items in the associative array. |
| [DispObj](#dispObj) | Returns a counted reference of the underlying dispatch pointer. |
| [DispPtr](#dispPtr) | Returns the underlying dispatch pointer. |
| [Exists](#exists) | Checks if a specified key exists in the associative array. |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [HashVal](#hashval) | Returns the hash value for a specified key in the associative array. |
| [Item](#item) | Sets or returns an item for a specified key in associative array. |
| [Items](#items) | Returns a safe array containing all the items in the associative array. |
| [Key](#key) | Sets or returns an item for a specified key in the associative array. |
| [Keys](#keys) | Returns an array containing all the keys in the associative array. |
| [NewEnum](#newenum) | Returns a reference to the standard enumerator. |
| [Remove](#remove) | Removes a key, item pair from the associative array. |
| [RemoveAll](#removeall) | Removes all key, item pairs from the associative array. |

## Error and result codes

| Name       | Description |
| ---------- | ----------- |
| [GetErrorInfo](#geterrorinfo) | Returns a localized description of the specified error code. |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [SetResult](#setresult) | Sets the last result code. |

---

### <a name="geterrorinfo"></a>GetErrorInfo

Returns a localized description of the specified error code. If the error is omited, it will return the value returned by the Windows API function **GetLastError**.
```
PRIVATE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
```
---

### <a name="getlastresult"></a>GetLastResult

Returns the last result code
```
FUNCTION GetLastResult () AS HRESULT
```
---

### <a name="setresult"></a>SetResult

Sets the last result code.
```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Result* | The **HRESULT** error code returned by the methods. |

---

### <a name="Add"></a>Add

Adds a key and item pair to the associtive array.

```
FUNCTION Add (BYREF dvKey AS DVARIANT, BYREF dvItem AS DVARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dvKey* | The key associated with the item being added. |
| *dvItem* | The item associated with the key being added. |

#### Return value

An error occurs if the key already exists.

---

### <a name="Count"></a>Count

Returns the number of items in the associative array.

```
FUNCTION Count () AS LONG
```
---

### <a name="DispObj"></a>DispObj

Returns a counted reference of the underlying dispatch pointer. You must call **IUnknown_Release** when no longer needs it.

```
FUNCTION DispObj () AS ANY PTR
```
---

### <a name="DispPtr"></a>DispPtr

Returns the underlying dispatch pointer. As it is a raw pointer, don't call **IUnknown_Release** on it.

```
FUNCTION DispPtr () AS ANY PTR
```
---

### <a name="Exists"></a>Exists

Checks if a specified key exists in the associative array.

```
FUNCTION Exists (BYREF dvKey AS DVARIANT) AS BOOLEAN
```

#### Return value

Returns True is the key exists; False, otherwise.

---

### <a name="HashVal"></a>HashVal

Returns the hash value for a specified key in the associative array.

```
FUNCTION HashVal (BYREF dvKey AS DVARIANT) AS DVARIANT
```
---

### <a name="Item"></a>Item

Sets or returns an item for a specified key in associative array.

```
PROPERTY Item (BYREF dvKey AS DVARIANT) AS DVARIANT
PROPERTY Item (BYREF dvKey AS DVARIANT, BYREF dvNewItem AS DVARIANT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dvKey* | Key associated with the item being retrieved or added. |
| *dvNewItem* | The new value associated with the specified key. |

#### Return value

The item value.

#### Remarks

If key is not found when changing an item, a new key is created with the specified *dvNewvItem*. Contrary to the Dictionary object, if key is not found when attempting to return an existing item, it returns and empty variant and sets the last result code to E_INVALIDARG, instead of creating a new key with its corresponding item empty.

---

### <a name="Items"></a>Items

Returns a safe array containing all the items in the associative array.

```
FUNCTION Items () AS DVARIANT
```

#### Return value

A DVARIANT containing all the items in a safe array.

---

### <a name="Key"></a>Key

Sets or returns an item for a specified key in the associative array.

```
PROPERTY Key (BYREF dvKey AS DVARIANT, BYREF dvNewKey AS DVARIANT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dvKey* | Key value being changed. |
| *dvNewKey* | New value that replaces the specified key. |

#### Remarks

Contrarily to the Dictionary object, if key is not found when changing a key, this method sets the last result code to E_INVALIDARG and exits, instead of creating a new key with its associated item empty.

#### Return value

The item value.

---

### <a name="Keys"></a>Keys

Returns an array containing all the keys in the associative array.

```
FUNCTION Keys () AS DVARIANT
```

#### Return value

A DVARIANT containing all the keys in a safe array.

---

### <a name="NewEnum"></a>NewEnum

Returns a reference to the standard enumerator.

```
FUNCTION NewEnum () AS IEnumVARIANT PTR
```

#### Return value

A pointer to the standard **IEnumVARIANT** interface.

#### Return value

IUnknown pointer that must be cast to an **IEnumVARIANT** interface.

---

### <a name="Remove"></a>Remove

Removes a key, item pair from the associative array.

```
FUNCTION Remove (BYREF dvKey AS DVARIANT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dvKey* | Key associated with the key, item pair you want to remove from the associative array.

#### Return value

An error occurs if the specified key, item pair does not exist.

---

### <a name="RemoveAll"></a>RemoveAll

Removes all key, item pairs from the associative array.

```
FUNCTION RemoveAll() AS HRESULT
```

#### Return value

Returns S_OK (0) or an HRESULT error code.

---
