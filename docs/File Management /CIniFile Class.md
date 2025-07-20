# CIniFile Class

**CIniFile** is a wrapper class to work with .ini files.

Pass the path and name of the .ini file to read or write to when you create an instance of the **CIniFile** class. If you omit the path, the class will use the current directory. If the .ini file does not exist, it will create it.

* To retrieve data from the class, use the methods **GetString**, **GetDouble** or **GetInt**.
* To set data values, call one of the overloaded **WriteValue** methods.
* To delete a key, call the **DeleteKey** method.
* To delete an entire section, call the **DeleteSection** method.

Other useful methods are **GetKeyNames**, that returns a safe array with the names of all the keys of the specified section, **GetSectionNames**, that returns a safe array with the names of all the sections in the .ini file, and **GetSectionValues**, that returns a **Dictionary** object with the names and values of all the keys of the specified section.

**Include file**: AfxNova/CIniFile.inc

---

### Constructors

```
CONSTRUCTOR CIniFile ((BYREF wszFileName AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileName* | The path and name of the .ini file. If you omit the path, the class will use the current directory. If the .ini file does not exist, it will create it. |

#### Example

```
#include once "AfxNova/CIniFile.inc"
USING AfxMova

DIM cIni AS CInifile = "Test.ini"
cIni.WriteValue("Test", "Name", "Joe Doe")
DIM wszName AS WSTRING * 260
wszName = cIni.GetString("Test", "Name")
print wszName
```
---

### Methods

| Name       | Description |
| ---------- | ----------- |
| [DeleteKey](#deletekey) | Deletes a key from the specified section of an initialization file. |
| [DeleteSection](#deletesection) | Deletes a section from an initialization file. |
| [GetDouble](#getdouble) | Retrieves a numeric value from the specified section in an initialization file. |
| [GetInt](#getint) | Retrieves a numeric value from the specified section in an initialization file. |
| [GetKeyNames](#getkeynames) | Returns a safe array with the names of all the keys of the specified section. |
| [GetPath](#getpath) | Returns the full path of ini file this object instance is operating on. |
| [GetSectionNames](#getsectionnames) | Returns a safe array with the names of all sections in the ini file. |
| [GetSectionValues](#getsectionvalues) | Returns the keys and values of the specified section as a dictionary object. |
| [GetString](#getstring) | Retrieves a string from the specified section in an initialization file. |
| [WriteValue](#writevalue) | Writes a value into the specified section of an initialization file. |

---

## DeleteKey

Deletes a key from the specified section of an initialization file.

```
FUNCTION DeleteKey (BYREF wszSectionName AS WSTRING, BYREF wszKeyName AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSectionName* | The name of the section. |
| *wszKeyName* | The name of the key to delete. |

#### Return value

BOOLEAN. True on success or False on failure.

---

## DeleteSection

Deletes a section from an initialization file.

```
FUNCTION DeleteSection (BYREF wszSectionName AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSectionName* | The name of the section to delete. |

#### Return value

BOOLEAN. True on success or False on failure.

---

## GetDouble

Retrieves a numeric value from the specified section in an initialization file.

```
FUNCTION GetDouble (BYREF wszSectionName AS WSTRING, _
   BYREF wszKeyName AS WSTRING, BYVAL nDefaultValue AS DOUBLE = 0) AS DOUBLE
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSectionName* | The name of the section. |
| *wszKeyName* | The name of the key. |
| *nDefaultValue* | A default value to be returned if the key key cannot be found in the initialization file. |

#### Return value

DOUBLE. The retrieved value. If the key key cannot be found in the initialization file, the default value is returned.

---

## GetInt

Retrieves a numeric value from the specified section in an initialization file.

```
FUNCTION GetInt (BYREF wszSectionName AS WSTRING, _
   BYREF wszKeyName AS WSTRING, BYVAL nDefaultValue AS LONG = 0) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSectionName* | The name of the section. |
| *wszKeyName* | The name of the key. |
| *nDefaultValue* | A default value to be returned if the key key cannot be found in the initialization file. |

#### Return value

LONG. The retrieved value. If the key key cannot be found in the initialization file, the default value is returned.

---

## GetKeyNames

Returns a safe array with the names of all the keys of the specified section.

```
FUNCTION GetKeyNames () AS CSafeArray
```

#### Return value

DSAFEARRAY. The names of all the keys of the specified section.

#### Example

```
DIM cIni AS CInifile = "Test.ini"
DIM dsa AS DSafeArray = cIni.GetKeyNames("Test")
FOR i AS LONG = dsa.LBound TO dsa.UBound
   print dsa.GetStr(i)
NEXT
```
---

## GetPath

Returns the full path of the initialization file this object instance is operating on.

```
FUNCTION GetPath () AS DWSTRING
```

#### Return value

The path of the initialization file.

---

## GetSectionNames

Returns a safe array with the names of all sections in the initialization file.

```
FUNCTION GetSectionNames () AS DSafeArray
```

#### Return value

DSAFEARRAY. The names of all the sections of the file.

#### Example

```
DIM cIni AS CInifile = "Test.ini"
DIM dsa AS DSafeArray = cIni.GetSectionNames
FOR i AS LONG = dsa.LBound TO dsa.UBound
   print dsa.GetStr(i)
NEXT
```
---

## GetSectionValues

Returns the keys and values of the specified section as a dictionary object.

```
FUNCTION GetSectionValues (BYREF wszSectionName AS WSTRING, BYREF pDic AS CDicObj) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSectionName* | The name of the section. |
| *pDic* | A reference to an instance of the *CDicObj* class. |

#### Return value

BOOLEAN. True on success or False on failure.

#### Examples
```
DIM pDic AS CDicObj
IF cIni.GetSectionValues("Test", pDic) THEN
   print pDic.Item("Name")
END IF
```
```
DIM pDic AS CDicObj
IF cIni.GetSectionValues("Test", pDic) THEN
   ' // Get all the keys and display them
   DIM dvKeys AS DVARIANT = pDic.Keys
   FOR i AS LONG = dvKeys.GetLBound TO dvKeys.GetUBound
      print dvKeys.GetVariantElement(i)
   NEXT
   ' // Get all the items and display them
   DIM dvItems AS DVARIANT = pDic.Items
   FOR i AS LONG = dvItems.GetLBound TO dvItems.GetUBound
      print dvItems.GetVariantElement(i)
   NEXT
END IF
```
---

## GetString

Retrieves a string from the specified section in an initialization file.

```
FUNCTION GetString (BYREF wszSectionName AS WSTRING, _
   BYREF wszKeyName AS WSTRING, BYREF wszDefaultValue AS WSTRING = "") AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSectionName* | The name of the section. |
| *wszKeyName* | The name of the key. |
| *wszDefaultValue* | A default string. If the key key cannot be found in the initialization file, the default string is returned. Avoid specifying a default string with trailing blank characters. The function inserts a null character in the returned buffer to strip any trailing blanks. |

#### Return value

The retrieved string. If the key key cannot be found in the initialization file, the default value is returned.

---

## WriteValue

Writes a value into the specified section of an initialization file.

```
FUNCTION WriteValue (BYREF wszSectionName AS WSTRING, _
   BYREF wszKeyName AS WSTRING, BYREF wszValue AS WSTRING) AS BOOLEAN
```
```
FUNCTION WriteValue (BYREF wszSectionName AS WSTRING, _
   BYREF wszKeyName AS WSTRING, BYVAL dblValue AS DOUBLE) AS BOOLEAN
```
```
FUNCTION WriteValue (BYREF wszSectionName AS WSTRING, _
   BYREF wszKeyName AS WSTRING, BYVAL intValue AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszSectionName* | The name of the section. |
| *wszKeyName* | The name of the key. |
| *wszValue* / *dblValue* / *intValue* | The value to write. |

#### Return value

BOOLEAN. True on success or False on failure.

---
