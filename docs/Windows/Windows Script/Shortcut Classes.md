# Shortcut Classes

**CShortcut** allows to create a shortcut programmatically.

**CURLShortcut** allows to create a URL shortcut programmatically.

**Include file**: AfxNova/CShortcut.inc

---

### Constructors

```
CONSTRUCTOR CShortcut (BYREF wszPathName AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPathName* | The full path and file name of the shortcut. |

#### Example

```
' Creates a shortcut programatically (if it already exists, opens it)
DIM pShortcut AS CShortcut = ExePath & "\Test.lnk"   ' --> change me
' Sets various properties and saves them to disk
pShortcut.Description = "Hello world"   ' --> change me
pShortcut.WorkingDirectory = ExePath & "\"   ' --> change me
pShortcut.Arguments = "/c"   ' --> change me
pShortcut.HotKey = "Ctrl+Alt+e"   ' --> change me
pShortcut.IconLocation = ExePath & "\Program.ico,0"   ' --> change me
pShortcut.RelativePath = ExePath & "\"   ' --> change me
pShortcut.TargetPath = ExePath & $"\HelloWord.exe"   ' --> change me
pShortcut.WindowStyle = WshNormalFocus
pShortcut.Save
```

```
CONSTRUCTOR CURLShortcut (BYREF wszPathName AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPathName* | The full path and file name of the shortcut. |

```
' Creates a shortcut programatically (if it already exists, opens it)
DIM pURLShortcut AS CURLShortcut = ExePath & "\Microsoft Web Site.url"   ' --> change me
pURLShortcut.TargetPath = "http://www.microsoft.com"   ' --> change me
pURLShortcut.Save
```
---

# CShortcut Methods

| Name       | Description |
| ---------- | ----------- |
| [Save](#save1) | Saves a shortcut object to disk. |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [GetErrorInfo](#geterrorinfo) | Returns a localized description of the specified error code. |

---

# CShortcut Properties

| Name       | Description |
| ---------- | ----------- |
| [Arguments](#arguments) | Gets/sets the arguments for a shortcut, or identifies a shortcut's arguments. |
| [Description](#description) | Returns or sets a shortcut's description. |
| [FullName](#fullname1) | Returns the fully qualified path of the shortcut object's target. |
| [Hotkey](#hotkey) | Assigns a key-combination to a shortcut, or identifies the key-combination assigned to a shortcut. |
| [IconLocation](#iconlocation) | Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut. |
| [RelativePath](#relativepath) | Assigns a relative path to a shortcut. |
| [TargetPath](#targetpath1) | Gets/sets the path of the shortcut's executable. |
| [WindowStyle](#windowstyle) | Assigns a window style to a shortcut, or identifies the type of window style used by a shortcut. |
| [WorkingDirectory](#workingdirectory) | Assigns a working directory to a shortcut, or identifies the working directory used by a shortcut. |

---

# CURLShortcut Methods

| Name       | Description |
| ---------- | ----------- |
| [Save](#save2) | Saves a URL shortcut object to disk. |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [GetErrorInfo](#geterrorinfo) | Returns a localized description of the specified error code. |

---

# CShortcut Properties

| Name       | Description |
| ---------- | ----------- |
| [FullName](#fullname2) | Returns the fully qualified path of the shortcut object's target. |
| [TargetPath](#targetpath2) | Gets/sets the path of the shortcut's executable. |

---

## <a name="save1"></a>Save (CShortcut)

Saves a shortcut object to disk.

```
FUNCTION Save () AS HRESULT
```

#### Remarks

After creating an instance of the **CShortcut* class to create to create a shortcut object and set the shortcut object's properties, the Save method must be used to save the shortcut object to disk. The **Save** method uses the information in the shortcut object's **FullName** property to determine where to save the shortcut object on a disk. You can only create shortcuts to system objects. This includes files, directories, and drives (but does not include printer links or scheduled tasks).

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## <a name="save2"></a>Save (CURLShortcut)

Saves a URL shortcut object to disk.

```
FUNCTION Save () AS HRESULT
```

#### Remarks

After creating an instance of the **CURLShortcut** class to create to create a shortcut object and set the shortcut object's properties, the Save method must be used to save the shortcut object to disk. The **Save** method uses the information in the shortcut object's **FullName** property to determine where to save the shortcut object on a disk. You can only create shortcuts to system objects. This includes files, directories, and drives (but does not include printer links or scheduled tasks).

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## <a name="getlastresult"></a>GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## <a name="geterrorinfo"></a>GetErrorInfo

Returns a localized description of the specified error code. If the error is omited, it will return the value returned by **GetLastResult**.

```
FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
```

---

## <a name="arguments"></a>Arguments

Gets/sets the arguments for a shortcut, or identifies a shortcut's arguments.

```
PROPERTY Arguments () AS DWSTRING
PROPERTY Arguments (BYREF wszArguments AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszArguments* | The arguments for the shortcut. |

#### Return value

The arguments of the shortcut.

---

## <a name="description"></a>Description

Returns or sets a shortcut's description.

```
PROPERTY Description () AS DWSTRING
PROPERTY Description (BYREF wszDescription AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDescription* | A string value describing a shortcut. |

#### Return value

The description of the shortcut.

---

## <a name="fullname1"></a>FullName (CShortcut)

Returns the fully qualified path of the shortcut object's target.

```
PROPERTY FullName () AS DWSTRING
```

#### Remarks

The **FullName** property contains a read-only string value indicating the fully qualified path to the shortcut's target.

---

## <a name="fullname2"></a>FullName (CURLShortcut)

Returns the fully qualified path of the shortcut object's target.

```
PROPERTY FullName () AS DWSTRING
```

#### Remarks

The **FullName** property contains a read-only string value indicating the fully qualified path to the shortcut's target.

---

## <a name="hotkey"></a>Hotkey

Assigns a key-combination to a shortcut, or identifies the key-combination assigned to a shortcut.

```
PROPERTY Hotkey () AS DWSTRING
PROPERTY Hotkey (BYREF wszHotkey AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszHotkey* | A string representing the key-combination to assign to the shortcut. |

#### Return value

The hotkey of the shortcut.

---

## <a name="iconlocation"></a>IconLocation

Assigns a an icon to a shortcut, or identifies the icon assigned to a shortcut.

```
PROPERTY IconLocation () AS DWSTRING
PROPERTY IconLocation (BYREF wszIconLocation AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszIconLocation* | A string that locates the icon. The string should contain a fully qualified path and an index associated with the icon. |

#### Return value

The location of the icon.

---

## <a name="relativepath"></a>RelativePath

Assigns a relative path to a shortcut.

```
PROPERTY RelativePath (BYREF wszRelativePath AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszRelativePath* | A string containing the relative path. |

---

## <a name="targetpath1"></a>TargetPath (CShortcut)

Gets/sets the path of the shortcut's executable.

```
PROPERTY TargetPath () AS DWSTRING
PROPERTY TargetPath (BYREF wszTargetPath AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszTargetPath* | A string containing the path of the shortcut's executable. |

#### Return value

The path of the shortcut's executable.

---

## <a name="targetpath2"></a>TargetPath (CURLShortcut)

Gets/sets the path of the shortcut's executable.

```
PROPERTY TargetPath () AS DWSTRING
PROPERTY TargetPath (BYREF wszTargetPath AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszTargetPath* | A string containing the path of the shortcut's executable. |

#### Return value

The path of the shortcut's executable.

---

## <a name="windowstyle"></a>WindowStyle

Assigns a window style to a shortcut, or identifies the type of window style used by a shortcut.

```
PROPERTY WindowStyle () AS LONG
PROPERTY WindowStyle (BYVAL nWindowStyle AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *nWindowStyle* | The window style for the program being run. |

The following table lists the available settings for *nWindowStyle*.

| Style      | Description |
| ---------- | ----------- |
| 1 | Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. |
| 3 | Activates the window and displays it as a maximized window. |
| 7 | Minimizes the window and activates the next top-level window. |

#### Return value

LONG. The window style.

---

## <a name="workingdirectory"></a>WorkingDirectory

Assigns a working directory to a shortcut, or identifies the working directory used by a shortcut.

```
PROPERTY WorkingDirectory () AS DWSTRING
PROPERTY WorkingDirectory (BYREF wszWorkingDirectory AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszWorkingDirectory* | Directory in which the shortcut starts. |

#### Return value

The working directory of the shortcut.

---
