# OpenFileDialog class

`OpenFileDialog`creates an dialog box that lets the user specify the drive, directory, and the name of a file or set of files to be opened.

# SaveFileDialog class

`SaveFileDialog` creates a dialog box that lets the user specify the drive, directory, and name of a file to save.

---

## Properties

| Name       | Description |
| ---------- | ----------- |
| [DefaultExt](#defaultext) | Gets/sets the default extension to be added to file names. |
| [FileName](#filename) | Gets/sets the file name that appears in the file name edit box when that dialog box is opened. |
| [FileType](#filetype) | Gets/sets the file type that appears as selected. |
| [Filter](#filter) | Gets/sets the file types that the dialog can open or save. |
| [FilterIndex](#filterindex) | Gets/sets the file type that appears as selected in the dialog. |
| [Flags](#flags) | Gets/sets flags that control the behavior of the dialog. |
| [FlagsEx](#flagsex) | Gets/sets extra flags that control the behavior of the dialog. |
| [InitialDir](#initialdir) | Gets/sets the folder used as a the initial directory. |
| [Title](#title) | Gets/sets the title of the dialog. |

---

## Methods

| Name       | Description |
| ---------- | ----------- |
| [ShowOpen](#showopen) | Displays the Open File Dialog. |
| [ShowSave](#showsave) | Displays the Save File Dialog. |

---

## <a name="defaultext"></a>DefaultExt

Gets/sets the default extension to be added to file names. This extension is appended to the file name if the user fails to type an extension. This string can be any length, but only the first three characters are appended. The string should not contain a period (.). If this property is empty and the user fails to type an extension, no extension is appended.

```
PROPERTY DefaultExt () AS DWSTRING
PROPERTY DefaultExt (BYREF wszDefaultExt AS WSTRING)
```
Example: DefaultExt = "BAS"

---

## <a name="filename"></a>FileName

Gets/sets the file name that appears in the file name edit box when that dialog box is opened.

```
PROPERTY FileName () AS DWSTRING
PROPERTY FileName (BYREF wszFileName AS WSTRING)
```
Example: FileName = "*.*"

---

## <a name="filter"></a>Filter

Gets/sets the file types that the dialog can open or save.
```
PROPERTY FileName () AS DWSTRING
PROPERTY FileName (BYREF dwsFilter AS DWSTRING)
```
A string containing pairs of "\|" separated strings. The first string in each pair is a display string that describes the filter (for example, "Text Files"), and the second string specifies the filter pattern (for example, "\*.TXT"). To specify multiple filter patterns for a single display string, use a semicolon to separate the patterns (for example, "\*.TXT;\*.DOC;\*.BAK"). A pattern string can be a combination of valid file name characters and the asterisk (\*) wildcard character. Do not include spaces in the pattern string. The system does not change the order of the filters. It displays them in the **File Types** combo box in the order specified in *dwsFilter*. |

Example: Filter = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"

---

## <a name="filterindex"></a>FilterIndex

Gets/sets the file type that appears as selected in the dialog.

```
PROPERTY FilterIndex () AS DWORD
PROPERTY FilterIndex (BYVAL dwFilterIndex AS DWORD)
```
Example: FilterIndex = 1

---

## <a name="flags"></a>Flags

Gets/sets the flags that control the behavior of the dialog.

```
PROPERTY Flags () AS DWORD
PROPERTY Flags (BYVAL dwFilterIndex AS DWORD)
```

| Flag       | Description |
| ---------- | ----------- |
| OFN_ALLOWMULTISELECT | **OpenFileDialog**: Allows multiple selections. |
| OFN_FILEMUSTEXIST | The user can type only names of existing files in the **File Name** entry field. If this flag is specified and the user enters an invalid name, the dialog box procedure displays a warning in a message box. If this flag is specified, the OFN_PATHMUSTEXIST flag is also used. |
| OFN_FORCESHOWHIDDEN | Forces the showing of system and hidden files, thus overriding the user setting to show or not show hidden files. However, a file that is marked both system and hidden is not shown. |
| OFN_NODEREFERENCELINKS | Directs the dialog box to return the path and file name of the selected shortcut (.LNK) file. If this value is not specified, the dialog box returns the path and file name of the file referenced by the shortcut. |
| OFN_PATHMUSTEXIST | The user can type only valid paths and file names. If this flag is used and the user types an invalid path and file name in the **File Name** entry field, the dialog box function displays a warning in a message box. |

Example: Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY

---

## <a name="flagsex"></a>FlagsEx

Gets/sets extra flags that control the behavior of the dialog. Currently, the only option is OFN_EX_NOPLACESBAR.

```
PROPERTY FlagsEx () AS DWORD
PROPERTY FlagsEx (BYVAL dwFilterIndex AS DWORD)
```
Example: FlagsEx = OFN_EX_NOPLACESBAR

---

## <a name="initialdir"></a>InitialDir

Gets/sets the folder used as a the initial directory. If no initial directory is specified, the default is the current directory.

```
PROPERTY InitialDir () AS DWSTRING
PROPERTY InitialDir (BYREF wszInitialDir AS WSTRING)
```
---

## <a name="title"></a>Title

Gets/sets the title of the dialog.

```
PROPERTY Title () AS DWSTRING
PROPERTY Title (BYREF wszTitle AS WSTRING)
```

If not specified, the default titles are "Open" for the **OpenFileDialog** and "Save" for the **SaveFileDialog**. These default names are localized.

---

## <a name="showopen"></a>ShopOpen

Displays the Open File Dialog.

```
FUNCTION ShowOpen (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, _
   BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| **hParent** | A handle to the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner. |
| **xPos**/**yPos** | Coordinates to position the dialog relative to the client area of the parent. If both *xPos* and *yPos* are -1, the dialog is centered. They are ignored if a hook is not used. |
| **bHook** | Boolean True or False. If True, the dialog is positioned according the specified coordinates and uses the old-style user interface. |

#### Remarks

If the OFN_ALLOWMULTISELECT flag is set and the user selects multiple files, the returned string contains the current directory followed by the file names of the selected files. For Explorer-style dialog boxes, the directory and file name strings are separated by semicolons. If the user selects only one file, the returned string does not have a separator between the path and file name.

Parse the number of ",". If only one, then the user has selected only a file and the string contains the full path. If more, The first substring contains the path and the others the files. If the user has not selected any file, an empty string is returned. On failure, an empty string is returned and, if not null, the pdwBufLen parameter will be filled by the size of the required buffer in characters.

---
