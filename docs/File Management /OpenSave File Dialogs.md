# OpenFileDialog class

`OpenFileDialog`creates a dialog box that lets the user specify the drive, directory, and the name of a file or set of files to be opened.

# SaveFileDialog class

`SaveFileDialog` creates a dialog box that lets the user specify the drive, directory, and name of a file to save.

---

## Constructors (OpenFileDialog/SaveFileDialog)

| Name       | Description |
| ---------- | ----------- |
| Constructor (OpenFileDialog) | Creates a dialog box that lets the user specify the drive, directory, and the name of a file or set of files to be opened. |
| Constructor (SaveFileDialog) | Creates a dialog box that lets the user specify the drive, directory, and name of a file to save. |

#### Usage

```
DIM pofd AS OpenFileDialog
DIM psfd AS SaveFileDialog
```
---

## Properties

| Name       | Description |
| ---------- | ----------- |
| [DefaultExt](#defaultext) | Gets/sets the default extension to be added to file names. |
| [FileName](#filename) | Gets/sets the file name that appears in the dialog's file name edit box when that dialog box is opened. |
| [Filter](#filter) | Gets/sets the file types that the dialog can open or save. |
| [FilterIndex](#filterindex) | Gets/sets the file type that appears as selected in the dialog. |
| [Flags](#flags) | Gets/sets flags that control the behavior of the dialog. |
| [FlagsEx](#flagsex) | Gets/sets extra flags that control the behavior of the dialog.  Currently, the only option is OFN_EX_NOPLACESBAR. |
| [InitialDir](#initialdir) | Gets/sets the folder used as the initial directory. |
| [Title](#title) | Gets/sets the title of the dialog. |

---

## Methods

| Name       | Description |
| ---------- | ----------- |
| [ShowOpen](#showopen) | Displays the Open File Dialog. |
| [ShowSave](#showsave) | Displays the Save File Dialog. |

---

## DefaultExt

Gets/sets the default extension to be added to file names. This extension is appended to the file name if the user fails to type an extension. This string can be any length, but only the first three characters are appended. The string should not contain a period (.). If this property is empty and the user fails to type an extension, no extension is appended.

```
PROPERTY DefaultExt () AS DWSTRING
PROPERTY DefaultExt (BYREF wszDefaultExt AS WSTRING)
```
Example: DefaultExt = "BAS"

---

## FileName

Gets/sets the file name that appears in the file name edit box when that dialog box is opened.

```
PROPERTY FileName () AS DWSTRING
PROPERTY FileName (BYREF wszFileName AS WSTRING)
```
Example: FileName = "*.*"

---

## Filter

Gets/sets the file types that the dialog can open or save.
```
PROPERTY FileName () AS DWSTRING
PROPERTY FileName (BYREF dwsFilter AS DWSTRING)
```
A string containing pairs of "\|" separated strings. The first string in each pair is a display string that describes the filter (for example, "Text Files"), and the second string specifies the filter pattern (for example, "\*.TXT"). To specify multiple filter patterns for a single display string, use a semicolon to separate the patterns (for example, "\*.TXT;\*.DOC;\*.BAK"). A pattern string can be a combination of valid file name characters and the asterisk (\*) wildcard character. Do not include spaces in the pattern string. The system does not change the order of the filters. It displays them in the **File Types** combo box in the order specified in *dwsFilter*. |

Example: Filter = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"

---

## FilterIndex

Gets/sets the file type that appears as selected in the dialog.

```
PROPERTY FilterIndex () AS DWORD
PROPERTY FilterIndex (BYVAL dwFilterIndex AS DWORD)
```
Example: FilterIndex = 1

---

## Flags

Gets/sets the flags that control the behavior of the dialog.

```
PROPERTY Flags () AS DWORD
PROPERTY Flags (BYVAL dwFilterIndex AS DWORD)
```

#### OpenFileDialog

| Flag       | Description |
| ---------- | ----------- |
| OFN_ALLOWMULTISELECT | The **File Name** list box allows multiple selections. |
| OFN_DONTADDTORECENT | Prevents the system from adding a link to the selected file in the file system directory that contains the user's most recently used documents. To retrieve the location of this directory, call the **SHGetSpecialFolderLocation** function with the CSIDL_RECENT flag. |
| OFN_EXTENSIONDIFFERENT | The user typed a file name extension that differs from the extension specified by *wszDefExt*. The function does not use this flag if *wszDefExt* is NULL. |
| OFN_FILEMUSTEXIST | The user can type only names of existing files in the **File Name** entry field. If this flag is specified and the user enters an invalid name, the dialog box procedure displays a warning in a message box. If this flag is specified, the OFN_PATHMUSTEXIST flag is also used. |
| OFN_FORCESHOWHIDDEN | Forces the showing of system and hidden files, thus overriding the user setting to show or not show hidden files. However, a file that is marked both system and hidden is not shown. |
| OFN_HIDEREADONLY | Hides the **Read Only** check box. |
| OFN_NODEREFERENCELINKS | Directs the dialog box to return the path and file name of the selected shortcut (.LNK) file. If this value is not specified, the dialog box returns the path and file name of the file referenced by the shortcut. |
| OFN_NONETWORKBUTTON | Hides and disables the **Network** button. |
| OFN_NOREADONLYRETURN | The returned file does not have the **Read Only** check box selected and is not in a write-protected directory. |
| OFN_PATHMUSTEXIST | The user can type only valid paths and file names. If this flag is used and the user types an invalid path and file name in the **File Name** entry field, the dialog box function displays a warning in a message box. |
| OFN_READONLY | Causes the **Read Only** check box to be selected initially when the dialog box is created. This flag indicates the state of the **Read Only** check box when the dialog box is closed. |
| OFN_SHOWHELP | Causes the dialog box to display the **Help** button. The *hwndOwner* member must specify the window to receive the HELPMSGSTRING registered messages that the dialog box sends when the user clicks the **Help** button. |

Example: Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY

#### SaveFileDialog

| Flag       | Description |
| ---------- | ----------- |
| OFN_CREATEPROMPT | If the user specifies a file that does not exist, this flag causes the dialog box to prompt the user for permission to create the file. If the user chooses to create the file, the dialog box closes and the function returns the specified name; otherwise, the dialog box remains open. |
| OFN_DONTADDTORECENT | Prevents the system from adding a link to the selected file in the file system directory that contains the user's most recently used documents. To retrieve the location of this directory, call the **SHGetSpecialFolderLocation** function with the CSIDL_RECENT flag. |
| OFN_EXTENSIONDIFFERENT | The user typed a file name extension that differs from the extension specified by *wszDefExt*. The function does not use this flag if *wszDefExt* is NULL. |
| OFN_FORCESHOWHIDDEN | Forces the showing of system and hidden files, thus overriding the user setting to show or not show hidden files. However, a file that is marked both system and hidden is not shown. |
| OFN_HIDEREADONLY | Hides the **Read Only** check box |
| OFN_NOCHANGEDIR | Restores the current directory to its original value if the user changed the directory while searching for files. |
| OFN_NODEREFERENCELINKS | Directs the dialog box to return the path and file name of the selected shortcut (.LNK) file. If this value is not specified, the dialog box returns the path and file name of the file referenced by the shortcut. |
| OFN_NOTESTFILECREATE | The file is not created before the dialog box is closed. This flag should be specified if the application saves the file on a create-nonmodify network share. When an application specifies this flag, the library does not check for write protection, a full disk, an open drive door, or network protection. Applications using this flag must perform file operations carefully, because a file cannot be reopened once it is closed. |
| OFN_NONETWORKBUTTON | Hides and disables the **Network** button. |
| OFN_NOREADONLYRETURN | The returned file does not have the **Read Only** check box selected and is not in a write-protected directory. |
| OFN_OVERWRITEPROMPT | The user can type only valid paths and file names. If this flag is used and the user types an invalid path and file name in the **File Name** entry field, the dialog box function displays a warning in a message box. |
| OFN_PATHMUSTEXIST | The user can type only valid paths and file names. If this flag is used and the user types an invalid path and file name in the **File Name** entry field, the dialog box function displays a warning in a message box. |
| OFN_SHOWHELP | Causes the dialog box to display the **Help** button. The hwndOwner member must specify the window to receive the HELPMSGSTRING registered messages that the dialog box sends when the user clicks the **Help** button. |

Example: Flags = OFN_HIDEREADONLY OR OFN_OVERWRITEPROMPT OR OFN_CREATEPROMPT

---

## FlagsEx

Gets/sets extra flags that control the behavior of the dialog. Currently, the only option is OFN_EX_NOPLACESBAR.

```
PROPERTY FlagsEx () AS DWORD
PROPERTY FlagsEx (BYVAL dwFilterIndex AS DWORD)
```
Example: FlagsEx = OFN_EX_NOPLACESBAR

---

## InitialDir

Gets/sets the folder used as a the initial directory. If no initial directory is specified, the default is the current directory.

```
PROPERTY InitialDir () AS DWSTRING
PROPERTY InitialDir (BYREF wszInitialDir AS WSTRING)
```
---

## Title

Gets/sets the title of the dialog.

```
PROPERTY Title () AS DWSTRING
PROPERTY Title (BYREF wszTitle AS WSTRING)
```

If not specified, the default titles are "Open" for the **OpenFileDialog** and "Save" for the **SaveFileDialog**. These default names are localized.

---

## ShowOpen

Displays the Open File Dialog.

```
FUNCTION ShowOpen (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, _
   BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *hParent* | A handle to the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner. |
| *xPos*/*yPos* | Coordinates to position the dialog relative to the client area of the parent. If both *xPos* and *yPos* are -1, the dialog is centered. They are ignored if a hook is not used. |
| *bHook* | Boolean True or False. If True, the dialog is positioned according the specified coordinates and uses the old-style user interface. |

#### Remarks

If the OFN_ALLOWMULTISELECT flag is set and the user selects multiple files, the returned string contains the current directory followed by the file names of the selected files. For Explorer-style dialog boxes, the directory and file name strings are separated by semicolons. If the user selects only one file, the returned string does not have a separator between the path and file name.

Parse the number of ",". If only one, then the user has selected only a file and the string contains the full path. If more, The first substring contains the path and the others the files. If the user has not selected any file, an empty string is returned. On failure, an empty string is returned and, if not null, the pdwBufLen parameter will be filled by the size of the required buffer in characters.

#### Usage example:

```
DIM pOFN AS OpenFileDialog
pOFN.FileName = "*.*"
pOFN.InitialDir = CURDIR
pOFN.Filter = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
pOFN.DefaultExt = "BAS"
pOFN.FilterIndex = 1
pOFN.Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY
DIM dwsFile AS DWSTRING = pOFN.ShowOpen(hParent)
```
To allow myltiple selection use:
```
pOFN.Flags = OFN_FILEMUSTEXIST OR OFN_HIDEREADONLY OR OFN_ALLOWMULTISELECT
```
To position the dialog box use:
```
DIM dwsFile AS DWSTRING = pOFN.ShowOpen(0, 20, 20, FALSE)
```
To center the dialog use:
```
DIM dwsFile AS DWSTRING = pOFN.ShowOpen(0, -1, -1, TRUE)
```
---

## ShowSave

Displays the Save File Dialog.

```
FUNCTION ShowSave (BYVAL hParent AS HWND = NULL, BYVAL xPos AS LONG = 0, BYVAL yPos AS LONG = 0, _
   BYVAL bHook AS BOOLEAN = FALSE) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *hParent* | A handle to the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner. |
| *xPos*/*yPos* | Coordinates to position the dialog relative to the client area of the parent. If both *xPos* and *yPos* are -1, the dialog is centered. They are ignored if a hook is not used. |
| *bHook* | Boolean True or False. If True, the dialog is positioned according the specified coordinates and uses the old-style user interface. |

#### Return value

The path of the file to be saved.

#### Usage example:

```
DIM pSFN AS SaveFileDialog
pSFN.FileName = "*.*"
pSFN.InitialDir = CURDIR
pSFN.Filter = "BAS files (*.BAS)|*.BAS|" & "All Files (*.*)|*.*|"
pSFN.DefaultExt = "BAS"
pSFN.FilterIndex = 1
pSFN.Flags = OFN_HIDEREADONLY OR OFN_OVERWRITEPROMPT OR OFN_CREATEPROMPT
DIM dwsFile AS DWSTRING = pSFN.ShowSave(hParent)
```
To position the dialog box use:
```
DIM dwsFile AS DWSTRING = pSFN.ShowSave(0, 20, 20, FALSE)
```
To center the dialog use:
```
DIM dwsFile AS DWSTRING = pSFN.ShowSave(0, -1, -1, TRUE)
```
---
