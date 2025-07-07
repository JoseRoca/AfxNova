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

---
