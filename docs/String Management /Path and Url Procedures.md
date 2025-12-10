# Path and Url procedures

Assorted path and url procedures.

**Include file**: AfxPath.inc

---

## Path procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxPathAddBackSlash](#afxpathaddbackslash) | Adds a backslash to the end of a string to create the correct syntax for a path. If the source path already has a trailing backslash, no backslash will be added. |
| [AfxPathAddExtension](#afxpathaddextension) | Adds a file name extension to a path string. |
| [AfxPathAllocCanonicalize](#afxpathalloccanonicalize) | Converts a path string into a canonical form. |
| [AfxPathAppend](#afxpathappend) | Appends one path to the end of another. |
| [AfxPathBuildRoot](#afxpathbuildroot) | Creates a root path from a given drive number. |
| [AfxPathCanonicalize](#afxpathcanonicalize) | Removes elements of a file path according to special strings inserted into that path. |
| [AfxPathCombine](#afxpathcombine) | Concatenates two strings that represent properly formed paths into one path; also concatenates any relative path elements. |
| [AfxPathCommonPrefix](#afxpathcommonprefix) | Compares two paths to determine if they share a common prefix. A prefix is one of these types: "C:\\", ".", "..", "..\\". |
| [AfxPathCompactPath](#afxpathcompactpath) | Truncates a file path to fit within a given pixel width by replacing path components with ellipses. |
| [AfxPathCompactPathEx](#afxpathcompactpathex) | Truncates a path to fit within a certain number of characters by replacing path components with ellipses. |
| [AfxPathCreateFromUrl](#afxpathcreatefromurl) | Converts a file URL to a Microsoft MS-DOS path. |
| [AfxPathFileExists](#afxpathfileexists) | Determines whether a path to a file system object such as a file or directory is valid. |
| [AfxPathFindExtension](#afxpathfindextension) | Searches a path for an extension. |
| [AfxPathFindFileName](#afxpathfindfilename) | Searches a path for a file name. |
| [AfxPathFindNextComponent](#afxpathfindnextcomponent) | Parses a path and returns the portion of that path that follows the first backslash. |
| [AfxPathFindOnPath](#afxpathfindonpath) | Searches for a file. |
| [AfxPathFindSuffixArray](#afxpathfindsuffixarray) | Determines whether a given file name has one of a list of suffixes. |
| [AfxPathGetArgs](#afxpathgetargs) | Finds the command line arguments within a given path. |
| [AfxPathGetCharType](#afxpathgetchartype) | Determines the type of character in relation to a path. |
| [AfxPathGetDriveNumber](#afxpathgetdrivenumber) | Searches a path for a drive letter within the range of 'A' to 'Z' and returns the corresponding drive number. |
| [AfxPathIsContentType](#afxpathiscontenttype) | Determines if a file's registered content type matches the specified content type. |
| [AfxPathIsDirectory](#afxpathisdirectory) | Verifies that a path is a valid directory. |
| [AfxPathIsDirectoryEmpty](#afxpathisdirectoryempty) | Determines whether a specified path is an empty directory. |
| [AfxPathIsFileSpec](#afxpathisfilespec) | Searches a path for any path-delimiting characters (for example, ':' or '\' ). If there are no path-delimiting characters present, the path is considered to be a File Spec path. |
| [AfxPathIsHTMLFile](#afxpathishtmlfile) | Determines if a file is an HTML file. The determination is made based on the content type that is registered for the file's extension. |
| [AfxPathIsLFNFileSpec](#afxpathislfnfileSpec) | Determines whether a file name is in long format. |
| [AfxPathIsNetworkPath](#afxpathisnetworkpath) | Determines whether a path string represents a network resource. |
| [AfxPathIsPrefix](#afxpathisprefix) | Searches a path to determine if it contains a valid prefix of the type passed by wszPrefix. A prefix is one of these types: "C:\\", ".", "..", "..\\". |
| [AfxPathIsRelative](#afxpathisrelative) | Searches a path and determines if it is relative. |
| [AfxPathIsRoot](#afxpathisroot) | Parses a path to determine if it is a directory root. |
| [AfxPathIsSameRoot](#afxpathissameroot) | Compares two paths to determine if they have a common root component. |
| [AfxPathIsSystemFolder](#afxpathissystemfolder) | Determines if an existing folder contains the attributes that make it a system folder. Alternately, this function indicates if certain attributes qualify a folder to be a system folder. |
| [AfxPathIsUNC](#afxpathisunc) | Determines if the string is a valid Universal Naming Convention (UNC) for a server and share path. |
| [AfxPathIsUNCServer](#afxpathisuncserver) | Determines if a string is a valid Universal Naming Convention (UNC) for a server path only. |
| [AfxPathIsUNCServerShare](#afxpathisuncservershare) | Determines if a string is a valid Universal Naming Convention (UNC) share path, \\\\*server*\\*share*. |
| [AfxPathIsURL](#afxpathisurl) | Tests a given string to determine if it conforms to a valid URL format. |
| [AfxPathMakePretty](#afxpathmakepretty) | Converts a path to all lowercase characters to give the path a consistent appearance. |
| [AfxPathMakeSystemFolder](#afxpathmakesystemfolder) | Gives an existing folder the proper attributes to become a system folder. |
| [AfxPathMatchSpec](#afxpathmatchspec) | Searches a string using a Microsoft MS-DOS wild card match type. |
| [AfxPathMatchSpecEx](#afxpathmatchspecex) | Searches a path to determine whether it contains a file of a specified file type extension. |
| [AfxPathParseIconLocation](#afxpathparseiconlocation) | Parses a file location string that contains a file location and icon index, and returns separate values. |
| [AfxPathQuoteSpaces](#afxpathquotespaces) | Searches a path for spaces. If spaces are found, the entire path is enclosed in quotation marks. |
| [AfxPathRelativePathTo](#afxpathrelativepathto) | Creates a relative path from one file or folder to another. |
| [AfxPathRemoveArgs](#afxpathremoveargs) | Removes any arguments from a given path. |
| [AfxPathRemoveBackslash](#afxpathremovebackslash) | Removes the trailing backslash from a given path. |
| [AfxPathRemoveBlanks](#afxpathremoveblanks) | Removes all leading and trailing spaces from a string. |
| [AfxPathRemoveExtension](#afxpathremoveextension) | Removes the file name extension from a path, if one is present. |
| [AfxPathRemoveFileSpec](#afxpathremovefilespec) | Removes the trailing file name and backslash from a path, if they are present. |
| [AfxPathRenameExtension](#afxpathrenameextension) | Replaces the extension of a file name with a new extension. If the file name does not contain an extension, the extension will be attached to the end of the string. |
| [AfxPathSearchAndQualify](#afxpathsearchandqualify) | Determines if a given path is correctly formatted and fully qualified. |
| [AfxPathSetDlgItemPath](#afxpathsetdlgitempath) | Sets the text of a child control in a window or dialog box, using **AfxCompactPath** to ensure the path fits in the control. |
| [AfxPathSkipRoot](#afxpathskiproot) | Parses a path, ignoring the drive letter or Universal Naming Convention (UNC) server/share path elements. |
| [AfxPathStripPath](#afxpathstrippath) | Removes the path portion of a fully qualified path and file. |
| [AfxPathStripToRoot](#afxpathstriptoroot) | Removes all parts of the path except for the root information. |
| [AfxPathUndecorate](#afxpathundecorate) | Removes the decoration from a path string. |
| [AfxPathUnExpandEnvStrings](#afxpathunexpandenvstrings) | Replaces certain folder names in a fully-qualified path with their associated environment string. |
| [AfxPathUnmakeSystemFolder](#afxpathunmakesystemfolder) | Removes the attributes from a folder that make it a system folder. This folder must actually exist in the file system. |
| [AfxPathUnquoteSpaces](#afxpathunquotespaces) | Removes quotes from the beginning and end of a path. |

---

## Url procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxIsValidUrl](#afxisvalidurl) | Determines if a specified string is a valid URL. |
| [AfxUrlApplyScheme](#afxurlapplyscheme) | Determines a scheme for a specified URL string, and returns a string with an appropriate prefix. |
| [AfxUrlCanonicalize](#afxurlcanonicalize) | Converts a URL string into canonical form. |
| [AfxUrlCombine](#afxurlcombine) | When provided with a relative URL and its base, returns a URL in canonical form. |
| [AfxUrlCompare](#afxurlcompare) | Makes a case-sensitive comparison of two URL strings. |
| [AfxUrlCreateFromPath](#afxurlcreatefrompath) | Converts a Microsoft MS-DOS path to a canonicalized URL. |
| [AfxUrlEscape](#afxurlescape) | Converts characters in a URL that might be altered during transport across the Internet ("unsafe" characters) into their corresponding escape sequences. |
| [AfxUrlEscapeSpaces](#afxurlescapespaces) | Converts space characters into their corresponding escape sequence. |
| [AfxUrlFixup](#afxurlfixup) | Attempts to correct a URL whose protocol identifier is incorrect. For example, htttp will be changed to http. |
| [AfxUrlGetHostName](#afxurlgethostname) | Accepts a URL string and returns the host name. |
| [AfxUrlGetLocation](#afxurlgetlocation) | Retrieves the location from a URL. |
| [AfxUrlGetPart](#afxurlgetoart) | Accepts a URL string and returns a specified part of that URL. |
| [AfxUrlGetPassword](#afxurlgetpassword) | Accepts a URL string and returns the password. |
| [AfxUrlGetPort](#afxurlgetport) | Accepts a URL string and returns the port number. |
| [AfxUrlGetQuery](#afxurlgetquery) | Accepts a URL string and returns the query portion. |
| [AfxUrlGetScheme](#afxurlgetscheme) | Accepts a URL string and returns the scheme. |
| [AfxUrlGetUserName](#afxurlgetusername) | Accepts a URL string and returns the user name. |
| [AfxUrlHash](#afxurlhash) | Hashes a URL string. |
| [AfxUrlIs](#afxurlis) | Tests whether or not a URL is a specified type. |
| [AfxUrlIsFileUrl](#afxurlisfileurl) | Tests a URL to determine if it is a file URL. |
| [AfxUrlIsNoHistory](#afxurlisnohistory) | Returns whether a URL is a URL that browsers typically do not include in navigation history. |
| [AfxUrlIsOpaque](#afxurlissopaque) | Returns whether a URL is opaque. |
| [AfxUrlUnescape](#afxurlunescape) | Converts escape sequences back into ordinary characters. |
| [AfxUrlUnescapeInPlace](#afxurlunescapeinplace) | Converts escape sequences back into ordinary characters and overwrites the original string. |

---

## AfxPathAddBackSlash

Adds a backslash to the end of a string to create the correct syntax for a path. If the source path already has a trailing backslash, no backslash will be added.

```
FUNCTION AfxPathAddBackSlash (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that represents a path. |

#### Return value

The changed path.

#### Usage example

```
DIM dws AS DWSTRING = AfxPathAddBackSlash("C:\dir_name\dir_name\file_name")
```
---

## AfxPathAddExtension

Adds a file name extension to a path string.

```
FUNCTION AfxPathAddExtension (BYREF wszPath AS CONST WSTRING, BYVAL pwszExt AS WSTRING PTR = NULL) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that represents a path. |
| *pwszExt* | A string that contains the file name extension. |

#### Return value

The changed path.

#### Remarks

If there is already a file name extension present, no extension will be added. If the *wszPath* is an empty string, the result will be the file name extension only. If *wszExtension* is an empty string, an ".exe" extension will be added.

#### Usage examples

```
DIM dws AS DWSTRING = AfxPathAddExtension("file") ' output: file.exe
DIM dws AS DWSTRING = AfxPathAddExtension("file.doc") ' output: file.doc
DIM dws AS DWSTRING = AfxPathAddExtension("file", ".txt") ' output: file.txt
DIM dws AS DWSTRING = AfxPathAddExtension("", ".txt") ' output: .txt
```
---

### AfxPathAllocCanonicalize

Converts a path string into a canonical form. This function differs from **AfxPathCanonicalize** in that it accepts paths with "\\", "\\?" and "\\?\\UNC" prefixes, as well as one or more flags.

```
FUNCTION AfxPathAllocCanonicalize (BYVAL pwszPathIn AS PCWSTR, BYVAL dwFlags AS ULONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszPathIn* | A string that represents a path. |
| *dwFlags* | One or more of the flags listed below. |

| Value  | Meaning |
| ------ | ------- |
| *PATHCCH_NONE* | Do not allow for the construction of \\\?\\ paths (ie, long paths) longer than MAX_PATH. |
| *PATHCCH_ALLOW_LONG_PATHS* | Allow the building of \\\?\\ paths longer than MAX_PATH. |
| *PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS* | Forces the API to treat the caller as long path enabled, independent of the process's long name enabled state. This option can be used only when *PATHCCH_ALLOW_LONG_PATHS* is specified, and cannot be used with *PATHCCH_FORCE_DISABLE_LONG_NAME_PROCESS*. Note: This value is available starting in Windows 10, version 1703. |
| *PATHCCH_FORCE_DISABLE_LONG_NAME_PROCESS* | Forces the API to treat the caller as long path disabled, independent of the process's long name enabled state. This option can be used only when *PATHCCH_ALLOW_LONG_PATHS* is specified, and cannot be used with *PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS*. Note: This value is available starting in Windows 10, version 1703. |
| *PATHCCH_DO_NOT_NORMALIZE_SEGMENTS* | Disables the normalization of path segments that includes removing trailing dots and spaces. This enables access to paths that win32 path normalization will block. Note: This value is available starting in Windows 10, version 1703. |
| *PATHCCH_ENSURE_IS_EXTENDED_LENGTH_PATH* | Converts the input path into the extended length DOS device path form (with the \\\?\\ prefix) if not already in that form. This enables access to paths that are otherwise not addressable due to Win32 normalization rules (that can strip trailing dots and spaces) and path length limitations. This option implies the same behavior of *PATHCCH_DO_NOT_NORMALIZE_SEGMENTS*. Note: This value is available starting in Windows 10, version 1703. |
| *PATHCCH_ENSURE_TRAILING_SLASH* | When combining or normalizing a path, ensure there is a trailing backslash. Note This value is available starting in Windows 10, version 1703. |
| *PATHCCH_CANONICALIZE_SLASHES* | Convert forward slashes to back slashes and collapse multiple slashes. Note: This value is available starting in Windows 11 (SDK version 10.0.22000.194). |

#### Usage example

```
AfxPathAllocCanonicalize(ExePath, PATHCCH_ENSURE_TRAILING_SLASH)
```
---

## AfxPathAppend

Appends one path to the end of another.

```
FUNCTION AfxPathAppend (BYREF wszPath AS CONST WSTRING, BYREF wszMore AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that represents a path. |
| *wszMore* | A string that contains the path to be appended. |

#### Return value

The changed path.

#### Remarks

This function automatically inserts a backslash between the two strings, if one is not already present.

The path supplied in *wszPath* cannot begin with "..\\" or ".\\" to produce a relative path string. If present, those periods are stripped from the output string. For example, appending "path3" to "..\\path1\\path2" results in an output of "\\path1\\path2\\path3" rather than "..\\path1\\path2\\path3".

#### Usage example

```
DIM dws AS DWSTRING = AfxPathAppend("name_1\name_2", "name_3")
```
---

## AfxPathBuildRoot

Creates a root path from a given drive number.

```
FUNCTION AfxPathBuildRoot (BYVAL nDrive AS LONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *nDrive* | A variable of type LONG that indicates the desired drive number. It should be between 0 and 25. |

#### Return value

The root path. If the call fails for any reason (for example, an invalid drive number), an empty string is returned.

#### Usage example

```
DIM dws AS DWSTRING = AfxPathBuildRoot(2) ' output: C:\
```
---

## AfxPathCanonicalize

Removes elements of a file path according to special strings inserted into that path.

```
FUNCTION AfxPathCanonicalize (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be canonicalized. |

#### Return value

The canonicalized path.

#### Remarks

This function allows the user to specify what to remove from a path by inserting special character sequences into the path. The ".." sequence indicates to remove a path segment from the current position to the previous path segment. The "." sequence indicates to skip over the next path segment to the following path segment. The root segment of the path cannot be removed.

If there are more ".." sequences than there are path segments, the contents of the returned string contains just the root, "\\".

#### Usage example

```
DIM dws AS DWSTRING = AfxPathCanonicalize("A:\name_1\.\name_2\..\sname_3") ' output: A:\name_1\name_3
DIM dws AS DWSTRING = AfxPathCanonicalize("A:\name_1\..\name_2\.\name_3") ' output: A:\name_2\name_3
DIM dws AS DWSTRING = AfxPathCanonicalize("A:\name_1\name_2\.\name_3\..\name_4") ' output: A:\name_1\name_2\name_4
DIM dws AS DWSTRING = AfxPathCanonicalize("A:\name_1\.\name_2\.\name_3\..\name_4\..") ' output: A:\name_1\name_2
DIM dws AS DWSTRING = AfxPathCanonicalize("C:\..") ' output: C:\
```
---

## AfxPathCombine

Concatenates two strings that represent properly formed paths into one path; also concatenates any relative path elements.

```
FUNCTION AfxPathCombine (BYREF wszDir AS CONST WSTRING, BYREF wszFile AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDir* | A string that contains the directory path. This value can be "". |
| *wszFile* | A string that contains the file path. This value can be "". |

#### Return value

The concatenated path if successful, or an empty string otherwise.

#### Remarks

The directory path should be in the form of A:,B:, ..., Z:. The file path should be in a correct form that represents the file part of the path. The file path must not be null, and if it ends with a backslash, the backslash will be maintained.

#### Usage example

```
DIM dws AS DWSTRING = AfxPathCombine("C:", "One\Two\Three") ' output: C:\One\Two\Three
```
---

## AfxPathCommonPrefix

Compares two paths to determine if they share a common prefix. A prefix is one of these types: "C:\\", ".", "..", "..\\".

```
FUNCTION AfxPathCommonPrefix (BYREF wszFile1 AS CONST WSTRING, BYREF wszFile2 AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile1* | A string that contains the first path name. |
| *wszFile2* | A string that contains the second path name. |

#### Return value

The common prefix.

#### Usage example

```
DIM dws AS DWSTRING = AfxPathCommonPrefix("C:\win\desktop\temp.txt", "c:\win\tray\sample.txt")
```
---

## AfxPathCompactPath

Truncates a file path to fit within a given pixel width by replacing path components with ellipses.

```
FUNCTION AfxPathCompactPath (BYVAL hDC AS HDC, BYREF wszPath AS CONST WSTRING, BYVAL dx AS DWORD) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *hDC* | A handle to the device context used for font metrics. This value can be NULL. |
| *wszPath* | A string that contains the path to be modified. |
| *dx* | The width, in pixels, in which the string must fit. |

#### Return value

The compacted path.

#### Remarks

This function uses the font currently selected in *hDC* to calculate the width of the text. This function will not compact the path beyond the base file name preceded by ellipses.

#### Usage example

```
DIM dws AS DWSTRING = AfxPathCompactPath(NULL, "C:\path1\path2\sample.txt", 180)
```
---

## AfxPathCompactPathEx

Truncates a path to fit within a certain number of characters by replacing path components with ellipses.

```
FUNCTION AfxPathCompactPathEx (BYREF wszPath AS CONST WSTRING, BYVAL cchMax AS DWORD) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be altered. |
| *cchMax* | The maximum number of characters to be contained in the new string, including the terminating null character. For example, if *cchMax* = 8, the resulting string can contain a maximum of 7 characters plus the terminating null character. |

#### Return value

The compacted path.

#### Remarks

The '/' separator will be used instead of '\\' if the original string used it. If *wszPath* points to a file name that is too long, instead of a path, the file name will be truncated to *cchMax* characters, including the ellipsis and the terminating NULL character. For example, if the input file name is "My Filename" and *cchMax* is 10, *AfxPathCompactPathEx* will return "My Fil...".

#### Usage example

```
DIM dws AS DWSTRING = AfxPathCompactPathEx("c:\test\My Filename", 18)
```
---

## AfxPathCreateFromUrl

Converts a file URL to a Microsoft MS-DOS path.

```
FUNCTION AfxPathCreateFromUrl (BYREF wszUrl AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. |

#### Return value

The MS-DOS path.

#### Usage example

```
DIM dws AS DWSTRING = AfxPathCreateFromUrl("file:///C:/Documents%20and%20Settings/URIS.txt")
```
---

## AfxPathFileExists

Determines whether a path to a file system object such as a file or directory is valid.

```
FUNCTION AfxPathFileExists (BYREF wszPath AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the full path of the object to verify. |

#### Return value

Returns TRUE if the file exists, or FALSE otherwise.

#### Remarks

This function tests the validity of the path.

A path specified by Universal Naming Convention (UNC) is limited to a file only; that is, \\\\server\\share\file is permitted. A UNC path to a server or server share is not permitted; that is, \\\\server or \\\\server\\share. This function returns FALSE if a mounted remote drive is out of service.

---

## AfxPathFindExtension

Searches a path for an extension.

```
FUNCTION AfxPathFindExtension (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to search, including the extension being searched for. |

#### Return value

The file name extension.

#### Remarks

Note that a valid file name extension cannot contain a space.

#### Usage example

```
DIM dws AS DWSTRING = AfxPathFindExtension("C:\TEST\filetxt")
```
---

## AfxPathFindFileName

Searches a path for a file name.

```
FUNCTION AfxPathFindFileName (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to search. |

#### Return value

The file name.

---

## AfxPathFindNextComponent

Searches a path for a file name.

```
FUNCTION AfxPathFindNextComponent (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to parse. This string must not be longer than MAX_PATH characters, plus the terminating null character. Path components are delimited by backslashes. For instance, the path "c:\\path1\\path2\\file.txt" has four components: c:, path1, path2, and file.txt. |

#### Return value

The truncated path.

#### Remarks

**AfxPathFindNextComponent** walks a path string until it encounters a backslash ("\\"), ignores everything up to that point including the backslash, and returns the rest of the path. Therefore, if a path begins with a backslash (such as \\path1\\path2), the function simply removes the initial backslash and returns the rest (path1\\path2).

#### Usage example

```
DIM dws AS DWSTRING = AfxPathFindNextComponent("C:\TEST\file.txt")
```
---

## AfxPathFindOnPath

Searches for a file.

```
FUNCTION AfxPathFindOnPath (BYREF wszFile AS CONST WSTRING, _
   BYVAL ppwszOtherDirs AS WSTRING PTR PTR = NULL) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | A string that contains the file name for which to search. If the search is successful, this parameter is used to return the fully-qualified path name. |
| *ppwszOtherDirs* | An optional, null-terminated array of directories to be searched first. This value can be NULL. |

#### Return value

The fully-qualified path name on success or an empty string on failure.

#### Remarks

**AfxPathFindOnPath** searches for the file specified by *wszFile*. If no directories are specified in *ppwszOtherDirs*, it attempts to find the file by searching standard directories such as System32 and the directories specified in the PATH environment variable. To expedite the process or enable **AfxPathFindOnPath** to search a wider range of directories, use the *ppwszOtherDirs* parameter to specify one or more directories to be searched first. If more than one file has the name specified by *wszFile*, **AfxPathFindOnPath** returns the first instance it finds.

---

## AfxPathFindSuffixArray

Determines whether a given file name has one of a list of suffixes.

```
FUNCTION AfxPathFindSuffixArray (BYREF wszPath AS WSTRING, BYVAL apwszSuffix AS WSTRING PTR PTR, _
   BYVAL iArraySize AS LONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the file name to be tested. A full path can be used. |
| *apwszSuffix* | An array of *iArraySize* string pointers. Each string pointed to is null-terminated and contains one suffix. The strings can be of variable lengths. |
| *iArraySize* | The number of elements in the array pointed to by *apwszSuffix*. |

#### Return value

The matching suffix if successful, or an empty string if bstrPath does not end with one of the specified suffixes.

#### Remarks

This function uses a case-sensitive comparison. The suffix must match exactly.

---

## AfxPathGetArgs

Finds the command line arguments within a given path.

```
FUNCTION AfxPathGetArgs (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be searched. |

#### Return value

A string that contains the arguments portion of the path if successful or an empty string if there are no arguments in the path.

#### Remarks

This function should not be used on generic command path templates (from users or the registry), but rather should be used only on templates that the application knows to be well formed.

---

## AfxPathGetCharType

Determines the type of character in relation to a path.

```
FUNCTION AfxPathGetCharType (BYREF wszCh AS CONST WSTRING) AS UINT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCh* | The character for which to determine the type. |

#### Return value

Returns one or more of the following values that define the type of character.

| Retrn code | Description |
| ---------- | ----------- |
| GCT_INVALID | The character is not valid in a path. |
| GCT_LFNCHAR | The character is valid in a long file name. |
| GCT_SEPARATOR | The character is a path separator. |
| GCT_SHORTCHAR | The character is valid in a short (8.3) file name. |
| GCT_WILD | The character is a wildcard character. |

---

## AfxPathGetDriveNumber

Searches a path for a drive letter within the range of 'A' to 'Z' and returns the corresponding drive number.

```
FUNCTION AfxPathGetDriveNumber (BYREF wszPath AS CONST WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be searched. |

#### Return value

Returns 0 through 25 (corresponding to 'A' through 'Z') if the path has a drive letter, or -1 otherwise.

#### Usage example

```
DIM n AS LONG = AfxPathGetDriveNumber("C:\TEST\bar.txt") ' output: 2
```
---

## AfxPathIsContentType

Determines if a file's registered content type matches the specified content type. This function obtains the content type for the specified file type and compares that string with the *wszContentType*. The comparison is not case-sensitive.

```
FUNCTION AfxPathIsContentType (BYREF wszPath AS CONST WSTRING, _
   BYREF wszContentType AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the file whose content type will be compared. |
| *wszContentType* | A string that contains the content type string to which the file's registered content type will be compared. |

#### Return value

Returns nonzero if the file's registered content type matches *wszContentType*, or zero otherwise.

#### Usage example

```
DIM b AS BOOLEAN = AfxPathIsContentType("C:\TEST\bar.txt", "text/plain") ' output: true
DIM b AS BOOLEAN = AfxPathIsContentType("C:\TEST\bar.txt", "image/gif") ' output: false
```
---

## AfxPathIsDirectory

Verifies that a path is a valid directory.

```
FUNCTION AfxPathIsDirectory (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to verify. |

#### Return value

Returns TRUE if the path is a valid directory (FILE_ATTRIBUTE_DIRECTORY is set), or FALSE otherwise.

---

## AfxPathIsDirectoryEmpty

Determines whether a specified path is an empty directory.

```
FUNCTION AfxPathIsDirectoryEmpty (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be tested. |

#### Return value

Returns TRUE if wszPath is an empty directory. Returns FALSE if *wszPath* is not a directory, or if it contains at least one file other than "." or "..".

#### Remarks

"C:\" is considered a directory.

---

## AfxPathIsFileSpec

Searches a path for any path-delimiting characters (for example, ':' or '\\' ). If there are no path-delimiting characters present, the path is considered to be a File Spec path.

```
FUNCTION AfxPathIsFileSpec (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be searched. |

#### Return value

Returns TRUE if there are no path-delimiting characters within the path, or FALSE if there are path-delimiting characters.

---

## AfxPathIsHTMLFile

Determines if a file is an HTML file. The determination is made based on the content type that is registered for the file's extension.

```
FUNCTION AfxPathIsHTMLFile (BYREF wszFile AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | A string that contains the path and name of the file. |

#### Return value

Returns True if the file is an HTML file, or False otherwise.

---

## AfxPathIsLFNFileSpec

Determines whether a file name is in long format.

```
FUNCTION AfxPathIsLFNFileSpec (BYREF wszName AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszName* | A string that contains the file name to be tested. |

#### Return value

Returns True if wszName exceeds the number of characters allowed by the 8.3 format, or False otherwise.

---

## AfxPathIsNetworkPath

Determines whether a path string represents a network resource.

```
FUNCTION AfxPathIsNetworkPath (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path. |

#### Return value

Returns True if the string represents a network resource, or False otherwise.

---

## AfxPathIsPrefix

Searches a path to determine if it contains a valid prefix of the type passed by wszPrefix. A prefix is one of these types: "C:\\", ".", "..", "..\\".

```
FUNCTION AfxPathIsPrefix (BYREF wszPrefix AS CONST WSTRING, BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrefix* | A string that contains the prefix for which to search. |
| *wszPath* | A string that contains contains the path to be searched. |

#### Return value

Returns True if the compared path is the full prefix for the path, or False otherwise.

---

## AfxPathIsRelative

Searches a path and determines if it is relative.

```
FUNCTION AfxPathIsRelative (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to search. |

#### Return value

Returns True if the path is relative, or False if it is absolute.

---

## AfxPathIsRoot

Parses a path to determine if it is a directory root.

```
FUNCTION AfxPathIsRoot (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to be validated. |

#### Return value

Returns True if the specified path is a root, or False otherwise.

#### Remarks

Returns True for paths such as "\\", "X:\\" or "\\\\server\\share". Paths such as "..\\path2" or "\\\\server\\" return FALSE. 

---

## AfxPathIsSameRoot

Compares two paths to determine if they have a common root component.

```
FUNCTION AfxPathIsSameRoot (BYREF wszPath1 AS CONST WSTRING, BYREF wszPath2 AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath1* | A string that contains contains the path to be compared. |
| *wszPath2* | A string that contains contains the second path to be compared. |

#### Return value

Returns True if both strings have the same root component, or False otherwise. If *wszPath1* contains only the server and share, this function also returns False.

---

## AfxPathIsSystemFolder

Determines if an existing folder contains the attributes that make it a system folder. Alternately, this function indicates if certain attributes qualify a folder to be a system folder.

```
FUNCTION AfxPathIsSystemFolder (BYREF wszPath AS CONST WSTRING, BYVAL dwAttrb AS DWORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the name of an existing folder. The attributes for this folder will be retrieved and compared with those that define a system folder. If this folder contains the attributes to make it a system folder, the function returns nonzero. If this value is NULL, this function determines if the attributes passed in dwAttrb qualify it to be a system folder. |
| *dwAttrb* | The file attributes to be compared. Used only if *wszPath* is NULL. In that case, the attributes passed in this value are compared with those that qualify a folder as a system folder. If the attributes are sufficient to make this a system folder, this function returns nonzero. These attributes are the attributes that are returned from **GetFileAttributes**. |

#### Return value

Returns True if the *wszPath* or *dwAttrb* represent a system folder, or False otherwise.

---

## AfxPathIsUNC

Determines if the string is a valid Universal Naming Convention (UNC) for a server and share path.

```
FUNCTION AfxPathIsUNC (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to validate. |

#### Return value

Returns True if the string is a valid UNC path, or False otherwise.

---

## AfxPathIsUNCServer

Determines if a string is a valid Universal Naming Convention (UNC) for a server path only.

```
FUNCTION AfxPathIsUNCServer (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to validate. |

#### Return value

Returns True if the string is a valid UNC path for a server only (no share name), or False otherwise.

---

## AfxPathIsUNCServerShare

Determines if a string is a valid Universal Naming Convention (UNC) share path, \\\\server\\share.

```
FUNCTION AfxPathIsUNCServerShare (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains contains the path to be validated. |

#### Return value

Returns True if the string is in the form \\\\server\\share, or False otherwise.

---

## AfxPathIsURL

Tests a given string to determine if it conforms to a valid URL format.

```
FUNCTION AfxPathIsURL (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the URL path to validate. |

#### Return value

Returns True if *wszPath* has a valid URL format, or Fañse otherwise.

#### Remarks

This function does not verify that the path points to an existing site—only that it has a valid URL format.

---

## AfxPathMakePretty

Converts a path to all lowercase characters to give the path a consistent appearance.

```
FUNCTION AfxPathMakePretty (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be converted. |

#### Return value

The changed path.

#### Remarks

This function only operates on paths that are entirely uppercase. For example: C:\WINDOWS will be converted to c:\windows, but c:\Windows will not be changed.

---

## AfxPathMakeSystemFolder

Gives an existing folder the proper attributes to become a system folder.

```
FUNCTION AfxPathMakeSystemFolder (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the name of an existing folder that will be made into a system folder. |

#### Return value

Returns True if successful, or False otherwise.

---

## AfxPathMatchSpec

Searches a string using a Microsoft MS-DOS wild card match type.

```
FUNCTION AfxPathMatchSpec (BYREF wszFile AS CONST WSTRING, BYREF wszSpec AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | A string that contains the path to be searched. |
| *wszSpec* | A string that contains the file type for which to search. For example, to test whether *wszFile* is a .doc file, *wszSpec* should be set to "\*.doc". |

#### Return value

Returns True if the string matches, or False otherwise.

---

## AfxPathMatchSpecEx

Searches a path to determine whether it contains a file of a specified file type extension.

```
FUNCTION AfxPathMatchSpecEx (BYREF wszFile AS CONST WSTRING, _
   BYREF wszSpec AS CONST WSTRING, BYVAL dwFlags AS DWORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | A string that contains the path to be searched. |
| *wszSpec* | A string that contains the file type extension or extensions for which to search. If exactly one extension is specified, set the PMSF_NORMAL flag in *dwFlags*. If more than one extension is specified, separate them with semicolons and set the PMSF_MULTIPLE flag. |
| *dwFlags* | Modifies the search condition. The following are valid flags:<br>*PMSF_NORMAL* (&h00000000): The *wszSpec* parameter points to a single file type extension to be matched.<br>*PMSF_MULTIPLE* (&h00000001): The *wszSpec* parameter points to a semicolon-delimited list of file type extensions to be matched.<br>*PMSF_DONT_STRIP_SPACES* (&h00010000): If PMSF_NORMAL is used, ignore leading spaces in the string pointed to by *wszSpec*. If PMSF_MULTIPLE is used, ignore leading spaces in each file type contained in the string pointed to by *wszSpec*. This flag can be combined with PMSF_NORMAL and PMSF_MULTIPLE. |

#### Return value

Returns one of the following values:

| Return code | Description |
| ----------- | ----------- |
| TRUE | A file type extension specified in *wszSpec* was found in the path pointed to by *wszFile*. |
| FALSE | The path pointed to by *wszFile* does not contain any file type extension specified in *wszSpec*. |

---

## AfxPathParseIconLocation

Parses a file location string that contains a file location and icon index, and returns separate values.

```
FUNCTION AfxPathParseIconLocation (BYREF wszIconFile AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszIconFile* | A string that contains a file location string. It should be in the form "*path,iconindex*". |

#### Return value

Returns the valid icon index value.

#### Remarks

This function is useful for taking a DefaultIcon value retrieved from the registry by **SHGetValue** and separating the icon index from the path.

#### Usage example

```
DIM dwsIconLocation AS DWSTRING = AfxPathParseIconLocation("C:\TEST\sample.txt,3")
```
---

## AfxPathQuoteSpaces

Searches a path for spaces. If spaces are found, the entire path is enclosed in quotation marks.

```
FUNCTION AfxPathQuoteSpaces (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be searched. |

#### Return value

True if spaces were found; otherwise, False.

---

## AfxPathRelativePathTo

Creates a relative path from one file or folder to another.

```
FUNCTION AfxPathRelativePathTo (BYREF wszFrom AS CONST WSTRING, BYVAL dwAttrFrom AS DWORD, _
   BYREF wszTo AS CONST WSTRING, BYVAL dwAttrTo AS DWORD) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFrom* | A string that contains the path that defines the start of the relative path. |
| *dwAttrFrom* | The file attributes of *wszFrom*. If this value contains FILE_ATTRIBUTE_DIRECTORY, *wszFrom* is assumed to be a directory; otherwise, *wszFrom* is assumed to be a file. |
| *wszTo* | A string that contains the path that defines the endpoint of the relative path. |
| *dwAttrTo* | The file attributes of wszTo. If this value contains FILE_ATTRIBUTE_DIRECTORY, wszTo is assumed to be directory; otherwise, *wszTo* is assumed to be a file. |

#### Return value

Returns True if successful, or False otherwise.

#### Remarks

This function takes a pair of paths and generates a relative path from one to the other. The paths do not have to be fully-qualified, but they must have a common prefix, or the function will fail and return False.

For example, let the starting point, *wszFrom*, be "c:\\FolderA\FolderB\\FolderC", and the ending point, *wszTo*, be "c:\\FolderA\\FolderD\\FolderE". **AfxPathRelativePathTo** will return the relative path from *wszFrom* to *wszTo* as: "..\\..\\FolderD\\FolderE". You will get the same result if you set *wszFrom* to "\\FolderA\\FolderB\\FolderC" and *wszTo* to "\\FolderA\\FolderD\\FolderE". On the other hand, "c:\\FolderA\\FolderB" and "a:\\FolderA\\FolderD do not share a common prefix, and the function will fail. Note that "\\\\" is not considered a prefix and is ignored. If you set *wszFrom* to "\\\\FolderA\\FolderB", and *wszTo* to "\\\\FolderC\\FolderD", the function will fail.

---

## AfxPathRemoveArgs

Removes any arguments from a given path.

```
FUNCTION AfxPathRemoveArgs (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to remove arguments. |

#### Return value

The changed path.

#### Remarks

This function should not be used on generic command path templates (from users or the registry), but rather it should be used only on templates that the application knows to be well formed.

---

## AfxPathRemoveBackslash

Removes the trailing backslash from a given path.

```
FUNCTION AfxPathRemoveBackslash (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to remove the backslash. |

#### Return value

The changed path.

---

## AfxPathRemoveBlanks

Removes all leading and trailing spaces from a string.

```
FUNCTION AfxPathRemoveBlanks (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to strip all leading and trailing spaces. |

#### Return value

The changed path.

---

## AfxPathRemoveExtension

Removes the file name extension from a path, if one is present.

```
FUNCTION AfxPathRemoveExtension (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to remove the extension. |

#### Return value

The changed path.

---

## AfxPathRemoveFileSpec

Removes the trailing file name and backslash from a path, if they are present.

```
FUNCTION AfxPathRemoveFileSpec (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path from which to remove the file name. |

#### Return value

The changed path.

---

## AfxPathRenameExtension

Replaces the extension of a file name with a new extension. If the file name does not contain an extension, the extension will be attached to the end of the string.

```
FUNCTION AfxPathRenameExtension (BYREF wszPath AS CONST WSTRING, BYREF wszExt AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path in which to replace the extension. |
| *wszExt* | A string that contains a '.' character followed by the new extension. |

#### Return value

The new path. If the new path and extension would exceed MAX_PATH characters, the path won't be changed.

---

## AfxPathSearchAndQualify

Determines if a given path is correctly formatted and fully qualified.

```
FUNCTION AfxPathSearchAndQualify (BYREF wszPath AS CONST WSTRING, BYREF wszFullyQualifiedPath AS CONST WSTRING, _
   BYVAL cchFullyQualifiedPath AS DWORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to search. |
| *wszFullyQualifiedPath* | A string of length MAX_PATH that contains the path to be referenced. |
| *cchFullyQualifiedPath* | The size of the buffer pointed to by *wszFullyQualifiedPath*, in characters. |

#### Return value

Returns True if the path is qualified, or False otherwise.

---

## AfxPathSetDlgItemPath

Sets the text of a child control in a window or dialog box, using **AfxCompactPath** to ensure the path fits in the control.

```
SUB AfxPathSetDlgItemPath (BYVAL hDlg AS HWND, BYVAL cID AS LONG, BYREF wszPath AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *hDlg* | A handle to the dialog box or window. |
| *cID* | The identifier of the control. |
| *wszPath* | A string that contains the path to set in the control. |

---

## AfxPathSkipRoot

Parses a path, ignoring the drive letter or Universal Naming Convention (UNC) server/share path elements.

```
FUNCTION AfxPathSkipRoot (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to parse. |

#### Return value

The changed path.

---

## AfxPathStripPath

Removes the path portion of a fully qualified path and file.

```
FUNCTION AfxPathStripPath (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path and file name. When this method returns successfully, the string contains only the file name, with the path removed. |

#### Return value

The changed path.

---

## AfxPathStripToRoot

Removes all parts of the path except for the root information.

```
FUNCTION AfxPathStripToRoot (BYREF wszRoot AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszRoot* | A string of length MAX_PATH that contains the path to be converted. |

#### Return value

Returns a string that contains only the root information taken from that path.

---

## AfxPathUndecorate

Removes the decoration from a path string.

```
FUNCTION AfxPathUndecorate (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path. |

#### Return value

The undecorated string.

#### Remarks

A decoration consists of a pair of square brackets with one or more digits in between, inserted immediately after the base name and before the file name extension.

#### Example

The following table illustrates how strings are modified by **AfxPathUndecorate**.

| Initial String | Undecorated String |
| -------------- | ------------------ |
| C:\\Path\\File\[5].txt | C:\\Path\\File.txt |
| C:\\Path\\File\[12] | C:\\Path\\File |
| C:\\Path\\File.txt | C:\\Path\\File.txt |
| C:\\Path\\\[3].txt | C:\\Path\\\[3].txt |

---

## AfxPathUnExpandEnvStrings

Replaces certain folder names in a fully-qualified path with their associated environment string.

```
FUNCTION AfxPathUnExpandEnvStrings (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path to be unexpanded. |

#### Return value

The unexpanded string.

---

## AfxPathUnmakeSystemFolder

Removes the attributes from a folder that make it a system folder. This folder must actually exist in the file system.

```
FUNCTION AfxPathUnmakeSystemFolder (BYREF wszPath AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the name of an existing folder that will have the system folder attributes removed. |

#### Return value

Returns nonzero if successful, or zero otherwise.

---

## AfxPathUnquoteSpaces

Removes quotes from the beginning and end of a path.

```
FUNCTION AfxPathUnquoteSpaces (BYREF wszPath AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the path. |

#### Return value

The unquoted path.

---

## AfxIsValidUrl

Determines if a specified string is a valid URL.

```
FUNCTION AfxIsValidUrl (BYREF wszUrl AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the full URL to check. |

#### Remarks

Return True if the URL is valid; FALSE, otherwise.

---


## AfxUrlApplyScheme

Determines a scheme for a specified URL string, and returns a string with an appropriate prefix.

```
FUNCTION AfxUrlApplyScheme (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains a URL. |
| *dwFlags* | The flags that specify how to determine the scheme. The flags can be combined. |

| Flag       | Description |
| ---------- | ----------- |
| URL_APPLY_DEFAULT | Apply the default scheme if **AfxUrlApplyScheme** can't determine one. The default prefix is stored in the registry but is typically "http". |
| URL_APPLY_GUESSSCHEME | Attempt to determine the scheme by examining *wszUrl*. |
| URL_APPLY_GUESSFILE | Attempt to determine a file URL from *wszUrl*. |
| URL_APPLY_FORCEAPPLY | Force **AfxUrlApplyScheme** to determine a scheme for *wszUrl*. |
| URL_APPLY_FORCEAPPLY | Force **AfxUrlApplyScheme** to determine a scheme for *wszUrl*. |

#### Return value

The changed url.

#### Remarks

If the URL has a valid scheme, the string will not be modified. However, almost any combination of two or more characters followed by a colon will be parsed as a scheme. Valid characters include some common punctuation marks, such as ".". If your input string fits this description, **AfxUrlApplyScheme** may treat it as valid and not apply a scheme. To force the function to apply a scheme to a URL, set the URL_APPLY_FORCEAPPLY and URL_APPLY_DEFAULT flags in *dwFlags*. This combination of flags forces the function to apply a scheme to the URL. Typically, the function will not be able to determine a valid scheme. The second flag guarantees that, if no valid scheme can be determined, the function will apply the default scheme to the URL.

---

## AfxUrlCanonicalize

Converts a URL string into canonical form.

```
FUNCTION AfxUrlCanonicalize (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains a URL. If the string does not refer to a file, it must include a valid scheme such as "http://". |
| *dwFlags* | The flags that specify how the URL is converted to canonical form. The flags can be combined. |

| Flag       | Description |
| ---------- | ----------- |
| URL_UNESCAPE | Un-escape any escape sequences that the URLs contain, with two exceptions. The escape sequences for "?" and "#" are not un-escaped. If one of the URL_ESCAPE_XXX flags is also set, the two URLs are first un-escaped, then combined, then escaped. |
| URL_ESCAPE_UNSAFE | Replace unsafe characters with their escape sequences. Unsafe characters are those characters that may be altered during transport across the Internet, and include the (<, >, ", #, {, }, \|, \\, ^, \[, ], and ') characters. This flag applies to all URLs, including opaque URLs. |
| URL_PLUGGABLE_PROTOCOL | Combine URLs with client-defined pluggable protocols, according to the W3C specification. This flag does not apply to standard protocols such as ftp, http, gopher, and so on. If this flag is set, **AfxUrlCombine** does not simplify URLs, so there is no need to also set URL_DONT_SIMPLIFY. |
| URL_ESCAPE_SPACES_ONLY | Replace only spaces with escape sequences. This flag takes precedence over URL_ESCAPE_UNSAFE, but does not apply to opaque URLs. |
| URL_DONT_SIMPLIFY | Treat "/./" and "/../" in a URL string as literal characters, not as shorthand for navigation. See Remarks for further discussion. |
| URL_NO_META | Defined to be the same as URL_DONT_SIMPLIFY. |
| URL_ESCAPE_PERCENT | Convert any occurrence of "%" to its escape sequence. |
| URL_ESCAPE_AS_UTF8 | Windows 7 and later. Percent-encode all non-ASCII characters as their UTF-8 equivalents. |

#### Return value

The canonicalized url.

#### Remarks

This function performs such tasks as replacing unsafe characters with their escape sequences and collapsing sequences like "..\\...".

If a URL string contains "/../" or "/./", **AfxUrlCanonicalize** normally treats the characters as indicating navigation in the URL hierarchy. The function simplifies the URLs before combining them. For instance "/hello/cruel/../world" is simplified to "/hello/world". If the URL_DONT_SIMPLIFY flag is set in *dwFlags*, the function does not simplify URLs. In this case, "/hello/cruel/../world" is left as it is.

---

## AfxUrlCombine

When provided with a relative URL and its base, returns a URL in canonical form.

```
FUNCTION AfxUrlCombine (BYREF wszBase AS WSTRING, BYREF wszRelative AS WSTRING, BYVAL dwFlags AS DWORD) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszBase* | A string that contains the base url. |
| *wszRelative* | A string that contains the relative url. |
| *dwFlags* | Flags that specify how the URL is converted to canonical form. The flags can be combined. |

| Flag       | Description |
| ---------- | ----------- |
| URL_DONT_SIMPLIFY | Treat "/./" and "/../" in a URL string as literal characters, not as shorthand for navigation. See Remarks for further discussion. |
| URL_ESCAPE_PERCENT | Convert any occurrence of "%" to its escape sequence. |
| URL_ESCAPE_SPACES_ONLY | Replace only spaces with escape sequences. This flag takes precedence over URL_ESCAPE_UNSAFE, but does not apply to opaque URLs. |
| URL_ESCAPE_UNSAFE | Replace unsafe characters with their escape sequences. Unsafe characters are those characters that may be altered during transport across the Internet, and include the (<, >, ", #, {, }, \|, \\, ^, \[, ], and ') characters. This flag applies to all URLs, including opaque URLs. |
| URL_NO_META | Defined to be the same as URL_DONT_SIMPLIFY. |
| URL_PLUGGABLE_PROTOCOL | Combine URLs with client-defined pluggable protocols, according to the W3C specification. This flag does not apply to standard protocols such as ftp, http, gopher, and so on. If this flag is set, **AfxUrlCombine** does not simplify URLs, so there is no need to also set URL_DONT_SIMPLIFY. |
| URL_UNESCAPE | Un-escape any escape sequences that the URLs contain, with two exceptions. The escape sequences for "?" and "#" are not un-escaped. If one of the URL_ESCAPE_XXX flags is also set, the two URLs are first un-escaped, then combined, then escaped. |
| URL_ESCAPE_AS_UTF8 | Windows 7 and later. Percent-encode all non-ASCII characters as their UTF-8 equivalents. |

#### Return value

The canonicalized url.

#### Remarks

Items between slashes are treated as hierarchical identifiers; the last item specifies the document itself. You must enter a slash (/) after the document name to append more items; otherwise, **AfxUrlCombine** changes one document for another.

If a URL string contains '/../' or '/./', **AfxUrlCombine** usually treats the characters as if they indicated navigation in the URL hierarchy. The function simplifies the URLs before combining them. For instance, "/hello/cruel/../world" is simplified to "/hello/world". If the URL_DONT_SIMPLIFY flag is set in dwFlags, the function does not simplify URLs. In this case, "/hello/cruel/../world" is left as it is.

---

## AfxUrlCompare

Makes a case-sensitive comparison of two URL strings.

```
FUNCTION AfxUrlCompare (BYREF wszUrl1 AS CONST WSTRING, BYREF wszUrl2 AS CONST WSTRING, _
   BYVAL fIgnoreSlash AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl1* | A string that contains the first URL. |
| *wszUrl2* | A string that contains the second URL. |
| *fIgnoreSlash* |  A value that is set to TRUE to have **AfxUrlCompare** ignore a trailing '/' character on either or both URLs. |

#### Return value

Returns zero if the two strings are equal. The function will also return zero if *fIgnoreSlash* is set to TRUE and one of the strings has a trailing '\\' character. The function returns a negative integer if the string pointed to by *wszUrl1* is less than the string pointed to by *wszUrl2*. Otherwise, it returns a positive integer.

#### Remarks

For best results, you should first canonicalize the URLs with **AfxUrlCanonicalize**. Then, compare the canonicalized URLs with **AfxUrlCompare**.

---

## AfxUrlCreateFromPath

Converts a Microsoft MS-DOS path to a canonicalized URL.

```
FUNCTION AfxUrlCreateFromPath (BYREF wszPath AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPath* | A string that contains the MS-DOS path. |

#### Return value

The canonicalized URL.

---

## AfxUrlEscape

Converts characters in a URL that might be altered during transport across the Internet ("unsafe" characters) into their corresponding escape sequences.

```
FUNCTION AfxUrlEscape (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the MS-DOS path. |
| *dwFlags* | The flags that indicate which portion of the URL is being provided in *wszURL* and which characters in that string should be converted to their escape sequences. |

| Flag       | Description |
| ---------- | ----------- |
| URL_DONT_ESCAPE_EXTRA_INFO | Used only in conjunction with URL_ESCAPE_SPACES_ONLY to prevent the conversion of characters in the query (the portion of the URL following the first # or ? character encountered in the string). This flag should not be used alone, nor combined with URL_ESCAPE_SEGMENT_ONLY. |
| URL_BROWSER_MODE | Defined to be the same as URL_DONT_ESCAPE_EXTRA_INFO. |
| URL_ESCAPE_SPACES_ONLY | Convert only space characters to their escape sequences, including those space characters in the query portion of the URL. Other unsafe characters are not converted to their escape sequences. This flag assumes that *wszURL* does not contain a full URL. It expects only the portions following the server specification.<br>Combine this flag with URL_DONT_ESCAPE_EXTRA_INFO to prevent the conversion of space characters in the query portion of the URL.<br>This flag cannot be combined with URL_ESCAPE_PERCENT or URL_ESCAPE_SEGMENT_ONLY. |
| URL_ESCAPE_PERCENT | Convert any % character found in the segment section of the URL (that section falling between the server specification and the first # or ? character). By default, the % character is not converted to its escape sequence. Other unsafe characters in the segment are also converted normally.<br>Combining this flag with URL_ESCAPE_SEGMENT_ONLY includes those % characters in the query portion of the URL. However, as the URL_ESCAPE_SEGMENT_ONLY flag causes the entire string to be considered the segment, any # or ? characters are also converted.<br>This flag cannot be combined with URL_ESCAPE_SPACES_ONLY. |
| URL_ESCAPE_SEGMENT_ONLY | Indicates that *wszURL* contains only that section of the URL following the server component but preceding the query. All unsafe characters in the string are converted. If a full URL is provided when this flag is set, all unsafe characters in the entire string are converted, including # and ? characters.<br>Combine this flag with URL_ESCAPE_PERCENT to include that character in the conversion.<br>This flag cannot be combined with URL_ESCAPE_SPACES_ONLY or URL_DONT_ESCAPE_EXTRA_INFO. |
| URL_ESCAPE_AS_UTF8 | Windows 7 and later. Percent-encode all non-ASCII characters as their UTF-8 equivalents. |

#### Return value

The converted URL.

#### Remarks

For the purposes of this document, a typical URL is divided into three sections: the server, the segment, and the query. For example:

http://microsoft.com/test.asp?url=/example/abc.asp?frame=true#fragment

The server portion is "http://microsoft.com/". The trailing forward slash is considered part of the server portion.

The segment portion is any part of the path found following the server portion, but before the first # or ? character, in this case simply "test.asp".

The query portion is the remainder of the path from the first # or ? character (inclusive) to the end. In the example, it is "?url=/example/abc.asp?frame=true#fragment".

Unsafe characters are those characters that might be altered during transport across the Internet. This function converts unsafe characters into their equivalent "%xy" escape sequences. The following table shows unsafe characters and their escape sequences.

| Character  | Escape Sequence |
| ---------- | ----------- |
| ^ | %5E5E |
| & | %26 |
| ` | %60 |
| { | %7B |
| } | %7D |
| \| | %7C |
| ] | %5D |
| \[ | %5B |
| " | %22 |
| < | %3C |
| > | %3E |
| \ | %5C |

Use of the URL_ESCAPE_SEGMENT_ONLY flag also causes the conversion of the # (%23), ? (%3F), and / (%2F) characters.

By default, AfxUrlEscape ignores any text following a # or ? character. The URL_ESCAPE_SEGMENT_ONLY flag overrides this behavior by regarding the entire string as the segment. The URL_ESCAPE_SPACES_ONLY flag overrides this behavior, but only for space characters.

Security Warning: This function should be called from client applications only.

---

## AfxUrlEscapeSpaces

Converts space characters into their corresponding escape sequence.

```
FUNCTION AfxUrlEscapeSpaces (BYREF wszUrl AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

The converted URL.

---

## AfxUrlFixup

Attempts to correct a URL whose protocol identifier is incorrect. For example, htttp will be changed to http.

```
FUNCTION AfxUrlFixup (BYREF wszUrl AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

The converted URL.

#### Remarks

The **AfxUrlFixup** function recognizes the schemes specified by the URL_SCHEME enumeration.

Priority is given to the first character in the protocol identifier section so htp will be converted to http instead of ftp.

**Notes**

Do not use this function for deterministic data transformation. The heuristics used by **AfxUrlFixup** can change from one release to the next. The function should only be used to correct possibly invalid user input.

This function is available through Windows 7 and Windows Server 2008. It might be altered or unavailable in subsequent versions of Windows.

---

## AfxUrlGetHostName

Accepts a URL string and returns the host name.

```
FUNCTION AfxUrlGetHostName (BYREF wszUrl AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | The URL string. |

#### Return value

If the function succeeds, it returns the host name; if it fails, it returns an empty string.

#### Usage example

```
print AfxUrlGetHostName("https://user@example.com:8080/path/to/resource?id=123&lang=es")
```
---

## AfxUrlGetLocation

Retrieves the location from a URL.

```
FUNCTION AfxUrlGetLocation (BYREF wszUrl AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the location. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

A string with the location or an empty string.

#### Remarks

The location is the segment of the URL starting with a ? or # character. If a file URL has a query string, the returned string includes the query string.

---

## AfxUrlGetPart

Accepts a URL string and returns a specified part of that URL.

```
FUNCTION AfxUrlGetPart (BYREF wszUrl AS WSTRING, _
   BYVAL dwPart AS DWORD, BYVAL dwFlags AS DWORD) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |
| *dwPart* | The flags that specify which part of the URL to retrieve. It can have one of the following values.<br>*URL_PART_HOSTNAME*: The host name.<br>*URL_PART_PASSWORD*: The password.<br>*URL_PART_PORT*: The port number.<br>*URL_PART_QUERY*: The query portion of the URL.<br>*URL_PART_SCHEME*: The URL scheme.<br>*URL_PART_USERNAME*: The username. |
| *dwFlags* | A flag that can be set to keep the URL scheme, in addition to the part that is specified by *dwPart*.<br>*URL_PARTFLAG_KEEPSCHEME*: Keep the URL scheme. |

#### Return value

A string with the specified part of that URL.

---

## AfxUrlGetPassword

Accepts a URL string and returns the password.

```
FUNCTION AfxUrlGetPassword (BYREF wszUrl AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | The URL string. |

#### Return value

If the function succeeds, it returns the password; if it fails, it returns an empty string.

#### Usage example

```
print AfxUrlGetPassword("https://user:secret@example.com:8080/path/to/resource?id=123&lang=es")
```
---

## AfxUrlGetPort

Accepts a URL string and returns the port number.

```
FUNCTION AfxUrlGetPort (BYREF wszUrl AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | The URL string. |

#### Return value

If the function succeeds, it returns the port number; if it fails, it returns an empty string.

#### Usage example

```
print AfxUrlGetPassword("https://user:secret@example.com:8080/path/to/resource?id=123&lang=es")
```
---

## AfxUrlGetQuery

Accepts a URL string and returns the query portion.

```
FUNCTION AfxUrlGetQuery (BYREF wszUrl AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | The URL string. |

#### Return value

If the function succeeds, it returns the query portion; if it fails, it returns an empty string.

#### Usage example

```
print AfxUrlGetQuery("https://user@example.com:8080/path/to/resource?id=123&lang=es")
```
---

## AfxUrlGetSchema

Accepts a URL string and returns the scheme.

```
FUNCTION AfxUrlGetScheme (BYREF wszUrl AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | The URL string. |

#### Return value

If the function succeeds, it returns the scheme; if it fails, it returns an empty string.

#### Usage example

```
print AfxUrlGetScheme("https://user@example.com:8080/path/to/resource?id=123&lang=es")
```
---

## AfxUrlGetUserName

Accepts a URL string and returns the user name.

```
FUNCTION AfxUrlGetUserName (BYREF wszUrl AS CONST WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | The URL string. |

#### Return value

If the function succeeds, it returns the user name; if it fails, it returns an empty string.

#### Usage example

```
FUNCTION AfxUrlGetUserName (BYREF wszUrl AS CONST WSTRING) AS DWSTRING
```
---

## AfxUrlHash

Hashes a URL string.

```
FUNCTION AfxUrlHash (BYREF wszUrl AS WSTRING, BYVAL pbHash AS BYTE PTR, _
   BYVAL cbHash AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |
| *pbHash* | A pointer to a buffer that, When this method returns successfully, receives the hashed array. |
| *cbHash* | The number of elements in the array at *pbHash*. It should be no larger than 256. |

#### Return value

If the function succeeds, it returns S_OK (0). Otherwise, it returns an HRESULT error code.

#### Remarks

To hash a URL into a single byte, set cbHash = SIZEOF(BYTE) and pbHash = VARPTR(bHashedValue), where bHashedValue is a one-byte buffer.

To hash a URL into a DWORD, set cbHash = SIZEOF(DWORD) and pbHash = VARPTR(dwHashedValue), where dwHashedValue is a DWORD buffer.

---

## AfxUrlIs

Tests whether or not a URL is a specified type.

```
FUNCTION AfxUrlIs (BYREF wszUrl AS WSTRING, BYVAL nUrlIs AS URLIS) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |
| *nUrlIs* | The type of URL to be tested for. This parameter can take one of the following values:<br>*URLIS_APPLIABLE*: Attempt to determine a valid scheme for the URL.<br>*URLIS_DIRECTORY*: Does the URL string end with a directory?<br>*URLIS_FILEURL*: Is the URL a file URL?<br>*URLIS_HASQUERY*: Does the URL have an appended query string?<br>*URLIS_NOHISTORY*: Is the URL a URL that is not typically tracked in navigation history?<br>*URLIS_OPAQUE*: Is the URL opaque?<br>*URLIS_URL*: Is the URL valid? |

#### Return value

For all but one of the URL types, **AfxUrlIs** returns True if the URL is the specified type, or False if not.

If **AfxUrlIs** is set to URLIS_APPLIABLE, **AfxUrlIs** will attempt to determine the URL scheme. If the function is able to determine a scheme, it returns True, or False otherwise.

---

## AfxUrlIsFileUrl

Tests a URL to determine if it is a file URL.

```
FUNCTION AfxUrlIsFileUrl (BYREF wszUrl AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

Returns True if the URL is a file URL, or False otherwise.

#### Remarks

A file URL has the form "File:// xxx".

---

## AfxUrlIsNoHistory

Returns whether a URL is a URL that browsers typically do not include in navigation history.

```
FUNCTION AfxUrlIsNoHistory (BYREF wszUrl AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

Returns a True value if the URL is a URL that is not included in navigation history, or False otherwise.

#### Remarks

This function is equivalent to the following: AfxUrlIs(wszURL, URLIS_NOHISTORY)

---

## AfxUrlIsOpaque

Returns whether a URL is opaque.

```
FUNCTION AfxUrlIsOpaque (BYVAL wszUrl AS CONST WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL to be corrected. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |

#### Return value

Returns True if the URL is opaque, or False otherwise.

#### Remarks

A URL that has a scheme that is not followed by two slashes (//) is opaque. For example, mailto:xyz@litwareinc.com is an opaque URL. Opaque URLs cannot be separated into the standard URL hierarchy. **AfxUrlIsOpaque** is equivalent to the following: AfxUrlIs(wszURL, URLIS_OPAQUE)

---

## AfxUrlUnescape

Converts escape sequences back into ordinary characters.

```
FUNCTION AfxUrlUnescape (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. |
| *dwFlags* | Flags that control which characters are unescaped. It can be the following flag:<br>*URL_DONT_UNESCAPE_EXTRA_INFO*: Do not convert the # or ? character, or any characters following them in the string. |

#### Return value

Returns S_OK (0) if successful or an HRESULT otherwise.

#### Remarks

An escape sequence has the form "%xy".

Input strings cannot be longer than INTERNET_MAX_URL_LENGTH.

---

## AfxUrlUnescapeInPlace

Converts escape sequences back into ordinary characters and overwrites the original string.

```
FUNCTION AfxUrlUnescapeInPlace (BYREF wszUrl AS WSTRING, BYVAL dwFlags AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszUrl* | A string that contains the URL. This string must not exceed INTERNET_MAX_PATH_LENGTH characters in length, including the terminating NULL character. The converted string is returned through this parameter. |
| *dwFlags* | Flags that control which characters are unescaped. It can be the following flag:<br>*URL_DONT_UNESCAPE_EXTRA_INFO*: Do not convert the # or ? character, or any characters following them in the string. |

#### Return value

Returns S_OK (0) if successful or an HRESULT otherwise.

---
