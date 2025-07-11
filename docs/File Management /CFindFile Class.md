# CFindFile Class

Wraps the **FindFirstFileW**, **FindNext** and **FindClose** API functions to perform file searches and retrieve information about the files found, offering an alternative to Free Basic's **DIR** function that does not work with unicode.

The **FindFile** method opens a search handle and returns information about the first file that the file system finds with a name that matches the specified pattern. This may or may not be the first file or directory that appears in a directory-listing application (such as the dir command) when given the same file name string pattern. This is because **FindFirstFileW** does no sorting of the search results. When the search is finished, call the **FindClose** method.

The following list identifies some other search characteristics:

* The search is performed strictly on the name of the file, not on any attributes such as a date or a file type.
* The search includes the long and short file names.
* An attempt to open a search with a trailing backslash always fails.
* Passing an invalid string, NULL, or empty string for the *wszFileSpec* parameter is not a valid use of this function. Results in this case are undefined.

You cannot use a trailing backslash \ in the *wszFileSpec* parameter for **FindFile**, therefore it may not be obvious how to search root directories.

* To examine files in a root directory, you can use "C:\*". Note: Prepending the string "\\?\" does not allow access to the root directory.
* On network shares, you can use an *wszFileSpec* in the form of the following: "\\Server\Share\*". However, you cannot use an *wszFileSpec* that points to the share itself; for example, "\\Server\Share" is not valid.
* To examine a directory that is not a root directory, use the path to that directory, without a trailing backslash. For example, an argument of "C:\Windows" returns information about the directory "C:\Windows", not about a directory or file in "C:\Windows". To examine the files and directories in "C:\Windows", use an *wszFileSpec* of "C:\Windows*".

**Include file**: AfxNova/CFindFile.inc

#### Usage example:

```
#include once "AfxNova/CFindFile.inc"
using Afx

DIM pFinder AS CFindFile
DIM wszPath AS WSTRING * 260 = "C:\Programs\WinFBX\AfxNova\*.bi"
IF pFinder.FindFile(wszPath) = S_OK THEN
   DO
      PRINT pFinder.FilePath
      IF pFinder.FindNext = 0 THEN EXIT DO
   LOOP
END IF
pFinder.Close
```

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Close](#close) | Ends the search, resets the context and releases all resources. |
| [CreationFileTime](#creationfiletime) | Returns the time the file was created as a FILETIME. |
| [CreationSerialTime](#creationserialtime) | Returns the time the file was created as a Date Serial. |
| [FileAttributes](#fileattributes) | Returns the attributes of the last found file. |
| [FileExt](#fileext) | Returns the extesion of the found file. |
| [FileName](#filename) | Returns the name of the found file. |
| [FileNameX](#filenamex) | Returns the name and extension of the found file. |
| [FilePath](#filepath) | Returns the full path of the found file. |
| [FileSize](#filesize) | Returns the size of the found file, in bytes. |
| [FileURL](#fileuRL) | Returns the file URL. |
| [FindFile](#findfile) | Opens a file search. |
| [FindNext](#findnext) | Searches the next file. |
| [IsCompressedFile](#iscompressedfile) | Checks if the found file is a compressed file. |
| [IsDots](#isdots) | Call this method to test for the current directory and parent directory markers while iterating through files. |
| [IsEncryptedFile](#isencryptedfile) | Checks if the found file is an encrypted file. |
| [IsFile](#isfile) | Checks if the found file is a file and not a folder. |
| [IsFolder](#isfolder) | Checks if the found file is a folder. |
| [IsHiddenFile](#ishiddenfile) | Checks if the found file is a hidden file. |
| [IsNormalFile](#IsNormalFile) | Checks if the found file is a normal file. |
| [IsNotContentIndexedFile](#isnotcontentindexedfile) | Checks if the found file is not to be indexed by the content indexing service.. |
| [IsOffLineFile](#isoffLinefile) | Checks if the found file is not available immediately. |
| [IsReadOnlyFile](#isreadonlyfile) | Checks if the found file is a read only file. |
| [IsReparsePointFile](#isreparsepointfile) | Checks if the found file is is file or directory that has an associated reparse point, or a file that is a symbolic link. |
| [IsSparseFile](#issparsefile) | Checks if the found file is a sparse file. |
| [IsSystemFile](#issystemfile) | Checks if the found file is a system file. |
| [IsTemporaryFile](#istemporaryfile) | Checks if the found file is a temporary file. |
| [LastAccessSerialTime](#lastaccessserialtime) | Returns the time the file was last accessed as a Date Serial. |
| [LastAccessFileTime](#lastaccessfiletime) | Returns the time the file was last accessed as a FILETIME. |
| [LastWriteFileTime](#lastwritefiletime) | Returns the time the file was written to, truncated, or overwritten, as a FILETIME. |
| [LastWriteSerialTime](#lastwriteserialtime) | Returns the time the file was written to, truncated, or overwritten, as a Date Serial. |
| [MatchesMask](#matchesmask) | Tests the file attributes on the found file. |
| [Root](#root) | Returns the root of the found file. |
| [ShortFileName](#shortfilename) | Returns an alternative name for the file. This name is in the classic 8.3 file name format. |

---

## Close

Call this method to end the search, reset the context, and release all resources. After calling **Close**, you do not have to create a new **CFindFile** instance before calling **FindFile** to begin a new search.

```
SUB Close
```
---

## CreationFileTime

Call this method to get the time, in local time format, the file was created.

```
FUNCTION CreationFileTime () AS FILETIME
```
---

## CreationSerialTime

Returns the file's creation date and time as Date Serial. The result can be formated using FreeBasic's Format function.

```
FUNCTION CreationSerialTime () AS DOUBLE
```

#### Usage example:
```
#INCLUDE ONCE "AfxNova/CFindFile.inc"
DIM pFinder AS CFindFile
DIM wszPath AS WSTRING * 260 = "C:\Programs\WinFBX\AfxNova\*.bi"
IF pFinder.FindFile(wszPath) = S_OK THEN
   DO
      PRINT pFinder.FilePath
      PRINT Format(pFinder.CreationSerialTime, "yyyy-mm-dd hh:mm AM/PM")
      IF pFinder.FindNext = 0 THEN EXIT DO
   LOOP
END IF
pFinder.Close
```
---

## FileAttributes

Call this method to get the time, in local time format, the file was created.

```
FUNCTION FileAttributes () AS DWORD
```

#### File Attribute Constants

File attributes are metadata values stored by the file system on disk and are used by the system and are available to developers via various file I/O APIs.

| Attribute  | Description |
| ---------- | ----------- |
| FILE_ATTRIBUTE_ARCHIVE | A file or directory that is an archive file or directory. Applications typically use this attribute to mark files for backup or removal. |
| FILE_ATTRIBUTE_DEVICE | This value is reserved for system use. |
| FILE_ATTRIBUTE_DIRECTORY | The handle that identifies a directory. |
| FILE_ATTRIBUTE_ENCRYPTED | A file or directory that is encrypted. For a file, all data streams in the file are encrypted. For a directory, encryption is the default for newly created files and subdirectories. |
| FILE_ATTRIBUTE_HIDDEN | The file or directory is hidden. It is not included in an ordinary directory listing. |
| FILE_ATTRIBUTE_INTEGRITY_STREAM | The directory or user data stream is configured with integrity (only supported on ReFS volumes). It is not included in an ordinary directory listing. The integrity setting persists with the file if it's renamed. If a file is copied the destination file will have integrity set if either the source file or destination directory have integrity set. Windows Server 2008 R2, Windows 7, Windows Server 2008, Windows Vista, Windows Server 2003, and Windows XP: This flag is not supported until Windows Server 2012. |
| FILE_ATTRIBUTE_NORMAL | A file that does not have other attributes set. This attribute is valid only when used alone. |
| FILE_ATTRIBUTE_NOT_CONTENT_INDEXED | The file or directory is not to be indexed by the content indexing service. |
| FILE_ATTRIBUTE_NO_SCRUB_DATA | The user data stream not to be read by the background data integrity scanner (AKA scrubber). When set on a directory it only provides inheritance. This flag is only supported on Storage Spaces and ReFS volumes. It is not included in an ordinary directory listing. Windows Server 2008 R2, Windows 7, Windows Server 2008, Windows Vista, Windows Server 2003, and Windows XP: This flag is not supported until Windows 8 and Windows Server 2012. |
| FILE_ATTRIBUTE_OFFLINE | The data of a file is not available immediately. This attribute indicates that the file data is physically moved to offline storage. This attribute is used by Remote Storage, which is the hierarchical storage management software. Applications should not arbitrarily change this attribute. |
| FILE_ATTRIBUTE_READONLY | A file that is read-only. Applications can read the file, but cannot write to it or delete it. This attribute is not honored on directories. For more information, see You cannot view or change the Read-only or the System attributes of folders in Windows Server 2003, in Windows XP, in Windows Vista or in Windows 7. |
| FILE_ATTRIBUTE_REPARSE_POINT | A file or directory that has an associated reparse point, or a file that is a symbolic link. |
| FILE_ATTRIBUTE_SPARSE_FILE | A file that is a sparse file. |
| FILE_ATTRIBUTE_SYSTEM | A file or directory that the operating system uses a part of, or uses exclusively. |
| FILE_ATTRIBUTE_TEMPORARY | A file that is being used for temporary storage. File systems avoid writing data back to mass storage if sufficient cache memory is available, because typically, an application deletes a temporary file after the handle is closed. In that scenario, the system can entirely avoid writing the data. Otherwise, the data is written after the handle is closed. |
| FILE_ATTRIBUTE_VIRTUAL | This value is reserved for system use. |

#### Remarks

The FILE_ATTRIBUTE_SPARSE_FILE attribute on the file is set if any of the streams of the file have ever been sparse.

---

## FileExt

Returns the extension of the most-recently-found file.

```
FUNCTION FileExt () AS DWSTRING
```
---

## FileName

Returns the name of the most-recently-found file, excluding the extension.

```
FUNCTION FileName () AS DWSTRING
```
---

## FileNameX

Returns the name of the most-recently-found file, including the extension.

```
FUNCTION FileNameX () AS DWSTRING
```
---

## FilePath

Call this method to get the entire path of the found file.

```
FUNCTION FilePath () AS DWSTRING
```
---

## FileSize

Call this method to get the size of the found file, in bytes.

```
FUNCTION FileSize  () AS ULONGLONG
```
---

## FileURL

Call this member function to retrieve the URL of the file.

```
FUNCTION FileURL () AS DWSTRING
```

#### Remark

**FileURL** is similar to the member function **FilePath**, except that it returns the URL in the form file:\//path. For example, calling FileURL to get the complete URL for myfile.txt returns the URL file:\//c:\myhtml\myfile.txt.

---

## FindFile

Call this function to open a file search.

```
FUNCTION FindFile (BYREF wszFileSpec AS WSTRING) AS LONG_PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFileSpec* | A string containing the name of the file to find (must not end in a trailing backslash). If you pass an empty string, FindFile does a wildcard (*.*) search.<br>If the string ends with a wildcard, period (.), or directory name, the user must have access permissions to the root and all subdirectories on the path.<br>To extend the MAX_PATH limit to 32,767 wide characters, prepend "\\?\" to the path. |

#### Return value

S_OK on success or an error code on failure. To get extended error information, call **GetLastError**. If the function fails because no matching files can be found, the GetLastError function returns ERROR_FILE_NOT_FOUND.

---

# FindNext

Call this method to continue a file search from a previous call to FindFile.

```
FUNCTION FindNext () AS LONG
```

#### Return value

Nonzero if there are more files; zero if the file found is the last one in the directory or if an error occurred. To get extended error information, call **GetLastError**. If the file found is the last file in the directory, or if no matching files can be  found, the **GetLastError** function returns ERROR_NO_MORE_FILES.

---

## IsCompressedFile

Retuns TRUE if the found file is a compressed file. Otherwise FALSE.

```
FUNCTION IsCompressedFile () AS BOOLEAN
```
---

## IsDots

Returns TRUE if the found file has the name "." or "..", which indicates that the found file is actually a directory. Otherwise FALSE.

```
FUNCTION IsDots () AS BOOLEAN
```
---

## IsEncryptedFile

Retuns TRUE if the found file is an encrypted file. Otherwise FALSE.

```
FUNCTION IsEncryptedFile () AS BOOLEAN
```
---

## IsFile 

Retuns TRUE if the found file is a file and not a folder. Otherwise FALSE.

```
FUNCTION IsFile () AS BOOLEAN
```
---

## IsFolder 

Retuns TRUE if the found file is a folder. Otherwise FALSE.

```
FUNCTION IsFolder () AS BOOLEAN
```
---

## IsHiddenFile 

Retuns TRUE if the found file is a hidden file. Otherwise FALSE.

```
FUNCTION IsHiddenFile  () AS BOOLEAN
```
---

## IsNormalFile 

Retuns TRUE if the found file is a normal file. Otherwise FALSE.

```
FUNCTION IsNormalFile () AS BOOLEAN
```
---

## IsNotContentIndexedFile 

Retuns TRUE if the found file is not to be indexed by the content indexing service.. Otherwise FALSE.

```
FUNCTION IsNotContentIndexedFile () AS BOOLEAN
```
---

## IsOffLineFile 

Retuns TRUE if the found file is not available immediately. Otherwise FALSE.

```
FUNCTION IsOffLineFile () AS BOOLEAN
```
---

## IsReadOnlyFile 

Retuns TRUE if the found file is a read only file. Otherwise FALSE.

```
FUNCTION IsReadOnlyFile () AS BOOLEAN
```
---

## IsReparsePointFile 

Retuns TRUE if the found file is a file or directory that has an associated reparse point, or a file that is a symbolic link. Otherwise FALSE.

```
FUNCTION IsReparsePointFile () AS BOOLEAN
```
---

## IsSparseFile 

Retuns TRUE if the found file is a sparse file. Otherwise FALSE.

```
FUNCTION IsSparseFile () AS BOOLEAN
```
---

## IsSystemFile 

Retuns TRUE if the found file is a system file. Otherwise FALSE.

```
FUNCTION IsSystemFile () AS BOOLEAN
```
---

## IsTemporaryFile 

Retuns TRUE if the found file is a temporary file. Otherwise FALSE.

```
FUNCTION IsTemporaryFile () AS BOOLEAN
```
---

## LastAccessFileTime 

Returns the time the file was last accessed, in local time format, as a FILETIME.

```
FUNCTION LastAccessFileTime () AS FILETIME
```
---

## LastAccessSerialTime 

Returns the time the file was last accessed, in local time format, as a Date Serial.

```
FUNCTION LastAccessSerialTime () AS DOUBLE
```

### Usage example
```
#INCLUDE ONCE "AfxNova/CFindFile.inc"
DIM pFinder AS CFindFile
DIM wszPath AS WSTRING * 260 = "C:\Programs\WinFBX\AfxNova\*.bi"
IF pFinder.FindFile(wszPath) = S_OK THEN
   DO
      PRINT pFinder.FilePath
      PRINT Format(pFinder.LastAccessSerialTime, "yyyy-mm-dd hh:mm AM/PM")
      IF pFinder.FindNext = 0 THEN EXIT DO
   LOOP
END IF
pFinder.Close
```
---

## LastWriteFileTime 

Returns the time the file was written to, truncated, or overwritten, in local time format, as a FILETIME.

```
FUNCTION LastWriteFileTime () AS FILETIME
```
---

## LastWriteSerialTime 

Returns the time the file was written to, truncated, or overwritten, in local time format, as a Date Serial.

```
FUNCTION LastWriteSerialTime () AS DOUBLE
```

### Usage example
```
#INCLUDE ONCE "AfxNova/CFindFile.inc"
DIM pFinder AS CFindFile
DIM wszPath AS WSTRING * 260 = "C:\Programs\WinFBX\AfxNova\*.bi"
IF pFinder.FindFile(wszPath) = S_OK THEN
   DO
      PRINT pFinder.FilePath
      PRINT Format(pFinder.LastWriteSerialTime, "yyyy-mm-dd hh:mm AM/PM")
      IF pFinder.FindNext = 0 THEN EXIT DO
   LOOP
END IF
pFinder.Close
```
---

## MatchesMask 

Call this method to test the file attributes on the found file.

```
FUNCTION MatchesMask (BYVAL dwMask AS DWORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwMask* | Specifies one or more file attributes, identified in the WIN32_FIND_DATAW structure, for the found file. To search for multiple attributes, use the bitwise OR operator. |

Any combination of the following attributes is acceptable:

| Attribute  | Description |
| ---------- | ----------- |
| FILE_ATTRIBUTE_ARCHIVE | The file is an archive file. Applications use this attribute to mark files for backup or removal. |
| FILE_ATTRIBUTE_COMPRESSED | The file or directory is compressed. For a file, this means that all of the data in the file is compressed. For a directory, this means that compression is the default for newly created files and subdirectories. |
| FILE_ATTRIBUTE_DIRECTORY | The file is a directory. |
| FILE_ATTRIBUTE_NORMAL | The file has no other attributes set. This attribute is valid only if used alone. All other file attributes override this attribute. |
| FILE_ATTRIBUTE_HIDDEN | The file is hidden. It is not to be included in an ordinary directory listing. |
| FILE_ATTRIBUTE_READONLY | The file is read only. Applications can read the file but cannot write to it or delete it. |
| FILE_ATTRIBUTE_SYSTEM | The file is part of or is used exclusively by the operating system. |
| FILE_ATTRIBUTE_TEMPORARY | The file is being used for temporary storage. Applications should write to the file only if absolutely necessary. Most of the file's data remains in memory without being flushed to the media because the file will soon be deleted. |

#### Return value

TRUE if successful; otherwise FALSE.

---

## Root 

Call this method to get the root of the found file.

```
FUNCTION Root () AS DWSTRING
```

#### Return value

This method returns the drive specifier and path name used to start a search. For example, calling **FindFile** with \*.dat results in **Root** returning an empty string. Passing a path, such as c:\windows\system\*.dll, to FindFile results **Root** returning c:\windows\system\.

---

## ShortFileName 

Call this method to get an alternative name for the file. This name is in the classic 8.3 file name format.

```
FUNCTION ShortFileName () AS DWSTRING
```

#### Return value

The alternate name of the most-recently-found file, including the extension.

---
