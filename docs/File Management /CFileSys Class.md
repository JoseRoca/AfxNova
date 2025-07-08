# CFileSys Class

The **CFileSys** class wraps the Microsoft File System Object and provides methods to work with files and folders, giving your application the ability to create, copy, alter, move, and delete files and folders, or to determine if and where particular files or folders exist. It also enables you to get information about files and folders, such as their names and the date they were created or last modified.

**Include file**: AfxNova/CFileSys.inc

### Methods

| Name       | Description |
| ---------- | ----------- |
| [BuildPath](#buildpath) | Appends a name to an existing path. |
| [CopyFile](#copyfile) | Copies one or more files from one location to another. |
| [CopyFolder](#copyfolder) | Recursively copies a folder from one location to another. |
| [CreateFolder](#createfolder) | Creates a folder. |
| [DateTimeToString](#datetimetostring) | Converts a DATE_ type to a string containing the date and the time. |
| [DeleteFile](#deletefile) | Deletes a specified file. |
| [DeleteFolder](#deletefolder) | Deletes a specified folder and its contents. |
| [DriveExists](#driveexists) | Checks if the specified drive exists. |
| [DriveLetters](#driveletters) | Returns a semicolon separated list with the driver letters. |
| [FileExists](#fileexists) | Checks for the existence of the specified file. |
| [FolderExists](#folderexists) | Checks for the existence of the specified folder. |
| [GetAbsolutePathName](#getabsolutepathname) | Returns complete and unambiguous path from a provided path specification. |
| [GetBaseName](#getbasename) | Returns a string containing the base name of the last component, less any file extension, in a path. |
| [GetDriveAvailableSpace](#getfriveavailablespace) | Returns the amount of space available to a user on the specified drive or network share. |
| [GetDriveFileSystem](#getdrivefilesystem) | Returns the type of file system in use for the specified drive or network share. |
| [GetDriveFreeSpace](#getdrivefreespace) | Returns the amount of free space available to a user on the specified drive or network share. |
| [GetDriveName](#getdrivename) | Returns a string containing the name of the drive for a specified path. |
| [GetDriveShareName](#getdrivesharename) | Returns the network share name for a specified drive. |
| [GetDriveTotalSize](#getdrivetotalsize) | Returns the total space, in bytes, of a drive or network share. |
| [GetDriveType](#getdrivetype) | Returns a value indicating the type of a specified drive. |
| [GetExtensionName](#getextensionname) | Returns a string containing the extension name of the file for a specified path. |
| [GetFileAttributes](#getfileattributes) | Returns the attributes of files. Read/write or read-only, depending on the attribute. |
| [GetFileDateCreated](#getfiledatecreated) | Returns the date and time that the specified file was created. |
| [GetFileDateLastAccessed](#getfiledatelastaccessed) | Returns the date and time that the specified file was accessed. |
| [GetFileDateLastModified](#getfiledatelastmodified) | Returns the date and time that the specified file was modified. |
| [GetFileName](#getfilename) | Returns a string containing the name of the file for a specified path. |
| [GetFileShortName](#getfileshortname) | Returns the short name used by programs that require the earlier 8.3 file naming convention. |
| [GetFileShortPath](#getfileshortpath) | Returns the short path used by programs that require the earlier 8.3 file naming convention. |
| [GetFileSize](#getfilesize) | Returns the size, in bytes, of the specified file. |
| [GetFileType](#getfiletype) | Returns information about the type of a file. |
| [GetFileVersion](#getfileversion) | Returns the version number of a specified file. |
| [GetFolderAttributes](#getfolderattributes) | Returns the attributes of folders. |
| [GetFolderDateCreated](#getfolderdatecreated) | Returns the date and time that the specified folder was created. |
| [GetFolderDateLastAccessed](#getfolderdatelastaccessed) | Returns the date and time that the specified folder was last accessed. |
| [GetFolderDateLastModified](#getfolderdatelastmodified) | Returns the date and time that the specified folder was last modified. |
| [GetFolderDriveLetter](#getfolderdriveletter) | Returns a string containing the drive letter for a specified folder. |
| [GetFolderName](#getfoldername) | Returns a string containing the name of the folder for a specified path, i.e. the path minus the file name. |
| [GetFolderShortName](#getfoldershortname) | Returns the short name used by programs that require the earlier 8.3 file naming convention. |
| [GetFolderShortPath](#getfoldershortpath) | Returns the short path used by programs that require the earlier 8.3 file naming convention. |
| [GetFolderSize](#getfoldersize) | Returns the size, in bytes, of all files and subfolders contained in the folder. |
| [GetFolderType](#getfoldertype) | Returns information about the type of a folder. |
| [GetNumDrives](#getnumdrives) | Returns the number of drives. |
| [GetNumFiles](#getnumfiles) | Returns the number of files contained in a specified folder, including those with hidden and system file attributes set. |
| [GetNumSubFolders](#getnumsubfolders) | Returns the number of folders contained in a specified folder, including those with hidden and system file attributes set. |
| [GetParentFolderName](#getparentfoldername) | Returns the folder name for the parent of the specified folder. |
| [GetSerialNumber](#getserialnumber) | Returns the decimal serial number used to uniquely identify a disk volume. |
| [GetStandardStream](#getstandardstream) | Returns a TextStream object corresponding to the standard input, output, or error stream. |
| [GetStdErrStream](#getstderrstream) | Returns a TextStream object corresponding to the standard error stream. |
| [GetStdInStream](#getstdinstream) | Returns a TextStream object corresponding to the standard input stream. |
| [GetStdOutStream](#getstdoutstream) | Returns a TextStream object corresponding to the standard output stream. |
| [GetTempName](#gettempname) | Returns a randomly generated temporary file or folder name that is useful for performing operations that require a temporary file or folder. |
| [GetVolumeName](#getvolumename) | Returns the volume name of the specified drive. |
| [IsDriveReady](#isdriveready) | Returns True if the specified drive is ready; False if it is not. |
| [IsRootFolder](#isrootfolder) | Returns True(-1) if the specified folder is the root folder; False(0) if it is not. |
| [MoveFile](#movefile) | Moves one or more files from one location to another. |
| [MoveFolder](#movefolder) | Moves one or more folders from one location to another. |
| [SetFileAttributes](#setfileattributes) | Sets the attributes of files. |
| [SetFileName](#setfilename) | Sets the name of a specified file. |
| [SetFolderAttributes](#setfolderattributes) | Sets the attributes of folders. |
| [SetFolderName](#setfoldername) | Sets the name of a specified folder. |
| [SetVolumeName](#setvolumename) | Sets the volume name of the specified drive. |

---

## Error and result codes

| Name       | Description |
| ---------- | ----------- |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [SetResult](#setresult) | Sets the last result code. |
| [GetErrorInfo](#geterrorinfo) | Returns a description of the specified error code. |

---

## GetLastResult

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```
---

## SetResult

Sets the result code.
```
FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
```
| Parameter | Description |
| --------- | ----------- |
| *Result* | The error code returned by the methods. |

---

## GetErrorInfo

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS DWSTRING
```
---

## BuildPath

Appends a name to an existing path.

```
FUNCTION BuildPath (BYREF wszPath AS WSTRING, BYREF wszName AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszPath* | Existing path to which name is appended. Path can be absolute or relative and need not specify an existing folder. |
| *wszName* | Name being appended to the existing path. |

#### Return value

The new path.

#### Remarks

The **BuildPath** method inserts an additional path separator between the existing path and the name, only if necessary.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsNewPath AS DWSTRING = pFileSys.BuildPath("C:\MyFolder", "Text.txt")
' Output: C:\MyFolder\Text.txt
```
Alternate syntax:
```
CFileSys().BuildPath("C:\MyFolder", "Text.txt")
```
---

## CopyFile

Copies one or more files from one location to another.

```
FUNCTION CopyFile (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING, _
   BYVAL OverWriteFiles AS BOOLEAN = TRUE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszSource* | Character string file specification, which can include wildcard characters, for one or more files to be copied. |
| *wszDestination* | Character string destination where the file or files from source are to be copied. Wildcard characters are not allowed. |
| *OverWriteFiles* | Boolean value that indicates if existing files are to be overwritten. If true, files are overwritten; if false, they are not. The default is true. Note that **CopyFile** will fail if destination has the read-only attribute set, regardless of the value of overwrite. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

Wildcard characters can only be used in the last path component of the *wszSource* argument. For example, you can use:

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.CopyFile("c:\mydocuments\letters\*.doc", "c:\tempfolder\")
```

But you can't use:

```
pFileSys.CopyFile("c:\mydocuments\*\R1???97.xls", "c:\tempfolder")
```

If *wszSource* contains wildcard characters, or destination ends with a path separator ("\\"), it is assumed that destination is an existing folder in which to copy matching files. Otherwise, destination is assumed to be the name of a file to create. In either case, three things can happen when an individual file is copied:

   - If destination does not exist, source gets copied. This is the usual case.

   - If destination is an existing file, an error occurs if overwrite is False. Otherwise, an attempt is made to copy source over the existing file.

   - If destination is a directory, an error occurs.

   - An error also occurs if a source using wildcard characters doesn't match any files. The **CopyFile** method stops on the first error it encounters. No attempt is made to roll back or undo any changes made before an error occurs.

Files copied to a new destination path will keep the same file name. To rename the copied file, simply include the new file name in the destination path. For example, this will copy the file to a new location and the file in the new location will have a different name:

```
pFileSys.CopyFile("c:\mydocuments\letters\sample.doc", "c:\tempfolder\sample_new.doc")
```

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.CopyFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\MyFile.txt")
```
Alternate syntax:
```
CFileSys().CopyFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\MyFile.txt")
```
---

## CopyFolder

Recursively copies a folder from one location to another.

```
FUNCTION CopyFolder (BYREF wszSource AS WSTRING, BYREF wsDestination AS WSTRING, _
   BYVAL OverWriteFiles AS BOOLEAN = TRUE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszSource* | Character string file specification, which can include wildcard characters, for one or more folders to be copied. |
| *wszDestination* |  Character string destination where the folder and subfolders from source are to be copied (must end with a "\\"). Wildcard characters are not allowed.  |
| *OverWriteFiles* | Boolean value that indicates if existing folders are to be overwritten. If true, files are overwritten; if false, they are not. The default is true. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

Wildcard characters can only be used in the last path component of the *wszource* argument. For example you can use:

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.CopyFolder "c:\mydocuments\letters\*", "c:\tempfolder\"
```

But you can't use:

```
pFileSys.CopyFolder "c:\mydocuments\*\*", "c:\tempfolder\"
```

If *wszSource* contains wildcard characters, or destination ends with a path separator ("\\"), it is assumed that destination is an existing folder in which to copy matching folders and subfolders. Otherwise, destination is assumed to be the name of a folder to create. In either case, four things can happen when an individual folder is copied:

   - If destination does not exist, the source folder and all its contents gets copied. This is the usual case.

   - If destination is an existing file, an error occurs.

   - If destination is a directory, an attempt is made to copy the folder and all its contents. If a file contained in source already exists in destination, an error occurs if overwrite is False. Otherwise, it will attempt to copy the file over the existing file.

   - If destination is a read-only directory, an error occurs if an attempt is made to copy an existing read-only file into that directory and *OverWriteFiles* is False.

An error also occurs if a source using wildcard characters doesn't match any folders.

The **CopyFolder** method stops on the first error it encounters. No attempt is made to roll back any changes made before an error occurs.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.CopyFolder("C:\MyFolder", "C:\MyOtherFolder\")
```
---

## CreateFolder

Creates a folder.

```
FUNCTION CreateFolder (BYREF wszFolderSpec AS WSTRING) AS Afx_IFolder PTR
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolderSpec* | String expression that identifies the folder to create. |

#### Return value

IDispatch. Reference to an **IFolder** object.

#### Remarks

An error occurs if the specified folder already exists.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pFolder AS Afx_IFolder PTR
pFolder = pFileSys.CreateFolder("C:\MyNewFolder")
IF pFolder THEN
   ' ....
   pFolder->Release
END IF
```
---

## DeleteFile

Deletes a specified file.

```
FUNCTION DeleteFile (BYREF wszFileSpec AS WSTRING, BYVAL bForce AS BOOLEAN = FALSE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszFileSpec* | The name of the file to delete. cbsFileSpec can contain wildcard characters in the last path component. |
| *bForce* | Boolean value that is true if files with the read-only attribute set are to be deleted; false (default) if they are not. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

An error occurs if no matching files are found. The **DeleteFile** method stops on the first error it encounters. No attempt is made to roll back or undo any changes that were made before an error occurred.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.DeleteFile("C:\MyFolder\MyFile.txt")
```
---

## DeleteFolder

Deletes a specified folder and its contents.

```
FUNCTION DeleteFolder (BYREF wszFolderSpec AS WSTRING, BYVAL bForce AS BOOLEAN = FALSE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolderSpec* | The name of the folder to delete. *wszFolderSpec* can contain wildcard characters in the last path component. |
| *bForce* | Boolean value that is true if folders with the read-only attribute set are to be deleted; false (default) if they are not. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

The **DeleteFolder** method does not distinguish between folders that have contents and those that do not. The specified folder is deleted regardless of whether or not it has contents. 

An error occurs if no matching folders are found. The **DeleteFolder** method stops on the first error it encounters. No attempt is made to roll back or undo any changes that were made before an error occurred.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.DeleteFolder("C:\MyFolder")
```
---

## DriveExists

Returns True if the specified drive exists; False if it does not.

```
FUNCTION DriveExists (BYREF wszDriveSpec AS WSTRING) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *wszDriveSpec* | A drive letter or a complete path specification. |

#### Return value

BOOLEAN. True if the specified drive exists; False if it does not.

#### Remarks

For drives with removable media, the **DriveExists** method returns true even if there are no media present. Use the **IsDriveReady** method to determine if a drive is ready.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM fExists AS BOOLEAN = pFileSys.DriveExists("C:")
```
---

## DriveLetters

Returns a semicolon separated list with the driver letters.

```
FUNCTION DriveLetters () AS DWSTRING
```

#### Return value

A semicolon separated list with the driver letters.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsLetters AS DWSTRING = pFileSys.DriveLetters
```
---

## FileExists

Returns True if the specified file exists; False if it does not.

```
FUNCTION FileExists (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *wszFileSpec* | CBSTR. The name of the file whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the file isn't expected to exist in the current folder. |

#### Return value

Boolean. True if the specified file exists; False if it does not.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM fExists AS BOOLEAN = pFileSys.FileExists("C:\MyFolder\Test.txt")
```
---

## FolderExists

Returns True if the specified folder exists; False if it does not.

```
FUNCTION FolderExists (BYREF wszFolderSpec AS WSTRING) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolderSpec* | CBSTR. The name of the folder whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the folder isn't expected to exist in the current folder. |

#### Return value

Boolean. True if the specified folder exists; False if it does not.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM fExists AS BOOLEAN = pFileSys.FolderExists("C:\MyFolder")
```
---

## GetAbsolutePathName

Returns complete and unambiguous path from a provided path specification.

```
FUNCTION GetAbsolutePathName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszPathSpec* | Path specification to change to a complete and unambiguous path. |

#### Return value

Boolean. The path name.

#### Remarks

A path is complete and unambiguous if it provides a complete reference from the root of the specified drive. A complete path can only end with a path separator character ("\\") if it specifies the root folder of a mapped drive.

Assuming the current directory is c:\mydocuments\reports, the following table illustrates the behavior of the **GetAbsolutePathName** method.

| pathspec   | Returned path |
| ---------- | ------------- |
| "c:" | "c:\mydocuments\reports" |
| "c:.." | "c:\mydocuments" |
| "c:\" | "c:\" |
| "c:*.*\may97" | "c:\mydocuments\reports\\\*\.\*\may97" |
| "region1" | "c:\mydocuments\reports\region1" |
| "c:\..\..\mydocuments" | "c:\mydocuments" |

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetAbsolutePathName("C:\MyFolder\Test.txt")
```
---

## GetBaseName

Returns a string containing the base name of the last component, less any file extension, in a path

```
FUNCTION GetBaseName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszPathSpec* | The path specification for the component whose base name is to be returned. |

#### Return value

Boolean. The base name.

#### Remarks

The **GetBaseName** method returns a zero-length string ("") if no component matches the path argument. The **GetBaseName** method works only on the provided path string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cwsName AS CWSTR = pFileSys.GetBaseName("C:\MyFolder\Test.txt")
```
---

## GetDriveAvailableSpace

Returns the amount of space available to a user on the specified drive or network share.

```
FUNCTION GetDriveAvailableSpace (BYREF wszDrive AS WSTRING) AS DOUBLE
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

DOUBLE. The amount of available space.

#### Remarks

The value returned by the **GetDriveAvailableSpace** method is typically the same as that returned by the **GetDriveFreeSpace** method. Differences may occur between the two for computer systems that support quotas. 

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
PRINT pFileSys.GetDriveAvailableSpace("C:")
```
---

## GetDriveFileSystem

Returns the type of file system in use for the specified drive or network share.

```
FUNCTION GetDriveFileSystem (BYREF wszDrive AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

The type of file system in use.

#### Remarks

Available return types include FAT, NTFS, and CDFS.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsFileSystem AS DWSTRING = pFileSys.GetDriveFileSystem("C:")
```
---

## GetDriveFreeSpace

Returns the amount of free space available to a user on the specified drive or network share

```
FUNCTION GetDriveFreeSpace (BYREF wszDrive AS WSTRING) AS DOUBLE
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

DOUBLE. The amount of free space.

#### Remarks

The value returned by the **GetDriveFreeSpace** property is typically the same as that returned by the **GetDriveAvailableSpace** property. Differences may occur between the two for computer systems that support quotas. 

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
PRINT pFileSys.GetDriveFreeSpace("C:")
```
---

## GetDriveName

Returns a string containing the name of the drive for a specified path.

```
FUNCTION GetDriveName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszPathSpec* | CBSTR. The path. |

#### Return value

The drive name.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsName AS DWSTRING = pFileSys.GetDriveName("C:\MyFolder\Test.txt")
```
---

## GetDriveShareName

Returns the network share name for a specified drive.

```
FUNCTION GetDriveShareName (BYREF wszDrive AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

The share name.

#### Remarks

If object is not a network drive, the **GetDriveShareName** method returns a zero-length string ("").

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsShareName AS DWSTRING = pFileSys.GetDriveShareName("H:")
```
---

## GetDriveTotalSize

Returns the total space, in bytes, of a drive or network share.

```
FUNCTION GetDriveTotalSize (BYREF wszDrive AS WSTRING) AS DOUBLE
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

DOUBLE. The total space in bytes.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
PRINT pFileSys.GetDriveTotalSize("C:")
```
---

## GetDriveType

Returns a value indicating the type of a specified drive.

```
FUNCTION GetDriveType (BYREF wszDrive AS WSTRING) AS DRIVETYPECONST
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

The type of the specified drive.

```
DriveType_UnknownType = 0
DriveType_Removable = 1
DriveType_Fixed = 2
DriveType_Remote = 3
DriveType_CDRom = 4
DriveType_RamDisk = 5
```
#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDriveType AS DRIVETYPECONST = pFileSys.GetDriveType("C:")
DIM t AS DWSTRING
SELECT CASE pDrive.DriveType
   CASE 0 : t = "Unknown"
   CASE 1 : t = "Removable"
   CASE 2 : t = "Fixed"
   CASE 3 : t = "Network"
   CASE 4 : t = "CD-ROM"
   CASE 5 : t = "RAM Disk"
END SELECT
```
---

## GetExtensionName

Returns a string containing the extension name of the file for a specified path.

```
FUNCTION GetExtensionName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszPathSpec* | The extension name. |

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsName AS DWSTRING = pFileSys.GetExtensionName("C:\MyFolder\Test.txt")
```
---

## GetFileAttributes

Returns the attributes of files.

```
FUNCTION GetFileAttributes (BYREF wszFile AS WSTRING) AS FILEATTRIBUTE
```

| Name       | Description |
| ---------- | ----------- |
| *wszFile* | The path to a specific file. |

#### Return value

The file attributes. Can be any of the following values or any logical combination of the following values:

| Constant   | Value       | Description |
| ---------- | ----------- | ----------- |
| FileAttribute_Normal     | 0 | Normal file. No attributes are set. |
| FileAttribute_ReadOnly   | 1 | Read-only file. Attribute is read/write. |
| FileAttribute_Hidden     | 2 | Hidden file. Attribute is read/write. |
| FileAttribute_System     | 4 | System file. Attribute is read/write. |
| FileAttribute_Volume     | 8 | Disk drive volume label. Attribute is read-only. |
| FileAttribute_Directory | 16 | Folder or directory. Attribute is read-only. |
| FileAttribute_Archive |   32 | File has changed since last backup. Attribute is read/write. |
| FileAttribute_Alias   | 1024 | Link or shortcut. Attribute is read-only. |
| FileAttribute_Compressed | 2048 | Compressed file. Attribute is read-only. |

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM lAttr FILEATTRIBUTE = pFileSys.GetFileAttributes("C:\MyPath\MyFile.txt")
```
---

## GetFileDateCreated

Returns the date and time that the specified file was created.

```
FUNCTION GetFileDateCreated (BYREF wszFile AS WSTRING) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *wszFile* | The path to a specific file. |

#### Return value

DATE_. The date and time of the file creation. This is a *Date Serial* number that can be formatted using the Free Basic's **Format** function. You can also use the wrapper function **AfxVariantDateTimeToStr**.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFileDateCreated("C:\MyPath\MyFile.txt")
```
---

## GetFileDateLastAccessed

Returns the date and time that the specified file was accesed.

```
FUNCTION GetFileDateLastAccessed (BYREF wszFile AS WSTRING) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *wszFile* | The path to a specific file. |

#### Return value

DATE_. The date and time that the file was last accessed. This is a *Date Serial* number that can be formatted using the Free Basic's **Format** function. You can also use the wrapper function **AfxVariantDateTimeToStr**.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFileDateLastAccessed("C:\MyPath\MyFile.txt")
```
---

## GetFileDateLastModified

Returns the date and time that the specified file was accessed.

```
FUNCTION GetFileDateLastModified (BYREF wszFile AS WSTRING) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *wszFile* | The path to a specific file. |

#### Return value

DATE_. The date and time that the file was last modified. This is a *Date Serial* number that can be formatted using the **DateTimeToString** method or the Free Basic's **Format** function.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFileDateLastModified("C:\MyPath\MyFile.txt")
PRINT pFilesys.DateTimeToString(nDate)
```
---

## GetFileName

Returns a string containing the name of the file for a specified path.

```
FUNCTION GetFileName (BYREF wszPathSpec AS WSTRING) AS DWSREING
```

| Name       | Description |
| ---------- | ----------- |
| *wszPathSpec* | The path to a specific file. |

#### Return value

The file name.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsName AS DWSTRING = pFileSys.GetFileName("C:\MyFolder\Test.txt")
```
---

## GetFileShortName

Returns the short name used by programs that require the earlier 8.3 file naming convention.

```
FUNCTION GetFileShortName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszPathSpec* | The path to a specific file. |

#### Return value

The short name for the specified file.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsName AS DWSTRING = pFileSys.GetFileShortName("C:\MyFolder\Test.txt")
```

## GetFileShortPath

Returns the short path used by programs that require the earlier 8.3 file naming convention.

```
FUNCTION GetFileShortPath (BYREF wszPathSpec AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszPathSpec* | The path to a specific file. |

#### Return value

The short path to a specific file.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsName AS DWSTRING = pFileSys.GetFileShortPath("C:\MyFolder\Test.txt")
```
---

## GetFileSize

Returns the size, in bytes, of the specified file.

```
FUNCTION GetFileSize (BYREF wszFile AS WSTRING) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *wszFile* | The path to a specific file. |

#### Return value

LONG. The size, in bytes, of the file.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nFileSize AS LONG = pFileSys.GetFileSize("C:\MyPath\MyFile.txt")
```
---

## GetFileType

Returns localized information about the type of a file. For example, for files ending in .TXT, "Text Document" is returned ("Documento de texto" in Spanish computers).

```
FUNCTION GetFileType (BYREF wszFile AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszFile* | The path to a specific file. |

#### Return value

The type of file.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsFileType AS DWSTRING = pFileSys.GetFileType("C:\MyPath\MyFile.txt")
```
---

## GetFileVersion

Returns the version number of a specified file.

```
FUNCTION GetFileVersion (BYREF dwsFile AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszFile* | The path (absolute or relative) to a specific file. |

#### Return value

The version number.

#### Remarks

The **GetFileVersion** method returns a zero-length string ("") if *wszFile* does not end with the named component. 

Note: The **GetFileVersion** method works only on the provided path string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsVersion AS DWSTRING = pFileSys.GetFileVersion("C:\MyFolder\MyFile.doc")
IF LEN(dwsVersion) THEN
   MSGBOX "File version: " & dwsVersion
ELSE
   MSGBOX "No version information available"
END IF
```
```
DIM pFileSys AS CFileSys
print pFileSys.GetFileVersion("c:\windows\system32\scrrun.dll")
```
---

## GetFolderAttributes

Returns the attributes of folders.

```
FUNCTION GetFolderAttributes (BYREF wszFolder AS WSTRING) AS FILEATTRIBUTE
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | CBSTR. The path to a specific folder. |

#### Return value

The attributes. Can be any of the following values or any logical combination of the following values:

| Constant   | Value       | Description |
| ---------- | ----------- | ----------- |
| FileAttribute_Normal     | 0 | Normal file. No attributes are set. |
| FileAttribute_ReadOnly   | 1 | Read-only file. Attribute is read/write. |
| FileAttribute_Hidden     | 2 | Hidden file. Attribute is read/write. |
| FileAttribute_System     | 4 | System file. Attribute is read/write. |
| FileAttribute_Volume     | 8 | Disk drive volume label. Attribute is read-only. |
| FileAttribute_Directory | 16 | Folder or directory. Attribute is read-only. |
| FileAttribute_Archive |   32 | File has changed since last backup. Attribute is read/write. |
| FileAttribute_Alias   | 1024 | Link or shortcut. Attribute is read-only. |
| FileAttribute_Compressed | 2048 | Compressed file. Attribute is read-only. |

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM lAttr FILEATTRIBUTE = pFileSys.GetFolderAttributes("C:\MyPath")
```
---

## GetFolderDateCreated

Returns the date and time that the specified folder was created.

```
FUNCTION GetFolderDateCreated (BYREF wszFolder AS WSTRING) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |

#### Return value

DATE_. The date and time that the folder was created. This is a *Date Serial* number that can be formatted using The **DateTimeoString**method pr the Free Basic's **Format** function.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFolderDateCreated("C:\MyPath")
PRINT pFilesys.DateTimeToString(nDate)
```
---

## GetFolderDateLastAccessed

Returns the date and time that the specified folder was last accessed.

```
FUNCTION GetFolderDateLastAccessed (BYREF dwsFolder AS WSTRING) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |

#### Return value

DATE_. The date and time that the folder was last accessed. This is a *Date Serial* number that can be formatted using the **DateTimeToString** method or the Free Basic's **Format** function.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFolderDateLastAccessed("C:\MyPath)
PRINT pFilesys.DateTimeToString(nDate)
```
---

## GetFolderDateLastModified

Returns the date and time that the specified folder was last modified.

```
FUNCTION GetFolderDateLastModified (BYREF wszFolder AS WSTRING) AS DATE_
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | CBSTR. The path to a specific folder. |

#### Return value

DATE_. The date and time that the folder was last modified. This is a *Date Serial* number that can be formatted using the **DateTimeToString** method or the Free Basic's **Format** function.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nDate AS DATE_ = pFileSys.GetFolderDateLastModified("C:\MyPath)
PRINT pFilesys.DateTimeToString(nDate
```
---

## GetFolderDriveLetter

Returns a string containing the drive letter for a specified folder.

```
FUNCTION GetFolderDriveLetter (BYREF wszFolder AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |

#### Return value

The drive letter of the folder.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsDriveLetter AS DWSTRING = pFileSys.GetFolderDriveLetter("c:\MyFolder)
```
---

## GetFolderName

Returns a string containing the name of the folder for a specified path, i.e. the path minus the file name.

```
FUNCTION GetFolderName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |

#### Return value

The name of the folder.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsName AS DWSTRING = pFileSys.GetFolderName("C:\MyFolder\Test.txt")
```
---

## GetFolderShortName

Returns the short name used by programs that require the earlier 8.3 file naming convention.

```
FUNCTION GetFolderShortName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |

#### Return value

The short name for the specified folder.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsFolderShorName AS DWSTRING = pFileSys.GetFolderShortName("c:\MyFolder)
```
---

## GetFolderShortPath

Returns the short path used by programs that require the earlier 8.3 file naming convention.

```
FUNCTION GetFolderShortPath (BYREF wszPathSpec AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |

#### Return value

The short path for the specified folder.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsFolderShortPath AS DWSTRING = pFileSys.GetFolderShortPath("c:\MyFolder)
```
---

## GetFolderSize

Returns the size, in bytes, of all files and subfolders contained in the folder.

```
FUNCTION GetFolderSize (BYREF wszFolder AS WsTRING) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | CBSTR. The path to a specific folder. |

#### Return value

LONG. The size, in bytes, of all files and subfolders contained in the folder.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nFolderSize AS LONG = pFileSys.GetFolderSize("C:\MyPath")
```
---

## GetFolderType

Returns information about the type of a folder.

```
FUNCTION GetFolderType (BYREF wszFolder AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | CBSTR. The path to a specific folder. |

#### Return value

CBSTR. The type of folder.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsFolderType AS DWSTRING = pFileSys.GetFolderType("c:\MyFolder)
```
---

## GetNumDrives

Returns the number of drives.

```
FUNCTION GetNumDrives () AS LONG
```
---

#### Return value

LONG. The number of drives.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM numDrives AS LONG = pFileSys.GetNumDrives
```
---

## GetNumFiles

Returns the number of files contained in a specified folder, including those with hidden and system file attributes set.

```
FUNCTION GetNumFiles (BYREF wszFolder AS WSTRING) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |

#### Return value

LONG. The number of files.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM numFiles AS LONG = pFileSys.GetNumFiles("C:\MyFolder")
```
---

## GetNumSubFolders

Returns the number of files contained in a specified folder, including those with hidden and system file attributes set.

```
FUNCTION GetNumSubFolders (BYREF wszFolder AS WsTRING) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |

#### Return value

LONG. The number of subfolders.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM numSubFolders AS LONG = pFileSys.GetNumSubFolders("C:\MyFolder")
```
---

## GetParentFolderName

Returns the folder name for the parent of the specified folder.

```
FUNCTION GetParentFolderName (BYREF wszFolder AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |

#### Return value

The name of the parent folder, or a empty string if the folder has no parent.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsParentFolderName AS DWsTRING = pFileSys.GetParentFolderName("C:\MyFolder\MySubfolder")
```
---

## GetSerialNumber

Returns the decimal serial number used to uniquely identify a disk volume.

```
FUNCTION GetSerialNumber (BYREF wszDrive AS WSTRING) AS LONG
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

LONG. The serial number.

#### Remarks

You can use the **GetSerialNumber** method to ensure that the correct disk is inserted in a drive with removable media. 

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM nSerialNumber AS LONG = pFileSys.GetSerialNumber("C:")
```
---

## GetStandardStream

Returns a TextStream object corresponding to the standard input, output, or error stream.

```
FUNCTION GetStandardStream (BYVAL nStreamType AS STANDARDSTREAMTYPES, _
   BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
```

| Name       | Description |
| ---------- | ----------- |
| *nStreamType* | LONG. Can be one of three constants: *StandardStreamTypes_StdErr*, *StandardStreamTypes_StdIn*, or *StandardStreamTypes_StdOut*. |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it is created as an ASCII file. If omitted, an ASCII file is assumed. |

#### Return value

A pointer to the *ITextStream* interface of the requested standard stream.

#### Settings

The nStreamType argument can have any of the following settings:

| Constant   | Constant    | Description |
| ---------- | ----------- | ----------- |
| StandardStreamTypes_StdIn  | 0 | Returns a TextStream object corresponding to the standard input stream. |
| StandardStreamTypes_StdOut | 1 | Returns a TextStream object corresponding to the standard output stream. |
| StandardStreamTypes_StdErr | 2 | Returns a TextStream object corresponding to the standard error stream. |

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pStm AS Afx_ITextStream PTR = pFileSys.GetStandardStream(StandardStreamTypes_StdOut)
```
---

## GetStdErrStream

Returns a **TextStream** object corresponding to the standard error stream.

```
FUNCTION GetStdErrStream (BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
```

| Name       | Description |
| ---------- | ----------- |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it is created as an ASCII file. If omitted, an ASCII file is assumed. |

#### Return value

A pointer to the *ITextStream* interface of the standard error stream.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pStm AS Afx_ITextStream PTR = pFileSys.GetStdErrStream
```
---

## GetStdInStream

Returns a TextStream object corresponding to the standard input stream.

```
FUNCTION GetStdInStream (BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
```

| Name       | Description |
| ---------- | ----------- |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it is created as an ASCII file. If omitted, an ASCII file is assumed. |

#### Return value

A pointer to the *ITextStream* interface of the standard input stream.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pStm AS Afx_ITextStream PTR = pFileSys.GetStdInStream
```
---

## GetStdOutStream

Returns a **TextStream** object corresponding to the standard output stream.

```
FUNCTION GetStdOutStream (BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
```

| Name       | Description |
| ---------- | ----------- |
| *bUnicode* | Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is true if the file is created as a Unicode file, false if it is created as an ASCII file. If omitted, an ASCII file is assumed. |

#### Return value

A pointer to the **ITextStream** interface of the standard output stream.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM pStm AS Afx_ITextStream PTR = pFileSys.GetStdOutStream
```
---

## GetTempName

Returns a randomly generated temporary file or folder name that is useful for performing operations that require a temporary file or folder.

```
FUNCTION GetTempName () AS DWSTRING
```

#### Return value

The temporary name.

#### Remarks

The **GetTempName** method does not create a file. It provides only a temporary file name that can be used with **CreateTextFile** to create a file.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM cbsName AS CBSTR = pFileSys.GetTempName
```
---

## GetVolumeName

Returns the volume name of the specified drive.

```
FUNCTION GetVolumeName (BYREF dwsDrive AS WSTRING) AS DWSTRING
```

| Name       | Description |
| ---------- | ----------- |
| *dwsDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

The volume name.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM dwsVolumeName AS DWSTRING = pFileSys.GetVolumeName("C:")
```
---

## IsDriveReady

Returns True if the specified drive is ready; False if it is not.

```
FUNCTION IsDriveReady (BYREF dwsDrive AS STRING) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

BOOLEAN. True if the specified drive is ready; False if it is not.

#### Remarks

For removable-media drives and CD-ROM drives, **IsDriveReady** returns True only when the appropriate media is inserted and ready for access. 

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM bIsReady AS BOOLEAN = pFileSys.IsDriveReady("C:")
```
---

## IsRootFolder

Returns True(-1) if the specified folder is the root folder; False(0) if it is not.

```
FUNCTION IsRootFolder (BYREF dwsFolder AS WSTRIN) AS BOOLEAN
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | CBSTR. The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |

#### Return value

BOOLEAN. True(-1) if the specified folder is the root folder; False(0) if it is not.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
DIM bIsRoot AS BOOLEAN = pFileSys.IsRootFolder("C:\MyFolder")
```
---

## MoveFile

Moves one or more files from one location to another.

```
FUNCTION MoveFile (BYREF wszSource AS DWSTRING, BYREF wszDestination AS WSTRING) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *cwszource* | The path to the file or files to be moved. The *wszSource* argument string can contain wildcard characters in the last path component only. |
| *wszDestination* | Destination where the file is to be moved. Wildcard characters are not allowed. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.MoveFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\")
```
---

## MoveFolder

Moves one or more folders from one location to another.

```
FUNCTION MoveFolder (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszSource* | CBSTR. The path to the folder or folders to be moved. The *cbsSource* argument string can contain wildcard characters in the last path component only. |
| *wszDestination* | CBSTR. Destination where the folder is to be moved (must end with a "\\"). Wildcard characters are not allowed. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.MoveFolder("C:\MyFolder", "C:\MyNewFolder\")
```
---

## SetFileAttributes

Sets the attributes of files.

```
FUNCTION SetFileAttributes (BYREF wszFile AS WSTRING, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszFile* | The path to a specific file. |
| *lAttr* | LONG. The new value for the attributes of the specified file. |

#### Return value

The file attributes. Can be any of the following values or any logical combination of the following values:

| Constant   | Value       | Description |
| ---------- | ----------- | ----------- |
| FileAttribute_Normal     | 0 | Normal file. No attributes are set. |
| FileAttribute_ReadOnly   | 1 | Read-only file. Attribute is read/write. |
| FileAttribute_Hidden     | 2 | Hidden file. Attribute is read/write. |
| FileAttribute_System     | 4 | System file. Attribute is read/write. |
| FileAttribute_Volume     | 8 | Disk drive volume label. Attribute is read-only. |
| FileAttribute_Directory | 16 | Folder or directory. Attribute is read-only. |
| FileAttribute_Archive |   32 | File has changed since last backup. Attribute is read/write. |
| FileAttribute_Alias   | 1024 | Link or shortcut. Attribute is read-only. |
| FileAttribute_Compressed | 2048 | Compressed file. Attribute is read-only. |

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetFileAttributes("C:\MyPath\MyFile.txt", 33)   ' FileAttribute_Archive OR FileAttribute_Normal
```
---

## SetFileName

Sets the name of a specified file.

```
FUNCTION SetFileName (BYREF wszFile AS WSTRING, BYREF wszName AS WSTRING) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszFile* | CBSTR. The path to a specific file. |
| *wszName* | CBSTR. The new name of the file. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

You only have to pass the new name of the file, not the full path.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetFileName("c:\MyFolder\Test.txt", "NewName")
```
---

## SetFolderAttributes

Sets the attributes of folders.

```
FUNCTION SetFolderAttributes (BYREF wszFolder AS WSTRING, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | CBSTR. The path to a specific folder. |
| *lAttr* | LONG. The new value for the attributes of the specified folder. |

#### Return value

The *lAttr* argument can have any of the following values or any logical combination of the following values:

| Constant   | Value       | Description |
| ---------- | ----------- | ----------- |
| FileAttribute_Normal     | 0 | Normal file. No attributes are set. |
| FileAttribute_ReadOnly   | 1 | Read-only file. Attribute is read/write. |
| FileAttribute_Hidden     | 2 | Hidden file. Attribute is read/write. |
| FileAttribute_System     | 4 | System file. Attribute is read/write. |
| FileAttribute_Volume     | 8 | Disk drive volume label. Attribute is read-only. |
| FileAttribute_Directory | 16 | Folder or directory. Attribute is read-only. |
| FileAttribute_Archive |   32 | File has changed since last backup. Attribute is read/write. |
| FileAttribute_Alias   | 1024 | Link or shortcut. Attribute is read-only. |
| FileAttribute_Compressed | 2048 | Compressed file. Attribute is read-only. |

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetFolderAttributes("C:\MyPath\MyFile.txt", 17)   ' FileAttribute_Directory OR FileAttribute_ReadOnly
```
---

## SetFolderName

Sets the name of a specified folder.

```
FUNCTION SetFolderName (BYREF wszFolder AS WSTRING, BYREF wszName AS WSTRING) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszFolder* | The path to a specific folder. |
| *wszName* | The new name of the folder. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Remarks

You only have to pass the new name of the folder, not the full path.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetFolderName("c:\MyFolder", "NewName")
```
---

## SetVolumeName

Sets the name of a specified drive.

```
FUNCTION SetVolumeName (BYREF wszDrive AS WSTRING, BYREF wszName AS WSTRING) AS HRESULT
```

| Name       | Description |
| ---------- | ----------- |
| *wszDrive* | The drive letter. For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\\. |
| *wszName* | The volume name. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

#### Usage example

```
#INCLUDE ONCE "AfxNova/CFileSys.inc"
DIM pFileSys AS CFileSys
pFileSys.SetVolumeName("C:", "VolumeName")
```
---

## DateTimeToString

Converts a DATE_ type to a string containing the date and the time.

```
FUNCTION DateTimeToString (BYVAL vbDate AS DATE_, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT, _
   BYVAL dwFlags AS DWORD = 0) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *vbDate* | A variant representation of time. |
| *lcid* | Optional. The language identifier used for the conversion. Default is LOCALE_USER_DEFAULT. |
| *dwFlags* | Optional. A value made from one or more flags. The following table shows the flags that can be set for this parameter. |

| Flag       | Description |
| ---------- | ----------- |
| LOCALE_NOUSEROVERRIDE | Uses the system default locale settings, rather than custom locale settings. |
| VAR_CALENDAR_HJRI | If set then the Hijri calendar is used. Otherwise the calendar sent in Control Panel is used. |
| VAR_DATEVALUEONLY | Omits the time portion of a **VT_DATE** and retrieves only the date.<br>Applies to conversions to or from dates.<br>Not used for **ChangeType** and **ChangeTypeEx**. |
| VAR_TIMEVALUEONLY | Omits the date portion of a **VT_DATE** and returns only the time. Applies to conversions to or from dates. Not used for **ChangeType** and **ChangeTypeEx**. |

#### Return value

The formatted date and time.

---
