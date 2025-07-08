# CFileStream Class

Allows to work with binary file streams. A binary stream consists of one or more bytes of arbitrary information. **CFileStream** provides a stream for a file, allowing read, write and seek operations.

**Include file**: AfxNova/CStream.inc

## Constructors

| Name       | Description |
| ---------- | ----------- |
| Constructor (File Name) | Opens or creates a file and retrieves a stream to read or write to that file. |
| Constructor (Stream) | Initializes an instance of the class from an existing stream and attaches it. |

```
CONSTRUCTOR CFileStream ( _
   BYREF wszFile AS WSTRING, _
   BYVAL grfMode AS DWORD = STGM_READ, _
   BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, _
   BYVAL fCreate AS BOOLEAN = FALSE)
```

```
CONSTRUCTOR CFileStream ( _
   BYVAL pwszFile AS WSTRING PTR, _
   BYVAL grfMode AS DWORD = STGM_READ, _
   BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, _
   BYVAL fCreate AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | The file name. |
| *pwszFile* | A pointer to a unicode null-terminated string that specifies the file name. |
| *grfMode* | One or more **STGM** values that are used to specify the file access mode and how the the stream is created and deleted. |
| *dwAttributes* | One or more flag values that specify file attributes in the case that a new file is created. |
| *fCreate* | A boolean value that helps specify, in conjunction with *grfMode*, how existing files should be treated when creating the stream. |

```
CONSTRUCTOR CFileStream (BYVAL pstm AS IStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstm* | A pointer to the **IStream** interface of an existing stream that will be attached to the class. |
| *fAddRef* | TRUE to increase the reference count of the stream; FALSE, otherwise. |

---

## Operators

| Name       | Description |
| ---------- | ----------- |
| LET | Initializes the class from an existing stream and attaches it. |
| CAST | Returns a pointer to the underlying **IStream** interface of the stream object. |

```
OPERATOR LET (BYVAL pstm AS IStream PTR)
OPERATOR CAST () AS IStream PTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstm* | A pointer to the **IStream** interface of an existing stream that will be attached to the class. |

---

## Methods

| Name       | Description |
| ---------- | ----------- |
| [Attach](#attach) | Attaches the passed stream to the class. |
| [Detach](#detach) | Detaches the stream from the class. |
| [Open](#open) | Opens or creates a file and retrieves a stream to read or write to that file. |
| [Close](#close) | Releases the stream object. |
| [Read](#read) | Reads a specified number of bytes from the stream into memory, starting at the current seek pointer. |
| [ReadTextA](#readTextA) | Reads a specified number of characters from the stream into memory, starting at the current seek pointer, and returns them as an ansi string. |
| [ReadTextW](#readTextW) | Reads a specified number of characters from the stream into memory, starting at the current seek pointer, and returns them as a unicode string. |
| [Write](#write) | Writes a specified number of bytes into the stream starting at the current seek pointer. |
| [WriteTextA](#writetexta) | Writes a ansi string into the stream starting at the current seek pointer. |
| [WriteTextW](#writetextw) | Writes a unicode string into the stream starting at the current seek pointer. |
| [Seek](#seek) | Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer. |
| [GetSeekPosition](#getseekposition) | Returns the seek position. |
| [ResetSeekPosition](#resetseekposition) | Sets the seek position at the beginning of the stream. |
| [SeekAtEndOfFile](#seekatendoffile) | Sets the seek position at the end of the stream. |
| [SeekAtEndOfStream](#seekatendofstream) | Sets the seek position at the end of the stream. |
| [GetSize](#getsize) | Returns the size of the stream. |
| [SetSize](#setsize) | Changes the size of the stream. |
| [CopyTo](#copyto) | Copies a specified number of bytes from the current seek pointer in the stream to the current seek pointer in another stream. |
| [LockRegion](#lockregion) | Restricts access to a specified range of bytes in the stream. |
| [UnlockRegion](#unlockregion) | Removes the access restriction on a range of bytes previously restricted with *LockRegion*. |
| [Stat](#stat) | Retrieves the **STATSTG** structure for this stream. |
| [Clone](#clone) | Creates a new stream with its own seek pointer that references the same bytes as the original stream. |
| [StreamPtr](#streamptr) | Returns a pointer to the underlying IStream interface. |

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

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

## GetErrorInfo

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS DWSTRING
```

#### Return value

DWSTRING. A description of the last result code.

---

## Attach

Attaches the passed stream to the class.

```
FUNCTION Attach (BYVAL pstm AS IStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstm* | A pointer to the **IStream** interface of an existing stream that will be attached to the class. |
| *fAddRef* | TRUE to increase the reference count of the stream; FALSE, otherwise. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## Detach

Detaches the stream from the class.

```
FUNCTION Detach () AS IStream PTR
```

#### Return value

IStream PTR. A pointer to the **IStream** interface of the stream object.

---

## Open

Opens or creates a file and retrieves a stream to read or write to that file.

```
FUNCTION Open (BYREF wszFile AS WSTRING, _
   BYVAL grfMode AS DWORD = STGM_READ, _
   BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, _
   BYVAL fCreate AS BOOLEAN = FALSE) AS HRESULT
```
```
FUNCTION Open (BYVAL pwszFile AS WSTRING PTR, _
   BYVAL grfMode AS DWORD = STGM_READ, _
   BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, _
   BYVAL fCreate AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszFile* | The file name. |
| *pwszFile* | A pointer to a unicode null-terminated string that specifies the file name. |
| *grfMode* | One or more **STGM** values that are used to specify the file access mode and how the stream is created and deleted. The STGM constants are flags that indicate conditions for creating and deleting the stream and access modes for the stream. These elements are often combined using an **OR** operator. They are interpreted in groups as listed in the following table. It is not valid to use more than one element from a single group. |
| *dwAttributes* | One or more flag values that specify file attributes in the case that a new file is created.<br>**_0_** = Prevents other processes from opening a file or device if they request delete, read, or write access.<br>**FILE_SHARE_DELETE** : Enables subsequent open operations on a file or device to request delete access. Otherwise, other processes cannot open the file or device if they request delete access. If this flag is not specified, but the file or device has been opened for delete access, the function fails. Delete access allows both delete and rename operations.<br>**FILE_SHARE_READ** : Enables subsequent open operations on a file or device to request read access. Otherwise, other processes cannot open the file or device if they request read access. If this flag is not specified, but the file or device has been opened for read access, the function fails.<br>**FILE_SHARE_WRITE** : Enables subsequent open operations on a file or device to request write access. Otherwise, other processes cannot open the file or device if they request write access. If this flag is not specified, but the file or device has been opened for write access or has a file mapping with write access, the function fails. |
| *fCreate* | BOOL value that helps specify, in conjunction with *grfMode*, how existing files should be treated when creating the stream |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

### STGM Constants

The STGM constants are flags that indicate conditions for creating and deleting the stream and access modes for the stream. These elements are often combined using an **OR** operator. They are interpreted in groups as listed in the following table. It is not valid to use more than one element from a single group.

| Group                       | Flag                  | Value      |
| --------------------------- | --------------------- | ---------- |
| Access                      | STGM_READ             | &h00000000 |
|                             | STGM_WRITE            | &h00000001 |
|                             | STGM_READWRITE        | &h00000002 |
| Sharing                     | STGM_SHARE_DENY_NONE  | &h00000040 |
|                             | STGM_SHARE_DENY_READ  | &h00000030 |
|                             | STGM_SHARE_DENY_WRITE | &h00000020 |
|                             | STGM_SHARE_EXCLUSIVE  | &h00000010 |
| Creation                    | STGM_CREATE           | &h00001000 |
|                             | STGM_FAILIFTHERE      | &h00000000 |
| Transactioning              | STGM_DIRECT           | &h00000000 |
|                             | STGM_TRANSACTED       | &h00010000 |

| Constant   | Meaning     |
| ---------- | ----------- |
| **STGM_READ** | Indicates that the stream is read-only, meaning that modifications cannot be made. For example, if a stream is opened with **STGM_READ**, the **Read** method may be called, but the **Write** method may not. |
| **STGM_WRITE** | Enables you to save changes to the stream, but does not permit access to its data. |
| **STGM_SHARE_DENY_NONE** | Specifies that subsequent openings of the stream are not denied read or write access. If no flag from the sharing group is specified, this flag is assumed. |
| **STGM_SHARE_DENY_READ** | Prevents others from subsequently opening the stream in **STGM_READ** mode. |
| **STGM_SHARE_DENY_WRITE** | Prevents others from subsequently opening the stream for **STGM_WRITE** or **STGM_READWRITE** access. In transacted mode, sharing of **STGM_SHARE_DENY_WRITE** or **STGM_SHARE_EXCLUSIVE** can significantly improve performance because they do not require snapshots. |
| **STGM_SHARE_EXCLUSIVE** | Prevents others from subsequently opening the stream in **STGM_READ** mode. |
| **STGM_CREATE** | Indicates that an existing stream should be removed before the new stream replaces it. A new stream is created when this flag is specified only if the existing stream has been successfully removed. This flag cannot be used with open operations. |
| **STGM_FAILIFTHERE** | Causes the create operation to fail if an existing stream with the specified name exists. In this case, **STG_E_FILEALREADYEXISTS** is returned. This is the default creation mode; that is, if no other create flag is specified, **STGM_FAILIFTHERE** is implied. |
| **STGM_DIRECT** | Indicates that, in direct mode, each change to a stream element is written as it occurs. This is the default if neither **STGM_DIRECT** nor **STGM_TRANSACTED** is specified. |
| **STGM_TRANSACTED** | Indicates that, in transacted mode, changes are buffered and written only if an explicit commit operation is called. To ignore the changes, call the **Revert** method. |

The *grfMode* and *fCreate* parameters work together to specify how the function should behave with respect to existing files.

| grfMode          | fCreate | File exists? | Behavior |
| ---------------- | ------- | ------------ | -------- |
| STGM_CREATE      | Ignored | Yes          | The file is recreated. |
| STGM_CREATE      | Ignored | No           | The file is created. |
| STGM_FAILIFTHERE | FALSE   | Yes          | The file is opened. |
| STGM_FAILIFTHERE | FALSE   | No           | The call fails. |
| STGM_FAILIFTHERE | TRUE    | Yes          | The call fails. |
| STGM_FAILIFTHERE | TRUE    | No           | The file is created. |

#### Example

```
#INCLUDE ONCE "AfxNova/CStream.inc"
USING AfxNova

DIM pStream AS CFileStream
IF pStream.Open("binary.txt", STGM_CREATE OR STGM_WRITE) = S_OK then
   DIM s AS STRING = "This is a test"
   pStream.Write(STRPTR(s), LEN(s))
   pStream.Close
END IF
```
---

## Close

Releases the stream object.

```
FUNCTION Close
```
---

## Read

Reads a specified number of bytes from the stream into memory, starting at the current seek pointer.

```
FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer which the stream data is read into. |
| *cb* | The number of bytes of data to read from the stream. |
| *pcbRead* | A pointer to a ULONG variable that receives the actual number of bytes read from the stream. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

```
FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer which the stream data is read into. |
| *cb* | The number of bytes of data to read from the stream. |

#### Return value

ULONG. The actual number of bytes read from the stream. Note: The number of bytes read may be zero.

#### Example

```
#INCLUDE ONCE "AfxNova/CStream.inc"
USING AfxNova

DIM pstm AS CFileStream
pstm.Open(ExePath & "\TextA1.txt", STGM_READ)
DIM strText AS STRING = SPACE(10)
pstm.Read(STRPTR(strText), LEN(strText))
print strText
```
---

## ReadTextA

Reads a specified number of characters from the stream into memory, starting at the current seek pointer. Ansi version.

```
FUNCTION ReadTextA (BYVAL numChars AS LONG) AS STRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *numChars* | The number of characters to read from the stream.<br>Pass -1 to read all the characters from the current seek position. |

#### Return value

STRING. The characters read.

#### Example

```
#INCLUDE ONCE "AfxNova/CStream.inc"
USING AfxNova

DIM pstm AS CFileStream
pstm.Open(ExePath & "\TextA1.txt", STGM_READ)
DIM strText AS STRING = pstm.ReadTextA(-1)
PRINT strText
```
---

## ReadTextW

Reads a specified number of characters from the stream into memory, starting at the current seek pointer. Unicode version.

```
FUNCTION ReadTextW (BYVAL numChars AS LONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *numChars* | The number of characters to read from the stream.<br>Pass -1 to read all the characters from the current seek position. |

#### Return value

DWSTRING. The characters read.

#### Example

```
#INCLUDE ONCE "AfxNova/CStream.inc"
USING AfxNova

DIM pstm AS CFileStream
pstm.Open(ExePath & "\TextW1.txt", STGM_READ)
DIM dwsText AS DWSTRING = pstm.ReadTextW(-1)
PRINT dwsText
```
---

## Write

Writes a specified number of bytes into the stream starting at the current seek pointer.

```
FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbWritten AS ULONG PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer that contains the data that is to be written to the stream. A valid pointer must be provided for this parameter even when cb is zero. |
| *cb* | The number of bytes of data to attempt to write into the stream. This value can be zero. |
| *pcbWritten* | A pointer to a ULONG variable where this method writes the actual number of bytes written to the stream. The caller can set this pointer to NULL, in which case this method does not provide the actual number of bytes written. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

```
FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer that contains the data that is to be written to the stream. A valid pointer must be provided for this parameter even when cb is zero. |
| *cb* | The number of bytes of data to attempt to write into the stream. This value can be zero. |

#### Return value

ULONG. The actual number of bytes written to the stream. Note: The number of bytes read may be zero.

---

## WriteTextA

Writes a string at the current seek position. Ansi version.

```
FUNCTION WriteTextA (BYREF strText AS STRING) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *strText* | STRING. The string to write. |

#### Return value

ULONG. The characters written.

#### Example

```
#INCLUDE ONCE "AfxNova/CStream.inc"
USING AfxNova

DIM pstm AS CFileStream
pstm.Open(ExePath & "\TextA1.txt", STGM_READWRITE)
pstm.Seek(5, STREAM_SEEK_SET)
pstm.WriteTextA(" 12345 ")
pstm.Seek(0, STREAM_SEEK_SET)
DIM s AS STRING = pstm.ReadTextA(50)
print s
```
---

## WriteTextW

Writes a string at the current seek position. Unicode version.

```
FUNCTION WriteTextW (BYREF wszText AS WSTRING) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | WSTRING. The string to write. |

#### Return value

ULONG. The characters written.

#### Example

```
#INCLUDE ONCE "AfxNova/CStream.inc"
USING AfxNova

DIM pstm AS CFileStream
pstm.Open(ExePath & "\TextW1.txt", STGM_READWRITE)
pstm.Seek(2, STREAM_SEEK_SET)   ' Skip BOM
pstm.Seek(5 * 2, STREAM_SEEK_CUR)
pstm.WriteTextW(" 12345 ")
pstm.Seek(2, STREAM_SEEK_SET)   ' Skip BOM
DIM dws AS DWSTRING = pstm.ReadTextW(50)
print dws
```
---

## Seek

Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer.

```
FUNCTION Seek (BYVAL dlibMove AS ULONGINT, BYVAL dwOrigin AS DWORD, _
   BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dlibMove* | ULONGINT. The displacement to be added to the location indicated by the dwOrigin parameter. If *dwOrigin* is **STREAM_SEEK_SET**, this is interpreted as an unsigned value rather than a signed value. |
| *dwOrigin* | DWORD. The origin for the displacement specified in *dlibMove*. The origin can be the beginning of the file (**STREAM_SEEK_SET**), the current seek pointer (**STREAM_SEEK_CUR**), or the end of the file (**STREAM_SEEK_END**).<br>**STREAM_SEEK_SET** : The new seek pointer is an offset relative to the beginning of the stream. In this case, the *dlibMove* parameter is the new seek position relative to the beginning of the stream.<br>**STREAM_SEEK_CUR** : The new seek pointer is an offset relative to the current seek pointer location. In this case, the *dlibMove* parameter is the signed displacement from the current seek position.<br>**STREAM_SEEK_END** : The new seek pointer is an offset relative to the end of the stream. In this case, the *dlibMove* parameter is the new seek position relative to the end of the stream. |
| *plibNewPosition* | ULONGINT PTR. A pointer to the location where this method writes the value of the new seek pointer from the beginning of the stream. You can set this pointer to NULL. In this case, this method does not provide the new seek pointer. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## GetSeekPosition

Returns the current seek position.

```
FUNCTION GetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The current seek position.

---

## ResetSeekPosition

Sets the seek position at the beginning of the stream.

```
FUNCTION ResetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

---

## SeekAtEndOfFile

Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfFile () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

---

## SeekAtEndOfFile

Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfStream () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

---

## GetSize

Returns the size of the stream in bytes.

```
FUNCTION GetSize () AS ULONGINT
```

#### Return value

ULONGINT. The size of the stream in bytes.

---

## SetSize

Changes the size of the stream object.

```
FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *libNewSize* | ULONGINT. Specifies the new size, in bytes, of the stream. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## CopyTo

Copies a specified number of bytes from the current seek pointer in the stream to the current seek pointer in another stream.

```
FUNCTION CopyTo (BYVAL pstm AS IStream PTR, _
   BYVAL cb AS ULONGINT, _
   BYVAL pcbRead AS ULONGINT PTR = NULL, _
   BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstm* | A pointer to the destination stream. The stream pointed to by *pstm* can be a new stream or a clone of the source stream. |
| *cb* | The number of bytes of data to attempt to copy into the stream. |
| *pcbRead* | A pointer to the location where this method writes the actual number of bytes read from the source. You can set this pointer to NULL. In this case, this method does not provide the actual number of bytes read. |
| *pcbWritten* | A pointer to the location where this method writes the actual number of bytes written to the destination. You can set this pointer to NULL. In this case, this method does not provide the actual number of bytes written. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## LockRegion

Restricts access to a specified range of bytes in the stream.

```
FUNCTION LockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *libOffset* | ULONGINT. Specifies the byte offset for the beginning of the range. |
| *cb* | ULONGINT. Specifies the length of the range, in bytes, to be restricted. |
| *dwLockType* | DWORD. Specifies the restrictions being requested on accessing the range.<br>- **LOCK_WRITE** : If this lock is granted, the specified range of bytes can be opened and read any number of times, but writing to the locked range is prohibited except for the owner that was granted this lock.<br>- **LOCK_EXCLUSIVE** : If this lock is granted, writing to the specified range of bytes is prohibited except by the owner that was granted this lock.<br>- **LOCK_ONLYONCE** : If this lock is granted, no other LOCK_ONLYONCE lock can be obtained on the range. Usually this lock type is an alias for some other lock type. Thus, specific implementations can have additional behavior associated with this lock type.|

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## UnlockRegion

Unlocks a region previously locked with the **LockRegion** method. Locked regions must later be explicitly unlocked by calling **UnlockRegion** with exactly the same values for the *libOffset*, *cb*, and *dwLockType* parameters. The region must be unlocked before the stream is released. Two adjacent regions cannot be locked separately and then unlocked with a single unlock call.

```
FUNCTION UnlockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *libOffset* | ULONGINT. Specifies the byte offset for the beginning of the range. |
| *cb* | ULONGINT. Specifies the length of the range, in bytes, to be restricted. |
| *dwLockType* | DWORD. Specifies the access restrictions previously placed on the range.|

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## Stat

Retrieves the **STATSTG** structure for this stream. For a detailed description of this structure see [STATSTG structure](https://docs.microsoft.com/en-us/windows/desktop/api/objidl/ns-objidl-tagstatstg).

```
FUNCTION Stat (BYVAL pstatstg AS STATSTG PTR, BYVAL grfStatFlag AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstatstg* | Pointer to a **STATSTG** structure where this method places information about this stream. |
| *grfStatFlag* | DWORD. Specifies that this method does not return some of the members in the **STATSTG** structure, thus saving a memory allocation operation. Values are taken from the **STATFLAG** enumeration.<br>- **STATFLAG_DEFAULT** : Requests that the statistics include the *pwcsName* member of the **STATSTG** structure.<br>- **STATFLAG_NONAME** : Requests that the statistics not include the *pwcsName* member of the **STATSTG structure**. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

```
FUNCTION Stat (BYVAL grfStatFlag AS DWORD) AS STATSTG
```

| Parameter  | Description |
| ---------- | ----------- |
| *grfStatFlag* | DWORD. Specifies that this method does not return some of the members in the **STATSTG** structure, thus saving a memory allocation operation. Values are taken from the **STATFLAG** enumeration.<br>- **STATFLAG_DEFAULT** : Requests that the statistics include the *pwcsName* member of the **STATSTG** structure.<br>- **STATFLAG_NONAME** : Requests that the statistics not include the *pwcsName* member of the **STATSTG structure**. |

#### Return value

STATFLAG. The **STATSTG** structure for this stream.

---

## Clone

Creates a new stream with its own seek pointer that references the same bytes as the original stream. The **Clone** method creates a new stream for accessing the same bytes but using a separate seek pointer. The new stream sees the same data as the source-stream. Changes written to one stream are immediately visible in the other. Range locking is shared between the streams. The initial setting of the seek pointer in the cloned stream instance is the same as the current setting of the seek pointer in the original stream at the time of the clone operation.

```
FUNCTION Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *ppstm* | When successful, pointer to the location of an **IStream** pointer to the new stream. If an error occurs, this parameter is NULL. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## StreamPtr

Returns a pointer to the underlying **IStream** interface.

```
FUNCTION StreamPtr () AS IStreamPtr
```
---

