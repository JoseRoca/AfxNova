# CMemStream Class

Creates a memory stream, allowing read, write and seek operations. The stream is thread-safe as of Windows 8. On earlier systems, the stream is not thread-safe. Cloning is supported as of Windows 8.

**Include file**: AfxNova/CStream.inc

### Constructor

```
CONSTRUCTOR CMemStream (BYVAL pInit AS CONST BYTE PTR, BYVAL cbInit AS UINT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pInit* | A pointer to a buffer of size *cbInit*. The contents of this buffer are used to set the initial contents of the memory stream. If this parameter is NULL, the returned memory stream does not have any initial content. |
| *cbInit* | The number of bytes in the buffer pointed to by *pInit*. If *pInit* is set to NULL, *cbInit* must be zero. |

### Operator CAST

Returns a pointer to the underlying **IStream** interface of the stream object.

```
OPERATOR CAST () AS IStream PTR
```

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Read](#read1) | Reads a specified number of characters from the stream into memory, starting at the current seek pointer. |
| [Write](#write1) | Writes a specified number of bytes into the stream starting at the current seek pointer. |
| [Seek](#seek1) | Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer. |
| [GetSeekPosition](#getseekposition1) | Returns the seek position. |
| [ResetSeekPosition](#resetseekposition1) | Sets the seek position at the beginning of the stream. |
| [SeekAtEndOfStream](#seekatendofstream1) | Sets the seek position at the end of the stream. |
| [CopyTo](#copyto) | Copies a specified number of bytes from the current seek pointer in the stream to the current seek pointer in another stream. |
| [GetSize](#getsize1) | Returns the size of the stream. |
| [SetSize](#setsize1) | Changes the size of the stream. |
| [Clone](#clone1) | Creates a new stream with its own seek pointer that references the same bytes as the original stream. |
| [StreamPtr](#streamptr) | Returns a pointer to the underlying IStream interface. |
| [GetLastResult](#getlastresult1) | Returns the last result code. |
| [GetErrorInfo](#geterrorinfo1) | Returns a description of the last result code. |

---

# CMemTextStream Class

Creates a memory text stream, allowing read, write and seek operations. The stream is thread-safe as of Windows 8. On earlier systems, the stream is not thread-safe. Cloning is supported as of Windows 8.

**Include file**: Afx/NovaCStream.inc

### Constructor (CMemTextStream)

```
CONSTRUCTOR CMemTextStream
CONSTRUCTOR CMemTextStream (BYVAL pwszText AS CONST WSTRING PTR)
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszText* | Creates a memory text stream and initializes it with the content of a string. If this parameter is NULL, the returned memory stream does not have any initial content. |

### Operator CAST

Returns a pointer to the underlying **IStream** interface of the stream object.

```
OPERATOR CAST () AS IStream PTR
```

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Read](#read2) | Reads a specified number of characters from the stream into memory, starting at the current seek pointer. |
| [Write](#write2) | Writes a specified number of bytes into the stream starting at the current seek pointer. |
| [Append](#append2) | Appends a string at the end of the stream. |
| [Seek](#seek2) | Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer. |
| [GetSeekPosition](#getseekposition2) | Returns the seek position. |
| [ResetSeekPosition](#resetseekposition2) | Sets the seek position at the beginning of the stream. |
| [SeekAtEndOfFile](#seekatendoffile2) | Sets the seek position at the end of the stream. |
| [SeekAtEndOfStream](#seekatendofstream2) | Sets the seek position at the end of the stream. |
| [CopyTo](#copyto2) | Copies a specified number of characters from the current seek pointer in the stream to the current seek pointer in another stream. |
| [GetSize](#getsize2) | Returns the size of the stream. |
| [SetSize](#setsize2) | Changes the size of the stream. |
| [Clone](#Ccone2) | Creates a new stream with its own seek pointer that references the same bytes as the original stream. |
| [StreamPtr](#streamptr) | Returns a pointer to the underlying IStream interface. |
| [GetLastResult](#getlastresult2) | Returns the last result code. |
| [GetErrorInfo](#geterrorinfo2) | Returns a description of the last result code. |

---

## <a name="streamptr"></a>StreamPtr

Returns a pointer to the underlying **IStream** interface.

```
FUNCTION StreamPtr () AS IStreamPtr
```

#### Remarks

To save a memory stream to a file you can:

1. Create an instance of the **CFileStream** class
2. Get a pointer to its underlying **IStream** interface with **CFileStream.StreamPtr**
3. Copy the contents of the memory stream to the file stream with **CMemStream.CopyTo**
5. Close the file with **CFileStream.Close**

---

## <a name="read1"></a>Read (CMemStream)

Reads a specified number of bytes from the stream into memory, starting at the current seek pointer.

```
FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer which the stream data is read into. |
| *cb* | The number of bytes of data to read from the stream. |
| *pcbRead* | A pointer to a ULONG variable that receives the actual number of bytes read from the stream. The number of bytes read may be zero. |

#### Return value

S_OK (0) or an HRESULT code.

```
FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer which the stream data is read into. |
| *cb* | The number of bytes of data to read from the stream. |

ULONG. The number characters read.

---

## <a name="write1"></a>Write (CMemStream)

Writes a specified number of bytes into the stream starting at the current seek pointer.

```
FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbWritten AS ULONG PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer that contains the data that is to be written to the stream. A valid pointer must be provided for this parameter even when *cb* is zero. |
| *cb* | The number of bytes of data to attempt to write into the stream. This value can be zero. |
| *pcbWritten* | A pointer to a ULONG variable where this method writes the actual number of bytes written to the stream. The caller can set this pointer to NULL, in which case this method does not provide the actual number of bytes written. |

#### Return value

S_OK (0) or an HRESULT code.

```
FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *pv* | A pointer to the buffer that contains the data that is to be written to the stream. A valid pointer must be provided for this parameter even when *cb* is zero. |
| *cb* | The number of bytes of data to attempt to write into the stream. This value can be zero. |

#### Return value

The number of characters actually written.

---

## <a name="seek1"></a>Seek (CMemStream)

Changes the seek pointer to a new location. The new location is relative to either the beginning of the stream, the end of the stream, or the current seek pointer.

```
FUNCTION Seek (BYVAL dlibMove AS ULONGINT, BYVAL dwOrigin AS DWORD, _
   BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *dlibMove* | The displacement to be added to the location indicated by the *dwOrigin* parameter. If *dwOrigin* is STREAM_SEEK_SET, this is interpreted as an unsigned value rather than a signed value. |
| *dwOrigin* | The origin for the displacement specified in *dlibMove*. The origin can be the beginning of the file (STREAM_SEEK_SET), the current seek pointer (STREAM_SEEK_CUR), or the end of the file (STREAM_SEEK_END). For more information about values, see the STREAM_SEEK enumeration. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## <a name="getseekposition1"></a>GetSeekPosition (CMemStream)

Returns the current seek position.

```
FUNCTION GetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The current seek position.

---

## <a name="resetseekposition1"></a>ResetSeekPosition (CMemStream)

Sets the seek position at the beginning of the stream.

```
FUNCTION ResetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

---

## <a name="seekatendofstream1"></a>SeekAtEndOfStream (CMemStream)
Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfStream () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

---

## <a name="getsize1"></a>GetSize (CMemStream)

Returns the size of the stream in bytes.

```
FUNCTION GetSize () AS ULONGINT
```

#### Return value

ULONGINT. The size of the stream in bytes.

---

## <a name="setsize1"></a>SetSize (CMemStream)

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

## <a name="copyto"></a>CopyTo (CMemStream)

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

## <a name="clone1"></a>Clone (CMemStream)

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

## <a name="getlastresult1"></a>GetLastResult (CMemStream)

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## <a name="geterrorinfo1"></a>GetErrorInfo (CMemStream)

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS CWSTR
```

#### Return value

A description of the last result code. If the result code is S_OK (0), it returns "Success"; otherwise, it returns the hexadecimal value of the error code and a description such "Seek error", "Write fault", "Read fault" or "Invalid argument".

---

## <a name="read2"></a>Read (CMemTextStream)

Reads a specified number of characters from the stream into memory, starting at the current seek pointer.

```
FUNCTION Read (BYVAL numChars AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *numChars* | The number of characters to read from the stream. Pass -1 to read all the characters from the current seek position. |

#### Return value

The characters read.

---

## <a name="write2"></a>Write (CMemTextStream)

Writes a string at the current seek position.

```
FUNCTION Write (BYREF wszText AS CONST WSTRING) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | The string to write. |

#### Return value

The number of characters actually written.

---

## <a name="append2"></a>Append (CMemTextStream)

Appends a string at the end of the stream.

```
FUNCTION Append (BYREF wszText AS CONST WSTRING) AS ULONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszText* | The string to append. |

#### Return value

The number of characters actually written.

---

## <a name="seek2"></a>Seek (CMemTextStream)

Sets the seek position as an absolute position from the start of the stream.

```
FUNCTION Seek (Seek (BYVAL nPos AS ULONGINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nPos* | The new seek position (from 1 to the end of the stream). |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## <a name="getseekposition2"></a>GetSeekPosition (CMemTextStream)

Returns the current seek position.

```
FUNCTION GetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The current seek position.

---

## <a name="resetseekposition2"></a>ResetSeekPosition (CMemTextStream)

Sets the seek position at the beginning of the stream.

```
FUNCTION ResetSeekPosition () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

---

## <a name="seekatendofstream2"></a>SeekAtEndOfStream (CMemTextStream)

Sets the seek position at the end of the stream.

```
FUNCTION SeekAtEndOfStream () AS ULONGINT
```

#### Return value

ULONGINT. The new seek position.

---

## <a name="getsize2"></a>GetSize (CMemTextStream)

Returns the size of the stream in characters.

```
FUNCTION GetSize () AS ULONGINT
```

#### Return value

ULONGINT. The size of the stream in bytes.

---

## <a name="setsize2"></a>SetSize (CMemTextStream)

Changes the size of the stream object.

```
FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *libNewSize* | ULONGINT. Specifies the new size, in characters, of the stream. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## <a name="copyto2"></a>CopyTo (CMemTextStream)

Copies a specified number of characters from the current seek pointer in the stream to the current seek pointer in another stream.

```
FUNCTION CopyTo (BYVAL pstm AS IStream PTR, _
   BYVAL cb AS ULONGINT, _
   BYVAL pcbRead AS ULONGINT PTR = NULL, _
   BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pstm* | A pointer to the destination stream. The stream pointed to by *pstm* can be a new stream or a clone of the source stream. |
| *cb* | The number of characters of data to attempt to copy into the stream. |
| *pcbRead* | A pointer to the location where this method writes the actual number of bytes read from the source. You can set this pointer to NULL. In this case, this method does not provide the actual number of bytes read. |
| *pcbWritten* | A pointer to the location where this method writes the actual number of bytes written to the destination. You can set this pointer to NULL. In this case, this method does not provide the actual number of bytes written. |

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## <a name="clone2"></a>Clone (CMemTextStream)

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

## <a name="getlastresult2"></a>GetLastResult (CMemTextStream)

Returns the last result code.

```
FUNCTION GetLastResult () AS HRESULT
```

#### Return value

HRESULT. S_OK (0) on success, or an error code on failure.

---

## <a name="geterrorinfo2"></a>GetErrorInfo (CMemTextStream)

Returns a description of the last result code.

```
FUNCTION GetErrorInfo () AS CWSTR
```

#### Return value

A description of the last result code. If the result code is S_OK (0), it returns "Success"; otherwise, it returns the hexadecimal value of the error code and a description such "Seek error", "Write fault", "Read fault" or "Invalid argument".

---
