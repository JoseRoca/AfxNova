' ########################################################################################
' Platform: Microsoft Windows
' Filename: CStream.inc
' Purpose:  Classes to work with streams.
' Compiler: Wrapper class for Microsoft WinHTTP Services, version 5.1
' Copyright (c) 2025 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "AfxNova/DWSTRING.inc"

NAMESPACE AfxNova

' ========================================================================================
' Opens or creates a file and retrieves a stream to read or write to that file.
' ========================================================================================
PRIVATE FUNCTION AfxSHCreateStreamOnFileEx (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD, _
BYVAL dwAttributes AS DWORD, BYVAL fCreate AS BOOLEAN, BYVAL ppStream AS IStream PTR PTR) AS HRESULT
   ' // See if the library is already loaded in the address space
   DIM AS ANY PTR pLib = GetModuleHandleW("shlwapi.dll")
   ' // If it is not loaded, load it
   IF pLib = NULL THEN pLib = DyLibLoad("shlwapi.dll")
   IF pLib = NULL THEN RETURN E_POINTER
   DIM pShCreateStreamOnFileEx AS FUNCTION (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD, BYVAL dwAttributes AS DWORD, _
       BYVAL fCreate AS BOOLEAN, BYVAL pstmTemplate AS IStream PTR, BYVAL ppStream AS IStream PTR PTR) AS HRESULT
   pShCreateStreamOnFileEx = DyLibSymbol(pLib, "SHCreateStreamOnFileEx")
   IF pShCreateStreamOnFileEx = NULL THEN RETURN E_POINTER
   DIM hr AS HRESULT = pShCreateStreamOnFileEx(pwszFile, grfMode, dwAttributes, fCreate, NULL, ppStream)
   ' For some reason, if we free the library, calling the methods that use the returned
   ' IStream pointer will GPF if compiled with the -O switch.
'   DyLibFree(pLib)
   RETURN hr
END FUNCTION
' ========================================================================================

' ========================================================================================
' CFileStream class
' ========================================================================================
TYPE CFileStream

Public:
   m_Result AS HRESULT
   m_pStream AS IStream PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYREF wszFile AS WSTRING, BYVAL grfMode AS DWORD = STGM_READ, _
           BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD = STGM_READ, _
           BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS BOOLEAN = FALSE)
   DECLARE CONSTRUCTOR (BYVAL pstm AS IStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   DECLARE OPERATOR LET (BYVAL pstm AS IStream PTR)
   DECLARE OPERATOR CAST () AS IStream PTR
   DECLARE FUNCTION StreamPtr () AS IStream PTR
   DECLARE FUNCTION Attach (BYVAL pstm AS IStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Detach () AS IStream PTR
   DECLARE FUNCTION Open (BYREF wszFile AS WSTRING, BYVAL grfMode AS DWORD = STGM_READ, _
           BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Open (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD = STGM_READ, _
           BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE SUB Close
   DECLARE FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
   DECLARE FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   DECLARE FUNCTION ReadTextA (BYVAL numChars AS LONG) AS STRING
   DECLARE FUNCTION ReadTextW (BYVAL numChars AS LONG) AS DWSTRING
   DECLARE FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbWritten AS ULONG PTR) AS HRESULT
   DECLARE FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   DECLARE FUNCTION WriteTextA (BYREF strText AS STRING) AS ULONG
   DECLARE FUNCTION WriteTextW (BYREF wszText AS WSTRING) AS ULONG
   DECLARE FUNCTION Seek (BYVAL dlibMove AS ULONGINT, BYVAL dwOrigin AS DWORD, BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
   DECLARE FUNCTION GetSeekPosition () AS ULONGINT
   DECLARE FUNCTION ResetSeekPosition () AS ULONGINT
   DECLARE FUNCTION SeekAtEndOfFile () AS ULONGINT
   DECLARE FUNCTION SeekAtEndOfStream () AS ULONGINT
   DECLARE FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
   DECLARE FUNCTION GetSize () AS ULONGINT
   DECLARE FUNCTION CopyTo (BYVAL pstm AS IStream PTR, BYVAL cb AS ULONGINT, _
           BYVAL pcbRead AS ULONGINT PTR = NULL, BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
   DECLARE FUNCTION LockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
   DECLARE FUNCTION UnlockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
   DECLARE FUNCTION Stat (BYVAL pstatstg AS STATSTG PTR, BYVAL grfStatFlag AS DWORD) AS HRESULT
   DECLARE FUNCTION Stat (BYVAL grfStatFlag AS DWORD) AS STATSTG
   DECLARE FUNCTION Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
   DECLARE FUNCTION Clone () AS IStream PTR

END TYPE
' ========================================================================================

' ========================================================================================
' CFileStream constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CFileStream
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Opens or creates a file and retrieves a stream to read or write to that file.
' Parameters:
' - pwszFile : A pointer to a unicode null-terminated string that specifies the file name.
' - grfMode : One or more STGM values that are used to specify the file access mode and how
'   the object that exposes the stream is created and deleted.
'   STGM constants: See https://docs.microsoft.com/en-us/windows/desktop/stg/stgm-constants
' - dwAttributes : One or more flag values that specify file attributes in the case that a
'   new file is created. For a complete list of possible values, see the dwFlagsAndAttributes
'   parameter of the CreateFile function.
'   See: https://docs.microsoft.com/en-us/windows/desktop/api/fileapi/nf-fileapi-createfilea
' - fCreate : A BOOL value that helps specify, in conjunction with grfMode, how existing
'   files should be treated when creating the stream. See Remarks for details.
'   https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-shcreatestreamonfileex
' ========================================================================================
PRIVATE CONSTRUCTOR CFileStream (BYREF wszFile AS WSTRING, BYVAL grfMode AS DWORD = STGM_READ, _
BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS BOOLEAN = FALSE)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   this.SetResult(AfxSHCreateStreamOnFileEx(@wszFile, grfMode, dwAttributes, fCreate, @m_pStream))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CFileStream (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD = STGM_READ, _
BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS BOOLEAN = FALSE)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   this.SetResult(AfxSHCreateStreamOnFileEx(pwszFile, grfMode, dwAttributes, fCreate, @m_pStream))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes an instance of the class from an existing IStream object and attaches it.
' ========================================================================================
PRIVATE CONSTRUCTOR CFileStream (BYVAL pstm AS IStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   IF pstm = NULL THEN EXIT CONSTRUCTOR
   IF fAddRef THEN pstm->lpvtbl->AddRef(pstm)
   m_pStream = pstm
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CFileStream destructor
' ========================================================================================
PRIVATE DESTRUCTOR CFileStream
   ' // Release the IStrean interface
   IF m_pStream THEN m_pStream->lpvtbl->Release(m_pStream)
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileStream.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileStream.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileStream.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   IF nError = -1 THEN nError = m_Result
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      dwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   IF nError THEN dwsMsg = "Error &h" + HEX(nError) & CHR(13, 10) + dwsMsg
   RETURN dwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Assignment operator
' ========================================================================================
PRIVATE OPERATOR CFileStream.LET (BYVAL pstm AS IStream PTR)
   m_Result = 0
   IF pstm = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the stream object
   IF m_pStream THEN m_pStream->lpvtbl->Release(m_pStream)
   ' // Attach the passed stream object to the class
   m_pStream = pstm
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CFileStream.CAST () AS IStream PTR
   m_Result = 0
   OPERATOR = m_pStream
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying IStream interface
' ========================================================================================
PRIVATE FUNCTION CFileStream.StreamPtr () AS IStream PTR
   m_Result = 0
   RETURN m_pStream
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches the passed stream object to the class
' ========================================================================================
PRIVATE FUNCTION CFileStream.Attach (BYVAL pstm AS IStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   m_Result = 0
   IF pstm = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the stream object
   IF m_pStream THEN m_Result = m_pStream->lpvtbl->Release(m_pStream)
   ' // Attach the passed stream object to the class
   IF fAddRef THEN pstm->lpvtbl->AddRef(pstm)
   m_pStream = pstm
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the stream object from the class
' ========================================================================================
PRIVATE FUNCTION CFileStream.Detach () AS IStream PTR
   m_Result = 0
   DIM pstm AS IStream PTR = m_pStream
   m_pStream = NULL
   RETURN pstm
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens or creates a file and retrieves a stream to read or write to that file.
' Parameters:
' - pwszFile : A pointer to a unicode null-terminated string that specifies the file name.
' - grfMode : One or more STGM values that are used to specify the file access mode and how
'   the object that exposes the stream is created and deleted.
'   STGM constants: See https://docs.microsoft.com/en-us/windows/desktop/stg/stgm-constants
' - dwAttributes : One or more flag values that specify file attributes in the case that a
'   new file is created. For a complete list of possible values, see the dwFlagsAndAttributes
'   parameter of the CreateFile function.
'   See: https://docs.microsoft.com/en-us/windows/desktop/api/fileapi/nf-fileapi-createfilea
' - fCreate : A BOOL value that helps specify, in conjunction with grfMode, how existing
'   files should be treated when creating the stream. See Remarks for details.
'   https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-shcreatestreamonfileex
' ========================================================================================
PRIVATE FUNCTION CFileStream.Open (BYREF wszFile AS WSTRING, BYVAL grfMode AS DWORD = STGM_READ, _
BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pStream <> NULL THEN m_pStream->lpvtbl->Release(m_pStream) : m_pStream = NULL
   RETURN this.SetResult(AfxSHCreateStreamOnFileEx(@wszFile, grfMode, dwAttributes, fCreate, @m_pStream))
END FUNCTION
' ========================================================================================
PRIVATE FUNCTION CFileStream.Open (BYVAL pwszFile AS WSTRING PTR, BYVAL grfMode AS DWORD = STGM_READ, _
BYVAL dwAttributes AS DWORD = FILE_ATTRIBUTE_NORMAL, BYVAL fCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pStream <> NULL THEN m_pStream->lpvtbl->Release(m_pStream) : m_pStream = NULL
   RETURN this.SetResult(AfxSHCreateStreamOnFileEx(pwszFile, grfMode, dwAttributes, fCreate, @m_pStream))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Releases the stream object.
' ========================================================================================
PRIVATE SUB CFileStream.Close
   ' // Release the IStrean interface
   IF m_pStream THEN m_pStream->lpvtbl->Release(m_pStream)
   m_pStream = NULL
   m_Result = 0
END SUB
' ========================================================================================

' ========================================================================================
' Reads a specified number of bytes from the stream into memory, starting at the
' current seek pointer.
' - pv : A pointer to the buffer which the stream data is read into.
' - cb : The number of bytes of data to read from the stream.
' - pcbRead: A pointer to a ULONG variable that receives the actual number of bytes read
'   from the stream. Note: The number of bytes read may be zero.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Read(m_pStream, pv, cb, pcbRead))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER)
   DIM cbRead AS ULONG
   this.SetResult(m_pStream->lpvtbl->Read(m_pStream, pv, cb, @cbRead))
   RETURN cbRead
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads a specified number of characters from the stream into memory, starting at the
' current seek pointer.
' - numChars : The number of characters to read from the stream.
'   Pass -1 to read all the characters from the current seek position.
' ========================================================================================
PRIVATE FUNCTION CFileStream.ReadTextA (BYVAL numChars AS LONG) AS STRING
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN ""
   DIM cbRead AS ULONG
   IF numChars < 0 THEN numChars = this.GetSize - this.GetSeekPosition
   DIM strText AS STRING = SPACE(numChars)
   this.SetResult(m_pStream->lpvtbl->Read(m_pStream, STRPTR(strText), numChars, @cbRead))
   IF cbRead < LEN(strText) THEN strText = LEFT(strText, cbRead)
   RETURN strText
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.ReadTextW (BYVAL numChars AS LONG) AS DWSTRING
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN ""
   DIM cbRead AS ULONG
   IF numChars < 0 THEN numChars = (this.GetSize \ 2) - this.GetSeekPosition * 2
   DIM dwsText AS DWSTRING = WSPACE(numChars)
   this.SetResult(m_pStream->lpvtbl->Read(m_pStream, dwsText, numChars * 2, @cbRead))
   IF cbRead < LEN(dwsText) THEN dwsText = LEFT(dwsText, cbRead)
   RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a specified number of bytes into the stream starting at the current seek pointer.
' - pv: A pointer to the buffer that contains the data that is to be written to the stream.
'   A valid pointer must be provided for this parameter even when cb is zero.
' - cb : The number of bytes of data to attempt to write into the stream. This value can be zero.
' - pcbWritten : A pointer to a ULONG variable where this method writes the actual number
'   of bytes written to the stream. The caller can set this pointer to NULL, in which
'   case this method does not provide the actual number of bytes written.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbWritten AS ULONG PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   IF pv = NULL THEN RETURN this.SetResult(E_INVALIDARG)
   RETURN this.SetResult(m_pStream->lpvtbl->Write(m_pStream, pv, cb, pcbWritten))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER)
   IF pv = NULL THEN this.SetResult(E_INVALIDARG)
   DIM cbWritten AS ULONG
   this.SetResult(m_pStream->lpvtbl->Write(m_pStream, pv, cb, @cbWritten))
   RETURN cbWritten
END FUNCTION
' ========================================================================================
' ========================================================================================
' Writes a string at the current seek position.
' ========================================================================================
PRIVATE FUNCTION CFileStream.WriteTextA (BYREF strText AS STRING) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM cbWritten AS ULONG
   this.SetResult(m_pStream->lpvtbl->Write(m_pStream, STRPTR(strText), LEN(strText), @cbWritten))
   RETURN cbWritten
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.WriteTextW (BYREF wszText AS WSTRING) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM cbWritten AS ULONG
   this.SetResult(m_pStream->lpvtbl->Write(m_pStream, STRPTR(wszText), LEN(wszText) * 2, @cbWritten))
   RETURN cbWritten \ 2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the seek pointer to a new location. The new location is relative to either the
' beginning of the stream, the end of the stream, or the current seek pointer.
' - dlibMove : The displacement to be added to the location indicated by the dwOrigin
'   parameter. If dwOrigin is STREAM_SEEK_SET, this is interpreted as an unsigned value
'   rather than a signed value.
' - dwOrigin : The origin for the displacement specified in dlibMove. The origin can be the
'   beginning of the file (STREAM_SEEK_SET), the current seek pointer (STREAM_SEEK_CUR), or
'   the end of the file (STREAM_SEEK_END). For more information about values, see the
'   STREAM_SEEK enumeration.
' - plibNewPosition : A pointer to the location where this method writes the value of the
'   new seek pointer from the beginning of the stream. You can set this pointer to NULL.
'   In this case, this method does not provide the new seek pointer.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Seek (BYVAL dlibMove AS ULONGINT, BYVAL dwOrigin AS DWORD, BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM lidlibMove AS LARGE_INTEGER : lidlibMove.QuadPart = dlibMove
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, lidlibMove, dwOrigin, @libNewPosition))
   IF plibNewPosition THEN *plibNewPosition = libNewPosition.QuadPart
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the current seek position.
' ========================================================================================
PRIVATE FUNCTION CFileStream.GetSeekPosition () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_CUR, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the seek position at the beginning of the stream.
' ========================================================================================
PRIVATE FUNCTION CFileStream.ResetSeekPosition () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_SET, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the seek position at the end of the stream.
' ========================================================================================
PRIVATE FUNCTION CFileStream.SeekAtEndOfFile () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_END, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the seek position at the end of the stream.
' ========================================================================================
PRIVATE FUNCTION CFileStream.SeekAtEndOfStream () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_END, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the size of the stream.
' libNewSize : Specifies the new size, in bytes, of the stream.
' ========================================================================================
PRIVATE FUNCTION CFileStream.SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulilibNewSize AS ULARGE_INTEGER : ulilibNewSize.QuadPart = libNewSize
   RETURN this.SetResult(m_pStream->lpvtbl->SetSize(m_pStream, ulilibNewSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the size of the stream in bytes.
' ========================================================================================
PRIVATE FUNCTION CFileStream.GetSize () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM tstatstg AS STATSTG
   this.SetResult(m_pStream->lpvtbl->Stat(m_pStream, @tstatstg, STATFLAG_NONAME))
   RETURN tstatstg.cbSize.QuadPart
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies a specified number of bytes from the current seek pointer in the stream to the
' current seek pointer in another stream.
' - pstm : A pointer to the destination stream. The stream pointed to by pstm can be a new
'   stream or a clone of the source stream.
' - cb : The number of bytes of data to attempt to copy into the stream.
' - pcbRead : A pointer to the location where this method writes the actual number of bytes
'   read from the source. You can set this pointer to NULL. In this case, this method does
'   not provide the actual number of bytes read.
' - pcbWritten : A pointer to the location where this method writes the actual number of
'   bytes written to the destination. You can set this pointer to NULL. In this case, this
'   method does not provide the actual number of bytes written.
' ========================================================================================
PRIVATE FUNCTION CFileStream.CopyTo (BYVAL pstm AS IStream PTR, BYVAL cb AS ULONGINT, _
BYVAL pcbRead AS ULONGINT PTR = NULL, BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulicb AS ULARGE_INTEGER: ulicb.QuadPart = cb
   DIM ulicbRead AS ULARGE_INTEGER, ulicbWritten AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->CopyTo(m_pStream, pstm, ulicb, @ulicbRead, @ulicbWritten))
   IF pcbRead THEN *pcbRead = ulicbRead.QuadPart
   IF pcbWritten THEN *pcbWritten = ulicbWritten.QuadPart
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Restricts access to a specified range of bytes in the stream.
' - libOffset : Integer that specifies the byte offset for the beginning of the range.
' - cb : Integer that specifies the length of the range, in bytes, to be restricted.
' - dwLockType : Specifies the restrictions being requested on accessing the range.
'   - LOCK_WRITE: If this lock is granted, the specified range of bytes can be opened and
'     read any number of times, but writing to the locked range is prohibited except for the
'     owner that was granted this lock.
'   - LOCK_EXCLUSIVE : If this lock is granted, writing to the specified range of bytes is
'     prohibited except by the owner that was granted this lock.
'   - LOCK_ONLYONCE: If this lock is granted, no other LOCK_ONLYONCE lock can be obtained
'     on the range. Usually this lock type is an alias for some other lock type. Thus,
'     specific implementations can have additional behavior associated with this lock type.
' ========================================================================================
PRIVATE FUNCTION CFileStream.LockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulilibOffset AS ULARGE_INTEGER : ulilibOffset.QuadPart = libOffset
   DIM ulicb AS ULARGE_INTEGER : ulicb.QuadPart = cb
   RETURN this.SetResult(m_pStream->lpvtbl->LockRegion(m_pStream, ulilibOffset, ulicb, dwLockType))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes the access restriction on a range of bytes previously restricted with IStream.LockRegion.
' - libOffset : Integer that specifies the byte offset for the beginning of the range.
' - cb : Integer that specifies the length of the range, in bytes, to be restricted.
' - dwLockType : Specifies the access restrictions previously placed on the range.
' Remarks:
'   IStream.UnlockRegion unlocks a region previously locked with the IStream.LockRegion
'   method. Locked regions must later be explicitly unlocked by calling IStream.UnlockRegion
'   with exactly the same values for the libOffset, cb, and dwLockType parameters. The
'   region must be unlocked before the stream is released. Two adjacent regions cannot be
'   locked separately and then unlocked with a single unlock call.
' ========================================================================================
PRIVATE FUNCTION CFileStream.UnlockRegion (BYVAL libOffset AS ULONGINT, BYVAL cb AS ULONGINT, BYVAL dwLockType AS DWORD) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulilibOffset AS ULARGE_INTEGER : ulilibOffset.QuadPart = libOffset
   DIM ulicb AS ULARGE_INTEGER : ulicb.QuadPart = cb
   RETURN this.SetResult(m_pStream->lpvtbl->UnlockRegion(m_pStream, ulilibOffset, ulicb, dwLockType))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the STATSTG structure for this stream.
' - pstatstg : Pointer to a STATSTG structure where this method places information about
'   this stream.
' - grfStatFlag : Specifies that this method does not return some of the members in the
'   STATSTG structure, thus saving a memory allocation operation. Values are taken from
'   the STATFLAG enumeration.
'   - STATFLAG_DEFAULT : Requests that the statistics include the pwcsName member of the
'     STATSTG structure.
'   - STATFLAG_NONAME : Requests that the statistics not include the pwcsName member of
'     the STATSTG structure.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Stat (BYVAL pstatstg AS STATSTG PTR, BYVAL grfStatFlag AS DWORD) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Stat(m_pStream, pstatstg, grfStatFlag))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.Stat (BYVAL grfStatFlag AS DWORD) AS STATSTG
   DIM tstatstg AS STATSTG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN tstatstg
   this.SetResult(m_pStream->lpvtbl->Stat(m_pStream, @tstatstg, grfStatFlag))
   RETURN tstatstg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new stream with its own seek pointer that references the same bytes as
' the original stream.
' - ppstm : When successful, pointer to the location of an IStream pointer to the new stream.
'   If an error occurs, this parameter is NULL.
' Remarks:
'    The Clone method creates a new stream for accessing the same bytes but using a
'    separate seek pointer. The new stream sees the same data as the source-stream.
'    Changes written to one stream are immediately visible in the other. Range
'    locking is shared between the streams.
'    The initial setting of the seek pointer in the cloned stream instance is the same as
'    the current setting of the seek pointer in the original stream at the time of the
'    clone operation.
' ========================================================================================
PRIVATE FUNCTION CFileStream.Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Clone(m_pStream, ppstm))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileStream.Clone () AS IStream PTR
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN NULL
   DIM ppstm AS ISTREAM PTR
   this.SetResult(m_pStream->lpvtbl->Clone(m_pStream, @ppstm))
   RETURN ppstm
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a memory stream.
' ========================================================================================
PRIVATE FUNCTION AfxSHCreateMemStream  (BYVAL pInit AS CONST BYTE PTR, BYVAL cbInit AS UINT) AS IStream PTR
   ' // See if the library is already loaded in the address space
   DIM AS ANY PTR pLib = GetModuleHandleW("shlwapi.dll")
   ' // If it is not loaded, load it
   IF pLib = NULL THEN pLib = DyLibLoad("shlwapi.dll")
   IF pLib = NULL THEN RETURN NULL
   DIM pSHCreateMemStream AS FUNCTION (BYVAL pInit AS CONST BYTE PTR, BYVAL cbInit AS UINT) AS IStream PTR
   pSHCreateMemStream = DyLibSymbol(pLib, "SHCreateMemStream")
   IF pSHCreateMemStream = NULL THEN RETURN NULL
   DIM pStream AS IStream PTR = pSHCreateMemStream(pInit, cbInit)
   ' For some reason, if we free the library, calling the methods that use the returned
   ' IStream pointer will GPF if compiled with the -O switch.
'   DyLibFree(pLib)
   RETURN pStream
END FUNCTION
' ========================================================================================

' ========================================================================================
' CMemStream class - Creates a memory stream.
' ========================================================================================
TYPE CMemStream

Public:
   m_Result AS HRESULT
   m_pStream AS IStream PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pInit AS CONST BYTE PTR, BYVAL cbInit AS UINT)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   DECLARE OPERATOR CAST () AS IStream PTR
   DECLARE FUNCTION StreamPtr () AS IStream PTR
   DECLARE FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
   DECLARE FUNCTION Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   DECLARE FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbWritten AS ULONG PTR) AS HRESULT
   DECLARE FUNCTION Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   DECLARE FUNCTION Seek (BYVAL dlibMove AS ULONGINT, BYVAL dwOrigin AS DWORD, BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
   DECLARE FUNCTION Seek (BYVAL dlibMove AS ULONGINT) AS HRESULT
   DECLARE FUNCTION GetSeekPosition () AS ULONGINT
   DECLARE FUNCTION ResetSeekPosition () AS ULONGINT
   DECLARE FUNCTION SeekAtEndOfStream () AS ULONGINT
   DECLARE FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
   DECLARE FUNCTION GetSize () AS ULONGINT
   DECLARE FUNCTION CopyTo (BYVAL pstm AS IStream PTR, BYVAL cb AS ULONGINT, _
           BYVAL pcbRead AS ULONGINT PTR = NULL, BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
   DECLARE FUNCTION Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
   DECLARE FUNCTION Clone () AS IStream PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Creates a memory stream.
' ========================================================================================
PRIVATE CONSTRUCTOR CMemStream
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   m_pStream = AfxSHCreateMemStream(NULL, 0)
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CMemStream (BYVAL pInit AS CONST BYTE PTR, BYVAL cbInit AS UINT)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   m_pStream = AfxSHCreateMemStream(pInit, cbInit)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CMemStream destructor
' ========================================================================================
PRIVATE DESTRUCTOR CMemStream
   ' // Release the IStrean interface
   IF m_pStream THEN m_pStream->lpvtbl->Release(m_pStream)
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CMemStream.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CMemStream.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CMemStream.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   IF nError = -1 THEN nError = m_Result
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      dwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   IF nError THEN dwsMsg = "Error &h" + HEX(nError) & CHR(13, 10) + dwsMsg
   RETURN dwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CMemStream.CAST () AS IStream PTR
   m_Result = 0
   OPERATOR = m_pStream
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying IStream interface
' ========================================================================================
PRIVATE FUNCTION CMemStream.StreamPtr () AS IStream PTR
   m_Result = 0
   RETURN m_pStream
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads a specified number of bytes from the stream into memory, starting at the
' current seek pointer.
' - pv : A pointer to the buffer which the stream data is read into.
' - cb : The number of bytes of data to read from the stream.
' - pcbRead: A pointer to a ULONG variable that receives the actual number of bytes read
'   from the stream. Note: The number of bytes read may be zero.
' ========================================================================================
PRIVATE FUNCTION CMemStream.Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbRead AS ULONG PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Read(m_pStream, pv, cb, pcbRead))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemStream.Read (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER)
   DIM cbRead AS ULONG
   this.SetResult(m_pStream->lpvtbl->Read(m_pStream, pv, cb, @cbRead))
   RETURN cbRead
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a specified number of bytes into the stream starting at the current seek pointer.
' - pv: A pointer to the buffer that contains the data that is to be written to the stream.
'   A valid pointer must be provided for this parameter even when cb is zero.
' - cb : The number of bytes of data to attempt to write into the stream. This value can be zero.
' - pcbWritten : A pointer to a ULONG variable where this method writes the actual number
'   of bytes written to the stream. The caller can set this pointer to NULL, in which
'   case this method does not provide the actual number of bytes written.
' ========================================================================================
PRIVATE FUNCTION CMemStream.Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG, BYVAL pcbWritten AS ULONG PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   IF pv = NULL THEN RETURN this.SetResult(E_INVALIDARG)
   RETURN this.SetResult(m_pStream->lpvtbl->Write(m_pStream, pv, cb, pcbWritten))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemStream.Write (BYVAL pv AS ANY PTR, BYVAL cb AS ULONG) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER)
   IF pv = NULL THEN this.SetResult(E_INVALIDARG)
   DIM cbWritten AS ULONG
   this.SetResult(m_pStream->lpvtbl->Write(m_pStream, pv, cb, @cbWritten))
   RETURN cbWritten
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the seek pointer to a new location. The new location is relative to either the
' beginning of the stream, the end of the stream, or the current seek pointer.
' - dlibMove : The displacement to be added to the location indicated by the dwOrigin
'   parameter. If dwOrigin is STREAM_SEEK_SET, this is interpreted as an unsigned value
'   rather than a signed value.
' - dwOrigin : The origin for the displacement specified in dlibMove. The origin can be the
'   beginning of the file (STREAM_SEEK_SET), the current seek pointer (STREAM_SEEK_CUR), or
'   the end of the file (STREAM_SEEK_END). For more information about values, see the
'   STREAM_SEEK enumeration.
' - plibNewPosition : A pointer to the location where this method writes the value of the
'   new seek pointer from the beginning of the stream. You can set this pointer to NULL.
'   In this case, this method does not provide the new seek pointer.
' ========================================================================================
PRIVATE FUNCTION CMemStream.Seek (BYVAL dlibMove AS ULONGINT, BYVAL dwOrigin AS DWORD, BYVAL plibNewPosition AS ULONGINT PTR = NULL) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM lidlibMove AS LARGE_INTEGER : lidlibMove.QuadPart = dlibMove
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, lidlibMove, dwOrigin, @libNewPosition))
   IF plibNewPosition THEN *plibNewPosition = libNewPosition.QuadPart
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the current seek position.
' ========================================================================================
PRIVATE FUNCTION CMemStream.GetSeekPosition () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_CUR, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the seek position at the beginning of the stream.
' ========================================================================================
PRIVATE FUNCTION CMemStream.ResetSeekPosition () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_SET, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the seek position at the end of the stream.
' ========================================================================================
PRIVATE FUNCTION CMemStream.SeekAtEndOfStream () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_END, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the seek position as an absolute position from the start of the stream.
' nPos ranges from 1 to the end of the strream.
' ========================================================================================
PRIVATE FUNCTION CMemStream.Seek (BYVAL nPos AS ULONGINT) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM lidlibMove AS LARGE_INTEGER : lidlibMove.QuadPart = (nPos - 1)
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, lidlibMove, STREAM_SEEK_SET, @libNewPosition))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the size of the stream.
' libNewSize : Specifies the new size, in bytes, of the stream.
' ========================================================================================
PRIVATE FUNCTION CMemStream.SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulilibNewSize AS ULARGE_INTEGER : ulilibNewSize.QuadPart = libNewSize
   RETURN this.SetResult(m_pStream->lpvtbl->SetSize(m_pStream, ulilibNewSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the size of the stream in bytes.
' ========================================================================================
PRIVATE FUNCTION CMemStream.GetSize () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM tstatstg AS STATSTG
   this.SetResult(m_pStream->lpvtbl->Stat(m_pStream, @tstatstg, STATFLAG_NONAME))
   RETURN tstatstg.cbSize.QuadPart
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies a specified number of bytes from the current seek pointer in the stream to the
' current seek pointer in another stream.
' - pstm : A pointer to the destination stream. The stream pointed to by pstm can be a new
'   stream or a clone of the source stream.
' - cb : The number of bytes of data to attempt to copy into the stream.
' - pcbRead : A pointer to the location where this method writes the actual number of bytes
'   read from the source. You can set this pointer to NULL. In this case, this method does
'   not provide the actual number of bytes read.
' - pcbWritten : A pointer to the location where this method writes the actual number of
'   bytes written to the destination. You can set this pointer to NULL. In this case, this
'   method does not provide the actual number of bytes written.
' ========================================================================================
PRIVATE FUNCTION CMemStream.CopyTo (BYVAL pstm AS IStream PTR, BYVAL cb AS ULONGINT, _
BYVAL pcbRead AS ULONGINT PTR = NULL, BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulicb AS ULARGE_INTEGER: ulicb.QuadPart = cb
   DIM ulicbRead AS ULARGE_INTEGER, ulicbWritten AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->CopyTo(m_pStream, pstm, ulicb, @ulicbRead, @ulicbWritten))
   IF pcbRead THEN *pcbRead = ulicbRead.QuadPart
   IF pcbWritten THEN *pcbWritten = ulicbWritten.QuadPart
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new stream with its own seek pointer that references the same bytes as
' the original stream.
' - ppstm : When successful, pointer to the location of an IStream pointer to the new stream.
'   If an error occurs, this parameter is NULL.
' Remarks:
'    The Clone method creates a new stream for accessing the same bytes but using a
'    separate seek pointer. The new stream sees the same data as the source-stream.
'    Changes written to one stream are immediately visible in the other. Range
'    locking is shared between the streams.
'    The initial setting of the seek pointer in the cloned stream instance is the same as
'    the current setting of the seek pointer in the original stream at the time of the
'    clone operation.
' ========================================================================================
PRIVATE FUNCTION CMemStream.Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Clone(m_pStream, ppstm))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemStream.Clone () AS IStream PTR
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN NULL
   DIM ppstm AS ISTREAM PTR
   this.SetResult(m_pStream->lpvtbl->Clone(m_pStream, @ppstm))
   RETURN ppstm
END FUNCTION
' ========================================================================================

' ========================================================================================
' CMemTextStream class
' ========================================================================================
TYPE CMemTextStream

Public:
   m_Result AS HRESULT
   m_pStream AS IStream PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pwszText AS CONST WSTRING PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo () AS DWSTRING
   DECLARE OPERATOR CAST () AS IStream PTR
   DECLARE FUNCTION StreamPtr () AS IStream PTR
   DECLARE FUNCTION Read (BYVAL numChars AS LONG) AS DWSTRING
   DECLARE FUNCTION Write (BYREF wszText AS CONST WSTRING) AS ULONG
   DECLARE FUNCTION Append (BYREF wszText AS CONST WSTRING) AS ULONG
   DECLARE FUNCTION Seek (BYVAL dlibMove AS ULONGINT) AS HRESULT
   DECLARE FUNCTION GetSeekPosition () AS ULONGINT
   DECLARE FUNCTION ResetSeekPosition () AS ULONGINT
   DECLARE FUNCTION SeekAtEndOfStream () AS ULONGINT
   DECLARE FUNCTION SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
   DECLARE FUNCTION GetSize () AS ULONGINT
   DECLARE FUNCTION CopyTo (BYVAL pstm AS IStream PTR, BYVAL cb AS ULONGINT, _
           BYVAL pcbRead AS ULONGINT PTR = NULL, BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
   DECLARE FUNCTION Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
   DECLARE FUNCTION Clone () AS IStream PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Creates a memory unicode text stream.
' ========================================================================================
PRIVATE CONSTRUCTOR CMemTextStream
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   m_pStream = AfxSHCreateMemStream(NULL, 0)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a memory text streams and initializes it with the content of a string.
' DIM pMemStm AS CMemTextStream = !"1234567890\r\l"
' pMemStm.Append(!"Second Test string\r\l")
' pMemStm.Append(!"Third Test string\r\l")
' DIM dws AS DWSTRING = pMemStm.Read(-1)   ' // Read all the lines
' print dws
' ========================================================================================
PRIVATE CONSTRUCTOR CMemTextStream (BYVAL pwszText AS CONST WSTRING PTR)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   IF pwszText = NULL THEN
      m_pStream = AfxSHCreateMemStream(NULL, 0)
   ELSE
      m_pStream = AfxSHCreateMemStream(CAST(BYTE PTR, pwszText), LEN(*pwszText) * 2)
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CMemTextStream destructor
' ========================================================================================
PRIVATE DESTRUCTOR CMemTextStream
   ' // Release the IStrean interface
   IF m_pStream THEN m_pStream->lpvtbl->Release(m_pStream)
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CMemTextStream.CAST () AS IStream PTR
   m_Result = 0
   OPERATOR = m_pStream
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying IStream interface
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.StreamPtr () AS IStream PTR
   m_Result = 0
   RETURN m_pStream
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads a specified number of characters from the stream into memory, starting at the
' current seek pointer.
' - numChars : The number of characters to read from the stream.
'   Pass -1 to read all the characters from the current seek position.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.Read (BYVAL numChars AS LONG) AS DWSTRING
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN ""
   IF numChars = 0 THEN RETURN ""
   DIM cbRead AS ULONG
   DIM cbMaxChars AS LONG = ((this.GetSize * 2) - (this.GetSeekPosition)) \ 2
   IF numChars < 0 OR numChars > cbMaxChars THEN numChars = cbMaxChars
   DIM dwsText AS DWSTRING = WSPACE(numChars)
   this.SetResult(m_pStream->lpvtbl->Read(m_pStream, *dwsText, numChars * 2, @cbRead))
   IF cbRead < LEN(dwsText) THEN dwsText = LEFT(dwsText, cbRead)
   RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a string at the current seek position.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.Write (BYREF wszText AS CONST WSTRING) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM cbWritten AS ULONG
   this.SetResult(m_pStream->lpvtbl->Write(m_pStream, STRPTR(wszText), LEN(wszText) * 2, @cbWritten))
   RETURN cbWritten \ 2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Appends a string at the end of the stream.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.Append (BYREF wszText AS CONST WSTRING) AS ULONG
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_END, @libNewPosition))
   IF m_Result THEN RETURN 0
   DIM cbWritten AS ULONG
   this.SetResult(m_pStream->lpvtbl->Write(m_pStream, STRPTR(wszText), LEN(wszText) * 2, @cbWritten))
   RETURN cbWritten \ 2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current seek position.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.GetSeekPosition () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_CUR, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the seek position at the beginning of the stream.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.ResetSeekPosition () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_SET, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================
' ========================================================================================
' Sets the seek position at the end of the stream.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.SeekAtEndOfStream () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM dlibMove AS LARGE_INTEGER : dlibMove.QuadPart = 0
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, dlibMove, STREAM_SEEK_END, @libNewPosition))
   RETURN libNewPosition.QuadPart
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the seek position as an absolute position from the start of the stream.
' nPos ranges from 1 to the end of the string.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.Seek (BYVAL nPos AS ULONGINT) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM lidlibMove AS LARGE_INTEGER : lidlibMove.QuadPart = (nPos - 1) * 2
   DIM libNewPosition AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->Seek(m_pStream, lidlibMove, STREAM_SEEK_SET, @libNewPosition))
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the size of the stream.
' libNewSize : Specifies the new size, in characters, of the stream.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.SetSize (BYVAL libNewSize AS ULONGINT) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulilibNewSize AS ULARGE_INTEGER : ulilibNewSize.QuadPart = libNewSize * 2
   RETURN this.SetResult(m_pStream->lpvtbl->SetSize(m_pStream, ulilibNewSize))
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the size of the stream in characters.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.GetSize () AS ULONGINT
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN 0
   DIM tstatstg AS STATSTG
   this.SetResult(m_pStream->lpvtbl->Stat(m_pStream, @tstatstg, STATFLAG_NONAME))
   RETURN tstatstg.cbSize.QuadPart \ 2
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies a specified number of characters from the current seek pointer in the stream to the
' current seek pointer in another stream.
' - pstm : A pointer to the destination stream. The stream pointed to by pstm can be a new
'   stream or a clone of the source stream.
' - cb : The number of characters of data to attempt to copy into the stream.
' - pcbRead : A pointer to the location where this method writes the actual number of bytes
'   read from the source. You can set this pointer to NULL. In this case, this method does
'   not provide the actual number of bytes read.
' - pcbWritten : A pointer to the location where this method writes the actual number of
'   bytes written to the destination. You can set this pointer to NULL. In this case, this
'   method does not provide the actual number of bytes written.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.CopyTo (BYVAL pstm AS IStream PTR, BYVAL cb AS ULONGINT, _
BYVAL pcbRead AS ULONGINT PTR = NULL, BYVAL pcbWritten AS ULONGINT PTR = NULL) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   DIM ulicb AS ULARGE_INTEGER: ulicb.QuadPart = cb * 2
   DIM ulicbRead AS ULARGE_INTEGER, ulicbWritten AS ULARGE_INTEGER
   this.SetResult(m_pStream->lpvtbl->CopyTo(m_pStream, pstm, ulicb, @ulicbRead, @ulicbWritten))
   IF pcbRead THEN *pcbRead = ulicbRead.QuadPart
   IF pcbWritten THEN *pcbWritten = ulicbWritten.QuadPart
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new stream with its own seek pointer that references the same bytes as
' the original stream.
' - ppstm : When successful, pointer to the location of an IStream pointer to the new stream.
'   If an error occurs, this parameter is NULL.
' Remarks:
'    The Clone method creates a new stream for accessing the same bytes but using a
'    separate seek pointer. The new stream sees the same data as the source-stream.
'    Changes written to one stream are immediately visible in the other. Range
'    locking is shared between the streams.
'    The initial setting of the seek pointer in the cloned stream instance is the same as
'    the current setting of the seek pointer in the original stream at the time of the
'    clone operation.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.Clone (BYVAL ppstm AS IStream PTR PTR) AS HRESULT
   IF m_pStream = NULL THEN RETURN this.SetResult(E_POINTER)
   RETURN this.SetResult(m_pStream->lpvtbl->Clone(m_pStream, ppstm))
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.Clone () AS IStream PTR
   IF m_pStream = NULL THEN this.SetResult(E_POINTER) : RETURN NULL
   DIM ppstm AS ISTREAM PTR
   this.SetResult(m_pStream->lpvtbl->Clone(m_pStream, @ppstm))
   RETURN ppstm
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CMemTextStream.GetErrorInfo () AS DWSTRING
   IF SUCCEEDED(m_Result) THEN RETURN "Success"
   DIM s AS DWSTRING = "Error &h" & HEX(m_Result, 8)
   SELECT CASE m_Result
      CASE E_POINTER : s += ": Null pointer"
      CASE E_INVALIDARG : s += ": Invalid argument"
      CASE STG_E_INVALIDFUNCTION : s += ": Unable to perform requested operation"
      CASE STG_E_ACCESSDENIED : s += ": Access denied"
      CASE STG_E_INVALIDHANDLE : s += ": Invalid handle"
      CASE STG_E_INSUFFICIENTMEMORY : s += ": Insufficient memory"
      CASE STG_E_INVALIDPOINTER : s += ": Invalid pointer"
      CASE STG_E_SEEKERROR : s += ": Seek error"
      CASE STG_E_WRITEFAULT : s += ": Write fault"
      CASE STG_E_INVALIDPARAMETER : s += ": Invalid parameter"
      CASE STG_E_UNKNOWN : s += ": Unknown error"
      CASE ELSE
         s += "Unknown error"
   END SELECT
   RETURN s
END FUNCTION
' ========================================================================================

END NAMESPACE
