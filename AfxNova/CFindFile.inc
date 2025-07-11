' ########################################################################################
' Platform: Microsoft Windows
' Filename: CFindFile.inc
' Purpose:  File find class. Performs local file searches.
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
#include once "vbcompat.bi"
#include once "AfxNova/AfxWin.inc"

NAMESPACE AfxNova

' ########################################################################################
' CFindFile class.
' ########################################################################################
TYPE CFindFile

   DIM m_hFile AS HANDLE                     ' // Handle of the file used to start the search
   DIM m_FileInfo AS WIN32_FIND_DATAW        ' // WIN32_FIND_DATAW structure
   DIM m_wszFileSpec AS WSTRING * MAX_PATH   ' // Path used to start the search

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE SUB Close
   DECLARE FUNCTION FindFile (BYREF wszFileSpec AS WSTRING) AS LONG_PTR
   DECLARE FUNCTION FindNext () AS LONG
   DECLARE FUNCTION FileAttributes () AS DWORD
   DECLARE FUNCTION IsDots () AS BOOLEAN
   DECLARE FUNCTION IsFile () AS BOOLEAN
   DECLARE FUNCTION IsFolder () AS BOOLEAN
   DECLARE FUNCTION IsSystemFile () AS BOOLEAN
   DECLARE FUNCTION IsHiddenFile () AS BOOLEAN
   DECLARE FUNCTION IsReadOnlyFile () AS BOOLEAN
   DECLARE FUNCTION IsTemporaryFile () AS BOOLEAN
   DECLARE FUNCTION IsNormalFile () AS BOOLEAN
   DECLARE FUNCTION IsCompressedFile () AS BOOLEAN
   DECLARE FUNCTION IsEncryptedFile () AS BOOLEAN
   DECLARE FUNCTION IsNotContentIndexedFile () AS BOOLEAN
   DECLARE FUNCTION IsOffLineFile () AS BOOLEAN
   DECLARE FUNCTION IsReparsePointFile () AS BOOLEAN
   DECLARE FUNCTION IsSparseFile () AS BOOLEAN
   DECLARE FUNCTION MatchesMask (BYVAL dwMask AS DWORD) AS BOOLEAN
   DECLARE FUNCTION FileName () AS DWSTRING
   DECLARE FUNCTION ShortFileName () AS DWSTRING
   DECLARE FUNCTION FileNameX () AS DWSTRING
   DECLARE FUNCTION FileExt () AS DWSTRING
   DECLARE FUNCTION FilePath () AS DWSTRING
   DECLARE FUNCTION FileURL () AS DWSTRING
   DECLARE FUNCTION Root () AS DWSTRING
   DECLARE FUNCTION FileSize () AS ULONGLONG
   DECLARE FUNCTION CreationFileTime () AS FILETIME
   DECLARE FUNCTION LastAccessFileTime () AS FILETIME
   DECLARE FUNCTION LastWriteFileTime () AS FILETIME
   DECLARE FUNCTION CreationSerialTime () AS DOUBLE
   DECLARE FUNCTION LastAccessSerialTime () AS DOUBLE
   DECLARE FUNCTION LastWriteSerialTime () AS DOUBLE

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CFindFile
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CFindFile
   IF m_hFile <> 0 AND m_hFile <> INVALID_HANDLE_VALUE THEN ..FindClose(m_hFile)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Close method
' Call this method to end the search, reset the context, and release all resources.
' After calling Close, you do not have to create a new CFindFile instance before
' calling FindFile to begin a new search.
' ========================================================================================
PRIVATE SUB CFindFile.Close
   IF m_hFile <> 0 AND m_hFile <> INVALID_HANDLE_VALUE THEN
      ..FindClose(m_hFile)
      m_hFile = 0
      m_wszFileSpec = ""
'      RESET m_FileInfo
   END IF
END SUB
' ========================================================================================

' ========================================================================================
' Method FindFile
' Call this method to open a file search.
' Parameter:
' wszName
'   A string containing the name of the file to find (must not end in a trailing
'   backslash). If you pass an empty string, FindFile does a wildcard (*.*) search.
'   If the string ends with a wildcard, period (.), or directory name, the user must
'   have access permissions to the root and all subdirectories on the path.
'   To extend the MAX_PATH limit to 32,767 wide characters, prepend "\\?\" to the path.
' Return value:
'   S_OK on success or an error code on failure.
'   To get extended error information, call GetLastError.
'   If the function fails because no matching files can be found, the GetLastError
'   function returns ERROR_FILE_NOT_FOUND.
' ========================================================================================
PRIVATE FUNCTION CFindFile.FindFile (BYREF wszFileSpec AS WSTRING) AS LONG_PTR
   this.Close   ' // Close previous search, if any
   m_wszFileSpec = wszFileSpec   ' // For the GetRoot method
   IF m_wszFileSpec = "" THEN m_wszFileSpec = "*.*"
   m_hFile = .FindFirstFileW(m_wszFileSpec, @m_FileInfo)
   IF m_hFile = INVALID_HANDLE_VALUE THEN RETURN cast(LONG_PTR, INVALID_HANDLE_VALUE)
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method FindNext
' Call this method to continue a file search from a previous call to FindFile.
' Return value:
'   Nonzero if there are more files; zero if the file found is the last one in the
'   directory or if an error occurred. To get extended error information, call GetLastError.
'   If the file found is the last file in the directory, or if no matching files can be
'   found, the GetLastError function returns ERROR_NO_MORE_FILES.
' ========================================================================================
PRIVATE FUNCTION CFindFile.FindNext () AS LONG
   RETURN .FindNextFileW(m_hFile, @m_FileInfo)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method FileAttributes
' Call this method to get the attributes of the last found file.
' ========================================================================================
PRIVATE FUNCTION CFindFile.FileAttributes () AS DWORD
   RETURN m_FileInfo.dwFileAttributes
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsDots
' Call this method to test for the current directory and parent directory markers while
' iterating through files.
' Return value:
'   TRUE if the found file has the name "." or "..", which indicates that the found
'   file is actually a directory. Otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsDots () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY THEN
      IF m_FileInfo.cFileName = "." OR m_FileInfo.cFileName = ".." THEN RETURN TRUE
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsFile
' Call this method to determine if the found file is a file.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsFolder
' Call this method to determine if the found file is a folder.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsFolder () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsSystemFile
' Call this method to determine if the found file is a system file.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsSystemFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsHiddenFile
' Call this method to determine if the found file is a hidden file.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsHiddenFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsReadOnlyFile
' Call this method to determine if the found file is a read only file.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsReadOnlyFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsTemporaryFile
' Call this method to determine if the found file is a temporary file.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsTemporaryFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsNormalFile
' Call this method to determine if the found file is a normal file.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsNormalFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsCompressedFile
' Call this method to determine if the found file is a compressed file.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsCompressedFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsEncryptedFile
' Call this method to determine if the found file is a encrypted file.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsEncryptedFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_ENCRYPTED) = FILE_ATTRIBUTE_ENCRYPTED THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsNotContentIndexedFile
' Call this method to determine if the found file is not to be indexed by the content indexing service.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsNotContentIndexedFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_NOT_CONTENT_INDEXED) = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsOffLineFile
' Call this method to determine if the found file is not available immediately.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsOffLineFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_OFFLINE) = FILE_ATTRIBUTE_OFFLINE THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsReparsePointFile
' Call this method to determine if the found file is file or directory that has an
' associated reparse point, or a file that is a symbolic link.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsReparsePointFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_REPARSE_POINT) = FILE_ATTRIBUTE_REPARSE_POINT THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method IsSparseFile
' Call this method to determine if the found file is a sparse file.
' Return value:
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.IsSparseFile () AS BOOLEAN
   IF (m_FileInfo.dwFileAttributes AND FILE_ATTRIBUTE_SPARSE_FILE) = FILE_ATTRIBUTE_SPARSE_FILE THEN RETURN TRUE
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method MatchesMask
' Call this method to test the file attributes on the found file.
' Parameter:
' dwMask
'   Specifies one or more file attributes, identified in the WIN32_FIND_DATA structure,
'   for the found file. To search for multiple attributes, use the bitwise OR operator.
'   Any combination of the following attributes is acceptable:
'   FILE_ATTRIBUTE_ARCHIVE: The file is an archive file. Applications use this attribute
'     to mark files for backup or removal.
'   FILE_ATTRIBUTE_COMPRESSED: The file or directory is compressed. For a file, this
'     means that all of the data in the file is compressed. For a directory, this means
'   that compression is the default for newly created files and subdirectories.
'   FILE_ATTRIBUTE_DIRECTORY: The file is a directory.
'   FILE_ATTRIBUTE_NORMAL: The file has no other attributes set. This attribute is valid
'     only if used alone. All other file attributes override this attribute.
'   FILE_ATTRIBUTE_HIDDEN: The file is hidden. It is not to be included in an ordinary
'     directory listing.
'   FILE_ATTRIBUTE_READONLY: The file is read only. Applications can read the file but
'     cannot write to it or delete it.
'   FILE_ATTRIBUTE_SYSTEM: The file is part of or is used exclusively by the operating system.
'   FILE_ATTRIBUTE_TEMPORARY: The file is being used for temporary storage. Applications
'     should write to the file only if absolutely necessary. Most of the file's data remains
'     in memory without being flushed to the media because the file will soon be deleted.
' Return value
'   TRUE if successful; otherwise FALSE.
' ========================================================================================
PRIVATE FUNCTION CFindFile.MatchesMask (BYVAL dwMask AS DWORD) AS BOOLEAN
   RETURN (m_FileInfo.dwFileAttributes AND dwMask)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method FileName
' Call this method to get the title of the found file.
' Return value
'   The name of the most-recently-found file, excluding the extension. For example,
'   calling GetFileNameX to generate a user message about the file c:\myhtml\myfile.txt
'   returns the file name myfile.
' ========================================================================================
PRIVATE FUNCTION CFindFile.FileName () AS DWSTRING
   DIM p AS LONG
   DIM wszFileName AS WSTRING * MAX_PATH = m_FileInfo.cFileName
   p = INSTRREV(wszFileName, ANY ":/\")
   IF p THEN wszFileName = MID(wszFileName, p + 1)
   p = INSTRREV(wszFileName, ".")
   IF p THEN wszFileName = LEFT(wszFileName, p - 1)
   RETURN wszFileName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method FileNameX
' Call this method to get the name of the found file, includig the extension.
' Return value
'   The name of the most-recently-found file, including the extension. For example,
'   calling GetFileName to generate a user message about the file c:\myhtml\myfile.txt
'   returns the file name myfile.txt.
' ========================================================================================
PRIVATE FUNCTION CFindFile.FileNameX () AS DWSTRING
   RETURN m_FileInfo.cFileName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method ShortFilename
' Call this method to get an alternative name for the file.
' This name is in the classic 8.3 file name format.
' Return value
'   The alternate file name of the most-recently-found file, including the extension.
' ========================================================================================
PRIVATE FUNCTION CFindFile.ShortFileName () AS DWSTRING
   RETURN m_FileInfo.cAlternateFileName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method FileExt
' Call this method to get the extesion of the found file.
' ========================================================================================
PRIVATE FUNCTION CFindFile.FileExt () AS DWSTRING
   DIM p AS LONG
   p = INSTRREV(m_FileInfo.cFileName, ".")
   IF p THEN RETURN MID(m_FileInfo.cFileName, p + 1)
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method FilePath
' FilePath returns the entire path for the file.
' ========================================================================================
PRIVATE FUNCTION CFindFile.FilePath () AS DWSTRING
   RETURN this.Root & m_FileInfo.cFileName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method FileURL
' Call this member function to retrieve the specified URL.
' GetFileURL is similar to the member function GetFilePath, except that it returns the
' URL in the form file://path. For example, calling GetFileURL to get the complete URL
' for myfile.txt returns the URL file://c:\myhtml\myfile.txt.
' ========================================================================================
PRIVATE FUNCTION CFindFile.FileURL () AS DWSTRING
   RETURN "file://" & this.FilePath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method Root
' Call this member function to get the root of the found file.
' This method returns the drive specifier and path name used to start a search.
' For example, calling FindFile with *.dat results in Root returning an empty string.
' Passing a path, such as c:\windows\system\*.dll, to FindFile results in Root
' returning c:\windows\system\.
' ========================================================================================
PRIVATE FUNCTION CFindFile.Root () AS DWSTRING
   DIM wszPath AS WSTRING * MAX_PATH
   wszPath = m_wszFileSpec
   .PathRemoveFileSpecW(wszPath)
   IF LEN(wszPath) THEN wszPath += "\"
   RETURN wszPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method FileSize
' Call this method to get the size of the found file, in bytes.
' ========================================================================================
PRIVATE FUNCTION CFindFile.FileSize () AS ULONGLONG
   RETURN (m_FileInfo.nFileSizeHigh * (&hFFFFFFFF + 1)) + m_FileInfo.nFileSizeLow
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method CreationTime
' Call this method to get the time the file was created.
' ========================================================================================
PRIVATE FUNCTION CFindFile.CreationFileTime () AS FILETIME
   RETURN m_FileInfo.ftCreationTime
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method LastAccessTime
' Call this method to get the time the file was last accessed.
' ========================================================================================
PRIVATE FUNCTION CFindFile.LastAccessFileTime () AS FILETIME
   RETURN m_FileInfo.ftLastAccessTime
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method LastWriteTime
' Call this method to get the time the file was last written to,
' truncated, or overwritten.
' ========================================================================================
PRIVATE FUNCTION CFindFile.LastWriteFileTime () AS FILETIME
   RETURN m_FileInfo.ftLastWriteTime
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method CreationDateTime
' Returns the file's creation date and time as Date Serial.
' The result can be formated using FreeBasic's Format function.
' Usage example:
' DIM dt AS DOUBLE = pFindFile.CreationSerialTime
' PRINT Format(dt, "yyyy-mm-dd hh:mm AM/PM")
' ========================================================================================
PRIVATE FUNCTION CFindFile.CreationSerialTime () AS DOUBLE
   DIM ft AS FILETIME
   FileTimeToLocalFileTime(@m_FileInfo.ftCreationTime, @ft)
   DIM st AS SYSTEMTIME
   FileTimeToSystemTime(@ft, @st)
   DIM dt AS DOUBLE
   SystemTimeToVariantTime @st, @dt
   RETURN dt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method LastAccessnDateTime
' Returns the date and time the file was last accessed as Date Serial.
' The result can be formated using FreeBasic's Format function.
' Usage example:
' DIM dt AS DOUBLE = pFindFile.LastAccessnSerialTime
' PRINT Format(dt, "yyyy-mm-dd hh:mm AM/PM")
' ========================================================================================
PRIVATE FUNCTION CFindFile.LastAccessSerialTime () AS DOUBLE
   DIM ft AS FILETIME
   FileTimeToLocalFileTime(@m_FileInfo.ftLastAccessTime, @ft)
   DIM st AS SYSTEMTIME
   FileTimeToSystemTime(@ft, @st)
   DIM dt AS DOUBLE
   SystemTimeToVariantTime @st, @dt
   RETURN dt
END FUNCTION
' ========================================================================================

' ========================================================================================
' Method LastAccessnDateTime
' Returns the date and time the file was last written to, truncated, or overwritten, as Date Serial.
' The result can be formated using FreeBasic's Format function.
' Usage example:
' DIM dt AS DOUBLE = pFindFile.LastWriteSerialTime
' PRINT Format(dt, "yyyy-mm-dd hh:mm AM/PM")
' ========================================================================================
PRIVATE FUNCTION CFindFile.LastWriteSerialTime () AS DOUBLE
   DIM ft AS FILETIME
   FileTimeToLocalFileTime(@m_FileInfo.ftLastWriteTime, @ft)
   DIM st AS SYSTEMTIME
   FileTimeToSystemTime(@ft, @st)
   DIM dt AS DOUBLE
   SystemTimeToVariantTime @st, @dt
   RETURN dt
END FUNCTION
' ========================================================================================

END NAMESPACE
