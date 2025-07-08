' ########################################################################################
' Platform: Microsoft Windows
' Filename: CFileSys.inc
' Purpose:  File System class
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
#include once "AfxNova/AfxScrRun.bi"
#include once "AfxNova/DWSTRING.inc"
USING AfxNova

NAMESPACE AfxNova

' ########################################################################################
' CFileSys - File system class
' ########################################################################################
TYPE CFileSys

Public:
   m_Result AS HRESULT
   m_pFileSys AS Afx_IFileSystem3 PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo () AS DWSTRING
   DECLARE FUNCTION BuildPath (BYREF wszPath AS WSTRING, BYREF wszName AS WSTRING) AS DWSTRING
   DECLARE FUNCTION CopyFile (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING, BYVAL OverWriteFiles AS BOOLEAN = TRUE) AS HRESULT
   DECLARE FUNCTION CopyFolder (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING, BYVAL OverWriteFiles AS BOOLEAN = TRUE) AS HRESULT
   DECLARE FUNCTION CreateFolder (BYREF wszFolder AS WSTRING) AS Afx_IFolder PTR
   DECLARE FUNCTION DeleteFile (BYREF wszFileSpec AS WSTRING, BYVAL bForce AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION DeleteFolder (BYREF wszFolder AS WSTRING, BYVAL bForce AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION DriveLetters () AS DWSTRING
   DECLARE FUNCTION DriveExists (BYREF wszDrive AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION FileExists (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION FolderExists (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION GetAbsolutePathName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetBaseName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetDriveAvailableSpace (BYREF wszDrive AS WSTRING) AS DOUBLE
   DECLARE FUNCTION GetDriveFileSystem (BYREF wszDrive AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetDriveFreeSpace (BYREF wszDrive AS WSTRING) AS DOUBLE
   DECLARE FUNCTION GetDriveName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetDriveShareName (BYREF wszDrive AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetDriveTotalSize (BYREF wszFrive AS WSTRING) AS DOUBLE

   DECLARE FUNCTION GetFolderName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetFileName (BYREF dwsPathSpec AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetExtensionName (BYREF dwsPathSpec AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetTempName () AS DWSTRING
   DECLARE FUNCTION GetDriveType (BYREF wszDrive AS WSTRING) AS DRIVETYPECONST
   DECLARE FUNCTION GetFolderType (BYREF wszFolder AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetFileType (BYREF wszFile AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetFileVersion (BYREF wszFileName AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetParentFolderName (BYREF wszFolder AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetFolderShortPath (BYREF wszFolder AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetFolderShortName (BYREF wszFolder AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetFolderDriveLetter (BYREF wszFolder AS WSTRING) AS DWSTRING
   DECLARE FUNCTION SetFolderName (BYREF wszFolder AS WSTRING, BYREF wszName AS WSTRING) AS HRESULT
   DECLARE FUNCTION GetSerialNumber (BYREF wszDrive AS WSTRING) AS LONG
   DECLARE FUNCTION GetVolumeName (BYREF wszDrive AS WSTRING) AS DWSTRING
   DECLARE FUNCTION SetVolumeName (BYREF wszDrive AS WSTRING, BYREF wszName AS WSTRING) AS HRESULT
   DECLARE FUNCTION GetNumDrives () AS LONG
   DECLARE FUNCTION GetNumSubFolders (BYREF wszFolder AS WSTRING) AS LONG
   DECLARE FUNCTION GetNumFiles (BYREF wszFolder AS WSTRING) AS LONG
   DECLARE FUNCTION IsRootFolder (BYREF wszFolder AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION GetStandardStream (BYVAL StandardStreamType AS STANDARDSTREAMTYPES, BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
   DECLARE FUNCTION GetStdErrStream (BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
   DECLARE FUNCTION GetStdInStream (BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
   DECLARE FUNCTION GetStdOutStream (BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
   DECLARE FUNCTION GetFileSize (BYREF wszFile AS WSTRING) AS LONG
   DECLARE FUNCTION SetFileName (BYREF wszFile AS WSTRING, BYREF wszName AS WSTRING) AS HRESULT
   DECLARE FUNCTION GetFileShortPath (BYREF wszFile AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetFileShortName (BYREF wszFile AS WSTRING) AS DWSTRING
   DECLARE FUNCTION GetFileAttributes (BYREF wszFile AS WSTRING) AS FILEATTRIBUTE
   DECLARE FUNCTION SetFileAttributes (BYREF wszFile AS WSTRING, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   DECLARE FUNCTION GetFileDateCreated (BYREF wszFile AS WSTRING) AS DATE_
   DECLARE FUNCTION GetFileDateLastModified (BYREF wszFile AS WSTRING) AS DATE_
   DECLARE FUNCTION GetFileDateLastAccessed (BYREF wszFile AS WSTRING) AS DATE_
   DECLARE FUNCTION GetFolderSize (BYREF wszFolder AS WSTRING) AS LONG
   DECLARE FUNCTION GetFolderAttributes (BYREF wszFolder AS WSTRING) AS FILEATTRIBUTE
   DECLARE FUNCTION SetFolderAttributes (BYREF wszFolder AS WSTRING, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   DECLARE FUNCTION GetFolderDateCreated (BYREF wszFolder AS WSTRING) AS DATE_
   DECLARE FUNCTION GetFolderDateLastModified (BYREF wszFolder AS WSTRING) AS DATE_
   DECLARE FUNCTION GetFolderDateLastAccessed (BYREF wszFolder AS WSTRING) AS DATE_
   DECLARE FUNCTION IsDriveReady (BYREF wszDrive AS WSTRING) AS BOOLEAN
   DECLARE FUNCTION MoveFolder (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING) AS HRESULT
   DECLARE FUNCTION MoveFile (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING) AS HRESULT

   '// Helper function
   DECLARE FUNCTION DateTimeToString (BYVAL vbDate AS DATE_, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT, BYVAL dwFlags AS DWORD = 0) AS DWSTRING

END TYPE
' ########################################################################################

' ========================================================================================
' Constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CFileSys
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Filesystem object
   DIM CLSID_FileSystemObject_ AS GUID = (&h0D43FE01, &hF093, &h11CF, {&h89, &h40, &h00, &hA0, &hC9, &h05, &h42, &h28})
'   DIM IID_IFileSystem AS GUID = (&h0AB5A3D0, &hE5B6, &h11D0, {&hAB, &hF5, &h00, &hA0, &hC9, &h0F, &hFF, &hC0})
   DIM IID_IFileSystem3 AS GUID = (&h2A0B9D10, &h4B87, &h11D3, {&hA9, &h7A, &h00, &h10, &h4B, &h36, &h5C, &h9F})
   SetResult(CoCreateInstance(@CLSID_FileSystemObject_, NULL, CLSCTX_INPROC_SERVER, @IID_IFileSystem3, @m_pFileSys))
   IF m_pFileSys = NULL THEN EXIT CONSTRUCTOR
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CFileSys
   IF m_pFileSys THEN m_pFileSys->Release
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetErrorInfo () AS DWSTRING
   DIM dws AS DWSTRING = "Error &h" & HEX(m_Result, 8)
   SELECT CASE UCASE(HEX(m_Result, 8))
      CASE "800A0030" : dws += ": Path not found"
      CASE "800A0034" : dws += ": Bad file name or number"
      CASE "800A0035" : dws += ": File not found"
      CASE "800A0036" : dws += ": Bad file mode"
      CASE "800A0037" : dws += ": File is already open"
      CASE "800A0039" : dws += ": Device I/O error"
      CASE "800A003A" : dws += ": File already exists"
      CASE "800A003D" : dws += ": Disk space is full"
      CASE "800A003E" : dws += ": Input beyond the end of the file"
      CASE "800A0043" : dws += ": Too many files"
      CASE "800A0044" : dws += ": Device not available"
      CASE "800A0046" : dws += ": Permission denied"
      CASE "800A0047" : dws += ": Disk not ready"
      CASE "800A004A" : dws += ": Cannot rename with different drive"
      CASE "800A004B" : dws += ": Path/file access error"
      CASE "800A004C" : dws += ": Path not found"
      CASE ELSE
         DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
         cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
                 FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
                 NULL, m_Result, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
                 cast(LPWSTR, @pBuffer), 0, NULL)
         IF cbLen THEN
            dwsMsg = *pBuffer
            LocalFree pBuffer
         END IF
         IF m_Result = 0 THEN
            dws = dwsMsg
         ELSE
            dws += " " & dwsMsg
         END IF
   END SELECT
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Appends a name to an existing path.
' Example:
' #include once "AfxNova/CFileSys.inc"
' DIM pFileSys AS CFileSys
' DIM dwsNewPath AS DWSTRING = pFileSys.BuildPath("C:\MyFolder", "Text.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.BuildPath (BYREF wszPath AS WSTRING, BYREF wszName AS WSTRING) AS DWSTRING
   DIM bstrNewPath AS AFX_BSTR
   DIM bstrPath AS AFX_BSTR = SysAllocString(wszPath)
   DIM bstrName AS AFX_BSTR = SysAllocString(wszName)
   SetResult(m_pFileSys->BuildPath(bstrPath, bstrName, @bstrNewPath))
   SysFreeString bstrPath
   SysFreeString bstrName
   DIM dwsNewPath AS DWSTRING = bstrNewPath
   SysFreeString bstrNewPath
   RETURN dwsNewPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Copies one or more files from one location to another.
' Parameters:
' - wszSource: Character string file specification, which can include wildcard characters,
'   for one or more files to be copied.
' - wszDestination: Character string destination where the file or files from source
'   are to be copied. Wildcard characters are not allowed.
' - OverWriteFiles: Boolean value that indicates if existing files are to be overwritten.
'   If true, files are overwritten; if false, they are not. The default is true. Note that
'   CopyFile will fail if destination has the read-only attribute set, regardless of the
'   value of overwrite.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.CopyFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.CopyFile (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING, BYVAL OverWriteFiles AS BOOLEAN = TRUE) AS HRESULT
   DIM bstrSource AS AFX_BSTR = SysAllocstring(wszSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocstring(wszDestination)
   SetResult(m_pFileSys->CopyFile(bstrSource, bstrDestination, OverWriteFiles))
   SysFreestring bstrSource
   SysFreestring bstrDestination
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Recursively copies a folder from one location to another.
' Parameters:
' - cbsSource: Character string file specification, which can include wildcard characters,
'   for one or more folders to be copied.
' - cbsDestination: Character string destination where the folder and subfolder from source
'   are to be copied. Wildcard characters are not allowed.
' - Boolean value that indicates if existing folders are to be overwritten. If true, files
'   are overwritten; if false, they are not. The default is true.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.CopyFolder("C:\MyFolder", "C:\MyOtherFolder\")
' ========================================================================================
PRIVATE FUNCTION CFileSys.CopyFolder (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING, BYVAL OverWriteFiles AS BOOLEAN = TRUE) AS HRESULT
   DIM bstrSource AS AFX_BSTR = SysAllocstring(wszSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocstring(wszDestination)
   SetResult(m_pFileSys->CopyFolder(bstrSource, bstrDestination, OverWriteFiles))
   SysFreestring bstrSource
   SysFreestring bstrDestination
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Creates a folder. An error occurs if the specified folder already exists.
' Example:
' DIM pFileSys AS CFileSys
' DIM pFolder AS Afx_Folder PTR
' pFolder = pFileSys.CreateFolder("C:\MyNewFolder")
' IF pFolder THEN
'    ....
'    pFolder->Release
' END IF
' ========================================================================================
PRIVATE FUNCTION CFileSys.CreateFolder (BYREF wszFolder AS WSTRING) AS Afx_IFolder PTR
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   SetResult(m_pFileSys->CreateFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   RETURN pFolder
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes a specified file.
' Parameters:
' - cbsFileSpec: The name of the file to delete. cbsFileSpec can contain wildcard characters
'   in the last path component.
' - bForce: Boolean value that is true if files with the read-only attribute set are to be
'   deleted; false (default) if they are not.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.DeleteFile("C:\MyFolder\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.DeleteFile (BYREF wszFileSpec AS WSTRING, BYVAL bForce AS BOOLEAN = FALSE) AS HRESULT
   DIM bstrFilespec AS AFX_BSTR = SysAllocString(wszFileSpec)
   SetResult(m_pFileSys->DeleteFile(bstrFilespec, bForce))
   SysFreeString bstrFilespec
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Deletes a specified folder and its contents.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.DeleteFolder("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.DeleteFolder (BYREF wszFolder AS WSTRING, BYVAL bForce AS BOOLEAN = FALSE) AS HRESULT
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   SetResult(m_pFileSys->DeleteFolder(bstrFolder, bForce))
   SysFreeString bstrFolder
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Checks if the specified drive exists and returns TRUE or FALSE.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' DIM fExists AS BOOLEAN = pFileSys.DriveExists("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.DriveExists (BYREF wszDrive AS WSTRING) AS BOOLEAN
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   DIM fExists AS VARIANT_BOOL
   SetResult(m_pFileSys->DriveExists(bstrDrive, @fExists))
   SysFreeString bstrDrive
   RETURN fExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a semicolon separated list with the driver letters.
' ========================================================================================
PRIVATE FUNCTION CFileSys.DriveLetters () AS DWSTRING
   IF m_pFileSys = NULL THEN RETURN ""
   ' // Get a reference to the IDriveCollection interface
   DIM pDrives AS Afx_IDriveCollection PTR
   SetResult(m_pFileSys->get_Drives(@pDrives))
   DIM nDrives AS LONG, bstrDriveLetter AS AFX_BSTR, dwsDrives AS DWSTRING
   IF pDrives THEN
      ' // Get a reference to the standard enumerator
      DIM pEnum AS IEnumVARIANT PTR, pDrive AS Afx_IDrive PTR
      DIM vItem AS VARIANT, celtFetched AS ULONG
      SetResult(pDrives->get__NewEnum(@cast(ANY PTR, pEnum)))
      ' // Enumerate the collection
      IF pEnum THEN
         DO
            pEnum->lpvtbl->Next(pEnum, 1, @vItem, @celtFetched)
            IF celtFetched = 0 THEN EXIT DO
            pDrive = cast(Afx_IDrive PTR, cast(ANY PTR, vItem.pdispVal))
            IF pDrive = NULL THEN
               VariantClear @vItem
               EXIT DO
            END IF
            SetResult(pDrive->get_DriveLetter(@bstrDriveLetter))
            IF bstrDriveLetter THEN
               dwsDrives += *bstrDriveLetter & ";"
               SysFreeString bstrDriveLetter
            END IF
            VariantClear @vItem
         LOOP
         pEnum->lpvtbl->Release(pEnum)
      END IF
      pDrives->Release
   END IF
   ' // Remove the last ";"
   IF LEN(dwsDrives) THEN dwsDrives = LEFT(dwsDrives, LEN(dwsDrives) - 1)
   ' // Return the drive letters
   RETURN dwsDrives
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True if the specified file exists; False if it does not.
' Example:
' DIM pFileSys AS CFileSys
' DIM fExists AS BOOLEAN = pFileSys.FileExists("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.FileExists (BYREF wszFileSpec AS WSTRING) AS BOOLEAN
   DIM fExists AS VARIANT_BOOL
   DIM bstrFileSpec AS AFX_BSTR = SysAllocString(wszFileSpec)
   SetResult(m_pFileSys->FileExists(bstrFileSpec, @fExists))
   SysFreeString bstrFileSpec
   RETURN fExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True if the specified folder exists; False if it does not.
' Example:
' DIM pFileSys AS CFileSys
' DIM fExists AS BOOLEAN = pFileSys.FolderExists("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.FolderExists (BYREF wszFolderSpec AS WSTRING) AS BOOLEAN
   DIM fExists AS VARIANT_BOOL
   DIM bstrFolderSpec AS AFX_BSTR = SysAllocString(wszFolderSpec)
   SetResult(m_pFileSys->FolderExists(bstrFolderSpec, @fExists))
   SysFreeString bstrFolderSpec
   RETURN fExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns complete and unambiguous path from a provided path specification.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsName AS DWSTRING = pFileSys.GetAbsolutePathName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetAbsolutePathName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(wszPathSpec)
   SetResult(m_pFileSys->GetAbsolutePathName(bstrPathSpec, @bstrName))
   SysFreeString bstrPathSpec
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the base name of the last component, less any file
' extension, in a path.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsName AS DWSTRING = pFileSys.GetBaseName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetBaseName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(wszPathSpec)
   IF INSTRREV(wszPathSpec, ".") THEN SetResult(m_pFileSys->GetBaseName(bstrPathSpec, @bstrName))
   SysFreeString bstrPathSpec
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the amount of space available to a user on the specified drive or network share.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetDriveAvailableSpace("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveAvailableSpace (BYREF wszDrive AS WSTRING) AS DOUBLE
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive THEN
      DIM vSize AS VARIANT
      SetResult(pDrive->get_AvailableSpace(@vSize))
      pDrive->Release
      RETURN vSize.dblVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the type of file system in use for the specified drive.
' Available return types include FAT, NTFS, and CDFS.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsFileSystem AS DWSTRING = pFileSys.GetDriveFileSystem("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveFileSystem (BYREF wszDrive AS WSTRING) AS DWSTRING
   ' // Get a reference to the IDrive interface
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeSTring bstrDrive
   DIM bstrFileSystem AS AFX_BSTR
   IF pDrive THEN
      SetResult(pDrive->get_FileSystem(@bstrFileSystem))
      pDrive->Release
   END IF
   DIM dwsFileSystem AS DWSTRING = bstrFileSystem
   SysFreeString bstrFileSystem
   RETURN dwsFileSystem
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the amount of free space available to a user on the specified drive or network share.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetDriveFreeSpace("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveFreeSpace (BYREF wszDrive AS WSTRING) AS DOUBLE
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeSTring bstrDrive
   IF pDrive THEN
      DIM vSize AS VARIANT
      SetResult(pDrive->get_FreeSpace(@vSize))
      pDrive->Release
      RETURN vSize.dblVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the name of the drive for a specified path.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsName AS DWSTRING = pFileSys.GetDriveName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(wszPathSpec)
   SetResult(m_pFileSys->GetDriveName(bstrPathSpec, @bstrName))
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' *** UNTESTED ***
' Returns the network share name for a specified drive.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' DIM pFileSys AS CFileSys
' DIM dwsShareName AS DWSTRING = pFileSys.GetDriveShareName("H:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveShareName (BYREF wszDrive AS WSTRING) AS DWSTRING
   ' // Get a reference to the IDrive interface
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   DIM bstrName AS AFX_BSTR
   IF pDrive THEN
      SetResult(pDrive->get_ShareName(@bstrName))
      pDrive->Release
   END IF
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the total space, in bytes, of a drive or network share.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetDriveTotalSize("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveTotalSize (BYREF wszDrive AS WSTRING) AS DOUBLE
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive THEN
      DIM vSize AS VARIANT
      SetResult(pDrive->get_TotalSize(@vSize))
      pDrive->Release
      RETURN vSize.dblVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a value indicating the type of a specified drive.
' Return value:
'   DriveType_UnknownType = 0
'   DriveType_Removable = 1
'   DriveType_Fixed = 2
'   DriveType_Remote = 3
'   DriveType_CDRom = 4
'   DriveType_RamDisk = 5
' Example:
' DIM pFileSys AS CFileSys
' DIM nDriveType AS DRIVETYPECONST = pFileSys.GetDriveType("C:")
' DIM t AS WSTRING
' SELECT CASE nDriveType
'    CASE 0 : t = "Unknown"
'    CASE 1 : t = "Removable"
'    CASE 2 : t = "Fixed"
'    CASE 3 : t = "Network"
'    CASE 4 : t = "CD-ROM"
'    CASE 5 : t = "RAM Disk"
' END SELECT
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetDriveType (BYREF wszDrive AS WSTRING) AS DRIVETYPECONST
   ' // Get a reference to the IDrive interface
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   DIM nDriveType AS DRIVETYPECONST
   IF pDrive THEN
      SetResult(pDrive->get_DriveType(@nDriveType))
      pDrive->Release
   END IF
   RETURN nDriveType
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the extension name of the file for a specified path.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsName AS DWSTRING = pFileSys.GetExtensionName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetExtensionName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR= SysAllocString(wszPathSpec)
   SetResult(m_pFileSys->GetExtensionName(bstrPathSpec, @bstrName))
   SysFreeString bstrPathSpec
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the attributes of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM lAttr FILEATTRIBUTE = pFileSys.GetFileAttributes("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileAttributes (BYREF wszFile AS WSTRING) AS FILEATTRIBUTE
   DIM lAttr AS FILEATTRIBUTE
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(wszFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile THEN
      SetResult(pFile->get_Attributes(@lAttr))
      pFile->Release
   END IF
   RETURN lAttr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was created.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFileDateCreated("C:\MyPath\MyFile.txt")
' Note: To convert the date to a string, call DateTimeToString.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileDateCreated (BYREF wszFile AS WSTRING) AS DATE_
   DIM nDate AS DATE_
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SYsAllocString(wszFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile THEN
      SetResult(pFile->get_DateCreated(@nDate))
      pFile->Release
   END IF
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was last modified.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFileDateLastAccessed("C:\MyPath\MyFile.txt")
' Note: To convert the date to a string, call DateTimeToString.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileDateLastAccessed (BYREF wszFile AS WSTRING) AS DATE_
   DIM nDate AS DATE_
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(wszFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile THEN
      SetResult(pFile->get_DateLastAccessed(@nDate))
      pFile->Release
   END IF
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was last modified.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFileDateLastModified("C:\MyPath\MyFile.txt")
' Note: To convert the date to a string, call DateTimeToString.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileDateLastModified (BYREF wszFile AS WSTRING) AS DATE_
   DIM nDate AS DATE_
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(wszFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile THEN
      SetResult(pFile->get_DateLastModified(@nDate))
      pFile->Release
   END IF
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the name of the file for a specified path.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsName AS DWSTRING = pFileSys.GetFileName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   DIM bstrName AS AFX_BSTR
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(wszPathSpec)
   IF INSTRREV(wszPathSpec, ".") THEN SetResult(m_pFileSys->GetFileName(bstrPathSpec, @bstrName))
   SysFreeString bstrPathSpec
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the short name used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFileName CBSTR = pFileSys.GetFileShortName("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileShortName (BYREF wszFile AS WSTRING) AS DWSTRING
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(wszFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   DIM bstrName AS AFX_BSTR
   IF pFile THEN
      SetResult(pFile->get_ShortName(@bstrName))
      pFile->Release
   END IF
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short path used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsFileName DWSTRING = pFileSys.GetFileShortPath("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileShortPath (BYREF wszFile AS WSTRING) AS DWSTRING
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(wszFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   DIM bstrPath AS AFX_BSTR
   IF pFile THEN
      SetResult(pFile->get_ShortPath(@bstrPath))
      pFile->Release
   END IF
   DIM dwsPath AS DWSTRING = bstrPath
   SysFreeString bstrPath
   RETURN dwsPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the size, in bytes, of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM nFileSize AS LONG = pFileSys.GetFileSize("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileSize (BYREF wszFile AS WSTRING) AS LONG
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysALlocString(wszFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile THEN
      DIM vSize AS VARIANT
      SetResult(pFile->get_Size(@vSize))
      pFile->Release
      IF m_Result = S_OK THEN RETURN vSize.lVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns localizedx information about the type of a file. For example, for files ending
' in .TXT, "Text Document" is returned ("Documento de texto" in Spanish).
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsFileType AS DWSTRING = pFileSys.FileType("C:\MyPath\MyFile.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileType (BYREF wszFile AS WSTRING) AS DWSTRING
   DIM bstrFile AS AFX_BSTR = SysAllocString(wszFile)
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   DIM bstrType AS AFX_BSTR
   IF pFile THEN
      SetResult(pFile->get_Type(@bstrType))
      pFile->Release
   END IF
   DIM dwsType AS DWSTRING = bstrType
   SysFreeString bstrType
   RETURN dwsType
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the version number of a specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsVersion AS DWSTRING = pFileSys.GetFileVersion("C:\MyFolder\MyFile.doc")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFileVersion (BYREF wszFile AS WSTRING) AS DWSTRING
   DIM bstrVersion AS AFX_BSTR
   DIM bstrFile AS AFX_BSTR = SysAllocstring(wszFile)
   SetResult(m_pFileSys->GetFileVersion(bstrFile, @bstrVersion))
   SysFreeString bstrFile
   DIM dwsVersion AS DWSTRING = bstrVersion
   SysFreestring bstrVersion
   RETURN dwsVersion
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the attributes of the specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM lAttr FILEATTRIBUTE = pFileSys.GetFolderAttributes("C:\MyPath")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderAttributes (BYREF wszFolder AS WSTRING) AS FILEATTRIBUTE
   DIM lAttr AS FILEATTRIBUTE
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      SetResult(pFolder->get_Attributes(@lAttr))
      pFolder->Release
   END IF
   RETURN lAttr
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified folder was created.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFolderDateCreated("C:\MyPath")
' Note: To convert the date to a string, call the DateTimeToString method.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderDateCreated (BYREF wszFolder AS WSTRING) AS DATE_
   DIM nDate AS DATE_
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      SetResult(pFolder->get_DateCreated(@nDate))
      pFolder->Release
   END IF
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was last modified.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFolderDateLastAccessed("C:\MyPath")
' Note: To convert the date to a string, call the DateTimeToString method.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderDateLastAccessed (BYREF wszFolder AS WSTRING) AS DATE_
   DIM nDate AS DATE_
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      SetResult(pFolder->get_DateLastAccessed(@nDate))
      pFolder->Release
   END IF
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the date and time that the specified file was last modified.
' Example:
' DIM pFileSys AS CFileSys
' DIM nDate AS DATE_ = pFileSys.GetFolderDateLastModified("C:\MyPath")
' Note: To convert the date to a string, call the DateTimeToString method.
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderDateLastModified (BYREF wszFolder AS WSTRING) AS DATE_
   DIM nDate AS DATE_
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      SetResult(pFolder->get_DateLastModified(@nDate))
      pFolder->Release
   END IF
   RETURN nDate
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the drive letter for a specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsDriveLetter AS DWSTRING = pFileSys.GetFolderDriveLetter("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderDriveLetter (BYREF wszFolder AS WSTRING) AS DWSTRING
   ' // Get a reference to the IFolder interface
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   DIM bstrDriveLetter AS AFX_BSTR
   IF pFolder THEN
      DIM pDrive AS Afx_IDrive PTR
      SetResult(pFolder->get_Drive(@pDrive))
      IF pDrive THEN
         SetResult(pDrive->get_DriveLetter(@bstrDriveLetter))
         pDrive->Release
      END IF
      pFolder->Release
   END IF
   DIM dwsDriveLetter AS DWSTRING = bstrDriveLetter
   SysFreeString bstrDriveLetter
   RETURN dwsDriveLetter
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a string containing the name of the folder for a specified path, i.e. the
' path minus the file name.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsName AS DWSTRING = pFileSys.GetFolderName("C:\MyFolder\Test.txt")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderName (BYREF wszPathSpec AS WSTRING) AS DWSTRING
   ' // Remove the file name
   DIM bstrPathSpec AS AFX_BSTR = SysAllocString(wszPathSpec)
   DIM dwsFolderName AS DWSTRING = wszPathSpec
   DIM bstrFileName AS AFX_BSTR
   IF INSTRREV(wszPathSpec, ".") THEN SetResult(m_pFileSys->GetFileName(bstrPathSpec, @bstrFileName))
   DIM dwsFileName AS DWSTRING = bstrFileName
   SysFreeString bstrFileName
   IF LEN(dwsFileName) THEN dwsFolderName = LEFT(dwsFolderName, LEN(dwsFolderName) - LEN(dwsFileName) - 1)
   IF RIGHT(dwsFolderName, 1) = "\" THEN dwsFolderName = LEFT(dwsFolderName, LEN(dwsFolderName) - 1)
   RETURN dwsFolderName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short name used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsFolderShortname AS DWSTRING = pFileSys.GetFolderShortName("c:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderShortName (BYREF wszFolder AS WSTRING) AS DWSTRING
   ' // Get a reference to the IFolder interface
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   DIM bstrName AS AFX_BSTR
   IF pFolder THEN
      SetResult(pFolder->get_ShortName(@bstrName))
      pFolder->Release
   END IF
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the short path used by programs that require the earlier 8.3 file naming convention.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsFolderShortPath AS DWSTRING = pFileSys.GetFolderShortPath("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderShortPath (BYREF wszFolder AS WSTRING) AS DWSTRING
   ' // Get a reference to the IFolder interface
   DIM bstrPath AS AFX_BSTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      SetResult(pFolder->get_ShortPath(@bstrPath))
      pFolder->Release
   END IF
   DIM dwsPath AS DWSTRING = bstrPath
   SysFreeString bstrPath
   RETURN dwsPath
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the size, in bytes, of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' DIM cbsFolderSize AS LONG = pFileSys.GetFolderSize("C:\MyPath")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderSize (BYREF wszFolder AS WSTRING) AS LONG
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      DIM vSize AS VARIANT
      SetResult(pFolder->get_Size(@vSize))
      pFolder->Release
      IF m_Result = S_OK THEN RETURN vSize.lVal
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns information about the type of a folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsFolderType AS DWSTRING = pFileSys.GetFolderType("c:\MyFolder)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetFolderType (BYREF wszFolder AS WSTRING) AS DWSTRING
   ' // Get a reference to the IFolder interface
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   DIM bstrType AS AFX_BSTR
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      SetResult(pFolder->get_Type(@bstrType))
      pFolder->Release
   END IF
   DIM dwsType AS DWSTRING = bstrType
   SysFreeString bstrType
   RETURN dwsType
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the number of drives.
' Example:
' DIM pFileSys AS CFileSys
' DIM numDrives AS LONG = pFileSys.GetNumDrives
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetNumDrives () AS LONG
   ' // Get a reference to the IDriveCollection interface
   DIM pDrives AS Afx_IDriveCollection PTR
   SetResult(m_pFileSys->get_Drives(@pDrives))
   DIM nDrives AS LONG
   IF pDrives THEN
      SetResult(pDrives->get_Count(@nDrives))
      pDrives->Release
   END IF
   RETURN nDrives
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the number of folders contained in a specified folder, including those with
' hidden and system file attributes set.
' Example:
' DIM pFileSys AS CFileSys
' DIM numSubFolders AS LONG = pFileSys.GetNumSubFolders("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetNumSubFolders (BYREF wszFolder AS WSTRING) AS LONG
   ' // Get a reference to the IFolder interface
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   DIM pFolder AS Afx_IFolder PTR
   DIM numSubFolders AS LONG
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      DIM pSubFolders AS Afx_IFolderCollection PTR
      SetResult(pFolder->get_SubFolders(@pSubFolders))
      IF pSubFolders THEN
         SetResult(pSubFolders->get_Count(@numSubFolders))
         pSubFolders->Release
      END IF
      pFolder->Release
   END IF
   RETURN numSubFolders
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the folder name for the parent of the specified folder.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetParentFolderName("C:\MyFolder\MySubfolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetParentFolderName (BYREF wszFolder AS WSTRING) AS DWSTRING
   DIM bstrName AS AFX_BSTR
   DIM bstrFolder AS AFX_BSTR = SysALlocString(wszFolder)
   SetResult(m_pFileSys->GetParentFolderName(bstrFolder, @bstrName))
   SysFreeString bstrFolder
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the volume serial number that the operating system assigns when a hard disk is formatted.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetSerialNumber("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetSerialNumber (BYREF wszDrive AS WSTRING) AS LONG
   ' // Get a reference to the IDrive interface
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   DIM pDrive AS Afx_IDrive PTR
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   DIM nSerialNumber AS LONG
   IF pDrive THEN
      SetResult(pDrive->get_SerialNumber(@nSerialNumber))
      pDrive->Release
   END IF
   RETURN nSerialNumber
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a TextStream object corresponding to the standard input, output, or error stream.
' DIM pFileSys AS CFileSys
' DIM pStm AS Afx_ITextStream PTR
' pStm = pFileSys.GetStandardStream(StandardStreamTypes_StdOut)
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetStandardStream (BYVAL StandardStreamType AS STANDARDSTREAMTYPES, BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
   DIM pStm AS Afx_ITextStream PTR
   SetResult(m_pFileSys->GetStandardStream(StandardStreamType, bUnicode, @pStm))
   RETURN pStm
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetStdErrStream (BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
   DIM pStm AS Afx_ITextStream PTR
   SetResult(m_pFileSys->GetStandardStream(StandardStreamTypes_StdErr, bUnicode, @pStm))
   RETURN pStm
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetStdInStream (BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
   DIM pStm AS Afx_ITextStream PTR
   SetResult(m_pFileSys->GetStandardStream(StandardStreamTypes_StdIn, bUnicode, @pStm))
   RETURN pStm
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetStdOutStream (BYVAL bUnicode AS BOOLEAN = FALSE) AS Afx_ITextStream PTR
   DIM pStm AS Afx_ITextStream PTR
   SetResult(m_pFileSys->GetStandardStream(StandardStreamTypes_StdOut, bUnicode, @pStm))
   RETURN pStm
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns a randomly generated temporary file or folder name that is useful for performing
' operations that require a temporary file or folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM dwsName AS DWSTRING = pFileSys.GetTempName
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetTempName () AS DWSTRING
   DIM bstrName AS AFX_BSTR
   SetResult(m_pFileSys->GetTempName(@bstrName))
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True (-1) if the specified drive is ready; False (0) if it is not.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' DIM fIsReady AS BOOLEAN = pFileSys.IsDriveReady("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.IsDriveReady (BYREF wszDrive AS WSTRING) AS BOOLEAN
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS Afx_BSTR = SysALlocString(wszDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   DIM fIsReady AS VARIANT_BOOL
   IF pDrive THEN
      SetResult(pDrive->get_IsReady(@fIsReady))
      pDrive->Release
   END IF
   RETURN fIsReady
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns True(-1) if the specified folder is the root folder; False(0) if it is not.
' DIM pFileSys AS CFileSys
' PRINT pFileSys.IsRootFolder("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.IsRootFolder (BYREF wszFolder AS WSTRING) AS BOOLEAN
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      DIM fRootFolder AS VARIANT_BOOL
      SetResult(pFolder->get_IsRootFolder(@fRootFolder))
      pFolder->Release
      RETURN fRootFolder
   END IF
   RETURN FALSE
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Moves one or more files from one location to another.
' Parameters:
' - cbsSource: The path to the file or files to be moved. The cbsSource argument string
'   can contain wildcard characters in the last path component only.
' - cbsDestination: The path where the file or files are to be moved. The cbsDestination
'   argument can't contain wildcard characters.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.MoveFile("C:\MyFolder\MyFile.txt", "C:\MyOtherFolder\")
' ========================================================================================
PRIVATE FUNCTION CFileSys.MoveFile (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING) AS HRESULT
   DIM bstrSource AS AFX_BSTR = SysAllocString(wszSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocString(wszDestination)
   SetResult(m_pFileSys->MoveFile(bstrSource, bstrDestination))
   SysFreestring bstrSource
   SysFreestring bstrDestination
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves one or more folders from one location to another.
' Parameters:
' - cbsSource: The path to the folder or folders to be moved. The cbsSource argument string
'   can contain wildcard characters in the last path component only.
' - cbsDestination: The path where the folder or folders are to be moved. The cbsDestination
'   argument can't contain wildcard characters.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.MoveFolder("C:\MyFolder", "C:\MyNewFolder\")
' ========================================================================================
PRIVATE FUNCTION CFileSys.MoveFolder (BYREF wszSource AS WSTRING, BYREF wszDestination AS WSTRING) AS HRESULT
   DIM bstrSource AS AFX_BSTR = SysAllocstring(wszSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocstring(wszDestination)
   SetResult(m_pFileSys->MoveFolder(bstrSource, bstrDestination))
   SysFreestring bstrSource
   SysFreestring bstrDestination
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the attributes of the specified file.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.SetFileAttributes("C:\MyPath\MyFile.txt", 33)
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFileAttributes (BYREF wszFile AS WSTRING, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(wszFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile THEN
      SetResult(pFile->put_Attributes(lAttr))
      pFile->Release
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the attributes of the specified folder.
' Example:
' DIM pFileSys AS CFileSys
' DIM lAttr FILEATTRIBUTE = pFileSys.SetFolderAttributes("C:\MyPath")
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFolderAttributes (BYREF wszFolder AS WSTRING, BYVAL lAttr AS FILEATTRIBUTE) AS HRESULT
   ' // Get a reference to the IFolder interface
   DIM pFolder AS Afx_IFolder PTR
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      SetResult(pFolder->put_Attributes(lAttr))
      pFolder->Release
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the name of a specified folder.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.SetFolderName("c:\MyFolder", "NewName")
' Note: You only have to pass the new name of the folder, not the full path.
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFolderName (BYREF wszFolder AS WSTRING, BYREF wszName AS WSTRING) AS HRESULT
   ' // Get a reference to the IFolder interface
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   DIM pFolder AS Afx_IFolder PTR
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      DIM bstrName AS AFX_BSTR = SysAllocString(wszName)
      SetResult(pFolder->put_Name(bstrName))
      SysFreeString bstrName
      pFolder->Release
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Returns the volume name of the specified drive.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' Example:
' DIM pFileSys AS CFileSys
' PRINT pFileSys.GetVolumeName("C:")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetVolumeName (BYREF wszDrive AS WSTRING) AS DWSTRING
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   DIM bstrName AS AFX_BSTR
   IF pDrive THEN
      SetResult(pDrive->get_VolumeName(@bstrName))
      pDrive->Release
   END IF
   DIM dwsName AS DWSTRING = bstrName
   SysFreeString bstrName
   RETURN dwsName
END FUNCTION
' ========================================================================================

' ========================================================================================
' *** UNTESTED ***
' Sets the volume name of the specified drive.
' For drive letters, the root drive is not included. For example, the path for the C drive
' is C:, not C:\.
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetVolumeName (BYREF wszDrive AS WSTRING, BYREF wszName AS WSTRING) AS HRESULT
   ' // Get a reference to the IDrive interface
   DIM pDrive AS Afx_IDrive PTR
   DIM bstrDrive AS AFX_BSTR = SysAllocString(wszDrive)
   SetResult(m_pFileSys->GetDrive(bstrDrive, @pDrive))
   SysFreeString bstrDrive
   IF pDrive THEN
      DIM bstrName AS AFX_BSTR = SysAllocString(wszName)
      SetResult(pDrive->put_VolumeName(bstrName))
      SysFreeString bstrName
      pDrive->Release
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of files contained in a specified folder, including those with
' hidden and system file attributes set.
' Example:
' DIM pFileSys AS CFileSys
' DIM numFiles AS LONG = pFileSys.GetNumFiles("C:\MyFolder")
' ========================================================================================
PRIVATE FUNCTION CFileSys.GetNumFiles (BYREF wszFolder AS WSTRING) AS LONG
   ' // Get a reference to the IFolder interface
   DIM bstrFolder AS AFX_BSTR = SysAllocString(wszFolder)
   DIM pFolder AS Afx_IFolder PTR
   DIM numFiles AS LONG
   SetResult(m_pFileSys->GetFolder(bstrFolder, @pFolder))
   SysFreeString bstrFolder
   IF pFolder THEN
      DIM pFiles AS Afx_IFileCollection PTR
      SetResult(pFolder->get_Files(@pFiles))
      IF pFiles THEN
         SetResult(pFiles->get_Count(@numFiles))
         pFiles->Release
      END IF
      pFolder->Release
   END IF
   RETURN numFiles
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Sets the name of a specified file.
' Example:
' DIM pFileSys AS CFileSys
' pFileSys.SetFileName("c:\MyFolder\Test.txt", "NewName")
' ========================================================================================
PRIVATE FUNCTION CFileSys.SetFileName (BYREF wszFile AS WSTRING, BYREF wszName AS WSTRING) AS HRESULT
   ' // Get a reference to the IFile interface
   DIM pFile AS Afx_IFile PTR
   DIM bstrFile AS AFX_BSTR = SysAllocString(wszFile)
   SetResult(m_pFileSys->GetFile(bstrFile, @pFile))
   SysFreeString bstrFile
   IF pFile THEN
      DIM bstrName AS AFX_BSTR = SysAllocString(wszName)
      SetResult(pFile->put_Name(bstrName))
      SysFreeString bstrName
      pFile->Release
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' * Converts a DATE_ type to a string containing the date and the time.
' ========================================================================================
PRIVATE FUNCTION CFileSys.DateTimeToString (BYVAL vbDate AS DATE_, BYVAL lcid AS LCID = LOCALE_USER_DEFAULT, BYVAL dwFlags AS DWORD = 0) AS DWSTRING
   DIM bstrOut AS AFX_BSTR
   VarBstrFromDate(vbDate, lcid, dwFlags, @bstrOut)
   DIM dwsOut AS DWSTRING = *bstrOut
   SysFreeString bstrOut
   RETURN dwsOut
END FUNCTION
' ========================================================================================

END NAMESPACE
