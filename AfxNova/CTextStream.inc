' ########################################################################################
' Platform: Microsoft Windows
' Filename: CTextStream.inc
' Purpose:  Wrapper class for reading and writing text files.
' Compiler: FreeBASIC 32 & 64 bit
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
' CTextStream - Class for reading and writing text files.
' ########################################################################################
TYPE CTextStream

Public:
   m_pFileSys AS Afx_IFileSystem PTR
   m_pTxtStm AS Afx_ITextStream PTR
   m_Result AS HRESULT

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pstm AS Afx_ITextStream PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR CAST () AS Afx_ITextStream PTR
   DECLARE OPERATOR LET (BYVAL pstm AS Afx_ITextStream PTR)
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo () AS DWSTRING
   DECLARE FUNCTION Create (BYREF wszFileName AS CONST WSTRING, BYVAL bOverwrite AS BOOLEAN = TRUE, BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Open (BYREF wszFileName AS CONST WSTRING, BYVAL IOMode AS LONG = 1, BYVAL bCreate AS BOOLEAN = FALSE, BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION OpenUnicode (BYREF wszFileName AS CONST WSTRING, BYVAL IOMode AS LONG = 1, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION OpenForInputA (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION OpenForOutputA (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION OpenForAppendA (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION OpenForInputW (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION OpenForOutputW (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION OpenForAppendW (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Close () AS HRESULT
   DECLARE PROPERTY Line () AS LONG
   DECLARE PROPERTY Column () AS LONG
   DECLARE FUNCTION EOS () AS BOOLEAN
   DECLARE FUNCTION EOL () AS BOOLEAN
   DECLARE FUNCTION Read (BYVAL numChars AS LONG) AS DWSTRING
   DECLARE FUNCTION ReadLine () AS DWSTRING
   DECLARE FUNCTION ReadAll () AS DWSTRING
   DECLARE FUNCTION Write (BYREF wszText AS CONST WSTRING) AS HRESULT
   DECLARE FUNCTION WriteLine (BYREF wszText AS CONST WSTRING) AS HRESULT
   DECLARE FUNCTION WriteBlankLines (BYVAL numLines AS LONG) AS HRESULT
   DECLARE FUNCTION Skip (BYVAL numChars AS LONG) AS HRESULT
   DECLARE FUNCTION SkipLine () AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Constructors
' ========================================================================================
PRIVATE CONSTRUCTOR CTextStream
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Filesystem object
   DIM CLSID_FileSystemObject_ AS GUID = (&h0D43FE01, &hF093, &h11CF, {&h89, &h40, &h00, &hA0, &hC9, &h05, &h42, &h28})
   DIM IID_IFileSystem AS GUID = (&h0AB5A3D0, &hE5B6, &h11D0, {&hAB, &hF5, &h00, &hA0, &hC9, &h0F, &hFF, &hC0})
   SetResult(CoCreateInstance(@CLSID_FileSystemObject_, NULL, CLSCTX_INPROC_SERVER, @IID_IFileSystem, @m_pFileSys))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Initializes the class with a pointer to an instance of the Afx_ITextStream interface.
' Example:
' // Create an instance of the CTextStream class initialized
' // with a pointer to the standard StdOut stream
' DIM pTxtStm AS CTextStream = CFileSys().GetStdOutStream(TRUE)
' // Write a string and an end of line to the stream
' pTxtStm.WriteLine "This is a test."
' ========================================================================================
PRIVATE CONSTRUCTOR CTextStream (BYVAL pstm AS Afx_ITextStream PTR)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   m_pTxtStm = pstm
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTextStream
   IF m_pTxtStm THEN
      m_pTxtStm->Close
      m_pTxtStm->Release
   END IF
   IF m_pFileSys THEN m_pFileSys->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code
' ========================================================================================
PRIVATE FUNCTION CTextStream.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CTextStream.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CTextStream.GetErrorInfo () AS DWSTRING
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
' Cast operator.
' ========================================================================================
PRIVATE OPERATOR CTextStream.CAST () AS Afx_ITextStream PTR
   OPERATOR = m_pTxtStm
END OPERATOR
' ========================================================================================
' ========================================================================================
' Assignment operator.
' ========================================================================================
PRIVATE OPERATOR CTextStream.LET (BYVAL pstm AS Afx_ITextStream PTR)
   IF pstm = NULL THEN m_Result = E_POINTER : EXIT OPERATOR
   IF m_pTxtStm THEN
      m_pTxtStm->Close
      m_pTxtStm->Release
   END IF
   m_pTxtStm = pstm
   m_Result = 0
END OPERATOR
' ========================================================================================

' ========================================================================================
' Creates a specified file name and returns a TextStream object that can be used to read
' from or write to the file.
' Parameters:
' - wszFileName: String expression that identifies the file to create.
' - bOverwrite: Boolean value that indicates whether you can overwrite an existing file.
'   The value is true if the file can be overwritten, false if it can't be overwritten.
'   If omitted, existing files are overwritten.
' - bUnicode: Boolean value that indicates whether the file is created as a Unicode or
'   ASCII file. The value is true if the file is created as a Unicode file, false if it's
'   created as an ASCII file. If omitted, an ASCII file is assumed.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Create (BYREF wszFileName AS CONST WSTRING, BYVAL bOverwrite AS BOOLEAN = TRUE, BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(@wszFileName)
   SetResult(m_pFileSys->CreateTextFile(bstrFileName, bOverwrite, bUnicode, @m_pTxtStm))
   SysFreeString(bstrFileName)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens a specified file and returns a TextStream object that can be used to read from,
' write to, or append to the file.
' Parameters:
' - wszFileName: String expression that identifies the file to open.
' - IOMode: Can be one of three constants: IOMODE_ForReading, IOMODE_ForWriting, or IOMODE_ForAppending.
' - bCreate: Value that indicates whether a new file can be created if the specified
'   filename doesn't exist. The value is True if a new file is created, False if it isn't
'   created. If omitted, a new file isn't created.
' - bUnicode: TRUE or FALSE. Indicates the format of the opened file. If omitted, the file
'   is opened as ASCII.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Open (BYREF wszFileName AS CONST WSTRING, BYVAL IOMode AS LONG = 1, BYVAL bCreate AS BOOLEAN = FALSE, BYVAL bUnicode AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(@wszFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode, bCreate, bUnicode, @m_pTxtStm))
   SysFreeString(bstrFileName)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens a specified file and returns a TextStream object that can be used to read from,
' write to, or append to the file.
' Parameters:
' - wszFileName: String expression that identifies the file to open.
' - IOMode: Can be one of three constants: IOMODE_ForReading, IOMODE_ForWriting, or IOMODE_ForAppending.
' - bCreate: Value that indicates whether a new file can be created if the specified
'   filename doesn't exist. The value is True if a new file is created, False if it isn't
'   created. If omitted, a new file isn't created.
' ========================================================================================
PRIVATE FUNCTION CTextStream.OpenUnicode (BYREF wszFileName AS CONST WSTRING, BYVAL IOMode AS LONG = 1, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(@wszFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode, bCreate, Tristate_True, @m_pTxtStm))
   SysFreeString(bstrFileName)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Additional methods to ease opening files
' ========================================================================================
PRIVATE FUNCTION CTextStream.OpenForInputA (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(@wszFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode_ForReading, bCreate, TriState_False, @m_pTxtStm))
   SysFreeString(bstrFileName)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTextStream.OpenForOutputA (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(@wszFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode_ForWriting, bCreate, TriState_False, @m_pTxtStm))
   SysFreeString(bstrFileName)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTextStream.OpenForAppendA (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(@wszFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode_ForAppending, bCreate, TriState_False, @m_pTxtStm))
   SysFreeString(bstrFileName)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CTextStream.OpenForInputW (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(@wszFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode_ForReading, bCreate, TriState_True, @m_pTxtStm))
   SysFreeString(bstrFileName)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTextStream.OpenForOutputW (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(@wszFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode_ForWriting, bCreate, TriState_True, @m_pTxtStm))
   SysFreeString(bstrFileName)
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CTextStream.OpenForAppendW (BYREF wszFileName AS CONST WSTRING, BYVAL bCreate AS BOOLEAN = FALSE) AS HRESULT
   IF m_pTxtStm THEN
      SetResult(m_pTxtStm->Close)
      m_pTxtStm->Release
      m_pTxtStm = NULL
   END IF
   DIM bstrFileName AS AFX_BSTR = SysAllocString(@wszFileName)
   SetResult(m_pFileSys->OpenTextFile(bstrFileName, IOMode_ForAppending, bCreate, TriState_True, @m_pTxtStm))
   SysFreeString(bstrFileName)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes an open TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Close () AS HRESULT
   IF m_pTxtStm THEN
      RETURN(SetResult(m_pTxtStm->Close))
      m_pTxtStm = NULL
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Read-only property that returns the current line number in a TextStream file.
' After a file is initially opened and before anything is written, Line is equal to 1.
' ========================================================================================
PRIVATE PROPERTY CTextStream.Line () AS LONG
   DIM nLine AS LONG
   IF m_pTxtStm THEN SetResult(m_pTxtStm->get_Line(@nLine))
   PROPERTY = nLine
END PROPERTY
' ========================================================================================

' ========================================================================================
' Read-only property that returns the column number of the current character position in a
' TextStream file. After a newline character has been written, but before any other character
' is written, Column is equal to 1.
' ========================================================================================
PRIVATE PROPERTY CTextStream.Column () AS LONG
   DIM nColumn AS LONG
   IF m_pTxtStm THEN SetResult(m_pTxtStm->get_Column(@nColumn))
   PROPERTY = nColumn
END PROPERTY
' ========================================================================================

' ========================================================================================
' Returns TRUE if the file pointer is at the end of a TextStream file; FALSE if it is not. Read-only.
' ========================================================================================
PRIVATE FUNCTION CTextStream.EOS () AS BOOLEAN
   DIM nEOS AS VARIANT_BOOL
   IF m_pTxtStm THEN SetResult(m_pTxtStm->get_AtEndOfStream(@nEOS))
   RETURN nEOS
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the file pointer is positioned immediately before the end-of-line marker
' in a TextStream file; FALSE if it is not. Read-only.
' ========================================================================================
PRIVATE FUNCTION CTextStream.EOL () AS BOOLEAN
   DIM nEOL AS VARIANT_BOOL
   IF m_pTxtStm THEN SetResult(m_pTxtStm->get_AtEndOfLine(@nEOL))
   RETURN nEOL
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads a specified number of characters from a TextStream file and returns the resulting string.
' After a file is initially opened and before anything is written, Line is equal to 1.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Read (BYVAL numChars AS LONG) AS DWSTRING
   IF m_pTxtStm = NULL THEN m_Result = E_POINTER : RETURN ""
   DIM bstrText AS AFX_BSTR
   SetResult(m_pTxtStm->Read(numChars, @bstrText))
   DIM dwsText AS DWSTRING = bstrText
   SysFreeString(bstrText)
   RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads an entire line (up to, but not including, the newline character) from a TextStream
' file and returns the resulting string.
' ========================================================================================
PRIVATE FUNCTION CTextStream.ReadLine () AS DWSTRING
   IF m_pTxtStm = NULL THEN m_Result = E_POINTER : RETURN ""
   DIM bstrText AS AFX_BSTR
   SetResult(m_pTxtStm->ReadLine(@bstrText))
   DIM dwsText AS DWSTRING = bstrText
   SysFreeString(bstrText)
   RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads an entire TextStream file and returns the resulting string.
' ========================================================================================
PRIVATE FUNCTION CTextStream.ReadAll () AS DWSTRING
   IF m_pTxtStm = NULL THEN m_Result = E_POINTER : RETURN ""
   DIM bstrText AS AFX_BSTR
   SetResult(m_pTxtStm->ReadAll(@bstrText))
   DIM dwsText AS DWSTRING = bstrText
   SysFreeString(bstrText)
   RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a specified string to a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Write (BYREF wszText AS CONST WSTRING) AS HRESULT
   IF m_pTxtStm = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   DIM bstrText AS AFX_BSTR = SysAllocString(@wszText)
   SetResult(m_pTxtStm->Write(bstrText))
   SysFreeString(bstrText)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a specified string and newline character to a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.WriteLine (BYREF wszText AS CONST WSTRING) AS HRESULT
   IF m_pTxtStm = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   DIM bstrText AS AFX_BSTR = SysAllocString(@wszText)
   SetResult(m_pTxtStm->WriteLine(bstrText))
   SysFreeString(bstrText)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a specified number of newline characters to a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.WriteBlankLines (BYVAL numLines AS LONG) AS HRESULT
   IF m_pTxtStm = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   RETURN SetResult(m_pTxtStm->WriteBlankLines(numLines))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Skips a specified number of characters when reading a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.Skip (BYVAL numChars AS LONG) AS HRESULT
   IF m_pTxtStm = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   RETURN SetResult(m_pTxtStm->Skip(numChars))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Skips the next line when reading a TextStream file.
' ========================================================================================
PRIVATE FUNCTION CTextStream.SkipLine () AS HRESULT
   IF m_pTxtStm = NULL THEN m_Result = E_POINTER : RETURN E_POINTER
   RETURN SetResult(m_pTxtStm->SkipLine)
END FUNCTION
' ========================================================================================

END NAMESPACE
