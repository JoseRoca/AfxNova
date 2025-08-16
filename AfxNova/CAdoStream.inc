' ########################################################################################
' Platform: Microsoft Windows
' File: CAdoStream.inc
' Contents: ADO Stream class
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

' ########################################################################################
' CAdoStream class
' ########################################################################################
TYPE CAdoStream EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pStream AS Afx_ADOStream PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pStream AS Afx_ADOStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOStream PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOStream PTR PTR

   ' // Attaches a recordset to the class
   DECLARE SUB Attach (BYVAL pStream AS Afx_ADOStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Indicates the size of the stream in number of bytes.
   DECLARE PROPERTY Size () AS LONG
   ' // Indicates whether the current position is at the end of the stream.
   DECLARE PROPERTY EOS () AS BOOLEAN
   ' // Indicates the current position within a Stream object.
   DECLARE PROPERTY Position () AS LONG
   DECLARE PROPERTY Position (BYVAL nPos AS LONG)
   ' // Indicates the type of data contained in the Stream (binary or text).
   DECLARE PROPERTY Type_ () AS StreamTypeEnum
   DECLARE PROPERTY Type_ (BYVAL nType AS StreamTypeEnum)
   ' // Indicates the binary character to be used as the line separator in text Stream objects.
   DECLARE PROPERTY LineSeparator () AS LineSeparatorEnum
   DECLARE PROPERTY LineSeparator (BYVAL pLS AS LineSeparatorEnum)
   ' // Indicates for whether the state of the Stream object is open or closed.
   DECLARE PROPERTY State () AS ObjectStateEnum
   ' // Indicates the available permissions for modifying data in a Stream object.
   DECLARE PROPERTY Mode () AS ConnectModeEnum
   DECLARE PROPERTY Mode (BYVAL nMode AS ConnectModeEnum)
   ' // Indicates the character set into which the contents of a text Stream should be
   ' // translated for storage in the Stream object's internal buffer.
   DECLARE PROPERTY Charset () AS DWSTRING
   DECLARE PROPERTY Charset (BYREF wszCharset AS WSTRING)
   ' // Reads a specified number of bytes from a binary Stream object.
   DECLARE FUNCTION Read (BYVAL NumBytes AS LONG = adReadAll) AS DVARIANT
   DECLARE FUNCTION Read (BYVAL NumBytes AS LONG = adReadAll, BYREF dvValue AS DVARIANT) AS HRESULT
   ' // Opens a connection to a data source.
   DECLARE FUNCTION Open (BYREF dvSource AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
           BYVAL nMode AS ConnectModeEnum = adModeUnknown, BYVAL Options AS StreamOpenOptionsEnum = adOpenStreamUnspecified, _
           BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WSTRING = "") AS HRESULT
   ' // Closes a Stream object and any dependent objects.
   DECLARE FUNCTION Close () AS HRESULT
   ' // Skips one entire line when reading a text stream.
   DECLARE FUNCTION SkipLine () AS HRESULT
   ' // Writes binary data to a Stream object.
   DECLARE FUNCTION Write (BYREF dvBuffer AS DVARIANT) AS HRESULT
   ' // Writes a string to a Stream object.
   DECLARE FUNCTION WriteText (BYREF wszData AS WSTRING, BYVAL Options AS StreamWriteEnum = adWriteChar) AS HRESULT
   ' // Sets the position that is the end of the stream.
   DECLARE FUNCTION SetEOS () AS HRESULT
   ' // Copies the specified number of characters or bytes (depending on Type) in the Stream to another Stream object.
   DECLARE FUNCTION CopyTo (BYVAL pDestStream AS Afx_ADOStream PTR, BYVAL CharNumber AS LONG = adReadAll) AS HRESULT
   ' // Forces the contents of the Stream remaining in the ADO buffer to the underlying object with which the Stream is associated.
   DECLARE FUNCTION Flush () AS HRESULT
   ' // Saves the binary contents of a Stream to a file.
   DECLARE FUNCTION SaveToFile (BYREF wszFileName AS WSTRING, BYVAL Options AS SaveOptionsEnum = adSaveCreateNotExist) AS HRESULT
   ' // Loads the contents of an existing file into a Stream.
   DECLARE FUNCTION LoadFromFile (BYREF wszFileName AS WSTRING) AS HRESULT
   ' // Reads a specified number of characters, an entire line, or the entire stream from a Stream object and returns the resulting string.
   DECLARE FUNCTION ReadText (BYVAL NumChars AS LONG = adReadAll) AS DWSTRING
   ' // Cancels execution of a pending, asynchronous method call.
   DECLARE FUNCTION Cancel () AS HRESULT

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoStream
   CADODB_DP("BEGIN Default")
   ' // Create an instance of the Stream object
   m_pStream = AfxNewCom("ADODB.Stream")
   CADODB_DP("END  " & WSTR(m_pStream))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoStream (BYVAL pStream AS Afx_ADOStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("")
   m_pStream = pStream
   IF fAddRef THEN
      IF m_pStream THEN m_pStream->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoStream
   CADODB_DP(WSTR(m_pStream))
   IF m_pStream THEN
      IF this.State = adStateOpen THEN this.Close
      m_pStream->Release
   END IF
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoStream.@ () AS Afx_ADOStream PTR PTR
'   OPERATOR = @m_pStream
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoStream.vptr () AS Afx_ADOStream PTR PTR
   CADODB_DP("")
   IF m_pStream THEN m_pStream->Release
   RETURN @m_pStream
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cStream AS CAdoStream) AS Afx_ADOStream PTR
   OPERATOR = cStream.m_pStream
END OPERATOR
' ========================================================================================

' ========================================================================================
' Attaches an ADO stream to the class.
' Useful to attach to the class a pointer returned by a function.
' ========================================================================================
PRIVATE SUB CAdoStream.Attach (BYVAL pStream AS Afx_ADOStream PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("")
   IF m_pStream THEN m_pStream->Release
   m_pStream = pStream
   IF fAddRef THEN
      IF m_pStream THEN m_pStream->AddRef
   END IF
END SUB
' ========================================================================================

' ===========================================================================================
' Indicates the size of the stream in number of bytes.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Size () AS LONG
   CADODB_DP("")
   DIM lSize AS LONG
   IF m_pStream THEN SetResult(m_pStream->get_Size(@lSize))
   PROPERTY = lSize
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates whether the current position is at the end of the stream.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.EOS () AS BOOLEAN
   CADODB_DP("")
   DIM bEOS AS VARIANT_BOOL
   IF m_pStream THEN SetResult(m_pStream->get_EOS(@bEOS))
   PROPERTY = bEOS
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the current position within a Stream object.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Position () AS LONG
   CADODB_DP("")
   DIM nPos AS LONG
   IF m_pStream THEN SetResult(m_pStream->get_Position(@nPos))
   PROPERTY = nPos
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Position (BYVAL nPos AS LONG)
   CADODB_DP("")
   IF m_pStream THEN SetResult(m_pStream->put_Position(nPos))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the type of data contained in the Stream (binary or text).
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Type_ () AS StreamTypeEnum
   CADODB_DP("")
   DIM nType AS StreamTypeEnum
   IF m_pStream THEN SetResult(m_pStream->get_Type(@nType))
   PROPERTY = nType
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Type_ (BYVAL nType AS StreamTypeEnum)
   CADODB_DP("")
   IF m_pStream THEN SetResult(m_pStream->put_Type(nType))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the binary character to be used as the line separator in text Stream objects.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.LineSeparator () AS LineSeparatorEnum
   CADODB_DP("")
   DIM LS AS LineSeparatorEnum
   IF m_pStream THEN SetResult(m_pStream->get_LineSeparator(@LS))
   PROPERTY = LS
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.LineSeparator (BYVAL LS AS LineSeparatorEnum)
   CADODB_DP("")
   IF m_pStream THEN SetResult(m_pStream->put_LineSeparator(LS))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates for whether the state of the Stream object is open or closed.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.State () AS ObjectStateEnum
   CADODB_DP("")
   DIM lState AS ObjectStateEnum
   IF m_pStream THEN SetResult(m_pStream->get_State(@lState))
   PROPERTY = lState
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the available permissions for modifying data in a Stream object.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Mode () AS ConnectModeEnum
   CADODB_DP("")
   DIM nMode AS ConnectModeEnum
   IF m_pStream THEN SetResult(m_pStream->get_Mode(@nMode))
   PROPERTY = nMode
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Mode (BYVAL nMode AS ConnectModeEnum)
   CADODB_DP("")
   IF m_pStream THEN SetResult(m_pStream->put_Mode(nMode))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the available permissions for modifying data in a Stream object.
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Charset () AS DWSTRING
   CADODB_DP("")
   DIM bstrCharset AS AFX_BSTR
   IF m_pStream THEN SetResult(m_pStream->get_Charset(@bstrCharset))
   DIM dwsCharSet AS DWSTRING = *bstrCharset
   SysFreeString bstrCharset
   PROPERTY = dwsCharSet
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoStream.Charset (BYREF wszCharset AS WSTRING)
   CADODB_DP("")
   DIM bstrCharSet AS AFX_BSTR = SysAllocString(wszCharset)
   IF m_pStream THEN SetResult(m_pStream->put_Charset(bstrCharSet))
   SysFreeString bstrCharSet
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Reads a specified number of bytes from a binary Stream object.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Read (BYVAL NumBytes AS LONG = adReadAll) AS DVARIANT
   CADODB_DP("")
   DIM dvValue AS DVARIANT
   IF m_pStream THEN SetResult(m_pStream->Read(NumBytes, dvValue))
   RETURN dvValue
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Read (BYVAL NumBytes AS LONG = adReadAll, BYREF dvValue AS DVARIANT) AS HRESULT
   CADODB_DP("")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Read(NumBytes, dvValue.vptr)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Open (BYREF dvSource AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYVAL nMode AS ConnectModeEnum = adModeUnknown, BYVAL Options AS StreamOpenOptionsEnum = adOpenStreamUnspecified, _
BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WSTRING = "") AS HRESULT
   CADODB_DP("")
   DIM bstrUserName AS AFX_BSTR = SysAllocString(wszUserName)
   DIM bstrPassword AS AFX_BSTR = SysAllocString(wszPassword)
   IF m_pStream THEN SetResult(m_pStream->Open(dvSource, nMode, Options, bstrUserName, bstrPassword))
   SysFreeString bstrUserName
   SysFreeString bstrPassword
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Closes a Stream object and any dependent objects.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Close () AS HRESULT
   CADODB_DP("")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Close))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Skips one entire line when reading a text stream.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.SkipLine () AS HRESULT
   CADODB_DP("")
   IF m_pStream THEN RETURN(SetResult(m_pStream->SkipLine))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Writes binary data to a Stream object.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Write (BYREF dvBuffer AS DVARIANT) AS HRESULT
   CADODB_DP("")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Write(dvBuffer)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Writes a string to a Stream object.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.WriteText (BYREF wszData AS WSTRING, BYVAL Options AS StreamWriteEnum = adWriteChar) AS HRESULT
   CADODB_DP("")
   DIM bstrData AS AFX_BSTR = SysAllocString(wszData)
   IF m_pStream THEN SetResult(m_pStream->WriteText(bstrData, Options))
   SysFreeString bstrData
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Sets the position that is the end of the stream.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.SetEOS () AS HRESULT
   CADODB_DP("")
   IF m_pStream THEN RETURN(SetResult(m_pStream->SetEOS))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Copies the specified number of characters or bytes (depending on Type) in the Stream to
' another Stream object.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.CopyTo (BYVAL pDestStream AS Afx_ADOStream PTR, BYVAL CharNumber AS LONG = adReadAll) AS HRESULT
   CADODB_DP("")
   IF m_pStream THEN RETURN(SetResult(m_pStream->CopyTo(pDestStream, CharNumber)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Forces the contents of the Stream remaining in the ADO buffer to the underlying object with
' which the Stream is associated.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Flush () AS HRESULT
   CADODB_DP("")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Flush))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves the binary contents of a Stream to a file.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.SaveToFile (BYREF wszFileName AS WSTRING, BYVAL Options AS SaveOptionsEnum = adSaveCreateNotExist) AS HRESULT
   CADODB_DP("")
   DIM bstrFileName AS AFX_BSTR = SysAllocString(wszFileName)
   IF m_pStream THEN SetResult(m_pStream->SaveToFile(bstrFileName, Options))
   SysFreeString bstrFileName
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Loads the contents of an existing file into a Stream.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.LoadFromFile (BYREF wszFileName AS WSTRING) AS HRESULT
   CADODB_DP("")
   DIM bstrFileName AS AFX_BSTR = SysAllocString(wszFileName)
   IF m_pStream THEN RETURN(SetResult(m_pStream->LoadFromFile(bstrFileName)))
   SysFreeString bstrFileName
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Reads a specified number of characters, an entire line, or the entire stream from a Stream
' object and returns the resulting string.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.ReadText (BYVAL NumChars AS LONG = adReadAll) AS DWSTRING
   CADODB_DP("")
   DIM bstrText AS AFX_BSTR
   IF m_pStream THEN SetResult(m_pStream->ReadText(NumChars, @bstrText))
   DIM dwsText AS DWSTRING = *bstrText
   RETURN dwsText
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call.
' ===========================================================================================
PRIVATE FUNCTION CAdoStream.Cancel () AS HRESULT
   CADODB_DP("")
   IF m_pStream THEN RETURN(SetResult(m_pStream->Cancel))
END FUNCTION
' ===========================================================================================
