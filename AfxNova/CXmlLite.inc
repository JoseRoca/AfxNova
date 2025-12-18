' ########################################################################################
' Platform: Microsoft Windows
' File: CXmlLite.inc
' Contents: XmlLite wrapper classes
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
#include once "AfxNova/AfxXmlLite.bi"
USING AfxNova

' ========================================================================================
' Macro for debug
' To allow debugging, define _CXMLLITE_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CXMLLITE_DEBUG_
   #define _CXMLLITE_DEBUG_ 0
#ENDIF
#ifndef _CXMLLITE_DP_
   #define _CXMLLITE_DP_ 1
   #MACRO CXMLLITE_DP(st)
      #IF (_CXMLLITE_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

NAMESPACE AfxNova

' ########################################################################################
' CXmlReader class
' ########################################################################################
TYPE CXmlReader

   m_Result AS HRESULT
   m_hLib AS HMODULE
   m_bUninitCOM AS BOOLEAN
   m_pXmlReader AS Afx_IXmlReader PTR
   m_pInputStream AS IUnknown PTR

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   ' // IXmlReader methods and properties
   DECLARE FUNCTION SetInput (BYVAL pInputStream AS IUnknown PTR) AS HRESULT
   DECLARE FUNCTION SetInput (BYVAL pInputStream AS IUnknown PTR, BYVAL nEncodingPage AS UINT) AS HRESULT
   DECLARE FUNCTION SetInput (BYVAL pInputStream AS IUnknown PTR, BYVAL nEncodingName AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION GetProperty (BYVAL nProperty AS UINT) AS LONG_PTR
   DECLARE FUNCTION SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
   DECLARE FUNCTION Read (BYVAL nodeType AS XmlNodeType PTR) AS HRESULT
   DECLARE FUNCTION GetNodeType () AS XmlNodeType 
   DECLARE FUNCTION MoveToFirstAttribute () AS HRESULT
   DECLARE FUNCTION MoveToNextAttribute () AS HRESULT
   DECLARE FUNCTION MoveToAttributeByName (BYVAL pwszLocalName AS LPCWSTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION MoveToElement () AS HRESULT
   DECLARE FUNCTION GetQualifiedName () AS DWStRING
   DECLARE FUNCTION GetNamespaceUri () AS DWSTRING
   DECLARE FUNCTION GetLocalName () AS DWSTRING
   DECLARE FUNCTION GetPrefix () AS DWSTRING
   DECLARE FUNCTION GetValue () AS DWSTRING
   DECLARE FUNCTION ReadValueChunk (BYVAL buffer AS WSTRING PTR, BYVAL chunkSize AS UINT, BYVAL pcwchRead AS UINT PTR) AS HRESULT
   DECLARE FUNCTION GetBaseUri () AS DWSTRING
   DECLARE FUNCTION IsDefault () AS BOOLEAN
   DECLARE FUNCTION IsEmptyElement () AS BOOLEAN
   DECLARE FUNCTION GetLineNumber () AS UINT
   DECLARE FUNCTION GetLinePosition () AS UINT
   DECLARE FUNCTION GetAttributeCount () AS UINT
   DECLARE FUNCTION GetDepth () AS UINT
   DECLARE FUNCTION IsEOF () AS BOOLEAN

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CXmlReader
   CXMLLITE_DP("")
   ' // Initialize the COM library
   IF m_bUninitCOM = FALSE THEN
      DIM hr AS HRESULT = CoInitialize(NULL)
      IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   END IF
   ' // Load the XmlLite library
   IF m_hLib = NULL THEN m_hLib = LoadLibraryW("XmlLite.dll")
   IF m_hLib = NULL THEN EXIT CONSTRUCTOR
   ' // Create the reader
   DIM pfnCreateXmlReader AS FUNCTION (BYREF riid AS IID, BYVAL ppvObject AS void PTR PTR, BYVAL pMalloc AS IMalloc PTR) AS HRESULT
   pfnCreateXmlReader = CAST(ANY PTR, GetProcAddress(m_hLib, "CreateXmlReader"))
   IF pfnCreateXmlReader = NULL THEN EXIT CONSTRUCTOR
   m_Result = pfnCreateXmlReader(AFX_IID_IXMlReader, @m_pXmlReader, NULL)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CXmlReader
   CXMLLITE_DP("")
   ' // Release the input source
   IF m_pInputStream THEN m_pInputStream->lpvtbl->Release(m_pInputStream)
   ' // Release the Afx_IXmlReader interface
   IF m_pXmlReader THEN m_pXmlReader->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
   ' // Free the XmlLite library
   IF m_hLib THEN FreeLibrary(m_hLib)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
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
' Sets the input source of the XML document to be parsed.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.SetInput (BYVAL pInputStream AS IUnknown PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      IF m_pInputStream THEN m_pInputStream->lpvtbl->Release(m_pInputStream)
      m_Result = m_pXmlReader->SetInput(pInputStream)
      IF m_Result = S_OK THEN m_pInputStream = pInputStream
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the input source of the XML document to be parsed.
' - pInputStream: A pointer to the input stream for the reader to read the data from.
'                 This can be either an ISequentialStream or an IStream interface.
' - nEncodingPage: Specifies the encoding of the input.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.SetInput (BYVAL pInputStream AS IUnknown PTR, BYVAL nEncodingPage AS UINT) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader = NULL THEN RETURN m_Result
   ' // Create the input interface
   DIM pfnCreateXmlReaderInputWithEncodingPage AS FUNCTION (BYVAL pInputStream AS IStream PTR, BYVAL pMalloc AS IMalloc PTR, _
       BYVAL nEncodingPage AS UINT, BYVAL fEncodingHint AS BOOL, BYVAL pwszBaseUri AS LPCWSTR, BYVAL ppInput AS IUnknown PTR PTR) AS HRESULT
   pfnCreateXmlReaderInputWithEncodingPage = CAST(ANY PTR, GetProcAddress(m_hLib, "CreateXmlReaderInputWithEncodingCodePage"))
   IF pfnCreateXmlReaderInputWithEncodingPage = NULL THEN RETURN m_Result
   ' // Release a previous input stream
   IF m_pInputStream THEN m_pInputStream->lpvtbl->Release(m_pInputStream)
   m_pInputStream = NULL
   m_Result = pfnCreateXmlReaderInputWithEncodingPage(cast(IStream PTR, pInputStream), NULL, nEncodingPage, TRUE, NULL, @m_pInputStream)
   IF FAILED(m_Result) THEN RETURN m_Result
   m_Result = m_pXmlReader->SetInput(m_pInputStream)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the input source of the XML document to be parsed.
' - pInputStream: A pointer to the input stream for the reader to read the data from.
'                 This can be either an ISequentialStream or an IStream interface.
' - nEncodingName: Specifies the encoding of the input.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.SetInput (BYVAL pInputStream AS IUnknown PTR, BYVAL encodingName AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader = NULL THEN RETURN m_Result
   ' // Create the input interface
   DIM pfnCreateXmlReaderInputWithEncodingName AS FUNCTION (BYVAL pInputStream AS IStream PTR, BYVAL pMalloc AS IMalloc PTR, _
       BYVAL pwszEncodingName AS LPCWSTR, BYVAL fEncodingHint AS BOOL, BYVAL pwszBaseUri AS LPCWSTR, BYVAL ppInput AS IUnknown PTR PTR) AS HRESULT
   pfnCreateXmlReaderInputWithEncodingName = CAST(ANY PTR, GetProcAddress(m_hLib, "CreateXmlReaderInputWithEncodingName"))
   IF pfnCreateXmlReaderInputWithEncodingName = NULL THEN RETURN m_Result
   ' // Release a previous input stream
   IF m_pInputStream THEN m_pInputStream->lpvtbl->Release(m_pInputStream)
   m_pInputStream = NULL
   m_Result = pfnCreateXmlReaderInputWithEncodingName(cast(IStream PTR, pInputStream), NULL, encodingName, TRUE, NULL, @m_pInputStream)
   IF FAILED(m_Result) THEN RETURN m_Result
   m_Result = m_pXmlReader->SetInput(m_pInputStream)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the specified property. A program can get properties at any time.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetProperty (BYVAL nProperty AS UINT) AS LONG_PTR
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      DIM pValue AS LONG_PTR
      m_Result = m_pXmlReader->GetProperty(nProperty, @pValue)
      RETURN pValue
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the specified property.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      m_Result = m_pXmlReader->SetProperty(nProperty, pValue)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Reads the next node from the stream and returns the type of the node.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.Read (BYVAL nodeType AS XmlNodeType PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      m_Result = m_pXmlReader->Read(nodeType)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Provides the type of the current node.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetNodeType () AS XmlNodeType 
   m_Result = E_POINTER
   DIM nodeType AS XmlNodeType
   IF m_pXmlReader THEN
      m_Result = m_pXmlReader->GetNodeType(@nodeType)
   END IF
   RETURN nodeType
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the reader position to the first attribute within the current node.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.MoveToFirstAttribute () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      m_Result = m_pXmlReader->MoveToFirstAttribute
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Advances the reader to the next attribute.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.MoveToNextAttribute () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      m_Result = m_pXmlReader->MoveToNextAttribute
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves the reader to the attribute with the specified name.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.MoveToAttributeByName (BYVAL pwszLocalName AS LPCWSTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      m_Result = m_pXmlReader->MoveToAttributeByName(pwszLocalName, pwszNamespaceUri)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Moves to the element that owns the current attribute node. After navigating through the
' attributes on an element, use this method to move the reader back to the element.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.MoveToElement () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      m_Result = m_pXmlReader->MoveToElement
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the qualified name of the node that the reader is currently positioned on. If no
' qualified name is available, returns an empty string.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetQualifiedName () AS DWSTRING
   m_Result = E_POINTER
   DIM dws AS DWSTRING
   IF m_pXmlReader THEN
      DIM pwszQualifiedName AS WSTRING PTR
      m_Result = m_pXmlReader->GetQualifiedName(@pwszQualifiedName, NULL)
      IF m_Result = S_OK THEN dws = *pwszQualifiedName
   END IF
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the namespace URI of the node that the reader is currently positioned on. If no
' namespace URI is available, returns an empty string.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetNamespaceUri () AS DWSTRING
   m_Result = E_POINTER
   DIM dws AS DWSTRING
   IF m_pXmlReader THEN
      DIM pwszNamespaceUri AS WSTRING PTR
      m_Result = m_pXmlReader->GetNamespaceUri(@pwszNamespaceUri, NULL)
      IF m_Result = S_OK THEN dws = *pwszNamespaceUri
   END IF
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' This method allows the client to determine the difference between elements that have a
' closing tag, but do not contain content, and elements that do not have a closing tag.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetLocalName () AS DWSTRING
   m_Result = E_POINTER
   DIM dws AS DWSTRING
   IF m_pXmlReader THEN
      DIM pwszLocalName AS WSTRING PTR
      m_Result = m_pXmlReader->GetLocalName(@pwszLocalName, NULL)
      IF m_Result = S_OK THEN dws = *pwszLocalName
   END IF
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the namespace prefix of the node that the reader is currently positioned on. If no
' prefix is available, returns an empty string.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetPrefix () AS DWSTRING
   m_Result = E_POINTER
   DIM dws AS DWSTRING
   IF m_pXmlReader THEN
      DIM pwszPrefix AS WSTRING PTR
      m_Result = m_pXmlReader->GetLocalName(@pwszPrefix, NULL)
      IF m_Result = S_OK THEN dws = *pwszPrefix
   END IF
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the value of the current token, if applicable.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetValue () AS DWSTRING
   m_Result = E_POINTER
   DIM dws AS DWSTRING
   IF m_pXmlReader THEN
      DIM pwszValue AS WSTRING PTR
      m_Result = m_pXmlReader->GetValue(@pwszValue, NULL)
      IF m_Result = S_OK THEN dws = *pwszValue
   END IF
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' This method will read up to a maximum of the specified chunk size (as available) from the
' value of the current node and copy the value into the specified buffer.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.ReadValueChunk (BYVAL buffer AS WSTRING PTR, BYVAL chunkSize AS UINT, BYVAL pcwchRead AS UINT PTR) AS HRESULT
   m_Result = E_POINTER
   m_Result = ReadValueChunk(buffer, chunkSize, pcwchRead)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the base URI of the token, if applicable.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetBaseUri () AS DWSTRING
   m_Result = E_POINTER
   DIM dws AS DWSTRING
   IF m_pXmlReader THEN
      DIM pwszBaseUri AS WSTRING PTR
      m_Result = m_pXmlReader->GetBaseUri(@pwszBaseUri, NULL)
      IF m_Result = S_OK THEN dws = *pwszBaseUri
   END IF
   RETURN dws
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether the attribute was specified in the source document or implied by the
' Document Type Definition (DTD).
' ========================================================================================
PRIVATE FUNCTION CXmlReader.IsDefault () AS BOOLEAN
   m_Result = 0
   IF m_pXmlReader THEN
      RETURN m_pXmlReader->IsDefault
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' This method allows the client to determine the difference between elements that have a
' closing tag, but do not contain content, and elements that do not have a closing tag.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.IsEmptyElement () AS BOOLEAN
   m_Result = 0
   IF m_pXmlReader THEN
      RETURN m_pXmlReader->IsEmptyElement
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the line number where the reader is positioned in the document.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetLineNumber () AS UINT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      DIM nLineNumber AS UINT
      m_Result = m_pXmlReader->GetLineNumber(@nLineNumber)
      RETURN nLineNumber
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the line position where the reader is positioned in the document.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetLinePosition () AS UINT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      DIM nLinePosition AS UINT
      m_Result = m_pXmlReader->GetLinePosition(@nLinePosition)
      RETURN nLinePosition
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of attributes in the current node.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetAttributeCount () AS UINT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      DIM nAttributeCount AS UINT
      m_Result = m_pXmlReader->GetAttributeCount(@nAttributeCount)
      RETURN nAttributeCount
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the depth of the current node in the document.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.GetDepth () AS UINT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      DIM nDepth AS UINT
      m_Result = m_pXmlReader->GetDepth(@nDepth)
      RETURN nDepth
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns TRUE if the end of the input is reached.
' ========================================================================================
PRIVATE FUNCTION CXmlReader.IsEOF () AS BOOLEAN
   m_Result = 0
   IF m_pXmlReader THEN
      RETURN m_pXmlReader->IsEOF
   END IF
END FUNCTION
' ========================================================================================


' ########################################################################################
' CXmlWriter class
' ########################################################################################
TYPE CXmlWriter

   m_Result AS HRESULT
   m_hLib AS HMODULE
   m_bUninitCOM AS BOOLEAN
   m_pXmlWriter AS Afx_IXmlWriter PTR
   m_pOutputStream AS IUnknown PTR

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   ' // IXmlWriter methods and properties
   DECLARE FUNCTION SetOutput (BYVAL pOutputStream AS IUnknown PTR) AS HRESULT
   DECLARE FUNCTION SetOutput (BYVAL pOutputStream AS IUnknown PTR, BYVAL nEncodingPage AS UINT) AS HRESULT
   DECLARE FUNCTION SetOutput (BYVAL pOutputStream AS IUnknown PTR, BYVAL pwszEncodingName AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION GetProperty (BYVAL nProperty AS UINT) AS LONG_PTR
   DECLARE FUNCTION SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
   DECLARE FUNCTION WriteAttributes (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION WriteAttributeString (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, _
                    BYVAL pwszNamespaceUri AS LPCWSTR, BYVAL pwszValue AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteCData (BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteCharEntity (BYVAL wch AS WCHAR PTR) AS HRESULT
   DECLARE FUNCTION WriteChars (BYVAL pwch AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION WriteComment (BYVAL pwszComment AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteDocType (BYVAL pwszName AS LPCWSTR, BYVAL pwszPublicId AS LPCWSTR, _
                    BYVAL pwszSystemId AS LPCWSTR, BYVAL pwszSubset AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteElementString (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, _
                    BYVAL pwszNamespaceUri AS LPCWSTR, BYVAL pwszValue AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteEndDocument () AS HRESULT
   DECLARE FUNCTION WriteEndElement () AS HRESULT
   DECLARE FUNCTION WriteEntityRef (BYVAL pwszName AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteFullEndElement () AS HRESULT
   DECLARE FUNCTION WriteName (BYVAL pwszName AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteNmToken (BYVAL pwszNmToken AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteNode (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION WriteNodeShallow (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION WriteProcessingInstruction (BYVAL pwszName AS LPCWSTR, BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteQualifiedName (BYVAL pwszLocalName AS LPCWSTR PTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteRaw (BYVAL pwszData AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteRawChars (BYVAL pwch AS CONST WSTRING PTR) AS HRESULT
   DECLARE FUNCTION WriteStartDocument (BYVAL standalone AS XmlStandalone) AS HRESULT
   DECLARE FUNCTION WriteStartElement (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteString (BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteSurrogateCharEntity (BYVAL wchLow AS WCHAR, BYVAL wchHigh AS WCHAR) AS HRESULT
   DECLARE FUNCTION WriteWhitespace (BYVAL pwszWhitespace AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION Flush () AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CXmlWriter
   CXMLLITE_DP("")
   ' // Initialize the COM library
   IF m_bUninitCOM = FALSE THEN
      DIM hr AS HRESULT = CoInitialize(NULL)
      IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   END IF
   ' // Load the XmlLite library
   IF m_hLib = NULL THEN m_hLib = LoadLibraryW("XmlLite.dll")
   IF m_hLib = NULL THEN EXIT CONSTRUCTOR
   ' // Create the reader
   DIM pfnCreateXmlWriter AS FUNCTION (BYREF riid AS IID, BYVAL ppvObject AS void PTR PTR, BYVAL pMalloc AS IMalloc PTR) AS HRESULT
   pfnCreateXmlWriter = CAST(ANY PTR, GetProcAddress(m_hLib, "CreateXmlWriter"))
   IF pfnCreateXmlWriter = NULL THEN EXIT CONSTRUCTOR
   m_Result = pfnCreateXmlWriter(AFX_IID_IXMlWriter, @m_pXmlWriter, NULL)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CXmlWriter
   CXMLLITE_DP("")
   ' // Release the input source
   IF m_pOutputStream THEN m_pOutputStream->lpvtbl->Release(m_pOutputStream)
   ' // Release the Afx_IXmlReader interface
   IF m_pXmlWriter THEN m_pXmlWriter->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
   ' // Free the XmlLite library
   IF m_hLib THEN FreeLibrary(m_hLib)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
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
' Sets the output stream for the writer.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.SetOutput (BYVAL pOutputStream AS IUnknown PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      IF m_pOutputStream THEN m_pOutputStream->lpvtbl->Release(m_pOutputStream)
      m_Result = m_pXmlWriter->SetOutput(pOutputStream)
      IF m_Result = S_OK THEN m_pOutputStream = pOutputStream
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the output source of the XML document to be writen.
' - pOutputStream: A pointer to the output stream for the writer to write data to.
'                  This can be either an ISequentialStream or an IStream interface.
' - nEncodingPage: Specifies the encoding of the output.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.SetOutput (BYVAL pOutputStream AS IUnknown PTR, BYVAL nEncodingPage AS UINT) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter = NULL THEN RETURN m_Result
   ' // Create the input interface
   DIM pfnCreateXmlWriterOutputWithEncodingCodePage AS FUNCTION (BYVAL pOutputStream AS IStream PTR, BYVAL pMalloc AS IMalloc PTR, _
       BYVAL nEncodingPage AS UINT, BYVAL ppOutput AS IUnknown PTR PTR) AS HRESULT
   pfnCreateXmlWriterOutputWithEncodingCodePage = CAST(ANY PTR, GetProcAddress(m_hLib, "CreateXmlWriterOutputWithEncodingCodePage"))
   IF pfnCreateXmlWriterOutputWithEncodingCodePage = NULL THEN RETURN m_Result
   ' // Release a previous output stream
   IF m_pOutputStream THEN m_pOutputStream->lpvtbl->Release(m_pOutputStream)
   m_pOutputStream = NULL
   m_Result = pfnCreateXmlWriterOutputWithEncodingCodePage(cast(IStream PTR, pOutputStream), NULL, nEncodingPage, @m_pOutputStream)
   IF FAILED(m_Result) THEN RETURN m_Result
   m_Result = m_pXmlWriter->SetOutput(m_pOutputStream)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the output source of the XML document to be writen.
' - pOutputStream: A pointer to the output stream for the writer to write data to.
'                  This can be either an ISequentialStream or an IStream interface.
' - nEncodingPage: Specifies the encoding of the output.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.SetOutput (BYVAL pOutputStream AS IUnknown PTR, BYVAL pwszEncodingName AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter = NULL THEN RETURN m_Result
   ' // Create the input interface
   DIM pfnCreateXmlWriterOutputWithEncodingName AS FUNCTION (BYVAL pOutputStream AS IStream PTR, BYVAL pMalloc AS IMalloc PTR, _
       BYVAL pwszEncodingName AS LPCWSTR, BYVAL ppOutput AS IUnknown PTR PTR) AS HRESULT
   pfnCreateXmlWriterOutputWithEncodingName = CAST(ANY PTR, GetProcAddress(m_hLib, "CreateXmlWriterOutputWithEncodingName"))
   IF pfnCreateXmlWriterOutputWithEncodingName = NULL THEN RETURN m_Result
   ' // Release a previous output stream
   IF m_pOutputStream THEN m_pOutputStream->lpvtbl->Release(m_pOutputStream)
   m_pOutputStream = NULL
   m_Result = pfnCreateXmlWriterOutputWithEncodingName(cast(IStream PTR, pOutputStream), NULL, pwszEncodingName, @m_pOutputStream)
   IF FAILED(m_Result) THEN RETURN m_Result
   m_Result = m_pXmlWriter->SetOutput(m_pOutputStream)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the specified property. A program can get properties at any time.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.GetProperty (BYVAL nProperty AS UINT) AS LONG_PTR
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      DIM pValue AS LONG_PTR
      m_Result = m_pXmlWriter->GetProperty(nProperty, @pValue)
      RETURN pValue
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the specified property.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->SetProperty(nProperty, pValue)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies attributes from the specified source IXmlReader to the IXmlWriter.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteAttributes (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteAttributes(pReader, fWriteDefaultAttributes)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes an attribute.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteAttributeString (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, _
BYVAL pwszNamespaceUri AS LPCWSTR, BYVAL pwszValue AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteAttributeString(pwszPrefix, pwszLocalName, pwszNamespaceUri, pwszValue)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out a CDATA section that contains the specified text.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteCData (BYVAL pwszText AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteCData(pwszText)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a character entity for the provided Unicode character value. This method writes
' the character entity in hexadecimal format.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteCharEntity (BYVAL wch AS WCHAR PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteCharEntity(wch)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified text content, escaping markup.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteChars (BYVAL pwch AS WSTRING PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteChars(pwch, LEN(*pwch))
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out a comment that contains the specified text.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteComment (BYVAL pwszComment AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteComment(pwszComment)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the <!DOCTYPE ...> declaration with the specified name and optional attributes.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteDocType (BYVAL pwszName AS LPCWSTR, BYVAL pwszPublicId AS LPCWSTR, _
BYVAL pwszSystemId AS LPCWSTR, BYVAL pwszSubset AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteDocType(pwszName, pwszPublicId, pwszSystemId, pwszSubset)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out an element with the specified prefix, name, namespace, and value.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteElementString (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, _
BYVAL pwszNamespaceUri AS LPCWSTR, BYVAL pwszValue AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteElementString(pwszPrefix, pwszLocalName, pwszNamespaceUri, pwszValue)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes any open elements or attributes, then closes the current document. To reinitialize
' the writer so that a new document can be written, you must call SetOutput.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteEndDocument () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteEndDocument
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes one element. If the element contains no content, this method writes a short
' end tag ("/>"). Otherwise, this method writes the full end tag.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteEndElement () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteEndElement
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out an entity reference with the specified name.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteEntityRef (BYVAL pwszName AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteEntityRef(pwszName)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes one element and pops the corresponding namespace scope. This method always writes
' the full end tag.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteFullEndElement () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteFullEndElement
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified name, ensuring that the name is valid according to the XML specification.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteName (BYVAL pwszName AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteName(pwszName)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified name, ensuring that the name is a valid NmToken according to the XML specification.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteNmToken (BYVAL pwszNmToken AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteNmToken(pwszNmToken)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the current node from the specified source IXmlReader to the IXmlWriter.
' This method moves the IXmlReader source to the start of the next sibling node, reading
' all of the child nodes of the current node. If the reader is in the initial state, this
' method copies all of the nodes in the reader.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteNode (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteNode(pReader, fWriteDefaultAttributes)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies the current node from the specified source IXmlReader to the IXmlWriter, without
' writing the children of the current node.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteNodeShallow (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteNodeShallow(pReader, fWriteDefaultAttributes)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a processing instruction.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteProcessingInstruction (BYVAL pwszName AS LPCWSTR, BYVAL pwszText AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteProcessingInstruction(pwszName, pwszText)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified namespace-qualified name by looking up the prefix that is in
' scope for the specified namespace.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteQualifiedName (BYVAL pwszLocalName AS LPCWSTR PTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteQualifiedName(pwszLocalName, pwszNamespaceUri)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables the caller to write out raw markup manually. You can avoid creating entities for
' special characters by using this method.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteRaw (BYVAL pwszData AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteRaw(pwszData)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out raw markup manually. Using this method allows an application to avoid creating
' entities for special characters.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteRawChars (BYVAL pwch AS CONST WSTRING PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteRawChars(pwch, LEN(*pwch))
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the XML declaration with the version "1.0". The encoding attribute is determined
' by the implementation of IXmlWriterOutput. By default, the encoding is UTF-8.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteStartDocument (BYVAL standalone AS XmlStandalone) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteStartDocument(standalone)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified start tag and associates it with the specified namespace.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteStartElement (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteStartElement(pwszPrefix, pwszLocalName, pwszNamespaceUri)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified text content, escaping any markup in the content.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteString (BYVAL pwszText AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteString(pwszText)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Forces the generation of a surrogate character entity for the specified string value.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteSurrogateCharEntity (BYVAL wchLow AS WCHAR, BYVAL wchHigh AS WCHAR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteSurrogateCharEntity(wchLow, wchHigh)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified white space.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.WriteWhitespace (BYVAL pwszWhitespace AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteWhitespace(pwszWhitespace)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Flushes whatever is in the buffer to the underlying stream and then flushes the underlying stream.
' ========================================================================================
PRIVATE FUNCTION CXmlWriter.Flush () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->Flush
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================


' ########################################################################################
' CXmlWriterLite class
' ########################################################################################
TYPE CXmlWriterLite

   m_Result AS HRESULT
   m_hLib AS HMODULE
   m_bUninitCOM AS BOOLEAN
   m_pXmlWriterLite AS Afx_IXmlWriterLite PTR
   m_pOutputStream AS IUnknown PTR

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   ' // IXmlWriter methods and properties
   DECLARE FUNCTION SetOutput (BYVAL pOutputStream AS IUnknown PTR) AS HRESULT
   DECLARE FUNCTION SetOutput (BYVAL pOutputStream AS IUnknown PTR, BYVAL nEncodingPage AS UINT) AS HRESULT
   DECLARE FUNCTION SetOutput (BYVAL pOutputStream AS IUnknown PTR, BYVAL pwszEncodingName AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION GetProperty (BYVAL nProperty AS UINT) AS LONG_PTR
   DECLARE FUNCTION SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
   DECLARE FUNCTION WriteAttributes (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION WriteAttributeString (BYVAL pwszQName AS WSTRING PTR, _
                    BYVAL pwszValue AS Const WSTRING PTR) AS HRESULT
   DECLARE FUNCTION WriteCData (BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteCharEntity (BYVAL wch AS WCHAR PTR) AS HRESULT
   DECLARE FUNCTION WriteChars (BYVAL pwch AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION WriteComment (BYVAL pwszComment AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteDocType (BYVAL pwszName AS LPCWSTR, BYVAL pwszPublicId AS LPCWSTR, _
                    BYVAL pwszSystemId AS LPCWSTR, BYVAL pwszSubset AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteElementString (BYVAL pwszQName AS WSTRING PTR, BYVAL pwszValue AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteEndDocument () AS HRESULT
   DECLARE FUNCTION WriteEndElement (BYVAL pwszQName AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION WriteEntityRef (BYVAL pwszName AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteFullEndElement (BYVAL pwszQName AS CONST WSTRING PTR) AS HRESULT
   DECLARE FUNCTION WriteName (BYVAL pwszName AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteNmToken (BYVAL pwszNmToken AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteNode (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION WriteNodeShallow (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION WriteProcessingInstruction (BYVAL pwszName AS LPCWSTR, BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteQualifiedName (BYVAL pwszLocalName AS LPCWSTR PTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteRaw (BYVAL pwszData AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteRawChars (BYVAL pwch AS CONST WSTRING PTR) AS HRESULT
   DECLARE FUNCTION WriteStartDocument (BYVAL standalone AS XmlStandalone) AS HRESULT
   DECLARE FUNCTION WriteStartElement (BYVAL pwszQName AS CONST WSTRING PTR) AS HRESULT
   DECLARE FUNCTION WriteString (BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION WriteSurrogateCharEntity (BYVAL wchLow AS WCHAR, BYVAL wchHigh AS WCHAR) AS HRESULT
   DECLARE FUNCTION WriteWhitespace (BYVAL pwszWhitespace AS LPCWSTR) AS HRESULT
   DECLARE FUNCTION Flush () AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CXmlWriterLite
   CXMLLITE_DP("")
   ' // Initialize the COM library
   IF m_bUninitCOM = FALSE THEN
      DIM hr AS HRESULT = CoInitialize(NULL)
      IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   END IF
   ' // Load the XmlLite library
   IF m_hLib = NULL THEN m_hLib = LoadLibraryW("XmlLite.dll")
   IF m_hLib = NULL THEN EXIT CONSTRUCTOR
   ' // Create the reader
   DIM pfnCreateXmlWriterLite AS FUNCTION (BYREF riid AS IID, BYVAL ppvObject AS void PTR PTR, BYVAL pMalloc AS IMalloc PTR) AS HRESULT
   pfnCreateXmlWriterLite = CAST(ANY PTR, GetProcAddress(m_hLib, "CreateXmlWriter"))
   IF pfnCreateXmlWriterLite = NULL THEN EXIT CONSTRUCTOR
   m_Result = pfnCreateXmlWriterLite(AFX_IID_IXMlWriterLite, @m_pXmlWriterLite, NULL)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Destructor
' ========================================================================================
PRIVATE DESTRUCTOR CXmlWriterLite
   CXMLLITE_DP("")
   ' // Release the input source
   IF m_pOutputStream THEN m_pOutputStream->lpvtbl->Release(m_pOutputStream)
   ' // Release the Afx_IXmlReader interface
   IF m_pXmlWriterLite THEN m_pXmlWriterLite->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
   ' // Free the XmlLite library
   IF m_hLib THEN FreeLibrary(m_hLib)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
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
' Sets the output stream for the writer.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.SetOutput (BYVAL pOutputStream AS IUnknown PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      IF m_pOutputStream THEN m_pOutputStream->lpvtbl->Release(m_pOutputStream)
      m_Result = m_pXmlWriterLite->SetOutput(pOutputStream)
      IF m_Result = S_OK THEN m_pOutputStream = pOutputStream
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================


' ========================================================================================
' Sets the output source of the XML document to be writen.
' - pOutputStream: A pointer to the output stream for the writer to write data to.
'                  This can be either an ISequentialStream or an IStream interface.
' - nEncodingPage: Specifies the encoding of the output.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.SetOutput (BYVAL pOutputStream AS IUnknown PTR, BYVAL nEncodingPage AS UINT) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite = NULL THEN RETURN m_Result
   ' // Create the input interface
   DIM pfnCreateXmlWriterOutputWithEncodingCodePage AS FUNCTION (BYVAL pOutputStream AS IStream PTR, BYVAL pMalloc AS IMalloc PTR, _
       BYVAL nEncodingPage AS UINT, BYVAL ppOutput AS IUnknown PTR PTR) AS HRESULT
   pfnCreateXmlWriterOutputWithEncodingCodePage = CAST(ANY PTR, GetProcAddress(m_hLib, "CreateXmlWriterOutputWithEncodingCodePage"))
   IF pfnCreateXmlWriterOutputWithEncodingCodePage = NULL THEN RETURN m_Result
   ' // Release a previous output stream
   IF m_pOutputStream THEN m_pOutputStream->lpvtbl->Release(m_pOutputStream)
   m_pOutputStream = NULL
   m_Result = pfnCreateXmlWriterOutputWithEncodingCodePage(cast(IStream PTR, pOutputStream), NULL, nEncodingPage, @m_pOutputStream)
   IF FAILED(m_Result) THEN RETURN m_Result
   m_Result = m_pXmlWriterLite->SetOutput(m_pOutputStream)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the output source of the XML document to be writen.
' - pOutputStream: A pointer to the output stream for the writer to write data to.
'                  This can be either an ISequentialStream or an IStream interface.
' - nEncodingPage: Specifies the encoding of the output.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.SetOutput (BYVAL pOutputStream AS IUnknown PTR, BYVAL pwszEncodingName AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite = NULL THEN RETURN m_Result
   ' // Create the input interface
   DIM pfnCreateXmlWriterOutputWithEncodingName AS FUNCTION (BYVAL pOutputStream AS IStream PTR, BYVAL pMalloc AS IMalloc PTR, _
       BYVAL pwszEncodingName AS LPCWSTR, BYVAL ppOutput AS IUnknown PTR PTR) AS HRESULT
   pfnCreateXmlWriterOutputWithEncodingName = CAST(ANY PTR, GetProcAddress(m_hLib, "CreateXmlWriterOutputWithEncodingName"))
   IF pfnCreateXmlWriterOutputWithEncodingName = NULL THEN RETURN m_Result
   ' // Release a previous output stream
   IF m_pOutputStream THEN m_pOutputStream->lpvtbl->Release(m_pOutputStream)
   m_pOutputStream = NULL
   m_Result = pfnCreateXmlWriterOutputWithEncodingName(cast(IStream PTR, pOutputStream), NULL, pwszEncodingName, @m_pOutputStream)
   IF FAILED(m_Result) THEN RETURN m_Result
   m_Result = m_pXmlWriterLite->SetOutput(m_pOutputStream)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the specified property. A program can get properties at any time.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.GetProperty (BYVAL nProperty AS UINT) AS LONG_PTR
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      DIM pValue AS LONG_PTR
      m_Result = m_pXmlWriterLite->GetProperty(nProperty, @pValue)
      RETURN pValue
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the specified property.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->SetProperty(nProperty, pValue)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies attributes from the specified source IXmlReader to the IXmlWriter.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteAttributes (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteAttributes(pReader, fWriteDefaultAttributes)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes an attribute.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteAttributeString (BYVAL pwszQName AS WSTRING PTR, _
BYVAL pwszValue AS Const WSTRING PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteAttributeString(pwszQName, LEN(*pwszQName), pwszValue, LEN(*pwszValue))
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out a CDATA section that contains the specified text.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteCData (BYVAL pwszText AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteCData(pwszText)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a character entity for the provided Unicode character value. This method writes
' the character entity in hexadecimal format.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteCharEntity (BYVAL wch AS WCHAR PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteCharEntity(wch)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified text content, escaping markup.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteChars (BYVAL pwch AS WSTRING PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteChars(pwch, LEN(*pwch))
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out a comment that contains the specified text.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteComment (BYVAL pwszComment AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteComment(pwszComment)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out a comment that contains the specified text.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteDocType (BYVAL pwszName AS LPCWSTR, BYVAL pwszPublicId AS LPCWSTR, _
BYVAL pwszSystemId AS LPCWSTR, BYVAL pwszSubset AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteDocType(pwszName, pwszPublicId, pwszSystemId, pwszSubset)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out an element with the specified prefix, name, namespace, and value.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteElementString (BYVAL pwszQName AS WSTRING PTR, BYVAL pwszValue AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteElementString(pwszQName, LEN(*pwszQName), pwszValue)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes any open elements or attributes, then closes the current document. To reinitialize
' the writer so that a new document can be written, you must call SetOutput.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteEndDocument () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteEndDocument
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes one element. If the element contains no content, this method writes a short
' end tag ("/>"). Otherwise, this method writes the full end tag.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteEndElement (BYVAL pwszQName AS WSTRING PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteEndElement(pwszQName, LEN(*pwszQName))
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes one element. If the element contains no content, this method writes a short
' end tag ("/>"). Otherwise, this method writes the full end tag.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteEntityRef (BYVAL pwszName AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteEntityRef(pwszName)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes one element and pops the corresponding namespace scope. This method always writes
' the full end tag.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteFullEndElement (BYVAL pwszQName AS CONST WSTRING PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteFullEndElement(pwszQName, LEN(*pwszQName))
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified name, ensuring that the name is valid according to the XML specification.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteName (BYVAL pwszName AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteName(pwszName)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified name, ensuring that the name is valid according to the XML specification.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteNmToken (BYVAL pwszNmToken AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteNmToken(pwszNmToken)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified name, ensuring that the name is valid according to the XML specification.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteNode (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteNode(pReader, fWriteDefaultAttributes)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified name, ensuring that the name is valid according to the XML specification.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteNodeShallow (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteNodeShallow(pReader, fWriteDefaultAttributes)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes a processing instruction.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteProcessingInstruction (BYVAL pwszName AS LPCWSTR, BYVAL pwszText AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteProcessingInstruction(pwszName, pwszText)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Enables the caller to write out raw markup manually. You can avoid creating entities for
' special characters by using this method.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteRaw (BYVAL pwszData AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteRaw(pwszData)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out raw markup manually. Using this method allows an application to avoid creating
' entities for special characters.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteRawChars (BYVAL pwch AS CONST WSTRING PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteRawChars(pwch, LEN(*pwch))
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the XML declaration with the version "1.0". The encoding attribute is determined
' by the implementation of IXmlWriterOutput. By default, the encoding is UTF-8.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteStartDocument (BYVAL standalone AS XmlStandalone) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteStartDocument(standalone)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified start tag and associates it with the specified namespace.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteStartElement (BYVAL pwszQName AS CONST WSTRING PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteStartElement(pwszQName, LEN(*pwszQName))
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified text content, escaping any markup in the content.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteString (BYVAL pwszText AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteString(pwszText)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Forces the generation of a surrogate character entity for the specified string value.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteSurrogateCharEntity (BYVAL wchLow AS WCHAR, BYVAL wchHigh AS WCHAR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteSurrogateCharEntity(wchLow, wchHigh)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified white space.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.WriteWhitespace (BYVAL pwszWhitespace AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->WriteWhitespace(pwszWhitespace)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Flushes whatever is in the buffer to the underlying stream and then flushes the underlying stream.
' ========================================================================================
PRIVATE FUNCTION CXmlWriterLite.Flush () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriterLite THEN
      m_Result = m_pXmlWriterLite->Flush
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

END NAMESPACE
