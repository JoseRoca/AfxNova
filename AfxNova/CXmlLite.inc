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
' CXmlLiteReader class
' ########################################################################################
TYPE CXmlLiteReader

   m_Result AS HRESULT
   m_hLib AS HMODULE
   m_bUninitCOM AS BOOLEAN
   m_pXmlReader AS Afx_IXmlReader PTR
   m_pInput AS IUnknown PTR

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   ' // IXmlReader methods and properties
   DECLARE FUNCTION SetInput (BYVAL pInput AS IUnknown PTR) AS HRESULT
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
PRIVATE CONSTRUCTOR CXmlLiteReader
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
PRIVATE DESTRUCTOR CXmlLiteReader
   CXMLLITE_DP("")
   ' // Release the input source
   IF m_pInput THEN m_pInput->lpvtbl->Release(m_pInput)
   ' // Release the Afx_IXmlReader interface
   IF m_pXmlReader THEN m_pXmlReader->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
   ' // Free the XmlLite library
   IF m_hLib THEN FreeLibrary(m_hLib)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last status code.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteReader.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteReader.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteReader.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
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
PRIVATE FUNCTION CXmlLiteReader.SetInput (BYVAL pInput AS IUnknown PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlReader THEN
      IF m_pInput THEN m_pInput->lpvtbl->Release(m_pInput)
      m_Result = m_pXmlReader->SetInput(pInput)
      IF m_Result = S_OK THEN m_pInput = pInput
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the specified property. A program can get properties at any time.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteReader.GetProperty (BYVAL nProperty AS UINT) AS LONG_PTR
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
PRIVATE FUNCTION CXmlLiteReader.SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteReader.Read (BYVAL nodeType AS XmlNodeType PTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteReader.GetNodeType () AS XmlNodeType 
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
PRIVATE FUNCTION CXmlLiteReader.MoveToFirstAttribute () AS HRESULT
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
PRIVATE FUNCTION CXmlLiteReader.MoveToNextAttribute () AS HRESULT
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
PRIVATE FUNCTION CXmlLiteReader.MoveToAttributeByName (BYVAL pwszLocalName AS LPCWSTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteReader.MoveToElement () AS HRESULT
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
PRIVATE FUNCTION CXmlLiteReader.GetQualifiedName () AS DWSTRING
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
PRIVATE FUNCTION CXmlLiteReader.GetNamespaceUri () AS DWSTRING
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
PRIVATE FUNCTION CXmlLiteReader.GetLocalName () AS DWSTRING
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
PRIVATE FUNCTION CXmlLiteReader.GetPrefix () AS DWSTRING
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
PRIVATE FUNCTION CXmlLiteReader.GetValue () AS DWSTRING
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
PRIVATE FUNCTION CXmlLiteReader.ReadValueChunk (BYVAL buffer AS WSTRING PTR, BYVAL chunkSize AS UINT, BYVAL pcwchRead AS UINT PTR) AS HRESULT
   m_Result = E_POINTER
   m_Result = ReadValueChunk(buffer, chunkSize, pcwchRead)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the base URI of the token, if applicable.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteReader.GetBaseUri () AS DWSTRING
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
PRIVATE FUNCTION CXmlLiteReader.IsDefault () AS BOOLEAN
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
PRIVATE FUNCTION CXmlLiteReader.IsEmptyElement () AS BOOLEAN
   m_Result = 0
   IF m_pXmlReader THEN
      RETURN m_pXmlReader->IsEmptyElement
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the line number where the reader is positioned in the document.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteReader.GetLineNumber () AS UINT
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
PRIVATE FUNCTION CXmlLiteReader.GetLinePosition () AS UINT
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
PRIVATE FUNCTION CXmlLiteReader.GetAttributeCount () AS UINT
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
PRIVATE FUNCTION CXmlLiteReader.GetDepth () AS UINT
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
PRIVATE FUNCTION CXmlLiteReader.IsEOF () AS BOOLEAN
   m_Result = 0
   IF m_pXmlReader THEN
      RETURN m_pXmlReader->IsEOF
   END IF
END FUNCTION
' ========================================================================================


' ########################################################################################
' CXmlLiteWriter class
' ########################################################################################
TYPE CXmlLiteWriter

   m_Result AS HRESULT
   m_hLib AS HMODULE
   m_bUninitCOM AS BOOLEAN
   m_pXmlWriter AS Afx_IXmlWriter PTR
   m_pOutput AS IUnknown PTR

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   ' // IXmlWriter methods and properties
   DECLARE FUNCTION SetOutput (BYVAL pOutput AS IUnknown PTR) AS HRESULT
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
PRIVATE CONSTRUCTOR CXmlLiteWriter
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
PRIVATE DESTRUCTOR CXmlLiteWriter
   CXMLLITE_DP("")
   ' // Release the input source
   IF m_pOutput THEN m_pOutput->lpvtbl->Release(m_pOutput)
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
PRIVATE FUNCTION CXmlLiteWriter.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last status code.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteWriter.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteWriter.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
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
PRIVATE FUNCTION CXmlLiteWriter.SetOutput (BYVAL pOutput AS IUnknown PTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      IF m_pOutput THEN m_pOutput->lpvtbl->Release(m_pOutput)
      m_Result = m_pXmlWriter->SetOutput(pOutput)
      IF m_Result = S_OK THEN m_pOutput = pOutput
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the specified property. A program can get properties at any time.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteWriter.GetProperty (BYVAL nProperty AS UINT) AS LONG_PTR
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
PRIVATE FUNCTION CXmlLiteWriter.SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteAttributes (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteAttributeString (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, _
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
PRIVATE FUNCTION CXmlLiteWriter.WriteCData (BYVAL pwszText AS LPCWSTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteCharEntity (BYVAL wch AS WCHAR PTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteChars (BYVAL pwch AS WSTRING PTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteComment (BYVAL pwszComment AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteComment(pwszComment)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out a comment that contains the specified text.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteWriter.WriteDocType (BYVAL pwszName AS LPCWSTR, BYVAL pwszPublicId AS LPCWSTR, _
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
PRIVATE FUNCTION CXmlLiteWriter.WriteElementString (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, _
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
PRIVATE FUNCTION CXmlLiteWriter.WriteEndDocument () AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteEndElement () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteEndElement
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Closes one element. If the element contains no content, this method writes a short
' end tag ("/>"). Otherwise, this method writes the full end tag.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteWriter.WriteEntityRef (BYVAL pwszName AS LPCWSTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteFullEndElement () AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteName (BYVAL pwszName AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteName(pwszName)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified name, ensuring that the name is valid according to the XML specification.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteWriter.WriteNmToken (BYVAL pwszNmToken AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteNmToken(pwszNmToken)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified name, ensuring that the name is valid according to the XML specification.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteWriter.WriteNode (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteNode(pReader, fWriteDefaultAttributes)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified name, ensuring that the name is valid according to the XML specification.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteWriter.WriteNodeShallow (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOLEAN) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteProcessingInstruction (BYVAL pwszName AS LPCWSTR, BYVAL pwszText AS LPCWSTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteQualifiedName (BYVAL pwszLocalName AS LPCWSTR PTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteRaw (BYVAL pwszData AS LPCWSTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteRawChars (BYVAL pwch AS CONST WSTRING PTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteStartDocument (BYVAL standalone AS XmlStandalone) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteStartElement (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteString (BYVAL pwszText AS LPCWSTR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteSurrogateCharEntity (BYVAL wchLow AS WCHAR, BYVAL wchHigh AS WCHAR) AS HRESULT
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
PRIVATE FUNCTION CXmlLiteWriter.WriteWhitespace (BYVAL pwszWhitespace AS LPCWSTR) AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->WriteWhitespace(pwszWhitespace)
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Writes out the specified white space.
' ========================================================================================
PRIVATE FUNCTION CXmlLiteWriter.Flush () AS HRESULT
   m_Result = E_POINTER
   IF m_pXmlWriter THEN
      m_Result = m_pXmlWriter->Flush
   END IF
   RETURN m_Result
END FUNCTION
' ========================================================================================

END NAMESPACE
