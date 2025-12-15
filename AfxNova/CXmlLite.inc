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


END NAMESPACE
