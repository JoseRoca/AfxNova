' ########################################################################################
' Microsoft Windows
' File: AfxXmlLite.bi
' Contents: Lightweight XML parser
' Copyright (c) 2025 José Roca
' Portions Copyright (c) Microsoft Corporation. All Rights Reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/unknwn.bi"
#include once "win/wtypes.bi"
#include once "win/objidl.bi"
#include once "win/oaidl.bi"
#include once "AfxNova/AfxCOM.inc"
USING AfxNova

' // These definitions don't work with this library (maybe a misalignment problem),
' // so I'm using AfxGuid, that calls the API function ClsidFromString.
'DIM SHARED AFX_IID_IXmlReader AS GUID = (&h7279FC81, &h709D, &h4095, {&hB6, &h3D, &h69, &hFE, &h48, &h0D, &h90, &h30})
'DIM SHARED AFX_IID_IXmlResolver AS GUID = (&h7279FC82, &h709D, &h4095, {&hB6, &h3D, &h69, &hFE, &h48, &h0D, &h90, &h30})
'DIM SHARED AFX_IID_IXmlWriter AS GUID = (&h7279FC88, &h709D, &h4095, {&hB6, &h3D, &h69, &hFE, &h48, &h0D, &h90, &h30})
'DIM SHARED AFX_IID_IXmlWriterLite AS GUID = (&h862494C6, &h1310, &h4AAD, {&h83, &hCD, &h2D, &hBE, &hEB, &hF6, &h70, &hD3})

#define AFX_IID_IXmlReader AfxGuid("{7279FC81-709D-4095-B63D-69FE4B0D9030}")
#define AFX_IID_IXmlResolver AfxGuid("{7279FC82-709D-4095-B63D-69FE4B0D9030}")
#define AFX_IID_IXmlWriter AfxGuid("{7279FC88-709D-4095-B63D-69FE4B0D9030}")
#define AFX_IID_IXmlWriterLite AfxGuid("{862494C6-1310-4AAD-B3CD-2DBEEBF670D3}")

#define Afx_IXmlReaderInput Afx_IUnknown
#define Afx_IXmlWriterOutput Afx_IUnknown

enum XmlNodeType
   XmlNodeType_None= 0
   XmlNodeType_Element = 1
   XmlNodeType_Attribute = 2
   XmlNodeType_Text = 3
   XmlNodeType_CDATA = 4
   XmlNodeType_ProcessingInstruction= 7
   XmlNodeType_Comment = 8
   XmlNodeType_DocumentType = 10
   XmlNodeType_Whitespace = 13
   XmlNodeType_EndElement = 15
   XmlNodeType_XmlDeclaration = 17
   _XmlNodeType_Last = 17
end enum

enum XmlConformanceLevel
   XmlConformanceLevel_Auto = 0
   XmlConformanceLevel_Fragment = 1
   XmlConformanceLevel_Document = 2
   _XmlConformanceLevel_Last = 2
end enum

enum DtdProcessing
   DtdProcessing_Prohibit = 0
   DtdProcessing_Parse = DtdProcessing_Prohibit + 1
   _DtdProcessing_Last = DtdProcessing_Parse
end enum

enum XmlReadState
   XmlReadState_Initial = 0
   XmlReadState_Interactive = 1
   XmlReadState_Error = 2
   XmlReadState_EndOfFile = 3
   XmlReadState_Closed = 4
end enum

enum XmlReaderProperty
   XmlReaderProperty_MultiLanguage = 0
   XmlReaderProperty_ConformanceLevel = XmlReaderProperty_MultiLanguage + 1
   XmlReaderProperty_RandomAccess = XmlReaderProperty_ConformanceLevel + 1
   XmlReaderProperty_XmlResolver = XmlReaderProperty_RandomAccess + 1
   XmlReaderProperty_DtdProcessing = XmlReaderProperty_XmlResolver + 1
   XmlReaderProperty_ReadState = XmlReaderProperty_DtdProcessing + 1
   XmlReaderProperty_MaxElementDepth = XmlReaderProperty_ReadState + 1
   XmlReaderProperty_MaxEntityExpansion = XmlReaderProperty_MaxElementDepth + 1
   _XmlReaderProperty_Last = XmlReaderProperty_MaxEntityExpansion
end enum

enum XmlError
   MX_E_MX = &hc00cee00
   MX_E_INPUTEND = MX_E_MX + 1
   MX_E_ENCODING = MX_E_INPUTEND + 1
   MX_E_ENCODINGSWITCH = MX_E_ENCODING + 1
   MX_E_ENCODINGSIGNATURE = MX_E_ENCODINGSWITCH + 1
   WC_E_WC = &hc00cee20
   WC_E_WHITESPACE = WC_E_WC + 1
   WC_E_SEMICOLON = WC_E_WHITESPACE + 1
   WC_E_GREATERTHAN = WC_E_SEMICOLON + 1
   WC_E_QUOTE = WC_E_GREATERTHAN + 1
   WC_E_EQUAL = WC_E_QUOTE + 1
   WC_E_LESSTHAN = WC_E_EQUAL + 1
   WC_E_HEXDIGIT = WC_E_LESSTHAN + 1
   WC_E_DIGIT = WC_E_HEXDIGIT + 1
   WC_E_LEFTBRACKET = WC_E_DIGIT + 1
   WC_E_LEFTPAREN = WC_E_LEFTBRACKET + 1
   WC_E_XMLCHARACTER = WC_E_LEFTPAREN + 1
   WC_E_NAMECHARACTER = WC_E_XMLCHARACTER + 1
   WC_E_SYNTAX = WC_E_NAMECHARACTER + 1
   WC_E_CDSECT = WC_E_SYNTAX + 1
   WC_E_COMMENT = WC_E_CDSECT + 1
   WC_E_CONDSECT = WC_E_COMMENT + 1
   WC_E_DECLATTLIST = WC_E_CONDSECT + 1
   WC_E_DECLDOCTYPE = WC_E_DECLATTLIST + 1
   WC_E_DECLELEMENT = WC_E_DECLDOCTYPE + 1
   WC_E_DECLENTITY = WC_E_DECLELEMENT + 1
   WC_E_DECLNOTATION = WC_E_DECLENTITY + 1
   WC_E_NDATA = WC_E_DECLNOTATION + 1
   WC_E_PUBLIC = WC_E_NDATA + 1
   WC_E_SYSTEM = WC_E_PUBLIC + 1
   WC_E_NAME = WC_E_SYSTEM + 1
   WC_E_ROOTELEMENT = WC_E_NAME + 1
   WC_E_ELEMENTMATCH = WC_E_ROOTELEMENT + 1
   WC_E_UNIQUEATTRIBUTE = WC_E_ELEMENTMATCH + 1
   WC_E_TEXTXMLDECL = WC_E_UNIQUEATTRIBUTE + 1
   WC_E_LEADINGXML = WC_E_TEXTXMLDECL + 1
   WC_E_TEXTDECL = WC_E_LEADINGXML + 1
   WC_E_XMLDECL = WC_E_TEXTDECL + 1
   WC_E_ENCNAME = WC_E_XMLDECL + 1
   WC_E_PUBLICID = WC_E_ENCNAME + 1
   WC_E_PESINTERNALSUBSET = WC_E_PUBLICID + 1
   WC_E_PESBETWEENDECLS = WC_E_PESINTERNALSUBSET + 1
   WC_E_NORECURSION = WC_E_PESBETWEENDECLS + 1
   WC_E_ENTITYCONTENT = WC_E_NORECURSION + 1
   WC_E_UNDECLAREDENTITY = WC_E_ENTITYCONTENT + 1
   WC_E_PARSEDENTITY = WC_E_UNDECLAREDENTITY + 1
   WC_E_NOEXTERNALENTITYREF = WC_E_PARSEDENTITY + 1
   WC_E_PI = WC_E_NOEXTERNALENTITYREF + 1
   WC_E_SYSTEMID = WC_E_PI + 1
   WC_E_QUESTIONMARK = WC_E_SYSTEMID + 1
   WC_E_CDSECTEND = WC_E_QUESTIONMARK + 1
   WC_E_MOREDATA = WC_E_CDSECTEND + 1
   WC_E_DTDPROHIBITED = WC_E_MOREDATA + 1
   WC_E_INVALIDXMLSPACE = WC_E_DTDPROHIBITED + 1
   NC_E_NC = &hc00cee60,
   NC_E_QNAMECHARACTER = NC_E_NC + 1
   NC_E_QNAMECOLON = NC_E_QNAMECHARACTER + 1
   NC_E_NAMECOLON = NC_E_QNAMECOLON + 1
   NC_E_DECLAREDPREFIX = NC_E_NAMECOLON + 1
   NC_E_UNDECLAREDPREFIX = NC_E_DECLAREDPREFIX + 1
   NC_E_EMPTYURI = NC_E_UNDECLAREDPREFIX + 1
   NC_E_XMLPREFIXRESERVED = NC_E_EMPTYURI + 1
   NC_E_XMLNSPREFIXRESERVED = NC_E_XMLPREFIXRESERVED + 1
   NC_E_XMLURIRESERVED = NC_E_XMLNSPREFIXRESERVED + 1
   NC_E_XMLNSURIRESERVED = NC_E_XMLURIRESERVED + 1
   SC_E_SC = &hc00cee80
   SC_E_MAXELEMENTDEPTH = SC_E_SC + 1
   SC_E_MAXENTITYEXPANSION = SC_E_MAXELEMENTDEPTH + 1
   WR_E_WR = &hc00cef00
   WR_E_NONWHITESPACE = WR_E_WR + 1
   WR_E_NSPREFIXDECLARED = WR_E_NONWHITESPACE + 1
   WR_E_NSPREFIXWITHEMPTYNSURI = WR_E_NSPREFIXDECLARED + 1
   WR_E_DUPLICATEATTRIBUTE = WR_E_NSPREFIXWITHEMPTYNSURI + 1
   WR_E_XMLNSPREFIXDECLARATION = WR_E_DUPLICATEATTRIBUTE + 1
   WR_E_XMLPREFIXDECLARATION = WR_E_XMLNSPREFIXDECLARATION + 1
   WR_E_XMLURIDECLARATION = WR_E_XMLPREFIXDECLARATION + 1
   WR_E_XMLNSURIDECLARATION = WR_E_XMLURIDECLARATION + 1
   WR_E_NAMESPACEUNDECLARED = WR_E_XMLNSURIDECLARATION + 1
   WR_E_INVALIDXMLSPACE = WR_E_NAMESPACEUNDECLARED + 1
   WR_E_INVALIDACTION = WR_E_INVALIDXMLSPACE + 1
   WR_E_INVALIDSURROGATEPAIR = WR_E_INVALIDACTION + 1
   XML_E_INVALID_DECIMAL = &hc00ce01d,
   XML_E_INVALID_HEXIDECIMAL = &hc00ce01e,
   XML_E_INVALID_UNICODE = &hc00ce01f,
   XML_E_INVALIDENCODING = &hc00ce06e
end enum

enum XmlStandalone
   XmlStandalone_Omit = 0
   XmlStandalone_Yes = 1
   XmlStandalone_No = 2
   _XmlStandalone_Last = 2
end enum

enum XmlWriterProperty
   XmlWriterProperty_MultiLanguage = 0
   XmlWriterProperty_Indent = XmlWriterProperty_MultiLanguage + 1
   XmlWriterProperty_ByteOrderMark = XmlWriterProperty_Indent + 1
   XmlWriterProperty_OmitXmlDeclaration = XmlWriterProperty_ByteOrderMark + 1
   XmlWriterProperty_ConformanceLevel = XmlWriterProperty_OmitXmlDeclaration + 1
   XmlWriterProperty_CompactEmptyElement = XmlWriterProperty_ConformanceLevel + 1
   _XmlWriterProperty_Last = XmlWriterProperty_CompactEmptyElement
end enum

' // Interfaces - Forward references
TYPE Afx_IXmlReader AS Afx_IXmlReader_
TYPE Afx_IXmlResolver AS Afx_IXmlResolver_
TYPE Afx_IXmlWriter AS Afx_IXmlWriter_
TYPE Afx_IXmlWriterLite AS Afx_IXmlWriterLite_

' ########################################################################################
' Interface name: IXmlReader
' IID: {7279FC81-709D-4095-B63D-69FE4B0D9030}
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_IXmlReader_INTERFACE_DEFINED__
#define __Afx_IXmlReader_INTERFACE_DEFINED__

TYPE Afx_IXmlReader_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION SetInput (BYVAL pInput AS IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetProperty (BYVAL nProperty AS UINT, BYVAL ppValue AS LONG_PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Read (BYVAL pNodeType AS XmlNodeType PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetNodeType (BYVAL pNodeType AS XmlNodeType PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveToFirstAttribute () AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveToNextAttribute () AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveToAttributeByName (BYVAL pwszLocalName AS LPCWSTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION MoveToElement () AS HRESULT
   DECLARE ABSTRACT FUNCTION GetQualifiedName (BYVAL ppwszQualifiedName AS LPCWSTR PTR, BYVAL pcwchQualifiedName AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetNamespaceUri (BYVAL ppwszNamespaceUri AS LPCWSTR PTR, BYVAL pcwchNamespaceUri AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetLocalName (BYVAL ppwszLocalName AS LPCWSTR PTR, BYVAL pcwchLocalName AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetPrefix (BYVAL ppwszPrefix AS LPCWSTR PTR, BYVAL pcwchPrefix AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetValue (BYVAL ppwszValue AS LPCWSTR PTR, BYVAL pcwchValue AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION ReadValueChunk (BYVAL pwchBuffer AS WSTRING PTR, BYVAL cwchChunkSize AS UINT, BYVAL pcwchRead AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetBaseUri (BYVAL ppwszBaseUri AS LPCWSTR PTR, BYVAL pcwchBaseUri AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsDefault () AS BOOL
   DECLARE ABSTRACT FUNCTION IsEmptyElement () AS BOOL
   DECLARE ABSTRACT FUNCTION GetLineNumber (BYVAL pnLineNumber AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetLinePosition (BYVAL pnLinePosition AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetAttributeCount (BYVAL pnAttributeCount AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDepth (BYVAL pnDepth AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsEOF () AS BOOL

END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: IXmlResolver
' IID: {7279FC82-709D-4095-B63D-69FE4B0D9030)
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_IXmlResolver_INTERFACE_DEFINED__
#define __Afx_IXmlResolver_INTERFACE_DEFINED__

TYPE Afx_IXmlResolver_ EXTENDS Afx_IUnknown

   DECLARE ABSTRACT FUNCTION ResolveUri (BYVAL pwszBaseUri AS LPCWSTR, BYVAL pwszPublicIdentifier AS LPCWSTR, _
      BYVAL pwszSystemIdentifier AS LPCWSTR, BYVAL ppResolvedInput AS IUnknown PTR) AS HRESULT

END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: IXmlWriter
' IID: (7279FC88-709D-4095-B63D-69FE4B0D9030)
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_IXmlWriter_INTERFACE_DEFINED__
#define __Afx_IXmlWriter_INTERFACE_DEFINED__

TYPE Afx_IXmlWriter_ EXTENDS Afx_IUnknown

   DECLARE ABSTRACT FUNCTION SetOutput (BYVAL pOutput AS IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetProperty (BYVAL nProperty AS UINT, BYVAL ppValue AS LONG_PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteAttributes (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteAttributeString (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, _
      BYVAL pwszNamespaceUri AS LPCWSTR, BYVAL pwszValue AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteCData (BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteCharEntity (BYVAL wch AS WCHAR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteChars (BYVAL pwch AS WSTRING PTR, BYVAL cwch AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteComment (BYVAL pwszComment AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteDocType (BYVAL pwszName AS LPCWSTR, BYVAL pwszPublicId AS LPCWSTR, _
      BYVAL pwszSystemId AS LPCWSTR, BYVAL pwszSubset AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteElementString (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, _
      BYVAL pwszNamespaceUri AS LPCWSTR, BYVAL pwszValue AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteEndDocument () AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteEndElement () AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteEntityRef (BYVAL pwszName AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteFullEndElement () AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteName (BYVAL pwszName AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteNmToken (BYVAL pwszNmToken AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteNode (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteNodeShallow (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteProcessingInstruction (BYVAL pwszName AS LPCWSTR, BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteQualifiedName (BYVAL pwszLocalName AS LPCWSTR PTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteRaw (BYVAL pwszData AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteRawChars (BYVAL pwch AS CONST WSTRING PTR, BYVAL cwch AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteStartDocument (BYVAL standalone AS XmlStandalone) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteStartElement (BYVAL pwszPrefix AS LPCWSTR, BYVAL pwszLocalName AS LPCWSTR, BYVAL pwszNamespaceUri AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteString (BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteSurrogateCharEntity (BYVAL wchLow AS WCHAR, BYVAL wchHigh AS WCHAR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteWhitespace (BYVAL pwszWhitespace AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Flush () AS HRESULT

END TYPE

#endif

' ########################################################################################


' ########################################################################################
' Interface name: IXmlWriterLite
' IID: {862494C6-1310-4AAD-B3CD-2DBEEBF670D3)
' Inherited interface = IUnknown
' ########################################################################################

#ifndef __Afx_IXmlWriterLite_INTERFACE_DEFINED__
#define __Afx_IXmlWriterLite_INTERFACE_DEFINED__

TYPE Afx_IXmlWriterLite_ EXTENDS Afx_IUnknown

   DECLARE ABSTRACT FUNCTION SetOutput (BYVAL SetOutput AS IUnknown PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetProperty (BYVAL nProperty AS UINT, BYVAL ppValue AS LONG_PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION SetProperty (BYVAL nProperty AS UINT, BYVAL pValue AS LONG_PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteAttributes (BYVAL pReader AS AFx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteAttributeString (BYVAL pwszQName AS WSTRING PTR, BYVAL cwszQName AS UINT, _
      BYVAL pwszValue AS Const WSTRING PTR, BYVAL cwszValue AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteCData (BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteCharEntity (BYVAL wch AS WSTRING PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteChars (BYVAL pwch AS CONST WSTRING PTR, BYVAL cwch AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteComment (BYVAL pwszComment AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteDocType (BYVAL pwszName AS LPCWSTR, BYVAL pwszPublicId AS LPCWSTR, _
      BYVAL pwszSystemId AS LPCWSTR, BYVAL pwszSubset AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteElementString (BYVAL pwszQName AS WSTRING PTR, BYVAL cwszQName AS UINT, BYVAL pwszValue AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteEndDocument () AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteEndElement (BYVAL pwszQName AS WSTRING PTR, BYVAL cwszQName AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteEntityRef (BYVAL pwszName AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteFullEndElement (BYVAL pwszQName AS CONST WSTRING PTR, BYVAL cwszQName AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteName (BYVAL pwszName AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteNmToken (BYVAL pwszNmToken AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteNode (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteNodeShallow (BYVAL pReader AS Afx_IXmlReader PTR, BYVAL fWriteDefaultAttributes AS BOOL) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteProcessingInstruction (BYVAL pwszName AS LPCWSTR, BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteRaw (BYVAL pwszData AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteRawChars (BYVAL pwch AS CONST WSTRING PTR, BYVAL cwch AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteStartDocument (BYVAL standalone AS XmlStandalone) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteStartElement (BYVAL pwszQName AS CONST WSTRING PTR, BYVAL cwszQName AS UINT) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteString (BYVAL pwszText AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteSurrogateCharEntity (BYVAL wchLow AS WCHAR, BYVAL wchHigh AS WCHAR) AS HRESULT
   DECLARE ABSTRACT FUNCTION WriteWhitespace (BYVAL pwszWhitespace AS LPCWSTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Flush () AS HRESULT

END TYPE

#endif

' ########################################################################################


'extern "Windows"
'' // XmlReader Constructor
'DECLARE FUNCTION CreateXmlReader LIB "xmllite.lib" ALIAS "CreateXmlReader" (BYVAL riid AS REFIID, BYVAL ppvObject AS void PTR PTR, BYVAL pMalloc AS IMalloc PTR) AS HRESULT
'' // XmlReaderInput Constructors
'DECLARE FUNCTION CreateXmlReaderInputWithEncodingCodePage LIB "xmllite.lib" ALIAS "CreateXmlReaderInputWithEncodingCodePage" (BYVAL pInputStream AS Afx_IUnknown PTR, BYVAL pMalloc AS IMalloc PTR, _
'   BYVAL nEncodingCodePage AS UINT, BYVAL fEncodingHint AS BOOL, BYVAL pwszBaseUri AS LPCWSTR, BYVAL ppInput AS Afx_IXmlReaderInput PTR PTR) AS HRESULT
'DECLARE FUNCTION CreateXmlReaderInputWithEncodingName LIB "xmllite.lib" ALIAS "CreateXmlReaderInputWithEncodingName" (BYVAL pInputStream AS Afx_IUnknown PTR, BYVAL pMalloc AS IMalloc PTR, _
'   BYVAL pwszEncodingName AS LPCWSTR, BYVAL fEncodingHint AS BOOL, BYVAL pwszBaseUri AS LPCWSTR, BYVAL ppInput AS Afx_IXmlReaderInput PTR PTR) AS HRESULT
'' // XmlWriter Constructor
'DECLARE FUNCTION CreateXmlWriter LIB "xmllite.lib" ALIAS "CreateXmlWriter" (BYVAL riid AS REFIID, BYVAL ppvObject AS void PTR PTR, BYVAL pMalloc AS IMalloc PTR) AS HRESULT
'' // XmlWriterOutput Constructors
'DECLARE FUNCTION CreateXmlWriterOutputWithEncodingCodePage LIB "xmllite.lib" ALIAS "CreateXmlWriterOutputWithEncodingCodePage" (BYVAL pOutputStream AS Afx_IUnknown PTR, BYVAL pMalloc AS IMalloc PTR, _
'   BYVAL nEncodingCodePage AS UINT, BYVAL ppOutput AS Afx_IXmlWriterOutput PTR PTR) AS HRESULT
'DECLARE FUNCTION CreateXmlWriterOutputWithEncodingName LIB "xmllite.lib" ALIAS "CreateXmlWriterOutputWithEncodingName" (BYVAL pOutputStream AS Afx_IUnknown PTR, BYVAL pMalloc AS IMalloc PTR, _
'   BYVAL nEncodingCodePage AS UINT, BYVAL ppOutput AS Afx_IXmlWriterOutput PTR PTR) AS HRESULT
'end extern
