' ########################################################################################
' Platform: Microsoft Windows
' File: CAdoFields.inc
' Contents: ADO Fields
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
' CAdoField class
' ########################################################################################
TYPE CAdoField EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pField AS Afx_ADOField PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pField AS Afx_ADOField PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOField PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOField PTR PTR

   ' // Attaches a reference to an Afx_ADOField object to the class
   DECLARE SUB Attach (BYVAL pField AS Afx_ADOField PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Retrieves the name of a field
   DECLARE PROPERTY Name () AS DWSTRING
   ' // Sets or returns a Variant value that indicates the value of the object
   DECLARE PROPERTY Value () AS DVARIANT
   DECLARE PROPERTY Value (BYREF dvValue AS DVARIANT)
   ' // Returns a reference to the Properties collection.
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Indicates the operational type or data type of a Field object.
   DECLARE PROPERTY Type_ () AS DataTypeEnum
   DECLARE PROPERTY Type_ (BYVAL nType AS DataTypeEnum)
   ' // Retrieves the data capacity of a field
   DECLARE PROPERTY ActualSize () AS LONG
   ' // Indicates the data capacity of a field.
   DECLARE PROPERTY DefinedSize () AS LONG
   DECLARE PROPERTY DefinedSize (BYVAL lSize AS LONG)
   ' // Indicates one or more characteristics of a field.
   DECLARE PROPERTY Attributes () AS LONG
   DECLARE PROPERTY Attributes (BYVAL lAttr AS LONG)
   ' // Indicates the degree of precision for numeric values for numeric fields.
   DECLARE PROPERTY Precision () AS BYTE
   DECLARE PROPERTY Precision (BYVAL bPrecision AS BYTE)
   ' // Indicates the scale of numeric values in a field.
   DECLARE PROPERTY NumericScale () AS BYTE
   DECLARE PROPERTY NumericScale (BYVAL bScale AS BYTE)
   ' // * Indicates the value of a field that existed in the record before any changes were made.
   DECLARE PROPERTY OriginalValue () AS DVARIANT
   ' // Indicates a field's current value in the database.
   DECLARE PROPERTY UnderlyingValue () AS DVARIANT
   ' // Appends data to a large text or binary data field.
   DECLARE FUNCTION AppendChunk (BYREF dvData AS DVARIANT) AS HRESULT
   ' // Returns all, or a portion, of the contents of a large text or binary data field.
   DECLARE FUNCTION GetChunk (BYVAL nLen AS LONG) AS DVARIANT
   ' // Links the current Field object to a data-bound control.
   DECLARE PROPERTY DataFormat () AS Afx_IUnknown PTR
   DECLARE PROPERTY DataFormat (BYVAL piDF AS Afx_IUnknown PTR)
   ' // * Indicates the status of a field. Returns a FieldStatusEnum value. The default value is adFieldOK.
   DECLARE PROPERTY Status () AS LONG

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoField
   CADODB_DP("Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoField (BYVAL pField AS Afx_ADOField PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("")
   IF pField THEN
      m_pField = pField
      IF fAddRef THEN m_pField->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoField
   CADODB_DP(WSTR(m_pField))
   IF m_pField THEN m_pField->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoField.@ () AS Afx_ADOField PTR PTR
'   OPERATOR = @m_pField
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoField.vptr () AS Afx_ADOField PTR PTR
   CADODB_DP("")
   IF m_pField THEN m_pField->Release
   RETURN @m_pField
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cField AS CAdoField) AS Afx_ADOField PTR
   OPERATOR = cField.m_pField
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOField object to the class
' ===========================================================================================
PRIVATE SUB CAdoField.Attach (BYVAL pField AS Afx_ADOField PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("")
   IF m_pField THEN m_pField->Release
   m_pField = pField
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pField THEN m_pField->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Retrieves the name of a field
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Name () AS DWSTRING
   CADODB_DP("")
   IF m_pField = NULL THEN SetResult E_POINTER : RETURN ""
   DIM bstrName AS AFX_BSTR
   SetResult(m_pField->get_Name(@bstrName))
   DIM dwsName AS DWSTRING = *bstrName
   RETURN dwsName
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Value () AS DVARIANT
   CADODB_DP("")
   DIM vValue AS VARIANT
   IF m_pField THEN SetResult(m_pField->get_Value(@vValue))
   RETURN vValue
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Value (BYREF dvValue AS DVARIANT)
   CADODB_DP("")
   IF m_pField THEN SetResult(m_pField->put_Value(dvValue))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the Properties collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Properties () AS Afx_ADOProperties PTR
   CADODB_DP("")
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pField THEN SetResult(m_pField->get_Properties(@pProperties))
   RETURN pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Type_ () AS DataTypeEnum
   CADODB_DP("")
   DIM nType AS DataTypeEnum
   IF m_pField THEN SetResult(m_pField->get_Type(@nType))
   RETURN nType
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Type_ (BYVAL nType AS DataTypeEnum)
   CADODB_DP("")
   IF m_pField THEN SetResult(m_pField->put_Type(nType))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the actual length of a field's value.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.ActualSize () AS LONG
   CADODB_DP("")
   DIM lSize AS LONG
   IF m_pField THEN SetResult(m_pField->get_ActualSize(@lSize))
   RETURN lSize
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.DefinedSize () AS LONG
   CADODB_DP("")
   DIM lSize AS LONG
   IF m_pField THEN SetResult(m_pField->get_DefinedSize(@lSize))
   RETURN lSize
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.DefinedSize (BYVAL lSize AS LONG)
   CADODB_DP("")
   IF m_pField THEN SetResult(m_pField->put_DefinedSize(lSize))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Attributes () AS LONG
   CADODB_DP("")
   DIM lAttr AS LONG
   IF m_pField THEN SetResult(m_pField->get_Attributes(@lAttr))
   RETURN lAttr
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the operational type or data type of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Attributes (BYVAL lAttr AS LONG)
   CADODB_DP("")
   IF m_pField THEN SetResult(m_pField->put_Attributes(lAttr))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the degree of precision for numeric values for numeric fields.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Precision () AS BYTE
   CADODB_DP("")
   DIM bPrecision AS BYTE
   IF m_pField THEN SetResult(m_pField->get_Precision(@bPrecision))
   RETURN bPrecision
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the degree of precision for numeric values for numeric fields.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Precision (BYVAL bPrecision AS BYTE)
   CADODB_DP("")
   IF m_pField THEN SetResult(m_pField->put_Precision(bPrecision))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.NumericScale () AS BYTE
   CADODB_DP("")
   DIM bScale AS BYTE
   IF m_pField THEN SetResult(m_pField->get_NumericScale(@bScale))
   RETURN bScale
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data capacity of a field.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.NumericScale (BYVAL bScale AS BYTE)
   CADODB_DP("")
   IF m_pField THEN SetResult(m_pField->put_NumericScale(bScale))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the value of a field that existed in the record before any changes were made.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.OriginalValue () AS DVARIANT
   CADODB_DP("")
   DIM vValue AS VARIANT
   IF m_pField THEN SetResult(m_pField->get_OriginalValue(@vValue))
   RETURN vValue
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates a field's current value in the database.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.UnderlyingValue () AS DVARIANT
   CADODB_DP("")
   DIM vValue AS VARIANT
   IF m_pField THEN SetResult(m_pField->get_UnderlyingValue(@vValue))
   RETURN vValue
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Appends data to a large text or binary data Field.
' ===========================================================================================
PRIVATE FUNCTION CAdoField.AppendChunk (BYREF dvData AS DVARIANT) AS HRESULT
   CADODB_DP("")
   IF m_pField THEN SetResult(m_pField->AppendChunk(dvData))
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns all, or a portion, of the contents of a large text or binary data field.
' ===========================================================================================
PRIVATE FUNCTION CAdoField.GetChunk (BYVAL nLen AS LONG) AS DVARIANT
   CADODB_DP("")
   DIM dvData AS DVARIANT
   IF m_pField THEN SetResult(m_pField->GetChunk(nLen, dvData))
   RETURN dvData
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Links the current Field object to a data-bound control
' ===========================================================================================
PRIVATE PROPERTY CAdoField.DataFormat () AS Afx_IUnknown PTR
   CADODB_DP("")
   DIM piDF AS Afx_IUnknown PTR
   IF m_pField THEN SetResult(m_pField->get_DataFormat(@piDF))
   RETURN piDF
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoField.DataFormat (BYVAL piDF AS Afx_IUnknown PTR)
   CADODB_DP("")
   IF m_pField THEN SetResult(m_pField->putref_DataFormat(piDF))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the status of a field. Returns a FieldStatusEnum value. The default value is adFieldOK.
' ===========================================================================================
PRIVATE PROPERTY CAdoField.Status () AS LONG
   CADODB_DP("")
   DIM lStatus AS LONG
   IF m_pField THEN SetResult(m_pField->get_Status(@lStatus))
   RETURN lStatus
END PROPERTY
' ===========================================================================================


' ########################################################################################
' CAdoFields class
' ########################################################################################
TYPE CAdoFields EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pFields AS Afx_ADOFields PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOFields PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOFields PTR PTR

   ' // Attaches a reference to an Afx_ADOFields object to the class
   DECLARE SUB Attach (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Retrieves the number of fields
   DECLARE PROPERTY Count () AS LONG
   ' // Indicates a specific member of the Fields collection, by name or ordinal number.
   DECLARE PROPERTY Item (BYREF dvIndex AS DVARIANT) AS Afx_ADOField PTR
   ' // Appends an object to a collection.
   DECLARE FUNCTION Append (BYREF wszName AS Wstring, BYVAL nType AS DataTypeEnum, BYVAL DefinedSize AS ADO_LONGPTR = 0, _
           BYVAL Attrib AS FieldAttributeEnum = 0, BYREF dvFieldValue AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   ' // Deletes an object from the Fields collection.
   DECLARE FUNCTION Delete_ (BYREF dvIndex AS DVARIANT) AS HRESULT
   ' // Returns a reference to the collection's enumerator.
   DECLARE FUNCTION NewEnum () AS AFX_IUnknown PTR
   ' // Refreshes the contents of the Properties collection.
   DECLARE FUNCTION Refresh () AS HRESULT
   ' // Resynchronizes the contents of the Fields collection.
   DECLARE FUNCTION Resync (BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   ' // Saves any changes you make to the current Fields collection of a Record object.
   DECLARE FUNCTION Update () AS HRESULT
   ' // Cancels any changes made to the Fields collection of a Record object before calling the Update method.
   DECLARE FUNCTION CancelUpdate () AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoFields
   CADODB_DP("Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoFields (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("Afx_ADOFields")
   IF pFields THEN
      m_pFields = pFields
      IF fAddRef THEN m_pFields->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoFields
   CADODB_DP(WSTR(m_pFields))
   IF m_pFields THEN m_pFields->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoFields.@ () AS Afx_ADOFields PTR PTR
'   OPERATOR = @m_pFields
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoFields.vptr () AS Afx_ADOFields PTR PTR
   CADODB_DP("")
   IF m_pFields THEN m_pFields->Release
   RETURN @m_pFields
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cFields AS CAdoFields) AS Afx_ADOFields PTR
   OPERATOR = cFields.m_pFields
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOFields object to the class
' ===========================================================================================
PRIVATE SUB CAdoFields.Attach (BYVAL pFields AS Afx_ADOFields PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("")
   IF m_pFields THEN m_pFields->Release
   m_pFields = pFields
   IF fAddRef THEN
      IF m_pFields THEN m_pFields->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Retrieves the number of fields
' ===========================================================================================
PRIVATE PROPERTY CAdoFields.Count () AS LONG
   CADODB_DP("")
   DIM nCount AS LONG
   IF m_pFields THEN SetResult(m_pFields->get_Count(@nCount))
   RETURN nCount
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates a specific member of the Fields collection, by name or ordinal number.
' ===========================================================================================
PRIVATE PROPERTY CAdoFields.Item (BYREF dvIndex AS DVARIANT) AS Afx_ADOField PTR
   CADODB_DP("")
   DIM pField AS Afx_ADOField PTR
   IF m_pFields THEN SetResult(m_pFields->get_Item(dvIndex, @pField))
   RETURN pField
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Appends an object to a collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Append (BYREF wszName AS WsTRING, BYVAL nType AS DataTypeEnum, BYVAL DefinedSize AS ADO_LONGPTR = 0, _
BYVAL Attrib AS FieldAttributeEnum = 0, BYREF dvFieldValue AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   CADODB_DP("")
   DIM bstrName AS AFX_BSTR = SysAllocString(wszName)
   IF m_pFields THEN SetResult(m_pFields->Append(bstrName, nType, DefinedSize, Attrib, dvFieldValue))
   SysFreeString bstrName
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Deletes an object from the Fields collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Delete_ (BYREF dvIndex AS DVARIANT) AS HRESULT
   CADODB_DP("")
   IF m_pFields THEN RETURN(SetResult(m_pFields->Delete_(dvIndex)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the collection's enumerator.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.NewEnum () AS Afx_IUnknown PTR
   CADODB_DP("")
   DIM pUnk AS Afx_IUnknown PTR
   IF m_pFields THEN SetResult(m_pFields->_NewEnum(@pUnk))
   RETURN pUnk
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Refreshes the contents of the Properties collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Refresh () AS HRESULT
   CADODB_DP("")
   IF m_pFields THEN RETURN(SetResult(m_pFields->Refresh))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Resynchronizes the contents of the Fields collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Resync (BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   CADODB_DP("")
   IF m_pFields THEN RETURN(SetResult(m_pFields->Resync))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves any changes you make to the current Fields collection of a Record object.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.Update () AS HRESULT
   CADODB_DP("")
   IF m_pFields THEN RETURN(SetResult(m_pFields->Update))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels any changes made to the Fields collection of a Record object before calling the Update method.
' ===========================================================================================
PRIVATE FUNCTION CAdoFields.CancelUpdate () AS HRESULT
   CADODB_DP("")
   IF m_pFields THEN RETURN(SetResult(m_pFields->CancelUpdate))
END FUNCTION
' ===========================================================================================
