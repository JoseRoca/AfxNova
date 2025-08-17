' ########################################################################################
' Platform: Microsoft Windows
' File: CAdoRecord.inc
' Contents: ADO Record class
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
' CAdoRecord class
' ########################################################################################
TYPE CAdoRecord EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pRecord AS Afx_ADORecord PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pRecord AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADORecord PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADORecord PTR PTR

   ' // Attaches a reference to an Afx_ADORecord object to the class
   DECLARE SUB Attach (BYVAL pRecord AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Returns a reference to the Properties collection.
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Returns a reference to the Fields collection.
   DECLARE PROPERTY Fields () AS Afx_ADOFields PTR
   ' // * Determines the Connection object over which the specified Command object will execute
   DECLARE PROPERTY ActiveConnection (BYREF wszConn AS WSTRING)
   DECLARE PROPERTY ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   DECLARE PROPERTY ActiveConnection (BYREF pconn AS CAdoConnection)
   DECLARE PROPERTY ActiveConnection () AS Afx_ADOConnection PTR
   ' // Indicates for all applicable objects whether the state of the object is open or closed.
   DECLARE PROPERTY State () AS ObjectStateEnum
   ' // Indicates the type of Record object.
   DECLARE PROPERTY RecordType () AS RecordTypeEnum
   ' // Indicates the data source or object represented by the Record.
   DECLARE PROPERTY Source () AS DVARIANT
   DECLARE PROPERTY Source (BYREF wszSource AS WSTRING)
   DECLARE PROPERTY PutRefSource (BYVAL pSource AS Afx_IDispatch PTR)
   ' // Sets or returns a ConnectModeEnum value. The default value for a Record object is adModeRead.
   DECLARE PROPERTY Mode () AS ConnectModeEnum
   DECLARE PROPERTY Mode (BYVAL Mode AS ConnectModeEnum)
   ' // Indicates an absolute URL string that points to the parent Record of the current Record object.
   DECLARE PROPERTY ParentURL () AS DWSTRING
   ' // Closes a Record object and any dependent objects.
   DECLARE FUNCTION Close () AS HRESULT
   ' // Cancels execution of a pending, asynchronous method call.
   DECLARE FUNCTION Cancel () AS HRESULT
   ' // Returns a Recordset whose rows represent the children of a collection Record.
   DECLARE FUNCTION GetChildren () AS Afx_ADORecordset PTR
   ' // Opens an existing Record object, or creates a new item represented by the Record (such as a file or directory).
   DECLARE FUNCTION Open (BYREF dvSource AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
         BYREF dvActiveConnection AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
         BYVAL nMode AS ConnectModeEnum = adModeUnknown, BYVAL CreateOptions AS RecordCreateOptionsEnum = adFailIfNotExists, _
         BYVAL Options AS RecordOpenOptionsEnum = -1, BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WSTRING = "") AS HRESULT
   ' // Moves a entity represented by a Record to another location.
   DECLARE FUNCTION MoveRecord (BYREF wszSource AS WSTRING = "", BYREF wszDestination AS WSTRING = "", _
         BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WSTRING = "", BYVAL Options AS MoveRecordOptionsEnum = adMoveUnspecified, _
         BYVAL Async AS BOOLEAN = FALSE) AS DWSTRING
   ' // Copies a entity represented by a Record to another location.
   DECLARE FUNCTION CopyRecord (BYREF wszSource AS WSTRING = "", BYREF wszDestination AS WSTRING = "", _
         BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WSTRING = "", BYVAL Options AS MoveRecordOptionsEnum = adCopyUnspecified, _
         BYVAL Async AS BOOLEAN = FALSE) AS DWSTRING
   ' // Deletes a the entity represented by a Record.
   DECLARE FUNCTION DeleteRecord (BYREF wszSource AS WSTRING = "", BYVAL Async AS BOOLEAN = FALSE) AS HRESULT
   ' // Retrieves information from the ADO Errors collection
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS HRESULT = 0) AS DWSTRING

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoRecord
   CADODB_DP("BEGIN")
   ' // Create an instance of the Record object
   m_pRecord = AfxNewCom("ADODB.Record")
   CADODB_DP("END - " & WSTR(m_pRecord))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoRecord (BYVAL pRecord AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("Afx_ADORecord")
   m_pRecord = pRecord
   IF fAddRef THEN
      IF m_pRecord THEN m_pRecord->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoRecord
   CADODB_DP(WSTR(m_pRecord))
   IF m_pRecord THEN
      IF this.State = adStateOpen THEN this.Close
      m_pRecord->Release
   END IF
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoRecord.@ () AS Afx_ADORecord PTR PTR
'   OPERATOR = @m_pRecord
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoRecord.vptr () AS Afx_ADORecord PTR PTR
   CADODB_DP("")
   IF m_pRecord THEN m_pRecord->Release
   RETURN @m_pRecord
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cRecord AS CAdoRecord) AS Afx_ADORecord PTR
   OPERATOR = cRecord.m_pRecord
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADORecord object to the class
' ===========================================================================================
PRIVATE SUB CAdoRecord.Attach (BYVAL pRecord AS Afx_ADORecord PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("")
   IF m_pRecord THEN m_pRecord->Release
   m_pRecord = pRecord
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pRecord THEN m_pRecord->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Properties collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Properties () AS Afx_ADOProperties PTR
   CADODB_DP("")
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pRecord THEN SetResult(m_pRecord->get_Properties(@pProperties))
   PROPERTY = pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Fields collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Fields () AS Afx_ADOFields PTR
   CADODB_DP("")
   DIM pFields AS Afx_ADOFields PTR
   IF m_pRecord THEN SetResult(m_pRecord->get_Fields(@pFields))
   PROPERTY = pFields
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets a reference to the active connection
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ActiveConnection (BYREF wszConn AS WSTRING)
   CADODB_DP("")
   DIM bstrConn AS AFX_BSTR = SysAllocString(wszConn)
   IF m_pRecord THEN SetResult(m_pRecord->put_ActiveConnection(bstrConn))
   SysFreeString bstrConn
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   CADODB_DP("Afx_ADOConnection")
   IF m_pRecord THEN SetResult(m_pRecord->putref_ActiveConnection(pconn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ActiveConnection (BYREF pconn AS CAdoConnection)
   CADODB_DP("CAdoConnection")
   IF m_pRecord THEN SetResult(m_pRecord->putref_ActiveConnection(pconn.m_pConnection))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Gets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ActiveConnection () AS Afx_ADOConnection PTR
   CADODB_DP("Afx_ADOConnection PTR")
   DIM vConnection AS VARIANT
   IF m_pRecord THEN SetResult(m_pRecord->get_ActiveConnection(@vConnection))
   DIM pConnection AS Afx_ADOConnection PTR
   IF vConnection.vt = VT_DISPATCH THEN
      pConnection = cast(Afx_ADOConnection PTR, cast(ULONG_PTR, vConnection.pdispVal))
      IF pConnection THEN pConnection->AddRef
   END IF
   PROPERTY = pConnection
   VariantClear(@vConnection)
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates for all applicable objects whether the state of the object is open or closed.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.State () AS ObjectStateEnum
   CADODB_DP("")
   DIM lState AS ObjectStateEnum
   IF m_pRecord THEN SetResult(m_pRecord->get_State(@lState))
   RETURN lState
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the type of Record object.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.RecordType () AS RecordTypeEnum
   CADODB_DP("")
   DIM lType AS LONG
   IF m_pRecord THEN SetResult(m_pRecord->get_RecordType(@lType))
   RETURN lType
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the data source or object represented by the Record.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Source () AS DVARIANT
   CADODB_DP("")
   DIM dvSource AS DVARIANT
   IF m_pRecord THEN SetResult(m_pRecord->get_Source(dvSource))
   RETURN dvSource
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Source (BYREF wszource AS WSTRING)
   CADODB_DP("")
   DIM bstrSource AS AFX_BSTR = SysAllocString(wszource)
   IF m_pRecord THEN SetResult(m_pRecord->put_Source(bstrSource))
   SysFreeString bstrSource
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.PutRefSource (BYVAL pSource AS Afx_IDispatch PTR)
   CADODB_DP("")
   IF m_pRecord THEN SetResult(m_pRecord->putref_Source(pSource))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a ConnectModeEnum value. The default value for a Record object is adModeRead.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Mode () AS ConnectModeEnum
   CADODB_DP("")
   DIM lMode AS LONG
   IF m_pRecord THEN SetResult(m_pRecord->get_Mode(@lMode))
   RETURN lMode
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.Mode (BYVAL lMode AS ConnectModeEnum)
   CADODB_DP("")
   IF m_pRecord THEN SetResult(m_pRecord->put_Mode(lMode))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates an absolute URL string that points to the parent Record of the current Record object.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecord.ParentURL () AS DWSTRING
   CADODB_DP("")
   DIM bstrURL AS AFX_BSTR
   IF m_pRecord THEN SetResult(m_pRecord->get_ParentURL(@bstrURL))
   DIM dwsURL AS DWSTRING = *bstrURL
   SysFreeString bstrURL
   RETURN dwsURL
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Closes a Record object and any dependent objects.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.Close () AS HRESULT
   CADODB_DP("")
   IF m_pRecord THEN RETURN(SetResult(m_pRecord->Close))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.Cancel () AS HRESULT
   CADODB_DP("")
   IF m_pRecord THEN RETURN(SetResult(m_pRecord->Cancel))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Returns a Recordset whose rows represent the children of a collection Record.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.GetChildren () AS Afx_ADORecordset PTR
   CADODB_DP("")
   DIM pRSet AS Afx_ADORecordset PTR
   IF m_pRecord THEN SetResult(m_pRecord->GetChildren(@pRSet))
   RETURN pRSet
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Opens an existing Record object, or creates a new item represented by the Record (such as a file or directory).
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.Open (BYREF dvSource AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYREF dvActiveConnection AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYVAL nMode AS ConnectModeEnum = adModeUnknown, BYVAL CreateOptions AS RecordCreateOptionsEnum = adFailIfNotExists, _
BYVAL Options AS RecordOpenOptionsEnum = adOpenRecordUnspecified, BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WSTRING = "") AS HRESULT
   CADODB_DP("")
   DIM bstrUserName AS AFX_BSTR = SysAllocString(wszUserName)
   DIM bstrPassword AS AFX_BSTR = SysAllocString(wszPassword)
   IF m_pRecord THEN SetResult(m_pRecord->Open(dvSource, dvActiveConnection, nMode, CreateOptions, Options, bstrUserName, bstrPassword))
   SysFreeString bstrUserName
   SysFreeString bstrPassword
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves a entity represented by a Record to another location.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.MoveRecord (BYREF wszSource AS WSTRING = "", BYREF wszDestination AS WSTRING = "", _
BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WSTRING = "", BYVAL Options AS MoveRecordOptionsEnum = adMoveUnspecified, _
BYVAL Async AS BOOLEAN = FALSE) AS DWSTRING
   CADODB_DP("")
   DIM bstrNewURL AS AFX_BSTR
   DIM bstrSource AS AFX_BSTR = SysAllocString(wszSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocString(wszDestination)
   DIM bstrUserName AS AFX_BSTR = SysAllocString(wszUserName)
   DIM bstrPassword AS AFX_BSTR = SysAllocString(wszPassword)
   IF m_pRecord THEN m_pRecord->MoveRecord(bstrSource, bstrDestination, bstrUserName, bstrPassword, Options, Async, @bstrNewURL)
   DIM dwsNewURL AS DWSTRING = *bstrNewURL
   SysFreeString bstrNewURL
   RETURN dwsNewURL
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Copies a entity represented by a Record to another location.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.CopyRecord (BYREF wszSource AS WSTRING = "", BYREF wszDestination AS WSTRING = "", _
BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WSTRING = "", BYVAL Options AS MoveRecordOptionsEnum = adMoveUnspecified, _
BYVAL Async AS BOOLEAN = FALSE) AS DWSTRING
   CADODB_DP("")
   DIM bstrNewURL AS AFX_BSTR
   DIM bstrSource AS AFX_BSTR = SysAllocString(wszSource)
   DIM bstrDestination AS AFX_BSTR = SysAllocString(wszDestination)
   DIM bstrUserName AS AFX_BSTR = SysAllocString(wszUserName)
   DIM bstrPassword AS AFX_BSTR = SysAllocString(wszPassword)
   IF m_pRecord THEN m_pRecord->CopyRecord(bstrSource, bstrDestination, bstrUserName, bstrPassword, Options, Async, @bstrNewURL)
   DIM dwsNewURL AS DWSTRING = *bstrNewURL
   SysFreeString bstrNewURL
   RETURN dwsNewURL
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Deletes a the entity represented by a Record.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.DeleteRecord (BYREF wszSource AS WSTRING = "", BYVAL Async AS BOOLEAN = FALSE) AS HRESULT
   CADODB_DP("")
   DIM bstrSource AS Afx_BSTR = SysAllocString(wszSource)
   IF m_pRecord THEN RETURN(m_pRecord->DeleteRecord(bstrSource, Async))
   SysFreeString bstrSource
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves information from the ADO Errors collection
' ===========================================================================================
PRIVATE FUNCTION CAdoRecord.GetErrorInfo (BYVAL nError AS HRESULT = 0) AS DWSTRING
   CADODB_DP("")
   DIM vConnection AS VARIANT
   IF m_pRecord THEN m_pRecord->get_ActiveConnection(@vConnection)
   DIM dwsError AS DWSTRING
   IF vConnection.vt = VT_DISPATCH THEN
      dwsError = AfxAdoGetErrorInfo(cast(Afx_ADOConnection PTR, cast(ULONG_PTR, vConnection.pdispVal)), nError)
   END IF
   VariantClear(@vConnection)
   RETURN dwsError
END FUNCTION
' ===========================================================================================
