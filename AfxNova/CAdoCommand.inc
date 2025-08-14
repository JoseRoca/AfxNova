' ########################################################################################
' Platform: Microsoft Windows
' File: CAdoCommand.inc
' Contents: ADO Command class
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

NAMESPACE AfxNova

' ########################################################################################
' CAdoCommand class
' ########################################################################################
TYPE CAdoCommand EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pCommand AS Afx_ADOCommand PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOCommand PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOCommand PTR PTR

   ' // Attaches a reference to an Afx_ADOCommand object to the class
   DECLARE SUB Attach (BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Returns a reference to the Properties collection.
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Returns a reference to the Parameters collection.
   DECLARE PROPERTY Parameters () AS Afx_ADOParameters PTR
   ' // * Determines the Connection object over which the specified Command object will execute
   DECLARE PROPERTY ActiveConnection (BYREF vConn AS DVARIANT)
   DECLARE PROPERTY ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   DECLARE PROPERTY ActiveConnection (BYREF pconn AS CAdoConnection)
   DECLARE PROPERTY ActiveConnection () AS Afx_ADOConnection PTR
   ' // * Sets or returns a BSTR value that contains a provider command
   DECLARE PROPERTY CommandText (BYREF wszText AS WSTRING)
   DECLARE PROPERTY CommandText () AS DWSTRING
   ' // * Executes the query, SQL statement, or stored procedure specified in the CommandText or CommandStream property
   DECLARE FUNCTION Execute (BYVAL RecordsAffected AS LONG PTR = NULL, BYREF dvParameters AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYVAL Options AS LONG = adCmdUnspecified) AS Afx_ADORecordset PTR
   ' // * Sets or returns a Long value that indicates, in seconds, how long to wait for a command to execute
   DECLARE PROPERTY CommandTimeout () AS LONG
   DECLARE PROPERTY CommandTimeout (BYVAL lTimeout AS LONG)
   ' // Sets or returns one or more CommandTypeEnum values.
   DECLARE PROPERTY CommandType (BYVAL lCmdType AS CommandTypeEnum)
   DECLARE PROPERTY CommandType () AS CommandTypeEnum
   ' // Sets or returns a Boolean value that, if set to True, indicates that the command should be prepared.
   DECLARE PROPERTY Prepared (BYVAL fPrepared AS BOOLEAN)
   DECLARE PROPERTY Prepared () AS BOOLEAN
   ' // Sets or returns a BSTR value that indicates the name of a Command object.
   DECLARE PROPERTY Name () AS DWSTRING
   DECLARE PROPERTY Name (BYREF wszName AS WSTRING)
   ' // Indicates the dialect of the CommandText or CommandStream properties.
   DECLARE PROPERTY Dialect (BYREF wszDialect AS WSTRING)
   DECLARE PROPERTY Dialect () AS DWSTRING
   ' // Indicates whether parameter names should be passed to the provider.
   DECLARE PROPERTY NamedParameters (BYVAL fNamedParameters AS BOOLEAN)
   DECLARE PROPERTY NamedParameters () AS BOOLEAN
   ' // Sets or returns the stream used as the input for a Command object.
   DECLARE PROPERTY CommandStream (BYVAL pStream AS Afx_ADOStream PTR)
   DECLARE PROPERTY CommandStream () AS Afx_ADOStream PTR
   ' // Indicates if the Command object is open or closed.
   DECLARE PROPERTY State () AS LONG
   ' // Cancels execution of a pending, asynchronous method call.
   DECLARE FUNCTION Cancel () AS HRESULT
   ' // Creates a new Parameter object with the specified properties.
   DECLARE FUNCTION CreateParameter (BYREF wszName AS WSTRING = "", BYVAL nType AS DataTypeEnum = 0, _
           BYVAL Direction AS ParameterDirectionEnum = adParamInput, BYVAL Size AS LONG = 0, _
           BYREF dvValue AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS Afx_ADOParameter PTR
   ' // Retrieves information from the ADO Errors collection
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS HRESULT = 0) AS DWSTRING

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoCommand
   CADODB_DP("BEGIN")
   ' // Create an instance of the Connection object
'   DIM CLSID_Command_ AS GUID = (&h00000507, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
'   DIM IID_Command_ AS GUID = (&h986761E8, &h7269, &h4890, {&hAA, &h65, &hAD, &h7C, &h03, &h69, &h7A, &h6D})
'   SetResult(CoCreateInstance(@CLSID_Command_, NULL, CLSCTX_INPROC_SERVER, @IID_Command_, @m_pCommand))
   m_pCommand = AfxNewCom("ADODB.Command")
   CADODB_DP("END - " & WSTR(m_pCommand))
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
' Attaches a reference to an Afx_ADOCommand object to the class
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoCommand (BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("BEGIN - Afx_ADOCommand")
   m_pCommand = pCommand
   IF m_pCommand THEN
      IF fAddRef THEN m_pCommand->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoCommand
   CADODB_DP("")
   IF m_pCommand THEN m_pCommand->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoCommand.@ () AS Afx_ADOCommand PTR PTR
'   OPERATOR = @m_pCommand
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoCommand.vptr () AS Afx_ADOCommand PTR PTR
   CADODB_DP("")
   IF m_pCommand THEN m_pCommand->Release
   RETURN @m_pCommand
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cCommand AS CAdoCommand) AS Afx_ADOCommand PTR
   OPERATOR = cCommand.m_pCommand
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOCommand object to the class
' ===========================================================================================
PRIVATE SUB CAdoCommand.Attach (BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("")
   IF m_pCommand THEN m_pCommand->Release
   m_pCommand = pCommand
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pCommand THEN m_pCommand->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Properties collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Properties () AS Afx_ADOProperties PTR
   CADODB_DP("")
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pCommand THEN SetResult(m_pCommand->get_Properties(@pProperties))
   PROPERTY = pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Parameters collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Parameters () AS Afx_ADOParameters PTR
   CADODB_DP("")
   DIM pParameters AS Afx_ADOParameters PTR
   IF m_pCommand THEN SetResult(m_pCommand->get_Parameters(@pParameters))
   PROPERTY = pParameters
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets a reference to the active connection
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.ActiveConnection (BYREF vConn AS DVARIANT)
   CADODB_DP("DVARIANT")
   IF m_pCommand THEN SetResult(m_pCommand->put_ActiveConnection(vConn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   CADODB_DP("Afx_ADOConnection PTR")
   IF m_pCommand THEN SetResult(m_pCommand->putref_ActiveConnection(pconn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.ActiveConnection (BYREF pconn AS CAdoConnection)
   CADODB_DP("CAdoConnection")
   IF m_pCommand THEN SetResult(m_pCommand->putref_ActiveConnection(pconn.m_pConnection))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Gets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.ActiveConnection () AS Afx_ADOConnection PTR
   CADODB_DP("")
   DIM pConnection AS Afx_ADOCOnnection PTR
   IF m_pCommand THEN SetResult(m_pCommand->get_ActiveConnection(@pConnection))
   PROPERTY = pConnection
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a BSTR value that contains a provider command
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandText () AS DWSTRING
   CADODB_DP("")
   DIM bstrText AS AFX_BSTR
   IF m_pCommand THEN SetResult(m_pCommand->get_CommandText(@bstrText))
   DIM dwsText AS DWSTRING = *bstrText
   SysFreeString bstrText
   RETURN dwsText
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandText (BYREF wszText AS WSTRING)
   CADODB_DP("")
   DIM bstrText AS AFX_BSTR = SysAllocString(wszText)
   IF m_pCommand THEN SetResult(m_pCommand->put_CommandText(bstrText))
   SysFreeString bstrText
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Executes the specified query, SQL statement, stored procedure, or provider-specific text
' ===========================================================================================
PRIVATE FUNCTION CAdoCommand.Execute (BYVAL RecordsAffected AS LONG PTR = NULL, BYREF dvParameters AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYVAL Options AS LONG = adCmdUnspecified) AS Afx_ADORecordset PTR
   CADODB_DP("")
   DIM vRecordsAffected AS VARIANT
   V_VT(@vRecordsAffected) = VT_I4
   DIM ppirs AS Afx_ADORecordset PTR
   IF m_pCommand THEN SetResult(m_pCommand->Execute(@vRecordsAffected, dvParameters, Options, @ppirs))
   IF RecordsAffected THEN *RecordsAffected = vRecordsAffected.lVal
   RETURN ppirs
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a Long value that indicates, in seconds, how long to wait for a command to
' execute. Default is 30.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandTimeout () AS LONG
   CADODB_DP("")
   DIM lTimeout AS LONG
   IF m_pCommand THEN SetResult(m_pCommand->get_CommandTimeout(@lTimeout))
   PROPERTY = lTimeout
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandTimeout (BYVAL lTimeout AS LONG)
   CADODB_DP("")
   IF m_pCommand THEN SetResult(m_pCommand->put_CommandTimeout(lTimeout))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns one or more CommandTypeEnum values.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandType (BYVAL lCmdType AS CommandTypeEnum)
   CADODB_DP("")
   IF m_pCommand THEN SetResult(m_pCommand->put_CommandType(lCmdType))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandType () AS CommandTypeEnum
   CADODB_DP("")
   DIM lCmdType AS CommandTypeEnum
   IF m_pCommand THEN SetResult(m_pCommand->get_CommandType(@lCmdType))
   RETURN lCmdType
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a Boolean value that, if set to True, indicates that the command should be prepared.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Prepared (BYVAL fPrepared AS BOOLEAN)
   CADODB_DP("")
   IF m_pCommand THEN SetResult(m_pCommand->put_Prepared(cast(VARIANT_BOOL, fPrepared)))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Prepared () AS BOOLEAN
   CADODB_DP("")
   DIM fPrepared AS VARIANT_BOOL
   IF m_pCommand THEN SetResult(m_pCommand->get_Prepared(@fPrepared))
   RETURN fPrepared
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a BSTR value that indicates the name of a Command object.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Name (BYREF wszName AS WSTRING)
   CADODB_DP("")
   DIM bstrName AS AFX_BSTR = SysAllocString(wszName)
   IF m_pCommand THEN SetResult(m_pCommand->put_Name(bstrName))
   SysFreestring bstrName
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Name () AS DWSTRING
   CADODB_DP("")
   DIM bstrName AS AFX_BSTR
   IF m_pCommand THEN SetResult(m_pCommand->get_Name(@bstrName))
   DIM dwsName AS DWSTRING = *bstrName
   RETURN dwsName
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the dialect of the CommandText or CommandStream properties.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Dialect (BYREF wszDialect AS WSTRING)
   CADODB_DP("")
   DIM bstrDialect AS AFX_BSTR = SysAllocString(wszDialect)
   IF m_pCommand THEN SetResult(m_pCommand->put_Dialect(bstrDialect))
   SysFreeString bstrDialect
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.Dialect () AS DWSTRING
   CADODB_DP("")
   DIM bstrDialect AS AFX_BSTR
   IF m_pCommand THEN SetResult(m_pCommand->get_Dialect(@bstrDialect))
   DIM dwsDialect AS DWSTRING = *bstrDialect
   SysFreeString bstrDialect
   RETURN dwsDialect
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates whether parameter names should be passed to the provider.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.NamedParameters (BYVAL fNamedParameters AS BOOLEAN)
   CADODB_DP("")
   IF m_pCommand THEN SetResult(m_pCommand->put_NamedParameters(fNamedParameters))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.NamedParameters () AS BOOLEAN
   CADODB_DP("")
   DIM fNamedParameters AS VARIANT_BOOL
   IF m_pCommand THEN SetResult(m_pCommand->get_NamedParameters(@fNamedParameters))
   RETURN fNamedParameters
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates whether parameter names should be passed to the provider.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandStream (BYVAL pStream AS Afx_ADOStream PTR)
   CADODB_DP("")
   IF m_pCommand THEN SetResult(m_pCommand->putref_CommandStream(cast(ANY PTR, pStream)))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.CommandStream () AS Afx_ADOStream PTR
   CADODB_DP("")
   DIM vStream AS VARIANT
   IF m_pCommand THEN SetResult(m_pCommand->get_CommandStream(@vStream))
   RETURN cast(Afx_ADOStream PTR, cast(ULONG_PTR, vStream.pdispVal))
   ' Don't clear the vStream variant becuase we are transfering owenership
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates if the Command object is open or closed.
' ===========================================================================================
PRIVATE PROPERTY CAdoCommand.State () AS LONG
   CADODB_DP("")
   DIM lState AS LONG
   IF m_pCommand THEN SetResult(m_pCommand->get_State(@lState))
   RETURN lState
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call.
' ===========================================================================================
PRIVATE FUNCTION CAdoCommand.Cancel () AS HRESULT
   CADODB_DP("")
   IF m_pCommand THEN RETURN(SetResult(m_pCommand->Cancel))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call.
' ===========================================================================================
PRIVATE FUNCTION CAdoCommand.CreateParameter (BYREF wszName AS WsTRING = "", BYVAL nType AS DataTypeEnum = 0, _
BYVAL Direction AS ParameterDirectionEnum = adParamInput, BYVAL Size AS LONG = 0, _
BYREF dvValue AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS Afx_ADOParameter PTR
   CADODB_DP("")
   DIM ppiprm AS Afx_ADOParameter PTR
   DIM bstrName AS AFX_BSTR = SysAllocString(wszName)
   IF m_pCommand THEN SetResult(m_pCommand->CreateParameter(bstrName, nType, Direction, Size, dvValue, @ppiprm))
   SysFreeString bstrName
   RETURN ppiprm
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves information from the ADO Errors collection
' ===========================================================================================
PRIVATE FUNCTION CAdoCommand.GetErrorInfo (BYVAL nError AS HRESULT = 0) AS DWSTRING
   CADODB_DP("")
   DIM pConnection AS Afx_ADOCOnnection PTR
   IF m_pCommand THEN SetResult(m_pCommand->get_ActiveConnection(@pConnection))
   DIM dwsError AS DWSTRING
   IF pConnection THEN
      dwsError = AfxAdoGetErrorInfo(pConnection, nError)
      pConnection->Release
   END IF
   RETURN dwsError
END FUNCTION
' ===========================================================================================

END NAMESPACE
