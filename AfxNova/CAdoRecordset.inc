' ########################################################################################
' Platform: Microsoft Windows
' File: CAdoRecordset.inc
' Contents: ADO Recordset class
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
' CAdoRecordset class
' ########################################################################################
TYPE CAdoRecordset EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pRecordset AS Afx_ADORecordset PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pRecordset AS Afx_AdoRecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADORecordset PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADORecordset PTR PTR

   DECLARE FUNCTION Open (BYREF dvSource AS DVARIANT, BYREF cActiveConnection AS CAdoConnection, _
           BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
           BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   DECLARE FUNCTION Open (BYREF dvSource AS DVARIANT, BYREF dvActiveConnection AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
           BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
           BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   DECLARE FUNCTION Open (BYREF dvSource AS DVARIANT, BYVAL pActiveConnection AS CAdoConnection PTR, _
           BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
           BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   DECLARE FUNCTION Open (BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, _
           BYVAL nLockType AS LockTypeEnum = adLockUnspecified, BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   ' // Attaches a recordset to the class
   DECLARE FUNCTION Attach (BYVAL pRecordset AS AFX_ADORecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   ' // * Returns a reference to the Properties collection
   DECLARE PROPERTY Properties () AS Afx_ADOProperties PTR
   ' // Gets a reference to the Fields collection of a Record object
   DECLARE PROPERTY Fields () AS Afx_ADOFields PTR
   ' // Indicates that the current record position is before the first record in a Recordset object
   DECLARE PROPERTY BOF () AS BOOLEAN
   ' // * Indicates that the current record position is after the last record in a Recordset object
   DECLARE PROPERTY EOF () AS BOOLEAN
   ' // Indicates the Command object that created the associated Recordset object
   DECLARE FUNCTION ActiveCommand () AS Afx_ADOCommand PTR
   ' // * Sets a reference to the active connection
   DECLARE PROPERTY ActiveConnection (BYREF dvConn AS DVARIANT)
   DECLARE PROPERTY ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   DECLARE PROPERTY ActiveConnection (BYREF pconn AS CAdoConnection)
   ' // * Gets a reference to the active connection
   DECLARE PROPERTY ActiveConnection () AS Afx_ADOConnection PTR
   ' // * Indicates the data source for a Recordset object.
   DECLARE PROPERTY Source (BYREF wszConn AS WsTRING)
   DECLARE PROPERTY Source (BYVAL pcmd AS Afx_ADOCommand PTR)
   DECLARE PROPERTY Source () AS DVARIANT
   ' // * Sets or returns a Variant expression that evaluates to a valid bookmark
   DECLARE PROPERTY Bookmark () AS DVARIANT
   DECLARE PROPERTY Bookmark (BYREF dvBookmark AS DVARIANT)
   ' // * Sets or returns a Long value that must be greater than 0. Default is 1.
   DECLARE PROPERTY CacheSize () AS LONG
   DECLARE PROPERTY CacheSize (BYVAL size AS LONG)
   ' // * Sets or returns a CursorTypeEnum value. The default value is adOpenForwardOnly.
   DECLARE PROPERTY CursorType () AS CursorTypeEnum
   DECLARE PROPERTY CursorType (BYVAL lCursorType AS CursorTypeEnum)
   ' // * Indicates the maximum number of records to return to a Recordset from a query.
   DECLARE PROPERTY CursorLocation () AS CursorLocationEnum
   DECLARE PROPERTY CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum)
   ' // * Sets or returns the lock type, a Long value that must be greater than 0. Default is 1.
   DECLARE PROPERTY LockType () AS LockTypeEnum
   DECLARE PROPERTY LockType (BYVAL lLockType AS LockTypeEnum)
   ' // Indicates the maximum number of records to return to a Recordset from a query.
   DECLARE PROPERTY MaxRecords () AS LONG
   DECLARE PROPERTY MaxRecords (BYVAL lMaxRecords AS LONG)
   ' // * Indicates how many records constitute one page in the Recordset.
   DECLARE PROPERTY PageSize () AS LONG
   DECLARE PROPERTY PageSize (BYVAL nPageSize AS LONG)
   ' // * Indicates how many pages of data the Recordset object contains
   DECLARE PROPERTY PageCount () AS LONG
   ' // * Indicates the number of records in a Recordset object.
   DECLARE PROPERTY RecordCount () AS LONG
   ' // Indicates which records are to be marshaled back to the server
   DECLARE PROPERTY MarshalOptions () AS MarshalOptionsEnum
   DECLARE PROPERTY MarshalOptions (BYVAL eMarshal AS MarshalOptionsEnum)
   ' // * Sets or returns a Variant value that indicates the value of the object
   DECLARE PROPERTY Collect (BYREF dvIndex AS DVARIANT) AS DVARIANT
   DECLARE PROPERTY Collect (BYREF dvIndex AS DVARIANT, BYREF dvValue AS DVARIANT)
   ' // Indicates the name of the index currently in effect for a Recordset object
   DECLARE PROPERTY Index () AS DWSTRING
   DECLARE PROPERTY Index (BYREF wszIndex AS WsTRING)
   ' // * Returns the Recordset as a string
   DECLARE FUNCTION GetString (BYVAL StringFormat AS StringFormatEnum = adClipString, BYVAL NumRows AS LONG = adReadAll, _
           BYREF wszColumnDelimeter AS WSTRING = "", BYREF wszRowDelimeter AS WSTRING = "", BYREF wszNullExpr AS WSTRING = "") AS DWSTRING
   ' // Indicates an object that contains data to be represented as a Recordset object
   DECLARE PROPERTY DataSource () AS Afx_IUnknown PTR
   DECLARE PROPERTY DataSource (BYVAL punkDataSource AS Afx_IUnknown PTR)
   ' // Indicates the name of the data member that will be retrieved from the object referenced by the DataSource property
   DECLARE PROPERTY DataMember () AS DWSTRING
   DECLARE PROPERTY DataMember (BYREF wszDataMember AS WSTRING)
   ' // * Sets or returns a value that indicates the current record position within the recordset
   DECLARE PROPERTY AbsolutePosition () AS PositionEnum
   DECLARE PROPERTY AbsolutePosition (BYVAL Position AS PositionEnum)
   ' // * Sets or returns a value that indicates the current page position within the recordset
   DECLARE PROPERTY AbsolutePage () AS PositionEnum
   DECLARE PROPERTY AbsolutePage (BYVAL nPage AS PositionEnum)
   ' // * Moves the position of the current record in a Recordset object
   DECLARE FUNCTION Move (BYVAL NumRecords AS LONG, BYVAL start AS BookmarkEnum = adBookmarkCurrent) AS HRESULT
   DECLARE FUNCTION Move (BYVAL NumRecords AS LONG, BYREF dvStart AS DVARIANT) AS HRESULT
   ' // * Moves the current record position to the next record in the Recordset
   DECLARE FUNCTION MoveNext () AS HRESULT
   ' // Moves to the previous record in a specified Recordset object and makes that record the current record
   DECLARE FUNCTION MovePrevious () AS HRESULT
   ' // Moves to the first record in a specified Recordset object and makes that record the current record
   DECLARE FUNCTION MoveFirst () AS HRESULT
   ' // Moves to the last record in a specified Recordset object and makes that record the current record
   DECLARE FUNCTION MoveLast () AS HRESULT
   ' // Indicates the status of the current record with respect to batch updates or other bulk operations
   DECLARE PROPERTY Status () AS LONG
   ' // Indicates for whether the state of the Recordset object is open or closed
   DECLARE PROPERTY State () AS LONG
   ' // Indicates the editing status of the current record
   DECLARE PROPERTY EditMode () AS EditModeEnum
   ' // Indicates, in a hierarchical Recordset object, whether the reference to the underlying child records
   ' // (that is, the chapter) changes when the parent row position changes.
   DECLARE PROPERTY StayInSync () AS BOOLEAN
   DECLARE PROPERTY StayInSync (BYVAL bStayInSync AS BOOLEAN)
   ' // Refreshes the data in the current Recordset object from the underlying database
   DECLARE FUNCTION Resync (BYVAL AffectRecords AS AffectEnum = adAffectAll, BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   ' // Updates the data in a Recordset object by re-executing the query on which the object is based
   DECLARE FUNCTION Requery (BYVAL Options AS LONG = adOpenUnspecified) AS HRESULT
   ' // Moves to the previous record in a specified Recordset object and makes that record the current record
   DECLARE FUNCTION NextRecordset (BYVAL RecordsAffected AS LONG PTR = NULL) AS Afx_AdoRecordset PTR
   ' // Determines whether a specified Recordset object supports a particular type of functionality
   DECLARE FUNCTION Supports (BYVAL CursorOptions AS CursorOptionEnum) AS BOOLEAN
   ' // Writes all pending batch updates to disk
   DECLARE FUNCTION UpdateBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   ' // Cancels a pending batch update
   DECLARE FUNCTION CancelBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   ' // Cancels any changes made to the current or new row of a Recordset object before calling the Update method
   DECLARE FUNCTION CancelUpdate () AS HRESULT
   ' // Closes a Recordset object and any dependent objects
   DECLARE FUNCTION Close () AS HRESULT
   ' // Creates a duplicate Recordset object from an existing Recordset object. Optionally, specifies that the clone be read-only
   DECLARE FUNCTION Clone (BYVAL nLockType AS LockTypeEnum = adLockUnspecified) AS Afx_AdoRecordset PTR
   ' // Cancels execution of a pending, asynchronous method call
   DECLARE FUNCTION Cancel () AS HRESULT
   ' // Indicates a filter for data in a Recordset
   DECLARE PROPERTY Filter () AS DVARIANT
   DECLARE PROPERTY Filter (BYVAL dvCriteria AS DVARIANT)
   ' // Indicates one or more field names on which the Recordset is sorted,
   ' // and whether each field is sorted in ascending or descending order
   DECLARE PROPERTY Sort () AS DWSTRING
   DECLARE PROPERTY Sort (BYREF wszCriteria AS WSTRING)
   ' // Deletes the current record or a group of records.
   DECLARE FUNCTION Delete_ (BYVAL AffectRecords AS AffectEnum = adAffectCurrent) AS HRESULT
   ' // Saves the Recordset in a file
   DECLARE FUNCTION Save (BYREF dvDestination AS DVARIANT, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
   ' // Saves the Recordset in a Stream object
   DECLARE FUNCTION Save (BYVAL pDestination AS Afx_AdoStream PTR, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
   ' // Compares two bookmarks and returns an indication of their relative values.
   DECLARE FUNCTION CompareBookmarks (BYREF dvBookmark1 AS DVARIANT, BYREF dvBookmark2 AS DVARIANT) AS CompareEnum
   ' // * Searches a Recordset for the row that satisfies the specified criteria
   DECLARE FUNCTION Find (BYREF wszCriteria AS WSTRING, BYREF dvStart AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
           BYVAL SkipRecords AS LONG = 0, BYVAL SearchDirection AS SearchDirectionEnum = adSearchForward) AS HRESULT
   ' // * Retrieves multiple records of a Recordset object into an array
   DECLARE FUNCTION GetRows (BYVAL Rows AS LONG = adGetRowsRest, BYREF dvStart AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF cvFields AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS SAFEARRAY PTR
   ' // * Searches the index of a Recordset to quickly locate the row that matches the specified values, and changes the current row position to that row
   DECLARE FUNCTION Seek (BYREF dvKeyValues AS DVARIANT, BYVAL SeekOption AS SeekEnum = adSeekFirstEQ) AS HRESULT
   ' // Creates a new record for an updatable Recordset object.
   DECLARE FUNCTION AddNew (BYREF dvFieldList AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF dvValues AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   DECLARE FUNCTION AddNew (BYVAL vFieldList AS VARIANT, BYVAL vValues AS VARIANT) AS HRESULT
   ' // Saves any changes you make to the current row of a Recordset object
   DECLARE FUNCTION Update (BYREF dvFieldList AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF dvValues AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   DECLARE FUNCTION Update (BYVAL vFieldList AS VARIANT, BYVAL vValues AS VARIANT) AS HRESULT
   ' // Retrieves information from the ADO Errors collection
   DECLARE FUNCTION GetErrorInfo (BYVAL nError AS HRESULT = 0) AS DWSTRING

END TYPE

' ========================================================================================
' Default constructor
' ========================================================================================
CONSTRUCTOR CAdoRecordset
   CADODB_DP("")
   ' // Create an instance of the Recordset object
'   DIM CLSID_Recordset_ AS GUID = (&h00000535, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
'   DIM IID_Recordset_ AS GUID = (&h00001556, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
'   SetResult(CoCreateInstance(@CLSID_Recordset_, NULL, CLSCTX_INPROC_SERVER, @IID_Recordset_, @m_pRecordset))
   m_pRecordset = AfxNewCom("ADODB.Recordset")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
CONSTRUCTOR CAdoRecordset (BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("Afx_ADORecordset")
   m_pRecordset = pRecordset
   IF fAddRef THEN
      IF m_pRecordset THEN m_pRecordset->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
DESTRUCTOR CAdoRecordset
   CADODB_DP("")
   ' // Close and release the recordset
   IF m_pRecordset THEN
      IF this.State = adStateOpen THEN this.Close
      m_pRecordset->Release
   END IF
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoRecordset.@ () AS Afx_ADORecordset PTR PTR
'   CADODB_DP("CAdoRecordset Operator @")
'   OPERATOR = @m_pRecordset
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoRecordset.vptr () AS Afx_ADORecordset PTR PTR
   CADODB_DP("")
   IF m_pRecordset THEN m_pRecordset->Release
   RETURN @m_pRecordset
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF crs AS CAdoRecordset) AS Afx_ADORecordset PTR
   CADODB_DP("")
   OPERATOR = crs.m_pRecordset
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Open (BYREF dvSource AS DVARIANT, BYREF cActiveConnection AS CAdoConnection, _
BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   CADODB_DP("cActiveConnection")
   ' // Make a VT_DISPATCH variant with the active connection pointer
   DIM vConnection AS VARIANT
   vConnection.vt = VT_DISPATCH
   vConnection.pdispVal = cast(IDispatch PTR, cast(ULONG_PTR, cActiveConnection.m_pConnection))
   IF vConnection.pdispVal THEN vConnection.pdispVal->lpVtbl->AddRef(vConnection.pdispVal)
   ' // Call the Open method
   IF m_pRecordset THEN SetResult(m_pRecordset->Open(dvSource, vConnection, nCursorType, nLockType, nOptions))
   ' // Clear the variant
   VariantClear(@vConnection)
   ' // Return the result code
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Open (BYREF dvSource AS DVARIANT, BYREF dvActiveConnection AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   CADODB_DP("dvActiveConnection")
   IF m_pRecordset THEN SetResult(m_pRecordset->Open(dvSource, dvActiveConnection, nCursorType, nLockType, nOptions))
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Open (BYREF dvSource AS DVARIANT, BYVAL pActiveConnection AS CAdoConnection PTR, _
BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, BYVAL nLockType AS LockTypeEnum = adLockUnspecified, _
BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   CADODB_DP("dvSource - pActiveConnection")

   ' // Make a VT_DISPATCH variant with the active connection pointer
   DIM vConnection AS VARIANT
   vConnection.vt = VT_DISPATCH
   vConnection.pdispVal = cast(IDispatch PTR, cast(ULONG_PTR, pActiveConnection->m_pConnection))
   IF vConnection.pdispVal THEN vConnection.pdispVal->lpVtbl->AddRef(vConnection.pdispVal)
   ' // Call the Open method
   IF m_pRecordset THEN SetResult(m_pRecordset->Open(dvSource, vConnection, nCursorType, nLockType, nOptions))
   ' // Clear the variant
   VariantClear(@vConnection)
   ' // Return the result code
   RETURN m_Result

END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Opens a connection to a data source.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Open (BYVAL nCursorType AS CursorTypeEnum = adOpenUnspecified, _
BYVAL nLockType AS LockTypeEnum = adLockUnspecified, BYVAL nOptions AS LONG = adCmdUnspecified) AS HRESULT
   CADODB_DP("")
   ' // Call the Open method
   DIM vSource AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)
   DIM vConnection AS VARIANT = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)
   IF m_pRecordset THEN SetResult(m_pRecordset->Open(vSource, vConnection, nCursorType, nLockType, nOptions))
   ' // Return the result code
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ========================================================================================
' Attaches a recordset to the class.
' Useful to attach to the class a pointer returned by a function.
' ========================================================================================
PRIVATE FUNCTION CAdoRecordset.Attach (BYVAL pRecordset AS AFX_ADORecordset PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   CADODB_DP("pRecordset")
   IF m_pRecordset = NULL THEN RETURN E_INVALIDARG
   ' // Close and release the recordset
   IF this.State = adStateOpen THEN this.Close
   m_pRecordset->Release
   ' // Attach the recordset
   m_pRecordset = pRecordset
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pRecordset THEN m_pRecordset->AddRef
   END IF
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ===========================================================================================
' * Returns a reference to the Properties collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Properties () AS Afx_ADOProperties PTR
   CADODB_DP("")
   DIM pProperties AS Afx_ADOProperties PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Properties(@pProperties))
   PROPERTY = pProperties
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Returns a reference to the Fields collection.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Fields () AS Afx_ADOFields PTR
   CADODB_DP("")
   DIM pFields AS Afx_ADOFields PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Fields(@pFields))
   PROPERTY = pFields
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates that the current record position is before the first record in a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.BOF () AS BOOLEAN
   CADODB_DP("")
   DIM pb AS VARIANT_BOOL
   IF m_pRecordset THEN SetResult(m_pRecordset->get_BOF(@pb))
   PROPERTY = pb
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Indicates that the current record position is after the last record in a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.EOF () AS BOOLEAN
   CADODB_DP("")
   DIM pb AS VARIANT_BOOL
   IF m_pRecordset THEN SetResult(m_pRecordset->get_EOF(@pb))
   PROPERTY = pb
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the Command object that created the associated Recordset object
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.ActiveCommand () AS Afx_ADOCommand PTR
   DIM pcmd AS Afx_IDispatch PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_ActiveCommand(@pcmd))
   RETURN cast(Afx_ADOCommand PTR, cast(ULONG_PTR, pcmd))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Sets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.ActiveConnection (BYVAL pconn AS Afx_ADOConnection PTR)
   CADODB_DP("Afx_ADOConnection")
   IF m_pRecordset THEN SetResult(m_pRecordset->putref_ActiveConnection(pconn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Sets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.ActiveConnection (BYREF pconn AS CAdoConnection)
   CADODB_DP("CAdoConnection")
   IF m_pRecordset THEN SetResult(m_pRecordset->putref_ActiveConnection(pconn.m_pConnection))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Sets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.ActiveConnection (BYREF vConn AS DVARIANT)
   CADODB_DP("DVARIANT")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_ActiveConnection(vConn))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
' * Gets a reference to the active connection
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.ActiveConnection () AS Afx_ADOConnection PTR
   CADODB_DP("Afx_ADOConnection PTR")
   DIM vConnection AS VARIANT
   IF m_pRecordset THEN SetResult(m_pRecordset->get_ActiveConnection(@vConnection))
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
' * Indicates the data source for a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Source (BYREF wszConn AS WSTRING)
   CADODB_DP(wszConn)
   DIM bstrConn AS AFX_BSTR = SysAllocString(wszConn)
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Source(bstrConn))
   SysFreeString bstrConn
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Source (BYVAL pcmd AS Afx_ADOCommand PTR)
   CADODB_DP("IDispatch PTR")
   IF m_pRecordset THEN SetResult(m_pRecordset->putref_Source(pcmd))
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Source () AS DVARIANT
   CADODB_DP("DVARIANT")
   DIM dvSource AS DVARIANT
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Source(dvSource))
   PROPERTY = dvSource
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a Variant expression that evaluates to a valid bookmark
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Bookmark () AS DVARIANT
   CADODB_DP("DVARIANT")
   DIM dvSource AS DVARIANT
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Bookmark(dvSource))
   PROPERTY = dvSource
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Bookmark (BYREF dvBookmark AS DVARIANT)
   CADODB_DP("DVARIANT")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Bookmark(dvBookmark))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a Long value that must be greater than 0. Default is 1.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CacheSize () AS LONG
   CADODB_DP("LONG")
   DIM size AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_CacheSize(@size))
   PROPERTY = size
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CacheSize (BYVAL size AS LONG)
   CADODB_DP("")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_CacheSize(size))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a CursorTypeEnum value. The default value is adOpenForwardOnly.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CursorType () AS CursorTypeEnum
   CADODB_DP("CursorTypeEnum")
   DIM lCursorType AS CursorTypeEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_CursorType(@lCursorType))
   PROPERTY = lCursorType
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CursorType (BYVAL lCursorType AS CursorTypeEnum)
   CADODB_DP("")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_CursorType(lCursorType))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the location of the cursor service
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CursorLocation () AS CursorLocationEnum
   CADODB_DP("CursorLocationEnum")
   DIM lCursorLoc AS CursorLocationEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_CursorLocation(@lCursorLoc))
   PROPERTY = lCursorLoc
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.CursorLocation (BYVAL lCursorLoc AS CursorLocationEnum)
   CADODB_DP("")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_CursorLocation(lCursorLoc))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a Long value that must be greater than 0. Default is 1.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.LockType () AS LockTypeEnum
   CADODB_DP("LockTypeEnum")
   DIM lLockType AS LockTypeEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_LockType(@lLockType))
   PROPERTY = lLockType
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.LockType (BYVAL lLockType AS LockTypeEnum)
   CADODB_DP("")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_LockType(lLockType))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the maximum number of records to return to a Recordset from a query
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.MaxRecords () AS LONG
   CADODB_DP("LONG")
   DIM lMaxRecords AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_MaxRecords(@lMaxRecords))
   PROPERTY = lMaxRecords
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.MaxRecords (BYVAL lMaxRecords AS LONG)
   CADODB_DP("")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_MaxRecords(lMaxRecords))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates how many records constitute one page in the Recordset
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.PageSize () AS LONG
   CADODB_DP("LONG")
   DIM nPageSize AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_PageSize(@nPageSize))
   PROPERTY = nPageSize
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.PageSize (BYVAL nPageSize AS LONG)
   CADODB_DP("")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_PageSize(nPageSize))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates how many pages of data the Recordset object contains
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.PageCount () AS LONG
   CADODB_DP("")
   DIM nPageCount AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_PageSize(@nPageCount))
   PROPERTY = nPageCount
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates how many pages of data the Recordset object contains
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.RecordCount () AS LONG
   CADODB_DP("")
   DIM pl AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_RecordCount(@pl))
   PROPERTY = pl
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates which records are to be marshaled back to the server
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.MarshalOptions () AS MarshalOptionsEnum
   CADODB_DP("MarshalOptionsEnum")
   DIM eMarshal AS MarshalOptionsEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_MarshalOptions(@eMarshal))
   PROPERTY = eMarshal
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.MarshalOptions (BYVAL eMarshal AS MarshalOptionsEnum)
   CADODB_DP("")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_MarshalOptions(eMarshal))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Sets or returns a Variant value that indicates the value of the object
' The ADO Recorset object exposes a hidden member: the Collect property. This property is
' functionally similar to the Field's Value property, but it doesn't need a reference (explicit
' or implicit) to the Field object. You can pass either a numeric index or a field's name to
' this property.
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Collect (BYREF dvIndex AS DVARIANT) AS DVARIANT
   CADODB_DP("DVARIANT")
   DIM dvValue AS DVARIANT
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Collect(dvIndex, dvValue))
   PROPERTY = dvValue
END PROPERTY
' ===========================================================================================

' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Collect (BYREF dvIndex AS DVARIANT, BYREF dvValue AS DVARIANT)
   CADODB_DP("Index, Value")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Collect(dvIndex, dvValue))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the name of the index currently in effect for a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Index () AS DWSTRING
   CADODB_DP("DWSTRING")
   DIM bstrIndex AS AFX_BSTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Index(@bstrIndex))
   DIM dwsIndex AS DWSTRING = *bstrIndex
   PROPERTY = dwsIndex
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Index (BYREF wszIndex AS WSTRING)
   CADODB_DP("")
   DIM bstrIndex AS AFX_BSTR = SysAllocString(wszIndex)
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Index(bstrIndex))
   SysFreeString bstrIndex
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns the Recordset as a string
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.GetString (BYVAL StringFormat AS StringFormatEnum = adClipString, BYVAL NumRows AS LONG = adReadAll, _
BYREF wszColumnDelimeter AS WSTRING = "", BYREF wszRowDelimeter AS WSTRING = "", BYREF wszNullExpr AS WSTRING = "") AS DWSTRING
   CADODB_DP("")
   DIM bstrOutput AS AFX_BSTR
   DIM bstrColumnDelimeter AS AFX_BSTR = SysAllocString(wszColumnDelimeter)
   DIM bstrRowDelimeter AS AFX_BSTR = SysAllocString(wszRowDelimeter)
   DIM bstrNullPtr AS AFX_BSTR = SysAllocString(wszNullExpr)
   IF m_pRecordset THEN SetResult(m_pRecordset->GetString(StringFormat, NumRows, bstrColumnDelimeter, bstrRowDelimeter, bstrNullPtr, @bstrOutput))
   DIM dwsOutput AS DWSTRING = *bstrOutput
   RETURN dwsOutput
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Indicates an object that contains data to be represented as a Recordset object
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.DataSource () AS Afx_IUnknown PTR
   CADODB_DP("Afx_IUnknown PTR")
   DIM pDataSource AS Afx_IUnknown PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_DataSource(@pDataSource))
   PROPERTY = pDataSource
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.DataSource (BYVAL punkDataSource AS Afx_IUnknown PTR)
   CADODB_DP("IUnknown PTR")
   IF m_pRecordset THEN SetResult(m_pRecordset->putref_DataSource(punkDataSource))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the name of the data member that will be retrieved from the object referenced by
' the DataSource property.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.DataMember () AS DWSTRING
   CADODB_DP("DWSTRING")
   DIM bstrDataMember AS AFX_BSTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_DataMember(@bstrDataMember))
   DIM dwsDataMember AS DWSTRING = *bstrDataMember
   PROPERTY = dwsDataMember
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.DataMember (BYREF wszDataMember AS WSTRING)
   CADODB_DP("DWSTRING")
   DIM bstrDataMember AS AFX_BSTR = SysALlocString(wszDataMember)
   IF m_pRecordset THEN SetResult(m_pRecordset->put_DataMember(bstrDataMember))
   SysFreeString bstrDataMember
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a value that indicates the current record position within the recordset
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.AbsolutePosition () AS PositionEnum
   CADODB_DP("PositionEnum")
   DIM Position AS PositionEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_AbsolutePosition(@Position))
   PROPERTY = Position
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.AbsolutePosition (BYVAL Position AS PositionEnum)
   CADODB_DP("PositionEnum")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_AbsolutePosition(Position))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Sets or returns a value that indicates the current page position within the recordset
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.AbsolutePage () AS PositionEnum
   CADODB_DP("PositionEnum")
   DIM nPage AS PositionEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_AbsolutePage(@nPage))
   PROPERTY = nPage
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.AbsolutePage (BYVAL nPage AS PositionEnum)
   CADODB_DP("PositionEnum")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_AbsolutePage(nPage))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Moves the position of the current record in a Recordset object.
' - NumRecords: A signed Long expression that specifies the number of records that the current
'   record position moves.
' - start: Optional. A BookmarkEnum value.
'   adBookmarkCurrent (0) Starts at the current record.
'   adBookmarkFirst (1) Starts at the first record.
'   adBookmarkLast (2) Starts at the last record.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Move (BYVAL NumRecords AS LONG, BYVAL start AS BookmarkEnum = adBookmarkCurrent) AS HRESULT
   CADODB_DP("BookmarkEnum")
   DIM vStart AS VARIANT : V_VT(@vStart) = VT_I4 : V_I4(@vStart) = start
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Move(NumRecords, vStart)))
END FUNCTION
' ===========================================================================================
' ===========================================================================================
' Moves the position of the current record in a Recordset object.
' - NumRecords: A signed Long expression that specifies the number of records that the current
'   record position moves.
' - cvStart: Optional. A Variant value that evaluates to a bookmark.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Move (BYVAL NumRecords AS LONG, BYREF dvStart AS DVARIANT) AS HRESULT
   CADODB_DP("DVARIANT")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Move(NumRecords, dvStart)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Moves the current record position to the next record in the Recordset
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.MoveNext () AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->MoveNext))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves the current record position to the previous record in the Recordset
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.MovePrevious () AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->MovePrevious))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves the current record position to the first record in the Recordset.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.MoveFirst () AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->MoveFirst))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves the current record position to the last record in the Recordset.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.MoveLast () AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->MoveLast))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Indicates the status of the current record with respect to batch updates or other bulk operations
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Status () AS LONG
   CADODB_DP("")
   DIM lStatus AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Status(@lStatus))
   PROPERTY = lStatus
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates for whether the state of the Recordset object is open or closed
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.State () AS LONG
   CADODB_DP("")
   DIM lState AS LONG
   IF m_pRecordset THEN SetResult(m_pRecordset->get_State(@lState))
   PROPERTY = lState
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the editing status of the current record
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.EditMode () AS EditModeEnum
   CADODB_DP("")
   DIM nMode AS EditModeEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->get_EditMode(@nMode))
   PROPERTY = nMode
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates, in a hierarchical Recordset object, whether the reference to the underlying child
' records (that is, the chapter) changes when the parent row position changes.
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.StayInSync () AS BOOLEAN
   CADODB_DP("")
   DIM bStayInSync AS VARIANT_BOOL
   IF m_pRecordset THEN SetResult(m_pRecordset->get_StayInSync(@bStayInSync))
   PROPERTY = bStayInSync
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.StayInSync (BYVAL bStayInSync AS BOOLEAN)
   CADODB_DP("")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_StayInSync(bStayInSync))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Refreshes the data in the current Recordset object from the underlying database
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Resync (BYVAL AffectRecords AS AffectEnum = adAffectAll, BYVAL ResyncValues AS ResyncEnum = adResyncAllValues) AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Resync(AffectRecords, ResyncValues)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Updates the data in a Recordset object by re-executing the query on which the object is based
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Requery (BYVAL Options AS LONG = adOpenUnspecified) AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Requery(Options)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Moves to the previous record in a specified Recordset object and makes that record the current record
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.NextRecordset (BYVAL RecordsAffected AS LONG PTR = NULL) AS Afx_AdoRecordset PTR
   CADODB_DP("")
   DIM vRecordsAffected AS VARIANT, pRecordset AS Afx_AdoRecordset PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->NextRecordset(@vRecordsAffected, @pRecordset))
   IF RecordsAffected THEN *RecordsAffected = vRecordsAffected.lVal
   RETURN pRecordset
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Determines whether a specified Recordset object supports a particular type of functionality
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Supports (BYVAL CursorOptions AS CursorOptionEnum) AS BOOLEAN
   CADODB_DP("")
   DIM pb AS VARIANT_BOOL
   IF m_pRecordset THEN SetResult(m_pRecordset->Supports(CursorOptions, @pb))
   RETURN pb
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Writes all pending batch updates to disk
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.UpdateBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->UpdateBatch))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels a pending batch update
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.CancelBatch (BYVAL AffectRecords AS AffectEnum = adAffectAll) AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->CancelBatch))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels any changes made to the current or new row of a Recordset object before calling the Update method
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.CancelUpdate () AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->CancelUpdate))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Closes a Recordset object and any dependent objects
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Close () AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Close))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Closes a Recordset object and any dependent objects
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Clone (BYVAL nLockType AS LockTypeEnum = adLockUnspecified) AS Afx_AdoRecordset PTR
   CADODB_DP("")
   DIM ppvObject AS Afx_AdoRecordset PTR
   IF m_pRecordset THEN SetResult(m_pRecordset->Clone(nLockType, @ppvObject))
   RETURN ppvObject
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Cancels execution of a pending, asynchronous method call
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Cancel () AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Cancel))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Indicates a filter for data in a Recordset
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Filter () AS DVARIANT
   CADODB_DP("")
   DIM dvCriteria AS DVARIANT
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Filter(dvCriteria))
   PROPERTY = dvCriteria
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Filter (BYVAL dvCriteria AS DVARIANT)
   CADODB_DP("")
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Filter(dvCriteria))
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates one or more field names on which the Recordset is sorted, and whether each field
' is sorted in ascending or descending order
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Sort () AS DWSTRING
   CADODB_DP("")
   DIM bstrCriteria AS AFX_BSTR
   IF m_pRecordset THEN SetResult(m_pRecordset->get_Sort(@bstrCriteria))
   DIM dwsCriteria AS DWSTRING = *bstrCriteria
   PROPERTY = dwsCriteria
END PROPERTY
' ===========================================================================================
' ===========================================================================================
PRIVATE PROPERTY CAdoRecordset.Sort (BYREF wszCriteria AS WSTRING)
   CADODB_DP("")
   DIM bstrCriteria AS AFX_BSTR = SysAllocString(wszCriteria)
   IF m_pRecordset THEN SetResult(m_pRecordset->put_Sort(bstrCriteria))
   SysFreeString bstrCriteria
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Deletes the current record or a group of records.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Delete_ (BYVAL AffectRecords AS AffectEnum = adAffectCurrent) AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Delete_(AffectRecords)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves the Recordset in a file
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Save (BYREF dvDestination AS DVARIANT, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
   CADODB_DP("DWSTRING")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Save(dvDestination, PersistFormat)))
END FUNCTION
' ===========================================================================================
' ===========================================================================================
' Saves the Recordset in a Stream object
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Save (BYVAL pDestination AS Afx_AdoStream PTR, BYVAL PersistFormat AS PersistFormatEnum = adPersistADTG) AS HRESULT
   CADODB_DP("Afx_AdoStream")
   DIM vDestination AS VARIANT : V_VT(@vDestination) = VT_BSTR : V_DISPATCH(@vDestination) = cast(IDispatch PTR, cast(ULONG_PTR, pDestination))
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Save(vDestination, PersistFormat)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Compares two bookmarks and returns an indication of their relative values.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.CompareBookmarks (BYREF dvBookmark1 AS DVARIANT, BYREF dvBookmark2 AS DVARIANT) AS CompareEnum
   CADODB_DP("")
   DIM pCompare AS CompareEnum
   IF m_pRecordset THEN SetResult(m_pRecordset->CompareBookmarks(dvBookmark1, dvBookmark2, @pCompare))
   RETURN pCompare
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Searches a Recordset for the row that satisfies the specified criteria
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Find (BYREF wszCriteria AS WSTRING, BYREF dvStart AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), _
BYVAL SkipRecords AS LONG = 0, BYVAL SearchDirection AS SearchDirectionEnum = adSearchForward) AS HRESULT
   CADODB_DP("")
   DIM bstrCriteria AS AFX_BSTR = SysAllocString(wszCriteria)
   IF m_pRecordset THEN SetResult(m_pRecordset->Find(bstrCriteria, SkipRecords, SearchDirection, dvStart))
   SysFreeString bstrCriteria
   RETURN m_Result
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Retrieves multiple records of a Recordset object into an array
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.GetRows (BYVAL Rows AS LONG = adGetRowsRest, BYREF dvStart AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF dvFields AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS SAFEARRAY PTR
   CADODB_DP("")
   DIM vRows AS VARIANT
   IF m_pRecordset THEN SetResult(m_pRecordset->GetRows(Rows, dvStart, dvFields, @vRows))
   ' // Don't clear the variant because we are transfering ownership of the safe array
   IF m_Result = S_OK THEN RETURN vRows.parray
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' * Searches the index of a Recordset to quickly locate the row that matches the specified values, and changes the current row position to that row
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Seek (BYREF dvKeyValues AS DVARIANT, BYVAL SeekOption AS SeekEnum = adSeekFirstEQ) AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN RETURN(SetResult(m_pRecordset->Seek(dvKeyValues, SeekOption)))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Creates a new record for an updatable Recordset object.
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.AddNew (BYREF dvFieldList AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF dvValues AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   CADODB_DP("DVARIANT")
   IF m_pRecordset THEN FUNCTION = SetResult(m_pRecordset->AddNew(dvFieldList, dvValues))
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.AddNew (BYVAL vFieldList AS VARIANT, BYVAL vValues AS VARIANT) AS HRESULT
   CADODB_DP("VARIANT")
   IF m_pRecordset THEN FUNCTION = SetResult(m_pRecordset->AddNew(vFieldList, vValues))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Saves any changes you make to the current row of a Recordset object
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Update (BYREF dvFieldList AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND), BYREF dvValues AS DVARIANT = TYPE<VARIANT>(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)) AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN FUNCTION = SetResult(m_pRecordset->Update(dvFieldList, dvValues))
END FUNCTION
' ===========================================================================================
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.Update (BYVAL vFieldList AS VARIANT, BYVAL vValues AS VARIANT) AS HRESULT
   CADODB_DP("")
   IF m_pRecordset THEN FUNCTION = SetResult(m_pRecordset->Update(vFieldList, vValues))
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Retrieves information from the ADO Errors collection
' ===========================================================================================
PRIVATE FUNCTION CAdoRecordset.GetErrorInfo (BYVAL nError AS HRESULT = 0) AS DWSTRING
   CADODB_DP("")
   DIM vConnection AS VARIANT
   IF m_pRecordset THEN m_pRecordset->get_ActiveConnection(@vConnection)
   DIM dwsError AS DWSTRING
   IF vConnection.vt = VT_DISPATCH THEN
      dwsError = AfxAdoGetErrorInfo(cast(Afx_ADOConnection PTR, cast(ULONG_PTR, vConnection.pdispVal)), nError)
   END IF
   VariantClear(@vConnection)
   RETURN dwsError
END FUNCTION
' ===========================================================================================

END NAMESPACE
