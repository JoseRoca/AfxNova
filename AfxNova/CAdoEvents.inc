' ########################################################################################
' Platform: Microsoft Windows
' File: CAdoEvents.inc
' Contents: Implementation of CAdoConnectionEvents and CAdoRecordsetEventsVt callback interfaces.
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
' CAdoConnectionEvents class
' Implementation of the ADO ConnectionEventsVt callback interface.
' ########################################################################################
TYPE CAdoConnectionEvents EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG

   ' ConnectionEvents
   DECLARE VIRTUAL FUNCTION InfoMessage (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION BeginTransComplete (BYVAL TransactionLevel AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION CommitTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION RollbackTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION WillExecute (BYVAL Source AS Afx_BSTR PTR, BYVAL CursorType AS CursorTypeEnum PTR, BYVAL LockType AS LockTypeEnum PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ExecuteComplete (BYVAL RecordsAffected AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION WillConnect (BYVAL ConnectionString AS Afx_BSTR PTR, BYVAL UserID AS Afx_BSTR PTR, BYVAL Password AS Afx_BSTR PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION ConnectComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION Disconnect (BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE VIRTUAL DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' ########################################################################################
' Template example to set an ADO ConnectionEventsVt innterface.
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
CONSTRUCTOR CAdoConnectionEvents
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CAdoConnectionEvents
END DESTRUCTOR
' =====================================================================================

' *** IUnknown interface methods ***

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAdoConnectionEvents.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CADODB_DP("riid: " & AfxGuidText(riid))
   DIM IID_ConnectionEventsVt_ AS GUID = (&h00001402, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_ConnectionEventsVt_) OR _
      IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
      ' // Increment the reference count
      this.AddRef
      RETURN S_OK
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' =====================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAdoConnectionEvents.AddRef () AS ULONG
   refCount += 1
   CADODB_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAdoConnectionEvents.Release () AS ULONG
   refCount -= 1
   CADODB_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CADODB_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' // CAdoConnectionEvents

' =====================================================================================
' The InfoMessage event is called whenever a warning occurs during a ConnectionEvent operation.
' =====================================================================================
FUNCTION CAdoConnectionEvents.InfoMessage (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called after the BeginTrans operation.
' =====================================================================================
FUNCTION CAdoConnectionEvents.BeginTransComplete (BYVAL TransactionLevel AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called after the BeginTrans operation.
' =====================================================================================
FUNCTION CAdoConnectionEvents.CommitTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called after the RollbackTrans operation.
' =====================================================================================
FUNCTION CAdoConnectionEvents.RollbackTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The WillExecute event is called just before a pending command executes on a connection.
' =====================================================================================
FUNCTION CAdoConnectionEvents.WillExecute (BYVAL Source AS Afx_BSTR PTR, BYVAL CursorType AS CursorTypeEnum PTR, BYVAL LockType AS LockTypeEnum PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called after a connection ends.
' =====================================================================================
FUNCTION CAdoConnectionEvents.ExecuteComplete (BYVAL RecordsAffected AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The WillConnect event is called before a connection starts.
' =====================================================================================
FUNCTION CAdoConnectionEvents.WillConnect (BYVAL ConnectionString AS Afx_BSTR PTR, BYVAL UserID AS Afx_BSTR PTR, BYVAL Password AS Afx_BSTR PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called after a connection starts.
' =====================================================================================
FUNCTION CAdoConnectionEvents.ConnectComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called after a connection ends.
' =====================================================================================
FUNCTION CAdoConnectionEvents.Disconnect (BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CAdoRecordsetEvents class
' Implementation of the ADO RecordsetEventsVt callback interface.
' ########################################################################################
TYPE CAdoRecordsetEvents EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG

   ' RecordsetEvents
   DECLARE VIRTUAL FUNCTION WillChangeField (BYVAL cFields AS LONG, BYVAL Fields AS VARIANT, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION FieldChangeComplete (BYVAL cFields AS LONG, BYVAL Fields AS VARIANT, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION WillChangeRecord (BYVAL adReason AS EventReasonEnum, BYVAL cRecords AS LONG, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION RecordChangeComplete (BYVAL adReason AS EventReasonEnum, BYVAL cRecords AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION WillChangeRecordset (BYVAL adReason AS EventReasonEnum, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION RecordsetChangeComplete (BYVAL adReason AS EventReasonEnum, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION WillMove (BYVAL adReason AS EventReasonEnum, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION MoveComplete (BYVAL adReason AS EventReasonEnum, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION EndOfRecordset (BYVAL fMoreData AS VARIANT_BOOL PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION FetchProgress (BYVAL Progress AS LONG, BYVAL MaxProgress AS LONG, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION FetchComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE VIRTUAL DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' ########################################################################################
' Template example to set an ADO RecordsetEventsVt innterface.
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
CONSTRUCTOR CAdoRecordsetEvents
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CAdoRecordsetEvents
END DESTRUCTOR
' =====================================================================================

' *** IUnknown interface methods ***

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CAdoRecordsetEvents.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CADODB_DP("riid: " & AfxGuidText(riid))
   DIM IID_RecordsetEventsVt_ AS GUID = (&h00001403, &h0000, &h0010, {&h80, &h00, &h00, &hAA, &h00, &h6D, &h2E, &hA4})
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_RecordsetEventsVt_) OR _
      IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
      ' // Increment the reference count
      this.AddRef
      RETURN S_OK
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' =====================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
FUNCTION CAdoRecordsetEvents.AddRef () AS ULONG
   refCount += 1
   CADODB_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CAdoRecordsetEvents.Release () AS ULONG
   refCount -= 1
   CADODB_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CADODB_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' // CAdoRecordsetEvents

' =====================================================================================
' The WillChangeField event is called before a pending operation changes the value of
' one or more Field objects in the Recordset.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.WillChangeField (BYVAL cFields AS LONG, BYVAL Fields AS VARIANT, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The FieldChangeComplete event is called after the value of one or more Field objects has changed.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.FieldChangeComplete (BYVAL cFields AS LONG, BYVAL Fields AS VARIANT, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The WillChangeRecord event is called before one or more records (rows) in the Recordset change. 
' =====================================================================================
FUNCTION CAdoRecordsetEvents.WillChangeRecord (BYVAL adReason AS EventReasonEnum, BYVAL cRecords AS LONG, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The RecordChangeComplete event is called after one or more records change.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.RecordChangeComplete (BYVAL adReason AS EventReasonEnum, BYVAL cRecords AS LONG, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The WillChangeRecordset event is called before a pending operation changes the Recordset.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.WillChangeRecordset (BYVAL adReason AS EventReasonEnum, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The RecordsetChangeComplete event is called after the Recordset has changed.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.RecordsetChangeComplete (BYVAL adReason AS EventReasonEnum, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The WillMove event is called before a pending operation changes the current position
' in the Recordset.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.WillMove (BYVAL adReason AS EventReasonEnum, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The MoveComplete event is called after the current position in the Recordset changes.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.MoveComplete (BYVAL adReason AS EventReasonEnum, BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The EndOfRecordset event is called when there is an attempt to move to a row past the
' end of the Recordset.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.EndOfRecordset (BYVAL fMoreData AS VARIANT_BOOL PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The FetchProgress event is called periodically during a lengthy asynchronous operation
' to report how many more rows have currently been retrieved into the Recordset.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.FetchProgress (BYVAL Progress AS LONG, BYVAL MaxProgress AS LONG, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' The FetchComplete event is called after all the records in a lengthy asynchronous
' operation have been retrieved into the Recordset.
' =====================================================================================
FUNCTION CAdoRecordsetEvents.FetchComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
   CADODB_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================
