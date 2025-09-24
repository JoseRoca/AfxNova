# Ado events

# Connection events

| Name       | Description |
| ---------- | ----------- |
| [BeginTransComplete](#begintranscomplete) | Called after the **BeginTrans** operation. |
| [CommitTransComplete](#committranscomplete) | Called after the **CommitTrans** operation. |
| [ConnectComplete](#connectcomplete) | Called after a connection starts. |
| [DisConnect](#disconnect) | Called after a connection ends. |
| [ExecuteComplete](#executecomplete) | Called after a command has finished executing. |
| [InfoMessage](#infomessage) | Called whenever a warning occurs during a **ConnectionEvent** operation. |
| [RollbackTransComplete](#rollabcktranscomplete) | Called after the **RollbackTrans** operation. |
| [WillConnect](#willconnect) | Called before a connection starts. |
| [WillExecute](#willexecute) | Called just before a pending command executes on this connection and affords the user an opportunity to examine and modify the pending execution parameters. |

---

# Recordset events

| Name       | Description |
| ---------- | ----------- |
| [EndOfRecordset](#endofrecordset) | Called when there is an attempt to move to a row past the end of the **Recordset**. |
| [FetchComplete](#fetchcomplete) | Called after all the records in a lengthy asynchronous operation have been retrieved into the **Recordset**. |
| [FetchProgress](#fetchprogress) | Called periodically during a lengthy asynchronous operation to report how many rows have currently been retrieved into the **Recordset**. |
| [FieldChangeComplete](#fieldchangecomplete) | Called after the value of one or more Field objects has changed. |
| [MoveComplete](#movecomplete) | Called after the current position in the **Recordset** changes. |
| [RecordChangeComplete](#recordchangecomplete) | Called after one or more records change. |
| [RecordsetChangeComplete](#recordsetchangecomplete) | Called after the **Recordset** has changed. |
| [WillChangeField](#willchangefield) | Called before a pending operation changes the value of one or more **Field** objects in the **Recordset**. |
| [WillChangeRecord](#willchangerecord) | Called before one or more records (rows) in the **Recordset** change. |
| [WillChangeRecordset](#willchangerecordset) | Called before a pending operation changes the **Recordset**. |
| [WillMove](#willmove) | The **WillMove** event is called before a pending operation changes the current position in the **Recordset**. |

---

## WillConnect

Called before a connection starts.

```
FUNCTION WillConnect (BYVAL ConnectionString AS Afx_BSTR PTR, BYVAL UserID AS Afx_BSTR PTR, _
   BYVAL Password AS Afx_BSTR PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, _
   BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *ConnectionString* | A BSTR that contains connection information for the pending connection. |
| *UserID* | A BSTR that contains a user name for the pending connection. |
| *Password* | A BSTR that contains a password for the pending connection. |
| *Options* | A Long value that indicates how the provider should evaluate the *ConnectionString*. Your only option is *adAsyncOpen*. |
| *adStatus* | *EventStatusEnum*. When this event is called, this parameter is set to *adStatusOK* by default. It is set to *adStatusCantDeny* if the event cannot request cancellation of the pending operation.Before this event returns, set this parameter to *adStatusUnwantedEvent* to prevent subsequent notifications. Set this parameter to *adStatusCancel* to request the connection operation that caused cancellation of this notification. |
| *pConnection* | The **Connection** object for which this event notification applies. Changes to the parameters of the **Connection** by the **WillConnect** event handler will have no effect on the **Connection**. |

#### Remarks

When **WillConnect** is called, the *ConnectionString*, *UserID*, *Password*, and *Options* parameters are set to the values established by the operation that caused this event (the pending connection), and can be changed before the event returns. WillConnect may return a request that the pending connection be canceled.

When this event is canceled, **ConnectComplete** will be called with its *adStatus* parameter set to *adStatusErrorsOccurred*.

---

## ConnectComplete

The **ConnectComplete** event is called after a connection starts.

```
FUNCTION ConnectComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, _
   BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *pError* | An **Error** object. It describes the error that occurred if the value of *EventStatusEnum* is *adStatusErrorsOccurred*; otherwise, it is not set. |
| *adStatus* | *EventStatusEnum*. When **ConnectComplete** is called, this parameter is set to *adStatusCancel* if a **WillConnect** event has requested cancellation of the pending connection. Before either event returns, set this parameter to *adStatusUnwantedEvent* to prevent subsequent notifications. However, closing and reopening the **Connection** causes these events to occur again. |
| *pConnection* | The **Connection** object that fired the event. |

---

## Disconnect

Called after a connection ends.

```
FUNCTION Disconnect (BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *adStatus* | An *EventStatusEnum* value indicating the success or failure of the operation that triggered the event. |
| *pConnection* | The **Connection** object that fired the event. |

---

## WillExecute

Called ust before a pending command executes on a connection.

```
FUNCTION WillExecute (BYVAL Source AS Afx_BSTR PTR, BYVAL CursorType AS CursorTypeEnum PTR, _
   BYVAL LockType AS LockTypeEnum PTR, BYVAL Options AS LONG PTR, BYVAL adStatus AS EventStatusEnum PTR, _
   BYVAL pCommand AS Afx_ADOCommand PTR, BYVAL pRecordset AS Afx_ADORecordset PTR, _
   BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *Source* | A BSTR that contains an SQL command or a stored procedure name. |
| *CursorType* | 	A *CursorTypeEnum* that contains the type of cursor for the **Recordset** that will be opened. With this parameter, you can change the cursor to any type during a **Recordset** Open operation. *CursorType* will be ignored for any other operation. |
| *LockType* | A *LockTypeEnum* that contains the lock type for the **Recordset** that will be opened. With this parameter, you can change the lock to any type during a **Recordset** Open operation. LockType will be ignored for any other operation. |
| *Options* | A Long value that indicates options that can be used to execute the command or open the **Recordset**. |
| *adStatus* | *EventStatusEnum*. Before this event returns, set this parameter to *adStatusUnwantedEvent* to prevent subsequent notifications, or *adStatusCancel* to request cancellation of the operation that caused this event. |
| *pCommand* | The **Command** object for which this event notification applies. |
| *pRecordset* | The **Recordset** object for which this event notification applies. |
| *pConnection* | The **Connection** object for which this event notification applies. |

#### Remarks

A **WillExecute** event may occur due to a **Connection.Execute**, **Command.Execute**, or **Recordset.Open** method The *pConnection* parameter should always contain a valid reference to a **Connection** object. If the event is due to **Connection.Execute**, the *pRecordset* and *pCommand* parameters are set to Nothing. If the event is due to Recordset.Open, the pRecordset parameter will reference the Recordset object and the pCommand parameter is set to NULL. If the event is due to **Command.Execute**, the *pCommand* parameter will reference the **Command** object and the *pRecordset* parameter is set to NULL.

**WillExecute** allows you to examine and modify the pending execution parameters. This event may return a request that the pending command be canceled.

---

## ExecuteComplete

Called after a command has finished executing.

```
FUNCTION ExecuteComplete (BYVAL RecordsAffected AS LONG, BYVAL pError AS Afx_ADOError PTR, _
   BYVAL adStatus AS EventStatusEnum PTR, BYVAL pCommand AS Afx_ADOCommand PTR, _
   BYVAL pRecordset AS Afx_ADORecordset PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *RecordsAffected* | A Long value indicating the number of records affected by the command. |
| *pError* | An **Error** object. It describes the error that occurred if the value of *adStatus* is *adStatusErrorsOccurred*; otherwise it is not set. |
| *adStatus* | *EventStatusEnum*. Before this event returns, set this parameter to *adStatusUnwantedEvent* to prevent subsequent notifications. |
| *pCommand* | The **Command** object that was executed. Contains a **Command** object even when calling **Connection.Execute** or **Recordset.Open** without explicitly creating a **Command**, in which cases the **Command** object is created internally by ADO. |
| *pRecordset* | A **Recordset** object that is the result of the executed command. This **Recordset** may be empty. You should never destroy this **Recordset** object from within this event handler. Doing so will result in an Access Violation when ADO tries to access an object that no longer exists. |
| *pConnection* | A **Connection** object. The connection over which the operation was executed. |

#### Remarks

An **ExecuteComplete** event may occur due to the **Connection.Execute**, **Command.Execute**, **Recordset.Open**, **Recordset.Requery**, or **Recordset.NextRecordset** methods.

---

## InfoMessage

Called whenever a warning occurs during a ConnectionEvent operation.

```
FUNCTION InfoMessage (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, _
   BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
```
| Parameter  | Description |
| ---------- | ----------- |
| *pError* | An **Error** object. This parameter contains any errors that are returned. If multiple errors are returned, enumerate the **Errors** collection to find them. |
| *adStatus* | *EventStatusEnum*. Before this event returns, set this parameter to *adStatusUnwantedEvent* to prevent subsequent notifications. |
| *pConnection* | A **Connection** object. The connection for which the warning occurred. For example, warnings can occur when opening a **Connection** object or executing a **Command** on a **Connection**. |

---

## BeginTransComplete

Called after the **BeginTrans** operation.

```
FUNCTION BeginTransComplete (BYVAL TransactionLevel AS LONG, BYVAL pError AS Afx_ADOError PTR, _
   BYVAL adStatus AS EventStatusEnum PTR, BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *TransactionLevel* | A Long value that contains the new transaction level of the **BeginTrans** that caused this event. |
| *pError* | An **Error** object. It describes the error that occurred if the value of *EventStatusEnum* is *adStatusErrorsOccurred*; otherwise, it is not set. |
| *adStatus* | *EventStatusEnum*. Subsequent notifications can be prevented by setting this parameter to *adStatusUnwantedEvent* before the event returns. |
| *pConnection* | The **Connection** object for which this event occurred. |

---

## CommitTransComplete

Called after the **CommitTrans** operation.

```
FUNCTION CommitTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, _
   BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pError* | An **Error** object. It describes the error that occurred if the value of *EventStatusEnum* is *adStatusErrorsOccurred*; otherwise, it is not set. |
| *adStatus* | *EventStatusEnum*. Subsequent notifications can be prevented by setting this parameter to *adStatusUnwantedEvent* before the event returns. |
| *pConnection* | The **Connection** object for which this event occurred. |

---

## RollbackTransComplete

Called after the **RollbackTrans** operation.

```
FUNCTION RollbackTransComplete (BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, _
   BYVAL pConnection AS Afx_ADOConnection PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pError* | An **Error** object. It describes the error that occurred if the value of *EventStatusEnum* is *adStatusErrorsOccurred*; otherwise, it is not set. |
| *adStatus* | *EventStatusEnum*. Subsequent notifications can be prevented by setting this parameter to *adStatusUnwantedEvent* before the event returns. |
| *pConnection* | The **Connection** object for which this event occurred. |

---

## WillChangeField

Called before a pending operation changes the value of one or more **Field** objects in the **Recordset**.

```
FUNCTION WillChangeField (BYVAL cFields AS LONG, BYVAL Fields AS VARIANT, _
   BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cFields* | A Long that indicates the number of **Field** objects in **Fields**. |
| *Fields* | For **WillChangeField**, the **Fields** parameter is an array of Variants that contains **Field** objects with the original values. |
| *adStatus* |	*EventStatusEnum*. When **WillChangeField** is called, this parameter is set to *adStatusOK* if the operation that caused the event was successful. It is set to *adStatusCantDeny* if this event cannot request cancellation of the pending operation. Before **WillChangeField** returns, set this parameter to *adStatusCancel* to request cancellation of the pending operation. |
| *pRecordset* | A **Recordset** object. The **Recordset** for which this event occurred. |

#### Remarks

A **WillChangeField** or **FieldChangeComplete** event may occur when setting the Value property and calling the **Update** method with field and value array parameters.

---

## FieldChangeComplete 

Called after the value of one or more **Field** objects has changed.

```
FUNCTION FieldChangeComplete (BYVAL cFields AS LONG, BYVAL Fields AS VARIANT, _
   BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, _
   BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cFields* | A Long that indicates the number of **Field** objects in **Fields**. |
| *Fields* | For **FieldChangeComplete**, the **Fields** parameter is an array of Variants that contains **Field** objects with the changed values. |
| *pError* | An **Error** object. It describes the error that occurred if the value of *adStatus* is *adStatusErrorsOccurred*; otherwise it is not set. |
| *adStatus* |	*EventStatusEnum*. When **FieldChangeComplete** is called, this parameter is set to *adStatusOK* if the operation that caused the event was successful, or to *adStatusErrorsOccurred* if the operation failed. Before **FieldChangeComplete** returns, set this parameter to *adStatusUnwantedEvent* to prevent subsequent notifications. |
| *pRecordset* | A **Recordset** object. The **Recordset** for which this event occurred. |

#### Remarks

A **WillChangeField** or **FieldChangeComplete** event may occur when setting the Value property and calling the **Update** method with field and value array parameters.

---

## WillChangeRecord

Called before one or more records (rows) in the **Recordset** change.

```
FUNCTION WillChangeRecord (BYVAL adReason AS EventReasonEnum, BYVAL cRecords AS LONG, _
   BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *adReason* | An *EventReasonEnum* value that specifies the reason for this event. Its value can be *adRsnAddNew*, *adRsnDelete*, *adRsnUpdate*, *adRsnUndoUpdate*, *adRsnUndoAddNew*, *adRsnUndoDelete*, or *adRsnFirstChange*. |
| *cRecords* | A Long value that indicates the number of records changing (affected). |
| *adStatus* | *EventStatusEnum*. When **WillChangeRecord** is called, this parameter is set to *adStatusOK* if the operation that caused the event was successful. It is set to *adStatusCantDeny* if this event cannot request cancellation of the pending operation. Before **WillChangeRecord** returns, set this parameter to *adStatusCancel* to request cancellation of the operation that caused this event or set this parameter to *adStatusUnwantedEvent* to prevent subsequent notications. |
| *pRecordset* | A **Recordset** object. The **Recordset** for which this event occurred. |

#### Remarks

A **WillChangeRecord** or **RecordChangeComplete** event may occur for the first changed field in a row due to the following **Recordset** operations: **Update**, **Delete**, **CancelUpdate**, **AddNew**, **UpdateBatch**, and **CancelBatch**. The value of the **Recordset** *CursorType* determines which operations cause the events to occur.

During the **WillChangeRecord** event, the **Recordset** **Filter** property is set to *adFilterAffectedRecords*. You cannot change this property while processing the event.

You must set the *adStatus* parameter to *adStatusUnwantedEvent* for each possible *adReason* value in order to completely stop event noticiation for any event that includes an *adReason* parameter.

---

## RecordChangeComplete

Called after one or more records change.

```
FUNCTION RecordChangeComplete (BYVAL adReason AS EventReasonEnum, BYVAL cRecords AS LONG, _
   BYVAL pError AS Afx_ADOError PTR, BYVAL adStatus AS EventStatusEnum PTR, _
   BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *adReason* | An *EventReasonEnum* value that specifies the reason for this event. Its value can be *adRsnAddNew*, *adRsnDelete*, *adRsnUpdate*, *adRsnUndoUpdate*, *adRsnUndoAddNew*, *adRsnUndoDelete*, or *adRsnFirstChange*. |
| *cRecords* | A Long value that indicates the number of records changing (affected). |
| *pError* | An **Error** object. It describes the error that occurred if the value of *adStatus* is *adStatusErrorsOccurred*; otherwise it is not set. |
| *adStatus* | *EventStatusEnum*. When **RecordChangeComplete** is called, this parameter is set to *adStatusOK* if the operation that caused the event was successful, or to *adStatusErrorsOccurred* if the operation failed. Before **RecordChangeComplete** returns, set this parameter to *adStatusUnwantedEvent* to prevent subsequent notifications. |
| *pRecordset* | A **Recordset** object. The **Recordset** for which this event occurred. |

#### Remarks

A **WillChangeRecord** or **RecordChangeComplete** event may occur for the first changed field in a row due to the following **Recordset** operations: **Update**, **Delete**, **CancelUpdate**, **AddNew**, **UpdateBatch**, and **CancelBatch**. The value of the **Recordset** *CursorType* determines which operations cause the events to occur.

During the **WillChangeRecord** event, the **Recordset** **Filter** property is set to *adFilterAffectedRecords*. You cannot change this property while processing the event.

You must set the *adStatus* parameter to *adStatusUnwantedEvent* for each possible *adReason* value in order to completely stop event noticiation for any event that includes an *adReason* parameter.

---

## WillChangeRecordset

Called before a pending operation changes the **Recordset**.

```
FUNCTION WillChangeRecordset (BYVAL adReason AS EventReasonEnum, _
   BYVAL adStatus AS EventStatusEnum PTR, BYVAL pRecordset AS Afx_ADORecordset PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *adReason* | An *EventReasonEnum* value that specifies the reason for this event. Its value can be *adRsnRequery*, *adRsnResynch*, *adRsnClose*, *adRsnOpen*. |
| *adStatus* | *EventStatusEnum*. When **WillChangeRecordset** is called, this parameter is set to *adStatusOK* if the operation that caused the event was successful. It is set to *adStatusCantDeny* if this event cannot request cancellation of the pending operation. Before **WillChangeRecordset** returns, set this parameter to *adStatusCancel* to request cancellation of the pending operation or set this parameter to *adStatusUnwantedEvent* to prevent subsequent notifications. |
| *pRecordset* | A **Recordset** object. The **Recordset** for which this event occurred. |

#### Remarks

A **WillChangeRecordset** or **RecordsetChangeComplete** event may occur due to the **Recordset** **Requery** or **Open** methods.

If the provider does not support bookmarks, a **RecordsetChange** event notification occurs each time new rows are retrieved from the provider. The frequency of this event depends on the **RecordsetCacheSize** property.

You must set the *adStatus* parameter to *adStatusUnwantedEvent* for each possible *adReason* value in order to completely stop event notification for any event that includes an *adReason* parameter.

---
