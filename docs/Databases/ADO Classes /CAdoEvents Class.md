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
| *adStatus* | An *EventStatusEnum* value indicating the success or failure of the operation that triggered the event.  |
| *pConnection* | The **Connection** object that fired the event. |

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
