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
