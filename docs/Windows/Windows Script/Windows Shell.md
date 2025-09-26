# Windows Shell

Provides functions to read, write and delete registry keys, log events, send keystrokes to the active window, run a program in a new process, and execute programs or commands in a new process, providing methods to interact with the provess's standard input, output, and error streams.

### Functions

| Name       | Description |
| ---------- | ----------- |
| [AfxLogEvent](#afxlogevent) | Adds an event entry to a log file. |
| [AfxRegDelete](#afxregdelete) | Deletes a key or one of its values from the registry. |
| [AfxRegRead](#afxregread) | Returns the value of a key or value-name from the registry. |
| [AfxRegWrite](#afxregwrite) | Creates a new key, adds another value-name to an existing key (and assigns it a value), or changes the value of an existing value-name. |
| [AfxRun](#afxrun) | Runs a program in a new process. |
| [AfxSendKeys](#afxsendkeys) | Sends keystrokes to the active window. |

### CWhsExec class

Executes a program or command in a new process, providing methods to interact with the provess's standard input, output, and error streams.

| Name       | Description |
| ---------- | ----------- |
| [CONSTRUCTOR](#constructor) | Creates an instance of the `CWhsExec` class. |
| [CAST](#cast) | Returns a pointer to the **IWshExec** interface. |
| [GetLastResult](#getlastresult) | Returns the last result code. |
| [GetErrorInfo](#geterrorinfo) | Returns a description of the last result code. |

---

| Name       | Description |
| ---------- | ----------- |
| [CurrentDirectory](#currentdirectory) | Retrieves or changes the current active directory. |
| [GetExitCode](#getexitcode) | Returns the exit code set by a script or program run using the Exec method. |
| [GetProcessID](#getprocessid) | The process ID (PID) for a process started with the Exec method. |
| [GetStatus](#getstatus) | Provides status information about a script run with the Exec method. |
| [GetStdErr](#getstderr) | Provides access to the stderr output stream of the IWshExec interface. |
| [GetStdIn](#getstdin) | Exposes the stdin input stream of the IWshExec interface. |
| [GetStdOut](#getstdout) | Exposes the stdout output stream of the IWshExec interface. |
| [Terminate](#terminate) | Instructs the script engine to end the process started by the Exec method. |

---

## Constructor

Creates an instance of the `CWhsExec` class.

Runs an application in a child command-shell, providing access to the StdIn/StdOut/StdErr streams.

```
CONSTRUCTOR CWhsExec (BYREF wszCommand AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCommand* | The command to execute. |

#### Usage example:

```
#include "AfxNova/AfxWsh.inc"
#include "AfxNova/CTextStream.inc"
USING AfxNova
DIM pWshExec AS CWhsExec = "%comspec% /c dir"
IF pWshExec THEN
   DIM pStdOutStm AS CTextStream = pWshExec.GetStdOut
   IF pStdOutStm THEN
      IF pStdOutStm.EOS = FALSE THEN
         ' // The console uses the CP_OEMCP code page
         DIM dwsOut AS DWSTRING
         dwsOut.OEM = pStdOutStm.ReadAll
         ' // You can also use
         ' DIM dwsOut AS DWSTRING = DWSTRING(pStdOutStm.ReadAll, CP_OEMCP)
         print dwsOut
      END IF
   END IF
END IF
```
---

## CAST

Returns a pointer to the **IWshExec** interface.

```
OPERATOR CAST () AS Afx_IWshExec PTR
```
---

## GetLastResult

Returns the result code returned by the last executed method.

```
FUNCTION GetLastResult () AS HRESULT
```
---

## GetErrorInfo

Returns a description of the last result code.
```
FUNCTION GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
```
---
