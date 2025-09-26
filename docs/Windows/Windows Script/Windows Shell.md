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

## AfxRegRead

Returns the value of a key or value-name from the registry.

```
FUNCTION AfxRegRead (BYREF wszKey AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszKey* | The registry key name. |

#### Usage examples

```
print AfxRegRead("HKLM\HARDWARE\DESCRIPTION\System\CentralProcessor\0\ProcessorNameString")
print AfxRegRead("HKLM\HARDWARE\DESCRIPTION\System\CentralProcessor\0\~Mhz")
```
---

## AfxRegWrite

Creates a new key, adds another value-name to an existing key (and assigns it a value), or changes the value of an existing value-name.

```
FUNCTION AfxRegWrite (BYREF wszKey AS WSTRING, BYREF dvValue AS DVARIANT, BYREF wszType AS WSTRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszKey* | String value indicating the key-name, value-name, or value you want to create, add, or change. |
| *dvValue* | The name of the new key you want to create, the name of the value you want to add to an existing key, or the new value you want to assign to an existing value-name. |
| *wszType* | String value indicating the value's data type. |

#### Usage examples

```
AfxRegWrite("HKCU\Software\ACME\FortuneTeller\", 1, "REG_BINARY")
AfxRegWrite("HKCU\Software\ACME\FortuneTeller\MindReader"", "Goocher!", "REG_SZ")
```
---

## AfxRegDelete

Deletes a key or one of its values from the registry.

```
FUNCTION AfxRegDelete (BYREF wszKey AS WSTRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszKey* | String value indicating the name of the registry key or key value you want to delete. |

#### Usage examples

```
AfxRegDelete("HKCU\Software\ACME\FortuneTeller\MindReader")
AfxRegDelete("HKCU\Software\ACME\FortuneTeller\")
AfxRegDelete("HKCU\Software\ACME\")
```

#### Return value

Returns S_OK or an HRESULT error code.

---

## AfxLogEvent

Adds an event entry to a log file.

```
FUNCTION AfxLogEvent (BYVAL nType AS LONG, BYREF wszMessage AS WSTRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *nType* | Integer value representing the event type. |
| *wszMessage* | String value containing the log entry text. |

There are six event types.

| Type  | Value |
| ----- | ----- |
| 0 | SUCCESS |
| 1 | ERROR |
| 2 | WARNING |
| 4 | INFORMATION |
| 8 | AUDIT_SUCCESS |
| 16 | AUDIT_FAILURE |

#### Usage example

```
AfxLogEvent(0, "Logon Script Completed Successfully")
```
---

## AfxRun

Runs a program in a new process.

```
AfxRun (BYREF wszCommand AS WSTRING, BYVAL WindowStyle AS LONG = 0, BYVAL WaitOnReturn AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszCommand* | String value indicating the command line you want to run. You must include any parameters you want to pass to the executable file. |
| *WindowStyle* | Integer value indicating the appearance of the program's window. Note that not all programs make use of this information. |
| *WaitOnReturn* | Boolean value indicating whether the script should wait for the program to finish executing before continuing to the next statement in your script. If set to true, script execution halts until the program finishes, and **Run** returns any error code returned by the program. If set to false (the default), the **Run** method returns immediately after starting the program, automatically returning 0 (not to be interpreted as an error code). |

#### Return value

Returns S_OK or an HRESULT error code.

#### Usage example

```
AfxRun("Calc")
```
---

## AfxSendKeys

Sends keystrokes to the active window.

```
FUNCTION AfxSendKeys (BYREF wszKeys AS WSTRING, BYVAL bWait AS BOOLEAN = FALSE) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszKeys* | Integer value representing the event type. |
| *bWait* | Boolean value: True or False. |

#### Return value

Returns S_OK or an HRESULT error code.

#### Usage example

```
AfxRun("calc")
AfxForceSetForegroundWindow(FindWindow("Windows.UI.Core.CoreWindow", NULL))
' ---or-- AfxForceSetForegroundWindow(FindWindow(NULL, "calc"))
SLEEP 1000
AfxSendKeys("1{+}")
SLEEP 500
AfxSendKeys("2")
SLEEP 500
AfxSendKeys("~")
SLEEP 500
AfxSendKeys("*3")
SLEEP 500
AfxSendKeys("~")
SLEEP 2500
```

#### Remarks

Use the **SendKeys** method to send keystrokes to applications that have no automation interface. Most keyboard characters are represented by a single keystroke. Some keyboard characters are made up of combinations of keystrokes (CTRL+SHIFT+HOME, for example). To send a single keyboard character, send the character itself as the string argument. For example, to send the letter x, send the string argument "x".

**Note**: To send a space, send the string " ".

You can use **SendKeys** to send more than one keystroke at a time. To do this, create a compound string argument that represents a sequence of keystrokes by appending each keystroke in the sequence to the one before it. For example, to send the keystrokes a, b, and c, you would send the string argument "abc". The **SendKeys** method uses some characters as modifiers of characters (instead of using their face-values). This set of special characters consists of parentheses, brackets, braces, and the:

•	plus sign      "+",
•	caret          "^",
•	percent sign "%",
•	and tilde      "~"

Send these characters by enclosing them within braces "{}". For example, to send the plus sign, send the string argument "{+}". Brackets "[ ]" have no special meaning when used with **SendKeys**, but you must enclose them within braces to accommodate applications that do give them a special meaning (for dynamic data exchange (DDE) for example).

•	To send bracket characters, send the string argument "{[}" for the left bracket and "{]}" for the right one.
•	To send brace characters, send the string argument "{{}" for the left brace and "{}}" for the right one.

Some keystrokes do not generate characters (such as ENTER and TAB). Some keystrokes represent actions (such as BACKSPACE and BREAK). To send these kinds of keystrokes, send the arguments shown in the following table:

| Key  | Argument |
| ---- | -------- |
| BACKSPACE | {BACKSPACE}, {BS}, or {BKSP} |
| BREAK | {BREAK} |
| CAPS LOCK | {CAPSLOCK} |
| DEL or DELETE | {DELETE} or {DEL} |
| DOWN ARROW | {DOWN} |
| END | {END} |
| ENTER | {ENTER} or ~ |
| ESC | {ESC} |
| HELP | {HELP} |
| HOME | {HOME} |
| INS or INSERT | {INSERT} or {INS} |
| LEFT ARROW | {LEFT} |
| NUM LOCK | {NUMLOCK} |
| PAGE DOWN | {PGDN} |
| PAGE UP | {PGUP} |
| PRINT SCREEN | {PRTSC} |
| RIGHT ARROW | {RIGHT} |
| SCROLL LOCK | {SCROLLLOCK} |
| TAB | {TAB} |
| UP ARROW | {UP} |
| F1 | {F1} |
| F2 | {F2} |
| F3 | {F3} |
| F4 | {F4} |
| F5 | {F5} |
| F6 | {F6} |
| F7 | {F7} |
| F8 | {F8} |
| F9 | {F9} |
| F10 | {F10} |
| F11 | {F11} |
| F12 | {F12} |
| F13 | {F13} |
| F14 | {F14} |
| F15 | {F15} |
| F16 | {F16} |

To send keyboard characters that are comprised of a regular keystroke in combination with a SHIFT, CTRL, or ALT, create a compound string argument that represents the keystroke combination. You do this by preceding the regular keystroke with one or more of the following special characters:

| Key  | Special Character |
| ---- | -------- |
| SHIFT | + |
| CTRL | ^ |
| ALT | % |

**Note**: When used this way, these special characters are not enclosed within a set of braces.

To specify that a combination of SHIFT, CTRL, and ALT should be held down while several other keys are pressed, create a compound string argument with the modified keystrokes enclosed in parentheses. For example, to send the keystroke combination that specifies that the SHIFT key is held down while:

•	e and c are pressed, send the string argument "+(ec)".
•	e is pressed, followed by a lone c (with no SHIFT), send the string argument "+ec".
 

You can use the SendKeys method to send a pattern of keystrokes that consists of a single keystroke pressed several times in a row. To do this, create a compound string argument that specifies the keystroke you want to repeat, followed by the number of times you want it repeated. You do this using a compound string argument of the form {keystroke number}. For example, to send the letter "x" ten times, you would send the string argument "{x 10}". Be sure to include a space between keystroke and number.

**Note**: The only keystroke pattern you can send is the kind that is comprised of a single keystroke pressed several times. For example, you can send "x" ten times, but you cannot do the same for "Ctrl+x".

**Note**: You cannot send the PRINT SCREEN key {PRTSC} to an application.

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

## CurrentDirectory

Retrieves or changes the current active directory.

```
[GET] PROPERTY CurrentDirectory () AS DWSTRING
[SET] PROPERTY CurrentDirectory (BYREF wszDir AS WSTRING) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszDir* | The directory to set. |

The [GET] **CurrentDirectory** returns a string that contains the fully qualified path of the current working directory of the active process.

The [SET] **CurrentDirectory** sets the current working directory of the active process.

---

## GetStatus

Provides status information about a script run with the Exec method.

```
FUNCTION GetStatus () AS WshExecStatus
```

#### Return value

The **Status** property returns a value from the *WshExecStatus* enumerated type. WshRunning = 0, WshFinished = 1, WshFailed = 2.

---

## GetStdIn

Exposes the stdin input stream of the **IWshExec** interface.

```
FUNCTION GetStdIn () AS Afx_ITextStream PTR
```

#### Return value

A pointer to the **ITextStream** interface.

---

## GetStdOut

Exposes the stdin output stream of the **IWshExec** interface.

```
FUNCTION GetStdOut () AS Afx_ITextStream PTR
```

#### Return value

A pointer to the **ITextStream** interface.

---

## GetStdErr

Provides access to the stderr output stream of the **IWshExec** interface.

```
FUNCTION GetStdErr () AS Afx_ITextStream PTR
```

#### Return value

A pointer to the **ITextStream** interface.

---

## GetExitCode

Returns the exit code set by a script or program run using the **Exec** method (called by the constructor).

```
FUNCTION GetExitCode () AS LONG
```

#### Remarks

Executables set an exit code when they finish running. This conveys the status information when a process ends. Often, it is used to send an error code (or some other piece of information) back to the caller. If the process has not finished, the **GetExitCode** property returns 0. The values returned from ^^GetExitCode** depend on the application that was called.

---

## GetProcessID

Returns the process ID (PID) for a process started with the **Exec** method (called by the constructor).

```
FUNCTION GetProcessID () AS LONG
```

#### Remarks

You can use the **GetProcessID** property to activate an application (as an argument to the **AppActivate** method).

---

## Terminate

Instructs the script engine to end the process started by the **Exec** method (called by the constructor).

```
FUNCTION Terminate
```

#### Remarks

The **Terminate** method does not return a value. Use the **Terminate** method only as a last resort since some applications do not clean up properly. As a general rule, let the process run its course and end on its own. The **Terminate** method attempts to end a process using the **WM_CLOSE** message. If that does not work, it kills the process immediately without going through the normal shutdown procedure.

---
