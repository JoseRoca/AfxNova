The `CIFileDialogEents` class implements the `IFileDialogEvents`interface, that exposes methods that allow notification of events within a common file dialog.

---

## Methods

| Name       | Description |
| ---------- | ----------- |
| [OnFileOk](#onfileok) | Called just before the dialog is about to return with a result. |
| [OnFolderChange](#onfolderchange) | Called when the user navigates to a new folder. |
| [OnFolderChanging](#onfolderchangig) | Called before IFileDialogEvents::OnFolderChange. This allows the implementer to stop navigation to a particular location. |
| [OnOverwrite](#onoverwrite) | Called from the save dialog when the user chooses to overwrite a file. |
| [OnSelectionChange](#onselectionchange) | Called when the user changes the selection in the dialog's view. |
| [OnShareViolation](#onshareviolation) | Enables an application to respond to sharing violations that arise from Open or Save operations. |
| [OnTypeChange](#ontypechange) | Called when the dialog is opened to notify the application of the initial chosen filetype. |

---

## OnFileOk

Called just before the dialog is about to return with a result.

```
FUNCTION OnFileOk (BYVAL pfd AS IFileDialog PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfd* | A pointer to the interface that represents the dialog. |

Implementations should return **S_OK** to accept the current result in the dialog or **S_FALSE** to refuse it. In the case of **S_FALSE**, the dialog should remain open.

#### Remarks

When this method is called, the **GetResult** and **GetResults** methods can be called.

The application can use this callback method to perform additional validation before the dialog closes, or to prevent the dialog from closing. If the application prevents the dialog from closing, it should display a UI to indicate a cause. To obtain a parent HWND for the UI, obtain the **IOleWindow** interface through **IFileDialog.QueryInterface** and call **IOleWindow.GetWindow**.

An application can also use this method to perform all of its work surrounding the opening or saving of files.

---

## OnFolderChange

Called when the user navigates to a new folder.

```
FUNCTION OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfd* | A pointer to the interface that represents the dialog. |

#### Return value

Returns **S_OK** if successful, or an error value otherwise.

#### Remarks

**OnFolderChange** is called when the dialog is opened.

---

## OnFolderChanging

Called before **OnFolderChange**. This allows the implementer to stop navigation to a particular location.

```
FUNCTION OnFolderChanging (BYVAL pfd AS IFileDialog PTR, BYVAL psiFolder AS IShellItem PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfd* | A pointer to the interface that represents the dialog. |
| *psiFolder* | A pointer to an interface that represents the folder to which the dialog is about to navigate. |

#### Return value

Returns **S_OK** if successful, or an error value otherwise. A return value of **S_OK** or **E_NOTIMPL** indicates that the folder change can proceed.

---

## OnOverwrite

Called from the save dialog when the user chooses to overwrite a file.

```
FUNCTION OnOverwrite (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, _
   BYVAL pResponse AS FDE_OVERWRITE_RESPONSE PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfd* | A pointer to the interface that represents the dialog. |
| *psi* | A pointer to the interface that represents the item that will be overwritten. |
| *pResponse* | A pointer to a value from the **FDE_OVERWRITE_RESPONSE** enumeration indicating the response to the potential overwrite action. |

**FDE_OVERWRITE_RESPONSE enumeration**

| Flag  | Value | Description |
| ----- | ----------- |
| **FDEOR_DEFAULT** | 0 | The application has not handled the event. The dialog displays a UI asking the user whether the file should be overwritten and returned from the dialog. |
| **FDEOR_ACCEPT** | 1 | The application has determined that the file should be returned from the dialog. |
| **FDEOR_REFUSE** | 2 | The application has determined that the file should not be returned from the dialog. |

#### Return value

The implementer should return **E_NOTIMPL** if this method is not implemented; **S_OK** or an appropriate error code otherwise.

#### Remarks

The **FOS_OVERWRITEPROMPT** flag must be set through **IFileDialog.SetOptions** before this method is called.

---

## OnSelectionChange

Called when the user changes the selection in the dialog's view.

```
FUNCTION OnSelectionChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfd* | A pointer to the interface that represents the dialog. |

#### Return value

If this method succeeds, it returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

---

## OnShareViolation

Enables an application to respond to sharing violations that arise from Open or Save operations.

```
FUNCTION OnShareViolation (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, _
   BYVAL pResponse AS FDE_SHAREVIOLATION_RESPONSE PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfd* | A pointer to the interface that represents the dialog. |
| *psi* | A pointer to the interface that represents the item that has the sharing violation. |
| *pResponse* | A pointer to a value from the **FDE_SHAREVIOLATION_RESPONSE** enumeration indicating the response to the sharing violation. |

**FDE_SHAREVIOLATION_RESPONSE enumeration**

| Flag  | Value | Description |
| ----- | ----------- |
| **FDESVR_DEFAULT** | 0 | The application has not handled the event. The dialog displays a UI that indicates that the file is in use and a different file must be chosen. |
| **FDESVR_ACCEPT** | 1 | The application has determined that the file should be returned from the dialog. |
| **FDESVR_REFUSE** | 2 | The application has determined that the file should not be returned from the dialog. |

#### Return value

The implementer should return **E_NOTIMPL** if this method is not implemented;** S_OK** or an appropriate error code otherwise.

#### Remarks

The **FOS_SHAREAWARE** flag must be set through **IFileDialog.SetOptions** before this method is called.

A sharing violation could possibly arise when the application attempts to open a file, because the file could have been locked between the time that the dialog tested it and the application opened it.

---

## OnTypeChange

Called when the dialog is opened to notify the application of the initial chosen filetype.

```
FUNCTION OnTypeChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfd* | A pointer to the interface that represents the dialog. |

#### Return value

Returns **S_OK** if successful, or an error value otherwise.

---
