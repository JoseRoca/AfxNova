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

Implementations should return **S_OK** to indicate success or an **HRESULT** error code.

#### Remarks

**OnFolderChange** is called when the dialog is opened.

---
