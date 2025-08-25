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
