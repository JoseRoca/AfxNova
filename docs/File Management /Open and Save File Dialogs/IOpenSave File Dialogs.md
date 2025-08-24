# IOpenFileDialog / IOpenFileDialog

'IOpenFileDialog`Wraps the **IFileOpenDialog** interface. Extends the **IFileDialog** interface by adding methods specific to the open dialog.

`IOpenFileDialog`Wraps the **IFileSaveDialog** interface. Extends the **IFileDialog** interface by adding methods specific to the save dialog, which include those that provide support for the collection of metadata to be persisted with the file.

Both also wrap the **IFileDialogEvents** interface, which exposes methods that allow notification of events within a common file dialog, and the **IFileDialogCustomize** interface, which exposes methods that allow an application to add controls to a common file dialog.

**Include file**: AfxNova/OpenSaveFile.inc

## Constructors

| Name       | Description |
| ---------- | ----------- |
| Constructors (IOpenFileDialog) | Creates a dialog box that lets the user specify the drive, directory, and the name of a file or set of files to be opened. |
| Constructors (ISaveFileDialog) | Creates a dialog box that lets the user specify the drive, directory, and name of a file to save. |

## Constructors (IOpenFileDialog/ISaveFileDialog)

```
CONSTRUCTOR
CONSTRUCTOR (BYVAL xPos AS LONG, BYVAL yPos AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *xPos*/*yPos* | Coordinates to position the dialog relative to the client area of the parent. |

## Methods of the IOpenFileDialog and ISaveFileDialog classes

| Name       | Description |
| ---------- | ----------- |
| [AddFileType](#addfiletype) | Adds a file type and pattern to the table. |
| [ShowOpen](#showopen) | Displays the open file dialog. |
| [ShowSave](#showsave) | Displays the save file dialog. |

## Common mehods inherited from **IFileDialog**

| Name       | Description |
| ---------- | ----------- |
| [AddPlace](#addplace) | Adds a folder to the list of places available for the user to open or save items. |
| [Advise](#advise) | Assigns an event handler that listens for events coming from the dialog. |
| [ClearClientData](#clearclientdata) | Instructs the dialog to clear all persisted state information. |
| [Close](#close) | Closes the dialog. |
| [GetCurrentSelection](#getcurrentselection) | Gets the user's current selection in the dialog. |
| [GetFileName](#getfilename) | Retrieves the text currently entered in the dialog's File name edit box. |
| [GetFileTypeIndex](#getfilename) | Gets the currently selected file type. |
| [GetFolder](#getfolder) | Gets either the folder currently selected in the dialog, or, if the dialog is not currently displayed, the folder that is to be selected when the dialog is opened. |
| [GetOptions](#getoptions) | Gets the current flags that are set to control dialog behavior. |
| [GetResult](#getresult) | Gets the choice that the user made in the dialog. |
| [SetClientGuid](#setclientguid) | Enables a calling application to associate a GUID with a dialog's persisted state. |
| [SetDefaultExtension](#setdefaultextension) | Sets the default extension to be added to file names. |
| [SetDefaultFolder](#setdefaultfolder) | Sets the folder used as a default if there is not a recently used folder value available. |
| [SetFileName](#setfilename) | Sets the file name that appears in the File name edit box when that dialog box is opened. |
| [SetFileNameLabel](#setfilenamelabel) | Sets the text of the label next to the file name edit box. |
| [SetFileTypeIndex](#setfiletypeindex) | Sets the file type that appears as selected in the dialog. |
| [SetFileTypes](#setfiletypes) | Sets the file types that the dialog can open or save. |
| [SetFolder](#SetFolder) | Sets a folder that is always selected when the dialog is opened, regardless of previous user action. |
| [SetOkButtonLabel](#setokBbttonlabel) | Sets the text of the Open or Save button. |
| [SetOptions](#setoptions) | Sets flags to control the behavior of the dialog. |
| [SetTitle](#settitle) | Sets the title of the dialog. |
| [Unadvise](#unadvise) | Removes an event handler that was attached through the IFileDialog::Advise method. |

