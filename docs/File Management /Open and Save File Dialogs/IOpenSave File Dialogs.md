# IOpenFileDialog / IOpenFileDialog

'IOpenFileDialog`Wraps the **IFileOpenDialog** interface. Extends the **IFileDialog** interface by adding methods specific to the open dialog.

`IOpenFileDialog`Wraps the **IFileSaveDialog** interface. Extends the **IFileDialog** interface by adding methods specific to the save dialog, which include those that provide support for the collection of metadata to be persisted with the file.

Both also wrap the **IFileDialogEvents** interface, which exposes methods that allow notification of events within a common file dialog, and the **IFileDialogCustomize** interface, which exposes methods that allow an application to add controls to a common file dialog.

**Include file**: AfxNova/OpenSaveFile.inc

---

## Constructors (IOpenFileDialog/ISaveFileDialog)

| Name       | Description |
| ---------- | ----------- |
| Constructor (IOpenFileDialog) | Creates a dialog box that lets the user specify the drive, directory, and the name of a file or set of files to be opened. |
| Constructor (ISaveFileDialog) | Creates a dialog box that lets the user specify the drive, directory, and name of a file to save. |

#### Usage

```
DIM pofd AS IOpenFileDialog
DIM psfd AS ISaveFileDialog
```
---

## Methods of the IOpenFileDialog and ISaveFileDialog classes

| Name       | Description |
| ---------- | ----------- |
| [AddFileType](#addfiletype) | Adds a file type and pattern to the table. |
| [ShowOpen](#showopen) | Displays the open file dialog. |
| [ShowSave](#showsave) | Displays the save file dialog. |

---

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

---

## Advise

Assigns an event handler that listens for events coming from the dialog.

```
FUNCTION AddPlace (BYVAL psi AS IShellItem PTR, BYVAL fdap AS FDAP) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *psi* | A pointer to an IShellItem that represents the folder to be made available to the user. This can only be a folder. |
| *fdap* | Specifies where the folder is placed within the list. See **FDAP** enumeration. |

**FDAP enumeration**

FDAP_BOTTOM = 0. The place is added to the bottom of the default list.
FDAP_TOP = 1. The place is added to the top of the default list.

---

## AddPlace

Adds a folder to the list of places available for the user to open or save items.

```
FUNCTION Advise (BYVAL pfde as IFileDialogEvents PTR) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfde* | A pointer to an **IFileDialogEvents** implementation that will receive events from the dialog. |

#### Return value

A DWORD that identifies this event handler

---

## ClearClientData 

Instructs the dialog to clear all persisted state information.

```
FUNCTION ClearClientData () AS HRESULT
```

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

Persisted information can be associated with an application or a GUID. If a GUID was set by using IFileDialog::SetClientGuid, that GUID is used to clear persisted information.

---

## Close 

Closes the dialog.

```
FUNCTION Close (BYVAL hr AS HRESULT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hr* | The code that will be returned by Show to indicate that the dialog was closed before a selection was made. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

An application can call this method from a callback method or function while the dialog is open. The dialog will close and the Show method will return with the HRESULT specified in hr.

If this method is called, there is no result available for the **GetResult** or **GetResults** methods, and they will fail if called.

---
