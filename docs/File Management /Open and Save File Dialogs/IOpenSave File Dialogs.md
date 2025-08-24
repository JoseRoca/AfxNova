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
| *psi* | A pointer to an **IShellItem** that represents the folder to be made available to the user. This can only be a folder. |
| *fdap* | Specifies where the folder is placed within the list. See **FDAP** enumeration. |

**FDAP enumeration**

| Constant | Value  | Description |
| -------- | ------ | ----------- |
| **FDAP_BOTTOM** | 0 | The place is added to the bottom of the default list. |
| **FDAP_TOP** | 1 | The place is added to the top of the default list. |

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

## GetCurrentSelection 

Gets the folder used as a default if there is not a recently used folder value available.

```
FUNCTION GetCurrentSelection (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *sigdnName* | Requests the form of an item's display name to retrieve . |

**SIGDN Enumeration**

| Constant | Description |
| -------- | ----------- |
| **SIGDN_NORMALDISPLAY** | Returns the display name relative to the parent folder. In UI this name is generally ideal for display to the user. |
| **SIGDN_PARENTRELATIVEPARSING** | Returns the parsing name relative to the parent folder. This name is not suitable for use in UI. |
| **SIGDN_DESKTOPABSOLUTEPARSING** | Returns the parsing name relative to the desktop. This name is not suitable for use in UI. |
| **SIGDN_PARENTRELATIVEEDITING** | Returns the editing name relative to the parent folder. In UI this name is suitable for display to the user. |
| **SIGDN_DESKTOPABSOLUTEEDITING** | Returns the editing name relative to the desktop. In UI this name is suitable for display to the user. |
| **SIGDN_FILESYSPATH** | Returns the item's file system path, if it has one. Only items that report SFGAO_FILESYSTEM have a file system path. When an item does not have a file system path, a call to IShellItem::GetDisplayName on that item will fail. In UI this name is suitable for display to the user in some cases, but note that it might not be specified for all items. |
| **SIGDN_URL** | Returns the item's URL, if it has one. Some items do not have a URL, and in those cases a call to IShellItem::GetDisplayName will fail. This name is suitable for display to the user in some cases, but note that it might not be specified for all items. |
| **SIGDN_PARENTRELATIVEFORADDRESSBAR** | Returns the path relative to the parent folder in a friendly format as displayed in an address bar. This name is suitable for display to the user. |
| **SIGDN_PARENTRELATIVE** | Returns the path relative to the parent folder. |
| **SIGDN_PARENTRELATIVEFORUI** | Introduced in Windows 8. |

#### Return value

The folder name.

---

## GetFileName 

Retrieves the text currently entered in the dialog's File name edit box.

```
FUNCTION GetFileName () AS DWSTRING
```

#### Return value

The folder name.

#### Remarks

The text in the File name edit box does not necessarily reflect the item the user chose. To get the item the user chose, use **GetResult**.

---

## GetFileTypeIndex 

Gets the file type that appears as selected in the dialog.

```
FUNCTION GetFileTypeIndex () AS UINT
```

#### Return value

Gets the currently selected file type.

#### Remarks

The index of the selected file type in the file type array passed to **SetFileTypes** in its *cFileTypes* parameter.

**Note**  This is a one-based index rather than zero-based.

---

## GetFolder 

Gets either the folder currently selected in the dialog, or, if the dialog is not currently displayed, the folder that is to be selected when the dialog is opened.

```
FUNCTION GetFolder (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *sigdnName* | Requests the form of an item's display name to retrieve . |

**SIGDN Enumeration**

| Constant | Description |
| -------- | ----------- |
| **SIGDN_NORMALDISPLAY** | Returns the display name relative to the parent folder. In UI this name is generally ideal for display to the user. |
| **SIGDN_PARENTRELATIVEPARSING** | Returns the parsing name relative to the parent folder. This name is not suitable for use in UI. |
| **SIGDN_DESKTOPABSOLUTEPARSING** | Returns the parsing name relative to the desktop. This name is not suitable for use in UI. |
| **SIGDN_PARENTRELATIVEEDITING** | Returns the editing name relative to the parent folder. In UI this name is suitable for display to the user. |
| **SIGDN_DESKTOPABSOLUTEEDITING** | Returns the editing name relative to the desktop. In UI this name is suitable for display to the user. |
| **SIGDN_FILESYSPATH** | Returns the item's file system path, if it has one. Only items that report SFGAO_FILESYSTEM have a file system path. When an item does not have a file system path, a call to IShellItem::GetDisplayName on that item will fail. In UI this name is suitable for display to the user in some cases, but note that it might not be specified for all items. |
| **SIGDN_URL** | Returns the item's URL, if it has one. Some items do not have a URL, and in those cases a call to IShellItem::GetDisplayName will fail. This name is suitable for display to the user in some cases, but note that it might not be specified for all items. |
| **SIGDN_PARENTRELATIVEFORADDRESSBAR** | Returns the path relative to the parent folder in a friendly format as displayed in an address bar. This name is suitable for display to the user. |
| **SIGDN_PARENTRELATIVE** | Returns the path relative to the parent folder. |
| **SIGDN_PARENTRELATIVEFORUI** | Introduced in Windows 8. |

#### Return value

The folder name.

---

## GetOptions 

Gets the current flags that are set to control dialog behavior.

```
FUNCTION GetOptions () AS FILEOPENDIALOGOPTIONS
```

#### Return value

A value made up of one or more of the **FILEOPENDIALOGOPTIONS** values.

| Constant | Value  | Description |
| -------- | ------ | ----------- |
| **FOS_OVERWRITEPROMPT** | &h2 | When saving a file, prompt before overwriting an existing file of the same name. This is a default value for the Save dialog. |
| **FOS_STRICTFILETYPES** | &h4 | In the Save dialog, only allow the user to choose a file that has one of the file name extensions specified through **SetFileTypes**. |
| **FOS_NOCHANGEDIR** | &h8 | Don't change the current working directory. |
| **FOS_PICKFOLDERS** | &h20 | Present an Open dialog that offers a choice of folders rather than files. |
| **FOS_FORCEFILESYSTEM** | &h40 | Ensures that returned items are file system items (SFGAO_FILESYSTEM). Note that this does not apply to items returned by **GetCurrentSelection**. |
| **FOS_ALLNONSTORAGEITEMS** | &h80 | Enables the user to choose any item in the Shell namespace, not just those with SFGAO_STREAM or SFAGO_FILESYSTEM attributes. This flag cannot be combined with FOS_FORCEFILESYSTEM. |
| **FOS_NOVALIDATE** | &h100 | Do not check for situations that would prevent an application from opening the selected file, such as sharing violations or access denied errors. |
| **FOS_ALLOWMULTISELECT** | &h200 | Enables the user to select multiple items in the open dialog. Note that when this flag is set, the IFileOpenDialog interface must be used to retrieve those items. |
| **FOS_PATHMUSTEXIST** | &h800 | The item returned must be in an existing folder. This is a default value. |
| **FOS_FILEMUSTEXIST** | &h1000 | The item returned must exist. This is a default value for the Open dialog. |
| **FOS_CREATEPROMPT** | &h2000 | Prompt for creation if the item returned in the open dialog does not exist. Note that this does not actually create the item. |
| **FOS_SHAREAWARE** | &h4000 | In the case of a sharing violation when an application is opening a file, call the application back through OnShareViolation for guidance. This flag is overridden by FOS_NOVALIDATE. |
| **FOS_NOREADONLYRETURN** | &h8000 | Do not return read-only items. This is a default value for the Save dialog. |
| **FOS_NOTESTFILECREATE** | &h10000 | Do not test whether creation of the item as specified in the Save dialog will be successful. If this flag is not set, the calling application must handle errors, such as denial of access, discovered when the item is created. |
| **FOS_HIDEMRUPLACES** | &h20000 | Hide the list of places from which the user has recently opened or saved items. This value is not supported as of Windows 7. |
| **FOS_HIDEPINNEDPLACES** | &h40000 | Hide items shown by default in the view's navigation pane. This flag is often used in conjunction with the **AddPlace** method, to hide standard locations and replace them with custom locations. Windows 7 and later. Hide all of the standard namespace locations (such as Favorites, Libraries, Computer, and Network) shown in the navigation pane. Windows Vista. Hide the contents of the Favorite Links tree in the navigation pane. Note that the category itself is still displayed, but shown as empty. |
| **FOS_NODEREFERENCELINKS** | &h100000 | Shortcuts should not be treated as their target items. This allows an application to open a .lnk file rather than what that file is a shortcut to. |
| **FOS_OKBUTTONNEEDSINTERACTION** | &h200000 | The OK button will be disabled until the user navigates the view or edits the filename (if applicable). Note: Disabling of the OK button does not prevent the dialog from being submitted by the Enter key. |
| **FOS_DONTADDTORECENT** | &h2000000 | Do not add the item being opened or saved to the recent documents list (SHAddToRecentDocs). |
| **FOS_FORCESHOWHIDDEN** | &h10000000 | Include hidden and system items. |
| **FOS_DEFAULTNOMINIMODE** | &h20000000 | Indicates to the Save As dialog box that it should open in expanded mode. Expanded mode is the mode that is set and unset by clicking the button in the lower-left corner of the Save As dialog box that switches between Browse Folders and Hide Folders when clicked. This value is not supported as of Windows 7. |
| **FOS_FORCEPREVIEWPANEON** | &h40000000 | Indicates to the Open dialog box that the preview pane should always be displayed. |
| **FOS_SUPPORTSTREAMABLEITEMS** | &h80000000 | Indicates that the caller is opening a file as a stream (BHID_Stream), so there is no need to download that file. |

---

## GetResult 

Gets the choice that the user made in the dialog.

```
GetResult (BYVAL sigdnName AS SIGDN = SIGDN_NORMALDISPLAY) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *sigdnName* | Requests the form of an item's display name to retrieve . |

**SIGDN Enumeration**

| Constant | Description |
| -------- | ----------- |
| **SIGDN_NORMALDISPLAY** | Returns the display name relative to the parent folder. In UI this name is generally ideal for display to the user. |
| **SIGDN_PARENTRELATIVEPARSING** | Returns the parsing name relative to the parent folder. This name is not suitable for use in UI. |
| **SIGDN_DESKTOPABSOLUTEPARSING** | Returns the parsing name relative to the desktop. This name is not suitable for use in UI. |
| **SIGDN_PARENTRELATIVEEDITING** | Returns the editing name relative to the parent folder. In UI this name is suitable for display to the user. |
| **SIGDN_DESKTOPABSOLUTEEDITING** | Returns the editing name relative to the desktop. In UI this name is suitable for display to the user. |
| **SIGDN_FILESYSPATH** | Returns the item's file system path, if it has one. Only items that report SFGAO_FILESYSTEM have a file system path. When an item does not have a file system path, a call to IShellItem::GetDisplayName on that item will fail. In UI this name is suitable for display to the user in some cases, but note that it might not be specified for all items. |
| **SIGDN_URL** | Returns the item's URL, if it has one. Some items do not have a URL, and in those cases a call to IShellItem::GetDisplayName will fail. This name is suitable for display to the user in some cases, but note that it might not be specified for all items. |
| **SIGDN_PARENTRELATIVEFORADDRESSBAR** | Returns the path relative to the parent folder in a friendly format as displayed in an address bar. This name is suitable for display to the user. |
| **SIGDN_PARENTRELATIVE** | Returns the path relative to the parent folder. |
| **SIGDN_PARENTRELATIVEFORUI** | Introduced in Windows 8. |

#### Return value

The choice that the user made in the dialog.

---

## SetClientGuid 

Enables a calling application to associate a GUID with a dialog's persisted state.

```
FUNCTION SetClientGuid (BYVAL guid AS GUID PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *guid* | The GUID to associate with this dialog state. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

A dialog's state can include factors such as the last visited folder and the position and size of the dialog.

Typically, this state is persisted based on the name of the executable file. By specifying a GUID, an application can have different persisted states for different versions of the dialog within the same application (for example, an import dialog and an open dialog).

**SetClientGuid** should be called immediately after creation of the dialog object.

---

## SetDefaultExtension 

Sets the default extension to be added to file names.

```
FUNCTION SetDefaultExtension (BYVAL pwszDefaultExtension AS WSTRING PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszDefaultExtension* | A pointer to a buffer that contains the extension text. This string should not include a leading period. For example, "jpg" is correct, while ".jpg" is not. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

#### Remarks

If this method is called before showing the dialog, the dialog will update the default extension automatically when the user chooses a new file type (see SetFileTypes).

---

## SetDefaultFolder 

Sets the folder used as a default if there is not a recently used folder value available.

```
FUNCTION SetDefaultFolder (BYVAL psi AS IShellItem PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *psi* | A pointer to the interface that represents the folder. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

---

## SetFileName

Sets the file name that appears in the File name edit box when that dialog box is opened.

```
FUNCTION SetFileName (BYVAL pwszName AS WSTRING PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszName* | A pointer to the name of the file. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

---

## SetFileNameLabel

Sets the file name that appears in the File name edit box when that dialog box is opened.

```
FUNCTION SetFileNameLabel (BYVAL pwszLabel AS WSTRING PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszLabel* | Sets the text of the label next to the file name edit box. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

---

## SetFileTypeIndex

Sets the file type that appears as selected in the dialog.

```
FUNCTION SetFileTypeIndex (BYVAL iFileType AS UINT) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *iFileType* | The index of the file type in the file type array passed to **SetFileTypes** in its *cFileTypes* parameter. Note that this is a one-based index, not zero-based. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

---

## SetFileTypes

Sets the file types that the dialog can open or save.

```
FUNCTION SetFileTypes (BYVAL cFileTypes AS UINT, BYVAL rgFilterSpec AS COMDLG_FILTERSPEC PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *cFileTypes* | The number of elements in the array specified by *rgFilterSpec*. |
| *rgFilterSpec* | A pointer to an array of *COMDLG_FILTERSPEC* structures, each representing a file type. |

#### Return value

If the method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code, including the following:

| Return code  | Description |
| ------------ | ----------- |
| **E_UNEXPECTED** | The FOS_PICKFOLDERS flag was set in the **SetOptions** method. |
| **E_INVALIDARG** | The *rgFilterSpec* parameter is NULL. |

---

## SetFolder

Sets a folder that is always selected when the dialog is opened, regardless of previous user action.

```
FUNCTION SetFolder (BYVAL pwszFolderName AS WSTRING PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszFolderName* | The number of elements in the array specified by *rgFilterSpec*. |

#### Return value

If the method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code, including the following:

#### Remarks

This folder overrides any "most recently used" folder. If this method is called while the dialog is displayed, it causes the dialog to navigate to the specified folder.

In general, we do not recommended the use of this method. If you call **SetFolder** before you display the dialog box, the most recent location that the user saved to or opened from is not shown. Unless there is a very specific reason for this behavior, it is not a good or expected user experience and should therefore be avoided. In almost all instances, **SetDefaultFolder** is the better method.

As of Windows 7, if the path of the folder specified through psi is the default path of a known folder, the known folder's current path is used in the dialog. That path might not be the same as the path specified in psi; for instance, if the known folder has been redirected. If the known folder is a library (virtual folders Documents, Music, Pictures, and Videos), the library's path is used in the dialog. If the specified library is hidden (as they are by default as of Windows 8.1), the library's default save location is used in the dialog, such as the Microsoft OneDrive Documents folder for the Documents library. Because of these mappings, the folder location used in the dialog might not be exactly as you specified when you called this method.

---

## SetOkButtonLabel

Sets the text of the Open or Save button.

```
FUNCTION SetOkButtonLabel (BYVAL pwszText AS WSTRING PTR) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszText* | A pointer to a buffer that contains the button text. |

#### Return value

If this method succeeds, it returns S_OK. Otherwise, it returns an HRESULT error code.

---

