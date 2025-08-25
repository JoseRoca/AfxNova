The `CIFileDialogEents` class implements the `IFileDialogEvents`interface, that exposes methods that allow notification of events within a common file dialog.

**Include file:** AfxNova/CIFileDialogEvents.inc

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

## Template code

| Name       | Description |
| ---------- | ----------- |
| [Template](#template) | Template code that implements the **FileDialogEvents** callback interface. |

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
| ----- | ----- | ----------- |
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
| ----- | ----- | ---------- |
| **FDESVR_DEFAULT** | 0 | The application has not handled the event. The dialog displays a UI that indicates that the file is in use and a different file must be chosen. |
| **FDESVR_ACCEPT** | 1 | The application has determined that the file should be returned from the dialog. |
| **FDESVR_REFUSE** | 2 | The application has determined that the file should not be returned from the dialog. |

#### Return value

The implementer should return **E_NOTIMPL** if this method is not implemented; **S_OK** or an appropriate error code otherwise.

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

## Template

The following template implements the **FileDialogEvents** callback interface. You can copy and modify it to implement your own or inherit from it and override its virtual methods.

```
' ########################################################################################
' CIFileDialogEvents class
' Implementation of the FileDialogEvents callback interface
' ########################################################################################
TYPE CIFileDialogEvents EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION OnFileOk (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnFolderChanging (BYVAL pfd AS IFileDialog PTR, BYVAL psiFolder AS IShellItem PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnSelectionChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnShareViolation (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_SHAREVIOLATION_RESPONSE PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnTypeChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnOverwrite (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_OVERWRITE_RESPONSE PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount As ULONG

END TYPE
' ########################################################################################

' ########################################################################################
' Template example to set a IFileDialogEvents innterface.
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CIFileDialogEvents
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CIFileDialogEvents
END DESTRUCTOR
' =====================================================================================

' *** IUnknown interface methods ***

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogEvents.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   OutputDebugStringW("CIFileDialogEvents.QueryInterface")
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_IFileOpenDialog) OR _
      IsEqualIID(riid, @IID_IFileSaveDialog) OR _
      IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
      ' // Increment the reference count
      this.AddRef
      RETURN S_OK
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' =====================================================================================

' ========================================================================================
' Increments the reference count for an interface on an object. This method should be called
' for every new copy of a pointer to an interface on an object.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogEvents.AddRef () AS ULONG
   refCount += 1
   OutputDebugStringW("CIFileDialogEvents.AddRef " & ..WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogEvents.Release () AS ULONG
   refCount -= 1
   OutputDebugStringW("CIFileDialogEvents.Release " & ..WSTR(refCount))
   IF refCount = 0 THEN
      OutputDebugStringW("CIFileDialogEvents - Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' *** Event Handlers ***

' =====================================================================================
' Called just before the dialog is about to return with a result.
' When this method is called, the IFileDialog::GetResult and GetResults methods can be called.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnFileOk (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   OutputDebugStringW("CIFileDialogEvents.OnFileOk")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called before OnFolderChange. This allows the implementer to stop navigation to a particular location.
' A pointer to an interface that represents the folder to which the dialog is about to navigate.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnFolderChanging (BYVAL pfd AS IFileDialog PTR, BYVAL psiFolder AS IShellItem PTR) AS HRESULT
   OutputDebugStringW("CIFileDialogEvents.OnFolderChanging")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called when the user navigates to a new folder.
' OnFolderChange is called when the dialog is opened.
' The code below allows to set the position of the dialog, but can't center it because
' the width and height of the dialog isn't settled until it is displayed.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   OutputDebugStringW("CIFileDialogEvents.OnFolderChange")

   ' // Before displaying the dialog get its coordinates
   DIM pOleWindow AS IOleWindow PTR
   ' // Get a pointer to the IOleWindow interface
   pfd->lpvtbl->QueryInterface(pfd, @IID_IOleWindow, @pOleWindow)
   ' // Get the window handle of the dialog
   DIM hOleWindow AS HWND
   IF pOleWindow THEN pOleWindow->lpvtbl->GetWindow(pOleWindow, @hOleWindow)
   OutputDebugStringW("CIFileDialogEvents.OnFolderChange - hOleWindow: " & ..WSTR(hOleWindow))
   IF hOleWindow THEN
      ' // Get he bounding rectangle of the parent window
      DIM AS RECT rcDlg, rcOwner
      GetWindowRect(GetParent(hOleWindow), @rcOwner)
      ' // Get he bounding rectangle of the open file dialog
      GetWindowRect(hOleWindow, @rcDlg)
      ' // Maps the open file dialog coordinates relative to its parent window
      MapWindowPoints(GetParent(hOleWindow), hOleWindow, CAST(POINT PTR, @rcDlg), 2)
   END IF

   RETURN S_OK

END FUNCTION
' =====================================================================================

' =====================================================================================
' Called when the user changes the selection in the dialog's view.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnSelectionChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   OutputDebugStringW("CIFileDialogEvents.OnSelectionChange")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Enables an application to respond to sharing violations that arise from Open or Save operations.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnShareViolation (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_SHAREVIOLATION_RESPONSE PTR) AS HRESULT
   OutputDebugStringW("CIFileDialogEvents.OnShareViolation")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called when the dialog is opened to notify the application of the initial chosen filetype.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnTypeChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   OutputDebugStringW("CIFileDialogEvents.OnTypeChange")
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called from the save dialog when the user chooses to overwrite a file.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnOverwrite (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_OVERWRITE_RESPONSE PTR) AS HRESULT
   OutputDebugStringW("CIFileDialogEvents.OnOverwrite")
   RETURN S_OK
END FUNCTION
' =====================================================================================

```
