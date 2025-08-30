# CIFileDialogEvents and CIFileDialogControlEvents classes.

The `CIFileDialogEents` class implements the `IFileDialogEvents` interface, that exposes methods that allow notification of events within a common file dialog.

The `CIFileDialogControlEvents` class implements the `IFileDialogControlEvents` interface, that exposes methods that allow an application to be notified of events that are related to controls that the application has added to a common file dialog.

**Include file:** AfxNova/CIFileDialogEvents.inc

---

## CIFileDialogEents Methods

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

## CIFileDialogControlEvents Methods

| Name       | Description |
| ---------- | ----------- |
| [OnButtonClicked](#onbuttonclicked) | Called when the user clicks a command button. |
| [OnCheckButtonToggled](#oncheckbuttontoggled) | Called when the user changes the state of a check button (check box). |
| [OnControlActivating](#oncontrolactivating) | Called when an **Open** button drop-down list customized through **EnableOpenDropDown** or a Tools menu is about to display its contents. |
| [OnItemSelected](#onitemselected) | Called when an item is selected in a combo box, when a user clicks an option button (also known as a radio button), or an item is chosen from the **Tools** menu. |

---

## Template code

| Name       | Description |
| ---------- | ----------- |
| [Template](#template) | Template code that implements the **FileDialogEvents** callback interface. |
| [Template2](#template2) | Template code that implements the **FileDialogEvents** and **IFileDialogControlEvents** callback interfaces. |

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

When this method is called, the **GetResult**, **GetResultString** and **GetResults** methods can be called.

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

## OnButtonClicked

Called when the user clicks a command button.

```
FUNCTION OnButtonClicked (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfdc* | A pointer to the interface through which the application added controls to the dialog. |
| *dwIDCtl* | The ID of the button that the user clicked. |

#### Return value

If this method succeeds, it returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

---

## OnCheckButtonToggled

Called when the user changes the state of a check button (check box).

```
FUNCTION OnCheckButtonToggled (BYVAL pfdc AS IFileDialogCustomize PTR, _
   BYVAL dwIDCtl AS DWORD, BYVAL bChecked AS WINBOOL) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfdc* | A pointer to the interface through which the application added controls to the dialog. |
| *dwIDCtl* | The ID of the button that the user clicked. |
| *bChecked* | A **BOOL** indicating the current state of the check button. **TRUE** if checked; **FALSE** otherwise. |

#### Return value

If this method succeeds, it returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

---

## OnControlActivating

Called when an **Open** button drop-down list customized through **EnableOpenDropDown** or a **Tools** menu is about to display its contents.

```
FUNCTION OnControlActivating (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
```

| Parameter  | Description |
| ---------- | ----------- |
| *pfdc* | A pointer to the interface through which the application added controls to the dialog. |
| *dwIDCtl* | The ID of the list or menu about to display. |

#### Return value

If this method succeeds, it returns **S_OK**. Otherwise, it returns an **HRESULT** error code.

---

## Template

The following template implements the **FileDialogEvents** callback interface.

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
   refCount AS ULONG = 0

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
   COSFD_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_IFileOpenDialog) OR _
      IsEqualIID(riid, @IID_IFileSaveDialog) OR _
      IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
      ' // Increment the reference count
      this.AddRef
      RETURN S_OK
   ELSE
      IF IsEqualIID(riid, @IID_IFileDialogControlEvents) THEN
         DIM pIFileDialogControlEvents AS CIFileDialogControlEvents PTR = NEW CIFileDialogControlEvents
         pIFileDialogControlEvents->AddRef
         *ppvObj = CAST(ANY PTR, pIFileDialogControlEvents)
         RETURN S_OK
      END IF
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
   IF refCount = 0 THEN
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
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called before OnFolderChange. This allows the implementer to stop navigation to a particular location.
' A pointer to an interface that represents the folder to which the dialog is about to navigate.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnFolderChanging (BYVAL pfd AS IFileDialog PTR, _
   BYVAL psiFolder AS IShellItem PTR) AS HRESULT
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

'   ' // Before displaying the dialog get its coordinates
'   DIM pOleWindow AS IOleWindow PTR
'   ' // Get a pointer to the IOleWindow interface
'   pfd->lpvtbl->QueryInterface(pfd, @IID_IOleWindow, @pOleWindow)
'   ' // Get the window handle of the dialog
'   DIM hOleWindow AS HWND
'   IF pOleWindow THEN pOleWindow->lpvtbl->GetWindow(pOleWindow, @hOleWindow)
'   OutputDebugStringW("CIFileDialogEvents.OnFolderChange - hOleWindow: " & ..WSTR(hOleWindow))
'   IF hOleWindow THEN
'      ' // Get he bounding rectangle of the parent window
'      DIM AS RECT rcDlg, rcOwner
'      GetWindowRect(GetParent(hOleWindow), @rcOwner)
'      ' // Get he bounding rectangle of the open file dialog
'      GetWindowRect(hOleWindow, @rcDlg)
'      ' // Maps the open file dialog coordinates relative to its parent window
'      MapWindowPoints(GetParent(hOleWindow), hOleWindow, CAST(POINT PTR, @rcDlg), 2)
'   END IF

   RETURN S_OK

END FUNCTION
' =====================================================================================

' =====================================================================================
' Called when the user changes the selection in the dialog's view.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnSelectionChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Enables an application to respond to sharing violations that arise from Open or Save operations.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnShareViolation (BYVAL pfd AS IFileDialog PTR, _
   BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_SHAREVIOLATION_RESPONSE PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called when the dialog is opened to notify the application of the initial chosen filetype.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnTypeChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called from the save dialog when the user chooses to overwrite a file.
' =====================================================================================
PRIVATE FUNCTION CIFileDialogEvents.OnOverwrite (BYVAL pfd AS IFileDialog PTR, _
   BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_OVERWRITE_RESPONSE PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

```

Example of how to override the template with yur own class:

```
' ########################################################################################
' CIFileDialogEventsImpl class
' Implementation of the FileDialogEvents callback interface
' ########################################################################################
TYPE CIFileDialogEventsImpl EXTENDS CIFileDialogEvents

   DECLARE FUNCTION OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT

END TYPE

PRIVATE FUNCTION CIFileDialogEventsImpl.OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT

   ' // Before displaying the dialog get its coordinates
   DIM pOleWindow AS IOleWindow PTR
   ' // Get a pointer to the IOleWindow interface
   pfd->lpvtbl->QueryInterface(pfd, @IID_IOleWindow, @pOleWindow)
   ' // Get the window handle of the dialog
   DIM hOleWindow AS HWND
   IF pOleWindow THEN pOleWindow->lpvtbl->GetWindow(pOleWindow, @hOleWindow)
   OutputDebugStringW ("CIFileDialogEventsImpl.OnFolderChange - hOleWindow: " & ..WSTR(hOleWindow))
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
' ========================================================================================
```

Set the events to your class with the **SetEvents** method:

```
' // Set events
DIM pfde AS ANY PTR = NEW CIFileDialogEventsImpl
pofd.SetEvents(pfde)
```

```
DIM pofd AS CIOpenFileDialog
' // Set the file types
pofd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
pofd.AddFileType("Executable files", "*.exe;*.dll")
pofd.AddFileType("All files", "*.*")
pofd.SetFileTypes()
' // Multiple selection (default is single selection)
DIM options AS FILEOPENDIALOGOPTIONS = pofd.GetOptions
pofd.SetOptions(options OR FOS_ALLOWMULTISELECT)
' // Optional: Set the title of the dialog
'   pofd.SetTitle("A Single-Selection Dialog")
' // Set the folder
pofd.SetFolder(CURDIR)
pofd.SetDefaultExtension("bas")
pofd.SetFileTypeIndex(1)
' // Set events
DIM pfde AS ANY PTR = NEW CIFileDialogEventsImpl
pofd.SetEvents(pfde)
' // Display the dialog
DIM hr AS HRESULT = pofd.ShowOpen(hwnd)
' // Folder name
OutputDebugStringW("Folder name: " & pofd.GetFolder)
' *** Single selection ***
' // Get the result
'IF hr = S_OK THEN
'   print pofd.GetResultString
'END IF
' *** Multiple selection ***
IF hr = S_OK THEN
   DIM dwsRes AS DWSTRING = pofd.GetResultsString
   FOR i AS LONG = 1 TO pofd.GetResultsCount
      OutputDebugStringW pofd.ParseResults(dwsRes, i)
   NEXT
END IF
```

You can use the same class of the Save file dialog:

```
' ########################################################################################
' CIFileDialogEventsImpl class
' Implementation of the FileDialogEvents callback interface
' ########################################################################################
TYPE CIFileDialogEventsImpl EXTENDS CIFileDialogEvents

   DECLARE FUNCTION OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE FUNCTION OnFileOk (BYVAL pfd AS IFileDialog PTR) AS HRESULT

END TYPE

PRIVATE FUNCTION CIFileDialogEventsImpl.OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   RETURN S_OK
END FUNCTION

PRIVATE FUNCTION CIFileDialogEventsImpl.OnFileOk (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' ########################################################################################
```

Setting the event's handler:

```
DIM psfd AS CISaveFileDialog
' // Set the file types
psfd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
psfd.AddFileType("Executable files", "*.exe;*.dll")
psfd.AddFileType("All files", "*.*")
psfd.SetFileTypes()
' // Optional: Set the title of the dialog
'   psfd.SetTitle("Save File Dialog")
' // Set the folder
psfd.SetFolder(CURDIR)
psfd.SetDefaultExtension("bas")
psfd.SetFileTypeIndex(1)
' // Set events
DIM pfde AS ANY PTR = NEW CIFileDialogEventsImpl
psfd.SetEvents(pfde)
' // Display the dialog
DIM hr AS HRESULT = psfd.ShowSave(hwnd)
' // Folder name
OutputDebugStringW("Folder name: " & psfd.GetFolder)
' // Get the result
IF hr = S_OK THEN
   OutputDebugStringW(psfd.GetResultString)
END IF
```
---

## Template2

The following template implements the **FileDialogEvents** and **IFileDialogControlEvents** callback interfaces.

```
' ########################################################################################
' Microsoft Windows
' File: CW_CIOpenFileDialog_Events_02.bas
' Contents: Open file dialog with dialog and control events
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José© Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#define _COSFD_DEBUG_ 1
#include once "AfxNova/CWindow.inc"
#include once "AfxNova/AfxCOM.inc"
#include once "AfxNova/CIOpenSaveFile.inc"
#include once "AfxNova/CIFileDialogCustomize.inc"
#include once "AfxNova/CIFileDialogEvents.inc"
USING AfxNova

CONST IDC_RICHEDIT = 1001
CONST IDC_TEST = 1002

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   DIM pWindow AS CWindow
   DIM hWin AS HWND = pWindow.Create(NULL, "DisplayIOpenFile", @WndProc)
   pWindow.SetClientSize(800, 450)
   pWindow.Center

   ' // Add buttons
   pWindow.AddControl("Button", hWin, IDC_TEST, "&Test", pWindow.ClientWidth - 195, pWindow.ClientHeight - 35, 75, 23)
   pWindow.AddControl("Button", hWin, IDCANCEL, "&Close", pWindow.ClientWidth - 95, pWindow.ClientHeight - 35, 75, 23)
  
   ' // Anchor the controls
   pWindow.AnchorControl(IDC_TEST, AFX_ANCHOR_BOTTOM_RIGHT)
   pWindow.AnchorControl(IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

   ' // Dispatch Windows messages
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ########################################################################################
' CIFileDialogControlEventsImpl class
' Implementation of the IFileDialogControlEvents callback interface
' ########################################################################################
TYPE CIFileDialogControlEventsImpl EXTENDS CIFileDialogControlEvents

   DECLARE FUNCTION OnItemSelected (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   DECLARE FUNCTION OnButtonClicked (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE FUNCTION OnCheckButtonToggled (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL bChecked AS WINBOOL) AS HRESULT
   DECLARE FUNCTION OnControlActivating (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT

END TYPE

' // Override the virtual methods

' ========================================================================================
PRIVATE FUNCTION CIFileDialogControlEventsImpl.OnItemSelected (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   COSFD_DP("")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIFileDialogControlEventsImpl.OnButtonClicked (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
   COSFD_DP("")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIFileDialogControlEventsImpl.OnCheckButtonToggled (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL bChecked AS WINBOOL) AS HRESULT
   COSFD_DP("")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CIFileDialogControlEventsImpl.OnControlActivating (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
   COSFD_DP("")
   RETURN S_OK
END FUNCTION
' ========================================================================================

' ########################################################################################
' CIFileDialogEventsImpl class
' Implementation of the FileDialogEvents callback interface
' ########################################################################################
TYPE CIFileDialogEventsImpl EXTENDS CIFileDialogEvents

   DECLARE FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   DECLARE FUNCTION OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   DECLARE FUNCTION OnItemSelected (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT

END TYPE

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CIFileDialogEventsImpl.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   COSFD_DP(AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_IFileOpenDialog) OR _
      IsEqualIID(riid, @IID_IFileSaveDialog) OR _
      IsEqualIID(riid, @IID_IUnknown) THEN
      *ppvObj = @this
      ' // Increment the reference count
      this.AddRef
      RETURN S_OK
   ELSE
      IF IsEqualIID(riid, @IID_IFileDialogControlEvents) THEN
         ' // Create an instance of our implemented IFileDialogControlEvents interface
         DIM pIFileDialogControlEvents AS CIFileDialogControlEventsImpl PTR = NEW CIFileDialogControlEventsImpl
         ' // Increment the reference count
         pIFileDialogControlEvents->AddRef
         ' // Return a pointer to the event's class
         *ppvObj = CAST(ANY PTR, pIFileDialogControlEvents)
         RETURN S_OK
      END IF
   END IF
   RETURN E_NOINTERFACE
END FUNCTION
' =====================================================================================

PRIVATE FUNCTION CIFileDialogEventsImpl.OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   ' // Before displaying the dialog get its coordinates
   DIM pOleWindow AS IOleWindow PTR
   ' // Get a pointer to the IOleWindow interface
   pfd->lpvtbl->QueryInterface(pfd, @IID_IOleWindow, @pOleWindow)
   ' // Get the window handle of the dialog
   DIM hOleWindow AS HWND
   IF pOleWindow THEN pOleWindow->lpvtbl->GetWindow(pOleWindow, @hOleWindow)
   COSFD_DP ("hOleWindow: " & WSTR(hOleWindow))
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
' ########################################################################################

' ========================================================================================
' Main window callback procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC hFocus AS HWND

   SELECT CASE uMsg

      CASE WM_NCACTIVATE
         ' // Save the handle of the control that has the focus
         IF wParam = 0 THEN hFocus = GetFocus
         ' // Note: Don't use EXIT FUNCTION

      CASE WM_SETFOCUS
         ' // Post a message to set the focus later, since some
         ' // Windows actions can steal it if we set it here
         IF hFocus THEN
            PostMessage hWnd, WM_USER + 999, cast(WPARAM, hFocus), 0
            hFocus = 0
         END IF

      CASE WM_USER + 999
         ' // Set the focus and show the line an column in the status bar
         IF wParam THEN SetFocus(cast(HWND, wParam))
         EXIT FUNCTION

      CASE WM_SYSCOMMAND
         ' // Capture this message and send a WM_CLOSE message
         IF (wParam AND &HFFF0) = SC_CLOSE THEN
            SendMessage hwnd, WM_CLOSE, 0, 0
            EXIT FUNCTION
         END IF

      CASE WM_COMMAND
         
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)

            ' // If ESC key pressed, close the application sending an WM_CLOSE message
            CASE IDCANCEL
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF

            ' // For testing purposes
            CASE IDC_TEST
               SCOPE
               DIM pofd AS CIOpenFileDialog
               ' // Set the file types
               pofd.AddFileType("FB code files", "*.bas;*.inc;*.bi")
               pofd.AddFileType("Executable files", "*.exe;*.dll")
               pofd.AddFileType("All files", "*.*")
               pofd.SetFileTypes()
               ' // Multiple selection (default is single selection)
               DIM options AS FILEOPENDIALOGOPTIONS = pofd.GetOptions
               pofd.SetOptions(options OR FOS_ALLOWMULTISELECT)
               ' // Optional: Set the title of the dialog
               '   pofd.SetTitle("A Single-Selection Dialog")
               ' // Set the folder
               pofd.SetFolder(CURDIR)
               pofd.SetDefaultExtension("bas")
               pofd.SetFileTypeIndex(1)
               ' // Customization
               ' // Get a pointer to the IFileDialog interface
               DIM pfdlg AS IFileDialog PTR = pofd.GetIFileDialogPtr
               IF pfdlg THEN
                  ' // Initialize an instance of the CIFileDialogCustomize class
                  DIM pCust AS CIFileDialogCustomize = pfdlg
                  ' // Relese the IFileDialog interface
                  pfdlg->lpvtbl->Release(pfdlg)
                  ' // Add controls to the dialog
                  IF pCust.GetPtr THEN
                     ' Call methods of the CIFileDialogCustomize class
                     pCust.AddCheckButton(1001, "My check button", TRUE)
                     pCust.AddComboBox(1002)
                     pCust.AddEditBox(1003, "My edit control")
                     pCust.AddMenu(1004, "My menu")
                     pCust.AddPushButton(1005, "My push button")
                     pCust.AddSeparator(1006)
                     pCust.AddText(1007, "My text")
                     pCust.EnableOpenDropDown(1008)
                  END IF
               END IF
               ' // Set events
               DIM pfde AS ANY PTR = NEW CIFileDialogEventsImpl
               pofd.SetEvents(pfde)
               ' // Display the dialog
               DIM hr AS HRESULT = pofd.ShowOpen(hwnd)
               ' // Folder name
               OutputDebugStringW("Folder name: " & pofd.GetFolder)
               ' *** Single selection ***
               ' // Get the result
               'IF hr = S_OK THEN
               '   OutputDebugStringW pofd.GetResult()
               'END IF
               ' *** Multiple selection ***
               IF hr = S_OK THEN
                  DIM dwsRes AS DWSTRING = pofd.GetResultsString
                  FOR i AS LONG = 1 TO pofd.GetResultsCount
                     OutputDebugStringW pofd.ParseResults(dwsRes, i)
                  NEXT
               END IF
               END SCOPE

         END SELECT

    	CASE WM_DESTROY
         ' // End the application by sending an WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
```
---


