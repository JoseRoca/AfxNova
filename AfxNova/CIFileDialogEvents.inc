' ########################################################################################
' Platform: Microsoft Windows
' File: CIFileDialogEvents.inc
' Contents: Implementation of IFileDialogEvents and IFileDialogControlEvents callback interfaces.
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2025 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#if _WIN32_WINNT >= &h0600


' ########################################################################################
' CIFileDialogControlEvents class
' Implementation of the IFileDialogControlEvents callback interfaces.
' ########################################################################################
TYPE CIFileDialogControlEvents EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG

   ' IFileDialogControlEvents
   DECLARE VIRTUAL FUNCTION OnItemSelected (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORD) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnButtonClicked (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnCheckButtonToggled (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL bChecked AS WINBOOL) AS HRESULT
   DECLARE VIRTUAL FUNCTION OnControlActivating (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE VIRTUAL DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' ########################################################################################
' Template example to set a IFileDialogControlEvents innterface.
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
CONSTRUCTOR CIFileDialogControlEvents
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CIFileDialogControlEvents
END DESTRUCTOR
' =====================================================================================

' *** IUnknown interface methods ***

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CIFileDialogControlEvents.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   COSFD_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @IID_IFileDialogControlEvents) OR _
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
FUNCTION CIFileDialogControlEvents.AddRef () AS ULONG
   refCount += 1
   COSFD_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CIFileDialogControlEvents.Release () AS ULONG
   refCount -= 1
   COSFD_DP(WSTR(refCount))
   IF refCount = 0 THEN
      COSFD_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' // IFileDialogControlEvents

' =====================================================================================
' Called when an item is selected in a combo box, when a user clicks an option button
' (also known as a radio button), or an item is chosen from the Tools menu.
' =====================================================================================
FUNCTION CIFileDialogControlEvents.OnItemSelected (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL dwIDItem AS DWORd) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called when the user clicks a command button.
' =====================================================================================
FUNCTION CIFileDialogControlEvents.OnButtonClicked (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called when the user changes the state of a check button (check box).
' =====================================================================================
FUNCTION CIFileDialogControlEvents.OnCheckButtonToggled (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD, BYVAL bChecked AS WINBOOL) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Enables an application to respond to sharing violations that arise from Open or Save operations.
' =====================================================================================
FUNCTION CIFileDialogControlEvents.OnControlActivating (BYVAL pfdc AS IFileDialogCustomize PTR, BYVAL dwIDCtl AS DWORD) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' ########################################################################################
' CIFileDialogEvents class
' Implementation of the FileDialogEvents and IFileDialogControlEvents callback interfaces.
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
   DECLARE VIRTUAL DESTRUCTOR

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
CONSTRUCTOR CIFileDialogEvents
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CIFileDialogEvents
END DESTRUCTOR
' =====================================================================================

' *** IUnknown interface methods ***

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CIFileDialogEvents.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
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
FUNCTION CIFileDialogEvents.AddRef () AS ULONG
   refCount += 1
   COSFD_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CIFileDialogEvents.Release () AS ULONG
   refCount -= 1
   COSFD_DP(WSTR(refCount))
   IF refCount = 0 THEN
      COSFD_DP("Delete class")
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
FUNCTION CIFileDialogEvents.OnFileOk (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called before OnFolderChange. This allows the implementer to stop navigation to a particular location.
' A pointer to an interface that represents the folder to which the dialog is about to navigate.
' =====================================================================================
FUNCTION CIFileDialogEvents.OnFolderChanging (BYVAL pfd AS IFileDialog PTR, BYVAL psiFolder AS IShellItem PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called when the user navigates to a new folder.
' OnFolderChange is called when the dialog is opened.
' The code below allows to set the position of the dialog, but can't center it because
' the width and height of the dialog isn't settled until it is displayed.
' =====================================================================================
FUNCTION CIFileDialogEvents.OnFolderChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT

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
FUNCTION CIFileDialogEvents.OnSelectionChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Enables an application to respond to sharing violations that arise from Open or Save operations.
' =====================================================================================
FUNCTION CIFileDialogEvents.OnShareViolation (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_SHAREVIOLATION_RESPONSE PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called when the dialog is opened to notify the application of the initial chosen filetype.
' =====================================================================================
FUNCTION CIFileDialogEvents.OnTypeChange (BYVAL pfd AS IFileDialog PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

' =====================================================================================
' Called from the save dialog when the user chooses to overwrite a file.
' =====================================================================================
FUNCTION CIFileDialogEvents.OnOverwrite (BYVAL pfd AS IFileDialog PTR, BYVAL psi AS IShellItem PTR, BYVAL pResponse AS FDE_OVERWRITE_RESPONSE PTR) AS HRESULT
   RETURN S_OK
END FUNCTION
' =====================================================================================

#endif
