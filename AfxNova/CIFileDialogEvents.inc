' ########################################################################################
' Platform: Microsoft Windows
' File: CIFileDialogEvents.inc
' Contents: Implementation of IFileDialogEvents callback interface.
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

#endif
