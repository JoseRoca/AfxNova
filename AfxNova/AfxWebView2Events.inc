' ########################################################################################
' Library name: AfxWebView2Events.inc
' (C) 2025 José Roca.
' ########################################################################################

#pragma once
#include once "AfxNova/AfxWebView2.bi"

' ########################################################################################
' CWebView2EnvironmentCompletedHandler class
' Implementation of the ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler callback interface.
' ########################################################################################
TYPE CWebView2EnvironmentCompletedHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL createdEnvironment AS Afx_ICoreWebView2Environment PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2EnvironmentCompletedHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2EnvironmentCompletedHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2EnvironmentCompletedHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler) OR _
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
PRIVATE FUNCTION CWebView2EnvironmentCompletedHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2EnvironmentCompletedHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2EnvironmentCompletedHandler.Invoke (BYVAL errorCode AS HRESULT, BYVAL createdEnvironment AS Afx_ICoreWebView2Environment PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2ControllerCompletedHandler class
' Implementation of the ICoreWebView2CreateCoreWebView2ControllerCompletedHandler callback interface.
' ########################################################################################
TYPE CWebView2ControllerCompletedHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL createdController AS Afx_ICoreWebView2Controller PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2ControllerCompletedHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2ControllerCompletedHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2ControllerCompletedHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler) OR _
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
PRIVATE FUNCTION CWebView2ControllerCompletedHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2ControllerCompletedHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2ControllerCompletedHandler.Invoke (BYVAL errorCode AS HRESULT, BYVAL createdController AS Afx_ICoreWebView2Controller PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2NavigationStartingdEventHandler class
' Implementation of the ICoreWebView2NavigationStartingEventHandler callback interface.
' ########################################################################################
TYPE CWebView2NavigationStartingdEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NavigationCompletedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2NavigationStartingdEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2NavigationStartingdEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2NavigationStartingdEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2NavigationStartingEventHandler) OR _
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
PRIVATE FUNCTION CWebView2NavigationStartingdEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2NavigationStartingdEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2NavigationStartingdEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NavigationCompletedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2NavigationCompletedEventHandler class
' Implementation of the ICoreWebView2NavigationCompletedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2NavigationCompletedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NavigationCompletedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2NavigationCompletedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2NavigationCompletedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2NavigationCompletedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2NavigationCompletedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2NavigationCompletedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2NavigationCompletedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2NavigationCompletedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NavigationCompletedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2ContentLoadingEventHandler class
' Implementation of the ICoreWebView2ContentLoadingEventHandler callback interface.
' ########################################################################################
TYPE CWebView2ContentLoadingEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ContentLoadingEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2ContentLoadingEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2ContentLoadingEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2ContentLoadingEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @Afx_IID_ICoreWebView2ContentLoadingEventHandler) OR _
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
PRIVATE FUNCTION CWebView2ContentLoadingEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2ContentLoadingEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2ContentLoadingEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ContentLoadingEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2SourceChangedEventHandler class
' Implementation of the ICoreWebView2SourceChangedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2SourceChangedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2SourceChangedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2SourceChangedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2SourceChangedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2SourceChangedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2SourceChangedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2SourceChangedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2SourceChangedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2SourceChangedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2SourceChangedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2HistoryChangedEventHandler class
' Implementation of the ICoreWebView2HistoryChangedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2HistoryChangedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2HistoryChangedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2HistoryChangedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2HistoryChangedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2HistoryChangedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2HistoryChangedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2HistoryChangedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2HistoryChangedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2ScriptDialogOpeningEventHandler class
' Implementation of the ICoreWebView2ScriptDialogOpeningEventHandler callback interface.
' ########################################################################################
TYPE CWebView2ScriptDialogOpeningEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ScriptDialogOpeningEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2ScriptDialogOpeningEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2ScriptDialogOpeningEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2ScriptDialogOpeningEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2ScriptDialogOpeningEventHandler) OR _
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
PRIVATE FUNCTION CWebView2ScriptDialogOpeningEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2ScriptDialogOpeningEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2ScriptDialogOpeningEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ScriptDialogOpeningEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2PermissionRequestedEventHandler class
' Implementation of the ICoreWebView2PermissionRequestedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2PermissionRequestedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2PermissionRequestedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2PermissionRequestedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2PermissionRequestedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2PermissionRequestedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2PermissionRequestedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2PermissionRequestedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2PermissionRequestedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2PermissionRequestedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2PermissionRequestedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2ProcessFailedEventHandler class
' Implementation of the ICoreWebView2ProcessFailedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2ProcessFailedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ProcessFailedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2ProcessFailedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2ProcessFailedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2ProcessFailedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2ProcessFailedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2ProcessFailedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2ProcessFailedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2ProcessFailedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2ProcessFailedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2WebMessageReceivedEventHandler class
' Implementation of the ICoreWebView2WebMessageReceivedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2WebMessageReceivedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2WebMessageReceivedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2WebMessageReceivedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2WebMessageReceivedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2WebMessageReceivedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2WebMessageReceivedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2WebMessageReceivedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2WebMessageReceivedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2WebMessageReceivedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2WebMessageReceivedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2NewWindowRequestedEventHandler class
' Implementation of the ICoreWebView2NewWindowRequestedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2NewWindowRequestedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NewWindowRequestedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2NewWindowRequestedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2NewWindowRequestedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2NewWindowRequestedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2NewWindowRequestedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2NewWindowRequestedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2NewWindowRequestedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2NewWindowRequestedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2NewWindowRequestedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2DocumentTitleChangedEventHandler class
' Implementation of the ICoreWebView2DocumentTitleChangedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2DocumentTitleChangedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2DocumentTitleChangedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2DocumentTitleChangedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2DocumentTitleChangedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2DocumentTitleChangedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2DocumentTitleChangedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2DocumentTitleChangedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2DocumentTitleChangedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2ContainsFullScreenElementChangedEventHandler class
' Implementation of the ICoreWebView2ContainsFullScreenElementChangedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2ContainsFullScreenElementChangedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2ContainsFullScreenElementChangedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2ContainsFullScreenElementChangedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2ContainsFullScreenElementChangedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2ContainsFullScreenElementChangedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2ContainsFullScreenElementChangedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2ContainsFullScreenElementChangedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2ContainsFullScreenElementChangedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2WebResourceRequestedEventHandler class
' Implementation of the ICoreWebView2WebResourceRequestedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2WebResourceRequestedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2WebResourceRequestedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2WebResourceRequestedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2WebResourceRequestedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2WebResourceRequestedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2WebResourceRequestedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2WebResourceRequestedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2WebResourceRequestedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2WebResourceRequestedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2WebResourceRequestedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2WindowCloseRequestedEventHandler class
' Implementation of the ICoreWebView2WindowCloseRequestedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2WindowCloseRequestedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2WindowCloseRequestedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2WindowCloseRequestedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2WindowCloseRequestedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2WindowCloseRequestedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2WindowCloseRequestedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2WindowCloseRequestedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2WindowCloseRequestedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS IUnknown PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2ExecuteScriptCompletedHandler class
' Implementation of the ICoreWebView2ExecuteScriptCompletedHandler callback interface.
' ########################################################################################
TYPE CWebView2ExecuteScriptCompletedHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL resultObjectAsJson AS LPCWSTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2ExecuteScriptCompletedHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2ExecuteScriptCompletedHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2ExecuteScriptCompletedHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2ExecuteScriptCompletedHandler) OR _
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
PRIVATE FUNCTION CWebView2ExecuteScriptCompletedHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2ExecuteScriptCompletedHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2ExecuteScriptCompletedHandler.Invoke (BYVAL errorCode AS HRESULT, BYVAL resultObjectAsJson AS LPCWSTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2AddScriptToExecuteOnDocumentCreatedCompleteHandler class
' Implementation of the ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler callback interface.
' ########################################################################################
TYPE CWebView2AddScriptToExecuteOnDocumentCreatedCompleteHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL resultObjectAsJson AS LPCWSTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2AddScriptToExecuteOnDocumentCreatedCompleteHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2AddScriptToExecuteOnDocumentCreatedCompleteHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2AddScriptToExecuteOnDocumentCreatedCompleteHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler) OR _
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
PRIVATE FUNCTION CWebView2AddScriptToExecuteOnDocumentCreatedCompleteHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2AddScriptToExecuteOnDocumentCreatedCompleteHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2AddScriptToExecuteOnDocumentCreatedCompleteHandler.Invoke (BYVAL errorCode AS HRESULT, BYVAL resultObjectAsJson AS LPCWSTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2AddCapturePreviewCompletedHandler class
' Implementation of the ICoreWebView2CapturePreviewCompletedHandler callback interface.
' ########################################################################################
TYPE CWebView2AddCapturePreviewCompletedHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL errorCode AS HRESULT) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2AddCapturePreviewCompletedHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2AddCapturePreviewCompletedHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2AddCapturePreviewCompletedHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2CapturePreviewCompletedHandler) OR _
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
PRIVATE FUNCTION CWebView2AddCapturePreviewCompletedHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2AddCapturePreviewCompletedHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2AddCapturePreviewCompletedHandler.Invoke (BYVAL errorCode AS HRESULT) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2DevToolsProtocolEventReceivedEventHandler class
' Implementation of the ICoreWebView2DevToolsProtocolEventReceivedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2DevToolsProtocolEventReceivedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2DevToolsProtocolEventReceivedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2DevToolsProtocolEventReceivedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2DevToolsProtocolEventReceivedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @Afx_IID_ICoreWebView2DevToolsProtocolEventReceivedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2DevToolsProtocolEventReceivedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2DevToolsProtocolEventReceivedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2DevToolsProtocolEventReceivedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2 PTR, BYVAL args AS Afx_ICoreWebView2DevToolsProtocolEventReceivedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2ZoomFactorChangedEventHandler class
' Implementation of the ICoreWebView2ZoomFactorChangedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2ZoomFactorChangedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS IUnknown PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2ZoomFactorChangedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2ZoomFactorChangedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2ZoomFactorChangedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2ZoomFactorChangedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2ZoomFactorChangedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2ZoomFactorChangedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2ZoomFactorChangedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS IUnknown PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2MoveFocusRequestedEventHandler class
' Implementation of the ICoreWebView2MoveFocusRequestedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2MoveFocusRequestedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS Afx_ICoreWebView2MoveFocusRequestedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2MoveFocusRequestedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2MoveFocusRequestedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2MoveFocusRequestedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2MoveFocusRequestedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2MoveFocusRequestedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2MoveFocusRequestedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2MoveFocusRequestedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS Afx_ICoreWebView2MoveFocusRequestedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2FocusChangedEventHandler Class
' Implementation of the ICoreWebView2FocusChangedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2FocusChangedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS IUnknown PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2FocusChangedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2FocusChangedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2FocusChangedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2FocusChangedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2FocusChangedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2FocusChangedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2FocusChangedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS IUnknown PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2AcceleratorKeyPressedEventHandler Class
' Implementation of the ICoreWebView2AcceleratorKeyPressedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2AcceleratorKeyPressedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS Afx_ICoreWebView2AcceleratorKeyPressedEventArgs PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2AcceleratorKeyPressedEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2AcceleratorKeyPressedEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2AcceleratorKeyPressedEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2AcceleratorKeyPressedEventHandler) OR _
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
PRIVATE FUNCTION CWebView2AcceleratorKeyPressedEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2AcceleratorKeyPressedEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2AcceleratorKeyPressedEventHandler.Invoke (BYVAL sender AS Afx_ICoreWebView2Controller PTR, BYVAL args AS Afx_ICoreWebView2AcceleratorKeyPressedEventArgs PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2NewBrowserVersionAvailableEventHandler Class
' Implementation of the ICoreWebView2NewBrowserVersionAvailableEventHandler callback interface.
' ########################################################################################
TYPE CWebView2NewBrowserVersionAvailableEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL webviewEnvironment AS Afx_ICoreWebView2Environment PTR, BYVAL args AS IUNknown PTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2NewBrowserVersionAvailableEventHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2NewBrowserVersionAvailableEventHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2NewBrowserVersionAvailableEventHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2NewBrowserVersionAvailableEventHandler) OR _
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
PRIVATE FUNCTION CWebView2NewBrowserVersionAvailableEventHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2NewBrowserVersionAvailableEventHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2NewBrowserVersionAvailableEventHandler.Invoke (BYVAL webviewEnvironment AS Afx_ICoreWebView2Environment PTR, BYVAL args AS IUNknown PTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================


' ########################################################################################
' CWebView2CallDevToolsProtocolMethodCompletedHandler Class
' Implementation of the ICoreWebView2CallDevToolsProtocolMethodCompletedHandler callback interface.
' ########################################################################################
TYPE CWebView2CallDevToolsProtocolMethodCompletedHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL FUNCTION Invoke (BYVAL errorCode AS HRESULT, BYVAL returnObjectAsJson AS LPCWSTR) AS HRESULT

   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR

   ' Reference count for COM
   refCount AS ULONG = 0

END TYPE
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
PRIVATE CONSTRUCTOR CWebView2CallDevToolsProtocolMethodCompletedHandler
   CWV2_DP("")
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
PRIVATE DESTRUCTOR CWebView2CallDevToolsProtocolMethodCompletedHandler
   CWV2_DP("")
END DESTRUCTOR
' =====================================================================================

' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
PRIVATE FUNCTION CWebView2CallDevToolsProtocolMethodCompletedHandler.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   CWV2_DP("riid: " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   IF IsEqualIID(riid, @AFX_IID_ICoreWebView2CallDevToolsProtocolMethodCompletedHandler) OR _
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
PRIVATE FUNCTION CWebView2CallDevToolsProtocolMethodCompletedHandler.AddRef () AS ULONG
   refCount += 1
   CWV2_DP(WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
PRIVATE FUNCTION CWebView2CallDevToolsProtocolMethodCompletedHandler.Release () AS ULONG
   refCount -= 1
   CWV2_DP(WSTR(refCount))
   IF refCount = 0 THEN
      CWV2_DP("Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' =====================================================================================
PRIVATE FUNCTION CWebView2CallDevToolsProtocolMethodCompletedHandler.Invoke (BYVAL errorCode AS HRESULT, BYVAL returnObjectAsJson AS LPCWSTR) AS HRESULT
   CWV2_DP("")
   RETURN S_OK
END FUNCTION
' =====================================================================================
