' ########################################################################################
' Library name: AfxWebView2Events.inc
' (C) 2025 José Roca.
' ########################################################################################

#pragma once
#include once "AfxWebView2.bi"

' ########################################################################################
' CWebView2EnvironmentCompletedHandler class
' Implementation of the ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler callback interface.
' ########################################################################################
TYPE CWebView2EnvironmentCompletedHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG

   ' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
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

   ' ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
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
' CWebView2NavigationCompletedEventHandler class
' Implementation of the ICoreWebView2NavigationCompletedEventHandler callback interface.
' ########################################################################################
TYPE CWebView2NavigationCompletedEventHandler EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG

   ' ICoreWebView2NavigationCompletedEventHandler
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
