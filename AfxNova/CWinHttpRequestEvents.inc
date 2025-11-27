' ########################################################################################
' CWinHttpRequestEvents class
' Implementation of the IWinHttpRequestEvents callback interface
' IID_IWinHttpRequestEvents = "{F97F4E15-B787-4212-80D1-D380CBBF982E}"
' ########################################################################################

TYPE CWinHttpRequestEvents EXTENDS OBJECT

   DECLARE VIRTUAL FUNCTION QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObject AS LPVOID PTR) AS HRESULT
   DECLARE VIRTUAL FUNCTION AddRef () AS ULONG
   DECLARE VIRTUAL FUNCTION Release () AS ULONG
   DECLARE VIRTUAL SUB OnResponseStart (BYVAL Status AS LONG, BYVAL ContentType AS AFX_BSTR)
   DECLARE VIRTUAL SUB OnResponseDataAvailable (BYVAL Data AS SAFEARRAY PTR)
   DECLARE VIRTUAL SUB OnResponseFinished ()
   DECLARE VIRTUAL SUB OnError (BYVAL ErrorNumber AS LONG, BYVAL ErrorDescription AS AFX_BSTR)

   DECLARE CONSTRUCTOR
   DECLARE VIRTUAL DESTRUCTOR

   ' Reference count for COM
   refCount As ULONG

END TYPE
' ########################################################################################

' ########################################################################################
' Template example to set a IWinHttpRequestEvents innterface.
' ########################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
CONSTRUCTOR CWinHttpRequestEvents
END CONSTRUCTOR
' =====================================================================================
' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CWinHttpRequestEvents
END DESTRUCTOR
' =====================================================================================

' IUnknown interface methods
' ========================================================================================
' Returns pointers to the implemented classes and supported interfaces.
' ========================================================================================
FUNCTION CWinHttpRequestEvents.QueryInterface (BYVAL riid AS REFIID, BYVAL ppvObj AS LPVOID PTR) AS HRESULT
   OutputDebugStringW("CWinHttpRequestEvents.QueryInterface " & AfxGuidText(riid))
   IF ppvObj = NULL THEN RETURN E_INVALIDARG
   DIM AFX_IID_IWinHttpRequestEvents_ AS GUID = (&hF97F4E15, &hB787, &h4212, {&h80, &hD1, &hD3, &h80, &hCB, &hBF, &h98, &h2E})
   IF IsEqualIID(riid, @AFX_IID_IWinHttpRequestEvents_) OR _
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
FUNCTION CWinHttpRequestEvents.AddRef () AS ULONG
   refCount += 1
   OutputDebugStringW("CWinHttpRequestEvents.AddRef " & ..WSTR(refCount))
   RETURN refCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Decrements the reference count for an interface on an object.
' If the count reaches 0, it deletes itself.
' ========================================================================================
FUNCTION CWinHttpRequestEvents.Release () AS ULONG
   refCount -= 1
   OutputDebugStringW("CWinHttpRequestEvents.Release " & ..WSTR(refCount))
   IF refCount = 0 THEN
      OutputDebugStringW("CWinHttpRequestEvents - Delete class")
      Delete @this
   END IF
   RETURN refCount
END FUNCTION
' =====================================================================================

' Event Handlers
' =====================================================================================
' The OnResponseStart event occurs when the response data starts to be received.
' Parameters:
' - Status [in]
'   Receives the standard status code returned with the response data. Status codes are defined in RFC 2616.
' - ContentType [in]
'   Specifies the type of content received, such as "text/html" or "image/gif".
' Return value
'   This event does not return a value.
' =====================================================================================
SUB CWinHttpRequestEvents.OnResponseStart (BYVAL Status AS LONG, BYVAL ContentType AS AFX_BSTR)
   OutputDebugStringW("CWinHttpRequestEvents.OnResponseStart")
END SUB
' =====================================================================================

' =====================================================================================
' The OnResponseDataAvailable event occurs when data is available from the response.
' - Data [in]
'   A zero-based array of bytes that receives the response data received by Microsoft
'   Windows HTTP Services (WinHTTP) up to the point that this event occurs.
'   This is a VARIANT of type VT_ARRAY | VT_UI1.
' Return value
'   This event does not return a value.
' =====================================================================================
SUB CWinHttpRequestEvents.OnResponseDataAvailable (BYVAL pData AS SAFEARRAY PTR)
   OutputDebugStringW("CWinHttpRequestEvents.OnResponseDataAvailable")
END SUB
' =====================================================================================

' =====================================================================================
' The OnResponseFinished event occurs when the response data is complete.
' This event signals that all of the response data that pertains to the most recent
' request has been received. OnResponseDataAvailable does not occur again until the next request.
' =====================================================================================
SUB CWinHttpRequestEvents.OnResponseFinished
   OutputDebugStringW("CWinHttpRequestEvents.OnResponseFinished")
END SUB
' =====================================================================================

' =====================================================================================
' The OnResponseFinished event occurs when the response data is complete.
' This event signals that all of the response data that pertains to the most recent
' request has been received. OnResponseDataAvailable does not occur again until the next request.
' =====================================================================================
SUB CWinHttpRequestEvents.OnError (BYVAL ErrorNumber AS LONG, BYVAL ErrorDescription AS AFX_BSTR)
   OutputDebugStringW("CWinHttpRequestEvents.OnError")
END SUB
' =====================================================================================
