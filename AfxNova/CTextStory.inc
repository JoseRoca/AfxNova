' ########################################################################################
' Microsoft Windows
' File: CTextStory.inc
' Contents: Class to work with the Richd Edit Text Object Model (TOM).
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' // To be included by CRichEditCtx.inc

' ########################################################################################
' CTextStory class
' ########################################################################################
TYPE CTextStory EXTENDS CTextObjectBase

Public:
   m_pTextStory AS ITextStory PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pTextStory AS ITextStory PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR LET (BYVAL pTextStory AS ITextStory PTR)
   DECLARE OPERATOR CAST () AS ITextStory PTR
   DECLARE FUNCTION TextStoryPtr () AS ITextStory PTR
   DECLARE FUNCTION Attach (BYVAL pTextStory AS ITextStory PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION Detach () AS ITextStory PTR

   ' // ITextStory interface methods
   DECLARE FUNCTION GetActive () AS LONG
   DECLARE FUNCTION SetActive (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetDisplay () AS IUnknown PTR
   DECLARE FUNCTION GetIndex () AS LONG
   DECLARE FUNCTION GetType () AS LONG
   DECLARE FUNCTION SetType (BYVAL Value AS LONG) AS HRESULT
   DECLARE FUNCTION GetProperty (BYVAL nType AS LONG) AS LONG
   DECLARE FUNCTION GetRange (BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG) AS ITextRange2 PTR
   DECLARE FUNCTION GetText (BYVAL Flags AS LONG) AS DWsTRING
   DECLARE FUNCTION SetFormattedText (BYVAL pUnk AS IUnknown PTR) AS HRESULT
   DECLARE FUNCTION SetProperty (BYVAL nType AS LONG, BYVAL VaLue AS LONG) AS HRESULT
   DECLARE FUNCTION SetText (BYVAL Flags AS LONG, BYREF wszText AS WSTRING) AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CTextStory constructor
' Called when a class variable is created.
' ========================================================================================
PRIVATE CONSTRUCTOR CTextStory
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CTextStory (BYVAL pTextStory AS ITextStory PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CREC_DP("BEGIN  - pTextStory = " & WSTR(m_pTextStory)) 
   ' // Store the pointer of ITextStory interface
   IF pTextStory = NULL THEN m_Result = E_INVALIDARG : EXIT CONSTRUCTOR
   IF fAddRef THEN pTextStory->lpvtbl->AddRef(pTextStory)
   m_pTextStory = pTextStory
   CREC_DP("END  - m_pTextStory = " & WSTR(m_pTextStory))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CTextStory destructor
' ========================================================================================
PRIVATE DESTRUCTOR CTextStory
   CREC_DP("m_pTextStory = " & WSTR(m_pTextStory))
   ' // Release the ITextStory interface
   IF m_pTextStory THEN m_pTextStory->lpvtbl->Release(m_pTextStory)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Assignment operator
' ========================================================================================
PRIVATE OPERATOR CTextStory.LET (BYVAL pTextStory AS ITextStory PTR)
   CREC_DP("m_pTextStory = " & WSTR(m_pTextStory))
   IF pTextStory = NULL THEN m_Result = E_INVALIDARG : EXIT OPERATOR
   ' // Release the interface
   IF m_pTextStory THEN m_pTextStory->lpvtbl->Release(m_pTextStory)
   ' // Attach the passed interface pointer to the class
   m_pTextStory = pTextStory
END OPERATOR
' ========================================================================================

' ========================================================================================
' Cast operator
' ========================================================================================
PRIVATE OPERATOR CTextStory.CAST () AS ITextStory PTR
   OPERATOR = m_pTextStory
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the underlying ITextStory interface
' ========================================================================================
PRIVATE FUNCTION CTextStory.TextStoryPtr () AS ITextStory PTR
   RETURN m_pTextStory
END FUNCTION
' ========================================================================================

' ========================================================================================
' Attaches the passed text document object to the class
' ========================================================================================
PRIVATE FUNCTION CTextStory.Attach (BYVAL pTextStory AS ITextStory PTR, BYVAL fAddRef AS BOOLEAN = FALSE) AS HRESULT
   IF pTextStory = NULL THEN m_Result = E_INVALIDARG : RETURN m_Result
   ' // Release the interface
   IF m_pTextStory THEN m_Result = m_pTextStory->lpvtbl->Release(m_pTextStory)
   ' // Attach the passed interface pointer to the class
   IF fAddRef THEN pTextStory->lpvtbl->AddRef(pTextStory)
   m_pTextStory = pTextStory
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Detaches the interface pointer from the class
' ========================================================================================
PRIVATE FUNCTION CTextStory.Detach () AS ITextStory PTR
   DIM pTextStory AS ITextStory PTR = m_pTextStory
   m_pTextStory = NULL
   RETURN pTextStory
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the active state of a story.
' ========================================================================================
PRIVATE FUNCTION CTextStory.GetActive () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextStory->lpvtbl->GetActive(m_pTextStory, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the active state of a story.
' ========================================================================================
PRIVATE FUNCTION CTextStory.SetActive (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextStory->lpvtbl->SetActive(m_pTextStory, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a new display for a story.
' ========================================================================================
PRIVATE FUNCTION CTextStory.GetDisplay () AS IUnknown PTR
   DIM pUnk AS IUnknown PTR
   m_Result = m_pTextStory->lpvtbl->GetDisplay(m_pTextStory, @pUnk)
   RETURN pUnk
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the index of this story.
' ========================================================================================
PRIVATE FUNCTION CTextStory.GetIndex () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextStory->lpvtbl->GetIndex(m_pTextStory, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets this story's type.
' ========================================================================================
PRIVATE FUNCTION CTextStory.GetType () AS LONG
   DIM Value AS LONG
   m_Result = m_pTextStory->lpvtbl->GetType(m_pTextStory, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the story type.
' ========================================================================================
PRIVATE FUNCTION CTextStory.SetType (BYVAL Value AS LONG) AS HRESULT
   m_Result = m_pTextStory->lpvtbl->SetType(m_pTextStory, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the value of the specified property.
' ========================================================================================
PRIVATE FUNCTION CTextStory.GetProperty (BYVAL nType AS LONG) AS LONG
   DIM Value AS LONG
   m_Result = m_pTextStory->lpvtbl->GetProperty(m_pTextStory, nType, @Value)
   RETURN Value
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a text range object for the story.
' ========================================================================================
PRIVATE FUNCTION CTextStory.GetRange (BYVAL cpActive AS LONG, BYVAL cpAnchor AS LONG) AS ITextRange2 PTR
   DIM pRange AS ITextRange2 PTR
   m_Result = m_pTextStory->lpvtbl->GetRange(m_pTextStory, cpActive, cpAnchor, @pRange)
   RETURN pRange
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the text in a story according to the specified conversion flags.
' ========================================================================================
PRIVATE FUNCTION CTextStory.GetText (BYVAL Flags AS LONG) AS DWSTRING
   DIM bstrText AS AFX_BSTR
   m_Result = m_pTextStory->lpvtbl->GetText(m_pTextStory, Flags, @bstrText)
   DIM dwsText AS DWSTRING = *bstrText
   SysFreeString bstrText
   RETURN dwsText
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces a story’s text with specified formatted text.
' ========================================================================================
PRIVATE FUNCTION CTextStory.SetFormattedText (BYVAL pUnk AS IUnknown PTR) AS HRESULT
   m_Result = m_pTextStory->lpvtbl->SetFormattedText(m_pTextStory, pUnk)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the value of the specified property.
' ========================================================================================
PRIVATE FUNCTION CTextStory.SetProperty (BYVAL nType AS LONG, BYVAL VaLue AS LONG) AS HRESULT
   m_Result = m_pTextStory->lpvtbl->SetProperty(m_pTextStory, nType, Value)
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text in this range.
' ========================================================================================
PRIVATE FUNCTION CTextStory.SetText (BYVAL Flags AS LONG, BYREF wszText AS WSTRING) AS HRESULT
   DIM bstrText AS AFX_BSTR = SysAllocString(wszText)
   m_Result = m_pTextStory->lpvtbl->SetText(m_pTextStory, Flags, bstrText)
   SysFreeString bstrText
   RETURN m_Result
END FUNCTION
' ========================================================================================
