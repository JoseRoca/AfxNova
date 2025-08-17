' ########################################################################################
' Platform: Microsoft Windows
' File: CAdoError.inc
' Contents: ADO errors
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

#pragma once

' ########################################################################################
' CAdoError class
' ########################################################################################
TYPE CAdoError EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pError AS Afx_ADOError PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pError AS Afx_ADOError PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOError PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOError PTR PTR

   ' // Attaches a reference to an Afx_ADOError object to the class
   DECLARE SUB Attach (BYVAL pError AS Afx_ADOError PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Returns a BSTR value that describes an Error object.
   DECLARE PROPERTY Description () AS DWSTRING
   ' // Returns a context ID, as a Long value, for a topic in a Help file.
   DECLARE PROPERTY HelpContext () AS LONG
   ' // Returns a BSTR value that evaluates to a fully resolved path to a Help file.
   DECLARE PROPERTY HelpFile () AS DWSTRING
   ' // Indicates the provider-specific error code for a given Error object.
   DECLARE PROPERTY NativeError () AS LONG
   ' // Indicates the number that uniquely identifies an Error object.
   DECLARE PROPERTY Number () AS LONG
   ' // Indicates the name of the object or application that originally generated an error.
   DECLARE PROPERTY Source () AS DWSTRING
   ' // Indicates the SQL state for a given Error object.
   DECLARE PROPERTY SQLState () AS DWSTRING

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoError
   CADODB_DP("CAdoError CONSTRUCTOR Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoError (BYVAL pError AS Afx_ADOError PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("-BEGIN CAdoError CONSTRUCTOR")
   IF pError THEN m_pError = pError
   IF fAddRef THEN
      IF m_pError THEN m_pError->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoError
   CADODB_DP("CAdoError DESTRUCTOR " & WSTR(m_pError))
   IF m_pError THEN m_pError->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoError.@ () AS Afx_ADOError PTR PTR
'   OPERATOR = @m_pError
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoError.vptr () AS Afx_ADOError PTR PTR
   CADODB_DP("CAdoError vptr")
   IF m_pError THEN m_pError->Release
   RETURN @m_pError
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cError AS CAdoError) AS Afx_ADOError PTR
   OPERATOR = cError.m_pError
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOError object to the class
' ===========================================================================================
PRIVATE SUB CAdoError.Attach (BYVAL pError AS Afx_ADOError PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("CAdoError.Attach")
   IF m_pError THEN m_pError->Release
   m_pError = pError
   ' // Increase the reference count
   IF fAddRef THEN
      IF m_pError THEN m_pError->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Returns a BSTR value that describes an Error object.
' ===========================================================================================
PRIVATE PROPERTY CAdoError.Description () AS DWSTRING
   CADODB_DP("CAdoError.Description")
   DIM bstrDesc AS AFX_BSTR
   m_pError->get_Description(@bstrDesc)
   DIM dwsDesc AS DWSTRING = *bstrDesc
   SysFreeString bstrDesc
   PROPERTY = dwsDesc
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a context ID, as a Long value, for a topic in a Help file.
' ===========================================================================================
PRIVATE PROPERTY CAdoError.HelpContext () AS LONG
   CADODB_DP("CAdoError.HelpContext")
   DIM lHelpContext AS LONG
   m_pError->get_HelpContext(@lHelpContext)
   PROPERTY = lHelpContext
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a BSTR value that describes an Error object.
' ===========================================================================================
PRIVATE PROPERTY CAdoError.HelpFile () AS DWSTRING
   CADODB_DP("CAdoError.HelpFile")
   DIM bstrHelpFile AS AFX_BSTR
   m_pError->get_HelpFile(@bstrHelpFile)
   DIM dwsHelpFile AS DWSTRING = *bstrHelpFile
   SysFreeString bstrHelpFile
   PROPERTY = dwsHelpFile
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a context ID, as a Long value, for a topic in a Help file.
' ===========================================================================================
PRIVATE PROPERTY CAdoError.NativeError () AS LONG
   CADODB_DP("CAdoError.NativeError")
   DIM lNativeError AS LONG
   m_pError->get_NativeError(@lNativeError)
   PROPERTY = lNativeError
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates the number that uniquely identifies an Error object.
' ===========================================================================================
PRIVATE PROPERTY CAdoError.Number () AS LONG
   CADODB_DP("CAdoError.Number")
   DIM lErrorNumber AS LONG
   m_pError->get_Number(@lErrorNumber)
   PROPERTY = lErrorNumber
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Indicates the name of the object or application that originally generated an error.
' ===========================================================================================
PRIVATE PROPERTY CAdoError.Source () AS DWSTRING
   CADODB_DP("CAdoError.Source")
   DIM bstrSource AS AFX_BSTR
   m_pError->get_Source(@bstrSource)
   DIM dwsSource AS DWsTRING = *bstrSource
   SysFreeString bstrSource
   PROPERTY = dwsSource
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' * Indicates the name of the object or application that originally generated an error.
' ===========================================================================================
PRIVATE PROPERTY CAdoError.SQLState () AS DWSTRING
   CADODB_DP("CAdoError.SQLState")
   DIM bstrState AS AFX_BSTR
   m_pError->get_SQLState(@bstrState)
   DIM dwsState AS DWSTRING = *bstrState
   SysFreeString bstrState
   PROPERTY = dwsState
END PROPERTY
' ===========================================================================================


' ########################################################################################
' CAdoErrors class
' ########################################################################################
TYPE CAdoErrors EXTENDS CAdoBase

   m_Result AS HRESULT
   m_pErrors AS Afx_ADOErrors PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pErrors AS Afx_ADOErrors PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   DECLARE DESTRUCTOR
'   DECLARE OPERATOR @ () AS Afx_ADOErrors PTR PTR
   DECLARE FUNCTION vptr () AS Afx_ADOErrors PTR PTR

   ' // Attaches a reference to an Afx_ADOErrors object to the class
   DECLARE SUB Attach (BYVAL pErrors AS Afx_ADOErrors PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   ' // Retrieves the number of errors
   DECLARE PROPERTY Count () AS LONG
   ' // Indicates a specific member of the Errors collection, by name or ordinal number.
   DECLARE PROPERTY Item (BYREF dvIndex AS DVARIANT) AS Afx_ADOError PTR
   ' // Returns a reference to the collection's enumerator.
   DECLARE FUNCTION NewEnum () AS Afx_IUnknown PTR
   ' // Refreshes the contents of the Errors collection.
   DECLARE FUNCTION Refresh () AS HRESULT
   ' // Removes all the Error objects from the Errors collection.
   DECLARE FUNCTION Clear () AS HRESULT

END TYPE
' ########################################################################################

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoErrors
   CADODB_DP("Default")
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Default constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CAdoErrors (BYVAL pErrors AS Afx_ADOErrors PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("pErrors, AddRef")
   IF pErrors THEN
      m_pErrors = pErrors
      IF fAddRef THEN m_pErrors->AddRef
   END IF
END CONSTRUCTOR
' ========================================================================================

' ===========================================================================================
' Cleanup
' ===========================================================================================
PRIVATE DESTRUCTOR CAdoErrors
   CADODB_DP("")
   IF m_pErrors THEN m_pErrors->Release
END DESTRUCTOR
' ===========================================================================================

' ========================================================================================
' Returns the address of the connection pointer
' ========================================================================================
'PRIVATE OPERATOR CAdoErrors.@ () AS Afx_ADOErrors PTR PTR
'   OPERATOR = @m_pErrors
'END OPERATOR
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CAdoErrors.vptr () AS Afx_ADOErrors PTR PTR
   CADODB_DP("")
   IF m_pErrors THEN m_pErrors->Release
   RETURN @m_pErrors
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the encapsulated IDispatch pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF cErrors AS CAdoErrors) AS Afx_ADOErrors PTR
   OPERATOR = cErrors.m_pErrors
END OPERATOR
' ========================================================================================

' ===========================================================================================
' Attaches a reference to an Afx_ADOErrors object to the class
' ===========================================================================================
PRIVATE SUB CAdoErrors.Attach (BYVAL pErrors AS Afx_ADOErrors PTR, BYVAL fAddRef AS BOOLEAN = FALSE)
   CADODB_DP("")
   IF m_pErrors THEN m_pErrors->Release
   m_pErrors = pErrors
   IF fAddRef THEN
      IF m_pErrors THEN m_pErrors->AddRef
   END IF
END SUB
' ===========================================================================================

' ===========================================================================================
' Retrieves the number of errors
' ===========================================================================================
PRIVATE PROPERTY CAdoErrors.Count () AS LONG
   CADODB_DP("")
   DIM nCount AS LONG
   IF m_pErrors THEN m_pErrors->get_Count(@nCount)
   RETURN nCount
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Indicates a specific member of the Errors collection, by name or ordinal number.
' ===========================================================================================
PRIVATE PROPERTY CAdoErrors.Item (BYREF dvIndex AS DVARIANT) AS Afx_ADOError PTR
   CADODB_DP("")
   DIM pField AS Afx_ADOError PTR
   IF m_pErrors THEN m_pErrors->get_Item(dvIndex, @pField)
   RETURN pField
END PROPERTY
' ===========================================================================================

' ===========================================================================================
' Returns a reference to the collection's enumerator.
' ===========================================================================================
PRIVATE FUNCTION CAdoErrors.NewEnum () AS Afx_IUnknown PTR
   CADODB_DP("")
   DIM pUnk AS Afx_IUnknown PTR
   IF m_pErrors THEN m_pErrors->_NewEnum(@pUnk)
   RETURN pUnk
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Refreshes the contents of the Errors collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoErrors.Refresh () AS HRESULT
   CADODB_DP("")
   IF m_pErrors THEN RETURN(m_pErrors->Refresh)
END FUNCTION
' ===========================================================================================

' ===========================================================================================
' Removes all the Error objects from the Errors collection.
' ===========================================================================================
PRIVATE FUNCTION CAdoErrors.Clear () AS HRESULT
   CADODB_DP("")
   IF m_pErrors THEN RETURN(m_pErrors->Clear)
END FUNCTION
' ===========================================================================================
