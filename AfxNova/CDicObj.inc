' ########################################################################################
' Platform: Microsoft Windows
' Filename: CDicObj.inc
' Purpose:  Wrapper class on top of the VBString Regular Expressions engine.
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
#include once "windows.bi"
#include once "win/ole2.bi"
#include once "AfxNova/AfxScrRun.bi"
#include once "AfxNova/DWSTRING.inc"
#include once "AfxNova/DVARIANT.inc"

NAMESPACE AfxNova

' ========================================================================================
' Macro for debug
' To allow debugging, define _CDICOBJ_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _CDICOBJ_DEBUG_
   #define _CDICOBJ_DEBUG_ 0
#ENDIF
#ifndef _CDICOBJ_DP_
   #define _CDICOBJ_DP_ 1
   #MACRO CDICOBJ_DP(st)
      #IF (_CDICOBJ_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

' ========================================================================================
' CDicObj class
' ========================================================================================
TYPE CDicObj

Public:
   m_Result AS HRESULT
   m_pDictionary AS Afx_IDictionary PTR

Private:
   m_bUninitCOM AS BOOLEAN

Public:
   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYREF pDic AS CDicObj)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION DispPtr () AS ANY PTR
   DECLARE FUNCTION DispObj () AS ANY PTR
   DECLARE FUNCTION GetLastResult () AS HRESULT
   DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
   DECLARE FUNCTION GetErrorInfo (BYVAL dwError AS LONG = -1) AS DWSTRING
   DECLARE PROPERTY Item (BYREF dvKey AS DVARIANT, BYREF dvNewItem AS DVARIANT)
   DECLARE PROPERTY Item (BYREF dvKey AS DVARIANT) AS DVARIANT
   DECLARE FUNCTION Add (BYREF dvKey AS DVARIANT, BYREF dvItem AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Count () AS LONG
   DECLARE FUNCTION Exists (BYREF dvKey AS DVARIANT) AS BOOLEAN
   DECLARE FUNCTION Items () AS DVARIANT
   DECLARE PROPERTY Key (BYREF dvKey AS DVARIANT, BYREF dvNewKey AS DVARIANT)
   DECLARE FUNCTION Keys () AS DVARIANT
   DECLARE FUNCTION Remove (BYREF dvKey AS DVARIANT) AS HRESULT
   DECLARE FUNCTION RemoveAll() AS HRESULT
   DECLARE FUNCTION NewEnum () AS IEnumVARIANT PTR
   DECLARE FUNCTION HashVal (BYREF dvKey AS DVARIANT) AS DVARIANT

END TYPE
' ========================================================================================

' ========================================================================================
' CDicObj constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CDicObj
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Dictionary object
   DIM CLSID_Dictionary_ AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   DIM IID_IDictionary_  AS GUID  = (&h42C642C1, &h97E1, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   CoCreateInstance(@CLSID_Dictionary_, NULL, CLSCTX_INPROC_SERVER, @IID_IDictionary_, @m_pDictionary)
   CDICOBJ_DP("Default - m_pDictionary = " & WSTR(m_pDictionary))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Creates a new instance of the Dictionary object and copies the contents of the
' passed CDicObj.
' ========================================================================================
PRIVATE CONSTRUCTOR CDicObj (BYREF pDic AS CDicObj)
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the Dictionary object
   DIM CLSID_Dictionary_ AS CLSID = (&hEE09B103, &h97E0, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   DIM IID_IDictionary_  AS GUID  = (&h42C642C1, &h97E1, &h11CF, {&h97, &h8F, &h00, &hA0, &h24, &h63, &hE0, &h6F})
   CoCreateInstance(@CLSID_Dictionary_, NULL, CLSCTX_INPROC_SERVER, @IID_IDictionary_, @m_pDictionary)
   ' // Copy the contents of the passed CDicObj
   DIm dvKey AS DVARIANT, dvValue AS DVARIANT
   DIM dvKeys AS DVARIANT = pDic.Keys
   FOR i AS LONG = dvKeys.GetLBound TO dvKeys.GetUBound
      dvKey = dvKeys.GetVariantElement(i)
      dvValue = pDic.Item(dvKey)
      pDic.Add(dvKey, dvValue)
   NEXT
   CDICOBJ_DP("pDic = " & WSTR(m_pDictionary) & " - m_pDictionary = " & WSTR(m_pDictionary))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CDicObj destructor
' ========================================================================================
PRIVATE DESTRUCTOR CDicObj
   CDICOBJ_DP("m_pDictionary = " & WSTR(m_pDictionary))
   ' // Release the Dictionary object
   IF m_pDictionary THEN
      m_pDictionary->RemoveAll
      m_pDictionary->Release
   END IF
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the underlying dispatch pointer.
' Don't call IUnknown_Release on it.
' ========================================================================================
PRIVATE FUNCTION CDicObj.DispPtr () AS ANY PTR
   RETURN m_pDictionary
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a counted reference of the underlying dispatch pointer.
' You must call IUnknown_Release when no longer need it.
' Note: It is not possible to clone the object this way because the destructor of the
' class removes all the contents.
' ========================================================================================
PRIVATE FUNCTION CDicObj.DispObj () AS ANY PTR
   IF m_pDictionary <> NULL THEN
      m_pDictionary->AddRef
      RETURN m_pDictionary
   END IF
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CDicObj.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CDicObj.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CDicObj.GetErrorInfo (BYVAL nError AS LONG = -1) AS DWSTRING
   IF nError = -1 THEN nError = m_Result
   DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
   cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
           FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
           NULL, nError, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
           cast(LPWSTR, @pBuffer), 0, NULL)
   IF cbLen THEN
      dwsMsg = *pBuffer
      LocalFree pBuffer
   END IF
   IF nError THEN dwsMsg = "Error &h" & HEX(nError) & CHR(13, 10) & dwsMsg
   RETURN dwsMsg
END FUNCTION
' ========================================================================================

' ========================================================================================
' Item property
' Sets or returns an item for a specified dvKey in the associative array.
' Arguments:
' - dvKey     : Key associated with the item being retrieved or added.
' - dvNewItem : The new value associated with the specified key.
' Remarks:
' If key is not found when changing an item, a new key is created with the specified dvNewItem.
' ========================================================================================
PRIVATE PROPERTY CDicObj.Item (BYREF dvKey AS DVARIANT, BYREF dvNewItem AS DVARIANT)
   CDICOBJ_DP("SET Item")
   m_Result = m_pDictionary->put_Item(dvKey, dvNewItem)
END PROPERTY
' ========================================================================================
' ========================================================================================
PRIVATE PROPERTY CDicObj.Item (BYREF dvKey AS DVARIANT) AS DVARIANT
   CDICOBJ_DP("GET Item")
   DIM dvItem AS DVARIANT
   IF this.Exists(dvKey) = FALSE THEN SetResult(E_INVALIDARG) : RETURN dvItem
   m_Result = m_pDictionary->get_Item(dvKey, dvItem)
   RETURN dvItem
END PROPERTY
' ========================================================================================

' ========================================================================================
' Add method
' Adds a key and item pair to the associative array.
' Arguments:
' - dvKey  : The key associated with the item being added.
' - dvItem : The item associated with the key being added.
' Remarks:
' An error occurs if the key already exists.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Add (BYREF dvKey AS DVARIANT, BYREF dvItem AS DVARIANT) AS HRESULT
   CDICOBJ_DP("Add")
   RETURN SetResult(m_pDictionary->Add(dvKey, dvItem))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Count property
' Returns the number of items in the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Count () AS LONG
   CDICOBJ_DP("Count")
   DIM nCount AS LONG
   m_Result = m_pDictionary->get_Count(@nCount)
   RETURN nCount
END FUNCTION
' ========================================================================================

' ========================================================================================
' Exists method
' Returns TRUE if a specified key exists in the associative array, FALSE if it does not.
' Argument:
' - dvKey : Key value being searched for the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Exists (BYREF dvKey AS DVARIANT) AS BOOLEAN
   CDICOBJ_DP("Exists")
   DIM nExists AS SHORT
   m_Result = m_pDictionary->Exists(dvKey, @nExists)
   RETURN nExists
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns an array containing all the items in a associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Items () AS DVARIANT
   CDICOBJ_DP("Items")
   DIM dvItems AS DVARIANT
   m_Result = m_pDictionary->Items(dvItems)
   RETURN dvItems
END FUNCTION
' ========================================================================================

' ========================================================================================
' Key property
' Arguments:
' - dvKey  : Key value being changed.
' - dvItem : New value that replaces the specified key.
' Usage example: pDic.Key("a") = "m"
' ========================================================================================
PRIVATE PROPERTY CDicObj.Key (BYREF dvKey AS DVARIANT, BYREF dvNewKey AS DVARIANT)
   CDICOBJ_DP("SET Key")
   IF this.Exists(dvKey) = FALSE THEN SetResult(E_INVALIDARG) : EXIT PROPERTY
   m_Result = m_pDictionary->put_Key(dvKey, dvNewKey)
END PROPERTY
' ========================================================================================

' ========================================================================================
' Keys method
' Returns an array containing all existing keys in the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Keys () AS DVARIANT
   CDICOBJ_DP("Keys")
   DIM dvKeys AS DVARIANT
   m_Result = m_pDictionary->Keys(dvKeys)
   RETURN dvKeys
END FUNCTION
' ========================================================================================

' ========================================================================================
' Remove method
' Removes a key, item pair from the associative array.
' Argument:
' - dvKey : Key associated with the key, item pair you want to remove from the associative array.
' Remarks:
' An error occurs if the specified key, item pair does not exist.
' ========================================================================================
PRIVATE FUNCTION CDicObj.Remove (BYREF dvKey AS DVARIANT) AS HRESULT
   CDICOBJ_DP("Remove")
   RETURN SetResult(m_pDictionary->Remove(dvKey))
END FUNCTION
' ========================================================================================

' ========================================================================================
' RemoveAll method
' Removes all key, item pairs from the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.RemoveAll() AS HRESULT
   CDICOBJ_DP("RemoveAll")
   RETURN SetResult(m_pDictionary->RemoveAll)
END FUNCTION
' ========================================================================================

' ========================================================================================
' NewEnum method
' Returns a reference to the standard IEnumVARIANT interface.
' ========================================================================================
PRIVATE FUNCTION CDicObj.NewEnum () AS IEnumVARIANT PTR
   CDICOBJ_DP("NewEnum")
   DIM ppUnk AS IUnknown PTR
   m_Result = m_pDictionary->get__NewEnum(cast(ANY PTR, @ppunk))
   RETURN CAST(IEnumVARIANT PTR, ppunk)
END FUNCTION
' ========================================================================================

' ========================================================================================
' HashVal property
' Returns the hash value for a specified key in the associative array.
' ========================================================================================
PRIVATE FUNCTION CDicObj.HashVal (BYREF dvKey AS DVARIANT) AS DVARIANT
   CDICOBJ_DP("HashVal")
   DIM dvHashVal AS DVARIANT
   m_Result = m_pDictionary->get_HashVal(dvKey, dvHashVal)
   RETURN dvHashVal
END FUNCTION
' ========================================================================================

END NAMESPACE
