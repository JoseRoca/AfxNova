' ########################################################################################
' Platform: Microsoft Windows
' Filename: CStack.inc
' Purpose:  Stack and Queue collections
' Compiler: FreeBasic 32 & 64-bit
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
#include once "AfxNova/DSafeArray.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' CStack class
' A Stack Collection is an ordered set of data items, which are accessed on a LIFO
' (Last-In / First-Out) basis. Each data item is passed and stored as a variant variable,
' using the Push and Pop methods.
' ========================================================================================
TYPE CStack

Private:
   DIM m_psa AS DSafeArray PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION Push (BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Pop () AS DVARIANT
   DECLARE FUNCTION Count () AS UINT
   DECLARE FUNCTION Clear () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CStack constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CStack
   ' // Create a safe array of 0 elements and a lower bound of 1
   m_psa = NEW DSafeArray(VT_VARIANT, 0, 1)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CStack destructor
' ========================================================================================
PRIVATE DESTRUCTOR CStack
   Delete m_psa
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Appends a variant at the end of the array.
' ========================================================================================
PRIVATE FUNCTION CStack.Push (BYREF dvData AS DVARIANT) AS HRESULT
   RETURN m_psa->AppendElement(dvData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets and removes the last element of the array.
' ========================================================================================
PRIVATE FUNCTION CStack.Pop () AS DVARIANT
   DIM dv AS DVARIANT
   DIM nPos AS UINT = this.Count
   IF nPos = 0 THEN RETURN dv
   dv = m_psa->GetVar(nPos)
   m_psa->DeleteVariantElement(nPos)
   RETURN dv
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of elements in the collection.
' ========================================================================================
PRIVATE FUNCTION CStack.Count () AS UINT
   RETURN m_psa->Count
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all the elements in the collection.
' ========================================================================================
PRIVATE FUNCTION CStack.Clear () AS HRESULT
   RETURN m_psa->Reset
END FUNCTION
' ========================================================================================

' ========================================================================================
' CQueue class
' A Queue Collection is an ordered set of data items, which are accessed on a FIFO
' (First-In / First-Out) basis. Each data item is passed and stored as a variant variable,
' using the Enqueue and Dequeue methods.
' ========================================================================================
TYPE CQueue

Private:
   DIM m_psa AS DSafeArray PTR

Public:
   DECLARE CONSTRUCTOR
   DECLARE DESTRUCTOR
   DECLARE FUNCTION Enqueue (BYREF dvData AS DVARIANT) AS HRESULT
   DECLARE FUNCTION Dequeue () AS DVARIANT
   DECLARE FUNCTION Count () AS UINT
   DECLARE FUNCTION Clear () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CQueue constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CQueue
   ' // Create a safe array of 0 elements and a lower bound of 1
   m_psa = NEW DSafeArray(VT_VARIANT, 0, 1)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CQueue destructor
' ========================================================================================
PRIVATE DESTRUCTOR CQueue
   Delete m_psa
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Appends a variant at the end of the array.
' ========================================================================================
PRIVATE FUNCTION CQueue.Enqueue (BYREF dvData AS DVARIANT) AS HRESULT
   RETURN m_psa->AppendElement(dvData)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets and removes the first element of the array.
' ========================================================================================
PRIVATE FUNCTION CQueue.Dequeue () AS DVARIANT
   DIM dv AS DVARIANT
   IF this.Count = 0 THEN RETURN dv
   dv = m_psa->GetVar(1)
   m_psa->DeleteVariantElement(1)
   RETURN dv
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the number of elements in the collection.
' ========================================================================================
PRIVATE FUNCTION CQueue.Count () AS UINT
   RETURN m_psa->Count
END FUNCTION
' ========================================================================================

' ========================================================================================
' Removes all the elements in the collection.
' ========================================================================================
PRIVATE FUNCTION CQueue.Clear () AS HRESULT
   RETURN m_psa->Reset
END FUNCTION
' ========================================================================================

END NAMESPACE
