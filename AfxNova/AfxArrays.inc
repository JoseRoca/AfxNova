' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxArrays.inc
' Purpose:  Macros to manipulate one-dimensional dynamic arrays.
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
#include once "/fbc-int/array.bi"
USING FBC

' ########################################################################################
'                                     ARRAY MACROS
' ########################################################################################

' ========================================================================================
' Appends a new element to a dynamic one-dimensional array.
' The array and the element to append can be of any type.
' Parameters:
' - rg:   The array
' - elem: The element to append
' Usage example:
' #define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
' DIM rg(ANY) AS XSTRING
' DIM xStr AS XSTRING = "String - "
' FOR i AS LONG = 1 TO 10
'    AppendElementToArray(rg, xStr & WSTR(i))
' NEXT
' Display the array
' FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'    print rg(i)
' NEXT
' ========================================================================================
#macro AppendElementToArray(rg, elem)
SCOPE
   DO   ' // Using DO...LOOP to allow an early exit
      ' // Check if the passed array is a fixed-length array
      DIM pDesc AS FBC.FBARRAY PTR = ArrayDescriptorPtr(rg())
      IF pDesc = NULL THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      IF (pDesc->flags AND FBARRAY_FLAGS_FIXED_LEN) <> 0 THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      ' // Redim the array
      REDIM PRESERVE rg(LBOUND(rg) TO UBOUND(rg) + 1) AS TYPEOF(rg)
      ' // Assign the passed string
      rg(UBOUND(rg)) = elem
      ' // Return success
      SetLastError(S_OK)
      ' // Exit the loop
      EXIT DO
   LOOP
END SCOPE
#endmacro
' ========================================================================================

' ========================================================================================
' Appends a one-dimensional array to another dynamic one-dimensional array.
' The array and the to append can be of any type.
' Parameters:
' - rg  : The array
' - elem: The array to append
' Usage example:
'   #define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
'   DIM rg(ANY) AS XSTRING
'   DIM rg2(ANY) AS XSTRING
'   DIM xStr AS XSTRING = "String - "
'   ' // Fill the array
'   FOR i AS LONG = 1 TO 10
'      AppendElementToArray(rg, xStr & WSTR(i))
'   NEXT
'   ' Fill the array to append
'   xStr = "String 2 - "
'   FOR i AS LONG = 1 TO 5
'      AppendElementToArray(rg2, xStr & WSTR(i))
'   NEXT
'   ' // Append the array to the first array
'   AppendArrayToArray(rg, rg2)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' Can also be used with numbers:
'   DIM rg(ANY) AS LONG
'   DIM rg2(ANY) AS LONG
'   ' // Fill the array
'   DIM nLong AS LONG = 1
'   FOR i AS LONG = 1 TO 10
'      AppendElementToArray(rg, nLong)
'      nLong += 1
'   NEXT
'   ' Fill the array to insert
'   nLong = 12345
'   FOR i AS LONG = 1 TO 5
'      AppendElementToArray(rg2, nLong)
'      nLong += 1
'   NEXT
'   AppendArrayToArray(rg, rg2)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' ========================================================================================
#macro AppendArrayToArray(rg, rg2)
SCOPE
   DO   ' // Using DO...LOOP to allow an early exit
      ' // Check if the passed array is a fixed-length array
      DIM pDesc AS FBC.FBARRAY PTR = ArrayDescriptorPtr(rg())
      IF pDesc = NULL THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      IF (pDesc->flags AND FBARRAY_FLAGS_FIXED_LEN) <> 0 THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      ' // Redim the array
      DIM targetUpper AS LONG = UBOUND(rg)
      DIM sourceUpper AS LONG = UBOUND(rg2)
      REDIM PRESERVE rg(LBOUND(rg) TO targetUpper + sourceUpper + 1) AS TYPEOF(rg)
      targetUpper += 1
      FOR i AS LONG = LBOUND(rg2) TO sourceUpper
         rg(targetUpper + i -LBOUND(rg2)) = rg2(i)
      NEXT
      ' // Return success
      SetLastError(S_OK)
      EXIT DO
   LOOP
END SCOPE
#endmacro
' ========================================================================================

' ========================================================================================
' Inserts a new element before the specified position in a dynamic one-dimensional array.
' Parameters:
' - rg:   The array
' - pos:  The position in the array where the new element will be added.
'         This position is relative to the lower bound of the array.
' - elem: The element to insert
' Usage example:
'   #define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
'   DIM rg(ANY) AS XSTRING
'   DIM xStr AS XSTRING = "String - "
'   ' // Fill the array
'   FOR i AS LONG = 1 TO 10
'      AppendElementToArray(rg, xStr & WSTR(i))
'   NEXT
'   InsertElementIntoArray(rg, 5, xStr & WSTR(11))
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' Can also be used with numbers:
'   DIM rg(ANY) AS LONG
'   DIM rg2(ANY) AS LONG
'   ' // Fill the array
'   DIM nLong AS LONG = 1
'   FOR i AS LONG = 1 TO 10
'      AppendElementToArray(rg, nLong)
'      nLong += 1
'   NEXT
'   ' Fill the array to insert
'   nLong = 12345
'   FOR i AS LONG = 1 TO 5
'      AppendElementToArray(rg2, nLong)
'      nLong += 1
'   NEXT
'   InsertElementIntoArray(rg, 5, nLong)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' ========================================================================================
#macro InsertElementIntoArray(rg, pos, elem)
SCOPE
   DO   ' // Using DO...LOOP to allow an early exit
      ' // Check if the passed array is a fixed-length array
      DIM pDesc AS FBC.FBARRAY PTR = ArrayDescriptorPtr(rg())
      IF pDesc = NULL THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      IF (pDesc->flags AND FBARRAY_FLAGS_FIXED_LEN) <> 0 THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      DIM lowerBound AS LONG = LBOUND(rg)
      DIM upperBound AS LONG = UBOUND(rg)
      DIM nPos AS LONG = pos   ' // We need to use a variable with numbers
      nPos = nPos - 1 + lowerBound
      IF nPos < lowerBound OR nPos > upperBound THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      ' // Redim the array
      REDIM PRESERVE rg(lowerBound TO upperBound + 1) AS TYPEOF(rg)
      ' // Move all the elements down
      FOR i AS LONG = UBOUND(rg) TO nPos + 1 STEP - 1
         SWAP rg(i), rg(i - 1)
      NEXT
      ' // Assign the element
      rg(nPos) = elem
      ' // Return success
      SetLastError(S_OK)
      EXIT DO
   LOOP
END SCOPE
#endmacro

' ========================================================================================
' Inserts a one-dimensional array before the specified position in another dynamic one-dimensional array.
' Parameters:
' - rg:  The array
' - pos: The position in the array where the second array will be added.
'        This position is relative to the lower bound of the array.
' - rg2: The array to insert
' Usage example:
'   #define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
'   DIM rg(ANY) AS XSTRING
'   DIM rg2(ANY) AS XSTRING
'   DIM xStr AS XSTRING = "String - "
'   ' // Fill the array
'   FOR i AS LONG = 1 TO 10
'      AppendElementToArray(rg, xStr & WSTR(i))
'   NEXT
'   ' Fill the array to insert
'   xStr = "String 2 - "
'   FOR i AS LONG = 1 TO 5
'      AppendElementToArray(rg2, xStr & WSTR(i))
'   NEXT
'   ' // Insert the array to the first array
'   InsertArrayIntoArray(rg, 5, rg2)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' Can also be used with numbers:
'   DIM rg(ANY) AS LONG
'   DIM rg2(ANY) AS LONG
'   ' // Fill the array
'   DIM nLong AS LONG = 1
'   FOR i AS LONG = 1 TO 10
'      AppendElementToArray(rg, nLong)
'      nLong += 1
'   NEXT
'   ' Fill the array to insert
'   nLong = 12345
'   FOR i AS LONG = 1 TO 5
'      AppendElementToArray(rg2, nLong)
'      nLong += 1
'   NEXT
'   ' // Insert the array to the first array
'   InsertArrayIntoArray(rg, 5, rg2)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' ========================================================================================
#macro InsertArrayIntoArray(rg, pos, rg2)
SCOPE
   DO   ' // Using DO...LOOP to allow an early exit
      ' // Check if the passed array is a fixed-length array
      DIM pDesc AS FBC.FBARRAY PTR = ArrayDescriptorPtr(rg())
      IF pDesc = NULL THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      IF (pDesc->flags AND FBARRAY_FLAGS_FIXED_LEN) <> 0 THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      DIM lowerBound AS LONG = LBOUND(rg)
      DIM upperBound AS LONG = UBOUND(rg)
      DIM nPos AS LONG = pos   ' // We need to use a variable with numbers
      nPos = nPos - 1 + lowerBound
      IF nPos < lowerBound OR nPos > upperBound THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      DIM targetUpper AS LONG = UBOUND(rg)
      DIM sourceUpper AS LONG = UBOUND(rg2)
      ' // Preallocate space in one step
      REDIM PRESERVE rg(LBOUND(rg) TO targetUpper + sourceUpper + 1) AS TYPEOF(rg)
      ' // Move all the elements down
      FOR i AS LONG = targetUpper + 1 TO nPos + 1 STEP - 1
         SWAP rg(i + sourceUpper), rg(i - 1)
      NEXT
      ' // Copy the array into position
      FOR i AS LONG = LBOUND(rg2) TO sourceUpper
         rg(nPos + i) = rg2(i)
      NEXT
      ' // Return success
      SetLastError(S_OK)
      EXIT DO
   LOOP
END SCOPE
#endmacro
' ========================================================================================

' ========================================================================================
' Removes the specified element of a dynamic one-dimensional array.
' Parameters:
' - rg:  The array
' - pos: The position in the array of the element to remove.
'        This position is relative to the lower bound of the array.
' Example:
'   #define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
'   DIM rg(ANY) AS XSTRING
'   DIM xStr AS XSTRING = "String - "
'   ' // Fill the array
'   FOR i AS LONG = 1 TO 10
'      AppendToArray(rg, xStr & WSTR(i))
'   NEXT
'   ' // Remove the fifth element
'   RemoveElementFromArray(rg, 5)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' Can also be used with numbers:
'   DIM rg(ANY) AS LONG
'   ' // Fill the array
'   DIM nLong AS LONG = 1
'   FOR i AS LONG = 1 TO 10
'      REDIM PRESERVE rg(UBOUND(rg) + 1)
'      rg(i - 1) = nLong
'      nLong += 1
'   NEXT
'   ' // Remove the specified element
'   RemoveElementFromArray(rg, 5)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' ========================================================================================
#macro RemoveElementFromArray(rg, pos)
SCOPE
   DO   ' // Using DO...LOOP to allow an early exit
      ' // Check if the passed array is a fixed-length array
      DIM pDesc AS FBC.FBARRAY PTR = ArrayDescriptorPtr(rg())
      IF pDesc = NULL THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      IF (pDesc->flags AND FBARRAY_FLAGS_FIXED_LEN) <> 0 THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      ' // Redim the array
      DIM lowerBound AS LONG = LBOUND(rg)
      DIM upperBound AS LONG = UBOUND(rg)
      DIM nPos AS LONG = pos - 1 + lowerBound
      IF nPos < lowerBound OR nPos > upperBound THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      FOR i AS LONG = nPos TO upperBound - 1
         rg(i) = rg(i + 1)  ' Direct copy instead of swap
      NEXT
      REDIM PRESERVE rg(lowerBound TO upperBound - 1) AS TYPEOF(rg)
      ' // Return success
      SetLastError(S_OK)
      ' // Exit the loop
      EXIT DO
   LOOP
END SCOPE
#endmacro
' ========================================================================================

' ========================================================================================
' Removes the first element of a dynamic one-dimensional array.
' Parameters:
' - rg:  The array
' Usage example:
'   #define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
'   DIM rg(ANY) AS XSTRING
'   DIM xStr AS XSTRING = "String - "
'   ' // Fill the array
'   FOR i AS LONG = 1 TO 10
'      AppendElementToArray(rg, xStr & WSTR(i))
'   NEXT
'   ' // Remove the frst element
'   RemoveFirstElementFromgArray(rg)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' Can also be used with numbers:
'   DIM rg(ANY) AS LONG
'   ' // Fill the array
'   DIM nLong AS LONG = 12345
'   FOR i AS LONG = 1 TO 10
'      REDIM PRESERVE rg(UBOUND(rg) + 1)
'      rg(i - 1) = nLong
'      nLong += 1
'   NEXT
'   ' // Remove the first element
'   RemoveFirstElementFromArray(rg)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' ========================================================================================
#macro RemoveFirstElementFromArray(rg)
SCOPE
   DO   ' // Using DO...LOOP to allow an early exit
      ' // Check if the passed array is a fixed-length array
      DIM pDesc AS FBC.FBARRAY PTR = ArrayDescriptorPtr(rg())
      IF pDesc = NULL THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      IF (pDesc->flags AND FBARRAY_FLAGS_FIXED_LEN) <> 0 THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      ' // Redim the array
      DIM lowerBound AS LONG = LBOUND(rg)
      DIM upperBound AS LONG = UBOUND(rg)
      DIM nPos AS LONG = lowerBound
      FOR i AS LONG = nPos TO upperBound - 1
         rg(i) = rg(i + 1)
      NEXT
      REDIM PRESERVE rg(lowerBound TO upperBound - 1) AS TYPEOF(rg)
      ' // Return success
      SetLastError(S_OK)
      ' // Exit the loop
      EXIT DO
   LOOP
END SCOPE
#endmacro
' ========================================================================================

' ========================================================================================
' Removes the last element of a dynamic one-dimensional array.
' Parameters:
' - rg : The array
' Usage example:
'   #define XSTRING DWSTRING ' // or STRING, BSTRING, etc.
'   DIM rg(ANY) AS XSTRING
'   DIM xStr AS XSTRING = "String - "
'   ' // Fill the array
'   FOR i AS LONG = 1 TO 10
'      AppendElementToArray(rg, xStr & WSTR(i))
'   NEXT
'   ' // Remove the last element
'   RemoveLastElementFromArray(rg)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' Can also be used with numbers:
'   DIM rg(ANY) AS LONG   ' // or DWORD, SINGLE, DOUBLE, etc.
'   ' // Fill the array
'   DIM nLong AS LONG = 12345
'   FOR i AS LONG = 1 TO 10
'      REDIM PRESERVE rg(UBOUND(rg) + 1)
'      rg(i - 1) = nLong
'      nLong += 1
'   NEXT
'   ' // Remove the last element
'   RemoveLastElementFromArray(rg)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' ========================================================================================
#macro RemoveLastElementFromArray(rg)
SCOPE
   DO   ' // Using DO...LOOP to allow an early exit
      ' // Check if the passed array is a fixed-length array
      DIM pDesc AS FBC.FBARRAY PTR = ArrayDescriptorPtr(rg())
      IF pDesc = NULL THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      IF (pDesc->flags AND FBARRAY_FLAGS_FIXED_LEN) <> 0 THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      ' // Redim the array
      DIM lowerBound AS LONG = LBOUND(rg)
      DIM upperBound AS LONG = UBOUND(rg)
      REDIM PRESERVE rg(lowerBound TO upperBound - 1) AS TYPEOF(rg)
      ' // Return success
      SetLastError(S_OK)
      ' // Exit the loop
      EXIT DO
   LOOP
END SCOPE
#endmacro
' ========================================================================================

' ========================================================================================
' Removes the specified elements of a dynamic one-dimensional array.
' Parameters:
' - rg : The array
' - rgElements: Array of longs indicating the 0-based positions of the elements to remove.
' - res: The result code. A boolean true of false value.
' Usage example:
'   DIM rg(ANY) AS LONG
'   ' // Fill the array
'   DIM nLong AS LONG = 1
'   FOR i AS LONG = 1 TO 10
'      AppendElementToArray(rg, nLong)
'      nLong += 1
'   NEXT
'   ' Fill the array of elements to delete
'   DIM rgElements(0 TO 3) AS LONG => {2, 4, 6, 8}
'   RemoveElementsFromArray(rg, rgElements)
'   ' // Display the array
'   FOR i AS LONG = LBOUND(rg) TO UBOUND(rg)
'      print rg(i)
'   NEXT
' ========================================================================================
#macro RemoveElementsFromArray(rg, rgElements)
SCOPE
   DO   ' // Using DO...LOOP to allow an early exit
      ' // Check if the passed array is a fixed-length array
      DIM pDesc AS FBC.FBARRAY PTR = ArrayDescriptorPtr(rg())
      IF pDesc = NULL THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      IF (pDesc->flags AND FBARRAY_FLAGS_FIXED_LEN) <> 0 THEN SetLastError(E_ACCESSDENIED) : EXIT DO
      DIM upperBound AS LONG = UBOUND(rg)
      DIM writeIndex AS LONG = LBOUND(rg)
      ' // Mark elements to be removed
      DIM rgRemove(UBOUND(rg)) AS BOOLEAN
      FOR i AS LONG = LBOUND(rgElements) TO UBOUND(rgElements)
         DIM index AS LONG = rgElements(i)
         IF index >= LBOUND(rg) AND index <= upperBound THEN
            rgRemove(index) = TRUE
         END IF
      NEXT
      ' // Shift elements down, skipping removed ones
      FOR i AS LONG = LBOUND(rg) TO upperBound
         IF NOT rgRemove(i) THEN
            rg(writeIndex) = rg(i)
            writeIndex += 1
         END IF
      NEXT
      ' // Resize array to the new size
      REDIM PRESERVE rg(LBOUND(rg) TO writeIndex - 1) AS TYPEOF(rg)
      ' // Return success
      SetLastError(S_OK)
      ' // Exit the loop
      EXIT DO
   LOOP
END SCOPE
#endmacro
