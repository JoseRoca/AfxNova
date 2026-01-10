' ########################################################################################
' Microsoft Windows
' File: GdipGetImageItemData.bas
' Contents: GDI+ Flat API - GdipGetImageItemData example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

' ========================================================================================
' The following example gets one piece of metadata from an Image object.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Load the image file
DIM image AS GpImage PTR
hStatus = GdipLoadImageFromFile("RYDER_Winona_01.jpg", @image)

DIM item AS ImageItemData
item.Size = SIZEOF(ImageItemData)
item.Data = NULL
item.DataSize = 0

' First call: get required size
hStatus = GdipGetImageItemData(image, @item)
IF hStatus = 0 THEN
   item.Data = Allocate(item.DataSize)
   ' Second call: retrieve actual data
   hStatus = GdipGetImageItemData(image, @item)
   IF hStatus = 0 THEN
      PRINT "Metadata retrieved, size = "; item.DataSize
      ' You can now parse item.Data
   ELSE
      PRINT "Failed to retrieve metadata"
   END IF
   IF item.Data THEN Deallocate(item.Data)
ELSE
   PRINT "Failed to get metadata size"
END IF

' Allocate buffers if needed
item.Desc = Allocate(item.DescSize)
item.Data = Allocate(item.DataSize)

' Retrieve the actual data
hStatus = GdipGetImageItemData(image, @item)

' Use item.Data and item.Desc as needed...

' Cleanup
IF item.Desc THEN Deallocate(item.Desc)
IF item.Data THEN Deallocate(item.Data)
IF image THEN GdipDisposeImage(image)

'Size	UINT	Integer that specifies the size of an ImageItemData object. Set to sizeof(ImageItemData).
'Position	UINT	Element of the ItemDataPosition enumeration that specifies the location in an image file that a piece of custom metadata should be stored.
'Desc	VOID *	A sequence of bytes that you create to identify a piece of custom metadata. For JPEG files, the description is one byte. For PNG files, the description is four bytes. For GIF files, the description is eleven bytes.
'DescSize	UINT	Integer that specifies the size, in bytes, of the identifier pointed to by Desc. Set to 1 for JPEG, 4 for PNG, and 11 for GIF.
'Data	UINT *	Pointer to a buffer that contains the metadata.
'DataSize	UINT	Integer that specifies the size, in bytes, of the buffer pointed to by Data.
'Cookie	UINT	Used internally by GDI+.

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP

' ========================================================================================
