' ########################################################################################
' Microsoft Windows
' File: GdipGetImageGraphicsContext.bas
' Contents: GDI+ Flat API - GdipGetImageGraphicsContext example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.inc"

TYPE GpImageStruct
   AS BYTE PTR decoder  ' AS IWICBitmapDecoder PTR decoder
   AS ImageType type
   AS GUID format
   AS UINT flags
   AS UINT frame_count
   AS UINT current_frame
   AS ColorPalette PTR palette
   AS REAL xres
   AS REAL yres
   AS LONG busy
END TYPE

TYPE GpBitmapStruct
   AS GpImageStruct image
   AS INT_ width
   AS INT_ height
   AS PixelFormat format
   AS ImageLockMode lockmode
   AS BYTE PTR bitmapbits   ' /* pointer to the buffer we passed in BitmapLockBits */
   AS HBITMAP hbitmap
   AS HDC hdc
   AS BYTE PTR bits   ' /* actual image bits if this is a DIB */
   AS INT_ stride   ' /* stride of bits if this is a DIB */
   AS BYTE PTR own_bits   ' /* image bits that need to be freed with this object */
   AS INT_ lockx
   AS INT_ locky   ' /* X and Y coordinates of the rect when a bitmap is locked for writing. */
   AS BYTE PTR metadata_reader  'AS IWICMetadataReader PTR metadata_reader   ' /* NULL if there is no metadata */
   AS UINT prop_count
   AS PropertyItem PTR prop_item   ' /* cached image properties */
END TYPE

' ========================================================================================
' The following example retrieves the height of an image.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Load the original image from file
'DIM image AS GpImage PTR
'hStatus = GdipLoadImageFromFile("climber.jpg", @image)

DIM img AS GpBitmapStruct
hStatus = GdipLoadImageFromFile("climber.jpg", @img)
print "hStatus: ", hStatus, img

' // Get the width of the image
'DIM nHeight AS UINT
'hStatus = GdipGetImageHeight(image, @nHeight)
'PRINT "Image height: " & WSTR(nHeight)

' PRINT "Image height: " & WSTR(img->height)

' // Cleanup
IF img THEN GdipDisposeImage(@img)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP


' ========================================================================================
