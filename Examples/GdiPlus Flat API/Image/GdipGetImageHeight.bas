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

' ========================================================================================
' The following example retrieves the height of an image.
' ========================================================================================

DIM hStatus AS LONG

' // Initialize GDI+
DIM token AS ULONG_PTR = AfxGdipInit

' // Load the original image from file
DIM image AS GpImage PTR
hStatus = GdipLoadImageFromFile("climber.jpg", @image)

' // Get the width of the image
DIM nHeight AS UINT
hStatus = GdipGetImageHeight(image, @nHeight)
PRINT "Image height: " & WSTR(nHeight)

' // Cleanup
IF image THEN GdipDisposeImage(image)

' // Shutdown GDI+
AfxGdipShutdown token

PRINT
PRINT "Press any key"
SLEEP


' ========================================================================================
