' ########################################################################################
' Microsoft Windows
' File: ImageGetHeight.bas
' Contents: GDI+ Flat API - ImageGetHeight example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'#CONSOLE ON
#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/CGdiPlus.inc"
USING AfxNova

' ========================================================================================
' The following example retrieves the height of an image.
' ========================================================================================

' // Create a Bitmap object from a JPEG file.
DIM image AS CGpImage = "climber.jpg"

' // Get the height of the image
DIM nHeight AS UINT = image.GetHeight
PRINT "Image height: " & WSTR(nHeight)

PRINT
PRINT "Press any key"
SLEEP

' ========================================================================================
