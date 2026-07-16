' ########################################################################################
' Microsoft Windows
' Compiler: FreeBasic 32 bit
' Copyright (c) 2026 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#cmdline "-s console"
'#define _CADODB_DEBUG_ 1
'#define _CADOX_DEBUG_ 1
#include once "AfxNova/CADODB.inc"
#include once "AfxNova/CADOX.inc"
USING AfxNova

DO   ' // Using a fake loop to avoid the use of GOTO or nested IFs/END IFs

' // Open the connection
DIM pConnection AS CADOConnection
IF *pConnection = NULL THEN EXIT DO
pConnection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb")

' // Set the ActiveConnection property of the Catalog
DIM pCatalog AS CADOXCatalog
pCatalog.ActiveConnection = *pConnection

' // Create the command representing the View
DIM pCommand AS CADOCommand
IF *pCommand = NULL THEN EXIT DO
pCommand.CommandText = "SELECT * FROM Authors"

' // Get a reference to the Views collection
DIM pViews AS CADOXViews = pCatalog.Views
IF *pViews = NULL THEN EXIT DO

' // Append the View to the collection
IF pViews.Append("AllAuthors", *pCommand) = S_OK THEN
   ? "View created"
ELSE
   ? "Failure"
END IF

' // Close the connection
' // If you don't close it, they will be closed when the application ends
pConnection.Close

EXIT DO' // Unconditional exit of the fake loop
LOOP

PRINT
PRINT "Press any key..."
SLEEP
