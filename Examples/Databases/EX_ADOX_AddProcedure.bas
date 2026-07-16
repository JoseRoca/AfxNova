' ########################################################################################
' Microsoft Windows
' Compiler: FreeBasic32 bit
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

' // Create the command representing the procedure
DIM pCommand AS CADOCommand
IF *pCommand = NULL THEN EXIT DO

' // Create the parameterized command (Microsoft Jet specific)
pCommand.ActiveConnection = *pConnection
DIM dwsCommandText AS DWSTRING = "PARAMETERS [AuthorId] Text; " & _
                                 "SELECT * FROM Authors WHERE Aud_ID = [AuthorId]"
pCommand.CommandText = dwsCommandText

' // Set the ActiveConnection property of the Catalog
DIM pCatalog AS CADOXCatalog
pCatalog.ActiveConnection = *pConnection

' // Get a reference to the Procedures collection
DIM pProcedures AS CADOXProcedures = pCatalog.Procedures
' // Append the procedure
IF pProcedures.Append("AuthorById", *pCommand) = S_OK THEN
   ? "Procedure created"
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
