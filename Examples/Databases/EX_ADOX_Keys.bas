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

' // Get a reference to the Tables collection
DIM pTables AS CADOXTables = pCatalog.Tables
IF *pTables = NULL THEN EXIT DO

' // Get a reference to the table "Titles"
DIM pTable AS CADOXTable = pTables.Item("Titles")
IF *pTable = NULL THEN EXIT DO

' // Get a reference to the Keys collection of the table
DIM pKeys AS CADOXKeys = pTable.Keys
IF *pKeys = NULL THEN EXIT DO

' // Get the number of objects in the collection
DIM nCount AS LONG = pKeys.Count

' // Enumerate the tables
FOR i AS LONG = 0 TO nCount - 1
   SCOPE
      DIM pKey AS CADOXKey = pKeys.Item(i)
      IF *pKey <> NULL THEN
         ? "Name: "; pKey.Name
         ? "Type: "; pKey.Type_
         ? "Delete rule: "; pKey.DeleteRule
         ? "Update rule: "; pKey.UpdateRule
         ? "Related table: "; pKey.RelatedTable
         ? "----------------------------------------------------"
      END IF
   END SCOPE
NEXT

' // Close the connection
' // If you don't close it, they will be closed when the application ends
pConnection.Close

EXIT DO' // Unconditional exit of the fake loop
LOOP

PRINT
PRINT "Press any key..."
SLEEP
