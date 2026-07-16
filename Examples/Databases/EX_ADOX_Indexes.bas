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

' // Get a reference to the Publishers table
DIM pTable AS CADOXTable = pTables.Item("Authors")

' // Get a reference to the Indexes collection of the table
DIM pIndexes AS CADOXIndexes = pTable.Indexes

' // Get the number of objects in the collection
DIM IdxCount AS LONG = pIndexes.Count
IF IdxCount = 0 THEN EXIT DO

' // Enumerate the indexes
FOR i AS LONG = 0 TO IdxCount - 1
   SCOPE
      DIM pIndex AS CADOXIndex = pIndexes.Item(i)
      IF pIndex.PrimaryKey THEN
         ? "Index name: " & pIndex.Name
         DIM pColumns AS CADOXColumns = pIndex.Columns
         DIM colCount AS LONG = pColumns.Count
         FOR x AS LONG = 0 TO colCount - 1
            DIM pColumn AS CADOXColumn = pColumns.Item(x)
            ? "Name: "; pColumn.Name
            ? "SortOrder: "; pColumn.SortOrder
            ? "ParentCatalog: "; pColumn.ParentCatalog
            ? "-------------------------------------------"
         NEXT
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
