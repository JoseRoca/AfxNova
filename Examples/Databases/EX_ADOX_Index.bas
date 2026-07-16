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

' // Create a Catalog object
DIM pCatalog AS CADOXCatalog
IF *pCatalog = NULL THEN EXIT DO

' // Open the connection
DIM pConnection AS CADOConnection
IF *pConnection = NULL THEN EXIT DO
pConnection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb")

' // Set the ActiveConnection property of the Catalog
pCatalog.ActiveConnection = *pConnection

' // Create a Table object
DIM pTable AS CADOXTable
IF *pTable = NULL THEN EXIT DO

' // Set the name of the table
pTable.Name = "myTable"

' // Get a reference to the Columns collection of the table
DIM pColumns AS CADOXColumns = pTable.Columns
IF *pColumns = NULL THEN EXIT DO

' // Append columns to the new table
pColumns.Append "Column1", adInteger
pColumns.Append "Column2", adInteger
' // Note: If you are using Jet 3.51 instead of 4.0 use adVarChar
pColumns.Append "Column3", adVarWChar, 50

' // Append the new table
DIM pTables AS CADOXTables = pCatalog.Tables
IF *pTables = NULL THEN EXIT DO
pTables.Append *pTable

' // Create an Index object
DIM pIndex AS CADOXIndex
IF *pIndex = NULL THEN EXIT DO

' // Define a multicolumn index
pIndex.Name = "multicolidx"
DIM pIdxColumns AS CADOXColumns = pIndex.Columns
pIdxColumns.Append "Column1", adVarWChar
pIdxColumns.Append "Column2", adVarWChar

' // Append the index to the table
DIM pIndexes AS CADOXIndexes = pTable.Indexes
pIndexes.Append *pIndex

' // Delete the table as this is a demonstration
IF pTables.Delete_("MyTable") = S_OK THEN
   ? "MyTable deleted"
ELSE
   ? pTables.GetErrorInfo
END IF

' // Close the connection
' // If you don't close it, they will be closed when the application ends
pConnection.Close

EXIT DO' // Unconditional exit of the fake loop
LOOP

PRINT
PRINT "Press any key..."
SLEEP
