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

' // Set the ActiveConnection property of the Catalog
pCatalog.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=nwind.mdb"

' // Create a Table object
DIM pTable AS CADOXTable
IF *pTable = NULL THEN EXIT DO
' // Set the name of the new table
pTable.Name = "NewTable"
' // Append a numeric and a text field to the new table
DIM pColumns AS CADOXColumns = pTable.Columns
pColumns.Append "NumField", adInteger
' // Note: If you are using Jet 3.51 instead of 4.0 use %adVarChar
pColumns.Append "TextField", adVarWChar, 20
' // Append the new table
DIM pTables AS CADOXTables = pCatalog.Tables
pTables.Append *pTable

print pTables.GetErrorInfo

' // Define the primary key
DIM pPrimaryKey AS CADOXKey
pPrimaryKey.Name = "adKeyPrimary"
pPrimaryKey.Type_ = adKeyPrimary
pPrimaryKey.RelatedTable = "Title Author"
pColumns.Append "NumField", adVarWChar
DIM pColumn AS CADOXColumn = pColumns.Item("NumField")
pColumn.RelatedColumn = "Au_ID"
pPrimaryKey.UpdateRule = adRICascade

' // Append the primary key
DIM pKeys AS CADOXKeys = pTable.Keys
pKeys.AppendKey *pPrimaryKey

print pKeys.GetErrorInfo

EXIT DO' // Unconditional exit of the fake loop
LOOP

PRINT
PRINT "Press any key..."
SLEEP
