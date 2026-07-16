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

' // Define the foreign key
DIM pForeignKey AS CADOXKey
pForeignKey.Name = "CustOrder"
pForeignKey.Type_ = adKeyForeign
pForeignKey.RelatedTable = "Customers"
DIM pColumns AS CADOXColumns = pForeignKey.Columns
pColumns.Append "CustomerId", adVarWChar
DIM pColumn AS CADOXColumn = pColumns.Item("CustomerId")
pColumn.RelatedColumn = "CustomerId"
pForeignKey.UpdateRule = adRICascade

' // Append the foreign key
DIM pTables AS CADOXTables = pCatalog.Tables
DIM pTable AS CADOXTable = pTables.Item("Orders")
DIM pKeys AS CADOXKeys = pTable.Keys
pKeys.AppendKey *pForeignKey

print pKeys.GetErrorInfo

EXIT DO' // Unconditional exit of the fake loop
LOOP

PRINT
PRINT "Press any key..."
SLEEP
