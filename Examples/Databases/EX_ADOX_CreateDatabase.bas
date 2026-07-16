' ########################################################################################
' Microsoft Windows
' Compiler: FreeBasic
' Copyright (c) 2026 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#cmdline "-s console"
'#define _CADOX_DEBUG_ 1
#include once "AfxNova/CADOX.inc"
USING AfxNova

' // Create a Catalog object
DIM pCatalog AS CADOXCatalog

' // Create the database
'IF pCatalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=new_db.mdb") = S_OK THEN
'   ? "Database created"
'ELSE
'   ? pCatalog.GetErrorInfo
'END IF

print pCatalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=new_db.mdb")

PRINT
PRINT "Press any key..."
SLEEP
