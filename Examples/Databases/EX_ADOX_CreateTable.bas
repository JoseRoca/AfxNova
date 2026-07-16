' ########################################################################################
' Microsoft Windows
' Compiler: FreeBasic
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

' // Create a new table called "Contacts3"
DIM pTable AS CADOXTable
IF *pTable = NULL THEN EXIT DO
pTable.Name = "Contacts3"

' // Create columns and append them to the Columns collection of the new Table object
' // Note that in ADOX the ADO Fields are called Columns
DIM pColumns AS CADOXColumns = pTable.Columns
pColumns.Append "FirstName", adVarWChar
pColumns.Append "LastName", adVarWChar
pColumns.Append "Phone", adVarWChar
pColumns.Append "Notes", adLongVarWChar

' // Add the new Table to the Tables collection of the database
IF pTables.Append(*pTable) = S_OK THEN
   ? "Table created"
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
