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

' ========================================================================================
' Checks if a table exists in the database
' ========================================================================================
FUNCTION ADOX_TableExists (BYVAL pCon AS Afx_ADOConnection PTR, BYREF wszTableName AS WSTRING) AS BOOLEAN

   IF pCon = NULL THEN RETURN FALSE

   ' // Create a Catalog object
   DIM pCatalog AS CADOXCatalog = AfxNewCom("ADOX.Catalog")
   IF *pCatalog = NULL THEN RETURN FALSE

   ' // Set the ActiveConnection property of the Catalog
   pCatalog.ActiveConnection = pCon
   ' // Get a reference to the Tables collection
   DIM pTables AS CADOXTables = pCatalog.Tables
   ' // Check if there are tables
   IF pTables.Count = 0 THEN RETURN FALSE
   ' // Check if the table exists
   DIM pTable AS CADOXTable = pTables.Item(wszTableName)
   RETURN (*pTable <> NULL)

END FUNCTION
' ========================================================================================

DO   ' // Using a fake loop to avoid the use of GOTO or nested IFs/END IFs

' // Open the connection
DIM pConnection AS CADOConnection
IF *pConnection = NULL THEN EXIT DO
pConnection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=biblio.mdb")

' // Set the ActiveConnection property of the Catalog
DIM pCatalog AS CADOXCatalog
pCatalog.ActiveConnection = *pConnection

IF ADOX_TableExists(*pConnection, "Publishers") THEN
   ? "Table exists"
ELSE
   ? "Table doesn't exist"
END IF

' // Close the connection
' // If you don't close it, they will be closed when the application ends
pConnection.Close

EXIT DO' // Unconditional exit of the fake loop
LOOP

PRINT
PRINT "Press any key..."
SLEEP
