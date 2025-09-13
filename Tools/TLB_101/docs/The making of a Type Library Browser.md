# The making of a Type Library browser

Type libraries contain the specification for one or more COM elements, including classes, interfaces, enumerations, and more. These files are stored in a standard binary format. A type library can be a stand-alone file with the .tlb file name extension, or it can be stored as a resource in an executable file, which can have an .ocx, .dll, or .exe file name extension. The type library viewers and conversion tools described following read this format to gain information about the COM elements in the library.
Before you can program an object in a particular programming language, you must be able to view its type library in that language. Doing this provides you with the proper syntax for the classes, interfaces, methods, properties, and events of the COM object.

https://msdn.microsoft.com/en-us/library/windows/desktop/ms680581(v=vs.85).aspx

Microsoft provides an OLE-COM Viewer, Oleview.exe, an application supplied with Visual C++ that displays the COM objects installed on your computer and the interfaces they support. You can use this object viewer to view type libraries. Visual Basic 6 has its own object browser.

What we need is a type library browser able to generate code in the proper syntax to be used with FreeBasic.

# Retrieving information from the registry

The first step is to retrieve the type libraries registered in the system.

All the registered type libraries have an entry in the registry under HKEY_CLASSES_ROOT\TypeLib. Under this section, every subkey is the CLSID of a TypeLibrary. Under the CLSID subkey there are one or more subkeys with the version numbers, that generally take the format MajorVersion.MinorVersion (e.g.: 1.0). Opening these version subkeys, there are other subkeys. The one that we need is the default one (0), which can contain one or two subkeys, "win32" and/or "win64". Opening these subkeys we can retrieve the path of the type library.

```
' ========================================================================================
' Searches for the win32 subkey
' ========================================================================================
FUNCTION TLB_RegSearchWin32 (BYVAL pwszKey AS WSTRING PTR) AS DWSTRING

   IF pwszKey = NULL THEN RETURN ""

   ' // Recursively searches for the win directory
   DIM hr AS LONG, hKey AS HKEY, dwIdx AS DWORD, ft AS FILETIME
   DIM wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH
   DIM wszKey AS WSTRING * MAX_PATH = *pwszKey
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      hr = RegOpenKeyExW (HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
      IF hr <> ERROR_SUCCESS THEN RETURN ""
      IF hKey = NULL THEN RETURN ""
      hr = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> S_OK THEN EXIT DO
      IF UCASE(wszKeyName) = "WIN32" THEN EXIT DO
      dwIdx += 1
   LOOP WHILE hr = S_OK

   ' // Closes the registry and returns the result
   RegCloseKey hKey
   IF hr <> S_OK OR LEN(wszKeyName) = 0 THEN RETURN ""
   RETURN wszKey

END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Searches for the win64 subkey
' ========================================================================================
FUNCTION TLB_RegSearchWin64 (BYVAL pwszKey AS WSTRING PTR) AS DWSTRING

   IF pwszKey = NULL THEN RETURN ""

   ' // Recursively searches for the win directory
   DIM hr AS LONG, hKey AS HKEY, dwIdx AS DWORD, ft AS FILETIME
   DIM wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH
   DIM wszKey AS WSTRING * MAX_PATH = *pwszKey
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      hr = RegOpenKeyExW (HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
      IF hr <> ERROR_SUCCESS THEN RETURN ""
      IF hKey = NULL THEN RETURN ""
      hr = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> S_OK THEN EXIT DO
      IF UCASE(wszKeyName) = "WIN64" THEN EXIT DO
      dwIdx += 1
   LOOP WHILE hr = S_OK

   ' // Closes the registry and returns the result
   RegCloseKey hKey
   IF hr <> S_OK OR LEN(wszKeyName) = 0 THEN RETURN ""
   RETURN wszKey

END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Returns the path of the typelib.
' ========================================================================================
FUNCTION TLB_RegEnumDirectory (BYVAL pwszKey AS WSTRING PTR) AS DWSTRING

   IF pwszKey = NULL THEN RETURN ""

   ' // Searches the HKEY_CLASSES_ROOT\TypeLib\<LIBID> node.
   DIM hKey AS HKEY, wszKey AS WSTRING * MAX_PATH = *pwszKey
   DIM hr AS LONG = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN RETURN ""
   IF hKey = 0 THEN RETURN ""
   DIM dwIdx AS DWORD, wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH, ft AS FILETIME
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH, wszSubkey AS WSTRING * MAX_PATH
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      hr = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> S_OK THEN EXIT DO
#ifdef __FB_64BIT__
      wszSubkey = TLB_RegSearchWin64(wszKey & "\" & wszKeyName)
      IF LEN(wszSubkey) THEN wszKey = wszSubkey & "\" & "win64"
#else
      wszSubkey = TLB_RegSearchWin32(wszKey & "\" & wszKeyName)
      IF LEN(wszSubkey) THEN wszKey = wszSubkey & "\" & "win32"
#endif
#ifdef __FB_64BIT__
      ' // Not all the typelibs have separate entries in the win64 subkey
      ' // See: https://msdn.microsoft.com/en-us/library/windows/desktop/ms724072(v=vs.85).aspx
      ' // If not found in the win64 subkey search in the win32 subkey
      IF LEN(wszSubkey) = 0 THEN wszSubkey = TLB_RegSearchWin32(wszKey & "\" & wszKeyName)
      IF LEN(wszSubkey) THEN wszKey = wszSubkey & "\" & "win32"
#endif
      IF LEN(wszSubkey) THEN EXIT DO
      dwIdx += 1
   LOOP
   RegCloseKey hKey
   IF hr <> S_OK OR LEN(wszSubkey) = 0 THEN RETURN ""

   hKey = NULL
   hr = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN RETURN ""
   DIM keyType AS DWORD
   DIM wszValueName AS WSTRING * MAX_PATH
   DIM wszKeyValue  AS WSTRING * MAX_PATH
   DIM cValueName AS DWORD = MAX_PATH
   DIM cbData AS DWORD = MAX_PATH
   dwIdx = 0
   hr = RegEnumValueW(hKey, dwIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData)

   ' // Closes the registry and returns the value
   RegCloseKey hKey
   RETURN wszKeyValue

END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Returns the different versions of the typelib.
' Parameter
' - hListView = Handle of the list view.
' - pwszLibId = Library guid.
' Return value: TRUE or FALSE.
' ========================================================================================
FUNCTION TLB_RegEnumVersions (BYVAL hListView AS HWND, BYVAL pwszLibId AS WSTRING PTR) AS BOOLEAN

   IF hListView = NULL OR pwszLibId = NULL THEN EXIT FUNCTION

   ' // Searches the HKEY_CLASSES_ROOT\TypeLib\<LIBID> node.
   DIM hKey AS HKEY
   DIM wszKey AS WSTRING * MAX_PATH = "TypeLib\" & *pwszLibId
   DIM hr AS LONG = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN RETURN FALSE
   IF hKey = NULL THEN RETURN FALSE

   ' // Opens the subtrees of the different versions of the TyleLib library
   DIM dwIdx AS DWORD, wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH, ft AS FILETIME
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      DIM hr AS LONG = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> ERROR_SUCCESS THEN RETURN FALSE
      ' // Gets the default value
      DIM hVerKey AS HKEY, wszSubKey AS WSTRING * MAX_PATH = wszKey & "\" & wszKeyName
      hr = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszSubKey, 0, KEY_READ, @hVerKey)
      IF hr <> ERROR_SUCCESS THEN RETURN FALSE
      DIM wszVer AS WSTRING * MAX_PATH = wszKeyName
      ' // Enumerate the entries until the default one, if any, is found.
      ' // RegEnumValue returns values in any order. This includes unnamed values.
      ' // A key may have one unnamed value. An unnamed value is displayed as (Default)
      ' // in Regedit.exe. If an unnamed value doesn't exist under a given key, it is
      ' // displayed as (value not set) in Regedit.exe.
      ' // Only existing unnamed values can be enumerated. If an unnamed value is enumerated, the
      ' // RegEnumValue function sets wszValueName to an empty string ("") and it sets cValueName to 0.
      DIM verIdx AS DWORD, cValueName AS DWORD, cbData AS DWORD, keyType AS DWORD
      DIM wszValueName AS WSTRING * 16383, wszKeyValue AS WSTRING * MAX_PATH
      DO
         cValueName = 16383 : cbData = MAX_PATH * 2 : wszValueName = "" : wszKeyValue = ""
         hr = RegEnumValueW(hVerKey, verIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData)
         IF LEN(wszValueName) = 0 THEN EXIT DO   ' // This is the default, unnamed value
         IF hr <> ERROR_SUCCESS THEN EXIT DO
         ' // Increment the index of the value to be retrieved.
         verIdx += 1
      LOOP
      ' // Closes the handle of the registry key
      RegCloseKey hVerKey
      DIM wszDesc AS WSTRING * MAX_PATH
      ' // If wszValueName is an empty string, assume that the value name is the key value.
      IF LEN(wszValueName) = 0 THEN wszDesc = wszKeyValue ELSE wszDesc = wszValueName
      ' // Increment the index of the subkey
      dwIdx += 1
      ' // Find the path of the type library
      DIM wszPath AS WSTRING * MAX_PATH = TLB_RegEnumDirectory(wszKey & "\" & wszKeyName)
      ' // If there is not path, skip the typelib because we can't retrieve it
      IF LEN(wszPath) = 0 THEN CONTINUE DO
      ' // Remove double quotes (if any)
      wszPath = DWStrRemove(wszPath, """")
      ' // Convert short paths to long paths
      ' // Dont use it with all files or these ending with version numbers
      ' // (a \ and a number) will we skipped.
      IF INSTR(wszPath, "%") THEN
         DIM wszDest AS WSTRING * MAX_PATH, cbLen AS DWORD
         cbLen = ExpandEnvironmentStringsW(@wszPath, @wszDest, MAX_PATH)
         IF cbLen THEN wszPath = wszDest
      END IF
      IF INSTR(wszPath, "~") <> 0 THEN
         DIM cbPath AS DWORD = LEN(wszPath)
         cbPath = GetLongPathNameW(wszPath, wszPath, cbPath)
      END IF
      DIM pathPos AS LONG = INSTRREV(wszPath, "\")
      DIM wszFile AS WSTRING * MAX_PATH = MID(wszPath, pathPos + 1)
      ' // Some have an added backslah and a number
      IF INSTR(wszFile, ".") = 0 THEN
         DIM wszTmpPath AS WSTRING * MAX_PATH = LEFT(wszPath, pathpos - 1)
         pathPos = INSTRREV(wszTmpPath, "\", LEN(wszFile) - 3)
         wszFile = MID(wszPath, pathPos + 1)
      END IF
      IF LEN(wszFile) = 0 THEN CONTINUE DO  ' // Skip files without a full path
      ' // Skip .OCA files
      DIM wszTemp AS WSTRING * MAX_PATH = wszPath
      IF MID(wszTemp, LEN(wszTemp) - 2, 1) = "\" THEN wszTemp = LEFT(wszTemp, LEN(wszTemp) - 2)
      IF MID(wszTemp, LEN(wszTemp) - 3, 1) = "\" THEN wszTemp = LEFT(wszTemp, LEN(wszTemp) - 3)
      ' // .OCA files are created by Visual Basic (we don't need them)
      IF UCASE(RIGHT(wszTemp, 4)) = ".OCA" THEN wszPath = ""
      IF LEN(wszPath) THEN
         ' // If the description is empty, use the file name
         IF wszDesc = "" THEN wszDesc = "[" & wszFile & "]"
         ' // Add the version number
         wszDesc += " (Ver " & wszVer & ")"
         ' // Add the items to the list view
         DIM lItemIdx AS LONG = ListView_AddItem(hListView, 0, 0, @wszFile)
         ListView_SetItemText(hListView, lItemIdx, 1, @wszDesc)
         ListView_SetItemText(hListView, lItemIdx, 2, @wszPath)
         ListView_SetItemText(hListView, lItemIdx, 3, pwszLibId)
      END IF
   LOOP

   ' // Closes the registry key
   RegCloseKey hKey

END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Enumerates all the typelibs.
' Parameter
' - hListView = Handle of the list view.
' Return value: TRUE or FALSE.
' ========================================================================================
FUNCTION TLB_RegEnumTypeLibs (BYVAL hListView AS HWND) AS BOOLEAN

   IF hListView = NULL THEN RETURN FALSE

   ' // Opens the HKEY_CLASSES_ROOT\TypeLib subtree
   DIM hKey AS HKEY
   DIM wszKey AS WSTRING * MAX_PATH = "TypeLib"
   DIM hr AS LONG = RegOpenKeyExW(HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN RETURN FALSE
   IF hKey = NULL THEN RETURN FALSE

   ' // Parses all the TypeLib subtree and gets the CLSIDs of all the TypeLibs
   DIM dwIdx AS DWORD, wszKeyName AS WSTRING * MAX_PATH, wszClass AS WSTRING * MAX_PATH, ft AS FILETIME
   DIM cchName AS DWORD = MAX_PATH, cchClass AS DWORD = MAX_PATH
   DO
      wszKeyName = "" : cchName = MAX_PATH : wszClass = "" : cchClass = MAX_PATH
      hr = RegEnumKeyExW(hKey, dwIdx, @wszKeyName, @cchName, 0, @wszClass, @cchClass, @ft)
      IF hr <> ERROR_SUCCESS THEN EXIT DO
      TLB_RegEnumVersions hListView, @wszKeyName
      dwIdx += 1
   LOOP
   ' // Closes the registry
   RegCloseKey hKey

   RETURN TRUE

END FUNCTION
' ========================================================================================
```

# Loading the type library

Once we have retrieved the path of the type library, the next step if to load it calling the API functions **LoadTypelib** or **LoadTypeLibEx**, that return a pointer of the **ITypeLib** interface.

This is my definition of that interface:

```
' ########################################################################################
' Interface name = ITypeLib
' Extracts information about a type library, the data that describes a set of objects.
' Inherited interface = IUnknown
' ########################################################################################

TYPE Afx_ITypeLib_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetTypeInfoCount () AS UINT
   DECLARE ABSTRACT FUNCTION GetTypeInfo (BYVAL index AS UINT, BYVAL ppTInfo AS Afx_ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfoType (BYVAL index AS UINT, BYVAL pTKind AS TYPEKIND PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeInfoOfGuid (BYVAL guid AS CONST GUID CONST PTR, BYVAL ppTinfo AS Afx_ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetLibAttr (BYVAL ppTLibAttr AS TLIBATTR PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeComp (BYVAL ppTComp AS ITypeComp PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDocumentation (BYVAL index AS INT_, BYVAL pBstrName AS AFX_BSTR PTR, BYVAL pBstrDocString AS AFX_BSTR PTR, BYVAL pdwHelpContext AS DWORD PTR, BYVAL pBstrHelpFile AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION IsName (BYVAL szNameBuf AS LPOLESTR, BYVAL lHashVal AS ULONG, BYVAL pfName AS WINBOOL PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION FindName (BYVAL szNameBuf AS LPOLESTR, BYVAL lHashVal AS ULONG, BYVAL ppTInfo AS Afx_ITypeInfo PTR PTR, BYVAL rgMemId AS MEMBERID PTR, BYVAL pcFound AS USHORT PTR) AS HRESULT
   DECLARE ABSTRACT SUB      ReleaseTLibAttr (BYVAL pTLibAttr AS TLIBATTR PTR)
END TYPE
```

and this is some code to load the type library and extract basic information:

```
' =====================================================================================
' Load the type library
' =====================================================================================
FUNCTION CParseTypeLib.LoadTypeLibEx (BYVAL pwszPath AS WSTRING PTR) AS HRESULT

   DIM pTypeLib AS ITypeLib PTR
   DIM hr AS HRESULT = .LoadTypeLibEx(pwszPath, REGKIND_NONE, @pTypeLib)
   m_pTypeLib = cast(Afx_ITypeLib PTR, cast(ULONG_PTR, pTypeLib))
   IF hr <> S_OK OR m_pTypeLib = NULL THEN
      TLB_MsgBox m_pWindow->hWindow, "Error &h" & HEX(hr, 8) & " loading " & *pwszPath, _
         MB_OK OR MB_ICONERROR OR MB_APPLMODAL, "CParseTypeLib.LoadTypeLibEx"
      RETURN hr
   END IF
   m_LibPath = *pwszPath

   ' // Gets the documentation
   DIM AS AFX_BSTR bstrLibName, bstrLibHelpString, bstrLibHelpFile
   hr = m_pTypeLib->GetDocumentation(-1, @bstrLibName, @bstrLibHelpString, @m_LibHelpContext, @bstrLibHelpFile)
   m_LibName = *bstrLibName : m_LibHelpString = *bstrLibHelpString : m_LibHelpFile = *bstrLibHelpFile
   SysFreeString bstrLibName : SysFreeString bstrLibHelpString : SysFreeString bstrLibHelpFile
   IF hr <> S_OK THEN
      TLB_MsgBox m_pWindow->hWindow, "Error &h" & HEX(hr, 8) & " retrieving the documentation", _
         MB_OK OR MB_ICONERROR OR MB_APPLMODAL, "CParseTypeLib.LoadTypeLibEx"
      RETURN hr
   END IF
   ' // Use the library name as a namespace
   m_Namespace = TRIM(m_LibName, ANY CHR(32, 34))
   DIM hEditNamespace AS HWND = cast(HWND, m_pWindow->UserData(AFX_HEDITNAMESPACE))
   SetWindowText hEditNamespace, m_Namespace

   ' // Gets the attributes of the library
   DIM pLibAttr AS TLIBATTR PTR
   hr = m_pTypeLib->GetLibAttr(@pLibAttr)
   IF hr <> S_OK OR pLibAttr = NULL THEN
      TLB_MsgBox m_pWindow->hWindow, "Error &h" & HEX(hr, 8) & " retrieving the attributes", _
         MB_OK OR MB_ICONERROR OR MB_APPLMODAL, "CParseTypeLib.LoadTypeLibEx"
      RETURN hr
   END IF
   m_LibGuid = AfxGuidText(pLibAttr->guid)
   m_LibLcid = pLibAttr->lcid
   m_LibSysKind = pLibAttr->syskind
   m_LibMajorVersion = pLibAttr->wMajorVerNum
   m_LibMinorVersion = pLibAttr->wMinorVerNum
   m_LibAttr = pLibAttr->wLibFlags
   m_pTypeLib->ReleaseTLibAttr(pLibAttr)

   ' // Treeview handle
   DIM hTreeView AS HWND = cast(HWND, m_pWindow->UserData(AFX_HTREEVIEW))
   ' // Delete all the items in the tree view
   TreeView_DeleteAllItems(hTreeView)
   ' // Create the nodes
   m_hRootNode = TreeView_AddItem(hTreeView, 0, TVI_ROOT, m_LibName)
   m_hDocNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Documentation")
   m_hProgIDsNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "ProgIDs (Program identifiers)")
   m_hVerIndProgIDsNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Version independent ProgIDs")
   m_hClsIDsNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "ClsIDs (Class identifiers)")
   m_hIIDsNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "IIDs (Interface identifiers)")
   m_hCoClassesNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "CoClasses")
   m_hTypeDefsNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Typedefs")
   m_hAliasesNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Aliases")
   m_hEnumsNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Enumerations")
   m_hRecordsNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Structures")
   m_hUnionsNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Unions")
   m_hModulesNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Modules")
   m_hIntNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Interfaces")
   m_hOleAutIntNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Ole automation interfaces")
   m_hDualIntNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Dual interfaces")
   m_hDispIntNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Dispatch interfaces")
   m_hDispblIntNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Dispatchable interfaces")
   m_hEventsNode = TreeView_AddItem(hTreeView, m_hRootNode, NULL, "Events interfaces")
   ' // Fill the documentation node
   IF m_hDocNode THEN
      IF LEN(m_LibHelpString) THEN TreeView_AddItem hTreeView, m_hDocNode, NULL, "Help string = " & m_LibHelpString
      TreeView_AddItem hTreeView, m_hDocNode, NULL, "GUID = " & m_LibGuid
      TreeView_AddItem hTreeView, m_hDocNode, NULL, "LCID = " & WSTR(m_LibLcid)
      TreeView_AddItem hTreeView, m_hDocNode, NULL, "Major version = " & WSTR(m_LibMajorVersion)
      TreeView_AddItem hTreeView, m_hDocNode, NULL, "Minor version = " & WSTR(m_LibMinorVersion)
      TreeView_AddItem hTreeView, m_hDocNode, NULL, "Path = " & m_LibPath
      IF m_LibHelpContext THEN TreeView_AddItem hTreeView, m_hDocNode, NULL, "Help context = " & WSTR(m_LibHelpContext)
      IF LEN(m_LibHelpFile) THEN TreeView_AddItem hTreeView, m_hDocNode, NULL, "Help file = " & m_LibHelpFile
      TreeView_AddItem hTreeView, m_hDocNode, NULL, "Attributes = " & WSTR(m_LibAttr) & " [&h" & HEX(m_LibAttr, 8) & "] " & TLB_LibFlagsToStr(m_LibAttr)
      SELECT CASE m_LibSysKind
         CASE SYS_WIN16 : TreeView_AddItem hTreeView, m_hDocNode, NULL, "Target OS = " & WSTR(m_LibSysKind) & " (Win16)"
         CASE SYS_WIN32 : TreeView_AddItem hTreeView, m_hDocNode, NULL, "Target OS = " & WSTR(m_LibSysKind) & " (Win32)"
         CASE SYS_MAC   : TreeView_AddItem hTreeView, m_hDocNode, NULL, "Target OS = " & WSTR(m_LibSysKind) & " (MAC)"
         CASE SYS_WIN64 : TreeView_AddItem hTreeView, m_hDocNode, NULL, "Target OS = " & WSTR(m_LibSysKind) & " (Win64)"
      END SELECT
   END IF

   ' // Parse the type infos
   this.ParseTypeInfos

   ' // Generate code
   this.GenerateCode

   ' // Expands the root node
   TreeView_Expand(hTreeView, m_hRootNode, TVE_EXPAND)

END FUNCTION
' =====================================================================================
```

# Parsing the information

To parse the type library information we need to call the methods of the **ITypeInfo** interface.

This is my definition of that interface:

```
' ########################################################################################
' Interface name = ITypeInfo
' Extracts information about the characteristics and capabilities of objects from type libraries.
' Inherited interface = IUnknown
' ########################################################################################

TYPE Afx_ITypeInfo AS Afx_ITypeInfo_

TYPE Afx_ITypeInfo_ EXTENDS Afx_IUnknown
   DECLARE ABSTRACT FUNCTION GetTypeAttr (BYVAL ppTypeAttr AS TYPEATTR PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetTypeComp (BYVAL ppTComp AS ITypeComp PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetFuncDesc (BYVAL index AS UINT, BYVAL ppFuncDesc AS FUNCDESC PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetVarDesc (BYVAL index AS UINT, BYVAL ppVarDesc AS VARDESC PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetNames (BYVAL memid AS MEMBERID, BYVAL rgBstrNames AS AFX_BSTR PTR, BYVAL cMaxNames AS UINT, BYVAL pcNames AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetRefTypeOfImplType (BYVAL index AS UINT, BYVAL pRefType AS HREFTYPE PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetImplTypeFlags (BYVAL index AS UINT, BYVAL pImplTypeFlags AS INT_ PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetIDsOfNames (BYVAL rgszNames AS LPOLESTR PTR, BYVAL cNames AS UINT, BYVAL pMemId AS MEMBERID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION Invoke (BYVAL pvInstance AS PVOID, BYVAL memid AS MEMBERID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDocumentation (BYVAL memid AS MEMBERID, BYVAL pBstrName AS AFX_BSTR PTR, BYVAL pBstrDocString AS AFX_BSTR PTR, BYVAL pdwHelpContext AS DWORD PTR, BYVAL pBstrHelpFile AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetDllEntry (BYVAL memid AS MEMBERID, BYVAL invKind AS INVOKEKIND, BYVAL pBstrDllName AS AFX_BSTR PTR, BYVAL pBstrName AS AFX_BSTR PTR, BYVAL pwOrdinal AS WORD PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetRefTypeInfo (BYVAL hRefType AS HREFTYPE, BYVAL ppTInfo AS Afx_ITypeInfo PTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION AddressOfMember (BYVAL memid AS MEMBERID, BYVAL invKind AS INVOKEKIND, BYVAL ppv AS PVOID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION CreateInstance (BYVAL pUnkOuter AS IUnknown PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS PVOID PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetMops (BYVAL memid AS MEMBERID, BYVAL pBstrMops AS AFX_BSTR PTR) AS HRESULT
   DECLARE ABSTRACT FUNCTION GetContainingTypeLib (BYVAL ppTLib AS Afx_ITypeLib PTR PTR, BYVAL pIndex AS UINT PTR) AS HRESULT
   DECLARE ABSTRACT SUB      ReleaseTypeAttr (BYVAL pTypeAttr AS TYPEATTR PTR)
   DECLARE ABSTRACT SUB      ReleaseFuncDesc (BYVAL pFuncDesc AS FUNCDESC PTR)
   DECLARE ABSTRACT SUB      ReleaseVarDesc (BYVAL pVarDesc AS VARDESC PTR)
END TYPE
```

To extract the type library information first wee need to rerieve how many **TypeInfos** it contains:

```
' // Retrieves the number of TypeInfos
DIM TypeInfoCount AS LONG = m_pTypeLib->GetTypeInfoCount
IF TypeInfoCount = 0 THEN
   TLB_MsgBox m_pWindow->hWindow, "This TypeLib doesn't have type infos", _
      MB_OK OR MB_ICONERROR OR MB_APPLMODAL, "CParseTypeLib.ParseTypeInfos"
   RETURN S_FALSE
END IF
```

For each type info we will retrieve the type of a type description, a pointer to its **ITypeInfo** interface and a pointer to the **TYPEATTR** structure that contains the attributes of the type description.

```
FOR i AS LONG = 0 TO TypeInfoCount - 1
   ' // Retrieves the TypeKind
   DIM hr AS HRESULT = m_pTypeLib->GetTypeInfoType(i, @pTKind)
   IF hr <> S_OK THEN
      TLB_MsgBox m_pWindow->hWindow, "Error &h" & HEX(hr, 8) & " retrieving the info type", _
         MB_OK OR MB_ICONERROR OR MB_APPLMODAL, "CParseTypeLib.ParseTypeInfos"
      EXIT FOR
   END IF

   ' // Retrieves the ITypeInfo interface
   hr = m_pTypeLib->GetTypeInfo(i, @pTypeInfo)
   IF hr <> S_OK OR pTypeInfo = NULL THEN
      TLB_MsgBox m_pWindow->hWindow, "Error &h" & HEX(hr, 8) & " retrieving the ITypeInfo interface", _
         MB_OK OR MB_ICONERROR OR MB_APPLMODAL, "CParseTypeLib.ParseTypeInfos"
      EXIT FOR
   END IF

   ' // Gets the address of a pointer to the TYPEATTR structure
   hr = pTypeInfo->GetTypeAttr(@pTypeAttr)
   IF hr <> S_OK OR pTypeAttr = NULL THEN
      TLB_MsgBox m_pWindow->hWindow, "Error &h" & HEX(hr, 8) & " retrieving the address of the TypeAttr structure", _
         MB_OK OR MB_ICONERROR OR MB_APPLMODAL, "CParseTypeLib.ParseTypeInfos"
      EXIT FOR
   END IF
      ...
      ...
      ...
NEXT
```

Inside the loop, we will process each block of information separately according its type:

```
      SELECT CASE pTKind
         CASE TKIND_COCLASS   ' // CoClasses
            ...
            ...
            ...
         CASE TKIND_RECORD, TKIND_UNION   ' // Structures and unions
            ...
            ...
            ...
         CASE TKIND_ALIAS   ' // Aliases and typedefs
            ...
            ...
            ...
         CASE TKIND_ENUM, TKIND_MODULE   ' // Enumerations and modules
            ...
            ...
            ...
         CASE TKIND_INTERFACE, TKIND_DISPATCH   ' // Interfaces
            ...
            ...
            ...
      END SELECT
```

# Enumerating the constants

If the retrieved type info is of type TKIND_ENUM or TKIND_MODULE, the *cvars* member of the **TYPEATTR** structure ( https://msdn.microsoft.com/en-us/library/windows/desktop/ms221003(v=vs.85).aspx ) contains the number of variables and the **GetVarDesc** method of the **ITypeInfo** interface retrieves a **VARDESC** structure ( https://msdn.microsoft.com/en-us/library/windows/desktop/ms221391(v=vs.85).aspx ) that describes the specified variable.

```
' =====================================================================================
' Retrieves information about constants
' =====================================================================================
FUNCTION CParseTypeLib.GetConstants (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL pTypeAttr AS TYPEATTR PTR, BYVAL hTreeView AS HWND, BYVAL hSubNode AS HTREEITEM) AS HRESULT

   IF pTypeInfo = NULL OR pTypeAttr = NULL THEN RETURN E_INVALIDARG

   FOR x AS LONG = 0 TO pTypeAttr->cVars - 1
      DIM pVarDesc AS VARDESC PTR
      DIM hr AS HRESULT = pTypeInfo->GetVarDesc(x, @pVarDesc)
      IF hr <> S_OK OR pVarDesc = NULL THEN EXIT FOR
      DIM dwsName AS DWSTRING, bstrName AS AFX_BSTR
      pTypeInfo->GetDocumentation(pVarDesc->memid, @bstrName, NULL, NULL, NULL)
      dwsName = *bstrName
      SysFreeString bstrName
      ' // Some components use spaces in the names of enumeration members!
      IF INSTR(dwsName, " ") THEN
         dwsName = DWStrReplace(dwsName, " ", "_")
      END IF
      ' // Pointer to the variant that stores the value of the constant
      DIM vtPtr AS tagVARIANT PTR = pVarDesc->lpvarvalue
      ' // Gets the value of the constant
      DIM dwsValue AS DWSTRING = AfxVarToStr(vtPtr)
      DIM dwsTypeKind AS DWSTRING = TLB_VarTypeToConstant(pVarDesc->elemdescVar.tdesc.vt)
      SELECT CASE pVarDesc->elemdescVar.tdesc.vt
         CASE VT_I1, VT_UI1, VT_I2, VT_UI2, VT_I4, VT_UI4, VT_INT, VT_UINT
            dwsValue = dwsValue & "   ' (&h" & HEX(VAL(dwsValue), 8) & ")"
         CASE VT_BSTR, VT_LPSTR, VT_LPWSTR
            ' // cdosys.dll contains VT_BSTR constants
            dwsValue = CHR(34) & dwsValue & CHR(34)
         CASE VT_PTR
            DIM ptdesc AS TYPEDESC PTR = pVarDesc->elemdescVar.tdesc.lptdesc
            IF ptdesc THEN
               ' WORD PTR (null terminated unicode string)
               ' hxds.dll contains a module with these kind of constants.
               IF ptdesc->vt = VT_UI2 THEN dwsValue = CHR(34) & dwsValue & CHR(34)
            END IF
         ' // Other types can be VT_CARRAY and VT_USERDEFINED, but don't have a typelib to test
      END SELECT
      DIM hSubNode2 AS HTREEITEM = TreeView_AddItem(hTreeView, hSubNode, NULL, dwsName & " = " & dwsValue)
      TreeView_AddItem(hTreeView, hSubNode2, NULL, "TYPE = " & dwsTypeKind)
      TreeView_AddItem(hTreeView, hSubNode2, NULL, "VALUE = " & dwsValue)
      TreeView_AddItem(hTreeView, hSubNode2, NULL, "ID = " & WSTR(pVarDesc->memid))
      TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
      pTypeInfo->ReleaseVarDesc(pVarDesc) : pVarDesc = NULL
   NEXT

END FUNCTION
' =====================================================================================
```

Modules can also contain string constants and declarations of functions and procedures of external DLLs. This is covered by the following code in the **GetFunctions** method.

```
DIM wOrdinal AS WORD, bstrDllName AS AFX_BSTR, bstrEntryPoint AS AFX_BSTR
hr = pTypeInfo->GetDllEntry(pFuncDesc->memid, pFuncDesc->invkind, @bstrDllName, @bstrEntryPoint, @wOrdinal)
dwsDllName = *bstrDllName
SysFreeString bstrDllName
dwsEntryPoint = *bstrEntryPoint
SysFreeString bstrEntryPoint
IF hr = S_OK THEN
   IF LEN(dwsDllName) THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "DLL name = " & dwsDllName)
   IF LEN(dwsEntryPoint) THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Entry point = " & dwsEntryPoint)
   IF wOrdinal THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Ordinal = " & WSTR(wOrdinal))
END IF
```

# Enumerating structures and unions

If the retrieved type info is of type TKIND_RECORD or TKIND_UNION, the *cvars* member of the **TYPEATTR** structure contains the number of members or data members and the **GetVarDesc** method of the **ITypeInfo** interface retrieves a **VARDESC** structure that describes the specified member or data member.

The parsing of this type info is more convoluted that in the case of the constants because they don't contain simple values, but the names and types of the members of an structure that can be simple data types, but also pointers, arrays or even other structures.

```
' =====================================================================================
' Retrieves information about the members of records and unions, and of datamembers.
' Note: Bined.dll fails to retrieve information of several members of the VSPROPSHEETPAGE structure.
' =====================================================================================
FUNCTION CParseTypeLib.GetMembers (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL pTypeAttr AS TYPEATTR PTR, BYVAL hTreeView AS HWND, BYVAL hSubNode AS HTREEITEM, BYVAL bIsRecord AS BOOLEAN = FALSE) AS HRESULT

   DIM wIndirectionLevel AS WORD           ' // Indirection level
   DIM pRefTypeInfo AS Afx_ITypeInfo PTR   ' // Referenced TypeInfo interface
   DIM pVarTypeAttr AS TYPEATTR PTR        ' // Type attribute for the member
   DIM hSubNode2 AS HTREEITEM              ' // Sub node handle
   DIM hSubNode3 AS HTREEITEM              ' // Sub node handle
   DIM dwsVarName AS DWSTRING              ' // Variable name
   DIM dwsVarType AS DWSTRING              ' // Variable type
   DIM dwsTypeKind AS DWSTRING             ' // Type kind
   DIM dwsFBKeyword AS DWSTRING            ' // FB keyword
   DIM dwsFBSyntax AS DWSTRING             ' // FB syntax

   IF pTypeInfo = NULL OR pTypeAttr = NULL THEN RETURN E_INVALIDARG

   FOR x AS LONG = 0 TO pTypeAttr->cVars - 1

      ' // Gets a reference to the VarDesc structure
      DIM pVarDesc AS VARDESC PTR
      DIM hr AS HRESULT = pTypeInfo->GetVarDesc(x, @pVarDesc)
      IF hr <> S_OK OR pVarDesc = NULL THEN EXIT FOR

      ' // Retrieve information
      DIM bstrVarName AS AFX_BSTR
      pTypeInfo->GetDocumentation(pVarDesc->memid, @bstrVarName, NULL, NULL, NULL)
      dwsVarName = *bstrVarName
      SysFreeString bstrVarName
      hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, dwsVarName)
      TreeView_AddItem(hTreeView, hSubNode2, NULL, "DispID = " & WSTR(pVarDesc->memid) & " [&h" & HEX(pVarDesc->memid, 8) & "]")
      IF pVarDesc->wVarFlags THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Attributes = " & WSTR(pVarDesc->wVarFlags) & " [&h" & HEX(pVarDesc->wVarFlags, 8) & "]" & TLB_VarFlagsToStr(pVarDesc->wVarFlags))
      wIndirectionLevel = 0
      IF pVarDesc->elemdescVar.tdesc.vt = VT_USERDEFINED THEN
         ' // If it is a user defined type, retrieve its name
         hr = pTypeInfo->GetRefTypeInfo(pVarDesc->elemdescVar.tdesc.hreftype, @pRefTypeInfo)
         IF hr = S_OK AND pRefTypeInfo <> NULL THEN
            DIM bstrVarType AS AFX_BSTR
            hr = pRefTypeInfo->GetDocumentation(-1, @bstrVarType, NULL, NULL, NULL)
            dwsVarType = *bstrVarType
            SysFreeString bstrVarType
            hr = pRefTypeInfo->GetTypeAttr(@pVarTypeAttr)
            IF hr = S_OK AND pVarTypeAttr <> NULL THEN
               IF pVarTypeAttr->typekind = TKIND_ALIAS THEN
                  dwsTypeKind = TLB_TypeKindToStr(pVarTypeAttr->typekind) & " | " & TLB_VarTypeToConstant(pVarTypeAttr->tdescalias.vt)
               ELSE
                  dwsTypeKind = TLB_TypeKindToStr(pVarTypeAttr->typekind)
               END IF
               TreeView_AddItem(hTreeView, hSubNode2, NULL, "TypeKind = " & dwsTypeKind)
               pRefTypeInfo->ReleaseTypeAttr(pVarTypeAttr)
               pVarTypeAttr = NULL
            END IF
            IF pRefTypeInfo THEN pRefTypeInfo->Release
         ELSE
            dwsVarType = "GetRefTypeInfo failed - Error: &h" & HEX(hr, 8)
         END IF
      ELSEIF pVarDesc->elemdescVar.tdesc.vt = VT_PTR THEN
         wIndirectionLevel = 1
         DIM ptdesc AS TYPEDESC PTR = pVarDesc->elemdescVar.tdesc.lptdesc
         DO
            SELECT CASE ptdesc->vt
               ' // If it is another pointer, loop again
               CASE VT_PTR
                  wIndirectionLevel += 1
                  ptdesc = ptdesc->lptdesc
               CASE VT_USERDEFINED
                  hr = pTypeInfo->GetRefTypeInfo(ptdesc->hreftype, @pRefTypeInfo)
                  IF hr = S_OK AND pRefTypeInfo <> NULL THEN
                     DIM bstrVarType AS AFX_BSTR
                     hr = pRefTypeInfo->GetDocumentation(-1, @bstrVarType, NULL, NULL, NULL)
                     dwsVarType = *bstrVarType
                     SysFreeString bstrVarType
                     IF hr = S_OK THEN
                        pRefTypeInfo->GetTypeAttr(@pVarTypeAttr)
                        IF hr = S_OK AND pVarTypeAttr <> NULL THEN
                           IF pVarTypeAttr->typekind = TKIND_ALIAS THEN
                              dwsTypeKind = TLB_TypeKindToStr(pVarTypeAttr->typekind) & " | " & TLB_VarTypeToConstant(pVarTypeAttr->tdescalias.vt)
                           ELSE
                              dwsTypeKind = TLB_TypeKindToStr(pVarTypeAttr->typekind)
                           END IF
                           TreeView_AddItem(hTreeView, hSubNode2, NULL, "TypeKind = " & dwsTypeKind)
                           pRefTypeInfo->ReleaseTypeAttr(pVarTypeAttr)
                           pVarTypeAttr = NULL
                        END IF
                     END IF
                     IF pRefTypeInfo THEN pRefTypeInfo->Release
                     EXIT DO
                  ELSE
                     dwsVarType = "GetRefTypeInfo failed - Error: &h" & HEX(hr, 8)
                  END IF
               CASE ELSE
                  ' // Get the equivalent type
                  dwsVarType = TLB_VarTypeToConstant(ptdesc->vt)
                  dwsFBKeyword = TLB_VarTypeToKeyword(ptdesc->vt)
                  EXIT DO
            END SELECT
         LOOP
      ELSE
         ' // Get the equivalent type
         dwsVarType = TLB_VarTypeToConstant(pVarDesc->elemdescVar.tdesc.vt)
         dwsFBKeyword = TLB_VarTypeToKeyword(pVarDesc->elemdescVar.tdesc.vt)
      END IF

      IF bIsRecord = FALSE THEN
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "VarType = " & dwsVarType)
      ELSE   ' // Records and unions only
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "Indirection level = " & WSTR(wIndirectionLevel))
'           ' // Add the "tag_" prefix to structures and unions
'            IF cbstrTypeKind = "TKIND_RECORD" OR cbstrTypeKind = "TKIND_UNION" THEN cbstrVarType = "tag" & cbstrVarType
         ' // END isn't allowed as a member name
         IF UCASE(dwsVarName) = "END" THEN dwsVarName += "_"
         ' // Use generic data types for enums and interfaces
         IF dwsFBKeyword = "" THEN dwsFBKeyword = dwsVarType
         ' // Add the "tag_" prefix to structures and unions
'         IF dwsTypeKind = "TKIND_RECORD" OR dwsTypeKind = "TKIND_UNION" THEN dwsFBKeyword = "tag_" & dwsFBKeyword
         IF wIndirectionLevel > 0 THEN dwsFBKeyword += " PTR"
'         IF dwsTypeKind = "TKIND_ALIAS | VT_PTR" THEN dwsFBKeyword = "VOID"
         IF dwsTypeKind = "TKIND_ALIAS | VT_PTR" THEN dwsFBKeyword = dwsVarType & " PTR"
         IF dwsTypeKind = "TKIND_ENUM" THEN dwsFBKeyword = dwsVarType
         IF dwsTypeKind = "TKIND_UNKNOWN" THEN dwsFBKeyword = "IUnknown PTR"
         IF dwsTypeKind = "TKIND_DISPATCH" THEN dwsFBKeyword = "IDispatch PTR"
         IF pVarDesc->elemdescVar.tdesc.vt = VT_CARRAY THEN
            DIM padesc AS ARRAYDESC PTR = pVarDesc->elemdescVar.tdesc.lpadesc
            dwsVarType += " | " & TLB_VarTypeToConstant(padesc->tdescElem.vt)
            dwsVarName += " ("
            FOR y AS LONG = 0 TO padesc->cDims - 1
               dwsVarName += WSTR(padesc->rgbounds(y).lLBound) & " TO "
               dwsVarName += WSTR(padesc->rgbounds(y).lLBound + padesc->rgbounds(y).cElements - 1)
               IF padesc->cDims > 1 THEN dwsVarName += ", "
            NEXT
            dwsVarName += ")"
            dwsFBKeyword = TLB_VarTypeToKeyword(padesc->tdescElem.vt)
         END IF
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "VarType = " & dwsVarType)
         IF pVarDesc->elemdescVar.tdesc.vt = VT_CARRAY THEN
            DIM padesc AS ARRAYDESC PTR = pVarDesc->elemdescVar.tdesc.lpadesc
            hSubNode3 = TreeView_AddItem(hTreeView, hSubNode2, NULL, "Dimensions = " & WSTR(padesc->cDims))
            FOR y AS LONG = 0 TO padesc->cDims - 1
               TreeView_AddItem(hTreeView, hSubNode3, NULL, "Dimension " & WSTR(y + 1) & " lower bound = " & WSTR(padesc->rgbounds(y).lLBound))
               TreeView_AddItem(hTreeView, hSubNode3, NULL, "Dimension " & WSTR(y + 1) & " elements = " & WSTR(padesc->rgbounds(y).cElements))
            NEXT
            TreeView_Expand(hTreeView, hSubNode3, TVE_EXPAND)
         END IF
'         ' // FB syntax
         SELECT CASE dwsVarType
            CASE "VT_LPSTR", "VT_CARRAY | VT_LPSTR"
               dwsFBSyntax = dwsVarName & " AS ZSTRING PTR"
            CASE "VT_LPWSTR", "VT_CARRAY | VT_LPWSTR"
               dwsFBSyntax = dwsVarName & " AS WSTRING PTR"
            CASE ELSE
               dwsFBSyntax = dwsVarName & " AS " & dwsFBKeyword
         END SELECT
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "FB syntax = " & dwsFBSyntax)
      END IF

      ' // Expand the nodes
'         TreeView_Expand(hTreeView, hSubNode2, TVE_EXPAND)
      TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
      ' // Release the VARDESC structure
      pTypeInfo->ReleaseVarDesc(pVarDesc) : pVarDesc = NULL

   NEXT

   ' // Just to satisfy the compiler rules; it has no useful meaning
   FUNCTION = S_OK

END FUNCTION
' =====================================================================================
```

# Aliases and typedefs

Some type libraries use aliases and typedefs to give alternate names to data types, enumerations or structures. For example, ADO uses ADO_LONGPTR as a typedef of a LongInteger and SearchDirection as an alias of SearchDirectionEnum.

```
' ----------------------------------------------------------------------------
' Aliases and typedefs
' ----------------------------------------------------------------------------
CASE TKIND_ALIAS
   dwsOrigName = "" : dwsAliasName = "" : dwsAliasName2 = "" : dwsTypedefName = ""
   DIM bstrName AS AFX_BSTR
   hr = m_pTypeLib->GetDocumentation(i, @bstrName, NULL, NULL, NULL)
   dwsName = *bstrName
   SysFreeString bstrName
   IF hr = S_OK THEN
      dwsOrigName = dwsName
      IF pTypeAttr->tdescAlias.vt = VT_USERDEFINED THEN
         ' // If it is a user defined type, retrieve its name
         hr = pTypeInfo->GetRefTypeInfo(pTypeAttr->tdescAlias.hreftype, @pRefTypeInfo)
         IF hr = S_OK AND pRefTypeInfo <> NULL THEN
            DIM bstrName AS AFX_BSTR
            hr = pRefTypeInfo->GetDocumentation(-1, @bstrName, NULL, NULL, NULL)
            dwsName = *bstrName
            SysFreeString bstrName
            IF hr = S_OK THEN
               dwsAliasName = dwsOrigName & " = " & dwsName
               dwsAliasName2 = dwsName & " = " & dwsOrigName
            END IF
            pRefTypeInfo->Release
            pRefTypeInfo = NULL
         END IF
      ELSEIF pTypeAttr->tdescAlias.vt = VT_PTR THEN
         ' // Pointer to a TYPEDESC structure
         DIM ptdesc AS TYPEDESC PTR = pTypeAttr->tdescAlias.lptdesc
         DO
            SELECT CASE ptdesc->vt
               ' // If it is a pointer, do it again
               CASE VT_PTR
                  ptdesc = ptdesc->lptdesc
               CASE VT_USERDEFINED
                  ' // Retrieve the name of the user defined type
                  hr = pTypeInfo->GetRefTypeInfo(ptdesc->hreftype, @pRefTypeInfo)
                  IF hr = S_OK AND pRefTypeInfo <> NULL THEN
                     DIM bstrName AS AFX_BSTR
                     hr = pRefTypeInfo->GetDocumentation(-1, @bstrName, NULL, NULL, NULL)
                     dwsName = *bstrName
                     SysFreeString bstrName
                     IF hr = S_OK THEN
                        dwsAliasName = dwsOrigName & " = " & dwsName
                        dwsAliasName2 = dwsName & " = " & dwsOrigName
                     END IF
                     pRefTypeInfo->Release
                     pRefTypeInfo = NULL
                  END IF
                  EXIT DO
               CASE ELSE
                  ' // Get the equivalent type
                  dwsTypedefName = dwsName & " = " & TLB_VarTypeToConstant(ptdesc->vt) & " <" & TLB_VarTypeToKeyword(pTypeAttr->tdescAlias.vt) & ">"
                  EXIT DO
            END SELECT
         LOOP
      ELSE
         ' // Get the equivalent type
'                  dwsTypedefName = dwsName & " = " & TLB_VarTypeToConstant(pTypeAttr->tdescAlias.vt) & " <" & TLB_VarTypeToKeyword(pTypeAttr->tdescAlias.vt) & ">"
         dwsTypedefName = dwsName & " = " & TLB_VarTypeToKeyword(pTypeAttr->tdescAlias.vt)  & " ' <" & TLB_VarTypeToConstant(pTypeAttr->tdescAlias.vt) & ">"
      END IF
      IF LEN(dwsTypedefName) THEN
         TreeView_AddItem hTreeView, m_hTypedefsNode, NULL, dwsTypedefName
      ELSE
         TreeView_AddItem hTreeView, m_hAliasesNode, NULL, dwsAliasName
'                  TreeView_AddItem hTreeView, m_hAliasesNode, NULL, dwsAliasName2
      END IF
   END IF
' ----------------------------------------------------------------------------
```

# CoClasses

**CoClasses** provide information about each COM class, such the ProgID (Program Identifier), CLSID (Class Identifier), attributes, the in-process server and the implemented interfaces.

```
' ----------------------------------------------------------------------------
' CoClasses
' ----------------------------------------------------------------------------
CASE TKIND_COCLASS
   ' // Get the name of the CoClass
   DIM AS AFX_BSTR bstrName, bstrHelpString, bstrHelpFile
   hr = m_pTypeLib->GetDocumentation(i, @bstrName, @bstrHelpString, @dwHelpContext, @bstrHelpFile)
   dwsName = *bstrName : SysFreeString bstrName
   dwsHelpString = *bstrHelpString : SysFreeString bstrHelpString
   dwsHelpFile = *bstrHelpFile : SysFreeString bstrHelpFile
   hNode = TreeView_AddItem(hTreeView, m_hCoClassesNode, NULL, dwsName)
   ' // ProgIDs node
   ' Some external programs, such McAffee Antivirus, modify the typelibs of
   ' components such Windows Script Host to redirect it to its own server.
   ' This originates duplicate ProgIDs, so we need to search if the ProgID
   ' is already listed to avoid duplicates.
   dwsProgID = TLB_GetProgID(AfxGuidText(pTypeAttr->guid))
   IF LEN(dwsProgID) THEN
      IF TreeView_ItemExists(hTreeView, m_hProgIDsNode, dwsProgID) = FALSE THEN
         TreeView_AddItem hTreeView, m_hProgIDsNode, NULL, dwsProgID
         hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "ProgID")
         TreeView_AddItem hTreeView, hSubNode, NULL, dwsProgID
         TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
      END IF
   END IF
   ' // Version independent ProgIDs node
   ' Note: Search if it already exists because there are components like
   ' MSXML that allow side-by-side installation of several versions that have
   ' different ProgIDs but, of course, the same independent version ProgID.
   dwsProgID = TLB_GetVersionIndependentProgID(AfxGuidText(pTypeAttr->guid))
   IF LEN(dwsProgID) THEN
      IF TreeView_ItemExists(hTreeView, m_hVerIndProgIDsNode, dwsProgID) = FALSE THEN
         TreeView_AddItem hTreeView, m_hVerIndProgIDsNode, NULL, dwsProgID
         hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Version independent ProgID")
         TreeView_AddItem hTreeView, hSubNode, NULL, dwsProgID
         TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
      END IF
   END IF
   ' // ClsIDs nodes
   TreeView_AddItem hTreeView, m_hClsIDsNode, NULL, "CLSID_" & dwsName & " = " & CHR(34) & AfxGuidText(pTypeAttr->guid) & CHR(34)
   hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "CLSID")
   TreeView_AddItem hTreeView, hSubNode, NULL, AfxGuidText(pTypeAttr->guid)
   TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
   ' // Attributes
   hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Attributes")
   TreeView_AddItem hTreeView, hSubNode, NULL, WSTR(pTypeAttr->wTypeFlags) & " [&h" & HEX(pTypeAttr->wTypeFlags, 8) & "]" & TLB_InterfaceFlagsToStr(pTypeAttr->wTypeFlags)
   TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
   ' // Help info
   IF LEN(dwsHelpString) THEN
      hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Help string")
      TreeView_AddItem hTreeView, hSubNode, NULL, dwsHelpString
      TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
   END IF
   IF dwHelpContext THEN
      hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Help context")
      TreeView_AddItem hTreeView, hSubNode, NULL, WSTR(dwHelpContext) & " [&h" & HEX(dwHelpContext, 8) & "]"
      TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
   END IF
   IF LEN(dwsHelpFile) THEN
      hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Help file")
      TreeView_AddItem hTreeView, hSubNode, NULL, dwsHelpFile
      TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
   END IF
   ' // InProcServer32
   dwsInProcServer = TLB_GetInprocServer32(AfxGuidText(pTypeAttr->guid))
   IF LEN(dwsInProcServer) THEN
      hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "InProcServer32")
      IF INSTR(dwsInProcServer, "%") THEN
         DIM wszSrc AS WSTRING * MAX_PATH, wszDest AS WSTRING * MAX_PATH, cbLen AS DWORD
         wszSrc = dwsInProcServer
         cbLen = ExpandEnvironmentStringsW(@wszSrc, @wszDest, MAX_PATH)
         IF cbLen THEN dwsInProcServer = wszDest
      END IF
      TreeView_AddItem hTreeView, hSubNode, NULL, dwsInProcServer
      TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
   END IF
   ' // Retrieve the implemented interfaces
   ' Note: Don't release pRefType or it will explode
   cImplTypes = pTypeAttr->cImplTypes
   IF cImplTypes THEN hImplIntSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Implemented interfaces")
   FOR x AS LONG = 0 TO cImplTypes - 1
      lImplTypeFlags = 0
      hr = pTypeInfo->GetImplTypeFlags(x, @lImplTypeFlags)
      IF hr <> S_OK THEN EXIT FOR
      pRefType = 0
      hr = pTypeInfo->GetRefTypeOfImplType(x, @pRefType)
      IF hr <> S_OK THEN EXIT FOR
      hr = pTypeInfo->GetRefTypeInfo(pRefType, @pImplTypeInfo)
      IF hr <> S_OK OR pImplTypeInfo = NULL THEN EXIT FOR
      DIM bstrName AS AFX_BSTR
      hr = pImplTypeInfo->GetDocumentation(-1, @bstrName, NULL, NULL, NULL)
      dwsName = *bstrName
      SysFreeString bstrName
      IF hr <> S_OK THEN EXIT FOR
      TreeView_AddItem hTreeView, hImplIntSubNode, NULL, dwsName
      TreeView_Expand(hTreeView, hImplIntSubNode, TVE_EXPAND)
      pImplTypeAttr = 0
      pImplTypeInfo->GetTypeAttr(@pImplTypeAttr)
      IF lImplTypeFlags = IMPLTYPEFLAG_FDEFAULT THEN   ' // Default interface
         hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Default interface")
         TreeView_AddItem hTreeView, hSubNode, NULL, dwsName
         TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
         hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Default interface IID")
         IF pImplTypeAttr THEN TreeView_AddItem hTreeView, hSubNode, NULL, AfxGuidText(pImplTypeAttr->guid)
         TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
      ELSEIF lImplTypeFlags = IMPLTYPEFLAG_FSOURCE THEN   ' // Events interface
         ' // Some components, such Office 12's AccWiz.dll, have deprecated CoClasses that
         ' // implement the same events interfaces that the new one. We need to check if the
         ' // interface is hidden to avoid listing them twice.
'                  IF (pTypeAttr->wTypeFlags AND TYPEFLAG_FHIDDEN) <> TYPEFLAG_FHIDDEN THEN
'                     IF TreeView_ItemExists(hTreeView, m_hEventsNode, dwsName) = FALSE THEN
'                        TreeView_AddItem hTreeView, m_hEventsNode, NULL, dwsName
'                     END IF
'                  END IF
      ELSEIF lImplTypeFlags = (IMPLTYPEFLAG_FDEFAULT OR IMPLTYPEFLAG_FSOURCE) THEN   ' // Default events interface
         hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Default events interface")
         TreeView_AddItem hTreeView, hSubNode, NULL, dwsName
         TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
         hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Default events interface IID")
         IF pImplTypeAttr THEN TreeView_AddItem hTreeView, hSubNode, NULL, AfxGuidText(pImplTypeAttr->guid)
         TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
         ' // Some components, such Office 12's AccWiz.dll, have deprecated CoClasses that
         ' // implement the same events interfaces that the new one. We need to check if the
         ' // interface is hidden to avoid listing them twice.
'                  IF (pTypeAttr->wTypeFlags AND TYPEFLAG_FHIDDEN) <> TYPEFLAG_FHIDDEN THEN
'                     IF TreeView_ItemExists(hTreeView, m_hEventsNode, dwsName) = FALSE THEN
'                        TreeView_AddItem hTreeView, m_hEventsNode, NULL, dwsName
'                     END IF
'                  END IF
      END IF
      IF pImplTypeAttr THEN
         IF pImplTypeInfo THEN pImplTypeInfo->ReleaseTypeAttr(pImplTypeAttr)
         pImplTypeAttr = NULL
      END IF
   NEXT
   IF pImplTypeAttr THEN
      IF pImplTypeInfo THEN pImplTypeInfo->ReleaseTypeAttr(pImplTypeAttr)
      pImplTypeAttr = NULL
   END IF
   IF pImplTypeInfo THEN
      pImplTypeInfo->Release
      pImplTypeInfo = NULL
   END IF
' ----------------------------------------------------------------------------
```

# Interfaces

The type infos TKIND_INTERFACE and TKIND_DISPATCH provide information about the implemented interfaces and its methods and properties.

```
' ----------------------------------------------------------------------------
' Interfaces
' ----------------------------------------------------------------------------
CASE TKIND_INTERFACE, TKIND_DISPATCH
   DIM AS AFX_BSTR bstrName, bstrHelpString
   hr = m_pTypeLib->GetDocumentation(i, @bstrName, @bstrHelpString, NULL, NULL)
   dwsName = *bstrName
   SysFreeString bstrName
   dwsHelpString = *bstrHelpString
   SysFreeString bstrHelpString
   IF hr = S_OK THEN
      TreeView_AddItem hTreeView, m_hIIDsNode, NULL, "IID_" & dwsName & " = " & CHR(34) & AfxGuidText(pTypeAttr->guid) & CHR(34)
      ' ------------------------------------------------------------------------------------------
      ' If it is a dual interface and the VTable option has been set, try to change the view.
      ' ------------------------------------------------------------------------------------------
      DIM VTableView AS BOOLEAN = m_VTableView
      IF pTKind = TKIND_DISPATCH AND (pTypeAttr->wTypeFlags AND TYPEFLAG_FDUAL) = TYPEFLAG_FDUAL THEN
         IF VTableView = TRUE THEN
            DO   ' // Fake DO LOOP to allow exit without GOTO
               ' // Attempt to change the view to VTable
               pRefType = NULL
               hr = pTypeInfo->GetRefTypeOfImplType(-1, @pRefType)
               IF hr <> S_OK OR pRefType = NULL THEN
                  VTableView = FALSE
                  EXIT DO
               END IF
               hr = pTypeInfo->GetRefTypeInfo(pRefType, @pRefTypeInfo)
               IF hr <> S_OK OR pRefTypeInfo = NULL THEN
                  VTableView = FALSE
                  EXIT DO
               END IF
               pRefTypeAttr = NULL
               hr = pRefTypeInfo->GetTypeAttr(@pRefTypeAttr)
               hSubNode = TreeView_AddItem(hTreeView, m_hDualIntNode, NULL, dwsName)
               IF AfxGuidText(pRefTypeAttr->guid) <> "{00000000-0000-0000-0000-000000000000}" THEN
                  TreeView_AddItem(hTreeView, hSubNode, NULL, "IID: " & AfxGuidText(pRefTypeAttr->guid))
               END IF
               IF LEN(dwsHelpString) THEN TreeView_AddItem(hTreeView, hSubNode, NULL, "Documentation string: " & dwsHelpString)
               IF pRefTypeAttr->wTypeFlags THEN TreeView_AddItem(hTreeView, hSubNode, NULL, "Attributes =  " & WSTR(pRefTypeAttr->wTypeFlags) & " [&h" & HEX(pRefTypeAttr->wTypeFlags, 8) & "]" & TLB_InterfaceFlagsToStr(pRefTypeAttr->wTypeFlags))
               dwsInheritedInterface = TLB_GetImplementedInterface(pRefTypeInfo)
               IF LEN(dwsInheritedInterface) THEN TreeView_AddItem(hTreeView, hSubNode, NULL, "Inherited interface = " & dwsInheritedInterface)
               ' /*** Datamembers ***/
               IF pRefTypeAttr->cVars THEN
                  hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "Number of datamembers = " & WSTR(pRefTypeAttr->cVars))
                  this.GetMembers (pRefTypeInfo, pRefTypeAttr, hTreeView, hSubNode2)
               END IF
               ' /*** Retrieves the methods and properties ***/
               IF @pRefTypeAttr->cFuncs THEN
                  hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "Number of methods = " & WSTR(pRefTypeAttr->cFuncs))
                  this.GetFunctions(pRefTypeInfo, pREfTypeAttr, hTreeView, hSubNode2, VTableView, TRUE, pTKind, dwsImplInterface)
               ELSE
                  hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "Number of methods = 0")
               END IF
               IF pRefTypeInfo THEN
                  IF pTypeAttr THEN pRefTypeInfo->ReleaseTypeAttr(pRefTypeAttr)
                  pRefTypeInfo->Release
               END IF
               ' // exit the fake loop
               EXIT DO
            LOOP
         END IF
      ELSE
         VTableView = FALSE
      END IF
      ' ------------------------------------------------------------------------------------------
      ' ...else use the default view
      ' ------------------------------------------------------------------------------------------
      IF pTKind = TKIND_INTERFACE OR pTKind = TKIND_DISPATCH AND CLNG(VTableView) = FALSE THEN
         dwsImplInterface = TLB_GetImplementedInterface(pTypeInfo)
         IF dwsImplInterface <> "" THEN
            IF UCASE(dwsImplInterface) <> "IUNKNOWN" AND UCASE(dwsImplInterface) <> "IDISPATCH" THEN
               dwsImplInterface = TLB_GetBaseClass(m_pTypeLib, dwsName)
            END IF
         END IF
         IF pTKind = TKIND_INTERFACE THEN
            IF UCASE(dwsImplInterface) = "IUNKNOWN" AND (pTypeAttr->wTypeFlags AND TYPEFLAG_FOLEAUTOMATION) = TYPEFLAG_FOLEAUTOMATION THEN
               hSubNode = TreeView_AddItem(hTreeView, m_hOleAutIntNode, NULL, dwsName)
            ELSEIF UCASE(dwsImplInterface) = "IDISPATCH" AND (pTypeAttr->wTypeFlags AND TYPEFLAG_FDUAL) <> TYPEFLAG_FDUAL THEN
               hSubNode = TreeView_AddItem(hTreeView, m_hDispblIntNode, NULL, dwsName)
            ELSE
               hSubNode = TreeView_AddItem(hTreeView, m_hIntNode, NULL, dwsName)
            END IF
         ELSEIF pTKind = TKIND_DISPATCH THEN
            IF (pTypeAttr->wTypeFlags AND TYPEFLAG_FDUAL) = TYPEFLAG_FDUAL THEN
               hSubNode = TreeView_AddItem(hTreeView, m_hDualIntNode, NULL, dwsName)
            ELSEIF INSTR(dwsEvents, "#" & dwsName & "#") THEN
               hSubNode = TreeView_AddItem(hTreeView, m_hEventsNode, NULL, dwsName)
            ELSE
               hSubNode = TreeView_AddItem(hTreeView, m_hDispIntNode, NULL, dwsName)
            END IF
         END IF
         IF AfxGuidText(pTypeAttr->guid) <> "{00000000-0000-0000-0000-000000000000}" THEN
            TreeView_AddItem(hTreeView, hSubNode, NULL, "IID: " & AfxGuidText(pTypeAttr->guid))
         END IF
         IF LEN(dwsHelpString) THEN TreeView_AddItem(hTreeView, hSubNode, NULL, "Documentation string: " & dwsHelpString)
         IF pTypeAttr->wTypeFlags THEN TreeView_AddItem(hTreeView, hSubNode, NULL, "Attributes =  " & WSTR(pTypeAttr->wTypeFlags) & " [&h" & HEX(pTypeAttr->wTypeFlags, 8) & "]" & TLB_InterfaceFlagsToStr(pTypeAttr->wTypeFlags))
         dwsInheritedInterface = TLB_GetImplementedInterface(pTypeInfo)
         IF LEN(dwsInheritedInterface) THEN TreeView_AddItem(hTreeView, hSubNode, NULL, "Inherited interface = " & dwsInheritedInterface)
         ' /*** Datamembers ***/
         IF pTypeAttr->cVars THEN
            hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "Number of datamembers = " & WSTR(pTypeAttr->cVars))
            this.GetMembers (pTypeInfo, pTypeAttr, hTreeView, hSubNode2)
         END IF
         ' /*** Retrieves the methods and properties ***/
         IF pTypeAttr->cFuncs THEN
            hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "Number of methods = " & WSTR(pTypeAttr->cFuncs))
            IF pTKind = TKIND_INTERFACE THEN
               this.GetFunctions(pTypeInfo, pTypeAttr, hTreeView, hSubNode2, TRUE, TRUE, pTKind, dwsImplInterface)
            ELSE
               this.GetFunctions(pTypeInfo, pTypeAttr, hTreeView, hSubNode2, VTableView, TRUE, pTKind, dwsImplInterface)
            END IF
         ELSE
            hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "Number of methods = 0")
         END IF
      END IF
   END IF
   ' ----------------------------------------------------------------------------
```

A very important particularity is that the information can be returned in two different kind of views, the **VTable** view and the **Automation** view.

To change the type of views from the default Automation one to the VTable one, we have to call the **GetRefTypeOfImplType** of the **ITypeInfo** interface. The meager documentation provided by Microsoft states that "If a type description describes a COM class, it retrieves the type description of the implemented interface types. For an interface, GetRefTypeOfImplType returns the type information for inherited interfaces, if any exist." See: https://msdn.microsoft.com/en-us/library/windows/desktop/ms221569(v=vs.85).aspx

There is a remark at the bottom: "If the TKIND_DISPATCH type description is for a dual interface, the TKIND_INTERFACE type description can be obtained by calling **GetRefTypeOfImplType** with an indexof –1, and by passing the returned *pRefTypehandle* to **GetRefTypeInfo** to retrieve the type information."

So, if we have a TKIND_DISPATCH description and be want a TKIND_INTERFACE description (assuming that the Dispatch interface is dual and not a dispatch only interface), we can get it passing -1 to **GetRefTypeOfImplType**.

```
' // Attempt to change the view to VTable
pRefType = NULL
hr = pTypeInfo->GetRefTypeOfImplType(-1, @pRefType)
IF hr <> S_OK OR pRefType = NULL THEN
   VTableView = FALSE
   EXIT DO
END IF
hr = pTypeInfo->GetRefTypeInfo(pRefType, @pRefTypeInfo)
IF hr <> S_OK OR pRefTypeInfo = NULL THEN
   VTableView = FALSE
   EXIT DO
END IF
pRefTypeAttr = NULL
hr = pRefTypeInfo->GetTypeAttr(@pRefTypeAttr)
```

# Retrieving the methods and properties

The *cFuncs* member of the **TYPEATTR** structure contains the number of methods and properties implemented in an interface and the **GetFuncDesc** method of the **ITypeInfo** interface retrieves the **FUNCDESC** structure that contains information about a specified function ( https://msdn.microsoft.com/en-us/library/windows/desktop/ms221425(v=vs.85).aspx ), as well as the return type.

```
' =====================================================================================
' Retrieve information about the methods, properties and functions.
' =====================================================================================
FUNCTION CParseTypeLib.GetFunctions (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL pTypeAttr AS TYPEATTR PTR, _
   BYVAL hTreeView AS HWND, BYVAL hSubNode AS HTREEITEM, BYVAL bVTableView AS BOOLEAN, BYVAL bIsMethod AS BOOLEAN = FALSE, _
   BYVAL pTKind AS TYPEKIND = -1, BYVAL pwszImplInterface AS WSTRING PTR = NULL) AS HRESULT

   DIM hSubNode2 AS HTREEITEM              ' // Sub node handle
   DIM dwHelpContext AS DWORD              ' // Help context number
   DIM pRefTypeInfo AS Afx_ITypeInfo PTR   ' // Referenced TypeInfo interface
   DIM pReturnTypeAttr AS TYPEATTR PTR     ' // Referenced TYPEATTR structure
   DIM ReturnTypeKind AS TYPEKIND          ' // Return value type kind
   DIM dwsName AS DWSTRING                 ' // Name
   DIM dwsHelpString AS DWSTRING           ' // Help string
   DIM dwsDllName AS DWSTRING              ' // DLL name
   DIM dwsEntryPoint AS DWSTRING           ' // Entry point
   DIM dwsType AS DWSTRING                 ' // Type

   IF pTypeInfo = NULL OR pTypeAttr = NULL THEN RETURN E_INVALIDARG

   FOR x AS LONG = 0 TO pTypeAttr->cFuncs - 1
      ' // Gets a reference to the FuncDesc structure
      DIM pFuncDesc AS FUNCDESC PTR
      DIM hr AS HRESULT = pTypeInfo->GetFuncDesc(x, @pFuncDesc)
      IF hr <> S_OK OR pFuncDesc = NULL THEN EXIT FOR
      ' // Retrieve the name
      DIM AS AFX_BSTR bstrName, bstrHelpString
      pTypeInfo->GetDocumentation(pFuncDesc->memid, @bstrName, @bstrHelpString, @dwHelpContext, NULL)
      dwsName = *bstrName
      SysFreeString bstrName
      dwsHelpString = *bstrHelpString
      SysFreeString bstrHelpString
      IF bIsMethod THEN
         ' ------------------------------------------------------------------
         ' Workaround for libraries that can have illegal method names.
         ' For example, TLBINF32.DLL has a property called GetTypeInfo.
         ' ------------------------------------------------------------------
         DIM vtOffset AS LONG
#ifdef __FB_64BIT__
         vtOffset = 48
#else
         vtOffset = 24
#endif
         IF UCASE(dwsName) = "QUERYINTERFACE" AND pFuncdesc->oVft > vtOffset THEN dwsName += "_"
         IF UCASE(dwsName) = "ADDREF" AND pFuncdesc->oVft > vtOffset THEN dwsName += "_"
         IF UCASE(dwsName) = "RELEASE" AND pFuncdesc->oVft > vtOffset THEN dwsName += "_"
         IF UCASE(dwsName) = "GETTYPEINFOCOUNT" AND pFuncdesc->oVft > vtOffset THEN dwsName += "_"
         IF UCASE(dwsName) = "GETTYPEINFO" AND pFuncdesc->oVft > vtOffset THEN dwsName += "_"
         IF UCASE(dwsName) = "GETIDSOFNAMES" AND pFuncdesc->oVft > vtOffset THEN dwsName += "_"
         IF UCASE(dwsName) = "INVOKE" AND pFuncdesc->oVft > vtOffset THEN dwsName += "_"
         IF UCASE(dwsName) = "DELETE" THEN dwsName += "_"
         IF UCASE(dwsName) = "PROPERTY" THEN dwsName += "_"

         IF pTKind = TKIND_INTERFACE OR pTKind = TKIND_DISPATCH THEN
            IF pFuncDesc->invkind = INVOKE_FUNC THEN dwsName = "METHOD " & dwsName
            IF pFuncDesc->invkind = INVOKE_PROPERTYGET THEN dwsName = "PROPERTY GET " & dwsName
            IF pFuncDesc->invkind = INVOKE_PROPERTYPUT THEN dwsName = "PROPERTY PUT " & dwsName
            IF pFuncDesc->invkind = INVOKE_PROPERTYPUTREF THEN dwsName = "PROPERTY PUTREF " & dwsName
         END IF
         hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, dwsName)
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "VTable offset = " &  WSTR(pFuncdesc->oVft) & " [&h" & HEX(pFuncdesc->oVft, 8) & "]")
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "DispID = " & WSTR(pFuncDesc->memid) & " [&h" & HEX(pFuncDesc->memid, 8) & "]")
         IF LEN(dwsHelpString) THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Help string = " & dwsHelpString)
         IF dwHelpContext THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Help context = " & WSTR(dwHelpContext))
      ELSE
         IF pFuncDesc->elemdescFunc.tdesc.vt = VT_VOID THEN
            hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "SUB " & dwsName)
         ELSE
            hSubNode2 = TreeView_AddItem(hTreeView, hSubNode, NULL, "FUNCTION " & dwsName)
         END IF
         IF LEN(dwsHelpString) THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Help string = " & dwsHelpString)
         IF dwHelpContext THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Help context = " & WSTR(dwHelpContext))
         TreeView_AddItem(hTreeView, hSubNode2, NULL, "DispID = " & WSTR(pFuncDesc->memid) & " [&h" & HEX(pFuncDesc->memid, 8) & "]")
         DIM wOrdinal AS WORD, bstrDllName AS AFX_BSTR, bstrEntryPoint AS AFX_BSTR
         hr = pTypeInfo->GetDllEntry(pFuncDesc->memid, pFuncDesc->invkind, @bstrDllName, @bstrEntryPoint, @wOrdinal)
         dwsDllName = *bstrDllName
         SysFreeString bstrDllName
         dwsEntryPoint = *bstrEntryPoint
         SysFreeString bstrEntryPoint
         IF hr = S_OK THEN
            IF LEN(dwsDllName) THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "DLL name = " & dwsDllName)
            IF LEN(dwsEntryPoint) THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Entry point = " & dwsEntryPoint)
            IF wOrdinal THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Ordinal = " & WSTR(wOrdinal))
         END IF
      END IF

      ' // Kind of function
      SELECT CASE pFuncDesc->funckind
         CASE FUNC_VIRTUAL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "FuncKind = Virtual")
         CASE FUNC_PUREVIRTUAL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "FuncKind = Pure virtual")
         CASE FUNC_NONVIRTUAL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "FuncKind = Non virtual")
         CASE FUNC_STATIC
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "FuncKind = Static")
         CASE FUNC_DISPATCH
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "FuncKind = Dispatch")
      END SELECT
      ' // Invoke kind
      SELECT CASE pFuncDesc->invkind
         CASE INVOKE_FUNC
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "InvokeKind = Function")
         CASE INVOKE_PROPERTYGET
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "InvokeKind = Get property")
         CASE INVOKE_PROPERTYPUT
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "InvokeKind = Put property")
         CASE INVOKE_PROPERTYPUTREF
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "InvokeKind = PutRef property")
      END SELECT
      ' // Calling convention
      SELECT CASE pFuncDesc->callconv
         CASE CC_FASTCALL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "Calling convention = FASTCALL")
         CASE CC_CDECL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "Calling convention = CDECL")
         CASE CC_MACPASCAL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "Calling convention = MACPASCAL")
         CASE CC_STDCALL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "Calling convention = STDCALL")
         CASE CC_FPFASTCALL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "Calling convention = FPFASTCALL")
         CASE CC_SYSCALL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "Calling convention = SYSCALL")
         CASE CC_MPWCDECL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "Calling convention = MPWCDECL")
         CASE CC_MPWPASCAL
            TreeView_AddItem(hTreeView, hSubNode2, NULL, "Calling convention = MPWPASCAL")
      END SELECT

      ' // More general information
      IF pFuncDesc->cParamsOpt THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Number of optional variant parameters = " & WSTR(pFuncDesc->cParamsOpt))
      IF pFuncDesc->cScodes THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Count of permitted return values = " & WSTR(pFuncDesc->cScodes))
      IF pFuncDesc->wFuncFlags THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Attributes = " & WSTR(pFuncDesc->wFuncFlags)& " [&h" & HEX(pFuncDesc->wFuncFlags, 8) & "]" & TLB_FuncFlagsToStr(pFuncDesc->wFuncFlags))
      ' // Return type
      ReturnTypeKind = -1  ' // Because the TYPEKIND enum starts at 0
      IF pFuncDesc->elemdescFunc.tdesc.vt = VT_USERDEFINED THEN
         ' // If it is a user defined type, retrieve its name
         hr = pTypeInfo->GetRefTypeInfo(pFuncDesc->elemdescFunc.tdesc.hreftype, @pRefTypeInfo)
         IF hr = S_OK AND pRefTypeInfo <> NULL THEN
            DIM bstrType AS AFX_BSTR
            hr = pRefTypeInfo->GetDocumentation(-1, @bstrType, NULL, NULL, NULL)
            dwsType = *bstrType
            SysFreeString bstrType
            hr = pRefTypeInfo->GetTypeAttr(@pReturnTypeAttr)
            IF hr = S_OK AND pReturnTypeAttr <> NULL THEN
               TreeView_AddItem(hTreeView, hSubNode2, NULL, "Return type typeKind = " & TLB_TypeKindToStr(pReturnTypeAttr->typekind))
               ReturnTypeKind = pReturnTypeAttr->typekind
               pRefTypeInfo->ReleaseTypeAttr(pReturnTypeAttr)
               pReturnTypeAttr = NULL
            END IF
            IF pRefTypeInfo THEN pRefTypeInfo->Release
         END IF
      ELSEIF pFuncDesc->elemdescFunc.tdesc.vt = VT_PTR THEN
         ' // Pointer to a TYPEDESC structure
         DIM ptdesc AS TYPEDESC PTR = pFuncDesc->elemdescFunc.tdesc.lptdesc
         DO
            SELECT CASE ptdesc->vt
               ' // If it is a pointer, do it again
               CASE VT_PTR
                  ptdesc = ptdesc->lptdesc
               CASE VT_USERDEFINED
                  ' // Retrieve the name of the user defined type
                  hr = pTypeInfo->GetRefTypeInfo(ptdesc->hreftype, @pRefTypeInfo)
                  IF hr = S_OK AND pRefTypeInfo <> NULL THEN
                     DIM bstrType AS AFX_BSTR
                     hr = pRefTypeInfo->GetDocumentation(-1, @bstrType, NULL, NULL, NULL)
                     dwsType = *bstrType
                     SysFreeString bstrType
                     hr = pRefTypeInfo->GetTypeAttr(@pReturnTypeAttr)
                     IF hr = S_OK AND pReturnTypeAttr <> NULL THEN
                        TreeView_AddItem(hTreeView, hSubNode2, NULL, "Return type typeKind = " & TLB_TypeKindToStr(pReturnTypeAttr->typekind))
                        ReturnTypeKind = pReturnTypeAttr->typekind
                        pRefTypeInfo->ReleaseTypeAttr(pReturnTypeAttr)
                        pReturnTypeAttr = NULL
                     END IF
                     IF pRefTypeInfo THEN pRefTypeInfo->Release
                  END IF
                  EXIT DO
               CASE ELSE
                  ' // Get the equivalent type
                  dwsType = TLB_VarTypeToConstant(ptdesc->vt)
                  EXIT DO
            END SELECT
         LOOP
      ELSE
         ' // Get the equivalent type
         dwsType = TLB_VarTypeToConstant(pFuncDesc->elemdescFunc.tdesc.vt)
      END IF
      ' // Return type
      TreeView_AddItem(hTreeView, hSubNode2, NULL, "Return type = " & dwsType)
      DIM strReturn AS STRING = ""
      IF bVTableView = FALSE THEN
         IF ReturnTypeKind = TKIND_INTERFACE OR ReturnTypeKind = TKIND_DISPATCH THEN
            strReturn = "BYVAL rhs AS " & dwsType & " PTR PTR"
         ELSEIF ReturnTypeKind = TKIND_ENUM THEN
            strReturn = "BYVAL rhs AS " & dwsType & " PTR"
         ELSEIF ReturnTypeKind = TKIND_ALIAS THEN
            strReturn = "BYVAL rhs AS " & dwsType & " PTR"
         ELSEIF pFuncDesc->elemdescFunc.tdesc.vt = VT_VOID THEN
            ' // With Automation view, VT_VOID means no return type
            strReturn = ""
         ELSEIF pFuncDesc->invkind <> INVOKE_PROPERTYGET AND pFuncDesc->cParams = 0 THEN
            ' // With Automation view, if it is not a get property and it has not
            ' // parameters, then it has not an OUT return value but returns the value
            ' //  directly as the result of the method, e.g. AddRef and Release.
            strReturn = ""
         ELSE
            ' // Returns the value as an OUT parameter
            strReturn = "BYVAL rhs AS " & TLB_VarTypeToKeyword(pFuncDesc->elemdescFunc.tdesc.vt) & " PTR"
         END IF
         IF LEN(strReturn) THEN TreeView_AddItem(hTreeView, hSubNode2, NULL, "Return type FB syntax = " & strReturn)
      END IF
      ' // Parameters
      IF pFuncDesc->cParams THEN this.GetParameters(pTypeInfo, pFuncDesc, hTreeView, hSubNode2, bVTableView)
      ' // Expand the nodes
'      TreeView_Expand(hTreeView, hSubNode2, TVE_EXPAND)
      TreeView_Expand(hTreeView, hSubNode, TVE_EXPAND)
      ' // Release the FUNCDESC structure
      pTypeInfo->ReleaseFuncDesc(pFuncDesc) : pFuncDesc = NULL
      ReturnTypeKind = -1
   NEXT

   ' // Just to satisfy the compiler rules; it has no useful meaning
   FUNCTION = S_OK

END FUNCTION
' =====================================================================================
```

# Retrieving the parameters

One of the members of the FUNDESC structure, cParams, contains the number of parameters of each function or method. Parameters can be of any kind of data type and be passed by value or by reference.

```
' =====================================================================================
' Retrieve information about the parameters
' =====================================================================================
FUNCTION CParseTypeLib.GetParameters (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL pFuncDesc AS FUNCDESC PTR, _
   BYVAL hTreeView AS HWND, BYVAL hSubNode2 AS HTREEITEM, BYVAL bVTableView AS BOOLEAN) AS HRESULT

   DIM hParamsNode AS HTREEITEM            ' // Parameters node
   DIM hParamNameNode AS HTREEITEM         ' // Parameter name node
   DIM pParamTypeAttr AS TYPEATTR PTR      ' // Referenced TYPEATTR structure
   DIM pReturnTypeAttr AS TYPEATTR PTR     ' // Referenced TYPEATTR structure
   DIM wIndirectionLevel AS WORD           ' // Indirection level
   DIM pRefTypeInfo AS Afx_ITypeInfo PTR   ' // Referenced TypeInfo interface
   DIM dwsParamName AS DWSTRING            ' // Parameter name
   DIM dwsVarType AS DWSTRING              ' // Variable type
   DIM dwsTypeKind AS DWSTRING             ' // Type kind
   DIM dwsFBKeyword AS DWSTRING            ' // FB keyword
   DIM dwsFBSyntax AS DWSTRING             ' // FB syntax

   hParamsNode = TreeView_AddItem(hTreeView, hSubNode2, NULL, "Number of parameters = " & WSTR(pFuncDesc->cParams))
   ' ----------------------------------------------------------------------------------
   ' Gets the name of all the parameters.
   ' The first one is the name of the function.
   ' If the member ID identifies a property that is implemented with property functions,
   ' the property name is returned. For property get functions, the names of the function
   ' and its parameters are always returned.
   ' For property put and put reference functions, the right side of the assignment is
   ' unnamed. If cMaxNames is less than is required to return all of the names of the
   ' parameters of a function, then only the names of the first cMaxNames-1 parameters
   ' are returned. The names of the parameters are returned in the array in the same
   ' order that they appear elsewhere in the interface (for example, the same order in
   ' the parameter array associated with the FUNCDESC enumeration).
   ' ----------------------------------------------------------------------------------
   REDIM rgbstrNames(pFuncDesc->cParams) AS AFX_BSTR
   DIM cNames AS DWORD   ' // Number of names
   DIM hr AS HRESULT = pTypeInfo->GetNames(pFuncDesc->memid, @rgbstrNames(0), pFuncDesc->cParams + 1, @cNames)
   IF hr = S_OK THEN
      ' // Pointer to an ELEMDESC structure
      DIM pParam AS ELEMDESC PTR = pFuncDesc->lprgelemdescParam
      ' // Retrieves information about the parameters
      FOR y AS LONG = 0 TO pFuncDesc->cParams - 1
         dwsVarType = "" : dwsTypeKind = "" : dwsFBKeyword = ""
         ' // Attributes
         DIM wFlags AS WORD = pParam[y].paramdesc.wParamFlags
         ' // When using the automation view, it does not return a name for the return type
         dwsParamName = rgbstrNames(y + 1)
         IF LEN(dwsParamName) = 0 THEN
            IF y = pFuncDesc->cParams - 1 THEN
               dwsParamName = "rhs"
            ELSE
               dwsParamName = "prm" & WSTR(y + 1)
            END IF
         END IF
         hParamNameNode = TreeView_AddItem(hTreeView, hParamsNode, NULL, dwsParamName)
         TreeView_AddItem(hTreeView, hParamNameNode, NULL, "Attributes = " & WSTR(wFlags) & " [&h" & HEX(wFlags, 8) & "] " & TLB_ParamflagsToStr(wFlags))
         wIndirectionLevel = 0
         IF pParam[y].tdesc.vt = VT_USERDEFINED THEN
            ' // If it is a user defined type, retrieve its name
            hr = pTypeInfo->GetRefTypeInfo(pParam[y].tdesc.hreftype, @pRefTypeInfo)
            IF hr = S_OK AND pRefTypeInfo <> NULL THEN
               DIM bstrVarType AS AFX_BSTR
               hr = pRefTypeInfo->GetDocumentation(-1, @bstrVarType, NULL, NULL, NULL)
               dwsVarType = *bstrVarType
               SysFreeString bstrVarType
               hr = pRefTypeInfo->GetTypeAttr(@pParamTypeAttr)
               IF hr = S_OK AND pParamTypeAttr <> NULL THEN
                  IF pParamTypeAttr->typekind = TKIND_ALIAS THEN
                     dwsTypeKind = TLB_TypeKindToStr(pParamTypeAttr->typekind) & " | " & TLB_VarTypeToConstant(pParamTypeAttr->tdescalias.vt)
                  ELSE
                     dwsTypeKind = TLB_TypeKindToStr(pParamTypeAttr->typekind)
                  END IF
                  TreeView_AddItem(hTreeView, hParamNameNode, NULL, "TypeKind = " & dwsTypeKind)
                  pRefTypeInfo->ReleaseTypeAttr(pParamTypeAttr)
                  pParamTypeAttr = NULL
               END IF
               IF pRefTypeInfo THEN pRefTypeInfo->Release
            END IF
         ELSEIF pParam[y].tdesc.vt = VT_PTR THEN
            ' // Pointer to a TYPEDESC structure
            DIM ptdesc AS TYPEDESC PTR = pParam[y].tdesc.lptdesc
            wIndirectionLevel = 1
            DO
               SELECT CASE ptdesc->vt
                  ' // If it is a pointer, do it again
                  CASE VT_PTR
                     wIndirectionLevel += 1
                     ptdesc = ptdesc->lptdesc
                  CASE VT_USERDEFINED
                     ' // Retrieve the name of the user defined type
                     hr = pTypeInfo->GetRefTypeInfo(ptdesc->hreftype, @pRefTypeInfo)
                     IF hr = S_OK AND pRefTypeInfo <> NULL THEN
                        DIM bstrVarType AS AFX_BSTR
                        hr = pRefTypeInfo->GetDocumentation(-1, @bstrVarType, NULL, NULL, NULL)
                        dwsVarType = *bstrVarType
                        SysFreeString bstrVarType
                        hr = pRefTypeInfo->GetTypeAttr(@pParamTypeAttr)
                        IF hr = S_OK AND pParamTypeAttr <> NULL THEN
                           IF pParamTypeAttr->typekind = TKIND_ALIAS THEN
                              dwsTypeKind = TLB_TypeKindToStr(pParamTypeAttr->typekind) & " | " & TLB_VarTypeToConstant(pParamTypeAttr->tdescalias.vt)
                           ELSE
                              dwsTypeKind = TLB_TypeKindToStr(pParamTypeAttr->typekind)
                           END IF
                           TreeView_AddItem(hTreeView, hParamNameNode, NULL, "TypeKind = " & dwsTypeKind)
                           pRefTypeInfo->ReleaseTypeAttr(pParamTypeAttr)
                           pParamTypeAttr = NULL
                        END IF
                        IF pRefTypeInfo THEN pRefTypeInfo->Release
                     END IF
                     EXIT DO
                  CASE ELSE
                     ' // Get the equivalent type
                     dwsVarType = TLB_VarTypeToConstant(ptdesc->vt)
                     dwsFBKeyword = TLB_VarTypeToKeyword(ptdesc->vt)
                     EXIT DO
               END SELECT
            LOOP
         ELSE
            ' // Get the equivalent type
            dwsVarType = TLB_VarTypeToConstant(pParam[y].tdesc.vt)
            dwsFBKeyword = TLB_VarTypeToKeyword(pParam[y].tdesc.vt)
            ' // Increment indirection level to pointers
            IF dwsTypeKind = "TKIND_INTERFACE" OR dwsTypeKind = "TKIND_DISPATCH" OR dwsTypeKind = "TKIND_COCLASS" THEN wIndirectionLevel += 1
            IF dwsVarType = "VT_SAFEARRAY" THEN wIndirectionLevel += 1
         END IF
         TreeView_AddItem(hTreeView, hParamNameNode, NULL, "Indirection level = " & WSTR(wIndirectionLevel))
         TreeView_AddItem(hTreeView, hParamNameNode, NULL, "VarType = " & dwsVarType)
         ' // Add a prefix to structures that begin with an underscore
'         IF LEFT$(dwsVarType, 1) = "_" THEN
'            IF dwsTypeKind = "TKIND_RECORD" OR dwsTypeKind = "TKIND_UNION" THEN dwsVarType = "tag" & dwsVarType
'         END IF
         ' // TODO: IF m_vTableView = TRUE then use BSTRING and DVARIANT
         SELECT CASE dwsTypeKind
            CASE "TKIND_INTERFACE", "TKIND_DISPATCH", "TKIND_COCLASS"
               IF wIndirectionLevel = 2 THEN
                  dwsFBSyntax = "BYVAL " & dwsParamName & " AS " & dwsVarType & " PTR PTR"
               ELSE
                  dwsFBSyntax = "BYVAL " & dwsParamName & " AS " & dwsVarType & " PTR"
               END IF
            CASE "TKIND_RECORD", "TKIND_UNION", "TKIND_ENUM"
               IF wIndirectionLevel = 2 THEN
                  dwsFBSyntax = "BYVAL " & dwsParamName & " AS " & dwsVarType & " PTR PTR"
               ELSEIF wIndirectionLevel = 1 THEN
                  dwsFBSyntax = "BYVAL " & dwsParamName &  " AS " & dwsVarType & " PTR"
               ELSE
                  dwsFBSyntax = "BYVAL " & dwsParamName & " AS " & dwsVarType
               END IF
            CASE ELSE
               IF LEFT(dwsTypeKind, 11) = "TKIND_ALIAS" THEN
                  IF wIndirectionLevel = 2 THEN
                     dwsFBSyntax = "BYVAL " & dwsParamName & " AS " & dwsVarType & " PTR PTR"
                  ELSEIF wIndirectionLevel = 1 THEN
                     dwsFBSyntax = "BYVAL " & dwsParamName &  " AS " & dwsVarType & " PTR"
                  ELSE
                     dwsFBSyntax = "BYVAL " & dwsParamName & " AS " & dwsVarType
                  END IF
               ELSE
                  SELECT CASE dwsVarType
                     CASE "VT_UNKNOWN"
                        IF wIndirectionLevel = 2 THEN
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS IUnknown PTR PTR"
                        ELSEIF wIndirectionLevel = 1 THEN
                           IF ((wFlags AND PARAMFLAG_FOUT) = PARAMFLAG_FOUT) OR ((wFlags AND PARAMFLAG_FRETVAL) = PARAMFLAG_FRETVAL) THEN
                              dwsFBSyntax = "BYVAL " & dwsParamName & " AS IUnknown PTR PTR"
                           ELSE
                              dwsFBSyntax = "BYVAL " & dwsParamName & " AS IUnknown PTR"
                           END IF
                        ELSE
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS IUnknown PTR"
                        END IF
                     CASE "VT_DISPATCH"
                        IF wIndirectionLevel = 2 THEN
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS IDispatch PTR PTR"
                        ELSEIF wIndirectionLevel = 1 THEN
                           IF ((wFlags AND PARAMFLAG_FOUT) = PARAMFLAG_FOUT) OR ((wFlags AND PARAMFLAG_FRETVAL) = PARAMFLAG_FRETVAL) THEN
                              dwsFBSyntax = "BYVAL " & dwsParamName & " AS IDispatch PTR PTR"
                           ELSE
                              dwsFBSyntax = "BYVAL " & dwsParamName & " AS IDispatch PTR"
                           END IF
                        ELSE
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS IDispatch PTR"
                        END IF
                     CASE "VT_VOID"
                        IF wIndirectionLevel = 2 THEN
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS ANY PTR PTR"
                        ELSEIF wIndirectionLevel = 1 THEN
                           IF ((wFlags AND PARAMFLAG_FOUT) = PARAMFLAG_FOUT) OR ((wFlags AND PARAMFLAG_FRETVAL) = PARAMFLAG_FRETVAL) THEN
                              dwsFBSyntax = "BYVAL " & dwsParamName & " AS ANY PTR PTR"
                           ELSE
                              dwsFBSyntax = "BYVAL " & dwsParamName & " AS ANY PTR"
                           END IF
                        ELSE
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS ANY PTR"
                        END IF
                     CASE "VT_LPSTR"
                        IF wIndirectionLevel = 1 OR wIndirectionLevel = 2 THEN
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS ZSTRING PTR PTR"
                        ELSE
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS ZSTRING PTR"
                        END IF
                     CASE "VT_LPWSTR"
                        IF wIndirectionLevel = 1 OR wIndirectionLevel = 2 THEN
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS WSTRING PTR PTR"
                        ELSE
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS WSTRING PTR"
                        END IF
                     CASE "VT_BSTR"
                        IF wIndirectionLevel = 2 THEN
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS BSTR PTR PTR"
                        ELSEIF wIndirectionLevel = 1 THEN
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS BSTR PTR"
                        ELSE
                           dwsFBSyntax = "BYVAL " & dwsParamName & " AS BSTR"
                        END IF
                     CASE ELSE
                        IF wIndirectionLevel = 2 THEN
                           dwsFBSyntax = "BYVAL " & dwsParamName &  " AS " & dwsFBKeyword & " PTR PTR"
                        ELSEIF wIndirectionLevel = 1 THEN
                           dwsFBSyntax = "BYVAL " & dwsParamName &  " AS " & dwsFBKeyword & " PTR"
                        ELSE
                           dwsFBSyntax = "BYVAL " & dwsParamName &  " AS " & dwsFBKeyword
                        END IF
                  END SELECT
               END IF
         END SELECT
         ' // See of it is an optional parameter without a default value
         IF (pParam[y].paramdesc.wParamFlags AND PARAMFLAG_FOPT) = PARAMFLAG_FOPT AND _
            (pParam[y].paramdesc.wParamFlags AND PARAMFLAG_FHASDEFAULT) <> PARAMFLAG_FHASDEFAULT THEN
            IF RIGHT(dwsFBSyntax, 4) = " PTR" THEN dwsFBSyntax += " = NULL"
            IF RIGHT(dwsFBSyntax, 11) = " AS VARIANT" THEN dwsFBSyntax += " = TYPE(VT_ERROR,0,0,0,DISP_E_PARAMNOTFOUND)"
         END IF
         ' // See if it has a default value
         IF (pParam[y].paramdesc.wParamFlags AND PARAMFLAG_FHASDEFAULT) = PARAMFLAG_FHASDEFAULT THEN
            DIM pex AS PARAMDESCEX PTR = pParam[y].paramdesc.pparamdescex
            DIM dwsDefaultValue AS DWSTRING = AfxVarToStr(@pex->vardefaultvalue)
            IF pex->vardefaultvalue.vt = VT_BSTR THEN
               TreeView_AddItem(hTreeView, hParamNameNode, NULL, "Default value = " & CHR(34) & dwsDefaultValue & CHR(34))
'               dwsFBSyntax += " = " & CHR(34) & dwsDefaultValue & CHR(34)
            ELSE
               TreeView_AddItem(hTreeView, hParamNameNode, NULL, "Default value = " & dwsDefaultValue)
               ' // Some typelibs have unprintable default values, e.g. wbemdisp.tlb,
               ' // that has unprintable IDispatch PTR values.
               IF LEN(dwsDefaultValue) THEN dwsFBSyntax += " = " & dwsDefaultValue
            END IF
         END IF
         TreeView_AddItem(hTreeView, hParamNameNode, NULL, "FB syntax = " & dwsFBSyntax)
         TreeView_Expand(hTreeView, hParamNameNode, TVE_EXPAND)
      NEXT
   END IF

   ' // Exand the parameters node
'   TreeView_Expand(hTreeView, hParamsNode, TVE_EXPAND)

   ' // DO NOT free the BSTRs; they are owned by the callee
   ' // Free the BSTRs of the array
'   FOR i AS LONG = LBOUND(rgbstrNames) TO UBOUND(rgbstrNames)
'      IF rgbstrNames(i) THEN SysFreeString(rgbstrNames(i))
'   NEXT

   ' // Just to satisfy the compiler rules; it has no useful meaning
   RETURN S_OK

END FUNCTION
' =====================================================================================
```

# Helper procedures

A number of helper procedures have been used to translate numeric values to more descriptive information:

```
' ========================================================================================
' Converts LibFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_LibFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND LIBFLAG_FRESTRICTED) = LIBFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND LIBFLAG_FCONTROL) = LIBFLAG_FCONTROL THEN strFlags += " [Control]"
   IF (wFlags AND LIBFLAG_FHIDDEN) = LIBFLAG_FHIDDEN THEN strFlags += " [Hidden]"
   IF (wFlags AND LIBFLAG_FHASDISKIMAGE) = LIBFLAG_FHASDISKIMAGE THEN strFlags += " [HasDiskImage]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Converts InterfaceFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_InterfaceFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND TYPEFLAG_FAPPOBJECT) = TYPEFLAG_FAPPOBJECT THEN strFlags += " [Application]"
   IF (wFlags AND TYPEFLAG_FCANCREATE) = TYPEFLAG_FCANCREATE THEN strFlags += " [Cancreate]"
   IF (wFlags AND TYPEFLAG_FLICENSED) = TYPEFLAG_FLICENSED THEN strFlags += " [Licensed]"
   IF (wFlags AND TYPEFLAG_FPREDECLID) = TYPEFLAG_FPREDECLID THEN strFlags += " [Predefined]"
   IF (wFlags AND TYPEFLAG_FHIDDEN) = TYPEFLAG_FHIDDEN THEN strFlags += " [Hidden]"
   IF (wFlags AND TYPEFLAG_FCONTROL) = TYPEFLAG_FCONTROL THEN strFlags += " [Control]"
   IF (wFlags AND TYPEFLAG_FDUAL) = TYPEFLAG_FDUAL THEN strFlags += " [Dual]"
   IF (wFlags AND TYPEFLAG_FNONEXTENSIBLE) = TYPEFLAG_FNONEXTENSIBLE THEN strFlags += " [Nonextensible]"
   IF (wFlags AND TYPEFLAG_FOLEAUTOMATION) = TYPEFLAG_FOLEAUTOMATION THEN strFlags += " [Oleautomation]"
   IF (wFlags AND TYPEFLAG_FRESTRICTED) = TYPEFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND TYPEFLAG_FAGGREGATABLE) = TYPEFLAG_FAGGREGATABLE THEN strFlags += " [Aggregatable]"
   IF (wFlags AND TYPEFLAG_FREPLACEABLE) = TYPEFLAG_FREPLACEABLE THEN strFlags += " [Replaceable]"
   IF (wFlags AND TYPEFLAG_FDISPATCHABLE) = TYPEFLAG_FDISPATCHABLE THEN strFlags += " [Dispatchable]"
   IF (wFlags AND TYPEFLAG_FREVERSEBIND) = TYPEFLAG_FREVERSEBIND THEN strFlags += " [Reversebind]"
   IF (wFlags AND TYPEFLAG_FPROXY) = TYPEFLAG_FPROXY THEN strFlags += " [Proxy]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Converts ImplTypeFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_ImplTypeFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND IMPLTYPEFLAG_FDEFAULT) = IMPLTYPEFLAG_FDEFAULT THEN strFlags += " [Default]"
   IF (wFlags AND IMPLTYPEFLAG_FSOURCE) = IMPLTYPEFLAG_FSOURCE THEN strFlags += " [Source]"
   IF (wFlags AND IMPLTYPEFLAG_FRESTRICTED) = IMPLTYPEFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND IMPLTYPEFLAG_FDEFAULTVTABLE) = IMPLTYPEFLAG_FDEFAULTVTABLE THEN strFlags += " [Default VTable]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Converts FuncFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_FuncFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND FUNCFLAG_FRESTRICTED) = FUNCFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND FUNCFLAG_FSOURCE) = FUNCFLAG_FSOURCE THEN strFlags += " [Source]"
   IF (wFlags AND FUNCFLAG_FBINDABLE) = FUNCFLAG_FBINDABLE THEN strFlags += " [Bindable]"
   IF (wFlags AND FUNCFLAG_FREQUESTEDIT) = FUNCFLAG_FREQUESTEDIT THEN strFlags += " [RequestEdit]"
   IF (wFlags AND FUNCFLAG_FDISPLAYBIND) = FUNCFLAG_FDISPLAYBIND THEN strFlags += " [DisplayBind]"
   IF (wFlags AND FUNCFLAG_FDEFAULTBIND) = FUNCFLAG_FDEFAULTBIND THEN strFlags += " [DefaultBind]"
   IF (wFlags AND FUNCFLAG_FHIDDEN) = FUNCFLAG_FHIDDEN THEN strFlags += " [Hidden]"
   IF (wFlags AND FUNCFLAG_FUSESGETLASTERROR) = FUNCFLAG_FUSESGETLASTERROR THEN strFlags += " [UsesGetLastError]"
   IF (wFlags AND FUNCFLAG_FDEFAULTCOLLELEM) = FUNCFLAG_FDEFAULTCOLLELEM THEN strFlags += " [DefaultCollELem]"
   IF (wFlags AND FUNCFLAG_FUIDEFAULT) = FUNCFLAG_FUIDEFAULT THEN strFlags += " [UserInterfaceDefault]"
   IF (wFlags AND FUNCFLAG_FNONBROWSABLE) = FUNCFLAG_FNONBROWSABLE THEN strFlags += " [NonBrowsable]"
   IF (wFlags AND FUNCFLAG_FREPLACEABLE) = FUNCFLAG_FREPLACEABLE THEN strFlags += " [Replaceable]"
   IF (wFlags AND FUNCFLAG_FIMMEDIATEBIND) = FUNCFLAG_FIMMEDIATEBIND THEN strFlags += " [InmediateBind]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Converts ParamFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_ParamflagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN FUNCTION = " [None]" : EXIT FUNCTION
   IF (wFlags AND PARAMFLAG_FOPT) = PARAMFLAG_FOPT THEN strFlags += " [opt]"
   IF (wFlags AND PARAMFLAG_FRETVAL) = PARAMFLAG_FRETVAL THEN strFlags += " [retval]"
   IF (wFlags AND PARAMFLAG_FIN) = PARAMFLAG_FIN THEN strFlags += " [in]"
   IF (wFlags AND PARAMFLAG_FOUT) = PARAMFLAG_FOUT THEN strFlags += " [out]"
   IF (wFlags AND PARAMFLAG_FLCID) = PARAMFLAG_FLCID THEN strFlags += " [lcid]"
   IF (wFlags AND PARAMFLAG_FHASDEFAULT) = PARAMFLAG_FHASDEFAULT THEN strFlags += " [hasdefault]"
   IF (wFlags AND PARAMFLAG_FHASCUSTDATA) = PARAMFLAG_FHASCUSTDATA THEN strFlags += " [hascustdata]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Converts VarFlags to a descriptive string.
' ========================================================================================
FUNCTION TLB_VarFlagsToStr (BYVAL wFlags AS WORD) AS STRING
   DIM strFlags AS STRING
   IF wFlags = 0 THEN strFlags = " [None]"
   IF (wFlags AND VARFLAG_FREADONLY) = VARFLAG_FREADONLY THEN strFlags += " [ReadOnly]"
   IF (wFlags AND VARFLAG_FSOURCE) = VARFLAG_FSOURCE THEN strFlags += " [Source]"
   IF (wFlags AND VARFLAG_FBINDABLE) = VARFLAG_FBINDABLE THEN strFlags += " [Bindable]"
   IF (wFlags AND VARFLAG_FREQUESTEDIT) = VARFLAG_FREQUESTEDIT THEN strFlags += " [RequestEdit]"
   IF (wFlags AND VARFLAG_FDISPLAYBIND) = VARFLAG_FDISPLAYBIND THEN strFlags += " [DisplayBind]"
   IF (wFlags AND VARFLAG_FDEFAULTBIND) = VARFLAG_FDEFAULTBIND THEN strFlags += " [DefaultBind]"
   IF (wFlags AND VARFLAG_FHIDDEN) = VARFLAG_FHIDDEN THEN strFlags += " [Hidden]"
   IF (wFlags AND VARFLAG_FRESTRICTED) = VARFLAG_FRESTRICTED THEN strFlags += " [Restricted]"
   IF (wFlags AND VARFLAG_FDEFAULTCOLLELEM) = VARFLAG_FDEFAULTCOLLELEM THEN strFlags += " [DefaultCollElem]"
   IF (wFlags AND VARFLAG_FUIDEFAULT) = VARFLAG_FUIDEFAULT THEN strFlags += " [User interface default]"
   IF (wFlags AND VARFLAG_FNONBROWSABLE) = VARFLAG_FNONBROWSABLE THEN strFlags += " [NoBrowsable]"
   IF (wFlags AND VARFLAG_FREPLACEABLE) = VARFLAG_FREPLACEABLE THEN strFlags += " [Replaceable]"
   IF (wFlags AND VARFLAG_FIMMEDIATEBIND) = VARFLAG_FIMMEDIATEBIND THEN strFlags += " [ImmediateBind]"
   FUNCTION = strFlags
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Converts a type kind to a descriptive string.
' ========================================================================================
FUNCTION TLB_TypeKindToStr (BYVAL dwTypeKind AS DWORD) AS STRING
   DIM strType AS STRING
   SELECT CASE dwTypeKind
      CASE TKIND_ENUM      : strType = "TKIND_ENUM"
      CASE TKIND_RECORD    : strType = "TKIND_RECORD"
      CASE TKIND_MODULE    : strType = "TKIND_MODULE"
      CASE TKIND_INTERFACE : strType = "TKIND_INTERFACE"
      CASE TKIND_DISPATCH  : strType = "TKIND_DISPATCH"
      CASE TKIND_COCLASS   : strType = "TKIND_COCLASS"
      CASE TKIND_ALIAS     : strType = "TKIND_ALIAS"
      CASE TKIND_UNION     : strType = "TKIND_UNION"
   END SELECT
   FUNCTION = strType
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Returns the VarType.
' ========================================================================================
FUNCTION TLB_VarTypeToStr (BYVAL VarType AS LONG, BYVAL fReturnType AS LONG = 0) AS STRING
   DIM s AS STRING
   SELECT CASE VarType
      CASE     0 : s = "VT_EMPTY"
      CASE     1 : s = "VT_NULL"
      CASE     2 : s = "VT_I2 <Short>"
      CASE     3 : s = "VT_I4 <Long>"
      CASE     4 : s = "VT_R4 <Single>"
      CASE     5 : s = "VT_R8 <Double>"
      CASE     6 : s = "VT_CY <CY>"
      CASE     7 : s = "VT_DATE <DATE_>"
      CASE     8 : s = "VT_BSTR <BSTR>"
      CASE     9 : s = "VT_DISPATCH <IDispatch>"
      CASE    10 : s = "VT_ERROR <SCode>"
      CASE    11 : s = "VT_BOOL <VARIANT_BOOL>"
      CASE    12 : s = "VT_VARIANT <Variant>"
      CASE    13 : s = "VT_UNKNOWN <IUnknown>"
      CASE    14 : s = "VT_DECIMAL <DECIMAL>"
      CASE    16 : s = "VT_I1 <Byte>"
      CASE    17 : s = "VT_UI1 <UByte>"
      CASE    18 : s = "VT_UI2 <Short>"
      CASE    19 : s = "VT_UI4 <Ulong>"
      CASE    20 : s = "VT_I8 <LongInt>"
      CASE    21 : s = "VT_UI8 <ULongInt>"
      CASE    22 : s = "VT_INT <Int_>"
      CASE    23 : s = "VT_UINT <Uint>"
      CASE    24 :
         IF fReturnType THEN
            s = "VT_VOID <void>"
         ELSE
            s = "VT_VOID <void>"
         END IF
      CASE    25 : s = "VT_HRESULT <HRESULT>"
      CASE    26 : s = "VT_PTR <PTR>"
      CASE    27 : s = "VT_SAFEARRAY <SAFEARRAY>"
      CASE    28 : s = "VT_CARRAY"
      CASE    29 : s = "VT_USERDEFINED"
      CASE    30 : s = "VT_LPSTR"
      CASE    31 : s = "VT_LPWSTR"
      CASE    36 : s = "VT_RECORD"
      CASE    64 : s = "VT_FILETIME <FILETIME>"
      CASE    65 : s = "VT_BLOB <BLOB>"
      CASE    66 : s = "VT_STREAM <IStream PTR>"
      CASE    67 : s = "VT_STORAGE <IStorage PTR>"
      CASE    68 : s = "VT_STREAMED_OBJECT"
      CASE    69 : s = "VT_STORED_OBJECT"
      CASE    70 : s = "VT_BLOB_OBJECT"
      CASE    71 : s = "VT_CF"
      CASE    72 : s = "VT_CLSID <Guid>"
      CASE  4096 : s = "VT_VECTOR"
      CASE  8192 : s = "VT_ARRAY"
      CASE 16384 : s = "VT_BYREF"
      CASE 32768 : s = "VT_RESERVED"
   END SELECT
   FUNCTION = s
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Returns the VarType
' ========================================================================================
FUNCTION TLB_VarTypeToConstant (BYVAL VarType AS LONG) AS STRING
   DIM s AS STRING
   SELECT CASE VarType
      CASE     0 : s = "VT_EMPTY"
      CASE     1 : s = "VT_NULL"
      CASE     2 : s = "VT_I2"
      CASE     3 : s = "VT_I4"
      CASE     4 : s = "VT_R4"
      CASE     5 : s = "VT_R8"
      CASE     6 : s = "VT_CY"
      CASE     7 : s = "VT_DATE"
      CASE     8 : s = "VT_BSTR"
      CASE     9 : s = "VT_DISPATCH"
      CASE    10 : s = "VT_ERROR"
      CASE    11 : s = "VT_BOOL"
      CASE    12 : s = "VT_VARIANT"
      CASE    13 : s = "VT_UNKNOWN"
      CASE    14 : s = "VT_DECIMAL"
      CASE    16 : s = "VT_I1"
      CASE    17 : s = "VT_UI1"
      CASE    18 : s = "VT_UI2"
      CASE    19 : s = "VT_UI4"
      CASE    20 : s = "VT_I8"
      CASE    21 : s = "VT_UI8"
      CASE    22 : s = "VT_INT"
      CASE    23 : s = "VT_UINT"
      CASE    24 : s = "VT_VOID"
      CASE    25 : s = "VT_HRESULT"
      CASE    26 : s = "VT_PTR"
      CASE    27 : s = "VT_SAFEARRAY"
      CASE    28 : s = "VT_CARRAY"
      CASE    29 : s = "VT_USERDEFINED"
      CASE    30 : s = "VT_LPSTR"
      CASE    31 : s = "VT_LPWSTR"
      CASE    36 : s = "VT_RECORD"
      CASE    64 : s = "VT_FILETIME"
      CASE    65 : s = "VT_BLOB"
      CASE    66 : s = "VT_STREAM"
      CASE    67 : s = "VT_STORAGE"
      CASE    68 : s = "VT_STREAMED_OBJECT"
      CASE    69 : s = "VT_STORED_OBJECT"
      CASE    70 : s = "VT_BLOB_OBJECT"
      CASE    71 : s = "VT_CF"
      CASE    72 : s = "VT_CLSID"
      CASE  4096 : s = "VT_VECTOR"
      CASE  8192 : s = "VT_ARRAY"
      CASE 16384 : s = "VT_BYREF"
      CASE 32768 : s = "VT_RESERVED"
   END SELECT
   FUNCTION = s
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Returns the VarType as a keyword
' ========================================================================================
FUNCTION TLB_VarTypeToKeyword OVERLOAD (BYVAL VarType AS LONG, BYVAL cElements AS WORD = 0) AS STRING
   ' Note: VT_I1 is an array of bytes; translate it to a fixed string
   DIM s AS STRING
   SELECT CASE VarType
      CASE     0 : s = "VOID"                              ' VT_EMPTY
      CASE     1 : s = "VOID"                              ' VT_NULL
      CASE     2 : s = "SHORT"                             ' VT_I2
      CASE     3 : s = "LONG"                              ' VT_I4
      CASE     4 : s = "SINGLE"                            ' VT_R4
      CASE     5 : s = "DOUBLE"                            ' VT_R8
      CASE     6 : s = "CY"                                ' VT_CY
      CASE     7 : s = "DATE_"                             ' VT_DATE
      CASE     8 : s = "BSTR"                              ' VT_BSTR
      CASE     9 : s = "IDispatch"                         ' VT_DISPATCH
      CASE    10 : s = "SCODE"                             ' VT_ERROR
      CASE    11 : s = "VARIANT_BOOL"                      ' VT_BOOL
      CASE    12 : s = "VARIANT"                           ' VT_VARIANT
      CASE    13 : s = "IUnknown"                          ' VT_UNKNOWN
      CASE    14 : s = "DECIMAL"                           ' VT_DECIMAL
      CASE 16, 17                                          ' VT_I1, VT_UI1
'         IF cElements THEN
'            s = "WSTRING * " & STR(cElements)              ' Byte array
'         ELSE
'            s = "BYTE"
'         END IF
         IF cElements THEN
            s = "(0 TO " & STR(cElements) & " AS " & IIF(VarType = 16, "BYTE", "UBYTE") & ")"
         ELSE
            s = IIF(VarType = 16, "BYTE", "UBYTE")
         END IF
      CASE    18 : s = "USHORT"                            ' VT_UI2
      CASE    19 : s = "ULONG"                             ' VT_UI4
      CASE    20 : s = "LONGINT"                           ' VT_I8
      CASE    21 : s = "ULONGINT"                          ' VT_UI8
      CASE    22 : s = "INT_"                              ' VT_INT
      CASE    23 : s = "UINT"                              ' VT_UINT
      CASE    24 : s = "VOID"                              ' VT_VOID
      CASE    25 : s = "HRESULT"                           ' VT_HRESULT
      CASE    26 : s = "PTR"                               ' VT_PTR
      CASE    27 : s = "SAFEARRAY"                         ' VT_SAFEARRAY
      CASE    28 : s = "VOID"                              ' VT_CARRAY
      CASE    29 : s = "VOID"                              ' VT_USERDEFINED
      CASE    30 : s = "ZTRING"                            ' VT_LPSTR
      CASE    31 : s = "WSTRING"                           ' VT_LPWSTR
      CASE    36 : s = "VOID"                              ' VT_RECORD
      CASE    64 : s = "FILETIME"                          ' VT_FILETIME
      CASE    65 : s = "BLOB"                              ' VT_BLOB
      CASE    66 : s = "IStream"                           ' VT_STREAM
      CASE    67 : s = "IStorage"                          ' VT_STORAGE
      CASE    68 : s = "VOID"                              ' VT_STREAMED_OBJECT
      CASE    69 : s = "VOID"                              ' VT_STORED_OBJECT
      CASE    70 : s = "VOID"                              ' VT_BLOB_OBJECT
      CASE    71 : s = "VOID"                              ' VT_CF
      CASE    72 : s = "CLSID"                             ' VT_CLSID
      CASE  4096 : s = "VOID"                              ' VT_VECTOR
      CASE  8192 : s = "VOID"                              ' VT_ARRAY
      CASE 16384 : s = "VT_BYREF"
      CASE 32768 : s = "VT_RESERVED"
      CASE ELSE
         s = "VOID"
   END SELECT
   FUNCTION = s
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Returns the VarType
' ========================================================================================
FUNCTION TLB_VarTypeToKeyword OVERLOAD (BYVAL VarType AS STRING) AS STRING
   DIM s AS STRING
   SELECT CASE VarType
      CASE "VT_EMPTY"           : s = "VOID"
      CASE "VT_NULL"            : s = "VOID"
      CASE "VT_I1"              : s = "BYTE"
      CASE "VT_UI1"             : s = "UBYTE"
      CASE "VT_I2"              : s = "SHORT"
      CASE "VT_UI2"             : s = "USHORT"
      CASE "VT_I4"              : s = "LONG"
      CASE "VT_UI4"             : s = "ULONG"
      CASE "VT_I8"              : s = "LONGINT"
      CASE "VT_UI8"             : s = "ULONGINT"
      CASE "VT_INT"             : s = "INT_"
      CASE "VT_UINT"            : s = "UINT"
      CASE "VT_R4"              : s = "SINGLE"
      CASE "VT_R8"              : s = "DOUBLE"
      CASE "VT_CY"              : s = "CY"
      CASE "VT_DATE"            : s = "DATE_"
      CASE "VT_BSTR"            : s = "BSTR"
      CASE "VT_UNKNOWN"         : s = "IUnknown"
      CASE "VT_DISPATCH"        : s = "IDispatch"
      CASE "VT_ERROR"           : s = "VOID"
      CASE "VT_BOOL"            : s = "BOOL"
      CASE "VT_VARIANT"         : s = "VARIANT"
      CASE "VT_DECIMAL"         : s = "DECIMAL"
      CASE "VT_VOID"            : s = "VOID"
      CASE "VT_HRESULT"         : s = "HRESULT"
      CASE "VT_PTR"             : s = "PTR"
      CASE "VT_SAFEARRAY"       : s = "SAFEARRAY"
      CASE "VT_CARRAY"          : s = "VOID"
      CASE "VT_USERDEFINED"     : s = "VOID"
      CASE "VT_LPSTR"           : s = "ZSTRING PTR"
      CASE "VT_LPWSTR"          : s = "WSTRING PTR"
      CASE "VT_RECORD"          : s = "VOID"
      CASE "VT_FILETIME"        : s = "FILETIME"
      CASE "VT_BLOB"            : s = "VOID"
      CASE "VT_STREAM"          : s = "IStream"
      CASE "VT_STORAGE"         : s = "IStorage"
      CASE "VT_STREAMED_OBJECT" : s = "VOID"
      CASE "VT_STORED_OBJECT"   : s = "VOID"
      CASE "VT_BLOB_OBJECT"     : s = "VOID"
      CASE "VT_CF"              : s = "VOID"
      CASE "VT_CLSID"           : s = "CLSID"
      CASE "VT_VECTOR"          : s = "VOID"
      CASE "VT_ARRAY"           : s = "VOID"
      CASE "VT_BYREF"           : s = "VOID"
      CASE "VT_RESERVED"        : s = "VOID"
      CASE ELSE
         s = VarType
   END SELECT
   FUNCTION = s
END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Gets the appropiate member name of the variant union for byref parameters.
' Note: VT_HRESULT isn't an automation compatible type, but the CreatePartnershipComplete
' event of Windows Media Player has a parameter of this type.
' ========================================================================================
FUNCTION TLB_GetUnionMemberName (BYVAL vt AS LONG) AS STRING
   DIM strvt AS STRING
   SELECT CASE vt
      CASE VT_I1, VT_UI1 : strvt = "pbVal"
      CASE VT_I2 : strvt = "piVal"
      CASE VT_I4, VT_INT, VT_UI4, VT_UINT, VT_HRESULT : strvt = "plVal"
      CASE VT_R4 : strvt = "pfltVal"
      CASE VT_R8, VT_I8, VT_UI8 : strvt = "pdblVal"
      CASE VT_BOOL : strvt = "pboolVal"
      CASE VT_ERROR : strvt = "pscode"
      CASE VT_CY : strvt = "pcyVal"
      CASE VT_DATE : strvt = "pdate"
      CASE VT_BSTR : strvt = "pbstrVal"
      CASE VT_UNKNOWN : strvt = "ppunkVal"
      CASE VT_DISPATCH : strvt = "ppdispVal"
      CASE VT_ARRAY : strvt = "psArray"
      CASE VT_VARIANT : strvt = "pVariant"
      CASE ELSE : strvt = "plVal"
   END SELECT
   FUNCTION = strvt
END FUNCTION
' ========================================================================================
```

And others to get information from the registry and to get the names of implemented and inherited interfaces, and the base class.

```
' ========================================================================================
' Gets the ProgID from the registry.
' ========================================================================================
FUNCTION TLB_GetProgID (BYVAL pwszGuid AS WSTRING PTR) AS DWSTRING

   ' // Name of the subkey to open
   DIM wszKey AS WSTRING * MAX_PATH = "CLSID\" & *pwszGuid & "\ProgID"
   DIM hKey AS HKEY   ' // Handle of the opened key
   RegOpenKeyExW HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey
   IF hKey = NULL THEN RETURN ""
   DIM dwIdx AS DWORD = 0                   ' // Index of the value to be retrieved
   DIM cValueName AS DWORD = MAX_PATH       ' // Size of wszValueName
   DIM cbData AS DWORD = MAX_PATH           ' // Size of wszKeyValue
   DIM keyType AS DWORD                     ' // Type of data
   DIM wszKeyValue AS WSTRING * MAX_PATH    ' // Buffer that receives the data
   DIM wszValueName AS WSTRING * MAX_PATH   ' // Name of the value
   RegEnumValueW hKey, dwIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData
   RegCloseKey hKey
   RETURN wszKeyValue

END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Gets the Version Independent ProgID from the registry.
' ========================================================================================
FUNCTION TLB_GetVersionIndependentProgID (BYVAL pwszGuid AS WSTRING PTR) AS DWSTRING

   ' // Name of the subkey to open
   DIM wszKey AS WSTRING * MAX_PATH = "CLSID\" & *pwszGuid & "\VersionIndependentProgID"
   DIM hKey AS HKEY   ' // Handle of the opened key
   RegOpenKeyExW HKEY_CLASSES_ROOT, @wszKey, 0, KEY_READ, @hKey
   IF hKey = NULL THEN RETURN ""
   DIM dwIdx AS DWORD = 0                   ' // Index of the value to be retrieved
   DIM cValueName AS DWORD = MAX_PATH       ' // Size of wszValueName
   DIM cbData AS DWORD = MAX_PATH           ' // Size of wszKeyValue
   DIM keyType AS DWORD                     ' // Type of data
   DIM wszKeyValue AS WSTRING * MAX_PATH    ' // Buffer that receives the data
   DIM wszValueName AS WSTRING * MAX_PATH   ' // Name of the value
   RegEnumValueW hKey, dwIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData
   RegCloseKey hKey
   RETURN wszKeyValue

END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Gets the InprocServer32 from the registry.
' ========================================================================================
FUNCTION TLB_GetInprocServer32 (BYVAL pwszGuid AS WSTRING PTR) AS DWSTRING

   ' // Name of the subkey to open
   DIM wszKey AS WSTRING * MAX_PATH = "CLSID\" & *pwszGuid & "\InprocServer32"
   DIM hKey AS HKEY   ' // Handle of the opened key
   RegOpenKeyExW HKEY_CLASSES_ROOT, @wszKey , 0, KEY_READ, @hKey
   IF hKey = NULL THEN RETURN ""
   DIM dwIdx AS DWORD = 0                   ' // Index of the value to be retrieved
   DIM cValueName AS DWORD = MAX_PATH       ' // Size of wszValueName
   DIM cbData AS DWORD = MAX_PATH           ' // Size of wszKeyValue
   DIM keyType AS DWORD                     ' // Type of data
   DIM wszKeyValue AS WSTRING * MAX_PATH    ' // Buffer that receives the data
   DIM wszValueName AS WSTRING * MAX_PATH   ' // Name of the value
   RegEnumValueW hKey, dwIdx, @wszValueName, @cValueName, NULL, @keyType, cast(BYTE PTR, @wszKeyValue), @cbData
   RegCloseKey hKey
   RETURN wszKeyValue

END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Retrieves the implemented interface.
' ========================================================================================
FUNCTION TLB_GetImplementedInterface (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL idx AS LONG = 0) AS DWSTRING

   DIM pRefType AS HREFTYPE   ' // Address to a referenced type description
   DIM hr AS HRESULT = pTypeInfo->GetRefTypeOfImplType(idx, @pRefType)
   IF hr <> S_OK OR pRefType = NULL THEN RETURN ""
   DIM pImplTypeInfo AS Afx_ITypeInfo PTR   ' // Implemented interface type info
   hr = pTypeInfo->GetRefTypeInfo(pRefType, @pImplTypeInfo)
   IF hr <> S_OK OR pImplTypeInfo = NULL THEN RETURN ""
   DIM dwsName AS DWSTRING, bstrName AS AFX_BSTR   ' // interface name
   pImplTypeInfo->GetDocumentation(-1, @bstrName, NULL, NULL, NULL)
   pImplTypeInfo->Release
   dwsName = *bstrName
   SysFreeString bstrName
   RETURN dwsName

END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Retrieves the inherited interface
' ========================================================================================
FUNCTION TLB_GetInheritedInterface (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL idx AS LONG = 0) AS DWSTRING

   DIM pRefType AS HREFTYPE   ' // Address to a referenced type description
   DIM hr AS HRESULT = pTypeInfo->GetRefTypeOfImplType(idx, @pRefType)
   IF hr <> S_OK OR pRefType = NULL THEN RETURN ""
   DIM pImplTypeInfo AS Afx_ITypeInfo PTR   ' // Implied interface type info
   hr = pTypeInfo->GetRefTypeInfo (pRefType, @pImplTypeInfo)
   IF hr <> S_OK OR pImplTypeInfo = NULL THEN RETURN ""
   DIM pTypeAttr AS TYPEATTR PTR   ' // Address of a pointer to the TYPEATTR structure
   hr = pImplTypeInfo->GetTypeAttr(@pTypeAttr)
   DIM dwsInterfaceName AS DWSTRING
   IF hr = S_OK AND pTypeAttr <> NULL THEN
      IF @pTypeAttr->cImplTypes = 1 THEN
         dwsInterfaceName = TLB_GetImplementedInterface(pImplTypeInfo, 0)
         pImplTypeInfo->ReleaseTypeAttr(pTypeAttr)
      END IF
   END IF
   pImplTypeInfo->Release
   RETURN dwsInterfaceName

END FUNCTION
' ========================================================================================
```
```
' ========================================================================================
' Retrieves the base class
' ========================================================================================
FUNCTION TLB_GetBaseClass (BYVAL pTypeLib AS Afx_ITypeLib PTR, BYREF dwsItemName AS DWSTRING) AS DWSTRING

   DIM pTypeInfo               AS Afx_ITypeInfo PTR   ' // TypeInfo interface
   DIM pTypeAttr               AS TYPEATTR PTR        ' // Address of a pointer to the TYPEATTR structure
   DIM pRefType                AS DWORD               ' // Address to a referenced type description
   DIM pRefTypeInfo            AS Afx_ITypeInfo PTR   ' // Referenced TypeInfo interface
   DIM pRefTypeAttr            AS TYPEATTR PTR        ' // Referenced TYPEATTR structure
   DIM dwsInheritedInterface   AS DWSTRING            ' // Inherited interface

   ' // Number of type infos
   DIM TypeInfoCount AS LONG = pTypeLib->GetTypeInfoCount
   IF TypeInfoCount = 0 THEN RETURN ""

   FOR i AS LONG = 0 TO TypeInfoCount - 1
      ' // Get the info type
      DIM pTKind AS TYPEKIND
      DIM hr AS HRESULT = pTypeLib->GetTypeInfoType(i, @pTKind)
      IF hr <> S_OK THEN EXIT FOR
      ' // Get the type info
      hr = pTypeLib->GetTypeInfo(i, @pTypeInfo)
      IF hr <> S_OK THEN EXIT FOR
      ' // Get the type attribute
      hr = pTypeInfo->GetTypeAttr(@pTypeAttr)
      IF hr <> S_OK OR pTypeAttr = NULL THEN EXIT FOR
      ' // If it is an interface...
      IF pTKind = TKIND_INTERFACE OR pTKind = TKIND_DISPATCH THEN
         ' // Get the name of the interface
         DIM dwsName AS DWSTRING, bstrName AS AFX_BSTR
         hr = pTypeLib->GetDocumentation(i, @bstrName, NULL, NULL, NULL)
         dwsName = *bstrName
         SysFreeString bstrName
         ' // If it is the one we are looking for...
         IF dwsName = dwsItemName THEN
            ' // If it inherits from another interface, recursively search the methods
            IF (pTypeAttr->wTypeFlags AND TYPEFLAG_FDUAL) = TYPEFLAG_FDUAL THEN
               dwsInheritedInterface = TLB_GetInheritedInterface(pTypeInfo, -1)
            ELSE
               dwsInheritedInterface = TLB_GetImplementedInterface(pTypeInfo)
            END IF
            ' // Check also that the interface doesn't inherit from itself!
            IF UCASE(dwsInheritedInterface) <> "IUNKNOWN" AND UCASE(dwsInheritedInterface) <> "IDISPATCH" AND UCASE(dwsInheritedInterface) <> UCASE(*bstrName) THEN
               dwsInheritedInterface = TLB_GetBaseClass(pTypeLib, dwsInheritedInterface)
            END IF
         END IF
      END IF
      pTypeInfo->ReleaseTypeAttr(pTypeAttr) : pTypeAttr = NULL
      pTypeInfo->Release : pTypeInfo = NULL
   NEXT

   IF pTypeAttr THEN pTypeInfo->ReleaseTypeAttr(pTypeAttr)
   IF pTypeInfo THEN pTypeInfo->Release

   RETURN dwsInheritedInterface

END FUNCTION
' ========================================================================================
```
