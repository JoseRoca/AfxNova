' ########################################################################################
' TypeLib Browser
' File: TLB_ParseLib.inc
' Contents: TypeLib Browser - Parsing class
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2016 José Roca. All Rights Reserved.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' ########################################################################################
' Load and parse the type library
' ########################################################################################
TYPE CParseTypeLib

   m_pTypeLib           AS Afx_ITypeLib PTR     ' // Reference to the ITypeLib interface
   m_LibName            AS WSTRING * MAX_PATH   ' // Library name
   m_LibPath            AS WSTRING * MAX_PATH   ' // Library path
   m_LibHelpString      AS WSTRING * 260        ' // Library help string
   m_LibHelpContext     AS DWORD                ' // Library Help context
   m_LibHelpFile        AS WSTRING * MAX_PATH   ' // Library help file
   m_LibGuid            AS WSTRING * 40         ' // Library GUID
   m_LibLcid            AS DWORD                ' // Library LCID
   m_LibSysKind         AS LONG                 ' // Library SysKind
   m_LibMajorVersion    AS SHORT                ' // Library major version
   m_LibMinorVersion    AS SHORT                ' // Library minor version
   m_LibAttr            AS SHORT                ' // Library attributes
   m_Namespace          AS WSTRING * 20         ' // Namespace

   m_hRootNode          AS HTREEITEM            ' // Root node
   m_hDocNode           AS HTREEITEM            ' // Documentation node
   m_hProgIDsNode       AS HTREEITEM            ' // ProgIDs node
   m_hVerIndProgIDsNode AS HTREEITEM            ' // Version independent ProgIDs node
   m_hClsIDsNode        AS HTREEITEM            ' // ClsIDs node
   m_hIIDsNode          AS HTREEITEM            ' // IIDs node
   m_hCoClassesNode     AS HTREEITEM            ' // CoClasses node
   m_hTypeDefsNode      AS HTREEITEM            ' // TypeDefs node
   m_hAliasesNode       AS HTREEITEM            ' // Aliases node
   m_hEnumsNode         AS HTREEITEM            ' // Enums node
   m_hModulesNode       AS HTREEITEM            ' // Modules node
   m_hRecordsNode       AS HTREEITEM            ' // Records node
   m_hUnionsNode        AS HTREEITEM            ' // Unions node
   m_hIntNode           AS HTREEITEM            ' // Interfaces node
   m_hDualIntNode       AS HTREEITEM            ' // Dual interfaces node
   m_hDispIntNode       AS HTREEITEM            ' // Dispatch interfaces node
   m_hDispblIntNode     AS HTREEITEM            ' // Dispatchable interfaces node
   m_hOleAutIntNode     AS HTREEITEM            ' // Ole automation interfaces node
   m_hEventsNode        AS HTREEITEM            ' // Events node

   DIM m_pWindow        AS CWindow PTR          ' // Reference to the CWindow interface
   DIM m_VtableView     AS BOOLEAN              ' // VtableView = TRUE or FALSE
   DIM dwsCode          AS DWSTRING             ' // Generated code
   DIM dwsEvents        AS DWSTRING             ' // Event interfaces names

   DECLARE CONSTRUCTOR (BYVAL pWindow AS CWindow PTR)
   DECLARE DESTRUCTOR
   DECLARE FUNCTION LoadTypeLibEx (BYVAL pwszPath AS WSTRING PTR) AS HRESULT
   DECLARE FUNCTION ParseTypeInfos () AS HRESULT
   DECLARE FUNCTION GetMembers (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL pTypeAttr AS TYPEATTR PTR, BYVAL hTreeView AS HWND, BYVAL hSubNode AS HTREEITEM, BYVAL bIsRecord AS BOOLEAN = FALSE) AS HRESULT
   DECLARE FUNCTION GetParameters (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL pFuncDesc AS FUNCDESC PTR, BYVAL hTreeView AS HWND, BYVAL hSubNode2 AS HTREEITEM, BYVAL bVTableView AS BOOLEAN) AS HRESULT
   DECLARE FUNCTION GetFunctions (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL pTypeAttr AS TYPEATTR PTR, BYVAL hTreeView AS HWND, BYVAL hSubNode AS HTREEITEM, BYVAL bVTableView AS BOOLEAN, BYVAL bIsMethod AS BOOLEAN = FALSE, BYVAL pTKind AS TYPEKIND = -1, BYVAL pwszImplInterface AS WSTRING PTR = NULL) AS HRESULT
   DECLARE FUNCTION GetConstants (BYVAL pTypeInfo AS Afx_ITypeInfo PTR, BYVAL pTypeAttr AS TYPEATTR PTR, BYVAL hTreeView AS HWND, BYVAL hSubNode AS HTREEITEM) AS HRESULT
   DECLARE SUB      GenerateCode

END TYPE
' #######################################################################################

' =====================================================================================
' Constructor
' =====================================================================================
CONSTRUCTOR CParseTypeLib (BYVAL pWindow AS CWindow PTR)
   m_pWindow = pWindow
END CONSTRUCTOR
' =====================================================================================

' =====================================================================================
' Destructor
' =====================================================================================
DESTRUCTOR CParseTypeLib
   ' // Release the ITypeLib interface
   IF m_pTypeLib THEN m_pTypeLib->Release
END DESTRUCTOR
' =====================================================================================

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
      dwsTypeKind = ""
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
         IF dwsTypeKind = "TKIND_UNION" THEN dwsFBKeyword = dwsVarType
         IF dwsTypeKind = "TKIND_RECORD" THEN dwsFBKeyword = dwsVarType
         IF dwsTypeKind = "TKIND_RECORD" AND wIndirectionLevel = 1 THEN dwsFBKeyword += " PTR"
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

' =====================================================================================
' Parses the type infos
' =====================================================================================
PRIVATE FUNCTION CParseTypeLib.ParseTypeInfos () AS HRESULT

   DIM hTreeView AS HWND                      ' // Tree view handle
   DIM pTKind AS TYPEKIND                     ' // Type kind
   DIM pTypeAttr AS TYPEATTR PTR              ' // Attributes
   DIM pTypeInfo AS Afx_ITypeInfo PTR         ' // ITypeInfo interface
   ' // Nodes
   DIM hNode AS HTREEITEM                     ' // Node handle
   DIM hSubNode AS HTREEITEM                  ' // Sub node handle
   DIM hSubNode2 AS HTREEITEM                 ' // Sub node handle
   ' // Documentation
   DIM dwsName AS DWSTRING                    ' // Name
   DIM dwsHelpString AS DWSTRING              ' // Help string
   DIM dwHelpContext AS DWORD                 ' // Help context number
   DIM dwsOrigName AS DWSTRING                ' // Original name
   DIM dwsAliasName AS DWSTRING               ' // Alias name
   DIM dwsAliasName2 AS DWSTRING              ' // Alias name (reversed)
   DIM dwsTypedefName AS DWSTRING             ' // Typedef name
   DIM dwsHelpFile AS DWSTRING                ' // Help file path
   ' // Type info
   DIM dwsProgID AS DWSTRING                  ' // ProgID
   DIM dwsInProcServer AS DWSTRING            ' // InProcServer path
   DIM cImplTypes AS LONG                     ' // Number of implemented interfaces
   DIM lImplTypeFlags AS LONG                 ' // Implemented interfaces flags
   DIM pRefType AS HREFTYPE                   ' // Address to a referenced type description
   DIM pImplTypeInfo AS Afx_ITypeInfo PTR     ' // Reference to the implemented interface TypeInfo
   DIM pImplTypeAttr AS TYPEATTR PTR          ' // Attributes of the implemented interface
   DIM hImplIntSubNode AS HTREEITEM           ' // Implemented interfaces subnode
   DIM pRefTypeInfo AS Afx_ITypeInfo PTR      ' // Referenced TypeInfo interface
   DIM pRefTypeAttr AS TYPEATTR PTR           ' // Attributes of referenced TypeInfo interface
   DIM dwsImplInterface AS DWSTRING           ' // Implemented interface
   DIM dwsInheritedInterface AS DWSTRING      ' // Inherited interface

   IF m_pTypeLib = NULL OR m_pWindow = NULL THEN RETURN E_INVALIDARG

   ' // Treeview handle
   hTreeView = cast(HWND, m_pWindow->UserData(AFX_HTREEVIEW))
   ' // Disables redraw to reduce flicker
'   SendMessageW(hTreeView, WM_SETREDRAW, FALSE, 0)

   ' // Retrieves the number of TypeInfos
   DIM TypeInfoCount AS LONG = m_pTypeLib->GetTypeInfoCount
   IF TypeInfoCount = 0 THEN
      TLB_MsgBox m_pWindow->hWindow, "This TypeLib doesn't have type infos", _
         MB_OK OR MB_ICONERROR OR MB_APPLMODAL, "CParseTypeLib.ParseTypeInfos"
      RETURN S_FALSE
   END IF

   ' // Get the use automation view option
   m_VtableView = TRUE
   DIM wszBuffer AS WSTRING * MAX_PATH
   DIM wszIniFileName AS WSTRING * MAX_PATH = ExePath & "\" & INIFILENAME
   GetPrivateProfileStringW("Browser options", "UseAutomationView", NULL, @wszBuffer, MAX_PATH, @wszIniFileName)
   IF VAL(wszBuffer) = 1 THEN m_VtableView = FALSE   ' // Use Automation view

   ' ==================================================================================
   ' // First retrieve the events interfaces
   ' ==================================================================================
   FOR i AS LONG = 0 TO TypeInfoCount - 1
      ' // Retrieves the TypeKind
      DIM hr AS HRESULT = m_pTypeLib->GetTypeInfoType(i, @pTKind)
      IF hr <> S_OK THEN EXIT FOR
      ' // Retrieves the ITypeInfo interface
      hr = m_pTypeLib->GetTypeInfo(i, @pTypeInfo)
      IF hr <> S_OK OR pTypeInfo = NULL THEN EXIT FOR
      ' // Gets the address of a pointer to the TYPEATTR structure
      hr = pTypeInfo->GetTypeAttr(@pTypeAttr)
      IF hr <> S_OK OR pTypeAttr = NULL THEN EXIT FOR
      IF pTKind = TKIND_COCLASS THEN
         ' // Retrieve the implemented interfaces
         cImplTypes = pTypeAttr->cImplTypes
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
            pImplTypeAttr = 0
            pImplTypeInfo->GetTypeAttr(@pImplTypeAttr)
            IF lImplTypeFlags = IMPLTYPEFLAG_FSOURCE OR _   ' // Events interface
               lImplTypeFlags = (IMPLTYPEFLAG_FDEFAULT OR IMPLTYPEFLAG_FSOURCE) THEN   ' // Default events interface
               dwsEvents += "#" & dwsName & "#"
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
      END IF
      IF pTypeAttr THEN
         IF pTypeInfo THEN pTypeInfo->ReleaseTypeAttr(pTypeAttr)
         pTypeAttr = NULL
      END IF
      IF pTypeInfo THEN
         pTypeInfo->Release
         pTypeInfo = NULL
      END IF
   NEXT
   ' // Cleanup
   IF pTypeAttr THEN pTypeInfo->ReleaseTypeAttr(pTypeAttr)
   IF pTypeInfo THEN pTypeInfo->Release
   ' ==================================================================================

   ' ==================================================================================
   ' // Parse the type infos
   ' ==================================================================================
   FOR i AS LONG = 0 TO TypeInfoCount - 1

      ' // Allow for a breath and check the abort flag
'      TLB_DoEvents hWnd
'      IF m_Abort THEN EXIT FOR
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

      ' // Retrieve information
      SELECT CASE pTKind

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

         ' ----------------------------------------------------------------------------
         ' Structures and unions
         ' ----------------------------------------------------------------------------
         CASE TKIND_RECORD, TKIND_UNION
            DIM AS AFX_BSTR bstrName, bstrHelpString, bstrHelpFile
            hr = m_pTypeLib->GetDocumentation(i, @bstrName, @bstrHelpString, @dwHelpContext, @bstrHelpFile)
            dwsName = *bstrName : SysFreeString bstrName
            dwsHelpString = *bstrHelpString : SysFreeString bstrHelpString
            dwsHelpFile = *bstrHelpFile : SysFreeString bstrHelpFile
            IF hr = S_OK THEN
               IF pTKind = TKIND_RECORD THEN
                  hNode = TreeView_AddItem(hTreeView, m_hRecordsNode, NULL, dwsName)
               ELSE
                  hNode = TreeView_AddItem(hTreeView, m_hUnionsNode, NULL, dwsName)
               END IF
               IF AfxGuidText(pTypeAttr->guid) <> "{00000000-0000-0000-0000-000000000000}" THEN TreeView_AddItem(hTreeView, hNode, NULL, "IID: " & AfxGuidText(pTypeAttr->guid))
               IF LEN(dwsHelpString) THEN TreeView_AddItem(hTreeView, hNode, NULL, "Documentation string: " & dwsHelpString)
               IF pTypeAttr->cVars THEN
                  hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Number of members: " & WSTR(pTypeAttr->cVars))
                  this.GetMembers(pTypeInfo, pTypeAttr, hTreeView, hSubNode, TRUE)
               END IF
            END IF
         ' ----------------------------------------------------------------------------

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

         ' ----------------------------------------------------------------------------
         ' Enumerations and modules
         ' ----------------------------------------------------------------------------
         CASE TKIND_ENUM, TKIND_MODULE
            DIM AS AFX_BSTR bstrName, bstrHelpString, bstrHelpFile
            hr = m_pTypeLib->GetDocumentation(i, @bstrName, @bstrHelpString, @dwHelpContext, @bstrHelpFile)
            dwsName = *bstrName : SysFreeString bstrName
            dwsHelpString = bstrHelpString : SysFreeString bstrHelpString
            dwsHelpFile = *bstrHelpFile : SysFreeString bstrHelpFile
            IF hr = S_OK THEN
               IF pTKind = TKIND_ENUM THEN
                  hNode = TreeView_AddItem(hTreeView, m_hEnumsNode, NULL, dwsName)
               ELSE
                  hNode = TreeView_AddItem(hTreeView, m_hModulesNode, NULL, dwsName)
               END IF
               IF AfxGuidText(pTypeAttr->guid) <> "{00000000-0000-0000-0000-000000000000}" THEN TreeView_AddItem(hTreeView, hNode, NULL, "IID: " & AfxGuidText(pTypeAttr->guid))
               IF LEN(dwsHelpString) THEN TreeView_AddItem(hTreeView, hNode, NULL, "Documentation string: " & dwsHelpString)
               ' /*** Constants ***/
               IF pTypeAttr->cVars THEN
                  hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Number of constants: " & WSTR(pTypeAttr->cVars))
                  this.GetConstants(pTypeInfo, pTypeAttr, hTreeView, hSubNode)
               END IF
               ' /*** Functions ***/
               IF pTypeAttr->cFuncs THEN
                  hSubNode = TreeView_AddItem(hTreeView, hNode, NULL, "Number of functions " & WSTR(pTypeAttr->cFuncs))
                  this.GetFunctions(pTypeInfo, pTypeAttr, hTreeView, hSubNode, FALSE)
               END IF
            END IF
         ' ----------------------------------------------------------------------------

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
'         ' ----------------------------------------------------------------------------

      END SELECT

      IF pTypeAttr THEN
         IF pTypeInfo THEN pTypeInfo->ReleaseTypeAttr(pTypeAttr)
         pTypeAttr = NULL
      END IF

      IF pTypeInfo THEN
         pTypeInfo->Release
         pTypeInfo = NULL
      END IF

   NEXT

   ' // Cleanup
   IF pTypeAttr THEN pTypeInfo->ReleaseTypeAttr(pTypeAttr)
   IF pTypeInfo THEN pTypeInfo->Release

   ' ==================================================================================

   ' // Sort the items
   TreeView_SortChildren(hTreeView, hImplIntSubNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hProgIDsNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hVerIndProgIDsNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hClsIDsNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hCoClassesNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hEnumsNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hModulesNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hRecordsNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hUnionsNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hAliasesNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hTypedefsNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hIIDsNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hIntNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hOleAutIntNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hDualIntNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hDispIntNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hDispblIntNode, FALSE)
   TreeView_SortChildren(hTreeView, m_hEventsNode, FALSE)

   ' // Remove empty nodes
   IF TreeView_GetChild(hTreeView, m_hProgIDsNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hProgIDsNode)
      m_hProgIDsNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hVerIndProgIDsNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hVerIndProgIDsNode)
      m_hVerIndProgIDsNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hClsIDsNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hClsIDsNode)
      m_hClsIDsNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hCoClassesNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hCoClassesNode)
      m_hCoClassesNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hEnumsNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hEnumsNode)
      m_hEnumsNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hModulesNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hModulesNode)
      m_hModulesNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hRecordsNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hRecordsNode)
      m_hRecordsNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hUnionsNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hUnionsNode)
      m_hUnionsNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hAliasesNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hAliasesNode)
      m_hAliasesNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hTypedefsNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hTypedefsNode)
      m_hTypedefsNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hIIDsNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hIIDsNode)
      m_hIIDsNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hIntNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hIntNode)
      m_hIntNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hOleAutIntNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hOleAutIntNode)
      m_hOleAutIntNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hDualIntNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hDualIntNode)
      m_hDualIntNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hDispIntNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hDispIntNode)
      m_hDispIntNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hDispblIntNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hDispblIntNode)
      m_hDispblIntNode = NULL
   END IF
   IF TreeView_GetChild(hTreeView, m_hEventsNode) = NULL THEN
      TreeView_DeleteItem(hTreeView, m_hEventsNode)
      m_hEventsNode = NULL
   END IF

   ' // Enables redraw and repaints the control
'   SendMessageW(hTreeView, WM_SETREDRAW, TRUE, 0)

   RETURN S_OK

END FUNCTION
' =====================================================================================

' =====================================================================================
SUB CParseTypeLib.GenerateCode

   IF m_pWindow = NULL THEN EXIT SUB
   DIM hTreeView AS HWND = cast(HWND, m_pWindow->UserData(AFX_HTREEVIEW))
   DIM dwsCode AS DWSTRING
   DIM dwsMacros AS DWSTRING
   ' // We can increase the capacity of the buffer for faster concatenation
   'dwsCode.Capacity = xx MB
   'dwsMacros.Capacity = xx MB

   ' // Header
   DIM wszFmt AS WSTRING * 15     ' // Date format
   DIM wszDate AS WSTRING * 256   ' // Formatted date
   wszFmt = "dd MMM yyyy"
   GetDateFormatW LOCALE_USER_DEFAULT, 0, NULL, @wszFmt, @wszDate, SIZEOF(wszDate)
   dwsCode += "' ########################################################################################" & CHR(13, 10)
   dwsCode += "' Library name: " & m_LibName & CHR(13, 10)
   IF LEN(m_LibHelpString) THEN dwsCode += "' Documentation string: " & m_LibHelpString & CHR(13, 10)
   dwsCode += "' GUID: " & m_LibGuid & CHR(13, 10)
   dwsCode += "' Version: " & WSTR(m_LibMajorVersion) & "." & WsTR(m_LibMinorVersion) & ", Locale ID = " & WSTR(m_LibLcid) & CHR(13, 10)
   dwsCode += "' Path: " & m_LibPath & CHR(13, 10)
   dwsCode += "' Attributes: " & STR(m_LibAttr) & " [&h" & HEX(m_LibAttr, 8) & "] " & TLB_LibFlagsToStr(m_LibAttr) & CHR(13, 10)
   IF LEN(m_LibHelpFile) THEN dwsCode += "' Help file: " & m_LibHelpFile & CHR(13, 10)
   'IF LEN(m_strLicKey) THEN dwsCode += "' License key: " & m_strLicKey & CHR(13, 10)
   dwsCode += "' Code generated by the " & TLBCAPTION & CHR(13, 10)
   dwsCode += "' " & TLBCOPYRIGHT & ". All rights reserved. Freeware. Use at your own risk." & CHR(13, 10)
   dwsCode += "' Date: " & wszDate & " Time: " & TIME & CHR(13, 10)
   dwsCode += "' ########################################################################################" & CHR(13, 10)
   dwsCode += CHR(13, 10)
   dwsCode += "#pragma once" & CHR(13, 10)
   dwsCode += "#include once " & CHR(34) & "AfxNova/AfxCOM.inc" & CHR(34) & CHR(13, 10)
   dwsCode += "USING AfxNova" & CHR(13, 10) & CHR(13, 10)

'   IF m_vTableView THEN
'      dwsCode += "NAMESPACE AfxNova" & CHR(13,10) & CHR(13,10)
'   ELSE
'      IF LEN(m_Namespace) THEN dwsCode += "NAMESPACE " & m_Namespace & CHR(13,10) & CHR(13,10)
'   END IF

   ' // ProgIDs
   IF m_hProgIDsNode THEN
      dwsCode += "' // ProgIDs (Program Identifiers)" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH, wszText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hProgIDsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            IF m_vTableView THEN
               wszText = "CONST AFX_PROGID_" & DWStrReplace(wszItemText, ".", "_") & " = "
            ELSE
               wszText = "CONST PROGID_" & DWStrReplace(wszItemText, ".", "_") & " = "
            END IF
            dwsCode += wszText & CHR(34) & wszItemText & CHR(34) & CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Version independent ProgIDs
   IF m_hVerIndProgIDsNode THEN
      dwsCode += "' // Version independent ProgIDs" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH, wszText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hVerIndProgIDsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            IF m_vTableView THEN
               wszText = "CONST AFX_PROGID_" & DWStrReplace(wszItemText, ".", "_") & " = "
            ELSE
               wszText = "CONST PROGID_" & DWStrReplace(wszItemText, ".", "_") & " = "
            END IF
            dwsCode += wszText & CHR(34) & wszItemText & CHR(34) & CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // ClsIDs (Class identifiers)
   IF m_hClsIDsNode THEN
      dwsCode += "' // ClsIDs (Class identifiers)" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hClsIDsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            DIM guidText AS DWSTRING = LEFT(RIGHT(wszItemText, 38), 36)
            DIM guidName AS DWSTRING = LEFT(wszItemText, LEN(wszItemText) - 43)
            IF m_vTableView THEN
'               dwsCode += "CONST AFX_" & wszItemText & CHR(13, 10)
               dwsCode += "DIM SHARED AFX_" & guidName & " AS GUID = " & AfxGuidToFmtStr(guidText) & CHR(13, 10)
            ELSE
'               dwsCode += "CONST " & wszItemText & CHR(13, 10)
               dwsCode += "DIM SHARED " & guidName & " AS GUID = "  & AfxGuidToFmtStr(guidText) & CHR(13, 10)
            END IF
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // IIDs (Interface identifiers)
   IF m_hIIDsNode THEN
      dwsCode += "' // IIDs (Interface identifiers)" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hIIDsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            DIM guidText AS DWSTRING = LEFT(RIGHT(wszItemText, 38), 36)
            DIM guidName AS DWSTRING = LEFT(wszItemText, LEN(wszItemText) - 43)
            IF m_vTableView THEN
'               dwsCode += "CONST AFX_" & wszItemText & CHR(13, 10)
               dwsCode += "DIM SHARED AFX_" & guidName & " AS GUID = " & AfxGuidToFmtStr(guidText) & CHR(13, 10)
            ELSE
'               dwsCode += "CONST " & wszItemText & CHR(13, 10)
               dwsCode += "DIM SHARED " & guidName & " AS GUID = "  & AfxGuidToFmtStr(guidText) & CHR(13, 10)
            END IF
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Typedefs
   IF m_hTypeDefsNode THEN
      dwsCode += "' // Typedefs" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hTypeDefsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            wszItemText = DWStrReplace(wszItemText, "=", "AS")
            dwsCode += "TYPE " & wszItemText & CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Aliases
   IF m_hAliasesNode THEN
      dwsCode += "' // Aliases" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hAliasesNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            wszItemText = DWStrReplace(wszItemText, "=", "AS")
            dwsCode += "TYPE " & wszItemText & CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Enumerations
   IF m_hEnumsNode THEN
      dwsCode += "' // Enumerations" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hEnumsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            dwsCode += "ENUM " & wszItemText & CHR(13, 10)
            ' // Search the first subnode
            DIM hSubNode AS HTREEITEM
            DIM dwsSubNodeText AS DWSTRING, dwsSubItemsText AS DWSTRING
            hSubNode = TreeView_GetChild(hTreeView, hItem)
            DO
               IF hSubNode = NULL THEN EXIT DO
               IF TreeView_GetItemText(hTreeView, hSubNode, @wszItemText, MAX_PATH) THEN
                  dwsSubNodeText += "   ' // " & wszItemText & CHR(13, 10)
               END IF
               ' // Search the second subnode
               IF INSTR(wszItemText, "Number of constants") THEN
                  DIM hSubItem AS HTREEITEM
                  hSubItem = TreeView_GetChild(hTreeView, hSubNode)
                  DO
                     IF hSubItem = NULL THEN EXIT DO
                     IF TreeView_GetItemText(hTreeView, hSubItem, @wszItemText, MAX_PATH) THEN
                        dwsSubItemsText += "   " & wszItemText & CHR(13, 10)
                     END IF
                     hSubItem = TreeView_GetNextSibling(hTreeView, hSubItem)
                  LOOP
               END IF
               hSubNode = TreeView_GetNextSibling(hTreeView, hSubNode)
            LOOP
            dwsCode += dwsSubNodeText & dwsSubItemsText
            dwsCode += "END ENUM" & CHR(13, 10) & CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Modules
   IF m_hModulesNode THEN
      dwsCode += "' // Modules" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hModulesNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            dwsCode += "' // Module: " & wszItemText & CHR(13, 10)
            ' // Search the first subnode
            DIM hSubNode AS HTREEITEM
            DIM dwsSubNodeText AS DWSTRING, dwsSubItemsText AS DWSTRING
            hSubNode = TreeView_GetChild(hTreeView, hItem)
            DO
               IF hSubNode = NULL THEN EXIT DO
               IF TreeView_GetItemText(hTreeView, hSubNode, @wszItemText, MAX_PATH) THEN
                  dwsSubNodeText += "' // " & wszItemText & CHR(13, 10)
               END IF
               ' // Search the second subnode
               IF INSTR(wszItemText, "Number of constants") THEN
                  DIM hSubItem AS HTREEITEM
                  hSubItem = TreeView_GetChild(hTreeView, hSubNode)
                  DO
                     IF hSubItem = NULL THEN EXIT DO
                     IF TreeView_GetItemText(hTreeView, hSubItem, @wszItemText, MAX_PATH) THEN
                        dwsSubItemsText += "CONST " & wszItemText & CHR(13, 10)
                     END IF
                     hSubItem = TreeView_GetNextSibling(hTreeView, hSubItem)
                  LOOP
               END IF
               hSubNode = TreeView_GetNextSibling(hTreeView, hSubNode)
            LOOP
            dwsCode += dwsSubNodeText & dwsSubItemsText
            dwsCode += CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

'   ' // Forward references for structures
'   IF m_hRecordsNode THEN
'      dwsCode += "' // Structures - Forward references" & CHR(13, 10) & CHR(13, 10)
'      DIM wszItemText AS WSTRING * MAX_PATH
'      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hRecordsNode)
'      DO
'         IF hItem = NULL THEN EXIT DO
'         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
'            dwsCode += "TYPE tag_" & wszItemText & " AS " & wszItemText & CHR(13, 10)
'         END IF
'         hItem = TreeView_GetNextSibling(hTreeView, hItem)
'      LOOP
'      dwsCode += CHR(13, 10)
'   END IF

'   ' // Forward references for unions
'   IF m_hUnionsNode THEN
'      dwsCode += "' // Unions - Forward references" & CHR(13, 10) & CHR(13, 10)
'      DIM wszItemText AS WSTRING * MAX_PATH
'      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hUnionsNode)
'      DO
'         IF hItem = NULL THEN EXIT DO
'         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
'            dwsCode += "TYPE tag_" & wszItemText & " AS " & wszItemText & CHR(13, 10)
'         END IF
'         hItem = TreeView_GetNextSibling(hTreeView, hItem)
'      LOOP
'      dwsCode += CHR(13, 10)
'   END IF

   ' // Structures
   IF m_hRecordsNode THEN
      dwsCode += "' // Structures" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hRecordsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            dwsCode += "TYPE " & wszItemText & CHR(13, 10)
            ' // Search the first subnode
            DIM hSubNode AS HTREEITEM
            DIM dwsSubNodeText AS DWSTRING, dwsSubItemsText AS DWSTRING
            hSubNode = TreeView_GetChild(hTreeView, hItem)
            DO
               IF hSubNode = NULL THEN EXIT DO
               IF TreeView_GetItemText(hTreeView, hSubNode, @wszItemText, MAX_PATH) THEN
                  dwsSubNodeText += "   ' // " & wszItemText & CHR(13, 10)
               END IF
               ' // Search the second subnode
               IF INSTR(wszItemText, "Number of members") THEN
                  DIM hSubItem AS HTREEITEM
                  hSubItem = TreeView_GetChild(hTreeView, hSubNode)
                  DO
                     IF hSubItem = NULL THEN EXIT DO
                     ' // Search the third subnode
                     DIM hSubItem2 AS HTREEITEM
                     hSubItem2 = TreeView_GetChild(hTreeView, hSubItem)
                     DO
                        IF hSubItem2 = NULL THEN EXIT DO
                        IF TreeView_GetItemText(hTreeView, hSubItem2, @wszItemText, MAX_PATH) THEN
                           IF INSTR(wszItemText, "FB syntax =") THEN
                              dwsSubItemsText += "   " & MID(wszItemText, 13) & CHR(13, 10)
                              EXIT DO
                           END IF
                        END IF
                        hSubItem2 = TreeView_GetNextSibling(hTreeView, hSubItem2)
                     LOOP
                     hSubItem = TreeView_GetNextSibling(hTreeView, hSubItem)
                  LOOP
               END IF
               hSubNode = TreeView_GetNextSibling(hTreeView, hSubNode)
            LOOP
            dwsCode += dwsSubNodeText & dwsSubItemsText
            dwsCode += "END TYPE" & CHR(13, 10) & CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Unions
   IF m_hUnionsNode THEN
      dwsCode += "' // Unions" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hUnionsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            dwsCode += "UNION " & wszItemText & CHR(13, 10)
            ' // Search the first subnode
            DIM hSubNode AS HTREEITEM
            DIM dwsSubNodeText AS DWSTRING, dwsSubItemsText AS DWSTRING
            hSubNode = TreeView_GetChild(hTreeView, hItem)
            DO
               IF hSubNode = NULL THEN EXIT DO
               IF TreeView_GetItemText(hTreeView, hSubNode, @wszItemText, MAX_PATH) THEN
                  dwsSubNodeText += "   ' // " & wszItemText & CHR(13, 10)
               END IF
               ' // Search the second subnode
               IF INSTR(wszItemText, "Number of members") THEN
                  DIM hSubItem AS HTREEITEM
                  hSubItem = TreeView_GetChild(hTreeView, hSubNode)
                  DO
                     IF hSubItem = NULL THEN EXIT DO
                     ' // Search the third subnode
                     DIM hSubItem2 AS HTREEITEM
                     hSubItem2 = TreeView_GetChild(hTreeView, hSubItem)
                     DO
                        IF hSubItem2 = NULL THEN EXIT DO
                        IF TreeView_GetItemText(hTreeView, hSubItem2, @wszItemText, MAX_PATH) THEN
                           IF INSTR(wszItemText, "FB syntax =") THEN
                              dwsSubItemsText += "   " & MID(wszItemText, 13) & CHR(13, 10)
                              EXIT DO
                           END IF
                        END IF
                        hSubItem2 = TreeView_GetNextSibling(hTreeView, hSubItem2)
                     LOOP
                     hSubItem = TreeView_GetNextSibling(hTreeView, hSubItem)
                  LOOP
               END IF
               hSubNode = TreeView_GetNextSibling(hTreeView, hSubNode)
            LOOP
            dwsCode += dwsSubNodeText & dwsSubItemsText
            dwsCode += "END UNION" & CHR(13, 10) & CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Forward references for interfaces
   IF m_hIntNode THEN
      dwsCode += "' // Interfaces - Forward references" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hIntNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            ' // ADO contains many interfaces that end with "_Deprecated"
            IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
               IF m_vTableView THEN
                  dwsCode += "TYPE Afx_" & wszItemText & " AS Afx_" & wszItemText & "_" & CHR(13, 10)
               ELSE
                  dwsCode += "TYPE " & wszItemText & " AS " & wszItemText & "_" & CHR(13, 10)
               END IF
            END IF
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Forward references for Ole Automation interfaces
   IF m_hOleAutIntNode THEN
      dwsCode += "' // Ole Automation interfaces - Forward references" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hOleAutIntNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            IF m_vTableView THEN
               dwsCode += "TYPE Afx_" & wszItemText & " AS Afx_" & wszItemText & "_" & CHR(13, 10)
            ELSE
               dwsCode += "TYPE " & wszItemText & " AS " & wszItemText & "_" & CHR(13, 10)
            END IF
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Forward references for dual interfaces
   IF m_hDualIntNode THEN
      dwsCode += "' // Dual interfaces - Forward references" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hDualIntNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            ' // ADO contains many interfaces that end with "_Deprecated"
            IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
               IF m_vTableView THEN
                  dwsCode += "TYPE Afx_" & wszItemText & " AS Afx_" & wszItemText & "_" & CHR(13, 10)
               ELSE
                  dwsCode += "TYPE " & wszItemText & " AS " & wszItemText & "_" & CHR(13, 10)
               END IF
            END IF
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Forward references for dispatch interfaces
   IF m_hDispIntNode THEN
      dwsCode += "' // Dispatch interfaces - Forward references" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hDispIntNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            ' // ADO contains many interfaces that end with "_Deprecated"
'            IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
               IF m_vTableView THEN
                  dwsCode += "TYPE Afx_" & wszItemText & " AS Afx_" & wszItemText & "_" & CHR(13, 10)
               ELSE
                  dwsCode += "TYPE " & wszItemText & " AS " & wszItemText & "_" & CHR(13, 10)
               END IF
'            END IF
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Forward references for dispatchable interfaces
   IF m_hDispblIntNode THEN
      dwsCode += "' // Dispatchable interfaces - Forward references" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hDispblIntNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            IF m_vTableView THEN
               dwsCode += "TYPE Afx_" & wszItemText & " AS Afx_" & wszItemText & "_" & CHR(13, 10)
            ELSE
               dwsCode += "TYPE " & wszItemText & " AS " & wszItemText & "_" & CHR(13, 10)
            END IF
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Forward references for event interfaces
   IF m_hEventsNode THEN
      dwsCode += "' // Event interfaces - Forward references" & CHR(13, 10) & CHR(13, 10)
      DIM wszItemText AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hEventsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            IF m_vTableView THEN
               dwsCode += "TYPE Afx_" & wszItemText & " AS Afx_" & wszItemText & "_" & CHR(13, 10)
            ELSE
               dwsCode += "TYPE " & wszItemText & " AS " & wszItemText & "_" & CHR(13, 10)
            END IF
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Interfaces
   IF m_hIntNode THEN
      dwsMacros.Clear
      dwsCode += "' // Interfaces" & CHR(13, 10) & CHR(13, 10)
      DIM hMethodsParentNode AS HTREEITEM
      DIM wszItemText AS WSTRING * MAX_PATH, wszInheritedInterface AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hIntNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            ' // ADO contains many interfaces that end with "_Deprecated"
            IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += "' Interface name: " & wszItemText & CHR(13, 10)
               DIM wszSubItemText AS WSTRING * MAX_PATH
               DIM hSubItem AS HTREEITEM = TreeView_GetChild(hTreeView, hItem)
               DO
                  IF hSubItem = NULL THEN EXIT DO
                  IF TreeView_GetItemText(hTreeView, hSubItem, @wszSubItemText, MAX_PATH) THEN
                     dwsCode += "' " & wszSubItemText & CHR(13, 10)
                     IF LEFT(wszSubItemText, 22) = "Inherited interface = " THEN
                        wszInheritedInterface = MID(wszSubItemText, 23)
                     END IF
                     hMethodsParentNode = NULL
                     IF LEFT(wszSubItemText, 17) = "Number of methods" AND VAL(MID(wszSubItemText, 21)) > 0 THEN
                        hMethodsParentNode = hSubItem
                     END IF
                  END IF
                  hSubItem = TreeView_GetNextSibling(hTreeView, hSubItem)
               LOOP
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += CHR(13, 10)
               IF m_VTableView THEN
                  dwsCode += "#ifndef __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               ELSE
                  dwsCode += "#ifndef __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               END IF
               dwsCode += CHR(13, 10)
               IF hMethodsParentNode THEN
                  DIM wszText AS WSTRING * MAX_PATH
                  IF m_VTableView THEN
                     IF wszItemText = "IUnknown" THEN
                        wszText = "_ EXTENDS OBJECT"
                     ELSEIF wszInheritedInterface = "IUnknown" THEN
                        wszText = "_ EXTENDS Afx_IUnknown"
                     ELSEIF wszInheritedInterface = "IDispatch" THEN
                        wszText = "_ EXTENDS Afx_IDispatch"
                     ELSE
                        wszText = "_ EXTENDS Afx_" & wszInheritedInterface
                     END IF
                     dwsCode += "TYPE Afx_" & wszItemText & wszText & CHR(13, 10)
                  ELSE
                     dwsCode += "TYPE " & wszItemText & "VTbl" & CHR(13, 10)
                  END IF
                  IF m_VTableView = FALSE THEN
                     IF wszItemText <> "IUnknown" THEN
                        dwsCode += "   QueryInterface AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT" & CHR(13, 10)
                        dwsCode += "   AddRef AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                        dwsCode += "   Release AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                        dwsMacros + = "#define " & wszItemText & "_QueryInterface(this, riid, ppvObject) (this)->lpVtbl->QueryInterface(this, riid, ppvObject)" & CHR(13, 10)
                        dwsMacros + = "#define " & wszItemText & "_AddRef(this) (this)->lpVtbl->AddRef(this)" & CHR(13, 10)
                        dwsMacros + = "#define " & wszItemText & "_Release(this) (this)->lpVtbl->Release(this)" & CHR(13, 10)
                     END IF
                  END IF
                  ' // Search the methods and properties
                  DIM hMethods AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsParentNode)
                  DIM wszMethodName AS WSTRING * MAX_PATH
                  DO
                     IF hMethods = NULL THEN EXIT DO
                     IF TreeView_GetItemText(hTreeView, hMethods, @wszMethodName, MAX_PATH) THEN
                        IF LEFT(wszMethodName, 6) = "METHOD" THEN wszMethodName = MID(wszMethodName, 8)
                        IF LEFT(wszMethodName, 12) = "PROPERTY GET" THEN wszMethodName = "get_" & MID(wszMethodName, 14)
                        IF LEFT(wszMethodName, 12) = "PROPERTY PUT" THEN wszMethodName = "put_" & MID(wszMethodName, 14)
                        IF LEFT(wszMethodName, 15) = "PROPERTY PUTREF" THEN wszMethodName = "putref_" & MID(wszMethodName, 17)
                        ' // Get the return type
                        DIM dwsParamNames AS DWSTRING, dwsParams AS DWSTRING
                        DIM wszReturnType AS WSTRING * MAX_PATH, wszMethodSubItem AS WSTRING * MAX_PATH
                        DIM hMethodsSubItems AS HTREEITEM = TreeView_GetChild(hTreeView, hMethods)
                        DO
                           IF hMethodsSubItems = NULL THEN EXIT DO
                           IF TreeView_GetItemText(hTreeView, hMethodsSubItems, @wszMethodSubItem, MAX_PATH) THEN
                              IF LEFT(wszMethodSubItem, 11) = "Return type" THEN
                                 wszReturnType = MID(wszMethodSubItem, 15)
                                 wszReturnType = TLB_VarTypeToKeyword(VAL(wszReturnType))
                              END IF
                              ' // Get the parameters
                              IF LEFT(wszMethodSubItem, 23) = "Number of parameters = " THEN
                                 DIM hParams AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsSubItems)
                                 DIM wszParamName AS WSTRING * MAX_PATH
                                 DO
                                    IF hParams = NULL THEN EXIT DO
                                    IF TreeView_GetItemText(hTreeView, hParams, @wszParamName, MAX_PATH) THEN
                                       dwsParamNames += wszParamName & ", "
                                       DIM hParam AS HTREEITEM = TreeView_GetChild(hTreeView, hParams)
                                       DIM wszParam AS WSTRING * MAX_PATH
                                       DIM wszTypeKind AS WSTRING * 260
                                       DO
                                          IF hParam = NULL THEN EXIT DO
                                          IF TreeView_GetItemText(hTreeView, hParam, @wszParam, MAX_PATH) THEN
'                                             IF LEFT(wszParam, 12) = "FB syntax = " THEN
'                                                dwsParams += MID(wszParam, 13) & ", "
'                                             END IF
                                             ' ------------------------------------------------------
                                             IF LEFT(wszParam, 11) = "TypeKind = " THEN
                                                wszTypeKind = MID(wszParam, 12)
                                             END IF
                                             IF LEFT(wszParam, 12) = "FB syntax = " THEN
                                                IF m_vTableView THEN
                                                   IF wszTypeKind = "TKIND_INTERFACE" OR wszTypeKind = "TKIND_DISPATCH" THEN
'                                                      dwsParams += "Afx_" & MID(wszParam, 13) & ", "
                                                      wszParam = DWStrReplace(wszParam, " AS ", " AS Afx_")
                                                      dwsParams += MID(wszParam, 13) & ", "
                                                   ELSE
                                                      dwsParams += MID(wszParam, 13) & ", "
                                                   END IF
                                                ELSE
                                                   dwsParams += MID(wszParam, 13) & ", "
                                                END IF
                                             END IF
                                             ' ------------------------------------------------------
                                          END IF
                                          hParam = TreeView_GetNextSibling(hTreeView, hParam)
                                       LOOP
                                    END IF
                                    hParams = TreeView_GetNextSibling(hTreeView, hParams)
                                 LOOP
                              END IF
                              ' --------------------------------------------------------------------------
                           END IF
                           hMethodsSubItems = TreeView_GetNextSibling(hTreeView, hMethodsSubItems)
                        LOOP
                        IF wszReturnType = "VOID" THEN wszReturnType = "HRESULT"
                        IF RIGHT(dwsParams, 2) = ", " THEN dwsParams = LEFT(dwsParams, LEN(dwsParams) - 2)
                        IF RIGHT(dwsParamNames, 2) = ", " THEN dwsParamNames = LEFT(dwsParamNames, LEN(dwsParamNames) - 2)
                        IF LEN(dwsParamNames) THEN dwsParamNames = ", " & dwsParamNames
                        IF m_VTableView THEN
                           dwsCode += "   DECLARE ABSTRACT FUNCTION " & wszMethodName & " (" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                        ELSE
                           IF LEN(dwsParams) THEN dwsParams = ", " & dwsParams
                           dwsCode += "   " & wszMethodName & " AS FUNCTION (BYVAL this AS " & wszItemText & " PTR" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                           dwsMacros += "#define " & wszItemText & "_" & wszMethodName & "(this" & dwsParamNames &") (this)->lpVtbl->" & wszMethodName & "(this" & dwsParamNames & ")" & CHR(13, 10)
                        END IF
                     END IF
                     hMethods = TreeView_GetNextSibling(hTreeView, hMethods)
                  LOOP
                  dwsCode += "END TYPE" & CHR(13, 10)
               END IF
            END IF
         END IF
         ' // ADO contains many interfaces that end with "_Deprecated"
         IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
            IF hMethodsParentNode THEN
               dwsCode += CHR(13, 10)
               IF m_VTableView = FALSE THEN
                  dwsCode += "TYPE " & wszItemText & "_" & CHR(13, 10)
                  dwsCode += "   lpVtbl AS " & wszItemText & "VTbl PTR" & CHR(13, 10)
                  dwsCode += "END TYPE" & CHR(13, 10) & CHR(13, 10)
                  dwsCode += "TYPE " & "LP" & UCASE(wszItemText) & " AS " & wszItemText & " PTR" & CHR(13, 10)
                  dwsCode += CHR(13, 10)
               END IF
            END IF
            dwsCode += "#endif" & CHR(13, 10) & CHR(13, 10)
            IF LEN(dwsMacros) THEN
               dwsCode += "' // C Object Macros" & CHR(13, 10)
               dwsCode += dwsMacros & CHR(13, 10)
               dwsMacros.Clear
            END IF
            dwsCode += "' ########################################################################################" & CHR(13, 10)
            dwsCode += CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Ole Automation interfaces
   IF m_hOleAutIntNode THEN
      dwsMacros.Clear
      dwsCode += "' // Ole Automation interfaces" & CHR(13, 10) & CHR(13, 10)
      DIM hMethodsParentNode AS HTREEITEM
      DIM wszItemText AS WSTRING * MAX_PATH, wszInheritedInterface AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hOleAutIntNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            ' // ADO contains many interfaces that end with "_Deprecated"
'            IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += "' Interface name: " & wszItemText & CHR(13, 10)
               DIM wszSubItemText AS WSTRING * MAX_PATH
               DIM hSubItem AS HTREEITEM = TreeView_GetChild(hTreeView, hItem)
               DO
                  IF hSubItem = NULL THEN EXIT DO
                  IF TreeView_GetItemText(hTreeView, hSubItem, @wszSubItemText, MAX_PATH) THEN
                     dwsCode += "' " & wszSubItemText & CHR(13, 10)
                     IF LEFT(wszSubItemText, 22) = "Inherited interface = " THEN
                        wszInheritedInterface = MID(wszSubItemText, 23)
'                        IF wszInheritedInterface = "IUnknown" OR wszInheritedInterface = "IDispatch" THEN wszInheritedInterface = "Afx_" & wszInheritedInterface
                     END IF
                     hMethodsParentNode = NULL
                     IF LEFT(wszSubItemText, 17) = "Number of methods" AND VAL(MID(wszSubItemText, 21)) > 0 THEN
                        hMethodsParentNode = hSubItem
                     END IF
                  END IF
                  hSubItem = TreeView_GetNextSibling(hTreeView, hSubItem)
               LOOP
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += CHR(13, 10)
               IF m_vTableView THEN
                  dwsCode += "#ifndef __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               ELSE
                  dwsCode += "#ifndef __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               END IF
               dwsCode += CHR(13, 10)
               IF hMethodsParentNode THEN
                  DIM wszText AS WSTRING * MAX_PATH
                  IF m_VTableView THEN
                     IF wszInheritedInterface = "IUnknown" THEN
                        wszText = "_ EXTENDS Afx_IUnknown"
                     ELSEIF wszInheritedInterface = "IDispatch" THEN
                        wszText = "_ EXTENDS Afx_IDispatch"
                     ELSE
                        wszText = "_ EXTENDS Afx_" & wszInheritedInterface
                     END IF
                     dwsCode += "TYPE Afx_" & wszItemText & wszText & CHR(13, 10)
                  ELSE
                     dwsCode += "TYPE " & wszItemText & "VTbl" & CHR(13, 10)
                  END IF
                  IF m_VTableView = FALSE THEN
                     dwsCode += "   QueryInterface AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT" & CHR(13, 10)
                     dwsCode += "   AddRef AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                     dwsCode += "   Release AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                     dwsMacros + = "#define " & wszItemText & "_QueryInterface(this, riid, ppvObject) (this)->lpVtbl->QueryInterface(this, riid, ppvObject)" & CHR(13, 10)
                     dwsMacros + = "#define " & wszItemText & "_AddRef(this) (this)->lpVtbl->AddRef(this)" & CHR(13, 10)
                     dwsMacros + = "#define " & wszItemText & "_Release(this) (this)->lpVtbl->Release(this)" & CHR(13, 10)
                  END IF
                  ' // Search the methods and properties
                  DIM hMethods AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsParentNode)
                  DIM wszMethodName AS WSTRING * MAX_PATH
                  DO
                     IF hMethods = NULL THEN EXIT DO
                     IF TreeView_GetItemText(hTreeView, hMethods, @wszMethodName, MAX_PATH) THEN
                        IF LEFT(wszMethodName, 6) = "METHOD" THEN wszMethodName = MID(wszMethodName, 8)
                        IF LEFT(wszMethodName, 12) = "PROPERTY GET" THEN wszMethodName = "get_" & MID(wszMethodName, 14)
                        IF LEFT(wszMethodName, 12) = "PROPERTY PUT" THEN wszMethodName = "put_" & MID(wszMethodName, 14)
                        IF LEFT(wszMethodName, 15) = "PROPERTY PUTREF" THEN wszMethodName = "putref_" & MID(wszMethodName, 17)
                        ' // Get the return type
                        DIM dwsParamNames AS DWSTRING, dwsParams AS DWSTRING
                        DIM wszReturnType AS WSTRING * MAX_PATH, wszMethodSubItem AS WSTRING * MAX_PATH
                        DIM hMethodsSubItems AS HTREEITEM = TreeView_GetChild(hTreeView, hMethods)
                        DO
                           IF hMethodsSubItems = NULL THEN EXIT DO
                           IF TreeView_GetItemText(hTreeView, hMethodsSubItems, @wszMethodSubItem, MAX_PATH) THEN
                              IF LEFT(wszMethodSubItem, 11) = "Return type" THEN
                                 wszReturnType = MID(wszMethodSubItem, 15)
                                 wszReturnType = TLB_VarTypeToKeyword(VAL(wszReturnType))
                              END IF
                              ' --------------------------------------------------------------------------
                              ' // Get the parameters
                              IF LEFT(wszMethodSubItem, 23) = "Number of parameters = " THEN
                                 DIM hParams AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsSubItems)
                                 DIM wszParamName AS WSTRING * MAX_PATH
                                 DO
                                    IF hParams = NULL THEN EXIT DO
                                    IF TreeView_GetItemText(hTreeView, hParams, @wszParamName, MAX_PATH) THEN
                                       dwsParamNames += wszParamName & ", "
                                       DIM hParam AS HTREEITEM = TreeView_GetChild(hTreeView, hParams)
                                       DIM wszParam AS WSTRING * MAX_PATH
                                       DIM wszTypeKind AS WSTRING * 260
                                       DO
                                          IF hParam = NULL THEN EXIT DO
                                          IF TreeView_GetItemText(hTreeView, hParam, @wszParam, MAX_PATH) THEN
'                                             IF LEFT(wszParam, 12) = "FB syntax = " THEN
'                                                dwsParams += MID(wszParam, 13) & ", "
'                                             END IF
                                             ' ------------------------------------------------------
                                             IF LEFT(wszParam, 11) = "TypeKind = " THEN
                                                wszTypeKind = MID(wszParam, 12)
                                             END IF
                                             IF LEFT(wszParam, 12) = "FB syntax = " THEN
                                                IF m_vTableView THEN
                                                   IF wszTypeKind = "TKIND_INTERFACE" OR wszTypeKind = "TKIND_DISPATCH" THEN
'                                                      dwsParams += "Afx_" & MID(wszParam, 13) & ", "
                                                      wszParam = DWStrReplace(wszParam, " AS ", " AS Afx_")
                                                      dwsParams += MID(wszParam, 13) & ", "
                                                   ELSE
                                                      dwsParams += MID(wszParam, 13) & ", "
                                                   END IF
                                                ELSE
                                                   dwsParams += MID(wszParam, 13) & ", "
                                                END IF
                                             END IF
                                             ' ------------------------------------------------------
                                          END IF
                                          hParam = TreeView_GetNextSibling(hTreeView, hParam)
                                       LOOP
                                    END IF
                                    hParams = TreeView_GetNextSibling(hTreeView, hParams)
                                 LOOP
                              END IF
                              ' --------------------------------------------------------------------------
                           END IF
                           hMethodsSubItems = TreeView_GetNextSibling(hTreeView, hMethodsSubItems)
                        LOOP
                        IF wszReturnType = "VOID" THEN wszReturnType = "HRESULT"
                        IF RIGHT(dwsParams, 2) = ", " THEN dwsParams = LEFT(dwsParams, LEN(dwsParams) - 2)
                        IF RIGHT(dwsParamNames, 2) = ", " THEN dwsParamNames = LEFT(dwsParamNames, LEN(dwsParamNames) - 2)
                        IF LEN(dwsParamNames) THEN dwsParamNames = ", " & dwsParamNames
                        IF m_VTableView THEN
                           dwsCode += "   DECLARE ABSTRACT FUNCTION " & wszMethodName & " (" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                        ELSE
                           IF LEN(dwsParams) THEN dwsParams = ", " & dwsParams
                           dwsCode += "   " & wszMethodName & " AS FUNCTION (BYVAL this AS " & wszItemText & " PTR" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                           dwsMacros += "#define " & wszItemText & "_" & wszMethodName & "(this" & dwsParamNames &") (this)->lpVtbl->" & wszMethodName & "(this" & dwsParamNames & ")" & CHR(13, 10)
                        END IF
                     END IF
                     hMethods = TreeView_GetNextSibling(hTreeView, hMethods)
                  LOOP
                  dwsCode += "END TYPE" & CHR(13, 10)
               END IF
'            END IF
         END IF
         ' // ADO contains many interfaces that end with "_Deprecated"
'         IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
            IF hMethodsParentNode THEN
               dwsCode += CHR(13, 10)
               IF m_VTableView = FALSE THEN
                  dwsCode += "TYPE " & wszItemText & "_" & CHR(13, 10)
                  dwsCode += "   lpVtbl AS " & wszItemText & "VTbl PTR" & CHR(13, 10)
                  dwsCode += "END TYPE" & CHR(13, 10) & CHR(13, 10)
                  dwsCode += "TYPE " & "LP" & UCASE(wszItemText) & " AS " & wszItemText & " PTR" & CHR(13, 10)
                  dwsCode += CHR(13, 10)
               END IF
            END IF
            dwsCode += "#endif" & CHR(13, 10) & CHR(13, 10)
            IF LEN(dwsMacros) THEN
               dwsCode += "' // C Object Macros" & CHR(13, 10)
               dwsCode += dwsMacros & CHR(13, 10)
               dwsMacros.Clear
            END IF
            dwsCode += "' ########################################################################################" & CHR(13, 10)
            dwsCode += CHR(13, 10)
'         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Dispatchable interfaces
   IF m_hDispblIntNode THEN
      dwsMacros.Clear
      dwsCode += "' // Dispatchable interfaces" & CHR(13, 10) & CHR(13, 10)
      DIM hMethodsParentNode AS HTREEITEM
      DIM wszItemText AS WSTRING * MAX_PATH, wszInheritedInterface AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hDispblIntNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            ' // ADO contains many interfaces that end with "_Deprecated"
'            IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += "' Interface name: " & wszItemText & CHR(13, 10)
               DIM wszSubItemText AS WSTRING * MAX_PATH
               DIM hSubItem AS HTREEITEM = TreeView_GetChild(hTreeView, hItem)
               DO
                  IF hSubItem = NULL THEN EXIT DO
                  IF TreeView_GetItemText(hTreeView, hSubItem, @wszSubItemText, MAX_PATH) THEN
                     dwsCode += "' " & wszSubItemText & CHR(13, 10)
                     IF LEFT(wszSubItemText, 22) = "Inherited interface = " THEN
                        wszInheritedInterface = MID(wszSubItemText, 23)
'                        IF wszInheritedInterface = "IUnknown" OR wszInheritedInterface = "IDispatch" THEN wszInheritedInterface = "Afx_" & wszInheritedInterface
                     END IF
                     hMethodsParentNode = NULL
                     IF LEFT(wszSubItemText, 17) = "Number of methods" AND VAL(MID(wszSubItemText, 21)) > 0 THEN
                        hMethodsParentNode = hSubItem
                     END IF
                  END IF
                  hSubItem = TreeView_GetNextSibling(hTreeView, hSubItem)
               LOOP
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += CHR(13, 10)
               IF m_vTableView THEN
                  dwsCode += "#ifndef __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               ELSE
                  dwsCode += "#ifndef __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               END IF
               dwsCode += CHR(13, 10)
               IF hMethodsParentNode THEN
                  DIM wszText AS WSTRING * MAX_PATH
                  IF m_VTableView THEN
                     IF wszInheritedInterface = "IUnknown" THEN
                        wszText = "_ EXTENDS Afx_IUnknown"
                     ELSEIF wszInheritedInterface = "IDispatch" THEN
                        wszText = "_ EXTENDS Afx_IDispatch"
                     ELSE
                        wszText = "_ EXTENDS Afx_" & wszInheritedInterface
                     END IF
                     dwsCode += "TYPE Afx_" & wszItemText & wszText & CHR(13, 10)
                  ELSE
                     dwsCode += "TYPE " & wszItemText & "VTbl" & CHR(13, 10)
                  END IF
                  IF m_VTableView = FALSE THEN
                     dwsCode += "   QueryInterface AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT" & CHR(13, 10)
                     dwsCode += "   AddRef AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                     dwsCode += "   Release AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                     dwsMacros + = "#define " & wszItemText & "_QueryInterface(this, riid, ppvObject) (this)->lpVtbl->QueryInterface(this, riid, ppvObject)"
                     dwsMacros + = "#define " & wszItemText & "_AddRef(this) (this)->lpVtbl->AddRef(this)"
                     dwsMacros + = "#define " & wszItemText & "_Release(this) (this)->lpVtbl->Release(this)"
                  END IF
                  ' // Search the methods and properties
                  DIM hMethods AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsParentNode)
                  DIM wszMethodName AS WSTRING * MAX_PATH
                  DO
                     IF hMethods = NULL THEN EXIT DO
                     IF TreeView_GetItemText(hTreeView, hMethods, @wszMethodName, MAX_PATH) THEN
                        IF LEFT(wszMethodName, 6) = "METHOD" THEN wszMethodName = MID(wszMethodName, 8)
                        IF LEFT(wszMethodName, 12) = "PROPERTY GET" THEN wszMethodName = "get_" & MID(wszMethodName, 14)
                        IF LEFT(wszMethodName, 12) = "PROPERTY PUT" THEN wszMethodName = "put_" & MID(wszMethodName, 14)
                        IF LEFT(wszMethodName, 15) = "PROPERTY PUTREF" THEN wszMethodName = "putref_" & MID(wszMethodName, 17)
                        ' // Get the return type
                        DIM dwsParamNames AS DWSTRING, dwsParams AS DWSTRING
                        DIM wszReturnType AS WSTRING * MAX_PATH, wszMethodSubItem AS WSTRING * MAX_PATH
                        DIM hMethodsSubItems AS HTREEITEM = TreeView_GetChild(hTreeView, hMethods)
                        DO
                           IF hMethodsSubItems = NULL THEN EXIT DO
                           IF TreeView_GetItemText(hTreeView, hMethodsSubItems, @wszMethodSubItem, MAX_PATH) THEN
                              IF LEFT(wszMethodSubItem, 11) = "Return type" THEN
                                 wszReturnType = MID(wszMethodSubItem, 15)
                                 wszReturnType = TLB_VarTypeToKeyword(VAL(wszReturnType))
                              END IF
                              ' --------------------------------------------------------------------------
                              ' // Get the parameters
                              IF LEFT(wszMethodSubItem, 23) = "Number of parameters = " THEN
                                 DIM hParams AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsSubItems)
                                 DIM wszParamName AS WSTRING * MAX_PATH
                                 DO
                                    IF hParams = NULL THEN EXIT DO
                                    IF TreeView_GetItemText(hTreeView, hParams, @wszParamName, MAX_PATH) THEN
                                       dwsParamNames += wszParamName & ", "
                                       DIM hParam AS HTREEITEM = TreeView_GetChild(hTreeView, hParams)
                                       DIM wszParam AS WSTRING * MAX_PATH
                                       DIM wszTypeKind AS WSTRING * 260
                                       DO
                                          IF hParam = NULL THEN EXIT DO
                                          IF TreeView_GetItemText(hTreeView, hParam, @wszParam, MAX_PATH) THEN
'                                             IF LEFT(wszParam, 12) = "FB syntax = " THEN
'                                                dwsParams += MID(wszParam, 13) & ", "
'                                             END IF
                                             ' ------------------------------------------------------
                                             IF LEFT(wszParam, 11) = "TypeKind = " THEN
                                                wszTypeKind = MID(wszParam, 12)
                                             END IF
                                             IF LEFT(wszParam, 12) = "FB syntax = " THEN
                                                IF m_vTableView THEN
                                                   IF wszTypeKind = "TKIND_INTERFACE" OR wszTypeKind = "TKIND_DISPATCH" THEN
'                                                      dwsParams += "Afx_" & MID(wszParam, 13) & ", "
                                                      wszParam = DWStrReplace(wszParam, " AS ", " AS Afx_")
                                                      dwsParams += MID(wszParam, 13) & ", "
                                                   ELSE
                                                      dwsParams += MID(wszParam, 13) & ", "
                                                   END IF
                                                ELSE
                                                   dwsParams += MID(wszParam, 13) & ", "
                                                END IF
                                             END IF
                                             ' ------------------------------------------------------
                                          END IF
                                          hParam = TreeView_GetNextSibling(hTreeView, hParam)
                                       LOOP
                                    END IF
                                    hParams = TreeView_GetNextSibling(hTreeView, hParams)
                                 LOOP
                              END IF
                              ' --------------------------------------------------------------------------
                           END IF
                           hMethodsSubItems = TreeView_GetNextSibling(hTreeView, hMethodsSubItems)
                        LOOP
                        IF wszReturnType = "VOID" THEN wszReturnType = "HRESULT"
                        IF RIGHT(dwsParams, 2) = ", " THEN dwsParams = LEFT(dwsParams, LEN(dwsParams) - 2)
                        IF RIGHT(dwsParamNames, 2) = ", " THEN dwsParamNames = LEFT(dwsParamNames, LEN(dwsParamNames) - 2)
                        IF LEN(dwsParamNames) THEN dwsParamNames = ", " & dwsParamNames
                        IF m_VTableView THEN
                           dwsCode += "   DECLARE ABSTRACT FUNCTION " & wszMethodName & " (" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                        ELSE
                           IF LEN(dwsParams) THEN dwsParams = ", " & dwsParams
                           dwsCode += "   " & wszMethodName & " AS FUNCTION (BYVAL this AS " & wszItemText & " PTR" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                           dwsMacros += "#define " & wszItemText & "_" & wszMethodName & "(this" & dwsParamNames &") (this)->lpVtbl->" & wszMethodName & "(this" & dwsParamNames & ")" & CHR(13, 10)
                        END IF
                     END IF
                     hMethods = TreeView_GetNextSibling(hTreeView, hMethods)
                  LOOP
                  dwsCode += "END TYPE" & CHR(13, 10)
               END IF
'            END IF
         END IF
         ' // ADO contains many interfaces that end with "_Deprecated"
'         IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
            IF hMethodsParentNode THEN
               dwsCode += CHR(13, 10)
               IF m_VTableView = FALSE THEN
                  dwsCode += "TYPE " & wszItemText & "_" & CHR(13, 10)
                  dwsCode += "   lpVtbl AS " & wszItemText & "VTbl PTR" & CHR(13, 10)
                  dwsCode += "END TYPE" & CHR(13, 10) & CHR(13, 10)
                  dwsCode += "TYPE " & "LP" & UCASE(wszItemText) & " AS " & wszItemText & " PTR" & CHR(13, 10)
                  dwsCode += CHR(13, 10)
               END IF
            END IF
            dwsCode += "#endif" & CHR(13, 10) & CHR(13, 10)
            IF LEN(dwsMacros) THEN
               dwsCode += "' // C Object Macros" & CHR(13, 10)
               dwsCode += dwsMacros & CHR(13, 10)
               dwsMacros.Clear
            END IF
            dwsCode += "' ########################################################################################" & CHR(13, 10)
            dwsCode += CHR(13, 10)
'         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Dual interfaces
   IF m_hDualIntNode THEN
      dwsMacros.Clear
      dwsCode += "' // Dual interfaces" & CHR(13, 10) & CHR(13, 10)
      DIM hMethodsParentNode AS HTREEITEM
      DIM wszItemText AS WSTRING * MAX_PATH, wszInheritedInterface AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hDualIntNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            ' // ADO contains many interfaces that end with "_Deprecated"
            IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += "' Interface name: " & wszItemText & CHR(13, 10)
               DIM wszSubItemText AS WSTRING * MAX_PATH
               DIM hSubItem AS HTREEITEM = TreeView_GetChild(hTreeView, hItem)
               DO
                  IF hSubItem = NULL THEN EXIT DO
                  IF TreeView_GetItemText(hTreeView, hSubItem, @wszSubItemText, MAX_PATH) THEN
                     dwsCode += "' " & wszSubItemText & CHR(13, 10)
                     IF LEFT(wszSubItemText, 22) = "Inherited interface = " THEN
                        wszInheritedInterface = MID(wszSubItemText, 23)
'                        IF wszInheritedInterface = "IUnknown" OR wszInheritedInterface = "IDispatch" THEN wszInheritedInterface = "Afx_" & wszInheritedInterface
                     END IF
                     hMethodsParentNode = NULL
                     IF LEFT(wszSubItemText, 17) = "Number of methods" AND VAL(MID(wszSubItemText, 21)) > 0 THEN
                        hMethodsParentNode = hSubItem
                     END IF
                  END IF
                  hSubItem = TreeView_GetNextSibling(hTreeView, hSubItem)
               LOOP
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += CHR(13, 10)
               IF m_vTableView THEN
                  dwsCode += "#ifndef __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               ELSE
                  dwsCode += "#ifndef __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               END IF
               dwsCode += CHR(13, 10)
               IF hMethodsParentNode THEN
                  DIM wszText AS WSTRING * MAX_PATH
                  IF m_VTableView THEN
                     IF wszInheritedInterface = "IUnknown" THEN
                        wszText = "_ EXTENDS Afx_IUnknown"
                     ELSEIF wszInheritedInterface = "IDispatch" THEN
                        wszText = "_ EXTENDS Afx_IDispatch"
                     ELSE
                        wszText = "_ EXTENDS Afx_" & wszInheritedInterface
                     END IF
                     dwsCode += "TYPE Afx_" & wszItemText & wszText & CHR(13, 10)
                  ELSE
                     dwsCode += "TYPE " & wszItemText & "VTbl" & CHR(13, 10)
                  END IF
                  ' // Search the methods and properties
                  DIM hMethods AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsParentNode)
                  DIM wszOrigMethodName AS WSTRING * MAX_PATH, wszMethodName AS WSTRING * MAX_PATH
                  DO
                     IF hMethods = NULL THEN EXIT DO
                     IF TreeView_GetItemText(hTreeView, hMethods, @wszOrigMethodName, MAX_PATH) THEN
                        IF LEFT(wszOrigMethodName, 6) = "METHOD" THEN wszMethodName = MID(wszOrigMethodName, 8)
                        IF LEFT(wszOrigMethodName, 12) = "PROPERTY GET" THEN wszMethodName = "get_" & MID(wszOrigMethodName, 14)
                        IF LEFT(wszOrigMethodName, 12) = "PROPERTY PUT" THEN wszMethodName = "put_" & MID(wszOrigMethodName, 14)
                        IF LEFT(wszOrigMethodName, 15) = "PROPERTY PUTREF" THEN wszMethodName = "putref_" & MID(wszOrigMethodName, 17)
                        IF m_VTableView THEN
                           ' --------------------------------------------------------------------------------
                           DIM dwsParams AS DWSTRING
                           ' // Get the return type
                           DIM wszReturnType AS WSTRING * MAX_PATH, wszMethodSubItem AS WSTRING * MAX_PATH
                           wszReturnType = "" : wszMethodSubItem = ""
                           DIM hMethodsSubItems AS HTREEITEM = TreeView_GetChild(hTreeView, hMethods)
                           DO
                              IF hMethodsSubItems = NULL THEN EXIT DO
                              IF TreeView_GetItemText(hTreeView, hMethodsSubItems, @wszMethodSubItem, MAX_PATH) THEN
                                 IF LEFT(wszMethodSubItem, 14) = "Return type = " THEN
                                    wszReturnType = MID(wszMethodSubItem, 15)
                                    IF wszReturnType = "VT_VOID" THEN
                                       wszReturnType = "HRESULT"
                                    ELSE
                                       wszReturnType = TLB_VarTypeToKeyword(wszReturnType)
                                    END IF
                                 END IF
                                 ' --------------------------------------------------------------------------
                                 ' // Get the parameters
                                 IF LEFT(wszMethodSubItem, 23) = "Number of parameters = " THEN
                                    DIM hParams AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsSubItems)
                                    DIM wszParamName AS WSTRING * MAX_PATH
                                    DO
                                       IF hParams = NULL THEN EXIT DO
                                       IF TreeView_GetItemText(hTreeView, hParams, @wszParamName, MAX_PATH) THEN
                                          DIM hParam AS HTREEITEM = TreeView_GetChild(hTreeView, hParams)
                                          DIM wszParam AS WSTRING * MAX_PATH
                                       DIM wszTypeKind AS WSTRING * 260
                                          DO
                                             IF hParam = NULL THEN EXIT DO
                                             IF TreeView_GetItemText(hTreeView, hParam, @wszParam, MAX_PATH) THEN
'                                                IF LEFT(wszParam, 12) = "FB syntax = " THEN
'                                                   dwsParams += MID(wszParam, 13) & ", "
'                                                END IF
                                                ' ------------------------------------------------------
                                                IF LEFT(wszParam, 11) = "TypeKind = " THEN
                                                   wszTypeKind = MID(wszParam, 12)
                                                END IF
                                                IF LEFT(wszParam, 12) = "FB syntax = " THEN
                                                   IF m_vTableView THEN
                                                      IF wszTypeKind = "TKIND_INTERFACE" OR wszTypeKind = "TKIND_DISPATCH" THEN
   '                                                      dwsParams += "Afx_" & MID(wszParam, 13) & ", "
                                                         wszParam = DWStrReplace(wszParam, " AS ", " AS Afx_")
                                                         dwsParams += MID(wszParam, 13) & ", "
                                                      ELSE
                                                         dwsParams += MID(wszParam, 13) & ", "
                                                      END IF
                                                   ELSE
                                                      dwsParams += MID(wszParam, 13) & ", "
                                                   END IF
                                                END IF
                                                ' ------------------------------------------------------
                                             END IF
                                             hParam = TreeView_GetNextSibling(hTreeView, hParam)
                                          LOOP
                                       END IF
                                       hParams = TreeView_GetNextSibling(hTreeView, hParams)
                                    LOOP
                                 END IF
                                 ' --------------------------------------------------------------------------
                              END IF
                              hMethodsSubItems = TreeView_GetNextSibling(hTreeView, hMethodsSubItems)
                           LOOP
                           IF RIGHT(dwsParams, 2) = ", " THEN dwsParams = LEFT(dwsParams, LEN(dwsParams) - 2)
                           dwsCode += "   DECLARE ABSTRACT FUNCTION " & wszMethodName & " (" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                           ' --------------------------------------------------------------------------------
                        ELSE
                           DIM dwsParamNames AS DWSTRING, dwsParams AS DWSTRING
                           DIM AS WSTRING * MAX_PATH wszReturnType, wszMethodSubItem, wszRhsParam
                           IF wszMethodName = "QueryInterface" THEN
                              dwsCode += "   QueryInterface AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT" & CHR(13, 10)
                              dwsMacros + = "#define " & wszItemText & "_QueryInterface(this, riid, ppvObject) (this)->lpVtbl->QueryInterface(this, riid, ppvObject)" & CHR(13, 10)
                           ELSEIF wszMethodName = "AddRef" THEN
                              dwsCode += "   AddRef AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                              dwsMacros + = "#define " & wszItemText & "_AddRef(this) (this)->lpVtbl->AddRef(this)" & CHR(13, 10)
                           ELSEIF wszMethodName = "Release" THEN
                              dwsCode += "   Release AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                              dwsMacros + = "#define " & wszItemText & "_Release(this) (this)->lpVtbl->Release(this)" & CHR(13, 10)
                           ELSEIF wszMethodName = "GetTypeInfoCount" THEN
                              dwsCode += "   GetTypeInfoCount AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT" & CHR(13, 10)
                              dwsMacros + = "#define " & wszItemText & "_GetTypeInfoCount(this, pctinfo) (this)->lpVtbl->GetTypeInfoCount(this, pctinfo)" & CHR(13, 10)
                           ELSEIF wszMethodName = "GetTypeInfo" THEN
                              dwsCode += "   GetTypeInfo AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT" & CHR(13, 10)
                              dwsMacros + = "#define " & wszItemText & "_GetTypeInfo(this, iTInfo, lcid, ppTInfo) (this)->lpVtbl->GetTypeInfo(this, iTInfo, lcid, ppTInfo)" & CHR(13, 10)
                           ELSEIF wszMethodName = "GetIDsOfNames" THEN
                              dwsCode += "   GetIDsOfNames AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT" & CHR(13, 10)
                              dwsMacros + = "#define " & wszItemText & "_GetIDsOfNames(this, riid, rgszNames, cNames, lcid, rgDispId) (this)->lpVtbl->GetIDsOfNames(this, riid, rgszNames, cNames, lcid, rgDispId)" & CHR(13, 10)
                           ELSEIF wszMethodName = "Invoke" THEN
                              dwsCode += "   Invoke AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT" & CHR(13, 10)
                              dwsMacros + = "#define " & wszItemText & "_Invoke(this, dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr) (this)->lpVtbl->Invoke(this, dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr)" & CHR(13, 10)
                           ELSE
                              ' -----------------------------------------------------------------------------
                              ' // Get the return type
                              DIM hMethodsSubItems AS HTREEITEM = TreeView_GetChild(hTreeView, hMethods)
                              DO
                                 IF hMethodsSubItems = NULL THEN EXIT DO
                                 IF TreeView_GetItemText(hTreeView, hMethodsSubItems, @wszMethodSubItem, MAX_PATH) THEN
                                    IF LEFT(wszMethodSubItem, 14) = "Return type = " THEN
                                       wszReturnType = MID(wszMethodSubItem, 15)
                                       IF LEFT(wszOrigMethodName, 12) <> "PROPERTY GET" THEN
                                          IF wszReturnType = "VT_VOID" THEN
                                             wszReturnType = "HRESULT"
                                          ELSE
                                             wszReturnType = TLB_VarTypeToKeyword(wszReturnType)
                                          END IF
                                       END IF
                                    END IF
                                    IF LEFT(wszMethodSubItem, 24) = "Return type FB syntax = " THEN
                                       wszRhsParam = MID(wszMethodSubItem, 25)
                                       wszReturnType = "HRESULT"
                                    END IF
                                    ' --------------------------------------------------------------------------
                                    ' // Get the parameters
                                    IF LEFT(wszMethodSubItem, 23) = "Number of parameters = " THEN
                                       dwsParamNames = ""
                                       DIM hParams AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsSubItems)
                                       DIM wszParamName AS WSTRING * MAX_PATH
                                       DO
                                          IF hParams = NULL THEN EXIT DO
                                          IF TreeView_GetItemText(hTreeView, hParams, @wszParamName, MAX_PATH) THEN
                                             dwsParamNames += wszParamName & ", "
                                             DIM hParam AS HTREEITEM = TreeView_GetChild(hTreeView, hParams)
                                             DIM wszParam AS WSTRING * MAX_PATH
                                             DIM wszTypeKind AS WSTRING * 260
                                             DO
                                                IF hParam = NULL THEN EXIT DO
                                                IF TreeView_GetItemText(hTreeView, hParam, @wszParam, MAX_PATH) THEN
'                                                   IF LEFT(wszParam, 12) = "FB syntax = " THEN
'                                                      dwsParams += MID(wszParam, 13) & ", "
'                                                   END IF
                                                   ' ------------------------------------------------------
                                                   IF LEFT(wszParam, 11) = "TypeKind = " THEN
                                                      wszTypeKind = MID(wszParam, 12)
                                                   END IF
                                                   IF LEFT(wszParam, 12) = "FB syntax = " THEN
                                                      IF m_vTableView THEN
                                                         IF wszTypeKind = "TKIND_INTERFACE" OR wszTypeKind = "TKIND_DISPATCH" THEN
      '                                                      dwsParams += "Afx_" & MID(wszParam, 13) & ", "
                                                            wszParam = DWStrReplace(wszParam, " AS ", " AS Afx_")
                                                            dwsParams += MID(wszParam, 13) & ", "
                                                         ELSE
                                                            dwsParams += MID(wszParam, 13) & ", "
                                                         END IF
                                                      ELSE
                                                         dwsParams += MID(wszParam, 13) & ", "
                                                      END IF
                                                   END IF
                                                   ' ------------------------------------------------------
                                                END IF
                                                hParam = TreeView_GetNextSibling(hTreeView, hParam)
                                             LOOP
                                          END IF
                                          hParams = TreeView_GetNextSibling(hTreeView, hParams)
                                       LOOP
                                    END IF
                                    ' --------------------------------------------------------------------------
                                 END IF
                                 hMethodsSubItems = TreeView_GetNextSibling(hTreeView, hMethodsSubItems)
                              LOOP
                              dwsParams += wszRhsParam
                              IF LEN(wszRhsParam) THEN dwsParamNames += "rhs"
                              ' -----------------------------------------------------------------------------
                              IF RIGHT(dwsParams, 2) = ", " THEN dwsParams = LEFT(dwsParams, LEN(dwsParams) - 2)
                              IF RIGHT(dwsParamNames, 2) = ", " THEN dwsParamNames = LEFT(dwsParamNames, LEN(dwsParamNames) - 2)
                              IF LEN(dwsParams) THEN dwsParams = ", " & dwsParams
                              IF LEN(dwsParamNames) THEN dwsParamNames = ", " & dwsParamNames
                              dwsCode += "   " & wszMethodName & " AS FUNCTION (BYVAL this AS " & wszItemText & " PTR" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                              dwsMacros += "#define " & wszItemText & "_" & wszMethodName & "(this" & dwsParamNames &") (this)->lpVtbl->" & wszMethodName & "(this" & dwsParamNames & ")" & CHR(13, 10)
                           END IF
                        END IF
                     END IF
                     hMethods = TreeView_GetNextSibling(hTreeView, hMethods)
                  LOOP
                  dwsCode += "END TYPE" & CHR(13, 10)
               END IF
            END IF
         END IF
         ' // ADO contains many interfaces that end with "_Deprecated"
         IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
            IF hMethodsParentNode THEN
               dwsCode += CHR(13, 10)
               IF m_VTableView = FALSE THEN
                  dwsCode += "TYPE " & wszItemText & "_" & CHR(13, 10)
                  dwsCode += "   lpVtbl AS " & wszItemText & "VTbl PTR" & CHR(13, 10)
                  dwsCode += "END TYPE" & CHR(13, 10) & CHR(13, 10)
                  dwsCode += "TYPE " & "LP" & UCASE(wszItemText) & " AS " & wszItemText & " PTR" & CHR(13, 10)
                  dwsCode += CHR(13, 10)
               END IF
            END IF
            dwsCode += "#endif" & CHR(13, 10) & CHR(13, 10)
            IF LEN(dwsMacros) THEN
               dwsCode += "' // C Object Macros" & CHR(13, 10)
               dwsCode += dwsMacros & CHR(13, 10)
               dwsMacros.Clear
            END IF
            dwsCode += "' ########################################################################################" & CHR(13, 10)
            dwsCode += CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
      dwsCode += CHR(13, 10)
   END IF

   ' // Event Interfaces
   IF m_hEventsNode THEN
      dwsMacros.Clear
      dwsCode += "' // Event interfaces" & CHR(13, 10) & CHR(13, 10)
      DIM hMethodsParentNode AS HTREEITEM
      DIM wszItemText AS WSTRING * MAX_PATH, wszInheritedInterface AS WSTRING * MAX_PATH
      DIM hItem AS HTREEITEM = TreeView_GetChild(hTreeView, m_hEventsNode)
      DO
         IF hItem = NULL THEN EXIT DO
         IF TreeView_GetItemText(hTreeView, hItem, @wszItemText, MAX_PATH) THEN
            ' // ADO contains many interfaces that end with "_Deprecated"
            IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += "' Interface name: " & wszItemText & CHR(13, 10)
               DIM wszSubItemText AS WSTRING * MAX_PATH
               DIM hSubItem AS HTREEITEM = TreeView_GetChild(hTreeView, hItem)
               DO
                  IF hSubItem = NULL THEN EXIT DO
                  IF TreeView_GetItemText(hTreeView, hSubItem, @wszSubItemText, MAX_PATH) THEN
                     dwsCode += "' " & wszSubItemText & CHR(13, 10)
                     IF LEFT(wszSubItemText, 22) = "Inherited interface = " THEN
                        wszInheritedInterface = MID(wszSubItemText, 23)
'                        IF wszInheritedInterface = "IUnknown" OR wszInheritedInterface = "IDispatch" THEN wszInheritedInterface = "Afx_" & wszInheritedInterface
                     END IF
                     hMethodsParentNode = NULL
                     IF LEFT(wszSubItemText, 17) = "Number of methods" AND VAL(MID(wszSubItemText, 21)) > 0 THEN
                        hMethodsParentNode = hSubItem
                     END IF
                  END IF
                  hSubItem = TreeView_GetNextSibling(hTreeView, hSubItem)
               LOOP
               dwsCode += "' ########################################################################################" & CHR(13, 10)
               dwsCode += CHR(13, 10)
               IF m_vTableView THEN
                  dwsCode += "#ifndef __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __Afx_" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               ELSE
                  dwsCode += "#ifndef __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
                  dwsCode += "#define __" & wszItemText & "_INTERFACE_DEFINED__" & CHR(13, 10)
               END IF
               dwsCode += CHR(13, 10)
               IF hMethodsParentNode THEN
                  DIM wszText AS WSTRING * MAX_PATH
                  IF m_VTableView THEN
                     IF wszInheritedInterface = "IUnknown" THEN
                        wszText = "_ EXTENDS Afx_IUnknown"
                     ELSEIF wszInheritedInterface = "IDispatch" THEN
                        wszText = "_ EXTENDS Afx_IDispatch"
                     ELSE
                        wszText = "_ EXTENDS Afx_" & wszInheritedInterface
                     END IF
                     dwsCode += "TYPE Afx_" & wszItemText & wszText & CHR(13, 10)
                  ELSE
                     dwsCode += "TYPE " & wszItemText & "VTbl" & CHR(13, 10)
                     dwsCode += "   QueryInterface AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL riid AS CONST IID CONST PTR, BYVAL ppvObj AS ANY PTR PTR) AS HRESULT" & CHR(13, 10)
                     dwsCode += "   AddRef AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                     dwsCode += "   Release AS FUNCTION (BYVAL this AS " & wszItemText & " PTR) AS ULONG" & CHR(13, 10)
                     IF wszInheritedInterface = "IDispatch" THEN
                        dwsCode += "   GetTypeInfoCount AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL pctinfo AS UINT PTR) AS HRESULT" & CHR(13, 10)
                        dwsCode += "   GetTypeInfo AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT" & CHR(13, 10)
                        dwsCode += "   GetIDsOfNames AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL iTInfo AS UINT, BYVAL lcid AS LCID, BYVAL ppTInfo AS ITypeInfo PTR PTR) AS HRESULT" & CHR(13, 10)
                        dwsCode += "   Invoke AS FUNCTION (BYVAL this AS " & wszItemText & " PTR, BYVAL dispIdMember AS DISPID, BYVAL riid AS CONST IID CONST PTR, BYVAL lcid AS LCID, BYVAL wFlags AS WORD, BYVAL pDispParams AS DISPPARAMS PTR, BYVAL pVarResult AS VARIANT PTR, BYVAL pExcepInfo AS EXCEPINFO PTR, BYVAL puArgErr AS UINT PTR) AS HRESULT" & CHR(13, 10)
                     END IF                           
                  END IF
                  ' // Search the methods and properties
                  DIM hMethods AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsParentNode)
                  DIM wszOrigMethodName AS WSTRING * MAX_PATH, wszMethodName AS WSTRING * MAX_PATH
                  DO
                     IF hMethods = NULL THEN EXIT DO
                     IF TreeView_GetItemText(hTreeView, hMethods, @wszOrigMethodName, MAX_PATH) THEN
                        IF LEFT(wszOrigMethodName, 6) = "METHOD" THEN wszMethodName = MID(wszOrigMethodName, 8)
                        IF LEFT(wszOrigMethodName, 12) = "PROPERTY GET" THEN wszMethodName = "get_" & MID(wszOrigMethodName, 14)
                        IF LEFT(wszOrigMethodName, 12) = "PROPERTY PUT" THEN wszMethodName = "put_" & MID(wszOrigMethodName, 14)
                        IF LEFT(wszOrigMethodName, 15) = "PROPERTY PUTREF" THEN wszMethodName = "putref_" & MID(wszOrigMethodName, 17)
                        IF m_VTableView THEN
                           ' --------------------------------------------------------------------------------
                           DIM dwsParams AS DWSTRING
                           ' // Get the return type
                           DIM wszReturnType AS WSTRING * MAX_PATH, wszMethodSubItem AS WSTRING * MAX_PATH
                           wszReturnType = "" : wszMethodSubItem = ""
                           DIM hMethodsSubItems AS HTREEITEM = TreeView_GetChild(hTreeView, hMethods)
                           DO
                              IF hMethodsSubItems = NULL THEN EXIT DO
                              IF TreeView_GetItemText(hTreeView, hMethodsSubItems, @wszMethodSubItem, MAX_PATH) THEN
                                 IF LEFT(wszMethodSubItem, 14) = "Return type = " THEN
                                    wszReturnType = MID(wszMethodSubItem, 15)
                                    IF wszReturnType = "VT_VOID" THEN
                                       wszReturnType = "HRESULT"
                                    ELSE
                                       wszReturnType = TLB_VarTypeToKeyword(wszReturnType)
                                    END IF
                                 END IF
                                 ' --------------------------------------------------------------------------
                                 ' // Get the parameters
                                 IF LEFT(wszMethodSubItem, 23) = "Number of parameters = " THEN
                                    DIM hParams AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsSubItems)
                                    DIM wszParamName AS WSTRING * MAX_PATH
                                    DO
                                       IF hParams = NULL THEN EXIT DO
                                       IF TreeView_GetItemText(hTreeView, hParams, @wszParamName, MAX_PATH) THEN
                                          DIM hParam AS HTREEITEM = TreeView_GetChild(hTreeView, hParams)
                                          DIM wszParam AS WSTRING * MAX_PATH
                                       DIM wszTypeKind AS WSTRING * 260
                                          DO
                                             IF hParam = NULL THEN EXIT DO
                                             IF TreeView_GetItemText(hTreeView, hParam, @wszParam, MAX_PATH) THEN
'                                                IF LEFT(wszParam, 12) = "FB syntax = " THEN
'                                                   dwsParams += MID(wszParam, 13) & ", "
'                                                END IF
                                                ' ------------------------------------------------------
                                                IF LEFT(wszParam, 11) = "TypeKind = " THEN
                                                   wszTypeKind = MID(wszParam, 12)
                                                END IF
                                                IF LEFT(wszParam, 12) = "FB syntax = " THEN
                                                   IF m_vTableView THEN
                                                      IF wszTypeKind = "TKIND_INTERFACE" OR wszTypeKind = "TKIND_DISPATCH" THEN
   '                                                      dwsParams += "Afx_" & MID(wszParam, 13) & ", "
                                                         wszParam = DWStrReplace(wszParam, " AS ", " AS Afx_")
                                                         dwsParams += MID(wszParam, 13) & ", "
                                                      ELSE
                                                         dwsParams += MID(wszParam, 13) & ", "
                                                      END IF
                                                   ELSE
                                                      dwsParams += MID(wszParam, 13) & ", "
                                                   END IF
                                                END IF
                                                ' ------------------------------------------------------
                                             END IF
                                             hParam = TreeView_GetNextSibling(hTreeView, hParam)
                                          LOOP
                                       END IF
                                       hParams = TreeView_GetNextSibling(hTreeView, hParams)
                                    LOOP
                                 END IF
                                 ' --------------------------------------------------------------------------
                              END IF
                              hMethodsSubItems = TreeView_GetNextSibling(hTreeView, hMethodsSubItems)
                           LOOP
                           IF RIGHT(dwsParams, 2) = ", " THEN dwsParams = LEFT(dwsParams, LEN(dwsParams) - 2)
                           dwsCode += "   DECLARE ABSTRACT FUNCTION " & wszMethodName & " (" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                           ' --------------------------------------------------------------------------------
                        ELSE
                           DIM dwsParamNames AS DWSTRING, dwsParams AS DWSTRING
                           DIM AS WSTRING * MAX_PATH wszReturnType, wszMethodSubItem, wszRhsParam
                           ' -----------------------------------------------------------------------------
                           ' // Get the return type
                           DIM hMethodsSubItems AS HTREEITEM = TreeView_GetChild(hTreeView, hMethods)
                           DO
                              IF hMethodsSubItems = NULL THEN EXIT DO
                              IF TreeView_GetItemText(hTreeView, hMethodsSubItems, @wszMethodSubItem, MAX_PATH) THEN
                                 IF LEFT(wszMethodSubItem, 14) = "Return type = " THEN
                                    wszReturnType = MID(wszMethodSubItem, 15)
                                    IF LEFT(wszOrigMethodName, 12) <> "PROPERTY GET" THEN
                                       IF wszReturnType = "VT_VOID" THEN
                                          wszReturnType = "HRESULT"
                                       ELSE
                                          wszReturnType = TLB_VarTypeToKeyword(wszReturnType)
                                       END IF
                                    END IF
                                 END IF
                                 IF LEFT(wszMethodSubItem, 24) = "Return type FB syntax = " THEN
                                    wszRhsParam = MID(wszMethodSubItem, 25)
                                    wszReturnType = "HRESULT"
                                 END IF
                                 ' --------------------------------------------------------------------------
                                 ' // Get the parameters
                                 IF LEFT(wszMethodSubItem, 23) = "Number of parameters = " THEN
                                    dwsParamNames = ""
                                    DIM hParams AS HTREEITEM = TreeView_GetChild(hTreeView, hMethodsSubItems)
                                    DIM wszParamName AS WSTRING * MAX_PATH
                                    DO
                                       IF hParams = NULL THEN EXIT DO
                                       IF TreeView_GetItemText(hTreeView, hParams, @wszParamName, MAX_PATH) THEN
                                          dwsParamNames += wszParamName & ", "
                                          DIM hParam AS HTREEITEM = TreeView_GetChild(hTreeView, hParams)
                                          DIM wszParam AS WSTRING * MAX_PATH
                                          DIM wszTypeKind AS WSTRING * 260
                                          DO
                                             IF hParam = NULL THEN EXIT DO
                                             IF TreeView_GetItemText(hTreeView, hParam, @wszParam, MAX_PATH) THEN
'                                                IF LEFT(wszParam, 12) = "FB syntax = " THEN
'                                                   dwsParams += MID(wszParam, 13) & ", "
'                                                END IF
                                                ' ------------------------------------------------------
                                                IF LEFT(wszParam, 11) = "TypeKind = " THEN
                                                   wszTypeKind = MID(wszParam, 12)
                                                END IF
                                                IF LEFT(wszParam, 12) = "FB syntax = " THEN
                                                   IF m_vTableView THEN
                                                      IF wszTypeKind = "TKIND_INTERFACE" OR wszTypeKind = "TKIND_DISPATCH" THEN
   '                                                      dwsParams += "Afx_" & MID(wszParam, 13) & ", "
                                                         wszParam = DWStrReplace(wszParam, " AS ", " AS Afx_")
                                                         dwsParams += MID(wszParam, 13) & ", "
                                                      ELSE
                                                         dwsParams += MID(wszParam, 13) & ", "
                                                      END IF
                                                   ELSE
                                                      dwsParams += MID(wszParam, 13) & ", "
                                                   END IF
                                                END IF
                                                ' ------------------------------------------------------
                                             END IF
                                             hParam = TreeView_GetNextSibling(hTreeView, hParam)
                                          LOOP
                                       END IF
                                       hParams = TreeView_GetNextSibling(hTreeView, hParams)
                                    LOOP
                                 END IF
                                 ' --------------------------------------------------------------------------
                              END IF
                              hMethodsSubItems = TreeView_GetNextSibling(hTreeView, hMethodsSubItems)
                           LOOP
                           dwsParams += wszRhsParam
                           IF LEN(wszRhsParam) THEN dwsParamNames += "rhs"
                           ' -----------------------------------------------------------------------------
                           IF RIGHT(dwsParams, 2) = ", " THEN dwsParams = LEFT(dwsParams, LEN(dwsParams) - 2)
                           IF RIGHT(dwsParamNames, 2) = ", " THEN dwsParamNames = LEFT(dwsParamNames, LEN(dwsParamNames) - 2)
                           IF LEN(dwsParams) THEN dwsParams = ", " & dwsParams
                           IF LEN(dwsParamNames) THEN dwsParamNames = ", " & dwsParamNames
                           dwsCode += "   " & wszMethodName & " AS FUNCTION (BYVAL this AS " & wszItemText & " PTR" & dwsParams & ") AS " & wszReturnType & CHR(13, 10)
                        END IF
                     END IF
                     hMethods = TreeView_GetNextSibling(hTreeView, hMethods)
                  LOOP
                  dwsCode += "END TYPE" & CHR(13, 10)
               END IF
            END IF
         END IF
         ' // ADO contains many interfaces that end with "_Deprecated"
         IF RIGHT(wszItemText, 11) <> "_Deprecated" THEN
            IF hMethodsParentNode THEN
               dwsCode += CHR(13, 10)
               IF m_VTableView = FALSE THEN
                  dwsCode += "TYPE " & wszItemText & "_" & CHR(13, 10)
                  dwsCode += "   lpVtbl AS " & wszItemText & "VTbl PTR" & CHR(13, 10)
                  dwsCode += "END TYPE" & CHR(13, 10) & CHR(13, 10)
                  dwsCode += "TYPE " & "LP" & UCASE(wszItemText) & " AS " & wszItemText & " PTR" & CHR(13, 10)
                  dwsCode += CHR(13, 10)
               END IF
            END IF
            dwsCode += "#endif" & CHR(13, 10) & CHR(13, 10)
            dwsCode += "' ########################################################################################" & CHR(13, 10)
            dwsCode += CHR(13, 10)
         END IF
         hItem = TreeView_GetNextSibling(hTreeView, hItem)
      LOOP
'      dwsCode += CHR(13, 10)
   END IF

'   IF m_vTableView THEN
'      dwsCode += "END NAMESPACE"
'   ELSE
'      IF LEN(m_Namespace) THEN dwsCode += "END NAMESPACE"
'   END IF

   IF m_pWindow THEN
      DIM hCodeView AS HWND = cast(HWND, m_pWindow->UserData(AFX_HCODEVIEW))
      IF hCodeView THEN
         ' // Set the text of the edit window
         Edit_SetText(hCodeView, dwsCode)
      END IF
   END IF

END SUB
' =====================================================================================
