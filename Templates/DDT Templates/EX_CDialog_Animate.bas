#define UNICODE
#define _WIN32_WINNT &h0602
#include once "windows.bi"
#include once "AfxNova/CDialog.inc"
#include once "AfxNova/AfxWin.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

' // Control identifiers
CONST IDC_START = WM_USER + 2048
CONST IDC_STOP = WM_USER + 2049
COnST IDC_ANIM = WM_USER + 2050

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // For testing purposes: The proper way is to use a manifest
   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles
   AfxEnableVisualStyles

   ' // Creeate an instance of the Dialog class.
'   DIM pDlg AS CDialog = CDialog("MS Sans Sherif", 8)   ' // original example
   ' // The default font is Segoe UI, 9 points.
   DIM pDlg AS CDialog
   pDlg.DialogNew(NULL, "CDialog - Animation Demo",,, 185, 60, WS_CAPTION OR WS_SYSMENU)
   ' // Add controls
   pDlg.ControlAdd("Button", IDC_START, "&Start", 6, 40, 55, 14)
   pDlg.ControlAdd("Button", IDC_STOP, "Sto&p", 65, 40, 55, 14, WS_DISABLED OR WS_TABSTOP)
   pDlg.ControlAdd("Button", IDCANCEL, "E&xit", 124, 40, 55, 14)
   pDlg.ControlAdd("SysAnimate32", IDC_ANIM, "", 6, 6, 172, 24, ACS_CENTER OR ACS_TRANSPARENT)

   ' // Display and activate the dialog as modal
   pDlg.DialogShowModal(@DlgProc)

   FUNCTION = pDlg.DialogEndResult

END FUNCTION
' ========================================================================================

' ========================================================================================
' Dialog callback procedure
' ========================================================================================
FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR
   
   ' // Pointer to the dialog class
   STATIC pDlg AS CDialog PTR
   STATIC hAnimate AS HWND
   
   SELECT CASE uMsg

      CASE WM_INITDIALOG
         ' // Get a pointer to the CDialog class
         pDlg = cast(CDialog PTR, lParam)
         hAnimate = pDlg->ControlHandle(IDC_ANIM)
         DIM wszName AS WSTRING * 260 = ExePath & "\Resources\FILECOPY.AVI"
         Animate_Open(hAnimate, @wszName)
         RETURN TRUE

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDC_START
               pDlg->ControlDisable IDC_START
               pDlg->ControlEnable IDC_STOP
               pDlg->ControlSetFocus IDC_STOP
               Animate_Play(hAnimate, 0, -1, -1)
            CASE IDC_STOP
               pDlg->ControlDisable IDC_STOP
               pDlg->ControlEnable IDC_START
               pDlg->ControlSetFocus IDC_START
               Animate_Stop(hAnimate)
            CASE IDCANCEL
               pDlg->DialogEnd(0)
         END SELECT

   END SELECT

   RETURN FALSE

END FUNCTION
