'#TEMPLATE DDT Dialog - Animation demo
#define UNICODE
#define _WIN32_WINNT &h0602
#include once "AfxNova/DDT.inc"
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
CONST IDC_ANIM = WM_USER + 2050

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles without including a manifest file
   AfxEnableVisualStyles

   ' // Create a new dialog
   DIM hDlg AS HWND = DialogNew(0, "Animation Demo",,, 185, 60, WS_CAPTION OR WS_SYSMENU)
   ' // Add controls to it
   ControlAddButton, hDlg, IDC_START, "&Start", 6, 40, 55, 14
   ControlAddButton, hDlg, IDC_STOP, "Sto&p", 65, 40, 55, 14, WS_DISABLED OR WS_TABSTOP
'   ControlAdd "SysAnimate32", hDlg, IDC_ANIM, "", 6, 6, 172, 24, ACS_CENTER OR ACS_TRANSPARENT
   ControlAddAnimate, hDlg, IDC_ANIM, "", 6, 6, 172, 24, ACS_CENTER OR ACS_TRANSPARENT
   ControlAddButton, hDlg, IDCANCEL, "E&xit", 124, 40, 55, 14

   ' // Display and activate the dialog as modal
   DialogShowModal hDlg, @DlgProc

   RETURN DialogEndResult(hDlg)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Dialog callback procedure
' ========================================================================================
FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR
   
   STATIC hAnimate AS HWND
   
   SELECT CASE uMsg

      CASE WM_INITDIALOG
         hAnimate = ControlHandle(hDlg, IDC_ANIM)
         DIM wszName AS WSTRING * 260 = ExePath & "\Resources\FILECOPY.AVI"
         Animate_Open(hAnimate, @wszName)
         RETURN TRUE

      CASE WM_COMMAND

         ' Trap only click events
         IF CBCTLMSG(wParam, lParam) <> BN_CLICKED THEN EXIT SELECT

         SELECT CASE CBCTL(wParam, lParam)

            CASE IDC_START
               ControlDisable hDlg, IDC_START
               ControlEnable hDlg, IDC_STOP
               ControlSetFocus hDlg, IDC_STOP
               Animate_Play(hAnimate, 0, -1, -1)

            CASE IDC_STOP
               ControlDisable hDlg, IDC_STOP
               ControlEnable hDlg, IDC_START
               ControlSetFocus hDlg, IDC_START
               Animate_Stop(hAnimate)

            CASE IDCANCEL
               DialogEnd hDlg

         END SELECT

   END SELECT

   RETURN FALSE

END FUNCTION
