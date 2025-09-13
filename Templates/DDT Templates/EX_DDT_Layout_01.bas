'#TEMPLATE DDT Dialog with colors and layout
#define UNICODE
#define _WIN32_WINNT &h0602
#include once "AfxNova/DDT.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION DlgProc(BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

' // Control identifiers
ENUM
   IDC_EDIT1 = 1001
   IDC_EDIT2
   IDC_LABEL
   IDC_GROUPBOX
   IDC_COMBOBOX
END ENUM

' ========================================================================================
' Main
' ========================================================================================
FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                   BYVAL hPrevInstance AS HINSTANCE, _
                   BYVAL pwszCmdLine AS WSTRING PTR, _
                   BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   SetProcessDpiAwareness(PROCESS_SYSTEM_DPI_AWARE)
   ' // Enable visual styles
   AfxEnableVisualStyles

   ' // Create a new dialog
   DIM hDlg AS HWND = DialogNewPixels(0, "DDT with FreeBasic - Colors and Layout Demo",,, 450, 180, WS_OVERLAPPEDWINDOW OR DS_CENTER)

   ' // Set the dialog's backgroung color
   DialogSetColor(hDlg, -1, RGB_GOLD)
   ' DialogDisableRepaintOnResize(hDlg, TRUE)    ' // optional: disable repaint on resizing

   ' // Add an Edit control
   ControlAddTextBox, hDlg, IDC_EDIT1, "", 15, 15, 305, 23
   ' // Add a multiline Edit control
   ControlAddTextbox, hDlg, IDC_EDIT2, "", 15, 45, 305, 80, WS_TABSTOP OR WS_VSCROLL OR ES_LEFT OR ES_AUTOHSCROLL OR ES_MULTILINE OR ES_NOHIDESEL OR ES_WANTRETURN
   ' // Add more controls
   ControlAddFrame, hDlg, IDC_GROUPBOX, "Groupbox", 335, 8, 100, 155
   ' // Disable themes for the GroupBox control; otherwise, the caption colors can't be changed
   SetWindowTheme(ControlHandle(hDlg, IDC_GROUPBOX), "", "")

   ControlAddButton, hDlg, IDCANCEL, "&Close", 245, 140, 75, 23
   ControlAddLabel, hDlg, IDC_LABEL, "Label", 15, 140, 75, 23
   ' // Add a combobox control
   ControlAddComboBox, hDlg, IDC_COMBOBOX, "", 345, 30, 80, 100
   ' // Fill the control with some data
   DIM wszText AS WSTRING * 260
   FOR i AS LONG = 1 TO 9
      wszText = "Item " & RIGHT("00" & STR(i), 2)
      ComboBoxAdd(hDlg, IDC_COMBOBOX, @wszText)
   NEXT

   ' // Set control's colors
   ControlSetColor(hDlg, IDC_EDIT1, RGB_WHITE, RGB_CORAL)
   ControlSetColor(hDlg, IDC_EDIT2, RGB_WHITE, RGB_LIGHTBLUE)
   ControlSetColor(hDlg, IDC_LABEL, RGB_YELLOW, RGB_GREEN)
   ControlSetColor(hDlg, IDC_GROUPBOX, RGB_WHITE, RGB_RED)

   ' // Anchor the controls
   ControlAnchor(hDlg, IDC_EDIT1, AFX_ANCHOR_WIDTH)
   ControlAnchor(hDlg, IDC_EDIT2, AFX_ANCHOR_HEIGHT_WIDTH)
   ControlAnchor(hDlg, IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)
   ControlAnchor(hDlg, IDC_LABEL, AFX_ANCHOR_BOTTOM)
   ControlAnchor(hDlg, IDC_GROUPBOX, AFX_ANCHOR_HEIGHT_RIGHT)
   ControlAnchor(hDlg, IDC_COMBOBOX, AFX_ANCHOR_RIGHT)

   ' // Display and activate the dialog as modal
   DialogShowModal hDlg, @DlgProc

   ' // Display and activate the dialog as modeless
'   DialogShowModeless hDlg, @DlgProc

   ' // Message handler loop
'   DO
'      DialogDoEvents
'   LOOP WHILE IsWin(hDlg)

   FUNCTION = DialogEndResult(hDlg)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Dialog callback procedure
' ========================================================================================
FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR
   
   SELECT CASE uMsg

      CASE WM_INITDIALOG
          RETURN TRUE

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  DialogSend hDlg, WM_CLOSE, 0, 0
               END IF
         END SELECT

      CASE WM_CLOSE
         ' // Multiline edit controls send a WM_CLOSE message when the Esc key is pressed.
         ' // This is a workaround; the proper way is to subclass the control and eat the ESC key.
         IF GetFocus = ControlHandle(hDlg, IDC_EDIT2) THEN RETURN TRUE   ' // abort
         ' // End the application
         DialogEnd(hDlg)
   END SELECT

END FUNCTION
