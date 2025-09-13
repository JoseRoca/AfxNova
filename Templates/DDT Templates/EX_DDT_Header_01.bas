'#TEMPLATE DDT Dialog with an Header control
#DEFINE UNICODE
#define _WIN32_WINNT &h0602
#include once "AfxNova/DDT.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

DECLARE FUNCTION DlgProc (BYVAL hDlg AS HWND, BYVAL uMsg AS DWORD, BYVAL wParam AS DWORD, BYVAL lParam AS LPARAM) AS INT_PTR

CONST IDC_HEADER = 1001

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

   ' // Create a new dialog using dialog units
   DIM hDlg AS HWND = DialogNew(0, "DDT Dialog with an Header control", , , 210, 100, WS_OVERLAPPEDWINDOW OR DS_CENTER, , "Segoe UI", 9)

   ' // Add the header control
   ControlAddHeader, hDlg, IDC_HEADER
   DIM hHeader AS HWND = ControlHandle(hDlg, IDC_HEADER)

   ' // Insert items
   DIM thdi AS HDITEMW
   DIM wszItem AS WSTRING * 260
   wszItem   = "Item 1"
   thdi.Mask = HDI_WIDTH OR HDI_FORMAT OR HDI_TEXT
   thdi.fmt  = HDF_LEFT OR HDF_STRING
   thdi.cxy  = ScaleForDpiX(hDlg, 50)
   thdi.pszText  = @wszItem
   Header_InsertItem(hHeader, 1, @thdi)
   wszItem   = "Item 2"
   Header_InsertItem(hHeader, 2, @thdi)
   wszItem   = "Item 3"
   Header_InsertItem(hHeader, 3, @thdi)
   wszItem   = "Item 4"
   Header_InsertItem(hHeader, 4, @thdi)

   ' // Add a button
   ControlAddButton, hDlg, IDCANCEL, "&Close", 135, 75, 50, 12
   ' // Anchor the button
   ControlAnchor(hDlg, IDCANCEL, AFX_ANCHOR_BOTTOM_RIGHT)

   ' // Display and activate the dialog as modal
   DialogShowModal(hDlg, @DlgProc)

   RETURN DialogEndResult(hDlg)

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
                  SendMessageW hDlg, WM_CLOSE, 0, 0
               END IF
         END SELECT

      CASE WM_NOTIFY
         ' // Retrieve the header item clicked
         DIM hd AS NMHEADERW
         CBNMTYPESET(hd, wParam, lParam)
         IF hd.hdr.idFrom = IDC_HEADER THEN
            SELECT CASE hd.hdr.code
               CASE HDN_ITEMCLICKW
                  MsgBox hDlg, "You have clicked item " & WSTR(hd.iItem + 1), MB_OK, "Message"
            END SELECT
         END IF

      CASE WM_SIZE
         ' // Resize the header control
         DIM hHeader AS HWND = ControlHandle(hDlg, IDC_HEADER)
         DIM thdl AS HDLAYOUT, twp AS WINDOWPOS, trc AS RECT
         GetClientRect hDlg, @trc
         thdl.prc = @trc
         thdl.pwpos = @twp
         Header_Layout(hHeader, @thdl)
         SetWindowPos hHeader, NULL, twp.x, twp.y, twp.cx, twp.cy, SWP_NOZORDER OR SWP_NOACTIVATE
         RETURN TRUE

      CASE WM_CLOSE
         ' // End the application
         DialogEnd(hDlg)

   END SELECT

   RETURN FALSE

END FUNCTION
