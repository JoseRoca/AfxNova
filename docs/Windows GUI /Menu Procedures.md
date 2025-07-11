# Menu Procedures

| Name       | Description |
| ---------- | ----------- |
| [AfxAddIconToMenuItem](#afxaddicontomenuitem) | Converts an icon handle to a bitmap and adds it to the specified *hbmpItem* field of HMENU item. |
| [AfxCheckMenuItem](#afxcheckmenuitem) | Checks a menu item. |
| [AfxDestroyMenu](#afxdestroymenu) | Destroys a menu. |
| [AfxDisableMenuItem](#afxdisablemenuitem) | Disables the specified menu item. |
| [AfxEnableMenuItem](#afxenablemenuitem) | Enables the specified menu item. |
| [AfxGetMenuFont](#afxgetmenufont) | Retrieves information about the font used in menu bars. |
| [AfxGetMenuFontPointSize](#afxgetmenufontpointsize) | Retrieves the point size of the font used in menu bars. |
| [AfxGetMenuItemState](#afxgetmenuitemstate) | Retrieves the state of the specified menu item. |
| [AfxGetMenuItemText](#afxgetmenuitemtext) | Retrieves the text of the specified menu item. |
| [AfxGetMenuItemTextLen](#afxgetmenuitemtextlen) | Retrieves length of the of the specified menu item. |
| [AfxGetMenuRect](#afxgetmenurect) | Calculates the size of a menu bar or a drop-down menu. |
| [AfxGrayMenuItem](#afxgraymenuitem) | Grays the specified menu item. |
| [AfxHiliteMenuItem](#afxhilitemenuitem) | Highlights the specified menu item. |
| [AfxIsMenuItemChecked](#afxismenuitemchecked) | Returns TRUE if the specified menu item is checked; FALSE otherwise. |
| [AfxIsMenuItemDisabled](#afxismenuitemdisabled) | Returns TRUE if the specified menu item is disabled; FALSE otherwise. |
| [AfxIsMenuItemEnabled](#afxismenuitemenabled) | Returns TRUE if the specified menu item is enabled; FALSE otherwise. |
| [AfxIsMenuItemGrayed](#afxismenuitemgrayed) | Returns TRUE if the specified menu item is grayed; FALSE otherwise. |
| [AfxIsMenuItemHighlighted](#afxismenuitemhighlighted) | Returns TRUE if the specified menu item is highlighted; FALSE otherwise. |
| [AfxIsMenuItemOwnerDraw](#afxismenuitemownerdraw) | Returns TRUE if the specified menu item is a ownerdraw; FALSE otherwise. |
| [AfxIsMenuItemPopup](#afxismenuitempopup) | Returns TRUE if the specified menu item is a submenu; FALSE otherwise. |
| [AfxIsMenuItemSeparator](#afxismenuitemseparator) | Returns TRUE if the specified menu item is a separator; FALSE otherwise. |
| [AfxRemoveCloseMenu](#Afxremoveclosemenu) | Removes the system menu close option and disables the X button. |
| [AfxRightJustifyMenuItem](#afxrightjustifymenuitem) | Right justifies a top level menu item. |
| [AfxSetMenuItemBold](#afxsetmenuitembold) | Changes the text of a menu item to bold. |
| [AfxSetMenuItemState](#afxsetmenuitemstate) | Sets the state of the specified menu item. |
| [AfxSetMenuItemText](#afxsetmenuitemtext) | Sets the text of the specified menu item. |
| [AfxToggleMenuItem](#afxtogglemenuitem) | Toggles the checked state of a menu item. |
| [AfxUnCheckMenuItem](#afxuncheckmenuitem) | Unchecks a menu item. |

---

## AfxAddIconToMenuItem

Converts an icon handle to a bitmap and adds it to the specified *hbmpItem field* of **HMENU** item.

```
FUNCTION AfxAddIconToMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS BOOLEAN, BYVAL hIcon AS HICON, BYVAL fAutoDestroy AS BOOLEAN = TRUE, _
   BYVAL phbmp AS HBITMAP PTR = NULL) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |
| *hIcon* | Handle of the icon to add to the menu. |
| *fAutoDestroy* | TRUE (the default) or FALSE. If TRUE, **AfxAddIconToMenuItem** destroys the icon before returning. |
| *phbmp* | Out. Location where the bitmap representation of the icon is stored. Can be NULL. |

#### Return value

TRUE or FALSE.

#### Remarks

The caller is responsible for destroying the bitmap generated. The icon will be destroyed if *fAutoDestroy* is set to true. The *hbmpItem* field of the menu item can be used to keep track of the bitmap by passing NULL to *phbmp*.

Using **AfxGdipAddIconFromFile** or **AfxGdipIconFromRes** to load the image from a file or resource and convert it to a icon you can use alphablended .png icons.

#### Usage example

Loading the icon from disc:

```
DIM hSubMenu AS HMENU = GetSubMenu(hMenu, 1)
DIM hIcon AS HICON = LoadImageW(NULL, ExePath & "\undo_32.ico", IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
IF hIcon THEN AfxAddIconToMenuItem(hSubMenu, 0, TRUE, hIcon)
```

Loading the icon from a resource file:

```
DIM hSubMenu AS HMENU = GetSubMenu(hMenu, 1)
AfxAddIconToMenuItem(hSubMenu, 0, TRUE, AfxGdipIconFromRes(hInstance, "IDI_UNDO_32"))
```
---

## AfxCheckMenuItem

Checks a menu item.

```
FUNCTION AfxCheckMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

The return value specifies the previous state of the menu item (either MF_CHECKED or MF_UNCHECKED). If the menu item does not exist, the return value is -1.

---

## AfxDestroyMenu

Destroys the specified menu and frees any memory that the menu occupies.
The second overloaded function destroys the menu attached to a window or dialog.

```
FUNCTION AfxDestroyMenu (BYVAL hMenu AS HMENU) AS BOOLEAN
FUNCTION AfxDestroyMenu (BYVAL hwnd AS HWND) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu to be destroyed. |

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window or dialog that owns the menu. |

#### Return value

A boolean true (-1) pr false (0).

---

## AfxDisableMenuItem

Disables the specified menu item.

```
FUNCTION AfxDisableMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

TRUE or FALSE. To get extended error information, use the **GetLastError** function.

#### Remarks

The application must call the **DrawMenuBar** function whenever a menu changes, whether or not the menu is in a displayed window.

---

## AfxEnableMenuItem

Enables the specified menu item.

```
FUNCTION AfxEnableMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

TRUE or FALSE. To get extended error information, use the **GetLastError** function.

#### Remarks

The application must call the **DrawMenuBar** function whenever a menu changes, whether or not the menu is in a displayed window.

---

## AfxGetMenuFont

Retrieves information about the font used in menu bars.

```
FUNCTION AfxGetMenuFont (BYVAL plfw AS LOGFONTW PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *plfw* | Pointer to a LOGFONTW structure.

#### Return value

TRUE or FALSE. To get extended error information, use the **GetLastError** function.

---

## AfxGetMenuFontPointSize

Retrieves the point size of the font used in menu bars.

```
FUNCTION AfxGetMenuFontPointSize () AS LONG
```

#### Return value

The point size of the font. If the function fails, the return value is 0.

---

## AfxGetMenuItemState

Retrieves the state of the specified menu item.

```
FUNCTION AfxGetMenuItemState (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS DWORD
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

0 on failure or one or more of the following values:
```
MFS_CHECKED   The item is checked
MFS_DEFAULT   The menu item is the default.
MFS_DISABLED  The item is disabled.
MFS_ENABLED   The item is enabled.
MFS_GRAYED    The item is grayed.
MFS_HILITE    The item is highlighted
MFS_UNCHECKED The item is unchecked.
MFS_UNHILITE  The item is not highlighted.
```
---

## AfxGetMenuItemText

Retrieves the text of the specified menu item.

```
FUNCTION AfxGetMenuItemText (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Usage example

```
DIM dwsText AS DWSTRING = AfxGetMenuItemText(hMenu, 1, TRUE)
```
---

## AfxGetMenuItemTextLen

Returns the lengnth of the specified menu item.

```
FUNCTION AfxGetMenuItemTextLen (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

---

## AfxGetMenuRect

Returns the dimensions of a menu bar or a drop-down menu.

```
FUNCTION AfxGetMenuRect (BYVAL hwnd AS HWND, BYVAL hmenu AS HMENU) AS RECT
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle of the window that owns the menu. |
| *hMenu* | Handle to the menu. |

---

## AfxGrayMenuItem

Grays the specified menu item.

```
FUNCTION AfxGrayMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

TRUE or FALSE. To get extended error information, use the **GetLastError** function.

#### Remarks

The application must call the **DrawMenuBar** function whenever a menu changes, whether or not the menu is in a displayed window.

---

## AfxHiliteMenuItem

Highlights the specified menu item.

```
FUNCTION AfxHiliteMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

TRUE or FALSE. To get extended error information, use the **GetLastError** function.

#### Remarks

The application must call the **DrawMenuBar** function whenever a menu changes, whether or not the menu is in a displayed window.

---

## AfxIsMenuItemChecked

Returns TRUE if the specified menu item is checked; FALSE otherwise.

```
FUNCTION AfxIsMenuItemChecked (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

---

## AfxIsMenuItemDisabled

Returns TRUE if the specified menu item is disabled; FALSE otherwise.

```
FUNCTION AfxIsMenuItemDisabled (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

---

## AfxIsMenuItemEnabled

Returns TRUE if the specified menu item is enabled; FALSE otherwise.

```
FUNCTION AfxIsMenuItemEnabled (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

---

## AfxIsMenuItemGrayed

Returns TRUE if the specified menu item is grayed; FALSE otherwise.

```
FUNCTION AfxIsMenuItemGrayed (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

---

## AfxIsMenuItemHighlighted

Returns TRUE if the specified menu item is highlighted; FALSE otherwise.

```
FUNCTION AfxIsMenuItemHighlighted (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

---

## AfxIsMenuItemOwnerDraw

Returns TRUE if the specified menu item is a ownerdraw; FALSE otherwise.

```
FUNCTION AfxIsMenuItemOwnerDraw (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

---

## AfxIsMenuItemPopup

Returns TRUE if the specified menu item is a submenu; FALSE otherwise.

```
FUNCTION AfxIsMenuItemPopup (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

---

## AfxIsMenuItemSeparator

Returns TRUE if the specified menu item is a separator; FALSE otherwise.

```
FUNCTION AfxIsMenuItemSeparator (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

---

## AfxRemoveCloseMenu

Removes the system menu close option and disables the X button.

```
SUB AfxRemoveCloseMenu (BYVAL hwnd AS HWND) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window that owns the menu. |

#### Return value

TRUE or FALSE.

---

## AfxRightJustifyMenuItem

Right justifies a top level menu item.

```
FUNCTION AfxRightJustifyMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window that owns the menu. |
| *uItem* | The zero-based position of the top level menu item to justify. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

#### Remarks

This is usually used to have the Help menu item right-justified on the menu bar.

---

## AfxSetMenuItemBold

Changes the text of a menu item to bold.

```
FUNCTION AfxSetMenuItemBold (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwnd* | Handle to the window that owns the menu. |
| *uItem* | The zero-based position of the top level menu item to bold. |

#### Return value

If the function succeeds, the return value is TRUE. If the function fails, the return value is FALSE.

---

## AfxSetMenuItemState

Sets the state of the specified menu item.

```
FUNCTION AfxSetMenuItemState (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fState AS DWORD, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fState* | The menu item state. It can be one or more of these values:<br>MFS_CHECKED: Checks the menu item.<br>MFS_DEFAULT: Specifies that the menu item is the default.<br>MFS_DISABLED: Disables the menu item and grays it so that it cannot be selected.<br>MFS_ENABLED: Enables the menu item so that it can be selected. This is the default state.<br>MFS_GRAYED: Disables the menu item and grays it so that it cannot be selected.<br>MFS_HILITE: Highlights the menu item.<br>MFS_UNCHECKED: Unchecks the menu item.<br>MFS_UNHILITE: Removes the highlight from the menu item. This is the default state. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

TRUE or FALSE. To get extended error information, use the **GetLastError** function.

#### Remarks

The application must call the **DrawMenuBar** function whenever a menu changes, whether or not the menu is in a displayed window.

---

## AfxSetMenuItemText

Sets the text of the specified menu item.

```
FUNCTION AfxSetMenuItemText (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYREF wszText AS WSTRING, BYVAL fByPosition AS LONG = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *wszText* | The text to set. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

TRUE or FALSE. To get extended error information, use the **GetLastError** function.

#### Remarks

The application must call the **DrawMenuBar** function whenever a menu changes, whether or not the menu is in a displayed window.

---

## AfxToggleMenuItem

Toggles the checked state of a menu item.

```
FUNCTION AfxToggleMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

The return value specifies the previous state of the menu item (either MF_CHECKED or MF_UNCHECKED). If the menu item does not exist, the return value is -1.

---

## AfxUnCheckMenuItem

Unchecks a menu item.

```
FUNCTION AfxUnCheckMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
   BYVAL fByPosition AS LONG = FALSE) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *uItem* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of *uItem*. If this parameter is FALSE, *uItem* is a menu item identifier. Otherwise, it is a menu item position. |

#### Return value

The return value specifies the previous state of the menu item (either MF_CHECKED or MF_UNCHECKED). If the menu item does not exist, the return value is -1.

---
