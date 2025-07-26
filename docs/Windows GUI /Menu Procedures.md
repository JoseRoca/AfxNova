# About Menus

A menu is a list of items that specify options or groups of options (a submenu) for an application. Clicking a menu item opens a submenu or causes the application to carry out a command. 

#### Menu Bars and Menus

A menu is arranged in a hierarchy. At the top level of the hierarchy is the *menu bar*; which contains a list of *menus*, which in turn can contain submenus. A menu bar is sometimes called a *top-level menu*, and the menus and submenus are also known as *pop-up menus*.

A menu item can either carry out a command or open a submenu. An item that carries out a command is called a *command item* or a *command*.

An item on the menu bar almost always opens a menu. Menu bars rarely contain command items. A menu opened from the menu bar drops down from the menu bar and is sometimes called a *drop-down menu*. When a drop-down menu is displayed, it is attached to the menu bar. A menu item on the menu bar that opens a drop-down menu is also called a menu name.

The menu names on a menu bar represent the main categories of commands that an application provides. Selecting a menu name from the menu bar typically opens a menu whose menu items correspond to the commands in a category. For example, a menu bar might contain a File menu name that, when clicked by the user, activates a menu with menu items such as New, Open, and Save. To get information about a menu bar, call **GetMenuBarInfo**.

Only an overlapped or pop-up window can contain a menu bar; a child window cannot contain one. If the window has a title bar, the system positions the menu bar just below it. A menu bar is always visible. A submenu is not visible, however, until the user selects a menu item that activates it. For more information about overlapped and pop-up windows, see [Window Types](https://learn.microsoft.com/en-us/windows/win32/winmsg/window-features).

Each menu must have an owner window. The system sends messages to a menu's owner window when the user selects the menu or chooses an item from the menu.

#### Shortcut Menus

The system also provides shortcut menus. A shortcut menu is not attached to the menu bar; it can appear anywhere on the screen. An application typically associates a shortcut menu with a portion of a window, such as the client area, or with a specific object, such as an icon. For this reason, these menus are also called *context menus*.

A shortcut menu remains hidden until the user activates it, typically by right-clicking a selection, a toolbar, or a taskbar button. The menu is usually displayed at the position of the caret or mouse cursor.

#### The Window Menu

The **Window** menu (also known as the **System** menu or **Control** menu) is a pop-up menu defined and managed almost exclusively by the operating system. The user can open the window menu by clicking the application icon on the title bar or by right-clicking anywhere on the title bar.

The **Window** menu provides a standard set of menu items that the user can choose to change a window's size or position, or close the application. Items on the window menu can be added, deleted, and modified, but most applications just use the standard set of menu items. An overlapped, pop-up, or child window can have a window menu. It is uncommon for an overlapped or pop-up window not to include a window menu.

When the user chooses a command from the **Window menu**, the system sends a **WM_SYSCOMMAND** message to the menu's owner window. In most applications, the window procedure does not process messages from the window menu. Instead, it simply passes the messages to the **DefWindowProc** function for system-default processing of the message. If an application adds a command to the window menu, the window procedure must process the command.

An application can use the **GetSystemMenu** function to create a copy of the default window menu to modify. Any window that does not use the **GetSystemMenu** function to make its own copy of the window menu receives the standard window menu.

See more documentation at [About Menus](https://learn.microsoft.com/en-us/windows/win32/menurc/about-menus) and [Usibg Menus](https://learn.microsoft.com/en-us/windows/win32/menurc/using-menus)

## Windows API Menu Procedures

See [Menu Functions](https://learn.microsoft.com/en-us/windows/win32/menurc/menu-functions)

| Name       | Description |
| ---------- | ----------- |
| **AppendMenu** | Appends a new item to the end of the specified menu bar, drop-down menu, submenu, or shortcut menu. |
| **CheckMenuItem** | Appends a new item to the end of the specified menu bar, drop-down menu, submenu, or shortcut menu. |
| **CheckMenuItem** | Sets the state of the specified menu item's check-mark attribute to either selected or clear. |
| **CheckMenuRadioItem** | Checks a specified menu item and makes it a radio item. |
| **CheckMenuRadioItem** | Checks a specified menu item and makes it a radio item. |
| **CreateMenu** | Creates a menu. |
| **CreatePopupMenu** | Creates a drop-down menu, submenu, or shortcut menu. |
| **CreatePopupMenu** | Creates a drop-down menu, submenu, or shortcut menu. |
| **DeleteMenu** | Deletes an item from the specified menu. |
| **DestroyMenu** | Destroys the specified menu and frees any memory that the menu occupies. |
| **DrawMenuBar** | Redraws the menu bar of the specified window. |
| **EnableMenuItem** | Enables, disables, or grays the specified menu item. |
| **EndMenu** | Ends the calling thread's active menu. |
| **GetMenu** | Retrieves a handle to the menu assigned to the specified window. |
| **GetMenuCheckMarkDimensions** | Retrieves the dimensions of the default check-mark bitmap. |
| **GetMenuDefaultItem** | Determines the default menu item on the specified menu. |
| **GetMenuInfo** | Retrieves information about a specified menu. |
| **GetMenuItemCount** | Determines the number of items in the specified menu. |
| **GetMenuItemID** | Retrieves the menu item identifier of a menu item located at the specified position in a menu. |
| **GetMenuItemInfo** | Retrieves information about a menu item. |
| **GetMenuItemRect** | Retrieves the bounding rectangle for the specified menu item. |
| **GetMenuState** | Retrieves the menu flags associated with the specified menu item. |
| **GetMenuString** | Copies the text string of the specified menu item into the specified buffer. |
| **GetSubMenu** | Retrieves a handle to the drop-down menu or submenu activated by the specified menu item. |
| **GetSubMenu** | Retrieves a handle to the drop-down menu or submenu activated by the specified menu item. |
| **GetSystemMenu** |Enables the application to access the window menu (also known as the system menu or the control menu) for copying and modifying. |
| **HiliteMenuItem** | Adds or removes highlighting from an item in a menu bar. |
| **InsertMenu** | Inserts a new menu item into a menu, moving other items down the menu. |
| **InsertMenuItem** | Inserts a new menu item at the specified position in a menu. |
| **IsMenu** | Determines whether a handle is a menu handle. |
| **LoadMenu** | Loads the specified menu resource from the executable (.exe) file associated with an application instance. |
| **LoadMenuIndirect** | Loads the specified menu template in memory. |
| **MenuItemFromPoint** | Determines which menu item, if any, is at the specified location. |
| **ModifyMenu** | Changes an existing menu item. |
| **RemoveMenu** | Deletes a menu item or detaches a submenu from the specified menu. |
| **SetMenu** | Assigns a new menu to the specified window. |
| **SetMenuDefaultItem** | Sets the default menu item for the specified menu. |
| **SetMenuInfo** | Sets information for a specified menu. |
| **SetMenuItemBitmaps** | Associates the specified bitmap with a menu item. |
| **SetMenuItemInfo** | Changes information about a menu item. |
| **TrackPopupMenu** | Displays a shortcut menu at the specified location and tracks the selection of items on the menu. |
| **TrackPopupMenuEx** | Displays a shortcut menu at the specified location and tracks the selection of items on the shortcut menu. |

---

## Menu Wrappers

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
| [AfxIsMenuItemOwnerdraw](#afxismenuitemownerdraw) | Returns TRUE if the specified menu item is a ownerdraw; FALSE otherwise. |
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

## DDT-like Menu Wrappers

These procedures replicate the PowerBASIC's menu procedures and add many more functionality using the same syntax. Contrarily to the *Afx menu* procedures, that use zero-based items, these ones are one-based.

| Name       | Description |
| ---------- | ----------- |
| [IsMenuHandle](#IsMenuHandle) | Determines whether a handle is a menu handle. |
| [IsMenuItemChecked](#ismenuitemchecked) | Returns TRUE if the specified menu item is checked; FALSE otherwise. |
| [IsMenuItemEnabled](#ismenuitemenabled) | Returns TRUE if the specified menu item is enabled; FALSE otherwise. |
| [IsMenuItemDisabled](#ismenuitemdisabled) | Returns TRUE if the specified menu item is disabled; FALSE otherwise. |
| [IsMenuItemGrayed](#ismenuitemgrayed) | Returns TRUE if the specified menu item is grayed; FALSE otherwise. |
| [IsMenuItemHighlighted](#ismenuitemghighlighted) | Returns TRUE if the specified menu item is grayed; FALSE otherwise. |
| [IsMenuItemSeparator](#ismenuitemseparator) | Returns TRUE if the specified menu item is a separator; FALSE otherwise. |
| [IsMenuItemOwnerdraw](#ismenuitemownerdraw) | Returns TRUE if the specified menu item is ownerdraw; FALSE otherwise. |
| [IsMenuItemPopup](#ismenuitempopup) |Returns TRUE if the specified menu item is a submenu; FALSE otherwise. |
| [MenuAddBitmapToItem](#menuaddbitmaptoitem) | Adds a bitmap to the menu item. |
| [MenuAddIconToItem](#menuaddicontoitem) | Adds an icon to the menu item. |
| [MenuAddPopup](#menuaddpopup) | Adds a popup child menu to an existing menu. |
| [MenuAddString](#menuaddstring) | Adds a string or separator to an existing menu. |
| [MenuAttach](#menuattach) | Attaches a menu to a window or dialog. |
| [MenuBoldItem](#menubolditem) | Changes the text of a menu item to bold. |
| [MenuCheckItem](#menucheckitem) | Checks a menu item. |
| [MenuCheckRadioButton](#menucheckradiobutton) | ' Checks a specified menu item and makes it a radio item. At the same time, the function clears all other menu items in the associated group and clears the radio-item type flag for those items. |
| [MenuContext](#MenuContext) | Creates a floating context menu. |
| [MenuDelete](#menudelete) | Deletes a menu item from an existing menu. |
| [MenuDestroy](#menudestroy) | Destroys the main menu from the window or dialog. |
| [MenuDisableItem](#menudisableitem) | Disables the specified menu item. |
| [MenuDrawBar](#menudrawbar) | Redraws the menu bar of the specified window or dialog. |
| [MenuEnableItem](#menuenableitem) | Enables the specified menu item. |
| [MenuFindItemPosition](#menufinditemposition) | Finds the position of the specified menu item. |
| [MenuGetBarInfo](#menugetbarinfo) | Retrieves information about the specified menu bar. |
| [MenuGetCheckMarkHeight](#menugetcheckmarkheight) | Retrieves the height of the default check-mark bitmap. |
| [MenuGetContextHelpId](#menugetcontexthelpid) | Retrieves the Help context identifier associated with the specified menu. |
| [MenuGetCheckMarkWidth](#menugetcheckmarkwidth) | Retrieves the dimensions of the default check-mark bitmap. |
| [MenuGetDefaultItem](#menugetdefaultitem) | Determines the default menu item on the specified menu. |
| [MenuGetFont](#menugetfont) | Retrieves information about the font used in menu bars. |
| [MenuGetFontPointSize](#menugetfontpointsize) | Retrieves the point size of the font used in menu bars. |
| [MenuGetHandle](#menugethandle) | Retrieves a handle to the menu assigned to the specified window or dialog.  |
| [MenuGetItemCount](#menugetitemcount) | Determines the number of items in the specified menu. |
| [MenuGetItemFromPoint](#menugetitemfrompoint) | Determines which menu item, if any, is at the specified location. |
| [MenuGetItemID](#menugetitemid) | Retrieves the menu item ID of a menu item located at the specified position in a menu. |
| [MenuGetRect](#menugetrect) | Calculates the size of a menu bar or a drop-down menu. |
| [MenuGetState](#menugetstate) | Retrieves the state of the specified menu item. |
| [MenuGetSubMenu](#menugetsubmenu) | Retrieves a handle to the drop-down menu or submenu activated by the specified menu item. |
| [MenuGetSubmenusCount](#menugetsubmenuscount) | Get the number of submenus of a menu. |
| [MenuGetText](#menugettext) | Retrieves the text of the specified menu item. |
| [MenuGetTextLen](#menugettexlLen) | Returns the lengnth of the text of the specified menu item. |
| [MenuGetWindowOwner](#menugetwindowowner) | Retrieve the window owner of the specified menu. |
| [MenuGetSytemMenuHandle](#menugetsytemmenuhandle) | Enables the application to access the window menu (also known as the system menu or the control menu) for copying and modifying.  |
| [MenuGrayItem](#menugrayitem) | Disables the specified menu item. |
| [MenuHiliteItem](#menuhiliteitem) | Highlights the specified menu item. |
| [MenuItemToggleCheckState](#menuitemtogglecheckstate) | Retrieves the state of the specified menu item. |
| [MenuNewBar](#menunewbar) | Creates a new menu bar. |
| [MenuNewPopup](#menunewpopup) | Creates a new popup menu. |
| [MenuRemoveCloseOptiom](#menuremovecloseoptiom) | Removes the system menu close option and disables the X button. |
| [MenuRestoreCloseOption](#menurestorecloseoption) | Restores the system menu close option and enables Alt+F4 and the X button. |
| [MenuRightJustifyItem](#menurightjustifyitem) | Right justifies a top level menu item. This is usually used to have the Help menu item. |
| [MenuSetContextHelpId](#menusetcontexthelpid) | Associates a Help context identifier with a menu. |
| [MenuSetDefaultItem](#menusetdefaultitem) | Sets the default menu item for the specified menu. |
| [MenuSetItemBitmaps](#MenuSetItemBitmaps) | Associates the specified bitmap with a menu item. Whether the menu item is selected or clear, the system displays the appropriate bitmap next to the menu item. |
| [MenuSetState](#menusetstate) | Sets the state of the specified menu item. |
| [MenuSetText](#menusettext) | Retrieves the text of the specified menu item. |
| [MenuUnCheckItem](#menuuncheckitem) | Unchecks a menu item. |

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

## AfxIsMenuItemOwnerdraw

Returns TRUE if the specified menu item is a ownerdraw; FALSE otherwise.

```
FUNCTION AfxIsMenuItemOwnerdraw (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
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

## IsMenuHandle

Determines whether a handle is a menu handle.

```
FUNCTION IsMenuHandle (BYVAL hMenu AS HMENU) AS BOOLEAN
```

#### Return value

Returns TRUE if the specified handle is a menu handle; FALSE otherwise.

---

## IsMenuItemChecked

Determines whether the specified menu item is checked.

```
FUNCTION IsMenuItemChecked (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

#### Return value

Returns TRUE if the specified menu item is checked; FALSE otherwise.

---

## IsMenuItemEnabled

Determines whether the specified menu item is enabled.

```
FUNCTION IsMenuItemEnabled (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

#### Return value

Returns TRUE if the specified menu item is enabled; FALSE otherwise.

---


## IsMenuItemDisabled

Determines whether the specified menu item is disabled.

```
FUNCTION IsMenuItemDisabled (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

#### Return value

Returns TRUE if the specified menu item is disabled; FALSE otherwise.

---

## IsMenuItemGrayed

Determines whether the specified menu item is grayed.

```
FUNCTION IsMenuItemGrayed (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

#### Return value

Returns TRUE if the specified menu item is grayed; FALSE otherwise.

---

## IsMenuItemHighlighted

Determines whether the specified menu item is highlighted.

```
FUNCTION IsMenuItemHighlighted (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

#### Return value

Returns TRUE if the specified menu item is highlighted; FALSE otherwise.

---

## IsMenuItemSeparator

Determines whether the specified menu item is a separator.

```
FUNCTION IsMenuItemSeparator (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

#### Return value

Returns TRUE if the specified menu item is a separator; FALSE otherwise.

---

## IsMenuItemOwnerdraw

Determines whether the specified menu item is ownerdraw.

```
FUNCTION IsMenuItemOwnerdraw (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

#### Return value

Returns TRUE if the specified menu item is ownerdraw; FALSE otherwise.

---

## IsMenuItemPopup

Determines whether the specified menu item is a popup item.

```
FUNCTION IsMenuItemPopup (BYVAL hMenu AS HMENU, BYVAL item AS LONG, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

#### Return value

Returns TRUE if the specified menu item is a popup item; FALSE otherwise.

---

## MenuAddBitmapToItem

Adds a bitmap to the menu item.

```
FUNCTION MenuAddBitmapToItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN, BYVAL hbmp AS HBITMAP) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to change. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of item. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position. |
| *hbmp* | The bitmap handle. |

#### Return value

Returns TRUE if the function succeeds; FALSE otherwise. To get extended error information, use the **GetLastError** function.

---
