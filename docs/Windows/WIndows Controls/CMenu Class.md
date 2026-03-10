# About Menus

A menu is a list of items that specify options or groups of options (a submenu) for an application. Clicking a menu item opens a submenu or causes the application to carry out a command. 

#### Menu Bars and Menus

A menu is arranged in a hierarchy. At the top level of the hierarchy is the *menu bar*; which contains a list of *menus*, which in turn can contain submenus. A menu bar is sometimaes called a *top-level menu*, and the menus and submenus are also known as *pop-up menus*.

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

See more MSDN documentation at [About Menus](https://learn.microsoft.com/en-us/windows/win32/menurc/about-menus) and [Using Menus](https://learn.microsoft.com/en-us/windows/win32/menurc/using-menus)

---

## Examples

| Name       | Description |
| ---------- | ----------- |
| [CWindow Menu](#cwindowmenu) | Builds a menu using CWindow. |
| [CDialog Menu](#cdialogmenu) | Builds a menu using CDialog. |

---

## Methods

| Name       | Description |
| ---------- | ----------- |
| [AddBitmapToItem](#addbitmaptoitem) | Adds a bitmap to the menu item. |
| [AddBitmapToMenuItem](#addbitmaptomenuinfo) | Adds a bitmap to the menu item. |
| [AddIconToMenuItem](#addicontomenuinfo) | Adds a bitmap to the menu item. |
| [AddPopup](#addpopup) | Adds a popup child menu to an existing menu. |
| [AddString](#addstring) | Adds a string or separator to an existing menu. |
| [Append](#append) | Appends a new item to the end of the specified menu bar, drop-down menu, submenu, or shortcut menu. |
| [Attach](#attach) | Attaches a menu to a window or dialog. |
| [BoldItem](#checkitem) | Changes the text of a menu item to bold. |
| [CheckItem](#checkitem) | Checks a menu item. |
| [CheckRadioButton](#checkradiobutton) | Checks a specified menu item and makes it a radio item. |
| [CheckRadioButton](#checkradiobutton) | Checks a specified menu item and makes it a radio item. |
| [ContextMenu](#contextmenu) | Creates a floating context menu. |
| [Create](#create) | Creates a new menu bar. |
| [CreatePopup](#createpopup) | Creates a drop-down menu, submenu, or shortcut menu. |
| [DeleteItem](#deleteitem) | Deletes a menu item from an existing menu. |
| [Destroy](#destroy) | Destroys the specified menu and frees any memory that the menu occupies. |
| [DisableItem](#disableitem) | Disables the specified menu item. |
| [DrawBar](#frawbar) |Redraws the menu bar of the specified window. |
| [EnableItem](#enableitem) | Enables the specified menu item. |
| [FindItemPosition](#finditemposition) | Finds the position of the specified menu item. |
| [GetBarInfo](#getarinfo) | Retrieves information about the specified menu bar. |
| [GetCheckMarkHeight](#getcheckmarkheight) | Retrieves the height of the default check-mark bitmap. |
| [GetCheckMarkWidth](#getcheckmarkwidth) | Retrieves the width of the default check-mark bitmap. |
| [GetContextHelpId](#getcontexthelpid) | Retrieves the Help context identifier associated with the specified menu. |
| [GetDefaultItem](#getdefaultitem) | Determines the default menu item on the specified menu. |
| [GetFont](#getfont) | Retrieves information about the font used in menu bars. |
| [GetFontPointSize](#getfontpointsize) | Retrieves the point size of the font used in menu bars. |
| [GetHandle](#gethandle) | Retrieves a handle to the menu assigned to the specified window or dialog.  |
| [GetInfo](#getinfo) | Gets information for a specified menu. |
| [GetItemCount](#getitemcount) | Determines the number of items in the specified menu. |
| [GetItemFromPoint](#getitemfrompoint) | Determines which menu item, if any, is at the specified location. |
| [GetItemID](#getitemid) | Retrieves the menu item ID of a menu item located at the specified position in a menu. |
| [GetItemInfo](#getiteminfo) | Retrieves information about a menu item. |
| [GetItemRect](#getitemrect) | Retrieves the bounding rectangle for the specified menu item. |
| [GetItemState](#getitemstate) | Retrieves the state of the specified menu item. |
| [GetItemText](#getitemtext) | Retrieves the text of the specified menu item. |
| [GetItemTextLen](#getitemtextlen) | Returns the lengnth of the specified menu item. |
| [GetRect](#getrect) | Calculates the size of a menu bar or a drop-down menu. |
| [GetState](#getstate) | Retrieves the state of the specified menu item. |
| [GetSubMenu](#ghetsubmenu) | Retrieves a handle to the drop-down menu or submenu activated by the specified menu item. |
| [GetSubmenusCount](#ghetsubmenuscount) | Retrieves the number of submenus of a menu. |
| [GetSystemMenuHandle](#getsystemmenuhandle) | Enables the application to access the window menu (also known as the system menu or the control menu) for copying and modifying. |
| [GetWindowOwner](#ghetwindowowner) | Retrieves the window owner of the specified menu |
| [GrayItem](#grayitem) | Grays the specified menu item. |
| [HiliteItem](#hiliteitem) | Highlights the specified menu item. |
| [IsItemChecked](#isitemchecked) | Returns TRUE if the specified menu item is checked; FALSE otherwise. |
| [IsItemDisabled](#isitemdisabled) | Returns TRUE if the specified menu item is enabled; FALSE otherwise. |
| [IsItemEnabled](#isitemenabled) | Returns TRUE if the specified menu item is disabled; FALSE otherwise. |
| [IsItemGrayed](#isitemgrayed) | Returns TRUE if the specified menu item is grayed; FALSE otherwise. |
| [IsItemHighlighted](#isitemhighlighted) | Returns TRUE if the specified menu item is highlighted; FALSE otherwise. |
| [IsItemOwnerdraw](#isitemownerdraw) | Returns TRUE if the specified menu item is ownerdraw; FALSE otherwise. |
| [IsItemPopup](#isitempopup) | Returns TRUE if the specified menu item is a submenu; FALSE otherwise. |
| [IsItemSeparator](#isitemseparator) | Returns TRUE if the specified menu item is a separator; FALSE otherwise. |
| [IsMenu](#ismenu) | Determines whether a handle is a menu handle. |
| [IsMenuHandle](#ismenuhadle) | Determines whether a handle is a menu handle. |
| [Load](#load) | Loads the specified menu resource from the executable (.exe) file associated with an application instance. |
| [LoadIndirect](#loadindirect) | Loads the specified menu template in memory. |
| [Modify](#modify) | Changes an existing menu item. |
| [NewBar](#newbar) | Creates a new menu bar. |
| [NewPopup](#newpopup) | Creates a drop-down menu, submenu, or shortcut menu. |
| [RemoveCloseMenu](#removeclosemenu) | Removes the system menu close option and disables the X button. |
| [RemoveCloseOptiom](#removecloseoption) | Removes the system menu close option and disables the X button. |
| [RemoveItem](#removeitem) | Deletes a menu item from an existing menu. |
| [RestoreCloseOption](#restorecloseoption) | Restores the system menu close option and enables Alt+F4 and the X button. |
| [RightJustifyItem](#rightjustifyitem) |  Right justifies a top level menu item. This is usually used to have the Help menu item right-justified on the menu bar. |
| [SetContextHelpId](#setcontexthelpid) | Associates a Help context identifier with a menu. |
| [SetDefaultItem](#setdefaultitem) | Sets the default menu item for the specified menu. |
| [SetInfo](#setifo) | Sets information for a specified menu. |
| [SetItemBitmaps](#setitembitmaps) | Associates the specified bitmap with a menu item. |
| [SetItemBold](#bolditem) | Changes the text of a menu item to bold. |
| [SetItemInfo](#setiteminfo) | Changes information about a menu item. |
| [SetItemText](#setitemtext) | Sets the text of the specified menu item. |
| [SetItemState](#setitemstate) | Sets the state of the specified menu item. |
| [SetState](#setstate) | Sets the state of the specified menu item. |
| [ToggleCheckState](#togglecheckstate) | Toggles the checked state of a menu item. |
| [ToggleItem](#toggleitem) | Toggles the checked state of a menu item. |
| [TrackPopupMenu](#trackpopupmenu) | Displays a shortcut menu at the specified location and tracks the selection of items on the menu. |
| [TrackPopupMenuEx](#trackpopupmenuex) | Displays a shortcut menu at the specified location and tracks the selection of items on the menu. |
| [UncheckItem](#uncheckitem) | Unchecks a menu item. |

---

### AddBitmapToItem

Adds a bitmap to the menu item.

```
FUNCTION AddBitmapToItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG, _
   BYVAL fByPosition AS BOOLEAN, BYVAL hbmp AS HBITMAP) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to change. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of item. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position, where position = 1 for the first position, position = 2 for the second, and so on. |
| *hbmp* | The bitmap handle. |

#### Return value

Returns TRUE if the function succeeds; FALSE otherwise.

---

### AddBitmapToMenuItem

Adds a bitmap to the menu item.

```
FUNCTION AddBitmapToMenuItem (BYVAL hMenu AS HMENU, BYVAL nMenuItem AS LONG, _
   BYVAL fByPosition AS BOOLEAN, BYVAL hbmp AS HBITMAP) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *nMenuItem* | The identifier or position of the menu item to change. The meaning of this parameter depends on the value of *fByPosition*. |
| *fByPosition* | The meaning of item. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position, where position = 1 for the first position, position = 2 for the second, and so on. |
| *hbmp* | The bitmap handle. |

#### Return value

Returns TRUE if the function succeeds; FALSE otherwise.

---

### AddIconToMenuItem

Converts an icon handle to a bitmap and adds it to the specified *hbmpItem field* of **HMENU** item.

```
FUNCTION AddIconToMenuItem (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD, _
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
IF hIcon THEN CMenu.AddIconToMenuItem(hSubMenu, 0, TRUE, hIcon)
```

Loading the icon from a resource file:

```
DIM hSubMenu AS HMENU = GetSubMenu(hMenu, 1)
CMenu.AddIconToMenuItem(hSubMenu, 0, TRUE, AfxGdipIconFromRes(hInstance, "IDI_UNDO_32"))
```
---

### AddPopup

Adds a popup child menu to an existing menu. A popup menu is a small window that "pops up" when a menu item is highlighted. This allows nesting, and gives the user an opportunity to choose from "sub-menu" items.

```
FUNCTION AddPopup (BYVAL hMenu AS HMENU, BYREF wszText AS WSTRING, BYVAL hPopup AS HMENU, _
   BYVAL fState AS UINT, BYVAL item AS LONG = 0, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle of the parent menu which holds the popup. |
| *wszText* | Text displayed in the parent menu. An ampersand (&) may be used in the  to make the following letter into a control accelerator (hot-key). The letter appears underscored to signify that it is an accelerator. |
| *hPopup* | Handle of the child popup menu to be added. |
| *fState* | The initial state of the menu item. It can be one of the following:<br>MFS_DISABLED: Disable the item so that it cannot be selected.<br>%MFS_ENABLED:  Enable the item so that it can be selected. |
| *item* | If the option is included, *item* is a unique numeric identifier for this popup menu. *item* may be used later with a *fByPosition* option to reference this popup. *item* is an integral numeric value in the range of -32768 to +32767. |
| *fByPosition* | Indicates the position in the parent menu where the popup child menu is to be inserted. If the *fByPosition* option is used, the popup menu is inserted prior to the menu item specified by *item*. Otherwise, the popup menu is inserted at the physical position within the parent menu, where position = 1 for the first position, position = 2 for the second, and so on. If position is not specified then the popup menu is appended to the end of the menu. |

#### Return value

Returns TRUE if the function succeeds; FALSE otherwise.

---

## AddString

Adds a string or separator to an existing menu. A string may contain an optional command accelerator key, and also describe an equivalent keyboard accelerator combination.

```
FUNCTION AddString OVERLOAD (BYVAL hMenu AS HMENU, BYREF wszText AS WSTRING, BYVAL id AS LONG, _
   BYVAL fState AS UINT, BYVAL item AS LONG = 0, BYVAL fByPosition AS BOOLEAN = FALSE) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle of the parent menu to which the string should be added. |
| *wszText* | Text to display in the parent menu. An ampersand (&) may be used in the string to make the following letter into a command accelerator (hot-key). The letter is underscored to signify that it is an accelerator. To create a horizontal separator instead of a text string, set wszText = "-", id = 0, fState = 0. Keyboard accelerators can be indicated in the text of a menu item, for the reference of the user. To include a keyboard accelerator description in a menu string, separate it from the menu item text with a CHR(9) character. For example: MenuAddString hMenu, "Cu&t" & CHR(9) & "CTRL+X", id, fState. |
| *id* | The unique  identifier for the menu item. When a menu item is selected, *id* is sent to the parent dialog callback function to notify the dialog which option was selected. |
| *fState* | The initial state of the menu item. It can be one or more of the following, combined together with the OR operator to form a bitmask: <br>MFS_CHECKED: Place a checkmark next to the item.<br>MFS_DEFAULT: The default menu item, displayed in bold.  Only one item may be the default.<br>MFS_DISABLED: Disable the menu item so that it cannot be selected.<br>MFS_ENABLED: Enable the menu item so that it can be selected.<br>MFS_GRAYED: Disable the menu item so that it cannot be selected, and draw it in a "grayed" state to indicate this.<br>MFS_HILITE: Highlight the menu item.<br>MFS_UNCHECKED: Do not place a checkmark next to the item.<br>MFS_UNHILITE: Item is not highlighted.<br>A state value of zero (0) provides MFS_ENABLED, MFS_UNCHECKED, and MFS_UNHILITE. |
| *item* | Optional position in the parent menu, where the menu item should be inserted. If the *fByPosition* option is used, the menu item is inserted prior to the menu item identifier specified by *item*. Otherwise, the menu item is inserted at the physical position within the parent menu, where position = 1 for the first position, position = 2 for the second, and so on. If position is not specified then the popup menu is appended to the end of the menu. |
| *fByPosition* | The meaning of item. If this parameter is FALSE, *item* is a menu item identifier. Otherwise, it is a menu item position, where position = 1 for the first position, position = 2 for the second, and so on. |

#### Return value

Returns TRUE if the function succeeds; FALSE otherwise.

#### Remarks

The application must call the **DrawBar** statement whenever a menu changes, whether or not the menu is in a displayed dialog.

#### Usage examples
```
CMenu.AddString hPopup1, "&Open", ID_OPEN, MF_ENABLED
```
Insert the item before the ID_OPEN item
```
CMenu.AddString hPopup1, "&Exit", ID_EXIT, MF_ENABLED, 1, TRUE          ' insert by position
CMenu.AddString hPopup1, "&Exit", ID_EXIT, MF_ENABLED, ID_OPEN, FALSE   ' insert by identifier
```
---

## Append

Appends a new item to the end of the specified menu bar, drop-down menu, submenu, or shortcut menu. You can use this function to specify the content, appearance, and behavior of the menu item.

```
FUNCTION Append (BYVAL hMenu AS HMENU, BYVAL uFlags AS UINT, BYVAL uIDNewItem AS UINT_PTR, _
   BYVAL pwszNewItem AS WSTRING PTR) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle of the parent menu to which the item should be added. |
| *uFlags* | Controls the appearance and behavior of the new menu item. This parameter can be a combination of the following values (see below). |
| *uIDNewItem* | The identifier of the new menu item or, if the *uFlags* parameter is set to **MF_POPUP**, a handle to the drop-down menu or submenu. |
| *pwszNewItem* | The content of the new menu item. The interpretation of *pwszNewItem* depends on whether the *uFlags* parameter includes the following values (see below). |

| Flag  | Meaning |
| ----- | ------- |
| **MF_BITMAP** | Contains a bitmap handle. |
| **MF_OWNERDRAW** | Contains an application-supplied value that can be used to maintain additional data related to the menu item. The value is in the **itemData** member of the structure pointed to by the *lParam* parameter of the **WM_MEASUREITEM** or **WM_DRAWITEM** message sent when the menu is created or its appearance is updated. |
| **MF_STRING** | Contains a pointer to a null-terminated string. |

#### Values for the *uFlags* parameter

| Flag  | Meaning |
| ----- | ------- |
| **MF_BITMAP** | Uses a bitmap as the menu item. The lpNewItem parameter contains a handle to the bitmap. |
| **MF_CHECKED** | Places a check mark next to the menu item. If the application provides check-mark bitmaps (see SetMenuItemBitmaps, this flag displays the check-mark bitmap next to the menu item. |
| **MF_DISABLED** | Disables the menu item so that it cannot be selected, but the flag does not gray it. |
| **MF_ENABLED** | Enables the menu item so that it can be selected, and restores it from its grayed state. |
| **MF_GRAYED** | Disables the menu item and grays it so that it cannot be selected. |
| **MF_MENUBARBREAK** | Functions the same as the **MF_MENUBREAK** flag for a menu bar. For a drop-down menu, submenu, or shortcut menu, the new column is separated from the old column by a vertical line. |
| **MF_MENUBREAK** | Places the item on a new line (for a menu bar) or in a new column (for a drop-down menu, submenu, or shortcut menu) without separating columns. |
| **MF_OWNERDRAW** | Specifies that the item is an owner-drawn item. Before the menu is displayed for the first time, the window that owns the menu receives a **WM_MEASUREITEM** message to retrieve the width and height of the menu item. The **WM_DRAWITEM** message is then sent to the window procedure of the owner window whenever the appearance of the menu item must be updated. |
| **MF_POPUP** | Specifies that the menu item opens a drop-down menu or submenu. The *uIDNewItem* parameter specifies a handle to the drop-down menu or submenu. This flag is used to add a menu name to a menu bar, or a menu item that opens a submenu to a drop-down menu, submenu, or shortcut menu. |
| **MF_SEPARATOR** | Draws a horizontal dividing line. This flag is used only in a drop-down menu, submenu, or shortcut menu. The line cannot be grayed, disabled, or highlighted. The *pwszNewItem* and *uIDNewItem* parameters are ignored. |
| **MF_STRING** | Specifies that the menu item is a text string; the lpNewItem parameter is a pointer to the string. |
| **MF_UNCHECKED** | Does not place a check mark next to the item (default). If the application supplies check-mark bitmaps (see **SetMenuItemBitmaps**), this flag displays the clear bitmap next to the menu item. |

---

### Attach

Attaches a menu to a window or dialog.

```
FUNCTION Attach (BYVAL hMenu AS HMENU, BYVAL hwnd AS HWND) AS BOOLEAN
```

#### Return value

Returns TRUE if the function succeeds; FALSE otherwise.

#### Remarks

The Windows API function **SetMenu** performs the same action.

---

## BoldItem

Changes the text of a menu item to bold.

```
FUNCTION BoldItem (BYVAL hMenu AS HMENU, BYVAL item AS LONG) AS BOOLEAN
FUNCTION SetItemBold (BYVAL hMenu AS HMENU, BYVAL uItem AS DWORD) AS BOOLEAN
```
| Parameter  | Description |
| ---------- | ----------- |
| *hMenu* | Handle to the menu that contains the menu item. |
| *item* | The identifier or position of the menu item to get information about. The meaning of this parameter depends on the value of *fByPosition*. |

#### Return value

Returns TRUE if the function succeeds; FALSE otherwise.

---

++++++++++++++

| [CheckItem](#checkitem) | Checks a menu item. |
| [CheckRadioButton](#checkradiobutton) | Checks a specified menu item and makes it a radio item. |
| [CheckRadioButton](#checkradiobutton) | Checks a specified menu item and makes it a radio item. |
| [ContextMenu](#contextmenu) | Creates a floating context menu. |
| [Create](#create) | Creates a new menu bar. |
| [CreatePopup](#createpopup) | Creates a drop-down menu, submenu, or shortcut menu. |
| [DeleteItem](#deleteitem) | Deletes a menu item from an existing menu. |
| [Destroy](#destroy) | Destroys the specified menu and frees any memory that the menu occupies. |
| [DisableItem](#disableitem) | Disables the specified menu item. |
| [DrawBar](#frawbar) |Redraws the menu bar of the specified window. |
| [EnableItem](#enableitem) | Enables the specified menu item. |
| [FindItemPosition](#finditemposition) | Finds the position of the specified menu item. |
| [GetBarInfo](#getarinfo) | Retrieves information about the specified menu bar. |
| [GetCheckMarkHeight](#getcheckmarkheight) | Retrieves the height of the default check-mark bitmap. |
| [GetCheckMarkWidth](#getcheckmarkwidth) | Retrieves the width of the default check-mark bitmap. |
| [GetContextHelpId](#getcontexthelpid) | Retrieves the Help context identifier associated with the specified menu. |
| [GetDefaultItem](#getdefaultitem) | Determines the default menu item on the specified menu. |
| [GetFont](#getfont) | Retrieves information about the font used in menu bars. |
| [GetFontPointSize](#getfontpointsize) | Retrieves the point size of the font used in menu bars. |
| [GetHandle](#gethandle) | Retrieves a handle to the menu assigned to the specified window or dialog.  |
| [GetInfo](#getinfo) | Gets information for a specified menu. |
| [GetItemCount](#getitemcount) | Determines the number of items in the specified menu. |
| [GetItemFromPoint](#getitemfrompoint) | Determines which menu item, if any, is at the specified location. |
| [GetItemID](#getitemid) | Retrieves the menu item ID of a menu item located at the specified position in a menu. |
| [GetItemInfo](#getiteminfo) | Retrieves information about a menu item. |
| [GetItemRect](#getitemrect) | Retrieves the bounding rectangle for the specified menu item. |
| [GetItemState](#getitemstate) | Retrieves the state of the specified menu item. |
| [GetItemText](#getitemtext) | Retrieves the text of the specified menu item. |
| [GetItemTextLen](#getitemtextlen) | Returns the lengnth of the specified menu item. |
| [GetRect](#getrect) | Calculates the size of a menu bar or a drop-down menu. |
| [GetState](#getstate) | Retrieves the state of the specified menu item. |
| [GetSubMenu](#ghetsubmenu) | Retrieves a handle to the drop-down menu or submenu activated by the specified menu item. |
| [GetSubmenusCount](#ghetsubmenuscount) | Retrieves the number of submenus of a menu. |
| [GetSystemMenuHandle](#getsystemmenuhandle) | Enables the application to access the window menu (also known as the system menu or the control menu) for copying and modifying. |
| [GetWindowOwner](#ghetwindowowner) | Retrieves the window owner of the specified menu |
| [GrayItem](#grayitem) | Grays the specified menu item. |
| [HiliteItem](#hiliteitem) | Highlights the specified menu item. |
| [IsItemChecked](#isitemchecked) | Returns TRUE if the specified menu item is checked; FALSE otherwise. |
| [IsItemDisabled](#isitemdisabled) | Returns TRUE if the specified menu item is enabled; FALSE otherwise. |
| [IsItemEnabled](#isitemenabled) | Returns TRUE if the specified menu item is disabled; FALSE otherwise. |
| [IsItemGrayed](#isitemgrayed) | Returns TRUE if the specified menu item is grayed; FALSE otherwise. |
| [IsItemHighlighted](#isitemhighlighted) | Returns TRUE if the specified menu item is highlighted; FALSE otherwise. |
| [IsItemOwnerdraw](#isitemownerdraw) | Returns TRUE if the specified menu item is ownerdraw; FALSE otherwise. |
| [IsItemPopup](#isitempopup) | Returns TRUE if the specified menu item is a submenu; FALSE otherwise. |
| [IsItemSeparator](#isitemseparator) | Returns TRUE if the specified menu item is a separator; FALSE otherwise. |
| [IsMenu](#ismenu) | Determines whether a handle is a menu handle. |
| [IsMenuHandle](#ismenuhadle) | Determines whether a handle is a menu handle. |
| [Load](#load) | Loads the specified menu resource from the executable (.exe) file associated with an application instance. |
| [LoadIndirect](#loadindirect) | Loads the specified menu template in memory. |
| [Modify](#modify) | Changes an existing menu item. |
| [NewBar](#newbar) | Creates a new menu bar. |
| [NewPopup](#newpopup) | Creates a drop-down menu, submenu, or shortcut menu. |
| [RemoveCloseMenu](#removeclosemenu) | Removes the system menu close option and disables the X button. |
| [RemoveCloseOptiom](#removecloseoption) | Removes the system menu close option and disables the X button. |
| [RemoveItem](#removeitem) | Deletes a menu item from an existing menu. |
| [RestoreCloseOption](#restorecloseoption) | Restores the system menu close option and enables Alt+F4 and the X button. |
| [RightJustifyItem](#rightjustifyitem) |  Right justifies a top level menu item. This is usually used to have the Help menu item right-justified on the menu bar. |
| [SetContextHelpId](#setcontexthelpid) | Associates a Help context identifier with a menu. |
| [SetDefaultItem](#setdefaultitem) | Sets the default menu item for the specified menu. |
| [SetInfo](#setifo) | Sets information for a specified menu. |
| [SetItemBitmaps](#setitembitmaps) | Associates the specified bitmap with a menu item. |
| [SetItemInfo](#setiteminfo) | Changes information about a menu item. |
| [SetItemText](#setitemtext) | Sets the text of the specified menu item. |
| [SetItemState](#setitemstate) | Sets the state of the specified menu item. |
| [SetState](#setstate) | Sets the state of the specified menu item. |
| [ToggleCheckState](#togglecheckstate) | Toggles the checked state of a menu item. |
| [ToggleItem](#toggleitem) | Toggles the checked state of a menu item. |
| [TrackPopupMenu](#trackpopupmenu) | Displays a shortcut menu at the specified location and tracks the selection of items on the menu. |
| [TrackPopupMenuEx](#trackpopupmenuex) | Displays a shortcut menu at the specified location and tracks the selection of items on the menu. |
| [UncheckItem](#uncheckitem) | Unchecks a menu item. |
