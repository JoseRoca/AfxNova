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

## Windows API Menu Procedures

Menu functions provided by Windows.

See [Menu Functions](https://learn.microsoft.com/en-us/windows/win32/menurc/menu-functions)

| Name       | Description |
| ---------- | ----------- |
| **AppendMenu** | Appends a new item to the end of the specified menu bar, drop-down menu, submenu, or shortcut menu. |
| **CheckMenuItem** | Sets the state of the specified menu item's check-mark attribute to either selected or clear. |
| **CheckMenuRadioItem** | Checks a specified menu item and makes it a radio item. |
| **CreateMenu** | Creates a menu. |
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

## SDK Style Menu Wrappers

Wwappers to add functionality or convenience of use to the above listed SDK functions.

| Name       | Description |
| ---------- | ----------- |
| [AddBitmapToItem](#addbitmaptoitem) | Adds a bitmap to the menu item. |
| [AddPopup](#addpopup) | Adds a popup child menu to an existing menu. |
| [AddString](#addstring) | Adds a string or separator to an existing menu. |
| [Append](#append) | Appends a new item to the end of the specified menu bar, drop-down menu, submenu, or shortcut menu. |
| [Attach](#attach) | Attaches a menu to a window or dialog. |
| [BoldItem](#checkitem) | Changes the text of a menu item to bold. |
| [CheckItem](#checkitem) | Checks a menu item. |
| [CheckRadioButton](#checkradiobutton) | Checks a specified menu item and makes it a radio item. |
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
| [GetItemCount](#getitemcount) | Determines the number of items in the specified menu. |
| [GetItemFromPoint](#getitemfrompoint) | Determines which menu item, if any, is at the specified location. |
| [GetItemID](#getitemid) | Retrieves the menu item ID of a menu item located at the specified position in a menu. |
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
| [NewBar](#newbar) | Creates a new menu bar. |
| [NewPopup](#newpopup) | Creates a drop-down menu, submenu, or shortcut menu. |
| [RemoveCloseMenu](#removeclosemenu) | Removes the system menu close option and disables the X button. |
| [RemoveCloseOptiom](#removecloseoption) | Removes the system menu close option and disables the X button. |
| [RemoveItem](#removeitem) | Deletes a menu item from an existing menu. |
| [RestoreCloseOption](#restorecloseoption) | Restores the system menu close option and enables Alt+F4 and the X button. |
| [RightJustifyItem](#rightjustifyitem) |  Right justifies a top level menu item. This is usually used to have the Help menu item right-justified on the menu bar. |
| [SetContextHelpId](#setcontexthelpid) | Associates a Help context identifier with a menu. |
| [SetDefaultItem](#setdefaultitem) | Sets the default menu item for the specified menu. |
| [SetItemBitmaps](#setitembitmaps) | Associates the specified bitmap with a menu item. |
| [SetItemBold](#setitembold) | Changes the text of a menu item to bold. |
| [SetItemText](#setitemtext) | Sets the text of the specified menu item. |
| [SetItemState](#setitemstate) | Sets the state of the specified menu item. |
| [SetState](#setstate) | Sets the state of the specified menu item. |
| [ToggleCheckState](#togglecheckstate) | Toggles the checked state of a menu item. |
| [ToggleItem](#toggleitem) | Toggles the checked state of a menu item. |
| [UncheckItem](#uncheckitem) | Unchecks a menu item. |

---
