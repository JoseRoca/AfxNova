# DWStrList

`DWStrList`implements an indexed double-linked list for the `DWSTRING` (Unicode dynamic string) data type

#### Usage example:

' // Build the linked list
DIM List AS DWStrList PTR = NEW DWStrList
List->Add("Result 1")
List->Add("Result 2")
List->Add("Result 3")
List->Insert(1, "New string")
List->Replace(2, "Replaced string")

' // Retrieve and print the results
FOR i AS LONG = 1 TO List->Count
   PRINT List->Item(i)
NEXT

' // Delete the list
Delete List

---

## Methods

| Name       | Description |
| ---------- | ----------- |
| [Add](#add) | Appends an item to the list. |
| [Clear](#clear) | Empties the entire list. |
| [Count](#count) | Returns the number of items in the list. |
| [Insert](#insert) | Inserts an item at the specific index. |
| [Item](#item) | Retrieves the item at the specified index. |
| [Remove](#remove) | Removes the specified item from the list. |

---

## <a name="add"></a>Add (DWStrList)

Appends a string to the list.

```
FUNCTION Add (BYREF dws AS DWSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *dws* | The string to append. |

#### Return value

Returns the number of items in the list.

---

## <a name="clear"></a>Clear (DWStrList)

Empties the entire list.

```
SUB Clear
```
---

## <a name="count"></a>Count (DWStrList)

Returns the number of items in the list.

```
FUNCTION Count () AS LONG
```

#### Return value

Returns the number of items in the list.

---

## <a name="insert"></a>Insert (DWStrList)

Inserts an item at the specific index.

```
FUNCTION Insert (BYVAL idx AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *idx* | The one-based index where to insert the item. |

#### Return value

If the method succeeds, it returns TRUE; otherwise, FALSE.

---
---

## <a name="item"></a>Item (DWStrList)

Retrieves the item at the specified index.

```
FUNCTION Item (BYVAL idx AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *idx* | The one-based index of the item to retrieve. |

#### Return value

If the method succeeds, it returns TRUE; otherwise, FALSE.

---
---

## <a name="remove"></a>Remove (DWStrList)

Removes the specified item from the list. Indexes are one-based.

```
FUNCTION Remove (BYVAL idx AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *idx* | The one-based index of the item to remove. |

#### Return value

If the method succeeds, it returns TRUE; otherwise, FALSE.

---
