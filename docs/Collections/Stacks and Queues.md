# CStack Class

A `Stack` collection is an ordered set of data items, which are accessed on a LIFO (Last-In / First-Out) basis. Each data item is passed and stored as a variant variable, using the **Push** and **Pop** methods.

**Include file**: AfxNova/CStack.inc

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Push](#push) | Appends a variant at the end of the collection. |
| [Pop](#pop) | Gets and removes the last element of the collection. |
| [Count](#count) | Returns the number of elements in the collection. |
| [Clear](#clear) | Removes all the elements in the collection. |

---

# CQueue Class

A Queue` collection is an ordered set of data items, which are accessed on a FIFO (First-In / First-Out) basis. Each data item is passed and stored as a variant variable, using the **Enqueue** and **Dequeue** methods.

**Include file**: AfxNova/CStack.inc

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Enqueue](#enqueue) | Appends a variant at the end of the collection. |
| [Dequeue](#dequeue) | Gets and removes the first element of the collection. |
| [Count](#count2) | Returns the number of elements in the collection. |
| [Clear](#clear2) | Removes all the elements in the collection. |

---

## <a name="push"></a>Push (CStack)

Appends a variant at the end of the collection.

```
FUNCTION Push (BYREF cv AS CVAR) AS HRESULT
```

#### Return value

Returns S_OK on success, or an error HRESULT on failure.
Error DISP_E_ARRAYISLOCKED: The array is locked.

#### Usage example:

```
#INCLUDE ONCE "AfxNova/CStack.inc"
USING AfxNova

DIM pStack AS CStack
pStack.Push "String 1"
pStack.Push "String 2"
DIM dv AS DVARIANT
dv = pStack.Pop
PRINT dv
dv = pStack.Pop
PRINT dv

' --or--

PRINT pStack.Pop
PRINT pStack.Pop
```
---

## <a name="pop"></a>Pop (CStack)

Gets and removes the last element of the collection.

```
FUNCTION Pop () AS DVARIANT
```

#### Return value

A Variant. If there are no elements to pop, the returned variant will be empty.

---

## <a name="count"></a>Count (CStack)

Returns the number of elements in the collection.

```
FUNCTION Count () AS UINT
```
---

## <a name="clear"></a>Clear (CStack)

Returns the number of elements in the collection.

```
FUNCTION Clear () AS HRESULT
```

#### Return value

Returns S_OK on success, or an error HRESULT on failure.

---

## <a name="enqueue"></a>Enqueue (CQueue)

Appends a variant at the end of the collection.

```
FUNCTION Enqueue (BYREF cv AS CVAR) AS HRESULT
```

#### Return value

Returns S_OK on success, or an error HRESULT on failure.

Error DISP_E_ARRAYISLOCKED: The array is locked.

#### Usage example:

```
#INCLUDE ONCE "AfxNova/CStack.inc"
USING AfxNova

DIM pQueue AS CQueue
pQueue.Enqueue "String 1"
pQueue.Enqueue 12345.12
print pQueue.Dequeue
print pQueue.Dequeue
```
---

## <a name="dequeue"></a>Dequeue (CQueue)

Gets and removes the first element of the collection.

```
FUNCTION Dequeue () AS CVAR
```

#### Return value

A Variant. If there are no elements to dequeue, the returned variant will be empty.

---

## <a name="count2"></a>Count (CQueue)

Returns the number of elements in the collection.

```
FUNCTION Count () AS UINT
```
---

## <a name="clear2"></a>Clear (CQueue)

Removes all the elements in the collection.

```
FUNCTION Clear () AS HRESULT
```
#### Return value

Returns S_OK on success, or an error HRESULT on failure.

---
