# JSON Reader and Writer classes

UTF‑16–native JSON reader/writer plus a string unquoter, designed to interop cleanly with BSTR/DWSTRING, DVARIANT, and DSafeArray.

Zero dependencies, COM‑friendly types, predictable pretty‑printing, and safe handling of escapes/control chars.

## JsonUnquoteW function

Decodes a JSON string literal into a DWSTRING, resolving escape sequences and control codes.

```
FUNCTION JSonUnquoteW (BYREF wszJson AS WSTRING) AS DWSTRING
```
| Parameter  | Description |
| ---------- | ----------- |
| *wszJson* | A JSON‑encoded string surrounded by quotes. |

#### Return value

Input/contract: Expects a JSON‑encoded string surrounded by quotes. Returns empty if the input isn’t a quoted JSON string.

Decoding rules: Handles \" \\ \/ \b \f \n \r \t and \uXXXX (one UTF‑16 code unit per escape). Surrogate pairs pass through correctly when present as consecutive \uXXXX sequences.

Typical use: Consume `WebView2` **ExecuteScript** string results (which arrive as JSON strings) and recover the plain text.

---

# JsonReader class

Tokenizes a UTF‑16 JSON text into a stream of **JsonToken** entries without allocating a DOM.

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructor1) | Creates a new **JsonReader** object initialized to the specified value. |
| [ReadNext](#readxext) | Reads the next token. |
| [ReadNumber](#readumber) | Parses a JSON number. |
| [ReadString](#readstring) | Reads a string by slicing the raw JSON string and unquoting via **JSonUnquoteW**. |
| [SkipWhitespace](#skipwhitespace) | Advances the position past spaces, tabs, CR and LF. |

---

# JsonWriter class

Builds JSON text into an internal `BSTRING` buffer with optional pretty‑printing and smart inline formatting.

| Name       | Description |
| ---------- | ----------- |
| [BeginArray](#beginarray) | Emits “[”, increases depth, pushes a “first item” flag, and inserts newline+indent. |
| [BeginObject](#beginobject) | Structures control with inline suppression. |
| [Clear](#clear) | Resets the buffer, depth, and firstItemStack to an empty state. |
| [EndArray](#endarray) | Closes the array with proper dedent and “]”, then pops the stack. |
| [EndObject](#endobject) | Closes the object with proper dedent and “}”. |
| [Name](#name) | Writes a JSON object key with correct escaping, colon, and spacing. |
| [SetIndentSize](#setindentsize) | Control spaces per indent level. |
| [SetInlineThreshold](#setinlinethreshold) | Control spaces per indent level. |
| [ToBString](#tobstring) | Returns the current buffer as BSTRING (UTF-16). |
| [ToUtf8](#toutf8) | Returns the current buffer as a utf-8 string. |
| [Value](#value) | Emits a JSON string with escaping for quotes, backslashes, control chars, and nonprintables via \uXXXX. |
| [ValueBool](#valuebool) | Emits a boolean true or false value. |
| [ValueNull](#valuenull) | Emits a literal null, respecting comma/indent rules. |
| [ValueSafeArray](#valuesafearray) | Serializes a 1‑D SAFEARRAY as JSON array with smart inline vs pretty layout. |
| [ValueVariant](#valuevariant) | Serializes a `DVARIANT` using JSON‑compatible mapping. |

---

## <a name="constructors1"></a>Constructor (JsonReader)

Creates a new **JsonReader** object initialized to the specified value.

```
CONSTRUCTOR JsonReader (BYREF source AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *source* | A JSON‑encoded string surrounded by quotes. |

---

#### Example

```
' ========================================================================================
' Json test
' ========================================================================================
'#CONSOLE ON
#define UNICODE
#define _WIN32_WINNT &h0602
#include once "AfxNova/AfxJson.inc"
USING AfxNova

SUB TestJsonReader()
   ' Craft some JSON with tricky bits:
   ' - Unicode \u00E9 (é) and surrogate pair \uD83D\uDE03 (??)
   ' - Quotes, backslashes, slash, control codes
   ' - Numbers with exponents, booleans, null
   DIM sample AS DWSTRING = _
       "{""name"":""Jos" & WCHR(&h00E9) & " " & WCHR(&hD83D) & WCHR(&hDE03) & """, " & _
       """quote"":""He said """"Hello/World""""!"", " & _
       """newline"":""Line1" & WCHR(10) & "Line2"", " & _
       """pi"":3.14159, ""exp"":-2.5e+3, " & _
       """ok"":true, ""missing"":null}"

   DIM rdr AS JsonReader = JsonReader(sample)
   DIM tok AS JsonToken
   DIM idx AS INTEGER = 0

   PRINT "Testing JSON Reader with JSonUnquoteW-backed ReadString"
   PRINT STRING(60,"-")

   WHILE rdr.ReadNext(tok)
      idx += 1
      PRINT idx; ". "; 
      SELECT CASE tok.kind
         CASE JSON_STRING
            PRINT "STRING: "; tok.value
         CASE JSON_NUMBER
            PRINT "NUMBER: "; tok.value
         CASE JSON_BOOL
            PRINT "BOOL: "; tok.value
         CASE JSON_NULL
            PRINT "NULL"
         CASE JSON_OBJECT_START
            PRINT "{"
         CASE JSON_OBJECT_END
            PRINT "}"
         CASE JSON_ARRAY_START
            PRINT "["
         CASE JSON_ARRAY_END
            PRINT "]"
         CASE JSON_COLON
            PRINT ":"
         CASE JSON_COMMA
            PRINT ","
         CASE ELSE
            PRINT "UNKNOWN"
      END SELECT
   WEND

   PRINT STRING(60,"-")
   PRINT "Done."
END SUB

TestJsonReader
' ========================================================================================
```

## SkipWhitespace

Advances the position past spaces, tabs, CR, and LF.

```
SUB JsonReader.SkipWhitespace
```
---

## ReadString

Reads a JSON string literal starting at the opening quote and returns its decoded text.

```
FUNCTION ReadString () AS DWSTRING
```

Walks raw characters, skips over escapes (including \uXXXX), slices the raw quoted segment, then calls JsonUnquoteW to decode.

---

## ReadNumber

Parses a JSON number checking for optional sign, integer, optional fraction and optional exponent.

```
FUNCTION ReadNumber () AS DWSTRING
```

Returns the exact textual representation captured.

If no digits are consumed, resets to start and returns empty. If an exponent marker lacks digits, truncates back to pre‑exponent.

---

## ReadNext

Produces the next token and advances the position.

```
FUNCTION ReadNext (BYREF tok AS JsonToken) AS BOOLEAN
```

Tokenization:

Structural: { } [ ] : , mapped to JSON_OBJECT_START/END, JSON_ARRAY_START/END, JSON_COLON, JSON_COMMA.

Values: JSON_STRING via ReadString, JSON_NUMBER via ReadNumber, JSON_BOOL (“true”/“false”), JSON_NULL (“null”).

End/invalid handling: Returns FALSE at end of buffer. On unrecognized input, sets JSON_NONE and fast‑forwards to avoid infinite loops.

---
---

## SetIndentSize

Sets the maximum character length used to decide if arrays/objects may be emitted inline.

```
SUB SetIndentSize(BYVAL n As LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *n* | The number of spaces to indent. |

---

## SetInlineThreshold

Sets the maximum character length used to decide if arrays/objects may be emitted inline.

```
SUB SetInlineThreshold(BYVAL n As LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *n* | The maximum character length. |

---

## BeginObject

Structures control with inline suppression.

```
SUB BeginObject
```
Purpose: Emits “{”, increases depth, pushes a “first item” flag, and inserts newline+indent.

#### Examples

Primitives + string escaping
```
#include once "AfxNova/AfxJson.inc"
USING AfxNova

DIM esc AS DWSTRING = "Quote: """ & " Newline: " & WCHR(10) & "Tab: " & WCHR(9) & "End"
DIM jw AS JsonWriter
jw.SetIndentSize(2)
jw.BeginObject()
  jw.Name("s")    : jw.Value(esc)
  jw.Name("i64")  : jw.Value(9223372036854775807)  ' max signed 64
  jw.Name("pi")   : jw.Value(3.141592653589793)
  jw.Name("ok")   : jw.ValueVariant(DVARIANT(TRUE, "BOOLEAN"))   ' VT_BOOL -> true
  jw.Name("none") : jw.ValueVariant(DVARIANT())                  ' VT_EMPTY -> null
jw.EndObject()
print jw.ToUtf8
```

Nested objects

```
#include once "AfxNova/AfxJson.inc"
USING AfxNova

Dim jw As JsonWriter
jw.SetIndentSize(2)
jw.BeginObject()
  jw.Name("app") : jw.Value("AfxNova")
  jw.Name("version") : jw.Value(1)
  jw.Name("author")
  jw.BeginObject()
    jw.Name("name")  : jw.Value("José Roca")
    jw.Name("site")  : jw.Value("https://github.com/JoseRoca/AfxNova")
    jw.Name("active"): jw.ValueVariant(DVARIANT(TRUE, "BOOLEAN"))
  jw.EndObject()
  jw.Name("notes") : jw.Value("All UTF-16 internally")
jw.EndObject()
print jw.ToBString
```
---

## EndObject

Closes the object with proper dedent and “}”, then pops the "first item" stack.

```
SUB EndObject
```
---

## BeginArray

Emits “[”, increases depth, pushes a “first item” flag, and inserts newline+indent.

```
SUB BeginArray
```
---

## EndArray

Closes the array with proper dedent and “]”, then pops the stack.

```
SUB EndArray
```
---

## Name

Writes a JSON object key with correct escaping, colon, and spacing.

```
SUB Name (BYREF s AS WSTRING)
```

| Parameter  | Description |
| ---------- | ----------- |
| *s* | The JSON string. |

Resets the current level’s "first item" flag so the next value decides comma/newline.

---

## Value

Emits a JSON string with escaping for quotes, backslashes, control chars, and nonprintables via \uXXXX.

```
SUB Value (BYREF s AS WSTRING)
SUB Value (BYVAL n AS LONGINT)
SUB Value (BYVAL n AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *s* | The string to emit. |
| *n* | The number to emit. |

The numbers are converted to string with **WSTR**.

#### Examples

Unicode text (BMP): accents, em dash, Greek

```
#include once "AfxNova/AfxJson.inc"
USING AfxNova

DIM uni AS DWSTRING = "España — café — ?x"
DIM jw AS JsonWriter
jw.SetIndentSize(2)
jw.BeginObject()
  jw.Name("title") : jw.Value("Unicode test")
  jw.Name("text")  : jw.Value(uni)
jw.EndObject()
AfxMsg jw.ToBString
```

## ValueNull

Emits a literal null, respecting comma/indent rules.

```
SUB ValueNull
```
---

## ValueBool

Emits a boolean true or false value.

```
SUB ValueBool (BYVAL b AS BOOLEAN)
```

| Parameter  | Description |
| ---------- | ----------- |
| *b* | A boolean true(-1) or false(0) value. |

---

## ValueVariant

Serializes a `DVARIANT` using JSON‑compatible mapping.

```
SUB ValueVariant (BYREF dv AS DVARIANT)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dv | a variant value. |

^^Mapping** :

Empty/Null: null

Boolean: true/false

BSTR: string (escaped)

R4/R8: number (double)

Integral types (signed/unsigned): number (integer)

Arrays (common VT_ARRAY types): delegates to ValueSafeArray

Other: dv.ToStr() as string fallback

#### Example

```
#include once "AfxNova/AfxJson.inc"
USING AfxNova

DIM jw AS JsonWriter
jw.SetIndentSize(0) ' 0 = no pretty-print
jw.BeginObject()
  jw.Name("k")  : jw.Value("v")
  jw.Name("n")  : jw.Value(123)
  jw.Name("b")  : jw.ValueVariant(DVARIANT(FALSE, "BOOLEAN"))
  jw.Name("nil"): jw.ValueVariant(DVARIANT())
jw.EndObject()
print jw.ToBString
```
---

## ValueSafeArray

Serializes a 1‑D SAFEARRAY as JSON array with smart inline vs pretty layout.

```
SUB ValueSafeArray (BYREF sa AS DSafeArray)
```

Escapes VT_BSTR directly; otherwise serializes each element with **ValueVariant**.

#### Example

SAFEARRAY serialization with inline suppression.

```
#include once "AfxNova/AfxJson.inc"
USING AfxNova

DIM dvArr AS DSafeArray
dvArr.Create(VT_VARIANT, 3, 0)
dvArr.PutVar(0, DVARIANT("first ??"))
dvArr.PutVar(1, DVARIANT(42))
dvArr.PutVar(2, DVARIANT(3.14159))
DIM jw AS JsonWriter
jw.SetIndentSize(2) ' 2-space indent
jw.BeginObject()
  jw.Name("title") : jw.Value("Test")
  jw.Name("data")  : jw.ValueSafeArray(dvArr)
jw.EndObject()
PRINT jw.ToBString()

' Outout:
{
  "title": "Test",
  "data": [
    "first",
    42,
    3.14159
  ]
}
```
```
#include once "AfxNova/AfxJson.inc"
USING AfxNova

DIM tiny AS DSafeArray
tiny.Create(VT_VARIANT, 3, 0)
tiny.PutVar(0, DVARIANT(1))
tiny.PutVar(1, DVARIANT(2))
tiny.PutVar(2, DVARIANT(3))
DIM jw As JsonWriter
jw.SetIndentSize(2)
jw.SetInlineThreshold(20)  ' tighten threshold
jw.BeginObject()
   jw.Name("a") : jw.ValueSafeArray(tiny)
   jw.Name("b") : jw.Value("longer text here will break inline")
jw.EndObject()

Output:
{
 "a": [
   1,
   2,
   3
 ],
 "b": "longer text here will break inline"
}
```
---

## ToBString

Returns the current buffer as a BSTRING (UTF‑16).

```
FUNCTION ToBString () AS BSTRING
```

---

## ToUtf8

Returns the current buffer as a UTF-8 string.

```
FUNCTION ToUtf8 () AS STRING
```
---

---

## Clear

Resets the buffer, depth, and m_firstItemStack to an empty state.

```
SUB Clear
```
---

#### Notes and tips

WebView2 bridge: Use **ToBString** with **PostWebMessageAsJson** for native→JS, and **JSonUnquoteW** for **ExecuteScript** string returns.

Pretty vs compact: **SetIndentSize(0)** for compact output; rely on *m_inlineThreshold* to keep small arrays inline even with pretty‑print enabled.

Robust reading: **JsonReader** is a tokenizer, not a full validator. Guard your consumer against JSON_NONE to handle malformed input.

Unicode fidelity: Writer escapes control characters and preserves BMP characters directly; astral plane characters in input strings are emitted as UTF‑16 code units and will be round‑tripped correctly by modern JS engines.

They’re built first and foremost to make life easier when shuttling data in and out of WebView2, but because they stick to clean JSON in/out and COM‑friendly types, they’re essentially drop‑in utilities anywhere you need structured text parsing or emission.

That means you could:

Feed them API responses from a local service or a remote REST endpoint

Serialize/deserialize data between FreeBASIC apps without pulling in heavier libraries

Store application state in a readable format for logs or config files

Transform Excel/Access/other COM‑capable app data into JSON for reporting or exchange

---
