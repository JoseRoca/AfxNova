Json components overview
Scope: UTF‑16–native JSON reader/writer plus a string unquoter, designed to interop cleanly with BSTR/DWSTRING, DVARIANT, and DSafeArray.

Philosophy: Zero dependencies, COM‑friendly types, predictable pretty‑printing, and safe handling of escapes/control chars.

#### JsonUnquoteW function
Signature: PRIVATE FUNCTION JSonUnquoteW(ByRef wszJson As WSTRING) As DWSTRING

Purpose: Decode a JSON string literal into a DWSTRING, resolving escape sequences and control codes.

Input/contract: Expects a JSON‑encoded string surrounded by quotes. Returns empty if the input isn’t a quoted JSON string.

Decoding rules: Handles \" \\ \/ \b \f \n \r \t and \uXXXX (one UTF‑16 code unit per escape). Surrogate pairs pass through correctly when present as consecutive \uXXXX sequences.

Typical use: Consume WebView2 ExecuteScript string results (which arrive as JSON strings) and recover the plain text.

#### JsonReader class
Role: Tokenize a UTF‑16 JSON text into a stream of JsonToken entries without allocating a DOM.

State: m_buf (DWSTRING buffer), m_pos (1‑based cursor).

Constructor and destructor
Constructor: PRIVATE CONSTRUCTOR JsonReader(ByRef source As WSTRING)

Purpose: Initialize the reader with source text and reset position to the start.

Destructor: PRIVATE DESTRUCTOR JsonReader

Purpose: No dynamic resources; present for symmetry.

#### SkipWhitespace
Signature: PRIVATE SUB JsonReader.SkipWhitespace

Purpose: Advance m_pos past spaces, tabs, CR, and LF.

Side effect: Updates m_pos; no tokens produced.

#### ReadString
Signature: PRIVATE FUNCTION JsonReader.ReadString() As DWSTRING

Purpose: Consume a JSON string literal starting at the opening quote and return its decoded text.

Parsing detail: Walks raw characters, skips over escapes (including \uXXXX), slices the raw quoted segment, then calls JsonUnquoteW to decode.

#### ReadNumber
Signature: PRIVATE FUNCTION JsonReader.ReadNumber() As DWSTRING

Purpose: Parse a JSON number per the grammar: optional sign, integer, optional fraction, optional exponent.

Output form: Returns the exact textual representation captured (as DWSTRING).

Robustness: If no digits are consumed, resets to start and returns empty. If an exponent marker lacks digits, truncates back to pre‑exponent.

ReadNext
Signature: PRIVATE FUNCTION JsonReader.ReadNext(ByRef tok As JsonToken) As Boolean

Purpose: Produce the next token and advance the cursor.

Tokenization:
Structural: { } [ ] : , mapped to JSON_OBJECT_START/END, JSON_ARRAY_START/END, JSON_COLON, JSON_COMMA.

Values: JSON_STRING via ReadString, JSON_NUMBER via ReadNumber, JSON_BOOL (“true”/“false”), JSON_NULL (“null”).

End/invalid handling: Returns FALSE at end of buffer. On unrecognized input, sets JSON_NONE and fast‑forwards to avoid infinite loops.

#### JsonWriter class
Role: Build JSON text into an internal BSTRING buffer with optional pretty‑printing and smart inline formatting.

State: buf (BSTRING), firstItemStack() (comma/first‑element tracking per nesting level), indentSize, depth, inlineThreshold.

Configuration

SetIndentSize(ByVal n As LONG):

Purpose: Control spaces per indent level (0 disables pretty‑printing newlines/indents for structural entries).

Constraints: Non‑negative values only.

SetInlineThreshold(ByVal n As LONG):

Purpose: Set the maximum character length used to decide if arrays/objects may be emitted inline.

Effect: Influences IsSmallInline decisions for array serialization.

Structure control
BeginObject():

Purpose: Emit “{”, increase depth, push a “first item” flag, and insert newline+indent.

Comma logic: Respects AppendCommaIfNeeded before emitting.

EndObject():

Purpose: Close the object with proper dedent and “}”, then pop the “first item” stack.

BeginArray():
Purpose: Emit “[”, increase depth, push a “first item” flag, and insert newline+indent.

EndArray():

Purpose: Close the array with proper dedent and “]”, then pop the stack.

Name(ByRef s As WString):

Purpose: Write a JSON object key with correct escaping, colon, and spacing.

Effect: Resets the current level’s “first item” flag so the next value decides comma/newline.

Primitive values
Value(ByRef s As WString):
Purpose: Emit a JSON string with escaping for quotes, backslashes, control chars, and nonprintables via \uXXXX.

Value(ByVal n As LongInt):
Purpose: Emit an integer number using WStr conversion (no quotes).

Value(ByVal n As Double):
Purpose: Emit a floating‑point number using WStr conversion (no quotes).

ValueNull():
Purpose: Emit literal null, respecting comma/indent rules.

ValueBool(ByVal b As Boolean):
Purpose: Emit true or false.

Composite values and COM interop
ValueVariant(ByRef dv As DVARIANT):

Purpose: Serialize a DVARIANT using JSON‑compatible mapping.

Mapping:
Empty/Null: null
Boolean: true/false
BSTR: string (escaped)
R4/R8: number (double)
Integral types (signed/unsigned): number (integer)
Arrays (common VT_ARRAY types): delegates to ValueSafeArray
Other: dv.ToStr() as string fallback
ValueSafeArray(ByRef sa As DSafeArray):

Purpose: Serialize a 1‑D SAFEARRAY as JSON array with smart inline vs pretty layout.

Inline decision: Renders a compact temporary “[a,b,c]” form and uses IsSmallInline(temp) to choose inline or multi‑line emission.

Element handling: Escapes VT_BSTR directly; otherwise serializes each element with ValueVariant.

Output and buffer management
ToBString() As BSTRING:
Purpose: Retrieve the current buffer as BSTRING (UTF‑16).

ToUtf8() As STRING:
Purpose: Retrieve a UTF‑8 copy, handy for file I/O or engines expecting UTF‑8.

Clear():
Purpose: Reset buffer, depth, and firstItemStack to an empty state.

Private helpers
IsSmallInline(ByRef tmp As String) As Boolean:

Purpose: Heuristic for one‑line emission; TRUE when length ≤ inlineThreshold and no “{” or “[”.

AppendEscaped(ByRef s As WString):
Purpose: Append a quoted, JSON‑escaped form of s into buf, encoding control chars < 32 as \uXXXX.

AppendCommaIfNeeded():
Purpose: Manage comma insertion between sibling elements based on firstItemStack; injects newline+indent for pretty‑print paths.

AppendNewlineAndIndent():
Purpose: Add line break and depth×indentSize spaces.

Notes and tips
WebView2 bridge: Use ToBString() with PostWebMessageAsJson for native→JS, and JSonUnquoteW for ExecuteScript string returns.

Pretty vs compact: SetIndentSize(0) for compact output; rely on inlineThreshold to keep small arrays inline even with pretty‑print enabled.

Robust reading: JsonReader is a tokenizer, not a full validator. Guard your consumer against JSON_NONE to handle malformed input.

Unicode fidelity: Writer escapes control characters and preserves BMP characters directly; astral plane characters in input DWSTRING are emitted as UTF‑16 code units and will be round‑tripped correctly by modern JS engines.

' ========================================================================================
' Json test
' ========================================================================================
''#CONSOLE ON
'#define UNICODE
'#define _WIN32_WINNT &h0602
'#include once "AfxNova/AfxJson.inc"
'USING AfxNova

'SUB TestJsonReader()
'   ' Craft some JSON with tricky bits:
'   ' - Unicode \u00E9 (é) and surrogate pair \uD83D\uDE03 (??)
'   ' - Quotes, backslashes, slash, control codes
'   ' - Numbers with exponents, booleans, null
'   DIM sample AS DWSTRING = _
'       "{""name"":""Jos" & WCHR(&h00E9) & " " & WCHR(&hD83D) & WCHR(&hDE03) & """, " & _
'       """quote"":""He said """"Hello/World""""!"", " & _
'       """newline"":""Line1" & WCHR(10) & "Line2"", " & _
'       """pi"":3.14159, ""exp"":-2.5e+3, " & _
'       """ok"":true, ""missing"":null}"

'   DIM rdr AS JsonReader = JsonReader(sample)
'   DIM tok AS JsonToken
'   DIM idx AS INTEGER = 0

'   PRINT "Testing JSON Reader with JSonUnquoteW-backed ReadString"
'   PRINT STRING(60,"-")

'   WHILE rdr.ReadNext(tok)
'      idx += 1
'      PRINT idx; ". "; 
'      SELECT CASE tok.kind
'         CASE JSON_STRING
'            PRINT "STRING: "; tok.value
'         CASE JSON_NUMBER
'            PRINT "NUMBER: "; tok.value
'         CASE JSON_BOOL
'            PRINT "BOOL: "; tok.value
'         CASE JSON_NULL
'            PRINT "NULL"
'         CASE JSON_OBJECT_START
'            PRINT "{"
'         CASE JSON_OBJECT_END
'            PRINT "}"
'         CASE JSON_ARRAY_START
'            PRINT "["
'         CASE JSON_ARRAY_END
'            PRINT "]"
'         CASE JSON_COLON
'            PRINT ":"
'         CASE JSON_COMMA
'            PRINT ","
'         CASE ELSE
'            PRINT "UNKNOWN"
'      END SELECT
'   WEND

'   PRINT STRING(60,"-")
'   PRINT "Done."
'END SUB

'TestJsonReader
' ========================================================================================

```
' ========================================================================================
' SAFEARRAY serialization with inline suppression.
' Example:
'   DIM dvArr AS DSafeArray
'   dvArr.Create(VT_VARIANT, 3, 0)
'   dvArr.PutVar(0, DVARIANT("first ??"))
'   dvArr.PutVar(1, DVARIANT(42))
'   dvArr.PutVar(2, DVARIANT(3.14159))
'   DIM jw AS JsonWriter
'   jw.SetIndentSize(2) ' 2-space indent
'   jw.BeginObject()
'      jw.Name("title") : jw.Value("Test")
'      jw.Name("data")  : jw.ValueSafeArray(dvArr)
'   jw.EndObject()
'   PRINT jw.ToBString()
' Example:
'   DIM tiny AS DSafeArray
'   tiny.Create(VT_VARIANT, 3, 0)
'   tiny.PutVar(0, DVARIANT(1))
'   tiny.PutVar(1, DVARIANT(2))
'   tiny.PutVar(2, DVARIANT(3))
'   DIM jw As JsonWriter
'   jw.SetIndentSize(2)
'   jw.SetInlineThreshold(20)  ' tighten threshold
'   jw.BeginObject()
'      jw.Name("a") : jw.ValueSafeArray(tiny)
'      jw.Name("b") : jw.Value("longer text here will break inline")
'  jw.EndObject()
' Output:
' {
'  "a": [
'    1,
'    2,
'    3
'  ],
'  "b": "longer text here will break inline"
'}
' ========================================================================================

```
' // More examples:
' Primitives + string escaping
' DIM esc AS DWSTRING = "Quote: """ & " Newline: " & WCHR(10) & "Tab: " & WCHR(9) & "End"
' DIM jw AS JsonWriter
' jw.SetIndentSize(2)
' jw.BeginObject()
'   jw.Name("s")    : jw.Value(esc)
'   jw.Name("i64")  : jw.Value(9223372036854775807)  ' max signed 64
'   jw.Name("pi")   : jw.Value(3.141592653589793)
'   jw.Name("ok")   : jw.ValueVariant(DVARIANT(TRUE, "BOOLEAN"))   ' VT_BOOL ? true
'   jw.Name("none") : jw.ValueVariant(DVARIANT())                  ' VT_EMPTY ? null
' jw.EndObject()
' print jw.ToUtf8()

' Unicode text (BMP): accents, em dash, Greek
' DIM uni AS DWSTRING = "España — café — ?x"
' DIM jw AS JsonWriter
' jw.SetIndentSize(2)
' jw.BeginObject()
'   jw.Name("title") : jw.Value("Unicode test")
'   jw.Name("text")  : jw.Value(uni)
' jw.EndObject()
' AfxMsg jw.ToBString

'Dim jw As JsonWriter
'jw.SetIndentSize(2)
'jw.BeginObject()
'  jw.Name("app") : jw.Value("AfxNova")
'  jw.Name("version") : jw.Value(1)
'  jw.Name("author")
'  jw.BeginObject()
'    jw.Name("name")  : jw.Value("Jose Roca")
'    jw.Name("site")  : jw.Value("https://github.com/JoseRoca/AfxNova")
'    jw.Name("active"): jw.ValueVariant(DVARIANT(TRUE, "BOOLEAN"))
'  jw.EndObject()
'  jw.Name("notes") : jw.Value("All UTF-16 internally")
'jw.EndObject()
'print jw.ToBString

'DIM jw AS JsonWriter
'jw.SetIndentSize(0) ' 0 = no pretty-print
'jw.BeginObject()
'  jw.Name("k")  : jw.Value("v")
'  jw.Name("n")  : jw.Value(123)
'  jw.Name("b")  : jw.ValueVariant(DVARIANT(FALSE, "BOOLEAN"))
'  jw.Name("nil"): jw.ValueVariant(DVARIANT())
'jw.EndObject()
'print jw.ToBString
```

```

Exactly — they’re built first and foremost to make life easier when shuttling data in and out of WebView2, but because they stick to clean JSON in/out and COM‑friendly types, they’re essentially drop‑in utilities anywhere you need structured text parsing or emission.

That means you could:

Feed them API responses from a local service or a remote REST endpoint

Serialize/deserialize data between PowerBASIC apps without pulling in heavier libraries

Store application state in a readable format for logs or config files

Transform Excel/Access/other COM‑capable app data into JSON for reporting or exchange

It’s like having a screwdriver that also happens to pry open paint cans — its design target is WebView2, but the solid core makes it handy in a lot of other jobs.
