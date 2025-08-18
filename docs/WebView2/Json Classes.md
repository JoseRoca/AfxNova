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

