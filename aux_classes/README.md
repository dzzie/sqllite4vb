# cSQLiteTable and friends

A small set of VB6 auxiliary classes that sit on top of David Zimmer's
[`cSQLite`](http://sandsprite.com) library. Where the base library gives you
live cursors and statement handles, these add a detached snapshot container
with JSON / CSV / SQL round-tripping, schema introspection, and ad-hoc data
construction.

## What problem this solves

`cSQLiteResults` is a live cursor — bound to a connection, forward-only,
holds a statement handle that must be finalized. Useful for streaming
through results, but awkward when you want to:

- Hand a result set to another part of the program after the connection
  closes
- Iterate the same data more than once
- Serialize results to disk and reload them later
- Build a table from CSV or JSON with no database involved
- Feed a `MSFlexGrid` from arbitrary tabular data
- Show an AI a result set, let it propose changes, and generate the SQL
  to apply them

`cSQLiteTable` is the answer: a detached snapshot — column names, column
type hints, and a 2D `Variant` block of values. No statement, no
connection, no iteration state.

## Files

| File                            | Type           | Role                                                   |
| ------------------------------- | -------------- | ------------------------------------------------------ |
| `cSQLiteTable.cls`              | Class          | The container itself                                   |
| `cSQLiteField.cls`              | Class          | Result-set column descriptor (returned by `GetSchema`) |
| `JsonParser_cSQLiteTable.cls`   | Class, private | Recursive-descent JSON parser used by `LoadJson`       |
| `modCsvParser.bas`              | Module         | RFC-4180-ish CSV parser used by `LoadCsv`              |
| `modSQLiteTypes.bas`            | Module         | `SQLiteType` enum + `SqlTypeName` helper               |

For an ActiveX DLL build, `JsonParser_cSQLiteTable` is set to `Private`
instancing so it doesn't leak out of the project type library. The CSV
parser is a `.bas` module since it's stateless. The descriptor classes
(`cSQLiteField`) use `Friend Sub SetFrom` for construction so external
callers can't `New` them directly — they only come out of `GetSchema()`.

## Quick start

```vb
Dim t As New cSQLiteTable

' --- From a live db ---
t.LoadQuery db, "SELECT name, age FROM users WHERE age > ?", 30

' --- From a CSV file (no db needed) ---
t.LoadCsvFile "C:\data\users.csv"

' --- Hand-built (no db, no file) ---
t.SetColumns Array("name", "age")
t.AddRow Array("alice", 30)
t.AddRow Array("bob",   45)
```

Once loaded, every path produces the same kind of object:

```vb
Debug.Print t.RowCount & " x " & t.ColumnCount

' Cell access - default property takes (row, col) where col is index or name
Debug.Print t(0, "name"), t(0, 1)
Debug.Print t.Value(0, "name")    ' explicit-name alias

' Whole row or column
Dim row As Variant
row = t.Row(0)                    ' 1D Variant array
Dim col As Variant
col = t.Column("age")             ' 1D Variant array
```

## Schema introspection

```vb
Dim f As cSQLiteField
For Each f In t.GetSchema()
    Debug.Print f.Ordinal, f.Name, f.TypeName
Next
```

Each `cSQLiteField` exposes `Ordinal`, `Name`, `SqlType` (the `SQLiteType`
enum value), and `TypeName` (friendly string: `"INTEGER"`, `"TEXT"`, ...).

Types come from `sqlite3_column_type()` when the data was loaded from a
live cursor, from the JSON file's metadata when the columnar shape was
used, or from row-sniffing for CSV and the portable JSON shape. Columns
that are NULL in every row fall back to `BLOB` affinity (the loosest in
SQLite).

## I/O surface

|       | Read in                                  | Write out                  |
| ----- | ---------------------------------------- | -------------------------- |
| Live  | `LoadFromResults`, `LoadQuery`           | —                          |
| JSON  | `LoadJson`, `LoadJsonFile`               | `ToJson`, `SaveJsonFile`   |
| CSV   | `LoadCsv`, `LoadCsvFile`                 | `ToCsv`, `SaveCsvFile`     |
| SQL   | —                                        | `ToSql`, `SaveSqlFile`     |
| Hand  | `SetColumns` + `AddRow`                  | —                          |
| Grid  | —                                        | `FillFlexGrid`             |

No `FromSql` — loading from SQL means executing it, which means a live
db, at which point `LoadQuery` is the right answer.

### JSON format

Columnar with explicit type metadata, designed to round-trip cleanly:

```json
{
  "version": 1,
  "columns": [
    {"name": "id",   "type": 1},
    {"name": "name", "type": 3},
    {"name": "data", "type": 4}
  ],
  "rows": [
    [1, "alice", {"$blob": "48656c6c6f"}],
    [2, null,    null]
  ]
}
```

BLOBs serialize as `{"$blob": "<lowercase hex>"}`. NULLs as JSON `null`.
Strings are properly escaped (control chars, non-ASCII, embedded quotes,
backslashes — all the things that bite naive escapers).

`LoadJson` also accepts a top-level array of objects (the "portable"
shape) for interop with other tools that emit JSON that way.

### CSV

`LoadCsv` handles RFC-4180-ish: quoted fields with embedded delimiters,
doubled quotes, embedded newlines, mixed quoted/unquoted on one row,
mixed line endings, optional trailing newline, UTF-8 BOM.

Light type inference is on by default:

- empty → `Null`
- pure digits (optional leading `-`) → `Long`, with overflow fallback to `Double`
- numeric-looking strings matching `[+-]? digit+ ( . digit+ )? ( [eE] [+-]? digit+ )?` → `Double`
- everything else → `String`

**No date parsing.** Locale and format ambiguity makes it more dangerous
than useful. `"2024-01-15"` stays a `String`.

Pass `typed:=False` to disable inference and keep every cell as a String.

### SQL

`ToSql(tableName)` produces a self-contained script: optional
`DROP TABLE IF EXISTS`, then `CREATE TABLE` with column types inferred
from data, then one `INSERT` per row.

```sql
DROP TABLE IF EXISTS "users";
CREATE TABLE "users" ("id" INTEGER, "name" TEXT, "age" INTEGER);
INSERT INTO "users" ("id", "name", "age") VALUES (1, 'alice', 30);
INSERT INTO "users" ("id", "name", "age") VALUES (2, 'O''Brien', 45);
```

Identifiers are always wrapped in `"..."` (reserved-word and space safe).
String literals use single quotes with apostrophes doubled. BLOBs go in
as `X'<hex>'`. NULLs as `NULL`. Numbers via `Str$()` for locale-safety.

The per-statement generators are also available individually:

```vb
sql = t.GenerateInsert("users", rowIdx)
sql = t.GenerateUpdate("users", rowIdx, "id = 42")
sql = t.GenerateDelete("users", "id = 42")
```

The container stays dumb: caller supplies the table name and any WHERE
clause. No source-table tracking, no dirty-flag state, no PK awareness.
`GenerateUpdate` writes all columns (no dirty tracking) and refuses to
generate an unconditional UPDATE or DELETE.

## The container is detached

Worth being explicit about what this means:

- A `cSQLiteTable` doesn't know which database (if any) it came from.
- It doesn't track constraints, indexes, defaults, foreign keys, or
  primary keys. `GenerateCreateTable` synthesizes column-name-and-type
  DDL only.
- Mutating a cell doesn't propagate anywhere. The generators emit SQL
  for the caller to execute; running it is the caller's choice.
- A loaded table will outlive the `cSQLiteResults` it came from. The
  `LoadFromResults` call drains and closes the cursor.

This is by design. The class is a marshalling format and an editor's
working copy, not a smart cursor. If you want smart-cursor behavior,
ADO Recordset already exists.

## Type system

`modSQLiteTypes.bas` defines a public enum that mirrors the SQLITE_*
storage class codes:

```vb
Public Enum SQLiteType
    sqlInteger = 1
    sqlFloat   = 2
    sqlText    = 3
    sqlBlob    = 4
    sqlNull    = 5
End Enum
```

Plus `SqlTypeName(t)` for friendly display strings. Values are
deliberately 1..5 to match what `sqlite3_column_type()` returns, so they
serialize directly into the JSON format and back.

## Setup

These classes use late-bound `Scripting.Dictionary` (via `CreateObject`),
so no project reference is required at compile time. The runtime
dependency is `scrrun.dll`, which ships with every supported Windows
version.

`FillFlexGrid` takes its argument as `Object` so the class doesn't drag
a hard reference to the FlexGrid OCX — you can compile a project that
uses `cSQLiteTable` without having that control loaded.

For an ActiveX DLL build, set instancing as follows:

| Class                          | Instancing            |
| ------------------------------ | --------------------- |
| `cSQLiteTable`                 | `MultiUse`            |
| `cSQLiteField`                 | `PublicNotCreatable`  |
| `JsonParser_cSQLiteTable`      | `Private` (already set) |

## What this is for

The original use case was building an AI loop around a SQLite database:
let an AI ask "what tables exist? what's the schema? show me 10 rows
where X." Marshall the result through `cSQLiteTable` → `ToJson()` →
feed to the model. Let the model propose an edit → set the new value
in the table → `GenerateUpdate()` → show the SQL → optionally execute.

Other places this earns its keep:

- **Test fixtures.** Build a `cSQLiteTable` with `SetColumns` + `AddRow`,
  hand to code that expects query results, no real db needed.
- **Snapshot diffing.** Run the same query against two databases, dump
  both as JSON, diff.
- **Caching.** Expensive query results saved as JSON, reloaded on next
  run without hitting the db. BLOBs survive via hex codec.
- **ETL-lite.** `LoadCsvFile` → `ToSql` is "CSV to SQLite" in three lines,
  no ODBC driver, no MDB middleware.
- **Form serialization.** A flexgrid full of user-edited data →
  `cSQLiteTable` built from it → JSON to disk → reload later.

## What got tested

Built and validated against real data:

- 30 adversarial string cases through write→read JSON round-trip
  (embedded quotes anywhere, backslashes, Windows paths, control chars,
  non-ASCII, empty strings, trailing-backslash trap, SQL-injection-shaped
  payloads, apostrophes, JSON-syntax-inside-strings)
- 18 CSV parsing cases compared against Python's `csv` module byte-for-byte
- 33 type-coercion edge cases (numbers, dates, sign-only, exponent
  variants, locale-y thousands separators, version numbers like `1.5.3`)
- Full CSV → SQL → sqlite3 pipeline executing cleanly with embedded
  delimiters, embedded newlines inside quoted fields, NULLs, apostrophes
- BLOB round-trip via hex codec including empty and binary content

## License

MIT, matching the rest of the `cSQLite` library.

## Credits

Written with [Claude (Anthropic)](https://www.anthropic.com/) as an
extension to David Zimmer's `cSQLite` library at http://sandsprite.com.
