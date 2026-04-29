# sqlite3_vb6

A VB6 wrapper library for [SQLite](https://sqlite.org), with full Unicode support, an ADO-flavored class API, and a complete schema introspection / migration toolkit.

---

## Table of Contents

- [Quick Start](#quick-start)
- [Architecture](#architecture)
- [Class Reference](#class-reference)
  - [cSQLite — connection](#csqlite--connection)
  - [cSQLiteStatement — prepared statement](#csqlitestatement--prepared-statement)
  - [cSQLiteResults — streaming result set](#csqliteresults--streaming-result-set)
  - [cSQLiteSchema — schema introspection & DDL](#csqliteschema--schema-introspection--ddl)
  - [cSQLiteColumn — column metadata](#csqlitecolumn--column-metadata)
- [Constants](#constants)
- [Examples](#examples)
  - [Hello, SQLite](#hello-sqlite)
  - [Insert a few rows](#insert-a-few-rows)
  - [Query and iterate](#query-and-iterate)
  - [Bulk insert in a transaction](#bulk-insert-in-a-transaction)
  - [Working with blobs](#working-with-blobs)
  - [Schema introspection](#schema-introspection)
  - [Idempotent migrations](#idempotent-migrations)
  - [Database maintenance](#database-maintenance)
  - [Unicode round-trip](#unicode-round-trip)

---

## Quick Start

1. Drop `sqlite3_vb6.dll` (the C shim) next to your `.exe`.
2. Register the ActiveX DLL: `regsvr32 sqlite4vb.dll` (one-time, per machine).
3. In the VB6 IDE: **Project → References → Browse**, pick the ActiveX DLL.
4. Code:

```vb
Dim db As New cSQLite
db.OpenDB App.Path & "\my.db"
db.Execute "CREATE TABLE IF NOT EXISTS notes (id INTEGER PRIMARY KEY, body TEXT)"
db.ExecInsert "notes", "body", "Hello, world!"

Dim rs As cSQLiteResults
Set rs = db.Query("SELECT id, body FROM notes")
Do While rs.MoveNext()
    Debug.Print rs("id"), rs("body")
Loop
' db closes automatically when it goes out of scope
```

---

## Architecture

The library has three layers, top to bottom:

| Layer | Files | What it is |
|---|---|---|
| **High-level OO API** | `cSQLite.cls`, `cSQLiteStatement.cls`, `cSQLiteResults.cls`, `cSQLiteSchema.cls`, `cSQLiteColumn.cls` | What you'll use 99% of the time. ADO-flavored, RAII cleanup, type-dispatched binding. |
| **Low-level API** | `cSQLiteAPI.cls` (a `GlobalMultiUse` class in the ActiveX DLL) | Direct passthroughs to the `sqlite3_*` C functions. Globally accessible by bare name (`sqlite3_open`, `sqlite3_step`, etc.) — no instance needed. Use when you want raw control the OO layer doesn't expose. |
| **C shim** | `sqlite3_vb6.dll` (built from `sqlite3_vb6.c` + the SQLite amalgamation) | stdcall wrapper around SQLite with Unicode marshalling. You don't touch this directly. |

The C shim handles the cdecl→stdcall conversion and UTF-16↔UTF-8 string marshalling so VB6 can bind cleanly without `_Foo@N` aliases or codepage loss.

### Usage flavors

Most apps just reference the ActiveX DLL and use the OO classes:

```vb
Dim db As New cSQLite
db.OpenDB "my.db"
```

If you want raw API access (porting C code, debugging, edge cases not exposed by the classes), the global names work without qualification:

```vb
Dim h As Long
sqlite3_open "my.db", h
' ... etc
```

Both styles work in the same project — the OO classes are built on top of the same low-level API surface.

---

## Class Reference

### cSQLite — connection

The main entry point. Open a database, run statements, prepare queries, manage transactions.

| Member | What it does |
|---|---|
| [`OpenDB`](#csqlite-opendb) | Open a database file (creates if missing) |
| [`CloseDB`](#csqlite-closedb) | Close the connection |
| [`IsOpen`](#csqlite-isopen) | True if a connection is currently open |
| [`Path`](#csqlite-path) | Path of the open database file |
| [`Handle`](#csqlite-handle) | Raw `sqlite3*` handle as Long |
| [`Version`](#csqlite-version) | SQLite library version string |
| [`Execute`](#csqlite-execute) | Run SQL with no parameters and no result rows |
| [`Prepare`](#csqlite-prepare) | Build a `cSQLiteStatement` for repeated execution |
| [`Query`](#csqlite-query) | Run a SELECT, return a `cSQLiteResults` iterator |
| [`Scalar`](#csqlite-scalar) | Run a SELECT, return the first column of the first row |
| [`ExecInsert`](#csqlite-execinsert) | Build & run a parameterized INSERT in one call |
| [`ExecUpdate`](#csqlite-execupdate) | Build & run a parameterized UPDATE in one call |
| [`BeginTrans`](#csqlite-begintrans) | BEGIN TRANSACTION |
| [`CommitTrans`](#csqlite-committrans) | COMMIT |
| [`RollbackTrans`](#csqlite-rollbacktrans) | ROLLBACK |
| [`InTransaction`](#csqlite-intransaction) | True while a transaction is active |
| [`LastInsertRowID`](#csqlite-lastinsertrowid) | Rowid of the last successful INSERT |
| [`RowsAffected`](#csqlite-rowsaffected) | Rows changed by the last INSERT/UPDATE/DELETE |
| [`LastErrorCode`](#csqlite-lasterrorcode) | Most recent SQLite result code |
| [`LastErrorMessage`](#csqlite-lasterrormessage) | Most recent SQLite error message |

#### Prototypes

<a name="csqlite-opendb"></a>
```vb
Sub OpenDB(ByVal sFilename As String)
```
Opens (or creates) the SQLite database at `sFilename`. Sets a 5-second busy timeout by default. Raises on failure.

<a name="csqlite-closedb"></a>
```vb
Sub CloseDB()
```
Closes the connection. Called automatically by `Class_Terminate`, so you usually don't need to invoke it explicitly.

<a name="csqlite-isopen"></a>
```vb
Property Get IsOpen() As Boolean
```
Returns True if a connection is currently open.

<a name="csqlite-path"></a>
```vb
Property Get Path() As String
```
Returns the file path of the currently open database, or `""` if not open.

<a name="csqlite-handle"></a>
```vb
Property Get Handle() As Long
```
The raw `sqlite3*` handle. Useful only if you need to call low-level functions in `modSQLite.bas` directly.

<a name="csqlite-version"></a>
```vb
Property Get Version() As String
```
The SQLite library version (e.g. `"3.46.0"`).

<a name="csqlite-execute"></a>
```vb
Sub Execute(ByVal sSql As String)
```
Runs SQL that doesn't return rows (DDL, INSERT/UPDATE/DELETE without parameters). Raises on error with the SQLite error message.

<a name="csqlite-prepare"></a>
```vb
Function Prepare(ByVal sSql As String) As cSQLiteStatement
```
Returns a prepared statement. Use for any SQL that runs more than once or takes parameters. The statement is finalized automatically when its variable goes out of scope.

<a name="csqlite-query"></a>
```vb
Function Query(ByVal sSql As String, ParamArray params()) As cSQLiteResults
```
Convenience wrapper that prepares a statement, binds positional `?` parameters from `params`, and returns a result-set iterator. Use [`MoveNext`](#csqliteresults-movenext) to walk rows.

<a name="csqlite-scalar"></a>
```vb
Function Scalar(ByVal sSql As String, ParamArray params()) As Variant
```
Runs a query and returns the first column of the first row (as a Variant). Returns `Empty` if the query produced no rows. Great for things like `SELECT COUNT(*) FROM foo`.

<a name="csqlite-execinsert"></a>
```vb
Function ExecInsert(ByVal tblName As String, ByVal fields As String, _
                    ParamArray params()) As Currency
```
Builds `INSERT INTO tblName (fields) VALUES (?,?,...)` with placeholder count matching `params`, binds them, executes. Returns the new rowid as Currency (use `Int64FromCurrency` to display). Type-correct for Strings, Longs, Doubles, Dates, Byte arrays, and Null/Empty — no SQL escaping concerns.

<a name="csqlite-execupdate"></a>
```vb
Function ExecUpdate(ByVal tblName As String, ByVal criteria As String, _
                    ByVal fields As String, ParamArray params()) As Long
```
Builds `UPDATE tblName SET f1=?,f2=?,... criteria` and binds. The `criteria` string can itself contain `?` placeholders; values for them go in `params` *after* the SET values. Returns rows affected.

<a name="csqlite-begintrans"></a>
```vb
Sub BeginTrans()
```
Starts a transaction. Pair with [`CommitTrans`](#csqlite-committrans) or [`RollbackTrans`](#csqlite-rollbacktrans).

<a name="csqlite-committrans"></a>
```vb
Sub CommitTrans()
```
Commits the current transaction.

<a name="csqlite-rollbacktrans"></a>
```vb
Sub RollbackTrans()
```
Rolls back the current transaction, discarding all changes since `BeginTrans`.

<a name="csqlite-intransaction"></a>
```vb
Property Get InTransaction() As Boolean
```
True while a transaction is currently open.

<a name="csqlite-lastinsertrowid"></a>
```vb
Property Get LastInsertRowID() As Currency
```
Rowid of the most recent successful INSERT. Currency carries 64 bits losslessly across the COM boundary.

<a name="csqlite-rowsaffected"></a>
```vb
Property Get RowsAffected() As Long
```
Rows changed by the most recent INSERT, UPDATE, or DELETE.

<a name="csqlite-lasterrorcode"></a>
```vb
Property Get LastErrorCode() As Long
```
Most recent SQLite result code.

<a name="csqlite-lasterrormessage"></a>
```vb
Property Get LastErrorMessage() As String
```
Most recent SQLite error message.

---

### cSQLiteStatement — prepared statement

You don't construct these directly. Get one via [`db.Prepare(sql)`](#csqlite-prepare).

| Member | What it does |
|---|---|
| [`Bind`](#csqlitestatement-bind) | Bind a value to a `?` parameter (type-dispatched) |
| [`BindNamed`](#csqlitestatement-bindnamed) | Bind by `:name` or `@name` parameter |
| [`Step`](#csqlitestatement-step) | Advance to next row; True if a row is available |
| [`Execute`](#csqlitestatement-execute) | Step a non-row-returning statement to completion |
| [`Reset`](#csqlitestatement-reset) | Reset cursor to start; keep bindings |
| [`ClearBindings`](#csqlitestatement-clearbindings) | Reset all bindings to NULL |
| [`Finalize`](#csqlitestatement-finalize) | Release the statement (auto-called on terminate) |
| [`SQL`](#csqlitestatement-sql) | The SQL text this was prepared from |
| [`Handle`](#csqlitestatement-handle) | Raw `sqlite3_stmt*` handle |
| [`ParameterCount`](#csqlitestatement-parametercount) | Number of `?` placeholders |
| [`ColumnCount`](#csqlitestatement-columncount) | Number of result columns |
| [`ColumnName`](#csqlitestatement-columnname) | Name of column N |
| [`ColumnDeclType`](#csqlitestatement-columndecltype) | Declared type of column N (e.g. `"TEXT"`) |
| [`ColumnType`](#csqlitestatement-columntype) | Storage type of column N (`SQLITE_INTEGER` etc.) |
| [`ColumnInt`](#csqlitestatement-columnint) | Get column N as Long |
| [`ColumnInt64`](#csqlitestatement-columnint64) | Get column N as Currency (carries int64) |
| [`ColumnDouble`](#csqlitestatement-columndouble) | Get column N as Double |
| [`ColumnText`](#csqlitestatement-columntext) | Get column N as String |
| [`ColumnBlob`](#csqlitestatement-columnblob) | Get column N as Byte() |
| [`ColumnValue`](#csqlitestatement-columnvalue) | Get column N as natural-typed Variant |

#### Prototypes

<a name="csqlitestatement-bind"></a>
```vb
Sub Bind(ByVal idx As Long, ByRef value As Variant)
```
Binds `value` to parameter `idx` (1-based). Routes to the right SQLite bind based on the Variant's type: String → `bind_text`, Long/Integer/Boolean → `bind_int`, Currency → `bind_int64`, Double → `bind_double`, Date → ISO 8601 text, Byte() → `bind_blob`, Null/Empty → `bind_null`.

<a name="csqlitestatement-bindnamed"></a>
```vb
Sub BindNamed(ByVal sName As String, ByRef value As Variant)
```
Binds by named parameter (e.g. `":userId"` or `"@email"`). Includes the leading `:` or `@`. Raises if the name doesn't exist.

<a name="csqlitestatement-step"></a>
```vb
Function Step() As Boolean
```
Advances the cursor. Returns True if a row is available, False on `SQLITE_DONE`. Raises on error.

<a name="csqlitestatement-execute"></a>
```vb
Sub Execute()
```
For statements that don't return rows. Steps to completion and resets, ready for re-use with new bindings.

<a name="csqlitestatement-reset"></a>
```vb
Sub Reset()
```
Resets the statement so it can be re-stepped from the beginning. **Does not clear bindings** — those persist across resets. Use this for re-executing the same statement with the same parameters.

<a name="csqlitestatement-clearbindings"></a>
```vb
Sub ClearBindings()
```
Sets all parameters to NULL. Call after `Reset` if you want to re-bind from scratch.

<a name="csqlitestatement-finalize"></a>
```vb
Sub Finalize()
```
Releases the statement. Called automatically when the variable goes out of scope; rarely need to call directly.

<a name="csqlitestatement-sql"></a>
```vb
Property Get SQL() As String
```
The SQL text this statement was prepared from.

<a name="csqlitestatement-handle"></a>
```vb
Property Get Handle() As Long
```
Raw `sqlite3_stmt*` handle.

<a name="csqlitestatement-parametercount"></a>
```vb
Property Get ParameterCount() As Long
```
Number of `?` placeholders.

<a name="csqlitestatement-columncount"></a>
```vb
Property Get ColumnCount() As Long
```
Number of columns in the result set.

<a name="csqlitestatement-columnname"></a>
```vb
Function ColumnName(ByVal iCol As Long) As String
```
Name of column `iCol` (0-based).

<a name="csqlitestatement-columndecltype"></a>
```vb
Function ColumnDeclType(ByVal iCol As Long) As String
```
Declared type from the CREATE TABLE statement (e.g. `"TEXT"`, `"INTEGER NOT NULL"`).

<a name="csqlitestatement-columntype"></a>
```vb
Function ColumnType(ByVal iCol As Long) As Long
```
Storage class of the value: `SQLITE_INTEGER`, `SQLITE_FLOAT`, `SQLITE_TEXT`, `SQLITE_BLOB`, or `SQLITE_NULL`.

<a name="csqlitestatement-columnint"></a>
```vb
Function ColumnInt(ByVal iCol As Long) As Long
```
Column value as Long.

<a name="csqlitestatement-columnint64"></a>
```vb
Function ColumnInt64(ByVal iCol As Long) As Currency
```
Column value as Currency (the bytes carry int64 losslessly; use `Int64FromCurrency` to convert to a displayable integer).

<a name="csqlitestatement-columndouble"></a>
```vb
Function ColumnDouble(ByVal iCol As Long) As Double
```
Column value as Double.

<a name="csqlitestatement-columntext"></a>
```vb
Function ColumnText(ByVal iCol As Long) As String
```
Column value as String (UTF-8 → UTF-16 conversion handled).

<a name="csqlitestatement-columnblob"></a>
```vb
Function ColumnBlob(ByVal iCol As Long) As Byte()
```
Column value as a fresh Byte array. Empty array for NULL/zero-length.

<a name="csqlitestatement-columnvalue"></a>
```vb
Function ColumnValue(ByVal iCol As Long) As Variant
```
Column value as a Variant, dispatched on the column's actual storage type. Returns `Long` (for in-range integers), `Decimal` (for big int64), `Double`, `String`, `Byte()`, or `Null`.

---

### cSQLiteResults — streaming result set

The recordset returned by [`db.Query`](#csqlite-query). Streaming — pulls one row at a time. Use this for any query that might be large.

| Member | What it does |
|---|---|
| [`MoveNext`](#csqliteresults-movenext) | Advance to the next row; True if a row is now current |
| [`AtEnd`](#csqliteresults-atend) | True after the last row has been consumed |
| [`Item`](#csqliteresults-item) | **Default property** — `rs("name")` or `rs(0)` |
| [`FieldCount`](#csqliteresults-fieldcount) | Number of columns |
| [`FieldName`](#csqliteresults-fieldname) | Name of column N |
| [`FieldNames`](#csqliteresults-fieldnames) | All column names as a String() |
| [`FieldType`](#csqliteresults-fieldtype) | Storage type of a column |
| [`FieldText`](#csqliteresults-fieldtext) | Typed getter — String |
| [`FieldInt`](#csqliteresults-fieldint) | Typed getter — Long |
| [`FieldInt64`](#csqliteresults-fieldint64) | Typed getter — Currency |
| [`FieldDouble`](#csqliteresults-fielddouble) | Typed getter — Double |
| [`FieldBlob`](#csqliteresults-fieldblob) | Typed getter — Byte() |
| [`LoadAll`](#csqliteresults-loadall) | Read all rows into a 2D Variant array |
| [`CloseRS`](#csqliteresults-closers) | Close early (auto-closes on terminate) |

#### Prototypes

<a name="csqliteresults-movenext"></a>
```vb
Function MoveNext() As Boolean
```
Steps to the next row. Returns True if a row is now current, False at EOF. Use as the loop condition.

<a name="csqliteresults-atend"></a>
```vb
Property Get AtEnd() As Boolean
```
True once `MoveNext` has returned False at least once.

<a name="csqliteresults-item"></a>
```vb
Property Get Item(ByVal key As Variant) As Variant
```
Default property: `rs("colname")` returns the column by name (case-insensitive); `rs(0)` returns by 0-based index. Returns a Variant of the natural type.

<a name="csqliteresults-fieldcount"></a>
```vb
Property Get FieldCount() As Long
```
Number of columns in the current row.

<a name="csqliteresults-fieldname"></a>
```vb
Function FieldName(ByVal iCol As Long) As String
```
Name of column `iCol` (0-based).

<a name="csqliteresults-fieldnames"></a>
```vb
Function FieldNames() As String()
```
All column names. Handy for printing headers.

<a name="csqliteresults-fieldtype"></a>
```vb
Function FieldType(ByVal key As Variant) As Long
```
Storage type constant for the column (`SQLITE_INTEGER`, `SQLITE_FLOAT`, etc.). `key` can be a column name or 0-based index.

<a name="csqliteresults-fieldtext"></a>
```vb
Function FieldText(ByVal key As Variant) As String
```
Typed access — returns the column as String. Avoids Variant overhead.

<a name="csqliteresults-fieldint"></a>
```vb
Function FieldInt(ByVal key As Variant) As Long
```
Typed access — returns the column as Long.

<a name="csqliteresults-fieldint64"></a>
```vb
Function FieldInt64(ByVal key As Variant) As Currency
```
Typed access — returns the column as Currency (int64-safe).

<a name="csqliteresults-fielddouble"></a>
```vb
Function FieldDouble(ByVal key As Variant) As Double
```
Typed access — returns the column as Double.

<a name="csqliteresults-fieldblob"></a>
```vb
Function FieldBlob(ByVal key As Variant) As Byte()
```
Typed access — returns the column as Byte().

<a name="csqliteresults-loadall"></a>
```vb
Function LoadAll() As Variant
```
Reads every remaining row into a 2D Variant array `arr(rowIdx, colIdx)`. Returns `Empty` if no rows. Consumes the cursor.

<a name="csqliteresults-closers"></a>
```vb
Sub CloseRS()
```
Closes the result set early. Called automatically when the variable goes out of scope.

---

### cSQLiteSchema — schema introspection & DDL

Schema metadata, table/index/view manipulation, app versioning, and database maintenance. Attach to a `cSQLite` instance with [`AttachDb`](#csqliteschema-attachdb).

| Member | What it does |
|---|---|
| [`AttachDb`](#csqliteschema-attachdb) | Attach to a `cSQLite` connection |
| **Listing** | |
| [`ListTables`](#csqliteschema-listtables) | All user tables (Collection) |
| [`ListTablesArr`](#csqliteschema-listtablesarr) | All user tables (String array) |
| [`ListViews`](#csqliteschema-listviews) | All views |
| [`ListIndexes`](#csqliteschema-listindexes) | All indexes (optionally filtered by table) |
| **Existence** | |
| [`TableExists`](#csqliteschema-tableexists) | True if a table exists |
| [`ColumnExists`](#csqliteschema-columnexists) | True if a column exists on a table |
| [`IndexExists`](#csqliteschema-indexexists) | True if an index exists |
| [`ViewExists`](#csqliteschema-viewexists) | True if a view exists |
| **Inspection** | |
| [`DescribeTable`](#csqliteschema-describetable) | Collection of `cSQLiteColumn` for a table |
| [`DescribeTableText`](#csqliteschema-describetabletext) | Pretty-printed table description |
| [`GetCreateSQL`](#csqliteschema-getcreatesql) | The CREATE TABLE statement |
| [`RowCount`](#csqliteschema-rowcount) | Row count of a table |
| **DDL** | |
| [`CreateTable`](#csqliteschema-createtable) | CREATE TABLE |
| [`CreateTableIfMissing`](#csqliteschema-createtableifmissing) | CREATE TABLE IF NOT EXISTS |
| [`DropTable`](#csqliteschema-droptable) | DROP TABLE |
| [`RenameTable`](#csqliteschema-renametable) | ALTER TABLE … RENAME |
| [`AddColumn`](#csqliteschema-addcolumn) | ALTER TABLE … ADD COLUMN |
| [`RenameColumn`](#csqliteschema-renamecolumn) | ALTER TABLE … RENAME COLUMN (3.25+) |
| [`DropColumn`](#csqliteschema-dropcolumn) | ALTER TABLE … DROP COLUMN (3.35+) |
| [`CreateIndex`](#csqliteschema-createindex) | CREATE INDEX |
| [`DropIndex`](#csqliteschema-dropindex) | DROP INDEX |
| **Versioning** | |
| [`UserVersion`](#csqliteschema-userversion) | Read/write `PRAGMA user_version` |
| [`SchemaVersion`](#csqliteschema-schemaversion) | Read `PRAGMA schema_version` |
| [`ApplicationID`](#csqliteschema-applicationid) | Read/write `PRAGMA application_id` |
| **Maintenance** | |
| [`CompactDB`](#csqliteschema-compactdb) | VACUUM the database |
| [`DatabaseSize`](#csqliteschema-databasesize) | File size in bytes |
| [`FreelistPages`](#csqliteschema-freelistpages) | Number of unused pages |
| [`FreeSpace`](#csqliteschema-freespace) | Free space in bytes |
| [`Optimize`](#csqliteschema-optimize) | PRAGMA optimize |
| [`Analyze`](#csqliteschema-analyze) | ANALYZE |
| [`Reindex`](#csqliteschema-reindex) | REINDEX |
| [`IntegrityCheck`](#csqliteschema-integritycheck) | Full integrity check |
| [`QuickCheck`](#csqliteschema-quickcheck) | Fast check (no index validation) |
| [`WALCheckpoint`](#csqliteschema-walcheckpoint) | Checkpoint the WAL |
| **Utility** | |
| [`QuoteIdent`](#csqliteschema-quoteident) | Safely quote a SQL identifier |

#### Prototypes

<a name="csqliteschema-attachdb"></a>
```vb
Sub AttachDb(ByRef db As cSQLite)
```
Attach to an open connection. Required before any other call.

<a name="csqliteschema-listtables"></a>
```vb
Function ListTables() As Collection
```
All user tables as a Collection of names. Excludes SQLite internal tables (`sqlite_*`).

<a name="csqliteschema-listtablesarr"></a>
```vb
Function ListTablesArr() As String()
```
Same as `ListTables` but as a String array — handy for combo-box population.

<a name="csqliteschema-listviews"></a>
```vb
Function ListViews() As Collection
```
All views as a Collection of names.

<a name="csqliteschema-listindexes"></a>
```vb
Function ListIndexes(Optional ByVal sTable As String = "") As Collection
```
All indexes (excluding internal `sqlite_*` indexes). Optionally filter to a single table.

<a name="csqliteschema-tableexists"></a>
```vb
Function TableExists(ByVal sTable As String) As Boolean
```
True if the named table exists.

<a name="csqliteschema-columnexists"></a>
```vb
Function ColumnExists(ByVal sTable As String, ByVal sColumn As String) As Boolean
```
True if the column exists on the table. Case-insensitive. Returns False if the table itself doesn't exist (no error).

<a name="csqliteschema-indexexists"></a>
```vb
Function IndexExists(ByVal sIndexName As String) As Boolean
```
True if the named index exists.

<a name="csqliteschema-viewexists"></a>
```vb
Function ViewExists(ByVal sViewName As String) As Boolean
```
True if the named view exists.

<a name="csqliteschema-describetable"></a>
```vb
Function DescribeTable(ByVal sTable As String) As Collection
```
Returns a Collection of [`cSQLiteColumn`](#csqlitecolumn--column-metadata) objects describing each column.

<a name="csqliteschema-describetabletext"></a>
```vb
Function DescribeTableText(ByVal sTable As String) As String
```
A formatted human-readable description of the table — useful for debug output.

<a name="csqliteschema-getcreatesql"></a>
```vb
Function GetCreateSQL(ByVal sTable As String) As String
```
The original `CREATE TABLE` statement from `sqlite_master`.

<a name="csqliteschema-rowcount"></a>
```vb
Function RowCount(ByVal sTable As String) As Currency
```
`SELECT COUNT(*)` on the table. Returned as Currency for int64 safety.

<a name="csqliteschema-createtable"></a>
```vb
Sub CreateTable(ByVal sTable As String, ByVal columnDefs As String)
```
Creates a table. `columnDefs` is the inside of the parens, e.g. `"id INTEGER PRIMARY KEY, name TEXT NOT NULL"`.

<a name="csqliteschema-createtableifmissing"></a>
```vb
Sub CreateTableIfMissing(ByVal sTable As String, ByVal columnDefs As String)
```
`CREATE TABLE IF NOT EXISTS`.

<a name="csqliteschema-droptable"></a>
```vb
Sub DropTable(ByVal sTable As String, Optional ByVal ifExists As Boolean = True)
```
Drops a table. Defaults to `DROP TABLE IF EXISTS`.

<a name="csqliteschema-renametable"></a>
```vb
Sub RenameTable(ByVal sOldName As String, ByVal sNewName As String)
```
Renames a table.

<a name="csqliteschema-addcolumn"></a>
```vb
Sub AddColumn(ByVal sTable As String, ByVal sColName As String, _
              ByVal sColType As String, _
              Optional ByVal sExtra As String = "")
```
Adds a column. `sExtra` is appended after the type — use it for `"NOT NULL DEFAULT 0"` or similar.

<a name="csqliteschema-renamecolumn"></a>
```vb
Sub RenameColumn(ByVal sTable As String, _
                 ByVal sOldCol As String, ByVal sNewCol As String)
```
Renames a column. Requires SQLite 3.25+.

<a name="csqliteschema-dropcolumn"></a>
```vb
Sub DropColumn(ByVal sTable As String, ByVal sColName As String)
```
Drops a column. Requires SQLite 3.35+.

<a name="csqliteschema-createindex"></a>
```vb
Sub CreateIndex(ByVal sIndexName As String, ByVal sTable As String, _
                ByVal sColumns As String, _
                Optional ByVal unique As Boolean = False)
```
Creates an index. `sColumns` is the comma-separated column list (e.g. `"name"` or `"lastname, firstname"`).

<a name="csqliteschema-dropindex"></a>
```vb
Sub DropIndex(ByVal sIndexName As String, _
              Optional ByVal ifExists As Boolean = True)
```
Drops an index. Defaults to `DROP INDEX IF EXISTS`.

<a name="csqliteschema-userversion"></a>
```vb
Property Get UserVersion() As Long
Property Let UserVersion(ByVal n As Long)
```
Read/write the database's `user_version` field — a 32-bit integer reserved for application schema versioning. Survives VACUUM. The standard mechanism for migrations.

<a name="csqliteschema-schemaversion"></a>
```vb
Property Get SchemaVersion() As Long
```
Read-only counter SQLite increments on every schema change. Useful for cache invalidation.

<a name="csqliteschema-applicationid"></a>
```vb
Property Get ApplicationID() As Long
Property Let ApplicationID(ByVal n As Long)
```
Read/write the `application_id` field — a 32-bit magic number identifying which app owns this database file.

<a name="csqliteschema-compactdb"></a>
```vb
Function CompactDB() As Double
```
Runs `VACUUM`. Rebuilds the database file from scratch, reclaiming free pages. Returns the bytes saved (negative if it grew). Locks the database for the duration.

<a name="csqliteschema-databasesize"></a>
```vb
Property Get DatabaseSize() As Double
```
Current file size in bytes.

<a name="csqliteschema-freelistpages"></a>
```vb
Property Get FreelistPages() As Long
```
Number of unused pages in the file (reclaimed by VACUUM).

<a name="csqliteschema-freespace"></a>
```vb
Property Get FreeSpace() As Double
```
Bytes equivalent of the freelist (`FreelistPages × page size`).

<a name="csqliteschema-optimize"></a>
```vb
Sub Optimize()
```
`PRAGMA optimize`. Cheap. Recommended before closing long-lived connections.

<a name="csqliteschema-analyze"></a>
```vb
Sub Analyze(Optional ByVal sTable As String = "")
```
Full `ANALYZE` (one table or all). Heavier than `Optimize` but more thorough.

<a name="csqliteschema-reindex"></a>
```vb
Sub Reindex(Optional ByVal sName As String = "")
```
Rebuilds indexes. Rarely needed except after collation changes.

<a name="csqliteschema-integritycheck"></a>
```vb
Function IntegrityCheck(Optional ByVal maxErrors As Long = 100) As String
```
Full `PRAGMA integrity_check`. Returns `"ok"` if healthy, otherwise a multi-line error report. Slow on big databases.

<a name="csqliteschema-quickcheck"></a>
```vb
Function QuickCheck(Optional ByVal maxErrors As Long = 100) As String
```
Faster than `IntegrityCheck` — doesn't validate index/table consistency.

<a name="csqliteschema-walcheckpoint"></a>
```vb
Sub WALCheckpoint(Optional ByVal mode As String = "PASSIVE")
```
Checkpoint the write-ahead log. Modes: `"PASSIVE"`, `"FULL"`, `"RESTART"`, `"TRUNCATE"`. No-op for non-WAL databases.

<a name="csqliteschema-quoteident"></a>
```vb
Function QuoteIdent(ByVal s As String) As String
```
Wraps an identifier in double-quotes and escapes embedded quotes — safe for identifiers that might collide with reserved words.

---

### cSQLiteColumn — column metadata

A simple immutable struct returned by [`DescribeTable`](#csqliteschema-describetable).

| Member | What it returns |
|---|---|
| `Name` | Column name |
| `DeclType` | Declared type (e.g. `"INTEGER"`, `"TEXT NOT NULL"`) |
| `NotNull` | True if NOT NULL constraint is set |
| `DefaultValue` | Default value as text (empty if none) |
| `IsPrimaryKey` | True if this column participates in the primary key |
| `PKOrdinal` | 1-based ordinal in composite keys; 0 for non-PK columns |
| `ColumnIndex` | 0-based position in the table |

---

## Constants

All of these are members of the global `SQLiteConsts` enum and accessible by bare name (no qualifier needed).

**Result codes:**
`SQLITE_OK` (0), `SQLITE_ERROR` (1), `SQLITE_BUSY` (5), `SQLITE_LOCKED` (6), `SQLITE_NOMEM` (7), `SQLITE_READONLY` (8), `SQLITE_INTERRUPT` (9), `SQLITE_IOERR` (10), `SQLITE_CORRUPT` (11), `SQLITE_FULL` (13), `SQLITE_CANTOPEN` (14), `SQLITE_CONSTRAINT` (19), `SQLITE_MISMATCH` (20), `SQLITE_MISUSE` (21), `SQLITE_RANGE` (25), `SQLITE_NOTADB` (26), `SQLITE_ROW` (100), `SQLITE_DONE` (101).

**Column storage types:**
`SQLITE_INTEGER` (1), `SQLITE_FLOAT` (2), `SQLITE_TEXT` (3), `SQLITE_BLOB` (4), `SQLITE_NULL` (5).

**Open flags (bitwise OR-able):**
`SQLITE_OPEN_READONLY`, `SQLITE_OPEN_READWRITE`, `SQLITE_OPEN_CREATE`, `SQLITE_OPEN_URI`, `SQLITE_OPEN_MEMORY`, `SQLITE_OPEN_NOMUTEX`, `SQLITE_OPEN_FULLMUTEX`.

---

## Examples

### Hello, SQLite

```vb
Dim db As New cSQLite
db.OpenDB App.Path & "\hello.db"
Debug.Print "SQLite version: " & db.Version
db.Execute "CREATE TABLE IF NOT EXISTS hello (msg TEXT)"
db.Execute "INSERT INTO hello (msg) VALUES ('Hi there!')"
Debug.Print db.Scalar("SELECT msg FROM hello")
```

### Insert a few rows

```vb
db.Execute "CREATE TABLE users (" & _
    "id INTEGER PRIMARY KEY AUTOINCREMENT," & _
    "name TEXT NOT NULL," & _
    "age INTEGER)"

db.ExecInsert "users", "name,age", "Alice", 30
db.ExecInsert "users", "name,age", "O'Brien", 45     ' apostrophe — no escaping needed
db.ExecInsert "users", "name,age", "Charlie", Null   ' NULL via VB6 Null
Debug.Print "last id = " & Int64FromCurrency(db.LastInsertRowID)
```

### Query and iterate

Streaming — the right pattern for any query:

```vb
Dim rs As cSQLiteResults
Set rs = db.Query("SELECT id, name, age FROM users WHERE age >= ? ORDER BY age DESC", 30)
Do While rs.MoveNext()
    Debug.Print rs("id"), rs("name"), rs("age")
Loop
```

Single-value queries:

```vb
Dim avgAge As Double
avgAge = db.Scalar("SELECT AVG(age) FROM users")
```

Materialize a small result into a 2D array if you need random access:

```vb
Dim data As Variant
data = db.Query("SELECT name, age FROM users ORDER BY age").LoadAll
If Not IsEmpty(data) Then
    Dim r As Long
    For r = LBound(data, 1) To UBound(data, 1)
        Debug.Print data(r, 0), data(r, 1)
    Next r
End If
```

### Bulk insert in a transaction

100x-1000x faster than auto-commit:

```vb
db.BeginTrans
Dim stmt As cSQLiteStatement
Set stmt = db.Prepare("INSERT INTO logs (ts, msg) VALUES (?, ?)")
Dim i As Long
For i = 1 To 10000
    stmt.Bind 1, Now
    stmt.Bind 2, "log entry " & i
    stmt.Execute
Next i
Set stmt = Nothing  ' finalize
db.CommitTrans
```

### Working with blobs

```vb
' Insert a blob from a Byte array
Dim bytes() As Byte
ReDim bytes(0 To 15)
Dim i As Long
For i = 0 To 15: bytes(i) = i: Next i
db.ExecInsert "files", "name,data", "test.bin", bytes

' Read it back
Dim rs As cSQLiteResults
Set rs = db.Query("SELECT data FROM files WHERE name=?", "test.bin")
If rs.MoveNext() Then
    Dim got() As Byte
    got = rs.FieldBlob("data")
    Debug.Print "got " & (UBound(got) + 1) & " bytes"
End If
```

### Schema introspection

```vb
Dim sch As New cSQLiteSchema
sch.AttachDb db

' List everything
Dim tbl As Variant
For Each tbl In sch.ListTables
    Debug.Print tbl & " has " & sch.RowCount(CStr(tbl)) & " rows"
Next

' Describe a single table
Dim col As cSQLiteColumn
For Each col In sch.DescribeTable("users")
    Debug.Print col.Name, col.DeclType, _
                IIf(col.NotNull, "NOT NULL", "nullable"), _
                IIf(col.IsPrimaryKey, "PK", "")
Next

' Or the formatted version
Debug.Print sch.DescribeTableText("users")
```

### Idempotent migrations

The simplest form — check-and-add:

```vb
Dim sch As New cSQLiteSchema
sch.AttachDb db

If Not sch.ColumnExists("users", "email") Then _
    sch.AddColumn "users", "email", "TEXT"

If Not sch.IndexExists("idx_users_email") Then _
    sch.CreateIndex "idx_users_email", "users", "email", unique:=True
```

Versioned migrations using `user_version`:

```vb
Do
    Select Case sch.UserVersion
        Case 0
            sch.CreateTableIfMissing "users", _
                "id INTEGER PRIMARY KEY AUTOINCREMENT," & _
                "name TEXT NOT NULL"
            sch.UserVersion = 1
        Case 1
            If Not sch.ColumnExists("users", "email") Then _
                sch.AddColumn "users", "email", "TEXT"
            sch.UserVersion = 2
        Case 2
            If Not sch.ColumnExists("users", "active") Then _
                sch.AddColumn "users", "active", "INTEGER", "NOT NULL DEFAULT 1"
            sch.UserVersion = 3
        Case Else
            Exit Do
    End Select
Loop
```

### Database maintenance

```vb
Dim sch As New cSQLiteSchema
sch.AttachDb db

' Health check
Dim health As String
health = sch.IntegrityCheck()
If health <> "ok" Then MsgBox "Database problems:" & vbCrLf & health

' Compact (after big deletes)
Debug.Print "before: " & sch.DatabaseSize & " bytes, " & sch.FreelistPages & " free pages"
Dim saved As Double
saved = sch.CompactDB
Debug.Print "VACUUM saved " & saved & " bytes"

' Update query planner statistics (recommended on close)
sch.Optimize
```

### Unicode round-trip

Full Unicode through the wrapper — no codepage loss:

```vb
db.Execute "CREATE TABLE i18n (label TEXT, payload TEXT)"

' Cyrillic, Chinese, accented Latin, emoji (surrogate pair)
db.ExecInsert "i18n", "label,payload", "Russian", _
    ChrW(&H41F) & ChrW(&H440) & ChrW(&H438) & ChrW(&H432) & ChrW(&H435) & ChrW(&H442)
db.ExecInsert "i18n", "label,payload", "Chinese", _
    ChrW(&H4F60) & ChrW(&H597D) & ChrW(&H4E16) & ChrW(&H754C)
db.ExecInsert "i18n", "label,payload", "Accented", _
    "caf" & ChrW(&HE9) & " naïve résumé"
db.ExecInsert "i18n", "label,payload", "Emoji", _
    "smile " & ChrW(&HD83D) & ChrW(&HDE00)

' Round-trip preserves every byte
Dim rs As cSQLiteResults
Set rs = db.Query("SELECT label, payload FROM i18n")
Do While rs.MoveNext()
    Debug.Print rs("label"), rs("payload")
Loop
```

---

## License

MIT License. See [LICENSE](LICENSE) for details.

This package wraps SQLite, which is in the public domain. SQLite itself is not covered by this license — see [sqlite.org](https://sqlite.org) for SQLite's own (non-)terms.
