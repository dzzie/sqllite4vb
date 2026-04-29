VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com
'
' cSQLiteSchema — schema introspection & light DDL helpers.
'
' Usage:
'   Dim sch As New cSQLiteSchema
'   sch.AttachDb db
'
'   For Each tbl In sch.ListTables
'       Debug.Print tbl
'   Next
'
'   For Each col In sch.DescribeTable("users")
'       Debug.Print col.Name, col.DeclType, col.IsPrimaryKey, col.NotNull
'   Next
'
'   sch.AddColumn "users", "email", "TEXT"
'   sch.RenameTable "users", "people"
'   sch.DropTable "old_table"
'
' For column types: SQLite has dynamic typing, so DeclType is just the
' string from the CREATE TABLE statement ("INTEGER", "TEXT NOT NULL",
' "REAL DEFAULT 0", whatever). We expose it raw so you can see exactly
' what's there.
'
Option Explicit

Private m_db As cSQLite

Public Sub AttachDb(ByRef db As cSQLite)
    Set m_db = db
End Sub

' --- table listing ------------------------------------------------

' Return all user tables (excludes sqlite internal tables) as a Collection
' of names. Walk it with "For Each tbl In sch.ListTables".
Public Function ListTables() As Collection
    EnsureAttached
    Dim rs As cSQLiteResults
    Set rs = m_db.Query( _
        "SELECT name FROM sqlite_master " & _
        "WHERE type='table' AND name NOT LIKE 'sqlite_%' " & _
        "ORDER BY name")

    Dim out As New Collection
    Do While rs.MoveNext()
        out.Add rs.FieldText(0)
    Loop
    Set ListTables = out
End Function

' Same but as a 1D array (handy for combo boxes).
Public Function ListTablesArr() As String()
    Dim c As Collection
    Set c = ListTables
    Dim arr() As String, i As Long
    If c.Count = 0 Then
        ReDim arr(0)
        ListTablesArr = arr
        Exit Function
    End If
    ReDim arr(0 To c.Count - 1)
    For i = 0 To c.Count - 1
        arr(i) = c(i + 1)
    Next i
    ListTablesArr = arr
End Function

Public Function ListViews() As Collection
    EnsureAttached
    Dim rs As cSQLiteResults
    Set rs = m_db.Query( _
        "SELECT name FROM sqlite_master WHERE type='view' ORDER BY name")
    Dim out As New Collection
    Do While rs.MoveNext()
        out.Add rs.FieldText(0)
    Loop
    Set ListViews = out
End Function

Public Function ListIndexes(Optional ByVal sTable As String = "") As Collection
    EnsureAttached
    Dim rs As cSQLiteResults
    If Len(sTable) = 0 Then
        Set rs = m_db.Query( _
            "SELECT name FROM sqlite_master WHERE type='index' " & _
            "AND name NOT LIKE 'sqlite_%' ORDER BY name")
    Else
        Set rs = m_db.Query( _
            "SELECT name FROM sqlite_master WHERE type='index' " & _
            "AND tbl_name=? AND name NOT LIKE 'sqlite_%' ORDER BY name", sTable)
    End If
    Dim out As New Collection
    Do While rs.MoveNext()
        out.Add rs.FieldText(0)
    Loop
    Set ListIndexes = out
End Function

Public Function TableExists(ByVal sTable As String) As Boolean
    EnsureAttached
    Dim v As Variant
    v = m_db.Scalar( _
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", sTable)
    TableExists = Not IsEmpty(v)
End Function

' Check whether a column exists on a given table. Useful for migration
' code:
'   If Not sch.ColumnExists("users", "email") Then
'       sch.AddColumn "users", "email", "TEXT"
'   End If
' Returns False if the table itself doesn't exist (no error).
' Comparison is case-insensitive (SQLite's column names are).
Public Function ColumnExists(ByVal sTable As String, ByVal sColumn As String) As Boolean
    EnsureAttached
    Dim rs As cSQLiteResults
    Set rs = m_db.Query("PRAGMA table_info(" & QuoteIdent(sTable) & ")")
    Dim target As String
    target = LCase$(sColumn)
    Do While rs.MoveNext()
        If LCase$(rs.FieldText("name")) = target Then
            ColumnExists = True
            Exit Function
        End If
    Loop
End Function

' Check whether a named index exists.
Public Function IndexExists(ByVal sIndexName As String) As Boolean
    EnsureAttached
    Dim v As Variant
    v = m_db.Scalar( _
        "SELECT 1 FROM sqlite_master WHERE type='index' AND name=?", sIndexName)
    IndexExists = Not IsEmpty(v)
End Function

' Check whether a view exists.
Public Function ViewExists(ByVal sViewName As String) As Boolean
    EnsureAttached
    Dim v As Variant
    v = m_db.Scalar( _
        "SELECT 1 FROM sqlite_master WHERE type='view' AND name=?", sViewName)
    ViewExists = Not IsEmpty(v)
End Function

Public Function GetCreateSQL(ByVal sTable As String) As String
    EnsureAttached
    Dim v As Variant
    v = m_db.Scalar( _
        "SELECT sql FROM sqlite_master WHERE name=?", sTable)
    If Not IsNull(v) And Not IsEmpty(v) Then GetCreateSQL = CStr(v)
End Function

Public Function RowCount(ByVal sTable As String) As Currency
    EnsureAttached
    ' Quote the identifier safely with double-quotes; escape any embedded
    ' quotes by doubling them per SQL identifier rules.
    Dim qid As String
    qid = """" & Replace(sTable, """", """""") & """"
    Dim v As Variant
    v = m_db.Scalar("SELECT COUNT(*) FROM " & qid)
    If Not IsEmpty(v) Then RowCount = CCur(v)
End Function

' --- table description --------------------------------------------

' Return an array of cSQLiteColumn for the given table.
' Uses PRAGMA table_info() which is the canonical SQLite introspection.
Public Function DescribeTable(ByVal sTable As String) As Collection
    EnsureAttached

    ' Use PRAGMA table_xinfo for full info including hidden columns;
    ' fall back to table_info if the SQLite is too old (xinfo since 3.26).
    ' We'll just use table_info — it's universally supported and gives
    ' us name, type, notnull, dflt_value, pk.
    Dim rs As cSQLiteResults
    Set rs = m_db.Query("PRAGMA table_info(" & QuoteIdent(sTable) & ")")

    Dim out As New Collection
    Dim col As cSQLiteColumn  ' declared once, NOT "As New"
    Do While rs.MoveNext()
        Dim defaultStr As String
        Dim dv As Variant
        dv = rs("dflt_value")
        If IsNull(dv) Or IsEmpty(dv) Then
            defaultStr = ""
        Else
            defaultStr = CStr(dv)
        End If

        ' CRITICAL: Set ... New (not Dim As New) — auto-instantiation
        ' would reuse the same object across iterations, causing every
        ' Collection entry to point at the same column with the last
        ' iteration's values.
        Set col = New cSQLiteColumn
        col.SetFromPragma _
            cid:=rs.FieldInt("cid"), _
            colName:=rs.FieldText("name"), _
            DeclType:=rs.FieldText("type"), _
            NotNull:=(rs.FieldInt("notnull") <> 0), _
            defaultVal:=defaultStr, _
            PKOrdinal:=rs.FieldInt("pk")
        out.Add col
    Loop
    Set DescribeTable = out
End Function

' Print a human-readable description to a String (handy for debugging).
Public Function DescribeTableText(ByVal sTable As String) As String
    Dim cols As Collection
    Set cols = DescribeTable(sTable)
    If cols.Count = 0 Then
        DescribeTableText = "(table '" & sTable & "' not found or has no columns)"
        Exit Function
    End If

    Dim s As String, c As cSQLiteColumn
    s = "Table: " & sTable & vbCrLf
    s = s & "  " & PadRight("Column", 20) & PadRight("Type", 16) & _
            PadRight("Null?", 8) & PadRight("PK", 6) & "Default" & vbCrLf
    s = s & "  " & String$(70, "-") & vbCrLf
    For Each c In cols
        s = s & "  " & PadRight(c.Name, 20) & _
                PadRight(IIf(Len(c.DeclType) > 0, c.DeclType, "(none)"), 16) & _
                PadRight(IIf(c.NotNull, "NOT NULL", "NULL"), 8) & _
                PadRight(IIf(c.IsPrimaryKey, "PK" & IIf(c.PKOrdinal > 1, _
                    "(" & c.PKOrdinal & ")", ""), ""), 6) & _
                c.DefaultValue & vbCrLf
    Next c
    DescribeTableText = s
End Function

' --- light DDL helpers --------------------------------------------

' SQLite's ALTER TABLE is limited compared to other engines:
'   - RENAME TABLE: yes
'   - ADD COLUMN: yes
'   - DROP COLUMN: yes (since 3.35)
'   - RENAME COLUMN: yes (since 3.25)
'   - ALTER COLUMN TYPE: NO — requires the rebuild dance.
'
' For unsupported operations the standard pattern is:
'   1. CREATE TABLE new_t WITH the desired schema
'   2. INSERT INTO new_t SELECT ... FROM old_t
'   3. DROP TABLE old_t
'   4. ALTER TABLE new_t RENAME TO old_t
' Wrap that in a transaction. We don't automate it here because
' the SELECT projection depends on what you're trying to change.

Public Sub CreateTable(ByVal sTable As String, ByVal columnDefs As String)
    EnsureAttached
    m_db.Execute "CREATE TABLE " & QuoteIdent(sTable) & " (" & columnDefs & ")"
End Sub

Public Sub CreateTableIfMissing(ByVal sTable As String, ByVal columnDefs As String)
    EnsureAttached
    m_db.Execute "CREATE TABLE IF NOT EXISTS " & QuoteIdent(sTable) & _
        " (" & columnDefs & ")"
End Sub

Public Sub DropTable(ByVal sTable As String, Optional ByVal ifExists As Boolean = True)
    EnsureAttached
    Dim s As String
    s = "DROP TABLE "
    If ifExists Then s = s & "IF EXISTS "
    s = s & QuoteIdent(sTable)
    m_db.Execute s
End Sub

Public Sub RenameTable(ByVal sOldName As String, ByVal sNewName As String)
    EnsureAttached
    m_db.Execute "ALTER TABLE " & QuoteIdent(sOldName) & _
                 " RENAME TO " & QuoteIdent(sNewName)
End Sub

Public Sub AddColumn(ByVal sTable As String, ByVal sColName As String, _
                     ByVal sColType As String, _
                     Optional ByVal sExtra As String = "")
    ' sExtra: extra DDL after the type, e.g. "NOT NULL DEFAULT 0"
    EnsureAttached
    Dim s As String
    s = "ALTER TABLE " & QuoteIdent(sTable) & _
        " ADD COLUMN " & QuoteIdent(sColName) & " " & sColType
    If Len(sExtra) > 0 Then s = s & " " & sExtra
    m_db.Execute s
End Sub

' Requires SQLite 3.25+. We don't version-gate; if it fails, the error
' message will say so plainly.
Public Sub RenameColumn(ByVal sTable As String, _
                        ByVal sOldCol As String, ByVal sNewCol As String)
    EnsureAttached
    m_db.Execute "ALTER TABLE " & QuoteIdent(sTable) & _
                 " RENAME COLUMN " & QuoteIdent(sOldCol) & _
                 " TO " & QuoteIdent(sNewCol)
End Sub

' Requires SQLite 3.35+.
Public Sub DropColumn(ByVal sTable As String, ByVal sColName As String)
    EnsureAttached
    m_db.Execute "ALTER TABLE " & QuoteIdent(sTable) & _
                 " DROP COLUMN " & QuoteIdent(sColName)
End Sub

Public Sub CreateIndex(ByVal sIndexName As String, ByVal sTable As String, _
                       ByVal sColumns As String, _
                       Optional ByVal unique As Boolean = False)
    EnsureAttached
    Dim s As String
    s = "CREATE "
    If unique Then s = s & "UNIQUE "
    s = s & "INDEX " & QuoteIdent(sIndexName) & " ON " & _
        QuoteIdent(sTable) & " (" & sColumns & ")"
    m_db.Execute s
End Sub

Public Sub DropIndex(ByVal sIndexName As String, _
                     Optional ByVal ifExists As Boolean = True)
    EnsureAttached
    Dim s As String
    s = "DROP INDEX "
    If ifExists Then s = s & "IF EXISTS "
    s = s & QuoteIdent(sIndexName)
    m_db.Execute s
End Sub

' --- application versioning --------------------------------------
'
' SQLite reserves two 32-bit integers in the database header for app use:
'   user_version    — for your schema version number (you set/read this)
'   application_id  — magic number identifying which app owns this file
' Both default to 0. Unlike everything else here, these survive across
' VACUUM and are part of the official "this DB belongs to my app v3"
' pattern.
'
' Typical migration code:
'   Select Case sch.UserVersion
'       Case 0
'           sch.CreateTable "users", "id INTEGER PRIMARY KEY, name TEXT"
'           sch.UserVersion = 1
'       Case 1
'           sch.AddColumn "users", "email", "TEXT"
'           sch.UserVersion = 2
'   End Select

Public Property Get UserVersion() As Long
    EnsureAttached
    Dim v As Variant
    v = m_db.Scalar("PRAGMA user_version")
    If Not IsEmpty(v) Then UserVersion = CLng(v)
End Property

Public Property Let UserVersion(ByVal n As Long)
    EnsureAttached
    ' user_version doesn't accept '?' bindings — must be inlined.
    ' n is a Long so no escaping concern.
    m_db.Execute "PRAGMA user_version = " & n
End Property

' Read-only counter that SQLite increments every time the schema changes.
' Useful for cache invalidation: stash this, do work, check if it moved.
Public Property Get SchemaVersion() As Long
    EnsureAttached
    Dim v As Variant
    v = m_db.Scalar("PRAGMA schema_version")
    If Not IsEmpty(v) Then SchemaVersion = CLng(v)
End Property

' Magic number identifying which app owns this database. Settable.
' Common practice: pick four ASCII bytes that mean something to you,
' e.g. ApplicationID = &H4D594150  ' "MYAP"
Public Property Get ApplicationID() As Long
    EnsureAttached
    Dim v As Variant
    v = m_db.Scalar("PRAGMA application_id")
    If Not IsEmpty(v) Then ApplicationID = CLng(v)
End Property

Public Property Let ApplicationID(ByVal n As Long)
    EnsureAttached
    m_db.Execute "PRAGMA application_id = " & n
End Property

' --- maintenance / compaction ------------------------------------
'
' VACUUM rebuilds the whole DB file from scratch — reclaims free pages,
' defragments, repacks indexes. Run this after big deletes or as periodic
' maintenance. Requires no other connections and locks the DB for the
' duration. On a multi-GB DB this can take a while, so don't call it
' from a UI thread without warning the user.
'
' Returns the size delta (bytes saved). Negative means the file grew.
Public Function CompactDB() As Double
    EnsureAttached
    Dim before As Double, after As Double
    before = DatabaseSize
    m_db.Execute "VACUUM"
    after = DatabaseSize
    CompactDB = before - after
End Function

' Return the current on-disk size of the main database in bytes.
' Computed from page_count * page_size — accurate even before VACUUM
' has flushed unused pages. Returned as Double because file sizes can
' easily exceed Long's 2GB ceiling, and Double represents whole bytes
' losslessly up to 2^53 (~9 petabytes).
Public Property Get DatabaseSize() As Double
    EnsureAttached
    Dim pageCount As Variant, pageSize As Variant
    pageCount = m_db.Scalar("PRAGMA page_count")
    pageSize = m_db.Scalar("PRAGMA page_size")
    If IsEmpty(pageCount) Or IsEmpty(pageSize) Then Exit Property
    DatabaseSize = CDbl(pageCount) * CDbl(pageSize)
End Property

' Number of unused pages in the file. After lots of deletes this grows;
' VACUUM resets it to zero. Useful for "should I run VACUUM?" decisions.
Public Property Get FreelistPages() As Long
    EnsureAttached
    Dim v As Variant
    v = m_db.Scalar("PRAGMA freelist_count")
    If Not IsEmpty(v) Then FreelistPages = CLng(v)
End Property

' Free space estimate in bytes (freelist_count * page_size).
Public Property Get FreeSpace() As Double
    EnsureAttached
    Dim pageSize As Variant
    pageSize = m_db.Scalar("PRAGMA page_size")
    If IsEmpty(pageSize) Then Exit Property
    FreeSpace = CDbl(FreelistPages) * CDbl(pageSize)
End Property

' Update the query planner's statistics. SQLite recommends running this
' periodically, especially before closing long-lived connections, and
' especially after bulk inserts. Cheap and worth doing.
Public Sub Optimize()
    EnsureAttached
    m_db.Execute "PRAGMA optimize"
End Sub

' Force a full ANALYZE on every table. Heavier than Optimize (which
' only re-analyzes tables that need it), but guaranteed thorough.
Public Sub Analyze(Optional ByVal sTable As String = "")
    EnsureAttached
    If Len(sTable) = 0 Then
        m_db.Execute "ANALYZE"
    Else
        m_db.Execute "ANALYZE " & QuoteIdent(sTable)
    End If
End Sub

' Rebuild indexes. Rarely needed; collation changes are the main reason.
Public Sub Reindex(Optional ByVal sName As String = "")
    EnsureAttached
    If Len(sName) = 0 Then
        m_db.Execute "REINDEX"
    Else
        m_db.Execute "REINDEX " & QuoteIdent(sName)
    End If
End Sub

' Run integrity_check. Returns "ok" if the database is healthy, or
' a multi-line error report if not. This reads every page and verifies
' indexes match table content — slow on big DBs but the gold standard
' check.
Public Function IntegrityCheck(Optional ByVal maxErrors As Long = 100) As String
    EnsureAttached
    Dim rs As cSQLiteResults
    Set rs = m_db.Query("PRAGMA integrity_check(" & maxErrors & ")")
    Dim out As String
    Do While rs.MoveNext()
        If Len(out) > 0 Then out = out & vbCrLf
        out = out & rs.FieldText(0)
    Loop
    IntegrityCheck = out
End Function

' Quick check is much faster than full integrity_check but doesn't
' verify index/table consistency. Returns "ok" or a list of problems.
Public Function QuickCheck(Optional ByVal maxErrors As Long = 100) As String
    EnsureAttached
    Dim rs As cSQLiteResults
    Set rs = m_db.Query("PRAGMA quick_check(" & maxErrors & ")")
    Dim out As String
    Do While rs.MoveNext()
        If Len(out) > 0 Then out = out & vbCrLf
        out = out & rs.FieldText(0)
    Loop
    QuickCheck = out
End Function

' Force a checkpoint of the write-ahead log if the DB is in WAL mode.
' Modes:
'   "PASSIVE"  - default; checkpoint as much as possible without waiting
'   "FULL"     - block writers, then checkpoint
'   "RESTART"  - like FULL plus restart the WAL file at offset 0
'   "TRUNCATE" - like RESTART plus truncate the WAL file to zero length
' No-op (and no error) if the DB is in rollback-journal mode.
Public Sub WALCheckpoint(Optional ByVal mode As String = "PASSIVE")
    EnsureAttached
    m_db.Execute "PRAGMA wal_checkpoint(" & mode & ")"
End Sub

' --- internals ----------------------------------------------------

Private Sub EnsureAttached()
    If m_db Is Nothing Then
        Err.Raise vbObjectError + 30400, "cSQLiteSchema", _
            "no database attached - call AttachDb first"
    End If
    If Not m_db.IsOpen Then
        Err.Raise vbObjectError + 30401, "cSQLiteSchema", _
            "attached database is not open"
    End If
End Sub

' Quote a SQLite identifier. Uses double-quotes (SQL-standard) and
' escapes any embedded double-quotes by doubling them.
Public Function QuoteIdent(ByVal s As String) As String
    QuoteIdent = """" & Replace(s, """", """""") & """"
End Function

Private Function PadRight(ByVal s As String, ByVal width As Long) As String
    If Len(s) >= width Then
        PadRight = s & " "
    Else
        PadRight = s & Space$(width - Len(s))
    End If
End Function
