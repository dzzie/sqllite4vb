VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SQLite VB6 Test"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRunTest 
      Caption         =   "Basic"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdUnicode 
      Caption         =   "Unicode"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdTxn 
      Caption         =   "Transaction"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdBlob 
      Caption         =   "Blob"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdClasses 
      Caption         =   "Class API"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSchema 
      Caption         =   "Schema"
      Height          =   495
      Left            =   6720
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdMaint 
      Caption         =   "Maintenance"
      Height          =   495
      Left            =   8040
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   9600
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   720
      Width           =   10575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Log(ByVal s As String)
    txtOutput.Text = txtOutput.Text & s & vbCrLf
    txtOutput.SelStart = Len(txtOutput.Text)
End Sub

Private Sub cmdClear_Click()
    txtOutput.Text = ""
End Sub

Private Function HexBytes(ByRef b() As Byte) As String
    Dim i As Long, s As String
    On Error Resume Next
    For i = LBound(b) To UBound(b)
        If i > LBound(b) Then s = s & " "
        s = s & Right$("0" & Hex$(b(i)), 2)
    Next i
    HexBytes = s
End Function

' ===== TEST 1: basic users table (low-level API) ==================

Private Sub cmdRunTest_Click()
    On Error GoTo Fail
    txtOutput.Text = ""

    Dim hDb As Long, hStmt As Long, rc As Long
    Dim sDbPath As String
    sDbPath = App.Path & "\test_users.db"

    Log "SQLite version: " & sqlite3_libversion() & _
        "  (number=" & sqlite3_libversion_number() & ")"
    Log "Threadsafe: " & sqlite3_threadsafe()
    Log ""

    If Dir(sDbPath) <> "" Then Kill sDbPath

    Log "[1] Opening " & sDbPath
    rc = sqlite3_open(sDbPath, hDb)
    If rc <> SQLITE_OK Then
        Log "    open failed rc=" & rc & " msg=" & sqlite3_errmsg(hDb)
        GoTo Cleanup
    End If
    Log "    db handle = &H" & Hex$(hDb)

    Log "[2] Creating table"
    SQLiteExec hDb, _
        "CREATE TABLE users (" & _
        "  id   INTEGER PRIMARY KEY AUTOINCREMENT," & _
        "  name TEXT NOT NULL," & _
        "  age  INTEGER NOT NULL" & _
        ")"

    Log "[3] Inserting rows"
    rc = sqlite3_prepare_v2(hDb, _
            "INSERT INTO users (name, age) VALUES (?, ?)", hStmt)
    If rc <> SQLITE_OK Then GoTo Cleanup

    InsertUser hDb, hStmt, "Alice", 30
    InsertUser hDb, hStmt, "Bob", 45
    InsertUser hDb, hStmt, "Charlie", 27
    InsertUser hDb, hStmt, "Diana", 52

    sqlite3_finalize hStmt: hStmt = 0
    Log "    " & sqlite3_total_changes(hDb) & " rows inserted"

    Log ""
    Log "[4] SELECT * FROM users"
    DumpTable hDb, "SELECT id, name, age FROM users ORDER BY id"

Cleanup:
    If hStmt <> 0 Then sqlite3_finalize hStmt
    If hDb <> 0 Then sqlite3_close hDb
    Log "": Log "[done]"
    Exit Sub
Fail:
    Log "*** ERROR " & Err.Number & ": " & Err.Description
    Resume Cleanup
End Sub

Private Sub InsertUser(ByVal hDb As Long, ByVal hStmt As Long, _
                       ByVal sName As String, ByVal nAge As Long)
    Dim rc As Long
    sqlite3_reset hStmt
    sqlite3_clear_bindings hStmt
    sqlite3_bind_text hStmt, 1, sName
    sqlite3_bind_int hStmt, 2, nAge
    rc = sqlite3_step(hStmt)
    If rc = SQLITE_DONE Then
        Log "    + " & sName & " (" & nAge & ")"
    Else
        Log "    insert failed: " & sqlite3_errmsg(hDb)
    End If
End Sub

Private Sub DumpTable(ByVal hDb As Long, ByVal sSql As String)
    Dim hStmt As Long, rc As Long, nCols As Long, i As Long, sLine As String

    rc = sqlite3_prepare_v2(hDb, sSql, hStmt)
    If rc <> SQLITE_OK Then Exit Sub

    nCols = sqlite3_column_count(hStmt)
    sLine = "    "
    For i = 0 To nCols - 1
        If i > 0 Then sLine = sLine & " | "
        sLine = sLine & sqlite3_column_name(hStmt, i)
    Next i
    Log sLine
    Log "    " & String$(60, "-")

    Do
        rc = sqlite3_step(hStmt)
        If rc = SQLITE_ROW Then
            sLine = "    "
            For i = 0 To nCols - 1
                If i > 0 Then sLine = sLine & " | "
                sLine = sLine & sqlite3_column_text(hStmt, i)
            Next i
            Log sLine
        ElseIf rc = SQLITE_DONE Then
            Exit Do
        Else
            Exit Do
        End If
    Loop

    sqlite3_finalize hStmt
End Sub

' ===== TEST 2: Unicode roundtrip ==================================

Private Sub cmdUnicode_Click()
    On Error GoTo Fail
    txtOutput.Text = ""
    Log "=== Unicode roundtrip test ===" & vbCrLf

    Dim hDb As Long, hStmt As Long, rc As Long
    Dim sDbPath As String
    sDbPath = App.Path & "\test_unicode.db"
    If Dir(sDbPath) <> "" Then Kill sDbPath

    rc = sqlite3_open(sDbPath, hDb)
    If rc <> SQLITE_OK Then GoTo Fail
    SQLiteExec hDb, "CREATE TABLE t (id INTEGER PRIMARY KEY, label TEXT, payload TEXT)"

    Dim cases(0 To 4) As String, labels(0 To 4) As String
    labels(0) = "ASCII baseline":          cases(0) = "Hello, World!"
    labels(1) = "Cyrillic":                cases(1) = ChrW$(&H41F) & ChrW$(&H440) & ChrW$(&H438) & ChrW$(&H432) & ChrW$(&H435) & ChrW$(&H442)
    labels(2) = "CJK (Chinese)":           cases(2) = ChrW$(&H4F60) & ChrW$(&H597D) & ChrW$(&H4E16) & ChrW$(&H754C)
    labels(3) = "Accented Latin":          cases(3) = "caf" & ChrW$(&HE9) & " na" & ChrW$(&HEF) & "ve r" & ChrW$(&HE9) & "sum" & ChrW$(&HE9)
    labels(4) = "Emoji (surrogate pair)":  cases(4) = "smile " & ChrW$(&HD83D) & ChrW$(&HDE00) & " here"

    rc = sqlite3_prepare_v2(hDb, "INSERT INTO t (label, payload) VALUES (?, ?)", hStmt)
    If rc <> SQLITE_OK Then GoTo Fail

    Dim i As Long
    For i = 0 To UBound(cases)
        sqlite3_reset hStmt
        sqlite3_clear_bindings hStmt
        sqlite3_bind_text hStmt, 1, labels(i)
        sqlite3_bind_text hStmt, 2, cases(i)
        sqlite3_step hStmt
    Next i
    sqlite3_finalize hStmt: hStmt = 0

    rc = sqlite3_prepare_v2(hDb, "SELECT id, label, payload FROM t ORDER BY id", hStmt)
    Dim allPass As Boolean: allPass = True

    Do
        rc = sqlite3_step(hStmt)
        If rc <> SQLITE_ROW Then Exit Do
        Dim id As Long, gotPayload As String
        id = sqlite3_column_int(hStmt, 0)
        gotPayload = sqlite3_column_text(hStmt, 2)
        Dim ok As Boolean
        ok = (StrComp(gotPayload, cases(id - 1), vbBinaryCompare) = 0)
        Log "[" & id & "] " & sqlite3_column_text(hStmt, 1) & " -> " & IIf(ok, "PASS", "FAIL")
        Log "    bytes: " & HexUtf16(gotPayload)
        If Not ok Then allPass = False
    Loop
    sqlite3_finalize hStmt: hStmt = 0
    Log ""
    Log "=== Result: " & IIf(allPass, "ALL PASS", "FAILURES") & " ==="

Cleanup:
    If hStmt <> 0 Then sqlite3_finalize hStmt
    If hDb <> 0 Then sqlite3_close hDb
    Exit Sub
Fail:
    Log "*** ERROR " & Err.Number & ": " & Err.Description
    Resume Cleanup
End Sub

Private Function HexUtf16(ByVal s As String) As String
    Dim b() As Byte
    b = s
    HexUtf16 = HexBytes(b)
End Function

' ===== TEST 3: transaction performance ============================

Private Sub cmdTxn_Click()
    On Error GoTo Fail
    txtOutput.Text = ""
    Log "=== Transaction performance test ===" & vbCrLf

    Dim hDb As Long, hStmt As Long, rc As Long
    Dim sDbPath As String
    sDbPath = App.Path & "\test_txn.db"
    If Dir(sDbPath) <> "" Then Kill sDbPath

    rc = sqlite3_open(sDbPath, hDb): If rc <> SQLITE_OK Then GoTo Fail
    SQLiteExec hDb, "CREATE TABLE items (id INTEGER PRIMARY KEY, name TEXT, value REAL)"

    Const n As Long = 10000
    Log "Inserting " & n & " rows in a single transaction..."

    SQLiteExec hDb, "BEGIN TRANSACTION"
    rc = sqlite3_prepare_v2(hDb, "INSERT INTO items (name, value) VALUES (?, ?)", hStmt)

    Dim t0 As Single, t1 As Single, i As Long
    t0 = Timer
    For i = 1 To n
        sqlite3_reset hStmt
        sqlite3_bind_text hStmt, 1, "item_" & i
        sqlite3_bind_double hStmt, 2, Rnd * 1000#
        sqlite3_step hStmt
    Next i
    sqlite3_finalize hStmt: hStmt = 0
    SQLiteExec hDb, "COMMIT"
    t1 = Timer

    Dim elapsed As Single
    elapsed = t1 - t0: If elapsed <= 0 Then elapsed = 0.001
    Log "  elapsed: " & Format$(elapsed, "0.000") & " sec"
    Log "  rate   : " & Format$(n / elapsed, "#,##0") & " inserts/sec"

Cleanup:
    If hStmt <> 0 Then sqlite3_finalize hStmt
    If hDb <> 0 Then sqlite3_close hDb
    Exit Sub
Fail:
    Log "*** ERROR " & Err.Number & ": " & Err.Description
    Resume Cleanup
End Sub

' ===== TEST 4: blob roundtrip =====================================

Private Sub cmdBlob_Click()
    On Error GoTo Fail
    txtOutput.Text = ""
    Log "=== Blob roundtrip test ===" & vbCrLf

    Dim hDb As Long, hStmt As Long, rc As Long
    Dim sDbPath As String
    sDbPath = App.Path & "\test_blob.db"
    If Dir(sDbPath) <> "" Then Kill sDbPath

    rc = sqlite3_open(sDbPath, hDb): If rc <> SQLITE_OK Then GoTo Fail
    SQLiteExec hDb, "CREATE TABLE files (id INTEGER PRIMARY KEY, name TEXT, data BLOB)"

    Dim tricky(0 To 9) As Byte, i As Long
    tricky(0) = &HDE: tricky(1) = &HAD: tricky(2) = &HBE: tricky(3) = &HEF
    tricky(4) = &H0:  tricky(5) = &H0:  tricky(6) = &H0:  tricky(7) = &HFF
    tricky(8) = &H7F: tricky(9) = &H80

    rc = sqlite3_prepare_v2(hDb, "INSERT INTO files (name, data) VALUES (?, ?)", hStmt)
    sqlite3_bind_text hStmt, 1, "tricky"
    sqlite3_bind_blob hStmt, 2, tricky
    sqlite3_step hStmt
    sqlite3_finalize hStmt: hStmt = 0

    rc = sqlite3_prepare_v2(hDb, "SELECT data FROM files", hStmt)
    sqlite3_step hStmt
    Dim got() As Byte
    got = sqlite3_column_blob(hStmt, 0)
    Log "Sent: " & HexBytes(tricky)
    Log "Got : " & HexBytes(got)
    sqlite3_finalize hStmt: hStmt = 0

Cleanup:
    If hStmt <> 0 Then sqlite3_finalize hStmt
    If hDb <> 0 Then sqlite3_close hDb
    Exit Sub
Fail:
    Log "*** ERROR " & Err.Number & ": " & Err.Description
    Resume Cleanup
End Sub

' ===== TEST 5: high-level class API ===============================
'
' This is what day-to-day code will look like.

Private Sub cmdClasses_Click()
    On Error GoTo Fail
    txtOutput.Text = ""
    Log "=== Class API test (cSQLite / cSQLiteStatement / cSQLiteResults) ===" & vbCrLf

    Dim sDbPath As String
    sDbPath = App.Path & "\test_classes.db"
    If Dir(sDbPath) <> "" Then Kill sDbPath

    ' --- Open & create -------------------------------------------------
    Dim db As New cSQLite
    db.OpenDB sDbPath
    Log "Opened: version=" & db.Version & "  path=" & db.Path
    Log ""

    db.Execute _
        "CREATE TABLE users (" & _
        "  id INTEGER PRIMARY KEY AUTOINCREMENT," & _
        "  name TEXT NOT NULL," & _
        "  age INTEGER," & _
        "  joined TEXT," & _
        "  notes TEXT" & _
        ")"

    ' --- ExecInsert with mixed types ----------------------------------
    Log "[1] ExecInsert (parameterized, safe vs apostrophes)"
    db.ExecInsert "users", "name,age,joined,notes", "Alice", 30, Now, "first user"
    db.ExecInsert "users", "name,age,joined,notes", "O'Brien", 45, Now, "tests apostrophe"
    db.ExecInsert "users", "name,age,joined,notes", "Charlie", 27, Now, Null
    db.ExecInsert "users", "name,age,joined,notes", "Diana", 52, Now, _
        "name with " & ChrW$(&HE9) & " accent and ; semicolons; etc."
    Log "    " & db.RowsAffected & " row(s) affected by last insert"
    Log "    last rowid = " & CStr(Int64FromCurrency(db.LastInsertRowID))
    Log ""

    ' --- Query with rs("name") access ---------------------------------
    Log "[2] Query with by-name field access"
    Dim rs As cSQLiteResults
    Set rs = db.Query("SELECT id, name, age, notes FROM users ORDER BY id")
    Do While rs.MoveNext()
        Log "    " & rs("id") & " | " & rs("name") & " | age=" & rs("age") & _
            " | " & IIf(IsNull(rs("notes")), "(null)", rs("notes"))
    Loop
    rs.CloseRS
    Log ""

    ' --- Parameterized query ------------------------------------------
    Log "[3] Parameterized query"
    Set rs = db.Query("SELECT name, age FROM users WHERE age >= ? ORDER BY age DESC", 30)
    Do While rs.MoveNext()
        Log "    " & rs.FieldText("name") & "  (age " & rs.FieldInt("age") & ")"
    Loop
    Log ""

    ' --- Scalar -------------------------------------------------------
    Log "[4] Scalar query"
    Log "    avg age = " & Format$(db.Scalar("SELECT AVG(age) FROM users"), "0.00")
    Log "    Bob exists? = " & (Not IsEmpty(db.Scalar("SELECT 1 FROM users WHERE name=?", "Bob")))
    Log "    Alice age = " & db.Scalar("SELECT age FROM users WHERE name=?", "Alice")
    Log ""

    ' --- ExecUpdate ---------------------------------------------------
    Log "[5] ExecUpdate (criteria can be literal OR parameterized)"
    Dim affected As Long
    ' Literal criteria:
    affected = db.ExecUpdate("users", "WHERE name='Alice'", "age,notes", 31, "updated")
    Log "    literal-criteria update: " & affected & " row(s)"
    ' Parameterized criteria (safe with untrusted values):
    affected = db.ExecUpdate("users", "WHERE id=?", "age", 99, 2)
    '                                          ^placeholder       ^SET val ^criteria val
    Log "    parameterized-criteria update: " & affected & " row(s)"
    Log "    Alice's age now = " & db.Scalar("SELECT age FROM users WHERE name='Alice'")
    Log "    user id=2 age now = " & db.Scalar("SELECT age FROM users WHERE id=?", 2)
    Log ""

    ' --- Transaction with prepared statement reuse --------------------
    Log "[6] 1000 inserts in a transaction with statement reuse"
    db.Execute "CREATE TABLE big (n INTEGER, label TEXT)"
    Dim t0 As Single, t1 As Single
    t0 = Timer
    db.BeginTrans
    Dim stmt As cSQLiteStatement
    Set stmt = db.Prepare("INSERT INTO big (n, label) VALUES (?, ?)")
    Dim i As Long
    For i = 1 To 1000
        stmt.Bind 1, i
        stmt.Bind 2, "row " & i
        stmt.Execute
    Next i
    Set stmt = Nothing  ' triggers Class_Terminate -> finalize
    db.CommitTrans
    t1 = Timer
    Log "    inserted 1000 rows in " & Format$(t1 - t0, "0.000") & " sec"
    Log "    count = " & db.Scalar("SELECT COUNT(*) FROM big")
    Log ""

    ' --- LoadAll for small result sets --------------------------------
    Log "[7] LoadAll() — read first 5 big rows into 2D array"
    Set rs = db.Query("SELECT n, label FROM big WHERE n <= 5 ORDER BY n")
    Dim data As Variant
    data = rs.LoadAll
    If Not IsEmpty(data) Then
        Dim r As Long
        For r = LBound(data, 1) To UBound(data, 1)
            Log "    " & data(r, 0) & " | " & data(r, 1)
        Next r
    End If
    Log ""

    ' --- Implicit cleanup ---------------------------------------------
    Log "[8] Implicit cleanup test"
    Log "    db is open? " & db.IsOpen
    Set db = Nothing  ' triggers Class_Terminate -> close
    Log "    db variable set to Nothing — implicit close happened"
    Log ""
    Log "=== done ==="

    Exit Sub
Fail:
    Log "*** ERROR " & Err.Number & ": " & Err.Description
End Sub

' ===== TEST 7: maintenance / compaction ===========================
'
' Populate a table with bulky data, delete most of it, observe that
' the file doesn't shrink on its own, then VACUUM and watch the space
' get reclaimed. Also demos integrity check, optimize, and freelist
' inspection.

Private Sub cmdMaint_Click()
    On Error GoTo Fail
    txtOutput.Text = ""
    Log "=== Maintenance / compaction test ===" & vbCrLf

    Dim sDbPath As String
    sDbPath = App.Path & "\test_maint.db"
    If Dir(sDbPath) <> "" Then Kill sDbPath

    Dim db As New cSQLite
    db.OpenDB sDbPath
    Dim sch As New cSQLiteSchema
    sch.AttachDb db

    ' Build a chunky table — 5000 rows with ~200 bytes of text each.
    sch.CreateTable "bulk", _
        "id INTEGER PRIMARY KEY, payload TEXT NOT NULL"

    Log "[1] Populating bulk table (5000 rows, ~200 bytes each)..."
    db.BeginTrans
    Dim stmt As cSQLiteStatement
    Set stmt = db.Prepare("INSERT INTO bulk (payload) VALUES (?)")
    Dim filler As String
    filler = String$(200, "x")
    Dim i As Long
    For i = 1 To 5000
        stmt.Bind 1, filler & " row " & i
        stmt.Execute
    Next i
    Set stmt = Nothing
    db.CommitTrans

    Dim sizeBefore As Double
    sizeBefore = sch.DatabaseSize
    Log "    rows         : " & db.Scalar("SELECT COUNT(*) FROM bulk")
    Log "    file size    : " & FmtBytes(sizeBefore)
    Log "    free pages   : " & sch.FreelistPages
    Log "    free space   : " & FmtBytes(sch.FreeSpace)
    Log ""

    ' --- delete most rows -----------------------------------------
    Log "[2] Deleting 4500 of the 5000 rows..."
    db.Execute "DELETE FROM bulk WHERE id > 500"
    Log "    rows now     : " & db.Scalar("SELECT COUNT(*) FROM bulk")
    Log "    file size    : " & FmtBytes(sch.DatabaseSize) & "  (note: still huge!)"
    Log "    free pages   : " & sch.FreelistPages
    Log "    free space   : " & FmtBytes(sch.FreeSpace)
    Log "    -> SQLite parks deleted pages on a freelist, doesn't shrink the file."
    Log ""

    ' --- VACUUM ---------------------------------------------------
    Log "[3] CompactDB (VACUUM) — rebuilding the file..."
    Dim t0 As Single
    t0 = Timer
    Dim saved As Double
    saved = sch.CompactDB
    Log "    elapsed      : " & Format$(Timer - t0, "0.000") & " sec"
    Log "    bytes saved  : " & FmtBytes(saved)
    Log "    file size    : " & FmtBytes(sch.DatabaseSize)
    Log "    free pages   : " & sch.FreelistPages & "  (back to zero)"
    Log ""

    ' --- integrity check -----------------------------------------
    Log "[4] Integrity check"
    Dim chk As String
    chk = sch.IntegrityCheck
    Log "    result: " & chk
    Log ""

    ' --- quick check ---------------------------------------------
    Log "[5] Quick check"
    Log "    result: " & sch.QuickCheck
    Log ""

    ' --- optimize ------------------------------------------------
    Log "[6] PRAGMA optimize (updates query planner stats)"
    sch.Optimize
    Log "    done"
    Log ""

    ' --- analyze -------------------------------------------------
    Log "[7] Full ANALYZE"
    sch.Analyze
    Log "    done"
    Log ""

    db.CloseDB
    Log "=== done ==="
    Exit Sub
Fail:
    Log "*** ERROR " & Err.Number & ": " & Err.Description
End Sub

' Pretty-print byte count
Private Function FmtBytes(ByVal n As Double) As String
    If n < 1024 Then
        FmtBytes = CStr(CLng(n)) & " B"
    ElseIf n < 1048576# Then
        FmtBytes = Format$(n / 1024, "0.0") & " KB"
    ElseIf n < 1073741824# Then
        FmtBytes = Format$(n / 1048576#, "0.00") & " MB"
    Else
        FmtBytes = Format$(n / 1073741824#, "0.00") & " GB"
    End If
End Function

' ===== TEST 6: schema introspection ===============================

Private Sub cmdSchema_Click()
    On Error GoTo Fail
    txtOutput.Text = ""
    Log "=== Schema introspection test ===" & vbCrLf

    Dim sDbPath As String
    sDbPath = App.Path & "\test_schema.db"
    If Dir(sDbPath) <> "" Then Kill sDbPath

    Dim db As New cSQLite
    db.OpenDB sDbPath

    Dim sch As New cSQLiteSchema
    sch.AttachDb db

    ' Build a few tables to introspect
    sch.CreateTable "users", _
        "id INTEGER PRIMARY KEY AUTOINCREMENT," & _
        "name TEXT NOT NULL," & _
        "email TEXT," & _
        "age INTEGER DEFAULT 0," & _
        "active INTEGER NOT NULL DEFAULT 1"

    sch.CreateTable "products", _
        "sku TEXT PRIMARY KEY," & _
        "title TEXT NOT NULL," & _
        "price REAL," & _
        "stock INTEGER DEFAULT 0"

    sch.CreateTableIfMissing "audit_log", _
        "ts TEXT, msg TEXT"

    sch.CreateIndex "idx_users_email", "users", "email", Unique:=True
    sch.CreateIndex "idx_users_name", "users", "name"

    ' --- list ---
    Log "[1] ListTables()"
    Dim tbl As Variant
    For Each tbl In sch.ListTables
        Log "    - " & tbl & "  (" & sch.RowCount(CStr(tbl)) & " rows)"
    Next
    Log ""

    Log "[2] ListIndexes('users')"
    Dim idx As Variant
    For Each idx In sch.ListIndexes("users")
        Log "    - " & idx
    Next
    Log ""

    Log "[3] TableExists tests"
    Log "    users     -> " & sch.TableExists("users")
    Log "    bogus     -> " & sch.TableExists("bogus")
    Log ""

    Log "[4] DescribeTable('users')"
    Log sch.DescribeTableText("users")

    Log "[5] GetCreateSQL('users')"
    Log "    " & sch.GetCreateSQL("users")
    Log ""

    ' --- modify ---
    Log "[6] AddColumn"
    sch.AddColumn "users", "phone", "TEXT"
    Log "    after add:"
    Log sch.DescribeTableText("users")

    Log "[7] RenameColumn (requires SQLite 3.25+)"
    sch.RenameColumn "users", "phone", "phone_number"
    Log "    after rename:"
    Log sch.DescribeTableText("users")

    Log "[8] RenameTable"
    sch.RenameTable "audit_log", "events"
    Log "    tables now:"
    For Each tbl In sch.ListTables
        Log "    - " & tbl
    Next
    Log ""

    Log "[9] DropTable (events) and DropIndex (idx_users_name)"
    sch.DropTable "events"
    sch.DropIndex "idx_users_name"
    Log "    tables: " & Join(sch.ListTablesArr, ", ")
    Log "    user indexes: "
    For Each idx In sch.ListIndexes("users")
        Log "      " & idx
    Next
    Log ""

    db.CloseDB
    Log "=== done ==="
    Exit Sub
Fail:
    Log "*** ERROR " & Err.Number & ": " & Err.Description
End Sub
