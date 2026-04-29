VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com
'
' cSQLite — ADO-flavored connection wrapper around sqlite3_vb6.dll
'
' Lifecycle:
'   Dim db As New cSQLite
'   db.OpenDB "myapp.db"
'   ... use it ...
'   ' db closes automatically when it goes out of scope (Class_Terminate)
'   ' or you can call db.CloseDB explicitly.
'
' One-shot DDL/DML:
'   db.Execute "CREATE TABLE t (id INTEGER PRIMARY KEY, name TEXT)"
'
' Convenience builders (parameterized under the hood — safe & fast):
'   db.ExecInsert "users", "name,age", "O'Brien", 42
'   db.ExecUpdate "users", "WHERE id=5", "name,age", "Jones", 31
'
' Queries:
'   Dim rs As cSQLiteResults
'   Set rs = db.Query("SELECT name, age FROM users WHERE age > ?", 30)
'   Do While rs.MoveNext()
'       Debug.Print rs("name"), rs("age")
'   Loop
'
' Prepared statements (for tight loops or custom binding):
'   Dim stmt As cSQLiteStatement
'   Set stmt = db.Prepare("INSERT INTO logs (msg, level) VALUES (?, ?)")
'   For i = 1 To 1000
'       stmt.Bind 1, "msg " & i
'       stmt.Bind 2, i Mod 5
'       stmt.Execute
'   Next
'   ' stmt finalized automatically when it goes out of scope
'
' Transactions:
'   db.BeginTrans
'   ... lots of inserts ...
'   db.CommitTrans  ' or db.RollbackTrans
'
Option Explicit

Private m_hDb       As Long
Private m_path      As String

' --- lifecycle ----------------------------------------------------

Public Property Get IsOpen() As Boolean
    IsOpen = (m_hDb <> 0)
End Property

Public Property Get Handle() As Long
    Handle = m_hDb
End Property

Public Property Get Path() As String
    Path = m_path
End Property

Public Property Get Version() As String
    Version = api.sqlite3_libversion()
End Property

Public Sub OpenDB(ByVal sFilename As String)
    If m_hDb <> 0 Then Err.Raise vbObjectError + 30100, "cSQLite.OpenDB", _
        "already open (path=" & m_path & ")"
    Dim rc As Long
    rc = api.sqlite3_open(sFilename, m_hDb)
    If rc <> SQLITE_OK Then
        Dim sErr As String
        If m_hDb <> 0 Then sErr = api.sqlite3_errmsg(m_hDb)
        If m_hDb <> 0 Then api.sqlite3_close m_hDb
        m_hDb = 0
        Err.Raise vbObjectError + 30000 + rc, "cSQLite.OpenDB", _
            "open failed (rc=" & rc & "): " & sErr
    End If
    m_path = sFilename

    ' Sensible defaults: a 5-second busy timeout makes shared-DB scenarios
    ' much more forgiving without changing semantics.
    api.sqlite3_busy_timeout m_hDb, 5000
End Sub

Public Sub CloseDB()
    If m_hDb <> 0 Then
         api.sqlite3_close m_hDb
        m_hDb = 0
        m_path = ""
    End If
End Sub

Private Sub Class_Terminate()
    CloseDB
End Sub

' --- error info ---------------------------------------------------

Public Property Get LastErrorCode() As Long
    If m_hDb <> 0 Then LastErrorCode = api.sqlite3_errcode(m_hDb)
End Property

Public Property Get LastErrorMessage() As String
    If m_hDb <> 0 Then LastErrorMessage = api.sqlite3_errmsg(m_hDb)
End Property

' --- one-shot SQL exec --------------------------------------------

' Run SQL with no parameters and no results. Raises on error.
' For SELECTs use Query() instead.
Public Sub Execute(ByVal sSql As String)
    EnsureOpen
    Dim sErr As String, rc As Long
    rc = api.sqlite3_exec(m_hDb, sSql, sErr)
    If rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 30000 + rc, "cSQLite.Execute", _
            "exec failed (rc=" & rc & "): " & sErr & vbCrLf & "SQL: " & sSql
    End If
End Sub

' --- prepared-statement factories ---------------------------------

Public Function Prepare(ByVal sSql As String) As cSQLiteStatement
    EnsureOpen
    Dim hStmt As Long, rc As Long
    rc = api.sqlite3_prepare_v2(m_hDb, sSql, hStmt)
    If rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 30000 + rc, "cSQLite.Prepare", _
            "prepare failed (rc=" & rc & "): " & api.sqlite3_errmsg(m_hDb) & _
            vbCrLf & "SQL: " & sSql
    End If
    Set Prepare = New cSQLiteStatement
    Prepare.AttachHandle m_hDb, hStmt, sSql
End Function

' Run a SELECT with optional bound parameters; return iterator-style results.
'   Set rs = db.Query("SELECT * FROM t WHERE age > ?", 30)
'   Do While rs.MoveNext()
'       ... rs("col") ...
'   Loop
Public Function Query(ByVal sSql As String, ParamArray params()) As cSQLiteResults
    EnsureOpen
    Dim stmt As cSQLiteStatement
    Set stmt = Prepare(sSql)

    ' Bind any positional parameters
    Dim p As Variant, i As Long
    i = 1
    For Each p In params
        stmt.Bind i, p
        i = i + 1
    Next

    Set Query = New cSQLiteResults
    Query.AttachStatement stmt
End Function

' --- ADO-flavored insert builder using prepared statements ---------
'
' Builds "INSERT INTO {table} ({fields}) VALUES (?,?,?,...)" with one
' placeholder per value, then binds and executes. Returns last_insert_rowid.
'
' Type-correct for Strings, Longs, Doubles, Currency, Dates, Byte arrays
' (blobs), and Null/Empty (-> NULL). No SQL injection risk.
'
' Usage:
'   db.ExecInsert "users", "name,age,joined", "O'Brien", 42, Now
'   db.ExecInsert "files", "name,data", "test.bin", binaryByteArray
Public Function ExecInsert(ByVal tblName As String, ByVal fields As String, _
                           ParamArray params()) As Currency
    Dim values() As Variant
    values = params  ' force a real array we can pass around
    ExecInsert = ExecInsertArr(tblName, fields, values)
End Function

' Same as ExecInsert but takes the values as an actual array (for
' programmatic callers that already have a Variant array). This is
' the one to use when you're forwarding a ParamArray from your own
' helper function.
Public Function ExecInsertArr(ByVal tblName As String, ByVal fields As String, _
                              ByRef values() As Variant) As Currency
    EnsureOpen
    Dim n As Long
    n = UBound(values) - LBound(values) + 1

    Dim placeholders As String, i As Long
    For i = 1 To n
        If i > 1 Then placeholders = placeholders & ","
        placeholders = placeholders & "?"
    Next i

    Dim sSql As String
    sSql = "INSERT INTO " & tblName & " (" & fields & ") VALUES (" & placeholders & ")"

    Dim stmt As cSQLiteStatement
    Set stmt = Prepare(sSql)
    For i = 0 To n - 1
        stmt.Bind i + 1, values(LBound(values) + i)
    Next i
    stmt.Execute

    ExecInsertArr = api.sqlite3_last_insert_rowid(m_hDb)
End Function

' Builds "UPDATE {table} SET f1=?,f2=?,... {criteria}" with placeholders
' and binds the values. {criteria} is the WHERE clause as you'd write
' it (e.g. "WHERE id=5") — interpolate any criteria values yourself,
' or use a separate parameterized query if untrusted.
'
' Returns rows affected.
'
' Usage:
'   db.ExecUpdate "users", "WHERE id=5", "name,age", "Jones", 31
' Builds "UPDATE {table} SET f1=?,f2=?,... {criteria}" with placeholders
' for the SET values. The {criteria} string can also contain '?'
' placeholders (e.g. "WHERE id=?"); supply values for those AFTER the
' SET values in the params list. Values are bound left-to-right matching
' SQLite's positional '?' numbering.
'
' Returns rows affected.
'
' Usage:
'   ' Simple: literal WHERE clause
'   db.ExecUpdate "users", "WHERE name='Alice'", "age", 31
'
'   ' Parameterized WHERE (safe with untrusted values):
'   db.ExecUpdate "users", "WHERE id=?", "age,notes", 31, "updated", 5
'   '                                                  ^^^ SET values  ^ criteria value
Public Function ExecUpdate(ByVal tblName As String, ByVal criteria As String, _
                           ByVal fields As String, _
                           ParamArray params()) As Long
    Dim values() As Variant
    values = params
    ExecUpdate = ExecUpdateArr(tblName, criteria, fields, values)
End Function

Public Function ExecUpdateArr(ByVal tblName As String, ByVal criteria As String, _
                              ByVal fields As String, _
                              ByRef values() As Variant) As Long
    EnsureOpen

    Dim fieldList() As String
    fieldList = Split(fields, ",")
    Dim nFields As Long
    nFields = UBound(fieldList) - LBound(fieldList) + 1

    Dim nValues As Long
    nValues = UBound(values) - LBound(values) + 1

    ' Count '?' placeholders in the criteria so we know how many extra
    ' values are expected after the SET values.
    Dim nCritPlaceholders As Long
    nCritPlaceholders = CountPlaceholders(criteria)

    Dim expectedValues As Long
    expectedValues = nFields + nCritPlaceholders

    If expectedValues <> nValues Then
        Err.Raise vbObjectError + 30101, "cSQLite.ExecUpdate", _
            "value count mismatch: expected " & expectedValues & _
            " (" & nFields & " for SET + " & nCritPlaceholders & _
            " for criteria), got " & nValues
    End If

    Dim setClause As String, i As Long
    For i = 0 To nFields - 1
        If i > 0 Then setClause = setClause & ","
        setClause = setClause & Trim$(fieldList(LBound(fieldList) + i)) & "=?"
    Next i

    Dim sSql As String
    sSql = "UPDATE " & tblName & " SET " & setClause
    If Len(criteria) > 0 Then sSql = sSql & " " & criteria

    Dim stmt As cSQLiteStatement
    Set stmt = Prepare(sSql)
    For i = 0 To nValues - 1
        stmt.Bind i + 1, values(LBound(values) + i)
    Next i
    stmt.Execute

    ExecUpdateArr = api.sqlite3_changes(m_hDb)
End Function

' Count unquoted '?' characters in an SQL fragment. Skips '?' inside
' string literals (single or double quoted). Good enough for typical
' UPDATE WHERE clauses; not a full SQL parser.
Private Function CountPlaceholders(ByVal s As String) As Long
    Dim i As Long, ch As String, n As Long
    Dim inSingle As Boolean, inDouble As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If inSingle Then
            If ch = "'" Then inSingle = False
        ElseIf inDouble Then
            If ch = """" Then inDouble = False
        Else
            Select Case ch
                Case "'": inSingle = True
                Case """": inDouble = True
                Case "?": n = n + 1
            End Select
        End If
    Next i
    CountPlaceholders = n
End Function

' --- transactions -------------------------------------------------

Public Sub BeginTrans()
    Execute "BEGIN TRANSACTION"
End Sub

Public Sub CommitTrans()
    Execute "COMMIT"
End Sub

Public Sub RollbackTrans()
    Execute "ROLLBACK"
End Sub

Public Property Get InTransaction() As Boolean
    If m_hDb <> 0 Then InTransaction = (api.sqlite3_get_autocommit(m_hDb) = 0)
End Property

' --- info & utility -----------------------------------------------

Public Property Get LastInsertRowID() As Currency
    If m_hDb <> 0 Then LastInsertRowID = api.sqlite3_last_insert_rowid(m_hDb)
End Property

Public Property Get RowsAffected() As Long
    If m_hDb <> 0 Then RowsAffected = api.sqlite3_changes(m_hDb)
End Property

' Run a SELECT that returns one scalar value. Returns Empty if no rows.
'   ageOfBob = db.Scalar("SELECT age FROM users WHERE name=?", "Bob")
Public Function Scalar(ByVal sSql As String, ParamArray params()) As Variant
    EnsureOpen
    Dim stmt As cSQLiteStatement
    Set stmt = Prepare(sSql)

    Dim p As Variant, i As Long
    i = 1
    For Each p In params
        stmt.Bind i, p
        i = i + 1
    Next

    If stmt.Step() Then
        Scalar = stmt.ColumnValue(0)
    Else
        Scalar = Empty
    End If
End Function

' --- internals ----------------------------------------------------

Private Sub EnsureOpen()
    If m_hDb = 0 Then
        Err.Raise vbObjectError + 30102, "cSQLite", "database not open"
    End If
End Sub
