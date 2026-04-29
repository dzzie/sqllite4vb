VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com
'
' cSQLiteStatement — prepared statement wrapper.
'
' You normally don't construct these directly. Get one via:
'   Set stmt = db.Prepare("INSERT INTO t (a,b) VALUES (?,?)")
'
' Then bind and step:
'   stmt.Bind 1, "hello"
'   stmt.Bind 2, 42
'   stmt.Execute        ' for INSERT/UPDATE/DELETE
'
' Or for SELECTs:
'   Do While stmt.Step()
'       Debug.Print stmt.ColumnText(0), stmt.ColumnInt(1)
'   Loop
'
' Reuse:
'   stmt.Reset          ' clear position, keep bindings
'   stmt.ClearBindings  ' clear bindings too
'
' The statement is finalized automatically when this object is destroyed.
'
Option Explicit

Private m_hDb       As Long
Private m_hStmt     As Long
Private m_sql       As String

' Cache column count for the lifetime of the statement.
Private m_nCols     As Long
Private m_colsCached As Boolean

' --- attach (called by cSQLite.Prepare) ---------------------------

Friend Sub AttachHandle(ByVal hDb As Long, ByVal hStmt As Long, ByVal sSql As String)
    m_hDb = hDb
    m_hStmt = hStmt
    m_sql = sSql
End Sub

Public Property Get Handle() As Long
    Handle = m_hStmt
End Property

Public Property Get SQL() As String
    SQL = m_sql
End Property

Public Property Get ParameterCount() As Long
    If m_hStmt <> 0 Then ParameterCount = api.sqlite3_bind_parameter_count(m_hStmt)
End Property

' --- binding ------------------------------------------------------

' Type-dispatching bind. Routes to the right sqlite3_bind_* based on
' the Variant's runtime type.
Public Sub Bind(ByVal idx As Long, ByRef value As Variant)
    EnsureValid
    Dim rc As Long
    Dim vt As Integer
    vt = VarType(value)

    Select Case vt
        Case vbNull, vbEmpty
            rc = api.sqlite3_bind_null(m_hStmt, idx)
        Case vbString
            rc = api.sqlite3_bind_text(m_hStmt, idx, CStr(value))
        Case vbLong, vbInteger, vbByte, vbBoolean
            rc = api.sqlite3_bind_int(m_hStmt, idx, CLng(value))
        Case vbCurrency
            rc = api.sqlite3_bind_int64(m_hStmt, idx, CCur(value))
        Case vbSingle, vbDouble
            rc = api.sqlite3_bind_double(m_hStmt, idx, CDbl(value))
        Case vbDate
            ' Store dates as ISO 8601 text — readable in SQLite browsers,
            ' sortable, and round-trips to Date via CDate().
            rc = api.sqlite3_bind_text(m_hStmt, idx, _
                Format$(value, "yyyy-mm-dd hh:nn:ss"))
        Case vbDecimal
            rc = api.sqlite3_bind_double(m_hStmt, idx, CDbl(value))
        Case Else
            ' Byte array -> blob
            If IsArray(value) Then
                If vt = (vbArray Or vbByte) Then
                    Dim bytes() As Byte
                    bytes = value
                    rc = api.sqlite3_bind_blob(m_hStmt, idx, bytes)
                Else
                    Err.Raise vbObjectError + 30200, "cSQLiteStatement.Bind", _
                        "unsupported array type for bind (only Byte() supported)"
                End If
            Else
                Err.Raise vbObjectError + 30201, "cSQLiteStatement.Bind", _
                    "unsupported VarType " & vt & " for bind"
            End If
    End Select

    If rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 30000 + rc, "cSQLiteStatement.Bind", _
            "bind failed (idx=" & idx & " rc=" & rc & "): " & api.sqlite3_errmsg(m_hDb)
    End If
End Sub

' Bind by named parameter (e.g. for ":name" or "@id" style placeholders)
Public Sub BindNamed(ByVal sName As String, ByRef value As Variant)
    EnsureValid
    Dim idx As Long
    idx = api.sqlite3_bind_parameter_index(m_hStmt, sName)
    If idx = 0 Then
        Err.Raise vbObjectError + 30202, "cSQLiteStatement.BindNamed", _
            "no parameter named '" & sName & "' in statement"
    End If
    Bind idx, value
End Sub

Public Sub Reset()
    If m_hStmt <> 0 Then api.sqlite3_reset m_hStmt
End Sub

Public Sub ClearBindings()
    If m_hStmt <> 0 Then api.sqlite3_clear_bindings m_hStmt
End Sub

' --- execution ----------------------------------------------------

' Step the statement. Returns True if a row is available, False on
' SQLITE_DONE. Raises on error.
Public Function Step() As Boolean
    EnsureValid
    Dim rc As Long
    rc = api.sqlite3_step(m_hStmt)
    Select Case rc
        Case SQLITE_ROW
            Step = True
        Case SQLITE_DONE
            Step = False
        Case Else
            Err.Raise vbObjectError + 30000 + rc, "cSQLiteStatement.Step", _
                "step failed (rc=" & rc & "): " & api.sqlite3_errmsg(m_hDb) & _
                vbCrLf & "SQL: " & m_sql
    End Select
End Function

' Execute a non-query (INSERT/UPDATE/DELETE/DDL). Steps to completion
' and raises if a row comes back unexpectedly.
Public Sub Execute()
    EnsureValid
    Dim rc As Long
    rc = api.sqlite3_step(m_hStmt)
    Select Case rc
        Case SQLITE_DONE
            ' good
        Case SQLITE_ROW
            Err.Raise vbObjectError + 30203, "cSQLiteStatement.Execute", _
                "Execute called on a statement that returns rows; use Step instead"
        Case Else
            Err.Raise vbObjectError + 30000 + rc, "cSQLiteStatement.Execute", _
                "execute failed (rc=" & rc & "): " & api.sqlite3_errmsg(m_hDb) & _
                vbCrLf & "SQL: " & m_sql
    End Select
    api.sqlite3_reset m_hStmt
End Sub

' --- column metadata ----------------------------------------------

Public Property Get ColumnCount() As Long
    If m_hStmt = 0 Then Exit Property
    If Not m_colsCached Then
        m_nCols = api.sqlite3_column_count(m_hStmt)
        m_colsCached = True
    End If
    ColumnCount = m_nCols
End Property

Public Function ColumnName(ByVal iCol As Long) As String
    EnsureValid
    ColumnName = api.sqlite3_column_name(m_hStmt, iCol)
End Function

Public Function ColumnDeclType(ByVal iCol As Long) As String
    EnsureValid
    ColumnDeclType = api.sqlite3_column_decltype(m_hStmt, iCol)
End Function

' Returns one of SQLITE_INTEGER / FLOAT / TEXT / BLOB / NULL
Public Function ColumnType(ByVal iCol As Long) As Long
    EnsureValid
    ColumnType = api.sqlite3_column_type(m_hStmt, iCol)
End Function

' --- column value access ------------------------------------------

Public Function ColumnInt(ByVal iCol As Long) As Long
    EnsureValid
    ColumnInt = api.sqlite3_column_int(m_hStmt, iCol)
End Function

Public Function ColumnInt64(ByVal iCol As Long) As Currency
    EnsureValid
    ColumnInt64 = api.sqlite3_column_int64(m_hStmt, iCol)
End Function

Public Function ColumnDouble(ByVal iCol As Long) As Double
    EnsureValid
    ColumnDouble = api.sqlite3_column_double(m_hStmt, iCol)
End Function

Public Function ColumnText(ByVal iCol As Long) As String
    EnsureValid
    ColumnText = api.sqlite3_column_text(m_hStmt, iCol)
End Function

Public Function ColumnBlob(ByVal iCol As Long) As Byte()
    EnsureValid
    ColumnBlob = api.sqlite3_column_blob(m_hStmt, iCol)
End Function

' Type-dispatched value: returns a Variant of the natural VB6 type
' for the column's actual SQLite storage class.
'
' For SQLITE_INTEGER we return Long when the value fits (the common
' case — rowids, ages, counts, flags). For values outside Long range
' we promote to Decimal via the Currency-as-int64 trick, multiplying
' the scaled Currency back up to its real integer value. This keeps
' the "rs("age") returns 30, not 0.003" expectation intact while still
' supporting genuine 64-bit values losslessly.
'
' If you specifically need int64 semantics (e.g. for a rowid that you
' know might overflow Long), call ColumnInt64 directly — that returns
' Currency without the magnitude check.
Public Function ColumnValue(ByVal iCol As Long) As Variant
    EnsureValid
    Select Case api.sqlite3_column_type(m_hStmt, iCol)
        Case SQLITE_INTEGER
            Dim cv As Currency
            cv = api.sqlite3_column_int64(m_hStmt, iCol)
            ' Currency stores int64/10000. Multiply back by 10000 (via
            ' Decimal to avoid overflow) to get the real integer value.
            Dim realVal As Variant
            realVal = CDec(cv) * CDec(10000)
            ' If it fits in Long, return Long (clean display, no scale).
            If realVal >= -2147483648# And realVal <= 2147483647# Then
                ColumnValue = CLng(realVal)
            Else
                ' Big value — return as Decimal Variant.
                ColumnValue = realVal
            End If
        Case SQLITE_FLOAT
            ColumnValue = api.sqlite3_column_double(m_hStmt, iCol)
        Case SQLITE_TEXT
            ColumnValue = api.sqlite3_column_text(m_hStmt, iCol)
        Case SQLITE_BLOB
            Dim b() As Byte
            b = api.sqlite3_column_blob(m_hStmt, iCol)
            ColumnValue = b
        Case SQLITE_NULL
            ColumnValue = Null
    End Select
End Function

' --- finalize -----------------------------------------------------

Public Sub Finalize()
    If m_hStmt <> 0 Then
        api.sqlite3_finalize m_hStmt
        m_hStmt = 0
    End If
End Sub

Private Sub Class_Terminate()
    Finalize
End Sub

Private Sub EnsureValid()
    If m_hStmt = 0 Then
        Err.Raise vbObjectError + 30204, "cSQLiteStatement", _
            "statement has been finalized or never attached"
    End If
End Sub
