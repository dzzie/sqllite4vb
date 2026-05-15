VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com
'
' cSQLiteTable - a static, detached, in-memory snapshot of a result set.
'
' Where cSQLiteResults is a live cursor (statement handle, forward-only,
' bound to a db connection), cSQLiteTable is just data: column names,
' column type hints, and a 2D Variant block of values. No statement, no
' connection, no iteration state. Safe to keep around, hand off, re-walk,
' or feed to an MSFlexGrid.
'
' Typical usage:
'
'   Dim rs As cSQLiteResults
'   Set rs = db.Query("SELECT name, age FROM users WHERE age > ?", 30)
'
'   Dim t As New cSQLiteTable
'   t.LoadFromResults rs        ' drains rs, closes it
'
'   Debug.Print t.RowCount, t.ColumnCount
'   Debug.Print t(0, "name"), t(0, 1)
'   Debug.Print t.Value(5, "age")
'
'   For r = 0 To t.RowCount - 1
'       For c = 0 To t.ColumnCount - 1
'           Debug.Print t(r, c),
'       Next
'       Debug.Print
'   Next
'
' One-liner from a db:
'   Set t = cSQLiteTable.Snapshot(db, "SELECT * FROM users")
'   ...but that's a static helper, so really:
'   Dim t As New cSQLiteTable: t.LoadQuery db, "SELECT * FROM users"
'
' Construct without a db (testing, manual data, flexgrid feed):
'   t.SetColumns Array("name", "age")
'   t.AddRow Array("alice", 30)
'   t.AddRow Array("bob",   45)
'
' Inspect the shape:
'   Dim f As cSQLiteField
'   For Each f In t.GetSchema()
'       Debug.Print f.Ordinal, f.Name, f.TypeName  ' "INTEGER", "TEXT", ...
'   Next
'
' Feed an MSFlexGrid:
'   t.FillFlexGrid MSFlexGrid1
'
' Generate SQL from current row data (the container stays dumb - caller
' supplies the table name and any WHERE clause):
'   sql = t.GenerateInsert("users", 0)
'   sql = t.GenerateUpdate("users", 0, "id = 42")
'   sql = t.GenerateDelete("users", "id = 42")
'
' Persist to JSON and load back (round-trips column names + types):
'   t.SaveJsonFile "C:\snapshot.json"
'   Dim t2 As New cSQLiteTable
'   t2.LoadJsonFile "C:\snapshot.json"
'
' JSON shape (version 1, columnar):
'   {
'     "version": 1,
'     "columns": [ {"name": "...", "type": <1..5>}, ... ],
'     "rows":    [ [v, v, v], [v, v, v], ... ]
'   }
'   BLOBs serialize as {"$blob": "<lowercase hex>"}; NULLs as JSON null.
'
Option Explicit

' Column metadata. Names are required; type hints are best-effort (carried
' from sqlite3_column_type at load time for the first row that produced a
' non-NULL value, else SQLITE_NULL = 5). We don't enforce types on writes.
Private m_colNames()      As String
Private m_colTypes()      As SQLiteType  ' sqlInteger=1..sqlNull=5
Private m_colCount        As Long
Private m_haveCols        As Boolean

' Row storage. We grow a Collection of 1D Variant arrays during AddRow,
' and lazily materialize a 2D Variant array (m_grid) on first read.
' This keeps AddRow O(1) amortized and reads O(1).
Private m_rows            As Collection   ' of Variant() - 0..colCount-1
Private m_grid()          As Variant      ' (row, col) once materialized
Private m_gridValid       As Boolean

' Lazy case-insensitive name -> index map
Private m_nameIndex       As Object       ' Scripting.Dictionary

Public Enum SQLiteType
    sqlInteger = 1
    sqlFloat = 2
    sqlText = 3
    sqlBlob = 4
    sqlNull = 5
End Enum

' Friendly name for logging/debug. Returns the SQLite-style uppercase
' string ("INTEGER", "FLOAT", "TEXT", "BLOB", "NULL") or "UNKNOWN".
Public Function SqlTypeName(ByVal t As SQLiteType) As String
    Select Case t
        Case sqlInteger: SqlTypeName = "INTEGER"
        Case sqlFloat:   SqlTypeName = "FLOAT"
        Case sqlText:    SqlTypeName = "TEXT"
        Case sqlBlob:    SqlTypeName = "BLOB"
        Case sqlNull:    SqlTypeName = "NULL"
        Case Else:       SqlTypeName = "UNKNOWN"
    End Select
End Function

Private Sub Class_Initialize()
    Set m_rows = New Collection
    m_haveCols = False
    m_gridValid = False
End Sub

' =================================================================
' Loaders
' =================================================================

' Drain a live cSQLiteResults into this table. Closes rs when done.
' Safe to call on an rs that hasn't had MoveNext called yet, or one
' that's been partially walked (we'll just take what's left).
Public Sub LoadFromResults(ByRef rs As cSQLiteResults)
    If rs Is Nothing Then
        Err.Raise vbObjectError + 30400, "cSQLiteTable.LoadFromResults", _
            "rs is Nothing"
    End If

    Reset

    Dim ncols As Long
    ncols = rs.FieldCount
    If ncols = 0 Then
        ' Statement with no result columns (e.g. a DML escapee). Leave empty.
        Exit Sub
    End If

    ' Capture column names up front. Type hints get filled in as we walk
    ' rows (first non-NULL wins per column).
    ReDim m_colNames(0 To ncols - 1)
    ReDim m_colTypes(0 To ncols - 1)
    Dim i As Long
    For i = 0 To ncols - 1
        m_colNames(i) = rs.FieldName(i)
        m_colTypes(i) = sqlNull   ' until we observe a real value
    Next i
    m_colCount = ncols
    m_haveCols = True

    Do While rs.MoveNext()
        Dim row() As Variant
        ReDim row(0 To ncols - 1)
        For i = 0 To ncols - 1
            row(i) = rs.Item(i)
            If m_colTypes(i) = sqlNull Then
                ' Only update the type hint if we haven't seen a real value yet.
                Dim t As Long
                t = rs.FieldType(i)
                If t <> sqlNull Then m_colTypes(i) = t
            End If
        Next i
        m_rows.Add row
    Loop

    rs.CloseRS
    m_gridValid = False
End Sub

' One-shot: run a query against a db and load the snapshot.
Public Sub LoadQuery(ByRef db As cSQLite, ByVal sql As String, ParamArray params() As Variant)
    If db Is Nothing Then
        Err.Raise vbObjectError + 30401, "cSQLiteTable.LoadQuery", _
            "db is Nothing"
    End If
    Dim rs As cSQLiteResults
    ' Forward ParamArray verbatim. cSQLite.Query is assumed to accept it the
    ' same way (the snippet in cSQLiteResults header shows db.Query taking
    ' positional params).
    Select Case UBound(params) - LBound(params) + 1
        Case 0:  Set rs = db.Query(sql)
        Case 1:  Set rs = db.Query(sql, params(LBound(params)))
        Case 2:  Set rs = db.Query(sql, params(LBound(params)), params(LBound(params) + 1))
        Case 3:  Set rs = db.Query(sql, params(LBound(params)), params(LBound(params) + 1), params(LBound(params) + 2))
        Case 4:  Set rs = db.Query(sql, params(LBound(params)), params(LBound(params) + 1), params(LBound(params) + 2), params(LBound(params) + 3))
        Case Else
            ' If you need more than 4 params, build the rs yourself and call LoadFromResults.
            Err.Raise vbObjectError + 30402, "cSQLiteTable.LoadQuery", _
                "too many parameters for ParamArray bridge - call LoadFromResults instead"
    End Select
    LoadFromResults rs
End Sub

' =================================================================
' Manual construction (no db required)
' =================================================================

' Define the columns. Pass either a 1D string array or a Variant array of
' names. Must be called before any AddRow. Calling again on a populated
' table is an error (Reset first).
Public Sub SetColumns(ByRef names As Variant)
    If m_rows.Count > 0 Then
        Err.Raise vbObjectError + 30403, "cSQLiteTable.SetColumns", _
            "cannot redefine columns on a populated table - call Reset first"
    End If

    Dim lo As Long, hi As Long, n As Long, i As Long
    If IsArray(names) Then
        lo = LBound(names)
        hi = UBound(names)
    Else
        Err.Raise vbObjectError + 30404, "cSQLiteTable.SetColumns", _
            "names must be an array"
    End If

    n = hi - lo + 1
    If n <= 0 Then
        Err.Raise vbObjectError + 30405, "cSQLiteTable.SetColumns", _
            "names is empty"
    End If

    ReDim m_colNames(0 To n - 1)
    ReDim m_colTypes(0 To n - 1)
    For i = 0 To n - 1
        m_colNames(i) = CStr(names(lo + i))
        m_colTypes(i) = sqlNull   ' unknown until populated
    Next i
    m_colCount = n
    m_haveCols = True
    Set m_nameIndex = Nothing
    m_gridValid = False
End Sub

' Append a row. Length must equal ColumnCount.
Public Sub AddRow(ByRef values As Variant)
    If Not m_haveCols Then
        Err.Raise vbObjectError + 30406, "cSQLiteTable.AddRow", _
            "no columns defined - call SetColumns first"
    End If
    If Not IsArray(values) Then
        Err.Raise vbObjectError + 30407, "cSQLiteTable.AddRow", _
            "values must be an array"
    End If
    Dim lo As Long, hi As Long, n As Long
    lo = LBound(values)
    hi = UBound(values)
    n = hi - lo + 1
    If n <> m_colCount Then
        Err.Raise vbObjectError + 30408, "cSQLiteTable.AddRow", _
            "row width " & n & " doesn't match column count " & m_colCount
    End If

    Dim row() As Variant
    ReDim row(0 To m_colCount - 1)
    Dim i As Long
    For i = 0 To m_colCount - 1
        row(i) = values(lo + i)
    Next i
    m_rows.Add row
    m_gridValid = False
End Sub

' Wipe everything. After this the table is in the same state as a fresh New.
Public Sub Reset()
    Set m_rows = New Collection
    Erase m_colNames
    Erase m_colTypes
    Erase m_grid
    m_colCount = 0
    m_haveCols = False
    m_gridValid = False
    Set m_nameIndex = Nothing
End Sub

' =================================================================
' Shape
' =================================================================

Public Property Get RowCount() As Long
    RowCount = m_rows.Count
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = m_colCount
End Property

Public Property Get ColumnName(ByVal iCol As Long) As String
    CheckCol iCol
    ColumnName = m_colNames(iCol)
End Property

' Return all column names as a 1D string array (handy for headers).
Public Function ColumnNames() As String()
    Dim names() As String
    If m_colCount = 0 Then
        ReDim names(0)
        ColumnNames = names
        Exit Function
    End If
    ReDim names(0 To m_colCount - 1)
    Dim i As Long
    For i = 0 To m_colCount - 1
        names(i) = m_colNames(i)
    Next i
    ColumnNames = names
End Function

' SQLite column type as last observed at load time. See SQLiteType enum:
'   sqlInteger=1, sqlFloat=2, sqlText=3, sqlBlob=4, sqlNull=5.
' This is a hint, not a constraint - Variant cells hold whatever you put in.
Public Property Get ColumnType(ByVal key As Variant) As SQLiteType
    ColumnType = m_colTypes(ResolveColIndex(key))
End Property

Public Property Get ColumnIndex(ByVal name As String) As Long
    ColumnIndex = ResolveColIndex(name)
End Property

' Return the full schema as a Collection of cSQLiteField, one per column,
' in column order. Walk it with For Each:
'
'   Dim f As cSQLiteField
'   For Each f In t.GetSchema()
'       Debug.Print f.Ordinal, f.Name, f.TypeName
'   Next
'
' Returns an empty Collection if no columns are defined.
Public Function GetSchema() As Collection
    Dim out As New Collection
    Dim i As Long
    For i = 0 To m_colCount - 1
        Dim f As cSQLiteField
        Set f = New cSQLiteField
        f.SetFrom i, m_colNames(i), m_colTypes(i)
        out.Add f
    Next i
    Set GetSchema = out
End Function

' =================================================================
' Cell access
' =================================================================

' Default property: t(row, col_key) where col_key is a Long index or
' String name. Returns Variant.
Public Property Get Item(ByVal iRow As Long, ByVal colKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    CheckRow iRow
    Dim c As Long
    c = ResolveColIndex(colKey)
    EnsureGrid
    Item = m_grid(iRow, c)
End Property

' Explicit alias for code clarity. Same as Item.
Public Property Get Value(ByVal iRow As Long, ByVal colKey As Variant) As Variant
    CheckRow iRow
    Dim c As Long
    c = ResolveColIndex(colKey)
    EnsureGrid
    Value = m_grid(iRow, c)
End Property

' Return an entire row as a 1D Variant array (0-based by column).
Public Function row(ByVal iRow As Long) As Variant
    CheckRow iRow
    ' Hand back a copy so callers can't mutate our internals.
    Dim src() As Variant
    src = m_rows(iRow + 1)
    Dim out() As Variant
    ReDim out(0 To m_colCount - 1)
    Dim i As Long
    For i = 0 To m_colCount - 1
        out(i) = src(i)
    Next i
    row = out
End Function

' Return an entire column as a 1D Variant array (0-based by row).
Public Function Column(ByVal colKey As Variant) As Variant
    Dim c As Long
    c = ResolveColIndex(colKey)
    If m_rows.Count = 0 Then
        Column = Empty
        Exit Function
    End If
    EnsureGrid
    Dim out() As Variant
    ReDim out(0 To m_rows.Count - 1)
    Dim r As Long
    For r = 0 To m_rows.Count - 1
        out(r) = m_grid(r, c)
    Next r
    Column = out
End Function

' Hand back the underlying 2D grid (rowIdx, colIdx). This is a copy of the
' materialized array, not a reference - mutating it doesn't affect us.
' Returns Empty if no rows.
Public Function ToArray() As Variant
    If m_rows.Count = 0 Or m_colCount = 0 Then
        ToArray = Empty
        Exit Function
    End If
    EnsureGrid
    ' VB6 arrays are by-value in Variant assignment, so this is already a copy.
    ToArray = m_grid
End Function

' =================================================================
' Output helpers
' =================================================================

' Pump our contents into an MSFlexGrid. Sets Rows/Cols, writes headers
' in the fixed row, then fills cells with CStr(value).
' Pass the control as an Object so we don't take a hard reference to
' the MSFlexGrid type library at compile time.
Public Sub FillFlexGrid(ByRef grid As Object, Optional ByVal includeHeader As Boolean = True)
    If grid Is Nothing Then Exit Sub

    Dim nRows As Long, ncols As Long
    nRows = m_rows.Count
    ncols = m_colCount

    grid.cols = ncols
    If includeHeader Then
        grid.FixedRows = 1
        grid.Rows = nRows + 1
    Else
        grid.FixedRows = 0
        grid.Rows = IIf(nRows = 0, 1, nRows)   ' MSFlexGrid wants Rows >= 1
    End If

    Dim r As Long, c As Long, rowOffset As Long
    rowOffset = IIf(includeHeader, 1, 0)

    If includeHeader Then
        For c = 0 To ncols - 1
            grid.TextMatrix(0, c) = m_colNames(c)
        Next c
    End If

    If nRows = 0 Then Exit Sub
    EnsureGrid
    For r = 0 To nRows - 1
        For c = 0 To ncols - 1
            grid.TextMatrix(r + rowOffset, c) = VariantToCell(m_grid(r, c))
        Next c
    Next r
End Sub

' Build a CSV string. Quotes fields containing delimiter, quote, CR, or LF.
' Doubles embedded quotes. NULLs render as empty fields.
Public Function ToCsv(Optional ByVal delim As String = ",", _
                      Optional ByVal includeHeader As Boolean = True, _
                      Optional ByVal lineSep As String = vbCrLf) As String
    Dim sb As String
    Dim r As Long, c As Long

    If includeHeader And m_colCount > 0 Then
        For c = 0 To m_colCount - 1
            If c > 0 Then sb = sb & delim
            sb = sb & CsvEscape(m_colNames(c), delim)
        Next c
        sb = sb & lineSep
    End If

    If m_rows.Count > 0 And m_colCount > 0 Then
        EnsureGrid
        For r = 0 To m_rows.Count - 1
            For c = 0 To m_colCount - 1
                If c > 0 Then sb = sb & delim
                sb = sb & CsvEscape(VariantToCell(m_grid(r, c)), delim)
            Next c
            sb = sb & lineSep
        Next r
    End If

    ToCsv = sb
End Function

' Save ToCsv() output to a UTF-8 file.
Public Sub SaveCsvFile(ByVal path As String, _
                        Optional ByVal delim As String = ",", _
                        Optional ByVal includeHeader As Boolean = True, _
                        Optional ByVal lineSep As String = vbCrLf)
    Dim s As String
    s = ToCsv(delim, includeHeader, lineSep)
    Dim bytes() As Byte
    bytes = StrConv(s, vbFromUnicode)
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Binary Access Write As #fnum
    If LenB(s) > 0 Then Put #fnum, , bytes
    Close #fnum
End Sub

' Replace contents from a CSV string. The parser handles:
'   - Quoted fields with embedded delimiters, quotes (doubled), and newlines
'   - Empty fields ("a,,b" -> three cells, middle is Null)
'   - Mixed quoted/unquoted fields on one row
'   - Trailing newline or its absence
'   - Mixed line endings (CRLF, LF, CR)
'
' Type handling:
'   - hasHeader=True: first row becomes column names
'   - hasHeader=False: columns are named "col1", "col2", ...
'   - typed=True (default): empty -> Null; integer-looking -> Long;
'                           number-with-decimal -> Double; else String
'   - typed=False: everything stays as String
'
' Column count is taken from the header row (or from the first data row
' if no header). Rows with too few cells get padded with Null; rows with
' too many raise.
Public Sub LoadCsv(ByVal csvText As String, _
                    Optional ByVal delim As String = ",", _
                    Optional ByVal hasHeader As Boolean = True, _
                    Optional ByVal typed As Boolean = True)
    Reset

    Dim fields As Collection
    Set fields = ParseCsv(csvText, delim)
    If fields.Count = 0 Then Exit Sub

    Dim firstRow As Variant
    firstRow = fields(1)

    Dim ncols As Long
    ncols = UBound(firstRow) - LBound(firstRow) + 1

    ' Set up columns
    ReDim m_colNames(0 To ncols - 1)
    ReDim m_colTypes(0 To ncols - 1)
    Dim i As Long
    If hasHeader Then
        For i = 0 To ncols - 1
            m_colNames(i) = CStr(firstRow(LBound(firstRow) + i))
            m_colTypes(i) = sqlNull
        Next i
    Else
        For i = 0 To ncols - 1
            m_colNames(i) = "col" & (i + 1)
            m_colTypes(i) = sqlNull
        Next i
    End If
    m_colCount = ncols
    m_haveCols = True

    ' Add rows
    Dim startIdx As Long
    If hasHeader Then startIdx = 2 Else startIdx = 1

    Dim r As Long
    For r = startIdx To fields.Count
        Dim rawRow As Variant
        rawRow = fields(r)
        Dim got As Long
        got = UBound(rawRow) - LBound(rawRow) + 1
        If got > ncols Then
            Err.Raise vbObjectError + 30440, "cSQLiteTable.LoadCsv", _
                "CSV row " & (r - 1) & " has " & got & " fields, expected " & ncols
        End If

        Dim outRow() As Variant
        ReDim outRow(0 To ncols - 1)
        For i = 0 To ncols - 1
            If i < got Then
                Dim raw As String
                raw = CStr(rawRow(LBound(rawRow) + i))
                If typed Then
                    outRow(i) = CoerceCsvCell(raw)
                Else
                    If Len(raw) = 0 Then
                        outRow(i) = ""    ' typed=False: empty stays empty string
                    Else
                        outRow(i) = raw
                    End If
                End If
            Else
                outRow(i) = Null    ' short row, pad with NULL
            End If
        Next i
        m_rows.Add outRow
    Next r

    m_gridValid = False
    Set m_nameIndex = Nothing

    ' Bake type hints from the actual loaded data so GetSchema reports
    ' real types instead of sqlNull. SniffColumnType walks rows looking
    ' for the first non-NULL value per column.
    Dim col As Long
    For col = 0 To m_colCount - 1
        m_colTypes(col) = SniffColumnType(col)
    Next col
End Sub

' Read a CSV file (UTF-8, BOM tolerated) and replace contents.
Public Sub LoadCsvFile(ByVal path As String, _
                        Optional ByVal delim As String = ",", _
                        Optional ByVal hasHeader As Boolean = True, _
                        Optional ByVal typed As Boolean = True)
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Binary Access Read As #fnum
    Dim n As Long
    n = LOF(fnum)
    Dim bytes() As Byte
    If n > 0 Then
        ReDim bytes(0 To n - 1)
        Get #fnum, , bytes
    End If
    Close #fnum

    Dim s As String
    s = StrConv(bytes, vbUnicode)

    ' Strip UTF-8 BOM if present
    If Len(s) >= 3 Then
        If Asc(Mid$(s, 1, 1)) = &HEF And Asc(Mid$(s, 2, 1)) = &HBB _
           And Asc(Mid$(s, 3, 1)) = &HBF Then
            s = Mid$(s, 4)
        End If
    End If

    LoadCsv s, delim, hasHeader, typed
End Sub

' Light type inference: empty -> Null, integer -> Long, decimal -> Double,
' else String. Deliberately does NOT try date parsing - locale and format
' ambiguity makes it more dangerous than useful.
Private Function CoerceCsvCell(ByVal raw As String) As Variant
    If Len(raw) = 0 Then
        CoerceCsvCell = Null
        Exit Function
    End If

    ' Probe for integer: optional leading -, then all digits.
    Dim start As Long
    start = 1
    If Left$(raw, 1) = "-" Then start = 2
    Dim allDigits As Boolean
    allDigits = (start <= Len(raw))   ' must have at least one digit
    Dim i As Long, ch As String
    For i = start To Len(raw)
        ch = Mid$(raw, i, 1)
        If ch < "0" Or ch > "9" Then
            allDigits = False
            Exit For
        End If
    Next i
    If allDigits Then
        ' Beware integer overflow - if too big, fall through to Double.
        On Error GoTo not_long
        CoerceCsvCell = CLng(raw)
        Exit Function
not_long:
        On Error GoTo 0
        ' Fall through to Double attempt
    End If

    ' Probe for double using strict grammar:
    '   [-]? digit+ ( . digit+ )? ( [eE] [+-]? digit+ )?
    ' This rejects strings like "2024-01-15" that contain digits and dashes
    ' but aren't actually numbers - the previous version used Val() which
    ' silently truncates at the first non-numeric character.
    If LooksLikeDouble(raw) Then
        CoerceCsvCell = CDbl(raw)
        Exit Function
    End If

    ' Fallback: string
    CoerceCsvCell = raw
End Function

' Strict number-grammar matcher. Returns True only for strings that fully
' match the floating-point grammar from end to end. No silent truncation,
' no locale-sensitive comma-as-decimal nonsense.
Private Function LooksLikeDouble(ByVal s As String) As Boolean
    Dim n As Long: n = Len(s)
    If n = 0 Then Exit Function

    Dim p As Long: p = 1
    Dim ch As String

    ' Optional sign
    ch = Mid$(s, p, 1)
    If ch = "+" Or ch = "-" Then
        p = p + 1
        If p > n Then Exit Function   ' sign alone isn't a number
    End If

    ' Integer part: one or more digits required (we allow ".5" via the
    ' branch below, so check that pattern explicitly)
    Dim sawIntDigit As Boolean
    Do While p <= n
        ch = Mid$(s, p, 1)
        If ch >= "0" And ch <= "9" Then
            sawIntDigit = True
            p = p + 1
        Else
            Exit Do
        End If
    Loop

    ' Optional fractional part: . followed by digits
    Dim sawFracDigit As Boolean
    If p <= n And Mid$(s, p, 1) = "." Then
        p = p + 1
        Do While p <= n
            ch = Mid$(s, p, 1)
            If ch >= "0" And ch <= "9" Then
                sawFracDigit = True
                p = p + 1
            Else
                Exit Do
            End If
        Loop
    End If

    ' Must have at least one digit somewhere in the mantissa
    If Not (sawIntDigit Or sawFracDigit) Then Exit Function

    ' Optional exponent: [eE] [+-]? digit+
    If p <= n Then
        ch = Mid$(s, p, 1)
        If ch = "e" Or ch = "E" Then
            p = p + 1
            If p > n Then Exit Function   ' exponent marker with no digits
            ch = Mid$(s, p, 1)
            If ch = "+" Or ch = "-" Then
                p = p + 1
                If p > n Then Exit Function
            End If
            Dim sawExpDigit As Boolean
            Do While p <= n
                ch = Mid$(s, p, 1)
                If ch >= "0" And ch <= "9" Then
                    sawExpDigit = True
                    p = p + 1
                Else
                    Exit Do
                End If
            Loop
            If Not sawExpDigit Then Exit Function
        End If
    End If

    ' If we consumed the entire string, it's a valid number
    LooksLikeDouble = (p > n)
End Function

' =================================================================
' SQL statement generators
' =================================================================
'
' These take a table name (and where applicable a WHERE clause) and use
' the table's row data as the value source. They generate SQL with values
' embedded as literals - apostrophes doubled, blobs as X'...', NULLs as
' NULL, numbers via Str$() for locale-safety.
'
' Identifiers (table name, column names) are always wrapped in double
' quotes, which SQLite accepts and which sidesteps reserved-word issues.
'
' These do NOT execute anything. They build a string. What you do with
' it is your problem.
'
' Usage:
'   Debug.Print t.GenerateInsert("users", 0)
'   ' INSERT INTO "users" ("name", "age") VALUES ('alice', 30);
'
'   Debug.Print t.GenerateUpdate("users", 0, "id = 42")
'   ' UPDATE "users" SET "name" = 'alice', "age" = 30 WHERE id = 42;
'
'   Debug.Print t.GenerateDelete("users", "id = 42")
'   ' DELETE FROM "users" WHERE id = 42;

Public Function GenerateInsert(ByVal tableName As String, ByVal iRow As Long) As String
    CheckRow iRow
    If m_colCount = 0 Then
        Err.Raise vbObjectError + 30430, "cSQLiteTable.GenerateInsert", _
            "no columns defined"
    End If
    EnsureGrid

    Dim cols As String, vals As String
    Dim c As Long
    For c = 0 To m_colCount - 1
        If c > 0 Then
            cols = cols & ", "
            vals = vals & ", "
        End If
        cols = cols & QuoteIdent(m_colNames(c))
        vals = vals & SqlLiteral(m_grid(iRow, c))
    Next c

    GenerateInsert = "INSERT INTO " & QuoteIdent(tableName) & _
                     " (" & cols & ") VALUES (" & vals & ");"
End Function

Public Function GenerateUpdate(ByVal tableName As String, ByVal iRow As Long, _
                                ByVal whereClause As String) As String
    CheckRow iRow
    If m_colCount = 0 Then
        Err.Raise vbObjectError + 30431, "cSQLiteTable.GenerateUpdate", _
            "no columns defined"
    End If
    If Len(Trim$(whereClause)) = 0 Then
        Err.Raise vbObjectError + 30432, "cSQLiteTable.GenerateUpdate", _
            "whereClause is required - refusing to generate unconditional UPDATE"
    End If
    EnsureGrid

    Dim sets As String
    Dim c As Long
    For c = 0 To m_colCount - 1
        If c > 0 Then sets = sets & ", "
        sets = sets & QuoteIdent(m_colNames(c)) & " = " & SqlLiteral(m_grid(iRow, c))
    Next c

    GenerateUpdate = "UPDATE " & QuoteIdent(tableName) & _
                     " SET " & sets & _
                     " WHERE " & whereClause & ";"
End Function

Public Function GenerateDelete(ByVal tableName As String, _
                                ByVal whereClause As String) As String
    If Len(Trim$(whereClause)) = 0 Then
        Err.Raise vbObjectError + 30433, "cSQLiteTable.GenerateDelete", _
            "whereClause is required - refusing to generate unconditional DELETE"
    End If
    GenerateDelete = "DELETE FROM " & QuoteIdent(tableName) & _
                     " WHERE " & whereClause & ";"
End Function

' Generate an INSERT for every row. Handy for "dump this table as a SQL
' script" - feeds straight into sqlite3 .read.
Public Function GenerateInsertAll(ByVal tableName As String, _
                                   Optional ByVal lineSep As String = vbCrLf) As String
    Dim sb As String
    Dim r As Long
    For r = 0 To m_rows.Count - 1
        sb = sb & GenerateInsert(tableName, r) & lineSep
    Next r
    GenerateInsertAll = sb
End Function

' Synthesize a CREATE TABLE statement from the table's own column names
' and SQLite type hints. This captures the SHAPE of the data (column
' names + types) but NOT constraints - we don't know about NOT NULL,
' PRIMARY KEY, DEFAULT, etc. because the container is just data, not
' a source-database mirror. By design: the table might have come from
' a query result, a JSON file, or AddRow calls with no db at all.
'
' For columns whose type hint is sqlNull (typically because the table
' was built by hand via AddRow), we scan the row data for the first
' non-NULL value and use its SQLite type. Columns that are NULL in
' every row fall back to BLOB affinity (SQLite's default for unknown
' types - it'll store anything).
Public Function GenerateCreateTable(ByVal tableName As String, _
                                     Optional ByVal IfNotExists As Boolean = False) As String
    If m_colCount = 0 Then
        Err.Raise vbObjectError + 30434, "cSQLiteTable.GenerateCreateTable", _
            "no columns defined"
    End If

    ' Resolve effective types: use the hint if known, otherwise sniff rows.
    Dim effTypes() As SQLiteType
    ReDim effTypes(0 To m_colCount - 1)
    Dim c As Long
    For c = 0 To m_colCount - 1
        If m_colTypes(c) = sqlNull Then
            effTypes(c) = SniffColumnType(c)
        Else
            effTypes(c) = m_colTypes(c)
        End If
    Next c

    Dim cols As String
    For c = 0 To m_colCount - 1
        If c > 0 Then cols = cols & ", "
        cols = cols & QuoteIdent(m_colNames(c)) & " " & SqlTypeName(effTypes(c))
    Next c
    Dim ifne As String
    If IfNotExists Then ifne = "IF NOT EXISTS "
    GenerateCreateTable = "CREATE TABLE " & ifne & QuoteIdent(tableName) & _
                          " (" & cols & ");"
End Function

' Scan column iCol's row data for the first non-NULL value and return
' its SQLite native type. Used by GenerateCreateTable when the stored
' hint is sqlNull. Falls back to sqlBlob if every row is NULL (BLOB
' has the loosest affinity in SQLite - it accepts any value).
Private Function SniffColumnType(ByVal iCol As Long) As SQLiteType
    If m_rows.Count = 0 Then
        SniffColumnType = sqlBlob
        Exit Function
    End If
    EnsureGrid
    Dim r As Long
    For r = 0 To m_rows.Count - 1
        Dim v As Variant
        v = m_grid(r, iCol)
        Select Case VarType(v)
            Case vbNull, vbEmpty
                ' keep looking
            Case vbInteger, vbLong, vbByte, vbBoolean
                SniffColumnType = sqlInteger: Exit Function
            Case vbSingle, vbDouble, vbCurrency, vbDecimal
                SniffColumnType = sqlFloat: Exit Function
            Case vbString, vbDate
                SniffColumnType = sqlText: Exit Function
            Case vbArray + vbByte
                SniffColumnType = sqlBlob: Exit Function
            Case Else
                SniffColumnType = sqlText: Exit Function
        End Select
    Next r
    SniffColumnType = sqlBlob   ' column is NULL in every row
End Function

' Full SQL serialization: CREATE TABLE followed by every row as an
' INSERT. The result is a self-contained SQL script you can feed to
' `sqlite3 newdb.sqlite` or `.read` into an existing connection.
'
' CREATE TABLE is synthesized from internal column names and type hints
' (with row sniffing for unknown types). Constraints (NOT NULL, PRIMARY
' KEY, defaults) are not reproduced - the container doesn't track them.
'
' Parallels ToJson and ToCsv.
Public Function ToSql(ByVal tableName As String, _
                       Optional ByVal lineSep As String = vbCrLf, _
                       Optional ByVal IfNotExists As Boolean = False, _
                       Optional ByVal DropFirst As Boolean = False) As String
    Dim sb As String
    If DropFirst Then
        sb = sb & "DROP TABLE IF EXISTS " & QuoteIdent(tableName) & ";" & lineSep
    End If
    sb = sb & GenerateCreateTable(tableName, IfNotExists) & lineSep
    sb = sb & GenerateInsertAll(tableName, lineSep)
    ToSql = sb
End Function

' Save ToSql() output to a file. Parallels SaveJsonFile.
Public Sub SaveSqlFile(ByVal path As String, ByVal tableName As String, _
                        Optional ByVal IfNotExists As Boolean = False, _
                        Optional ByVal DropFirst As Boolean = False)
    Dim s As String
    s = ToSql(tableName, vbCrLf, IfNotExists, DropFirst)
    Dim bytes() As Byte
    bytes = StrConv(s, vbFromUnicode)
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Binary Access Write As #fnum
    If LenB(s) > 0 Then Put #fnum, , bytes
    Close #fnum
End Sub

' --- SQL literal helpers -----------------------------------------

' Wrap an identifier in double quotes, doubling any embedded double quote.
Private Function QuoteIdent(ByVal s As String) As String
    QuoteIdent = """" & Replace$(s, """", """""") & """"
End Function

' Render a Variant cell as a SQL literal.
Private Function SqlLiteral(ByRef v As Variant) As String
    Select Case VarType(v)
        Case vbNull, vbEmpty
            SqlLiteral = "NULL"
        Case vbBoolean
            ' SQLite has no bool; 1/0 is the convention.
            SqlLiteral = IIf(CBool(v), "1", "0")
        Case vbInteger, vbLong, vbByte
            SqlLiteral = CStr(CLng(v))
        Case vbCurrency
            SqlLiteral = Trim$(Str$(CCur(v)))
        Case vbSingle, vbDouble, vbDecimal
            ' Str$ uses "." regardless of locale.
            SqlLiteral = Trim$(Str$(CDbl(v)))
        Case vbString
            SqlLiteral = SqlQuoteString(CStr(v))
        Case vbDate
            SqlLiteral = SqlQuoteString(Format$(v, "yyyy-mm-dd HH:Nn:Ss"))
        Case vbArray + vbByte
            ' BLOB as X'...' literal. Same coercion trick as JsonEncodeValue.
            Dim bytes() As Byte
            bytes = v
            SqlLiteral = "X'" & BytesToHex(bytes) & "'"
        Case Else
            ' Best-effort fallback.
            SqlLiteral = SqlQuoteString(CStr(v))
    End Select
End Function

' Wrap a string in single quotes, doubling any embedded single quote.
' This is the SQLite-standard string escape - no backslash escapes,
' no special control char handling. SQLite stores whatever bytes you
' give it.
Private Function SqlQuoteString(ByVal s As String) As String
    SqlQuoteString = "'" & Replace$(s, "'", "''") & "'"
End Function

' =================================================================
' JSON I/O (columnar shape, hex-encoded blobs, no external deps)
' =================================================================
'
' Format (version 1):
'   {
'     "version": 1,
'     "columns": [ {"name": "...", "type": <SQLite type 1..5>}, ... ],
'     "rows":    [ [v, v, ...], ... ]
'   }
'
' Cell values:
'   - integers  -> JSON number
'   - doubles   -> JSON number
'   - strings   -> JSON string (with full \uXXXX escaping for control chars)
'   - blobs     -> {"$blob": "<lowercase hex>"}
'   - NULL/Empty-> JSON null
'   - dates     -> ISO-ish string yyyy-mm-ddTHH:MM:SS (cell type stays TEXT)
'
' The reader accepts either:
'   - Top-level object with columns+rows (the format above), or
'   - Top-level array of objects (portable shape) - inferred, column order
'     comes from the first object's key order.

' Serialize the table to a JSON string.
' Pretty=False produces a single-line dense form (best for storage).
' Pretty=True indents with 2 spaces (best for diffing/eyeballing).
Public Function ToJson(Optional ByVal Pretty As Boolean = False) As String
    Dim sb As String
    Dim nl As String, ind1 As String, ind2 As String
    Dim kvSep As String

    If Pretty Then
        nl = vbCrLf
        ind1 = "  "
        ind2 = "    "
        kvSep = ": "
    Else
        nl = ""
        kvSep = ":"
    End If

    sb = "{" & nl
    sb = sb & ind1 & """version""" & kvSep & "1," & nl
    sb = sb & ind1 & """columns""" & kvSep & "["

    Dim i As Long
    If m_colCount > 0 Then
        sb = sb & nl
        For i = 0 To m_colCount - 1
            sb = sb & ind2 & "{""name""" & kvSep & JsonEscapeString(m_colNames(i)) & _
                 "," & IIf(Pretty, " ", "") & """type""" & kvSep & CStr(m_colTypes(i)) & "}"
            If i < m_colCount - 1 Then sb = sb & ","
            sb = sb & nl
        Next i
        sb = sb & ind1
    End If
    sb = sb & "]," & nl
    sb = sb & ind1 & """rows""" & kvSep & "["

    Dim nRows As Long
    nRows = m_rows.Count
    If nRows > 0 And m_colCount > 0 Then
        EnsureGrid
        sb = sb & nl
        Dim r As Long, c As Long
        For r = 0 To nRows - 1
            sb = sb & ind2 & "["
            For c = 0 To m_colCount - 1
                If c > 0 Then sb = sb & "," & IIf(Pretty, " ", "")
                sb = sb & JsonEncodeValue(m_grid(r, c))
            Next c
            sb = sb & "]"
            If r < nRows - 1 Then sb = sb & ","
            sb = sb & nl
        Next r
        sb = sb & ind1
    End If
    sb = sb & "]" & nl & "}"

    ToJson = sb
End Function

' Save ToJson(Pretty) to a UTF-8 file (no BOM).
Public Sub SaveJsonFile(ByVal path As String, Optional ByVal Pretty As Boolean = True)
    Dim s As String
    s = ToJson(Pretty)
    Dim bytes() As Byte
    bytes = StrConv(s, vbFromUnicode)   ' good enough if all chars are ASCII
                                        ' (JsonEscapeString emits \uXXXX for >7F)
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Binary Access Write As #fnum
    If LenB(s) > 0 Then Put #fnum, , bytes
    Close #fnum
End Sub

' Replace contents from a JSON string. Accepts either the columnar shape
' written by ToJson, or a top-level array of objects.
Public Sub LoadJson(ByVal jsonText As String)
    Dim parser As New CJsonParser
    ' Result-set JSON always has an object or array at the root, both of
    ' which our parser returns as Objects (Dictionary or Collection).
    ' Set is required when the RHS is a Variant holding an object reference.
    Dim root As Object
    Set root = parser.Parse(jsonText)

    Reset

    If TypeName(root) = "Dictionary" Then
        ' Columnar shape: object with "columns" and "rows".
        Dim d As Object
        Set d = root
        If Not d.Exists("columns") Then
            Err.Raise vbObjectError + 30420, "cSQLiteTable.LoadJson", _
                "object form requires 'columns' key"
        End If
        If Not d.Exists("rows") Then
            Err.Raise vbObjectError + 30421, "cSQLiteTable.LoadJson", _
                "object form requires 'rows' key"
        End If

        Dim colsArr As Collection
        Set colsArr = d("columns")
        Dim n As Long
        n = colsArr.Count
        If n = 0 Then Exit Sub

        ReDim m_colNames(0 To n - 1)
        ReDim m_colTypes(0 To n - 1)
        Dim i As Long
        For i = 0 To n - 1
            Dim colDef As Object
            Set colDef = colsArr(i + 1)
            m_colNames(i) = CStr(colDef("name"))
            If colDef.Exists("type") Then
                m_colTypes(i) = CLng(colDef("type"))
            Else
                m_colTypes(i) = sqlNull
            End If
        Next i
        m_colCount = n
        m_haveCols = True

        Dim rowsArr As Collection
        Set rowsArr = d("rows")
        Dim r As Long
        For r = 1 To rowsArr.Count
            Dim parsedRow As Collection
            Set parsedRow = rowsArr(r)
            If parsedRow.Count <> m_colCount Then
                Err.Raise vbObjectError + 30422, "cSQLiteTable.LoadJson", _
                    "row " & (r - 1) & " has " & parsedRow.Count & _
                    " values, expected " & m_colCount
            End If
            Dim row() As Variant
            ReDim row(0 To m_colCount - 1)
            For i = 0 To m_colCount - 1
                row(i) = DecodeJsonCell(parsedRow(i + 1))
            Next i
            m_rows.Add row
        Next r

    ElseIf TypeName(root) = "Collection" Then
        ' Portable shape: array of objects. Take column order from first row.
        Dim arr As Collection
        Set arr = root
        If arr.Count = 0 Then Exit Sub

        Dim first As Object
        Set first = arr(1)
        Dim keys As Variant
        keys = first.keys
        Dim k As Long
        Dim ncols As Long
        ncols = UBound(keys) - LBound(keys) + 1
        If ncols = 0 Then Exit Sub

        ReDim m_colNames(0 To ncols - 1)
        ReDim m_colTypes(0 To ncols - 1)
        For k = 0 To ncols - 1
            m_colNames(k) = CStr(keys(LBound(keys) + k))
            m_colTypes(k) = sqlNull
        Next k
        m_colCount = ncols
        m_haveCols = True

        For r = 1 To arr.Count
            Dim obj As Object
            Set obj = arr(r)
            Dim row2() As Variant
            ReDim row2(0 To ncols - 1)
            For k = 0 To ncols - 1
                If obj.Exists(m_colNames(k)) Then
                    row2(k) = DecodeJsonCell(obj(m_colNames(k)))
                Else
                    row2(k) = Null
                End If
            Next k
            m_rows.Add row2
        Next r
    Else
        Err.Raise vbObjectError + 30423, "cSQLiteTable.LoadJson", _
            "unexpected top-level JSON type"
    End If

    m_gridValid = False
    Set m_nameIndex = Nothing

    ' For the portable-shape branch the type hints came in as sqlNull
    ' (since portable JSON doesn't carry them). Sniff actual row data to
    ' populate them, so GetSchema returns useful answers.
    ' For the columnar branch the hints came from the JSON metadata
    ' already, but only update sqlNull slots so we don't clobber known
    ' types with sniffed approximations.
    Dim col As Long
    For col = 0 To m_colCount - 1
        If m_colTypes(col) = sqlNull Then
            m_colTypes(col) = SniffColumnType(col)
        End If
    Next col
End Sub

' Read a JSON file (UTF-8, BOM tolerated) and replace our contents.
Public Sub LoadJsonFile(ByVal path As String)
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Binary Access Read As #fnum
    Dim n As Long
    n = LOF(fnum)
    Dim bytes() As Byte
    If n > 0 Then
        ReDim bytes(0 To n - 1)
        Get #fnum, , bytes
    End If
    Close #fnum

    Dim s As String
    s = StrConv(bytes, vbUnicode)

    ' Strip UTF-8 BOM if present (EF BB BF -> three chars in this representation)
    If Len(s) >= 3 Then
        If Asc(Mid$(s, 1, 1)) = &HEF And Asc(Mid$(s, 2, 1)) = &HBB _
           And Asc(Mid$(s, 3, 1)) = &HBF Then
            s = Mid$(s, 4)
        End If
    End If

    LoadJson s
End Sub

' --- JSON writer helpers -----------------------------------------

Private Function JsonEncodeValue(ByRef v As Variant) As String
    Select Case VarType(v)
        Case vbNull, vbEmpty
            JsonEncodeValue = "null"
        Case vbBoolean
            JsonEncodeValue = IIf(CBool(v), "true", "false")
        Case vbInteger, vbLong, vbByte
            JsonEncodeValue = CStr(CLng(v))
        Case vbCurrency
            ' Currency is a 64-bit scaled int. Emit with up to 4 decimals.
            JsonEncodeValue = FormatCurrencyForJson(CCur(v))
        Case vbSingle, vbDouble, vbDecimal
            JsonEncodeValue = FormatDoubleForJson(CDbl(v))
        Case vbString
            JsonEncodeValue = JsonEscapeString(CStr(v))
        Case vbDate
            JsonEncodeValue = JsonEscapeString(Format$(v, "yyyy-mm-dd\THH:Nn:Ss"))
        Case vbArray + vbByte
            ' Extract into a typed Byte() local. VB6 won't let us pass a
            ' Variant directly to a ByRef Byte() parameter, even when the
            ' Variant holds a byte array - the compiler refuses, and the
            ' error surfaces (confusingly) on the Case line itself.
            Dim bytes() As Byte
            bytes = v
            JsonEncodeValue = "{""$blob"":" & JsonEscapeString(BytesToHex(bytes)) & "}"
        Case Else
            ' Fall back to string form
            JsonEncodeValue = JsonEscapeString(CStr(v))
    End Select
End Function

Private Function FormatDoubleForJson(ByVal d As Double) As String
    ' VB6 Str$ uses "." regardless of locale - perfect for JSON.
    Dim s As String
    s = Trim$(Str$(d))
    ' Str$ prefixes positive numbers with a space which Trim handles; it does
    ' NOT prefix anything for the integer/zero case beyond that. Good.
    FormatDoubleForJson = s
End Function

Private Function FormatCurrencyForJson(ByVal c As Currency) As String
    ' Avoid locale-dependent thousands separators / commas.
    Dim s As String
    s = Trim$(Str$(c))
    FormatCurrencyForJson = s
End Function

Private Function JsonEscapeString(ByVal s As String) As String
    Dim sb As String
    Dim i As Long, ch As String, code As Long
    sb = """"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536   ' AscW returns signed
        Select Case code
            Case 34:  sb = sb & "\"""
            Case 92:  sb = sb & "\\"
            Case 8:   sb = sb & "\b"
            Case 9:   sb = sb & "\t"
            Case 10:  sb = sb & "\n"
            Case 12:  sb = sb & "\f"
            Case 13:  sb = sb & "\r"
            Case Else
                If code < 32 Or code > 126 Then
                    sb = sb & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    sb = sb & ch
                End If
        End Select
    Next i
    sb = sb & """"
    JsonEscapeString = sb
End Function

Private Function BytesToHex(ByRef b() As Byte) As String
    On Error GoTo empty_blob
    Dim lo As Long, hi As Long, n As Long, i As Long
    lo = LBound(b)
    hi = UBound(b)
    n = hi - lo + 1
    If n <= 0 Then BytesToHex = "": Exit Function
    Dim out As String
    out = Space$(n * 2)
    Dim hx As String
    For i = 0 To n - 1
        hx = Hex$(b(lo + i))
        If Len(hx) = 1 Then hx = "0" & hx
        Mid$(out, i * 2 + 1, 2) = LCase$(hx)
    Next i
    BytesToHex = out
    Exit Function
empty_blob:
    BytesToHex = ""
End Function

Private Function HexToBytes(ByVal s As String) As Byte()
    Dim n As Long, i As Long
    n = Len(s) \ 2
    Dim out() As Byte
    If n = 0 Then
        ' Zero-length blob: leave the array uninitialized. UBound on the
        ' result will raise; consumers should bracket with On Error like
        ' BytesToHex does on the write side.
        HexToBytes = out
        Exit Function
    End If
    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        out(i) = CByte("&H" & Mid$(s, i * 2 + 1, 2))
    Next i
    HexToBytes = out
End Function

' Translate a single cell value coming back from the JSON parser into the
' Variant we want to store. The parser hands us: String, Double, Boolean,
' Null, Collection (array), or Dictionary (object). We only special-case
' {"$blob": "<hex>"}; anything else passes through.
Private Function DecodeJsonCell(ByRef v As Variant) As Variant
    If IsObject(v) Then
        Dim d As Object
        Set d = v
        If d.Exists("$blob") Then
            DecodeJsonCell = HexToBytes(CStr(d("$blob")))
            Exit Function
        End If
        ' Unknown object - stringify so we don't lose it silently.
        DecodeJsonCell = "[object]"
    Else
        DecodeJsonCell = v
    End If
End Function


Private Sub EnsureGrid()
    If m_gridValid Then Exit Sub
    Dim n As Long
    n = m_rows.Count
    If n = 0 Or m_colCount = 0 Then
        Erase m_grid
        m_gridValid = True
        Exit Sub
    End If
    ReDim m_grid(0 To n - 1, 0 To m_colCount - 1)
    Dim r As Long, c As Long
    Dim row() As Variant
    For r = 0 To n - 1
        row = m_rows(r + 1)
        For c = 0 To m_colCount - 1
            m_grid(r, c) = row(c)
        Next c
    Next r
    m_gridValid = True
End Sub

Private Sub CheckRow(ByVal iRow As Long)
    If iRow < 0 Or iRow >= m_rows.Count Then
        Err.Raise vbObjectError + 30410, "cSQLiteTable", _
            "row index " & iRow & " out of range (0.." & (m_rows.Count - 1) & ")"
    End If
End Sub

Private Sub CheckCol(ByVal iCol As Long)
    If iCol < 0 Or iCol >= m_colCount Then
        Err.Raise vbObjectError + 30411, "cSQLiteTable", _
            "column index " & iCol & " out of range (0.." & (m_colCount - 1) & ")"
    End If
End Sub

Private Function ResolveColIndex(ByRef key As Variant) As Long
    If VarType(key) = vbString Then
        If m_nameIndex Is Nothing Then BuildNameIndex
        Dim k As String
        k = LCase$(CStr(key))
        If Not m_nameIndex.Exists(k) Then
            Err.Raise vbObjectError + 30412, "cSQLiteTable", _
                "no column named '" & key & "'"
        End If
        ResolveColIndex = m_nameIndex(k)
    Else
        Dim i As Long
        i = CLng(key)
        CheckCol i
        ResolveColIndex = i
    End If
End Function

Private Sub BuildNameIndex()
    Set m_nameIndex = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To m_colCount - 1
        m_nameIndex(LCase$(m_colNames(i))) = i
    Next i
End Sub

' Best-effort Variant -> displayable string for flexgrid / csv.
Private Function VariantToCell(ByRef v As Variant) As String
    Select Case VarType(v)
        Case vbNull, vbEmpty
            VariantToCell = ""
        Case vbArray + vbByte
            ' BLOB - show length placeholder; consumer who cares can fetch raw.
            On Error Resume Next
            VariantToCell = "<blob " & (UBound(v) - LBound(v) + 1) & " bytes>"
            On Error GoTo 0
        Case Else
            VariantToCell = CStr(v)
    End Select
End Function

Private Function CsvEscape(ByVal s As String, ByVal delim As String) As String
    If InStr(s, delim) = 0 And InStr(s, """") = 0 _
       And InStr(s, vbCr) = 0 And InStr(s, vbLf) = 0 Then
        CsvEscape = s
    Else
        CsvEscape = """" & Replace$(s, """", """""") & """"
    End If
End Function




' modCsvParser - RFC-4180-ish CSV parser. Reads a CSV string and yields
' rows as 1D Variant arrays of raw field strings. Type coercion is the
' caller's job.
'
' Handles:
'   - Quoted fields with embedded delimiters, quotes (doubled), newlines
'   - Empty fields ("a,,b" -> three cells, middle empty)
'   - Mixed quoted/unquoted fields on one row
'   - Trailing newline or absence thereof
'   - Mixed line endings (CRLF, LF, CR)
'   - Single-char delimiters (",", ";", vbTab, "|", etc.)
'
' Usage:
'   Dim rows As Collection
'   Set rows = ParseCsv(csvText, ",")
'   ' rows is a Collection of Variant() arrays. Each array's elements
'   ' are the raw String fields for that row.
 
' Parse a full CSV document. Returns a Collection of 1D Variant arrays.
' Each array holds the raw String fields for one row. Field count can
' vary between rows - the caller decides what to do about that.
Private Function ParseCsv(ByVal csvText As String, _
                          Optional ByVal delim As String = ",") As Collection
    Dim out As New Collection
    Dim n As Long
    n = Len(csvText)
    If n = 0 Then
        Set ParseCsv = out
        Exit Function
    End If

    Dim delimCh As String
    delimCh = Left$(delim, 1)

    Dim pos As Long
    pos = 1

    Do While pos <= n
        Dim row As Collection
        Set row = New Collection
        Dim rowEnded As Boolean
        rowEnded = False

        Do While Not rowEnded
            Dim field As String
            field = ParseField(csvText, pos, n, delimCh, rowEnded)
            row.Add field
            If pos > n Then rowEnded = True
        Loop

        ' Convert Collection -> 1D Variant array
        Dim arr() As Variant
        ReDim arr(0 To row.Count - 1)
        Dim i As Long
        For i = 0 To row.Count - 1
            arr(i) = row(i + 1)
        Next i
        out.Add arr
    Loop

    Set ParseCsv = out
End Function

' --- internals ---------------------------------------------------

' Parse one field starting at pos. Advances pos past the field and
' its trailing delimiter or row terminator. Sets rowEnded=True if the
' terminator was a newline or EOF.
Private Function ParseField(ByRef s As String, ByRef pos As Long, _
                             ByVal n As Long, ByVal delimCh As String, _
                             ByRef rowEnded As Boolean) As String
    Dim sb As String
    Dim ch As String

    If pos > n Then
        rowEnded = True
        ParseField = ""
        Exit Function
    End If

    ch = Mid$(s, pos, 1)
    If ch = """" Then
        ' Quoted field. Consume until we find a closing quote not followed
        ' by another quote.
        pos = pos + 1
        Do While pos <= n
            ch = Mid$(s, pos, 1)
            If ch = """" Then
                If pos < n And Mid$(s, pos + 1, 1) = """" Then
                    ' Escaped doubled quote
                    sb = sb & """"
                    pos = pos + 2
                Else
                    ' End of quoted field
                    pos = pos + 1
                    Exit Do
                End If
            Else
                sb = sb & ch
                pos = pos + 1
            End If
        Loop
        ' After closing quote, expect delim, newline, or EOF.
        ConsumeTerminator s, pos, n, delimCh, rowEnded
    Else
        ' Unquoted field. Read until delim or newline.
        Do While pos <= n
            ch = Mid$(s, pos, 1)
            If ch = delimCh Then
                pos = pos + 1
                Exit Do
            ElseIf ch = vbCr Or ch = vbLf Then
                ConsumeNewline s, pos, n
                rowEnded = True
                Exit Do
            Else
                sb = sb & ch
                pos = pos + 1
            End If
        Loop
        If pos > n Then rowEnded = True
    End If

    ParseField = sb
End Function

' After a closing quote in a quoted field. The next char should be
' delim, CR/LF, or EOF.
Private Sub ConsumeTerminator(ByRef s As String, ByRef pos As Long, _
                               ByVal n As Long, ByVal delimCh As String, _
                               ByRef rowEnded As Boolean)
    If pos > n Then
        rowEnded = True
        Exit Sub
    End If
    Dim ch As String
    ch = Mid$(s, pos, 1)
    If ch = delimCh Then
        pos = pos + 1
    ElseIf ch = vbCr Or ch = vbLf Then
        ConsumeNewline s, pos, n
        rowEnded = True
    Else
        ' Strict CSV would error here. We tolerate: treat as end-of-field.
        rowEnded = True
    End If
End Sub

' Consume a newline, accepting CRLF, LF, or CR.
Private Sub ConsumeNewline(ByRef s As String, ByRef pos As Long, ByVal n As Long)
    Dim ch As String
    ch = Mid$(s, pos, 1)
    If ch = vbCr Then
        pos = pos + 1
        If pos <= n And Mid$(s, pos, 1) = vbLf Then pos = pos + 1
    ElseIf ch = vbLf Then
        pos = pos + 1
    End If
End Sub


