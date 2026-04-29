VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com
'
' cSQLiteResults — iterator-style result set, ADO Recordset flavor.
'
' You don't construct these directly. Get one via:
'   Set rs = db.Query("SELECT name, age FROM users WHERE age > ?", 30)
'
' Access pattern (preferred):
'   Do While rs.MoveNext()
'       Debug.Print rs("name"), rs("age")     ' default property: by-name
'       Debug.Print rs(0), rs(1)              ' or by index
'   Loop
'
' Other access:
'   rs.Fields.Count
'   rs.Fields(0).Name
'   rs.Fields("name").Value
'   rs.AtEnd
'   rs.LoadAll  ' read every remaining row into memory as a 2D array
'
Option Explicit

Private m_stmt        As cSQLiteStatement
Private m_eof         As Boolean
Private m_started     As Boolean   ' have we issued at least one MoveNext?
Private m_nameIndex   As Object    ' Scripting.Dictionary: name -> col index
                                   ' (built lazily; case-insensitive)

' --- attach (called by cSQLite.Query) -----------------------------

Friend Sub AttachStatement(ByRef stmt As cSQLiteStatement)
    Set m_stmt = stmt
    m_eof = False
    m_started = False
End Sub

' --- iteration ----------------------------------------------------

' Advance to the next row. Returns True if a row is now current,
' False if EOF.
Public Function MoveNext() As Boolean
    If m_stmt Is Nothing Then
        Err.Raise vbObjectError + 30300, "cSQLiteResults.MoveNext", _
            "results not attached"
    End If
    If m_eof Then
        MoveNext = False
        Exit Function
    End If
    m_started = True
    MoveNext = m_stmt.Step()
    If Not MoveNext Then m_eof = True
End Function

Public Property Get AtEnd() As Boolean
    AtEnd = m_eof
End Property

Public Property Get FieldCount() As Long
    If Not m_stmt Is Nothing Then FieldCount = m_stmt.ColumnCount
End Property

' Default property: rs(name_or_index) returns the current value as Variant.
' Marked as default via VB_UserMemId = 0 attribute below.
Public Property Get Item(ByVal key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    EnsureRow
    Dim idx As Long
    idx = ResolveIndex(key)
    Item = m_stmt.ColumnValue(idx)
End Property

' Strongly-typed getters for when you don't want Variant overhead.
Public Function FieldText(ByVal key As Variant) As String
    EnsureRow
    FieldText = m_stmt.ColumnText(ResolveIndex(key))
End Function

Public Function FieldInt(ByVal key As Variant) As Long
    EnsureRow
    FieldInt = m_stmt.ColumnInt(ResolveIndex(key))
End Function

Public Function FieldInt64(ByVal key As Variant) As Currency
    EnsureRow
    FieldInt64 = m_stmt.ColumnInt64(ResolveIndex(key))
End Function

Public Function FieldDouble(ByVal key As Variant) As Double
    EnsureRow
    FieldDouble = m_stmt.ColumnDouble(ResolveIndex(key))
End Function

Public Function FieldBlob(ByVal key As Variant) As Byte()
    EnsureRow
    FieldBlob = m_stmt.ColumnBlob(ResolveIndex(key))
End Function

Public Function FieldName(ByVal iCol As Long) As String
    If m_stmt Is Nothing Then Exit Function
    FieldName = m_stmt.ColumnName(iCol)
End Function

Public Function FieldType(ByVal key As Variant) As Long
    EnsureRow
    FieldType = m_stmt.ColumnType(ResolveIndex(key))
End Function

' --- bulk read ----------------------------------------------------

' Read every remaining row into a 2D Variant array.
' Returns Empty if no rows. Otherwise: arr(rowIdx, colIdx).
' The row index is 0-based; the column index is 0-based.
Public Function LoadAll() As Variant
    Dim rows As Collection
    Set rows = New Collection

    Dim nCols As Long
    nCols = FieldCount
    If nCols = 0 Then
        LoadAll = Empty
        Exit Function
    End If

    Do While MoveNext()
        Dim row() As Variant
        ReDim row(0 To nCols - 1)
        Dim i As Long
        For i = 0 To nCols - 1
            row(i) = m_stmt.ColumnValue(i)
        Next i
        rows.Add row
    Loop

    If rows.Count = 0 Then
        LoadAll = Empty
        Exit Function
    End If

    Dim out() As Variant
    ReDim out(0 To rows.Count - 1, 0 To nCols - 1)
    Dim r As Long, c As Long
    For r = 0 To rows.Count - 1
        Dim thisRow() As Variant
        thisRow = rows(r + 1)
        For c = 0 To nCols - 1
            out(r, c) = thisRow(c)
        Next c
    Next r
    LoadAll = out
End Function

' Return all column names as a 1D string array (handy for headers).
Public Function FieldNames() As String()
    Dim n As Long, i As Long
    Dim names() As String
    n = FieldCount
    If n = 0 Then
        ReDim names(0)
        FieldNames = names
        Exit Function
    End If
    ReDim names(0 To n - 1)
    For i = 0 To n - 1
        names(i) = m_stmt.ColumnName(i)
    Next i
    FieldNames = names
End Function

' --- close --------------------------------------------------------

Public Sub CloseRS()
    If Not m_stmt Is Nothing Then
        m_stmt.Finalize
        Set m_stmt = Nothing
    End If
    m_eof = True
End Sub

Private Sub Class_Terminate()
    CloseRS
End Sub

' --- internals ----------------------------------------------------

Private Sub EnsureRow()
    If m_stmt Is Nothing Then
        Err.Raise vbObjectError + 30301, "cSQLiteResults", "results not attached"
    End If
    If Not m_started Then
        Err.Raise vbObjectError + 30302, "cSQLiteResults", _
            "no current row — call MoveNext first"
    End If
    If m_eof Then
        Err.Raise vbObjectError + 30303, "cSQLiteResults", "past end of results"
    End If
End Sub

' Resolve a Variant key to a column index. Accepts:
'   - Long/Integer: used directly as 0-based index
'   - String: looked up in the (lazily-built) name index, case-insensitive
Private Function ResolveIndex(ByRef key As Variant) As Long
    If VarType(key) = vbString Then
        If m_nameIndex Is Nothing Then BuildNameIndex
        Dim k As String
        k = LCase$(CStr(key))
        If Not m_nameIndex.Exists(k) Then
            Err.Raise vbObjectError + 30304, "cSQLiteResults", _
                "no column named '" & key & "'"
        End If
        ResolveIndex = m_nameIndex(k)
    Else
        ResolveIndex = CLng(key)
    End If
End Function

Private Sub BuildNameIndex()
    Set m_nameIndex = CreateObject("Scripting.Dictionary")
    Dim n As Long, i As Long
    n = FieldCount
    For i = 0 To n - 1
        m_nameIndex(LCase$(m_stmt.ColumnName(i))) = i
    Next i
End Sub
