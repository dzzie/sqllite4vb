VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com
'
' cSQLiteField - describes one column of a result set.
'
' This is the result-set-level analog of cSQLiteColumn (which describes a
' table column from PRAGMA table_info). A cSQLiteField carries just what
' the engine reports for the columns of a query result: position, name,
' and the runtime SQLite type (sqlInteger / sqlFloat / sqlText / sqlBlob /
' sqlNull).
'
' Usage:
'   Dim fields As Collection
'   Set fields = t.GetSchema()
'   Dim f As cSQLiteField
'   For Each f In fields
'       Debug.Print f.Ordinal, f.Name, SqlTypeName(f.SqlType)
'   Next
'
Option Explicit

Private m_ordinal As Long
Private m_name    As String
Private m_sqlType As SQLiteType

Friend Sub SetFrom(ByVal Ordinal As Long, ByVal name As String, ByVal SqlType As SQLiteType)
    m_ordinal = Ordinal
    m_name = name
    m_sqlType = SqlType
End Sub

' Zero-based column index within the result set.
Public Property Get Ordinal() As Long
    Ordinal = m_ordinal
End Property

Public Property Get name() As String
    name = m_name
End Property

' SQLite native type observed at load time. May be sqlNull if the column
' was NULL in every loaded row (or if the table was built manually without
' type information).
Public Property Get SqlType() As SQLiteType
    SqlType = m_sqlType
End Property

' Convenience: friendly type name for logging.
Public Property Get TypeName() As String
    TypeName = SqlTypeName(m_sqlType)
End Property

Private Function SqlTypeName(ByVal t As SQLiteType) As String
    Select Case t
        Case sqlInteger: SqlTypeName = "INTEGER"
        Case sqlFloat:   SqlTypeName = "FLOAT"
        Case sqlText:    SqlTypeName = "TEXT"
        Case sqlBlob:    SqlTypeName = "BLOB"
        Case sqlNull:    SqlTypeName = "NULL"
        Case Else:       SqlTypeName = "UNKNOWN"
    End Select
End Function

