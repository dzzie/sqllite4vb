VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com
'
' cSQLiteColumn — column metadata returned by cSQLiteSchema.DescribeTable.
'
Option Explicit

Private m_cid          As Long
Private m_name         As String
Private m_declType     As String
Private m_notNull      As Boolean
Private m_defaultValue As String
Private m_pkOrdinal    As Long

Friend Sub SetFromPragma(ByVal cid As Long, ByVal colName As String, _
                          ByVal DeclType As String, ByVal NotNull As Boolean, _
                          ByVal defaultVal As String, ByVal PKOrdinal As Long)
    m_cid = cid
    m_name = colName
    m_declType = DeclType
    m_notNull = NotNull
    m_defaultValue = defaultVal
    m_pkOrdinal = PKOrdinal
End Sub

Public Property Get ColumnIndex() As Long: ColumnIndex = m_cid: End Property
Public Property Get Name() As String: Name = m_name: End Property
Public Property Get DeclType() As String: DeclType = m_declType: End Property
Public Property Get NotNull() As Boolean: NotNull = m_notNull: End Property
Public Property Get DefaultValue() As String: DefaultValue = m_defaultValue: End Property
Public Property Get PKOrdinal() As Long: PKOrdinal = m_pkOrdinal: End Property
Public Property Get IsPrimaryKey() As Boolean: IsPrimaryKey = (m_pkOrdinal > 0): End Property
