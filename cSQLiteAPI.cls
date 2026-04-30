VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSQLiteAPI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com
'
' modSQLite.bas
'
' VB6 declarations for sqlite3_vb6.dll.
'
' All exports are __stdcall with undecorated names (via .def file),
' so no Alias gymnastics are needed.
'
' STRING MARSHALLING:
'   VB6's "ByVal x As String" in a Declare auto-converts UTF-16 -> ANSI
'   on the way out, which would lose any character outside the current
'   Windows code page. To preserve full Unicode we declare string inputs
'   as ByVal Long and pass StrPtr(s) at the call site — the C side
'   receives the raw UTF-16 buffer pointer and converts to UTF-8 itself.
'
'   For string OUTPUTS, we use "As Variant" rather than "As String" on
'   the raw declares. Why: VB6's Declare treats "As String" returns as
'   LPSTR (ANSI char*), reading bytes until NUL — which truncates
'   UTF-16 at the first 0x00 high byte (so "Alice" becomes "A").
'   "As Variant" doesn't have that baggage; it carries a real BSTR
'   with length prefix and full Unicode round-trips correctly. The
'   public wrappers do CStr() to convert back to a normal VB6 String.
'
'   The Public wrapper functions hide all this — callers just pass
'   strings normally.
'
' HANDLES:
'   sqlite3* and sqlite3_stmt* are 32-bit pointers (VB6 is x86), held
'   as Long in VB6.
'
' INT64:
'   Marshalled as Currency. Currency is an 8-byte signed integer with
'   a /10000 scale factor in VB6's interpretation, but the bytes on
'   the stack ARE int64. So bind/column round-trips are byte-correct;
'   only the *displayed* number is scaled. Use Int64FromCurrency /
'   CurrencyFromInt64 below to convert between the two views when
'   you want to show the raw integer.
'
Option Explicit

' ------------------------------------------------------------------
' Result codes
' ------------------------------------------------------------------
'Public Const SQLITE_OK         As Long = 0
'Public Const SQLITE_ERROR      As Long = 1
'Public Const SQLITE_BUSY       As Long = 5
'Public Const SQLITE_LOCKED     As Long = 6
'Public Const SQLITE_NOMEM      As Long = 7
'Public Const SQLITE_READONLY   As Long = 8
'Public Const SQLITE_INTERRUPT  As Long = 9
'Public Const SQLITE_IOERR      As Long = 10
'Public Const SQLITE_CORRUPT    As Long = 11
'Public Const SQLITE_FULL       As Long = 13
'Public Const SQLITE_CANTOPEN   As Long = 14
'Public Const SQLITE_CONSTRAINT As Long = 19
'Public Const SQLITE_MISMATCH   As Long = 20
'Public Const SQLITE_MISUSE     As Long = 21
'Public Const SQLITE_RANGE      As Long = 25
'Public Const SQLITE_NOTADB     As Long = 26
'Public Const SQLITE_ROW        As Long = 100
'Public Const SQLITE_DONE       As Long = 101
'
'' Column types
'Public Const SQLITE_INTEGER As Long = 1
'Public Const SQLITE_FLOAT   As Long = 2
'Public Const SQLITE_TEXT    As Long = 3
'Public Const SQLITE_BLOB    As Long = 4
'Public Const SQLITE_NULL    As Long = 5
'
'' open_v2 flags
'Public Const SQLITE_OPEN_READONLY      As Long = &H1
'Public Const SQLITE_OPEN_READWRITE     As Long = &H2
'Public Const SQLITE_OPEN_CREATE        As Long = &H4
'Public Const SQLITE_OPEN_URI           As Long = &H40
'Public Const SQLITE_OPEN_MEMORY        As Long = &H80
'Public Const SQLITE_OPEN_NOMUTEX       As Long = &H8000
'Public Const SQLITE_OPEN_FULLMUTEX     As Long = &H10000

Public Enum SQLiteConsts
    ' --- result codes ---
    SQLITE_OK = 0
    SQLITE_ERROR = 1
    SQLITE_BUSY = 5
    SQLITE_LOCKED = 6
    SQLITE_NOMEM = 7
    SQLITE_READONLY = 8
    SQLITE_INTERRUPT = 9
    SQLITE_IOERR = 10
    SQLITE_CORRUPT = 11
    SQLITE_FULL = 13
    SQLITE_CANTOPEN = 14
    SQLITE_CONSTRAINT = 19
    SQLITE_MISMATCH = 20
    SQLITE_MISUSE = 21
    SQLITE_RANGE = 25
    SQLITE_NOTADB = 26
    SQLITE_ROW = 100
    SQLITE_DONE = 101

    ' --- column storage types ---
    ' (note: SQLITE_INTEGER shares value 1 with SQLITE_ERROR � different
    '  namespaces in the C API, harmless in VB6 since they're separate names)
    SQLITE_INTEGER = 1
    SQLITE_FLOAT = 2
    SQLITE_TEXT = 3
    SQLITE_BLOB = 4
    SQLITE_NULL = 5
End Enum

Public Enum sqlLiteOpenConst
    ' --- sqlite3_open_v2 flags (bitwise OR-able) ---
    SQLITE_OPEN_READONLY = &H1&
    SQLITE_OPEN_READWRITE = &H2&
    SQLITE_OPEN_CREATE = &H4&
    SQLITE_OPEN_URI = &H40&
    SQLITE_OPEN_MEMORY = &H80&
    SQLITE_OPEN_NOMUTEX = &H8000&
    SQLITE_OPEN_FULLMUTEX = &H10000
End Enum


' ------------------------------------------------------------------
' Raw declares (Private). Public wrappers below.
' ------------------------------------------------------------------

' Lifecycle / version
Private Declare Function raw_libversion Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_libversion" () As Variant
Private Declare Function raw_libversion_number Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_libversion_number" () As Long
Private Declare Function raw_threadsafe Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_threadsafe" () As Long
Private Declare Function raw_initialize Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_initialize" () As Long
Private Declare Function raw_shutdown Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_shutdown" () As Long

' Connection
Private Declare Function raw_open Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_open" ( _
    ByVal pFilename As Long, ByRef ppDb As Long) As Long
Private Declare Function raw_open_v2 Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_open_v2" ( _
    ByVal pFilename As Long, ByRef ppDb As Long, _
    ByVal flags As Long, ByVal pVfs As Long) As Long
Private Declare Function raw_close Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_close" (ByVal db As Long) As Long
Private Declare Function raw_close_v2 Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_close_v2" (ByVal db As Long) As Long

' Errors
Private Declare Function raw_errcode Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_errcode" (ByVal db As Long) As Long
Private Declare Function raw_extended_errcode Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_extended_errcode" (ByVal db As Long) As Long
Private Declare Function raw_errmsg Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_errmsg" (ByVal db As Long) As Variant
Private Declare Function raw_errstr Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_errstr" (ByVal rc As Long) As Variant

' Exec
' NEW � correct:
Private Declare Function sqlite3_exec_simple Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_exec_simple" _
    (ByVal hDb As Long, ByVal sql As Long, ByRef errMsg As Variant) As Long
    
' Prepare / step / reset / finalize
Private Declare Function raw_prepare_v2 Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_prepare_v2" ( _
    ByVal db As Long, ByVal pSql As Long, ByRef ppStmt As Long) As Long
Private Declare Function raw_step Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_step" (ByVal stmt As Long) As Long
Private Declare Function raw_reset Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_reset" (ByVal stmt As Long) As Long
Private Declare Function raw_finalize Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_finalize" (ByVal stmt As Long) As Long
Private Declare Function raw_clear_bindings Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_clear_bindings" (ByVal stmt As Long) As Long

' Bind
Private Declare Function raw_bind_parameter_count Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_bind_parameter_count" (ByVal stmt As Long) As Long
Private Declare Function raw_bind_parameter_index Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_bind_parameter_index" ( _
    ByVal stmt As Long, ByVal pName As Long) As Long
Private Declare Function raw_bind_null Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_bind_null" (ByVal stmt As Long, ByVal idx As Long) As Long
Private Declare Function raw_bind_int Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_bind_int" ( _
    ByVal stmt As Long, ByVal idx As Long, ByVal value As Long) As Long
Private Declare Function raw_bind_int64 Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_bind_int64" ( _
    ByVal stmt As Long, ByVal idx As Long, ByVal value As Currency) As Long
Private Declare Function raw_bind_double Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_bind_double" ( _
    ByVal stmt As Long, ByVal idx As Long, ByVal value As Double) As Long
Private Declare Function raw_bind_text Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_bind_text" ( _
    ByVal stmt As Long, ByVal idx As Long, ByVal pValue As Long) As Long
Private Declare Function raw_bind_blob Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_bind_blob" ( _
    ByVal stmt As Long, ByVal idx As Long, _
    ByVal pData As Long, ByVal length As Long) As Long
Private Declare Function raw_bind_zeroblob Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_bind_zeroblob" ( _
    ByVal stmt As Long, ByVal idx As Long, ByVal n As Long) As Long

' Column metadata
Private Declare Function raw_column_count Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_count" (ByVal stmt As Long) As Long
Private Declare Function raw_data_count Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_data_count" (ByVal stmt As Long) As Long
Private Declare Function raw_column_name Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_name" ( _
    ByVal stmt As Long, ByVal iCol As Long) As Variant
Private Declare Function raw_column_decltype Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_decltype" ( _
    ByVal stmt As Long, ByVal iCol As Long) As Variant
Private Declare Function raw_column_type Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_type" ( _
    ByVal stmt As Long, ByVal iCol As Long) As Long

' Column values
Private Declare Function raw_column_int Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_int" ( _
    ByVal stmt As Long, ByVal iCol As Long) As Long
Private Declare Function raw_column_int64 Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_int64" ( _
    ByVal stmt As Long, ByVal iCol As Long) As Currency
Private Declare Function raw_column_double Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_double" ( _
    ByVal stmt As Long, ByVal iCol As Long) As Double
Private Declare Function raw_column_text Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_text" ( _
    ByVal stmt As Long, ByVal iCol As Long) As Variant
Private Declare Function raw_column_bytes Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_bytes" ( _
    ByVal stmt As Long, ByVal iCol As Long) As Long
Private Declare Function raw_column_blob_copy Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_column_blob_copy" ( _
    ByVal stmt As Long, ByVal iCol As Long, _
    ByVal pDest As Long, ByVal destSize As Long) As Long

' Misc
Private Declare Function raw_last_insert_rowid Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_last_insert_rowid" (ByVal db As Long) As Currency
Private Declare Function raw_changes Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_changes" (ByVal db As Long) As Long
Private Declare Function raw_total_changes Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_total_changes" (ByVal db As Long) As Long
Private Declare Function raw_busy_timeout Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_busy_timeout" (ByVal db As Long, ByVal ms As Long) As Long
Private Declare Function raw_get_autocommit Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_get_autocommit" (ByVal db As Long) As Long
Private Declare Sub raw_interrupt Lib "sqlite3_vb6.dll" _
    Alias "vb6_sqlite3_interrupt" (ByVal db As Long)

' ==================================================================
' Public API — clean wrappers that hide the StrPtr trick.
' ==================================================================

' Convert a Variant returned from the C side to a String. VT_NULL
' (which the C side uses to mean "input was NULL") becomes "".
Private Function VariantToString(ByRef v As Variant) As String
    If IsNull(v) Then
        VariantToString = vbNullString
    Else
        VariantToString = CStr(v)
    End If
End Function

Public Function sqlite3_libversion() As String
    sqlite3_libversion = VariantToString(raw_libversion())
End Function

Public Function sqlite3_libversion_number() As Long
    sqlite3_libversion_number = raw_libversion_number()
End Function

Public Function sqlite3_threadsafe() As Long
    sqlite3_threadsafe = raw_threadsafe()
End Function

Public Function sqlite3_initialize() As Long
    sqlite3_initialize = raw_initialize()
End Function

Public Function sqlite3_shutdown() As Long
    sqlite3_shutdown = raw_shutdown()
End Function

Public Function sqlite3_open(ByVal sFilename As String, ByRef hDb As Long) As Long
    sqlite3_open = raw_open(StrPtr(sFilename), hDb)
End Function

Public Function sqlite3_open_v2(ByVal sFilename As String, ByRef hDb As Long, _
                                ByVal flags As Long, _
                                Optional ByVal sVfs As String = vbNullString) As Long
    Dim pVfs As Long
    If LenB(sVfs) > 0 Then pVfs = StrPtr(sVfs)
    sqlite3_open_v2 = raw_open_v2(StrPtr(sFilename), hDb, flags, pVfs)
End Function

Public Function sqlite3_close(ByVal hDb As Long) As Long
    sqlite3_close = raw_close(hDb)
End Function

Public Function sqlite3_close_v2(ByVal hDb As Long) As Long
    sqlite3_close_v2 = raw_close_v2(hDb)
End Function

Public Function sqlite3_errcode(ByVal hDb As Long) As Long
    sqlite3_errcode = raw_errcode(hDb)
End Function

Public Function sqlite3_extended_errcode(ByVal hDb As Long) As Long
    sqlite3_extended_errcode = raw_extended_errcode(hDb)
End Function

Public Function sqlite3_errmsg(ByVal hDb As Long) As String
    sqlite3_errmsg = VariantToString(raw_errmsg(hDb))
End Function

Public Function sqlite3_errstr(ByVal rc As Long) As String
    sqlite3_errstr = VariantToString(raw_errstr(rc))
End Function

Public Function sqlite3_exec(ByVal hDb As Long, ByVal sql As String, _
                             Optional ByRef sErr As String) As Long
    Dim vErr As Variant
    sqlite3_exec = sqlite3_exec_simple(hDb, StrPtr(sql), vErr)
    If IsNull(vErr) Or IsEmpty(vErr) Then
        sErr = ""
    Else
        sErr = CStr(vErr)
    End If
End Function

Public Function sqlite3_prepare_v2(ByVal hDb As Long, ByVal sSql As String, _
                                   ByRef hStmt As Long) As Long
    sqlite3_prepare_v2 = raw_prepare_v2(hDb, StrPtr(sSql), hStmt)
End Function

Public Function sqlite3_step(ByVal hStmt As Long) As Long
    sqlite3_step = raw_step(hStmt)
End Function

Public Function sqlite3_reset(ByVal hStmt As Long) As Long
    sqlite3_reset = raw_reset(hStmt)
End Function

Public Function sqlite3_finalize(ByVal hStmt As Long) As Long
    sqlite3_finalize = raw_finalize(hStmt)
End Function

Public Function sqlite3_clear_bindings(ByVal hStmt As Long) As Long
    sqlite3_clear_bindings = raw_clear_bindings(hStmt)
End Function

Public Function sqlite3_bind_parameter_count(ByVal hStmt As Long) As Long
    sqlite3_bind_parameter_count = raw_bind_parameter_count(hStmt)
End Function

Public Function sqlite3_bind_parameter_index(ByVal hStmt As Long, _
                                             ByVal sName As String) As Long
    sqlite3_bind_parameter_index = raw_bind_parameter_index(hStmt, StrPtr(sName))
End Function

Public Function sqlite3_bind_null(ByVal hStmt As Long, ByVal idx As Long) As Long
    sqlite3_bind_null = raw_bind_null(hStmt, idx)
End Function

Public Function sqlite3_bind_int(ByVal hStmt As Long, ByVal idx As Long, _
                                 ByVal value As Long) As Long
    sqlite3_bind_int = raw_bind_int(hStmt, idx, value)
End Function

Public Function sqlite3_bind_int64(ByVal hStmt As Long, ByVal idx As Long, _
                                   ByVal value As Currency) As Long
    sqlite3_bind_int64 = raw_bind_int64(hStmt, idx, value)
End Function

Public Function sqlite3_bind_double(ByVal hStmt As Long, ByVal idx As Long, _
                                    ByVal value As Double) As Long
    sqlite3_bind_double = raw_bind_double(hStmt, idx, value)
End Function

Public Function sqlite3_bind_text(ByVal hStmt As Long, ByVal idx As Long, _
                                  ByVal value As String) As Long
    sqlite3_bind_text = raw_bind_text(hStmt, idx, StrPtr(value))
End Function

Public Function sqlite3_bind_blob(ByVal hStmt As Long, ByVal idx As Long, _
                                  ByRef bytes() As Byte) As Long
    Dim n As Long
    On Error Resume Next
    n = UBound(bytes) - LBound(bytes) + 1
    If Err.Number <> 0 Then
        Err.Clear
        sqlite3_bind_blob = raw_bind_null(hStmt, idx)
        Exit Function
    End If
    On Error GoTo 0
    If n <= 0 Then
        sqlite3_bind_blob = raw_bind_null(hStmt, idx)
    Else
        sqlite3_bind_blob = raw_bind_blob(hStmt, idx, _
                                          VarPtr(bytes(LBound(bytes))), n)
    End If
End Function

Public Function sqlite3_bind_zeroblob(ByVal hStmt As Long, ByVal idx As Long, _
                                      ByVal n As Long) As Long
    sqlite3_bind_zeroblob = raw_bind_zeroblob(hStmt, idx, n)
End Function

Public Function sqlite3_column_count(ByVal hStmt As Long) As Long
    sqlite3_column_count = raw_column_count(hStmt)
End Function

Public Function sqlite3_data_count(ByVal hStmt As Long) As Long
    sqlite3_data_count = raw_data_count(hStmt)
End Function

Public Function sqlite3_column_name(ByVal hStmt As Long, ByVal iCol As Long) As String
    sqlite3_column_name = VariantToString(raw_column_name(hStmt, iCol))
End Function

Public Function sqlite3_column_decltype(ByVal hStmt As Long, ByVal iCol As Long) As String
    sqlite3_column_decltype = VariantToString(raw_column_decltype(hStmt, iCol))
End Function

Public Function sqlite3_column_type(ByVal hStmt As Long, ByVal iCol As Long) As Long
    sqlite3_column_type = raw_column_type(hStmt, iCol)
End Function

Public Function sqlite3_column_int(ByVal hStmt As Long, ByVal iCol As Long) As Long
    sqlite3_column_int = raw_column_int(hStmt, iCol)
End Function

Public Function sqlite3_column_int64(ByVal hStmt As Long, ByVal iCol As Long) As Currency
    sqlite3_column_int64 = raw_column_int64(hStmt, iCol)
End Function

Public Function sqlite3_column_double(ByVal hStmt As Long, ByVal iCol As Long) As Double
    sqlite3_column_double = raw_column_double(hStmt, iCol)
End Function

Public Function sqlite3_column_text(ByVal hStmt As Long, ByVal iCol As Long) As String
    sqlite3_column_text = VariantToString(raw_column_text(hStmt, iCol))
End Function

Public Function sqlite3_column_bytes(ByVal hStmt As Long, ByVal iCol As Long) As Long
    sqlite3_column_bytes = raw_column_bytes(hStmt, iCol)
End Function

' Returns a Byte array sized to the blob (zero-length if NULL/empty).
Public Function sqlite3_column_blob(ByVal hStmt As Long, ByVal iCol As Long) As Byte()
    Dim n As Long, copied As Long
    Dim out() As Byte
    n = raw_column_bytes(hStmt, iCol)
    If n <= 0 Then
        sqlite3_column_blob = vbNullString  ' coerces to empty Byte array
        Exit Function
    End If
    ReDim out(0 To n - 1)
    copied = raw_column_blob_copy(hStmt, iCol, VarPtr(out(0)), n)
    If copied < n And copied > 0 Then ReDim Preserve out(0 To copied - 1)
    sqlite3_column_blob = out
End Function

Public Function sqlite3_last_insert_rowid(ByVal hDb As Long) As Currency
    sqlite3_last_insert_rowid = raw_last_insert_rowid(hDb)
End Function

Public Function sqlite3_changes(ByVal hDb As Long) As Long
    sqlite3_changes = raw_changes(hDb)
End Function

Public Function sqlite3_total_changes(ByVal hDb As Long) As Long
    sqlite3_total_changes = raw_total_changes(hDb)
End Function

Public Function sqlite3_busy_timeout(ByVal hDb As Long, ByVal ms As Long) As Long
    sqlite3_busy_timeout = raw_busy_timeout(hDb, ms)
End Function

Public Function sqlite3_get_autocommit(ByVal hDb As Long) As Long
    sqlite3_get_autocommit = raw_get_autocommit(hDb)
End Function

Public Sub sqlite3_interrupt(ByVal hDb As Long)
    raw_interrupt hDb
End Sub

' ==================================================================
' Convenience helpers
' ==================================================================

' Open a DB and return the handle, raising on error.
Public Function SQLiteOpen(ByVal sFilename As String) As Long
    Dim hDb As Long, rc As Long
    rc = sqlite3_open(sFilename, hDb)
    If rc <> SQLITE_OK Then
        Dim sErr As String
        If hDb <> 0 Then sErr = sqlite3_errmsg(hDb)
        If hDb <> 0 Then sqlite3_close hDb
        Err.Raise vbObjectError + 30000 + rc, "SQLiteOpen", _
            "sqlite3_open failed (rc=" & rc & "): " & sErr
    End If
    SQLiteOpen = hDb
End Function

' Execute a no-result SQL statement. Raises on error.
Public Sub SQLiteExec(ByVal hDb As Long, ByVal sSql As String)
    Dim sErr As String, rc As Long
    rc = sqlite3_exec(hDb, sSql, sErr)
    If rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 30000 + rc, "SQLiteExec", _
            "exec failed (rc=" & rc & "): " & sErr & vbCrLf & "SQL: " & sSql
    End If
End Sub

' Currency<->Int64 conversion. Currency stores int64/10000 — multiply
' or divide by 10000 to convert between the two views.
Public Function Int64FromCurrency(ByVal c As Currency) As Variant
    Int64FromCurrency = CDec(c) * CDec(10000)
End Function

Public Function CurrencyFromInt64(ByVal v As Variant) As Currency
    CurrencyFromInt64 = CCur(CDec(v) / CDec(10000))
End Function

Private Sub Class_Initialize()
    EnsureDLL
End Sub
