Attribute VB_Name = "modSQLite"
'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com

Global api As New cSQLiteAPI


Option Explicit

Private Declare Function GetModuleHandleEx Lib "kernel32" _
    Alias "GetModuleHandleExA" ( _
    ByVal dwFlags As Long, _
    ByVal lpModuleName As Long, _
    ByRef phModule As Long) As Long

Private Declare Function GetModuleFileName Lib "kernel32" _
    Alias "GetModuleFileNameA" ( _
    ByVal hModule As Long, _
    ByVal lpFilename As String, _
    ByVal nSize As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Const GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS As Long = &H4
Private Const GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT As Long = &H2

Dim hDll As Long

Function EnsureDLL() As Boolean

    Dim f As String
    
    If hDll <> 0 Then
        EnsureDLL = True
        Exit Function
    End If
    
    hDll = GetModuleHandle("sqlite3_vb6.dll")
    If hDll <> 0 Then
        EnsureDLL = True
        Exit Function
    End If
    
    f = DllFolder() & "\sqlite3_vb6.dll"
    hDll = LoadLibrary(f)
    If hDll <> 0 Then
        EnsureDLL = True
        Exit Function
    End If
    
End Function


' Returns the full path of the DLL that contains this code.
' Inside an ActiveX DLL, this is the DLL's own path — *not* App.Path,
' which would be the host process.
Public Function DllFullPath() As String
    Dim hMod As Long
    Dim buf As String

    ' Pass the address of a function in this module — GetModuleHandleEx
    ' looks up which module owns that address. UNCHANGED_REFCOUNT means
    ' we don't have to FreeLibrary later.
    Dim flags As Long
    flags = GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS Or _
            GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT

    If GetModuleHandleEx(flags, AddressOf DllFullPath_Anchor, hMod) = 0 Then
        Exit Function
    End If

    buf = String$(260, vbNullChar)
    Dim n As Long
    n = GetModuleFileName(hMod, buf, Len(buf))
    If n > 0 Then
        DllFullPath = Left$(buf, n)
    End If
End Function

Public Function DllFolder() As String
    Dim p As String
    p = DllFullPath
    If Len(p) > 0 Then
        Dim i As Long
        i = InStrRev(p, "\")
        If i > 0 Then DllFolder = Left$(p, i - 1)
    End If
End Function

' Anchor function whose address we use to identify this module.
' AddressOf only works on functions in .bas modules (not classes), and
' only on functions that are actually used somewhere — so we use it
' inside DllFullPath above.
Public Sub DllFullPath_Anchor()
    ' empty by design — we just need its address
End Sub



