VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJsonParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'License: MIT
'Author:  David Zimmer
'Site:    http://sandsprite.com
'
' Minimal JSON parser used by cSQLiteTable.LoadJson.
' Recursive descent, no external dependencies.
'
' Output value types:
'   object -> Scripting.Dictionary (case-sensitive keys, insertion order)
'   array  -> VBA.Collection (1-based)
'   string -> String
'   number -> Double  (JSON doesn't distinguish int vs float anyway)
'   bool   -> Boolean
'   null   -> Null    (vbNull, distinguishable from Empty)
'
' Usage:
'   Dim p As New CJsonParser
'   Dim root As Variant
'   root = p.Parse(jsonText)
'   ' inspect with IsObject(root), TypeName(root)
'
Option Explicit

Private m_s    As String
Private m_pos  As Long
Private m_len  As Long

' Top-level. Returns a Variant which may hold an Object (Dictionary or
' Collection) or a primitive. Use IsObject() to discriminate.
Public Function Parse(ByVal jsonText As String) As Variant
    m_s = jsonText
    m_pos = 1
    m_len = Len(jsonText)

    SkipWs
    Dim v As Variant
    Select Case Peek()
        Case "{": Set v = ParseObject()
        Case "[": Set v = ParseArray()
        Case """", "t", "f", "n", "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            v = ParseScalar()
        Case Else
            Err.Raise vbObjectError + 30500, "JsonParser.Parse", _
                "unexpected character at position " & m_pos
    End Select

    SkipWs
    If m_pos <= m_len Then
        Err.Raise vbObjectError + 30501, "JsonParser.Parse", _
            "trailing garbage at position " & m_pos
    End If

    If IsObject(v) Then Set Parse = v Else Parse = v
End Function

' --- object ------------------------------------------------------

Private Function ParseObject() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0   ' binary / case-sensitive

    ConsumeChar "{"
    SkipWs

    If Peek() = "}" Then
        m_pos = m_pos + 1
        Set ParseObject = d
        Exit Function
    End If

    Do
        SkipWs
        If Peek() <> """" Then
            Err.Raise vbObjectError + 30503, "JsonParser", _
                "expected string key at position " & m_pos
        End If
        Dim k As String
        k = ParseString()
        SkipWs
        ConsumeChar ":"

        Dim v As Variant
        SkipWs
        Select Case Peek()
            Case "{": Set v = ParseObject()
            Case "[": Set v = ParseArray()
            Case Else: v = ParseScalar()
        End Select

        If d.Exists(k) Then d.Remove k   ' last write wins
        d.Add k, v

        SkipWs
        Select Case Peek()
            Case ","
                m_pos = m_pos + 1
            Case "}"
                m_pos = m_pos + 1
                Set ParseObject = d
                Exit Function
            Case Else
                Err.Raise vbObjectError + 30504, "JsonParser", _
                    "expected ',' or '}' at position " & m_pos
        End Select
    Loop
End Function

' --- array -------------------------------------------------------

Private Function ParseArray() As Collection
    Dim c As New Collection
    ConsumeChar "["
    SkipWs

    If Peek() = "]" Then
        m_pos = m_pos + 1
        Set ParseArray = c
        Exit Function
    End If

    Do
        Dim v As Variant
        SkipWs
        Select Case Peek()
            Case "{": Set v = ParseObject()
            Case "[": Set v = ParseArray()
            Case Else: v = ParseScalar()
        End Select
        c.Add v

        SkipWs
        Select Case Peek()
            Case ","
                m_pos = m_pos + 1
            Case "]"
                m_pos = m_pos + 1
                Set ParseArray = c
                Exit Function
            Case Else
                Err.Raise vbObjectError + 30505, "JsonParser", _
                    "expected ',' or ']' at position " & m_pos
        End Select
    Loop
End Function

' --- scalar (string/number/bool/null) ----------------------------

Private Function ParseScalar() As Variant
    SkipWs
    Dim ch As String
    ch = Peek()
    Select Case ch
        Case """":      ParseScalar = ParseString()
        Case "t", "f":  ParseScalar = ParseBool()
        Case "n"
            Expect "null"
            ParseScalar = Null
        Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            ParseScalar = ParseNumber()
        Case Else
            Err.Raise vbObjectError + 30502, "JsonParser", _
                "unexpected character '" & ch & "' at position " & m_pos
    End Select
End Function

' --- string ------------------------------------------------------

Private Function ParseString() As String
    ConsumeChar """"
    Dim sb As String
    Dim ch As String
    Do
        If m_pos > m_len Then
            Err.Raise vbObjectError + 30506, "JsonParser", "unterminated string"
        End If
        ch = Mid$(m_s, m_pos, 1)
        m_pos = m_pos + 1

        If ch = """" Then
            ParseString = sb
            Exit Function
        ElseIf ch = "\" Then
            If m_pos > m_len Then
                Err.Raise vbObjectError + 30507, "JsonParser", "bad escape"
            End If
            Dim esc As String
            esc = Mid$(m_s, m_pos, 1)
            m_pos = m_pos + 1
            Select Case esc
                Case """": sb = sb & """"
                Case "\":  sb = sb & "\"
                Case "/":  sb = sb & "/"
                Case "b":  sb = sb & Chr$(8)
                Case "t":  sb = sb & vbTab
                Case "n":  sb = sb & vbLf
                Case "f":  sb = sb & Chr$(12)
                Case "r":  sb = sb & vbCr
                Case "u"
                    If m_pos + 3 > m_len Then
                        Err.Raise vbObjectError + 30508, "JsonParser", "bad \u escape"
                    End If
                    Dim hex4 As String
                    hex4 = Mid$(m_s, m_pos, 4)
                    m_pos = m_pos + 4
                    sb = sb & ChrW$(CLng("&H" & hex4))
                Case Else
                    Err.Raise vbObjectError + 30509, "JsonParser", _
                        "unknown escape '\" & esc & "'"
            End Select
        Else
            sb = sb & ch
        End If
    Loop
End Function

' --- number ------------------------------------------------------

Private Function ParseNumber() As Double
    Dim start As Long
    start = m_pos
    If Mid$(m_s, m_pos, 1) = "-" Then m_pos = m_pos + 1
    Do While m_pos <= m_len
        Dim ch As String
        ch = Mid$(m_s, m_pos, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "e" Or ch = "E" _
           Or ch = "+" Or ch = "-" Then
            m_pos = m_pos + 1
        Else
            Exit Do
        End If
    Loop
    Dim numStr As String
    numStr = Mid$(m_s, start, m_pos - start)
    ' Val() uses "." regardless of locale - exactly what JSON requires.
    ParseNumber = Val(numStr)
End Function

' --- bool --------------------------------------------------------

Private Function ParseBool() As Boolean
    If Mid$(m_s, m_pos, 4) = "true" Then
        m_pos = m_pos + 4
        ParseBool = True
    ElseIf Mid$(m_s, m_pos, 5) = "false" Then
        m_pos = m_pos + 5
        ParseBool = False
    Else
        Err.Raise vbObjectError + 30510, "JsonParser", _
            "expected true/false at position " & m_pos
    End If
End Function

' --- helpers -----------------------------------------------------

Private Sub SkipWs()
    Do While m_pos <= m_len
        Select Case Mid$(m_s, m_pos, 1)
            Case " ", vbTab, vbCr, vbLf
                m_pos = m_pos + 1
            Case Else
                Exit Sub
        End Select
    Loop
End Sub

Private Function Peek() As String
    If m_pos > m_len Then
        Peek = ""
    Else
        Peek = Mid$(m_s, m_pos, 1)
    End If
End Function

Private Sub ConsumeChar(ByVal expected As String)
    If m_pos > m_len Or Mid$(m_s, m_pos, 1) <> expected Then
        Err.Raise vbObjectError + 30511, "JsonParser", _
            "expected '" & expected & "' at position " & m_pos
    End If
    m_pos = m_pos + 1
End Sub

Private Sub Expect(ByVal lit As String)
    If Mid$(m_s, m_pos, Len(lit)) <> lit Then
        Err.Raise vbObjectError + 30512, "JsonParser", _
            "expected '" & lit & "' at position " & m_pos
    End If
    m_pos = m_pos + Len(lit)
End Sub
