VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   1260
      TabIndex        =   2
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1260
      TabIndex        =   1
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1500
      TabIndex        =   0
      Top             =   420
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim t As cSQLiteTable
    Dim db As New cSQLite
    Dim f
    
    db.OpenDB App.path & "\sample.db"
    
    ' Basic case from the header
    Set t = New cSQLiteTable
    t.LoadQuery db, "SELECT name, age FROM users WHERE age > ?", 30
    ' Expect: 5 rows, 2 cols, both sqlInteger/sqlText
    
    ' Schema introspection
    For Each f In t.GetSchema()
        Debug.Print f.Ordinal, f.name, f.TypeName
    Next
    
    ' Mixed types - confirms sqlFloat hint
    t.LoadQuery db, "SELECT * FROM events"
    Debug.Print t.ColumnType("severity") = sqlFloat   ' True
    
    ' BLOB round-trip
    t.LoadQuery db, "SELECT label, payload FROM blobs"
    t.SaveJsonFile App.path & "\blobs.json"
    ' open the JSON in notepad - you should see {"$blob":"89504e470d0a1a0a"} etc.
    
    Dim t2 As New cSQLiteTable
    t2.LoadJsonFile App.path & "\blobs.json"
    Dim b() As Byte
    b = t2(1, "payload")     ' row 1 = png-header
    Debug.Print UBound(b) - LBound(b) + 1   ' 8
    
    ' NULL handling
    t.LoadQuery db, "SELECT * FROM nullable"
    Debug.Print IsNull(t(0, "a_int"))      ' True (all-null row)
    Debug.Print t.ColumnType("a_real") = sqlFloat   ' True (learned from row 3)

    t.LoadQuery db, "SELECT label, payload FROM blobs"
    t.SaveJsonFile App.path & "\blobs.json", True       ' pretty, so you can eyeball it
    
    ' open blobs.json in notepad - confirm you see lowercase hex like:
    '   ["png-header", {"$blob": "89504e470d0a1a0a"}]
    
    t2.LoadJsonFile App.path & "\blobs.json"

    b = t2(1, "payload")                       ' row 1 = png-header
    Debug.Print UBound(b) - LBound(b) + 1      ' should print 8
    Debug.Print Hex(b(0)) & " " & Hex(b(1))    ' should print 89 50


End Sub

Private Sub Command2_Click()
    
    Dim t As New cSQLiteTable
    t.SetColumns Array("id", "label")
    t.AddRow Array(1, "has ""quotes"" in it")
    t.AddRow Array(2, "path C:\temp\file.txt")
    t.AddRow Array(3, "O'Brien said ""hi""")
    t.AddRow Array(4, "json: {""k"":""v""}")
    t.AddRow Array(5, "trailing slash\")
    t.AddRow Array(6, "")
    t.AddRow Array(7, "line1" & vbCrLf & "line2" & vbTab & "tabbed")
    t.AddRow Array(8, "O'Brien")
    t.AddRow Array(9, "''; DROP TABLE users;--")
    t.AddRow Array(10, "it's a 'test'")
    t.AddRow Array(11, "'wrapped'")
    t.SaveJsonFile "C:\quote_test.json", True
    
    t.SaveJsonFile App.path & "\quote_test.json", True
    
    Dim t2 As New cSQLiteTable
    t2.LoadJsonFile App.path & "\quote_test.json"
    
    Dim i As Long
    For i = 0 To t2.RowCount - 1
        Debug.Print t2(i, "id"), "[" & t2(i, "label") & "]"
        Debug.Print "match: " & (CStr(t2(i, "label")) = CStr(t(i, "label")))
    Next

End Sub

Private Sub Command3_Click()
 
    Dim basePath As String
     basePath = App.path & "\"

    ' --- 1. Load the CSV ---
    Dim t As New cSQLiteTable
    t.LoadCsvFile basePath & "test_import.csv"   ' defaults: comma, hasHeader, typed

    Debug.Print "Loaded " & t.RowCount & " rows, " & t.ColumnCount & " columns"

    ' --- 2. Dump the schema as inferred from row data ---
    Debug.Print "--- Schema ---"
    Dim f As cSQLiteField
    For Each f In t.GetSchema()
        Debug.Print "  " & f.Ordinal & " " & f.name & " " & f.TypeName
    Next

    ' --- 3. Spot-check a few cells ---
    Debug.Print "--- Sample cells ---"
    Debug.Print "Row 0, name:        [" & CStr(t(0, "name")) & "]"
    Debug.Print "Row 1, name:        [" & CStr(t(1, "name")) & "]"   ' O'Brien
    Debug.Print "Row 2, name:        [" & CStr(t(2, "name")) & "]"   ' "Smith, John"
    Debug.Print "Row 2, notes:       [" & CStr(t(2, "notes")) & "]"  ' has embedded quotes
    Debug.Print "Row 3, notes:       [" & CStr(t(3, "notes")) & "]"  ' sql-injection-ish
    Debug.Print "Row 4, name isNull: " & IsNull(t(4, "name"))         ' empty in csv
    Debug.Print "Row 4, score isNull:" & IsNull(t(4, "score"))
    Debug.Print "Row 6, name:        [" & CStr(t(6, "name")) & "]"   ' multi-line value

    ' --- 4. Generate and save SQL dump ---
    t.SaveSqlFile basePath & "test_import.sql", "users", _
                  IfNotExists:=False, DropFirst:=True

    Debug.Print "--- SQL dumped to " & basePath & "test_import.sql ---"

    ' --- 5. Print first portion of the SQL for eyeballing ---
    Dim sql As String
    sql = t.ToSql("users", IfNotExists:=False, DropFirst:=True)
    Debug.Print Left$(sql, 1500)
    If Len(sql) > 1500 Then Debug.Print "...(" & (Len(sql) - 1500) & " more chars)"
 
End Sub
