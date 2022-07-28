Attribute VB_Name = "Module4"
' functions

Option Explicit

' ------------------------------------
' main

Public Function makeTestSheet(q As Variant, sn As Integer, en As Integer)
    Dim sname As String
    sname = copyTemplate
    If sname = "" Then
        Exit Function
    End If
    
    Dim s As Worksheet
    Set s = ActiveSheet
    
    s.Cells(cover_row, cover_col).Value = "(" & sn & " - " & en & ")"
    s.range(s.Cells(q_srow, q_scol), s.Cells(q_srow + numQ - 1, q_scol)) = q
End Function

Public Function getAlldb() As Variant
    Dim arr As Variant
    Dim erow As Long
    
    erow = db.Cells(1, 1).End(xlDown).Row
    arr = db.range(db.Cells(1, 1), db.Cells(erow, 3))
    
    getAlldb = arr
End Function

Public Function copyTemplate() As String
    Dim sname As String
    sname = Format(Now, "yyyymmdd_hhmmss")
    
    temp.Copy after:=this.Worksheets(this.Worksheets.count)
    
    On Error GoTo err
    ActiveSheet.name = sname
    copyTemplate = sname
    Exit Function
    
err:
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    MsgBox _
        "シート名が重複しています。" & vbLf & _
        "シート名：" & sname & vbLf & vbLf & _
        "時間をあけて再度お試しください。", vbInformation
    sname = ""
    copyTemplate = sname
End Function

Public Function shuffleArray(ar() As Integer) As Integer()
    Dim iRnd As Long
    Dim s
    Dim i As Integer
    
    Call Randomize
    
    For i = UBound(ar) To 0 Step -1
        iRnd = Int(i * Rnd)
        
        s = ar(iRnd)
        ar(iRnd) = ar(i)
        ar(i) = s
    Next
    
    shuffleArray = ar()
End Function

' ------------------------------------
' student

Public Function whoIsStudent() As Variant
    Dim erow As Long
    Dim i As Integer
    Dim count As Integer
    Dim student(1) As Variant
    
    erow = top.Cells(Rows.count, 2).End(xlUp).Row
    
    If erow = 1 Then
        MsgBox "生徒が存在しません。"
        Exit Function
    End If
    
    count = 0
    For i = 2 To erow
        If top.Cells(i, 1) <> "" Then
            student(0) = i
            student(1) = top.Cells(i, 2).Value
            count = count + 1
        End If
    Next
    
    If count = 0 Then
        MsgBox "生徒を指定してください。"
        GoTo returnNull
    ElseIf count = 1 Then
        whoIsStudent = student
    Else
        MsgBox "生徒は1人ずつ指定してください。"
        GoTo returnNull
    End If
    Exit Function
    
returnNull:
    student(0) = 0
    student(1) = ""
    whoIsStudent = student
    Exit Function
End Function

Public Function makeStudentSheet(sname As String) As String
    Dim db As Variant
    db = getAlldb
    
    Dim s As Worksheet
    this.Worksheets.Add after:=this.Worksheets(this.Worksheets.count)
    Set s = ActiveSheet
    
    On Error GoTo err
    s.name = sname
    
    s.Cells(1, 1).Value = commentA1
    s.range(s.Cells(s_srow, s_scol), s.Cells(s_srow + UBound(db) - 1, s_scol)).Value = WorksheetFunction.Transpose(WorksheetFunction.index(WorksheetFunction.Transpose(db), 2))
    s.Cells.EntireColumn.ColumnWidth = col2width
    s.Columns(s_scol).ColumnWidth = col1width
    s.Rows(1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    makeStudentSheet = sname
    Exit Function
    
err:
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    MsgBox _
        "シート名が重複しています。" & vbLf & _
        "シート名：" & sname & vbLf & vbLf & _
        "生徒氏名または重複するシート名を変更して、再度お試しください。", vbInformation
    sname = ""
    makeStudentSheet = sname
    Exit Function
End Function

Public Function getStudentdb(name As Variant) As Variant
    Dim sheet As Worksheet
    Set sheet = this.Worksheets(name)
    
    Dim erow As Long
    Dim ecol As Long
    
    erow = sheet.Cells(1, 1).End(xlDown).Row
    ecol = sheet.Cells(1, Columns.count).End(xlToLeft).Column
    
    Dim db As Variant
    db = sheet.range(sheet.Cells(1, 1), sheet.Cells(erow, ecol + 1))
    getStudentdb = db
End Function

Function writeStudentdb(db As Variant, name As Variant)
    Dim sheet As Worksheet
    Set sheet = this.Worksheets(name)
    
    sheet.range(sheet.Cells(1, 1), sheet.Cells(UBound(db), UBound(db, 2))) = db
End Function
