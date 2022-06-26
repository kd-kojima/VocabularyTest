Attribute VB_Name = "Module4"
' functions

Option Explicit

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
    
    temp.Copy after:=Worksheets(Worksheets.count)
    
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
