Attribute VB_Name = "Module3"
' students

Option Explicit

Sub mainStudentCorrect()
    Call init
    Call makemainStudent(2, 0)
End Sub

Sub spellStudentCorrect()
    Call init
    Call makemainStudent(3, 0)
End Sub

Sub mainStudentFailYet()
    Call init
    Call makemainStudent(2, 1)
End Sub

Sub spellStudentFailYet()
    Call init
    Call makemainStudent(3, 1)
End Sub

'
' type_
'   2: en->ja test
'   3: ja->en test
'
' mode
'   0: correct word test (covers ">= threshold")
'   1: failed/notyet word test (covers "<= threshold" and "not yet tested")
'
Function makemainStudent(type_ As Integer, mode As Integer)
    Dim i As Integer
    
    Dim student As Variant
    student = whoIsStudent
    If student(0) = 0 Then Exit Function
    
    Dim db As Variant
    db = getStudentdb(student(1))
    If UBound(db, 2) < 3 Then
        MsgBox "���k�̉𓚌��ʂ��o�^����Ă��܂���B"
        Exit Function
    End If
    
    ' get where test covers
    Dim sn As Integer
    Dim en As Integer
    
    On Error GoTo nullerr
    sn = Application.InputBox("�J�n�ԍ�")
    If sn = 0 Or sn = False Then
        MsgBox "�L�����Z�����܂����B"
        Exit Function
    End If
    
    en = Application.InputBox("�I���ԍ�")
    If sn = 0 Or sn = False Then
        MsgBox "�L�����Z�����܂����B"
        Exit Function
    End If
    On Error GoTo 0
    
    If en < sn Then
        MsgBox _
            "�I���ԍ��Ƃ��āA�J�n�ԍ���菬���������͓��͂ł��܂���B" & vbLf & _
            "�J�n�ԍ��F" & sn & vbLf & "�I���ԍ��F" & en & vbLf & vbLf & _
            "�e�X�g�͈͂��m�F���čēx���������������B", vbInformation
        Exit Function
    End If
    
    ' get threshold
    Dim threshold As Variant
    
    On Error GoTo nullerr
    If mode = 0 Then
        threshold = Application.InputBox("���͂��ꂽ�l�ȏ�̐��𐔂̒P����A�e�X�g�͈͂Ƃ��܂��B")
    Else
        threshold = Application.InputBox("���͂��ꂽ�l�ȉ��̐��𐔂̒P��ƁA���o��̒P����A�e�X�g�͈͂Ƃ��܂��B")
    End If
    On Error GoTo 0
    
    If threshold = "" Or threshold = "False" Then
        MsgBox "�L�����Z�����܂����B"
        Exit Function
    End If
    
    threshold = CInt(threshold)
    
    ' convert sn/en as serial number (for if sn/en is page number etc)
    Dim alldb As Variant
    
    Dim isn As Integer
    Dim ien As Integer
    
    alldb = getAlldb
    
    isn = 0
    ien = 0
    For i = 1 To UBound(alldb)
        If alldb(i, 1) < sn Then
            isn = i
        ElseIf alldb(i, 1) <= en Then
            ien = i
        End If
    Next
    
    isn = isn + 1
    ien = ien
    
    ' set index of test words
    Dim index() As Integer
    ReDim index(1) As Integer
    Dim count As Integer
    
    count = 0
    For i = isn + 1 To ien + 1
        If _
        (mode = 0 And db(i, UBound(db, 2) - 1) <> "" And db(i, UBound(db, 2) - 1) >= threshold) _
        Or (mode = 1 And (db(i, UBound(db, 2) - 1) <= threshold Or db(i, UBound(db, 2) - 1) = "")) Then
            If count > 1 Then
                ReDim Preserve index(count)
            End If
            index(count) = i - 1
            count = count + 1
        End If
    Next
    
    If count < numQ Then
        MsgBox _
            "�e�X�g�͈͂̒P�ꐔ�́A����̖�萔�i" & numQ & "��j�ȏ�ł���K�v������܂��B" & vbLf & _
            "�J�n�ԍ��F" & sn & vbLf & "�I���ԍ��F" & en & vbLf & "�P�ꐔ�F" & count & vbLf & vbLf & _
            "�e�X�g�͈͂�ύX���čēx���������������B", vbInformation
        Exit Function
    End If
    
    index = shuffleArray(index)
    
    ' make questions db
    Dim Qdb() As String
    ReDim Qdb(1 To numQ, 1 To 1) As String
    
    For i = 1 To numQ
        Qdb(i, 1) = alldb(index(i - 1), type_)
    Next
    
    ' make test sheet
    Call makeTestSheet(Qdb, sn, en)
    Exit Function
    
nullerr:
    MsgBox _
        "�G���[�ԍ��F" & err.number & vbLf & _
        "�G���[���e�F" & err.Description & vbLf & vbLf & _
        "��������͂��Ă��������B", vbCritical
    Exit Function
End Function


Sub addStudent()
    Call init
    
    Dim name As String
    name = Application.InputBox("���k����")
    If name = "False" Then
        MsgBox "�L�����Z�����܂����B"
        Exit Sub
    End If
    
    Dim erow As Long
    Dim students As Variant
    
    erow = top.Cells(top.Rows.count, 2).End(xlUp).Row
    
    If erow > 1 Then
    
        students = top.range(top.Cells(1, 2), top.Cells(erow, 2))
    
        Dim i As Integer
        For i = 1 To UBound(students)
            If students(i, 1) = name Then
                MsgBox name & " �͊��ɑ��݂��܂��B"
                Exit Sub
            End If
        Next
    End If
    
    Dim res As String
    res = makeStudentSheet(name)
    If res = "" Then Exit Sub
    
    top.Cells(erow + 1, 2).Value = name
End Sub

Sub delStudent()
    Call init
    Dim student As Variant
    
    student = whoIsStudent
    If student(0) = 0 Then Exit Sub
    
    If MsgBox(student(1) & " �̃f�[�^���폜���܂��B", vbOKCancel + vbExclamation) = vbCancel Then
        Exit Sub
    End If
    
    top.Cells(student(0), 1).Value = ""
    top.Cells(student(0), 2).Delete shift:=xlUp
    
    Application.DisplayAlerts = False
    this.Worksheets(student(1)).Delete
    Application.DisplayAlerts = True
End Sub

Sub registResult()
    Call init
    
    Dim student As Variant
    
    student = whoIsStudent
    If student(0) = 0 Then Exit Sub
    
    Dim db As Variant
    db = getStudentdb(student(1))
    
    Dim correct As Variant
    Dim fail As Variant
    Dim ecorr As Long
    Dim efail As Long
    
    ecorr = top.Cells(Rows.count, c_col).End(xlUp).Row
    efail = top.Cells(Rows.count, f_col).End(xlUp).Row
    
    If ecorr = 1 And efail = 1 Then
        MsgBox ("�o�^����P�ꂪ����܂���B")
        Exit Sub
    End If
    
    If MsgBox("�o�^����P�ꐔ�F" & (ecorr + efail - 2), vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim j As Integer
    Dim flag As Boolean
    
    ' duplicate result since last tested
    If UBound(db, 2) > 2 Then
        For i = 2 To UBound(db)
            If db(i, UBound(db, 2) - 1) <> "" Then
                db(i, UBound(db, 2)) = db(i, UBound(db, 2) - 1)
            End If
        Next
    End If
    
    If ecorr > 1 Then
        correct = top.range(top.Cells(1, c_col), top.Cells(ecorr, c_col))
        For i = 2 To UBound(correct)
            flag = False
            For j = 2 To UBound(db)
                If correct(i, 1) = db(j, 1) Then
                    If db(j, UBound(db, 2)) = "" Or db(j, UBound(db, 2)) < 1 Then
                        db(j, UBound(db, 2)) = 1
                    Else
                        db(j, UBound(db, 2)) = db(j, UBound(db, 2)) + 1
                    End If
                    flag = True
                    Exit For
                End If
            Next
            If flag = False Then
                MsgBox "���݂��Ȃ��p�P��F" & correct(i, 1)
                Exit Sub
            End If
        Next
    End If
            
    If efail > 1 Then
        fail = top.range(top.Cells(1, f_col), top.Cells(efail, f_col))
        For i = 2 To UBound(fail)
            flag = False
            For j = 2 To UBound(db)
                If fail(i, 1) = db(j, 1) Then
                    If db(j, UBound(db, 2)) = "" Or db(j, UBound(db, 2)) > 0 Then
                        db(j, UBound(db, 2)) = 0
                    Else
                        db(j, UBound(db, 2)) = db(j, UBound(db, 2)) - 1
                    End If
                    flag = True
                    Exit For
                End If
            Next
            If flag = False Then
                MsgBox "���݂��Ȃ��p�P��F" & fail(i, 1)
                Exit Sub
            End If
        Next
    End If
    
    db(1, UBound(db, 2)) = UBound(db, 2) - 1
    
    Call writeStudentdb(db, student(1))
    
    ' delete correct and fail
    If ecorr > 1 Then
        For i = 2 To UBound(correct)
            correct(i, 1) = ""
        Next
        top.range(top.Cells(1, c_col), top.Cells(ecorr, c_col)) = correct
    End If
    
    If efail > 1 Then
        For i = 2 To UBound(fail)
            fail(i, 1) = ""
        Next
        top.range(top.Cells(1, f_col), top.Cells(efail, f_col)) = fail
    End If
End Sub

Sub resetResult()
    Call init
    
    Dim student As Variant
    
    student = whoIsStudent
    If student(0) = 0 Then Exit Sub
    
    If MsgBox(student(1) & " �̉𓚌��ʂ̃J�E���g�����Z�b�g���܂��B" & vbLf & "�i�ߋ��̃f�[�^�͎c��܂��B�j", vbOKCancel + vbInformation) = vbCancel Then
        Exit Sub
    End If
    
    Dim db As Variant
    db = getStudentdb(student(1))
    
    db(1, UBound(db, 2)) = UBound(db, 2) - 1
    
    Call writeStudentdb(db, student(1))
End Sub

Sub resultViewer()
    Call init
    
    Dim i As Integer
    Dim j As Integer
    
    Dim student As Variant
    student = whoIsStudent
    If student(0) = 0 Then Exit Sub
    
    Dim studentdb As Variant
    Dim withoutheader() As Variant
    
    studentdb = getStudentdb(student(1))
    
    ReDim withoutheader(1 To UBound(studentdb) - 1, 1 To UBound(studentdb, 2) - 1)
    For i = 1 To UBound(withoutheader)
        For j = 1 To UBound(withoutheader, 2)
            withoutheader(i, j) = studentdb(i + 1, j + 1)
        Next
    Next
    
    Dim max_ As Integer
    Dim min_ As Integer
    Dim corr_step As Double
    Dim fail_step As Double
    max_ = WorksheetFunction.Max(withoutheader)
    min_ = WorksheetFunction.Min(withoutheader)
    min_ = min_ * -1 + 1
    
    corr_step = 1 / max_
    fail_step = 1 / min_
    
    Debug.Print (max_)
    Debug.Print (min_)
    Debug.Print (corr_step)
    Debug.Print (fail_step)
    
    
    Dim s As Worksheet
    Dim sname As String
    this.Worksheets.Add after:=this.Worksheets(this.Worksheets.count)
    Set s = ActiveSheet
    sname = student(1) & "_resultViewer"
    
    ' On Error GoTo err
    s.name = sname
    
    Dim color As Integer
    
    For i = 2 To UBound(studentdb)
        For j = 2 To UBound(studentdb, 2) - 1
            If studentdb(i, j) = "" Then
                s.Cells(i, j + 1).Interior.color = RGB(127, 127, 127)
            ElseIf studentdb(i, j) >= 1 Then
                color = Round(127 + corr_step * 127 * studentdb(i, j))
                s.Cells(i, j + 1).Interior.color = RGB(color, color, color)
            Else
                color = Round(255 - fail_step * 255 * (studentdb(i, j) * -1 + 1))
                s.Cells(i, j + 1).Interior.color = RGB(255, color, color)
            End If
        Next
    Next
    
    studentdb(1, 1) = ""
    For i = 2 To UBound(studentdb)
        For j = 2 To UBound(studentdb, 2)
            studentdb(i, j) = ""
        Next
    Next
    s.range(s.Cells(1, 2), s.Cells(UBound(studentdb), UBound(studentdb, 2) + 1)).Value = studentdb
    
    Dim db As Variant
    db = getAlldb
    s.range(s.Cells(2, 1), s.Cells(UBound(db) + 1, 1)).Value = WorksheetFunction.index(db, 0, 1)
    
    s.Cells.EntireColumn.ColumnWidth = col2width
    s.Columns(2).ColumnWidth = col1width
    With s.range(s.Cells(1, UBound(studentdb, 2)), s.Cells(UBound(studentdb), UBound(studentdb, 2)))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
    
    For i = 1 To UBound(db)
        If (i - 1) Mod 5 = 0 Then
            s.range(s.Cells(i, 1), s.Cells(i, UBound(studentdb, 2))).Borders(xlEdgeBottom).LineStyle = xlContinuous
        End If
    Next

    Exit Sub
    
err:
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    MsgBox _
        "�V�[�g�����d�����Ă��܂��B" & vbLf & _
        "�V�[�g���F" & sname & vbLf & vbLf & _
        "�Â� Result Viewer ���폜���āA�ēx���������������B", vbInformation
    Exit Sub
End Sub

