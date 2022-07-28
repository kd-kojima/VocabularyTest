Attribute VB_Name = "Module2"
' main

Option Explicit

Sub main()
    Call init
    Call makemain(2)
End Sub

Sub spell()
    Call init
    Call makemain(3)
End Sub

'
' type_
'   2: en->ja test
'   3: ja->en test
'
Function makemain(type_ As Integer)
    Dim i As Integer
    
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
    
    ' get db and set/shuffle index
    Dim db As Variant
    Dim index() As Integer
    ReDim index(1) As Integer
    Dim count As Integer
    
    count = 0
    db = getAlldb
    
    For i = 1 To UBound(db)
        If sn <= db(i, 1) And db(i, 1) <= en Then
            If count > 1 Then
                ReDim Preserve index(count)
            End If
            index(count) = i
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
        Qdb(i, 1) = db(index(i - 1), type_)
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

Sub create_2in1()
    Call init
    Dim i As Integer
    Dim j As Integer
    
    Dim sIndex() As Integer
    Dim regName As Object
    Set regName = CreateObject("VBScript.RegExp")
    regName.Pattern = "^\d{8}_\d{6}"
    
    j = 1
    For i = 1 To this.Worksheets.count
        If regName.test(this.Worksheets(i).name) Then
            ReDim Preserve sIndex(j)
            sIndex(j) = i
            j = j + 1
        End If
    Next
    
    If j = 1 Then
        MsgBox "�e�X�g�V�[�g������܂���B", vbInformation
        Exit Sub
    End If
    
    Dim s_name As String
    Dim tSheet As Worksheet
    s_name = Format(Now, "yyyymmdd_hhmmss")
    temp2in1.Copy after:=this.Worksheets(this.Worksheets.count)
    ActiveSheet.name = "2in1_" + s_name
    Set tSheet = ActiveSheet
    
    Dim sr As Integer
    Dim sc As Integer
    Dim thisSheet As Worksheet
    
    For i = 1 To UBound(sIndex)
        sr = (Int((i - 1) / 2)) * last_row + 1
        sc = ((i - 1) Mod 2) * last_col + 1
        Set thisSheet = this.Worksheets(sIndex(i))
        thisSheet.range(thisSheet.Cells(1, 1), thisSheet.Cells(last_row, last_col)).Copy tSheet.range(tSheet.Cells(sr, sc), tSheet.Cells(sr + last_row - 1, sc + last_col - 1))
    Next
End Sub
