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
    sn = Application.InputBox("開始番号")
    If sn = 0 Or sn = False Then
        MsgBox "キャンセルしました。"
        Exit Function
    End If
    
    en = Application.InputBox("終了番号")
    If sn = 0 Or sn = False Then
        MsgBox "キャンセルしました。"
        Exit Function
    End If
    On Error GoTo 0
    
    If en < sn Then
        MsgBox _
            "終了番号として、開始番号より小さい数字は入力できません。" & vbLf & _
            "開始番号：" & sn & vbLf & "終了番号：" & en & vbLf & vbLf & _
            "テスト範囲を確認して再度お試しください。", vbInformation
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
            "テスト範囲の単語数は、既定の問題数（" & numQ & "問）以上である必要があります。" & vbLf & _
            "開始番号：" & sn & vbLf & "終了番号：" & en & vbLf & "単語数：" & count & vbLf & vbLf & _
            "テスト範囲を変更して再度お試しください。", vbInformation
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
        "エラー番号：" & err.number & vbLf & _
        "エラー内容：" & err.Description & vbLf & vbLf & _
        "整数を入力してください。", vbCritical
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
        MsgBox "テストシートがありません。", vbInformation
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
