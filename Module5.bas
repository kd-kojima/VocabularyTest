Attribute VB_Name = "Module5"
' setup

Option Explicit

Dim i As Integer
Dim j As Integer
Dim numQ As Integer

Dim this As Workbook
    
Dim setup_sn As String
Dim top_sn As String
Dim db_sn As String
Dim t_sn As String
Dim t2_sn As String
    
Dim setup_s As Worksheet
Dim top_s As Worksheet
Dim db_s As Worksheet
Dim t_s As Worksheet
Dim t2_s As Worksheet

Dim top_rows As Integer
Dim top_cols As Integer
Dim top_row_h() As Double
Dim top_col_w() As Double

Dim db_rows As Integer
Dim db_cols As Integer
Dim db_row_h() As Double
Dim db_col_w() As Double

Dim t_rows As Integer
Dim t_cols As Integer
Dim t_row_h() As Double
Dim t_col_w() As Double
Dim t_numcol As Integer
Dim t_anscol As Integer
Dim t_titlerow As Integer
Dim t_titlecol As Integer
Dim t_stestrow As Integer
Dim t_etestrow As Integer
Dim t_fontsize As Integer

Dim t2_rows As Integer
Dim t2_cols As Integer
Dim t2_row_h() As Double
Dim t2_col_w() As Double

Sub setup()
    numQ = 20
    
    Set this = ThisWorkbook
    
    setup_sn = "setup"
    t_sn = "T"
    t2_sn = "T2"
    top_sn = "Top"
    db_sn = "db"
    
    ' ---------------------------------------------
    ' temp sheet variables
    
    t_rows = 21
    t_cols = 6
    ReDim t_row_h(1 To t_rows) As Double
    For i = LBound(t_row_h) To UBound(t_row_h)
        t_row_h(i) = 33.6
    Next
    ReDim t_col_w(1 To t_cols) As Double
    t_col_w(1) = 7.8
    t_col_w(2) = 1.1
    t_col_w(3) = 19.8
    t_col_w(4) = 2
    t_col_w(5) = 28.6
    t_col_w(6) = 5.9
    
    t_numcol = 1
    t_anscol = 5
    t_titlerow = 1
    t_titlecol = 1
    
    t_stestrow = 2
    t_etestrow = t_stestrow + numQ - 1
    
    t_fontsize = 14
    
    ' ---------------------------------------------
    
    If isNotExistSheet(t_sn) Then
        Call makeTempsheet
    End If
    
End Sub

Function makeTempsheet()
    this.Worksheets.Add after:=this.Worksheets(this.Worksheets.count)
    ActiveSheet.name = t_sn
    Set t_s = this.Worksheets(t_sn)
    
    Call setWH(t_s, t_col_w, t_row_h)
    
    Dim title As String
    title = Application.InputBox("テストのタイトルを入力してください。" & vbLf & "（テスト用紙の左上に記載されます。）" & vbLf & vbLf & "default: 単語テスト")
    If title = "" Then title = "単語テスト"
    
    t_s.Cells(t_titlerow, t_titlecol).Value = title
    
    For i = t_stestrow To t_etestrow
        t_s.Cells(i, t_numcol).Value = i - 1
        t_s.Cells(i, t_anscol).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Next
    
    With t_s.PageSetup
        .TopMargin = Application.CentimetersToPoints(1.8)
        .BottomMargin = Application.CentimetersToPoints(1.8)
        .LeftMargin = Application.CentimetersToPoints(2.6)
        .RightMargin = Application.CentimetersToPoints(2.6)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        .CenterHorizontally = True
        .CenterVertically = True
    End With
    
    t_s.range(Cells(1, 1), Cells(t_rows, t_cols)).Font.Size = t_fontsize
    
    Dim br As range
    Set br = t_s.range("H2")
    With t_s.Buttons.Add(br.Left, br.top, br.Width * 3, br.Height * 2)
        .OnAction = "create_2in1"
        .Characters.Text = "2in1"
        .Characters.Font.Size = 14
    End With
    
End Function

Function isNotExistSheet(sname As String) As Boolean
    For i = 1 To this.Worksheets.count
        If this.Worksheets(i).name = sname Then
            isNotExistSheet = False
            Exit Function
        End If
    Next
    
    isNotExistSheet = True
End Function

Function setW(s As Worksheet, w() As Double)
    For i = LBound(w) To UBound(w)
        s.Columns(i).ColumnWidth = w(i)
    Next
End Function

Function setH(s As Worksheet, h() As Double)
    For i = LBound(h) To UBound(h)
        s.Rows(i).RowHeight = h(i)
    Next
End Function

Function setWH(s As Worksheet, w() As Double, h() As Double)
    For i = LBound(w) To UBound(w)
        s.Columns(i).ColumnWidth = w(i)
    Next
    For i = LBound(h) To UBound(h)
        s.Rows(i).RowHeight = h(i)
    Next
End Function
