Attribute VB_Name = "Module1"
' variables

Option Explicit

Public this As Workbook

Public temp_sname As String
Public temp2in1_sname As String
Public db_sname As String
Public top_sname As String

Public temp As Worksheet
Public temp2in1 As Worksheet
Public db As Worksheet
Public top As Worksheet

Public numQ As Integer
Public q_srow As Integer
Public q_scol As Integer
Public cover_row As Integer
Public cover_col As Integer
Public last_row As Integer
Public last_col As Integer

Public commentA1 As String
Public s_srow As Integer
Public s_scol As Integer
Public col1width As Double
Public col2width As Double

Public c_col As Integer
Public f_col As Integer


Public Function init()

    ' -----------------------------------------
    ' workbook
    Set this = ThisWorkbook
    
    
    
    ' -----------------------------------------
    ' sheet names
    ' (Edit here when you change sheet's name.)
    temp_sname = "T"
    temp2in1_sname = "T2"
    db_sname = "db"
    top_sname = "Top"
    
    
    
    ' -----------------------------------------
    ' worksheets
    Set temp = this.Worksheets(temp_sname)
    Set temp2in1 = this.Worksheets(temp2in1_sname)
    Set db = this.Worksheets(db_sname)
    Set top = this.Worksheets(top_sname)
    
    
    
    ' -----------------------------------------
    ' template sheet
    ' (Edit here when you change template sheet.)
    
    ' number of questions
    numQ = 20
    
    ' start row/col of test word
    q_srow = 2
    q_scol = 3
    
    ' row/col of displaying where test covers
    cover_row = 1
    cover_col = 5
    
    ' end row/col of sheet
    last_row = 21
    last_col = 6
    
    
    
    ' -----------------------------------------
    ' student sheet
    
    ' message displayed A1
    commentA1 = ">=1: correct, <=0: fail"
    
    ' words start row/col
    s_srow = 2
    s_scol = 1
    
    ' column width
    col1width = 18
    col2width = 3
    

    
    ' -----------------------------------------
    ' top sheet
    
    ' column of correct/fail
    c_col = 4
    f_col = 5
    
    
End Function

