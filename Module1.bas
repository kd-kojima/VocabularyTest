Attribute VB_Name = "Module1"
' arguments

Option Explicit

Public numQ As Integer

Public temp_sname As String
Public temp2in1_sname As String
Public db_sname As String
Public top_sname As String

Public temp As Worksheet
Public temp2in1 As Worksheet
Public db As Worksheet
Public top As Worksheet

Public q_srow As Integer
Public q_scol As Integer
Public cover_row As Integer
Public cover_col As Integer
Public last_row As Integer
Public last_col As Integer



Public Function init()
    
    ' -----------------------------------------
    ' (Edit here when you change sheet's name.)
    temp_sname = "T"
    temp2in1_sname = "T2"
    db_sname = "db"
    top_sname = "Top"
    
    
    ' -----------------------------------------
    ' worksheets
    Set temp = Worksheets(temp_sname)
    Set temp2in1 = Worksheets(temp2in1_sname)
    Set db = Worksheets(db_sname)
    Set top = Worksheets(top_sname)
    
    
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
    
End Function

