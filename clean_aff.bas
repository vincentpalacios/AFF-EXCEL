Attribute VB_Name = "clean_aff"
Sub clean_aff_detail_xls()
Attribute clean_aff_detail_xls.VB_ProcData.VB_Invoke_Func = " \n14"
'
' clean_aff_detail_xls
' version 1.0
' Author: Vincent Palacios
' Date: 04/25/18

' Overview:
'   Unwrap cells
'   Unmerge cells
'   Autofit rows
'   Set column width
'   Delete empty columns
'   Set data to general format
'   Remove +/- symobls from data
'   Convert numbers stored as text to numeric data

    Range("A1").Select
    
    'Credit to: https://www.mrexcel.com/forum/excel-questions/608080-vba-code-select-all-data-worksheet.html
    On Error Resume Next
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    mylastcelladd = Cells(mylastcell.row, mylastcell.Column).Address
    myRange = "A1:" & mylastcelladd
    Range(myRange).Select
    
    With Selection
        .WrapText = False
        .MergeCells = False
    End With
    Selection.Rows.AutoFit
    Selection.ColumnWidth = 10.71
    Range("B:B,C:C,E:E,F:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.ColumnWidth = 30
    
    
    'Find first numeric cell in column B
    Dim i As Long, lastrow As Long, lngVal As Long

    lastrow = Cells(Rows.Count, "B").End(xlUp).row

    For i = 9 To lastrow Step 1
        If IsNumeric(Cells(i, "B").Value) Then
            myfirstcell = Cells(i, "B").Address
            Exit For
        End If
    Next i
        
    Range(myfirstcell).Select
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    mylastcelladd = Cells(mylastcell.row, mylastcell.Column).Address
    myRange = myfirstcell & ":" & mylastcelladd
    Range(myRange).Select
    
    'http://www.ozgrid.com/forum/forum/help-forums/excel-general/55398-converting-numbers-stored-as-text-to-numbers-via-macro?t=64027
    With Selection
        .NumberFormat = "General"
        .Value = .Value
        .HorizontalAlignment = xlRight
        .Replace What:="+/-", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
    End With
    
End Sub

Sub clean_aff_detail_transpose_xls()
'
' clean_aff_detail_transpose_xls
' version 1.0
' Author: Vincent Palacios
' Date: 04/25/18

' Overview:
'   Unwrap cells
'   Unmerge cells
'   Autofit rows
'   Set column width
'   Delete empty columns
'   Set data to general format
'   Remove +/- symobls from data
'   Convert numbers stored as text to numeric data

    Range("A1").Select
    
    'Credit to: https://www.mrexcel.com/forum/excel-questions/608080-vba-code-select-all-data-worksheet.html
    On Error Resume Next
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    mylastcelladd = Cells(mylastcell.row, mylastcell.Column).Address
    myRange = "A1:" & mylastcelladd
    Range(myRange).Select
    
    With Selection
        .WrapText = False
        .MergeCells = False
    End With
    Selection.Rows.AutoFit
    Selection.ColumnWidth = 10.71
    Range("B:B,D:D,F:F,G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.ColumnWidth = 30
    
    
    'Find first numeric cell in column C
    Dim i As Long, lastrow As Long, lngVal As Long

    lastrow = Cells(Rows.Count, "C").End(xlUp).row

    For i = 9 To lastrow Step 1
        If IsNumeric(Cells(i, "C").Value) Then
            myfirstcell = Cells(i, "C").Address
            Exit For
        End If
    Next i
        
    Range(myfirstcell).Select
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    mylastcelladd = Cells(mylastcell.row, mylastcell.Column).Address
    myRange = myfirstcell & ":" & mylastcelladd
    Range(myRange).Select
    
    'http://www.ozgrid.com/forum/forum/help-forums/excel-general/55398-converting-numbers-stored-as-text-to-numbers-via-macro?t=64027
    With Selection
        .NumberFormat = "General"
        .Value = .Value
        .HorizontalAlignment = xlRight
        .Replace What:="+/-", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
    End With
    
End Sub
