Attribute VB_Name = "Module1"
Sub format_cells()
Attribute format_cells.VB_Description = "to format cells to center"
Attribute format_cells.VB_ProcData.VB_Invoke_Func = " \n14"
'
' format_cells Macro
' to format cells to center
'

'
    ActiveCell.FormulaR1C1 = "A"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "B"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "C"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "D"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "E"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "F"
    Range("A1:A6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ChDir "E:\work\excel"
    ActiveWorkbook.SaveAs Filename:="E:\work\excel\format_cells.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub
Sub summary()
Attribute summary.VB_Description = "to find summary of the series\n"
Attribute summary.VB_ProcData.VB_Invoke_Func = " \n14"
'
' summary Macro
' to find summary of the series
'

'
    ActiveWindow.SmallScroll Down:=18
    Range("C33").Select
    ActiveCell.FormulaR1C1 = "=mean(R[-31]C:R[-2]C)"
    Range("C33").Select
    Selection.ClearContents
    Range("C33").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-31]C:R[-2]C)"
    Range("B33").Select
    ActiveCell.FormulaR1C1 = "average"
    Range("B34").Select
    ActiveCell.FormulaR1C1 = "median"
    Range("C34").Select
    ActiveCell.FormulaR1C1 = "=medi"
    Range("C34").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=media"
    Range("C34").Select
    ActiveCell.FormulaR1C1 = "=MEDIAN(R[-32]C:R[-3]C)"
    Range("B35").Select
    ActiveCell.FormulaR1C1 = "mode"
    Range("C35").Select
    ActiveCell.FormulaR1C1 = "=MODE(R[-33]C:R[-4]C)"
    Range("C35").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=MODE(R[-33]C:R[-4]C)"
    Range("C35").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=MODE(R[-33]C:R[-4]C)"
    Range("B37").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B37").Select
    ActiveCell.FormulaR1C1 = "kurtosis"
    Range("C37").Select
    ActiveCell.FormulaR1C1 = "=KURT(R[-35]C:R[-6]C)"
    Range("B38").Select
    ActiveCell.FormulaR1C1 = "skewness"
    Range("C38").Select
    ActiveCell.FormulaR1C1 = "=SKEW(R[-36]C:R[-7]C)"
    Range("C33").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("D33").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("D28").Select
End Sub

Sub active_cell()
Range("A1:B3").Select


End Sub

Sub range_offset()
Range("A1:B3").Offset(3, 2).Select

End Sub

Sub Resize()
Range("B6").Resize(2, 2).Select

End Sub

Sub variable_ex()
Dim my_number As Integer
Dim my_other_number As Double
my_number = 10
my_other_number = my_number + 0.5
Range("A11:A20").Value = my_number
Range("B11:B20").Value = my_other_number
End Sub

'action explicit & general declations
Option Explicit
Sub explicit_ex()
Dim my_num As Double
my_num = 1
my_nums = 2
Range("A1").Value = my_num
Range("B2").Value = my_nums


End Sub

'arithmetic operations
Sub arith()
Dim num_1 As Integer
Dim num_2 As Integer
Dim num_3 As Integer

num_1 = 3
num_2 = 2
num_3 = 1

Ans1 = num_1 + num_2
Ans2 = num_1 - num_2
Ans3 = num_1 + num_2 - num_3
Ans4 = (num_1 + num_2) * num_3
Ans5 = (num_1 + num_2 + num_3) / 3

Worksheets(7).Range("A1").Value = "simple addition"
Worksheets(7).Range("B1").Value = Ans1
Worksheets(7).Range("A2").Value = "simple subtraction"
Worksheets(7).Range("B2").Value = Ans2
Worksheets(7).Range("A3").Value = "subtraction from addition"
Worksheets(7).Range("B3").Value = Ans3
Worksheets(7).Range("A4").Value = "multiplication"
Worksheets(7).Range("B4").Value = Ans4
Worksheets(7).Range("A5").Value = "simple addition"
Worksheets(7).Range("B5").Value = Ans5

End Sub

'number of rows
Sub num_of_rows_()
Dim num_of_rows As Long
num_of_rows = Worksheets(1).Rows.Count
MsgBox num_of_rows
End Sub

'how to use function?
Sub work_sheet_fun()
Dim work_sheet_function1 As Single
Dim work_sheet_fucntion2 As Double
Dim work_sheet_fucntion3 As Variant
Dim work_sheet_fucntion4 As String
work_sheet_function1 = WorksheetFunction.Pi
work_sheet_function2 = WorksheetFunction.Pi
work_sheet_function3 = "this is only a varient"
work_sheet_function4 = "You know! this is string"
Range("a1").Value = work_sheet_function1
Range("a2").Value = work_sheet_function2
Range("a3").Value = work_sheet_function3
Range("a4").Value = work_sheet_function4
End Sub
