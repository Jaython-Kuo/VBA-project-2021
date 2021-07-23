Attribute VB_Name = "Module1"
Option Explicit

Sub 口罩特約診所()
Attribute 口罩特約診所.VB_Description = "口罩特殊藥局 口罩篩選  計算總合"
Attribute 口罩特約診所.VB_ProcData.VB_Invoke_Func = "z\n14"
'
' 口罩特約診所 Macro
' 口罩特殊藥局 口罩篩選  計算總合
'
' 快速鍵: Ctrl+z
'
    Columns("A:B").Select
    Selection.AutoFilter
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=sum"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("E2").Select
    ActiveWindow.SmallScroll Down:=-42
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]/413"
    Range("G1").Select

    Selection.NumberFormatLocal = "0_ "
End Sub
Sub 口罩特約藥局()
Attribute 口罩特約藥局.VB_Description = "口罩數量排序  小到大  計算總合及平均數"
Attribute 口罩特約藥局.VB_ProcData.VB_Invoke_Func = "x\n14"
'
' 口罩特約藥局 Macro
' 口罩數量排序  小到大  計算總合及平均數
'
' 快速鍵: Ctrl+x
'
    Range("B2").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[413]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("G2").Select
End Sub
