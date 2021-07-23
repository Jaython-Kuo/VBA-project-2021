Attribute VB_Name = "Module1"
Sub 診所_疫苗數量()
Attribute 診所_疫苗數量.VB_Description = "遞增並計算總和平均"
Attribute 診所_疫苗數量.VB_ProcData.VB_Invoke_Func = "z\n14"
'
' 診所_疫苗數量 Macro
' 遞增並計算總和平均
'
' 快速鍵: Ctrl+z
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
    ActiveCell.FormulaR1C1 = "=aver"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[1]C[-5]:R[413]C[-5])"
    Range("G2").Select
End Sub
