Attribute VB_Name = "Module1"
Option Explicit

Sub �f�n�S���E��()
Attribute �f�n�S���E��.VB_Description = "�f�n�S���ħ� �f�n�z��  �p���`�X"
Attribute �f�n�S���E��.VB_ProcData.VB_Invoke_Func = "z\n14"
'
' �f�n�S���E�� Macro
' �f�n�S���ħ� �f�n�z��  �p���`�X
'
' �ֳt��: Ctrl+z
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
Sub �f�n�S���ħ�()
Attribute �f�n�S���ħ�.VB_Description = "�f�n�ƶq�Ƨ�  �p��j  �p���`�X�Υ�����"
Attribute �f�n�S���ħ�.VB_ProcData.VB_Invoke_Func = "x\n14"
'
' �f�n�S���ħ� Macro
' �f�n�ƶq�Ƨ�  �p��j  �p���`�X�Υ�����
'
' �ֳt��: Ctrl+x
'
    Range("B2").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
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
