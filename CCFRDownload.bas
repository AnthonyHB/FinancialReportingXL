Attribute VB_Name = "CCFRDownload"
Sub FormatDownload()
'
' FormatDownload Macro
'
    Application.ScreenUpdating = False
    
' Application of Origin Delete
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    
' Amount Column
    Range("F1").Select
    ActiveCell.Value = "Amount"
    Range("F2").Select
    
    Do While Cells(ActiveCell.Row, 1) <> Empty
        If Cells(ActiveCell.Row, 1) <> Empty Then
            ActiveCell.Formula = "=" & Cells(ActiveCell.Row, ActiveCell.Column - 2) & "+" & Cells(ActiveCell.Row, ActiveCell.Column - 1)
            ActiveCell.Value = ActiveCell.Value
            Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
        Else
            Exit Do
        End If
    Loop
    
' Delete Debits/Credits
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").EntireColumn.AutoFit
    Columns("D:D").Style = "Comma"
    Range("A1").Select
    
' SUMIF for all accounts
    Range("F4").Select
        ActiveCell.Value = "'  4128-1099.0000"
        ActiveCell.Offset(0, 1).FormulaR1C1 = "=SUMIF(C1,RC[-1],C[-3])"
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4128-1205.0000"
        ActiveCell.Offset(0, 1).FillDown
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4135-1099.0000"
        ActiveCell.Offset(0, 1).FormulaR1C1 = "=SUMIF(C1,RC[-1],C[-3])"
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4135-1205.0000"
        ActiveCell.Offset(0, 1).FillDown
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4234-1099.0000"
        ActiveCell.Offset(0, 1).FormulaR1C1 = "=SUMIF(C1,RC[-1],C[-3])"
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4234-1205.0000"
        ActiveCell.Offset(0, 1).FillDown
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4236-1099.0000"
        ActiveCell.Offset(0, 1).FormulaR1C1 = "=SUMIF(C1,RC[-1],C[-3])"
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4236-1205.0000"
        ActiveCell.Offset(0, 1).FillDown
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4338-1099.0000"
        ActiveCell.Offset(0, 1).FormulaR1C1 = "=SUMIF(C1,RC[-1],C[-3])"
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4338-1205.0000"
        ActiveCell.Offset(0, 1).FillDown
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4350-1099.0000"
        ActiveCell.Offset(0, 1).FormulaR1C1 = "=SUMIF(C1,RC[-1],C[-3])"
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4350-1205.0000"
        ActiveCell.Offset(0, 1).FillDown
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4369-1099.0000"
        ActiveCell.Offset(0, 1).FormulaR1C1 = "=SUMIF(C1,RC[-1],C[-3])"
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "'  4369-1205.0000"
        ActiveCell.Offset(0, 1).FillDown
        ActiveCell.Offset(1, 0).Select
        Columns("F:F").EntireColumn.AutoFit
        Range("G4:G17").Style = "Comma"
        Columns("G:G").EntireColumn.AutoFit

    Application.ScreenUpdating = True

End Sub

Sub FilterByDate()
'
' FilterDate Macro
'

'
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$D$100000").AutoFilter Field:=2, Criteria1:="=11/*" _
        , Operator:=xlAnd, Criteria2:="=*/2016"
End Sub
