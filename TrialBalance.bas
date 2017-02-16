Attribute VB_Name = "TrialBalance"
Sub TrialBalance1()
Attribute TrialBalance1.VB_ProcData.VB_Invoke_Func = "t\n14"

' TrialBalance Macro

' Will malfunction if new accounts are introduced. This will take you all the way to the point
' to where you need to Paste Values for the Equity lines. The macro takes too long to calculate
' the SUMPRODUCT before giving you the value to copy, so it needs to be done separately.

Application.ScreenUpdating = False

' Setup workbook
    Rows("1:1").Select
    Range("B1").Activate
    Selection.Insert Shift:=xlDown
    Selection.Insert Shift:=xlDown
    Columns("C:C").Select
    Range("C4").Activate
    Selection.Insert Shift:=xlToRight
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Type"
    Range("D4").Select
    ActiveWindow.FreezePanes = True
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "Balance Sheet"
    Columns("I:I").EntireColumn.AutoFit
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "Income Statement"
    Columns("J:J").EntireColumn.AutoFit
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "Equity"
    Range("L3").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("I4").Select
    Range("G3").Select
    Selection.Copy
    Rows("3:3").Select
    Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    
' B/S, IS, and Total Formulas
    Range("C4:C2073") = "=VLOOKUP(A:A,'Accounting:Books and records:Mieneke:Financial Statements:2017:January 2017:[1.31.2017 Financial Statements.xlsx]GL Account Classification'!$A:$B,2,FALSE)"
    Range("I4:I2073") = "=SUM(IF(RC[-6]=""B/S"",RC[-4]:RC[-3]))"
    Range("J4:J2073") = "=SUMIF(RC[-7],""IS"",RC[-4])"
    Range("L4:L2073") = "=SUM(RC[-3]:RC[-1])"
    
' Retained Earnings
    'Loops through all cells, finds Retained Earnings, deletes B/S, adds formula to Col N
    
    Range("B3").Select
 
    Do While ActiveCell.Value <> Empty
        If InStr(1, ActiveCell, "Retained earnings") Then
            Cells(ActiveCell.Row, ActiveCell.Column + 7).Clear
            Cells(ActiveCell.Row, ActiveCell.Column + 12).Formula = "=SUMPRODUCT((LEFT(C1,6)=LEFT(RC[-13],6))+0,C12)"
            Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
        ElseIf ActiveCell.Value <> Empty Then
            Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
        Else
            Application.ScreenUpdating = True
        End If
    Loop

End Sub


Sub PasteEquity()
Attribute PasteEquity.VB_ProcData.VB_Invoke_Func = "p\n14"

' PasteEquity Macro
' DO NOT HAVE FILTERS ON. This will use lots of CMD+Down's starting from the N Column
' N Column should have the SUMPRODUCT formula pasted from the TrialBalance macro

    Application.ScreenUpdating = False
    Range("N3").Select
    
    Do While Cells(ActiveCell.Row, 1) <> Empty
        If Cells(ActiveCell.Row, 1) <> Empty Then
            'Paste value to the right
            Selection.End(xlDown).Select
            Selection.Copy
            Cells(ActiveCell.Row, ActiveCell.Column + 1).Select
            Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
            Application.CutCopyMode = False
            'Paste negative to Equity
            Cells(ActiveCell.Row, ActiveCell.Column - 4).Formula = "=-Round(RC[4],2)"
            Cells(ActiveCell.Row, ActiveCell.Column - 1).Select
        ElseIf ActiveCell.Row > 3000 Then
            Range("N3").Select
            Application.ScreenUpdating = True
            Exit Do
        End If
    Loop
    
End Sub
