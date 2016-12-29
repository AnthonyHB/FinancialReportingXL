Attribute VB_Name = "WF"
Sub clearRow()
Attribute clearRow.VB_ProcData.VB_Invoke_Func = "n\n14"
' Copy all the Date/Amounts, move them one cell down so that they match with the Description.
' Keyboard Shortcut: Option+Cmd+n
'
If Cells(ActiveCell.Row, 1) = Empty Then
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Delete Shift:=xlUp
    ActiveCell.Offset(1, 0).Range("A1").Select
Else
    ActiveCell.Offset(1, 0).Range("A1").Select
End If
End Sub

