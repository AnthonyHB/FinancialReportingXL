Attribute VB_Name = "B34"
Sub clearRow()
Attribute clearRow.VB_ProcData.VB_Invoke_Func = "n\n14"
' Copy all the Date/Amounts, move them one cell down so that they match with the Description.
' Keyboard Shortcut: Option+Cmd+n
'
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub

