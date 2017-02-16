Attribute VB_Name = "FSFile"
Sub LastMonth()
Attribute LastMonth.VB_ProcData.VB_Invoke_Func = "n\n14"

    With ActiveCell
        .Formula = _
            "=""=VLOOKUP("" & RC2 & "","" & R5C17 & ""!$B:$O,14,FALSE)"" "
        .Value = .Value
        .Value = .Value
        .Offset(1, 0).Select
    End With
    
End Sub
