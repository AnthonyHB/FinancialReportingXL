Sub Macro2()


Dim i As Integer
Dim targetCells As Range
Dim cell As Range
Dim referenceRange As Range
Dim thisSheet As Worksheet

Set referenceRange = ActiveSheet.Range("CA1")

With referenceRange
    For Each thisSheet In ThisWorkbook.Sheets
        If thisSheet.Index >= referenceRange.Parent.Index Then
            Set targetCells = thisSheet.Cells.SpecialCells(xlCellTypeFormulas, 23)
            For Each cell In targetCells
                If cell.HasFormula Then
                    .Offset(i, 0).Value = thisSheet.Name
                    .Offset(i, 1).Value = cell.Address
                    .Offset(i, 2).Value = CStr(cell.Formula)
                    i = i + 1
                End If
            Next
        End If
    Next
End With

End Sub