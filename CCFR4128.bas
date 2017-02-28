Attribute VB_Name = "CCFR"

Sub DataTransfer4128()

Dim cc4128 As Worksheet
Dim fr4128 As Worksheet
Dim data As Worksheet

Set cc4128 = ThisWorkbook.Sheets("4128CC")
Set fr4128 = ThisWorkbook.Sheets("4128FR")
Set data = ThisWorkbook.Sheets("Data")

cc4128LR = cc4128.ListObjects("CC4128A").TotalsRowRange.Row
fr4128LR = fr4128.ListObjects("FR4128A").TotalsRowRange.Row

dataLR = data.Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0

If myIB1 = Empty Then Exit Sub

y = cc4128LR - 1 ' leaves 1 row before total row
z = fr4128LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4128-1099.0000" Then ' (TK) Fix to populate all stores
            cc4128.Cells(y, 1) = data.Cells(x, 1) 'name
            cc4128.Cells(y, 2) = data.Cells(x, 2) 'date
            cc4128.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc4128.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc4128.ListObjects("CC4128A").ListRows.Add AlwaysInsert:=True
        ElseIf Left(data.Cells(x, 1), 16) = "  4128-1205.0000" Then ' (TK) Fix to populate all stores
            fr4128.Cells(z, 1) = data.Cells(x, 1) 'name
            fr4128.Cells(z, 2) = data.Cells(x, 2) 'date
            fr4128.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr4128.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr4128.ListObjects("FR4128A").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

cc4128.Visible = True
cc4128.Select
End Sub

Sub Deleter4128()
Dim cc4128 As Worksheet
Dim fr4128 As Worksheet

Set cc4128 = ThisWorkbook.Sheets("4128CC")
Set fr4128 = ThisWorkbook.Sheets("4128FR")

    With cc4128.ListObjects("CC4128A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With fr4128.ListObjects("FR4128A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
End Sub
