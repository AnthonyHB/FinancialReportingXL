Attribute VB_Name = "CCFR"

Sub DataTransfer4338()

Dim cc4338 As Worksheet
Dim fr4338 As Worksheet
Dim data As Worksheet

Set cc4338 = ThisWorkbook.Sheets("4338CC")
Set fr4338 = ThisWorkbook.Sheets("4338FR")
Set data = ThisWorkbook.Sheets("Data")

cc4338LR = cc4338.ListObjects("CC4338A").TotalsRowRange.Row
fr4338LR = fr4338.ListObjects("FR4338A").TotalsRowRange.Row

dataLR = data.Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0

If myIB1 = Empty Then Exit Sub

y = cc4338LR - 1 ' leaves 1 row before total row
z = fr4338LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4338-1099.0000" Then ' (TK) Fix to populate all stores
            cc4338.Cells(y, 1) = data.Cells(x, 1) 'name
            cc4338.Cells(y, 2) = data.Cells(x, 2) 'date
            cc4338.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc4338.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc4338.ListObjects("CC4338A").ListRows.Add AlwaysInsert:=True
        ElseIf Left(data.Cells(x, 1), 16) = "  4338-1205.0000" Then ' (TK) Fix to populate all stores
            fr4338.Cells(z, 1) = data.Cells(x, 1) 'name
            fr4338.Cells(z, 2) = data.Cells(x, 2) 'date
            fr4338.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr4338.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr4338.ListObjects("FR4338A").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

cc4338.Visible = True
cc4338.Select
End Sub

Sub Deleter4338()
Dim cc4338 As Worksheet
Dim fr4338 As Worksheet

Set cc4338 = ThisWorkbook.Sheets("4338CC")
Set fr4338 = ThisWorkbook.Sheets("4338FR")

    With cc4338.ListObjects("CC4338A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With fr4338.ListObjects("FR4338A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
End Sub
