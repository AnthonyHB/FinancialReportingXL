Attribute VB_Name = "CCFR"

Sub DataTransfer4236()

Dim cc4236 As Worksheet
Dim fr4236 As Worksheet
Dim data As Worksheet

Set cc4236 = ThisWorkbook.Sheets("4236CC")
Set fr4236 = ThisWorkbook.Sheets("4236FR")
Set data = ThisWorkbook.Sheets("Data")

cc4236LR = cc4236.ListObjects("CC4236A").TotalsRowRange.Row
fr4236LR = fr4236.ListObjects("FR4236A").TotalsRowRange.Row

dataLR = data.Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0

If myIB1 = Empty Then Exit Sub

y = cc4236LR - 1 ' leaves 1 row before total row
z = fr4236LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4236-1099.0000" Then ' (TK) Fix to populate all stores
            cc4236.Cells(y, 1) = data.Cells(x, 1) 'name
            cc4236.Cells(y, 2) = data.Cells(x, 2) 'date
            cc4236.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc4236.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc4236.ListObjects("CC4236A").ListRows.Add AlwaysInsert:=True
        ElseIf Left(data.Cells(x, 1), 16) = "  4236-1205.0000" Then ' (TK) Fix to populate all stores
            fr4236.Cells(z, 1) = data.Cells(x, 1) 'name
            fr4236.Cells(z, 2) = data.Cells(x, 2) 'date
            fr4236.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr4236.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr4236.ListObjects("FR4236A").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

cc4236.Visible = True
cc4236.Select
End Sub

Sub Deleter4236()
Dim cc4236 As Worksheet
Dim fr4236 As Worksheet

Set cc4236 = ThisWorkbook.Sheets("4236CC")
Set fr4236 = ThisWorkbook.Sheets("4236FR")

    With cc4236.ListObjects("CC4236A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With fr4236.ListObjects("FR4236A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
End Sub
