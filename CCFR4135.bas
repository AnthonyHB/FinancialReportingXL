Attribute VB_Name = "CCFR"

Sub DataTransfer4135()

Dim cc4135 As Worksheet
Dim fr4135 As Worksheet
Dim data As Worksheet

Set cc4135 = ThisWorkbook.Sheets("4135CC")
Set fr4135 = ThisWorkbook.Sheets("4135FR")
Set data = ThisWorkbook.Sheets("Data")

cc4135LR = cc4135.ListObjects("CC4135A").TotalsRowRange.Row
fr4135LR = fr4135.ListObjects("FR4135A").TotalsRowRange.Row

dataLR = data.Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0

If myIB1 = Empty Then Exit Sub

y = cc4135LR - 1 ' leaves 1 row before total row
z = fr4135LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4135-1099.0000" Then ' (TK) Fix to populate all stores
            cc4135.Cells(y, 1) = data.Cells(x, 1) 'name
            cc4135.Cells(y, 2) = data.Cells(x, 2) 'date
            cc4135.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc4135.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc4135.ListObjects("CC4135A").ListRows.Add AlwaysInsert:=True
        ElseIf Left(data.Cells(x, 1), 16) = "  4135-1205.0000" Then ' (TK) Fix to populate all stores
            fr4135.Cells(z, 1) = data.Cells(x, 1) 'name
            fr4135.Cells(z, 2) = data.Cells(x, 2) 'date
            fr4135.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr4135.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr4135.ListObjects("FR4135A").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

cc4135.Visible = True
cc4135.Select
End Sub

Sub Deleter4135()
Dim cc4135 As Worksheet
Dim fr4135 As Worksheet

Set cc4135 = ThisWorkbook.Sheets("4135CC")
Set fr4135 = ThisWorkbook.Sheets("4135FR")

    With cc4135.ListObjects("CC4135A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With fr4135.ListObjects("FR4135A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
End Sub
