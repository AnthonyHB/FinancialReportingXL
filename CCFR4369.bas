Attribute VB_Name = "CCFR"

Sub DataTransfer4369()

Dim cc4369 As Worksheet
Dim fr4369 As Worksheet
Dim data As Worksheet

Set cc4369 = ThisWorkbook.Sheets("4369CC")
Set fr4369 = ThisWorkbook.Sheets("4369FR")
Set data = ThisWorkbook.Sheets("Data")

cc4369LR = cc4369.ListObjects("CC4369A").TotalsRowRange.Row
fr4369LR = fr4369.ListObjects("FR4369A").TotalsRowRange.Row

dataLR = data.Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0

If myIB1 = Empty Then Exit Sub

y = cc4369LR - 1 ' leaves 1 row before total row
z = fr4369LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4369-1099.0000" Then ' (TK) Fix to populate all stores
            cc4369.Cells(y, 1) = data.Cells(x, 1) 'name
            cc4369.Cells(y, 2) = data.Cells(x, 2) 'date
            cc4369.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc4369.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc4369.ListObjects("CC4369A").ListRows.Add AlwaysInsert:=True
        ElseIf Left(data.Cells(x, 1), 16) = "  4369-1205.0000" Then ' (TK) Fix to populate all stores
            fr4369.Cells(z, 1) = data.Cells(x, 1) 'name
            fr4369.Cells(z, 2) = data.Cells(x, 2) 'date
            fr4369.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr4369.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr4369.ListObjects("FR4369A").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

cc4369.Visible = True
cc4369.Select
End Sub

Sub Deleter4369()
Dim cc4369 As Worksheet
Dim fr4369 As Worksheet

Set cc4369 = ThisWorkbook.Sheets("4369CC")
Set fr4369 = ThisWorkbook.Sheets("4369FR")

    With cc4369.ListObjects("CC4369A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With fr4369.ListObjects("FR4369A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
End Sub
