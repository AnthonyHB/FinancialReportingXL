Attribute VB_Name = "CCFR"

Sub DataTransfer4350()

Dim cc4350 As Worksheet
Dim fr4350 As Worksheet
Dim data As Worksheet

Set cc4350 = ThisWorkbook.Sheets("4350CC")
Set fr4350 = ThisWorkbook.Sheets("4350FR")
Set data = ThisWorkbook.Sheets("Data")

cc4350LR = cc4350.ListObjects("CC4350A").TotalsRowRange.Row
fr4350LR = fr4350.ListObjects("FR4350A").TotalsRowRange.Row

dataLR = data.Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0

If myIB1 = Empty Then Exit Sub

y = cc4350LR - 1 ' leaves 1 row before total row
z = fr4350LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4350-1099.0000" Then ' (TK) Fix to populate all stores
            cc4350.Cells(y, 1) = data.Cells(x, 1) 'name
            cc4350.Cells(y, 2) = data.Cells(x, 2) 'date
            cc4350.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc4350.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc4350.ListObjects("CC4350A").ListRows.Add AlwaysInsert:=True
        ElseIf Left(data.Cells(x, 1), 16) = "  4350-1205.0000" Then ' (TK) Fix to populate all stores
            fr4350.Cells(z, 1) = data.Cells(x, 1) 'name
            fr4350.Cells(z, 2) = data.Cells(x, 2) 'date
            fr4350.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr4350.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr4350.ListObjects("FR4350A").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

cc4350.Visible = True
cc4350.Select
End Sub

Sub Deleter4350()
Dim cc4350 As Worksheet
Dim fr4350 As Worksheet

Set cc4350 = ThisWorkbook.Sheets("4350CC")
Set fr4350 = ThisWorkbook.Sheets("4350FR")

    With cc4350.ListObjects("CC4350A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With fr4350.ListObjects("FR4350A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
End Sub
