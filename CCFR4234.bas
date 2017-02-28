Attribute VB_Name = "CCFR"

Sub DataTransfer4234()

Dim cc4234 As Worksheet
Dim fr4234 As Worksheet
Dim data As Worksheet

Set cc4234 = ThisWorkbook.Sheets("4234CC")
Set fr4234 = ThisWorkbook.Sheets("4234FR")
Set data = ThisWorkbook.Sheets("Data")

cc4234LR = cc4234.ListObjects("CC4234A").TotalsRowRange.Row
fr4234LR = fr4234.ListObjects("FR4234A").TotalsRowRange.Row

dataLR = data.Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0

If myIB1 = Empty Then Exit Sub

y = cc4234LR - 1 ' leaves 1 row before total row
z = fr4234LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4234-1099.0000" Then ' (TK) Fix to populate all stores
            cc4234.Cells(y, 1) = data.Cells(x, 1) 'name
            cc4234.Cells(y, 2) = data.Cells(x, 2) 'date
            cc4234.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc4234.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc4234.ListObjects("CC4234A").ListRows.Add AlwaysInsert:=True
        ElseIf Left(data.Cells(x, 1), 16) = "  4234-1205.0000" Then ' (TK) Fix to populate all stores
            fr4234.Cells(z, 1) = data.Cells(x, 1) 'name
            fr4234.Cells(z, 2) = data.Cells(x, 2) 'date
            fr4234.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr4234.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr4234.ListObjects("FR4234A").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

cc4234.Visible = True
cc4234.Select
End Sub

Sub Deleter4234()
Dim cc4234 As Worksheet
Dim fr4234 As Worksheet

Set cc4234 = ThisWorkbook.Sheets("4234CC")
Set fr4234 = ThisWorkbook.Sheets("4234FR")

    With cc4234.ListObjects("CC4234A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With fr4234.ListObjects("FR4234A")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
End Sub
