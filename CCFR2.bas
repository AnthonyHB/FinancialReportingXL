Attribute VB_Name = "CCFR"

Sub DataTransfer()
Attribute DataTransfer.VB_ProcData.VB_Invoke_Func = "n\n14"

Dim data As Worksheet
Dim cc As Worksheet
Dim fr As Worksheet
Dim myIB1 as Long

Set data = ThisWorkbook.Sheets("Data")
Set cc = ThisWorkbook.Sheets("CSA CC Detail")
Set fr = ThisWorkbook.Sheets("CSA FR Detail")

dataLR = data.Cells(Rows.Count, 1).End(xlUp).Row
cc4128LR = cc.ListObjects("CC_4128").TotalsRowRange.Row 
cc4135LR = cc.ListObjects("CC_4135").TotalsRowRange.Row 
cc4234LR = cc.ListObjects("CC_4234").TotalsRowRange.Row 
cc4236LR = cc.ListObjects("CC_4236").TotalsRowRange.Row 
cc4338LR = cc.ListObjects("CC_4338").TotalsRowRange.Row 
cc4350LR = cc.ListObjects("CC_4350").TotalsRowRange.Row 
cc4369LR = cc.ListObjects("CC_4369").TotalsRowRange.Row 
fr4128LR = fr.ListObjects("FR_4128").TotalsRowRange.Row 
fr4135LR = fr.ListObjects("FR_4135").TotalsRowRange.Row 
fr4234LR = fr.ListObjects("FR_4234").TotalsRowRange.Row 
fr4236LR = fr.ListObjects("FR_4236").TotalsRowRange.Row 
fr4338LR = fr.ListObjects("FR_4338").TotalsRowRange.Row 
fr4350LR = fr.ListObjects("FR_4350").TotalsRowRange.Row 
fr4369LR = fr.ListObjects("FR_4369").TotalsRowRange.Row 


myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0
If myIB1 = Empty Then Exit Sub

On Error Resume Next

Application.ScreenUpdating = False

' 4128
' leaves 1 row before total row
y = cc4128LR - 1 
z = fr4128LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Finds current month. Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4128-1099.0000" Then ' (TK) Fix to populate all stores
            cc.Cells(y, 1) = data.Cells(x, 1) 'name
            cc.Cells(y, 2) = data.Cells(x, 2) 'date
            cc.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc.ListObjects("CC_4128").ListRows.Add AlwaysInsert:=True

        ElseIf Left(data.Cells(x, 1), 16) = "  4128-1205.0000" Then ' (TK) Fix to populate all stores
            fr.Cells(z, 1) = data.Cells(x, 1) 'name
            fr.Cells(z, 2) = data.Cells(x, 2) 'date
            fr.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr.ListObjects("FR_4128").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

' 4135
' leaves 1 row before total row
y = cc4135LR - 1
z = fr4135LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Finds current month. Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4135-1099.0000" Then ' (TK) Fix to populate all stores
            cc.Cells(y, 1) = data.Cells(x, 1) 'name
            cc.Cells(y, 2) = data.Cells(x, 2) 'date
            cc.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc.ListObjects("CC_4135").ListRows.Add AlwaysInsert:=True

        ElseIf Left(data.Cells(x, 1), 16) = "  4135-1205.0000" Then ' (TK) Fix to populate all stores
            fr.Cells(z, 1) = data.Cells(x, 1) 'name
            fr.Cells(z, 2) = data.Cells(x, 2) 'date
            fr.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr.ListObjects("FR_4135").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

' 4234
' leaves 1 row before total row
y = cc4234LR - 1
z = fr4234LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Finds current month. Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4234-1099.0000" Then ' (TK) Fix to populate all stores
            cc.Cells(y, 1) = data.Cells(x, 1) 'name
            cc.Cells(y, 2) = data.Cells(x, 2) 'date
            cc.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc.ListObjects("CC_4234").ListRows.Add AlwaysInsert:=True

        ElseIf Left(data.Cells(x, 1), 16) = "  4234-1205.0000" Then ' (TK) Fix to populate all stores
            fr.Cells(z, 1) = data.Cells(x, 1) 'name
            fr.Cells(z, 2) = data.Cells(x, 2) 'date
            fr.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr.ListObjects("FR_4234").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

' 4236
' leaves 1 row before total row
y = cc4236LR - 1
z = fr4236LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Finds current month. Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4236-1099.0000" Then ' (TK) Fix to populate all stores
            cc.Cells(y, 1) = data.Cells(x, 1) 'name
            cc.Cells(y, 2) = data.Cells(x, 2) 'date
            cc.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc.ListObjects("CC_4236").ListRows.Add AlwaysInsert:=True

        ElseIf Left(data.Cells(x, 1), 16) = "  4236-1205.0000" Then ' (TK) Fix to populate all stores
            fr.Cells(z, 1) = data.Cells(x, 1) 'name
            fr.Cells(z, 2) = data.Cells(x, 2) 'date
            fr.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr.ListObjects("FR_4236").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

' 4338
' leaves 1 row before total row
y = cc4338LR - 1
z = fr4338LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Finds current month. Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4338-1099.0000" Then ' (TK) Fix to populate all stores
            cc.Cells(y, 1) = data.Cells(x, 1) 'name
            cc.Cells(y, 2) = data.Cells(x, 2) 'date
            cc.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc.ListObjects("CC_4338").ListRows.Add AlwaysInsert:=True

        ElseIf Left(data.Cells(x, 1), 16) = "  4338-1205.0000" Then ' (TK) Fix to populate all stores
            fr.Cells(z, 1) = data.Cells(x, 1) 'name
            fr.Cells(z, 2) = data.Cells(x, 2) 'date
            fr.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr.ListObjects("FR_4338").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

' 4350
' leaves 1 row before total row
y = cc4350LR - 1
z = fr4350LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Finds current month. Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4350-1099.0000" Then ' (TK) Fix to populate all stores
            cc.Cells(y, 1) = data.Cells(x, 1) 'name
            cc.Cells(y, 2) = data.Cells(x, 2) 'date
            cc.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc.ListObjects("CC_4350").ListRows.Add AlwaysInsert:=True

        ElseIf Left(data.Cells(x, 1), 16) = "  4350-1205.0000" Then ' (TK) Fix to populate all stores
            fr.Cells(z, 1) = data.Cells(x, 1) 'name
            fr.Cells(z, 2) = data.Cells(x, 2) 'date
            fr.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr.ListObjects("FR_4350").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

' 4369
' leaves 1 row before total row
y = cc4369LR - 1
z = fr4369LR - 1

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Finds current month. Will not support multiple digit months
        If Left(data.Cells(x, 1), 16) = "  4369-1099.0000" Then ' (TK) Fix to populate all stores
            cc.Cells(y, 1) = data.Cells(x, 1) 'name
            cc.Cells(y, 2) = data.Cells(x, 2) 'date
            cc.Cells(y, 3) = data.Cells(x, 3) 'desc
            cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            y = y + 1
            cc.ListObjects("CC_4369").ListRows.Add AlwaysInsert:=True

        ElseIf Left(data.Cells(x, 1), 16) = "  4369-1205.0000" Then ' (TK) Fix to populate all stores
            fr.Cells(z, 1) = data.Cells(x, 1) 'name
            fr.Cells(z, 2) = data.Cells(x, 2) 'date
            fr.Cells(z, 3) = data.Cells(x, 3) 'desc
            fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
            z = z + 1
            fr.ListObjects("FR_4369").ListRows.Add AlwaysInsert:=True
        End If
    End If
Next x

Application.ScreenUpdating = True

cc4128.Visible = True
cc4128.Select

End Sub


Sub Deleter()

Dim cc As Worksheet
Dim fr As Worksheet

Set cc = ThisWorkbook.Sheets("CSA CC Detail")
Set fr = ThisWorkbook.Sheets("CSA FR Detail")

    With cc.ListObjects("CC_4128")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    With fr.ListObjects("FR_4128")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
End Sub
