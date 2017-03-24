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

myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0
If myIB1 = Empty Then Exit Sub
On Error Resume Next

Application.ScreenUpdating = False
dataLR = data.Cells(Rows.Count, 1).End(xlUp).Row

For x = 2 To dataLR
    If Left(data.Cells(x, 2), 2) = " " & myIB1 And Right(data.Cells(x, 2), 5) = "17   " Then ' (TK) Finds current month. Will not support multiple digit months

        ' 4128
        If Left(data.Cells(x, 1), 7) = "  4128-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccLR = cc.ListObjects("CC_4128").TotalsRowRange.Row 
                y = ccR - 1      

                cc.Cells(y, 1) = data.Cells(x, 1) 'name
                cc.Cells(y, 2) = data.Cells(x, 2) 'date
                cc.Cells(y, 3) = data.Cells(x, 3) 'desc
                cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

                y = y + 1
                cc.ListObjects("CC_4128").ListRows.Add AlwaysInsert:=True

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frLR = fr.ListObjects("FR_4128").TotalsRowRange.Row 
                z = frR - 1

                fr.Cells(z, 1) = data.Cells(x, 1) 'name
                fr.Cells(z, 2) = data.Cells(x, 2) 'date
                fr.Cells(z, 3) = data.Cells(x, 3) 'desc
                fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

                z = z + 1
                fr.ListObjects("FR_4128").ListRows.Add AlwaysInsert:=True
        
        ' 4135
        If Left(data.Cells(x, 1), 7) = "  4135-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccLR = cc.ListObjects("CC_4135").TotalsRowRange.Row 
                y = ccR - 1      

                cc.Cells(y, 1) = data.Cells(x, 1) 'name
                cc.Cells(y, 2) = data.Cells(x, 2) 'date
                cc.Cells(y, 3) = data.Cells(x, 3) 'desc
                cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

                y = y + 1
                cc.ListObjects("CC_4135").ListRows.Add AlwaysInsert:=True

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frLR = fr.ListObjects("FR_4135").TotalsRowRange.Row 
                z = frR - 1

                fr.Cells(z, 1) = data.Cells(x, 1) 'name
                fr.Cells(z, 2) = data.Cells(x, 2) 'date
                fr.Cells(z, 3) = data.Cells(x, 3) 'desc
                fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
                
                z = z + 1
                fr.ListObjects("FR_4135").ListRows.Add AlwaysInsert:=True

        ' 4234
        If Left(data.Cells(x, 1), 7) = "  4234-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccLR = cc.ListObjects("CC_4234").TotalsRowRange.Row 
                y = ccR - 1      

                cc.Cells(y, 1) = data.Cells(x, 1) 'name
                cc.Cells(y, 2) = data.Cells(x, 2) 'date
                cc.Cells(y, 3) = data.Cells(x, 3) 'desc
                cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

                y = y + 1
                cc.ListObjects("CC_4234").ListRows.Add AlwaysInsert:=True

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frLR = fr.ListObjects("FR_4234").TotalsRowRange.Row 
                z = frR - 1

                fr.Cells(z, 1) = data.Cells(x, 1) 'name
                fr.Cells(z, 2) = data.Cells(x, 2) 'date
                fr.Cells(z, 3) = data.Cells(x, 3) 'desc
                fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
                
                z = z + 1
                fr.ListObjects("FR_4234").ListRows.Add AlwaysInsert:=True

        ' 4236
        If Left(data.Cells(x, 1), 7) = "  4236-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccLR = cc.ListObjects("CC_4236").TotalsRowRange.Row 
                y = ccR - 1      

                cc.Cells(y, 1) = data.Cells(x, 1) 'name
                cc.Cells(y, 2) = data.Cells(x, 2) 'date
                cc.Cells(y, 3) = data.Cells(x, 3) 'desc
                cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

                y = y + 1
                cc.ListObjects("CC_4236").ListRows.Add AlwaysInsert:=True

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frLR = fr.ListObjects("FR_4236").TotalsRowRange.Row 
                z = frR - 1

                fr.Cells(z, 1) = data.Cells(x, 1) 'name
                fr.Cells(z, 2) = data.Cells(x, 2) 'date
                fr.Cells(z, 3) = data.Cells(x, 3) 'desc
                fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
                
                z = z + 1
                fr.ListObjects("FR_4236").ListRows.Add AlwaysInsert:=True

        ' 4338
        If Left(data.Cells(x, 1), 7) = "  4338-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccLR = cc.ListObjects("CC_4338").TotalsRowRange.Row 
                y = ccR - 1      

                cc.Cells(y, 1) = data.Cells(x, 1) 'name
                cc.Cells(y, 2) = data.Cells(x, 2) 'date
                cc.Cells(y, 3) = data.Cells(x, 3) 'desc
                cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

                y = y + 1
                cc.ListObjects("CC_4338").ListRows.Add AlwaysInsert:=True

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frLR = fr.ListObjects("FR_4338").TotalsRowRange.Row 
                z = frR - 1

                fr.Cells(z, 1) = data.Cells(x, 1) 'name
                fr.Cells(z, 2) = data.Cells(x, 2) 'date
                fr.Cells(z, 3) = data.Cells(x, 3) 'desc
                fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
                
                z = z + 1
                fr.ListObjects("FR_4338").ListRows.Add AlwaysInsert:=True

        ' 4350
        If Left(data.Cells(x, 1), 7) = "  4350-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccLR = cc.ListObjects("CC_4350").TotalsRowRange.Row 
                y = ccR - 1      

                cc.Cells(y, 1) = data.Cells(x, 1) 'name
                cc.Cells(y, 2) = data.Cells(x, 2) 'date
                cc.Cells(y, 3) = data.Cells(x, 3) 'desc
                cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

                y = y + 1
                cc.ListObjects("CC_4350").ListRows.Add AlwaysInsert:=True

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frLR = fr.ListObjects("FR_4350").TotalsRowRange.Row 
                z = frR - 1

                fr.Cells(z, 1) = data.Cells(x, 1) 'name
                fr.Cells(z, 2) = data.Cells(x, 2) 'date
                fr.Cells(z, 3) = data.Cells(x, 3) 'desc
                fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
                
                z = z + 1
                fr.ListObjects("FR_4350").ListRows.Add AlwaysInsert:=True

        ' 4369
        If Left(data.Cells(x, 1), 7) = "  4369-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccLR = cc.ListObjects("CC_4369").TotalsRowRange.Row 
                y = ccR - 1      

                cc.Cells(y, 1) = data.Cells(x, 1) 'name
                cc.Cells(y, 2) = data.Cells(x, 2) 'date
                cc.Cells(y, 3) = data.Cells(x, 3) 'desc
                cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

                y = y + 1
                cc.ListObjects("CC_4369").ListRows.Add AlwaysInsert:=True

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frLR = fr.ListObjects("FR_4369").TotalsRowRange.Row 
                z = frR - 1

                fr.Cells(z, 1) = data.Cells(x, 1) 'name
                fr.Cells(z, 2) = data.Cells(x, 2) 'date
                fr.Cells(z, 3) = data.Cells(x, 3) 'desc
                fr.Cells(z, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount
                
                z = z + 1
                fr.ListObjects("FR_4369").ListRows.Add AlwaysInsert:=True

            End If
        End If
    End If
Next x


Application.ScreenUpdating = True

cc.Visible = True
cc.Select

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
