Attribute VB_Name = "CCFR"

Sub DataTransfer()
Attribute DataTransfer.VB_ProcData.VB_Invoke_Func = "n\n14"

Dim data As Worksheet
Dim cc As Worksheet
Dim fr As Worksheet
Dim t As ListObject
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
        
        If Left(data.Cells(x, 1), 7) = "  4128-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                Set t = cc.ListObjects("CC_4128")
                Call TableCC(t, x)

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                Set t = fr.ListObjects("FR_4128")
                Call TableFR(t, x)

            End If
        
        ' 4135
        ElseIf Left(data.Cells(x, 1), 7) = "  4135-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccR = cc.ListObjects("CC_4135").TotalsRowRange.Row 
                Call TableCC(t, x)

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frR = fr.ListObjects("FR_4135").TotalsRowRange.Row 
                Call TableFR(t, x)

            End If

        ' 4234
        ElseIf Left(data.Cells(x, 1), 7) = "  4234-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccR = cc.ListObjects("CC_4234").TotalsRowRange.Row 
                Call TableCC(t, x)

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frR = fr.ListObjects("FR_4234").TotalsRowRange.Row 
                Call TableFR(t, x)

            End If

        ' 4236
        ElseIf Left(data.Cells(x, 1), 7) = "  4236-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccR = cc.ListObjects("CC_4236").TotalsRowRange.Row 
                Call TableCC(t, x)

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frR = fr.ListObjects("FR_4236").TotalsRowRange.Row 
                Call TableFR(t, x)

            End If

        ' 4338
        ElseIf Left(data.Cells(x, 1), 7) = "  4338-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccR = cc.ListObjects("CC_4338").TotalsRowRange.Row 
                Call TableCC(t, x)

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frR = fr.ListObjects("FR_4338").TotalsRowRange.Row 
                Call TableFR(t, x)

            End If

        ' 4350
        ElseIf Left(data.Cells(x, 1), 7) = "  4350-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccR = cc.ListObjects("CC_4350").TotalsRowRange.Row 
                Call TableCC(t, x)

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frR = fr.ListObjects("FR_4350").TotalsRowRange.Row 
                Call TableFR(t, x)

            End If

        ' 4369
        ElseIf Left(data.Cells(x, 1), 7) = "  4369-" Then ' (TK) Fix to populate all stores

            If Right(data.Cells(x, 1), 9) = "1099.0000" Then   
                ccR = cc.ListObjects("CC_4369").TotalsRowRange.Row 
                Call TableCC(t, x)

            ElseIf Right(data.Cells(x, 1), 9) = "1205.0000" Then 
                frR = fr.ListObjects("FR_4369").TotalsRowRange.Row 
                Call TableFR(t, x)

            End If
        End If
    End If
Next x

Application.ScreenUpdating = True
cc.Visible = True
cc.Select
End Sub


Sub TableCC(ByVal t As ListObject, ByVal x As Long)
Dim data As Worksheet
Dim cc As Worksheet

Set cc = ThisWorkbook.Sheets("CSA CC Detail")

ccR = t.TotalsRowRange.Row 
y = ccR - 1      

cc.Cells(y, 1) = data.Cells(x, 1) 'name
cc.Cells(y, 2) = data.Cells(x, 2) 'date
cc.Cells(y, 3) = data.Cells(x, 3) 'desc
cc.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

y = y + 1
t.ListRows.Add AlwaysInsert:=True
End Sub


Sub TableFR(ByVal t As ListObject, ByVal x As Long)
Dim data As Worksheet
Dim fr As Worksheet

Set fr = ThisWorkbook.Sheets("CSA FR Detail")

frR = t.TotalsRowRange.Row 
y = frR - 1      

fr.Cells(y, 1) = data.Cells(x, 1) 'name
fr.Cells(y, 2) = data.Cells(x, 2) 'date
fr.Cells(y, 3) = data.Cells(x, 3) 'desc
fr.Cells(y, 4) = data.Cells(x, 4) + data.Cells(x, 5) 'Amount

y = y + 1
t.ListRows.Add AlwaysInsert:=True
End Sub

