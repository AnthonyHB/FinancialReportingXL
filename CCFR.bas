Attribute VB_Name = "CCFR"

Sub Deleter()
Attribute Deleter.VB_ProcData.VB_Invoke_Func = "p\n14"

Call Deleter4128
Call Deleter4135
Call Deleter4234
Call Deleter4236
Call Deleter4338
Call Deleter4350
Call Deleter4369

End Sub

Sub DataTransfer()
Attribute DataTransfer.VB_ProcData.VB_Invoke_Func = "n\n14"

Dim myIB1 As Long

myIB1 = InputBox("What month is this for? (1-12)", "Month", "1") + 0
If myIB1 = Empty Then Exit Sub

Application.ScreenUpdating = False

Call DataTransfer4128(myIB1)
Call DataTransfer4135(myIB1)
Call DataTransfer4234(myIB1)
Call DataTransfer4236(myIB1)
Call DataTransfer4338(myIB1)
Call DataTransfer4350(myIB1)
Call DataTransfer4369(myIB1)

Application.ScreenUpdating = True
End Sub

Sub DataTransfer4128(myIB1 As Long)
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

Sub DataTransfer4135(myIB1 As Long)

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

Sub DataTransfer4234(myIB1 As Long)

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

Sub DataTransfer4236(myIB1 As Long)

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

Sub DataTransfer4338(myIB1 As Long)

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

Sub DataTransfer4350(myIB1 As Long)

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

Sub DataTransfer4369(myIB1 As Long)

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


