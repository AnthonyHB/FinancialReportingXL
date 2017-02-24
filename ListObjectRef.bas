Attribute VB_Name = "Module1"
Sub myFirstLoop()

For x = 1 To 10
      Cells(x, 1) = x * 12.75
      If Cells(x, 1) > 50 Then
            Cells(x, 2) = True
            Cells(x, 2).Font.Bold = False
        Else
            Cells(x, 2) = False
            Cells(x, 2).Font.Bold = True
        End If
Next x

End Sub
Sub myFirstReport()
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

myIB1 = InputBox("How much money should they make?", "How much?", "200") + 0

For x = 2 To lastRow
    If Cells(x, 4) >= myIB1 Then
        myMsg = myMsg & vbNewLine & Cells(x, 1) & ", " & Cells(x, 2)
    End If
Next x

MsgBox myMsg

End Sub

Sub myPrintableReport()
Dim dsheet As Worksheet
Dim rptsheet As Worksheet

Set dsheet = ThisWorkbook.Sheets("dsheet")
Set rptsheet = ThisWorkbook.Sheets("rptsheet")

rptLR = rptsheet.ListObjects("Table1").TotalsRowRange.Row

lastRow = dsheet.Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

myIB1 = InputBox("How much money should they make?", "How much?", "200") + 0
If myIB1 = Empty Then Exit Sub

y = rptLR - 1 ' leaves 1 row before total row

For x = 2 To lastRow
    If dsheet.Cells(x, 4) >= myIB1 Then
            rptsheet.Cells(y, 1) = dsheet.Cells(x, 1) 'name
            rptsheet.Cells(y, 2) = dsheet.Cells(x, 4) 'sale amt
            y = y + 1
            rptsheet.ListObjects("Table1").ListRows.Add AlwaysInsert:=True
    End If
Next x

rptsheet.Visible = True
rptsheet.Select

End Sub

Sub StepLoop()
For x = 20 To 2 Step -5
    MsgBox x
Next x
End Sub

Sub forEachLoop()

For Each cell In Range("names")
    If cell = "Alan" Then Exit For
    MsgBox cell
Next cell

End Sub
Sub myExample1()

For Each sht In ActiveWorkbook.Sheets
    MsgBox sht.Name
Next sht

End Sub

Sub myExample2()
For Each pvt In ActiveSheet.PivotTables
    'if, then, etc.
    MsgBox pvt.Name
Next pvt
End Sub

Sub DoLoop()
' Until/While
x = 2
Do
    myVar = Cells(x, 1)
    x = x + 1
Loop Until Cells(x, 1) = ""

End Sub

Sub DoLoop2()
x = 2
Do
    If Cells(x, 1) = "" Then Exit Do
    myVar = Cells(x, 1)
    x = x + 1
Loop

End Sub
Sub Deleter()
Dim dsheet As Worksheet
Dim rptsheet As Worksheet

Set dsheet = ThisWorkbook.Sheets("dsheet")
Set rptsheet = ThisWorkbook.Sheets("rptsheet")

    With rptsheet.ListObjects("Table1")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
End Sub
