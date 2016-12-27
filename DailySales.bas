Attribute VB_Name = "DailySales"
Sub ClearAll()
Attribute ClearAll.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' ClearWorksheet Macro
'
' Keyboard Shortcut: Option+Cmd+p
'
    Range("D34:AHD55").Select
    Selection.ClearContents
    Selection.ClearComments

    Range("D59:AH59").Select
    Selection.ClearContents
    
    Range("D34").Select

End Sub

Sub PullDataSingle()
Attribute PullDataSingle.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' PullDataSingle Macro
' CMD + alt + n

' go to CurrentRow, B Column
    If Cells(ActiveCell.Row, 2) <> Empty Then
    
    'This is the logic that pulls the data
        If ActiveCell.Row <> 43 Then
            With ActiveCell
                    .Formula = _
                        "=""=""&R34C1&R8C&R35C1&RC2"
                    .Value = .Value
                    .Value = .Value
                End With
            
        Else
            With ActiveCell
                .Formula = _
                    "=""=""&R34C1&R8C&R35C1&RC1&""+""&R34C1&R8C&R35C1&RC2 "
                .Value = .Value
                .Value = .Value
            End With
        End If
        
    End If
    ActiveCell.Offset(1, 0).Select
End Sub

Sub PullDataColumn()
Attribute PullDataColumn.VB_ProcData.VB_Invoke_Func = "n\n14"
' Cmd + Option + N
    
    If Cells(6, ActiveCell.Column) <> Empty And Cells(6, ActiveCell.Column) <> "Sunday" Then
    
        Cells(34, ActiveCell.Column).Select
        
            Do While ActiveCell.Row < 55
                If ActiveCell.Row < 55 Then
                    Call PullDataSingle
                Else
                    Exit Do
                End If
            Loop
            'After Loop
            ' 0.05 difference to discounts
            If Cells(32, ActiveCell.Column).Value < 0.05 And Cells(32, ActiveCell.Column).Value > -0.05 Then
                If Cells(32, ActiveCell.Column).Value <> 0 Then
                    Cells(37, ActiveCell.Column).Formula = "=" & Cells(37, ActiveCell.Column).Value & "-" & Round(Cells(32, ActiveCell.Column).Value, 2)
                Else
                    Cells(37, ActiveCell.Column).Formula = "=" & Cells(37, ActiveCell.Column).Value
                End If
            Else
            End If
            
            'End Here
            Cells(34, ActiveCell.Column + 1).Select
    
    'Handle Sundays'
    ElseIf Cells(6, ActiveCell.Column) = "Sunday" Then
        Cells(34, ActiveCell.Column + 1).Select
    Else
        MsgBox "Not a valid area for this macro."
    End If
        
End Sub

Sub SaveCSV()
Attribute SaveCSV.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' SaveCSV Macro
' Intro to creating CSV file with DS Workbook

' If test = 0
    If Range("$C$32") = 0 Then
    
    ' If 31, then add Column AH, else only go up to Column AG
        If Cells(8, 34) <> Empty Then
            ActiveWorkbook.Names.Add Name:="Mydata", RefersTo:= _
                "=$D$86:$AH$104"
        Else
            ActiveWorkbook.Names.Add Name:="Mydata", RefersTo:= _
                "=$D$86:$AG$104"
        End If
        
        ' Enter and drag INDEX formula
        Range("E110").Formula = "=INDEX(Mydata,1+INT((ROW(A1)-1)/COLUMNS(Mydata)),MOD(ROW(A1)-1+COLUMNS(Mydata),COLUMNS(Mydata))+1)"
        Range("E110").Select
        Selection.AutoFill Destination:=Range("E110:E700"), Type:=xlFillDefault
        Range("E110:E700").Select
        
        ' Copy and Paste
        Selection.Copy
        ExecuteExcel4Macro "WINDOW.SIZE(398,85,"""")"
        ExecuteExcel4Macro "WINDOW.MOVE(2,-43,"""")"
        Workbooks.Add
        Range("A1").Select
        Selection.PasteSpecial (xlPasteValues)
        
        ' Text to Columns
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
            Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1))
            
        ' Date Formating'
        Columns("C:C").NumberFormat = "m/d/yy"
        
        ' Delete #REF's
        Range("A1").Select
        Selection.End(xlDown).Select
        
        Do While ActiveCell.Text = "#REF!"
            If ActiveCell.Text = "#REF!" Then
                Selection.ClearContents
                Cells(ActiveCell.Row - 1, 1).Select
            Else
                Exit Do
            End If
        Loop
            
        ' Save and Close
        ActiveWorkbook.SaveAs Filename:= _
            ActiveWorkbook.Path & ":Uploads:" & Left(Cells(1, 2), 4) & "Upload.csv", FileFormat:=xlCSV, _
            CreateBackup:=False
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    Else
        MsgBox "Totals out of Balance"
    End If
End Sub
