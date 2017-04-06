Attribute VB_Name = "PrintValidation"
Option Explicit

Sub PrintPDF()
Dim r As Long, i As Long
Dim list As Range
Dim strFile As String
Dim path As String
Dim report As String
Dim ws As Worksheet
Dim info As Worksheet

Application.ScreenUpdating = False

Set info = ThisWorkbook.Sheets("Info")
Set list = Range("Info!$B$9:$B$28")
Set ws = ActiveSheet
path = ThisWorkbook.path & ":"
report = "Report.pdf"
r = list.Cells.Count
strFile = ""

For i = 1 To r
    Range("C6") = list.Cells(i)
    strFile = path & Range("C6").Value & report
    ws.SaveAs Filename:=strFile, FileFormat:=xlPDF
Next i

Application.ScreenUpdating = True

End Sub
