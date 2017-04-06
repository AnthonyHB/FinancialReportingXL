Attribute VB_Name = "PrintValidation"
Sub PrintPDF()
Dim r As Long, i As Long
Dim list As Range
Dim strFile As String
Dim path As String
Dim sdf As String
Dim ws As Worksheet

Application.ScreenUpdating = False

list = Range("=Info!$B$9:$B$28")

ws = ActiveSheet
path = ThisWorkbook.Path & ":""
pdf = "Report.pdf"
r = list.Cells.Count

For i = 1 To r
    Range("C6") = list.Cells(i)
    strFile = path & Range("C6") & pdf
    ws.SaveAs Filename:=strFile, FileFormat:=xlPDF
Next i

Application.ScreenUpdating = True

End Sub
