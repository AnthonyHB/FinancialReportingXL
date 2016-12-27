Attribute VB_Name = "Module1"
Option Explicit
Sub ReadFilesInSequence()

  Dim FileName As String
  Dim StrFile As String
  Dim FileNumber As Long
  Dim PathCrnt As String
  Dim RowDestCrnt As Long
  Dim SheetDest As String
  Dim TgtValue As String
  Dim WBookSrc As Workbook

  PathCrnt = ActiveWorkbook.Path & ":"

  SheetDest = "Sheet1"
  RowDestCrnt = 1
  
  Application.ScreenUpdating = False

  With Worksheets(SheetDest)
    ' Delete current contents of destination sheet
    .Cells.EntireRow.Delete
  End With

  FileNumber = 1
    
  Do While True
    StrFile = PathCrnt & FileNumber & ".xlsm"
    
    On Error GoTo Err1
    FileName = Dir(StrFile)
    If FileName = "" Then
        Exit Do
Err1:
    Application.ScreenUpdating = True
    Exit Do
        End If
        
    Set WBookSrc = Workbooks.Open(PathCrnt & FileName)
    With WBookSrc.Worksheets("Daily Sales Recon")
      TgtValue = .Cells(1, "E").Value
    End With
    WBookSrc.Close SaveChanges:=False
    With Worksheets(SheetDest)
      .Cells(RowDestCrnt, "A").Value = FileName
      .Cells(RowDestCrnt, "B").Value = TgtValue
    End With
    
    RowDestCrnt = RowDestCrnt + 1

    FileNumber = FileNumber + 1
    

  Loop

End Sub
