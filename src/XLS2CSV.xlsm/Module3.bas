Attribute VB_Name = "Module3"
Function Str2Basename(path) As String
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  Str2Basename = fso.GetBaseName(path)
  Set fso = Nothing
  
End Function

Function JoinPaths(x, y) As String
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  JoinPaths = fso.Buildpath(x, y)
  Set fso = Nothing
  
End Function


Sub AllCellsStandard(ws As Worksheet)
  Dim c As Range
  Application.ScreenUpdating = False
  With ws.UsedRange
    .NumberFormat = "General"
    .Application.ErrorCheckingOptions.NumberAsText = False
  End With
  Application.ScreenUpdating = True
End Sub

Function IsExcelFile(filename As String) As Boolean
  Dim extension As String
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  extension = LCase(fso.GetExtensionName(filename))
  If extension = "xls" Or extension = "xlsx" Or extension = "xlsm" Then
    IsExcelFile = True
  Else
    IsExcelFile = False
  End If
  Set fso = Nothing
  
End Function


