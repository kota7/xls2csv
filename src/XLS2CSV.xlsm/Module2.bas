Attribute VB_Name = "Module2"
Sub Dir2CSVRecursive(directory As String, outdir As String, prefix As String, _
                     separator As String, ForceGeneralFormat As Boolean, SkipHidden As Boolean)
  
  Dim fso As Object
  Dim myFolder As Object
  Dim muSubFolder As Object
  Dim newprefix As String
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  Dir2CSV directory:=directory, outdir:=outdir, prefix:=prefix, _
    separator:=separator, ForceGeneralFormat:=ForceGeneralFormat, SkipHidden:=SkipHidden
  
  
  ' look for next directory
  Set myFolder = fso.GetFolder(directory)
  For Each mySubFolder In myFolder.SubFolders
    If prefix <> "" Then
      newprefix = prefix & separator & Str2Basename(mySubFolder.path)
    Else
      newprefix = Str2Basename(mySubFolder.path)
    End If
    
    Dir2CSVRecursive directory:=mySubFolder.path, outdir:=outdir, prefix:=newprefix, _
      separator:=separator, ForceGeneralFormat:=ForceGeneralFormat, SkipHidden:=SkipHidden
  Next
  
  Set fso = Nothing
  Set myFolder = Nothing
  Set mysubforlder = Nothing
End Sub



Sub Dir2CSV(directory As String, outdir As String, prefix As String, _
            separator As String, ForceGeneralFormat As Boolean, SkipHidden As Boolean)
  Dim xlspath As String
  Dim fso As Object
  Dim myFolder As Object
  Dim curFile As Object
  
  If Dir(directory, vbDirectory) = "" Then
    Debug.Print "ERROR. Directory not found: " & directory
    Exit Sub
  End If
  Debug.Print "Directory: " & directory
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myFolder = fso.GetFolder(directory)
  For Each curFile In myFolder.Files
    If IsExcelFile(curFile.Name) Then
      XLS2CSV xlspath:=curFile.path, outdir:=outdir, prefix:=prefix, _
        separator:=separator, ForceGeneralFormat:=ForceGeneralFormat, SkipHidden:=SkipHidden
    End If
  Next
  
  Set fso = Nothing
  Set myFolder = Nothing
  Set curFile = Nothing

End Sub


Sub XLS2CSV(xlspath As String, outdir As String, prefix As String, _
            separator As String, ForceGeneralFormat As Boolean, SkipHidden As Boolean)
  ' Save all sheets of xlspath as separate CSV files.
  ' Each CSV is named as <basename>--<xls name>--<sheet name>.csv
  
  Application.DisplayAlerts = False
  Dim b As Workbook
  If Dir(xlspath) = "" Then
    Debug.Print "ERROR. File not exist: " & xlspath
    Exit Sub
  End If
  Debug.Print "Excel file: " & xlspath
  
  ' Call TryOpen function to handle errors
  Set b = TryOpen(xlspath)
  If b Is Nothing Then
    Exit Sub
  End If
  
  Dim name1 As String
  If prefix <> "" Then
    name1 = prefix & separator & Str2Basename(xlspath)
  Else
    name1 = Str2Basename(xlspath)
  End If
  
  Dim savename As String
  Dim savepath As String
  Dim ws As Worksheet
  For Each ws In b.Worksheets
    Debug.Print "  ", ws.Name, "->",
    If SkipHidden And (Not ws.Visible) Then
      Debug.Print "  hidden and ignored"
    Else
      savename = name1 & separator & ws.Name & ".csv"
    
      ws.Activate
      If ForceGeneralFormat Then
        Call AllCellsStandard(ws)
      End If
      savepath = JoinPaths(outdir, savename)
    
      b.SaveAs filename:=savepath, FileFormat:=xlCSV
      Debug.Print "  " & savepath & " saved."
    End If
  Next
  
  b.Close
  
  Application.DisplayAlerts = True
End Sub


Function TryOpen(filepath As String) As Workbook

  On Error GoTo SecondTry
  Set TryOpen = Workbooks.Open(filepath, _
                               UpdateLinks:=0, _
                               IgnoreReadOnlyRecommended:=True)
  Exit Function
  
SecondTry:
  On Error GoTo ThirdTry
  Debug.Print "Opening in Repair Mode..."
  Set TryOpen = Workbooks.Open(filepath, _
                               UpdateLinks:=0, _
                               IgnoreReadOnlyRecommended:=True, _
                               CorruptLoad:=xlRepairFile)
  Exit Function
  
  
ThirdTry:
  Dim ans As Integer
  ans = MsgBox("Problem in opening: " & vbCr & filepath & vbCr & _
               "You may fix the problem and retry, or skip this file.", _
               vbRetryCancel)
  If ans = vbRetry Then
    ' retry
    Set TryOpen = TryOpen(filepath)
    
  ElseIf ans = vbCancel Then
    Set TryOpen = Nothing
    Exit Function
  End If
  
End Function


