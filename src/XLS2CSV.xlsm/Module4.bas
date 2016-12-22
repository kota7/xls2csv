Attribute VB_Name = "Module4"
' Interface handlers


Sub MainDir2CSV()
  Dim directory As String
  Dim outdir As String
  Dim separator As String
  Dim ForceGeneralFormat As Boolean
  Dim SkipHidden As Boolean
  
  Application.DisplayAlerts = False
  Debug.Print "****************"
  
  'read config
  With Worksheets("config")
    directory = .Range("InputDirectory").Value
    outdir = .Range("OutputDirectory").Value
    separator = .Range("Separator").Value
    If .Range("GeneralFlag").Value = "Yes" Then
      ForceGeneralFormat = True
    Else
      ForceGeneralFormat = False
    End If
    If .Range("SkipHiddenFlag").Value = "Yes" Then
      SkipHidden = True
    Else
      SkipHidden = False
    End If
  End With
  
  'input validity check
  If directory = "" Then
    MsgBox "Input directory is empty!"
    Exit Sub
  ElseIf outdir = "" Then
    MsgBox "Output directory is empty!"
    Exit Sub
  ElseIf Dir(directory, vbDirectory) = "" Then
    MsgBox "Input directory not found: " & directory & "!"
    Exit Sub
  ElseIf Dir(outdir, vbDirectory) = "" Then
    MsgBox "Output directory not found: " & directory & "!"
    Exit Sub
  End If
  
  Dir2CSV directory:=directory, outdir:=outdir, prefix:="", _
    separator:=separator, ForceGeneralFormat:=ForceGeneralFormat, SkipHidden:=SkipHidden
  MsgBox "Done:D"

  Application.DisplayAlerts = True
  
End Sub

Sub MainDir2CSVRecursive()
  
  Dim directory As String
  Dim outdir As String
  Dim separator As String
  Dim ForceGeneralFormat As Boolean
  Dim SkipHidden As Boolean
  
  Application.DisplayAlerts = False
  Debug.Print "****************"

  'read config
  With Worksheets("config")
    directory = .Range("InputDirectory").Value
    outdir = .Range("OutputDirectory").Value
    separator = .Range("Separator").Value
    If .Range("GeneralFlag").Value = "Yes" Then
      ForceGeneralFormat = True
    Else
      ForceGeneralFormat = False
    End If
    If .Range("SkipHiddenFlag").Value = "Yes" Then
      SkipHidden = True
    Else
      SkipHidden = False
    End If
  End With
  
  
  'input validity check
  If directory = "" Then
    MsgBox "Input directory is empty!"
    Exit Sub
  ElseIf outdir = "" Then
    MsgBox "Output directory is empty!"
    Exit Sub
  ElseIf Dir(directory, vbDirectory) = "" Then
    MsgBox "Input directory not found: " & directory & "!"
    Exit Sub
  ElseIf Dir(outdir, vbDirectory) = "" Then
    MsgBox "Output directory not found: " & directory & "!"
    Exit Sub
  End If

  Dir2CSVRecursive directory:=directory, outdir:=outdir, prefix:="", _
    separator:=separator, ForceGeneralFormat:=ForceGeneralFormat, SkipHidden:=SkipHidden
  MsgBox "Done:D"

  Application.DisplayAlerts = True

End Sub

Sub MainFile2CSV()
  Dim xlspath As String
  Dim outdir As String
  Dim separator As String
  Dim ForceGeneralFormat As Boolean
  Dim SkipHidden As Boolean
  
  Application.DisplayAlerts = False
  Debug.Print "****************"
  
  'read config
  With Worksheets("config")
    xlspath = .Range("InputFile").Value
    directory = .Range("InputDirectory").Value
    outdir = .Range("OutputDirectory").Value
    separator = .Range("Separator").Value
    If .Range("GeneralFlag").Value = "Yes" Then
      ForceGeneralFormat = True
    Else
      ForceGeneralFormat = False
    End If
    If .Range("SkipHiddenFlag").Value = "Yes" Then
      SkipHidden = True
    Else
      SkipHidden = False
    End If
  End With

  'input validity check
  If Dir(xlspath) = "" Then
    MsgBox "File not found: " & xlspath
    Exit Sub
  End If
  'input validity check
  If outdir = "" Then
    MsgBox "Output directory is empty!"
    Exit Sub
  ElseIf Dir(outdir, vbDirectory) = "" Then
    MsgBox "Output directory not found: " & directory & "!"
    Exit Sub
  End If


  XLS2CSV xlspath:=xlspath, outdir:=outdir, prefix:="", _
    separator:=separator, ForceGeneralFormat:=ForceGeneralFormat, SkipHidden:=SkipHidden
  MsgBox "Done:D"
  
  Application.DisplayAlerts = True

End Sub

