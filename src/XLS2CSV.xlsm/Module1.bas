Attribute VB_Name = "Module1"
Sub XlsDirDialog()
  With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = True Then
      Worksheets(1).Range("C2").Value = .SelectedItems(1)
    End If
  End With
End Sub


Sub OutDirDialog()
  With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show = True Then
      Worksheets(1).Range("C5").Value = .SelectedItems(1)
    End If
  End With
End Sub


Sub XlsFileDialog()
  Dim fname As String

  fname = Application.GetSaveAsFilename(InitialFileName:="", FileFilter:="Excel ƒtƒ@ƒCƒ‹ (*.xls; *.xlsx; *.xlsm),*.xls; *.xlsx; *.xlsm")
  If fname <> "False" Then
    Worksheets("config").Range("C3").Value = fname
  End If
End Sub
