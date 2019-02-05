Sub RenameFiles()


Dim xDir As String
Dim xFile As String
Dim xRow As Long
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .Title = "Please select a folder"
    .ButtonName = "Pick Folder"
    If .Show = 0 Then
        MsgBox "Nothing was selected"
        Exit Sub
    Else
        xDir = .SelectedItems(1)
        xFile = Dir(xDir & Application.PathSeparator & "*")
        
        Do Until xFile = ""
            xRow = 0
            On Error Resume Next
            xRow = Application.Match(xFile, Range("A:A"), 0)
            If xRow > 0 Then
                Name xDir & Application.PathSeparator & xFile As _
                xDir & Application.PathSeparator & Cells(xRow, "D").Value
            End If
            xFile = Dir
        Loop
    End If
End With

End Sub