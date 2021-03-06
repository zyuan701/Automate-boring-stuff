Sub Loop_Inside_Folder()

    Dim FileDir As String
    Dim FiletoList As String
    Dim FileName As String
    Dim nrow As Integer

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Please select a folder"
        .ButtonName = "Pick Folder"
        'Cancel has show value of 0 and -1 means something was selected
        If .Show = 0 Then
            MsgBox "Nothing was selected"
            Exit Sub
        Else
            'Dir finds the first file in the folder
            FileDir = .SelectedItems(1) & "\"

        End If
    End With
    
    'FiletoList = Dir(FileDir & "Test_*xls*")
    FiletoList = Dir(FileDir, vbDirectory)
    
    nrow = 2
    
    Do Until FiletoList = ""
        If FiletoList <> "." And FiletoList <> ".." Then
            Cells(nrow, 1) = FiletoList
            nrow = nrow + 1
        End If
        
        FiletoList = Dir
    Loop
     
End Sub
Sub ClearAll()
    ActiveSheet.Range("A2").Select
    Range(Selection, Selection.End(xlDown)).ClearContents
End Sub