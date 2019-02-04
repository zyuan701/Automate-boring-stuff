Option Explicit

Sub CreateFile()
    
    Dim FileDir As String
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
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim FileOut As Object
    Dim myFileName As String
    Dim lastRow As Long
    lastRow = Cells("2,1").End(xlDown).Row
    
    Dim i As Long

        For i = 2 To lastRow
            If Not IsEmpty(Cells(i, 1)) Then
                myFileName = FileDir & Cells(i, 1) & ".txt"
                Set FileOut = fso.CreateTextFile(myFileName)
                FileOut.Close
            End If
        Next
    MsgBox "File Creation Completed"
    Set fso = Nothing
    Set FileOut = Nothing
End Sub
