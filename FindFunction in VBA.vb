Option Compare Text

Sub CheckCategory()

    Dim Listi As Integer
    Dim ListLrow As Integer
    Dim Namej As Integer
    Dim NameLrow As Integer
    Dim fndWord As Integer
    Dim Sht1 As Worksheet
    Dim Sht2 As Worksheet
    
    On Error Resume Next 'Suppress Errors... for when we don't find a match
    
    'Define worksheet that has data on it....
    Set Sht1 = Sheets("Checklist")
    Set Sht2 = Sheets("FileName")
    
    'Get last row for words based on column D in Checklist sheet
    ListLrow = Sht1.Cells(Rows.Count, "D").End(xlUp).Row
    
    'Get last row for comments based on column A in FileName sheet
    NameLrow = Sht2.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through lists and find matches....
    For Listi = 2 To ListLrow
    
        For Namej = 2 To NameLrow

            'Look for word...
            fndWord = Application.WorksheetFunction.Search(Sht1.Cells(Listi, "D"), Sht2.Cells(Namej, "A"))
            
            'If we found the word....then
            If fndWord > 0 Then
                Sht2.Cells(Namej, "B") = Sht1.Cells(Listi, "D")
                fndWord = 0 'Reset Variable for next loop
            End If
        
        Next Namej
    
    Next Listi
    
    With Sht2.Columns("B")
        .SpecialCells(xlCellTypeBlanks).Font.Color = vbRed
        .SpecialCells(xlCellTypeBlanks).Value = "N/A"
        .HorizontalAlignment = xlLeft
    End With
    
End Sub
