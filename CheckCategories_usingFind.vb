Sub CheckCategory()
    Application.ScreenUpdating = False
    
    Dim Listi As Integer
    Dim ListLrow As Integer
    Dim Namej As Integer
    Dim NameLrow As Integer
    Dim fndWord As Range
    Dim Sht1 As Worksheet
    Dim Sht2 As Worksheet
    Dim kw As Variant
    
    
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
            For Each kw In Split(Sht1.Cells(Listi, "D"), ";")
            'Look for word...
            Set fndWord = Sht2.Cells(Namej, "A").Find(What:=kw, LookIn:=xlValues)
            
            'If we found the word....then
            If Not fndWord Is Nothing Then
                Sht2.Cells(Namej, "B") = Sht1.Cells(Listi, "E")
                fndWord = Nothing 'Reset Variable for next loop
            End If
            Next kw
        Next Namej
    
    Next Listi
    
    With Sht2.Range("A2:B2").Select
        Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeBlanks).Font.Color = vbRed
        Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeBlanks).Value = "N/A"
        Range(Selection, Selection.End(xlDown)).HorizontalAlignment = xlLeft
    End With
    
    Application.ScreenUpdating = True

End Sub
