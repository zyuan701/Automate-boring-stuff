Sub FillInFormula()

    Dim nrow As Integer
    nrow = 2
    
    Do Until Cells(nrow, 1) = ""
        Cells(nrow, 3).Formula = "=IFNA(RIGHT(A" & nrow & ",LEN(A" & nrow & ")-IFERROR(LastPosition(A" & nrow & ","".""),LEN(A" & nrow & ")+2)+2),"""")"
        Cells(nrow, 4).Formula = "=IF(B" & nrow & "="""",A" & nrow & ",B" & nrow & "&C" & nrow & ")"
        Cells(nrow, 5).Formula = "=IFERROR(LEFT(A" & nrow & ",IFERROR(LastPosition(A" & nrow & ",""."")-2,LEN(A" & nrow & "))),"""")"
        nrow = nrow + 1
    Loop
        
End Sub
