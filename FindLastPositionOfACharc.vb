Function LastPosition(xCell As Range, xChar As String)

    Dim rLen As Integer
    xLen = Len(xCell)
        For i = xLen To 1 Step -1
        If Mid(xCell, i - 1, 1) = xChar Then
            LastPosition = i
            Exit Function
        End If
    Next i
End Function