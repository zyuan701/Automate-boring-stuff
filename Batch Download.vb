Option Explicit

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
    ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Sub Batch_Download()
    
    Dim ws As Worksheet
    Dim strurl As String
    Dim strpath As String
    Dim strname As String
    Dim i As Long
    Dim LastRow As Long
    
    Set ws = Sheets("Sheet1")

    LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    strpath = "C:\Users\replace_with_yourname\Downloads\"

    For i = 2 To LastRow

        strurl = Cells(i, 2)
        strname = Cells(i, 1)
        URLDownloadToFile 0, strurl, strpath & strname & ".pdf", 0, 0
    

    Next i
    
End Sub
