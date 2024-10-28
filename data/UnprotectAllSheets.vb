Sub UnprotectAllSheets()
    Dim ws As Worksheet
    Dim password As String
    password = "tsushyd@24"
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect password
    Next ws
End Sub