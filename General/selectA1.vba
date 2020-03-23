Sub selectA1()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
		ws.Activate
        ActiveSheet.Range("A1").Select

    Next

End Sub
