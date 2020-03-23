Sub unlockSheets()

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
		'Change to used password
        ws.unprotect Password:="password"

    Next

End Sub
