
'Export file v1

Sub Export_File()

Validation1 = Validate_Version()
If Validation1 = "fail" Then
    MsgBox "Unknown form version, please contact admin!", vbExclamation, "Error"
    GoTo Terminate
End If
Validation2 = Validate_Cells()
If Validation2 = "pass" Then
    ThisFile = Range("E3").Value & "_" & Range("E2").Value & "_" & Range("B2").Value & "_" & Range("B3").Value & "_" & Range("E1").Value & ".xlsm"
    ActiveWorkbook.SaveAs Filename:=ThisFile
    MsgBox "Exported successfully as:" & ThisFile, vbInformation, "Export successful"
Else
    MsgBox "Not all cells are filled correctly, please review the input!", vbExclamation, "Export canceled"
End If
Terminate:
End Sub

Function Validate_Cells() As String
Version = Range("E1").Value
'Version 1 possibility
If Version = "Countries_TST01" Then
    If ((WorksheetFunction.CountA(Range("B1:B2")) = 0) Or (WorksheetFunction.CountA(Range("B5:B7")) = 0) Or (WorksheetFunction.CountA(Range("B9:B15")) = 0) Or (WorksheetFunction.CountA(Range("E2:F3")) = 0) Or (WorksheetFunction.CountA(Range("e5:e13")) = 0)) Then
    Validate_Cells = "fail"
    Else
    Validate_Cells = "pass"
    End If
'Version 2 possibility	
ElseIf Version = "IA_TST01" Then
    Validate_Cells = "pass"
Else
    Validate_Cells = "fail"
End If
End Function

Function Validate_Version() As String
Version = Range("E1").Value
'
If ((Version = "Countries_TST01") Or (Version = "IA_TST01")) Then
    Validate_Version = "pass"
Else
    Validate_Version = "fail"
End If
End Function