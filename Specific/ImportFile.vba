
'Import v1

Sub File_import()
'Define files to import
Dim uploadfile As Variant
Dim importedfile As Workbook

Dim Country As String

    
MsgBox ("Please select file to be imported")

'Change to working directory
ChDir "SRCDIRECTORY"
uploadfile = Application.GetOpenFilename

If uploadfile = "False" Then
    MsgBox ("Upload false")
    Exit Sub
End If

Workbooks.Open uploadfile
Set importedfile = ActiveWorkbook
With importedfile
    
    Country = Range("B2").Value
    
End With
'Change filename and sheetname
Windows("File.xlsm").Activate
Sheets("Data").Select

Range("D4").Value = Country

End Sub
