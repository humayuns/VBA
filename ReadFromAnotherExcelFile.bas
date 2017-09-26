Sub ReadFromAnotherExcelFile()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim vFile As Variant

    'Open the target workbook
    vFile = Application.GetOpenFilename("Excel-files,*.xls", _
        1, "Select One File To Open", , False)
    
    'if the user didn't select a file, exit sub
    If TypeName(vFile) = "Boolean" Then Exit Sub
    Workbooks.Open vFile
    'Set targetworkbook
    Set wb = ActiveWorkbook
   
    ' Print value from first cell in debug window.
    Debug.Print wb.Worksheets("Sheet1").Range("A1").Value
End Sub
