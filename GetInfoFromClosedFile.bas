Private Function GetInfoFromClosedFile(ByVal wbPath As String, _
    wbName As String, wsName As String, cellRef As String) As Variant
Dim arg As String
    GetInfoFromClosedFile = ""
    If Right(wbPath, 1) <> "" Then wbPath = wbPath & ""
    If Dir(wbPath & "" & wbName) = "" Then Exit Function
    arg = "'" & wbPath & "[" & wbName & "]" & _
        wsName & "'!" & Range(cellRef).Address(True, True, xlR1C1)
    On Error Resume Next
    GetInfoFromClosedFile = ExecuteExcel4Macro(arg)
End Function
