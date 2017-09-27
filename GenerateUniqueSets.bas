Dim i As Integer, j As Integer
Dim counter As Integer
counter = 1
For i = 1 To options
    For j = 1 + counter To options
        Debug.Print i & " vs " & j
    Next j
    counter = counter + 1
Next i
