
Sub daoru():
    Dim i As Byte
    Dim j As Byte
    Dim k As Integer
    For i = 2 To 21 Step 1
        For j = 1 To Worksheets(1).Cells(i + 1, 4).Value Step 1
            For k = 2 To 3607 Step 1
                If Worksheets(i).Cells(j + 2, 2).Value = Worksheets(22).Cells(k, 3) Then
                    Worksheets(i).Cells(j + 2, 4).Value = Worksheets(22).Cells(k, 15)
                End If
            Next k
        Next j
    Next i
    MsgBox "Complete"
End Sub
