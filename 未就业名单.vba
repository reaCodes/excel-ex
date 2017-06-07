Option Explicit

Sub wjy()
    Dim i As Byte
    Dim j As Byte
    Dim m As Byte
    m = 3
    For i = 3 To 22 Step 1
        For j = 1 To Worksheets(1).Cells(i, 4).Value Step 1
            If Worksheets(i).Cells(j + 2, 4) = "" Then
                Worksheets(2).Cells(m, 2) = Worksheets(i).Cells(j + 2, 1)
                Worksheets(2).Cells(m, 3) = Worksheets(1).Cells(i, 1)
                Worksheets(2).Cells(m, 4) = Worksheets(i).Cells(j + 2, 2)
                Worksheets(2).Cells(m, 6) = Worksheets(1).Cells(i, 2)
                Worksheets(2).Cells(m, 7) = Worksheets(1).Cells(i, 3)
                m = m + 1
            End If
        Next j
    Next i
    MsgBox "Complete"
End Sub

