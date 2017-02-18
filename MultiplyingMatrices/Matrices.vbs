
Private Sub FunctionName()
    Dim currentCol As Integer, currentRow As Integer
    Dim i As Integer, j As Integer
    Dim sumCol As Integer, sumRow As Integer
    Dim Ah As Integer, Aw As Integer, Bh As Integer, Bw As Integer, Ch As Integer, Cw As Integer

    Ah=Cells(6,2)
    Aw=Cells(6,3)
    Bh=Cells(6,4)
    Bw=Cells(6,6)
    Ch=Bh
    Cw=Aw
    For currentRow = 0 To Ch
        For currentCol = 0 To Cw
            sumCol=0
            For i = 0 To Bh
                sumCol = sumCol + Cells(i+24,currentCol+3)
            Next i

            sumRow=0
            For j = 0 To Aw
                sumRow = sumRow + Cells(currentRow+10, j+3)
            Next j

            Cells(currentCol + 38, currentRow + 3) = sumRow * sumCol
        Next j
    Next i

End Sub