
Private Sub btnMagic_Click()
    Dim currentCol As Integer, currentRow As Integer
    Dim i As Integer, j As Integer
    Dim sum As Integer
    Dim Ah As Integer, Aw As Integer, Bh As Integer, Bw As Integer, Ch As Integer, Cw As Integer

    Ah=Cells(6,2)
    Aw=Cells(6,3)
    Bh=Cells(6,4)
    Bw=Cells(6,6)
    Ch=Bh
    Cw=Aw
    For currentRow = 0 To Ch
        For currentCol = 0 To Cw
            sum=0
            For i = 0 To Bh
                Cells(3, 3)=Cells(i+24,currentCol+3)
                Cells(4, 4)=Cells(currentRow+10, j+3)
                sum = sum + (Cells(i+24,currentCol+3) * Cells(currentRow+10, i+3))
            Next i

            Cells(currentCol + 38, currentRow + 3) = sum
        Next currentCol
    Next currentRow

End Sub