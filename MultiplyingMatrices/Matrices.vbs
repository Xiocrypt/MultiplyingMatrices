
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
                Cells(4, 14)=i+24
                Cells(5, 14)=currentCol+3
                Cells(6, 14)=Cells(i+24,currentCol+3)
                
                Cells(8, 14)=currentRow+10
                Cells(9, 14)=j+3
                Cells(10, 14)=Cells(currentRow+10, j+3)

                Cells(12, 14)=currentRow
                Cells(13, 14)=currentCol

                Cells(15, 14)=Ah
                Cells(16, 14)=Aw

                Cells(18, 14)=Bh
                Cells(19, 14)=Bw

                Cells(21, 14)=Ch
                Cells(22, 14)=Cw
                sum = sum + (Cells(i+24,currentCol+3) * Cells(currentRow+10, i+3))
            Next i

            Cells(currentCol + 38, currentRow + 3) = sum
        Next currentCol
    Next currentRow

End Sub