Module Module1
    'clever.code Interleaved 2 of 5
    Public Function barCode(ByVal barText As String, ByVal addText As Boolean, ByRef barContainer As Object)
        ' X is dimension of the width of the smallest element of the barcode symbol .
        ' N is the ratio between wide to narrow symbol
        Dim pattern(9, 4) As Integer
        Const N = 3
        Const symbol = "00110100010100111000001011010001100000111001001010"
        Dim Xdim As Double = 0.254
        Const QuietZone = 6.35
        Const TextZone = 0.5
        barContainer.Cls()
        barText = Trim(barText)
        ' check if Number only
        For i = 1 To Len(barText)
            If Asc(Mid(barText, i, 1)) > 47 And Asc(Mid(barText, i, 1)) < 58 Then
            Else
                barText = ""
                Exit For
            End If
        Next
        If barText = "" Then
            barContainer.Print("Number")
            Exit Function
        End If
        Dim spacecolor As Color = barContainer.BackColor
        Dim barcolor As Color = barContainer.ForeColor
        For i = 0 To 9
            For j = 0 To 4
                pattern(i, j) = Val(Mid(symbol, i * 5 + j + 1, 1))
            Next
        Next
        'make number even
        If Len(barText) Mod 2 <> 0 Then barText = "0" & barText
        ' get pixel to mm ratio
        barContainer.ScaleMode = 3
        Dim lpix As Double = barContainer.ScaleWidth
        barContainer.ScaleMode = 6
        Dim hmm As Double = barContainer.ScaleHeight
        Dim lmm As Double = barContainer.ScaleWidth
        Dim ratio As Double = lpix / lmm
        Dim lbarcode As Double = lmm - (2 * QuietZone)
        ' check whether text required below barcode and then adjust height accordingly
        If addText Then
            hmm = hmm - barContainer.TextHeight("8") - TextZone
        End If
        Xdim = lbarcode / (6 + N + (Len(barText) * (2 * N + 3)))
        barContainer.DrawWidth = Int(ratio * Xdim) + Abs(Int(ratio * Xdim) < 1)
        Xdim = barContainer.DrawWidth / ratio
        Dim length As Double = (6 + N + (Len(barText) * (2 * N + 3))) * Xdim
        Dim Height As Double = 0.15 * length
        If Height < 6.35 Then Height = 6.35
        If hmm < Height Or lbarcode < length Then
            barContainer.Print("Size")
            Exit Function
        End If
        ' pair the symbols    & draw barcode
        ' start symbol  NnNn
        Dim x As Double = QuietZone
        Dim y As Double = 1
        For i = 0 To 1
        barContainer.Line (x, y)-(x, hmm), barcolor, BF
            x = x + Xdim
        barContainer.Line (x, y)-(x, hmm), spacecolor, BF
            x = x + Xdim
        Next
        For i = 0 To Len(barText) - 2 Step 2
            Dim pair As String = Mid(barText, i + 1, 2)
            Dim txtX As Double = x
            For j = 0 To 4
                ' draw bar
                For k = 0 To (N - 1) * pattern(Val(Left(pair, 1)), j)
                barContainer.Line (x, y)-(x, hmm), barcolor, BF
                    x = x + Xdim
                Next
                ' draw space
                For k = 0 To (N - 1) * pattern(Val(Right(pair, 1)), j)
                barContainer.Line (x, y)-(x, hmm), spacecolor, BF
                    x = x + Xdim
                Next
            Next
            ' write text below
            If addText Then
                barContainer.CurrentX = txtX
                Dim txty As Double = y
                barContainer.CurrentY = hmm + TextZone
                barContainer.Print(pair)
                barContainer.CurrentX = x
                barContainer.CurrentY = txty
            End If
        Next
        ' stop symbol  WnN
        For i = 0 To 2
            barContainer.Line (x, y)-(x, hmm), barcolor, BF
            x = x + Xdim
        Next i
            barContainer.Line (x, y)-(x, hmm), spacecolor, BF
        x = x + Xdim
            barContainer.Line (x, y)-(x, hmm), barcolor, BF
    End Function

End Module
