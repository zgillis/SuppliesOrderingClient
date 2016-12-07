Module CustomFunctions

    'This function.
    Private Sub CalculateLineTotals()
        Dim intLineCount As Integer

        intItemQuantity(0) = Integer.Parse(txtQuantity1.Text)
        intItemQuantity(1) = Integer.Parse(txtQuantity2.Text)
        intItemQuantity(2) = Integer.Parse(txtQuantity3.Text)
        intItemQuantity(3) = Integer.Parse(txtQuantity4.Text)
        intItemQuantity(4) = Integer.Parse(txtQuantity5.Text)
        intItemQuantity(5) = Integer.Parse(txtQuantity6.Text)
        intItemQuantity(6) = Integer.Parse(txtQuantity7.Text)
        intItemQuantity(7) = Integer.Parse(txtQuantity8.Text)
        intItemQuantity(8) = Integer.Parse(txtQuantity9.Text)

        intItemQuantity(9) = Integer.Parse(txtQuantity10.Text)

        For intItemLoopCount = 0 To 10
            If intItemQuantity(intLineCount) > 1000 Or intItemQuantity(intLineCount) < 0 Then
                MsgBox("Please enter a quantity between 0 and 1000 for " & strItemName(intLineCount))
            Else
                decLineTotals(intLineCount) = decItemUnitPrice(intLineCount) * intItemQuantity(intLineCount)
            End If
        Next

        'FILL IN TEXTBOXES WITH CALCULATED LINE TOTALS 


    End Sub


End Module
