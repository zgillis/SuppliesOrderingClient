Public Class frmOrderForm

    'Declares a constant value for state sales tax.
    Public Const SALES_TAX As Decimal = 0.055

    'Declares two String arrays and one Decimal array, each 10 objects in size to
    'represent the item name, unit name, and unit prices for the 10 items for sale.
    'Additionally, an array is declared to hold the user's specified quantity for
    'each item and their decimal line total values that are calculated by the program.
    Dim strItemName As String() = New String(10) {}
    Dim strItemUnitName As String() = New String(10) {}
    Dim decItemUnitPrice As Decimal() = New Decimal(10) {}
    Dim intItemQuantity As Integer() = New Integer(10) {}
    Dim decLineTotals As Decimal() = New Decimal(10) {}


    Private Sub frmOrderForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
#Region "Initializing Arrays"
        'Fills in each of the 10 elements of the item name array with the appropriate
        'item names.
        strItemName = {
                "2"" x 6"" Lumber",
                "2"" x 4"" Lumber",
                "4' x 8' x 1/2"" Plywood",
                "4' x 8' x 5/8"" Plywood",
                "3/4"" Nails",
                "2 1/2"" Brads",
                "1 1/2"" Galvanized Screw",
                "10lb Sledge Hammer",
                "Five Drill Bits 1/16"" - 5/16",
                "Staple Gun"}

        'Fills in each of the 10 elements of the item unit name array with the appropriate
        'item units.
        strItemUnitName = {
                "10' Length",
                "12' length",
                "1 Sheet",
                "1 Sheet",
                "5lb Box",
                "10 lb Box",
                "1lb Box",
                "One",
                "1 Pack",
                "One"}

        'Fills in each of the 10 elements of the unit price array with the appropriate
        'unit prices.
        decItemUnitPrice = {
                3.49,
                2.45,
                6.25,
                6.25,
                5.95,
                5.49,
                4.65,
                11.95,
                8.85,
                10.95}

#End Region



#Region "Filling Labels From Arrays"

        'Assigns the value of each item name in the item name array to the
        'text value of the corresponding output labels.
        lblItemName1.Text = strItemName(0)
        lblItemName2.Text = strItemName(1)
        lblItemName3.Text = strItemName(2)
        lblItemName4.Text = strItemName(3)
        lblItemName5.Text = strItemName(4)
        lblItemName6.Text = strItemName(5)
        lblItemName7.Text = strItemName(6)
        lblItemName8.Text = strItemName(7)
        lblItemName9.Text = strItemName(8)
        lblItemName10.Text = strItemName(9)

        'Assigns the value of each item name in the unit name array to the
        'text value of the corresponding output labels.
        lblUnitName1.Text = strItemUnitName(0)
        lblUnitName2.Text = strItemUnitName(1)
        lblUnitName3.Text = strItemUnitName(2)
        lblUnitName4.Text = strItemUnitName(3)
        lblUnitName5.Text = strItemUnitName(4)
        lblUnitName6.Text = strItemUnitName(5)
        lblUnitName7.Text = strItemUnitName(6)
        lblUnitName8.Text = strItemUnitName(7)
        lblUnitName9.Text = strItemUnitName(8)
        lblUnitName10.Text = strItemUnitName(9)

        'Assigns the value of each item's price in the unit price array to the
        'text value of the corresponding output labels.
        lblUnitPrice1.Text = CDec(decItemUnitPrice(0))
        lblUnitPrice2.Text = CDec(decItemUnitPrice(1))
        lblUnitPrice3.Text = CDec(decItemUnitPrice(2))
        lblUnitPrice4.Text = CDec(decItemUnitPrice(3))
        lblUnitPrice5.Text = CDec(decItemUnitPrice(4))
        lblUnitPrice6.Text = CDec(decItemUnitPrice(5))
        lblUnitPrice7.Text = CDec(decItemUnitPrice(6))
        lblUnitPrice8.Text = CDec(decItemUnitPrice(7))
        lblUnitPrice9.Text = CDec(decItemUnitPrice(8))
        lblUnitPrice10.Text = CDec(decItemUnitPrice(9))
#End Region
        'Sets all quantity textboxes to '0'.
        ResetQuantities()

        'Sets all line total labels to zero currency decimal format.
        'E.G. (0.00)
        InitializeLineTotals()
    End Sub


#Region "Custom-Written Functions"

    'The following function is responsible for taking in arrays as parameters, and will calculate the line
    'total for a specific line on the invoice, as specified in this case by the index being passed from
    'TextBox.LostFocus events. In addition, this function verifies whether the quantity specified fits within
    'the 0-1000 requirement, and also performs a try/catch check to avoid the program from crashing if a user
    'is to use invalid, non-numeric characters.
    Private Sub CalculateLineTotal(intLineNum As Integer, ByRef strItemName As String(), ByRef decUnitPrice As Decimal(), ByRef txtLineQtyTextBox As TextBox, ByRef lblLineTotal As Label, ByRef decLineTotals As Decimal())
        'Reduces the value of intLineNum by one to account for array indexing.
        intLineNum -= 1

        'A try statement is used to catch any invalid characters from crashing the program when attempting to
        'do mathematical equations.
        Try
            'The line total of the item line given by the argument intLineNum. The program attempts to convert
            'the text box value on that line into an integer and multiply it by the unit price of that item line.
            'If successful, the result is stored in decLineTotals at the corresponding index.
            decLineTotals(intLineNum) = Integer.Parse(txtLineQtyTextBox.Text) * decItemUnitPrice(intLineNum)
            'If the line quantity exceeded 1000 or fell below 0, the program stops the user to display a message
            'box, and resets the value of that quantity text box to 0.
            If (CInt(txtLineQtyTextBox.Text) > 1000) Or CInt(txtLineQtyTextBox.Text < 0) Then
                MsgBox("Please enter a quantity between 0 and 1000" & vbCrLf & "for " & strItemName(intLineNum) & ".")
                txtLineQtyTextBox.Text = 0

            Else
                'If the user entered valid characters and the quantity fell within the range, the program continues to
                'display the line total, and then recalculate the entire invoice and display the most updated calculations.
                lblLineTotal.Text = Format(decLineTotals(intLineNum))
                CalculateTotal(decLineTotals, SALES_TAX)
            End If

            'If invalid data was entered into the textbox, the user receives a message box prompting them to try again
            ', and the value of the text box is reset to 0.
        Catch ex As Exception
            MsgBox("Please enter a valid numeric quantity between 0 and 1000" & vbCrLf & "for " & strItemName(intLineNum) & ".")
            txtLineQtyTextBox.Text = 0
        End Try
    End Sub

    'The following function calculates the invoice total by passing in the array with each line's total
    ', adding it up, calculating the tax from a specified tax rate, and then displaying the results.
    Private Sub CalculateTotal(ByRef decLineTotals As Decimal(), decTaxRate As Decimal)
        'Declares variables for the resulting calculations, and one to count the program's index location
        'while looping through the arrays.
        Dim decSubtotal As Decimal
        Dim decTotalTax As Decimal
        Dim decGrandTotal As Decimal
        Dim intLineLoopCount As Integer

        'The program loops through each line total, keeping a running total of the total charges.
        For intLineLoopCount = 0 To (decLineTotals.Length - 1) Step 1
            decSubtotal += decLineTotals(intLineLoopCount)
        Next

        'Tax is calculated by multiplying the subtotal and tax rate.
        decTotalTax = decSubtotal * decTaxRate
        'Total is calculated by adding the subtotal and tax.
        decGrandTotal = decSubtotal + decTotalTax

        'Displays the results by assigning the values to the subtotal, tax and total
        'textboxes, after being currency formatted.
        lblSubtotal.Text = Format(decSubtotal, "currency")
        lblTax.Text = Format(decTotalTax, "currency")
        lblTotal.Text = Format(decGrandTotal, "currency")
    End Sub


    Private Sub ResetQuantities()
        'Each of the quantity text boxes are assigned the value
        'of zero.
        txtQuantity1.Text = 0
        txtQuantity2.Text = 0
        txtQuantity3.Text = 0
        txtQuantity4.Text = 0
        txtQuantity5.Text = 0
        txtQuantity6.Text = 0
        txtQuantity7.Text = 0
        txtQuantity8.Text = 0
        txtQuantity9.Text = 0
        txtQuantity10.Text = 0
    End Sub

    Private Sub InitializeLineTotals()
        'Each line total label is set initially to $0.00, and this also
        'happens when the "Clear" button is clicked.
        lblLineTotal1.Text = Format(0, "currency")
        lblLineTotal2.Text = Format(0, "currency")
        lblLineTotal3.Text = Format(0, "currency")
        lblLineTotal4.Text = Format(0, "currency")
        lblLineTotal5.Text = Format(0, "currency")
        lblLineTotal6.Text = Format(0, "currency")
        lblLineTotal7.Text = Format(0, "currency")
        lblLineTotal8.Text = Format(0, "currency")
        lblLineTotal9.Text = Format(0, "currency")
        lblLineTotal10.Text = Format(0, "currency")

        'Additionally, the subtotal, tax and total labels are also reset.
        lblSubtotal.Text = Format(0, "currency")
        lblTax.Text = Format(0, "currency")
        lblTotal.Text = Format(0, "currency")
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Calls the Reset Quantities function to reset all quantity
        'text boxes to 0.
        ResetQuantities()

        'Calls the Initialize Line Totals function to reset all line totals to
        '$0.00.
        InitializeLineTotals()
    End Sub
#End Region

#Region "TextChanged Event Handlers"
    'The 'TextChanged' event for each text box triggers the CalculateLineTotal function, sending it the
    'line number and the text box object that contains a specified quantity. These events are triggered
    'by any change to the text value in any of the 10 text boxes.
    '--------------------------------------------------------------------------------------------------

    Private Sub txtQuantity1_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity1.LostFocus
        CalculateLineTotal(1, strItemName, decItemUnitPrice, sender, lblLineTotal1, decLineTotals)
    End Sub
    Private Sub txtQuantity2_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity2.LostFocus
        CalculateLineTotal(2, strItemName, decItemUnitPrice, sender, lblLineTotal2, decLineTotals)
    End Sub
    Private Sub txtQuantity3_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity3.LostFocus
        CalculateLineTotal(3, strItemName, decItemUnitPrice, sender, lblLineTotal3, decLineTotals)
    End Sub
    Private Sub txtQuantity4_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity4.LostFocus
        CalculateLineTotal(4, strItemName, decItemUnitPrice, sender, lblLineTotal4, decLineTotals)
    End Sub
    Private Sub txtQuantity5_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity5.LostFocus
        CalculateLineTotal(5, strItemName, decItemUnitPrice, sender, lblLineTotal5, decLineTotals)
    End Sub
    Private Sub txtQuantity6_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity6.LostFocus
        CalculateLineTotal(6, strItemName, decItemUnitPrice, sender, lblLineTotal6, decLineTotals)
    End Sub
    Private Sub txtQuantity7_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity7.LostFocus
        CalculateLineTotal(7, strItemName, decItemUnitPrice, sender, lblLineTotal7, decLineTotals)
    End Sub
    Private Sub txtQuantity8_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity8.LostFocus
        CalculateLineTotal(8, strItemName, decItemUnitPrice, sender, lblLineTotal8, decLineTotals)
    End Sub
    Private Sub txtQuantity9_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity9.LostFocus
        CalculateLineTotal(9, strItemName, decItemUnitPrice, sender, lblLineTotal9, decLineTotals)
    End Sub
    Private Sub txtQuantity10_LostFocus(sender As Object, e As EventArgs) Handles txtQuantity10.TextChanged
        CalculateLineTotal(10, strItemName, decItemUnitPrice, sender, lblLineTotal10, decLineTotals)
    End Sub
#End Region


#Region "GUI Components"

    'Upon clicking the exit button on the menu, the program exits.
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub ClearQuantitiesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearQuantitiesToolStripMenuItem.Click
        'Calls the Reset Quantities function to reset all quantity
        'text boxes to 0.
        ResetQuantities()
        'Calls the Initialize Line Totals function to reset all line totals to
        '$0.00.
        InitializeLineTotals()
    End Sub

    'Displays some information about the application.
    Private Sub AboutThisAppToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutThisAppToolStripMenuItem.Click
        MsgBox("Copyright © 2016, Zachary A. Gillis" & vbCrLf & vbTab & "-Senior IS Major at UW-Eau Claire", MsgBoxStyle.Information, "IS 304 Building Supplies Assignment")
    End Sub
#End Region
End Class
