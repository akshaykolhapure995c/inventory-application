Public Class mdi
    Dim str1 As Label
    '  Dim addr As String = str1

    Private Sub ItemDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItemDetailsToolStripMenuItem.Click
        Add_Item_Details.Show()
    End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        Add_Item_Name.Show()
    End Sub

    Private Sub SalesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesToolStripMenuItem.Click
        billing.Show()
    End Sub

    Private Sub BuyersProfileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BuyersProfileToolStripMenuItem.Click
        buyersprofile.Show()
    End Sub

    Private Sub ChangeFirmAddressToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Label2.Text = InputBox("Enter Firm Address :")
        ' My.Settings.sd = Label2.Text
        ' My.Settings.Save()
    End Sub

    Private Sub mdi_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        For Each prog As Process In Process.GetProcesses
            If prog.ProcessName = "Bakale Billing" Then
                prog.Kill()
            End If
        Next
    End Sub

    Private Sub mdi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TextBox1.Text = My.Settings.sd
        ' Label2.Text = My.Settings.sd
        'Label4.Text = Date.Now()
    End Sub

    Private Sub NewBankDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewBankDetailsToolStripMenuItem.Click
        bank_details.Show()

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        showstock.Show()
    End Sub

    Private Sub PurchaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurchaseToolStripMenuItem.Click
        purchase.show()
    End Sub

    Private Sub PurchaseReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PurchaseReportToolStripMenuItem.Click
        purchase_report.show()
    End Sub

    Private Sub SalesReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesReportToolStripMenuItem.Click
        sales_report.Show()
    End Sub

    Private Sub ChangeBankDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub RectangleShape1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RectangleShape1.Click

    End Sub

    Private Sub PictureBox1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub SearchBillToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SearchBillToolStripMenuItem.Click
        search_bill.Show()
    End Sub

    Private Sub DeductItemQuantityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeductItemQuantityToolStripMenuItem.Click
        deduct_without_sales.Show()
    End Sub

    Private Sub ChangePasswordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangePasswordToolStripMenuItem.Click
        Change_Password.Show()
    End Sub

    Private Sub SalesWithoutStockToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

   

    Private Sub UpdateBuyersProfileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateBuyersProfileToolStripMenuItem.Click
        update_buyer.show()
    End Sub

    Private Sub SalesWithoutStockToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesWithoutStockToolStripMenuItem.Click
        sales_without_stock.Show()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        expenses.Show()
    End Sub

    Private Sub PartywiseBillToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PartywiseBillToolStripMenuItem.Click
        Party_wise_bill.Show()
    End Sub

    Private Sub LedgerBalanceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LedgerBalanceToolStripMenuItem.Click
        ledger_balance.Show()
    End Sub

    Private Sub ProfitLossToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProfitLossToolStripMenuItem.Click
        profit_loss.Show()
    End Sub

    Private Sub ViewBankStatmentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewBankStatmentToolStripMenuItem.Click
        Bank_Statement.Show()
    End Sub

    Private Sub DepreciationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DepreciationToolStripMenuItem.Click
        depreciation.Show()
    End Sub
End Class
