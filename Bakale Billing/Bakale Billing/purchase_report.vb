Public Class purchase_report
    Dim con As New ADODB.Connection
    Dim cmd2 As New ADODB.Recordset
   

    Private Sub purchase_report_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim dt As New DataTable
        With dt
            .Columns.Add("srno")
            .Columns.Add("sellername")
            .Columns.Add("invoicenum")
            .Columns.Add("date_")
            .Columns.Add("tin_num")
            .Columns.Add("net_rs")
            .Columns.Add("tax_rs")
            .Columns.Add("val_of_incl_tax")
            .Columns.Add("valofcomp")
            .Columns.Add("taxfreepur")
            .Columns.Add("exemptedpur")
            .Columns.Add("labourcharge")
            .Columns.Add("otherchatge")
            .Columns.Add("gross")
            .Columns.Add("action")
            .Columns.Add("returnfromnum")
            .Columns.Add("transcode")
            .Columns.Add("description")


        End With
        For Each dr As DataGridViewRow In Me.DataGridView1.Rows
            dt.Rows.Add(dr.Cells(0).Value, dr.Cells(1).Value, dr.Cells(2).Value, dr.Cells(3).Value, dr.Cells(4).Value, dr.Cells(5).Value, dr.Cells(6).Value, dr.Cells(7).Value, dr.Cells(8).Value, dr.Cells(9).Value, dr.Cells(10).Value, dr.Cells(11).Value, Label2.Text, Label4.Text, Label5.Text, Label6.Text, Label7.Text, Label8.Text)
        Next
        Dim rptdoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptdoc = New CrystalReport3
        rptdoc.SetDataSource(dt)
        ' rptdoc.PrintToPrinter(1, False, 0, 0)

        purchase_report_viewer.CrystalReportViewer1.ReportSource = rptdoc
        purchase_report_viewer.ShowDialog()
        purchase_report_viewer.Dispose()

        ' purchase_report_viewer.Show()

    End Sub

    
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        DataGridView1.Rows.Clear()
        Dim dte As String
        Dim i As Integer
        i = 1
        cmd2.Open("select * from purchase", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd2.EOF
            dte = cmd2.Fields("date_").Value
            If (dte.ToString.Contains(ComboBox1.Text) And dte.ToString.Contains(ComboBox2.Text)) Then
                DataGridView1.Rows.Add(i, cmd2.Fields("date_").Value, cmd2.Fields("bill_no").Value, cmd2.Fields("name_").Value, cmd2.Fields("gst").Value, cmd2.Fields("hsn").Value, cmd2.Fields("purchase").Value, cmd2.Fields("igst").Value, cmd2.Fields("sgst").Value, cmd2.Fields("cgst").Value, cmd2.Fields("round").Value, cmd2.Fields("gross").Value, cmd2.Fields("gst").Value)
                i = i + 1
                cmd2.MoveNext()
            Else
                cmd2.MoveNext()
            End If
        End While
        cmd2.Close()
        Dim y, purchase, igst, sgst, cgst, gross, r As Integer
        purchase = 0
        igst = 0
        cgst = 0
        gross = 0
        r = 0
        sgst = 0
        For y = 0 To DataGridView1.RowCount - 1
            purchase = purchase + Val(DataGridView1.Item("Column7", DataGridView1.Rows.Item(y).Index).Value)
            igst = igst + Val(DataGridView1.Item("Column8", DataGridView1.Rows.Item(y).Index).Value)
            sgst = sgst + Val(DataGridView1.Item("Column9", DataGridView1.Rows.Item(y).Index).Value)
            cgst = cgst + Val(DataGridView1.Item("Column10", DataGridView1.Rows.Item(y).Index).Value)
            r = r + Val(DataGridView1.Item("Column11", DataGridView1.Rows.Item(y).Index).Value)
            gross = gross + Val(DataGridView1.Item("Column12", DataGridView1.Rows.Item(y).Index).Value)
        Next y
        Label2.Text = Val(purchase)
        Label4.Text = Val(igst)
        Label5.Text = Val(sgst)
        Label6.Text = Val(cgst)
        Label8.Text = Val(r)
        Label7.Text = Val(gross)
    End Sub

    
    Private Sub RectangleShape1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RectangleShape1.Click

    End Sub
End Class