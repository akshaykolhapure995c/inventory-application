Public Class search_bill
    Dim con As New ADODB.Connection
    Dim cmd, cmd2 As New ADODB.Recordset
    Dim v As New Integer
    Dim srno As New Integer
    Dim names As New HashSet(Of String)
    Dim id As Integer
    Dim printcopy As Boolean
    Dim line As Integer
    Dim count As Integer = 0
    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        cmd.Open("select * from sales_report where bill_no = '" & TextBox28.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            TextBox33.Text = cmd.Fields("round").Value
            cmd.MoveNext()
        End While
        cmd.Close()

        cmd.Open("select * from sales where bill_no = '" & TextBox28.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If Not cmd.EOF Then
            While Not cmd.EOF

                TextBox20.Text = cmd.Fields("bill_no").Value
                Me.DataGridView1.Rows.Add(cmd.Fields("sr_no").Value, cmd.Fields("particulars").Value, cmd.Fields("hsn").Value, cmd.Fields("size").Value, cmd.Fields("piece_qty").Value, cmd.Fields("per_piece_rate").Value, cmd.Fields("item_amount").Value, cmd.Fields("remaining_stock").Value, "X")

                ComboBox1.Text = cmd.Fields("buyer_name").Value
                TextBox1.Text = cmd.Fields("buyer_address").Value
                TextBox26.Text = cmd.Fields("buyer_state").Value
                TextBox25.Text = cmd.Fields("buyer_gst").Value
                TextBox2.Text = cmd.Fields("buyer_tin").Value
                TextBox21.Text = cmd.Fields("buyer_pan").Value
                DateTimePicker2.Text = cmd.Fields("order_date").Value
                TextBox5.Text = cmd.Fields("transport_name").Value
                TextBox7.Text = cmd.Fields("supplier_ref").Value
                TextBox8.Text = cmd.Fields("order_no").Value
                TextBox3.Text = cmd.Fields("dispatch_by").Value
                '  cmd.Fields("dispatch_through").Value = TextBox9.Text
                '  cmd.Fields("document_through").Value = TextBox4.Text
                TextBox22.Text = cmd.Fields("bakale_gst").Value
                TextBox30.Text = cmd.Fields("bank_name").Value
                TextBox29.Text = cmd.Fields("bank_account_no").Value
                TextBox11.Text = cmd.Fields("branch_name").Value
                TextBox27.Text = cmd.Fields("b_city").Value
                TextBox12.Text = cmd.Fields("ifsc_code").Value

                TextBox13.Text = cmd.Fields("total").Value
                TextBox14.Text = cmd.Fields("tax_rate").Value
                TextBox15.Text = cmd.Fields("tax_amount").Value
                TextBox16.Text = cmd.Fields("sgst").Value
                TextBox17.Text = cmd.Fields("cgst").Value
                TextBox18.Text = cmd.Fields("igst").Value
                TextBox19.Text = cmd.Fields("total_taxinclude").Value
                TextBox4.Text = cmd.Fields("sgst_amt").Value
                TextBox10.Text = cmd.Fields("cgst_amt").Value
                TextBox9.Text = cmd.Fields("igst_amt").Value

                cmd.MoveNext()
            End While
        Else
            MsgBox("Record not found.", vbCritical, "Message")
        End If
        cmd.Close()


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        DataGridView1.Rows.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox10.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""
        TextBox11.Text = ""
        TextBox28.Text = ""
        TextBox12.Text = ""
        TextBox27.Text = ""
        TextBox30.Text = ""
        TextBox29.Text = ""
        ComboBox1.Text = ""
    End Sub

    Private Sub search_bill_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        printcopy = False
        ComboBox1.Focus()
        srno = 0
        con.Open("dsn=bkl")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Label41.Text = AmtInWord((TextBox19.Text))
        Catch ex As Exception
            Label41.Text = ex.ToString
        End Try
        printfun()

    End Sub

    Private Sub printfun()
        Dim dt As New DataTable
        With dt
            .Columns.Add("sr_no")
            .Columns.Add("particular")
            .Columns.Add("hsn_no")
            .Columns.Add("size")
            .Columns.Add("pieces")
            .Columns.Add("pprate")
            .Columns.Add("total")
            .Columns.Add("bill_no")
            .Columns.Add("grand_total")
            .Columns.Add("buyer_name")
            .Columns.Add("buyer_details")
            .Columns.Add("buyer_state")
            .Columns.Add("buyer_gst_number")
            .Columns.Add("buyer_pan_number")
            .Columns.Add("buyer_tin_number")
            .Columns.Add("firm_address")
            .Columns.Add("transport_name")
            .Columns.Add("s_reference")
            .Columns.Add("order_number")
            .Columns.Add("dispatched_by")
            .Columns.Add("date_")
            .Columns.Add("bank_name")
            .Columns.Add("acc_number")
            .Columns.Add("ifsc_code")
            .Columns.Add("branch_name")
            .Columns.Add("city_")
            .Columns.Add("total_cost")
            .Columns.Add("tax_amt")
            .Columns.Add("sgst")
            .Columns.Add("cgst")
            .Columns.Add("igst")
            .Columns.Add("cgst_cost")
            .Columns.Add("sgst_cost")
            .Columns.Add("igst_cost")
            .Columns.Add("lr_number")
            .Columns.Add("bale_number")
            .Columns.Add("sizes")
            .Columns.Add("rndoff")
        End With
        '   line = 18
        '  n = 1
        For Each dr As DataGridViewRow In Me.DataGridView1.Rows
            dt.Rows.Add(dr.Cells(0).Value, dr.Cells(1).Value, dr.Cells(2).Value, dr.Cells(3).Value, dr.Cells(4).Value, dr.Cells(6).Value, dr.Cells(7).Value, TextBox20.Text, TextBox19.Text, ComboBox1.Text, TextBox1.Text, TextBox26.Text, TextBox25.Text, TextBox21.Text, TextBox28.Text, mdi.Label2.Text, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox3.Text, DateTimePicker2.Text, TextBox30.Text, TextBox29.Text, TextBox12.Text, TextBox11.Text, TextBox27.Text, TextBox13.Text, Me.Label41.Text, TextBox16.Text, TextBox17.Text, TextBox18.Text, TextBox4.Text, TextBox10.Text, TextBox9.Text, TextBox22.Text, TextBox3.Text, dr.Cells(3).Value, TextBox33.Text)
            '    line = line - 1
            '     n = n + 1
        Next

        '    While (line <> 0)
        ' dt.Rows.Add(n, " ", " ", " ", " ", " ", " ", TextBox20.Text, TextBox19.Text, ComboBox1.Text, TextBox1.Text, TextBox26.Text, TextBox25.Text, TextBox21.Text, TextBox2.Text, mdi.Label2.Text, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox3.Text, DateTimePicker2.Text, TextBox30.Text, TextBox29.Text, TextBox12.Text, TextBox11.Text, TextBox27.Text, TextBox13.Text, Me.Label41.Text, TextBox16.Text, TextBox17.Text, TextBox18.Text, TextBox4.Text, TextBox10.Text, TextBox9.Text, TextBox22.Text, TextBox3.Text, "", TextBox33.Text)
        ' line = line - 1
        ' n = n + 1
        ' End While



        Dim rptdoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptdoc = New CrystalReport1
        rptdoc.SetDataSource(dt)
        rptdoc.PrintToPrinter(1, False, 0, 0)



        '  Form1.CrystalReportViewer1.ReportSource = rptdoc
        '  Form1.ShowDialog()
        ' form1.dispose()

        ' temp_print.Show()

        ' Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If (DataGridView1.Rows.Count = 0) Then

            MsgBox("Please Add atleast one item to the cart", vbCritical, "Message")
        Else
            Dim y, w As New Integer
            Dim val1, val2 As New Integer
            val1 = Val(TextBox13.Text)
            val2 = Val(TextBox14.Text)
            w = 0
            For y = 0 To DataGridView1.RowCount - 1
                w = w + DataGridView1.Item("Column7", DataGridView1.Rows.Item(y).Index).Value
            Next y
            TextBox13.Text = w
            TextBox15.Text = 0.05 * Val(TextBox13.Text)
            TextBox14.Text = 5

            If TextBox26.Text.Contains("Maharashtra") Or TextBox26.Text.Contains("maharashtra") Then
                TextBox16.Text = 2.5
                TextBox17.Text = 2.5
                TextBox4.Text = Val(TextBox15.Text) / 2
                TextBox10.Text = Val(TextBox15.Text) / 2
                TextBox18.Text = "NA"
                TextBox9.Text = 0.0
                TextBox18.Enabled = False
                TextBox9.Enabled = False


            Else
                TextBox16.Text = "NA"
                TextBox17.Text = "NA"
                TextBox4.Text = 0.0
                TextBox10.Text = 0.0
                TextBox18.Text = 5
                TextBox9.Text = Val(TextBox15.Text)
                TextBox16.Enabled = False
                TextBox17.Enabled = False
                TextBox4.Enabled = False
                TextBox10.Enabled = False

            End If
            TextBox19.Text = "Rs." & Val(TextBox15.Text) + Val(TextBox13.Text)
        End If
    End Sub
End Class