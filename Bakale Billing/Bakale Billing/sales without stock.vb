Public Class sales_without_stock
    Dim con As New ADODB.Connection
    Dim cmd, cmd2 As New ADODB.Recordset
    Dim v As New Integer
    Dim srno As New Integer
    Dim names As New HashSet(Of String)
    Dim id As Integer
    Dim printcopy As Boolean
    Dim line As Integer
    Dim count As Integer = 0
    Dim n As Integer
    Private Sub sales_without_stock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        printcopy = False
        ComboBox1.Focus()
        srno = 0
        con.Open("dsn=bkl")

        cmd.Open("select bill_no from sales ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        If Not cmd.EOF() Then
            cmd.MoveLast()
            TextBox20.Text = cmd.Fields("bill_no").Value + 1
        Else
            TextBox20.Text = 1
        End If

        cmd.Close()


        cmd.Open("select * from buyerinfo ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            ComboBox1.Items.Add(cmd.Fields("buyer_name").Value)
            cmd.MoveNext()
        End While
        cmd.Close()
        

        cmd2.Open("select * from bankinfo ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        TextBox30.Text = cmd2.Fields("b_name").Value.ToString
        TextBox29.Text = cmd2.Fields("acc_number").Value.ToString
        TextBox11.Text = cmd2.Fields("branch_name").Value.ToString
        TextBox27.Text = cmd2.Fields("city").Value.ToString
        TextBox12.Text = cmd2.Fields("ifsc").Value.ToString

        cmd2.Close()



    End Sub
    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim a, amount As String
        If (DataGridView1.Item("Column9", DataGridView1.CurrentRow.Index).Selected) Then
n:          a = InputBox("Please Enter Changed Quantity")
            If (a = "" Or Not IsNumeric(Val(a)) Or a <= 0) Then
                MsgBox("Please Enter Valid Quantity", vbCritical, "Invalid Quantity Value")
                GoTo n

            End If
            amount = a * DataGridView1.Item("Column6", DataGridView1.CurrentRow.Index).Value
            DataGridView1.Item("Column4", DataGridView1.CurrentRow.Index).Value = a
            DataGridView1.Item("Column7", DataGridView1.CurrentRow.Index).Value = amount
            DataGridView1.Item("Column8", DataGridView1.CurrentRow.Index).Value = DataGridView1.Item("Column8", DataGridView1.CurrentRow.Index).Value - a
        End If

        If (DataGridView1.Item("Column10", DataGridView1.CurrentRow.Index).Selected) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        End If

    End Sub

    Private Sub TextBox23_PreviewKeyDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles TextBox23.PreviewKeyDown
        If e.KeyCode = Keys.Enter Then
            If (ComboBox2.Text = "" Or ComboBox3.Text = "" Or TextBox6.Text = "" Or TextBox23.Text = "") Then
                'MessageBox.Show(, vbCritical, "Message")
                MsgBox("Please Fill All Fields of Order", vbCritical, "Message")
            Else
                If (IsNumeric(TextBox23.Text) And IsNumeric(TextBox31.Text) And Val(TextBox23.Text) > 0 And Val(TextBox31.Text) >= 0 And Val(TextBox6.Text) > 0) Then
                    Dim discount As Double
                    discount = (Val(TextBox31.Text) / 100) * Val(TextBox6.Text)


                    Dim dis_price As Double
                    dis_price = TextBox6.Text - discount
                    Dim decimalpoints As Double
                    decimalpoints = Math.Abs(dis_price - Math.Floor(dis_price))
                    Dim a, b, u As Double
                    If (decimalpoints >= 0.5) Then
                        dis_price = Math.Ceiling(dis_price)
                    Else
                        dis_price = Math.Round(dis_price)
                    End If
                    srno = srno + 1

                    a = Val(TextBox6.Text) * Val(TextBox23.Text)
                    b = (dis_price * TextBox23.Text)
                    DataGridView1.Rows.Add(srno, ComboBox2.Text, TextBox32.Text, ComboBox3.Text, Val(TextBox23.Text), TextBox6.Text + ".00", dis_price, b.ToString + ".00", "", "E", "X")
                    ComboBox2.Text = ""


                    TextBox6.Text = ""
                    TextBox23.Text = ""
                    TextBox32.Text = ""
                    ComboBox2.Focus()
                Else
                    MsgBox("Invalid Opearion !! Check Your Values", vbCritical, "Alert")
                End If
            End If
        End If

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If (ComboBox2.Text = "" Or ComboBox3.Text = "" Or TextBox6.Text = "" Or TextBox23.Text = "") Then
            'MessageBox.Show(, vbCritical, "Message")
            MsgBox("Please Fill All Fields of Order", vbCritical, "Message")
        Else
            If (IsNumeric(TextBox23.Text) And IsNumeric(TextBox31.Text) And Val(TextBox23.Text) > 0 And Val(TextBox31.Text) >= 0 And Val(TextBox6.Text) > 0) Then
                Dim discount As Double
                discount = (Val(TextBox31.Text) / 100) * Val(TextBox6.Text)


                Dim dis_price As Double
                dis_price = TextBox6.Text - discount
                Dim decimalpoints As Double
                decimalpoints = Math.Abs(dis_price - Math.Floor(dis_price))
                Dim a, b, u As Double
                If (decimalpoints >= 0.5) Then
                    dis_price = Math.Ceiling(dis_price)
                Else
                    dis_price = Math.Round(dis_price)
                End If
                srno = srno + 1

                a = Val(TextBox6.Text) * Val(TextBox23.Text)
                b = (dis_price * TextBox23.Text)
                DataGridView1.Rows.Add(srno, ComboBox2.Text, TextBox32.Text, ComboBox3.Text, Val(TextBox23.Text), TextBox6.Text + ".00", dis_price, b.ToString + ".00", "", "E", "X")
                ComboBox2.Text = ""


                TextBox6.Text = ""
                TextBox23.Text = ""
                TextBox32.Text = ""
                ComboBox2.Focus()
            Else
                MsgBox("Invalid Opearion !! Check Your Values", vbCritical, "Alert")
            End If
        End If
    End Sub

    Private Sub ComboBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.LostFocus
        cmd.Open("select * from buyerinfo where buyer_name = '" & ComboBox1.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            TextBox1.Text = cmd.Fields("address").Value
            TextBox26.Text = cmd.Fields("state").Value
            TextBox25.Text = cmd.Fields("gst_number").Value
            'TextBox2.Text = cmd.Fields("tin_number").Value
            TextBox21.Text = cmd.Fields("pan").Value
            TextBox31.Text = cmd.Fields("dis").Value
            TextBox28.Text = cmd.Fields("contact_number").Value
            cmd.MoveNext()
        End While

        cmd.Close()
        TextBox31.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        

        If (DataGridView1.Rows.Count = 0) Then

            MsgBox("Please Add atleast one item to the cart", vbCritical, "Message")
        Else
            Dim y, w, p As New Integer
            Dim val1, val2 As New Integer
            val1 = Val(TextBox13.Text)
            val2 = Val(TextBox14.Text)
            w = 0
            p = 0
            For y = 0 To DataGridView1.RowCount - 1
                w = w + DataGridView1.Item("Column7", DataGridView1.Rows.Item(y).Index).Value
                p = p + DataGridView1.Item("Column4", DataGridView1.Rows.Item(y).Index).Value
            Next y
            TextBox13.Text = w.ToString + ".00"
            TextBox15.Text = 0.05 * Val(TextBox13.Text)
            TextBox15.Text = Math.Round(Val(TextBox15.Text), 2)
            TextBox14.Text = 5
            TextBox24.Text = p

            If TextBox26.Text.Contains("Maharashtra") Or TextBox26.Text.Contains("maharashtra") Then
                TextBox16.Text = 2.5
                TextBox17.Text = 2.5
                TextBox4.Text = Val(TextBox15.Text) / 2 + ".00"
                TextBox4.Text = Math.Round(Val(TextBox4.Text), 2)
                TextBox10.Text = Val(TextBox15.Text) / 2 + ".00"
                TextBox10.Text = Math.Round(Val(TextBox10.Text), 2)
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
                TextBox9.Text = Math.Round(Val(TextBox9.Text), 2)
                TextBox16.Enabled = False
                TextBox17.Enabled = False
                TextBox4.Enabled = False
                TextBox10.Enabled = False

            End If
            Dim tot As New Single
            tot = Val(TextBox15.Text) + Val(TextBox13.Text).ToString + ".00"

            Dim z As New Double

            Dim x As Double = tot
            z = Convert.ToInt32(x)
            TextBox33.Text = z - tot
            TextBox33.Text = Math.Round(Val(TextBox33.Text), 2)
            If Val(TextBox33.Text) > 50 Then
                TextBox33.Text = " " + TextBox33.Text
            Else
                TextBox33.Text = " " + TextBox33.Text
            End If
            TextBox19.Text = z.ToString + ".00"
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ' DataGridView1.Rows.Clear()
        '   TextBox1.Text = ""
        '  TextBox31.Text = ""
        '  TextBox2.Text = ""
        '  TextBox3.Text = ""
        ' TextBox4.Text = ""
        'TextBox5.Text = ""
        ' TextBox6.Text = ""
        ' TextBox7.Text = ""
        '  TextBox8.Text = ""
        '  TextBox9.Text = ""
        '  TextBox13.Text = ""
        ' TextBox14.Text = ""
        ' TextBox15.Text = ""
        ' TextBox16.Text = ""
        ' TextBox17.Text = ""
        '  TextBox18.Text = ""
        ' TextBox19.Text = ""
        ' TextBox20.Text = ""
        ' TextBox10.Text = ""
        ' TextBox21.Text = ""
        ' TextBox22.Text = ""
        ' TextBox23.Text = ""
        ''  TextBox24.Text = ""
        ' TextBox25.Text = ""
        ' TextBox26.Text = ""
        ' TextBox11.Text = ""
        'TextBox28.Text = ""
        ' TextBox12.Text = ""
        'TextBox27.Text = ""
        ' TextBox30.Text = ""
        ' TextBox29.Text = ""
        'ComboBox1.Text = ""


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Label41.Text = AmtInWord((TextBox19.Text))
        Catch ex As Exception
            Label41.Text = ex.ToString
        End Try

        If (TextBox13.Text = "") Then
            MsgBox("Please Make total then try again.", vbCritical, "Message")
        Else
            If (TextBox26.Text = "" Or ComboBox1.Text = "") Then
                MsgBox("Please select Buyer's Profile", vbCritical, "Message")
            Else
                If (TextBox5.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox3.Text = "") Then
                    MsgBox("Please Fill all Fields of other details", vbCritical, "Message")
                Else
                    If (TextBox30.Text = "" Or TextBox29.Text = "") Then
                        MsgBox("Please select bank details", vbCritical, "Message")
                    Else
                        If RadioButton1.Checked Then
                            If (printcopy = False) Then
                                printcopy = True
                                MsgBox("Are you Sure to Print Original Invoice ?", MsgBoxStyle.Question, "Message")

                                '  Dim str_item As String
                                ' Dim item_size, rmn_stock As New Integer
                                ' For Me.v = 0 To Me.DataGridView1.RowCount - 1
                                '   str_item = Me.DataGridView1(1, v).Value
                                '  item_size = Me.DataGridView1(3, v).Value
                                ' rmn_stock = Me.DataGridView1(7, v).Value
                                '   cmd.Open("select * from item_name where itemnm ='" & str_item & "' and size ='" & item_size & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                'MsgBox(rmn_stock)
                                '  While Not cmd.EOF
                                '   cmd.Fields("quantity").Value = rmn_stock
                                '  cmd.MoveNext()
                                ' End While
                                '  cmd.Close()
                                ' Next
                              
                                cmd.Open("select * from sales ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                                For Me.v = 0 To DataGridView1.RowCount - 1
                                    cmd.AddNew()
                                    cmd.Fields("bill_no").Value = TextBox20.Text
                                    cmd.Fields("sr_no").Value = DataGridView1.Item("Column0", DataGridView1.Rows.Item(v).Index).Value
                                    cmd.Fields("particulars").Value = DataGridView1.Item("Column1", DataGridView1.Rows.Item(v).Index).Value
                                    cmd.Fields("hsn").Value = DataGridView1.Item("Column2", DataGridView1.Rows.Item(v).Index).Value
                                    cmd.Fields("size").Value = DataGridView1.Item("Column3", DataGridView1.Rows.Item(v).Index).Value
                                    cmd.Fields("piece_qty").Value = DataGridView1.Item("Column4", DataGridView1.Rows.Item(v).Index).Value
                                    cmd.Fields("per_piece_rate").Value = DataGridView1.Item("Column5", DataGridView1.Rows.Item(v).Index).Value
                                    cmd.Fields("discounted_price").Value = DataGridView1.Item("Column6", DataGridView1.Rows.Item(v).Index).Value

                                    cmd.Fields("item_amount").Value = DataGridView1.Item("Column7", DataGridView1.Rows.Item(v).Index).Value

                                    cmd.Fields("remaining_stock").Value = DataGridView1.Item("Column8", DataGridView1.Rows.Item(v).Index).Value


                                    cmd.Fields("buyer_name").Value = ComboBox1.Text
                                    cmd.Fields("buyer_address").Value = TextBox1.Text
                                    cmd.Fields("buyer_state").Value = TextBox26.Text
                                    cmd.Fields("buyer_gst").Value = TextBox25.Text
                                    cmd.Fields("buyer_tin").Value = TextBox2.Text
                                    cmd.Fields("buyer_pan").Value = TextBox21.Text
                                    cmd.Fields("order_date").Value = DateTimePicker2.Text
                                    cmd.Fields("transport_name").Value = TextBox5.Text
                                    cmd.Fields("supplier_ref").Value = TextBox7.Text
                                    cmd.Fields("order_no").Value = TextBox8.Text
                                    cmd.Fields("dispatch_by").Value = TextBox3.Text
                                    '  cmd.Fields("dispatch_through").Value = TextBox9.Text
                                    '  cmd.Fields("document_through").Value = TextBox4.Text
                                    cmd.Fields("bakale_gst").Value = TextBox22.Text
                                    cmd.Fields("bank_name").Value = TextBox30.Text
                                    cmd.Fields("bank_account_no").Value = TextBox29.Text
                                    cmd.Fields("branch_name").Value = TextBox11.Text
                                    cmd.Fields("b_city").Value = TextBox27.Text
                                    cmd.Fields("ifsc_code").Value = TextBox12.Text
                                    cmd.Fields("total").Value = TextBox13.Text
                                    cmd.Fields("tax_rate").Value = TextBox14.Text
                                    cmd.Fields("tax_amount").Value = TextBox15.Text
                                    cmd.Fields("sgst").Value = TextBox16.Text
                                    cmd.Fields("cgst").Value = TextBox17.Text
                                    cmd.Fields("igst").Value = TextBox18.Text
                                    cmd.Fields("total_taxinclude").Value = TextBox19.Text
                                    cmd.Fields("sgst_amt").Value = TextBox4.Text
                                    cmd.Fields("cgst_amt").Value = TextBox10.Text
                                    cmd.Fields("igst_amt").Value = TextBox9.Text
                                    cmd.MoveNext()
                                Next v
                                cmd.Update()
                                cmd.Close()
                                cmd.Open("select * from sales_report ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                cmd.AddNew()
                                cmd.Fields("date_").Value = DateTimePicker2.Text
                                cmd.Fields("bill_no").Value = TextBox20.Text
                                cmd.Fields("name_").Value = ComboBox1.Text
                                cmd.Fields("gst").Value = TextBox25.Text
                                cmd.Fields("hsn").Value = ""
                                cmd.Fields("sales").Value = TextBox13.Text
                                cmd.Fields("igst").Value = TextBox9.Text
                                cmd.Fields("sgst").Value = TextBox4.Text
                                cmd.Fields("cgst").Value = TextBox10.Text
                                cmd.Fields("round").Value = TextBox33.Text
                                cmd.Fields("gross").Value = TextBox19.Text
                                cmd.Fields("place_of_supply").Value = TextBox26.Text
                                cmd.Fields("invoice_type").Value = "Regular"
                                cmd.Fields("cess_amount").Value = "NA"
                                cmd.Update()
                                cmd.Close()
                                printfun()
                            Else
                                MsgBox("Original Copy Already printed ! Please Select Duplicate Copy ", vbCritical)
                            End If
                        ElseIf RadioButton2.Checked = True Then
                            If printcopy = False Then
                                MsgBox("Please print original invoice copy first.", vbCritical, "Message")
                                RadioButton1.Focus()
                            Else

                                printfun()
                            End If

                        ElseIf RadioButton3.Checked = True Then
                            If printcopy = False Then
                                MsgBox("Please print original invoice copy first.", vbCritical, "Message")
                                RadioButton1.Focus()
                            Else
                                MsgBox("select radio 3")
                                printfun()
                            End If

                        Else : MsgBox("Select copy type to print invoice.")
                        End If
                        End If
                End If
            End If
        End If


        Me.Focus()


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

    Private Sub ComboBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.LostFocus
        ComboBox3.Text = "38 X 44"
        TextBox32.Text = "6205"
        TextBox6.Focus()
    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click

    End Sub

   
    
    Private Sub RectangleShape6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RectangleShape6.Click

    End Sub
End Class