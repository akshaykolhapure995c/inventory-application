Imports System.IO

Public Class billing
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

    




    Private Sub billing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        cmd.Open("select * from item_name ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Dim str, convertstr As String
        While Not cmd.EOF
            str = cmd.Fields("itemnm").Value
            convertstr = str.ToUpper()
            names.Add(convertstr)
            cmd.MoveNext()
        End While
        For a As Integer = 0 To names.Count - 1
            ComboBox2.Items.Add(names(a))
        Next
        cmd.Close()

        cmd2.Open("select * from bankinfo ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        TextBox30.Text = cmd2.Fields("b_name").Value.ToString
        TextBox29.Text = cmd2.Fields("acc_number").Value.ToString
        TextBox11.Text = cmd2.Fields("branch_name").Value.ToString
        TextBox27.Text = cmd2.Fields("city").Value.ToString
        TextBox12.Text = cmd2.Fields("ifsc").Value.ToString

        cmd2.Close()



    End Sub

    Private Sub ComboBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        ComboBox3.Text = ""
        ComboBox3.Items.Clear()
        cmd.Open("select * from item_name where itemnm = '" & ComboBox2.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            ComboBox3.Items.Add(cmd.Fields("size").Value)
            cmd.MoveNext()
        End While

        For i As Int16 = 0 To Me.ComboBox3.Items.Count - 2
            For j As Int16 = Me.ComboBox3.Items.Count - 1 To i + 1 Step -1
                If Me.ComboBox3.Items(i).ToString = Me.ComboBox3.Items(j).ToString Then
                    Me.ComboBox3.Items.RemoveAt(j)
                End If
            Next
        Next
        cmd.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If (Val(TextBox23.Text) > Val(TextBox24.Text)) Then
            MessageBox.Show("Entered Items are not available in stock!!")
            TextBox23.Focus()

        Else
            srno = srno + 1
            Dim a As Integer
            a = Val(TextBox6.Text) * Val(TextBox23.Text)
            TextBox24.Text = Val(TextBox24.Text) - Val(TextBox23.Text)
            '  DataGridView1.Rows.Add(srno, ComboBox2.Text, "62063000", ComboBox3.Text, ComboBox4.Text, TextBox23.Text, TextBox6.Text, a, TextBox24.Text, "X")
        End If
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
            Dim tot As New Single
            tot = Val(TextBox15.Text) + Val(TextBox13.Text)

            Dim z As New Double

            Dim x As Double = tot
            z = Convert.ToInt32(x)
            TextBox33.Text = z - tot
            If Val(TextBox33.Text) > 50 Then
                TextBox33.Text = " " + TextBox33.Text
            Else
                TextBox33.Text = " " + TextBox33.Text
            End If
            TextBox19.Text = z

        End If
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
                            printcopy = True
                            MsgBox("Are you Sure to Print Original Invoice ?", MsgBoxStyle.Question, "Message")
                            Dim str_item As String
                            Dim item_size, rmn_stock As New Integer
                            For Me.v = 0 To Me.DataGridView1.RowCount - 1
                                str_item = Me.DataGridView1(1, v).Value
                                item_size = Me.DataGridView1(3, v).Value
                                rmn_stock = Me.DataGridView1(7, v).Value
                                cmd.Open("select * from item_name where itemnm ='" & str_item & "' and size ='" & item_size & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                                'MsgBox(rmn_stock)
                                While Not cmd.EOF
                                    cmd.Fields("quantity").Value = rmn_stock
                                    cmd.MoveNext()
                                End While
                                cmd.Close()
                            Next

                            cmd.Open("select * from sales ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                            For Me.v = 0 To DataGridView1.RowCount - 1
                                cmd.AddNew()
                                cmd.Fields("bill_no").Value = TextBox20.Text
                                cmd.Fields("sr_no").Value = DataGridView1.Item("Column10", DataGridView1.Rows.Item(v).Index).Value
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
                            printfun()
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
        line = 18
        n = 1
        For Each dr As DataGridViewRow In Me.DataGridView1.Rows
            dt.Rows.Add(dr.Cells(0).Value, dr.Cells(1).Value, dr.Cells(2).Value, dr.Cells(3).Value, dr.Cells(4).Value, dr.Cells(6).Value, dr.Cells(7).Value, TextBox20.Text, TextBox19.Text, ComboBox1.Text, TextBox1.Text, TextBox26.Text, TextBox25.Text, TextBox21.Text, TextBox28.Text, mdi.Label2.Text, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox3.Text, DateTimePicker2.Text, TextBox30.Text, TextBox29.Text, TextBox12.Text, TextBox11.Text, TextBox27.Text, TextBox13.Text, Me.Label41.Text, TextBox16.Text, TextBox17.Text, TextBox18.Text, TextBox4.Text, TextBox10.Text, TextBox9.Text, TextBox22.Text, TextBox3.Text, dr.Cells(3).Value, TextBox33.Text)
            line = line - 1
            n = n + 1
        Next

        While (line <> 0)
            dt.Rows.Add(n, " ", " ", " ", " ", " ", " ", TextBox20.Text, TextBox19.Text, ComboBox1.Text, TextBox1.Text, TextBox26.Text, TextBox25.Text, TextBox21.Text, TextBox2.Text, mdi.Label2.Text, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox3.Text, DateTimePicker2.Text, TextBox30.Text, TextBox29.Text, TextBox12.Text, TextBox11.Text, TextBox27.Text, TextBox13.Text, Me.Label41.Text, TextBox16.Text, TextBox17.Text, TextBox18.Text, TextBox4.Text, TextBox10.Text, TextBox9.Text, TextBox22.Text, TextBox3.Text, "", TextBox33.Text)
            line = line - 1
            n = n + 1
        End While



        Dim rptdoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptdoc = New CrystalReport1
        rptdoc.SetDataSource(dt)
        rptdoc.PrintToPrinter(1, False, 0, 0)

        ' Form1.CrystalReportViewer1.ReportSource = rptdoc
        'Form1.ShowDialog()
        ' form1.dispose()

        ' temp_print.Show()

        ' Me.Close()
    End Sub


   


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim dt1 As New DataTable
        With dt1
            .Columns.Add("sr_no")
            .Columns.Add("particular")
            .Columns.Add("hsn_no")
            .Columns.Add("size")
            .Columns.Add("pieces")
            .Columns.Add("waist")
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

        End With
        line = 13
        For Each dr1 As DataGridViewRow In Me.DataGridView1.Rows
            dt1.Rows.Add(dr1.Cells(0).Value, dr1.Cells(1).Value, dr1.Cells(2).Value, dr1.Cells(3).Value, dr1.Cells(4).Value, dr1.Cells(5).Value, dr1.Cells(7).Value, TextBox20.Text, TextBox19.Text, ComboBox1.Text, TextBox1.Text, TextBox26.Text, TextBox25.Text, TextBox21.Text, TextBox2.Text, mdi.Label2.Text, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox3.Text, DateTimePicker2.Text, TextBox30.Text, TextBox29.Text, TextBox12.Text, TextBox11.Text, TextBox27.Text, TextBox13.Text, TextBox15.Text, TextBox16.Text, TextBox17.Text, TextBox18.Text)
            line = line - 1
        Next
        While (line <> 0)
            dt1.Rows.Add("", "", "", "", "", "", "") ', TextBox20.Text, TextBox19.Text, ComboBox1.Text, TextBox1.Text, TextBox26.Text, TextBox25.Text, TextBox21.Text, TextBox2.Text, mdi.Label2.Text, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox3.Text, DateTimePicker2.Text, TextBox30.Text, TextBox29.Text, TextBox12.Text, TextBox11.Text, TextBox27.Text, TextBox13.Text, TextBox15.Text, TextBox16.Text, TextBox17.Text, TextBox18.Text)
            line = line - 1

        End While

        Dim rptdoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptdoc = New CrystalReport2
        rptdoc.SetDataSource(dt1)
        rptdoc.PrintToPrinter(1, False, 0, 0)

        ' form1.CrystalReportviewer1.reportSource = rptdoc
        ' form1.showdialog()
        ' form1.dispose()

        ' temp_print.Show()

        ' Me.Close()

    End Sub
    Private Sub ComboBox1_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.LostFocus
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




    Private Sub textbox29_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        cmd.Open("select * from bankinfo where b_name = '" & TextBox30.Text & "' and acc_number='" & TextBox29.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            TextBox11.Text = cmd.Fields("branch_name").Value
            TextBox27.Text = cmd.Fields("city").Value
            TextBox12.Text = cmd.Fields("ifsc").Value
            cmd.MoveNext()
        End While
        cmd.Close()
    End Sub

    Private Sub ComboBox2_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.LostFocus
        ComboBox3.Items.Clear()
        ComboBox3.Text = ""
        cmd.Open("select * from item_name where itemnm = '" & ComboBox2.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            ComboBox3.Items.Add(cmd.Fields("size").Value)
            cmd.MoveNext()
        End While
        cmd.Close()

        For i As Int16 = 0 To Me.ComboBox3.Items.Count - 2
            For j As Int16 = Me.ComboBox3.Items.Count - 1 To i + 1 Step -1
                If Me.ComboBox3.Items(i).ToString = Me.ComboBox3.Items(j).ToString Then
                    Me.ComboBox3.Items.RemoveAt(j)
                End If
            Next
        Next


    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If (ComboBox2.Text = "" Or TextBox6.Text = "" Or TextBox24.Text = "" Or TextBox23.Text = "") Then
            'MessageBox.Show(, vbCritical, "Message")
            MsgBox("Please Fill All Fields of Order", vbCritical, "Message")
        Else

            If (Val(TextBox23.Text) > Val(TextBox24.Text)) Then
                MsgBox("Entered Items are not available in stock!!", vbCritical, "Alert")
                TextBox23.Focus()
            ElseIf (IsNumeric(TextBox23.Text) And IsNumeric(TextBox31.Text) And Val(TextBox23.Text) > 0 And Val(TextBox31.Text) >= 0 And Val(TextBox6.Text) > 0) Then
                Dim discount As Integer
                discount = (Val(TextBox31.Text) / 100) * Val(TextBox6.Text)


                Dim dis_price As Integer
                dis_price = TextBox6.Text - discount
                srno = srno + 1
                Dim a As Integer
                a = Val(TextBox6.Text) * Val(TextBox23.Text)
                TextBox24.Text = Val(TextBox24.Text) - Val(TextBox23.Text)
                DataGridView1.Rows.Add(srno, ComboBox2.Text, TextBox32.Text, "-", TextBox23.Text, TextBox6.Text, dis_price, dis_price * TextBox23.Text, TextBox24.Text, "E", "X")
                ComboBox2.Text = ""
                ComboBox3.Text = ""
                TextBox24.Text = ""
                TextBox6.Text = ""
                TextBox23.Text = ""
                TextBox32.Text = ""
                ComboBox2.Focus()
            Else
                MsgBox("Invalid Opearion !! Check Your Values", vbCritical, "Alert")
            End If
        End If

    End Sub


    Function strReplicate(ByVal str As String, ByVal intD As Integer) As String
        'This fucntion padded "0" after the number to evaluate hundred, thousand and on....
        'using this function you can replicate any Charactor with given string.
        Dim i As Integer
        strReplicate = ""
        For i = 1 To intD
            strReplicate = strReplicate + str
        Next
        Return strReplicate
    End Function
    Function AmtInWord(ByVal Num As Decimal) As String
        'I have created this function for converting amount in indian rupees (INR). 
        'You can manipulate as you wish like decimal setting, Doller (any currency) Prefix.

        Dim strNum As String
        Dim strNumDec As String
        Dim StrWord As String
        strNum = Num

        If InStr(1, strNum, ".") <> 0 Then
            strNumDec = Mid(strNum, InStr(1, strNum, ".") + 1)

            If Len(strNumDec) = 1 Then
                strNumDec = strNumDec + "0"
            End If
            If Len(strNumDec) > 2 Then
                strNumDec = Mid(strNumDec, 1, 2)
            End If

            strNum = Mid(strNum, 1, InStr(1, strNum, ".") - 1)
            StrWord = IIf(CDbl(strNum) = 1, " Rupee ", " Rupees ") + NumToWord(CDbl(strNum)) + IIf(CDbl(strNumDec) > 0, " and Paise" + cWord3(CDbl(strNumDec)), "")
        Else
            StrWord = IIf(CDbl(strNum) = 1, " Rupee ", " Rupees ") + NumToWord(CDbl(strNum))
        End If
        AmtInWord = StrWord & " Only"
        Return AmtInWord
    End Function
    Function NumToWord(ByVal Num As Decimal) As String
        'I divided this function in two part.
        '1. Three or less digit number.
        '2. more than three digit number.
        Dim strNum As String
        Dim StrWord As String
        strNum = Num

        If Len(strNum) <= 3 Then
            StrWord = cWord3(CDbl(strNum))
        Else
            StrWord = cWordG3(CDbl(Mid(strNum, 1, Len(strNum) - 3))) + " " + cWord3(CDbl(Mid(strNum, Len(strNum) - 2)))
        End If
        NumToWord = StrWord
    End Function
    Function cWordG3(ByVal Num As Decimal) As String
        '2. more than three digit number.
        Dim strNum As String = ""
        Dim StrWord As String = ""
        Dim readNum As String = ""
        strNum = Num
        If Len(strNum) Mod 2 <> 0 Then
            readNum = CDbl(Mid(strNum, 1, 1))
            If readNum <> "0" Then
                StrWord = retWord(readNum)
                readNum = CDbl("1" + strReplicate("0", Len(strNum) - 1) + "000")
                StrWord = StrWord + " " + retWord(readNum)
            End If
            strNum = Mid(strNum, 2)
        End If
        While Not Len(strNum) = 0
            readNum = CDbl(Mid(strNum, 1, 2))
            If readNum <> "0" Then
                StrWord = StrWord + " " + cWord3(readNum)
                readNum = CDbl("1" + strReplicate("0", Len(strNum) - 2) + "000")
                StrWord = StrWord + " " + retWord(readNum)
            End If
            strNum = Mid(strNum, 3)
        End While
        cWordG3 = StrWord
        Return cWordG3
    End Function
    Function cWord3(ByVal Num As Decimal) As String
        '1. Three or less digit number.
        Dim strNum As String = ""
        Dim StrWord As String = ""
        Dim readNum As String = ""
        If Num < 0 Then Num = Num * -1
        strNum = Num

        If Len(strNum) = 3 Then
            readNum = CDbl(Mid(strNum, 1, 1))
            StrWord = retWord(readNum) + " Hundred"
            strNum = Mid(strNum, 2, Len(strNum))
        End If

        If Len(strNum) <= 2 Then
            If CDbl(strNum) >= 0 And CDbl(strNum) <= 20 Then
                StrWord = StrWord + " " + retWord(CDbl(strNum))
            Else
                StrWord = StrWord + " " + retWord(CDbl(Mid(strNum, 1, 1) + "0")) + " " + retWord(CDbl(Mid(strNum, 2, 1)))
            End If
        End If

        strNum = CStr(Num)
        cWord3 = StrWord
        Return cWord3
    End Function

    Function retWord(ByVal Num As Decimal) As String
        'This two dimensional array store the primary word convertion of number.
        retWord = ""
        Dim ArrWordList(,) As Object = {{0, ""}, {1, "One"}, {2, "Two"}, {3, "Three"}, {4, "Four"}, _
                                        {5, "Five"}, {6, "Six"}, {7, "Seven"}, {8, "Eight"}, {9, "Nine"}, _
                                        {10, "Ten"}, {11, "Eleven"}, {12, "Twelve"}, {13, "Thirteen"}, {14, "Fourteen"}, _
                                        {15, "Fifteen"}, {16, "Sixteen"}, {17, "Seventeen"}, {18, "Eighteen"}, {19, "Nineteen"}, _
                                        {20, "Twenty"}, {30, "Thirty"}, {40, "Forty"}, {50, "Fifty"}, {60, "Sixty"}, _
                                        {70, "Seventy"}, {80, "Eighty"}, {90, "Ninety"}, {100, "Hundred"}, {1000, "Thousand"}, _
                                        {100000, "Lakh"}, {10000000, "Crore"}}

        Dim i As Integer
        For i = 0 To UBound(ArrWordList)
            If Num = ArrWordList(i, 0) Then
                retWord = ArrWordList(i, 1)
                Exit For
            End If
        Next
        Return retWord
    End Function

    

    ' Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    cmd.Open("select * from sales where bill_no = '" & TextBox28.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
    '   If Not cmd.EOF Then
    '      While Not cmd.EOF
    '
    '           cmd.Fields("bill_no").Value = TextBox20.Text
    '          Me.DataGridView1.Rows.Add(cmd.Fields("sr_no").Value, cmd.Fields("particulars").Value, cmd.Fields("hsn").Value, cmd.Fields("size").Value, cmd.Fields("piece_qty").Value, cmd.Fields("per_piece_rate").Value, cmd.Fields("item_amount").Value, cmd.Fields("remaining_stock").Value, "X")

    '         ComboBox1.Text = cmd.Fields("buyer_name").Value
    '        TextBox1.Text = cmd.Fields("buyer_address").Value
    '       TextBox26.Text = cmd.Fields("buyer_state").Value
    '      TextBox25.Text = cmd.Fields("buyer_gst").Value
    '     TextBox2.Text = cmd.Fields("buyer_tin").Value
    ''    TextBox21.Text = cmd.Fields("buyer_pan").Value
    '  DateTimePicker2.Text = cmd.Fields("order_date").Value
    '  TextBox5.Text = cmd.Fields("transport_name").Value
    ' TextBox7.Text = cmd.Fields("supplier_ref").Value
    'TextBox8.Text = cmd.Fields("order_no").Value
    'TextBox3.Text = cmd.Fields("dispatch_by").Value
    '  cmd.Fields("dispatch_through").Value = TextBox9.Text
    '  cmd.Fields("document_through").Value = TextBox4.Text
    '  TextBox22.Text = cmd.Fields("bakale_gst").Value
    '  TextBox30.Text = cmd.Fields("bank_name").Value
    '  TextBox29.Text = cmd.Fields("bank_account_no").Value
    '  TextBox11.Text = cmd.Fields("branch_name").Value
    '  TextBox27.Text = cmd.Fields("b_city").Value
    ' TextBox12.Text = cmd.Fields("ifsc_code").Value
    '
    '               TextBox13.Text = cmd.Fields("total").Value
    '                TextBox14.Text = cmd.Fields("tax_rate").Value
    '     TextBox15.Text = cmd.Fields("tax_amount").Value
    '    TextBox16.Text = cmd.Fields("sgst").Value
    '     TextBox17.Text = cmd.Fields("cgst").Value
    '     TextBox18.Text = cmd.Fields("igst").Value
    ''     TextBox19.Text = cmd.Fields("total_taxinclude").Value
    '     cmd.MoveNext()
    'End While
    'Else
    '    MsgBox("Record not found.", vbCritical, "Message")
    'End If
    'cmd.Close()


    '   End Sub
    '
    'Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    DataGridView1.Rows.Clear()
    '    TextBox1.Text = ""
    '    TextBox2.Text = ""
    '    TextBox3.Text = ""
    '    TextBox4.Text = ""
    '    TextBox5.Text = ""
    '   TextBox6.Text = ""
    '   TextBox7.Text = ""
    '   TextBox8.Text = ""
    ''   TextBox9.Text = ""
    '   TextBox13.Text = ""
    '   TextBox14.Text = ""
    '   TextBox15.Text = ""
    '   TextBox16.Text = ""
    '   TextBox17.Text = ""
    '   TextBox18.Text = ""
    '   TextBox19.Text = ""
    '  TextBox20.Text = ""
    '  TextBox10.Text = ""
    ''  TextBox21.Text = ""
    '  TextBox22.Text = ""
    '  TextBox23.Text = ""
    '  TextBox24.Text = ""
    '  TextBox25.Text = ""
    '  TextBox26.Text = ""
    ' TextBox11.Text = ""
    ' TextBox28.Text = ""
    ' TextBox12.Text = ""
    ' TextBox27.Text = ""
    ' TextBox30.Text = ""
    ' TextBox29.Text = ""
    '      ComboBox1.Text = ""
    '   End'' Sub
    '
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

        If (DataGridView1.Item("Column5", DataGridView1.CurrentRow.Index).Selected) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        End If

    End Sub

    Private Sub TextBox6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.GotFocus
        cmd.Open("select * from item_name where itemnm= '" & ComboBox2.Text & "' and size= '" & ComboBox3.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            TextBox24.Text = cmd.Fields("quantity").Value
            TextBox6.Text = cmd.Fields("per_piece_rate").Value
            '  TextBox32.Text = cmd.Fields("hsn").Value
            cmd.MoveNext()
        End While
        TextBox24.Enabled = False

        cmd.Close()
    End Sub

    

    Private Sub TextBox23_PreviewKeyDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles TextBox23.PreviewKeyDown
        If e.KeyCode = Keys.Enter Then
            If (ComboBox2.Text = "" Or TextBox6.Text = "" Or TextBox24.Text = "" Or TextBox23.Text = "") Then
                'MessageBox.Show(, vbCritical, "Message")
                MsgBox("Please Fill All Fields of Order", vbCritical, "Message")
            Else

                If (Val(TextBox23.Text) > Val(TextBox24.Text)) Then
                    MsgBox("Entered Items are not available in stock!!", vbCritical, "Alert")
                    TextBox23.Focus()
                ElseIf (IsNumeric(TextBox23.Text) And IsNumeric(TextBox31.Text)) Then
                    Dim discount As Integer
                    discount = (Val(TextBox31.Text) / 100) * Val(TextBox6.Text)


                    Dim dis_price As Integer
                    dis_price = TextBox6.Text - discount
                    srno = srno + 1
                    Dim a As Integer
                    a = Val(TextBox6.Text) * Val(TextBox23.Text)
                    TextBox24.Text = Val(TextBox24.Text) - Val(TextBox23.Text)
                    DataGridView1.Rows.Add(srno, ComboBox2.Text, TextBox32.Text, "-", TextBox23.Text, TextBox6.Text, dis_price, dis_price * TextBox23.Text, TextBox24.Text, "E", "X")
                    ComboBox2.Text = ""
                    ComboBox3.Text = ""
                    TextBox24.Text = ""
                    TextBox6.Text = ""
                    TextBox23.Text = ""
                    TextBox32.Text = ""
                    ComboBox2.Focus()
                Else
                    MsgBox("Please Enter Numeric Value For Pieces", vbCritical, "Alert")
                End If
            End If
        End If

    End Sub


    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        DataGridView1.Rows.Clear()
        TextBox1.Text = ""
        TextBox33.Text = ""
        TextBox31.Text = ""
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

    Private Sub ComboBox3_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.LostFocus
        cmd.Open("select * from item_name where itemnm= '" & ComboBox2.Text & "' and size= '" & ComboBox3.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            TextBox24.Text = cmd.Fields("quantity").Value
            TextBox6.Text = cmd.Fields("per_piece_rate").Value
            TextBox32.Text = cmd.Fields("hsn").Value

            cmd.MoveNext()
        End While
        cmd.Close()
    End Sub

    
End Class
