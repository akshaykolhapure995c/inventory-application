Public Class ledger_balance
    Dim con As New ADODB.Connection
    Dim cmd, cmd1 As New ADODB.Recordset
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim x As Integer
        x = 0
        DataGridView2.Rows.Clear()
        cmd.Open("select * from sales_report where name_='" & ComboBox2.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd1.Open("select * from balance where cust_name='" & ComboBox2.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF Or Not cmd1.EOF
            If Not cmd.EOF And cmd1.EOF Then
                DataGridView2.Rows.Add(cmd.Fields("date_").Value, cmd.Fields("bill_no").Value, cmd.Fields("gross").Value, "-", "-", "-")
                cmd.MoveNext()
            ElseIf Not cmd1.EOF And cmd.EOF Then
                DataGridView2.Rows.Add(cmd1.Fields("paid_date").Value, "-", "-", cmd1.Fields("paid").Value, cmd1.Fields("payment_type").Value, cmd1.Fields("cheque_number").Value)
                cmd1.MoveNext()
            Else
                If cmd.Fields("date_").Value = cmd1.Fields("paid_date").Value Then
                    DataGridView2.Rows.Add(cmd.Fields("date_").Value, cmd.Fields("bill_no").Value, cmd.Fields("gross").Value, cmd1.Fields("paid").Value)
                Else
                    DataGridView2.Rows.Add(cmd.Fields("date_").Value, cmd.Fields("bill_no").Value, cmd.Fields("gross").Value, "-", "-", "-")
                    DataGridView2.Rows.Add(cmd1.Fields("paid_date").Value, "-", "-", cmd1.Fields("paid").Value, cmd1.Fields("payment_type").Value, cmd1.Fields("cheque_number").Value)
                End If
                cmd.MoveNext()
                cmd1.MoveNext()
            End If
        End While
        cmd.Close()
        cmd1.Close()
        For i = 0 To DataGridView2.Rows.Count - 1
            If (IsNumeric(DataGridView2.Rows(i).Cells(2).Value)) Then
                x = x + DataGridView2.Rows(i).Cells(2).Value
            End If
        Next
        Label1.Text = x
        x = 0
        For i = 0 To DataGridView2.Rows.Count - 1
            If (IsNumeric(DataGridView2.Rows(i).Cells(3).Value)) Then
                x = x + DataGridView2.Rows(i).Cells(3).Value
            End If
        Next
        Label9.Text = x
        Label4.Text = Val(Label1.Text) - Val(Label9.Text)

        Label6.Text = "Amount To be Paid By " + ComboBox2.Text + " Is Rs."

    End Sub

    Private Sub ledger_balance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
        cmd.Open("select buyer_name from sales ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            ComboBox2.Items.Add(cmd.Fields("buyer_name").Value)
            ComboBox1.Items.Add(cmd.Fields("buyer_name").Value)
            cmd.MoveNext()
        End While
        cmd.Close()
        For i As Int16 = 0 To Me.ComboBox2.Items.Count - 2
            For j As Int16 = Me.ComboBox2.Items.Count - 1 To i + 1 Step -1
                If Me.ComboBox2.Items(i).ToString = Me.ComboBox2.Items(j).ToString Then
                    Me.ComboBox2.Items.RemoveAt(j)
                End If
            Next
        Next
        For i As Int16 = 0 To Me.ComboBox1.Items.Count - 2
            For j As Int16 = Me.ComboBox1.Items.Count - 1 To i + 1 Step -1
                If Me.ComboBox1.Items(i).ToString = Me.ComboBox1.Items(j).ToString Then
                    Me.ComboBox1.Items.RemoveAt(j)
                End If
            Next
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cmd.Open("select * from balance", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd.AddNew()
        cmd.Fields("paid_date").Value = DateTimePicker1.Value.Date
        cmd.Fields("cust_name").Value = ComboBox1.Text
        cmd.Fields("paid").Value = TextBox21.Text
        If RadioButton1.Checked Then
            cmd.Fields("payment_type").Value = RadioButton1.Text
        Else
            cmd.Fields("payment_type").Value = RadioButton2.Text
        End If
        cmd.Fields("cheque_number").Value = TextBox1.Text
        MsgBox("Submitted")
        cmd.Update()
        cmd.Close()

        cmd.Open("select * from bank_statement", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd.AddNew()
        cmd.Fields("paid_date").Value = DateTimePicker1.Value.Date
        cmd.Fields("cust_name").Value = ComboBox1.Text
        cmd.Fields("paid").Value = TextBox21.Text
        If RadioButton1.Checked Then
            cmd.Fields("payment_type").Value = RadioButton1.Text
        Else
            cmd.Fields("payment_type").Value = RadioButton2.Text
        End If
        cmd.Fields("cheque_number").Value = TextBox1.Text

        cmd.Update()
        cmd.Close()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        TextBox1.Text = ""
        TextBox21.Text = ""
        RadioButton1.Checked = False
        RadioButton2.Checked = False
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim x As Integer
        x = 0
        DataGridView2.Rows.Clear()
        cmd.Open("select * from sales_report where name_='" & ComboBox2.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd1.Open("select * from balance where cust_name='" & ComboBox2.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF Or Not cmd1.EOF
            If Not cmd.EOF And cmd1.EOF Then
                If cmd.Fields("date_").Value = cmd1.Fields("paid_date").Value Then
                    DataGridView2.Rows.Add(cmd.Fields("date_").Value, cmd.Fields("bill_no").Value, cmd.Fields("gross").Value, cmd1.Fields("paid").Value)
                    cmd.MoveNext()
                End If
            ElseIf Not cmd1.EOF And cmd.EOF Then
                DataGridView2.Rows.Add(cmd1.Fields("paid_date").Value, "-", "-", cmd1.Fields("paid").Value)
                cmd1.MoveNext()
            Else
                If cmd.Fields("date_").Value = cmd1.Fields("paid_date").Value Then
                    DataGridView2.Rows.Add(cmd.Fields("date_").Value, cmd.Fields("bill_no").Value, cmd.Fields("gross").Value, cmd1.Fields("paid").Value)
                Else
                    DataGridView2.Rows.Add(cmd.Fields("date_").Value, cmd.Fields("bill_no").Value, cmd.Fields("gross").Value, "-")
                    DataGridView2.Rows.Add(cmd1.Fields("paid_date").Value, "-", "-", cmd1.Fields("paid").Value)
                End If
                cmd.MoveNext()
                cmd1.MoveNext()
            End If
        End While
        cmd.Close()
        cmd1.Close()
        For i = 0 To DataGridView2.Rows.Count - 1
            If (IsNumeric(DataGridView2.Rows(i).Cells(2).Value)) Then
                x = x + DataGridView2.Rows(i).Cells(2).Value
            End If
        Next
        Label1.Text = x
        x = 0
        For i = 0 To DataGridView2.Rows.Count - 1
            If (IsNumeric(DataGridView2.Rows(i).Cells(3).Value)) Then
                x = x + DataGridView2.Rows(i).Cells(3).Value
            End If
        Next
        Label9.Text = x
        Label4.Text = Val(Label1.Text) - Val(Label9.Text)

        Label6.Text = "Amount To be Paid By " + ComboBox2.Text + "Is Rs."
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            TextBox1.Visible = True
            TextBox1.Text = ""
            TextBox1.Focus()
            Label10.Visible = True
        Else
            TextBox1.Visible = False
            Label10.Visible = False
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            TextBox1.Text = "-----"
            TextBox1.Visible = False
            Label10.Visible = False

        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked Then
            TextBox2.Text = "-----"
            TextBox2.Visible = False
            Label11.Visible = False

        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked Then
            TextBox2.Visible = True
            TextBox2.Text = ""
            TextBox2.Focus()
            Label11.Visible = True
        Else
            TextBox2.Visible = False
            Label11.Visible = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        cmd.Open("select * from bank_statement", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd.AddNew()
        cmd.Fields("paid_date").Value = DateTimePicker2.Value.Date
        cmd.Fields("cust_name").Value = ComboBox3.Text
        cmd.Fields("paid").Value = TextBox3.Text
        If RadioButton4.Checked Then
            cmd.Fields("payment_type").Value = RadioButton4.Text
        Else
            cmd.Fields("payment_type").Value = RadioButton3.Text
        End If
        cmd.Fields("cheque_number").Value = TextBox2.Text
        MsgBox("Submitted")
        cmd.Update()
        cmd.Close()

        ComboBox3.Text = ""
        TextBox3.Text = ""
        TextBox2.Text = ""
        RadioButton4.Checked = False
        RadioButton3.Checked = False
    End Sub
End Class