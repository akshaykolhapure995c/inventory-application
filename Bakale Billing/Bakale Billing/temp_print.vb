Public Class temp_print
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset

    Private Sub temp_print_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")

        If billing.DataGridView1.RowCount = 0 Then
            MsgBox("No Item")
            Button1.Enabled = False
        Else
            For v = 0 To billing.DataGridView1.RowCount - 1
                Me.ListBox1.Items.Add(billing.DataGridView1.Item("Column1", billing.DataGridView1.Rows.Item(v).Index).Value)
                Me.ListBox2.Items.Add(billing.DataGridView1.Item("Column2", billing.DataGridView1.Rows.Item(v).Index).Value)
                Me.ListBox3.Items.Add(billing.DataGridView1.Item("Column3", billing.DataGridView1.Rows.Item(v).Index).Value)
                Me.ListBox4.Items.Add(billing.DataGridView1.Item("Column5", billing.DataGridView1.Rows.Item(v).Index).Value)
                Me.ListBox5.Items.Add(billing.DataGridView1.Item("Column6", billing.DataGridView1.Rows.Item(v).Index).Value)
                Me.ListBox6.Items.Add(billing.DataGridView1.Item("Column7", billing.DataGridView1.Rows.Item(v).Index).Value)
                Me.ListBox7.Items.Add(billing.DataGridView1.Item("Column10", billing.DataGridView1.Rows.Item(v).Index).Value)
            Next v


        End If
        Me.Label13.Text = billing.TextBox15.Text
        Me.Label12.Text = billing.TextBox16.Text
        Me.Label11.Text = billing.TextBox17.Text
        Me.Label10.Text = billing.TextBox18.Text
        Me.Label15.Text = billing.TextBox13.Text
        Me.TextBox1.Text = billing.TextBox1.Text
        Me.Label62.Text = billing.TextBox26.Text
        Me.Label63.Text = billing.TextBox25.Text
        Me.Label61.Text = billing.TextBox21.Text
        Me.Label43.Text = billing.TextBox20.Text
        Me.Label44.Text = billing.DateTimePicker2.Text
        Me.Label45.Text = billing.TextBox8.Text
        '   Me.Label46.Text = billing.TextBox4.Text
        '  Me.Label47.Text = billing.TextBox9.Text
        Me.Label48.Text = billing.TextBox26.Text
        Me.Label49.Text = billing.TextBox5.Text
        Me.Label50.Text = billing.DateTimePicker2.Text
        Me.Label51.Text = billing.TextBox7.Text
        Me.Label14.Text = billing.TextBox19.Text


       
       
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim str_item As String
        Dim item_size, rmn_stock As New Integer
        For v = 0 To billing.DataGridView1.RowCount - 1
            str_item = billing.DataGridView1(1, v).Value
            item_size = billing.DataGridView1(3, v).Value
            rmn_stock = billing.DataGridView1(8, v).Value
            cmd.Open("select * from item_name where itemnm ='" & str_item & "' and size ='" & item_size & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            MsgBox(rmn_stock)
            While Not cmd.EOF
                cmd.Fields("quantity").Value = rmn_stock
                cmd.MoveNext()
            End While
            cmd.Close()
            'MsgBox((billing.DataGridView1.Item("Column2", billing.DataGridView1.Rows.Item(v).Index).Value))
        Next



        Me.PrintForm1.Print()
        ' Me.Close()
        ' billing.Close()
    End Sub

    Private Sub RectangleShape1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RectangleShape1.Click

        billing.DataGridView1.Rows.Clear()
        billing.TextBox1.Text = ""
        billing.TextBox2.Text = ""
        billing.TextBox3.Text = ""
        ' billing.TextBox4.Text = ""
        billing.TextBox5.Text = ""
        billing.TextBox6.Text = ""
        billing.TextBox7.Text = ""
        billing.TextBox8.Text = ""
        ' billing.TextBox9.Text = ""
        billing.TextBox13.Text = ""
        billing.TextBox14.Text = ""
        billing.TextBox15.Text = ""
        billing.TextBox16.Text = ""
        billing.TextBox17.Text = ""
        billing.TextBox18.Text = ""
        billing.TextBox19.Text = ""
        billing.TextBox20.Text = ""
        billing.TextBox21.Text = ""
        billing.TextBox22.Text = ""
        billing.TextBox23.Text = ""
        billing.TextBox24.Text = ""
        billing.TextBox25.Text = ""
        billing.TextBox26.Text = ""



    End Sub
End Class