Public Class update_buyer
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset
    Private Sub update_buyer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
        cmd.Open("select * from buyerinfo ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            ComboBox3.Items.Add(cmd.Fields("buyer_name").Value)
            cmd.MoveNext()
        End While
        cmd.Close()
    End Sub


   
    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        cmd.Open("select * from buyerinfo where buyer_name = '" & ComboBox3.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            TextBox3.Text = cmd.Fields("address").Value
            ComboBox2.Text = cmd.Fields("state").Value
            ComboBox1.Text = cmd.Fields("city").Value
            TextBox5.Text = cmd.Fields("gst_number").Value
            'TextBox2.Text = cmd.Fields("tin_number").Value
            TextBox6.Text = cmd.Fields("pan").Value
            TextBox4.Text = cmd.Fields("dis").Value
            TextBox2.Text = cmd.Fields("contact_number").Value
            cmd.MoveNext()
        End While

        cmd.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cmd.Open("select * from buyerinfo where buyer_name = '" & ComboBox3.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            cmd.Fields("buyer_name").Value = ComboBox3.Text
            cmd.Fields("address").Value = TextBox3.Text
            cmd.Fields("state").Value = ComboBox2.Text
            cmd.Fields("city").Value = ComboBox1.Text
            cmd.Fields("gst_number").Value = TextBox5.Text
            'TextBox2.Text = cmd.Fields("tin_number").Value
            cmd.Fields("pan").Value = TextBox6.Text
            cmd.Fields("dis").Value = TextBox4.Text
            cmd.Fields("contact_number").Value = TextBox2.Text
            cmd.MoveNext()
        End While
        cmd.Close()
        MsgBox("Updated Succesfully")
        TextBox3.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        ComboBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        TextBox2.Text = ""

    End Sub
End Class