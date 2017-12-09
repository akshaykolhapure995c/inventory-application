Public Class buyersprofile
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset
  
    Private Sub buyersprofile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub

    
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If (ComboBox1.Text = "" Or ComboBox2.Text = "" Or TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "") Then
            MsgBox("Please fill all fields", vbCritical, "Message")

        Else
            cmd.Open("select * from buyerinfo", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            cmd.AddNew()
            cmd.Fields("buyer_name").Value = TextBox1.Text
            cmd.Fields("contact_number").Value = TextBox2.Text
            cmd.Fields("address").Value = TextBox3.Text
            cmd.Fields("city").Value = ComboBox1.Text
            cmd.Fields("state").Value = ComboBox2.Text
            cmd.Fields("dis").Value = TextBox4.Text
            cmd.Fields("gst_number").Value = TextBox5.Text
            cmd.Fields("pan").Value = TextBox6.Text
            ' cmd.Fields("tin_number").Value = TextBox7.Text
            MsgBox("Added Successfully", MsgBoxStyle.Information, "Success!")

            cmd.Update()
            cmd.Close()

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            ' TextBox7.Text = ""
            ComboBox1.Text = ""
            ComboBox2.Text = ""
        End If
    End Sub


End Class