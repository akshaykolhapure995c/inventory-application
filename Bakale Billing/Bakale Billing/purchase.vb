Public Class purchase
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If (TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "") Then
            MsgBox("Please Fill All Fields ", vbCritical, "Message")




        Else





            cmd.Open("select * from purchase", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            cmd.AddNew()
            cmd.Fields("date_").Value = DateTimePicker1.Text
            cmd.Fields("bill_no").Value = TextBox2.Text
            cmd.Fields("name_").Value = TextBox1.Text
            cmd.Fields("gst").Value = TextBox3.Text
            cmd.Fields("hsn").Value = TextBox10.Text
            cmd.Fields("purchase").Value = TextBox4.Text
            cmd.Fields("igst").Value = TextBox5.Text
            cmd.Fields("sgst").Value = TextBox6.Text
            cmd.Fields("cgst").Value = TextBox7.Text
            cmd.Fields("round").Value = TextBox8.Text
            cmd.Fields("gross").Value = TextBox9.Text

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""

            MessageBox.Show("Saved Successfully")
            cmd.Update()
            cmd.Close()
        End If
    End Sub

    Private Sub RectangleShape1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RectangleShape1.Click

    End Sub

    Private Sub purchase_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub
End Class