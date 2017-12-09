Public Class bank_details
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset

   
    Private Sub bank_details_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (TextBox10.Text = "" Or TextBox11.Text = "" Or TextBox27.Text = "" Or TextBox28.Text = "" Or TextBox12.Text = "") Then
            MsgBox("Please Fill all fields", vbCritical, "message")
        Else

            cmd.Open("select * from bankinfo", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            cmd.AddNew()
            cmd.MoveFirst()
            cmd.Fields("b_name").Value = Me.TextBox28.Text
            cmd.Fields("acc_number").Value = Me.TextBox10.Text
            cmd.Fields("branch_name").Value = Me.TextBox11.Text
            cmd.Fields("city").Value = Me.TextBox27.Text
            cmd.Fields("ifsc").Value = Me.TextBox12.Text

            MsgBox("Updated Successfully.")
            TextBox28.Text = ""
            TextBox10.Text = ""
            TextBox11.Text = ""
            TextBox27.Text = ""
            TextBox12.Text = ""

            cmd.Update()
            cmd.Close()
        End If
    End Sub
End Class