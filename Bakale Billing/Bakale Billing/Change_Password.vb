Public Class Change_Password
    Dim con As New ADODB.Connection
    Dim cmd, cmd2 As New ADODB.Recordset
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = ""
        cmd.Open("select * from  other", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If Not TextBox2.Text = TextBox3.Text Then
            MsgBox("Password mismatch.")
        End If
        If TextBox2.Text = TextBox3.Text Then
            cmd.Fields("pw").Value = StrReverse(TextBox3.Text)
            MsgBox("Password changed successfully.", MsgBoxStyle.Information, "Success!")
            Me.Close()
        End If
        cmd.Update()
        cmd.Close()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub

    Private Sub Change_Password_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Enabled = False
        mdi.Enabled = True
    End Sub

    Private Sub Change_Password_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub

    Private Sub TextBox2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.GotFocus
      
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        cmd2.Open("select * from  other", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        If Not cmd2.Fields("pw").Value = StrReverse(TextBox1.Text) Then
            MsgBox("Please enter valid password.")
        End If
        cmd2.Close()
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class