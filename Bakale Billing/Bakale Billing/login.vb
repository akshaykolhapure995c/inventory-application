Public Class login
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cmd.Open("select * from other ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd.MoveFirst()
        If (TextBox2.Text = StrReverse(cmd.Fields("pw").Value) And TextBox1.Text = cmd.Fields("un").Value) Then
            mdi.Show()
            Me.Hide()
        Else
            MsgBox("Incorrect Username or Password", vbCritical, "Invalid User")
            TextBox1.Focus()
        End If
        cmd.Close()
    End Sub

    Private Sub login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub

    Private Sub TextBox2_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles TextBox2.PreviewKeyDown
        If e.KeyCode = Keys.Enter Then
            cmd.Open("select * from other ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            cmd.MoveFirst()
            If (TextBox2.Text = StrReverse(cmd.Fields("pw").Value) And TextBox1.Text = cmd.Fields("un").Value) Then
                mdi.Show()
                Me.Hide()
            Else
                MsgBox("Incorrect Username or Password", vbCritical, "Invalid User")
                TextBox1.Focus()
            End If
            cmd.Close()
        End If
    End Sub

End Class