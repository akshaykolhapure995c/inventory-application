Public Class expenses
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset
    Private Sub expenses_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cmd.Open("select * from expenses ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd.AddNew()
        cmd.Fields("da_te").Value = DateTimePicker2.Text
        cmd.Fields("particulars").Value = TextBox21.Text
        cmd.Fields("totalamt").Value = TextBox1.Text
        cmd.Update()
        cmd.Close()
        MsgBox("Added.")
        TextBox1.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView2.Rows.Clear()
        Dim str1, str2 As String
        Dim y As Integer
        str1 = DateTimePicker1.Text
        str2 = DateTimePicker3.Text
        y = 0
        While (Not (str1 = str2))
            cmd.Open("select * from expenses where da_te  ='" & str1 & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not cmd.EOF)
                DataGridView2.Rows.Add(cmd.Fields("da_te").Value, cmd.Fields("particulars").Value, cmd.Fields("totalamt").Value)
                cmd.MoveNext()
            End While
            cmd.Close()
            y = y + 1
            str1 = DateTimePicker1.Value.Date.AddDays(y)
        End While
        If (str1 = str2) Then
            cmd.Open("select * from expenses where da_te  = '" & str1 & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not cmd.EOF)
                DataGridView2.Rows.Add(cmd.Fields("da_te").Value, cmd.Fields("particulars").Value, cmd.Fields("totalamt").Value)
                cmd.MoveNext()
            End While
            cmd.Close()
        End If

        Dim ab As New Integer
        For y = 0 To DataGridView2.RowCount - 1
            ab = ab + Val(DataGridView2.Item("Column20", DataGridView2.Rows.Item(y).Index).Value)
        Next y
        Label5.Text = ab
    End Sub
End Class