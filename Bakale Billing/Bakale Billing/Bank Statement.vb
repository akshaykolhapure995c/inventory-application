Public Class Bank_Statement
    Dim con As New ADODB.Connection
    Dim cmd, cmd1 As New ADODB.Recordset
    Dim str1, str2, str3, str4 As String
    Private Sub Bank_Statement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
        str1 = "By Cheque(Debit)"
        str2 = "By Cash(Debit)"
        str3 = "By Cheque(Credit)"
        str4 = "By Cash(Credit)"
        cmd.Open("select * from bank_statement where payment_type='" & str1 & "' or payment_type='" & str2 & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd1.Open("select * from bank_statement where payment_type='" & str3 & "' or payment_type='" & str4 & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF Or Not cmd1.EOF
            If Not cmd.EOF And cmd1.EOF Then
                DataGridView2.Rows.Add(cmd.Fields("paid_date").Value, cmd.Fields("cust_name").Value, cmd.Fields("cheque_number").Value, cmd.Fields("paid").Value, "-")
                cmd.MoveNext()
            ElseIf Not cmd1.EOF And cmd.EOF Then
                DataGridView2.Rows.Add(cmd1.Fields("paid_date").Value, cmd1.Fields("cust_name").Value, cmd1.Fields("cheque_number").Value, "-", cmd1.Fields("paid").Value)
                cmd1.MoveNext()
            Else
                If cmd.Fields("paid_date").Value = cmd1.Fields("paid_date").Value Then
                    DataGridView2.Rows.Add(cmd.Fields("paid_date").Value, cmd.Fields("cust_name").Value, cmd.Fields("cheque_number").Value, cmd.Fields("paid").Value, "-")
                Else
                    DataGridView2.Rows.Add(cmd.Fields("paid_date").Value, cmd.Fields("cust_name").Value, cmd.Fields("cheque_number").Value, cmd.Fields("paid").Value, "-")
                    DataGridView2.Rows.Add(cmd1.Fields("paid_date").Value, cmd1.Fields("cust_name").Value, cmd1.Fields("cheque_number").Value, "-", cmd1.Fields("paid").Value)
                End If
                cmd.MoveNext()
                cmd1.MoveNext()
            End If

        End While

        cmd.Close()
        cmd1.Close()

        Dim ab As New Integer
        For y = 0 To DataGridView2.RowCount - 1
            ab = ab + Val(DataGridView2.Item("Column20", DataGridView2.Rows.Item(y).Index).Value)
        Next y
        Label4.Text = ab
    End Sub
End Class