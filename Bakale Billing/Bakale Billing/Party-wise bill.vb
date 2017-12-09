Public Class Party_wise_bill
    Dim con As New ADODB.Connection
    Dim cmd2, cmd As New ADODB.Recordset
    Private Sub Party_wise_bill_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
        cmd.Open("select buyer_name from sales ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            ComboBox1.Items.Add(cmd.Fields("buyer_name").Value)
            cmd.MoveNext()
        End While
        cmd.Close()
        For i As Int16 = 0 To Me.ComboBox1.Items.Count - 2
            For j As Int16 = Me.ComboBox1.Items.Count - 1 To i + 1 Step -1
                If Me.ComboBox1.Items(i).ToString = Me.ComboBox1.Items(j).ToString Then
                    Me.ComboBox1.Items.RemoveAt(j)
                End If
            Next
        Next
    End Sub

    Private Sub ComboBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.LostFocus
        ComboBox2.Items.Clear()
        cmd.Open("select bill_no from sales where buyer_name ='" & ComboBox1.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            ComboBox2.Items.Add(cmd.Fields("bill_no").Value)
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

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView2.Rows.Clear()
        If ComboBox1.Text = "" Or ComboBox2.Text = "" Then
            MsgBox("Incorrect Party Name or Bill Number")
        Else
            cmd.Open("select * from sales where buyer_name ='" & ComboBox1.Text & "' and bill_no='" & ComboBox2.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While Not cmd.EOF
                DataGridView2.Rows.Add(cmd.Fields("order_date").Value, cmd.Fields("particulars").Value, cmd.Fields("piece_qty").Value, cmd.Fields("per_piece_rate").Value, cmd.Fields("item_amount").Value)
                Label5.Text = cmd.Fields("total_taxinclude").Value
                cmd.MoveNext()
            End While

            cmd.Close()
        End If
    End Sub
End Class