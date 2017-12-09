Public Class depreciation
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset
    Private Sub depreciation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")



        cmd.Open("select * from expenses ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF

            ComboBox2.Items.Add(cmd.Fields("particulars").Value)
            cmd.MoveNext()
        End While
        cmd.Close()

        cmd.Open("select * from depreciation_cost ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            ComboBox1.Items.Add(cmd.Fields("name_of_asset").Value)

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

        For i As Int16 = 0 To Me.ComboBox2.Items.Count - 2
            For j As Int16 = Me.ComboBox2.Items.Count - 1 To i + 1 Step -1
                If Me.ComboBox2.Items(i).ToString = Me.ComboBox2.Items(j).ToString Then
                    Me.ComboBox2.Items.RemoveAt(j)
                End If
            Next
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView2.Rows.Clear()
        Dim str1, str2 As String
        Dim y As Integer
        str1 = DateTimePicker1.Text
        str2 = DateTimePicker3.Text
        y = 0
        While (Not (str1 = str2))
            cmd.Open("select * from expenses where da_te  ='" & str1 & "' and particulars='" & ComboBox1.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not cmd.EOF)
                DataGridView2.Rows.Add(cmd.Fields("da_te").Value, cmd.Fields("particulars").Value, cmd.Fields("totalamt").Value, "Add Depreciation")
                cmd.MoveNext()
            End While
            cmd.Close()
            y = y + 1
            str1 = DateTimePicker1.Value.Date.AddDays(y)
        End While



        If (str1 = str2) Then
            cmd.Open("select * from expenses where da_te  ='" & str1 & "' and particulars='" & ComboBox1.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not cmd.EOF)
                DataGridView2.Rows.Add(cmd.Fields("da_te").Value, cmd.Fields("particulars").Value, cmd.Fields("totalamt").Value, "Add Depreciation")
                cmd.MoveNext()
            End While
            cmd.Close()
        End If

    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If (DataGridView2.Item("Column1", DataGridView2.CurrentRow.Index).Selected) Then
            depreciation_details.Show()
            depreciation_details.TextBox6.Text = DataGridView2.Item("Column19", DataGridView2.CurrentRow.Index).Value()
            depreciation_details.TextBox1.Text = DataGridView2.Item("Column20", DataGridView2.CurrentRow.Index).Value()

            cmd.Open("select * from depreciation_cost where name_of_asset  ='" & DataGridView2.Item("Column19", DataGridView2.CurrentRow.Index).Value() & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not cmd.EOF)
                depreciation_details.TextBox2.Text = cmd.Fields("depreciation_value").Value
                cmd.MoveNext()
            End While
            cmd.Close()
            depreciation_details.TextBox3.Text = (Val(depreciation_details.TextBox2.Text) / 100) * Val(depreciation_details.TextBox1.Text)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cmd.Open("select * from depreciation_cost ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd.AddNew()
        cmd.Fields("name_of_asset").Value = ComboBox2.Text
        cmd.Fields("depreciation_value").Value = TextBox1.Text
        cmd.Update()
        cmd.Close()
    End Sub
End Class