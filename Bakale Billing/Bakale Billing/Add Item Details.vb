Public Class Add_Item_Details
    Dim con As New ADODB.Connection
    Dim cmd, cmd2, cmd3 As New ADODB.Recordset
    Dim item_name As String
    Dim ramount, total_qty, sizes As New Integer
    Dim check As Boolean
    Dim names As New HashSet(Of String)
    Dim b As Boolean

    Private Sub Add_Item_Details_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
        cmd.Open("select * from item_name ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            names.Add(cmd.Fields("itemnm").Value.ToString.ToUpper)
            cmd.MoveNext()
        End While
        cmd.Close()
        For a As Integer = 0 To names.Count - 1
            ComboBox1.Items.Add(names(a))
        Next
        names.Clear()
    End Sub

    Private Sub TextBox1_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs)
        If e.KeyCode = Keys.Enter Then
            total_qty = Val(TextBox3.Text) + Val(TextBox1.Text)
            TextBox2.Text = Val(TextBox4.Text) * (Val(TextBox3.Text) + Val(TextBox1.Text))
            DataGridView1.Rows.Add(ComboBox1.Text, ComboBox2.Text, TextBox4.Text, total_qty, TextBox2.Text)
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            TextBox4.Text = ""
            TextBox1.Text = ""
            TextBox3.Text = ""
            TextBox2.Text = ""
            ComboBox1.Focus()
        End If
    End Sub


    Public Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        additem()
    End Sub

    Private Sub additem()
        For v = 0 To DataGridView1.RowCount - 1
            item_name = DataGridView1.Item("Column1", DataGridView1.Rows.Item(v).Index).Value
            sizes = DataGridView1.Item("Column2", DataGridView1.Rows.Item(v).Index).Value
            cmd2.Open("select * from item_name where itemnm= '" & item_name & "' and size= '" & sizes & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            While Not cmd2.EOF
                cmd2.Fields("itemnm").Value = DataGridView1.Item("Column1", DataGridView1.Rows.Item(v).Index).Value
                cmd2.Fields("size").Value = DataGridView1.Item("Column2", DataGridView1.Rows.Item(v).Index).Value
                cmd2.Fields("per_piece_rate").Value = DataGridView1.Item("Column3", DataGridView1.Rows.Item(v).Index).Value
                cmd2.Fields("quantity").Value = DataGridView1.Item("Column4", DataGridView1.Rows.Item(v).Index).Value
                cmd2.Fields("total_amount").Value = DataGridView1.Item("Column5", DataGridView1.Rows.Item(v).Index).Value
                cmd2.Fields("date_").Value = DateTimePicker1.Text
                cmd2.MoveNext()
            End While
            cmd2.Close()
        Next

        'cmd2.Update()

        sizes = 0
        item_name = ""
        MsgBox("Added Successfully.")
        DataGridView1.Rows.Clear()
        '   ComboBox3.Text = ""

    End Sub

    
    Private Sub ComboBox3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        cmd.Open("select * from item_name where itemnm='" & ComboBox1.Text & "' and size='" & ComboBox2.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd.EOF
            TextBox3.Text = cmd.Fields("quantity").Value
            TextBox4.Text = cmd.Fields("per_piece_rate").Value
            TextBox2.Text = cmd.Fields("total_amount").Value
            cmd.MoveNext()
        End While
        TextBox2.Text = Val(TextBox4.Text) * (Val(TextBox3.Text) + Val(TextBox1.Text))
        cmd.Close()
        TextBox4.Focus()
    End Sub

    Private Sub ComboBox1_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.LostFocus

        ComboBox2.Text = ""
        ComboBox2.Items.Clear()
      

        cmd2.Open("select * from item_name where itemnm ='" & ComboBox1.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd2.EOF
            ComboBox2.Items.Add(cmd2.Fields("size").Value)
            cmd2.MoveNext()
        End While

        For i As Int16 = 0 To Me.ComboBox2.Items.Count - 2
            For j As Int16 = Me.ComboBox2.Items.Count - 1 To i + 1 Step -1
                If Me.ComboBox2.Items(i).ToString = Me.ComboBox2.Items(j).ToString Then
                    Me.ComboBox2.Items.RemoveAt(j)
                End If
            Next
        Next

        cmd2.Close()
    End Sub

    Private Sub ComboBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.LostFocus
        

        cmd2.Open("select * from item_name where itemnm= '" & ComboBox1.Text & "' and size= '" & ComboBox2.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd2.EOF
            TextBox3.Text = cmd2.Fields("quantity").Value
            TextBox4.Text = cmd2.Fields("per_piece_rate").Value
            TextBox2.Text = Val(TextBox3.Text) * Val(TextBox4.Text)
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            cmd2.MoveNext()
        End While
        cmd2.Close()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (TextBox1.Text = "") Then
            MsgBox("Fill all fields !!", vbCritical, "Alert")

        Else
            If (Not IsNumeric(TextBox1.Text) Or Not IsNumeric(TextBox4.Text)) Then
                MsgBox("Please Enter Numeric values", vbCritical, "Invalid Input!")
            Else

                total_qty = Val(TextBox3.Text) + Val(TextBox1.Text)
                TextBox2.Text = Val(TextBox4.Text) * (Val(TextBox3.Text) + Val(TextBox1.Text))
                DataGridView1.Rows.Add(ComboBox1.Text, ComboBox2.Text, TextBox4.Text, total_qty, TextBox2.Text)
                
                ComboBox1.Text = ""
                ComboBox2.Text = ""
                TextBox1.Text = ""
                TextBox4.Text = ""
                TextBox3.Text = ""
                TextBox2.Text = ""
                ComboBox1.Focus()
            End If
        End If
    End Sub
End Class