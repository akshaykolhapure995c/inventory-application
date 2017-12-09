Public Class profit_loss
    Dim con As New ADODB.Connection
    Dim cmd, cmd1 As New ADODB.Recordset
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
        For g = 0 To DataGridView2.RowCount - 1
            ab = ab + Val(DataGridView2.Item("Column20", DataGridView2.Rows.Item(g).Index).Value)
        Next g
        Label1.Text = ab
    End Sub

    Private Sub profit_loss_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub
    Sub n2()
        Dim str1, str2 As String
        Dim y As New Integer
        str1 = DateTimePicker1.Text
        str2 = DateTimePicker3.Text
        If (str1 = str2) Then
            cmd1.Open("select * from sales_report where date_ ='" & str1 & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not cmd1.EOF)
                DataGridView1.Rows.Add(cmd1.Fields("date_").Value, cmd1.Fields("gross").Value)
                cmd1.MoveNext()
            End While
            cmd1.Close()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.Rows.Clear()
        Dim str1, str2 As String
        Dim y As Integer
        str1 = DateTimePicker1.Text
        str2 = DateTimePicker3.Text
        y = 0
        While (Not (str1 = str2))
            cmd1.Open("select * from sales_report where date_ ='" & str1 & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not cmd1.EOF)
                DataGridView1.Rows.Add(cmd1.Fields("date_").Value, cmd1.Fields("gross").Value)
                cmd1.MoveNext()
            End While
            cmd1.Close()
            y = y + 1
            str1 = DateTimePicker1.Value.Date.AddDays(y)

        End While

        If (str1 = str2) Then
            cmd1.Open("select * from sales_report where date_  = '" & str1 & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While (Not cmd1.EOF)
                DataGridView1.Rows.Add(cmd1.Fields("date_").Value, cmd1.Fields("gross").Value)
                cmd1.MoveNext()
            End While
            cmd1.Close()
        End If

        Dim ab As New Integer
        For g = 0 To DataGridView1.RowCount - 1
            ab = ab + Val(DataGridView1.Item("Column2", DataGridView1.Rows.Item(g).Index).Value)
        Next g
        Label2.Text = ab

        Label6.Text = Val(Label2.Text) - Val(Label1.Text)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim x, y As New Integer
        Dim dt As New DataTable
        With dt
            .Columns.Add("particular")
            .Columns.Add("expensesamt")
            .Columns.Add("sales_amount")
            .Columns.Add("date_a")
            .Columns.Add("date_b")
            .Columns.Add("total_expense")
            .Columns.Add("total_sales")
            .Columns.Add("gross_profit")
        End With
        x = DataGridView2.Rows.Count
        y = DataGridView1.Rows.Count

        If (x > y) Then
            For i = 0 To x - 1
                If i > y - 1 Then
                    dt.Rows.Add(DataGridView2.Rows(i).Cells(1).Value, DataGridView2.Rows(i).Cells(2).Value, "", DateTimePicker1.Value.ToShortDateString, DateTimePicker3.Value.ToShortDateString, Label1.Text, Label2.Text, Label6.Text)
                Else
                    dt.Rows.Add(DataGridView2.Rows(i).Cells(1).Value, DataGridView2.Rows(i).Cells(2).Value, DataGridView1.Rows(i).Cells(1).Value, DateTimePicker1.Value.ToShortDateString, DateTimePicker3.Value.ToShortDateString, Label1.Text, Label2.Text, Label6.Text)
                End If
            Next
        Else
            For i = 0 To y - 1
                If i > x - 1 Then
                    dt.Rows.Add("", "", DataGridView1.Rows(i).Cells(1).Value, DateTimePicker1.Value.ToShortDateString, DateTimePicker3.Value.ToShortDateString, Label1.Text, Label2.Text, Label6.Text)
                Else
                    dt.Rows.Add(DataGridView2.Rows(i).Cells(1).Value, DataGridView2.Rows(i).Cells(2).Value, DataGridView1.Rows(i).Cells(1).Value, DateTimePicker1.Value.ToShortDateString, DateTimePicker3.Value.ToShortDateString, Label1.Text, Label2.Text, Label6.Text)
                End If
            Next
        End If




        '  For Each dr As DataGridViewRow In Me.DataGridView2.Rows
        'dt.Rows.Add(dr.Cells(1).Value, dr.Cells(2).Value, "", DateTimePicker1.Text, DateTimePicker3.Text, Label1.Text, Label2.Text, Label6.Text)
        '  Next


        ' For Each dr2 As DataGridViewRow In Me.DataGridView1.Rows
        '   dt.Rows.Add("", "", dr2.Cells(1).Value, DateTimePicker1.Text, DateTimePicker3.Text, Label1.Text, Label2.Text, Label6.Text)
        '  Next

        Dim rptdoc As CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptdoc = New CrystalReport5
        rptdoc.SetDataSource(dt)
        rptdoc.PrintToPrinter(1, False, 0, 0)
    End Sub
End Class