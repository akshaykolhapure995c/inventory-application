Public Class showstock

    Dim con As New ADODB.Connection
    Dim cmd, cmd2, cmd3 As New ADODB.Recordset
    Dim names As New HashSet(Of String)

    Private Sub showstock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        con.Open("dsn=bkl")


        cmd.Open("select * from item_name ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        Dim str, convertstr As String
        While Not cmd.EOF
            str = cmd.Fields("itemnm").Value
            convertstr = str.ToUpper()
            names.Add(convertstr)
            cmd.MoveNext()
        End While
        For a As Integer = 0 To names.Count - 1
            ComboBox1.Items.Add(names(a))
            ComboBox2.Items.Add(names(a))
        Next
        cmd.Close()



        'cmd2.Open("select * from item_name ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        ' While Not cmd2.EOF
        'ComboBox1.Items.Add(cmd2.Fields("itemnm").Value)
        'ComboBox2.Items.Add(cmd2.Fields("itemnm").Value)
        'cmd2.MoveNext()
        'End While

        'cmd2.Close()


        Dim i, j As Integer
        Dim arr(16) As Integer
        For i = 0 To ComboBox1.Items.Count - 1
            For j = 0 To DataGridView2.Columns.Count - 2
                ' MsgBox(DataGridView2.Columns(j + 1).HeaderCell.Value)
                cmd2.Open("select quantity from item_name where itemnm='" & ComboBox1.Items(i).ToString & "' and size='" & DataGridView2.Columns(j + 1).HeaderCell.Value & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If cmd2.EOF Then
                    arr(j) = 0
                Else
                    arr(j) = cmd2.Fields("quantity").Value
                End If

                cmd2.Close()
            Next
            DataGridView2.Rows.Add(ComboBox1.Items(i).ToString, arr(0), arr(1), arr(2), arr(3), arr(4), arr(5), arr(6), arr(7), arr(8), arr(9), arr(10), arr(11), arr(12), arr(13), arr(14), arr(15))
        Next
        DataGridView2.EnableHeadersVisualStyles = False
        DataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.Beige
        ' DataGridView2.ColumnHeadersDefaultCellStyle.DefaultCellStyle.BackColor = Color.Blue
 

    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim j As Integer
        Dim arr(16) As Integer
        DataGridView2.Rows.Clear()
        '   For i = 0 To ComboBox1.Items.Count - 1
        For j = 0 To DataGridView2.Columns.Count - 2
            ' MsgBox(DataGridView2.Columns(j + 1).HeaderCell.Value)
            cmd2.Open("select quantity from item_name where itemnm='" & ComboBox2.Text & "' and size='" & DataGridView2.Columns(j + 1).HeaderCell.Value & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If cmd2.EOF Then
                arr(j) = 0
            Else
                arr(j) = cmd2.Fields("quantity").Value
            End If

            cmd2.Close()
        Next
        DataGridView2.Rows.Add(ComboBox2.Text, arr(0), arr(1), arr(2), arr(3), arr(4), arr(5), arr(6), arr(7), arr(8), arr(9), arr(10), arr(11), arr(12), arr(13), arr(14), arr(15))
        ' Next
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView2.Rows.Clear()
        Dim i, j As Integer
        Dim arr(16) As Integer
        For i = 0 To ComboBox1.Items.Count - 1
            For j = 0 To DataGridView2.Columns.Count - 2
                ' MsgBox(DataGridView2.Columns(j + 1).HeaderCell.Value)
                cmd2.Open("select quantity from item_name where itemnm='" & ComboBox1.Items(i).ToString & "' and size='" & DataGridView2.Columns(j + 1).HeaderCell.Value & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If cmd2.EOF Then
                    arr(j) = 0
                Else
                    arr(j) = cmd2.Fields("quantity").Value
                End If

                cmd2.Close()
            Next
            DataGridView2.Rows.Add(ComboBox1.Items(i).ToString, arr(0), arr(1), arr(2), arr(3), arr(4), arr(5), arr(6), arr(7), arr(8), arr(9), arr(10), arr(11), arr(12), arr(13), arr(14), arr(15))
        Next
    End Sub
End Class