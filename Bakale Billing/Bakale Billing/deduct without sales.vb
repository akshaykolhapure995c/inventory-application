Public Class Deduct_without_sales
    Dim con As New ADODB.Connection
    Dim cmd, cmd2 As New ADODB.Recordset
    Dim v As New Integer
    'Dim srno As New Integer
    Dim names As New HashSet(Of String)
    Dim rm As Integer
    ' Dim printcopy As Boolean
    'Dim line As Integer
    'Dim count As Integer = 0
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If (Val(TextBox2.Text) > Val(TextBox1.Text) Or Not IsNumeric(TextBox2.Text)) Then
            MsgBox("Invalid Operation", vbCritical, "Error")
        Else
            cmd2.Open("select * from item_name where itemnm = '" & ComboBox1.Text & "' and size='" & ComboBox2.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            While Not cmd2.EOF

                cmd2.Fields("quantity").Value = Val(TextBox1.Text) - Val(TextBox2.Text)
                cmd2.MoveNext()
            End While
            MsgBox("Successfully Deducted")
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            TextBox1.Text = ""
            TextBox2.Text = ""
            cmd2.Close()
        End If
    End Sub

    Private Sub Deduct_without_sales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
        cmd2.Open("select * from item_name ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        While Not cmd2.EOF
            names.Add(cmd2.Fields("itemnm").Value.ToString.ToUpper)
            cmd2.MoveNext()
        End While

        cmd2.Close()
        For a As Integer = 0 To names.Count - 1
            ComboBox1.Items.Add(names(a))
        Next
        names.Clear()
    End Sub

    Private Sub ComboBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.LostFocus
        ComboBox2.Items.Clear()
        ComboBox2.Text = ""
        cmd2.Open("select * from item_name where itemnm = '" & ComboBox1.Text & "' ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd2.EOF
            names.Add(cmd2.Fields("size").Value)
            cmd2.MoveNext()
        End While
        cmd2.Close()
        For a As Integer = 0 To names.Count - 1
            ComboBox2.Items.Add(names(a))
        Next

    End Sub

    Private Sub ComboBox2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.LostFocus
        TextBox1.Text = ""
        cmd2.Open("select * from item_name where itemnm = '" & ComboBox1.Text & "' and size='" & ComboBox2.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        While Not cmd2.EOF
            TextBox1.Text = cmd2("quantity").Value
            TextBox1.Enabled = False
            cmd2.MoveNext()
        End While
        cmd2.Close()
    End Sub
End Class