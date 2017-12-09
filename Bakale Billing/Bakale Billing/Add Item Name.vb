Public Class Add_Item_Name
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset
    Private Sub Add_Item_Name_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub





    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (ComboBox1.Text = "" Or TextBox2.Text = "" Or ComboBox2.Text = "" Or TextBox1.Text = "" Or TextBox3.Text = "") Then
            MsgBox("Please fill all fields.", vbCritical, "Message")
            ComboBox1.Focus()

        Else
            cmd.Open("Select * from item_name where itemnm='" & ComboBox1.Text & "' and size='" & ComboBox2.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            If Not cmd.EOF Then
                MsgBox("Item Name already exists please enter another item name or size.", vbCritical, "Message")
                TextBox1.Focus()
                cmd.Close()

            Else
                
                cmd.Close()
                cmd.Open("select * from item_name where itemnm='" & ComboBox1.Text & "' and size='" & ComboBox1.Text & "'", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If Not cmd.EOF Then
                    MsgBox("Item name already exists please enter another Item name.", vbCritical, "Message")
                    ComboBox1.Focus()
                    cmd.Close()
                Else
                    cmd.AddNew()
                    cmd.Fields("itemnm").Value = ComboBox1.Text
                    cmd.Fields("size").Value = ComboBox2.Text
                    cmd.Fields("quantity").Value = TextBox2.Text
                    cmd.Fields("per_piece_rate").Value = TextBox1.Text
                    cmd.Fields("total_amount").Value = ""
                    cmd.Fields("date_").Value = ""

                    cmd.Fields("hsn").Value = TextBox3.Text

                    MsgBox("Added Successfully.")
                    cmd.Update()
                    cmd.Close()
                End If
                ComboBox1.Text = ""
                TextBox2.Text = ""
                ComboBox2.Text = ""
                TextBox1.Text = ""
                TextBox3.Text = ""
            End If
        End If
    End Sub
End Class