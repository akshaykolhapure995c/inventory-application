Public Class depreciation_details
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Recordset
    Private Sub depreciation_details_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con.Open("dsn=bkl")
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        cmd.Open("select * from expenses ", con, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        cmd.AddNew()
        cmd.Fields("da_te").Value = DateTimePicker1.Text
        cmd.Fields("particulars").Value = TextBox6.Text
        cmd.Fields("totalamt").Value = TextBox3.Text
        cmd.Update()
        cmd.Close()
        MsgBox("Successfully Added.")
    End Sub
End Class