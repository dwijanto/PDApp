Imports System.Windows.Forms
Imports System.Text

Public Class DialogProjectHelper
    Public BS As BindingSource
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub New(ByVal bs As BindingSource)
        InitializeComponent()

        Dim tbl As New DataTable
        tbl = CType(bs.DataSource, DataTable).Copy
        Me.BS = New BindingSource
        Me.BS.DataSource = tbl
        Me.BS.Filter = "statusname = 'Active'"

        DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.DataSource = Me.BS

    End Sub





    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim obj As TextBox = DirectCast(sender, TextBox)
        Dim sb As New StringBuilder

        If obj.Text = "" Then
            BS.Filter = "statusname = 'Active'"
        Else
            sb.Append(String.Format("(projectid like '*{0}*' or projectname like '*{0}*' or shortname like '*{0}*')  and statusname = 'Active'", obj.Text))
            BS.Filter = sb.ToString
        End If

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex > -1 Then
            OK_Button.PerformClick()
        End If
    End Sub
End Class
