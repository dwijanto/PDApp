Imports System.Windows.Forms
Imports System.Text

Public Class DialogTopVendorV2

    Public Property BS As BindingSource
    Private Property VBS As BindingSource
    Public Shared myformTopVendor As DialogTopVendorV2

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub UpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdateToolStripMenuItem.Click, DataGridView1.CellDoubleClick
        ShowTx(TxRecord.UpdateRecord)
    End Sub

    Public Shared Function getInstance(vendorbs As BindingSource, TopBS As BindingSource)
        If myformTopVendor Is Nothing Then
            myformTopVendor = New DialogTopVendorV2(vendorbs, TopBS)

        End If
        Return myformTopVendor
    End Function

    Public Sub New(vendorbs As BindingSource, TopBS As BindingSource)
        InitializeComponent()
        Me.BS = TopBS
        Me.VBS = vendorbs
        InitData()
    End Sub

    Private Sub InitData()
        CShortname.DisplayMember = "shortname"
        CShortname.ValueMember = "id"
        CShortname.DataSource = VBS

        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = Me.BS
    End Sub


    Private Sub RoleAssignmentToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AddToolStripMenuItem1.Click
        ShowTx(TxRecord.AddRecord)
    End Sub

    Private Sub ShowTx(ByVal StatusTx As TxRecord)
        Dim drv As DataRowView = Nothing
        Select Case StatusTx
            Case TxRecord.AddRecord
                drv = Me.BS.AddNew
            Case TxRecord.UpdateRecord
                drv = Me.BS.Current
        End Select
        Dim myform As New DialogTopVendorInputV2(drv, VBS)
        myform.Show()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim obj As TextBox = DirectCast(sender, TextBox)
        Dim sb As New StringBuilder

        If obj.Text = "" Then
            BS.Filter = ""
        Else
            sb.Append(String.Format("year like '*{0}*' or shortname like '*{0}*'", obj.Text))
            If IsNumeric(obj.Text) Then
                sb.Append(String.Format(" or orderline = {0}", obj.Text))
            End If
            BS.Filter = sb.ToString
        End If
    End Sub

    Private Sub DialogTopVendor_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        myformTopVendor.Dispose()
        myformTopVendor = Nothing
    End Sub

    Public Shared Sub RefreshDataGridView(ByRef sender As Object, e As EventArgs)
        myformTopVendor.DataGridView1.Invalidate()
    End Sub

End Class
