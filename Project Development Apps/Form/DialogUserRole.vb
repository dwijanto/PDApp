Imports System.Windows.Forms

Public Class DialogUserRole
    Public Property _bsRole As BindingSource

    Public Property _bsUserRole As BindingSource
    Private Property bsUserRoleOri As New BindingSource
    Public Property _bsUser As BindingSource

    Dim datarv As DataRowView
    Dim myadapter As UserAdapter
    Public Sub New(myadapter As UserAdapter)

        ' This call is required by the designer.
        InitializeComponent()
        _bsUser = myadapter.BS
        datarv = _bsUser.Current
        Me.myadapter = myadapter
        _bsRole = New BindingSource
        _bsRole.DataSource = myadapter.DS.Tables(1).Copy

        bsUserRoleOri.DataSource = myadapter.DS.Tables(2)

        Me._bsUserRole = New BindingSource
        Me._bsUserRole.DataSource = myadapter.DS.Tables(2).Copy 'Get A new DataSource ' similar as get from database

        Me.Text = Me.Text + " - " + datarv.Item("userid")

        InitData()
        ' Add any initialization after the InitializeComponent() call.

    End Sub
 

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    'Private Sub RoleAssignmentToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RoleAssignmentToolStripMenuItem1.Click
    '    Dim myform As New DialogAddRemoveUserRole(_bsRole)
    '    If myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
    '        'add new userRole show in datagridview
    '        Dim drv As DataRowView = _bsUserRole.AddNew
    '        drv.Row.Item("userid") = datarv.Item("id")
    '        drv.Row.Item("roleid") = myform.drv.Row.Item("id")
    '        drv.Row.Item("rolename") = myform.drv.Row.Item("rolename")
    '        drv.EndEdit()
    '        DataGridView1.Invalidate()
    '        'Don't forget to add to the real database         
    '        'myadapter.DS.Tables(2).Rows.Add(drv.Row)
    '        Dim dr = myadapter.DS.Tables(2).NewRow
    '        dr.Item("userid") = datarv.Item("id")
    '        dr.Item("roleid") = myform.drv.Row.Item("id")
    '        dr.Item("rolename") = myform.drv.Row.Item("rolename")
    '        myadapter.DS.Tables(2).Rows.Add(dr)
    '    End If
    'End Sub


    Private Sub RoleAssignmentToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RoleAssignmentToolStripMenuItem1.Click
        Dim myform As New DialogAddRemoveUserRole(_bsUser, _bsRole, Me)
        myform.Show()
    End Sub
    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If Not IsNothing(_bsUserRole.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    Dim dr = bsUserRoleOri.Find("id", _bsUserRole.Item(drv.Index).row.item("id"))
                    bsUserRoleOri.RemoveAt(dr)
                    'bsUserRoleOri.EndEdit()
                    _bsUserRole.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub DialogUserRole_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub DialogUserRole_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub InitData()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = _bsUserRole
        _bsUserRole.Filter = String.Format("userid = {0}", datarv.Item("id"))
    End Sub

    Public Sub endtx(o As Object, e As DialogAddRemoveUserRoleException)

        Try
            Dim drv As DataRowView = _bsUserRole.AddNew
            drv.Row.Item("userid") = datarv.Item("id")
            drv.Row.Item("roleid") = o.drv.Row.Item("id")
            drv.Row.Item("rolename") = o.drv.Row.Item("rolename")
            drv.EndEdit()
            DataGridView1.Invalidate()
            'Don't forget to add to the real database         
            'myadapter.DS.Tables(2).Rows.Add(drv.Row)
            Dim dr = myadapter.DS.Tables(2).NewRow
            dr.Item("userid") = datarv.Item("id")
            dr.Item("roleid") = o.drv.Row.Item("id")
            dr.Item("rolename") = o.drv.Row.Item("rolename")
            myadapter.DS.Tables(2).Rows.Add(dr)
        Catch ex As Exception
            _bsUserRole.CancelEdit()
            MessageBox.Show(ex.Message)
        End Try

    End Sub
End Class
