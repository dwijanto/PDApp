Imports System.Windows.Forms

Public Class DialogRole
    Public Property bsRole As BindingSource
    Public Shared myform As DialogRole
    'Public Shared myhandler As DialogAddRemoveRole.SaveRecordEventHandler = AddressOf myform.AssignRecord
    'Shared myhandler As New DialogAddRemoveRole.SaveRecordEventHandler(AddressOf myform.AssignRecord)
    'Public Shared myhandler As DialogAddRemoveRole.SaveRecordEventHandler

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.Validate()
        Me.bsRole.EndEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        For Each drv As DataRowView In bsRole.List
            If drv.Row.RowState = DataRowState.Modified Then
                drv.Row.Item("rolename") = drv.Row.Item("rolename", DataRowVersion.Original)
                drv.Row.Item("area") = drv.Row.Item("area", DataRowVersion.Original)
            End If
        Next
        Me.bsRole.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub UpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) 'Handles UpdateToolStripMenuItem.Click, DataGridView1.CellDoubleClick
        Dim myform = New DialogAddRemoveRole(bsRole, TxRecord.UpdateRecord)
        myform.Show()
    End Sub

    Public Shared Function getInstance(myadapter As UserAdapter)
        If myform Is Nothing Then

            myform = New DialogRole(myadapter)
            'myhandler = New DialogAddRemoveRole.SaveRecordEventHandler(AddressOf myform.AssignRecord)
            'RemoveHandler DialogAddRemoveRole.SaveRecord, myhandler
            'AddHandler DialogAddRemoveRole.SaveRecord, myhandler
            ''AddHandler DialogAddRemoveRole.SaveRecord, AddressOf myform.AssignRecord

        ElseIf myform.IsDisposed Then
            'myform = Nothing
            myform = New DialogRole(myadapter)
        End If
        Return myform
    End Function

    Private Sub New(myadapter As UserAdapter)

        ' This call is required by the designer.
        InitializeComponent()

        _bsRole = New BindingSource
        _bsRole.DataSource = myadapter.DS.Tables(1)

        'RemoveHandler DialogAddRemoveRole.SaveRecord, AddressOf Me.AssignRecord
        InitData()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub InitData()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = _bsRole
    End Sub

    Private Sub RoleAssignmentToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RoleAssignmentToolStripMenuItem1.Click
        Dim myform = New DialogAddRemoveRole(bsRole, TxRecord.AddRecord)
        myform.Show()
        'Me.bsRole.AddNew()
    End Sub



    Private Sub AssignRecord(ByRef sender As Object, e As UserEventArgs)
        Dim drv As DataRowView = DirectCast(sender, DataRowView)
        Dim mydrv As DataRowView = Nothing
        Try
            Select Case e.StatusTx
                Case TxRecord.AddRecord
                    mydrv = bsRole.AddNew()
                Case TxRecord.UpdateRecord
                    Dim result = bsRole.Find("id", drv.Item("id"))
                    mydrv = bsRole.Item(result)
            End Select
            mydrv.Item("rolename") = drv.Item("rolename")

            mydrv.EndEdit()
            DataGridView1.Invalidate()
        Catch ex As Exception
            bsRole.CancelEdit()
        End Try
        

    End Sub



    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If Not IsNothing(bsRole.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    bsRole.RemoveAt(drv.Index)
                Next
            End If
        End If       
    End Sub


    Private Sub DialogRole_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        myform.Dispose()
        myform = Nothing
        Me.Dispose()

    End Sub


End Class
