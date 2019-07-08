Imports System.Windows.Forms
Imports System.Text

Public Class DialogTopVendor

    Public Property BS As BindingSource
    Public Shared myformTopVendor As DialogTopVendor
    Public myadapter As VendorAdapter
    Public Shared Event SaveRecord(ByRef sender As Object, ByVal e As EventArgs)
    Private Property VBS As BindingSource

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.Validate()
        Me.BS.EndEdit()
        DataGridView1.Refresh()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub UpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdateToolStripMenuItem.Click, DataGridView1.CellDoubleClick
        'Dim myform = New DialogTopVendorInput(Me.BS, myadapter.BS, TxRecord.UpdateRecord)
        'myform.Show()
    End Sub

    Public Shared Function getInstance(myadapter As VendorAdapter)
        If myformTopVendor Is Nothing Then
            myformTopVendor = New DialogTopVendor(myadapter)
            AddHandler DialogTopVendorInput.SaveRecord, AddressOf myformTopVendor.AssignRecord
        ElseIf myformTopVendor.IsDisposed Then
            'myform = Nothing
            myformTopVendor = New DialogTopVendor(myadapter)
        End If
        Return myformTopVendor
    End Function

    Public Sub New(myadapter As VendorAdapter)

        ' This call is required by the designer.
        InitializeComponent()

        Dim mytbl As New DataTable
        mytbl = TryCast(myadapter.BS.DataSource, DataTable).Copy
        VBS = New BindingSource
        VBS.DataSource = mytbl

        'Me.BS = New BindingSource
        myadapter.topBS.Filter = ""
        Me.BS = myadapter.topBS 'mytbl

        TextBox1.Text = ""
        'Me.BS = myadapter.topBS
        Me.myadapter = myadapter

        InitData()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    'Private Sub New(myadapter As VendorAdapter)

    '    ' This call is required by the designer.
    '    InitializeComponent()

    '    Dim mytbl As New DataTable
    '    mytbl = TryCast(myadapter.topBS.DataSource, DataTable).Copy

    '    Me.BS = New BindingSource
    '    Me.BS.DataSource = mytbl

    '    TextBox1.Text = ""
    '    'Me.BS = myadapter.topBS
    '    Me.myadapter = myadapter

    '    InitData()
    '    ' Add any initialization after the InitializeComponent() call.

    'End Sub

    Private Sub InitData()


        CShortname.DisplayMember = "shortname"
        CShortname.ValueMember = "id"
        CShortname.DataSource = VBS 'myadapter.BS

        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = Me.BS
    End Sub

    Private Sub RoleAssignmentToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RoleAssignmentToolStripMenuItem1.Click
        'DataGridView1.Refresh()
        'Dim myform = New DialogTopVendorInput(Me.BS, myadapter.BS, TxRecord.AddRecord)
        'myform.Show()
        Me.BS.AddNew()
    End Sub

    Private Sub AssignRecord(ByRef sender As Object, e As UserEventArgs)
        Dim drv As DataRowView = DirectCast(sender, DataRowView)
        Dim mydrv As DataRowView = Nothing
        Select Case e.StatusTx
            Case TxRecord.AddRecord
                mydrv = Me.BS.AddNew()
                'mydrv = myadapter.topBS.AddNew
            Case TxRecord.UpdateRecord
                Dim result = Me.BS.Find("id", drv.Item("id"))
                mydrv = Me.BS.Item(result)
        End Select
        mydrv.Item("year") = drv.Item("year")
        mydrv.Item("vendorid") = drv.Item("vendorid")
        mydrv.Item("orderline") = drv.Item("orderline")
        mydrv.Item("shortname") = drv.Item("shortname")
        mydrv.EndEdit()
        Application.DoEvents()
        DataGridView1.DataSource = Me.BS
        DataGridView1.Invalidate()
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
        'Me.Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.BS.AddNew()
    End Sub


End Class
