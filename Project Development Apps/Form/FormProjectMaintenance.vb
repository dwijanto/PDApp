Imports System.Text
Imports System.Threading

Public Class FormProjectMaintenance
    Public Shared myForm As FormProjectMaintenance
    Private myUserInfo As UserInfo

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormProjectMaintenance
            ' AddHandler DialogProjectInput.SaveRecord, AddressOf myForm.AssignRecord
            'AddHandler DialogProjectInput2.SaveRecord, AddressOf myForm.AssignRecord
            AddHandler DialogProjectInput3.SaveRecord, AddressOf myForm.AssignRecord
        ElseIf myForm.IsDisposed Then
            myForm = New FormProjectMaintenance
        End If
        Return myForm
    End Function


    Private Sub FormProjectMaintenance_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As ProjectAdapter


    Private Sub FormUser_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub FormUser_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        myAdapter = ProjectAdapter.getInstance
        Try
            If myAdapter.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4                   
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myAdapter.BS
                    'VendorCB.DataSource = myAdapter.VendorBS
                    'VendorCB.DisplayMember = "shortname"
                    'VendorCB.ValueMember = "id"

                    'RoleCB.DataSource = myAdapter.RoleBS
                    'RoleCB.DisplayMember = "rolename"
                    'RoleCB.ValueMember = "id"

                    'PICCB.DataSource = myAdapter.PICBS
                    'PICCB.DisplayMember = "username"
                    'PICCB.ValueMember = "id"

                    'Column5.DataSource = myAdapter.ProjectStatusBS
                    'Column5.DisplayMember = "statusname"
                    'Column5.ValueMember = "id"
                    'Column5.DataPropertyName = "project_status"
            End Select
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ToolStripStatusLabel1.Text = String.Format("Row selected : {0}", DataGridView1.SelectedRows.Count)
    End Sub
    Private Sub DataGridView1_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseUp
        ToolStripStatusLabel1.Text = String.Format("Row selected : {0}", DataGridView1.SelectedRows.Count)
    End Sub


    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim action As TxRecord = TxRecord.UpdateRecord
        If e.RowIndex <> -1 Then
            If Not myUserInfo.isAdmin Then
                Dim drv As DataRowView = myAdapter.BS.Current
                If String.Format("{0:MMM-yyyy}", drv.Item("creationdate")) <> String.Format("{0:MMM-yyyy}", Date.Today) Then
                    'MessageBox.Show("You cannot modify this record. Please contact Admin.")

                    'Exit Sub
                    'action = TxRecord.ViewRecord

                End If
            End If
            'Dim myform = New DialogProjectInput(myAdapter, action)
            'Dim myform = New DialogProjectInput2(myAdapter, action)
            Dim myform = New DialogProjectInput3(myAdapter, action)
            myform.ShowDialog()
        End If

    End Sub
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        'Dim myform As New DialogProjectInput(myAdapter, TxRecord.AddRecord)
        'Dim myform As New DialogProjectInput2(myAdapter, TxRecord.AddRecord)
        Dim myform As New DialogProjectInput3(myAdapter, TxRecord.AddRecord)
        myform.ShowDialog()
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        loaddata()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        myAdapter.Save()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(myAdapter.BS.Current) Then
            If Not myUserInfo.isAdmin Then
                Dim drv As DataRowView = myAdapter.BS.Current
                If String.Format("{0:MMM-yyyy}", drv.Item("creationdate")) <> String.Format("{0:MMM-yyyy}", Date.Today) Then

                    MessageBox.Show("You cannot delete this record. Please contact Admin.")
                    Exit Sub
                Else

                    If myUserInfo.Userid.ToLower <> drv.Item("modifiedby").ToString.ToLower Then
                        MessageBox.Show("This record does not belong to you, cannot be deleted.")
                        Exit Sub
                    End If
                End If
            End If
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myAdapter.BS.Position = drv.Index
                    Dim currdrv As DataRowView = myAdapter.BS.Current
                    Dim sameuser As Boolean = currdrv.item("modifiedby").ToString.ToLower = myUserInfo.Userid.ToLower
                    Dim usercandelete As Boolean = String.Format("{0:MMM-yyyy}", currdrv.Item("creationdate")) = String.Format("{0:MMM-yyyy}", Date.Today)
                    If myUserInfo.isAdmin Or (usercandelete And sameuser) Then
                        myAdapter.BS.RemoveAt(drv.Index)
                    Else
                        currdrv.Row.RowError = "This record does not belong to you.cannot be deleted"
                    End If

                Next
            End If
        End If
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim obj As ToolStripTextBox = DirectCast(sender, ToolStripTextBox)
        Dim sb As New StringBuilder

        If obj.Text = "" Then
            myAdapter.BS.Filter = ""
        Else
            sb.Append(String.Format("projectid like '*{0}*' or projectname like '*{0}*' or shortname like '*{0}*' or rolename like '*{0}*' or statusname like '*{0}*' or username like '*{0}*'", obj.Text))
            myAdapter.BS.Filter = sb.ToString
        End If
    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        myUserInfo = UserInfo.getInstance
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub AssignRecord(ByRef sender As Object, ByVal e As UserEventArgs)
        Dim drv As DataRowView = DirectCast(sender, DataRowView)
        Dim mydrv As DataRowView = Nothing
        Select Case e.StatusTx
            Case TxRecord.AddRecord
                mydrv = myAdapter.BS.AddNew()

            Case TxRecord.UpdateRecord
                Dim position = myAdapter.BS.Find("id", drv.Item("id"))
                mydrv = myAdapter.BS.Item(position)
        End Select
        For i = 1 To mydrv.Row.ItemArray.Length - 1
            mydrv.Item(i) = drv.Item(i)
        Next       
        mydrv.EndEdit()
        DataGridView1.Invalidate()
    End Sub





End Class