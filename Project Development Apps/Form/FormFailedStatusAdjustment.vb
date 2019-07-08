Imports System.Threading
Imports System.Text

Public Class FormFailedStatusAdjustment

    Public Shared myForm As FormFailedStatusAdjustment
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As ProjectAdjustmentAdapter
    Private myuserinfo As UserInfo


    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormFailedStatusAdjustment
            AddHandler DialogFailedStatusInput.SaveRecord, AddressOf myForm.AssignRecord
        ElseIf myForm.IsDisposed Then
            myForm = New FormFailedStatusAdjustment
        End If
        Return myForm
    End Function


    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim myform = New DialogFailedStatusInput(myAdapter, TxRecord.AddRecord)
        myform.Show()
    End Sub

    Private Sub FormProjectSigned_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub


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
        myAdapter = ProjectAdjustmentAdapter.getInstance
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
                    'Init Combobox
                    'CProjectId.DataSource = myAdapter.ProjectBS
                    'CProjectId.DisplayMember = "projectid"
                    'CProjectId.ValueMember = "id"
                    'CProjectId.DataPropertyName = "projectid"

                    'CProjectName.DataSource = myAdapter.ProjectBS
                    'CProjectName.DisplayMember = "projectname"
                    'CProjectName.ValueMember = "id"
                    'CProjectName.DataPropertyName = "projectid"

                    'CShortName.DataSource = myAdapter.ProjectBS
                    'CShortName.DisplayMember = "shortname"
                    'CShortName.ValueMember = "id"
                    'CShortName.DataPropertyName = "projectid"

                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myAdapter.BS
            End Select
        End If
    End Sub



    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex <> -1 Then
            If Not myuserinfo.isAdmin Then
                Dim drv As DataRowView = myAdapter.BS.Current
                If String.Format("{0:MMM-yyyy}", drv.Item("postingdate")) <> String.Format("{0:MMM-yyyy}", Date.Today) Then
                    MessageBox.Show("You cannot modify this record. Please contact Admin.")
                    Exit Sub
                End If
            End If
            Dim myform = New DialogFailedStatusInput(myAdapter, TxRecord.UpdateRecord)
            myform.Show()
        End If
    End Sub



    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        loaddata()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        myAdapter.Save()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(myAdapter.BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myAdapter.BS.RemoveAt(drv.Index)
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
            sb.Append(String.Format("pidname like '*{0}*' or projectname like '*{0}*' or shortname like '*{0}*' or projectstage like '*{0}*' or phase_status_name like '*{0}*'", obj.Text))
            myAdapter.BS.Filter = sb.ToString

        End If
    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        myuserinfo = UserInfo.getInstance


    End Sub

    Private Sub AssignRecord(ByRef sender As Object, ByVal e As UserEventArgs)
        Dim drv As DataRowView = DirectCast(sender, DataRowView)
        Dim mydrv As DataRowView = Nothing
        Select Case e.StatusTx
            Case TxRecord.AddRecord
                mydrv = myAdapter.BS.AddNew()
            Case TxRecord.UpdateRecord
                Dim result = myAdapter.BS.Find("id", drv.Item("id"))
                mydrv = myAdapter.BS.Item(result)
        End Select
        For i = 1 To mydrv.Row.ItemArray.Length - 1
            mydrv.Item(i) = drv.Item(i)
        Next

        mydrv.EndEdit()
        DataGridView1.Invalidate()
    End Sub

End Class