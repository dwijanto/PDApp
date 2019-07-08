Imports System.Threading
Imports System.Text

'Rule:
' User can delete at the same period (postingdate month = current month)
' Admin can delete any

Public Class FormProjectSigned
    Public Shared myForm As FormProjectSigned
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As ProjectSignedAdapter
    Private myuserinfo As UserInfo

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormProjectSigned
            AddHandler DialogProjectSignedInput.SaveRecord, AddressOf myForm.AssignRecord
        ElseIf myForm.IsDisposed Then
            myForm = New FormProjectSigned
        End If
        Return myForm
    End Function


    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim myform = New DialogProjectSignedInput(myAdapter, TxRecord.AddRecord)
        myform.Show()
    End Sub

    Private Sub FormProjectSigned_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub


    Private Sub FormUser_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub FormProjectSigned_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not ValidateRP4() Then
            e.Cancel = True
        End If
    End Sub

    Private Function ValidateRP4() As Boolean
        Dim myresult As Boolean = True
        Dim myAdapter = New ProjectSignedAdapter
        Dim result As String = String.Empty
        If myAdapter.ValidateRP4(result) Then
            If result.Length > 0 Then
                MessageBox.Show(String.Format("Missing one of RP4 Ontime or RP4 Price. Please check your transaction. {0}", vbCrLf & vbCrLf & result))
                myresult = False
            End If
        End If
        'Throw New NotImplementedException
        Return myresult
    End Function

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
        myAdapter = ProjectSignedAdapter.getInstance
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
            Dim myform = New DialogProjectSignedInput(myAdapter, TxRecord.UpdateRecord)
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
            If Not myuserinfo.isAdmin Then
                Dim drv As DataRowView = myAdapter.BS.Current
                If String.Format("{0:MMM-yyyy}", drv.Item("postingdate")) <> String.Format("{0:MMM-yyyy}", Date.Today) Then
                    MessageBox.Show("You cannot delete this record. Please contact Admin.")
                    Exit Sub
                End If
            End If
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
        'mydrv.Item("userid") = drv.Item("userid")
        'mydrv.Item("username") = drv.Item("username")
        'mydrv.Item("email") = drv.Item("email")
        'mydrv.Item("isactive") = drv.Item("isactive")
        'mydrv.Item("isadmin") = drv.Item("isadmin")
        'mydrv.Item("title") = drv.Item("title")
        mydrv.EndEdit()
        DataGridView1.Invalidate()
    End Sub

End Class