Imports System.Windows.Forms
Imports System.Threading

Public Class DialogOpenUser
    Public Shared myForm As DialogOpenUser
    Dim myselecteddate As Date = Today.Date.Date

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As OpenUserAdapter

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New DialogOpenUser
            'AddHandler DialogUserInput.SaveRecord, AddressOf myForm.AssignRecord
        ElseIf myForm.IsDisposed Then
            myForm = New DialogOpenUser
        End If
        Return myForm
    End Function


    Private Sub FormProjectSigned_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    'Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    '    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    '    Me.Close()
    'End Sub

    'Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    '    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    '    Me.Close()
    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        myAdapter.mydate = myselecteddate
        loaddata()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        myAdapter = OpenUserAdapter.getInstance

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub FormUser_Load(sender As Object, e As EventArgs) Handles Me.Load
        myAdapter.mydate = myselecteddate
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
        myAdapter = OpenUserAdapter.getInstance
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
                    ToolStripStatusLabel1.Text = String.Format("Record count: {0}", myAdapter.BS.Count)
            End Select
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        myselecteddate = DateTimePicker1.Value.Date
    End Sub
End Class
