Imports System.Windows.Forms
Imports System.Threading

Public Class DialogExtendPeriod
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As PeriodAdapter

    Public Shared myFormDialogExtendPeriod As DialogExtendPeriod

    Public Shared Function getInstance()
        If myFormDialogExtendPeriod Is Nothing Then
            myFormDialogExtendPeriod = New DialogExtendPeriod
        End If
        Return myFormDialogExtendPeriod
    End Function
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Sub DoWork()
        myAdapter = PeriodAdapter.getInstance
        Try
            If myAdapter.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Private Sub DialogExtendPeriod_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        myFormDialogExtendPeriod = Nothing
        Me.Dispose()      
    End Sub
    Private Sub DialogExtendPeriod_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadData()
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

                    TextBox1.Text = String.Format("{0:dd-MMM-yyyy}", myAdapter.DS.Tables(0).Rows(0).Item("startdate"))
            End Select
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If IsNumeric(TextBox2.Text) Then
            myAdapter.GeneratePeriod(TextBox2.Text)
            myAdapter.Save()
            LoadData()
        Else
            TextBox2.Text = ""

        End If

    End Sub
End Class
