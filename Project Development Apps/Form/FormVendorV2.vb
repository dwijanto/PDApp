Imports System.Text
Imports System.Threading
Public Class FormVendorV2

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Public myAdapter As VendorAdapter
    Public Shared myFormVendor As FormVendorV2

    Public Shared Function getInstance()
        If myFormVendor Is Nothing Then
            myFormVendor = New FormVendorV2
            RemoveHandler DialogTopVendorInputV2.EndTransaction, AddressOf DialogTopVendorV2.RefreshDataGridView
            AddHandler DialogTopVendorInputV2.EndTransaction, AddressOf DialogTopVendorV2.RefreshDataGridView
        End If
        Return myFormVendor
    End Function

    Private Sub FormUser_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        myFormVendor = Nothing
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
        myAdapter = VendorAdapter.getInstance
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
            End Select
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ShowTx(TxRecord.UpdateRecord)
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        ShowTx(TxRecord.AddRecord)
    End Sub

    Private Sub ShowTx(ByVal StatusTx As TxRecord)
        Dim drv As DataRowView = Nothing
        Select Case StatusTx
            Case TxRecord.AddRecord
                drv = myAdapter.BS.AddNew
            Case TxRecord.UpdateRecord
                drv = myAdapter.BS.Current
        End Select
        Dim myform As New DialogVendorInputV2(drv)
        myform.Show()
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
            sb.Append(String.Format("shortname like '*{0}*'", obj.Text))
            myAdapter.BS.Filter = sb.ToString
        End If
    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click

        Dim VendorBS As New BindingSource
        Dim mytbl As New DataTable
        mytbl = TryCast(myAdapter.BS.DataSource, DataTable).Copy
        VendorBS.DataSource = mytbl

        Dim myform As DialogTopVendorV2 = DialogTopVendorV2.getInstance(VendorBS, myAdapter.topBS)
        myform.WindowState = FormWindowState.Normal
        myform.Show()
        myform.Activate()
    End Sub
End Class