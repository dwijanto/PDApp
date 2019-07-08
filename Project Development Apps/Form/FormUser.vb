Imports System.Threading
Imports System.Text
Imports Microsoft.Office.Interop

Public Enum TxRecord
    AddRecord = 0
    UpdateRecord = 1
    CancelRecord = 2
    DeleteRecord = 3
    ViewRecord = 4
End Enum

Public Class FormUser

    Dim SqlStrReport As String

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As UserAdapter
    Public Shared myForm As FormUser

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormUser
            AddHandler DialogUserInput.SaveRecord, AddressOf myForm.AssignRecord
            AddHandler DialogAddRemoveRole.SaveRecord, AddressOf myForm.TVAssignRecord
        ElseIf myForm.IsDisposed Then
            'myForm = Nothing
            myForm = New FormUser
        End If
        Return myForm
    End Function

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
        myAdapter = UserAdapter.getInstance
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
        Dim myform = New DialogUserInput(myAdapter.BS, TxRecord.UpdateRecord)
        myform.Show()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim myform As New DialogUserInput(myAdapter.BS, TxRecord.AddRecord)
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
            sb.Append(String.Format("userid like '*{0}*' or username like '*{0}*' or email like '*{0}*'", obj.Text))
            myAdapter.BS.Filter = sb.ToString
        End If
    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

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
        mydrv.Item("userid") = drv.Item("userid")
        mydrv.Item("username") = drv.Item("username")
        mydrv.Item("email") = drv.Item("email")
        mydrv.Item("isactive") = drv.Item("isactive")
        mydrv.Item("isadmin") = drv.Item("isadmin")
        mydrv.Item("title") = drv.Item("title")
        mydrv.EndEdit()
        DataGridView1.Invalidate()
    End Sub

    Private Sub RoleAssignmentToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RoleAssignmentToolStripMenuItem1.Click, ToolStripButton7.Click
        Dim myform = New DialogUserRole(myAdapter)
        myform.Show()
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Dim myform As DialogRole = DialogRole.getInstance(myAdapter)
        myform.Show()
        myform.Activate()
    End Sub

    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        SqlStrReport = String.Format("select u.username,u.userid,u.email,u.isadmin,u.isactive,pd.gettitle(u.title) as title,r.rolename from pd.user u " &
                                     " left join pd.user_role ur on ur.userid = u.id" &
                                     " left join pd.role r on r.id = ur.roleid" &
                                     " order by r.rolename,u.username,title;", "")
        loadReport()
    End Sub
    Private Sub loadReport()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""

            myThread = New Thread(AddressOf DoReport)
            myThread.SetApartmentState(ApartmentState.STA)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoReport()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Dim sqlstr As String = String.Empty
        Dim sb As New StringBuilder
        'DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        ProgressReport(11, "InitFilter")

        sqlstr = String.Format(SqlStrReport, sb.ToString.ToLower)

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("UserRole-{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            ' Dim myreport As New ExportToExcel(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\Templates\PrintingPromotionPrice.xltx", "A1", False, False, False)
            Dim myreport As New ExportToExcel(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet)
            myreport.Run(Me, New System.EventArgs)
        End If

        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub
    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent
        owb.Worksheets(1).select()
        Dim osheet As Excel.Worksheet = owb.Worksheets(1)
       
    End Sub
    Private Sub PivotTable()
        ' Throw New NotImplementedException
    End Sub

    Private Sub TVAssignRecord(ByRef sender As Object, e As UserEventArgs)
        Dim drv As DataRowView = DirectCast(sender, DataRowView)
        Dim mydrv As DataRowView = Nothing
        Try
            Select Case e.StatusTx
                Case TxRecord.AddRecord
                    mydrv = myAdapter.RoleBS.AddNew 'bsRole.AddNew()
                Case TxRecord.UpdateRecord
                    Dim result = myAdapter.RoleBS.Find("id", drv.Item("id")) 'bsRole.Find("id", drv.Item("id"))
                    mydrv = myAdapter.RoleBS.Item(result) 'bsRole.Item(result)
            End Select
            mydrv.Item("rolename") = drv.Item("rolename")

            mydrv.EndEdit()
            'DataGridView1.Invalidate()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub


    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click

    End Sub
End Class