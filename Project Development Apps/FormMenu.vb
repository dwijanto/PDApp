Imports System.Reflection
Public Class FormMenu

    Private dbAdapter1 As DbAdapter
    Private userinfo1 As UserInfo
    Dim HasError As Boolean = True

    Private Sub FormMenu_Load(ByVal sender As Object, ByVal e As System.EventArgs) ' Handles Me.Load !! Remarks is on purpose
        If HasError Then
            Me.Close()
            Exit Sub
        End If

        Try
            'HelperClass1 = New HelperClass
            dbAdapter1 = DbAdapter.getInstance

            'HelperClass1.UserInfo.IsAdmin = DbAdapter1.IsAdmin(HelperClass1.UserId)
            'HelperClass1.UserInfo.AllowUpdateDocument = DbAdapter1.AllowUpdateDocument(HelperClass1.UserId)

            loglogin(dbAdapter1.userid)

            Me.Text = GetMenuDesc()
            Me.Location = New Point(300, 10)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try

    End Sub
    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "PD Team Apps"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub
    Public Function GetMenuDesc() As String
        'Label1.Text = "Welcome, " & HelperClass1.UserInfo.DisplayName
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & dbAdapter1.ConnectionStringDict.Item("HOST") & ", Database: " & dbAdapter1.ConnectionStringDict.Item("DATABASE") & ", Userid: " & dbAdapter1.UserInfo.Userid 'HelperClass1.UserId

    End Function
    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormMenu))
        Dim frm As Object = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        Dim myform = frm.GetInstance
        'Dim myform = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Object).GetInstance
        myform.show()
        myform.windowstate = FormWindowState.Normal
        myform.activate()
    End Sub

    Private Sub ExecuteForm(ByVal obj As Windows.Forms.Form)
        With obj
            .WindowState = FormWindowState.Normal
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
            .Focus()
        End With
    End Sub

    Private Sub FormMenu_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.FormMenu_Load(Me, New EventArgs)
        AddHandler UserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler VendorToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ProjectMaintenanceToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ProjectStageStatusToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler FailedStatusAdjustmentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ShowOpenUserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler GeneratePeriodToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'Admin
        MasterToolStripMenuItem.Visible = userinfo1.isAdmin
        ReportToolStripMenuItem.Visible = userinfo1.isAdmin


    End Sub

    Private Sub FormMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.CloseReason = CloseReason.ApplicationExitCall Then
            If MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                If ValidateRP4() Then
                    Me.CloseOpenForm()
                    dbAdapter1.Dispose()
                Else
                    e.Cancel = True
                End If

            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Private Sub CloseOpenForm()
        For i = 1 To (My.Application.OpenForms.Count - 1)
            My.Application.OpenForms.Item(1).Close()
        Next
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        Try
            userinfo1 = UserInfo.getInstance
            userinfo1.Userid = Environment.UserDomainName & "\" & Environment.UserName
            userinfo1.computerName = My.Computer.Name
            userinfo1.ApplicationName = "PD Team Apps"
            userinfo1.Username = "N/A"
            userinfo1.isAuthenticate = False
            userinfo1.Role = 0

            dbAdapter1 = DbAdapter.getInstance
            dbAdapter1.UserInfo = userinfo1
            dbAdapter1.UserInfo.isAdmin = dbAdapter1.IsAdmin(userinfo1.Userid)

            HasError = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
       
    End Sub

    Private Sub ClosingPeriodToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClosingPeriodToolStripMenuItem.Click
        If MessageBox.Show("Close this period?", "Closing Period.", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.OK Then
            If dbAdapter1.CloseConfirmation(userinfo1.Userid.ToLower) Then
                MessageBox.Show("OK")
            End If


        End If
    End Sub

    Private Sub ReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportToolStripMenuItem.Click

    End Sub


    Private Sub FailedStatusAdjustmentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FailedStatusAdjustmentToolStripMenuItem.Click

    End Sub

    Private Sub RawDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RawDataToolStripMenuItem.Click
        Dim myform As FormReportKPI = FormReportKPI.getInstance
        myform.Show()
        myform.Activate()
    End Sub

    Private Sub GeneratePeriodToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GeneratePeriodToolStripMenuItem.Click

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

    Private Sub SaveHistoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveHistoryToolStripMenuItem.Click
        Dim myform As FormSaveHistory = FormSaveHistory.getInstance
        myform.Show()
        myform.Activate()
    End Sub

    Private Sub GuidelineToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GuidelineToolStripMenuItem.Click
        Dim p As New System.Diagnostics.Process
        'p.StartInfo.FileName = "\\172.22.10.44\SharedFolder\PriceCMMF\New\template\Supplier Management Task User Guide-Assets Purchase.pdf"
        p.StartInfo.FileName = Application.StartupPath & "\help\PD scoreboard user guideline.pdf"
        p.Start()
    End Sub



    Private Sub ProjectMaintenanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProjectMaintenanceToolStripMenuItem.Click

    End Sub
End Class