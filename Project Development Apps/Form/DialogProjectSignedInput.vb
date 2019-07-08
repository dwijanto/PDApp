Imports System.ComponentModel

Public Class DialogProjectSignedInput
    Implements INotifyPropertyChanged


    Private WithEvents BS As BindingSource
    Private ProjectNameBS As BindingSource
    Private PhaseBS As BindingSource

    Public StatusTx As TxRecord
    Public Shared Event SaveRecord(ByRef sender As Object, ByVal e As EventArgs)
    Private titlelist As New List(Of UserTitle)
    Private userinfo1 As UserInfo
    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    Private _ProjectDrv As DataRowView
    Private _ProjectSignedDRV As DataRowView
    Private _myadapter As ProjectSignedAdapter

    Public Property PhaseStatus As Integer
        Get
            If RadioButton1.Checked = True Then
                Return 1
            ElseIf RadioButton2.Checked = True Then
                Return 2
            Else
                Return Nothing
            End If
        End Get
        Set(value As Integer)
            If value = 1 Then
                RadioButton1.Checked = True
            ElseIf value = 2 Then
                RadioButton2.Checked = True
            End If
        End Set
    End Property

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not Me.validate Then
            Exit Sub
        End If
        Try
            'reminder
            Dim drv As DataRowView = ComboBox2.SelectedItem
            If StatusTx = TxRecord.AddRecord Then
                If drv.Item("input_name").ToString.Contains("RP4") Then
                    Dim displaymessage As String = "RP4 Price"
                    If drv.Item("input_name") = displaymessage Then
                        displaymessage = "RP4 On Time"
                    End If
                    MessageBox.Show(String.Format("Just a reminder. Don't forget to update ""{0}"" also.", displaymessage))
                End If
            End If
            _ProjectSignedDRV.EndEdit()

            'Me.BS.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            'RaiseEvent SaveRecord(BS.Current, New UserEventArgs(StatusTx))
            Me.Close()
        Catch ex As ConstraintException
            MessageBox.Show("Found duplicate data. Please check.")
            'Me.BS.CancelEdit()
            _ProjectSignedDRV.CancelEdit()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            'Me.BS.CancelEdit()
            _ProjectSignedDRV.CancelEdit()
        End Try
        'Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        'BS.CancelEdit()
        _ProjectSignedDRV.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub


    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(ComboBox1, "")
        ErrorProvider1.SetError(ComboBox1, "")
        ErrorProvider1.SetError(RadioButton2, "")
        'Dim drvProject As DataRowView = Nothing
        Dim drvProjectPhase As DataRowView = Nothing


        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
            myret = False
        Else
            'drvProject = ComboBox1.SelectedItem
            _ProjectDrv = ComboBox1.SelectedItem

            'Dim curdrv = _ProjectSignedDRV 'Me.BS.Current
            'curdrv.Row.Item("pidname") = _ProjectDrv.Row.Item("projectid")
            'curdrv.Row.Item("projectname") = _ProjectDrv.Row.Item("projectname")
            'curdrv.Row.Item("shortname") = _ProjectDrv.Row.Item("shortname")
            'curdrv.Row.Item("vendorid") = _ProjectDrv.Row.Item("vendorid")
            'Dim curdrv = _ProjectSignedDRV 'Me.BS.Current
            _ProjectSignedDRV.Row.Item("pidname") = _ProjectDrv.Row.Item("projectid")
            _ProjectSignedDRV.Row.Item("projectname") = _ProjectDrv.Row.Item("projectname")
            _ProjectSignedDRV.Row.Item("shortname") = _ProjectDrv.Row.Item("shortname")
            _ProjectSignedDRV.Row.Item("vendorid") = _ProjectDrv.Row.Item("vendorid")
        End If

        If IsNothing(ComboBox2.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox2, "Please select from the list.")
            myret = False
        Else

            drvProjectPhase = ComboBox2.SelectedItem

            If Not IsDBNull(_ProjectDrv.Item("variant")) Then
                If _ProjectDrv.Item("variant") Then
                    If Not (drvProjectPhase.Item("id") = 3 Or drvProjectPhase.Item("id") = 7) Then
                        ErrorProvider1.SetError(ComboBox2, "Please select PSA/RRA or RAMP UP from the list for variant project.")
                        myret = False

                    End If
                End If
                
            End If
            'Dim curdrv = _ProjectSignedDRV 'Me.BS.Current
            'curdrv.row.item("phase_status_name") = ""
            _ProjectSignedDRV.Row.Item("phase_status_name") = ""
            If drvProjectPhase.Row.Item("id") > 2 Then
                If (RadioButton1.Checked = False And RadioButton2.Checked = False) Then
                    ErrorProvider1.SetError(RadioButton2, "Please select one option.")
                    myret = False
                End If
                'curdrv.row.item("phase_status_name") = IIf(PhaseStatus = 1, "Passed", "Failed")
                _ProjectSignedDRV.Row.Item("phase_status_name") = IIf(PhaseStatus = 1, "Passed", "Failed")
                If PhaseStatus = 1 Then
                    TextBox1.Text = String.Empty
                End If
            Else
                'curdrv.row.item("phase_status") = DBNull.Value
                'curdrv.row.item("phase_status_name") = "n/a"
                'curdrv.row.item("lognumber") = DBNull.Value
                _ProjectSignedDRV.Row.Item("phase_status") = DBNull.Value
                _ProjectSignedDRV.Row.Item("phase_status_name") = "n/a"
                _ProjectSignedDRV.Row.Item("lognumber") = DBNull.Value
            End If
            'check ProjectId and shortname for Project_Phase_id > RP1 (1)
            If drvProjectPhase.Row.Item("id") > 1 Then
                'Move record position to correct location (Find projectId)
                'Dim pk(0) As Object
                'pk(0) = _ProjectSignedDRV.Item("projectid")
                'Dim result = _myadapter.DS.Tables(1).Rows.Find(pk) '_myadapter.ProjectBS.Find("id", _ProjectSignedDRV.Item("projectid"))

                'Dim pos = _myadapter.ProjectBS.Find("id", _ProjectSignedDRV.Item("projectid"))
                'Dim result As DataRowView = _myadapter.ProjectBS.Current

                'If IsDBNull(result.Item("projectid")) Or IsDBNull(result.Item("shortname")) Then
                If (IsDBNull(_ProjectDrv.Item("projectid")) Or IsDBNull(_ProjectDrv.Item("shortname"))) And _ProjectDrv.Item("sbu_id") <> 6 Then
                    'show interface

                    Dim myform = New DialogProjectIDVendor(_ProjectDrv, _myadapter)
                    If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        myret = False
                    Else
                        'assign projectId
                        Me._ProjectSignedDRV.Item("pidname") = _ProjectDrv.Item("projectid")

                        Dim mypos = _myadapter.ProjectBS.Find("id", _ProjectSignedDRV.Item("projectid"))
                        _myadapter.ProjectBS.Position = mypos
                        Dim result As DataRowView = _myadapter.ProjectBS.Current
                        result.Item("projectid") = _ProjectDrv.Item("projectid")
                        _ProjectSignedDRV.Row.Item("vendorid") = _ProjectDrv.Item("vendorid")
                        _ProjectSignedDRV.Row.Item("shortname") = _ProjectDrv.Item("shortname")
                        ' result.Item("vendorid") = _ProjectDrv.Item("vendorid")
                        ' result.Item("shortname") = _ProjectDrv.Item("shortname")
                        'currprojectdrv.Item("projectid") = _ProjectDrv.Item("projectid")

                    End If



                End If
            End If

            'curdrv.Row.Item("projectstage") = drvProjectPhase.Row.Item("input_name")
            _ProjectSignedDRV.Row.Item("projectstage") = drvProjectPhase.Row.Item("input_name")

        End If



        Return myret
    End Function
    Public Sub New(ByRef drv As DataRowView, ByRef ProjectDrv As DataRowView, ByRef myadapter As ProjectSignedAdapter)
        InitializeComponent()
        userinfo1 = UserInfo.getInstance
        Me._myadapter = myadapter
        Me._ProjectSignedDRV = drv
        Me._ProjectDrv = ProjectDrv
        If Not IsDBNull(drv.Item("project_phase_id")) Then
            ShowRadioButton(drv.Item("project_phase_id") > 2)
        End If


        Dim mytbl As New DataTable
        mytbl = CType(myadapter.ProjectBS.DataSource, DataTable).Copy

        ProjectNameBS = New BindingSource
        PhaseBS = New BindingSource

        ProjectNameBS.DataSource = mytbl 'myadapter.ProjectBS
        PhaseBS.DataSource = myadapter.PhaseBS.DataSource

        InitData()
        Panel1.Visible = False

        If PhaseStatus = 2 Then
            If userinfo1.isAdmin Then
                Panel1.Visible = True

            End If

        End If
    End Sub


    Public Sub New(ByRef myadapter As ProjectSignedAdapter, ByVal StatusTx As TxRecord)

        '     This call is required by the designer.
        InitializeComponent()
        userinfo1 = UserInfo.getInstance
        Me.StatusTx = StatusTx
        Dim position = myadapter.BS.Position

        Dim mytbl As New DataTable
        mytbl = CType(myadapter.BS.DataSource, DataTable).Copy

        Me.BS = New BindingSource
        Me.BS.DataSource = mytbl

        ProjectNameBS = New BindingSource
        PhaseBS = New BindingSource

        ProjectNameBS = myadapter.ProjectBS
        PhaseBS.DataSource = myadapter.PhaseBS.DataSource
        Label4.Text = ""
        ShowRadioButton(False)

        If StatusTx = TxRecord.AddRecord Then
            Dim drv As DataRowView = BS.AddNew()
            drv.Row.Item("postingdate") = Date.Today
            drv.Row.Item("signed_date") = Date.Today
            drv.Row.Item("stamp") = Date.Now
            drv.Row.Item("modifiedby") = userinfo1.Userid
            drv.EndEdit()
        Else
            'Prevent incorrect row while in Sort Mode
            Dim drv As DataRowView = myadapter.BS.Current
            position = Me.BS.Find("id", drv.Item("id"))
            Label4.Text = "" & drv.Item("pidname")
            ShowRadioButton(drv.Item("project_phase_id") > 2)
            Me.BS.Position = position
        End If

        InitData()
        Panel1.Visible = False

        If PhaseStatus = 2 Then
            If userinfo1.isAdmin Then
                Panel1.Visible = True
            End If


        End If

        '     Add any initialization after the InitializeComponent() call.


    End Sub

    Private Sub ShowRadioButton(value As Boolean)
        RadioButton1.Visible = value
        RadioButton2.Visible = value
    End Sub
    Private Sub InitData()



        ComboBox1.DisplayMember = "projectname"
        ComboBox1.ValueMember = "id"
        ComboBox1.DataSource = ProjectNameBS


        ComboBox2.DisplayMember = "input_name"
        ComboBox2.ValueMember = "id"
        ComboBox2.DataSource = PhaseBS

        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        Me.DataBindings.Clear()

        DateTimePicker1.DataBindings.Add(New Binding("value", Me._ProjectSignedDRV, "signed_date", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", Me._ProjectSignedDRV, "projectid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", Me._ProjectSignedDRV, "project_phase_id", True, DataSourceUpdateMode.OnPropertyChanged))
        Me.DataBindings.Add(New Binding("PhaseStatus", Me._ProjectSignedDRV, "phase_status", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox1.DataBindings.Add(New Binding("text", Me._ProjectSignedDRV, "lognumber", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub


    Private Sub DialogUserInput_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not DialogResult = Windows.Forms.DialogResult.OK Then
            _ProjectSignedDRV.CancelEdit()
        End If
        Me.Dispose()
    End Sub

    Sub onPropertyChanged(ByVal name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        onPropertyChanged("PhaseStatus")
        Panel1.Visible = False

        If PhaseStatus = 2 Then
            If userinfo1.isAdmin Then
                Panel1.Visible = True
            End If


        End If
    End Sub

    Private Sub DialogProjectSignedInput_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted

        Select Case DirectCast(sender, ComboBox).Name
            Case "ComboBox1"
                Dim drv As DataRowView = ComboBox1.SelectedItem
                Label4.Text = "" & drv.Row.Item("projectid")

            Case "ComboBox2"
                Dim drv As DataRowView = ComboBox2.SelectedItem
                If Not IsNothing(drv) Then
                    ShowRadioButton(drv.Row.Item("id") > 2)
                End If


        End Select

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myform = New DialogProjectHelper(Me.ProjectNameBS)
        If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim CurrSelectedDrv As DataRowView = myform.BS.Current
            If Not IsNothing(CurrSelectedDrv) Then
                Dim pos = ProjectNameBS.Find("id", CurrSelectedDrv("id"))
                ProjectNameBS.Position = pos
                'ComboBox1.SelectedItem = CurrSelectedDrv
                'Application.DoEvents()
                Label4.Text = "" & CurrSelectedDrv.Item("projectid")
            End If


        End If
    End Sub
    Private Sub Button1_Click1(sender As Object, e As EventArgs)
        Dim myform = New DialogProjectHelper(Me.ProjectNameBS)
        If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim drv As DataRowView = _ProjectSignedDRV 'myform.BS.Current
            Dim pos = ProjectNameBS.Find("id", drv.Item("id"))
            If pos > -1 Then
                Me.ProjectNameBS.Position = pos
                Dim projectnamebscurrent As DataRowView = Me.ProjectNameBS.Current
                ComboBox1.SelectedItem = projectnamebscurrent 'Me.ProjectNameBS.Current
                Application.DoEvents()
                'Label4.Text = "" & drv.Item("projectid")
                'Label4.Text = "" & drv.Item("pidname")
                Label4.Text = "" & projectnamebscurrent.Item("projectid")
            Else
                MessageBox.Show("Record cannot be found.")
            End If


        End If
    End Sub

    'Private Sub BS_ListChanged(sender As Object, e As ListChangedEventArgs) Handles BS.ListChanged
    '    If Not IsNothing(BS) Then
    '        If Not IsNothing(BS.Current) Then
    '            Panel1.Visible = False
    '            If PhaseStatus = 2 Then
    '                Panel1.Visible = True
    '            End If
    '        End If
    '    End If
    'End Sub
End Class
