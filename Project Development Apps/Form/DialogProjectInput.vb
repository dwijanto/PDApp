Imports System.Windows.Forms
Imports System.ComponentModel
Public Class DialogProjectInput
    Implements INotifyPropertyChanged
    Public StatusTx As TxRecord
    Private WithEvents BS As BindingSource

    Private SBUBS As BindingSource
    Private VendorBS As BindingSource
    Private ProjectStatusBS As BindingSource
    Private RoleBS As BindingSource
    Private PICBS As BindingSource

    Public myadapter As ProjectAdapter
    Public Shared Event SaveRecord(ByRef sender As Object, ByVal e As EventArgs)
    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged
    Private _myval As Integer
    Private myUserInfo As UserInfo

    Private SelectedStatus As String = String.Empty

    Public Property ProjectType As Integer
        Get
            If RadioButton1.Checked = True Then
                Return 1
            ElseIf RadioButton2.Checked = True Then
                Return 2
            Else
                Return 3
            End If
        End Get
        Set(value As Integer)
            If value = 1 Then
                RadioButton1.Checked = True
            ElseIf value = 2 Then
                RadioButton2.Checked = True
            ElseIf value = 3 Then
                RadioButton3.Checked = True
            End If
        End Set
    End Property
    

    Public Sub New(ByRef myadapter As ProjectAdapter, ByVal StatusTx As TxRecord)

        'This call is required by the designer.
        InitializeComponent()
        myUserInfo = UserInfo.getInstance
        Me.StatusTx = StatusTx

        Dim mytbl As New DataTable
        mytbl = CType(myadapter.BS.DataSource, DataTable).Copy

        Me.BS = New BindingSource
        Me.BS.DataSource = mytbl

        Me.SBUBS = New BindingSource
        Me.VendorBS = New BindingSource
        Me.ProjectStatusBS = New BindingSource
        Me.RoleBS = New BindingSource
        Me.PICBS = New BindingSource

        Me.SBUBS.DataSource = myadapter.SBUBS
        Me.VendorBS.DataSource = myadapter.VendorBS
        Me.ProjectStatusBS.DataSource = myadapter.ProjectStatusBS
        Me.RoleBS.DataSource = myadapter.RoleBS
        Me.PICBS.DataSource = myadapter.PICBS

        If StatusTx = TxRecord.AddRecord Then
            Dim drv As DataRowView = Me.BS.AddNew()
            'find user
            Dim position = PICBS.Find("userid", myUserInfo.Userid.ToLower)
            drv.Item("modifiedby") = myUserInfo.Userid.ToLower
            drv.Item("creationdate") = Date.Now
            drv.Item("countstartdate") = Today.Date
            If Me.RoleBS.Count = 1 Then
                drv.Item("role") = TryCast(RoleBS.Item(0), DataRowView).Item("id")
            End If
            If position >= 0 Then
                drv.Item("pic") = TryCast(PICBS.Item(position), DataRowView).Item("id")
            End If
            drv.EndEdit()

        Else
            Dim drv As DataRowView = myadapter.BS.Current
            Dim position = Me.BS.Find("id", drv.Item("id"))
            Me.BS.Position = position
        End If
        'DateTimePicker1.Visible = myUserInfo.isAdmin
        'Label9.Visible = myUserInfo.isAdmin
        InitData()
        '     Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If StatusTx = TxRecord.ViewRecord Then
            Me.BS.CancelEdit()
            Me.Close()
        Else
            If Not Me.validate Then
                Exit Sub
            End If
            Try
                Me.BS.EndEdit()
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                RaiseEvent SaveRecord(BS.Current, New UserEventArgs(StatusTx))
                Me.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Me.BS.CancelEdit()

            End Try
        End If
        
    End Sub


    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        If StatusTx = TxRecord.ViewRecord Then
            Me.BS.CancelEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
        Else
            Me.BS.CancelEdit()
            Me.BS.RemoveCurrent()
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
        End If
       
    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        Dim drv As DataRowView = Me.BS.Current
        ErrorProvider1.SetError(MaskedTextBox1, "")
        ErrorProvider1.SetError(TextBox1, "")
        ErrorProvider1.SetError(RadioButton3, "")
        ErrorProvider1.SetError(ComboBox1, "")
        ErrorProvider1.SetError(ComboBox2, "")
        ErrorProvider1.SetError(ComboBox4, "")
        ErrorProvider1.SetError(ComboBox3, "")
        ErrorProvider1.SetError(ComboBox5, "")
        ErrorProvider1.SetError(CheckBox1, "")
        ErrorProvider1.SetError(CheckBox2, "")

        If CheckBox1.Checked Or ComboBox1.SelectedValue <> 6 Then
            'If MaskedTextBox1.Text = "00-0000-00" Then
            '    ErrorProvider1.SetError(MaskedTextBox1, "ProjectId cannot be blank.")
            '    myret = False
            'End If
            If IsNothing(ComboBox2.SelectedItem) Then
                ErrorProvider1.SetError(ComboBox2, "Please select from list.")
                myret = False
            Else
                drv.Row.Item("shortname") = DirectCast(ComboBox2.SelectedItem, DataRowView).Row.Item("shortname")
            End If
        End If
        

        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Project Name cannot be blank.")
            myret = False
        End If

        If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False Then
            ErrorProvider1.SetError(RadioButton3, "Select one option.")
            myret = False
        End If

        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from list.")
            myret = False
       
        End If

        If IsNothing(ComboBox3.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox3, "Please select from list.")
            myret = False
        Else
            drv.Row.Item("statusname") = DirectCast(ComboBox3.SelectedItem, DataRowView).Row.Item("statusname")
        End If
        If IsNothing(ComboBox4.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox4, "Please select from list.")
            myret = False
        Else
            drv.Row.Item("rolename") = DirectCast(ComboBox4.SelectedItem, DataRowView).Row.Item("rolename")
        End If
        
        If IsNothing(ComboBox5.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox5, "Please select from list.")
            myret = False
        Else
            drv.Row.Item("username") = DirectCast(ComboBox5.SelectedItem, DataRowView).Row.Item("username")
        End If

        If SelectedStatus <> "" Then
            If "Stop,End".Contains(SelectedStatus) Then
                drv.Row.Item("countenddate") = DateTimePicker2.Value
                Panel1.Visible = True
            Else
                Panel1.Visible = False
                drv.Row.Item("countenddate") = DBNull.Value
            End If
        End If

        If CheckBox1.Checked And CheckBox2.Checked Then
            ErrorProvider1.SetError(CheckBox1, "Variant and Count Project cannot be enabled together. Please select One!")
            ErrorProvider1.SetError(CheckBox2, "Variant and Count Project cannot be enabled together. Please select One!")
            myret = False
        End If



        Return myret
    End Function


    Private Sub InitData()

        ComboBox1.DataSource = Me.SBUBS
        ComboBox1.DisplayMember = "sbuname"
        ComboBox1.ValueMember = "id"

        ComboBox2.DataSource = Me.VendorBS
        ComboBox2.DisplayMember = "shortname"
        ComboBox2.ValueMember = "id"

        ComboBox3.DataSource = Me.ProjectStatusBS
        ComboBox3.DisplayMember = "statusname"
        ComboBox3.ValueMember = "id"

        ComboBox4.DataSource = Me.RoleBS
        ComboBox4.DisplayMember = "rolename"
        ComboBox4.ValueMember = "id"

        ComboBox5.DataSource = Me.PICBS
        ComboBox5.DisplayMember = "username"
        ComboBox5.ValueMember = "id"

        MaskedTextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        CheckBox1.DataBindings.Clear()
        CheckBox2.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()
        ComboBox4.DataBindings.Clear()
        ComboBox5.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()

        MaskedTextBox1.DataBindings.Add(New Binding("text", Me.BS, "projectid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox1.DataBindings.Add(New Binding("text", Me.BS, "projectname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        Me.DataBindings.Add(New Binding("ProjectType", Me.BS, "project_type_id", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", Me.BS, "sbu_id", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", Me.BS, "vendorid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", Me.BS, "project_status", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox4.DataBindings.Add(New Binding("SelectedValue", Me.BS, "role", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox5.DataBindings.Add(New Binding("SelectedValue", Me.BS, "pic", True, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox1.DataBindings.Add(New Binding("checked", Me.BS, "variant", True, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox2.DataBindings.Add(New Binding("checked", Me.BS, "countasrunningproject", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("text", Me.BS, "remarks", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker1.DataBindings.Add(New Binding("text", Me.BS, "countstartdate", True, DataSourceUpdateMode.OnPropertyChanged))
        DateTimePicker2.DataBindings.Add(New Binding("text", Me.BS, "countenddate", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox3.DataBindings.Add(New Binding("text", Me.BS, "lognumber", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    End Sub


    Private Sub DialogProjectInput_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub

    Sub onPropertyChanged(ByVal name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton3.CheckedChanged, RadioButton2.CheckedChanged
        onPropertyChanged("ProjectType")
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        Panel2.Visible = CheckBox2.Checked
        If myUserInfo.isAdmin Then

        End If
    End Sub

    Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox3.SelectionChangeCommitted
        Dim dr As DataRowView = ComboBox3.SelectedItem
        If Not IsNothing(dr) Then
            SelectedStatus = dr.Item("statusname")
            Panel1.Visible = False
            If SelectedStatus = "Stopped" Or SelectedStatus = "End" Then
                Panel1.Visible = True
            End If
        End If
        
    End Sub

    Private Sub BS_ListChanged(sender As Object, e As ListChangedEventArgs) Handles BS.ListChanged
        CheckPanel1()       

    End Sub

    Private Sub DialogProjectInput_Load(sender As Object, e As EventArgs) Handles Me.Load
        Panel3.Visible = False
        Panel2.Visible = CheckBox2.Checked
        If myUserInfo.isAdmin Then

            Panel3.Visible = True
        End If
        CheckPanel1()
        If StatusTx = TxRecord.ViewRecord Then
            GroupBox1.Enabled = False
        End If
    End Sub

    Public Sub CheckPanel1()
        If Not IsNothing(BS) Then
            If Not IsNothing(BS.Current) Then
                Dim drv As DataRowView = BS.Current
                Dim statusname = "" & drv.Item("statusname")
                Panel1.Visible = (statusname = "Stopped" Or statusname = "End")
            End If

        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Panel2.Visible = CheckBox2.Checked
    End Sub


End Class
