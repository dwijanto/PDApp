﻿Imports System.Windows.Forms
Imports System.ComponentModel
Public Class DialogProjectInput3

    Implements INotifyPropertyChanged
    Public StatusTx As TxRecord
    Private WithEvents BS As BindingSource

    Private SBUBS As BindingSource
    Private VendorBS As BindingSource
    Private ProjectStatusBS As BindingSource
    Private RoleBS As BindingSource
    Private PICBS As BindingSource
    Private SUBSBUBS As BindingSource
    Private CategoriesBS As BindingSource
    Private PTypeBS As BindingSource
    Private QualityLevelBS As BindingSource


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
            ElseIf RadioButton3.Checked = True Then
                Return 3
            Else
                Return 4
            End If
        End Get
        Set(value As Integer)
            If value = 1 Then
                RadioButton1.Checked = True
            ElseIf value = 2 Then
                RadioButton2.Checked = True
            ElseIf value = 3 Then
                RadioButton3.Checked = True
            ElseIf value = 4 Then
                RadioButton4.Checked = True
            End If
        End Set
    End Property

    Public Property StartFromRP1 As Boolean
        Get
            Return RP1RadioButton.Checked
        End Get
        Set(value As Boolean)
            If value Then
                RP1RadioButton.Checked = True
            Else
                RP3RadioButton.Checked = True
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
        Me.SUBSBUBS = New BindingSource
        Me.CategoriesBS = New BindingSource
        Me.PTypeBS = New BindingSource
        Me.QualityLevelBS = New BindingSource

        Me.SBUBS.DataSource = myadapter.SBUBS
        Me.VendorBS.DataSource = myadapter.VendorBS
        Me.ProjectStatusBS.DataSource = myadapter.ProjectStatusBS
        Me.RoleBS.DataSource = myadapter.RoleBS
        Me.PICBS.DataSource = myadapter.PICBS
        Me.SUBSBUBS.DataSource = myadapter.SubSBUBS
        Me.CategoriesBS.DataSource = myadapter.CategoriesBS
        Me.PTypeBS.DataSource = myadapter.PTypeBS
        Me.QualityLevelBS.DataSource = myadapter.QualityLevelBS

        If StatusTx = TxRecord.AddRecord Then
            Dim drv As DataRowView = Me.BS.AddNew()
            'find user
            Dim position = PICBS.Find("userid", myUserInfo.Userid.ToLower)
            drv.Item("modifiedby") = myUserInfo.Userid.ToLower
            drv.Item("creationdate") = Date.Now
            drv.Item("countstartdate") = Today.Date
            drv.Item("countasrp1") = True
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
        ErrorProvider1.SetError(RP3RadioButton, "")

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

        If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False Then
            ErrorProvider1.SetError(RadioButton4, "Select one option.")
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

        If CheckBox2.Checked Then
            If (Not RP1RadioButton.Checked) And (Not RP3RadioButton.Checked) Then
                ErrorProvider1.SetError(RP3RadioButton, "Please select one.")
                myret = False
            End If
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

        ComboBox6.DataSource = Me.SUBSBUBS
        ComboBox6.DisplayMember = "subsbuname"
        ComboBox6.ValueMember = "subsbuid"

        ComboBox7.DataSource = Me.CategoriesBS
        ComboBox7.DisplayMember = "categories"
        ComboBox7.ValueMember = "id"

        ComboBox8.DataSource = Me.PTypeBS
        ComboBox8.DisplayMember = "ptype"
        ComboBox8.ValueMember = "id"

        ComboBox9.DataSource = Me.QualityLevelBS
        ComboBox9.DisplayMember = "qualitylevel"
        ComboBox9.ValueMember = "id"

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

        ComboBox6.DataBindings.Clear()
        ComboBox7.DataBindings.Clear()
        ComboBox8.DataBindings.Clear()
        ComboBox9.DataBindings.Clear()

        DateTimePicker1.DataBindings.Clear()

        MaskedTextBox1.DataBindings.Add(New Binding("text", Me.BS, "projectid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox1.DataBindings.Add(New Binding("text", Me.BS, "projectname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        Me.DataBindings.Add(New Binding("ProjectType", Me.BS, "project_type_id", True, DataSourceUpdateMode.OnPropertyChanged))
        Me.DataBindings.Add(New Binding("StartFromRP1", Me.BS, "countasrp1", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", Me.BS, "sbu_id", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", Me.BS, "vendorid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox3.DataBindings.Add(New Binding("SelectedValue", Me.BS, "project_status", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox4.DataBindings.Add(New Binding("SelectedValue", Me.BS, "role", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox5.DataBindings.Add(New Binding("SelectedValue", Me.BS, "pic", True, DataSourceUpdateMode.OnPropertyChanged))

        ComboBox6.DataBindings.Add(New Binding("SelectedValue", Me.BS, "subsbuid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox7.DataBindings.Add(New Binding("SelectedValue", Me.BS, "categoryid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox8.DataBindings.Add(New Binding("SelectedValue", Me.BS, "ptypeid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox9.DataBindings.Add(New Binding("SelectedValue", Me.BS, "qualitylevelid", True, DataSourceUpdateMode.OnPropertyChanged))

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

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton3.CheckedChanged, RadioButton2.CheckedChanged, RadioButton4.CheckedChanged
        onPropertyChanged("ProjectType")
    End Sub

    Private Sub RP1RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles RP1RadioButton.CheckedChanged
        onPropertyChanged("StartFromRP1")
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
            'Panel3.Visible = True
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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim mydrv As DataRowView = ComboBox1.SelectedItem
        If Not IsNothing(mydrv) Then
            If mydrv.Item("sbuname") = "Cookware" Then
                GroupBox2.Visible = True

                If ComboBox6.Visible = False Then
                    ComboBox6.Visible = True
                    If ComboBox6.SelectedIndex < 0 Then
                        ComboBox6.SelectedIndex = 0
                    End If

                End If
                If ComboBox7.Visible = False Then
                    ComboBox7.Visible = True
                    If ComboBox7.SelectedIndex < 0 Then
                        ComboBox7.SelectedIndex = 0
                    End If

                End If
                If ComboBox8.Visible = False Then
                    ComboBox8.Visible = True
                    If ComboBox8.SelectedIndex < 0 Then
                        ComboBox8.SelectedIndex = 0
                    End If

                End If
                If ComboBox9.Visible = False Then
                    ComboBox9.Visible = True
                    If ComboBox9.SelectedIndex < 0 Then
                        ComboBox9.SelectedIndex = 0
                    End If

                End If

            Else
                GroupBox2.Visible = False

                ComboBox6.Visible = False
                ComboBox6.SelectedIndex = -1
                ComboBox7.Visible = False
                ComboBox7.SelectedIndex = -1
                ComboBox8.Visible = False
                ComboBox8.SelectedIndex = -1
                ComboBox9.Visible = False
                ComboBox9.SelectedIndex = -1
                
            End If
        End If

    End Sub
End Class
