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

    Public Property ProjectDrv As DataRowView


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
            
            Me.BS.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            RaiseEvent SaveRecord(BS.Current, New UserEventArgs(StatusTx))
            Me.Close()
        Catch ex As ConstraintException
            MessageBox.Show("Found duplicate data. Please check.")
            Me.BS.CancelEdit()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.BS.CancelEdit()
        End Try
        'Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        BS.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub


    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(ComboBox1, "")
        ErrorProvider1.SetError(ComboBox1, "")
        ErrorProvider1.SetError(RadioButton2, "")

        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
            myret = False
        Else
            Dim drv = ComboBox1.SelectedItem
            Dim curdrv = Me.BS.Current
            curdrv.row.item("pidname") = drv.Row.Item("projectid")
            curdrv.row.item("projectname") = drv.Row.Item("projectname")
            curdrv.row.item("shortname") = drv.row.item("shortname")
        End If

        If IsNothing(ComboBox2.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox2, "Please select from the list.")
            myret = False
        Else
            Dim drv = ComboBox2.SelectedItem
            Dim curdrv = Me.BS.Current
            curdrv.row.item("phase_status_name") = ""
            If drv.row.item("id") > 2 Then
                If (RadioButton1.Checked = False And RadioButton2.Checked = False) Then
                    ErrorProvider1.SetError(RadioButton2, "Please select one option.")
                    myret = False
                End If
                curdrv.row.item("phase_status_name") = IIf(PhaseStatus = 1, "Passed", "Failed")
                If PhaseStatus = 1 Then
                    TextBox1.Text = String.Empty
                End If
            Else
                curdrv.row.item("phase_status") = DBNull.Value
                curdrv.row.item("phase_status_name") = "n/a"
                curdrv.row.item("lognumber") = DBNull.Value
            End If
            'check ProjectId and shortname for Project_Phase_id > RP1 (1)
            If drv.row.item("id") > 1 Then
                If ProjectDrv.Item("projectid") = DBNull.Value Or ProjectDrv.Item("shortname") = DBNull.Value Then
                    'show interface
                    MessageBox.Show("Show Interface")
                End If
            End If
            curdrv.row.item("projectstage") = drv.row.item("input_name")

        End If


       
        Return myret
    End Function



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
            Panel1.Visible = True
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
     
        DateTimePicker1.DataBindings.Add(New Binding("value", Me.BS, "signed_date", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", Me.BS, "projectid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("SelectedValue", Me.BS, "project_phase_id", True, DataSourceUpdateMode.OnPropertyChanged))
        Me.DataBindings.Add(New Binding("PhaseStatus", Me.BS, "phase_status", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox1.DataBindings.Add(New Binding("text", Me.BS, "lognumber", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub


    Private Sub DialogUserInput_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub

    Sub onPropertyChanged(ByVal name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        onPropertyChanged("PhaseStatus")
        Panel1.Visible = False
        If PhaseStatus = 2 Then
            Panel1.Visible = True
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
                ShowRadioButton(drv.Row.Item("id") > 2)

        End Select

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myform = New DialogProjectHelper(Me.ProjectNameBS)
        If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim drv As DataRowView = myform.BS.Current
            Dim pos = ProjectNameBS.Find("id", drv.Item("id"))
            If pos > -1 Then
                Me.ProjectNameBS.Position = pos
                ComboBox1.SelectedItem = Me.ProjectNameBS.Current
                Application.DoEvents()
                Label4.Text = "" & drv.Item("projectid")
            Else
                MessageBox.Show("Record cannot be found.")
            End If
            

        End If
    End Sub

    Private Sub BS_ListChanged(sender As Object, e As ListChangedEventArgs) Handles BS.ListChanged
        If Not IsNothing(BS) Then
            If Not IsNothing(BS.Current) Then
                Panel1.Visible = False
                If PhaseStatus = 2 Then
                    Panel1.Visible = True
                End If
            End If
        End If
    End Sub
End Class
