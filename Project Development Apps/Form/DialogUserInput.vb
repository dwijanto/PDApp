Imports System.Windows.Forms
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Public Class DialogUserInput
    Private _bs As BindingSource
    Public StatusTx As TxRecord
    Public Shared Event SaveRecord(ByRef sender As Object, ByVal e As EventArgs)
    Private titlelist As New List(Of UserTitle)
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not Me.validate Then
            Exit Sub
        End If
        Try
            _bs.EndEdit()
            'RaiseEvent SaveRecord(_bs.Current, New EventArgs)
            'drv = myUserAdapter.BS.Current
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            RaiseEvent SaveRecord(_bs.Current, New UserEventArgs(StatusTx))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            _bs.CancelEdit()
        End Try

        
        Me.Close()
        'StatusTx = TxRecord.UpdateRecord

    End Sub
    'Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    '    If Not Me.validate Then
    '        Exit Sub
    '    End If
    '    _bs.EndEdit()
    '    'RaiseEvent SaveRecord(_bs.Current, New EventArgs)
    '    'drv = myUserAdapter.BS.Current
    '    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    '    Me.Close()
    '    StatusTx = TxRecord.UpdateRecord
    '    RaiseEvent SaveRecord(_bs.Current, New UserEventArgs(StatusTx))
    'End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        _bs.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
        'If StatusTx = TxRecord.AddRecord Then
        '    StatusTx = TxRecord.DeleteRecord
        'ElseIf StatusTx = TxRecord.UpdateRecord Then
        '    StatusTx = TxRecord.CancelRecord
        'End If
        'RaiseEvent SaveRecord(_bs.Current, New UserEventArgs(StatusTx))
    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(TextBox1, "")
        ErrorProvider1.SetError(TextBox2, "")
        ErrorProvider1.SetError(TextBox3, "")

        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Userid cannot be blank.")
            myret = False
        End If
        If TextBox2.Text = "" Then
            ErrorProvider1.SetError(TextBox2, "User Name cannot be blank.")
            myret = False
        End If
        If TextBox3.Text = "" Then
            ErrorProvider1.SetError(TextBox3, "email cannot be blank.")
            myret = False
        End If
        'If IsNothing(ComboBox1.SelectedItem) Then
        '    ErrorProvider1.SetError(ComboBox1, "Please select from the list.")
        '    myret = False
        'End If
        If ComboBox1.Text = "" Then
            ErrorProvider1.SetError(ComboBox1, "Title cannot be blank.")
            myret = False
        End If
        Return myret
    End Function



    Public Sub New(ByRef bs As BindingSource, ByVal StatusTx As TxRecord)

        '     This call is required by the designer.
        InitializeComponent()
        Me.StatusTx = StatusTx
        Dim position = bs.Position

        Dim mytbl As New DataTable
        mytbl = CType(bs.DataSource, DataTable).Copy

        Me._bs = New BindingSource
        Me._bs.DataSource = mytbl ' = New BindingSource(bs.DataSource, bs.DataMember)


        If StatusTx = TxRecord.AddRecord Then
            Dim drv As DataRowView = _bs.AddNew()
            drv.Row.Item("isactive") = True
            drv.Row.Item("isadmin") = False
            'drv.EndEdit()
        Else
            'Prevent incorrect row while in Sort Mode
            Dim drv As DataRowView = bs.Current
            position = _bs.Find("id", drv.Item("id"))
            Me._bs.Position = position
        End If

        InitData()

        '     Add any initialization after the InitializeComponent() call.


    End Sub
    Public Sub New(ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        Dim position = bs.Position
        Me._bs = New BindingSource(bs.DataSource, bs.DataMember)
        Me._bs.Position = position
        InitData()

        ' Add any initialization after the InitializeComponent() call.


    End Sub
    'Public Sub New(ByRef bs As BindingSource, ByVal addrecord As Integer)

    '    ' This call is required by the designer.
    '    InitializeComponent()
    '    Me._bs = New BindingSource(bs.DataSource, bs.DataMember)
    '    Me._bs.AddNew()
    '    InitData()

    '    ' Add any initialization after the InitializeComponent() call.

    'End Sub
    'Public Sub New(ByVal id As Integer)

    '    ' This call is required by the designer.
    '    InitializeComponent()
    '    Me._myId = id
    '    StatusTx = TxRecord.UpdateRecord
    '    If _myId = 0 Then
    '        StatusTx = TxRecord.AddRecord
    '    End If
    '    InitDataId()

    '    ' Add any initialization after the InitializeComponent() call.

    'End Sub
    Private Sub InitData()

        titlelist.Clear()
        titlelist.Add(New UserTitle(3, "PJM"))
        titlelist.Add(New UserTitle(4, "PDM"))
        titlelist.Add(New UserTitle(5, "SPE"))
        titlelist.Add(New UserTitle(2, "SPDM"))
        titlelist.Add(New UserTitle(1, "Director"))
        titlelist.Add(New UserTitle(0, "Admin"))

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        CheckBox1.DataBindings.Clear()
        CheckBox2.DataBindings.Clear()

        ComboBox1.DataSource = titlelist
        ComboBox1.ValueMember = "id"
        ComboBox1.DisplayMember = "titlename"

        ComboBox1.DataBindings.Clear()

        TextBox1.DataBindings.Add(New Binding("text", _bs, "userid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("text", _bs, "username", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("text", _bs, "email", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        CheckBox1.DataBindings.Add(New Binding("checked", _bs, "isactive", True, DataSourceUpdateMode.OnPropertyChanged))
        CheckBox2.DataBindings.Add(New Binding("checked", _bs, "isadmin", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", _bs, "title", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub

    'Private Sub InitDataId()
    '    myUserAdapter = New UserAdapter
    '    _bs = myUserAdapter.getData(_myId)

    '    TextBox1.DataBindings.Clear()
    '    TextBox2.DataBindings.Clear()
    '    TextBox3.DataBindings.Clear()


    '    TextBox1.DataBindings.Add(New Binding("text", _bs, "userid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    '    TextBox2.DataBindings.Add(New Binding("text", _bs, "username", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    '    TextBox3.DataBindings.Add(New Binding("text", _bs, "email", True, DataSourceUpdateMode.OnPropertyChanged, ""))

    'End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myActiveDirectory As New ActiveDirectory
        Try
            If myActiveDirectory.getDataAD(TextBox1.Text) Then
                TextBox1.Text = myActiveDirectory.UserInfo.Domain & "\" & myActiveDirectory.UserInfo.Userid
                TextBox2.Text = myActiveDirectory.UserInfo.DisplayName
                TextBox3.Text = myActiveDirectory.UserInfo.email
            End If
        Catch ex As Exception
            MessageBox.Show("User not found in AD.")
        End Try
    End Sub




    Private Sub DialogUserInput_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub
End Class

Public Class UserTitle
    Public Property id As Integer
    Public Property titlename As String

    Public Sub New(id As Integer, titlename As String)
        Me.id = id
        Me.titlename = titlename
    End Sub
    Public Shadows Function tostring() As String
        Return titlename
    End Function

End Class