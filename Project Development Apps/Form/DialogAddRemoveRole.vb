Imports System.Windows.Forms

Public Class DialogAddRemoveRole
    Private BSRole As BindingSource
    Public Property StatusTx As TxRecord

    Public Shared Event SaveRecord(ByRef sender As Object, ByVal e As EventArgs)

    Public Sub New(ByRef bs As BindingSource, statustx As TxRecord)

        ' This call is required by the designer.
        InitializeComponent()
        Me.StatusTx = statustx
        Dim position = bs.Position

        Dim mytbl As New DataTable
        mytbl = CType(bs.DataSource, DataTable).Copy

        Me.BSRole = New BindingSource
        Me.BSRole.DataSource = mytbl ' = New BindingSource(bs.DataSource, bs.DataMember)


        If statustx = TxRecord.AddRecord Then
            Dim drv As DataRowView = BSRole.AddNew()
            'drv.EndEdit()
        Else
            'Prevent incorrect row while in Sort Mode
            Dim drv As DataRowView = bs.Current
            position = BSRole.Find("id", drv.Item("id"))
            Me.BSRole.Position = position
        End If

        InitData()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not Me.Validate Then
            Exit Sub
        End If
        Try
            BSRole.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            RaiseEvent SaveRecord(BSRole.Current, New UserEventArgs(StatusTx))

        Catch ex As Exception

            MessageBox.Show(ex.Message)
            BSRole.CancelEdit()
        End Try

        Me.Close()


    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Private Sub InitData()

        TextBox1.DataBindings.Clear()       
        TextBox1.DataBindings.Add(New Binding("text", BSRole, "rolename", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    

    End Sub
    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(TextBox1, "")
       

        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Role Name cannot be blank.")
            myret = False
        End If
      
        Return myret
    End Function

End Class
