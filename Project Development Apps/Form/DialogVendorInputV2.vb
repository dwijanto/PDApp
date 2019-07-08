Imports System.Windows.Forms

Public Class DialogVendorInputV2

    Private BSVendor As BindingSource
    Public Property StatusTx As TxRecord
    Public Shared Event SaveRecord(ByRef sender As Object, ByVal e As EventArgs)
    Private DRV As DataRowView


    Public Sub New(ByVal drv As DataRowView)
        ' This call is required by the designer.
        InitializeComponent()
        Me.DRV = drv
        InitDataDRV()
    End Sub

    Public Sub New(ByRef bs As BindingSource, statustx As TxRecord)

        ' This call is required by the designer.
        InitializeComponent()
        Me.StatusTx = statustx
        Dim position = bs.Position

        Dim mytbl As New DataTable
        mytbl = CType(bs.DataSource, DataTable).Copy

        Me.BSVendor = New BindingSource
        Me.BSVendor.DataSource = mytbl ' = New BindingSource(bs.DataSource, bs.DataMember)


        If statustx = TxRecord.AddRecord Then
            Dim drv As DataRowView = BSVendor.AddNew()
            'drv.EndEdit()
        Else
            'Prevent incorrect row while in Sort Mode
            Dim drv As DataRowView = bs.Current
            position = BSVendor.Find("id", drv.Item("id"))
            Me.BSVendor.Position = position
        End If

        InitData()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not Me.Validate Then
            Exit Sub
        End If
        Try
            DRV.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            DRV.CancelEdit()
        End Try
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        DRV.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Private Sub InitData()

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox1.DataBindings.Add(New Binding("text", BSVendor, "shortname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("text", BSVendor, "panel", True, DataSourceUpdateMode.OnPropertyChanged, ""))


    End Sub
    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(TextBox1, "")
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Short Name cannot be blank.")
            myret = False
        End If

        Return myret
    End Function

    Private Sub InitDataDRV()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox1.DataBindings.Add(New Binding("text", DRV, "shortname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("text", DRV, "panel", True, DataSourceUpdateMode.OnPropertyChanged, ""))
    End Sub
End Class
