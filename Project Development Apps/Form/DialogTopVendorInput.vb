Imports System.Windows.Forms

Public Class DialogTopVendorInput
    Private BSTopVendor As BindingSource
    Private VendorBS As BindingSource
    Public Property StatusTx As TxRecord

    Public Shared Event SaveRecord(ByRef sender As Object, ByVal e As EventArgs)

    Public Sub New(ByRef bs As BindingSource, ByRef VendorBS As BindingSource, statustx As TxRecord)

        ' This call is required by the designer.
        InitializeComponent()
        Me.StatusTx = statustx
        Dim position = bs.Position

        Dim mytbl As New DataTable
        mytbl = CType(bs.DataSource, DataTable).Copy

        Me.BSTopVendor = New BindingSource
        Me.BSTopVendor.DataSource = mytbl ' = New BindingSource(bs.DataSource, bs.DataMember)

        mytbl = CType(VendorBS.DataSource, DataTable).Copy
        ' Me.VendorBS = New BindingSource(VendorBS.DataSource, VendorBS.DataMember)
        Me.VendorBS = New BindingSource
        Me.VendorBS.DataSource = mytbl

        If statustx = TxRecord.AddRecord Then
            Dim drv As DataRowView = Me.BSTopVendor.AddNew()
            'drv.EndEdit()
        Else
            'Prevent incorrect row while in Sort Mode
            Dim drv As DataRowView = bs.Current
            position = Me.BSTopVendor.Find("id", drv.Item("id"))
            Me.BSTopVendor.Position = position
        End If

        InitData()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not Me.Validate Then
            Exit Sub
        End If
        Try

            BSTopVendor.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            RaiseEvent SaveRecord(BSTopVendor.Current, New UserEventArgs(StatusTx))
        Catch ex As ConstraintException
            MessageBox.Show("Duplicate data found. Please check.")
            BSTopVendor.CancelEdit()
            Exit Sub
        Catch ex As Exception

            MessageBox.Show(ex.Message)
            BSTopVendor.CancelEdit()
            Exit Sub
        End Try

        Me.Close()


    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Private Sub InitData()

        ComboBox1.DataSource = VendorBS
        ComboBox1.DisplayMember = "shortname"
        ComboBox1.ValueMember = "id"

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()


        TextBox1.DataBindings.Add(New Binding("text", BSTopVendor, "year", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("text", BSTopVendor, "orderline", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", BSTopVendor, "vendorid", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub
    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(TextBox1, "")


        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Role Name cannot be blank.")
            myret = False
        End If
        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from list")
            myret = False
        Else
            Dim drv = ComboBox1.SelectedItem
            Dim mydrv As DataRowView = Me.BSTopVendor.Current
            mydrv.Row.Item("shortname") = drv.row.item("shortname")
        End If
        If TextBox2.Text = "" Then
            ErrorProvider1.SetError(TextBox2, "Order Line No cannot be blank.")
            myret = False
        End If
        Return myret
    End Function

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        'Dim drv = ComboBox1.SelectedItem
        'Dim mydrv As DataRowView = Me.BSTopVendor.Current
        'mydrv.Row.Item("shortname") = drv.row.item("shortname")
        'mydrv.EndEdit()

    End Sub
End Class
