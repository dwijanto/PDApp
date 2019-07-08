Imports System.Windows.Forms

Public Class DialogTopVendorInputV2

    Private BSTopVendor As BindingSource
    Private VendorBS As BindingSource
    Private DRV As DataRowView
    Public Shared Event EndTransaction(ByRef sender As Object, ByVal e As EventArgs)

    Public Sub New(drv As DataRowView, vendorbs As BindingSource)
        InitializeComponent()
        Me.DRV = drv

        Me.VendorBS = New BindingSource
        Dim mytbl As New DataTable
        mytbl = TryCast(vendorbs.DataSource, DataTable).Copy
        Me.VendorBS.DataSource = mytbl


        'Me.VendorBS = vendorbs

        InitDataDRV()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not Me.validate Then
            Exit Sub
        End If
        Try
            DRV.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            RaiseEvent EndTransaction(sender, e)
        Catch ex As ConstraintException
            MessageBox.Show("Duplicate data found. Please check.")
            DRV.CancelEdit()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            DRV.CancelEdit()
            Exit Sub
        End Try
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        DRV.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        RaiseEvent EndTransaction(sender, e)
        Me.Close()
    End Sub


    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        ErrorProvider1.SetError(TextBox1, "")
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Year cannot be blank.")
            myret = False
        End If
        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from list")
            myret = False
        Else
            Dim cbdrv = ComboBox1.SelectedItem
            DRV.Row.Item("shortname") = cbdrv.row.item("shortname")
        End If
        If TextBox2.Text = "" Then
            ErrorProvider1.SetError(TextBox2, "Order Line No cannot be blank.")
            myret = False
        End If
        Return myret
    End Function

    Private Sub InitDataDRV()
        ComboBox1.DataSource = VendorBS
        ComboBox1.DisplayMember = "shortname"
        ComboBox1.ValueMember = "id"

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()

        TextBox1.DataBindings.Add(New Binding("text", DRV, "year", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("text", DRV, "orderline", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", DRV, "vendorid", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub

    Private Sub DialogTopVendorInputV2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
