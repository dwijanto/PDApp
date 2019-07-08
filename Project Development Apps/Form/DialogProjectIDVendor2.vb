Imports System.Windows.Forms

Public Class DialogProjectIDVendor2

    Public _ProjectDRV As DataRowView
    Public myadapter As ProjectSignedAdapter2


    Public Sub New(projectDRV As DataRowView, myadapter As ProjectSignedAdapter2)

        ' This call is required by the designer.
        InitializeComponent()
        Me._ProjectDRV = projectDRV
        Me.myadapter = myadapter
        ' Add any initialization after the InitializeComponent() call.
        initData()
    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not Me.Validate Then
            Exit Sub
        End If
        Try
            _ProjectDRV.EndEdit() 'Me.myadapter.ProjectBS.EndEdit()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK

            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            _ProjectDRV.CancelEdit()
        End Try

    End Sub

    Public Overloads Function validate()
        Dim myret As Boolean = True
        ErrorProvider1.SetError(MaskedTextBox1, "")
        ErrorProvider1.SetError(ComboBox1, "")
        If MaskedTextBox1.Text = "00-0000-00" Then
            ErrorProvider1.SetError(MaskedTextBox1, "ProjectId cannot be blank.")
            myret = False
        End If
        If IsNothing(ComboBox1.SelectedItem) Then
            ErrorProvider1.SetError(ComboBox1, "Please select from list.")
            myret = False
        Else
            _ProjectDRV.Item("shortname") = DirectCast(ComboBox1.SelectedItem, DataRowView).Row.Item("shortname")
            _ProjectDRV.Item("vendorid") = DirectCast(ComboBox1.SelectedItem, DataRowView).Row.Item("id")
        End If
        Return myret

    End Function

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub initData()
        Dim mytbl As New DataTable
        mytbl = CType(myadapter.VendorBS.DataSource, DataTable).Copy
        Dim VendorBS As New BindingSource
        VendorBS.DataSource = mytbl

        ComboBox1.DataSource = VendorBS 'Me.myadapter.VendorBS
        ComboBox1.DisplayMember = "shortname"
        ComboBox1.ValueMember = "id"
        MaskedTextBox1.DataBindings.Add(New Binding("text", _ProjectDRV, "projectid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("SelectedValue", _ProjectDRV, "vendorid", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub

    Private Sub DialogProjectIDVendor_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not Me.DialogResult = System.Windows.Forms.DialogResult.OK Then
            _ProjectDRV.CancelEdit()
        End If
    End Sub
End Class
