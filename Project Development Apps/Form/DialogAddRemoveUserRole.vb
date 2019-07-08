Imports System.Windows.Forms


Public Class DialogAddRemoveUserRole
    Public Property bsrole As New BindingSource
    Public Shared Event EndTx(ByVal o As Object, ByVal e As DialogAddRemoveUserRoleException)
    Public Property bsUser As BindingSource
    Public Property bsuserrole As BindingSource
    Public Property bsuserroleTX As BindingSource
    Public drv As DataRowView
    Public userid As Integer
    Private myform As DialogUserRole
    Public Sub New(bsuser As BindingSource, bsRole As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        ComboBox1.DataSource = bsRole
        ComboBox1.DisplayMember = "rolename"
        ComboBox1.ValueMember = "id"
        Me.bsUser = bsuser
        userid = CType(bsuser.Current, DataRowView).Item("id")

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(bsuser As BindingSource, bsRole As BindingSource, myform As DialogUserRole)

        ' This call is required by the designer.
        InitializeComponent()
        ComboBox1.DataSource = bsRole
        ComboBox1.DisplayMember = "rolename"
        ComboBox1.ValueMember = "id"
        Me.myform = myform
        Me.bsUser = bsuser
        userid = CType(bsuser.Current, DataRowView).Item("id")
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(bsRole As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        ComboBox1.DataSource = bsRole
        ComboBox1.DisplayMember = "rolename"
        ComboBox1.ValueMember = "id"

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        If Me.validate Then
            drv = ComboBox1.SelectedItem
            Me.DialogResult = System.Windows.Forms.DialogResult.OK

            Dim mye = New DialogAddRemoveUserRoleException(userid)
            myform.endtx(Me, mye)
            Me.Close()
        Else

        End If


    End Sub



    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DialogAddRemoveUserRole_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub
    Public Overloads Function validate() As Boolean
        Dim drv As DataRowView = ComboBox1.SelectedItem
        Dim myret = False

        If Not IsNothing(drv) Then
            ErrorProvider1.SetError(ComboBox1, "")
            myret = True
        Else
            ErrorProvider1.SetError(ComboBox1, "Please select from list.")
        End If
        Return myret
    End Function


End Class
Public Class DialogAddRemoveUserRoleException
    Inherits EventArgs
    Public Property userid As Integer
    Public Sub New(userid As Integer)
        Me.userid = userid
    End Sub
End Class