'
Imports System.Text
Public Class BaseAdapter
    Public Property DbAdapter1 As DbAdapter
    Public Property DS As DataSet
    Public Property BS As BindingSource
    Public Property SB As StringBuilder

    Public Sub New()
        DbAdapter1 = DbAdapter.getInstance 
        SB = New StringBuilder
    End Sub
End Class
