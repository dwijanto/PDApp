Public Class UserInfo
    Private Const VBLastErrorNumber As Integer = 513
    Public Property Userid As String = "N/A"
    Public Property Username As String = Environment.UserDomainName & "\" & Environment.UserName
    Public Property email As String
    Public Property Domain As String
    Public Property DisplayName As String
    Public Property Department As String
    Public Property computerName As String = My.Computer.Name
    Public Property ApplicationName As String = "PD Team Apps"
    Public Property isAuthenticate As Boolean = False
    Public Property isAdmin As Boolean = False
    Public Property Role As Integer

    Private Shared myInstance As UserInfo

    Public Sub New()

    End Sub

    Public Shared Function getInstance() As UserInfo
        If myInstance Is Nothing Then
            myInstance = New UserInfo
        End If
        Return myInstance
    End Function
End Class
