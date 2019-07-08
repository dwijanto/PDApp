Public Class OpenUserAdapter
    Inherits BaseAdapter
    Implements IAdapter


    Private Shared myInstance As OpenUserAdapter
    Public Property mydate As Date = Today.Date

    Public Shared Function getInstance()
        If myInstance Is Nothing Then
            myInstance = New OpenUserAdapter
        End If
        Return myInstance
    End Function

    Public Sub New()
        MyBase.New()
    End Sub

    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        DS = New DataSet
        BS = New BindingSource

        SB.Clear()
        SB.Append(String.Format("with u as ((select id from pd.user where isactive and title > 2" &
                  " order by username) except" &
                  " select userid from pd.taskcompletionconf where to_char(creationdate,'Mon YYYY') = to_char('{0}'::date,'Mon YYYY'))" &
                  " select us.username,pd.gettitle(title) as titlename,us.email from u left join pd.user us on us.id = u.id order by username;", mydate.Date))

        DbAdapter1.getDataSet(SB.ToString, DS)

        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("id")

        BS.DataSource = DS.Tables(0)

        Return True
    End Function

    Public Function Save() As Boolean Implements IAdapter.Save       
        Return False
    End Function
End Class
