Public Class UserAdapter
    Inherits BaseAdapter
    Implements IAdapter

    Public Property RoleBS As BindingSource
    Public Property UserRoleBS As BindingSource
    Private Shared myInstance As UserAdapter


    Public Shared Function getInstance()
        If myInstance Is Nothing Then
            myInstance = New UserAdapter            
        End If
        Return myInstance
    End Function

    Public Sub New()
        MyBase.New()

    End Sub
    Public Function getData(ByVal id As Integer) As BindingSource
        DS = New DataSet
        BS = New BindingSource
        RoleBS = New BindingSource
        SB.Clear()
        SB.Append(String.Format("select u.* from pd.user u where id = {0} order by u.userid;", id))
        DbAdapter1.getDataSet(SB.ToString, DS)
        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("userid")
        DS.Tables(0).PrimaryKey = pk
        BS.DataSource = DS.Tables(0)
        If IsNothing(BS.Current) Then
            BS.AddNew()
        End If
        Return BS
    End Function
    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        DS = New DataSet
        BS = New BindingSource
        RoleBS = New BindingSource
        UserRoleBS = New BindingSource
        SB.Clear()
        SB.Append("select u.* from pd.user u order by u.username;")
        SB.Append("select * from  pd.role r order by rolename;")
        SB.Append("select ur.*,r.rolename from  pd.user_role ur left join pd.role r on r.id = ur.roleid;")
        DbAdapter1.GetDataset(SB.ToString, DS)
        'Set Primary Key
        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("id")
        DS.Tables(0).PrimaryKey = pk
        DS.Tables(0).Columns("id").AutoIncrement = True
        DS.Tables(0).Columns("id").AutoIncrementSeed = -1
        DS.Tables(0).Columns("id").AutoIncrementStep = -1


        Dim pk1(0) As DataColumn
        pk1(0) = DS.Tables(1).Columns("id")
        DS.Tables(1).PrimaryKey = pk1
        DS.Tables(1).Columns("id").AutoIncrement = True
        DS.Tables(1).Columns("id").AutoIncrementSeed = -1
        DS.Tables(1).Columns("id").AutoIncrementStep = -1

        Dim pk2(0) As DataColumn
        pk2(0) = DS.Tables(2).Columns("id")
        DS.Tables(2).PrimaryKey = pk2
        DS.Tables(2).Columns("id").AutoIncrement = True
        DS.Tables(2).Columns("id").AutoIncrementSeed = -1
        DS.Tables(2).Columns("id").AutoIncrementStep = -1


        'Unique Constrain
        Dim U0(0) As DataColumn
        U0(0) = DS.Tables(0).Columns("userid")
        Dim UUser As UniqueConstraint = New UniqueConstraint(U0)
        DS.Tables(0).Constraints.Add(UUser)


        'Unique Constrain
        Dim U1(0) As DataColumn
        U1(0) = DS.Tables(1).Columns("rolename")
        Dim URole As UniqueConstraint = New UniqueConstraint(U1)
        DS.Tables(1).Constraints.Add(URole)

        'Unique Constrain
        Dim U2(1) As DataColumn
        U2(0) = DS.Tables(2).Columns("userid")
        U2(1) = DS.Tables(2).Columns("roleid")
        Dim userRole As UniqueConstraint = New UniqueConstraint(U2)
        DS.Tables(2).Constraints.Add(userRole)


        BS.DataSource = DS.Tables(0)
        RoleBS.DataSource = DS.Tables(1)
        UserRoleBS.DataSource = DS.Tables(2)
        Return True
    End Function

    Public Function Save() As Boolean Implements IAdapter.Save
        Dim myret As Boolean = False
        myInstance.BS.EndEdit()
        myInstance.RoleBS.EndEdit()
        myInstance.UserRoleBS.EndEdit()
        Dim ds2 As DataSet = myInstance.DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If DbAdapter1.UserTx(Me, mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If
        Return myret
    End Function

End Class

Public Class UserEventArgs
    Inherits EventArgs
    Public StatusTx As TxRecord

    Public Sub New(ByVal status As TxRecord)
        StatusTx = status
    End Sub
End Class