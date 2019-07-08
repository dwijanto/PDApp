Imports System.Text

Public Class ProjectAdapter
    Inherits BaseAdapter
    Implements IAdapter

    Public Property SBUBS As BindingSource
    Public Property VendorBS As BindingSource
    Public Property ProjectStatusBS As BindingSource
    Public Property RoleBS As BindingSource
    Public Property PICBS As BindingSource
    Public Property SubSBUBS As BindingSource

    Private Shared myInstance As ProjectAdapter
    Private myUserInstance As UserInfo

    Private ProjectCriteria As String
    Private RoleCriteria As String
    Private UserCriteria As String

    Public Shared Function getInstance()
        If myInstance Is Nothing Then
            myInstance = New ProjectAdapter
        End If
        Return myInstance
    End Function

    Public Sub New()
        MyBase.New()
        myUserInstance = UserInfo.getInstance
    End Sub

    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        SB.Clear()

        If Not myUserInstance.isAdmin Then
            ProjectCriteria = String.Format("left join pd.user_role ur on ur.roleid = p.role left join pd.user u on u.id = ur.userid   where lower(u.userid) = '{0}'", myUserInstance.Userid.ToLower)
            RoleCriteria = String.Format("left join pd.user_role ur on ur.roleid = r.id left join pd.user u on u.id = ur.userid where lower(u.userid) = '{0}'", myUserInstance.Userid.ToLower)
            UserCriteria = String.Format("and roleid in (select roleid from pd.user_role ur" &
                                         " left join pd.user u on u.id = ur.userid" &
                                         " where lower(u.userid) = '{0}')", myUserInstance.Userid.ToLower)
        Else
            RoleCriteria = " where r.id <> 10"
        End If

        DS = New DataSet
        BS = New BindingSource
        SBUBS = New BindingSource
        VendorBS = New BindingSource
        ProjectStatusBS = New BindingSource
        RoleBS = New BindingSource
        PICBS = New BindingSource
        SubSBUBS = New BindingSource


        SB.Clear()
        SB.Append(String.Format("select p.*,v.shortname,pic.username,ps.statusname,r.rolename from pd.project p left join pd.vendor v on v.id = p.vendorid  left join pd.user pic on pic.id = p.pic left join pd.role r on r.id = p.role left join pd.projectstatus ps on ps.id = p.project_status {0};", ProjectCriteria))
        SB.Append("select * from pd.sbu order by sbuname;")
        SB.Append("select * from pd.vendor order by shortname;")
        SB.Append("select * from pd.projectstatus order by id;")
        SB.Append(String.Format("select * from  pd.role r  {0}  order by rolename;", RoleCriteria))
        SB.Append(String.Format("select distinct u.userid,u.username,u.id from  pd.role r" &
                  " left join pd.user_role ur on ur.roleid = r.id " &
                  " left join pd.user u on u.id = ur.userid" &
                  " where title <> 0 and isactive {0}" &
                  " order by username;", UserCriteria))
        SB.Append("select * from pd.subsbu order by subsbuname;")

        DbAdapter1.getDataSet(SB.ToString, DS)

        'Set Primary Key
        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("id")
        DS.Tables(0).PrimaryKey = pk
        DS.Tables(0).Columns("id").AutoIncrement = True
        DS.Tables(0).Columns("id").AutoIncrementSeed = -1
        DS.Tables(0).Columns("id").AutoIncrementStep = -1

        BS.DataSource = DS.Tables(0)
        SBUBS.DataSource = DS.Tables(1)
        VendorBS.DataSource = DS.Tables(2)
        ProjectStatusBS.DataSource = DS.Tables(3)
        RoleBS.DataSource = DS.Tables(4)
        PICBS.DataSource = DS.Tables(5)
        SubSBUBS.DataSource = DS.Tables(6)

        Return True
    End Function

    Public Function Save() As Boolean Implements IAdapter.Save
        Dim myret As Boolean = False
        myInstance.BS.EndEdit()
        myInstance.RoleBS.EndEdit()
        myInstance.VendorBS.EndEdit()
        Dim ds2 As DataSet = myInstance.DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If DbAdapter1.ProjectTx(Me, mye) Then
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
