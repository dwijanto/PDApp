Imports System.Text

Public Class ProjectSignedAdapter
    Inherits BaseAdapter
    Implements IAdapter

    Public Property ProjectBS As BindingSource
    Public Property PhaseBS As BindingSource
    Public Property VendorBS As BindingSource
    Private Shared myInstance As ProjectSignedAdapter

    Private myUserInstance As UserInfo

    Private ProjectCriteria As String
    Private RoleCriteria As String
    Private UserCriteria As String

    Public Shared Function getInstance()
        If myInstance Is Nothing Then
            myInstance = New ProjectSignedAdapter
        End If
        Return myInstance
    End Function

    Public Sub New()
        MyBase.New()
        myUserInstance = UserInfo.getInstance
    End Sub


    Public Function ValidateRP4(ByRef result As String) As Boolean
        Dim sb As New StringBuilder
        Dim myCriteria As String = String.Empty
        If Not myUserInstance.isAdmin Then
            myCriteria = String.Format(" and  lower(u.userid) = '{0}' and p.pic = ur.userid", myUserInstance.Userid.ToLower)
        Else
            Return True
        End If
        Dim sqlstr = String.Format("select tx.projectid,count(pp.project_phase_name_alias),p.projectname from pd.projecttx tx" &
                                   " left join pd.project_phase pp on pp.id = tx.project_phase_id" &
                                   " left join pd.project p on p.id = tx.projectid" &
                                   " left join pd.user_role ur on ur.roleid = p.role" &
                                   " left join pd.user u on u.id = ur.userid" &
                                   " where pp.project_phase_name_alias = 'RP4' {0}" &
                                   " group by tx.projectid ,projectname" &
                                   " having count(pp.project_phase_name_alias) <> 2 " &
                                   " order by p.projectname;", myCriteria)
        Dim ds As New DataSet
        If DbAdapter1.getDataSet(sqlstr, ds) Then
            For Each dr As DataRow In ds.Tables(0).Rows()
                If sb.Length > 0 Then
                    sb.Append(vbCrLf)
                End If

                sb.Append(dr.Item("projectname"))

            Next
            result = sb.ToString
        End If
        Return True
    End Function

    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        SB.Clear()


        If Not myUserInstance.isAdmin Then
            UserCriteria = String.Format(" where tx.modified = '{0}'", myUserInstance.Userid)
            ProjectCriteria = String.Format("left join pd.user_role ur on ur.roleid = p.role left join pd.user u on u.id = ur.userid   where lower(u.userid) = '{0}' ", myUserInstance.Userid.ToLower)
        End If

        DS = New DataSet
        BS = New BindingSource
        ProjectBS = New BindingSource
        PhaseBS = New BindingSource
        VendorBS = New BindingSource


        SB.Clear()
        SB.Append(String.Format("select tx.*,p.projectid as pidname,p.projectname,v.id as vendorid,v.shortname,pp.input_name as projectstage,case phase_status when 1 then 'Passed'  when 2 then 'Failed' end as phase_status_name,pic.username as pic from pd.projecttx tx left join pd.project p on p.id = tx.projectid left join pd.vendor v on v.id = p.vendorid left join pd.project_phase pp on pp.id = tx.project_phase_id left join pd.user pic on pic.id = p.pic {0};", ProjectCriteria))
        SB.Append(String.Format("select p.*,v.shortname,pic.username,ps.statusname,r.rolename from pd.project p left join pd.vendor v on v.id = p.vendorid  left join pd.user pic on pic.id = p.pic left join pd.role r on r.id = p.role left join pd.projectstatus ps on ps.id = p.project_status {0}   order by p.projectname;", ProjectCriteria))
        'SB.Append(String.Format("select p.*,v.shortname,pic.username,ps.statusname,r.rolename from pd.project p left join pd.vendor v on v.id = p.vendorid  left join pd.user pic on pic.id = p.pic left join pd.role r on r.id = p.role left join pd.projectstatus ps on ps.id = p.project_status {0}  statusname = 'Active'  order by p.projectname;", IIf(ProjectCriteria = "", " where ", ProjectCriteria & " and ")))
        SB.Append("select * from pd.vendor order by shortname;")
        SB.Append("select * from pd.project_phase order by order_line;")

        DbAdapter1.getDataSet(SB.ToString, DS)

        'Set Primary Key
        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("id")
        DS.Tables(0).PrimaryKey = pk
        DS.Tables(0).Columns("id").AutoIncrement = True
        DS.Tables(0).Columns("id").AutoIncrementSeed = -1
        DS.Tables(0).Columns("id").AutoIncrementStep = -1

        Dim uc(1) As DataColumn
        uc(0) = DS.Tables(0).Columns("projectid")
        uc(1) = DS.Tables(0).Columns("project_phase_id")
        Dim UniqueConstraint As UniqueConstraint = New UniqueConstraint("ProjectPhase", uc)
        DS.Tables(0).Constraints.Add(UniqueConstraint)

        BS.DataSource = DS.Tables(0)

        Dim pk1(0) As DataColumn
        pk1(0) = DS.Tables(1).Columns("id")
        DS.Tables(1).PrimaryKey = pk1

        ProjectBS.DataSource = DS.Tables(1)
        VendorBS.DataSource = DS.Tables(2)
        PhaseBS.DataSource = DS.Tables(3)

        Return True
    End Function

    Public Function Save() As Boolean Implements IAdapter.Save
        Dim myret As Boolean = False
        myInstance.BS.EndEdit()
        myInstance.ProjectBS.EndEdit()
        Dim ds2 As DataSet = myInstance.DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If DbAdapter1.ProjectSignedTx(Me, mye) Then
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
