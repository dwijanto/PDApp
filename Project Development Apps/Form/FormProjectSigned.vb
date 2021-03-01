Imports System.Threading
Imports System.Text
Imports Microsoft.Office.Interop

'Rule:
' User can delete at the same period (postingdate month = current month)
' Admin can delete any

Public Class FormProjectSigned
    Public Shared myForm As FormProjectSigned
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    'Dim myAdapter As ProjectSignedAdapter
    Dim myAdapter As ProjectSignedAdapter2
    Private myuserinfo As UserInfo

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormProjectSigned
            'AddHandler DialogProjectSignedInput.SaveRecord, AddressOf myForm.AssignRecord
            AddHandler DialogProjectSignedInput2.SaveRecord, AddressOf myForm.AssignRecord
        ElseIf myForm.IsDisposed Then
            myForm = New FormProjectSigned
        End If
        Return myForm
    End Function


    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        ShowTx(TxRecord.AddRecord)
        'Dim myform = New DialogProjectSignedInput(myAdapter, TxRecord.AddRecord)
        'myform.Show()
    End Sub
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex <> -1 Then
            If Not myuserinfo.isAdmin Then
                Dim drv As DataRowView = myAdapter.BS.Current
                If String.Format("{0:MMM-yyyy}", drv.Item("postingdate")) <> String.Format("{0:MMM-yyyy}", Date.Today) Then
                    MessageBox.Show("You cannot modify this record. Please contact Admin.")
                    Exit Sub
                End If
            End If
            ShowTx(TxRecord.UpdateRecord)
            'Dim myform = New DialogProjectSignedInput(myAdapter, TxRecord.UpdateRecord)
            'myform.Show()
        End If
    End Sub

    Private Sub ShowTx(txRecord As TxRecord)
        Dim ProjectDRV As DataRowView = Nothing
        Dim drv As DataRowView = Nothing
        Select Case txRecord
            Case txRecord.UpdateRecord
                drv = myAdapter.BS.Current
            Case txRecord.AddRecord
                drv = myAdapter.BS.AddNew
                drv.Row.Item("postingdate") = Date.Today
                drv.Row.Item("signed_date") = Date.Today
                drv.Row.Item("stamp") = Date.Now
                drv.Row.Item("modifiedby") = myuserinfo.Userid
                drv.Row.Item("docreceived") = False
                drv.Row.Item("pdoc") = False
        End Select
        'Dim myform = New DialogProjectSignedInput(drv, ProjectDRV, myAdapter)
        Dim myform = New DialogProjectSignedInput2(drv, ProjectDRV, myAdapter)
        myform.ShowDialog()
    End Sub

    Private Sub FormProjectSigned_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        'MessageBox.Show("Activate")
        DataGridView1.Invalidate()
    End Sub


    Private Sub FormProjectSigned_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub


    Private Sub FormUser_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub FormProjectSigned_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not ValidateRP4() Then
            e.Cancel = True
        End If
        If myAdapter.DS.HasErrors Then
            MessageBox.Show("Failed to save to database. Please check!")
            e.Cancel = True
        End If
    End Sub

    Private Function ValidateRP4() As Boolean
        Dim myresult As Boolean = True
        'Dim myAdapter = New ProjectSignedAdapter
        Dim myAdapter = New ProjectSignedAdapter2
        Dim result As String = String.Empty
        If myAdapter.ValidateRP4(result) Then
            If result.Length > 0 Then
                MessageBox.Show(String.Format("Missing one of RP4 Ontime or RP4 Price. Please check your transaction. {0}", vbCrLf & vbCrLf & result))
                myresult = False
            End If
        End If
        'Throw New NotImplementedException
        Return myresult
    End Function

    Private Sub FormUser_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        'myAdapter = ProjectSignedAdapter.getInstance
        myAdapter = ProjectSignedAdapter2.getInstance
        Try
            If myAdapter.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    'Init Combobox
                    'CProjectId.DataSource = myAdapter.ProjectBS
                    'CProjectId.DisplayMember = "projectid"
                    'CProjectId.ValueMember = "id"
                    'CProjectId.DataPropertyName = "projectid"

                    'CProjectName.DataSource = myAdapter.ProjectBS
                    'CProjectName.DisplayMember = "projectname"
                    'CProjectName.ValueMember = "id"
                    'CProjectName.DataPropertyName = "projectid"

                    'CShortName.DataSource = myAdapter.ProjectBS
                    'CShortName.DisplayMember = "shortname"
                    'CShortName.ValueMember = "id"
                    'CShortName.DataPropertyName = "projectid"

                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myAdapter.BS

                    DataGridView1.Columns("Column1").ReadOnly = Not (myuserinfo.isAdmin) 'must reverse
            End Select
        End If
    End Sub







    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        loaddata()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Me.Validate()
        myAdapter.Save()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(myAdapter.BS.Current) Then
            If Not myuserinfo.isAdmin Then
                Dim drv As DataRowView = myAdapter.BS.Current
                If String.Format("{0:MMM-yyyy}", drv.Item("postingdate")) <> String.Format("{0:MMM-yyyy}", Date.Today) Then
                    MessageBox.Show("You cannot delete this record. Please contact Admin.")
                    Exit Sub
                Else
                    If myuserinfo.Userid.ToLower <> drv.Item("modifiedby").ToString.ToLower Then
                        MessageBox.Show("This record is not belong to you, cannot be deleted.")
                        Exit Sub
                    End If
                End If
            End If
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    'myAdapter.BS.RemoveAt(drv.Index)
                    myAdapter.BS.Position = drv.Index
                    Dim currdrv As DataRowView = myAdapter.BS.Current
                    Dim sameuser As Boolean = currdrv.Item("modifiedby").ToString.ToLower = myuserinfo.Userid.ToLower
                    Dim usercandelete As Boolean = (String.Format("{0:MMM-yyyy}", currdrv.Item("postingdate")) = String.Format("{0:MMM-yyyy}", Date.Today))
                    If myuserinfo.isAdmin Or (usercandelete And sameuser) Then
                        myAdapter.BS.RemoveAt(drv.Index)
                    Else
                        currdrv.Row.RowError = "This record does not belong to you.cannot be deleted"
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim obj As ToolStripTextBox = DirectCast(sender, ToolStripTextBox)
        Dim sb As New StringBuilder

        If obj.Text = "" Then
            myAdapter.BS.Filter = ""
        Else
            sb.Append(String.Format("pidname like '*{0}*' or projectname like '*{0}*' or shortname like '*{0}*' or projectstage like '*{0}*' or phase_status_name like '*{0}*' or pic like '*{0}*'", obj.Text))
            myAdapter.BS.Filter = sb.ToString

        End If
    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        myuserinfo = UserInfo.getInstance


    End Sub

    Private Sub AssignRecord(ByRef sender As Object, ByVal e As UserEventArgs)
        Dim drv As DataRowView = DirectCast(sender, DataRowView)
        Dim mydrv As DataRowView = Nothing
        Select Case e.StatusTx
            Case TxRecord.AddRecord
                mydrv = myAdapter.BS.AddNew()
            Case TxRecord.UpdateRecord
                Dim result = myAdapter.BS.Find("id", drv.Item("id"))
                mydrv = myAdapter.BS.Item(result)
        End Select
        For i = 1 To mydrv.Row.ItemArray.Length - 1
            mydrv.Item(i) = drv.Item(i)
        Next
        'mydrv.Item("userid") = drv.Item("userid")
        'mydrv.Item("username") = drv.Item("username")
        'mydrv.Item("email") = drv.Item("email")
        'mydrv.Item("isactive") = drv.Item("isactive")
        'mydrv.Item("isadmin") = drv.Item("isadmin")
        'mydrv.Item("title") = drv.Item("title")
        mydrv.EndEdit()
        DataGridView1.Invalidate()
    End Sub




    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Dim GenerateReport As New ClassGenReport(Me, String.Format("ReportTx-{0:yyyy-MM-dd}.xlsx", Date.Today))
        Dim mylist As New List(Of QueryWorksheet)
        Dim q1 As New QueryWorksheet
        Dim criteria As String = String.Empty
        Dim classReportKPI1 As New ClassReportKPI

        If Not myuserinfo.isAdmin Then
            criteria = " where p.role in (select * from ur)"
            'criteria = String.Format("left join pd.user_role ur on ur.roleid = p.role left join pd.user u on u.id = ur.userid   where lower(u.userid) = '{0}' ", myuserinfo.Userid.ToLower)
        End If
        'q1.Sqlstr = String.Format("with ur as(select ur.roleid from pd.user_role ur left join pd.user u on u.id = ur.userid   where lower(u.userid) = '{1}'), " &
        '                          " ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
        '            " from pd.projecttx" &
        '            " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
        '            " txadjust as( select tx.projectid,tx.project_phase_id,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
        '            " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
        '            " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
        '            " group by projectid" &
        '            " having count(projectid) = 2)," &
        '            " lognumber as (select * from crosstab('select projectid,project_phase_id,lognumber from pd.projecttx where not lognumber isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,failreportlogpsarra text,failreportlogrp4price text,failreportlogrp4ontime text,failreportlogrp5closure text,failreportlogrampup text))" &
        '            " select r.rolename,u.username, s.sbu_shortname,p.projectid, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rp5closure," &
        '            " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time""," &
        '            " case when pc.status = 2 then case when field5 > field7 then field5 else field7 end end as datelaunchesontime," &
        '            " statusname as projectstatus,p.countstartdate as startdate,p.lognumber as lognumberpsarra," &
        '            " failreportlogpsarra,failreportlogrp4price,failreportlogrp4ontime,failreportlogrp5closure,failreportlogrampup" &
        '            " , case when txpsarra.status isnull then null	" &
        '            " when txpsarra.status = 1 then 'pass' else  'failed'" &
        '            " end as psarrastatus," &
        '            " case when txprice.status isnull then null	" &
        '            " when txprice.status = 1 then 'pass' else  'failed'" &
        '            " end as pricestatus," &
        '            " case when txontime.status isnull then null" &
        '            " when txontime.status = 1 then 'pass' else  'failed'" &
        '            " end as ontimestatus," &
        '            " case when txrampup.status isnull then null" &
        '            " when txrampup.status = 1 then 'pass' else  'failed'" &
        '            " end as rampupstatus" &
        '            ", case p.variant when true then 'yes' else '' end as variant" &
        '            " from ct" &
        '            " left join pd.project p on p.id = ct.id" &
        '            " left join pd.sbu s on s.id = p.sbu_id" &
        '            " left join pd.role r on r.id = p.role" &
        '            " left join pd.user u on u.id = p.pic" &
        '            " left join pd.vendor v on v.id = p.vendorid" &
        '            " left join pd.projecttype pt on pt.id = p.project_type_id" &
        '            " left join pc on pc.projectid = ct.id" &
        '            " left join pd.projectstatus ps on ps.id = p.project_status" &
        '            " left join lognumber l on l.id = ct.id" &
        '            " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
        '            " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
        '            " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
        '            " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
        '            " {2} " &
        '            " order by p.projectname", String.Format("{0:yyyyMM}", Today.Date), myuserinfo.Userid.ToLower, criteria)
        'q1.Sqlstr = String.Format("with ur as(select ur.roleid from pd.user_role ur left join pd.user u on u.id = ur.userid   where lower(u.userid) = '{1}'), " &
        '                  " ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
        '    " from pd.projecttx" &
        '    " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
        '    " txadjust as( select tx.projectid,lognum,docreceived,tx.project_phase_id,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
        '    " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
        '    " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
        '    " group by projectid" &
        '    " having count(projectid) = 2)," &
        '    " lognumber as (select * from crosstab('select projectid,project_phase_id,pd.getlognumber(signed_date,project_phase_id,lognum) as lognumber from pd.projecttx where not lognum isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,failreportlogpsarra text,failreportlogrp4price text,failreportlogrp4ontime text,failreportlogrp5closure text,failreportlogrampup text))" &
        '    " select r.rolename,u.username, s.sbu_shortname,p.projectid, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rp5closure," &
        '    " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time""," &
        '    " case when pc.status = 2 then case when field5 > field7 then field5 else field7 end end as datelaunchesontime," &
        '    " statusname as projectstatus,p.countstartdate as startdate,pd.getlognumber(ptx.signed_date,0,p.lognum) as lognumberpsarra," &
        '    " failreportlogpsarra,failreportlogrp4price,failreportlogrp4ontime,failreportlogrp5closure,failreportlogrampup," &
        '    " case when txpsarra.status isnull then null	when txpsarra.status = 1 then 'pass' else  'failed' end as psarrastatus," &
        '    " case when txprice.status isnull then null	when txprice.status = 1 then 'pass' else  'failed' end as pricestatus," &
        '    " case when txontime.status isnull then null when txontime.status = 1 then 'pass' else  'failed' end as ontimestatus," &
        '    " case when txrampup.status isnull then null when txrampup.status = 1 then 'pass' else  'failed' end as rampupstatus" &
        '    ", case p.variant when true then 'yes' else '' end as variant," &
        '    " case when p.lognum isnull then null when p.docreceived then 'received' else 'not received' end  as pdoc," &
        '    " case when txpsarra.lognum isnull then null when txpsarra.docreceived then 'received' else 'not received' end as psarradocf," &
        '    " case when txprice.lognum isnull then null when txprice.docreceived then 'received' else 'not received' end  as pricedocf," &
        '    " case when txontime.lognum isnull then null when  txontime.docreceived then 'received' else 'not received' end  as ontimedocf," &
        '    " case when txrampup.lognum isnull then null when txrampup.docreceived then 'received' else 'not received' end as rampupdocf," &
        '    " case when txclosure.lognum isnull then null when txclosure.docreceived then 'received' else 'not received' end  as closuredocf" &
        '    " from ct" &
        '    " left join pd.project p on p.id = ct.id" &
        '    " left join pd.projecttx ptx on ptx.projectid = ct.id and project_phase_id = 3" &
        '    " left join pd.sbu s on s.id = p.sbu_id" &
        '    " left join pd.role r on r.id = p.role" &
        '    " left join pd.user u on u.id = p.pic" &
        '    " left join pd.vendor v on v.id = p.vendorid" &
        '    " left join pd.projecttype pt on pt.id = p.project_type_id" &
        '    " left join pc on pc.projectid = ct.id" &
        '    " left join pd.projectstatus ps on ps.id = p.project_status" &
        '    " left join lognumber l on l.id = ct.id" &
        '    " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
        '    " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
        '    " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
        '    " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
        '    " left join txadjust txclosure on txclosure.projectid = ct.id and txclosure.project_phase_id = 6 " &
        '    " {2} " &
        '    " order by p.projectname", String.Format("{0:yyyyMM}", Today.Date), myuserinfo.Userid.ToLower, criteria)
        'q1.Sqlstr = String.Format("with ur as(select ur.roleid from pd.user_role ur left join pd.user u on u.id = ur.userid   where lower(u.userid) = '{1}'), " &
        '                 " ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
        '   " from pd.projecttx" &
        '   " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
        '   " txadjust as( select tx.projectid,lognum,docreceived,tx.project_phase_id,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
        '   " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
        '   " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
        '   " group by projectid" &
        '   " having count(projectid) = 2)," &
        '   " lognumber as (select * from crosstab('select projectid,project_phase_id,pd.getlognumber(yearlognum,project_phase_id,lognum) as lognumber from pd.projecttx where not lognum isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,failreportlogpsarra text,failreportlogrp4price text,failreportlogrp4ontime text,failreportlogrp5closure text,failreportlogrampup text))" &
        '   " select r.rolename,u.username, s.sbu_shortname,p.projectid, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rp5closure," &
        '   " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time""," &
        '   " case when pc.status = 2 then case when field5 > field7 then field5 else field7 end end as datelaunchesontime," &
        '   " statusname as projectstatus,p.countstartdate as startdate,pd.getlognumber(p.yearlognum,0,p.lognum) as lognumberpsarra," &
        '   " failreportlogpsarra,failreportlogrp4price,failreportlogrp4ontime,failreportlogrampup,failreportlogrp5closure," &
        '   " case when txpsarra.status isnull then null	when txpsarra.status = 1 then 'pass' else  'failed' end as psarrastatus," &
        '   " case when txprice.status isnull then null	when txprice.status = 1 then 'pass' else  'failed' end as pricestatus," &
        '   " case when txontime.status isnull then null when txontime.status = 1 then 'pass' else  'failed' end as ontimestatus," &
        '   " case when txrampup.status isnull then null when txrampup.status = 1 then 'pass' else  'failed' end as rampupstatus," &
        '   " case when txclosure.status isnull then null when txclosure.status = 1 then 'pass' else  'failed' end as closurestatus," &
        '   " case p.variant when true then 'yes' else '' end as variant," &
        '   " case when p.lognum isnull then null when p.docreceived then 'received' else 'not received' end  as pdoc," &
        '   " case when txpsarra.lognum isnull then null when txpsarra.docreceived then 'received' else 'not received' end as psarradocf," &
        '   " case when txprice.lognum isnull then null when txprice.docreceived then 'received' else 'not received' end  as pricedocf," &
        '   " case when txontime.lognum isnull then null when  txontime.docreceived then 'received' else 'not received' end  as ontimedocf," &
        '   " case when txrampup.lognum isnull then null when txrampup.docreceived then 'received' else 'not received' end as rampupdocf," &
        '   " case when txclosure.lognum isnull then null when txclosure.docreceived then 'received' else 'not received' end  as closuredocf" &
        '   " from ct" &
        '   " left join pd.project p on p.id = ct.id" &
        '   " left join pd.projecttx ptx on ptx.projectid = ct.id and project_phase_id = 3" &
        '   " left join pd.sbu s on s.id = p.sbu_id" &
        '   " left join pd.role r on r.id = p.role" &
        '   " left join pd.user u on u.id = p.pic" &
        '   " left join pd.vendor v on v.id = p.vendorid" &
        '   " left join pd.projecttype pt on pt.id = p.project_type_id" &
        '   " left join pc on pc.projectid = ct.id" &
        '   " left join pd.projectstatus ps on ps.id = p.project_status" &
        '   " left join lognumber l on l.id = ct.id" &
        '   " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
        '   " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
        '   " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
        '   " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
        '   " left join txadjust txclosure on txclosure.projectid = ct.id and txclosure.project_phase_id = 6 " &
        '   " {2} " &
        '   " order by p.projectname", String.Format("{0:yyyyMM}", Today.Date), myuserinfo.Userid.ToLower, criteria)
        q1.Sqlstr = classReportKPI1.getSignedProject(myuserinfo.Userid.ToLower, criteria)
        q1.DataSheet = 1
        q1.SheetName = "Transaction"
        mylist.Add(q1)
        GenerateReport.LoadReport(mylist)

    End Sub

    Public Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim mye = DirectCast(e, FormatReportEventargs)
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent

        Select Case mye.sheetno            
            Case 1
                'owb.Worksheets(7).select()
                Dim osheet As Excel.Worksheet = owb.Worksheets(1)
                osheet.Columns("H:N").numberformat = "MMM-yyyy"
                osheet.Columns("P:P").numberformat = "MMM-yyyy"
                osheet.Rows("1:1").AutoFilter()
                osheet.Cells.EntireColumn.AutoFit()

                Dim check As String
                Dim i As Integer = 2
                Do While True
                    If IsNothing(osheet.Cells(i, 1).value) Then
                        Exit Do
                    End If
                    'Signed_date Failed
                    For j = 0 To 4
                        ' check = osheet.Cells(i, 22 + j).value
                        check = osheet.Cells(i, 25 + j).value
                        If Not IsNothing(check) Then
                            If check = "failed" Then
                                osheet.Cells(i, 10 + j).interior.color = 255
                            End If
                        End If
                    Next

                    'Document Received
                    For k = 0 To 5
                        ' check = osheet.Cells(i, 22 + j).value
                        check = osheet.Cells(i, 31 + k).value
                        If Not IsNothing(check) Then
                            If check = "not received" Then
                                osheet.Cells(i, 19 + k).interior.color = 65535
                            End If
                        End If
                    Next


                    i = i + 1
                Loop
                osheet.Columns("Y:AB").EntireColumn.Hidden = True
                osheet.Columns("AD:AI").EntireColumn.Hidden = True
                owb.Worksheets(1).select()
                osheet = owb.Worksheets(1)
        End Select
    End Sub
    Public Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)

    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub
End Class