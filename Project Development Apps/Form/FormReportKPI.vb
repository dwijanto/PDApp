Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.Text

Public Class FormReportKPI

    Private Shared myInstance As FormReportKPI
    Private SqlStrReport As String = String.Empty
    Dim myThread As New System.Threading.Thread(AddressOf DoReport)
    Dim dateSelected As Date
    Dim ClassReportKPI1 As New ClassReportKPI
    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    dateSelected = DateTimePicker1.Value.Date
    '    'SqlStrReport = String.Format("select * from pd.sp_getkpi('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) where postingdate >='{0:yyyy-MM-dd}'::date", CDate(DateTimePicker1.Value.Year.ToString + "-01-01"), lastDate(DateTimePicker1.Value.Date))
    '    SqlStrReport = String.Format("select * from pd.sp_getkpi('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date)" &
    '                                 " union all " &
    '                                 " select * from pd.sp_getrp14running('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) " &
    '                                 " union all " &
    '                                 " select * from pd.sp_getrp15running('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) ", CDate(DateTimePicker1.Value.Year.ToString + "-01-01"), lastDate(DateTimePicker1.Value.Date))
    '    loadReport()
    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        dateSelected = DateTimePicker1.Value.Date
        'SqlStrReport = String.Format("select * from pd.sp_getkpi('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) where postingdate >='{0:yyyy-MM-dd}'::date", CDate(DateTimePicker1.Value.Year.ToString + "-01-01"), lastDate(DateTimePicker1.Value.Date))
        'SqlStrReport = String.Format("select * from pd.sp_getkpi('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date)" &                                    
        '                             " union all " &
        '                             " select * from pd.sp_getrp15running('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) ", CDate(DateTimePicker1.Value.Year.ToString + "-01-01"), lastDate(DateTimePicker1.Value.Date))
        SqlStrReport = ClassReportKPI1.getRawData(CDate(DateTimePicker1.Value.Year.ToString + "-01-01"), lastDate(DateTimePicker1.Value.Date))
        loadReport()
    End Sub
    Private Function lastDate(mydate As Date) As Date
        Dim mymonth As Integer = mydate.Month
        Dim myyear As Integer = mydate.Year

        If mymonth = 12 Then
            myyear = myyear + 1
            mymonth = 1
        Else
            mymonth = mymonth + 1
        End If
        Return CDate(String.Format("{0}-{1}-1", myyear, mymonth)).AddDays(-1)
    End Function

    Public Shared Function getInstance()
        If myInstance Is Nothing Then
            myInstance = New FormReportKPI
        ElseIf myInstance.IsDisposed Then
            myInstance = New FormReportKPI
        End If
        Return myInstance
    End Function


    Private Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub FormReportKPI_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        myInstance = Nothing
        Me.Dispose()
    End Sub

    Private Sub loadReport()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""

            myThread = New Thread(AddressOf DoReport)
            myThread.SetApartmentState(ApartmentState.STA)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoReport()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Dim sqlstr As String = String.Empty
        Dim sb As New StringBuilder
        'DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        ProgressReport(11, "InitFilter")

        sqlstr = String.Format(SqlStrReport, sb.ToString.ToLower)

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("KPI-{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)
            Dim datasheet As Integer = 6

            Dim mylist As New List(Of QueryWorksheet)
            Dim q1 As New QueryWorksheet
            q1.DataSheet = 6
            q1.Sqlstr = sqlstr
            q1.SheetName = "RawData"
            mylist.Add(q1)

            Dim q2 As New QueryWorksheet
            q2.DataSheet = 7

            'q2.Sqlstr = String.Format("select * from pd.txhistory where period >= {0} and period <= {1}", String.Format("{0:yyyy01}", DateTimePicker1.Value.Date), String.Format("{0:yyyyMM}", DateTimePicker1.Value.Date))

            q2.Sqlstr = String.Format("select 'Total FP' as sbu_shortname,tx.* from pd.txhistory tx where period >= {0} and period <= {1} and sbu <> 'SA' " &
                                      " union all select 'Total KE',tx.* from pd.txhistory tx where period >= {0} and period <= {1} and  sbu in('EC','BnB','FP')" &
                                      " union all select 'Total HPC',tx.* from pd.txhistory tx where period >= {0} and period <= {1} and sbu in('PC','LC','H.Comf','H.Clean')" &
                                      " union all select 'Total CW',tx.* from pd.txhistory tx where period >= {0} and period <= {1} and sbu = 'CK'" &
                                      " union all select 'Total SA',tx.* from pd.txhistory tx where period >= {0} and period <= {1} and sbu = 'SA'" &
                                      " union all select 'Total OBH',tx.* from pd.txhistory tx where period >= {0} and period <= {1} and sbu = 'OBH'" &
                                      " union all select 'Total KE + HPC',tx.* from pd.txhistory tx where period >= {0} and period <= {1} and sbu in ('EC','BnB','FP','PC','LC','H.Comf','H.Clean')", String.Format("{0:yyyy01}", DateTimePicker1.Value.Date), String.Format("{0:yyyyMM}", DateTimePicker1.Value.Date))

            q2.SheetName = "TxHistory"

            mylist.Add(q2)


            Dim q3 As New QueryWorksheet
            q3.DataSheet = 8

            'q3.Sqlstr = "with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
            '            " from pd.projecttx" &
            '            " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
            '            " pc as (select projectid,sum(phase_status) as status from pd.projecttx where project_phase_id in (4,5,7)" &
            '            " group by projectid" &
            '            " having count(projectid) = 3)" &
            '            " select r.rolename,u.username, s.sbu_shortname, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rpclosure," &
            '            " case when status isnull then null when status = 3 then 'Pass' else 'Fail' end as ""RP4_RampUp"",statusname as projectstatus,p.countstartdate as startdate" &
            '            " from ct" &
            '            " left join pd.project p on p.id = ct.id" &
            '            " left join pd.sbu s on s.id = p.sbu_id" &
            '            " left join pd.role r on r.id = p.role" &
            '            " left join pd.user u on u.id = p.pic" &
            '            " left join pd.vendor v on v.id = p.vendorid" &
            '            " left join pd.projecttype pt on pt.id = p.project_type_id" &
            '            " left join pc on pc.projectid = ct.id" &
            '            " left join pd.projectstatus ps on ps.id = p.project_status" &
            '            " order by p.projectname"
            'q3.Sqlstr = "with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
            '            " from pd.projecttx" &
            '            " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
            '            " pc as (select projectid,sum(phase_status) as status from pd.projecttx where project_phase_id in (4,5,7)" &
            '            " group by projectid" &
            '            " having count(projectid) = 3)," &
            '            " lognumber as (select * from crosstab('select projectid,project_phase_id,lognumber from pd.projecttx where not lognumber isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,logpsarra text,logrp4price text,logrp4ontime text,logrp5closure text,logrampup text))" &
            '            " select r.rolename,u.username, s.sbu_shortname, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rpclosure," &
            '            " case when status isnull then null when status = 3 then 'Pass' else 'Fail' end as ""RP4_RampUp"",statusname as projectstatus,p.countstartdate as startdate," &
            '            " logpsarra,logrp4price,logrp4ontime,logrp5closure,logrampup" &
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
            '            " order by p.projectname"
            'q3.Sqlstr = "with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
            '           " from pd.projecttx" &
            '           " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
            '           " pc as (select projectid,sum(phase_status) as status from pd.projecttx where project_phase_id in (4,5,7)" &
            '           " group by projectid" &
            '           " having count(projectid) = 3)," &
            '           " lognumber as (select * from crosstab('select projectid,project_phase_id,lognumber from pd.projecttx where not lognumber isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,logpsarra text,logrp4price text,logrp4ontime text,logrp5closure text,logrampup text))" &
            '           " select r.rolename,u.username, s.sbu_shortname, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rpclosure," &
            '           " case when status isnull then null when status = 3 then 'Pass' else 'Fail' end as ""Project Launches On Time"",statusname as projectstatus,p.countstartdate as startdate," &
            '           " logpsarra,logrp4price,logrp4ontime,logrp5closure,logrampup" &
            '           " , case when txpsarra.phase_status isnull then null	" &
            '           " when txpsarra.phase_status = 1 then 'pass' else  'failed'" &
            '           " end as psarrastatus," &
            '           " case when txprice.phase_status isnull then null	" &
            '           " when txprice.phase_status = 1 then 'pass' else  'failed'" &
            '           " end as pricestatus," &
            '           " case when txontime.phase_status isnull then null" &
            '           " when txontime.phase_status = 1 then 'pass' else  'failed'" &
            '           " end as ontimestatus," &
            '           " case when txrampup.phase_status isnull then null" &
            '           " when txrampup.phase_status = 1 then 'pass' else  'failed'" &
            '           " end as rampupstatus" &
            '           " from ct" &
            '           " left join pd.project p on p.id = ct.id" &
            '           " left join pd.sbu s on s.id = p.sbu_id" &
            '           " left join pd.role r on r.id = p.role" &
            '           " left join pd.user u on u.id = p.pic" &
            '           " left join pd.vendor v on v.id = p.vendorid" &
            '           " left join pd.projecttype pt on pt.id = p.project_type_id" &
            '           " left join pc on pc.projectid = ct.id" &
            '           " left join pd.projectstatus ps on ps.id = p.project_status" &
            '           " left join lognumber l on l.id = ct.id" &
            '           " left join pd.projecttx txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
            '           " left join pd.projecttx txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
            '           " left join pd.projecttx txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
            '           " left join pd.projecttx txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
            '           " order by p.projectname"
            'q3.Sqlstr = String.Format("with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
            '          " from pd.projecttx" &
            '          " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
            '          " txadjust as( select tx.projectid,tx.project_phase_id,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
            '          " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
            '          " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
            '          " group by projectid" &
            '          " having count(projectid) = 2)," &
            '          " lognumber as (select * from crosstab('select projectid,project_phase_id,lognumber from pd.projecttx where not lognumber isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,logpsarra text,logrp4price text,logrp4ontime text,logrp5closure text,logrampup text))" &
            '          " select r.rolename,u.username, s.sbu_shortname, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rpclosure," &
            '          " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time"",statusname as projectstatus,p.countstartdate as startdate," &
            '          " logpsarra,logrp4price,logrp4ontime,logrp5closure,logrampup" &
            '          " , case when txpsarra.status isnull then null	" &
            '          " when txpsarra.status = 1 then 'pass' else  'failed'" &
            '          " end as psarrastatus," &
            '          " case when txprice.status isnull then null	" &
            '          " when txprice.status = 1 then 'pass' else  'failed'" &
            '          " end as pricestatus," &
            '          " case when txontime.status isnull then null" &
            '          " when txontime.status = 1 then 'pass' else  'failed'" &
            '          " end as ontimestatus," &
            '          " case when txrampup.status isnull then null" &
            '          " when txrampup.status = 1 then 'pass' else  'failed'" &
            '          " end as rampupstatus" &
            '          " from ct" &
            '          " left join pd.project p on p.id = ct.id" &
            '          " left join pd.sbu s on s.id = p.sbu_id" &
            '          " left join pd.role r on r.id = p.role" &
            '          " left join pd.user u on u.id = p.pic" &
            '          " left join pd.vendor v on v.id = p.vendorid" &
            '          " left join pd.projecttype pt on pt.id = p.project_type_id" &
            '          " left join pc on pc.projectid = ct.id" &
            '          " left join pd.projectstatus ps on ps.id = p.project_status" &
            '          " left join lognumber l on l.id = ct.id" &
            '          " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
            '          " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
            '          " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
            '          " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
            '          " order by p.projectname", String.Format("{0:yyyyMM}", DateTimePicker1.Value.Date))

            'q3.Sqlstr = String.Format("with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
            '         " from pd.projecttx" &
            '         " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
            '         " txadjust as( select tx.projectid,tx.project_phase_id,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
            '         " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
            '         " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
            '         " group by projectid" &
            '         " having count(projectid) = 2)," &
            '         " lognumber as (select * from crosstab('select projectid,project_phase_id,lognumber from pd.projecttx where not lognumber isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,failreportlogpsarra text,failreportlogrp4price text,failreportlogrp4ontime text,failreportlogrp5closure text,failreportlogrampup text))" &
            '         " select r.rolename,u.username, s.sbu_shortname,p.projectid, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rpclosure," &
            '         " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time"",statusname as projectstatus,p.countstartdate as startdate,p.lognumber as lognumberpsarra," &
            '         " failreportlogpsarra,failreportlogrp4price,failreportlogrp4ontime,failreportlogrp5closure,failreportlogrampup" &
            '         " , case when txpsarra.status isnull then null	" &
            '         " when txpsarra.status = 1 then 'pass' else  'failed'" &
            '         " end as psarrastatus," &
            '         " case when txprice.status isnull then null	" &
            '         " when txprice.status = 1 then 'pass' else  'failed'" &
            '         " end as pricestatus," &
            '         " case when txontime.status isnull then null" &
            '         " when txontime.status = 1 then 'pass' else  'failed'" &
            '         " end as ontimestatus," &
            '         " case when txrampup.status isnull then null" &
            '         " when txrampup.status = 1 then 'pass' else  'failed'" &
            '         " end as rampupstatus" &
            '         " from ct" &
            '         " left join pd.project p on p.id = ct.id" &
            '         " left join pd.sbu s on s.id = p.sbu_id" &
            '         " left join pd.role r on r.id = p.role" &
            '         " left join pd.user u on u.id = p.pic" &
            '         " left join pd.vendor v on v.id = p.vendorid" &
            '         " left join pd.projecttype pt on pt.id = p.project_type_id" &
            '         " left join pc on pc.projectid = ct.id" &
            '         " left join pd.projectstatus ps on ps.id = p.project_status" &
            '         " left join lognumber l on l.id = ct.id" &
            '         " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
            '         " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
            '         " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
            '         " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
            '         " order by p.projectname", String.Format("{0:yyyyMM}", DateTimePicker1.Value.Date))
            'q3.Sqlstr = String.Format("with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
            '        " from pd.projecttx" &
            '        " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
            '        " txadjust as( select tx.projectid,tx.project_phase_id,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
            '        " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
            '        " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
            '        " group by projectid" &
            '        " having count(projectid) = 2)," &
            '        " lognumber as (select * from crosstab('select projectid,project_phase_id,lognumber from pd.projecttx where not lognumber isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,failreportlogpsarra text,failreportlogrp4price text,failreportlogrp4ontime text,failreportlogrp5closure text,failreportlogrampup text))" &
            '        " select r.rolename,u.username, s.sbu_shortname,p.projectid, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rp5closure," &
            '        " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time""," &
            '        " case when pc.status = 2 then case when field5 > field7 then field5 else field7 end end as datelaunchesontime," &
            '        " statusname as projectstatus,p.countstartdate as startdate,p.lognumber as lognumberpsarra," &
            '        " failreportlogpsarra,failreportlogrp4price,failreportlogrp4ontime,failreportlogrp5closure,failreportlogrampup" &
            '        " , case when txpsarra.status isnull then null	" &
            '        " when txpsarra.status = 1 then 'pass' else  'failed'" &
            '        " end as psarrastatus," &
            '        " case when txprice.status isnull then null	" &
            '        " when txprice.status = 1 then 'pass' else  'failed'" &
            '        " end as pricestatus," &
            '        " case when txontime.status isnull then null" &
            '        " when txontime.status = 1 then 'pass' else  'failed'" &
            '        " end as ontimestatus," &
            '        " case when txrampup.status isnull then null" &
            '        " when txrampup.status = 1 then 'pass' else  'failed'" &
            '        " end as rampupstatus" &
            '        ", case p.variant when true then 'yes' else '' end as variant" &
            '        " from ct" &
            '        " left join pd.project p on p.id = ct.id" &
            '        " left join pd.sbu s on s.id = p.sbu_id" &
            '        " left join pd.role r on r.id = p.role" &
            '        " left join pd.user u on u.id = p.pic" &
            '        " left join pd.vendor v on v.id = p.vendorid" &
            '        " left join pd.projecttype pt on pt.id = p.project_type_id" &
            '        " left join pc on pc.projectid = ct.id" &
            '        " left join pd.projectstatus ps on ps.id = p.project_status" &
            '        " left join lognumber l on l.id = ct.id" &
            '        " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
            '        " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
            '        " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
            '        " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
            '        " order by p.projectname", String.Format("{0:yyyyMM}", DateTimePicker1.Value.Date))
            'q3.Sqlstr = String.Format("with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
            '        " from pd.projecttx" &
            '        " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
            '        " txadjust as( select tx.projectid,tx.project_phase_id,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
            '        " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
            '        " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
            '        " group by projectid" &
            '        " having count(projectid) = 2)," &
            '        " lognumber as (select * from crosstab('select projectid,project_phase_id,pd.getlognumber(signed_date,project_phase_id,lognum) from pd.projecttx where not lognum isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,failreportlogpsarra text,failreportlogrp4price text,failreportlogrp4ontime text,failreportlogrp5closure text,failreportlogrampup text))" &
            '        " select r.rolename,u.username, s.sbu_shortname,p.projectid, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rp5closure," &
            '        " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time""," &
            '        " case when pc.status = 2 then case when field5 > field7 then field5 else field7 end end as datelaunchesontime," &
            '        " statusname as projectstatus,p.countstartdate as startdate,pd.getlognumber(ptx.signed_date,0,p.lognum) as lognumberpsarra," &
            '        " failreportlogpsarra,failreportlogrp4price,failreportlogrp4ontime,failreportlogrp5closure,failreportlogrampup" &
            '        " , case when txpsarra.status isnull then null	when txpsarra.status = 1 then 'pass' else  'failed' end as psarrastatus," &
            '        " case when txprice.status isnull then null	 when txprice.status = 1 then 'pass' else  'failed' end as pricestatus," &
            '        " case when txontime.status isnull then null when txontime.status = 1 then 'pass' else  'failed' end as ontimestatus," &
            '        " case when txrampup.status isnull then null when txrampup.status = 1 then 'pass' else  'failed' end as rampupstatus" &
            '        ", case p.variant when true then 'yes' else '' end as variant" &
            '        " from ct" &
            '        " left join pd.project p on p.id = ct.id" &
            '        " left join pd.projecttx ptx on ptx.projectid = ct.id and project_phase_id = 3" &
            '        " left join pd.sbu s on s.id = p.sbu_id" &
            '        " left join pd.role r on r.id = p.role" &
            '        " left join pd.user u on u.id = p.pic" &
            '        " left join pd.vendor v on v.id = p.vendorid" &
            '        " left join pd.projecttype pt on pt.id = p.project_type_id" &
            '        " left join pc on pc.projectid = ct.id" &
            '        " left join pd.projectstatus ps on ps.id = p.project_status" &
            '        " left join lognumber l on l.id = ct.id" &
            '        " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
            '        " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
            '        " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
            '        " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
            '        " order by p.projectname", String.Format("{0:yyyyMM}", DateTimePicker1.Value.Date))

            'q3.Sqlstr = String.Format("with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
            '       " from pd.projecttx" &
            '       " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
            '       " txadjust as( select tx.projectid,tx.project_phase_id,lognum,docreceived,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
            '       " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
            '       " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
            '       " group by projectid" &
            '       " having count(projectid) = 2)," &
            '       " lognumber as (select * from crosstab('select projectid,project_phase_id,pd.getlognumber(yearlognum,project_phase_id,lognum) from pd.projecttx where not lognum isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,failreportlogpsarra text,failreportlogrp4price text,failreportlogrp4ontime text,failreportlogrp5closure text,failreportlogrampup text))" &
            '       " select r.rolename,u.username, s.sbu_shortname,p.projectid, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rp5closure," &
            '       " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time""," &
            '       " case when pc.status = 2 then case when field5 > field7 then field5 else field7 end end as datelaunchesontime," &
            '       " statusname as projectstatus,p.countstartdate as startdate,pd.getlognumber(p.yearlognum,0,p.lognum) as lognumberpsarra," &
            '       " failreportlogpsarra,failreportlogrp4price,failreportlogrp4ontime,failreportlogrp5closure,failreportlogrampup" &
            '       " , case when txpsarra.status isnull then null	when txpsarra.status = 1 then 'pass' else  'failed' end as psarrastatus," &
            '       " case when txprice.status isnull then null	 when txprice.status = 1 then 'pass' else  'failed' end as pricestatus," &
            '       " case when txontime.status isnull then null when txontime.status = 1 then 'pass' else  'failed' end as ontimestatus," &
            '       " case when txrampup.status isnull then null when txrampup.status = 1 then 'pass' else  'failed' end as rampupstatus," &
            '       " case when txclosure.status isnull then null when txclosure.status = 1 then 'pass' else  'failed' end as closurestatus," &
            '       " case p.variant when true then 'yes' else '' end as variant," &
            '       " case when p.lognum isnull then null when p.docreceived then 'received' else 'not received' end  as pdoc," &
            '       " case when txpsarra.lognum isnull then null when txpsarra.docreceived then 'received' else 'not received' end as psarradocf," &
            '       " case when txprice.lognum isnull then null when txprice.docreceived then 'received' else 'not received' end  as pricedocf," &
            '       " case when txontime.lognum isnull then null when  txontime.docreceived then 'received' else 'not received' end  as ontimedocf," &
            '       " case when txclosure.lognum isnull then null when txclosure.docreceived then 'received' else 'not received' end  as closuredocf," &
            '       " case when txrampup.lognum isnull then null when txrampup.docreceived then 'received' else 'not received' end as rampupdocf" &
            '       " from ct" &
            '       " left join pd.project p on p.id = ct.id" &
            '       " left join pd.projecttx ptx on ptx.projectid = ct.id and project_phase_id = 3" &
            '       " left join pd.sbu s on s.id = p.sbu_id" &
            '       " left join pd.role r on r.id = p.role" &
            '       " left join pd.user u on u.id = p.pic" &
            '       " left join pd.vendor v on v.id = p.vendorid" &
            '       " left join pd.projecttype pt on pt.id = p.project_type_id" &
            '       " left join pc on pc.projectid = ct.id" &
            '       " left join pd.projectstatus ps on ps.id = p.project_status" &
            '       " left join lognumber l on l.id = ct.id" &
            '       " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
            '       " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
            '       " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
            '       " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
            '       " left join txadjust txclosure on txclosure.projectid = ct.id and txclosure.project_phase_id = 6 " &
            '       " order by p.projectname", String.Format("{0:yyyyMM}", DateTimePicker1.Value.Date))

            q3.Sqlstr = ClassReportKPI1.getSignedProject(DateTimePicker1.Value.Date)

            q3.SheetName = "Signed Project"

            mylist.Add(q3)
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            ' Dim myreport As New ExportToExcel(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\Templates\PrintingPromotionPrice.xltx", "A1", False, False, False)
            'Dim myreport As New ExportToExcel(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\Templates\KPI-Work001.xltx")

            Dim myreport As New ExportToExcel(Me, mylist, filename, reportname, mycallback, PivotCallback, False, "\Templates\KPI-Work001New.xltx", "A1")
            myreport.Run(Me, New EventArgs)
        End If

        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub

    'Private Sub DoReport()
    '    ProgressReport(6, "Marquee")
    '    ProgressReport(1, "Loading Data.")
    '    Dim sqlstr As String = String.Empty
    '    Dim sb As New StringBuilder
    '    'DS = New DataSet

    '    Dim mymessage As String = String.Empty
    '    sb.Clear()
    '    ProgressReport(11, "InitFilter")

    '    sqlstr = String.Format(SqlStrReport, sb.ToString.ToLower)

    '    Dim mysaveform As New SaveFileDialog
    '    mysaveform.FileName = String.Format("KPI-{0:yyyyMMdd}.xlsx", Date.Today)

    '    If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
    '        Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
    '        Dim reportname = IO.Path.GetFileName(mysaveform.FileName)
    '        Dim datasheet As Integer = 6

    '        Dim mylist As New List(Of QueryWorksheet)
    '        Dim q1 As New QueryWorksheet
    '        q1.DataSheet = 6
    '        q1.Sqlstr = sqlstr
    '        q1.SheetName = "RawData"
    '        mylist.Add(q1)


    '        Dim q2 As New QueryWorksheet
    '        q2.DataSheet = 7

    '        q2.Sqlstr = "with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
    '                    " from pd.projecttx" &
    '                    " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
    '                    " pc as (select projectid,sum(phase_status) as status from pd.projecttx where project_phase_id in (4,5,7)" &
    '                    " group by projectid" &
    '                    " having count(projectid) = 3)" &
    '                    " select r.rolename,u.username, s.sbu_shortname, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rpclosure," &
    '                    " case when status isnull then null when status = 3 then 'Pass' else 'Fail' end as ""RP4_RampUp"",statusname as projectstatus" &
    '                    " from ct" &
    '                    " left join pd.project p on p.id = ct.id" &
    '                    " left join pd.sbu s on s.id = p.sbu_id" &
    '                    " left join pd.role r on r.id = p.role" &
    '                    " left join pd.user u on u.id = p.pic" &
    '                    " left join pd.vendor v on v.id = p.vendorid" &
    '                    " left join pd.projecttype pt on pt.id = p.project_type_id" &
    '                    " left join pc on pc.projectid = ct.id" &
    '                    " left join pd.projectstatus ps on ps.id = p.project_status" &
    '                    " order by p.projectname"
    '        q2.SheetName = "Signed Project"

    '        mylist.Add(q2)
    '        Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
    '        Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

    '        ' Dim myreport As New ExportToExcel(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\Templates\PrintingPromotionPrice.xltx", "A1", False, False, False)
    '        'Dim myreport As New ExportToExcel(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\Templates\KPI-Work001.xltx")

    '        Dim myreport As New ExportToExcel(Me, mylist, filename, reportname, mycallback, PivotCallback, False, "\Templates\KPI-Work001.xltx", "A1")
    '        myreport.Run(Me, New EventArgs)
    '    End If

    '    ProgressReport(1, "Loading Data.Done!")
    '    ProgressReport(5, "Continuous")
    'End Sub
    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        Dim mye = DirectCast(e, FormatReportEventargs)
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent

        Select Case mye.sheetno
            Case 6
                owb.Worksheets(6).select()
                Dim osheet As Excel.Worksheet = owb.Worksheets(6)

                Dim orange = osheet.Range("A2")
                If osheet.Cells(2, 2).text.ToString = "" Then
                    Err.Raise(100, Description:="Data not available.")
                End If

                osheet.Name = "RawData"

                owb.Names.Add("dbRange", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C1),COUNTA('RawData'!R1))")

                osheet.Columns("J:L").numberformat = "dd-MMM-yyyy"
                osheet.Rows("1:1").AutoFilter()
                osheet.Cells.EntireColumn.AutoFit()
                owb.Worksheets(1).select()
                osheet = owb.Worksheets(1)

                osheet.Cells(1, 2).value = String.Format("{0:dd-MMM-yyyy}", Date.Today)
                osheet.Cells(2, 2).value = String.Format("{0:MMM}", dateSelected)
                osheet.Cells(3, 2).value = String.Format("{0:yyyy}", dateSelected)

                osheet = owb.Worksheets(2)
                osheet.PivotTables("PivotTable1").PivotCache.Refresh()

                owb.Worksheets("Other").select()
                osheet = owb.Worksheets("other")

                osheet.PivotTables("PivotTable6").PivotFields("rank").AutoSort(Excel.XlSortOrder.xlAscending, "rank")
                For Each p As Excel.PivotItem In osheet.PivotTables("PivotTable6").pivotfields("rank").pivotitems
                    If IsNumeric(p.Value) Then
                        p.Visible = True
                    Else
                        p.Visible = False
                    End If
                Next
            Case 7
                'owb.Worksheets(7).select()
                Dim osheet As Excel.Worksheet = owb.Worksheets(7)
                osheet.Columns("S:U").numberformat = "dd-MMM-yyyy"
                osheet.Rows("1:1").AutoFilter()
                osheet.Cells.EntireColumn.AutoFit()
                owb.Names.Add("dbrangehistory", RefersToR1C1:="=OFFSET('TxHistory'!R1C1,0,0,COUNTA('TxHistory'!C1),COUNTA('TxHistory'!R1))")

                osheet = owb.Worksheets(5)
                osheet.PivotTables("PivotTable8").PivotCache.Refresh()
                osheet.PivotTables("PivotTable8").AddDataField(osheet.PivotTables("PivotTable8").PivotFields("rpphase1-4"), "Sum of rpphase1-4", Excel.XlConsolidationFunction.xlSum)

            Case 8
                'owb.Worksheets(7).select()
                Dim osheet As Excel.Worksheet = owb.Worksheets(8)
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
                osheet.Columns("AE:AJ").EntireColumn.Hidden = True
                osheet.Columns("Y:AC").EntireColumn.Hidden = True
                owb.Worksheets(1).select()
                osheet = owb.Worksheets(1)
        End Select


        'osheet = owb.Worksheets(3)
        'osheet.PivotTables("PivotTable2").PivotCache.Refresh()

        'osheet = owb.Worksheets(4)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(5)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(4)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(5)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(6)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(1)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

        'osheet = owb.Worksheets(7)
        'osheet.PivotTables("PivotTable1").PivotCache.Refresh()

    End Sub
    Private Sub PivotTable()
        ' Throw New NotImplementedException
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

            End Select
        End If
    End Sub
End Class