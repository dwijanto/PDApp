Public Class ClassReportKPI

    Public Function getRawData(ByVal startdate As Date, ByVal lastdate As Date)
        'Return String.Format(" with r as ( select * from pd.sp_getkpi01('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date)" &
        '                             " union all " &
        '                             " select * from pd.sp_getrp15running01('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) )" &
        '                             " select r.*," &
        '                              " case p.project_type_id " &
        '                             " when 1 then 'OEM'" &
        '                             " when 2 then 'ODM'" &
        '                             " when 3 then 'PPSA' end as projecttypename," &
        '                             " case when sub.subsbuname isnull then ""SBU"" else sub.subsbuname end as subfamily" &
        '                             " from r " &
        '                             " left join pd.project p on p.id = r.projectid" &
        '                             " left join pd.subsbu sub on sub.subsbuid = p.subsbuid", startdate, lastdate)
        Return String.Format(" with r as ( select * from pd.sp_getkpi01('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date)" &
                                     " union all " &
                                     " select * from pd.sp_getrp15running01('{0:yyyy-MM-dd}'::date,'{1:yyyy-MM-dd}'::date) )" &
                                     " select r.*," &
                                      " case p.project_type_id " &
                                     " when 1 then 'OEM'" &
                                     " when 2 then 'ODM'" &
                                     " when 3 then 'PPSA' end as projecttypename," &
                                     " case when sub.subsbuname isnull then ""SBU"" else sub.subsbuname end as subfamily," &
                                     " cat.categories,pt.ptype,q.qualitylevel" &
                                     " from r " &
                                     " left join pd.project p on p.id = r.projectid" &
                                     " left join pd.subsbu sub on sub.subsbuid = p.subsbuid" &
                                     " left join pd.categories cat on cat.id = p.categoryid" &
                                     " left join pd.ptype pt on pt.id = p.ptypeid" &
                                     " left join pd.qualitylevel q on q.id = p.qualitylevelid", startdate, lastdate)

    End Function
    Public Function getSignedProject(ByVal mydate As Date) As String
        Return String.Format("with ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
                  " from pd.projecttx" &
                  " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
                  " txadjust as( select tx.projectid,tx.project_phase_id,lognum,docreceived,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
                  " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
                  " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
                  " group by projectid" &
                  " having count(projectid) = 2)," &
                  " lognumber as (select * from crosstab('select projectid,project_phase_id,pd.getlognumber(yearlognum,project_phase_id,lognum) from pd.projecttx where not lognum isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,failreportlogpsarra text,failreportlogrp4price text,failreportlogrp4ontime text,failreportlogrp5closure text,failreportlogrampup text))" &
                  " select r.rolename,u.username, s.sbu_shortname,p.projectid, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rp5closure," &
                  " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time""," &
                  " case when not pc.status isnull then case when field5 > field7 then field5 else field7 end end as datelaunchesontime," &
                  " statusname as projectstatus,p.countstartdate as startdate,pd.getlognumber(p.yearlognum,0,p.lognum) as lognumberpsarra," &
                  " failreportlogpsarra,failreportlogrp4price,failreportlogrp4ontime,failreportlogrp5closure,failreportlogrampup" &
                  " , case when txpsarra.status isnull then null	when txpsarra.status = 1 then 'pass' else  'failed' end as psarrastatus," &
                  " case when txprice.status isnull then null	 when txprice.status = 1 then 'pass' else  'failed' end as pricestatus," &
                  " case when txontime.status isnull then null when txontime.status = 1 then 'pass' else  'failed' end as ontimestatus," &
                  " case when txrampup.status isnull then null when txrampup.status = 1 then 'pass' else  'failed' end as rampupstatus," &
                  " case when txclosure.status isnull then null when txclosure.status = 1 then 'pass' else  'failed' end as closurestatus," &
                  " case p.variant when true then 'yes' else '' end as variant," &
                  " case when p.lognum isnull then null when p.docreceived then 'received' else 'not received' end  as pdoc," &
                  " case when txpsarra.lognum isnull then null when txpsarra.docreceived then 'received' else 'not received' end as psarradocf," &
                  " case when txprice.lognum isnull then null when txprice.docreceived then 'received' else 'not received' end  as pricedocf," &
                  " case when txontime.lognum isnull then null when  txontime.docreceived then 'received' else 'not received' end  as ontimedocf," &
                  " case when txclosure.lognum isnull then null when txclosure.docreceived then 'received' else 'not received' end  as closuredocf," &
                  " case when txrampup.lognum isnull then null when txrampup.docreceived then 'received' else 'not received' end as rampupdocf," &
                  " case p.project_type_id " &
                  " when 1 then 'OEM'" &
                  " when 2 then 'ODM'" &
                  " when 3 then 'PPSA' end as projecttypename," &
                  " case when sub.subsbuname isnull then sbu_shortname else sub.subsbuname end as subfamily" &
                  " from ct" &
                  " left join pd.project p on p.id = ct.id" &
                  " left join pd.projecttx ptx on ptx.projectid = ct.id and project_phase_id = 3" &
                  " left join pd.sbu s on s.id = p.sbu_id" &
                  " left join pd.role r on r.id = p.role" &
                  " left join pd.user u on u.id = p.pic" &
                  " left join pd.vendor v on v.id = p.vendorid" &
                  " left join pd.projecttype pt on pt.id = p.project_type_id" &
                  " left join pc on pc.projectid = ct.id" &
                  " left join pd.projectstatus ps on ps.id = p.project_status" &
                  " left join lognumber l on l.id = ct.id" &
                  " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
                  " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
                  " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
                  " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
                  " left join txadjust txclosure on txclosure.projectid = ct.id and txclosure.project_phase_id = 6 " &
                  " left join pd.subsbu sub on sub.subsbuid = p.subsbuid" &
                  " order by p.projectname", String.Format("{0:yyyyMM}", mydate))

    End Function

    Public Function getSignedProject(ByVal userid As String, ByVal criteria As String) As String
        Return String.Format("with ur as(select ur.roleid from pd.user_role ur left join pd.user u on u.id = ur.userid   where lower(u.userid) = '{1}'), " &
                        " ct as (select * from crosstab('select projectid,project_phase_id,postingdate " &
          " from pd.projecttx" &
          " order by projectid','select m from generate_series(1,7) m') as  ct(id bigint,field1 date,field2 date,field3 date,field4 date,field5 date,field6 date,field7 date))," &
          " txadjust as( select tx.projectid,lognum,docreceived,tx.project_phase_id,case when  pa.phase_status isnull then tx.phase_status else pa.phase_status end as status" &
          " from pd.projecttx tx left join pd.projectadjustment pa on pa.projectid = tx.projectid and pa.project_phase_id = tx.project_phase_id and to_char(pa.postingdate,'yyyyMM')::integer <= {0} )," &
          " pc as (select projectid,sum(status) as status from txadjust where project_phase_id in (5,7)" &
          " group by projectid" &
          " having count(projectid) = 2)," &
          " lognumber as (select * from crosstab('select projectid,project_phase_id,pd.getlognumber(yearlognum,project_phase_id,lognum) as lognumber from pd.projecttx where not lognum isnull order by projectid','select m from generate_series(3,7) m') as  ct(id bigint,failreportlogpsarra text,failreportlogrp4price text,failreportlogrp4ontime text,failreportlogrp5closure text,failreportlogrampup text))" &
          " select r.rolename,u.username, s.sbu_shortname,p.projectid, p.projectname,v.shortname,project_type_rpt,field1 as rp1,field2 as rp2,field3 as ""psa/rra"",field4 as rp4price,field5 as rp4ontime,field7 as rampup,field6 as rp5closure," &
          " case when pc.status isnull then null when pc.status = 2 then 'Pass' else 'Fail' end as ""Project Launches On Time""," &
          " case when not pc.status isnull then case when field5 > field7 then field5 else field7 end end as datelaunchesontime," &
          " statusname as projectstatus,p.countstartdate as startdate,pd.getlognumber(p.yearlognum,0,p.lognum) as lognumberpsarra," &
          " failreportlogpsarra,failreportlogrp4price,failreportlogrp4ontime,failreportlogrampup,failreportlogrp5closure," &
          " case when txpsarra.status isnull then null	when txpsarra.status = 1 then 'pass' else  'failed' end as psarrastatus," &
          " case when txprice.status isnull then null	when txprice.status = 1 then 'pass' else  'failed' end as pricestatus," &
          " case when txontime.status isnull then null when txontime.status = 1 then 'pass' else  'failed' end as ontimestatus," &
          " case when txrampup.status isnull then null when txrampup.status = 1 then 'pass' else  'failed' end as rampupstatus," &
          " case when txclosure.status isnull then null when txclosure.status = 1 then 'pass' else  'failed' end as closurestatus," &
          " case p.variant when true then 'yes' else '' end as variant," &
          " case when p.lognum isnull then null when p.docreceived then 'received' else 'not received' end  as pdoc," &
          " case when txpsarra.lognum isnull then null when txpsarra.docreceived then 'received' else 'not received' end as psarradocf," &
          " case when txprice.lognum isnull then null when txprice.docreceived then 'received' else 'not received' end  as pricedocf," &
          " case when txontime.lognum isnull then null when  txontime.docreceived then 'received' else 'not received' end  as ontimedocf," &
          " case when txrampup.lognum isnull then null when txrampup.docreceived then 'received' else 'not received' end as rampupdocf," &
          " case when txclosure.lognum isnull then null when txclosure.docreceived then 'received' else 'not received' end  as closuredocf," &
          " case p.project_type_id " &
          " when 1 then 'OEM'" &
          " when 2 then 'ODM'" &
          " when 3 then 'PPSA' end as projecttypename," &
          " case when sub.subsbuname isnull then sbu_shortname else sub.subsbuname end as subfamily" &
          " from ct" &
          " left join pd.project p on p.id = ct.id" &
          " left join pd.projecttx ptx on ptx.projectid = ct.id and project_phase_id = 3" &
          " left join pd.sbu s on s.id = p.sbu_id" &
          " left join pd.role r on r.id = p.role" &
          " left join pd.user u on u.id = p.pic" &
          " left join pd.vendor v on v.id = p.vendorid" &
          " left join pd.projecttype pt on pt.id = p.project_type_id" &
          " left join pc on pc.projectid = ct.id" &
          " left join pd.projectstatus ps on ps.id = p.project_status" &
          " left join lognumber l on l.id = ct.id" &
          " left join txadjust txpsarra on txpsarra.projectid = ct.id and txpsarra.project_phase_id = 3 " &
          " left join txadjust txprice on txprice.projectid = ct.id and txprice.project_phase_id = 4 " &
          " left join txadjust txontime on txontime.projectid = ct.id and txontime.project_phase_id = 5 " &
          " left join txadjust txrampup on txrampup.projectid = ct.id and txrampup.project_phase_id = 7 " &
          " left join txadjust txclosure on txclosure.projectid = ct.id and txclosure.project_phase_id = 6 " &
          " left join pd.subsbu sub on sub.subsbuid = p.subsbuid" &
          " {2} " &
          " order by p.projectname", String.Format("{0:yyyyMM}", Today.Date), userid.ToString.ToLower, criteria)
    End Function
End Class
