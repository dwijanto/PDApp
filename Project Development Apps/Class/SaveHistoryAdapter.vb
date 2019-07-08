Imports System.Text

Public Class TxHistoryAdapter
    Inherits BaseAdapter
    Implements IAdapter
    Private _period As String 'yyyyMM
    Private _startdate As Date
    Private _enddate As Date



    Public Sub New(ByVal period As Date)
        Me._period = String.Format("{0:yyyyMM}", period)
        Me._startdate = String.Format("{0:yyyy-MM-01}", period)
        Me._enddate = CDate(getLastDate(period))
    End Sub
    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        Throw New NotImplementedException
    End Function

    Public Function Save() As Boolean Implements IAdapter.Save
        Dim myret As Boolean = False
        Dim sb As New StringBuilder

        'Save to DataBase
        'Steps:
        'Delete Current Period
        'Create New Period
        sb.Append(String.Format("delete from pd.txhistory where period = {0};", _period))
        sb.Append(String.Format("insert into pd.txhistory(rolename,username,sbu,projectname,statusname,shortname,projecttype,signeddate,inputdate,postingdate,projectphase,rpstagename,rank,panel,rpphase12,rpphase34,rpphase5closure,area,period) " &
                                " select rolename,username,""SBU"",projectname,statusname,shortname,""Project Type"",signed_date,inputdate,postingdate,""Project Phase"",rp_stage_name,rank,panel,""RP Phase1-2"",""RP Phase3-4"",""RP Phase5 Closure"",""Area"",{2}" &
                                " from (select * from pd.sp_getrp15running({0},{1}) where ""Group"" = 1)foo", String.Format("'{0:yyyy-MM-dd}'", _startdate), String.Format("'{0:yyyy-MM-dd}'", _enddate), _period))
        Dim message As String = String.Empty
        If DbAdapter1.ExecuteNonQuery(sb.ToString, message:=message) Then
            If message = "" Then
                myret = True
            End If
        End If
        Return myret
    End Function

    Public Function Delete() As Boolean
        Dim myret As Boolean = False
        Dim sb As New StringBuilder

        sb.Append(String.Format("delete from pd.txhistory where period = {0};", _period))
        Dim message As String = String.Empty
        If DbAdapter1.ExecuteNonQuery(sb.ToString, message:=message) Then
            If message = "" Then
                myret = True
            End If
        End If
        Return myret
    End Function
    Private Function getLastDate(period As Date) As Object
        Dim myyear = period.Year
        Dim mymonth = period.Month
        If period.Month = 12 Then
            myyear = myyear + 1
            mymonth = 1
        Else
            mymonth = mymonth + 1
        End If
        Return CDate(String.Format("{0}-{1}-{2}", myyear, mymonth, 1)).AddDays(-1)
    End Function

End Class
