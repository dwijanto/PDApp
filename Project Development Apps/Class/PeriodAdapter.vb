Public Class PeriodAdapter
    Inherits BaseAdapter
    Implements IAdapter

    Private Shared myInstance As PeriodAdapter

    Public Shared Function getInstance()
        If myInstance Is Nothing Then
            myInstance = New PeriodAdapter
        End If
        Return myInstance
    End Function

    Public Sub New()
        MyBase.New()
    End Sub

    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        Dim myret As Boolean = True
        DS = New DataSet
        BS = New BindingSource
        Dim sqlstr = "select * from pd.period order by id desc limit 1;"
        If Not DbAdapter1.getDataSet(sqlstr, DS) Then
            Return False
        End If
        BS.DataSource = DS.Tables(0)
        Return myret
    End Function

    Public Function Save() As Boolean Implements IAdapter.Save
        Dim myret As Boolean = False
        BS.EndEdit()

        'Dim ds2 As DataSet = myInstance.DS.GetChanges
        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If DbAdapter1.PeriodTx(Me, mye) Then
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

    Public Sub GeneratePeriod(ByVal count As Integer)
        Dim currDrv As DataRowView = BS.Current
        For i = 1 To count
            For j = 1 To 12
                Dim drv As DataRowView = BS.AddNew()
                drv.Row.Item("periodid") = String.Format("{0}{1:00}", CDate(currDrv.Item("startdate")).Year + i, j)
                drv.Row.Item("startdate") = CDate(String.Format("{0}-{1}-{2}", CDate(currDrv.Item("startdate")).Year + i, j, 1))
            Next
        Next
        BS.EndEdit()

    End Sub

End Class
