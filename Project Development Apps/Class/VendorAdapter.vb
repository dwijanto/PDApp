Public Class VendorAdapter
    Inherits BaseAdapter
    Implements IAdapter

    Public Property topBS As BindingSource
    Private Shared myInstance As VendorAdapter

    Public Shared Function getInstance()
        If myInstance Is Nothing Then
            myInstance = New VendorAdapter
        End If
        Return myInstance
    End Function

    Public Sub New()
        MyBase.New()
    End Sub

    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        DS = New DataSet
        BS = New BindingSource
        topBS = New BindingSource

        SB.Clear()

        SB.Append("select v.* from pd.vendor v order by v.shortname;")
        SB.Append("select t.year::character varying,t.vendorid,t.orderline,t.id,v.shortname from  pd.topvendor t left join pd.vendor v on v.id = t.vendorid order by t.year,t.orderline,v.shortname;")
        DbAdapter1.getDataSet(SB.ToString, DS)
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


        'Unique Constrain
        Dim U0(0) As DataColumn
        U0(0) = DS.Tables(0).Columns("shortname")
        Dim UUser As UniqueConstraint = New UniqueConstraint(U0)
        DS.Tables(0).Constraints.Add(UUser)


        'Unique Constrain
        Dim U1(1) As DataColumn
        U1(0) = DS.Tables(1).Columns("year")
        U1(1) = DS.Tables(1).Columns("vendorid")
        Dim userRole As UniqueConstraint = New UniqueConstraint(U1)
        DS.Tables(1).Constraints.Add(userRole)

        DS.Tables(0).TableName = "Vendor"
        DS.Tables(1).TableName = "Top15"

        BS.DataSource = DS.Tables(0)
        topBS.DataSource = DS.Tables(1)

        Return True
    End Function

    Public Function Save() As Boolean Implements IAdapter.Save
        Dim myret As Boolean = False
        'myInstance.BS.EndEdit()
        'myInstance.topBS.EndEdit()
        BS.EndEdit()
        topBS.EndEdit()

        'Dim ds2 As DataSet = myInstance.DS.GetChanges
        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If DbAdapter1.VendorTx(Me, mye) Then
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
