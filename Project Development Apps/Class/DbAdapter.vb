Imports Npgsql


Imports System.IO
Public Class DbAdapter
    Implements IDisposable

    Dim _ConnectionStringDict As Dictionary(Of String, String)
    Dim _connectionstring As String
    Private CopyIn1 As NpgsqlCopyIn
    Dim _userid As String
    Dim _password As String
    Dim mytransaction As NpgsqlTransaction
    Public Property UserInfo As UserInfo
    Public Shared myInstance As DbAdapter

    Public Shared Function getInstance() As DbAdapter
        If myInstance Is Nothing Then
            myInstance = New DbAdapter
        End If
        Return myInstance
    End Function

    Public ReadOnly Property userid As String
        Get
            Return _userid
        End Get
    End Property
    Public ReadOnly Property password As String
        Get
            Return _password
        End Get
    End Property

    Public Property Connectionstring As String
        Get
            Return _connectionstring

        End Get
        Set(ByVal value As String)
            _connectionstring = value
        End Set
    End Property

    Private Sub New()
        InitConnectionStringDict()
        '_connectionstring = getConnectionString()
    End Sub

    Public ReadOnly Property ConnectionStringDict As Dictionary(Of String, String)
        Get
            Return _ConnectionStringDict
        End Get
    End Property

    Private Sub InitConnectionStringDict()
        _ConnectionStringDict = New Dictionary(Of String, String)
        _connectionstring = getConnectionString()
        'Dim connectionstring = getConnectionString()
        Dim connectionstrings() As String = _connectionstring.Split(";")
        For i = 0 To (connectionstrings.Length - 1)
            Dim mystrs() As String = connectionstrings(i).Split("=")
            _ConnectionStringDict.Add(mystrs(0), mystrs(1))
        Next i

    End Sub

    Private Function getConnectionString() As String
        _userid = "admin"
        _password = "admin"
        Dim builder As New NpgsqlConnectionStringBuilder()
        builder.ConnectionString = My.Settings.ConnectionString1
        builder.Add("User Id", _userid)
        builder.Add("password", _password)
        Return builder.ConnectionString
    End Function

#Region "GetDataSet"
    Public Overloads Function getDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter

        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            Dim obj = TryCast(ex.Errors(0), NpgsqlError)
            Dim myerror As String = String.Empty
            If Not IsNothing(obj) Then
                myerror = obj.InternalQuery
            End If
            message = ex.Message & " " & myerror
        End Try
        Return myret
    End Function
#End Region

    Public Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(getConnectionString())
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message & ", " & ex.Detail
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message & ", " & ex2.Detail
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function


    Public Function validint(ByVal myqty As String) As Object

        If myqty = "" Then
            Return DBNull.Value
        Else
            Return CInt(myqty.Replace(",", "").Replace("""", ""))
        End If
    End Function
    Public Function validbool(ByVal mybool As String) As Object
        If mybool = "Y" Then
            Return "True"
        Else
            Return "False"
        End If
    End Function
    Public Function validdec(ByVal mydec As String) As Object
        If mydec = "" Then
            Return DBNull.Value
        Else
            Return CDec(mydec.Replace(",", "").Replace("""", ""))
        End If
    End Function

    Public Function validlong(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return DBNull.Value
        Else
            Return CLng(myvalue)
        End If
    End Function
    Public Function validlongNull(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return "Null"
        Else
            Return CLng(myvalue)
        End If
    End Function
    Public Function validchar(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return ""
        Else
            Return Trim(myvalue.Replace("'", "''").Replace("""", ""))
        End If
    End Function
    Public Function validcharNull(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return "Null"
        Else
            Return Trim(myvalue.Replace("'", "''").Replace("""", ""))
        End If
    End Function
    Public Function CDateddMMyyyy(ByVal updatedate As String) As Object
        Dim mydata() As String
        If updatedate.Contains(".") Then
            mydata = updatedate.Split(".")
        Else
            mydata = updatedate.Split("/")
        End If

        If mydata.Length > 1 Then
            Return CDate(mydata(2) & "-" & mydata(1) & "-" & mydata(0))
        End If
        Return DBNull.Value
    End Function
    Public Function ddMMyyyytoyyyyMMdd(ByVal updatedate As String) As Object
        Dim mydata() As String
        If updatedate.Contains(".") Then
            mydata = updatedate.Split(".")
        Else
            mydata = updatedate.Split("/")
        End If

        If mydata.Length > 1 Then
            Return "'" & mydata(2) & "-" & mydata(1) & "-" & mydata(0) & "'"
        End If
        Return DBNull.Value
    End Function


    Public Function ExNonQuery(ByVal sqlstr As String) As Long
        Dim myRet As Long
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                myRet = command.ExecuteNonQuery
            End Using
        End Using
        Return myRet
    End Function

    Public Function ExecuteNonQuery(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteNonQuery
                    'recordAffected = command.ExecuteNonQuery
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteNonQueryAsync(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteNonQuery

                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteScalar(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteScalar(ByVal sqlstr As String, ByRef recordAffected As Object, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function

    Sub ExecuteStoreProcedure(ByVal storeprocedurename As String)
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand(storeprocedurename, conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.ExecuteScalar()
            Catch ex As Exception

            End Try

        End Using
    End Sub

    Public Function getproglock(ByVal programname As String, ByVal userid As String, ByVal status As Integer) As Boolean
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("proglock", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = programname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = status
            result = cmd.ExecuteScalar
        End Using

        Return result
    End Function

    Function dateformatdot(ByVal myrecord As String) As Object
        Dim myreturn = "Null"
        If myrecord = "" Then
            Return myreturn
        End If
        Dim mysplit = Split(myrecord, ".")
        myreturn = "'" & mysplit(2) & "-" & mysplit(1) & "-" & mysplit(0) & "'"
        Return myreturn
    End Function

    Function dateformatdotdate(ByVal myrecord As String) As Date
        Dim myreturn = "Null"
        If myrecord = "" Then
            Return myreturn
        End If
        Dim mysplit = Split(myrecord, ".")
        myreturn = CDate(mysplit(2) & "-" & mysplit(1) & "-" & mysplit(0))
        Return myreturn
    End Function

    Function dateformatYYYYMMdd(ByVal myrecord As Object) As Object
        Dim myreturn = "Null"

        myreturn = "'" & CDate(myrecord).Year & "-" & CDate(myrecord).Month & "-" & CDate(myrecord).Day & "'"
        Return myreturn
    End Function

    Private Sub onRowInsertUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        'Table with autoincrement
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub

    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date)
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function

    Function IsAdmin(ByVal p1 As String) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim Command As New NpgsqlCommand
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Command = New NpgsqlCommand("pd.sp_isadmin", conn)
                Command.CommandType = CommandType.StoredProcedure
                Command.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Value = p1.ToLower
                Dim result As Object = Command.ExecuteScalar
                If Not IsNothing(result) Then
                    myret = result
                End If
            End Using
        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            Return False
        End Try
        Return myret
    End Function

    Function CloseConfirmation(ByVal userid As String) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim Command As New NpgsqlCommand
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Command = New NpgsqlCommand("pd.sp_closeconf", conn)
                Command.CommandType = CommandType.StoredProcedure
                Command.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").Value = userid.ToLower
                Dim result As Object = Command.ExecuteScalar
                If Not IsNothing(result) Then
                    myret = result
                End If
            End Using
        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            Return False
        End Try
        Return myret
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    'Public Function GetDataset(sqlstr, dataset) As Boolean
    '    Dim DataAdapter As New NpgsqlDataAdapter
    '    Dim myret As Boolean = False
    '    Using conn As New NpgsqlConnection(connectionString)
    '        conn.Open()
    '        DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
    '        DataAdapter.Fill(dataset)
    '        myret = True
    '    End Using
    '    Return myret
    'End Function

    Function UserTx(ByVal formUser As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)

        Using conn As New NpgsqlConnection(connectionString)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            sqlstr = "pd.sp_updateuser"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isadmin").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "title").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_insertuser"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isadmin").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "title").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_deleteuser"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(0))


            'Update
            sqlstr = "pd.sp_updaterole"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "rolename").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "area").SourceVersion = DataRowVersion.Current

            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_insertrole"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "rolename").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "area").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_deleterole"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


            sqlstr = "pd.sp_insertuserrole"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "userid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "roleid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_deleteuserrole"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(2))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function VendorTx(vendorAdapter As VendorAdapter, mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)

        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            mytransaction = conn.BeginTransaction

            'Update
            sqlstr = "pd.sp_updatevendor"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "panel").SourceVersion = DataRowVersion.Current

            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_insertvendor"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "shortname").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "panel").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_deletevendor"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

            'Update
            sqlstr = "pd.sp_updatetopvendor"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "year").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "orderline").SourceVersion = DataRowVersion.Current

            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_inserttopvendor"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "year").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "orderline").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_deletetopvendor"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function ProjectTx(projectAdapter As ProjectAdapter, mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)

        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            sqlstr = "pd.sp_updateproject"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "project_type_id").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sbu_id").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "countasrunningproject").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "role").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pic").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "project_status").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "creationdate").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "variant").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "countstartdate").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "countenddate").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "lognumber").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "countasrp1").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "subsbuid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_insertproject"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectname").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "project_type_id").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sbu_id").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "countasrunningproject").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "role").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pic").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "project_status").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "creationdate").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "variant").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "countstartdate").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "countenddate").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "lognumber").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "countasrp1").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "subsbuid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_deleteproject"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

           

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function ProjectSignedTx(projectSignedAdapter As ProjectSignedAdapter, mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)

        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            sqlstr = "pd.sp_updateprojecttx"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "postingdate").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "project_phase_id").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "signed_date").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "phase_status").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "stamp").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "lognumber").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_insertprojecttx"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "postingdate").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "project_phase_id").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "signed_date").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "phase_status").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "stamp").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "lognumber").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_deleteprojecttx"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

            sqlstr = "pd.sp_updateproject"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
            mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function
    Function ProjectSignedTx(projectSignedAdapter As ProjectSignedAdapter2, mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)



        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            sqlstr = "pd.sp_updateprojecttx"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "postingdate").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "project_phase_id").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "signed_date").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "phase_status").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "stamp").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lognum").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "yearlognum").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "docreceived").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_insertprojecttx"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "postingdate").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "project_phase_id").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "signed_date").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "phase_status").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "stamp").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lognum").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "yearlognum").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "docreceived").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_deleteprojecttx"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

            sqlstr = "pd.sp_updateproject"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "lognum").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "yearlognum").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "docreceived").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
            mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


            mytransaction.Commit()
            myret = True
        End Using

        Return myret
    End Function
    Function ProjectAdjustmentTx(projectTxAdjustment As ProjectAdjustmentAdapter, mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)

        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            sqlstr = "pd.sp_updateprojectadjustment"
            DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "postingdate").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "project_phase_id").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "signed_date").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reason").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "phase_status").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "stamp").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_insertprojectadjustment"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "projectid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "postingdate").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "project_phase_id").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "signed_date").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "reason").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "phase_status").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "stamp").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "modifiedby").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "pd.sp_deleteprojectadjustment"
            DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            DataAdapter.InsertCommand.Transaction = mytransaction
            DataAdapter.UpdateCommand.Transaction = mytransaction
            DataAdapter.DeleteCommand.Transaction = mytransaction

            mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function PeriodTx(periodAdapter As PeriodAdapter, mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)

        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            sqlstr = "pd.sp_insertperiod"
            DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "periodid").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "startdate").SourceVersion = DataRowVersion.Current
            DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure         
            DataAdapter.InsertCommand.Transaction = mytransaction        
            mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function


End Class
