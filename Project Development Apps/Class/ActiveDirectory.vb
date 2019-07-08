Imports System.DirectoryServices
Public Class ActiveDirectory
    Public Property UserInfo As UserInfo

    Public Function getDataAD(ByRef userid As String) As Boolean
        _UserInfo = New UserInfo

        Dim entry As DirectoryEntry = New DirectoryEntry
        entry.Path = "Cannot be blank"

        Dim myuser() As String = userid.Split("\")
        _UserInfo.Domain = myuser(0)
        _UserInfo.Userid = myuser(1)
        Select Case _UserInfo.Domain.ToString.ToLower
            Case "as"
                entry.Path = "LDAP://as/DC=as;DC=seb;DC=com"
            Case "eu"
                entry.Path = "LDAP://eu/DC=eu;DC=seb;DC=com"
            Case "supor"
                entry.Path = "LDAP://supor/DC=supor;DC=seb;DC=com"
            Case "sa"
                entry.Path = "LDAP://sa/DC=sa;DC=seb;DC=com"
            Case "na"
                entry.Path = "LDAP://na/DC=na;DC=seb;DC=com"
        End Select

        'Try
        Dim mysearcher As DirectorySearcher = New DirectorySearcher(entry)
        mysearcher.Filter = "sAMAccountName=" & _UserInfo.Userid
        Dim result As SearchResult = mysearcher.FindOne
        If Not (result Is Nothing) Then
            Dim myPerson As DirectoryEntry = New DirectoryEntry
            myPerson.Path = result.Path

            Try
                _UserInfo.email = myPerson.Properties("mail").Value.ToString
            Catch ex As Exception
            End Try

            Try
                _UserInfo.DisplayName = UCase(myPerson.Properties("givenname").Value.ToString) & " " & UCase(myPerson.Properties("sn").Value.ToString)
            Catch
            End Try

            Try
                _UserInfo.Department = UCase(myPerson.Properties("department").Value.ToString)
            Catch
            End Try
            'Dim counter As Integer = 0
            'For Each elemantName In myPerson.Properties.PropertyNames

            '    Dim valuecollection As PropertyValueCollection = myPerson.Properties(elemantName)
            '    For i = 0 To valuecollection.Count - 1
            '        'Debug.WriteLine(elemantName + "[" + i.ToString() + "]=" + valuecollection(i).ToString())
            '        Debug.WriteLine("{0} {1}[{2}] = {3}", counter, elemantName, i, valuecollection(i).ToString)
            '    Next
            '    counter += 1
            'Next
        Else
            Err.Raise(1)
        End If
        'Catch ex As Exception

        'End Try
        myuser = Nothing
        Return True
    End Function
End Class
