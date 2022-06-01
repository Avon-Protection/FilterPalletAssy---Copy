Imports System.IO
Imports System.Data.SqlClient
Imports System.Web.Security
Imports System.Data
Imports System.Net.Mail
Imports System.Security.Principal.WindowsIdentity
Imports System.Collections
Imports System.Collections.Specialized
Imports ActiveDs
Imports System.DirectoryServices
Imports System.Xml
Imports System.Xml.Xsl
Imports System.Web
Imports System.DBNull
Imports System.TimeSpan
Imports Microsoft.VisualBasic


Namespace commonClasses
    Public Class sqlInterface
        Public Sub executeFileQuery(ByVal connectionString As String, ByVal fileName As String)
            Dim objStream As StreamReader
            Dim sqlCmd As String
            Dim cmd As New SqlCommand
            Dim sqlCon As New SqlConnection

            sqlCon = New SqlConnection(connectionString)

            If sqlCon.State <> System.Data.ConnectionState.Open Then 'If state is anything other than OPEN
                sqlCon.Close() 'Close the connection to eliminate possible error conditions
                sqlCon.Open() 'Reopen to run the sqlcommand
            End If

            objStream = File.OpenText(fileName)
            sqlCmd = objStream.ReadToEnd()
            objStream.Close()

            cmd = New SqlCommand(sqlCmd, sqlCon)
            cmd.CommandType = System.Data.CommandType.Text
            cmd.ExecuteNonQuery()
        End Sub
        Public Function executeStoredProcNonQuery(ByVal connectionString As String, ByVal storedProc As String, ByVal params() As SqlParameter) As ArrayList
            Dim sqlCon As New SqlConnection
            Dim cmd As New SqlCommand
            Dim sqlParam As New SqlParameter
            Dim sqlAdap As New SqlDataAdapter
            Dim aList As New ArrayList
            Dim numParams As Integer

            sqlCon = New SqlConnection(connectionString)

            If sqlCon.State <> System.Data.ConnectionState.Open Then 'If state is anything other than OPEN
                sqlCon.Close() 'Close the connection to eliminate possible error conditions
                sqlCon.Open() 'Reopen to run the sqlcommand
            End If

            cmd = New SqlCommand(storedProc, sqlCon)
            cmd.CommandType = CommandType.StoredProcedure

            Try
                If params IsNot Nothing Then
                    For Each sqlParam In params
                        cmd.Parameters.Add(sqlParam)
                    Next
                    numParams = cmd.Parameters.Count
                End If
                cmd.ExecuteNonQuery()
                sqlCon.Close()
                SqlConnection.ClearPool(sqlCon)

                If params IsNot Nothing Then
                    For Each sqlParam In cmd.Parameters
                        aList.Add(sqlParam.Value)
                    Next
                End If

                Return aList
            Catch ex As Exception
                sqlCon.Close()
                Return Nothing
            End Try

        End Function
        Public Function executeStoredProcRtrnValue(ByVal connectionString As String, ByVal storedProc As String, ByVal params() As SqlParameter) As Integer
            Dim sqlCon As New SqlConnection
            Dim cmd As New SqlCommand
            Dim sqlParam As New SqlParameter
            Dim sqlAdap As New SqlDataAdapter
            Dim result As Integer
            Dim numParams As Integer

            sqlCon = New SqlConnection(connectionString)

            If sqlCon.State <> System.Data.ConnectionState.Open Then 'If state is anything other than OPEN
                sqlCon.Close() 'Close the connection to eliminate possible error conditions
                sqlCon.Open() 'Reopen to run the sqlcommand
            End If

            cmd = New SqlCommand(storedProc, sqlCon)
            cmd.CommandType = CommandType.StoredProcedure

            If params IsNot Nothing Then
                For Each sqlParam In params
                    cmd.Parameters.Add(sqlParam)
                Next
                numParams = cmd.Parameters.Count
            End If
            cmd.ExecuteNonQuery()
            sqlCon.Close()
            SqlConnection.ClearPool(sqlCon)

            If params IsNot Nothing Then
                For Each sqlParam In cmd.Parameters
                    If sqlParam.ParameterName = "@ReturnValue" Then
                        result = sqlParam.Value
                    End If
                Next
            End If

            Return result
        End Function
        Public Function executeStoredProcRtrnObj(ByVal connectionString As String, ByVal storedProc As String, ByVal params() As SqlParameter) As Object
            ' THIS PROCEDURE RETURNS A SINGLE VALUE WHICH IS AN INTEGER WHICH CAN BE USED TO 
            ' DETERMINE THAT STATUS OF A STORED PROCEDURE.
            ' ** NOTE ** 
            ' ENSURE THAT THE STORED PROCEDURE SELECTS A SINGLE VARIABLE TO RETURN BECAUSE
            ' THE CMD.EXECUTESCALAR DOES NOT RETRIEVE THE RETURN CODE
            Dim sqlCon As New SqlConnection
            Dim cmd As New SqlCommand
            Dim sqlParam As New SqlParameter
            Dim sqlAdap As New SqlDataAdapter
            Dim value As Object

            sqlCon = New SqlConnection(connectionString)

            If sqlCon.State <> System.Data.ConnectionState.Open Then 'If state is anything other than OPEN
                sqlCon.Close() 'Close the connection to eliminate possible error conditions
                sqlCon.Open() 'Reopen to run the sqlcommand
            End If

            cmd = New SqlCommand(storedProc, sqlCon)
            cmd.CommandType = CommandType.StoredProcedure

            Try
                If params IsNot Nothing Then
                    For Each sqlParam In params
                        cmd.Parameters.Add(sqlParam)
                    Next
                End If
                value = cmd.ExecuteScalar()
                cmd.Parameters.Clear()
                cmd.Dispose()
                sqlCon.Close()
                SqlConnection.ClearPool(sqlCon)

                Return value
            Catch ex As Exception
                sqlCon.Close()
                Return Nothing
            Finally
                sqlCon.Dispose()
            End Try

        End Function
        Public Function executeStoredProc(ByVal connectionString As String, ByVal storedProc As String, ByVal params() As SqlParameter, ByVal tableName As String) As DataSet
            Dim ds As New DataSet
            Dim sqlCon As New SqlConnection
            Dim cmd As New SqlCommand
            Dim sqlAdap As New SqlDataAdapter
            Dim sqlParam As New SqlParameter
            Dim sqlParam1 As New SqlParameter
            Dim conTimeOut As Integer = 600

            sqlCon = New SqlConnection(connectionString)

            If sqlCon.State <> System.Data.ConnectionState.Open Then 'If state is anything other than OPEN
                sqlCon.Close() 'Close the connection to eliminate possible error conditions
                sqlCon.Open() 'Reopen to run the sqlcommand
            End If

            cmd = New SqlCommand(storedProc, sqlCon)
            cmd.CommandTimeout = 600
            cmd.CommandType = CommandType.StoredProcedure
            sqlAdap = New SqlDataAdapter(cmd)

            If params IsNot Nothing Then
                For Each sqlParam In params
                    sqlAdap.SelectCommand.Parameters.Add(sqlParam)
                Next
            End If

            Try
                sqlAdap.Fill(ds, tableName)
                sqlCon.Close()
                SqlConnection.ClearPool(sqlCon)

                If params IsNot Nothing Then
                    For Each sqlParam In cmd.Parameters
                        If sqlParam.ParameterName = "@ReturnValue" Then
                            params(params.Length - 1).SqlValue = sqlParam.Value
                        End If
                    Next
                End If

                Return ds
            Catch ex As System.Exception
                sqlCon.Close()
                Return Nothing
            Finally
                sqlCon.Dispose()
            End Try
        End Function
        Public Function executeStoredProcOutput(ByVal connectionString As String, ByVal storedProc As String, ByRef params() As SqlParameter) As DataSet
            Dim ds As New DataSet
            Dim sqlCon As New SqlConnection
            Dim cmd As New SqlCommand
            Dim sqlParam As New SqlParameter
            Dim sqlParam1 As New SqlParameter
            Dim conTimeOut As Integer = 600
            Dim paramsIndex As Integer = 0

            sqlCon = New SqlConnection(connectionString)

            If sqlCon.State <> System.Data.ConnectionState.Open Then 'If state is anything other than OPEN
                sqlCon.Close() 'Close the connection to eliminate possible error conditions
                sqlCon.Open() 'Reopen to run the sqlcommand
            End If

            cmd = New SqlCommand(storedProc, sqlCon)
            cmd.CommandTimeout = 600
            cmd.CommandType = CommandType.StoredProcedure

            If params IsNot Nothing Then
                For Each sqlParam In params
                    cmd.Parameters.Add(sqlParam)
                Next
            End If

            Try
                cmd.ExecuteNonQuery()

                 If params IsNot Nothing Then
                    For Each sqlParam In cmd.Parameters
                        params(paramsIndex).SqlValue = sqlParam.Value
                    Next
                End If

                sqlCon.Close()
                SqlConnection.ClearPool(sqlCon)

               

                Return ds
            Catch ex As System.Exception
                sqlCon.Close()
                Return Nothing
            Finally
                sqlCon.Dispose()
            End Try
        End Function
        Public Function executeSelectQuery(ByVal connectionString As String, ByVal sqlStmt As String, ByVal tableName As String) As DataTable
            Dim dt As New DataTable(tableName)
            Dim dr As DataRow
            Dim sqlCon As New SqlConnection
            Dim cmd As New SqlCommand
            Dim sqlReader As SqlDataReader
            Dim header As String

            Try
                sqlCon = New SqlConnection(connectionString)

                If sqlCon.State <> System.Data.ConnectionState.Open Then 'If state is anything other than OPEN
                    sqlCon.Close() 'Close the connection to eliminate possible error conditions
                    sqlCon.Open() 'Reopen to run the sqlcommand
                End If

                cmd = New SqlCommand(sqlStmt, sqlCon)
                cmd.CommandTimeout = 600

                sqlReader = cmd.ExecuteReader()

                For i As Integer = 0 To sqlReader.FieldCount - 1
                    header = sqlReader.GetName(i)
                    dt.Columns.Add(header)
                Next

                While sqlReader.Read()
                    dr = dt.NewRow()

                    dr.Item(0) = sqlReader.GetValue(0).ToString()

                    For i As Integer = 1 To sqlReader.FieldCount - 1
                        dr.Item(i) = sqlReader.GetValue(i).ToString()
                    Next

                    dt.Rows.Add(dr)

                End While
                sqlReader.Close()
                Return dt
            Catch ex As Exception
                sqlCon.Close()
                Return Nothing
            Finally
                sqlCon.Dispose()
            End Try

        End Function
    End Class
    Public Class windowsInterface
        Public Function getLoggedInUser() As String
            Dim userName As String

            'userName = HttpContext.Current.Request.LogonUserIdentity.Name()
            userName = System.Web.HttpContext.Current.User.Identity.Name
            'userName = HttpContext.Current.Request.ServerVariables("AUTH_USER").ToString()

            Return userName
        End Function
    End Class
    Public Class emailInterface
        Public Sub sendMessage(ByVal toAddress As String, ByVal ccAddress As String, ByVal subject As String, ByVal body As String)
            Dim mailMessage As New MailMessage
			Dim smtpClient As New SmtpClient("smtp-na.avon.local")

			Try
                mailMessage.From = New MailAddress("DONOTREPLY@avon-rubber.com")
                mailMessage.To.Add(New MailAddress(toAddress))
                If (ccAddress <> Nothing) Then
                    mailMessage.CC.Add(New MailAddress(ccAddress))
                End If
                mailMessage.Subject = subject
                mailMessage.Body = body
                mailMessage.IsBodyHtml = True
                mailMessage.Priority = MailPriority.Normal

                smtpClient.Send(mailMessage)
            Catch ex As System.Exception
                Dim mailMessage2 As New MailMessage
				Dim smtpClient2 As New SmtpClient("smtp-na.avon.local")

				mailMessage2.From = New MailAddress("DONOTREPLY@avon-rubber.com")
                mailMessage2.To.Add(New MailAddress("scott.jesweak@avon-rubber.com"))
                mailMessage2.Subject = "EMAIL ERROR"
                mailMessage2.Body = ex.Message
                mailMessage2.IsBodyHtml = True
                mailMessage2.Priority = MailPriority.Normal

                smtpClient2.Send(mailMessage2)

            End Try
        End Sub
    End Class
    Public Class ADInterface
        Dim emailMessage As New emailInterface
        Dim emailSubject As String
        Dim emailBody As String

        Public Function GetDirectoryEntry(ByVal dirPath As String) As DirectoryEntry
            Dim dirEntry As DirectoryEntry
            dirEntry = New DirectoryEntry
            dirEntry.Path = dirPath
            dirEntry.Username = "avon\cadadmin"
            dirEntry.Password = "lz396gd417"
            Return dirEntry
        End Function
        Public Function UserExists(ByVal UserName As String, ByVal dirEntry As DirectoryEntry) As Boolean
            Dim deSearch As DirectorySearcher = New DirectorySearcher()
            Dim results As SearchResultCollection

            Try
                deSearch.SearchRoot = dirEntry
                deSearch.Filter = "(&(objectClass=user) (cn=" & UserName & "))"
                results = deSearch.FindAll()
                If results.Count = 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As System.Exception
                Return False
            End Try
        End Function
        Public Function getADGroupName(ByVal adPath As String) As String
            Dim groupName As String = ""
            Dim myValue As ResultPropertyValueCollection
            Dim retEntry As DirectoryEntry = New DirectoryEntry()
            Dim objResult As SearchResult
            Dim path As String
            Dim dirEntry As DirectoryEntry = New DirectoryEntry("GC://" & adPath, "avon\cadadmin", "lz396gd417")
            Dim search As DirectorySearcher = New DirectorySearcher(dirEntry)

            Try
                objResult = search.FindOne()

                retEntry = objResult.GetDirectoryEntry()
                path = retEntry.Path

                Dim myResultPropColl As ResultPropertyCollection
                myResultPropColl = objResult.Properties

                ' CODE TO GO THROUGH EACH PROPERTY
                'Dim myKey As String
                'For Each myKey In myResultPropColl.PropertyNames
                '    Dim myCollection As Object
                '    For Each myCollection In myResultPropColl(myKey)
                '        myValue = myResultPropColl("mail")
                '    Next myCollection
                'Next myKey

                myValue = myResultPropColl("name")
                If myValue.Count <> 0 Then
                    groupName = myValue.Item(0)
                End If

                Return groupName
            Catch ex As Exception
                emailSubject = "APPLICATION ERROR: SPARES MANAGEMENT : getADGroupName"
                emailBody = ex.Message
                emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
            End Try

        End Function
        Public Function createMemberList(ByVal adPath As String) As ArrayList
            Dim memberList As New ArrayList
            Dim objType As String
            Dim groupName As String = ""
            Dim tempList As New ArrayList
            Dim groupList As ArrayList
            Dim aString As String = ""
            Dim bString As String = ""

            Try
                objType = getObjectType(adPath)
                If objType = "group" Then
                    groupName = getADGroupName(adPath)
                    groupList = getADGroupMembers(groupName)

                    For Each aString In groupList
                        tempList = createMemberList(aString)
                        If tempList.Count > 1 Then
                            For Each bString In tempList
                                memberList.Add(bString)
                            Next
                        Else
                            memberList.Add(tempList.Item(0))
                        End If
                    Next
                Else
                    memberList.Add(adPath)
                End If

                Return memberList
            Catch ex As Exception
                emailSubject = "APPLICATION ERROR: SPARES MANAGEMENT : createMemberList"
                emailBody = ex.Message
                emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
            End Try
        End Function
        Public Function mergeArrayListMembers(ByVal aList1 As ArrayList, ByVal aList2 As ArrayList) As ArrayList
            Dim aList As New ArrayList(aList1.Count + aList2.Count)

            aList.AddRange(aList1)

            Dim o As Object
            For Each o In aList2
                If Not aList1.Contains(o) Then
                    aList.Add(o)
                End If
            Next

            aList.TrimToSize()
            Return aList
        End Function
        Public Function getADGroupMembers(ByVal groupName As String) As ArrayList
            Dim result As SearchResult
            Dim search As DirectorySearcher = New DirectorySearcher()
            Dim userNames As ArrayList
            Dim counter As Integer
            Dim user As String
            Dim numResults As Integer
            Dim lineNum As Integer = 0

            Try
                lineNum += 1
                search.Filter = String.Format("(cn={0})", groupName)
                lineNum += 1
                search.PropertiesToLoad.Add("member")
                lineNum += 1
                result = search.FindOne()
                lineNum += 1

                userNames = New ArrayList()

                lineNum += 1
                If Not IsDBNull(result) Then

                    numResults = result.Properties("member").Count

                    For counter = 0 To numResults - 1
                        user = result.Properties("member")(counter).ToString
                        userNames.Add(user)
                    Next counter

                End If

                lineNum += 1

                Return (userNames)
            Catch ex As Exception
                emailSubject = "APPLICATION ERROR: SPARES MANAGEMENT : getADGroupMembers at lineNum : " & lineNum
                emailBody = ex.Message
                emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
            End Try

        End Function
        Public Function getADUserPath(ByVal userName As String) As String
            Dim result As SearchResult
            Dim search As DirectorySearcher = New DirectorySearcher()
            Dim userPath As String

            Try
                search.Filter = "samaccountname=" + userName
                result = search.FindOne()

                If Not IsDBNull(result) Then
                    userPath = result.Path
                    Return (userPath)
                End If

            Catch ex As System.Exception
                Return Nothing
            End Try
            Return Nothing
        End Function
        Public Function getUserInfo(ByVal userPath As String) As adUser
            Dim dirEntry As DirectoryEntry
            Dim adPath As String
            Dim propColl As System.DirectoryServices.PropertyCollection
            Dim userInfo As New adUser

            'Try
            '    adPath = "GC://" & userPath
            '    dirEntry = GetDirectoryEntry(adPath)
            '    propColl = dirEntry.Properties
            '    userInfo.userName = propColl("sAMAccountName").Value
            '    userInfo.givenName = propColl("name").Value
            '    userInfo.emailAddress = propColl("mail").Value
            '    userInfo.passwordNoExpire = getPwdExpires(userInfo.givenName)
            '    userInfo.pwdLastSet = getPwdSetDate(userInfo.givenName)
            'Catch ex As Exception
            '    emailSubject = "APPLICATION ERROR: getUserInfo Method"
            '    emailBody = ex.Message
            '    emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
            'End Try

            Try

                adPath = "GC://" & userPath
                dirEntry = New DirectoryEntry(adPath, "avon\cadadmin", "lz396gd417")
                propColl = dirEntry.Properties
                userInfo.userName = propColl("sAMAccountName").Value
                userInfo.givenName = propColl("name").Value
                userInfo.emailAddress = propColl("mail").Value
                userInfo.passwordNoExpire = getPwdExpires(userInfo.givenName)
                userInfo.pwdLastSet = getPwdSetDate(userInfo.givenName)
            Catch ex As System.Exception
                userInfo.userName = "ERROR"
                emailSubject = "APPLICATION ERROR: SPARES SCHEDULING - getUserInfo Method"
                emailBody = ex.Message
                emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
            End Try

            Return userInfo
        End Function
        Public Function getUserInfoByUserName(ByVal userName As String) As adUser
            Dim aUser As New adUser
            Dim path As String
            Dim myValue As ResultPropertyValueCollection
            Dim retEntry As DirectoryEntry = New DirectoryEntry()
            Dim objResult As SearchResult
            Dim dirEntry As DirectoryEntry = New DirectoryEntry("GC://dc=avon,dc=local")
            Dim search As DirectorySearcher = New DirectorySearcher(dirEntry)

            Try
                search.Filter = "(&(objectCategory=User)" + "(samAccountName=" + userName + "))"
                objResult = search.FindOne()

                If objResult Is System.DBNull.Value Then
                    Return Nothing
                End If

                retEntry = objResult.GetDirectoryEntry()
                path = retEntry.Path

                Dim myResultPropColl As ResultPropertyCollection
                myResultPropColl = objResult.Properties

                ' CODE TO GO THROUGH EACH PROPERTY
                'Dim myKey As String
                'For Each myKey In myResultPropColl.PropertyNames
                '    Dim myCollection As Object
                '    For Each myCollection In myResultPropColl(myKey)
                '        myValue = myResultPropColl("mail")
                '    Next myCollection
                'Next myKey

                myValue = myResultPropColl("displayname")
                If myValue.Count <> 0 Then
                    aUser.givenName = myValue.Item(0)
                End If
                aUser.userName = userName
                myValue = myResultPropColl("mail")
                If myValue.Count <> 0 Then
                    aUser.emailAddress = myValue.Item(0)
                End If

            Catch ex As System.Exception
                emailSubject = "APPLICATION ERROR: getUserInfoByUserName Method: " & userName
                emailBody = ex.Message
                emailMessage.sendMessage("scott.jesweak@avon-rubber.com", Nothing, emailSubject, emailBody)
                Return Nothing
            End Try

            Return aUser
        End Function
        Public Function getObjectType(ByVal adPath As String) As String
            Dim objType As String = ""
            Dim myValue As ResultPropertyValueCollection
            Dim retEntry As DirectoryEntry = New DirectoryEntry()
            Dim objResult As SearchResult
            Dim path As String
            Dim dirEntry As DirectoryEntry = New DirectoryEntry("GC://" & adPath)
            Dim search As DirectorySearcher = New DirectorySearcher(dirEntry)

            objResult = search.FindOne()

            retEntry = objResult.GetDirectoryEntry()
            path = retEntry.Path

            Dim myResultPropColl As ResultPropertyCollection
            myResultPropColl = objResult.Properties

            myValue = myResultPropColl("objectClass")
            ''CODE TO GO THROUGH EACH PROPERTY
            'Dim myKey As String
            'For Each myKey In myResultPropColl.PropertyNames
            '    Dim myCollection As Object
            '    For Each myCollection In myResultPropColl(myKey)
            '        'myValue = myResultPropColl("mail")
            '    Next myCollection
            'Next myKey

            objType = myValue.Item(1).ToString

            Return objType
        End Function
        Private Function getPwdSetDate(ByVal givenName As String) As Date
            Dim search As New DirectorySearcher
            Dim result As SearchResult
            Dim lastSet As Long
            Dim dateLastSet As Date

            search.Filter = String.Format("(cn={0})", givenName)
            search.PropertiesToLoad.Add("pwdLastSet")
            search.PropertiesToLoad.Add("userAccountControl")
            result = search.FindOne()

            If Not IsDBNull(result) Then
                lastSet = result.Properties("pwdLastSet")(0)
                dateLastSet = Date.FromFileTimeUtc(lastSet)
                Return dateLastSet
            End If
        End Function
        Private Function getPwdExpires(ByVal givenName As String) As Long
            Dim search As New DirectorySearcher
            Dim result As SearchResult
            Dim passExpires As Long

            search.Filter = String.Format("(cn={0})", givenName)
            search.PropertiesToLoad.Add("userAccountControl")
            result = search.FindOne()

            If Not IsDBNull(result) Then
                ' userAccountControl value of 66048 - Password Does NOT Expire
                passExpires = result.Properties("userAccountControl")(0)
                Return passExpires
            End If
        End Function
        Public Function getMaxPwdAge() As Integer
            Dim search As New DirectorySearcher
            Dim dirEntry As DirectoryEntry
            Dim result As SearchResult
            Dim maxAge As System.TimeSpan

            dirEntry = New DirectoryEntry("LDAP://172.24.16.29/DC=avon;DC=local", Nothing, Nothing, AuthenticationTypes.Secure)
            '        dirEntry = GetDirectoryEntry("LDAP://172.24.16.29/DC=avon;DC=local")
            search = New DirectorySearcher(dirEntry, "(objectClass=*)", Nothing, SearchScope.Base)

            result = search.FindOne()
            maxAge = System.TimeSpan.MinValue

            If (result.Properties.Contains("maxPwdAge")) Then
                maxAge = System.TimeSpan.FromTicks(CLng(result.Properties("maxPwdAge")(0)))
                Return (System.Math.Abs(maxAge.Days))
            End If

            Return System.Math.Abs(System.Math.Ceiling(maxAge.Days))
        End Function
        Public Function removeDomainFromUserName(ByVal currentUser As String) As String
            Dim currentUserLgth As Integer
            Dim userStart As Integer

            currentUserLgth = currentUser.Length
            userStart = currentUser.IndexOf("\") + 1

            currentUser = currentUser.Substring(userStart, (currentUserLgth - currentUser.IndexOf("\")) - 1)
            Return currentUser
        End Function
    End Class
    Public Class adUser
        Public givenName As String
        Public userName As String
        Public pwdLastSet As Date
        Public emailAddress As String
        Public passwordNoExpire As Long
    End Class
End Namespace
