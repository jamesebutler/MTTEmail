'Imports Oracle.DataAccess.Client
Imports System.Configuration
Imports System.Net.Mail

Module Main

    Dim tracing As Boolean = System.Configuration.ConfigurationManager.AppSettings("Tracing")


    Public Function CallDROraclePackage(ByRef parms As OracleParameterCollection, ByVal packageName As String) As String 'OracleClient.OracleDataReader
        Dim conCust As OracleConnection = Nothing
        Dim cmdSQL As OracleCommand = Nothing
        Dim connection As String = String.Empty
        Dim provider As String = String.Empty
        Dim dr As OracleDataReader = Nothing
        Dim cnConnection As OracleConnection = Nothing
        Dim returnParamName As String = String.Empty
        Dim returnValue As String = String.Empty
        Dim returnParms As New Collection
        Try

            'check to see if the CountExceptions is 30 or above.
            'do not want to keep sending out errors emails.
            'JEB 2/7/2019
            If CountExceptions >= 30 Then
                End
            End If

            If strDB = "RIDEV" Then
                connection = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connection = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If

            cmdSQL = New OracleCommand

            With cmdSQL
                cnConnection = New OracleConnection(connection)
                cnConnection.Open()
                .Connection = cnConnection
                .CommandText = packageName
                .CommandType = CommandType.StoredProcedure
                Dim sb As New System.Text.StringBuilder
                For i As Integer = 0 To parms.Count - 1
                    If parms.Item(i).Value Is Nothing Then parms.Item(i).Value = DBNull.Value
                    Dim parm As New OracleParameter
                    parm.Direction = parms.Item(i).Direction
                    parm.DbType = parms.Item(i).DbType
                    parm.OracleDbType = parms.Item(i).DbType
                    parm.Size = parms.Item(i).Size
                    If parms.Item(i).Direction = ParameterDirection.Input Or parms.Item(i).Direction = ParameterDirection.InputOutput Then
                        If parms.Item(i).Value IsNot Nothing Then
                            parm.Value = parms.Item(i).Value
                            If parm.Value.ToString = "" Then
                                parm.IsNullable = True
                                parm.Value = System.DBNull.Value
                            End If
                        Else
                            If parm.OracleDbType = OracleDbType.NVarChar Then
                                'parm.Value = DBNull.Value
                                'parm.Size = 2
                            End If
                        End If
                    ElseIf parms.Item(i).Direction = ParameterDirection.Output Then
                        returnParms.Add(parms.Item(i).ParameterName)
                        returnParamName = parms.Item(i).ParameterName
                    End If
                    parm.ParameterName = parms.Item(i).ParameterName
                    .Parameters.Add(parm)
                    If sb.Length > 0 Then sb.Append(",")
                    If parm.OracleDbType = OracleDbType.VarChar Then
                        If parm.Value IsNot Nothing Then
                            sb.Append(parm.ParameterName & "= '" & parm.Value.ToString & "' Type=" & parm.OracleDbType.ToString)
                        Else
                            sb.Append(parm.ParameterName & "= '" & "Null" & "' Type=" & parm.OracleDbType.ToString)
                        End If
                    Else
                        If parm.Value IsNot Nothing Then
                            sb.Append(parm.ParameterName & "= '" & parm.Value.ToString & "' Type=" & parm.OracleDbType.ToString)
                        Else
                            sb.Append(parm.ParameterName)
                        End If
                    End If
                    sb.AppendLine()
                Next
            End With

            cmdSQL.ExecuteNonQuery()

            'Populate the original parms collection with the data from the output parameters
            For i As Integer = 0 To returnParms.Count - 1
                parms.Item(cmdSQL.Parameters(returnParms.Item(i)).ToString).Value = cmdSQL.Parameters(returnParms.Item(i)).Value.ToString
            Next
            '// return the return value if there is one
            If returnParamName.Length > 0 Then
                returnValue = cmdSQL.Parameters(returnParamName).Value.ToString
            Else
                returnValue = CStr(0)
            End If

        Catch ex As Exception
            ''Trace("MTTEmail", "CallDROraclePackage:Error " & ex.Message, tracing)
            If returnValue.Length = 0 Then returnValue = "Error Occurred"
            If Not conCust Is Nothing Then conCust = Nothing
            HandleError("CallDROraclePackage", ex.Message, ex)
        Finally
            CallDROraclePackage = returnValue
            If Not dr Is Nothing Then dr = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
            If cnConnection IsNot Nothing Then
                If cnConnection.State = ConnectionState.Open Then cnConnection.Close()
                cnConnection = Nothing
            End If
        End Try
    End Function

    Public Sub HandleError(Optional ByVal MethodName As String = "MTTEmail", Optional ByVal additionalErrMsg As String = "", Optional ByVal excep As Exception = Nothing)
        Dim le As Exception = Nothing
        Dim errorMessage As New System.Text.StringBuilder
        Dim errorCount As Integer = 0
        Dim errMsg As String = String.Empty
        Dim chunkLength As Integer = 0
        Dim maxLen As Integer = 3500

        'check to see if the CountExceptions is 30 or above.
        'do not want to keep sending out errors emails.
        'JEB 2/7/2019
        If CountExceptions >= 30 Then
            End
        End If


        Trace("MTTEmail", "HandleError", tracing)
        Try
            If excep IsNot Nothing Then
                le = excep
            End If

            If le IsNot Nothing Then
                'send email
                SendEmail(developmentEmail, ManufacturingEmail, "MTTEmail Error - " + MethodName, le.ToString)

                Do While le IsNot Nothing
                    errorCount = errorCount + 1
                    'errorMessage.Length = 0
                    errorMessage.Append("<Table width=100% border=1 cellpadding=2 cellspacing=2 bgcolor='#cccccc'>")
                    errorMessage.Append("<tr><th colspan=2><h2>Error</h2></th>")
                    errorMessage.Append("<tr><td><b>Program:</b></td><td>{0}</td></tr>")
                    errorMessage.Append("<tr><td><b>Exception #</b></td><td>{1}</td></tr>")
                    errorMessage.Append("<tr><td><b>Time:</b></td><td>{2}</td></tr>")
                    errorMessage.Append("<tr><td><b>Details:</b></td><td>{3}</td></tr>")
                    errorMessage.Append("<tr><td><b>Additional Info:</b></td><td>{4}</td></tr>")
                    errorMessage.Append("</table>")
                    errMsg = errorMessage.ToString
                    errMsg = String.Format(errMsg, My.Application.Info.AssemblyName, errorCount, FormatDateTime(Now, DateFormat.LongDate), le.ToString, additionalErrMsg)
                    additionalErrMsg = ""
                    le = le.InnerException
                    'MsgBox(errMsg)
                    errorMessage.Length = 0

                    For i As Integer = 0 To errMsg.Length Step maxLen
                        If errMsg.Length < maxLen Then
                            chunkLength = errMsg.Length - 1
                        Else
                            If errMsg.Length - i < maxLen Then
                                chunkLength = errMsg.Length - i
                            Else
                                chunkLength = maxLen
                            End If
                        End If

                        Dim errValue As String = errMsg.Substring(i, chunkLength)

                        'insert record into audit
                        InsertAuditRecord(MethodName, errValue)

                        System.Threading.Thread.Sleep(1000) ' Sleep for 1 second
                    Next
                Loop


            End If

        Catch ex As Exception
        Finally
            le = Nothing
            Try
                ' HttpContext.Current.Server.ClearError()
                'HttpContext.Current.Response.Redirect(redirectURL, False)
            Catch e As Exception
                'HttpContext.Current.Server.ClearError()
            End Try
        End Try
    End Sub

    Sub InsertAuditRecord(ByVal sourceName As String, ByVal errorMessage As String)
        Dim paramCollection As New OracleParameterCollection
        Dim param As New OracleParameter
        Dim ds As System.Data.DataSet = Nothing

        'check to see if the CountExceptions is 30 or above.
        'do not want to keep sending out errors emails.
        'JEB 2/7/2019
        If CountExceptions >= 30 Then
            End
        End If

        Try

            param = New OracleParameter
            param.ParameterName = "in_name"
            param.OracleDbType = OracleDbType.VarChar
            param.Direction = Data.ParameterDirection.Input
            param.Value = sourceName
            paramCollection.Add(param)

            param = New OracleParameter
            param.ParameterName = "in_desc"
            param.OracleDbType = OracleDbType.VarChar
            param.Direction = Data.ParameterDirection.Input
            param.Value = Mid(errorMessage, 1, 200)
            paramCollection.Add(param)

            Dim returnStatus As String = CallDROraclePackage(paramCollection, "Reladmin.RIAUDIT.InsertErrorRecord")
        Catch ex As Exception
            SendEmail(developmentEmail, ManufacturingEmail, "MTTEmail Error - InsertAuditRecord", ex.InnerException.ToString)

        Finally
            param = Nothing
            paramCollection = Nothing
        End Try
    End Sub
    Public Function cleanString(ByVal strEdit As String, ByVal defaultValue As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(strEdit, "[\n]", defaultValue).Trim
    End Function
    Sub SendEmail(ByVal toaddress As String, ByVal fromAddress As String, ByVal subject As String, ByVal body As String, Optional ByVal displayName As String = "Manufacturing Task", Optional ByVal carbonCopy As String = "", Optional ByVal blindCarbonCopy As String = "", Optional ByVal IsBodyHtml As Boolean = True)
        Dim mail As System.Net.Mail.MailMessage = New MailMessage '= New MailMessage(New MailAddress(fromAddress, displayName), New MailAddress(toaddress))

        Dim OkToSend As Boolean = False
        Dim inputAddress As New System.Text.StringBuilder


        Trace("MTTEmail", "SendEmail", tracing)
        'MsgBox(toaddress)
        Try
            'Comment following line after test runs
            'subject = subject & toaddress

            'hold the email addresses that is sent  - JEB 1/4/2019
            strEmailAddressSentTo = strEmailAddressSentTo + "|" + toaddress


            If strDefaultEmail <> "" Then
                toaddress = strDefaultEmail
                subject = subject & " : " & toaddress
            End If
            'toaddress = "amy.albrinck@ipaper.com"



            inputAddress.Append("<p>ToAddress:")
            inputAddress.Append(toaddress)
            inputAddress.Append("<br>")
            inputAddress.Append("fromAddress:")
            inputAddress.Append(fromAddress)
            inputAddress.Append("<br>")
            inputAddress.Append("carbonCopy:")
            inputAddress.Append(carbonCopy)
            inputAddress.Append("<br>")
            inputAddress.Append("blindCarbonCopy:")
            inputAddress.Append(blindCarbonCopy)
            inputAddress.Append("</p>")

            If toaddress.Length > 0 Then
                Dim toEmail As String() = Split(toaddress, ",")
                For i As Integer = 0 To toEmail.Length - 1
                    If toEmail(i).Length > 0 Then 'And isEmail(toEmail(i)) Then
                        mail.To.Add(toEmail(i))
                    End If
                Next
                If mail.To.Count > 0 Then OkToSend = True
            End If

            'carbonCopy = "cathy.cox@ipaper.com,amy.albrinck@ipaper.com"
            If carbonCopy.Length > 0 Then
                Dim copyEmail As String() = Split(carbonCopy, ",")
                For i As Integer = 0 To copyEmail.Length - 1
                    If copyEmail(i).Length > 0 Then 'And isEmail(copyEmail(i)) Then
                        mail.CC.Add(copyEmail(i))
                    End If
                Next
                If mail.CC.Count > 0 Then OkToSend = True
            End If

            If strEmailBCC <> "" Then
                blindCarbonCopy = strEmailBCC
            End If
            'blindCarbonCopy = "amy.albrinck@ipaper.com"
            If blindCarbonCopy.Length > 0 Then
                Dim bccEmail As String() = Split(blindCarbonCopy, ",")
                For i As Integer = 0 To bccEmail.Length - 1
                    If bccEmail(i).Length > 0 Then 'And isEmail(bccEmail(i)) Then
                        mail.Bcc.Add(bccEmail(i))
                    End If
                Next
                If mail.Bcc.Count > 0 Then OkToSend = True
            End If

            If fromAddress.Trim.Length > 0 Then ' And isEmail(fromAddress) Then
                mail.From = New MailAddress(fromAddress, displayName)
            Else
                mail.From = New MailAddress(ManufacturingEmail, "Manufacturing Task")
            End If
            mail.Priority = MailPriority.High
            mail.IsBodyHtml = IsBodyHtml

            'Send the email message
            mail.Subject = subject
            mail.Body = body

            If OkToSend = True Then
                'InsertAuditRecord("MTTEmail", "The following email has been sent -  " & Mid(body, 1, 3000) & "<br> Recipients:" & inputAddress.ToString)
                Dim emailTryCount As Integer = 0
                Dim emailSuccess As Boolean = False
                Do While emailTryCount < 5 And emailSuccess = False
                    Dim client As SmtpClient = New SmtpClient()
                    Try
                        With client
                            emailTryCount += 1
                            .Host = "gpimail.na.graphicpkg.pri"
                            .Timeout = 1000000
                            .Send(mail)
                            emailSuccess = True
                        End With
                        client.Dispose()
                        'InsertAuditRecord("MTTEmail", "The following email has been sent -  " & Mid(body, 1, 1800) & "<br> Recipients:" & inputAddress.ToString)

                    Catch ex As SmtpException
                        System.Threading.Thread.Sleep(1000)
                        'InsertAuditRecord("Send Email Error", "The following email was not sent - Retry(" & emailTryCount & ") -  " & Mid(body, 1, 3000) & "<br> Recipients:" & inputAddress.ToString)
                    Finally
                        client = Nothing
                        'mail.Dispose()
                        'InsertAuditRecord("MTTEmail", "Setting client to Nothing")
                    End Try
                Loop
                'Else
                '    InsertAuditRecord("Send Email", "This attempted email message was not sent b/c of a missing recipient.  " & Mid(body, 1, 3000) & inputAddress.ToString)
                '    Dim additionalErrorMsg As String = String.Format("<b>" & "An error occurred while trying to send an email to the following email address:" & "[{0}].<br>" & "Please forward the below information to the person assigned the task and contact MTT administrator to correct their email address." & "</b>", inputAddress.ToString)
                '    mail.Body = additionalErrorMsg & "<br><br>" & mail.Body
                '    mail.To.Add(mail.From)
                '    mail.Subject = "Email Error - " & mail.Subject
                '    Dim emailTryCount As Integer = 0
                '    Dim emailSuccess As Boolean = False
                '    Do While emailTryCount < 5 And emailSuccess = False
                '        Dim client As SmtpClient = New SmtpClient()
                '        Try
                '            With client
                '                emailTryCount += 1
                '                .Timeout = 1000000
                '                .Send(mail)
                '                InsertAuditRecord("Send Email", "The following email has been sent to the sender because the To Address was invalid -  " & Mid(body, 1, 3000) & "<br> Recipients:" & inputAddress.ToString)
                '                emailSuccess = True
                '            End With
                '        Catch ex As SmtpException
                '            System.Threading.Thread.Sleep(1000)
                '        Finally
                '            client = Nothing
                '        End Try

                '    Loop


            End If

        Catch ex As SmtpException
            CountExceptionsErrors()
            'Trace("MTTEmail", "SendEmail:Error SmtpException: " & ex.Message, tracing)
            'HandleError("Send Email", "This attempted email message was not sent b/c :" & ex.Message & "<br>" & body & inputAddress.ToString, ex)
        Catch ex As Exception
            CountExceptionsErrors()
            'Trace("MTTEmail", "SendEmail:Error " & ex.Message, tracing)
            'HandleError("Send Email", "This attempted email message was not sent b/c :" & ex.Message & "<br>" & body & inputAddress.ToString, ex)
        Finally
            'check to see if the CountExceptions is 30 or above.
            'do not want to keep sending out errors emails.
            'JEB 2/7/2019
            If CountExceptions >= 30 Then
                End
            End If
            If mail IsNot Nothing Then mail = Nothing
        End Try
    End Sub

    Public Function GetLocalizedDateTime(ByVal dateTime As String, ByVal locale As String, format As String) As String
        Trace("MTTEmail", "GetLocalizedDateTime", tracing)
        Dim returnValue As String
        Dim cI As System.Globalization.CultureInfo
        cI = System.Globalization.CultureInfo.GetCultureInfo(locale)
        Dim dt As DateTime
        dt = dateTime
        returnValue = dt.ToString(format, cI)
        Return returnValue
    End Function

    Public Sub Trace(ByVal sourceName As String, ByVal traceMessage As String, ByVal log As Boolean)

        Try
            If log Then
                InsertAuditRecord(sourceName + " TRACE", DateTime.Now.ToString + " " + traceMessage)
            End If

        Catch ex As Exception
            CountExceptionsErrors()
        Finally

        End Try
    End Sub


    Public Sub CountExceptionsErrors()

        CountExceptions = CountExceptions + 1

    End Sub



End Module
