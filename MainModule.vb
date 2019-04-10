Imports System.Runtime.CompilerServices

Module MainModule
    'ALA 2/9/2015 - File attachments are not showing correct url link so they 
    ' are not opening when users click on the link in emails.

    Public strDefaultEmail, strDB, strEmailBCC, strSP, strTaskID, strPlantCode As String
    Public strRole As String = ""
    Public strUrl As String = "http://gpitasktracker.graphicpkg.com"
    Public strFileURL As String = ""
    Public dtRunDate As Date
    Public intUniqueEmailSent As Integer
    Public strEmailAddressSentTo As String = String.Empty

    'JEB added 12/12/2018  
    Public failureEmail As String = System.Configuration.ConfigurationManager.AppSettings("failureEmail")
    Public developmentEmail As String = System.Configuration.ConfigurationManager.AppSettings("developmentEmail")
    Public ManufacturingEmail As String = System.Configuration.ConfigurationManager.AppSettings("ManufacturingEmail")
    Dim tracing As Boolean = System.Configuration.ConfigurationManager.AppSettings("Tracing")
    Public CountExceptions As Int16 = 0

    'JEB ended
    Sub Main()

        Dim cmdLineParams() As String 'Get Parameters

        'First Parameter = RunDate
        'Second Parameter = Role
        'Third Parameter = default EMAIL - if you want to run and send all emails to same email account
        'Forth Parameter = BCC Email - if you blind cc anyone to check program
        'Fifth Parameter = Database - will default to prod unless RIDEV entered.
        'Sixth Parameter = 
        'Seventh Parameter = Plant Code - will only run for users at PlantCode entered.

        Dim i As Integer

        Try

            Dim st As StackTrace = New StackTrace()

            Trace("MTTEmail", "Main - Start", tracing)

            Dim args As String = String.Empty
            'dim temp as Environment.GetCommandLineArgs()

            '=================================================
            'Used for testing  JEB 12/12/2018
            'Dim testit As String = String.Empty
            'testit = ",JEB,james.butler@graphicpkg.com,james.butler@graphicpkg.com,,"
            'testit = "2/16/2019,RESPONSIBLE,,james.butler@graphicpkg.com,,"
            'testit = "2/16/2019,CREATOR,,james.butler@graphicpkg.com,,"
            'testit = ",BUSUNITMGR,,james.butler@graphicpkg.com,,"

            'cmdLineParams = Split(testit, ",")
            '=================================================

            If Command() = "" Then
                SendEmail(developmentEmail, ManufacturingEmail, strRole + " - No parameters were entered.", strEmailAddressSentTo,)   'JEB 12/12/2018 added
                End
            Else
                cmdLineParams = Split(Command(), ",")
                For i = 0 To UBound(cmdLineParams)
                    If i = 0 Then
                        If cmdLineParams(i) Is DBNull.Value Or cmdLineParams(i) = "" Then
                            dtRunDate = Today()
                        Else
                            dtRunDate = cmdLineParams(i)
                        End If

                    ElseIf i = 1 Then
                        strRole = cmdLineParams(i)
                    ElseIf i = 2 Then
                        strDefaultEmail = cmdLineParams(i)
                    ElseIf i = 3 Then
                        strEmailBCC = cmdLineParams(i)
                    ElseIf i = 4 Then
                        strDB = cmdLineParams(i)
                        If strDB = "" Then
                            strDB = "RI"
                            strFileURL = "//GPIAZRELFPRD01/uploads_ri/"
                            'strFileURL = "//S02ARELPRD01/MEAS/PRODUCTION/MTT/UPLOADS/"
                            'strDB = "RITEST"
                        Else
                            strDB = "RIDEV"
                            strFileURL = "//S02ARELPRD01/MEAS/DEVELOPMENT/MTT/UPLOADS/"
                        End If
                    ElseIf i = 5 Then
                        strTaskID = cmdLineParams(i)
                        If strTaskID <> "" Then
                            GetIndTask(strTaskID)
                        End If
                    ElseIf i = 6 Then
                        strPlantCode = cmdLineParams(i)
                    End If
                    'i = i + 1
                Next

                If strPlantCode = "" Then
                    strPlantCode = "ALL"
                End If
                If strTaskID = "" Then
                    Trace("MTTEmail", "Main - Role:" + strRole, tracing)
                    SendEmail(developmentEmail, ManufacturingEmail, strRole + " - START - MTTEmail", strEmailAddressSentTo,)   'JEB 12/12/2018 added
                    GetProfile()
                End If
                Trace("MTTEmail", "Main - End:" + strRole, tracing)
                SendEmail(developmentEmail, ManufacturingEmail, strRole + " - SUCCESS - MTTEmail", strEmailAddressSentTo,)   'JEB 12/12/2018 added

            End If


        Catch ex As Exception
            CountExceptionsErrors()
            Trace("MTTEmail", "Main - Role:" + strRole + " ERROR: " + ex.InnerException.ToString, tracing)
            SendEmail(developmentEmail, ManufacturingEmail, strRole + " - ERROR - MTTEmail", ex.InnerException.ToString,)   'JEB 12/12/2018 added
            HandleError("MTTEmail.MAIN", ex.Message, ex)

        Finally
            End
        End Try

    End Sub

    Sub GetProfile()
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim dr As DataRow
        Dim drEmail As DataRow
        Dim dsProfileDefault As DataSet

        Dim strErr As String
        Dim strMsg, strBody, strLanguage, strUserid, strEmailAddress As String
        Dim previous_recType, previous_Email As String
        Dim param As New OracleParameter
        Dim intProfileCnt, intNonDefaultProfileCnt, intEmailCnt, intEmailSent As Integer
        Dim sbEmailBody As New System.Text.StringBuilder
        Trace("MTTEmail", "GetProfile:" + strRole, tracing)
        Try
            'Set up db connection
            strSP = My.Application.Info.AssemblyName

            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            'Set up initial procedure which returns the profile for emails that should be sent based on run date passed in.
            'This package checks the vw_mtt_notify_profile view.
            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MTTBATCHEMAILS.EMAILSCHED"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "in_role"
                param.OracleDbType = OracleDbType.VarChar
                param.Direction = Data.ParameterDirection.Input
                param.Value = strRole
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "in_date"
                param.OracleDbType = OracleDbType.Date
                param.Direction = Data.ParameterDirection.Input
                param.Value = dtRunDate
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsProfileDefault"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsProfileOther"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            Dim daProfile = New OracleDataAdapter(cmdSQL)
            dsProfileDefault = New DataSet()
            daProfile.Fill(dsProfileDefault)

            Dim strEmailType, strDateRange As String

            'Get Profile counts so we can determine if all emails were sent
            'The first recordset retrieves the default profile.  The second recordset retrieves any  unique profiles
            'that have been setup by support or by users.
            intProfileCnt = dsProfileDefault.Tables(0).Rows.Count()
            intNonDefaultProfileCnt = dsProfileDefault.Tables(1).Rows.Count()

            If intProfileCnt = 0 Then
                InsertAuditRecord(strSP, "NO DEFAULT " & strRole & " profile defined for " & dtRunDate.ToLongDateString)
            Else
                InsertAuditRecord(strSP, "DEFAULT " & strRole & " profile defined for " & dtRunDate.ToLongDateString)
            End If

            If intNonDefaultProfileCnt = 0 Then
                InsertAuditRecord(strSP, "NO unique " & strRole & " profiles defined for " & dtRunDate.ToLongDateString)
            End If

            connDB.Close()
            Trace("MTTEmail", "GetProfile-LOOP:" + strRole, tracing)
            'Loop thru records for the DEFAULT Profile
            For Each dr In dsProfileDefault.Tables(0).Rows

                strEmailType = dr("EmailType")
                strDateRange = dr("daterange")

                If strEmailType = "FUTURE" Or strEmailType = "ENTERED" Then
                    Try
                        connDB.Open()
                        cmdSQL = New OracleCommand
                        With cmdSQL
                            .Connection = connDB
                            .CommandText = "MTTBATCHEMAILS.emaillisting"
                            .CommandType = CommandType.StoredProcedure

                            param = New OracleParameter
                            param.ParameterName = "in_role"
                            param.OracleDbType = OracleDbType.VarChar
                            param.Direction = Data.ParameterDirection.Input
                            param.Value = strRole
                            .Parameters.Add(param)

                            param = New OracleParameter
                            param.ParameterName = "in_date"
                            param.OracleDbType = OracleDbType.Date
                            param.Direction = Data.ParameterDirection.Input
                            param.Value = dtRunDate
                            .Parameters.Add(param)

                            param = New OracleParameter
                            param.ParameterName = "in_daterange"
                            param.OracleDbType = OracleDbType.VarChar
                            param.Direction = Data.ParameterDirection.Input
                            param.Value = strDateRange
                            .Parameters.Add(param)

                            param = New OracleParameter
                            param.ParameterName = "in_plantcode"
                            param.OracleDbType = OracleDbType.VarChar
                            param.Direction = Data.ParameterDirection.Input
                            param.Value = strPlantCode
                            .Parameters.Add(param)

                            param = New OracleParameter
                            param.ParameterName = "rsUserIds"
                            param.OracleDbType = OracleDbType.Cursor
                            param.Direction = ParameterDirection.Output
                            .Parameters.Add(param)

                        End With

                        Dim daEmail = New OracleDataAdapter(cmdSQL)
                        Dim dsEmail As DataSet
                        dsEmail = New DataSet()
                        daEmail.Fill(dsEmail)

                        connDB.Close()

                        'Get email count so we can determine if all emails were sent
                        intEmailCnt = dsEmail.Tables(0).Rows.Count()
                        intEmailSent = 0

                        strMsg = ""
                        previous_Email = ""

                        'Loop thru all email records 
                        For Each drEmail In dsEmail.Tables(0).Rows

                            previous_recType = ""
                            strUserid = drEmail("username")
                            strEmailAddress = drEmail("email")
                            strLanguage = drEmail("language")

                            If strEmailType = "FUTURE" Then
                                'Get all tasks and email appropriate user.
                                GetTasks(strDateRange, strUserid, strLanguage, strEmailAddress)

                                intEmailSent = intEmailSent + 1

                                strBody = ""
                            Else
                                GetEnteredTasks(strUserid, strLanguage, strEmailAddress)

                                intEmailSent = intEmailSent + 1
                            End If

                        Next
                        drEmail = Nothing
                        dsEmail = Nothing

                        'If counts indicate not all emails were sent, send email to support and write record to audit table.
                        If intEmailSent <> intEmailCnt Then
                            HandleError(strSP, "Only " & intEmailSent & " emails were sent.  " & intEmailCnt & " Emails should have been sent")
                        Else
                            InsertAuditRecord(strSP, intEmailSent & " of " & intEmailCnt & " emails were sent.")
                        End If

                    Catch ex As Exception
                        'Exception handling
                        strErr = "Error occurred." & ex.Message
                        HandleError(strSP, strErr, ex)
                    End Try
                    connDB.Close()
                End If

            Next

            'intUniqueEmailSent = 0

            'Loop thru records for profile records for individual users.
            For Each dr In dsProfileDefault.Tables(1).Rows
                strUserid = dr("Username")
                strDateRange = dr("daterange")
                strLanguage = dr("Language")
                strEmailAddress = dr("Email")

                'InsertAuditRecord(strSP, "Unique " & strRole & " " & strDateRange & " profile defined for " & strUserid & " to run on " & dtRunDate.ToLongDateString)

                'Get all tasks and email appropriate user.
                If strDateRange = "DAILY" Then
                    GetEnteredTasks(strUserid, strLanguage, strEmailAddress)
                Else
                    GetTasks(strDateRange, strUserid, strLanguage, strEmailAddress)
                End If
                ' intUniqueEmailSent = intUniqueEmailSent + 1

            Next
            'If counts indicate not all emails were sent, send email to support and write record to audit table.
            '*** we really can't monitor how many emails should have been sent versus what was sent because we are 
            ' looping through all the unique profiles.  Many of which won't get any emails because no tasks were
            ' added.
            'If intUniqueEmailSent <> intUniqueEmailCnt Then
            'SendEmail("amy.albrinck@ipaper.com", "Manufacturing.task@ipaper.com", "AutoEmailError", "Only " & intUniqueEmailSent & " emails were sent.  " & intUniqueEmailCnt & " Emails should have been sent.")
            'HandleError(strSP, "Only " & intUniqueEmailSent & " emails were sent.  " & intUniqueEmailCnt & " Emails should have been sent")
            'Else
            'InsertAuditRecord(strSP, intUniqueEmailSent & " of " & intUniqueEmailCnt & " emails were sent.")
            'End If

            dsProfileDefault = Nothing
            daProfile = Nothing
            connDB.Close()

        Catch ex As Exception

            CountExceptionsErrors()
            HandleError(strSP, ex.Message, ex)
        Finally
            connDB.Close()
            dsProfileDefault = Nothing
            'daProfile = Nothing
            If Not connDB Is Nothing Then connDB = Nothing
            If Not cmdSQL Is Nothing Then cmdSQL = Nothing
        End Try
    End Sub

    Sub GetTasks(ByVal strDateRange As String, ByVal strUserid As String, ByVal strLanguage As String, ByVal strEmailAddress As String)
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim drTasks As OracleDataReader
        'Dim drTaskHeaderDocs As OracleDataReader
        'Dim drTaskItemDocs As OracleDataReader
        Dim strTaskID, strTaskHeaderID, strTaskItemID As String
        Dim strDueDate, strResponsible, strTitle, strBusUnitMgr, strComments As String

        Dim strHeading As String = ""
        Dim strHeading1 As String = ""
        Dim strErr As String = ""
        Dim strSubject As String = ""
        Dim strHeaderTitle As String
        Dim strRootTaskID As String = ""
        Dim strTaskLink, strTaskLinkCount As String
        Dim strMsg, strSiteName As String
        Dim strSourceSystem As String
        Dim strSourceRef As String
        Dim strHeaderURL, strDetailURL As String
        Dim strBody As String = ""
        Dim strFooter As String = ""
        Dim previous_recType As String
        Dim strRecType As String = ""
        Dim intLeadTime As Integer
        Dim param As New OracleParameter
        Dim sbEmailBody As New System.Text.StringBuilder
        Dim strTasksFound As String = "N"

        Trace("MTTEmail", "GetTasks", tracing)

        Try
            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            previous_recType = ""
            strSP = My.Application.Info.AssemblyName

            'Dim IPLoc As New IP.MEASFramework.ExtensibleLocalizationAssembly.WebLocalization(strLanguage, "MTT")

            'Get all tasks associated with userid 
            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MTTBATCHEMAILS.TASKLISTING"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "in_date"
                param.OracleDbType = OracleDbType.Date
                param.Direction = Data.ParameterDirection.Input
                param.Value = dtRunDate
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "in_daterange"
                param.OracleDbType = OracleDbType.VarChar
                param.Direction = Data.ParameterDirection.Input
                param.Value = strDateRange
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "in_userid"
                param.OracleDbType = OracleDbType.VarChar
                param.Direction = Data.ParameterDirection.Input
                param.Value = strUserid
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "in_role"
                param.OracleDbType = OracleDbType.VarChar
                param.Direction = Data.ParameterDirection.Input
                param.Value = strRole
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsAllTasks"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With


            sbEmailBody = New System.Text.StringBuilder
            Dim v_td As String() = {"<TD>", "</TD>"}

            drTasks = cmdSQL.ExecuteReader()

            While drTasks.Read
                strTaskHeaderID = drTasks("taskheaderseqid")
                strTaskItemID = drTasks("taskitemseqid")

                If drTasks("roottaskitemseqid") Is DBNull.Value Then
                    '                    strTaskID = drTasks("taskitemseqid")
                    strRootTaskID = ""
                Else
                    strRootTaskID = drTasks("roottaskitemseqid")
                End If 'strHeaderLinkCount = drTasks("HEADERLINKCOUNT")

                'strHeaderLinkDesc = ""
                'strHeaderLinkLocation = ""
                'strHeaderLinkFile = ""
                'strHeaderLink = ""
                'If strHeaderLinkCount > 0 Then
                '    cmdSQL = New OracleClient.OracleCommand
                '    With cmdSQL
                '        .Connection = connDB
                '        .CommandText = "MTTBATCHEMAILS.TASKHEADERDOCS"
                '        .CommandType = CommandType.StoredProcedure

                '        param = New OracleParameter
                '        param.ParameterName = "in_taskheaderid"
                '        param.OracleType = OracleType.VarChar
                '        param.Direction = Data.ParameterDirection.Input
                '        param.Value = strTaskHeaderID
                '        .Parameters.Add(param)

                '        param = New OracleParameter
                '        param.ParameterName = "rsTaskHeaderDocs"
                '        param.OracleType = OracleType.Cursor
                '        param.Direction = ParameterDirection.Output
                '        .Parameters.Add(param)

                '    End With
                '    drTaskHeaderDocs = cmdSQL.ExecuteReader()
                '    While drTaskHeaderDocs.Read
                '        strHeaderLinkDesc = drTaskHeaderDocs("description")
                '        strHeaderLinkLocation = drTaskHeaderDocs("location")
                '        strHeaderLinkFile = drTaskHeaderDocs("filename")

                '        strHeaderLink = strHeaderLink & "<BR>" & "<A HREF='" & GetFileLocation(strHeaderLinkFile, strHeaderLinkLocation) & "'>" & strHeaderLinkDesc & "</A>"
                '    End While

                'End If


                strTaskLinkCount = drTasks("TASKLINKCOUNT")
                strTaskLink = ""
                If strTaskLinkCount > 0 And strRole = "RESPONSIBLE" Then
                    If strRootTaskID = "" Then
                        strTaskLink = GetTaskItemDocs(strTaskItemID)
                    Else
                        strTaskLink = GetTaskItemDocs(strRootTaskID)
                    End If
                End If

                strRecType = Trim(drTasks("RECTYPE"))
                strSiteName = Trim(drTasks("sitename"))

                If drTasks("RECTYPE") = "Overdue" Then
                    strRecType = "<font color=red>" & strRecType & " <i >- " & strSiteName & "</i></font>"
                Else
                    strRecType = strRecType & "<i> - " & strSiteName & "</i>"
                End If

                strDueDate = drTasks("Item_DueDate")
                strDueDate = GetLocalizedDateTime(strDueDate, strLanguage, "MM/dd/yyyy")
                'strDueDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strDueDate, strLanguage, "d")

                'strTitle = Trim(IPLoc.GetResourceValue(drTasks("Item_Title")))
                strTitle = Trim(drTasks("Item_Title"))

                If drTasks("Whole_Name_Responsible_Person") Is DBNull.Value Or drTasks("Whole_Name_Responsible_Person") = " " Then
                    strResponsible = drTasks("RoleDescription") & " (" & drTasks("Responsible_Role_Names") & ")"
                Else
                    strResponsible = drTasks("Whole_Name_Responsible_Person")
                End If
                strBusUnitMgr = drTasks("Mgr")
                'strTypeMgr = drTasks("TypeMgr")
                strTaskID = drTasks("taskitemseqid")
                strHeaderTitle = drTasks("taskheadertitle")
                'strActivity = IPLoc.GetResourceValue(drTasks("ActivityName"))
                If drTasks("Mttcomment") Is DBNull.Value Then
                    strComments = ""
                Else
                    strComments = Trim(drTasks("mttcomment"))
                End If
                strTaskHeaderID = drTasks("taskheaderseqid")
                strTaskItemID = drTasks("taskitemseqid")
                intLeadTime = drTasks("LEADTIME")
                If drTasks("sourcesystem") Is DBNull.Value Then
                    strSourceSystem = ""
                    strSourceRef = ""
                    'strHeaderURL = "<A HREF=HTTP://" & strDB & "/TaskTracker/TaskHeader.aspx?HeaderNumber=" & strTaskHeaderID & ">" & strHeaderTitle & " (" & strActivity & ")</A>"
                    strHeaderURL = strHeaderTitle
                    strDetailURL = "<A HREF=" & strUrl & "/TaskDetails.aspx?HeaderNumber=" & strTaskHeaderID & "&TaskNumber=" & strTaskID & ">" & strTitle & "</A>"
                ElseIf drTasks("sourcesystem") = "MOC" And drTasks("sourcesystemref") IsNot DBNull.Value Then
                    strSourceSystem = drTasks("sourcesystem")
                    strSourceRef = drTasks("Sourcesystemref")
                    'strHeaderURL = "<A HREF=HTTP://" & strDB & "/RI/MOC/EnterMOC.aspx?MOCNumber=" & strSourceRef & ">" & strHeaderTitle & " (" & strActivity & ")</A>"
                    strHeaderURL = strHeaderTitle
                    strDetailURL = "<A HREF=" & strUrl & "/TaskDetails.aspx?HeaderNumber=" & strTaskHeaderID & "&TaskNumber=" & strTaskID & "&RefSite=MOC>" & strTitle & "</A>"
                Else
                    strSourceSystem = ""
                    strSourceRef = ""
                    'strHeaderURL = "<A HREF=HTTP://" & strDB & "/TaskTracker/TaskHeader.aspx?HeaderNumber=" & strTaskHeaderID & ">" & strHeaderTitle & " (" & strActivity & ")</A>"
                    strHeaderURL = strHeaderTitle
                    strDetailURL = "<A HREF=" & strUrl & "/TaskDetails.aspx?HeaderNumber=" & strTaskHeaderID & "&TaskNumber=" & strTaskID & ">" & strTitle & "</A>"
                End If

                If previous_recType <> strRecType Or previous_recType = "" Then
                    sbEmailBody.Append("</table><P><font size=2 face=Arial><B><U>" & strRecType & "</B></U></FONT><BR>")
                    sbEmailBody.Append("<TABLE border=1 width=100%><TR valign=top width=5% wrap=hard><font size=2 face=Arial><B>{0}" & "Due Date" & "{1}<TD width=25%>" & "Header Info" & "{1}<TD width=25%>" & "Description" & "{1}")
                    'sbEmailBody.Append("{0}" & IPLoc.GetResourceValue("Responsible") & "{1}<TD width=15% wrap=hard>" & IPLoc.GetResourceValue("Comments/Links") & "{1}{0}" & IPLoc.GetResourceValue("BU/Type Manager"))
                    sbEmailBody.Append("{0}" & "Responsible" & "{1}<TD width=15% wrap=hard>" & "Comments/Links" & "{1}{0}" & "BU/Type Manager")


                    sbEmailBody.Append("</B></TR>")
                End If
                If intLeadTime > 0 Then
                    sbEmailBody.Append("<TR valign=top><font size=2 face=Arial>{0}*" & strDueDate & " (" & intLeadTime & "){1}")
                Else
                    sbEmailBody.Append("<TR valign=top><font size=2 face=Arial>{0}" & strDueDate & "{1}")
                End If
                sbEmailBody.Append("{0}" & strHeaderURL & "{1}")

                sbEmailBody.Append("{0}" & strDetailURL & "{1}")

                sbEmailBody.Append("{0}" & strResponsible & "{1}")
                If strComments = "" Then
                    sbEmailBody.Append("{0}" & strTaskLink & "{1}")
                Else
                    sbEmailBody.Append("{0}" & strComments & "<BR>" & strTaskLink & "{1}")
                End If
                sbEmailBody.Append("{0}" & strBusUnitMgr)

                previous_recType = strRecType
                strTasksFound = "Y"
                'Console.WriteLine(strMsg & "  ")

            End While

            If strRecType <> "" Then
                strMsg = sbEmailBody.ToString
                strMsg = String.Format(strMsg, v_td)
                strHeading1 = "* " & "Task Items with Lead Time"
                If strRole = "CREATOR" Then
                    strSubject = "Manufacturing Task Tracker tasks that you have created."
                    strHeading = "<HTML><BODY><font size=2 face=Arial><B>Here are the tasks from Manufacturing Task Tracker created by you that require your attention.</B>"
                    strFooter = "</HTML></BODY>"
                ElseIf strRole = "RESPONSIBLE" Then
                    If strDateRange = "DAILY" Then
                        'strSubject = "Manufacturing Task Tracker tasks that were entered for you."
                        strSubject = "Manufacturing Task Tracker tasks that were entered for you."
                        'strHeading = "<HTML><BODY><font size=2 face=Arial><B>Here are the tasks from Manufacturing Task Tracker that were entered yesterday that you are responsible for.</B>"
                        strHeading = "<HTML><BODY><font size=2 face=Arial><B>" & "Manufacturing Task Tracker tasks that were entered that you are responsible for." & "</B>"
                    Else
                        strSubject = "Manufacturing Task Tracker tasks that list you as Responsible."
                        'strHeading = "<HTML><BODY><font size=2 face=Arial><B>Following are tasks that you are responsible for.  Click Task Description to view or update (assign to another person, add comments, complete task by entering the closed date).</B>"
                        strHeading = "<HTML><BODY><font size=2 face=Arial><B>" & "Task Items You Are Responsible For" & "</b>"
                    End If
                    strFooter = "</HTML></BODY>"
                ElseIf strRole = "BUSUNITMGR" Then
                    'strSubject = "Manufacturing Task Tracker tasks that list you as Business Unit Manager OR Type Manager."
                    strSubject = "Manufacturing Task Tracker tasks that list you as Business Unit Manager OR Type Manager."
                    'strHeading = "<HTML><BODY><font size=2 face=Arial><B>Here are the tasks from Manufacturing Task Tracker that list you as the Business Unit Manager or Type Manager that require your attention.</B>"
                    strHeading = "<HTML><BODY><font size=2 face=Arial><B>" & "Tasks For Business Unit Manager or Type Manager that require your attention." & "</B>"
                    strFooter = "</HTML></BODY>"
                ElseIf strRole = "TYPEMGR" Then
                    strSubject = "Manufacturing Task Tracker tasks that list you as Type Manager."
                    strHeading = "<HTML><BODY><font size=2 face=Arial><B>Here are the tasks from Manufacturing Task Tracker that list you as the Type Manager that require your attention.</B>"
                    strFooter = "</HTML></BODY>"
                End If

                If dtRunDate <> Now().Date Then
                    strBody = "<P><font size =1 face=Arial><B>MTT BATCH EMAIL RERUN for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR><font size=1>" & strHeading & "<BR><font size=1>" & strHeading1 & "<BR>" & strMsg.ToString & strFooter
                Else
                    strBody = "<P><font size =1 face=Arial><B>" & strHeading & "<BR><font size=1>" & strHeading1 & "<BR>" & strMsg.ToString & strFooter
                End If

                ' strBody = cleanString(strBody, "<br>")
                Trace("MTTEmail", "GetTasks:SendEmail", tracing)
                SendEmail(strEmailAddress, ManufacturingEmail, strSubject, strBody)

                strBody = ""
            End If

            If strTasksFound = "Y" Then
                InsertAuditRecord(strSP, "Emailing " & strRole & " tasks for " & strUserid & " for " & strDateRange & " for " & dtRunDate.ToLongDateString)
            Else
                InsertAuditRecord(strSP, "NO " & strRole & " tasks for " & strUserid & " for " & strDateRange & " for " & dtRunDate.ToLongDateString)
            End If

            drTasks.Close()
            drTasks = Nothing
            connDB.Close()

        Catch ex As Exception
            CountExceptionsErrors()
            'Exception handling
            strErr = "Error occurred for user " & strUserid & "." & ex.Message
            'SendEmail("amy.albrinck@ipaper.com", "Manufacturing.task@ipaper.com", "AutoEmailError", strErr)
            HandleError(strSP, strErr, ex)
        End Try
    End Sub

    Sub GetEnteredTasks(ByVal strUserid As String, ByVal strLanguage As String, ByVal strEmailAddress As String)
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim drTasks As OracleDataReader
        Dim drRecurringTasks As OracleDataReader
        Dim strTaskID, strTaskHeaderID As String
        Dim strDueDate, strResponsible, strTitle, strBusUnitMgr As String

        Dim strHeading As String = ""
        Dim strHeading1 As String = ""
        Dim strErr As String
        Dim strSubject As String = ""
        Dim strHeaderTitle, strTaskDescription, strCreatedBy As String
        Dim strMsg, strSiteName, strRootTaskID As String
        Dim previous_SiteName As String = ""
        Dim strBody As String = ""
        Dim strFooter As String = ""
        Dim strSourceSystem As String
        Dim strSourceRef As String
        Dim strDetailURL As String
        Dim previous_recRootTask As String = ""
        Dim strRecurringDates As String
        Dim intLeadTime As Integer
        Dim param As New OracleParameter
        Dim sbEmailBody As New System.Text.StringBuilder
        Dim strTasksFound As String = "N"
        Dim strTaskLink As String = ""
        Dim strTaskLinkCount As String

        Trace("MTTEmail", "GetEnteredTasks:" + strRole, tracing)

        Try
            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            strSP = My.Application.Info.AssemblyName

            'Dim IPLoc As New IP.MEASFramework.ExtensibleLocalizationAssembly.WebLocalization(strLanguage, "MTT")

            'Get all endtered tasks associated with userid based on profile settings
            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MTTBATCHEMAILS.ENTEREDTASKLISTING"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "in_date"
                param.OracleDbType = OracleDbType.Date
                param.Direction = Data.ParameterDirection.Input
                param.Value = dtRunDate
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "in_userid"
                param.OracleDbType = OracleDbType.VarChar
                param.Direction = Data.ParameterDirection.Input
                param.Value = strUserid
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsEnteredTasks"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            sbEmailBody = New System.Text.StringBuilder
            Dim v_td As String() = {"<TD>", "</TD>"}

            drTasks = cmdSQL.ExecuteReader()

            Dim h As Integer = 0
            While drTasks.Read
                h = h + 1
                If drTasks("roottaskitemseqid") Is DBNull.Value Then
                    '                    strTaskID = drTasks("taskitemseqid")
                    strRootTaskID = ""
                Else
                    strRootTaskID = drTasks("roottaskitemseqid")
                End If

                strTaskID = drTasks("taskitemseqid")
                strTaskLinkCount = drTasks("tasklinkcount")
                If strTaskLinkCount > 0 Then
                    'Check the roottaskid.  If it is not null then this is a recurring task and we need to get links from the root task.  
                    If strRootTaskID = "" Then
                        strTaskLink = GetTaskItemDocs(strTaskID)
                    Else
                        strTaskLink = GetTaskItemDocs(strRootTaskID)
                    End If
                End If

                strSiteName = drTasks("sitename")
                strCreatedBy = drTasks("WHOLE_NAME_CREATEDBY_PERSON")
                strDueDate = drTasks("Item_DueDate")
                strDueDate = GetLocalizedDateTime(strDueDate, strLanguage, "MM/dd/yyyy")

                strRecurringDates = drTasks("Item_DueDate")
                strRecurringDates = GetLocalizedDateTime(strRecurringDates, strLanguage, "MM/dd/yyyy")

                strTitle = drTasks("Item_Title")
                If drTasks("Whole_Name_Responsible_Person") Is DBNull.Value Or drTasks("Whole_Name_Responsible_Person") = " " Then
                    strResponsible = drTasks("RoleDescription") & " (" & drTasks("Responsible_Role_Names") & ")"
                Else
                    strResponsible = drTasks("Whole_Name_Responsible_Person")
                End If
                strBusUnitMgr = drTasks("Mgr")

                strHeaderTitle = drTasks("taskheadertitle")
                strTaskHeaderID = drTasks("taskheaderseqid")
                intLeadTime = drTasks("leadtime")
                If drTasks("item_description") Is DBNull.Value Then
                    strTaskDescription = ""
                Else
                    strTaskDescription = drTasks("ITEM_DESCRIPTION")
                End If

                If drTasks("sourcesystem") Is DBNull.Value Then
                    strSourceSystem = ""
                    strSourceRef = ""
                    strDetailURL = "<A HREF=" & strUrl & "/TaskDetails.aspx?HeaderNumber=" & strTaskHeaderID & "&TaskNumber=" & strTaskID & ">" & strTitle & "</A>"""
                ElseIf drTasks("sourcesystem") = "MOC" Then
                    strDetailURL = "<A HREF=" & strUrl & "/TaskDetails.aspx?HeaderNumber=" & strTaskHeaderID & "&TaskNumber=" & strTaskID & "&RefSite=MOC>" & strTitle & "</A>"""
                Else
                    strSourceSystem = drTasks("sourcesystem")
                    strDetailURL = "<A HREF=" & strUrl & "/TaskDetails.aspx?HeaderNumber=" & strTaskHeaderID & "&TaskNumber=" & strTaskID & ">" & strTitle & "</A>"""
                End If

                'Only show the headers for the first record.
                If h = 1 Or strSiteName <> previous_SiteName Then
                    sbEmailBody.Append("</table><P><font size=2 face=Arial><B><U><I>" & strSiteName & "</I></B></U></FONT><BR>")
                    sbEmailBody.Append("<TABLE border=1 width=100%><font size =2 face=Arial>")
                    sbEmailBody.Append("<TR BGCOLOR=#AAAAAA valign=top><B><TD width=15%>" & "Responsible" & "{1}")
                    sbEmailBody.Append("<TD width=30%>" & "Title" & "{1}<TD width=45%>" & "Description" & "{1}/" & "Links" & "{0}" & "Created By" & "{1}</TR></B>")
                End If

                sbEmailBody.Append("<TR valign=top>")
                sbEmailBody.Append("{0}" & strResponsible & "{1}")

                'sbEmailBody.Append("{0}<A HREF=HTTP://" & strDB & "/TaskTracker/TaskHeader.aspx?HeaderNumber=" & strTaskHeaderID & ">" & strHeaderTitle & " (" & strActivity & ")</A>{1}")
                sbEmailBody.Append("{0}<A HREF=" & strUrl & "/TaskDetails.aspx?HeaderNumber=" & strTaskHeaderID & "&TaskNumber=" & strTaskID & ">" & strTitle & "</A>{1}")
                sbEmailBody.Append("{0}" & strTaskDescription & "<BR>" & strTaskLink & "{1}")
                sbEmailBody.Append("{0}" & strCreatedBy & "{1}")
                sbEmailBody.Append("</TR>")

                If strRootTaskID <> "" Then
                    cmdSQL = New OracleCommand
                    With cmdSQL
                        .Connection = connDB
                        .CommandText = "MTTBATCHEMAILS.GetRecurringTasks"
                        .CommandType = CommandType.StoredProcedure

                        param = New OracleParameter
                        param.ParameterName = "in_date"
                        param.OracleDbType = OracleDbType.Date
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = dtRunDate
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "in_TaskITem"
                        param.OracleDbType = OracleDbType.VarChar
                        param.Direction = Data.ParameterDirection.Input
                        param.Value = strRootTaskID
                        .Parameters.Add(param)

                        param = New OracleParameter
                        param.ParameterName = "rsRecurringTasks"
                        param.OracleDbType = OracleDbType.Cursor
                        param.Direction = ParameterDirection.Output
                        .Parameters.Add(param)

                    End With

                    drRecurringTasks = cmdSQL.ExecuteReader()
                    Dim i As Integer = 0
                    While drRecurringTasks.Read
                        Dim dtDueDate As Date = (drRecurringTasks("DueDate"))
                        i = i + 1
                        If i = 1 Then
                            'strRecurringDates = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(dtDueDate, strLanguage, "dd MMM yyyy") ' & "-" & GetTaskStatus(drRecurringTasks("statusseqid"), True)
                            strRecurringDates = GetLocalizedDateTime(dtDueDate, strLanguage, "MM/dd/yyyy")
                        Else
                            'strRecurringDates = strRecurringDates & ", " & drRecurringTasks("DueDate") '& "-" & GetTaskStatus(drRecurringTasks("statusseqid"), True)
                            strRecurringDates = GetLocalizedDateTime(dtDueDate, strLanguage, "MM/dd/yyyy")
                        End If
                    End While

                    'sbEmailBody.Append("<TR><TD colspan=4>Due Date(s)-")
                    sbEmailBody.Append("<TR><TD colspan=4>" & "Due Date" & "-")
                    sbEmailBody.Append(strRecurringDates & "</TD></TR>")

                    drRecurringTasks = Nothing
                Else
                    sbEmailBody.Append("<TR><TD colspan=4>" & "Due Date" & "-")
                    sbEmailBody.Append(strRecurringDates & "</TD></TR>")
                End If

                If h <> 1 And strSiteName <> previous_SiteName Then
                    sbEmailBody.Append("<tr BGCOLOR=#AAAAAA><TD colspan=4></td></tr>")
                End If
                previous_recRootTask = strRootTaskID
                previous_SiteName = strSiteName
                strTasksFound = "Y"
                'sbEmailBody.Append("<BR>")
            End While
            sbEmailBody.Append("</TABLE><BR>")

            If strTasksFound = "Y" Then
                strMsg = sbEmailBody.ToString
                strMsg = String.Format(strMsg, v_td)
                Dim dtDate As Date = dtRunDate
                dtDate = dtDate.AddDays(-1)
                dtDate = GetLocalizedDateTime(dtDate, strLanguage, "MM/dd/yyyy")
                Dim strDate As String
                strDate = GetLocalizedDateTime(dtDate, strLanguage, "MM/dd/yyyy")
                'Dim strDate As String = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(dtDate, strLanguage, "dd MMM yyyy")
                'dtDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(dtDate, strLanguage, "dd MMM yyyy")

                strSubject = "New tasks assigned to you"

                strHeading = "<HTML><BODY><font size=2 face=Arial><B>" & "Task Items Entered On " & strDate
                strHeading = String.Format(strHeading, strDate) & "</b>"
                'strHeading = "<HTML><BODY><font size=2 face=Arial><B>Here are the tasks from Manufacturing Task Tracker that were entered on " & dtDate.ToShortDateString & " that you are responsible for.  Click Title to view or update (assign to another person, add comments, complete task by entering the closed date).</B>"
                'strHeading1 = "<font size =1 face=Arial>* Task Items with Lead Time<BR>"
                strFooter = "</HTML></BODY>"

                If dtRunDate <> Now().Date Then
                    strBody = "<P><font size =2 face=Arial><B>MTT BATCH EMAIL RERUN for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR>" & strHeading & "<BR>" & strHeading1 & "<BR>" & strMsg.ToString & strFooter
                Else
                    'strBody = "<P><font size =2 face=Arial><B>MTT BATCH EMAIL for " & dtRunDate & "<BR>" & strEmailAddress & "</P><BR>" & strHeading & "<BR>" & strHeading1 & "<BR>" & strMsg.ToString & strFooter
                    strBody = strHeading & "<BR>" & strHeading1 & "<BR>" & strMsg.ToString & strFooter
                End If

                strBody = cleanString(strBody, "<br>")

                intUniqueEmailSent = intUniqueEmailSent + 1
                Trace("MTTEmail", "GetEnteredTasks:SendEmail" + strRole, tracing)
                SendEmail(strEmailAddress, ManufacturingEmail, strSubject, strBody)

                strBody = ""
            End If

            If strTasksFound = "Y" Then
                InsertAuditRecord(strSP, "Emailing " & strRole & " entered tasks for " & strUserid & " for " & dtRunDate.ToLongDateString)
            Else
                InsertAuditRecord(strSP, "NO " & strRole & " entered tasks for " & strUserid & " for " & dtRunDate.ToLongDateString)
            End If
            connDB.Close()

            drTasks = Nothing

        Catch ex As Exception
            CountExceptionsErrors()
            'Exception handling
            strErr = "Error occurred for user " & strUserid & "." & ex.Message
            'SendEmail("amy.albrinck@ipaper.com", "Manufacturing.task@ipaper.com", "AutoEmailError", strErr)
            HandleError(strSP, strErr, ex)
        End Try
    End Sub



    Sub GetIndTask(ByVal strTaskID As String)
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim drIndTask As OracleDataReader
        Dim strTaskHeaderID, strTaskItemID, strEmailAddress As String
        Dim strLanguage As String = "en-US"
        Dim strDueDate, strResponsible, strActivity, strTitle, strBusUnitMgr, strComments As String

        Dim strHeading, strErr As String
        Dim strSubject, strHeaderTitle As String
        Dim strMsg, strSiteName As String
        Dim strBody As String = ""
        Dim strFooter As String = ""
        Dim strRecType As String = ""
        Dim param As New OracleParameter
        Dim sbEmailBody As New System.Text.StringBuilder

        Trace("MTTEmail", "GetIndTask:" + strRole, tracing)

        Try
            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            strSP = My.Application.Info.AssemblyName

            InsertAuditRecord(strSP, "Emailing for taskid " & strTaskID)

            'Dim IPLoc As New IP.MEASFramework.ExtensibleLocalizationAssembly.WebLocalization(strLanguage, "RI")

            'Get all tasks associated with userid based on profile settings
            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MTTBATCHEMAILS.INDTASKLISTING"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "in_taskid"
                param.OracleDbType = OracleDbType.VarChar
                param.Direction = Data.ParameterDirection.Input
                param.Value = strTaskID
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsIndTask"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With

            Dim v_td As String() = {"<TD>", "</TD>"}

            drIndTask = cmdSQL.ExecuteReader()

            While drIndTask.Read

                sbEmailBody = New System.Text.StringBuilder

                strLanguage = drIndTask("RESPONSIBLE_DEFAULTLANGUAGE")
                strEmailAddress = drIndTask("RESPONSIBLE_EMAIL")

                strRecType = drIndTask("RECTYPE")

                strSiteName = drIndTask("sitename") 'Cannot be NULL

                strDueDate = drIndTask("Item_DueDate") 'Cannot be NULL
                strDueDate = GetLocalizedDateTime(strDueDate, strLanguage, "MM/dd/yyyy")
                'strDueDate = IP.MEASFramework.ExtensibleLocalizationAssembly.DateTime.GetLocalizedDateTime(strDueDate, strLanguage, "d")
                strTitle = Trim(drIndTask("Item_Title"))
                If drIndTask("Whole_Name_Responsible_Person") Is DBNull.Value Or drIndTask("Whole_Name_Responsible_Person") = " " Then
                    strResponsible = drIndTask("RoleDescription") & " (" & drIndTask("Responsible_Role_Names") & ")"
                Else
                    strResponsible = drIndTask("Whole_Name_Responsible_Person")
                End If
                strBusUnitMgr = drIndTask("Mgr")
                strTaskID = drIndTask("taskitemseqid")
                strHeaderTitle = drIndTask("taskheadertitle")
                strActivity = drIndTask("ActivityName")
                If drIndTask("Mttcomment") Is DBNull.Value Then
                    strComments = ""
                Else
                    strComments = drIndTask("mttcomment")
                End If
                strTaskHeaderID = drIndTask("taskheaderseqid")
                strTaskItemID = drIndTask("taskitemseqid")

                sbEmailBody.Append("</table><P><font size =2 face=Arial><B><U>" & strRecType & "</B></U></FONT><BR>")
                sbEmailBody.Append("<TABLE border=1><font size =2 face=Arial><TR valign=top><B>{0}" & "Due Date" & "{1}<TD width=25%>Header Info{1}<TD width=25%>Task Description{1}")
                sbEmailBody.Append("{0}Responsible{1}<TD width=15% wrap=hard>Comments{1}{0}BU/Type Manager{1}")
                sbEmailBody.Append("</B></TR>")
                sbEmailBody.Append("<BR><TR valign=top><font size=2>{0}" & strDueDate & "{1}")

                sbEmailBody.Append("{0}<A HREF=" & strUrl & "/TaskHeader.aspx?HeaderNumber=" & strTaskHeaderID & ">" & strHeaderTitle & " (" & strActivity & ")</A>{1}")
                sbEmailBody.Append("{0}<A HREF=" & strUrl & "/TaskDetails.aspx?HeaderNumber=" & strTaskHeaderID & "&TaskNumber=" & strTaskID & ">" & strTitle & "</A>{1}")
                sbEmailBody.Append("{0}" & strResponsible & "{1}")
                sbEmailBody.Append("{0}" & strComments & "{1}")
                sbEmailBody.Append("{0}" & strBusUnitMgr & "{1}")

                strMsg = sbEmailBody.ToString
                strMsg = String.Format(strMsg, v_td)
                strSubject = "Manufacturing Task Tracker tasks that were entered and you."
                strHeading = "<HTML><BODY><font size=3 face=Arial><B>Here are the tasks from Manufacturing Task Tracker that were entered yesterday that you are responsible for.</B>"
                strFooter = "</HTML></BODY>"

                strBody = strHeading & "<BR>" & strMsg.ToString & strFooter
                strBody = cleanString(strBody, "<br>")

                Trace("MTTEmail", "GetIndTask:SendEmail" + strRole, tracing)
                SendEmail(strEmailAddress, "Manufacturing.task@@graphicpkg.com", strSubject, strBody)

                strBody = String.Empty

            End While

            drIndTask = Nothing
            connDB = Nothing

        Catch ex As Exception
            'Exception handling
            CountExceptionsErrors()
            strErr = "Error occurred." & ex.Message
            Trace("MTTEmail", "GetIndTask:Error" + strRole + " " + strErr, tracing)
            'SendEmail("amy.albrinck@@graphicpkg.com", "Manufacturing.task@@graphicpkg.com", "AutoEmailError", strErr)
            HandleError(strSP, strErr, ex)
        End Try
    End Sub
    Public Function GetTaskItemDocs(ByVal strtaskItemId As String) As String
        Dim strErr, strTaskLink, strTaskLinkDesc, strTaskLinkLocation, strTaskLinkFile As String
        Dim connDB As New OracleConnection
        Dim cmdSQL As OracleCommand = Nothing
        Dim drTaskItemDocs As OracleDataReader
        Dim param As New OracleParameter
        Dim sbEmailBody As New System.Text.StringBuilder

        strTaskLinkDesc = ""
        strTaskLinkLocation = ""
        strTaskLinkFile = ""
        strTaskLink = ""

        Try
            strSP = My.Application.Info.AssemblyName

            If strDB = "RIDEV" Then
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFATST").ToString
            Else
                connDB.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("connectionRCFAPRD").ToString
            End If
            connDB.Open()

            cmdSQL = New OracleCommand
            With cmdSQL
                .Connection = connDB
                .CommandText = "MTTBATCHEMAILS.TASKITEMDOCS"
                .CommandType = CommandType.StoredProcedure

                param = New OracleParameter
                param.ParameterName = "in_taskid"
                param.Direction = Data.ParameterDirection.Input
                param.Value = strtaskItemId
                .Parameters.Add(param)

                param = New OracleParameter
                param.ParameterName = "rsTaskDocs"
                param.OracleDbType = OracleDbType.Cursor
                param.Direction = ParameterDirection.Output
                .Parameters.Add(param)

            End With
            drTaskItemDocs = cmdSQL.ExecuteReader()
            While drTaskItemDocs.Read
                strTaskLinkDesc = drTaskItemDocs("description")
                strTaskLinkLocation = drTaskItemDocs("location")
                strTaskLinkFile = drTaskItemDocs("filename")

                strTaskLink = strTaskLink & "<A HREF='" & GetFileLocation(Replace(strTaskLinkFile, "'", "&#39"), Replace(strTaskLinkLocation, "'", "&#39")) & "'>" & strTaskLinkDesc & "</A>" & " <BR> "
            End While

            connDB.Close()

        Catch ex As Exception
            CountExceptionsErrors()
            'Exception handling
            strErr = "Error occurred." & ex.Message
            'SendEmail("amy.albrinck@ipaper.com", "Manufacturing.task@ipaper.com", "AutoEmailError", strErr)
            HandleError(strSP, strErr, ex)
        Finally
            If connDB IsNot Nothing Then
                If connDB.State = ConnectionState.Open Then connDB.Close()
                connDB = Nothing
            End If
        End Try

        Return strTaskLink
    End Function


    Public Function GetFileLocation(ByVal file As String, ByVal location As String) As String
        Trace("MTTEmail", "GetFileLocation:" + strRole, tracing)
        If file <> "na" And location.Length > 0 Then 'Attachment is a file
            Dim strpath As String = ""
            strpath = String.Format(strFileURL & "{0}", file)
            Return strpath
        Else 'Attachment is a URL
            If location.StartsWith("www", StringComparison.CurrentCulture) Then
                location = "http://" & location
            End If
            Return location
        End If
    End Function

    'Public Function GetTaskStatus(ByVal statusID As Integer, ByVal includeLabel As Boolean, Optional ByVal dueDate As String = "") As String
    '    Dim returnVal As String
    '    Try
    '        Dim NoWorkNeeded As String
    '        Dim Cancelled As String
    '        Dim Completed As String
    '        Dim LateNotCompleted As String
    '        Dim WorkInProcess As String
    '        'Dim A As System.Net.Mail.Attachment = New System.Net.Mail.Attachment("C:\Visual Studio Web Sites\MTTEmail\MTTEmail\Images\complete.gif")

    '        If includeLabel = True Then
    '            'Localization code can be added here
    '            NoWorkNeeded = "No Work Needed"
    '            Cancelled = "Cancelled"
    '            Completed = "Completed"
    '            LateNotCompleted = "Late/Not Completed"
    '            WorkInProcess = "Open"
    '        Else
    '            NoWorkNeeded = String.Empty
    '            Cancelled = String.Empty
    '            Completed = String.Empty
    '            LateNotCompleted = String.Empty
    '            WorkInProcess = String.Empty
    '        End If
    '        'Dim imagePath As String = "\\MTTEMail\Images\"
    '        'Dim imgNoWorkNeeded As String = "<img src='" & imagePath & "noworkneeded.gif' align=center width=15 height=15 title='No Work Needed' alt='No Work Needed'>" & NoWorkNeeded
    '        'Dim imgCanceled As String = "<img src='" & imagePath & "cancelled.gif' align=center width=15 height=15 title='Cancelled' alt='Cancelled'>" & Cancelled
    '        'Dim imgCompleted As String = "<img src='" & imagePath & "complete.gif' align=center width=15 height=15 title='Closed Complete' alt='Closed Complete'>" & Completed
    '        'Dim imgLateNotCompleted As String = "<img src='" & imagePath & "late_notcomp.gif' align=center width=15 height=15  title='Late/Not Completed' alt='Late/Not Completed'/>" & LateNotCompleted
    '        'Dim imgWorkInProcess As String = "<img src='" & imagePath & "wip.gif' align=center width=15 height=15 title='Open' alt='Open'/>" & WorkInProcess
    '        Dim imgNoWorkNeeded As String = NoWorkNeeded
    '        Dim imgCanceled As String = Cancelled
    '        Dim imgCompleted As String = Completed
    '        Dim imgLateNotCompleted As String = LateNotCompleted
    '        Dim imgWorkInProcess As String = WorkInProcess

    '        If dueDate.Length > 0 AndAlso IsDate(dueDate) Then
    '            If statusID = 1 And CDate(dueDate) < Now Then
    '                statusID = 0
    '            End If
    '        End If
    '        Select Case statusID
    '            Case 0 'Overdue
    '                Return imgLateNotCompleted
    '            Case 1 'Open
    '                returnVal = imgWorkInProcess
    '            Case 2 'Complete
    '                returnVal = imgCompleted
    '            Case 3 'No Work Needed
    '                returnVal = imgNoWorkNeeded
    '            Case 4 'Cancelled
    '                returnVal = imgCanceled
    '            Case Else
    '                returnVal = ""
    '        End Select
    '    Catch
    '        Throw
    '    End Try
    '    Return returnVal
    'End Function
    'Declaration

End Module
