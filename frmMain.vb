Option Strict Off
Option Explicit On

Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports DBPM_Server.SQLHelper
Imports DBPM_Server.siteConstants
Imports System.Data.Common
Imports System.Linq
Imports System.Collections

Imports System.Net.Mail

Partial Friend Class frmMain
    Inherits System.Windows.Forms.Form

    Private Sub CheckDataConnection(ByVal myConnection As SqlConnection)
        Select Case myConnection.ToString
            Case "cnDBPM"
                If cnDBPM.State <> ConnectionState.Open Then OpenConnectionDBPM()
            Case "cnQuickBooks"
                If cnQuickBooks.State <> ConnectionState.Open Then OpenConnectionQB()
            Case "cnMax"
                If cnMax.State <> ConnectionState.Open Then OpenConnectionMax()
        End Select
    End Sub


    Private Sub cmdNumberDropShips_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdNumberDropShips.Click
        If booQBRefreshInProgress Then Exit Sub
        If gstrCompany = "" Then Exit Sub

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "PutJobInvoicesIn" '"SUBNAME"

        Debug.WriteLine(Now.ToString & " : NumberDropShips.Click")

        Try
            NumberDropShipsNew()
        Catch ex As Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "")
            Exit Sub
        End Try

    End Sub


    Public Sub NumberDropShipsNew()
        Dim strErrorLine As String = ""

        'Used for ErrorHandling
        Dim strObjName As String = "frmMain"
        Dim strSubName As String = "NumberDropShips"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        CheckDataConnection(cnMax)

        Dim strNumDropShipsSQL As String = "SELECT a.StepN, a.SalesOrderN, a.ShipToDescription, (SELECT Max(ShipToDescription) FROM JOB_STEP WHERE Job_Step.SalesOrderN = a.SalesOrderN) as CurrentMaxDescription FROM JOB_Step a WHERE a.ShipToDescription = '' AND a.Steptype = 'Ship' AND a.StepJobNum NOT LIKE '%F&B%' Order by a.SalesOrderN, a.StepN"

        Using rs1_JOB_Step As New DataSet("Job_Step")
            Using job_Step_DA As New SqlDataAdapter(strNumDropShipsSQL, cnMax)
                job_Step_DA.FillSchema(rs1_JOB_Step, SchemaType.Source, "Job_Step")
                job_Step_DA.Fill(rs1_JOB_Step, "Job_Step")

                Dim tblJobSteps As DataTable = rs1_JOB_Step.Tables("Job_Step")
                Dim curSalesOrderN, oldSalesOrderN, MaxShipTo As String
                Dim iMax As Integer = 0
                oldSalesOrderN = ""
                For Each dr As DataRow In tblJobSteps.Rows
                    curSalesOrderN = NCStr(dr("SalesOrderN"))
                    If curSalesOrderN = oldSalesOrderN Then
                        'keep updating the shipToDescription Column
                        iMax += 1
                    Else
                        'this is a new salesorder number, so restart the maxshipto with the maximum already in the DB
                        MaxShipTo = NCStr(dr("CurrentMaxDescription"), "000")
                        If MaxShipTo = "" Then
                            iMax = 1
                        Else
                            iMax = CType(MaxShipTo, Integer) + 1
                        End If

                        oldSalesOrderN = NCStr(dr("SalesOrderN"))
                    End If
                    MaxShipTo = iMax.ToString.PadLeft(3, "0")

                    dr.BeginEdit()
                    dr("ShipToDescription") = MaxShipTo
                    dr.EndEdit()

                Next

                ShowUserMessage(strObjName, " NumberDropShips: " & tblJobSteps.Rows.Count & " records updated...", True)

                Using objCommandBuilder As New SqlCommandBuilder(job_Step_DA)
                    job_Step_DA.Update(rs1_JOB_Step, "Job_Step")
                End Using

            End Using
        End Using



    End Sub


    Private Sub cmdRefreshQBTables_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdRefreshQBTables.Click
        If booQBRefreshInProgress Then Exit Sub
        If gstrCompany = "" Then Exit Sub
        Debug.WriteLine(Now.ToString & " : RefreshQBTables.Click")

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "cmdRefreshQBTables_Click" '"SUBNAME"

        If booQBRefreshInProgress Then
            MsgBox("QB is currently being refreshed.")
        Else
            RefreshQBTables()
        End If

    End Sub


    Private Sub RunDocFinderRefresh()

        If Me.chkPauseProcessing.CheckState = CheckState.Checked Then Exit Sub
        If booQBRefreshInProgress Then Exit Sub

        ShowUserMessage("DocFinderRefresh", "Running DocFinderRefresh", "Starting DocFinderRefresh", True)

        booQBRefreshInProgress = True

        Try
            Dim curYearFiles As String = "2" & Now.ToString("yy") & "*"
            DocFinderRefresh2(curYearFiles)

            If Now.Month < 4 Then
                'january - march, check last year files as well
                Dim lastYearFiles As String = "2" & Now.AddYears(-1).ToString("yy") & "*"
                DocFinderRefresh2(lastYearFiles)
            End If

        Catch ex As Exception
            ShowUserMessage("DocFinderRefresh", "DocFinderRefresh had an error: " & ex.Message)
            HaveError("frmMain", "RefreshQBTables", Information.Err.Number, ex.Message, Information.Err.Source, "", "")
        End Try

        booQBRefreshInProgress = False

        'show finished
        ShowUserMessage("DocFinderRefresh", "Finished DocFinderRefresh", "Finished DocFinderRefresh", True)

    End Sub

    Private Sub cmdRunDocFinderRefresh_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdRunDocFinderRefresh.Click
        If booQBRefreshInProgress Then Exit Sub
        If gstrCompany = "" Then Exit Sub

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "DocFinderRefresh" '"SUBNAME"

        Debug.WriteLine(Now.ToString & " : DocFinderRefresh.Click")

        Try
            RunDocFinderRefresh()
        Catch ex As Exception
            HaveError(strObjName, strSubName & "RunDocFinderRefresh", Information.Err.Number, ex.Message, Information.Err.Source, "", "")
        End Try

    End Sub

    Private Sub frmMain_Load(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles MyBase.Load

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "Form_Load" '"SUBNAME"
        lblVersion.Text = CStr(My.Application.Info.Version.Major) & "." & CStr(My.Application.Info.Version.Minor) & "." & CStr(My.Application.Info.Version.Revision)
        Me.Show()

        OpenConnectionDBPM()
        OpenConnectionMax()
        OpenConnectionQB()

        Dim autoRun As Boolean = isAutoRun()
        If autoRun Then
            chkSeeProcessing.Checked = False
            gShowProcessing = False
            ShowUserMessage(strObjName, "Program started. Auto-running Processes...", "Program started...", True)

            RunDBPMProcesses()

            Me.Close()
        Else
            chkSeeProcessing.Checked = True
            gShowProcessing = True
        End If
        Using sqlhelp As New SQLHelper
            Dim maxJobNum As String = sqlhelp.GetDataItem("SELECT TOP 1 jobnum FROM job_header where ProcessType = 'Job' order by SalesOrderN desc")
            ShowUserMessage("Loading", "MAX JOB NUMBER: " & maxJobNum, , True)
        End Using
    End Sub
    Private Function isAutoRun() As Boolean

        Dim retVal As Boolean = False
        For Each s As String In My.Application.CommandLineArgs
            If s = "-autorun" Then retVal = True
        Next
        Return retVal
    End Function

    Private Sub frmMain_Closed(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles MyBase.Closed

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "Form_Unload" '"SUBNAME"



        Try
            If iErrCount > 0 Then
                'had errors... email
                Dim client As New SmtpClient(mailHost, mailPort)
                Dim msg As New MailMessage(mailErrorSendFrom, mailErrorSendTo)
                msg.Body = "Error has occurred on DBPM_Server. Errors are listed below and should also be recorded in SQL Server" & Environment.NewLine & Environment.NewLine & strProgramErrors.ToString
                msg.BodyEncoding = System.Text.Encoding.UTF8
                msg.Subject = "Errors on DBPM"
                client.Port = mailPort
                client.Credentials = New System.Net.NetworkCredential(mailErrorSendFrom, mailPassword)
                client.EnableSsl = True
                client.Send(msg)


            End If
        Catch ex As Exception

        End Try

        'Error handling
        Try

            If Not (cnQuickBooks Is Nothing) Then
                If cnQuickBooks.State = ConnectionState.Open Then cnQuickBooks.Close()
                cnQuickBooks = Nothing
            End If
            Debug.WriteLine("cnQuickBooks Closed")

            If Not (cnDBPM Is Nothing) Then
                If cnDBPM.State = ConnectionState.Open Then cnDBPM.Close()
                cnDBPM = Nothing
            End If
            Debug.WriteLine("cnDBPM Closed")

            If Not (cnMax Is Nothing) Then
                If cnMax.State = ConnectionState.Open Then cnMax.Close()
                cnMax = Nothing
            End If
            Debug.WriteLine("cnMax Closed")


            Environment.Exit(0)

        Catch exc As System.Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "")
        End Try

    End Sub







    Private Sub RunDBPMProcesses()
        If booQBRefreshInProgress Then Exit Sub
        If gbooEndProgram Then Environment.Exit(0)
        If gstrCompany = "" Then Exit Sub


        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "RunDBPMProcesses" '"SUBNAME"

        Debug.WriteLine(Now.ToString & "   Timer")

        'clear the listbox if getting full
        If Me.lstConversionProgress.Items.Count < -30000 Then Me.lstConversionProgress.Items.Clear()

        Try
            RefreshQBTables()
        Catch ex As Exception
            HaveError(strObjName, strSubName & "RefreshQBTables", Information.Err.Number, ex.Message, Information.Err.Source, "", "")
            Exit Try
        End Try

        Try
            PutJobInvoicesIn()
        Catch ex As Exception
            HaveError(strObjName, strSubName & "PutJobInvoicesIn", Information.Err.Number, ex.Message, Information.Err.Source, "", "")
            Exit Try
        End Try

        Try
            NumberDropShipsNew()
        Catch ex As Exception
            HaveError(strObjName, strSubName & "NumberDropShipsNew", Information.Err.Number, ex.Message, Information.Err.Source, "", "")
            Exit Try
        End Try

        Try
            RunDocFinderRefresh()
        Catch ex As Exception
            HaveError(strObjName, strSubName & "RunDocFinderRefresh", Information.Err.Number, ex.Message, Information.Err.Source, "", "")
            Exit Try
        End Try

        Try
            Dim rightNow As Date = DateTime.Now
            If rightNow.TimeOfDay >= New TimeSpan(19, 0, 0) AndAlso rightNow.TimeOfDay <= New TimeSpan(19, 15, 0) Then
                ShowUserMessage("RunDBPMProcesses", "Fix Max Deleted QB Customers", "Fix Max Deleted QB Customers")
                DeletedInQBSoFixInMax()
            End If

        Catch ex As Exception
            HaveError("RunDBPMProcesses", "RunDBPMProcesses : Fix Deleted QB Customers", CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
            Exit Try
        End Try

        Try
            'Update the teather clock
            Dim strSQL As String = ""
            If Not booQBRefreshInProgress Then
                strSQL = "UPDATE DBPM_Status SET StatusDate = getdate(), StatusUser = 'PROGRAM', StatusComputer = 'Server06' WHERE Condition = 'DBPM Server' AND Status = 'Last Run'"
                SQLHelper.ExecuteSQL(cnDBPM, strSQL)
            End If

        Catch ex As Exception
            HaveError("RunDBPMProcesses", "RunDBPMProcesses : Update Tether Clock", CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
            Exit Try
        End Try
    End Sub












    Private Sub DeletedInQBSoFixInMax()
        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "DeletedInQBSoFixInMax" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling

        Try
            Dim strSQL1 As String = ""
            strSQL1 = "UPDATE AMGR_Client_Tbl SET Division = '' WHERE Firm not in (SELECT ListID FROM QB_Customer ) AND Name_Type = 'C' AND Firm <> '' AND Division <> 'DELETED IN QB' "
            SQLHelper.ExecuteSQL(cnMax, strSQL1)
        Catch ex As Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "")
        End Try



    End Sub








    Public Sub DocFinderRefresh(ByVal sFilter As String)

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "DocFinderRefresh" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'FOR PART 1_ - Get records from JOB_DocsToFind
        Debug.WriteLine("List1_JOB_DocsToFind")

        Dim str1_JOB_DocsToFindSQL, str1_SearchTitle, str1_SearchPath, str1_SearchExtention As String
        'This routine gets the 1_JOB_DocsToFind from the database according to the selection in str1_JOB_DocsToFindSQL.
        'It then puts those 1_JOB_DocsToFind in the list box

        'FOR PART 2_ - Get records from JOB_DocsFound
        Debug.WriteLine("List2_JOB_DocsFound")
        Dim str2_MAXDocCreatedDate As Date
        'This routine gets the 2_JOB_DocsFound from the database according to the selection in str2_JOB_DocsFoundSQL.
        'It then puts those 2_JOB_DocsFound in the list box

        Using rs1_JOB_DocsToFind As New DataSet()

            str1_JOB_DocsToFindSQL = "SELECT * FROM JOB_DocsToFind WHERE IsActive = '1' AND IsRefresh = '1' "

            Dim adap As SqlDataAdapter = New SqlDataAdapter(str1_JOB_DocsToFindSQL, cnDBPM)
            rs1_JOB_DocsToFind.Tables.Clear()
            adap.Fill(rs1_JOB_DocsToFind)
            Dim strDirCommand As String = ""
            Dim iCount As Integer = 0
            Dim totalRows As String = ""
            Dim numFiles As Integer = 0

            If rs1_JOB_DocsToFind.Tables(0).Rows.Count > 0 Then
                totalRows = rs1_JOB_DocsToFind.Tables(0).Rows.Count.ToString

                For Each iteration_row As DataRow In rs1_JOB_DocsToFind.Tables(0).Rows
                    iCount += 1
                    'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs1_JOB_DocsToFind.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                    ShowUserMessage(strSubName, "Processing files starting with " & sFilter.Replace("*", ""))

                    Try


                        'Clear strings
                        str1_SearchPath = ""
                        str1_SearchExtention = ""

                        'get the columns from the database
                        str1_SearchPath = NCStr(iteration_row("SearchPath"))
                        str1_SearchExtention = NCStr(iteration_row("SearchExtention"))
                        str1_SearchTitle = NCStr(iteration_row("SearchTitle"))

                        'Get the current highest date of a file that has the SearchPath
                        Dim str2_JOB_DocsFoundSQL As String = "SELECT MAX(DocCreatedDate) AS DocCreatedDate FROM JOB_DocsFound WHERE DocPath like '" & str1_SearchPath & "%'"
                        Debug.WriteLine(str2_JOB_DocsFoundSQL)

                        Try
                            str2_MAXDocCreatedDate = CType(SQLHelper.ExecuteScalerDate(cnDBPM, CommandType.Text, str2_JOB_DocsFoundSQL), Date)
                        Catch ex As Exception
                            str2_MAXDocCreatedDate = Date.Now.AddYears(-1)
                        End Try


                        Dim mypath As String = str1_SearchPath
                        mypath = "\\devserver\Accounting\FileShare\ASI"     'use for dev testing
                        ShowUserMessage("DocFinderRefresh", "Path to Search: " & mypath, , True)

                        Dim dir As System.IO.DirectoryInfo
                        Dim fileList As FileInfo()
                        Dim dirs As IEnumerable(Of String)

                        dirs = Directory.EnumerateDirectories(mypath, sFilter, SearchOption.AllDirectories)

                        For Each folder As String In dirs
                            Try
                                dir = New System.IO.DirectoryInfo(folder)
                                fileList = dir.GetFiles("*" & str1_SearchExtention, System.IO.SearchOption.AllDirectories)
                                ShowUserMessage(folder, fileList.Count)

                                ' Search the contents of each file.
                                ' A regular expression created with the RegEx class
                                ' could be used instead of the Contains method.
                                Dim queryMatchingFiles = From file In fileList
                                                         Where file.CreationTime > str2_MAXDocCreatedDate
                                                         Select file

                                For Each myFile As FileInfo In queryMatchingFiles

                                    Dim mySQL As String = "INSERT INTO JOB_DocsFound (CreatedDate, CreatedBy, CreatedOnComputer, DocTitle, DocPath, DocCreatedDate, DocModifiedDate, DocFileSize, DocFileName, DocExtension) " & _
                                                           "VALUES (@CreatedDate, @CreatedBy, @CreatedOnComputer, @DocTitle, @DocPath, @DocCreatedDate, @DocModifiedDate, @DocFileSize, @DocFileName, @DocExtension)"

                                    SQLHelper.ExecuteSQL(cnDBPM, mySQL, _
                                                         New SqlParameter("@CreatedDate", Date.Now), _
                                                         New SqlParameter("@CreatedBy", gstrUserName), _
                                                         New SqlParameter("@CreatedOnComputer", gstrComputerName), _
                                                         New SqlParameter("@DocTitle", str1_SearchTitle), _
                                                         New SqlParameter("@DocPath", myFile.DirectoryName), _
                                                         New SqlParameter("@DocCreatedDate", myFile.CreationTime), _
                                                         New SqlParameter("@DocModifiedDate", myFile.LastWriteTime), _
                                                         New SqlParameter("@DocFileSize", myFile.Length), _
                                                         New SqlParameter("@DocFileName", Path.GetFileNameWithoutExtension(myFile.Name)), _
                                                         New SqlParameter("@DocExtension", myFile.Extension))
                                    numFiles += 1
                                Next

                            Catch ex As Exception
                                HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                                Continue For
                            End Try
                        Next

                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For
                    End Try
                Next iteration_row
            End If

            'Update status listbox
            ShowUserMessage(strObjName, numFiles.ToString & " Files Processed", , True)
        End Using

        Try
            Dim TempCommand_2 As SqlCommand
            TempCommand_2 = cnDBPM.CreateCommand()
            TempCommand_2.CommandText = "sp_JobDocsRunAfterSearch"
            TempCommand_2.ExecuteNonQuery()
        Catch ex As Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
        End Try






    End Sub




    Public Sub DocFinderRefresh2(ByVal sFilter As String)

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "DocFinderRefresh" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'FOR PART 1_ - Get records from JOB_DocsToFind
        Debug.WriteLine("List1_JOB_DocsToFind")

        Dim str1_JOB_DocsToFindSQL, str1_SearchTitle, str1_SearchPath, str1_SearchExtention As String
        'This routine gets the 1_JOB_DocsToFind from the database according to the selection in str1_JOB_DocsToFindSQL.
        'It then puts those 1_JOB_DocsToFind in the list box

        'FOR PART 2_ - Get records from JOB_DocsFound
        Debug.WriteLine("List2_JOB_DocsFound")
        Dim str2_MAXDocCreatedDate As Date
        'This routine gets the 2_JOB_DocsFound from the database according to the selection in str2_JOB_DocsFoundSQL.
        'It then puts those 2_JOB_DocsFound in the list box

        Using rs1_JOB_DocsToFind As New DataSet()

            str1_JOB_DocsToFindSQL = "SELECT * FROM JOB_DocsToFind WHERE IsActive = '1' AND IsRefresh = '1' "

            Dim adap As SqlDataAdapter = New SqlDataAdapter(str1_JOB_DocsToFindSQL, cnDBPM)
            rs1_JOB_DocsToFind.Tables.Clear()
            adap.Fill(rs1_JOB_DocsToFind)
            Dim strDirCommand As String = ""
            Dim iCount As Integer = 0
            Dim totalRows As String = ""
            Dim numFiles As Integer = 0

            If rs1_JOB_DocsToFind.Tables(0).Rows.Count > 0 Then
                totalRows = rs1_JOB_DocsToFind.Tables(0).Rows.Count.ToString

                For Each iteration_row As DataRow In rs1_JOB_DocsToFind.Tables(0).Rows
                    iCount += 1
                    'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs1_JOB_DocsToFind.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                    ShowUserMessage(strSubName, "Processing files starting with " & sFilter.Replace("*", ""))

                    Try


                        'Clear strings
                        str1_SearchPath = ""
                        str1_SearchExtention = ""

                        'get the columns from the database
                        str1_SearchPath = NCStr(iteration_row("SearchPath"))
                        str1_SearchExtention = NCStr(iteration_row("SearchExtention"))
                        str1_SearchTitle = NCStr(iteration_row("SearchTitle"))

                        'Get the current highest date of a file that has the SearchPath
                        Dim str2_JOB_DocsFoundSQL As String = "SELECT MAX(DocCreatedDate) AS DocCreatedDate FROM JOB_DocsFound WHERE DocPath like '" & str1_SearchPath & "%'"
                        Debug.WriteLine(str2_JOB_DocsFoundSQL)

                        Try
                            str2_MAXDocCreatedDate = CType(SQLHelper.ExecuteScalerDate(cnDBPM, CommandType.Text, str2_JOB_DocsFoundSQL), Date)
                        Catch ex As Exception
                            str2_MAXDocCreatedDate = Date.Now.AddYears(-1)
                        End Try


                        Dim mypath As String = str1_SearchPath
                        'mypath = "\\devserver\Accounting\FileShare\ASI"     'use for dev testing
                        ShowUserMessage("DocFinderRefresh", "Path to Search: " & mypath, , True)

                        Dim diTop As New DirectoryInfo(mypath)
                        Try
                            For Each di As DirectoryInfo In diTop.EnumerateDirectories(sFilter)
                                Try

                                    Dim queryMatchingFiles = From file In di.EnumerateFiles("*" & str1_SearchExtention, SearchOption.TopDirectoryOnly)
                                                                         Where file.CreationTime > str2_MAXDocCreatedDate
                                                                         Select file

                                    For Each myFile As FileInfo In queryMatchingFiles
                                        Try
                                            ShowUserMessage("refreshDocs", myFile.Name, , True)

                                            Dim mySQL As String = "INSERT INTO JOB_DocsFound (CreatedDate, CreatedBy, CreatedOnComputer, DocTitle, DocPath, DocCreatedDate, DocModifiedDate, DocFileSize, DocFileName, DocExtension) " & _
                                                         "VALUES (@CreatedDate, @CreatedBy, @CreatedOnComputer, @DocTitle, @DocPath, @DocCreatedDate, @DocModifiedDate, @DocFileSize, @DocFileName, @DocExtension)"

                                            SQLHelper.ExecuteSQL(cnDBPM, mySQL, _
                                                                 New SqlParameter("@CreatedDate", Date.Now), _
                                                                 New SqlParameter("@CreatedBy", gstrUserName), _
                                                                 New SqlParameter("@CreatedOnComputer", gstrComputerName), _
                                                                 New SqlParameter("@DocTitle", str1_SearchTitle), _
                                                                 New SqlParameter("@DocPath", myFile.DirectoryName), _
                                                                 New SqlParameter("@DocCreatedDate", myFile.CreationTime), _
                                                                 New SqlParameter("@DocModifiedDate", myFile.LastWriteTime), _
                                                                 New SqlParameter("@DocFileSize", myFile.Length), _
                                                                 New SqlParameter("@DocFileName", Path.GetFileNameWithoutExtension(myFile.Name)), _
                                                                 New SqlParameter("@DocExtension", myFile.Extension))
                                            numFiles += 1


                                        Catch ex As Exception
                                            Continue For
                                        End Try
                                    Next

                                Catch ex As Exception
                                    Continue For
                                End Try
                            Next
                        Catch UnAuthDir As UnauthorizedAccessException
                            Console.WriteLine("UnAuthDir: {0}", UnAuthDir.Message)
                        End Try
                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For
                    End Try
                Next iteration_row
            End If

            'Update status listbox
            ShowUserMessage(strObjName, numFiles.ToString & " Files Processed", , True)
        End Using

        Try
            Dim TempCommand_2 As SqlCommand
            TempCommand_2 = cnDBPM.CreateCommand()
            TempCommand_2.CommandText = "sp_JobDocsRunAfterSearch"
            TempCommand_2.ExecuteNonQuery()
        Catch ex As Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
        End Try






    End Sub

    Private Sub cmdPutJobInvoicesIn_Click(sender As Object, e As EventArgs) Handles cmdPutJobInvoicesIn.Click

       
        If booQBRefreshInProgress Then Exit Sub
        If gstrCompany = "" Then Exit Sub

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "PutJobInvoicesIn" '"SUBNAME"

        Debug.WriteLine(Now.ToString & " : PutJobInvoicesIn.Click")

        Try
            PutJobInvoicesIn()
        Catch ex As Exception
            HaveError(strObjName, strSubName & "PutJobInvoicesIn", Information.Err.Number, ex.Message, Information.Err.Source, "", "")
        End Try

    End Sub


    Private Sub btnRunDBPMProcesses_Click(sender As Object, e As EventArgs) Handles btnRunDBPMProcesses.Click


        ShowUserMessage("Run All Processes", "Program started. Auto-running Processes...", "Program started...", True)
        RunDBPMProcesses()
    End Sub


    Private Sub chkPauseForErrors_CheckedChanged(sender As Object, e As EventArgs) Handles chkPauseForErrors.CheckedChanged
        Select Case chkPauseForErrors.Checked
            Case True
                gbooVerboseErrors = True
            Case Else
                gbooVerboseErrors = False
        End Select
    End Sub

    Private Sub chkSeeProcessing_CheckedChanged(sender As Object, e As EventArgs) Handles chkSeeProcessing.CheckedChanged
        Select Case chkSeeProcessing.Checked
            Case True
                gShowProcessing = True
            Case Else
                gShowProcessing = False
        End Select
    End Sub
End Class