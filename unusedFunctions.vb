Public Class unusedFunctions



    Public Sub DocFinder()
        'COPIED OUT OF frmMain.vb
        'Not used currently
        Exit Sub

        Dim strShipInfo_JOB_DocsToFindSQL, strTemp As String

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "DocFinder" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub


        CheckDataConnection(cnDBPM)

        strShipInfo_JOB_DocsToFindSQL = "SELECT * FROM JOB_DocsToFind WHERE IsActive = '1' --ORDER BY SearchPathN"

        'FOR PART 1_ - Get records from JOB_DocsToFind
        Debug.WriteLine("List1_JOB_DocsToFind")
        Dim rs1_JOB_DocsToFind As DataSet
        Dim str1_JOB_DocsToFindSQL, str1_JOB_DocsToFindRow, str1_SearchPathN, str1_CreatedDate, str1_CreatedBy, str1_CreatedOnComputer, str1_SearchTitle, str1_CreatedLead, str1_SearchPath, str1_SearchExtention, str1_IsActive As String
        'This routine gets the 1_JOB_DocsToFind from the database according to the selection in str1_JOB_DocsToFindSQL.
        'It then puts those 1_JOB_DocsToFind in the list box

        'For insert
        Dim str2_CreatedDate, str2_CreatedBy, str2_CreatedOnComputer, str2_DocTitle, str2_DocPath, str2_DocCreatedDate, str2_DocModifiedDate, str2_DocFileSize, str2_DocCreator, str2_DocFileName, str2_DocExtension As String
        'This routine inserts data into JOB_DocsFound table.

        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        Me.lstConversionProgress.AddItem("")
        Me.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "   Running DocFinder")
        ListBoxHelper.SetSelected(Me.lstConversionProgress, Me.lstConversionProgress.Items.Count - 1, True)
        Me.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Running DocFinder"
        Me.lblStatus.Text = "Starting DocFinder"
        Application.DoEvents()


        'New recordset
        rs1_JOB_DocsToFind = New DataSet()
        str1_JOB_DocsToFindSQL = "SELECT TOP 100 * FROM JOB_DocsToFind"
        str1_JOB_DocsToFindSQL = "SELECT * FROM JOB_DocsToFind WHERE IsActive = '1' --ORDER BY SearchPathN"
        Debug.WriteLine(str1_JOB_DocsToFindSQL)
        Dim adap As SqlDataAdapter = New SqlDataAdapter(str1_JOB_DocsToFindSQL, cnDBPM)
        rs1_JOB_DocsToFind.Tables.Clear()
        adap.Fill(rs1_JOB_DocsToFind)
        Dim strSQL, strDirCommand As String
        Dim retval As Integer
        Dim strLineFromFile, strPathFromFile, strNow As String
        Dim strSQL1, strSQL2, strTableInsert As String
        If rs1_JOB_DocsToFind.Tables(0).Rows.Count > 0 Then

            'First clear the table
            strSQL = "DELETE FROM JOB_DocsFound"
            'Stop
            Debug.WriteLine(strSQL)
            Dim TempCommand As SqlCommand
            TempCommand = cnDBPM.CreateCommand()
            TempCommand.CommandText = strSQL
            TempCommand.ExecuteNonQuery()

            For Each iteration_row As DataRow In rs1_JOB_DocsToFind.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs1_JOB_DocsToFind.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Me.lblListboxStatus.Text = "Processing Record " & rs1_JOB_DocsToFind.Tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs1_JOB_DocsToFind.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str1_SearchPathN = ""
                str1_CreatedDate = ""
                str1_CreatedBy = ""
                str1_CreatedOnComputer = ""
                str1_SearchTitle = ""
                str1_CreatedLead = ""
                str1_SearchPath = ""
                str1_SearchExtention = ""
                str1_IsActive = ""

                'get the columns from the database
                If iteration_row("SearchPathN") <> "" Then str1_SearchPathN = iteration_row("SearchPathN")
                If iteration_row("CreatedDate") <> "" Then str1_CreatedDate = iteration_row("CreatedDate")
                If iteration_row("CreatedBy") <> "" Then str1_CreatedBy = iteration_row("CreatedBy")
                If iteration_row("CreatedOnComputer") <> "" Then str1_CreatedOnComputer = iteration_row("CreatedOnComputer")
                If iteration_row("SearchTitle") <> "" Then str1_SearchTitle = iteration_row("SearchTitle")
                If iteration_row("CreatedLead") <> "" Then str1_CreatedLead = iteration_row("CreatedLead")
                If iteration_row("SearchPath") <> "" Then str1_SearchPath = iteration_row("SearchPath")
                If iteration_row("SearchExtention") <> "" Then str1_SearchExtention = iteration_row("SearchExtention")
                If iteration_row("IsActive") <> "" Then str1_IsActive = iteration_row("IsActive")

                'Strip quote character out of strings
                str1_SearchPathN = str1_SearchPathN.Replace("'"c, "`"c)
                str1_CreatedDate = str1_CreatedDate.Replace("'"c, "`"c)
                str1_CreatedBy = str1_CreatedBy.Replace("'"c, "`"c)
                str1_CreatedOnComputer = str1_CreatedOnComputer.Replace("'"c, "`"c)
                str1_SearchTitle = str1_SearchTitle.Replace("'"c, "`"c)
                str1_CreatedLead = str1_CreatedLead.Replace("'"c, "`"c)
                str1_SearchPath = str1_SearchPath.Replace("'"c, "`"c)
                str1_SearchExtention = str1_SearchExtention.Replace("'"c, "`"c)
                str1_IsActive = str1_IsActive.Replace("'"c, "`"c)

                '        'Strip colon character out of strings
                '        str1_SearchPathN = Replace(str1_SearchPathN, ":", ";")
                '        str1_CreatedDate = Replace(str1_CreatedDate, ":", ";")
                '        str1_CreatedBy = Replace(str1_CreatedBy, ":", ";")
                '        str1_CreatedOnComputer = Replace(str1_CreatedOnComputer, ":", ";")
                '        str1_SearchTitle = Replace(str1_SearchTitle, ":", ";")
                '        str1_CreatedLead = Replace(str1_CreatedLead, ":", ";")
                '        str1_SearchPath = Replace(str1_SearchPath, ":", ";")
                '        str1_SearchExtention = Replace(str1_SearchExtention, ":", ";")
                '        str1_IsActive = Replace(str1_IsActive, ":", ";")

                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str1_JOB_DocsToFindRow = "" & _
                                         Strings.Left(str1_SearchTitle & "                  ", 20) & "   " & _
                                         Strings.Left(str1_SearchExtention & "                  ", 8) & "   " & _
                                         Strings.Left(str1_SearchPath & "                  ", 60) & "   " & _
                                         Strings.Left(str1_CreatedLead & "                  ", 18) & "   " & _
                                         Strings.Left(str1_SearchPathN & "                  ", 18) & "   " & _
                                         Strings.Left(str1_CreatedDate & "                  ", 18) & "   " & _
                                         Strings.Left(str1_CreatedBy & "                  ", 18) & "   " & _
                                         Strings.Left(str1_CreatedOnComputer & "                  ", 18) & "   " & _
                                         Strings.Left(str1_IsActive & "                  ", 18) & "   " & _
                                         "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs1_JOB_DocsToFind.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs1_JOB_DocsToFind.Tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs1_JOB_DocsToFind.Tables(0).Rows.Count))
                If Me.chkSeeProcessing.CheckState = CheckState.Checked Then
                    Me.lstConversionProgress.AddItem("1_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str1_JOB_DocsToFindRow)
                    'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str1_SearchPathN
                    ListBoxHelper.SetSelected(Me.lstConversionProgress, Me.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                'Poll the Path via Batch Cmd.  Output to specific filename
                strDirCommand = "CMD /c DIR " & Strings.Chr(34).ToString() & str1_SearchPath & "\*" & str1_SearchExtention & Strings.Chr(34).ToString() & " /Q /S > C:\DirOutputFile.txt"
                Debug.WriteLine(strDirCommand)
                'Stop

                'OLD
                'Dim intX As Integer
                'intX = Shell(strDirCommand)

                'NEW  -Waits for the CMD window to close before continuing
                'retval = ExecCmd("notepad.exe")
                retval = ExecCmd(strDirCommand)
                'MsgBox "Process Finished, Exit Code " & retval
                'Stop




                'Parse the output file into JOB_DocsFound table

                strNow = DateTimeHelper.ToString(DateTime.Now)

                'open <output file> for read
                'Open "C:\DirOutputFile.txt" For Append As #1
                'Open "C:\DirOutputFile.txt" For Output As #1
                FileSystem.FileOpen(1, "C:\DirOutputFile.txt", OpenMode.Input)

                'Print #1, "UPDATE " & strChosenCallNumber & "  " & strGlobalOperatorID & "           " & strCallDateM & "   " & txtCallTimeM & "   " & txtCallerM & "   " & strCommentsM & strCommentsM2 & strCommentsM3 & strCommentsM4
                'Input #1, s

                'For intI = 1 To XX '<numfilelines>
                Do While Not FileSystem.EOF(1)

                    'MyChar = Input(1, #1)   ' Read a character.
                    'strLineFromFile = Input(1000, #1)
                    'Input #1, strLineFromFile
                    strLineFromFile = FileSystem.LineInput(1)

                    Debug.WriteLine(strLineFromFile)


                    'frmMain.lblListboxStatus.Caption = "Processing Record " & rs1_JOB_DocsToFind.tables(0).Rows.IndexOf(iteration_row) & " of " & rs1_JOB_DocsToFind.RecordCount & ""
                    Me.lblListboxStatus.Text = strLineFromFile '"Processing Record " & rs1_JOB_DocsToFind.tables(0).Rows.IndexOf(iteration_row) & " of " & rs1_JOB_DocsToFind.RecordCount & ""
                    Application.DoEvents()


                    'If Right(Trim(strLineFromFile), 12) = "Previous Job" Then Stop


                    'Load line intI into strLineFromFile
                    Dim dbNumericTemp As Double
                    If Strings.Mid(strLineFromFile, 1, 1) = "" Then 'skip
                        'Skip
                        'Stop
                    ElseIf Strings.Mid(strLineFromFile, 2, 6) = "Volume" Then
                        'Skip
                        'Stop
                    ElseIf Strings.Mid(strLineFromFile, 2, 9) = "Directory" Then
                        'Capture Path
                        strPathFromFile = Strings.Mid(strLineFromFile, 15, 1000)
                        'Stop
                    ElseIf Double.TryParse(Strings.Mid(strLineFromFile, 1, 2), NumberStyles.Float, CultureInfo.CurrentCulture.NumberFormat, dbNumericTemp) Then  'Date of file
                        'Split and put pieces into table

                        'Vars to fill
                        'str2_DocFoundN = ""
                        str2_CreatedDate = strNow
                        str2_CreatedBy = gstrUserName
                        str2_CreatedOnComputer = gstrComputerName
                        str2_DocTitle = str1_SearchTitle
                        str2_DocPath = strPathFromFile
                        str2_DocCreatedDate = Strings.Mid(strLineFromFile, 1, 20)
                        str2_DocModifiedDate = ""
                        str2_DocFileSize = Strings.Mid(strLineFromFile, 21, 18).Trim()
                        str2_DocCreator = Strings.Mid(strLineFromFile, 40, 22).Trim() 'str1_CreatedLead

                        'str2_DocFileName = Mid(strLineFromFile, 63, 1000) 'WITH EXTENSION!
                        strTemp = Strings.Mid(strLineFromFile, 63, 1000) 'AND EXTENSION!
                        str2_DocFileName = Strings.Mid(strTemp, 1, Strings.Len(strTemp) - Strings.Len(str1_SearchExtention)) 'NO EXTENSION!

                        str2_DocExtension = str1_SearchExtention

                        '' strings to use
                        'str1_SearchPath = ""


                        'FIX
                        'replace apostrophy in str2_DocPath
                        str2_DocPath = str2_DocPath.Replace("'"c, "`"c)
                        str2_DocFileName = str2_DocFileName.Replace("'"c, "`"c)


                        'dim SQL strings

                        'Build the SQL string
                        strSQL1 = "INSERT INTO JOB_DocsFound " & Environment.NewLine & _
                                  "   ( CreatedDate " & Environment.NewLine & _
                                  "   , CreatedBy " & Environment.NewLine & _
                                  "   , CreatedOnComputer " & Environment.NewLine & _
                                  "   , DocTitle " & Environment.NewLine & _
                                  "   , DocPath " & Environment.NewLine & _
                                  "   , DocCreatedDate " & Environment.NewLine & _
                                  "   , DocModifiedDate " & Environment.NewLine & _
                                  "   , DocFileSize " & Environment.NewLine & _
                                  "   , DocCreator " & Environment.NewLine & _
                                  "   , DocFileName " & Environment.NewLine & _
                                  "   , DocExtension ) " & Environment.NewLine
                        strSQL2 = "VALUES " & Environment.NewLine & _
                                  "   ( '" & str2_CreatedDate & "'  --CreatedDate" & Environment.NewLine & _
                                  "   , '" & str2_CreatedBy & "'  --CreatedBy" & Environment.NewLine & _
                                  "   , '" & str2_CreatedOnComputer & "'  --CreatedOnComputer" & Environment.NewLine & _
                                  "   , '" & str2_DocTitle & "'  --DocTitle" & Environment.NewLine & _
                                  "   , '" & str2_DocPath & "'  --DocPath" & Environment.NewLine & _
                                  "   , '" & str2_DocCreatedDate & "'  --DocCreatedDate" & Environment.NewLine & _
                                  "   , NULL --'" & str2_DocModifiedDate & "'  --DocModifiedDate" & Environment.NewLine & _
                                  "   , '" & str2_DocFileSize & "'  --DocFileSize" & Environment.NewLine & _
                                  "   , '" & str2_DocCreator & "'  --DocCreator" & Environment.NewLine & _
                                  "   , '" & str2_DocFileName & "'  --DocFileName" & Environment.NewLine & _
                                  "   , '" & str2_DocExtension & "' ) --DocExtension" & Environment.NewLine

                        'Combine the strings
                        strTableInsert = strSQL1 & strSQL2 '& strSQL3 & strSQL4
                        'Debug.Print strTableInsert

                        'Execute the insert
                        Dim TempCommand_2 As SqlCommand
                        TempCommand_2 = cnDBPM.CreateCommand()
                        TempCommand_2.CommandText = strTableInsert
                        TempCommand_2.ExecuteNonQuery()

                    Else
                        'see the line
                        'Stop
                        Debug.WriteLine(strLineFromFile)
                    End If

                    'Next intI
                Loop

                FileSystem.FileClose(1)




            Next iteration_row
        Else
            If Me.chkSeeProcessing.CheckState = CheckState.Checked Then
                'frmMain.lstConversionProgress.AddItem txtTypeRadNum
                'frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            End If
        End If

        'Update status listbox
        Me.lblListboxStatus.Text = CStr(rs1_JOB_DocsToFind.Tables(0).Rows.Count) & " Records Processed"
        If Me.chkSeeProcessing.CheckState = CheckState.Checked Then
            'frmMain.lstConversionProgress.AddItem ""
            'frmMain.lstConversionProgress.AddItem ""
            ListBoxHelper.SetSelected(Me.lstConversionProgress, Me.lstConversionProgress.Items.Count - 1, True)
        End If


        'Fix stuff (delete "Previous Job" entries)
        Dim TempCommand_3 As SqlCommand
        TempCommand_3 = cnDBPM.CreateCommand()
        TempCommand_3.CommandText = "EXEC sp_JobDocsRunAfterSearch"
        TempCommand_3.ExecuteNonQuery()


        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            rs1_JOB_DocsToFind = Nothing

            Exit Sub



        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("ERROR IN DOCFINDER")
        End Try

    End Sub

    Public Sub DeleteOldFiles()
        'COPIED FROM frmMain.vb



        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "DeleteOldBackups" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub


        'FOR PART 1_ - Get records from DBPM_FilesToDelete
        Debug.WriteLine("List1_DBPM_FilesToDelete")
        'This routine gets the 1_DBPM_FilesToDelete from the database according to the selection in str1_DBPM_FilesToDeleteSQL.
        'It then puts those 1_DBPM_FilesToDelete in the list box



        'For insert
        Dim str2_CreatedDate, str2_CreatedBy, str2_CreatedOnComputer, str2_DocTitle, str2_DocPath, str2_DocCreatedDate, str2_DocModifiedDate, str2_DocFileSize, str2_DocCreator, str2_DocFileName, str2_DocExtension As String
        'This routine inserts data into DBPM_FilesFound table.

        'Show what's processing
        Me.lstConversionProgress.AddItem("")
        Me.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "   Deleting Old Files")
        ListBoxHelper.SetSelected(Me.lstConversionProgress, Me.lstConversionProgress.Items.Count - 1, True)
        Me.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Deleting Old Files"
        Me.lblStatus.Text = "Starting DocFinder"
        Application.DoEvents()

        Using Sql As New SQLHelper(gstrSQLConnectionString)
            Dim str1_DBPM_FilesToDeleteSQL, str1_SearchPathN, str1_CreatedDate, str1_CreatedBy, str1_CreatedOnComputer, str1_SearchTitle, str1_CreatedLead, str1_SearchPath, str1_SearchExtention, str1_IsActive, str1_DaysToKeep As String
            str1_DBPM_FilesToDeleteSQL = "SELECT * FROM DBPM_FilesToDelete WHERE IsActive = '1' ORDER BY SearchPathN"


            Using rs1_DBPM_FilesToDelete As SqlDataReader = Sql.ExecuteReader(CommandType.Text, str1_DBPM_FilesToDeleteSQL)

                Dim strSQL As String = ""
                Dim strDirCommand As String = ""
                Dim retval As Integer
                Dim strLineFromFile As String = ""
                Dim strPathFromFile As String = ""
                Dim strNow As String = ""
                Dim rowCount As Integer = 0, strTemp As String = ""
                Dim strSQL1 As String = ""
                Dim strSQL2 As String = ""
                Dim strTableInsert As String = ""
                If rs1_DBPM_FilesToDelete.HasRows Then

                    'First clear the table?
                    strSQL = "DELETE FROM DBPM_FilesFound"
                    Sql.ExecuteSQL(strSQL)


                    While rs1_DBPM_FilesToDelete.Read
                        rowCount += 1
                        'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs1_DBPM_FilesToDelete.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                        Me.lblListboxStatus.Text = "Processing Record: " & rowCount.ToString
                        Application.DoEvents()

                        'get the columns from the database
                        str1_SearchPathN = NCStr(rs1_DBPM_FilesToDelete("SearchPathN")).Replace("'"c, "`"c)
                        str1_CreatedDate = NCStr(rs1_DBPM_FilesToDelete("CreatedDate")).Replace("'"c, "`"c)
                        str1_CreatedBy = NCStr(rs1_DBPM_FilesToDelete("CreatedBy")).Replace("'"c, "`"c)
                        str1_CreatedOnComputer = NCStr(rs1_DBPM_FilesToDelete("CreatedOnComputer")).Replace("'"c, "`"c)
                        str1_SearchTitle = NCStr(rs1_DBPM_FilesToDelete("SearchTitle")).Replace("'"c, "`"c)
                        str1_CreatedLead = NCStr(rs1_DBPM_FilesToDelete("CreatedLead")).Replace("'"c, "`"c)
                        str1_SearchPath = NCStr(rs1_DBPM_FilesToDelete("SearchPath")).Replace("'"c, "`"c)
                        str1_SearchExtention = NCStr(rs1_DBPM_FilesToDelete("SearchExtention")).Replace("'"c, "`"c)
                        str1_IsActive = NCStr(rs1_DBPM_FilesToDelete("IsActive")).Replace("'"c, "`"c)
                        str1_DaysToKeep = NCStr(rs1_DBPM_FilesToDelete("DaysToKeep")).Replace("'"c, "`"c)


                        'Put the information together into a string
                        'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                        'Dim str1_DBPM_FilesToDeleteRow As String
                        'str1_DBPM_FilesToDeleteRow = "" & _
                        'Strings.Left(str1_SearchTitle & "                  ", 20) & "   " & _
                        'Strings.Left(str1_SearchExtention & "                  ", 8) & "   " & _
                        'Strings.Left(str1_SearchPath & "                  ", 60) & "   " & _
                        'Strings.Left(str1_CreatedLead & "                  ", 18) & "   " & _
                        'Strings.Left(str1_SearchPathN & "                  ", 18) & "   " & _
                        'Strings.Left(str1_CreatedDate & "                  ", 18) & "   " & _
                        'Strings.Left(str1_CreatedBy & "                  ", 18) & "   " & _
                        'Strings.Left(str1_CreatedOnComputer & "                  ", 18) & "   " & _
                        'Strings.Left(str1_IsActive & "                  ", 18) & "   " & _
                        '"" & Strings.Chr(9)

                        'put the line in the listbox
                        If Me.chkSeeProcessing.CheckState = CheckState.Checked Then
                            'Me.lstConversionProgress.AddItem("1_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str1_DBPM_FilesToDeleteRow)
                            'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str1_SearchPathN
                            ListBoxHelper.SetSelected(Me.lstConversionProgress, Me.lstConversionProgress.Items.Count - 1, True)
                        End If

                        'DO WORK: With each record


                        'Poll the Path via Batch Cmd.  Output to specific filename
                        strDirCommand = "CMD /c DIR " & Strings.Chr(34).ToString() & str1_SearchPath & "\*" & str1_SearchExtention & Strings.Chr(34).ToString() & " /Q /S > C:\DirOutputFile.txt"
                        Debug.WriteLine(strDirCommand)
                        'Stop

                        'OLD
                        'Dim intX As Integer
                        'intX = Shell(strDirCommand)

                        'NEW  -Waits for the CMD window to close before continuing
                        'retval = ExecCmd("notepad.exe")
                        retval = ExecCmd(strDirCommand)
                        'MsgBox "Process Finished, Exit Code " & retval
                        'Stop


                        'Parse the output file into DBPM_FilesFound table

                        strNow = DateTimeHelper.ToString(DateTime.Now)

                        'open <output file> for read
                        'Open "C:\DirOutputFile.txt" For Append As #1
                        'Open "C:\DirOutputFile.txt" For Output As #1
                        FileSystem.FileOpen(1, "C:\DirOutputFile.txt", OpenMode.Input)

                        'Print #1, "UPDATE " & strChosenCallNumber & "  " & strGlobalOperatorID & "           " & strCallDateM & "   " & txtCallTimeM & "   " & txtCallerM & "   " & strCommentsM & strCommentsM2 & strCommentsM3 & strCommentsM4
                        'Input #1, s

                        'For intI = 1 To XX '<numfilelines>
                        Do While Not FileSystem.EOF(1)

                            'MyChar = Input(1, #1)   ' Read a character.
                            'strLineFromFile = Input(1000, #1)
                            'Input #1, strLineFromFile
                            strLineFromFile = FileSystem.LineInput(1)

                            Debug.WriteLine(strLineFromFile)


                            'frmMain.lblListboxStatus.Caption = "Processing Record " & rs1_DBPM_FilesToDelete.tables(0).Rows.IndexOf(iteration_row) & " of " & rs1_DBPM_FilesToDelete.RecordCount & ""
                            Me.lblListboxStatus.Text = strLineFromFile '"Processing Record " & rs1_DBPM_FilesToDelete.tables(0).Rows.IndexOf(iteration_row) & " of " & rs1_DBPM_FilesToDelete.RecordCount & ""
                            Application.DoEvents()


                            'If Right(Trim(strLineFromFile), 12) = "Previous Job" Then Stop


                            'Load line intI into strLineFromFile
                            Dim dbNumericTemp As Double
                            If Strings.Mid(strLineFromFile, 1, 1) = "" Then 'skip
                                'Skip
                                'Stop
                            ElseIf Strings.Mid(strLineFromFile, 2, 6) = "Volume" Then
                                'Skip
                                'Stop
                            ElseIf Strings.Mid(strLineFromFile, 2, 9) = "Directory" Then
                                'Capture Path
                                strPathFromFile = Strings.Mid(strLineFromFile, 15, 1000)
                                'Stop
                            ElseIf Double.TryParse(Strings.Mid(strLineFromFile, 1, 2), NumberStyles.Float, CultureInfo.CurrentCulture.NumberFormat, dbNumericTemp) Then  'Date of file
                                'Split and put pieces into table


                                'Vars to fill
                                'str2_DocFoundN = ""
                                str2_CreatedDate = strNow
                                str2_CreatedBy = gstrUserName
                                str2_CreatedOnComputer = gstrComputerName
                                str2_DocTitle = str1_SearchTitle
                                str2_DocPath = strPathFromFile
                                str2_DocCreatedDate = Strings.Mid(strLineFromFile, 1, 20)
                                str2_DocModifiedDate = ""
                                str2_DocFileSize = Strings.Mid(strLineFromFile, 21, 18).Trim()
                                str2_DocCreator = Strings.Mid(strLineFromFile, 40, 22).Trim() 'str1_CreatedLead

                                'str2_DocFileName = Mid(strLineFromFile, 63, 1000) 'WITH EXTENSION!
                                strTemp = Strings.Mid(strLineFromFile, 63, 1000) 'AND EXTENSION!
                                str2_DocFileName = Strings.Mid(strTemp, 1, Strings.Len(strTemp) - Strings.Len(str1_SearchExtention)) 'NO EXTENSION!

                                str2_DocExtension = str1_SearchExtention


                                'FIX
                                'replace apostrophy in str2_DocPath
                                str2_DocPath = str2_DocPath.Replace("'"c, "`"c)
                                str2_DocFileName = str2_DocFileName.Replace("'"c, "`"c)


                                'NEW FOR FILE DELETE
                                'Dim strDirCommand As String
                                str1_DaysToKeep = "-" & str1_DaysToKeep

                                ''Check the date of the file
                                If CDate(str2_DocCreatedDate) > DateTime.Now.AddDays(CInt(str1_DaysToKeep)) Then

                                    'Dim strDirCommand As String
                                    ''strDirCommand = "CMD /c DIR " & Chr(34) & str1_SearchPath & "\*" & str1_SearchExtention & Chr(34) & " /Q /S > C:\DirOutputFile.txt"
                                    strDirCommand = "CMD /c DEL " & Strings.Chr(34).ToString() & str2_DocPath & "\" & str2_DocFileName & str2_DocExtension & Strings.Chr(34).ToString() & " /Q /S"
                                    Debug.WriteLine(strDirCommand)
                                    ''Stop

                                    ''NEW  -Waits for the CMD window to close before continuing
                                    'Dim retval As Long
                                    ''retval = ExecCmd("notepad.exe")
                                    retval = ExecCmd(strDirCommand)



                                    'dim SQL strings

                                    'Build the SQL string
                                    strSQL1 = "INSERT INTO DBPM_FilesFound " & Environment.NewLine & _
                                              "   ( CreatedDate " & Environment.NewLine & _
                                              "   , CreatedBy " & Environment.NewLine & _
                                              "   , CreatedOnComputer " & Environment.NewLine & _
                                              "   , DocTitle " & Environment.NewLine & _
                                              "   , DocPath " & Environment.NewLine & _
                                              "   , DocCreatedDate " & Environment.NewLine & _
                                              "   , DocModifiedDate " & Environment.NewLine & _
                                              "   , DocFileSize " & Environment.NewLine & _
                                              "   , DocCreator " & Environment.NewLine & _
                                              "   , DocFileName " & Environment.NewLine & _
                                              "   , DocExtension ) " & Environment.NewLine
                                    strSQL2 = "VALUES " & Environment.NewLine & _
                                              "   ( '" & str2_CreatedDate & "'  --CreatedDate" & Environment.NewLine & _
                                              "   , '" & str2_CreatedBy & "'  --CreatedBy" & Environment.NewLine & _
                                              "   , '" & str2_CreatedOnComputer & "'  --CreatedOnComputer" & Environment.NewLine & _
                                              "   , '" & str2_DocTitle & "'  --DocTitle" & Environment.NewLine & _
                                              "   , '" & str2_DocPath & "'  --DocPath" & Environment.NewLine & _
                                              "   , '" & str2_DocCreatedDate & "'  --DocCreatedDate" & Environment.NewLine & _
                                              "   , NULL --'" & str2_DocModifiedDate & "'  --DocModifiedDate" & Environment.NewLine & _
                                              "   , '" & str2_DocFileSize & "'  --DocFileSize" & Environment.NewLine & _
                                              "   , '" & str2_DocCreator & "'  --DocCreator" & Environment.NewLine & _
                                              "   , '" & str2_DocFileName & "'  --DocFileName" & Environment.NewLine & _
                                              "   , '" & str2_DocExtension & "' ) --DocExtension" & Environment.NewLine

                                    'Combine the strings
                                    strTableInsert = strSQL1 & strSQL2
                                    'Debug.Print strTableInsert

                                    'Execute the insert
                                    Dim TempCommand_2 As SqlCommand
                                    TempCommand_2 = cnDBPM.CreateCommand()
                                    TempCommand_2.CommandText = strTableInsert
                                    TempCommand_2.ExecuteNonQuery()

                                End If


                            Else
                                'see the line
                                'Stop
                                Debug.WriteLine(strLineFromFile)
                            End If

                            'Next intI
                        Loop

                        FileSystem.FileClose(1)

                    End While
                Else
                    If Me.chkSeeProcessing.CheckState = CheckState.Checked Then
                        'frmMain.lstConversionProgress.AddItem txtTypeRadNum
                        'frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
                    End If
                End If

                Me.lblListboxStatus.Text = rowCount.ToString & " Records Processed"

            End Using

        End Using
        'Update status listbox
        If Me.chkSeeProcessing.CheckState = CheckState.Checked Then
            ListBoxHelper.SetSelected(Me.lstConversionProgress, Me.lstConversionProgress.Items.Count - 1, True)
        End If


    End Sub


    Private Sub cmdDeleteOldFiles_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdDeleteOldFiles.Click
        'COPIED FROM frmMAIN.vb


        If booQBRefreshInProgress Then Exit Sub
        If Me.chkPauseProcessing.CheckState = CheckState.Checked Then Exit Sub

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "cmdDeleteOldBackups_Click" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Show what's processing in the listbox
        Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   Deleting Old Files")
        Me.lstConversionProgress.AddItem("" & Now & "   Running DocFinder")
        Me.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Deleting Old Files"
        Me.lblStatus.Text = "Deleting Old Files"
        Application.DoEvents()

        Try
            booQBRefreshInProgress = True
            'CheckDataConnection(cnDBPM)
            DeleteOldFiles()

            'show finished
            Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   Finished Deleting Old Files")
            Me.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "   Finished Deleting Old Files")
            Me.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(Me.lstConversionProgress, Me.lstConversionProgress.Items.Count - 1, True)
            Me.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Finished Deleting Old Files"
            Me.lblStatus.Text = "Finished Deleting Old Files"
            Application.DoEvents()

            'reset flag
            booQBRefreshInProgress = False

        Catch ex As Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "")
        End Try


    End Sub


    Private Sub cmdRunDocFinder_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs)
        'COPIED FROM frmMain.vb

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "frmMain" '"OBJNAME"
        Dim strSubName As String = "cmdRunDocFinder_Click" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub
        If Me.chkPauseProcessing.CheckState = CheckState.Checked Then Exit Sub

        If Not booQBRefreshInProgress Then
            'Show what's processing in the listbox
            Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   Running DocFinder")
            Me.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Running DocFinder"
            Me.lblStatus.Text = "Starting DocFinder"
            Application.DoEvents()

            booQBRefreshInProgress = True
            Try
                DocFinder()
            Catch ex As Exception
                HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "")
            End Try
            booQBRefreshInProgress = False
            Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   Finished DocFinder")
            Me.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "   Finished DocFinder")
            Me.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(Me.lstConversionProgress, Me.lstConversionProgress.Items.Count - 1, True)
            Me.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Finished DocFinder"
            Me.lblStatus.Text = "Finished DocFinder"
            Application.DoEvents()
        End If

    End Sub


    Public Sub OpenConnectionMaxLL()


        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modGen" '"OBJNAME"
        Dim strSubName As String = "OpenConnectionMaxLL" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then On Error GoTo ErrorFunc
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then Resume Next Else Exit Sub
RunCode:



        'Purpose: Open connection to database to be used throughout program

        'On Error GoTo SubError

        'open a new connection
        cnMaxLL = New SqlConnection()
        'UPGRADE_ISSUE: (2064) ADODB.CursorLocationEnum property CursorLocationEnum.adUseClient was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
        'UPGRADE_ISSUE: (2064) ADODB.Connection property cnMaxLL.CursorLocation was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
        'cnMaxLL.CursorLocation = adUseClient
        'UPGRADE_ISSUE: (2064) ADODB.Connection property cnMaxLL.CommandTimeout was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
        'cnMaxLL.CommandTimeout = 120
        'Debug.Print cnMaxLL.CommandTimeout

        'cnMaxLL.Open "Provider=SQLOLEDB;Server=AGFCCOMM5;uid=MaxUsers;pwd=thisthingisgreat;database=DrummondPrinting;" ', , , adAsyncConnect
        'cnMaxLL.Open "Provider=SQLOLEDB;Server=Server02;uid=MaxUsers;pwd=thisthingisgreat;database=DrummondPrinting;" ', , , adAsyncConnect
        'cnMaxLL.Open "Provider=SQLOLEDB;Server=Server02;uid=MaxUsers;pwd=thisthingisgreat;database=DrumTest;" ', , , adAsyncConnect
        'cnMaxLL.Open "Provider=SQLOLEDB;Server=Server02;uid=MASTER;pwd=control;database=DrumTest;" ', , , adAsyncConnect

        'cnMaxLL.Open "Provider=SQLOLEDB;Server=Server02;uid=MASTER;pwd=control;database=ZBTest;" ', , , adAsyncConnect
        'cnMaxLL.Open "Provider=SQLOLEDB;Server=Server02;uid=MASTER;pwd=control;database=DrummondPrinting;" ', , , adAsyncConnect
        'DoEvents

        If gstrCompany = "DrummondPrinting" Then
            'UPGRADE_TODO: (7010) The connection string must be verified to fullfill the .NET data provider connection string requirements. More Information: http://www.vbtonet.com/ewis/ewi7010.aspx
            cnMaxLL.ConnectionString = "Provider=SQLOLEDB;Server=" & gstrServerName & ";uid=" & sqlMaxUser & ";pwd=" & sqlMaxPass & ";database=DrummondLoadLeads;"
            cnMaxLL.Open() ', , , adAsyncConnect
        ElseIf gstrCompany = "FrazzledAndBedazzled" Then
            'UPGRADE_TODO: (7010) The connection string must be verified to fullfill the .NET data provider connection string requirements. More Information: http://www.vbtonet.com/ewis/ewi7010.aspx
            cnMaxLL.ConnectionString = "Provider=SQLOLEDB;Server=" & gstrServerName & ";uid=" & sqlMaxUser & ";pwd=" & sqlMaxPass & ";database=FrazzledLoadLeads;"
            cnMaxLL.Open() ', , , adAsyncConnect
        End If

        '  'Or just do this
        '  'cnMaxLL.Open "Provider=SQLOLEDB;Server=Server02;uid=MASTER;pwd=control;database=" & gstrCompany & ";" ', , , adAsyncConnect


        If cnMaxLL.State = ConnectionState.Open Then
            frmMain.lblStatus.Text = "MaxLL Connection Established " & DateTimeHelper.ToString(DateTime.Now)
            Debug.WriteLine("MaxLL Connection Established")
            Application.DoEvents()
        End If

        Exit Sub


        'MsgBox "<<modConnections  OpenConnectionMaxLL>> " & Err.Description
        frmMain.lblStatus.Text = "MaxLL Connection Failed.  Attempting Reconnection " & DateTimeHelper.ToString(DateTime.Now)
        Debug.WriteLine("MaxLL Connection Failed")
        Debug.WriteLine("<<modGen  OpenConnectionMaxLL>> " & Information.Err().Description)
        Application.DoEvents()

        Resume Next

    End Sub


    'THIS WAS IN frmMAIN

    'Private Const NORMAL_PRIORITY_CLASS As Integer = &H20
    'Private Const INFINITE As Integer = -1
    'Public Sub New()
    '    MyBase.New()
    '    If m_vb6FormDefInstance Is Nothing Then
    '        If m_InitializingDefInstance Then
    '            m_vb6FormDefInstance = Me
    '        Else
    '            Try
    '                'For the start-up form, the first instance created is the default instance.
    '                If System.Reflection.Assembly.GetExecutingAssembly().EntryPoint <> Nothing AndAlso System.Reflection.Assembly.GetExecutingAssembly().EntryPoint.DeclaringType = Me.GetType() Then
    '                    m_vb6FormDefInstance = Me
    '                End If

    '            Catch
    '            End Try
    '        End If
    '    End If
    '    'This call is required by the Windows Form Designer.
    '    isInitializingComponent = True
    '    InitializeComponent()
    '    isInitializingComponent = False
    '    ReLoadForm(False)
    'End Sub



    'Public Function ExecCmd(ByRef cmdline As String) As Integer
    '    Dim ret As Integer
    '    Dim proc As New DBPM_Server_4_28_2015Support.UnsafeNative.Structures.PROCESS_INFORMATION()
    '    Dim start As DBPM_Server_4_28_2015Support.UnsafeNative.Structures.STARTUPINFO = DBPM_Server_4_28_2015Support.UnsafeNative.Structures.STARTUPINFO.CreateInstance()

    '    ' Initialize the STARTUPINFO structure:
    '    'UPGRADE_WARNING: (2081) Len has a new behavior. More Information: http://www.vbtonet.com/ewis/ewi2081.aspx
    '    start.cb = Marshal.SizeOf(start)

    '    ' Start the shelled application:
    '    ret = DBPM_Server_4_28_2015Support.SafeNative.kernel32.CreateProcessA(Nothing, cmdline, 0, 0, 1, NORMAL_PRIORITY_CLASS, 0, Nothing, start, proc)

    '    ' Wait for the shelled application to finish:
    '    ret = DBPM_Server_4_28_2015Support.SafeNative.kernel32.WaitForSingleObject(proc.hProcess, INFINITE)
    '    DBPM_Server_4_28_2015Support.SafeNative.kernel32.GetExitCodeProcess(proc.hProcess, ret)
    '    DBPM_Server_4_28_2015Support.SafeNative.kernel32.CloseHandle(proc.hThread)
    '    DBPM_Server_4_28_2015Support.SafeNative.kernel32.CloseHandle(proc.hProcess)
    '    Return ret
    'End Function
    ''FINISH FOR ExecCmd() FUNCTION:



    'Private isInitializingComponent As Boolean
    'Private Sub chkPauseProcessing_CheckStateChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles chkPauseProcessing.CheckStateChanged
    '    If isInitializingComponent Then
    '        Exit Sub
    '    End If

    '    lblPaused.Visible = chkPauseProcessing.CheckState = CheckState.Checked

    'End Sub


    'from haveerrors...




    'COMMON ERROR:
    'TAMI    12/03/2008 2:48:17 PM   frmFindCustomerG3   HitMatchSearchEngine    3704    Operation is not allowed when the object is closed. DBPM    1.0.666 1/3/2009 4:41:00 PM 2   Server02    C:\Program Files\DBPM\DBPM.exe  High                26124   95987                       vw_DBPM_Errors  TAMIXP
            If strObjName = "frmFindCustomerG3" And strSubName = "HitMatchSearchEngine" And strErrNumber = "3704" And strErrDescription = "Operation is not allowed when the object is closed." Then
    'UPGRADE_TODO: (1067) Member lblLoadingDealerHistory is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
                frmFindCustomerG3.lblLoadingDealerHistory = ""
    'MsgBox "This Procedure was aborted due to errors. " & vbCrLf & "Try it again or restart the program if errors persist." & vbCrLf & ""
                MessageBox.Show("This Procedure was aborted due to errors. " & Environment.NewLine & "Try it again." & Environment.NewLine & "", Application.ProductName)
                result = "ES"
                gintErrorCounter_AllErrors = 0
                gintErrorCounter_SameError = 0
                gintErrorCounter_SameSub = 0
    'frmMessage.Hide()
                Application.DoEvents()
            End If




    'COMMON ERROR:
    'OUT OF MEMORY
    'If UCase(strErrDescription) = "OBJECT VARIABLE OR WITH BLOCK VARIABLE NOT SET" Then
    'If UCase(strErrDescription) = "BAD FILE NAME OR NUMBER" Then
    'If UCase(strErrDescription) = "OBJECT REQUIRED" Then
    Dim intX As Integer
            If strErrDescription.ToUpper() = "OUT OF MEMORY." Then

    'Exit the offending sub
    'HaveError = "RN"
                result = "ES"
                gintErrorCounter_AllErrors = 0
                gintErrorCounter_SameError = 0
                gintErrorCounter_SameSub = 0
    'frmMessage.Hide
                Application.DoEvents()

    'Close QB?
                CloseConnectionQB()

    'Kill QB?   'QB32W.exe ?    C:\pskill


    'Kill QODBC?   'FQ???.exe ?


    'Restart QB?


    'Start the DBPM_Server_Restarter (Shell it)
    'Temp: Start new DBPM_Server instance
                strPathAndFilenameOfDBPMS = Strings.Chr(34).ToString() & "C:\Program Files\DBPM\DBPM_Server.exe" & Strings.Chr(34).ToString()
                Debug.WriteLine(strPathAndFilenameOfDBPMS)
    'intX = Shell("C:\Program Files\DBPM")
    'intX = Shell(Chr(34) & "C:\Program Files\DBPM\DBPM_Server.exe" & Chr(34))
    'UPGRADE_TODO: (7005) parameters (if any) must be set using the Arguments property of ProcessStartInfo More Information: http://www.vbtonet.com/ewis/ewi7005.aspx
                intX = Process.Start(strPathAndFilenameOfDBPMS).Id

    'Log the event
                FileSystem.FileOpen(1, "C:\DBPMS_EventLog.txt", OpenMode.Append)
                FileSystem.PrintLine(1, "|" & strNow & _
                                     "|" & gstrUserName & _
                                     "|" & gstrComputerName & _
                                     "|" & strObjName & _
                                     "|" & strSubName & _
                                     "|" & strErrNumber & _
                                     "|" & strErrDescription & _
                                     "|" & strErrSource & _
                                     "|" & strJobNum & _
                                     "|" & strCust & _
                                     "|" & gstrApplicationName & _
                                     "|" & gstrVersionNum & _
                                     "|" & DateTimeHelper.ToString(gdteVersionExpirationDate) & _
                                     "|" & gstrInstancePID & _
                                     "|" & gstrServerName & _
                                     "|" & gstrModuleFileName & _
                                     "|" & gstrScreenResolution & _
                                     "|" & strTask & _
                                     "|" & strInformation & _
                                     "|" & strTableName & _
                                     "|" & strKeyNName & _
                                     "|" & strKeyN & _
                                     "|" & strSalesOrderN & _
                                     "|" & strStepN & "|")
                FileSystem.FileClose(1)


    'End this instance
                gbooEndProgram = True

            End If



End Class
Public Sub UpdateQBReceivePaymentLine(ByRef strCustomerRefFullName As String)
    'FINISHED WITH FIRST RUN_THROUGH


    'Permission and ErrorHandling          (Auto built)
    Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
    Dim strSubName As String = "UpdateQBReceivePaymentLine" '"SUBNAME"

    'Check permission to run
    If Not HavePermission(strObjName, strSubName) Then Exit Sub

    ShowUserMessage(strSubName, "Updating QB with Receive Payment Lines", strSubName)

    ''This routine gets the 1MaxOfCopy_QBTable from the database according to the selection in str1MaxOfCopy_QBTableSQL.
    ''It then puts those 1MaxOfCopy_QBTable in the list box

    'FOR PART 2SrcQB_ - Get records from QB_ReceivePaymentLine
    Debug.WriteLine("List2SrcQB_QB_ReceivePaymentLine")

    Dim str2SrcQB_QB_ReceivePaymentLineSQL, str2SrcQB_QB_ReceivePaymentLineRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_TotalAmount, str2SrcQB_PaymentMethodRefListID, str2SrcQB_PaymentMethodRefFullName, str2SrcQB_Memo, str2SrcQB_DepositToAccountRefListID, str2SrcQB_DepositToAccountRefFullName, str2SrcQB_CreditCardTxnInfoInputCreditCardNumber, str2SrcQB_CreditCardTxnInfoInputExpirationMonth, str2SrcQB_CreditCardTxnInfoInputExpirationYear, str2SrcQB_CreditCardTxnInfoInputNameOnCard, str2SrcQB_CreditCardTxnInfoInputCreditCardAddress, str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode, str2SrcQB_CreditCardTxnInfoInputCommercialCardCode, str2SrcQB_CreditCardTxnInfoResultResultCode, str2SrcQB_CreditCardTxnInfoResultResultMessage, str2SrcQB_CreditCardTxnInfoResultCreditCardTransID, str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber, str2SrcQB_CreditCardTxnInfoResultAuthorizationCode, str2SrcQB_CreditCardTxnInfoResultAVSStreet, str2SrcQB_CreditCardTxnInfoResultAVSZip, str2SrcQB_CreditCardTxnInfoResultReconBatchID, str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode, str2SrcQB_CreditCardTxnInfoResultPaymentStatus, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp, str2SrcQB_IsAutoApply, str2SrcQB_UnusedPayment, str2SrcQB_UnusedCredits, str2SrcQB_AppliedToTxnTxnID, str2SrcQB_AppliedToTxnPaymentAmount, str2SrcQB_AppliedToTxnTxnType, str2SrcQB_AppliedToTxnTxnDate, str2SrcQB_AppliedToTxnRefNumber, str2SrcQB_AppliedToTxnBalanceRemaining, str2SrcQB_AppliedToTxnAmount, str2SrcQB_AppliedToTxnSetCreditCreditTxnID, str2SrcQB_AppliedToTxnSetCreditAppliedAmount, str2SrcQB_AppliedToTxnDiscountAmount, str2SrcQB_AppliedToTxnDiscountAccountRefListID, str2SrcQB_AppliedToTxnDiscountAccountRefFullName, str2SrcQB_FQSaveToCache, str2SrcQB_FQPrimaryKey As String
    'This routine gets the 2SrcQB_QB_ReceivePaymentLine from the database according to the selection in str2SrcQB_QB_ReceivePaymentLineSQL.
    'It then puts those 2SrcQB_QB_ReceivePaymentLine in the list box

    'FOR PART 3TestID_
    Debug.WriteLine("List3TestID_QBTable")

    'dim SQL strings
    Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String

    Dim strNowMinusThirtyDays As String = ""
    strNowMinusThirtyDays = DateTime.Now.AddDays(-60).ToString("yyyy-MM-dd HH:mm:ss.000s")


    'PART 2SrcQB_: Get the new records from Actual QB
    'Get a recordset of records from QB that are newer than QBTable
    Using rs2SrcQB_QB_ReceivePaymentLine As DataSet = New DataSet()


        str2SrcQB_QB_ReceivePaymentLineSQL = "SELECT * FROM ReceivePaymentLine WHERE CustomerRefFullName = '" & strCustomerRefFullName & "' AND TimeModified > {ts '" & DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd HH:mm:ss.000") & "'}"
        Debug.WriteLine(str2SrcQB_QB_ReceivePaymentLineSQL)
        Using adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_ReceivePaymentLineSQL, cnQuickBooks)
            rs2SrcQB_QB_ReceivePaymentLine.Tables.Clear()
            adap.Fill(rs2SrcQB_QB_ReceivePaymentLine)
        End Using


        If rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows.Count > 10 Then
            ShowUserMessage(strSubName, "Too many RPL items found (more than 10)")
            Exit Sub
        End If

        Dim curRow As Integer = 0
        Dim rowCount As Integer = rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows.Count

        If rowCount > 0 Then


            For Each iteration_row As DataRow In rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows
                curRow += 1

                ShowUserMessage(strSubName, "Processing Record " & curRow.ToString & " of " & rowCount.ToString)

                'get the columns from the database
                str2SrcQB_TxnID = NCStr(iteration_row("TxnID")).Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                str2SrcQB_TxnNumber = NCStr(iteration_row("TxnNumber"), "0").Replace("'"c, "`"c)
                str2SrcQB_CustomerRefListID = NCStr(iteration_row("CustomerRefListID")).Replace("'"c, "`"c)
                str2SrcQB_CustomerRefFullName = NCStr(iteration_row("CustomerRefFullName")).Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefListID = NCStr(iteration_row("ARAccountRefListID")).Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefFullName = NCStr(iteration_row("ARAccountRefFullName")).Replace("'"c, "`"c)
                str2SrcQB_TxnDate = NCStr(iteration_row("TxnDate")).Replace("'"c, "`"c)
                str2SrcQB_TxnDateMacro = NCStr(iteration_row("TxnDateMacro")).Replace("'"c, "`"c)
                str2SrcQB_RefNumber = NCStr(iteration_row("RefNumber")).Replace("'"c, "`"c)
                str2SrcQB_TotalAmount = NCStr(iteration_row("TotalAmount"), "0").Replace("'"c, "`"c)
                str2SrcQB_PaymentMethodRefListID = NCStr(iteration_row("PaymentMethodRefListID")).Replace("'"c, "`"c)
                str2SrcQB_PaymentMethodRefFullName = NCStr(iteration_row("PaymentMethodRefFullName")).Replace("'"c, "`"c)
                str2SrcQB_Memo = NCStr(iteration_row("Memo")).Replace("'"c, "`"c)
                str2SrcQB_DepositToAccountRefListID = NCStr(iteration_row("DepositToAccountRefListID")).Replace("'"c, "`"c)
                str2SrcQB_DepositToAccountRefFullName = NCStr(iteration_row("DepositToAccountRefFullName")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCreditCardNumber = NCStr(iteration_row("CreditCardTxnInfoInputCreditCardNumber")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputExpirationMonth = NCStr(iteration_row("CreditCardTxnInfoInputExpirationMonth"), "0").Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputExpirationYear = NCStr(iteration_row("CreditCardTxnInfoInputExpirationYear"), "0").Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputNameOnCard = NCStr(iteration_row("CreditCardTxnInfoInputNameOnCard")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCreditCardAddress = NCStr(iteration_row("CreditCardTxnInfoInputCreditCardAddress")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode = NCStr(iteration_row("CreditCardTxnInfoInputCreditCardPostalCode")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCommercialCardCode = NCStr(iteration_row("CreditCardTxnInfoInputCommercialCardCode")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultResultCode = NCStr(iteration_row("CreditCardTxnInfoResultResultCode"), "0").Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultResultMessage = NCStr(iteration_row("CreditCardTxnInfoResultResultMessage")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultCreditCardTransID = NCStr(iteration_row("CreditCardTxnInfoResultCreditCardTransID")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber = NCStr(iteration_row("CreditCardTxnInfoResultMerchantAccountNumber")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultAuthorizationCode = NCStr(iteration_row("CreditCardTxnInfoResultAuthorizationCode")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultAVSStreet = NCStr(iteration_row("CreditCardTxnInfoResultAVSStreet")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultAVSZip = NCStr(iteration_row("CreditCardTxnInfoResultAVSZip")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultReconBatchID = NCStr(iteration_row("CreditCardTxnInfoResultReconBatchID")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode = NCStr(iteration_row("CreditCardTxnInfoResultPaymentGroupingCode"), "0").Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultPaymentStatus = NCStr(iteration_row("CreditCardTxnInfoResultPaymentStatus")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime = NCStr(iteration_row("CreditCardTxnInfoResultTxnAuthorizationTime")).Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp = NCStr(iteration_row("CreditCardTxnInfoResultTxnAuthorizationStamp"), "0").Replace("'"c, "`"c)
                str2SrcQB_IsAutoApply = NCStr(iteration_row("IsAutoApply")).Replace("'"c, "`"c)
                str2SrcQB_UnusedPayment = NCStr(iteration_row("UnusedPayment"), "0").Replace("'"c, "`"c)
                str2SrcQB_UnusedCredits = NCStr(iteration_row("UnusedCredits"), "0").Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnTxnID = NCStr(iteration_row("AppliedToTxnTxnID")).Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnPaymentAmount = NCStr(iteration_row("AppliedToTxnPaymentAmount"), "0").Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnTxnType = NCStr(iteration_row("AppliedToTxnTxnType")).Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnTxnDate = NCStr(iteration_row("AppliedToTxnTxnDate")).Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnRefNumber = NCStr(iteration_row("AppliedToTxnRefNumber")).Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnBalanceRemaining = NCStr(iteration_row("AppliedToTxnBalanceRemaining"), "0").Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnAmount = NCStr(iteration_row("AppliedToTxnAmount"), "0").Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnSetCreditCreditTxnID = NCStr(iteration_row("AppliedToTxnSetCreditCreditTxnID")).Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnSetCreditAppliedAmount = NCStr(iteration_row("AppliedToTxnSetCreditAppliedAmount"), "0").Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnDiscountAmount = NCStr(iteration_row("AppliedToTxnDiscountAmount"), "0").Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnDiscountAccountRefListID = NCStr(iteration_row("AppliedToTxnDiscountAccountRefListID")).Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnDiscountAccountRefFullName = NCStr(iteration_row("AppliedToTxnDiscountAccountRefFullName")).Replace("'"c, "`"c)
                str2SrcQB_FQSaveToCache = NCStr(iteration_row("FQSaveToCache")).Replace("'"c, "`"c)
                str2SrcQB_FQPrimaryKey = NCStr(iteration_row("FQPrimaryKey")).Replace("'"c, "`"c)

                'Change flags back to binary
                str2SrcQB_IsAutoApply = IIf(str2SrcQB_IsAutoApply = "True", "1", "0")
                str2SrcQB_FQSaveToCache = IIf(str2SrcQB_FQSaveToCache = "True", "1", "0")

                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_ReceivePaymentLineRow = "" & _
                                                     Strings.Left("RPL Upd" & "                  ", 18) & "   " & _
                                                     Strings.Left(str2SrcQB_TimeCreated & "                  ", 16) & "   " & _
                                                     Strings.Left(str2SrcQB_TimeModified & "                  ", 16) & "   " & _
                                                     Strings.Left(str2SrcQB_CustomerRefFullName & "                  ", 18) & "   " & _
                                                     Strings.Left(str2SrcQB_TxnDate & "                  ", 10) & "   " & _
                                                     Strings.Left(str2SrcQB_RefNumber & "                  ", 10) & "   " & _
                                                     Strings.Left(str2SrcQB_TotalAmount & "                  ", 18) & "   " & _
                                                     Strings.Left(str2SrcQB_PaymentMethodRefFullName & "                  ", 18) & "   " & _
                                                     "" & Strings.Chr(9)

                'Left(str2SrcQB_TxnID + "                  ", 18) & "   " & _
                '
                'put the line in the listbox


                ShowUserMessage(strSubName, "Processing Record " & curRow.ToString & " of " & rowCount.ToString)
                ShowUserMessage(strSubName, str2SrcQB_QB_ReceivePaymentLineRow)

                'Check to see if ListID or TxnID is in QBTable            'Yes then UPDATE record            'No then INSERT record
                Dim iRowCount As Integer = 0
                iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(FQPrimaryKey) FROM QB_ReceivePaymentLine WHERE FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'")
                If iRowCount > 1 Then Stop 'Should only be one
                If iRowCount = 1 Then 'record exists  -UPDATE

                    'DO UPDATE WORK:
                    Debug.WriteLine("UPDATE")

                    'Build the SQL string
                    strSQL1 = "UPDATE  " & Environment.NewLine & _
                              "       QB_ReceivePaymentLine " & Environment.NewLine & _
                              "SET " & Environment.NewLine & _
                              "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine & _
                              "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                              "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                              "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                              "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & Environment.NewLine & _
                              "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & Environment.NewLine & _
                              "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & Environment.NewLine & _
                              "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & Environment.NewLine & _
                              "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & Environment.NewLine & _
                              "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & Environment.NewLine & _
                              "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & Environment.NewLine & _
                              "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & Environment.NewLine & _
                              "     , TotalAmount = " & str2SrcQB_TotalAmount & "" & Environment.NewLine & _
                              "     , PaymentMethodRefListID = '" & str2SrcQB_PaymentMethodRefListID & "'" & Environment.NewLine & _
                              "     , PaymentMethodRefFullName = '" & str2SrcQB_PaymentMethodRefFullName & "'" & Environment.NewLine & _
                              "     , Memo = '" & str2SrcQB_Memo & "'" & Environment.NewLine & _
                              "     , DepositToAccountRefListID = '" & str2SrcQB_DepositToAccountRefListID & "'" & Environment.NewLine
                    strSQL2 = "     , DepositToAccountRefFullName = '" & str2SrcQB_DepositToAccountRefFullName & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoInputCreditCardNumber = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardNumber & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoInputExpirationMonth = " & str2SrcQB_CreditCardTxnInfoInputExpirationMonth & "" & Environment.NewLine & _
                              "     , CreditCardTxnInfoInputExpirationYear = " & str2SrcQB_CreditCardTxnInfoInputExpirationYear & "" & Environment.NewLine & _
                              "     , CreditCardTxnInfoInputNameOnCard = '" & str2SrcQB_CreditCardTxnInfoInputNameOnCard & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoInputCreditCardAddress = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardAddress & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoInputCreditCardPostalCode = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoInputCommercialCardCode = '" & str2SrcQB_CreditCardTxnInfoInputCommercialCardCode & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultResultCode = " & str2SrcQB_CreditCardTxnInfoResultResultCode & "" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultResultMessage = '" & str2SrcQB_CreditCardTxnInfoResultResultMessage & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultCreditCardTransID = '" & str2SrcQB_CreditCardTxnInfoResultCreditCardTransID & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultMerchantAccountNumber = '" & str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultAuthorizationCode = '" & str2SrcQB_CreditCardTxnInfoResultAuthorizationCode & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultAVSStreet = '" & str2SrcQB_CreditCardTxnInfoResultAVSStreet & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultAVSZip = '" & str2SrcQB_CreditCardTxnInfoResultAVSZip & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultReconBatchID = '" & str2SrcQB_CreditCardTxnInfoResultReconBatchID & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultPaymentGroupingCode = " & str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode & "" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultPaymentStatus = '" & str2SrcQB_CreditCardTxnInfoResultPaymentStatus & "'" & Environment.NewLine
                    strSQL3 = "     , CreditCardTxnInfoResultTxnAuthorizationTime = '" & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime & "'" & Environment.NewLine & _
                              "     , CreditCardTxnInfoResultTxnAuthorizationStamp = " & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp & "" & Environment.NewLine & _
                              "     , IsAutoApply = '" & str2SrcQB_IsAutoApply & "'" & Environment.NewLine & _
                              "     , UnusedPayment = " & str2SrcQB_UnusedPayment & "" & Environment.NewLine & _
                              "     , UnusedCredits = " & str2SrcQB_UnusedCredits & "" & Environment.NewLine & _
                              "     , AppliedToTxnTxnID = '" & str2SrcQB_AppliedToTxnTxnID & "'" & Environment.NewLine & _
                              "     , AppliedToTxnPaymentAmount = " & str2SrcQB_AppliedToTxnPaymentAmount & "" & Environment.NewLine & _
                              "     , AppliedToTxnTxnType = '" & str2SrcQB_AppliedToTxnTxnType & "'" & Environment.NewLine & _
                              "     , AppliedToTxnTxnDate = '" & str2SrcQB_AppliedToTxnTxnDate & "'" & Environment.NewLine & _
                              "     , AppliedToTxnRefNumber = '" & str2SrcQB_AppliedToTxnRefNumber & "'" & Environment.NewLine & _
                              "     , AppliedToTxnBalanceRemaining = " & str2SrcQB_AppliedToTxnBalanceRemaining & "" & Environment.NewLine & _
                              "     , AppliedToTxnAmount = " & str2SrcQB_AppliedToTxnAmount & "" & Environment.NewLine & _
                              "     , AppliedToTxnSetCreditCreditTxnID = '" & str2SrcQB_AppliedToTxnSetCreditCreditTxnID & "'" & Environment.NewLine & _
                              "     , AppliedToTxnSetCreditAppliedAmount = " & str2SrcQB_AppliedToTxnSetCreditAppliedAmount & "" & Environment.NewLine & _
                              "     , AppliedToTxnDiscountAmount = " & str2SrcQB_AppliedToTxnDiscountAmount & "" & Environment.NewLine & _
                              "     , AppliedToTxnDiscountAccountRefListID = '" & str2SrcQB_AppliedToTxnDiscountAccountRefListID & "'" & Environment.NewLine & _
                              "     , AppliedToTxnDiscountAccountRefFullName = '" & str2SrcQB_AppliedToTxnDiscountAccountRefFullName & "'" & Environment.NewLine & _
                              "     , FQSaveToCache = '" & str2SrcQB_FQSaveToCache & "'" & Environment.NewLine & _
                              "     , FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & Environment.NewLine & _
                              "WHERE " & Environment.NewLine & _
                              "       FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & Environment.NewLine

                    'Combine the strings
                    strTableUpdate = strSQL1 & strSQL2 & strSQL3 '& strSQL4 & strSQL5 & strSQL6
                    SQLHelper.ExecuteSQL(cnMax, strTableUpdate)

                Else
                    'record not exist  -INSERT
                    Debug.WriteLine("INSERT")

                    'Build the SQL string
                    strSQL1 = "INSERT INTO QB_ReceivePaymentLine " & Environment.NewLine & _
                              "   ( TxnID " & Environment.NewLine & _
                              "   , TimeCreated " & Environment.NewLine & _
                              "   , TimeModified " & Environment.NewLine & _
                              "   , EditSequence " & Environment.NewLine & _
                              "   , TxnNumber " & Environment.NewLine & _
                              "   , CustomerRefListID " & Environment.NewLine & _
                              "   , CustomerRefFullName " & Environment.NewLine & _
                              "   , ARAccountRefListID " & Environment.NewLine & _
                              "   , ARAccountRefFullName " & Environment.NewLine & _
                              "   , TxnDate " & Environment.NewLine & _
                              "   , TxnDateMacro " & Environment.NewLine & _
                              "   , RefNumber " & Environment.NewLine & _
                              "   , TotalAmount " & Environment.NewLine & _
                              "   , PaymentMethodRefListID " & Environment.NewLine & _
                              "   , PaymentMethodRefFullName " & Environment.NewLine & _
                              "   , Memo " & Environment.NewLine & _
                              "   , DepositToAccountRefListID " & Environment.NewLine
                    strSQL2 = "   , DepositToAccountRefFullName " & Environment.NewLine & _
                              "   , CreditCardTxnInfoInputCreditCardNumber " & Environment.NewLine & _
                              "   , CreditCardTxnInfoInputExpirationMonth " & Environment.NewLine & _
                              "   , CreditCardTxnInfoInputExpirationYear " & Environment.NewLine & _
                              "   , CreditCardTxnInfoInputNameOnCard " & Environment.NewLine & _
                              "   , CreditCardTxnInfoInputCreditCardAddress " & Environment.NewLine & _
                              "   , CreditCardTxnInfoInputCreditCardPostalCode " & Environment.NewLine & _
                              "   , CreditCardTxnInfoInputCommercialCardCode " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultResultCode " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultResultMessage " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultCreditCardTransID " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultMerchantAccountNumber " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultAuthorizationCode " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultAVSStreet " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultAVSZip " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultReconBatchID " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultPaymentGroupingCode " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultPaymentStatus " & Environment.NewLine
                    strSQL3 = "   , CreditCardTxnInfoResultTxnAuthorizationTime " & Environment.NewLine & _
                              "   , CreditCardTxnInfoResultTxnAuthorizationStamp " & Environment.NewLine & _
                              "   , IsAutoApply " & Environment.NewLine & _
                              "   , UnusedPayment " & Environment.NewLine & _
                              "   , UnusedCredits " & Environment.NewLine & _
                              "   , AppliedToTxnTxnID " & Environment.NewLine & _
                              "   , AppliedToTxnPaymentAmount " & Environment.NewLine & _
                              "   , AppliedToTxnTxnType " & Environment.NewLine & _
                              "   , AppliedToTxnTxnDate " & Environment.NewLine & _
                              "   , AppliedToTxnRefNumber " & Environment.NewLine & _
                              "   , AppliedToTxnBalanceRemaining " & Environment.NewLine & _
                              "   , AppliedToTxnAmount " & Environment.NewLine & _
                              "   , AppliedToTxnSetCreditCreditTxnID " & Environment.NewLine & _
                              "   , AppliedToTxnSetCreditAppliedAmount " & Environment.NewLine & _
                              "   , AppliedToTxnDiscountAmount " & Environment.NewLine & _
                              "   , AppliedToTxnDiscountAccountRefListID " & Environment.NewLine & _
                              "   , AppliedToTxnDiscountAccountRefFullName " & Environment.NewLine & _
                              "   , FQSaveToCache " & Environment.NewLine & _
                              "   , FQPrimaryKey ) " & Environment.NewLine
                    strSQL4 = "VALUES " & Environment.NewLine & _
                              "   ( '" & str2SrcQB_TxnID & "'  --TxnID" & Environment.NewLine & _
                              "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                              "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                              "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                              "   , " & str2SrcQB_TxnNumber & "  --TxnNumber" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CustomerRefListID & "'  --CustomerRefListID" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CustomerRefFullName & "'  --CustomerRefFullName" & Environment.NewLine & _
                              "   , '" & str2SrcQB_ARAccountRefListID & "'  --ARAccountRefListID" & Environment.NewLine & _
                              "   , '" & str2SrcQB_ARAccountRefFullName & "'  --ARAccountRefFullName" & Environment.NewLine & _
                              "   , '" & str2SrcQB_TxnDate & "'  --TxnDate" & Environment.NewLine & _
                              "   , '" & str2SrcQB_TxnDateMacro & "'  --TxnDateMacro" & Environment.NewLine & _
                              "   , '" & str2SrcQB_RefNumber & "'  --RefNumber" & Environment.NewLine & _
                              "   , " & str2SrcQB_TotalAmount & "  --TotalAmount" & Environment.NewLine & _
                              "   , '" & str2SrcQB_PaymentMethodRefListID & "'  --PaymentMethodRefListID" & Environment.NewLine & _
                              "   , '" & str2SrcQB_PaymentMethodRefFullName & "'  --PaymentMethodRefFullName" & Environment.NewLine & _
                              "   , '" & str2SrcQB_Memo & "'  --Memo" & Environment.NewLine & _
                              "   , '" & str2SrcQB_DepositToAccountRefListID & "'  --DepositToAccountRefListID" & Environment.NewLine
                    strSQL5 = "   , '" & str2SrcQB_DepositToAccountRefFullName & "'  --DepositToAccountRefFullName" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoInputCreditCardNumber & "'  --CreditCardTxnInfoInputCreditCardNumber" & Environment.NewLine & _
                              "   , " & str2SrcQB_CreditCardTxnInfoInputExpirationMonth & "  --CreditCardTxnInfoInputExpirationMonth" & Environment.NewLine & _
                              "   , " & str2SrcQB_CreditCardTxnInfoInputExpirationYear & "  --CreditCardTxnInfoInputExpirationYear" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoInputNameOnCard & "'  --CreditCardTxnInfoInputNameOnCard" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoInputCreditCardAddress & "'  --CreditCardTxnInfoInputCreditCardAddress" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode & "'  --CreditCardTxnInfoInputCreditCardPostalCode" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoInputCommercialCardCode & "'  --CreditCardTxnInfoInputCommercialCardCode" & Environment.NewLine & _
                              "   , " & str2SrcQB_CreditCardTxnInfoResultResultCode & "  --CreditCardTxnInfoResultResultCode" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoResultResultMessage & "'  --CreditCardTxnInfoResultResultMessage" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoResultCreditCardTransID & "'  --CreditCardTxnInfoResultCreditCardTransID" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber & "'  --CreditCardTxnInfoResultMerchantAccountNumber" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoResultAuthorizationCode & "'  --CreditCardTxnInfoResultAuthorizationCode" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoResultAVSStreet & "'  --CreditCardTxnInfoResultAVSStreet" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoResultAVSZip & "'  --CreditCardTxnInfoResultAVSZip" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoResultReconBatchID & "'  --CreditCardTxnInfoResultReconBatchID" & Environment.NewLine & _
                              "   , " & str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode & "  --CreditCardTxnInfoResultPaymentGroupingCode" & Environment.NewLine & _
                              "   , '" & str2SrcQB_CreditCardTxnInfoResultPaymentStatus & "'  --CreditCardTxnInfoResultPaymentStatus" & Environment.NewLine
                    strSQL6 = "   , '" & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime & "'  --CreditCardTxnInfoResultTxnAuthorizationTime" & Environment.NewLine & _
                              "   , " & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp & "  --CreditCardTxnInfoResultTxnAuthorizationStamp" & Environment.NewLine & _
                              "   , '" & str2SrcQB_IsAutoApply & "'  --IsAutoApply" & Environment.NewLine & _
                              "   , " & str2SrcQB_UnusedPayment & "  --UnusedPayment" & Environment.NewLine & _
                              "   , " & str2SrcQB_UnusedCredits & "  --UnusedCredits" & Environment.NewLine & _
                              "   , '" & str2SrcQB_AppliedToTxnTxnID & "'  --AppliedToTxnTxnID" & Environment.NewLine & _
                              "   , " & str2SrcQB_AppliedToTxnPaymentAmount & "  --AppliedToTxnPaymentAmount" & Environment.NewLine & _
                              "   , '" & str2SrcQB_AppliedToTxnTxnType & "'  --AppliedToTxnTxnType" & Environment.NewLine & _
                              "   , '" & str2SrcQB_AppliedToTxnTxnDate & "'  --AppliedToTxnTxnDate" & Environment.NewLine & _
                              "   , '" & str2SrcQB_AppliedToTxnRefNumber & "'  --AppliedToTxnRefNumber" & Environment.NewLine & _
                              "   , " & str2SrcQB_AppliedToTxnBalanceRemaining & "  --AppliedToTxnBalanceRemaining" & Environment.NewLine & _
                              "   , " & str2SrcQB_AppliedToTxnAmount & "  --AppliedToTxnAmount" & Environment.NewLine & _
                              "   , '" & str2SrcQB_AppliedToTxnSetCreditCreditTxnID & "'  --AppliedToTxnSetCreditCreditTxnID" & Environment.NewLine & _
                              "   , " & str2SrcQB_AppliedToTxnSetCreditAppliedAmount & "  --AppliedToTxnSetCreditAppliedAmount" & Environment.NewLine & _
                              "   , " & str2SrcQB_AppliedToTxnDiscountAmount & "  --AppliedToTxnDiscountAmount" & Environment.NewLine & _
                              "   , '" & str2SrcQB_AppliedToTxnDiscountAccountRefListID & "'  --AppliedToTxnDiscountAccountRefListID" & Environment.NewLine & _
                              "   , '" & str2SrcQB_AppliedToTxnDiscountAccountRefFullName & "'  --AppliedToTxnDiscountAccountRefFullName" & Environment.NewLine & _
                              "   , '" & str2SrcQB_FQSaveToCache & "'  --FQSaveToCache" & Environment.NewLine & _
                              "   , '" & str2SrcQB_FQPrimaryKey & "' ) --FQPrimaryKey" & Environment.NewLine

                    'Combine the strings
                    strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
                    'Debug.Print strTableInsert

                    'Execute the insert
                    SQLHelper.ExecuteSQL(cnMax, strTableInsert)

                End If
            Next iteration_row

        Else
            ShowUserMessage(strSubName, "No receive payment lines to process")

        End If

    End Using
    ShowUserMessage(strSubName, "Finished processing receive payment lines")

End Sub
