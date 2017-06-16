Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Windows.Forms
Imports DBPM_Server.siteConstants
Imports System.Net.Mail


Module modPermissions

	'For Permissions and ErrorHandling
	Public gbooUseErrorHandling As Boolean
	Public gbooIsPermSub As Boolean

	'For Permissions
	Public gstrObjName As String = ""
	Public gstrSubName As String = ""
	Public gbooUsePermissions As Boolean
	Public gbooPermSelfAwarenessOn As Boolean
    Public gbooVerboseErrors As Boolean = False

	Public gstrLastError_Now As String = ""
	Public gstrLastError_ObjName As String = ""
	Public gstrLastError_SubName As String = ""
	Public gstrLastError_ErrNumber As String = ""
	Public gstrLastError_ErrDescription As String = ""

	Public gintErrorCounter_AllErrors As Integer
	Public gintErrorCounter_SameError As Integer
	Public gintErrorCounter_SameSub As Integer

    Public gbooEndProgram As Boolean

    Public iErrCount As Integer = 0
    Public strProgramErrors As New Text.StringBuilder

    Public Sub HaveError(ByRef strObjName As String, ByRef strSubName As String, ByRef strErrNumber As String, ByRef strErrDescription As String, ByRef strErrSource As String, ByRef strJobNum As String, ByRef strCust As String, Optional ByVal thrownException As Exception = Nothing)
        Dim HaveErrorErrorHandlingError As Boolean = False
        Dim result As String = ""
        Dim gstrSalesOrderN As String = "", gstrStepN As String = ""

        Dim strTask As String = ""
        Dim strInformation As String = ""
        Dim strTableName As String = ""
        Dim strKeyNName As String = ""
        Dim strKeyN As String = ""
        Dim strSalesOrderN As String = ""
        Dim strStepN As String = ""
        Dim strNow As String = ""
        Dim strErrStackTrace As String = ""

        strProgramErrors.Append(DateTime.Now & " : " & strSubName & " : " & strErrDescription & Environment.NewLine & Environment.NewLine)
        ShowUserMessage(strSubName, strErrDescription, strObjName & ":" & strSubName & ":ERROR", True)

        Try
            If gbooUseErrorHandling Then HaveErrorErrorHandlingError = True

            Dim strFileName As String = "" ' 50  1
            Dim strExtension As String = "" ' 50  1
            Dim strClient_Id As String = "" ' 50  1
            Dim strContact_Number As String = "" ' 50  1
            Dim strListID As String = "" ' 50  1


            'Insert into error tracking table & notification
            strNow = DateTime.Now.ToString

            'Bypass these later when real data is passed
            strSalesOrderN = NCStr(gstrSalesOrderN)
            strStepN = NCStr(gstrStepN)

            If gbooVerboseErrors Then

                If Not thrownException Is Nothing Then
                    strErrDescription += Environment.NewLine & thrownException.Source & "-" & thrownException.Message & Environment.NewLine
                    strErrStackTrace = thrownException.StackTrace
                End If

                'ASK THE USER WHAT TO DO
                If MessageBox.Show("Error:" & Environment.NewLine & Environment.NewLine & "" & strObjName & " - " & strSubName & "" & Environment.NewLine & Environment.NewLine & strErrNumber & "  " & strErrDescription & "" & Environment.NewLine & Environment.NewLine & strJobNum & "   " & strCust & Environment.NewLine & Environment.NewLine & "Continue Processing Anyway?   (Usually choose Yes)", "Error. Continue Anyway?", MessageBoxButtons.YesNoCancel) = System.Windows.Forms.DialogResult.Yes Then
                    result = "RN" 'keep running 
                Else
                    result = "ES" 'end program
                    Application.DoEvents()
                End If


            Else
                result = "RN"
            End If


            'ERROR COUNTER - ALL ERRORS
            gintErrorCounter_AllErrors += 1
            If gintErrorCounter_AllErrors > 1000 Then
                'MessageBox.Show("Too many errors in the same session. " & Environment.NewLine & "The current Procedure was aborted." & Environment.NewLine & "Try again or restart the program if errors persist.", Application.ProductName)
                result = "ES"
                gintErrorCounter_AllErrors = 0
                gintErrorCounter_SameError = 0
                gintErrorCounter_SameSub = 0
                Application.DoEvents()
            End If



            'ERROR COUNTER - SAME SUB
            'Reset error counter if not the same error as last time
            If gstrLastError_ObjName <> strObjName Or gstrLastError_SubName <> strSubName Then
                gintErrorCounter_SameSub = 0
            End If

            'Stop error looping from happening
            'If gstrLastError_Now = strNow And gstrLastError_ObjName = strObjName And gstrLastError_SubName = strSubName And gstrLastError_ErrNumber = strErrNumber And gstrLastError_ErrDescription = strErrDescription Then
            Dim TempDate As String = Date.Now.ToString


            gstrLastError_Now = siteConstants.NCDate(gstrLastError_Now).ToString

            If TempDate = gstrLastError_Now Then
                gintErrorCounter_SameSub += 1
                If gintErrorCounter_SameSub > 100 Then
                    'MsgBox "This Procedure was aborted due to many errors. " & vbCrLf & "Try again or restart the program if errors persist."
                    'MessageBox.Show("Too many errors in the same procedure within the same second. " & Environment.NewLine & "The current Procedure was aborted." & Environment.NewLine & "Try again or restart the program if errors persist.", Application.ProductName)
                    result = "ES"
                    ' frmMessage.Hide()
                    gintErrorCounter_AllErrors = 0
                    gintErrorCounter_SameError = 0
                    gintErrorCounter_SameSub = 0
                    Application.DoEvents()
                End If
            End If


            'ERROR COUNTER - SAME ERROR
            'Reset error counter if not the same error as last time
            If gstrLastError_ObjName <> strObjName OrElse gstrLastError_SubName <> strSubName OrElse gstrLastError_ErrNumber <> strErrNumber OrElse gstrLastError_ErrDescription <> strErrDescription Then
                gintErrorCounter_SameError = 0
            End If

            'Stop error looping from happening
            If TempDate = gstrLastError_Now AndAlso gstrLastError_ObjName = strObjName AndAlso gstrLastError_SubName = strSubName AndAlso gstrLastError_ErrNumber = strErrNumber AndAlso gstrLastError_ErrDescription = strErrDescription Then
                gintErrorCounter_SameError += 1
                If gintErrorCounter_SameError > 15 Then
                    'MessageBox.Show("Too many of the same error in the same procedure within the same second. " & Environment.NewLine & "The current Procedure was aborted." & Environment.NewLine & "Try again or restart the program if errors persist.", Application.ProductName)
                    result = "ES"
                    'frmMessage.Hide()
                    gintErrorCounter_AllErrors = 0
                    gintErrorCounter_SameError = 0
                    gintErrorCounter_SameSub = 0
                    Application.DoEvents()
                End If
            End If

            gstrLastError_Now = strNow
            gstrLastError_ObjName = strObjName
            gstrLastError_SubName = strSubName
            gstrLastError_ErrNumber = strErrNumber
            gstrLastError_ErrDescription = strErrDescription

            'Fix strings
            strErrDescription = strErrDescription.Replace("'"c, "`"c)
            strErrDescription = strErrDescription.Replace(Environment.NewLine, " : ")

           
            'Insert into error tracking table & notification
            Dim strSQL1, strSQL2, strSQL3, strSQL4, strTableInsert As String

            'Build the SQL string
            strSQL1 = "INSERT INTO DBPM_Errors " & Environment.NewLine & _
                      "   ( CreatedDate " & Environment.NewLine & _
                      "   , CreatedBy " & Environment.NewLine & _
                      "   , CreatedOnComputer " & Environment.NewLine & _
                      "   , ObjName " & Environment.NewLine & _
                      "   , SubName " & Environment.NewLine & _
                      "   , ErrNumber " & Environment.NewLine & _
                      "   , ErrDescription " & Environment.NewLine & _
                      "   , ErrSource " & Environment.NewLine & _
                      "   , JobNum " & Environment.NewLine & _
                      "   , Cust " & Environment.NewLine & _
                      "   , ApplicationName " & Environment.NewLine & _
                      "   , VersionNum " & Environment.NewLine & _
                      "   , VersionExpirationDate " & Environment.NewLine & _
                      "   , InstancePID " & Environment.NewLine
            strSQL2 = "   , ServerName " & Environment.NewLine & _
                      "   , ExeName " & Environment.NewLine & _
                      "   , ScreenResolution " & Environment.NewLine & _
                      "   , RefTask " & Environment.NewLine & _
                      "   , RefInformation " & Environment.NewLine & _
                      "   , RefTableName " & Environment.NewLine & _
                      "   , RefKeyNName " & Environment.NewLine & _
                      "   , RefKeyN " & Environment.NewLine & _
                      "   , RefSalesOrderN " & Environment.NewLine & _
                      "   , RefStepN " & Environment.NewLine & _
                      "   , RefFileName " & Environment.NewLine & _
                      "   , RefExtension " & Environment.NewLine & _
                      "   , RefClient_Id " & Environment.NewLine & _
                      "   , RefContact_Number " & Environment.NewLine & _
                      "   , RefListID, StackTrace ) " & Environment.NewLine
            strSQL3 = "VALUES " & Environment.NewLine & _
                      "   ( '" & strNow & "'  --CreatedDate" & Environment.NewLine & _
                      "   , '" & gstrUserName & "'  --CreatedBy" & Environment.NewLine & _
                      "   , '" & gstrComputerName & "'  --CreatedOnComputer" & Environment.NewLine & _
                      "   , '" & strObjName & "'  --ObjName" & Environment.NewLine & _
                      "   , '" & strSubName & "'  --SubName" & Environment.NewLine & _
                      "   , '" & strErrNumber & "'  --ErrNumber" & Environment.NewLine & _
                      "   , '" & strErrDescription & "'  --ErrDescription" & Environment.NewLine & _
                      "   , '" & strErrSource & "'  --ErrSource" & Environment.NewLine & _
                      "   , '" & strJobNum & "'  --JobNum" & Environment.NewLine & _
                      "   , '" & strCust & "'  --Cust" & Environment.NewLine & _
                      "   , '" & gstrApplicationName & "'  --ApplicationName" & Environment.NewLine & _
                      "   , '" & gstrVersionNum & "'  --VersionNum" & Environment.NewLine & _
                      "   , '" & gdteVersionExpirationDate.ToString & "'  --VersionExpirationDate" & Environment.NewLine & _
                      "   , '" & gstrInstancePID & "'  --InstancePID" & Environment.NewLine
            strSQL4 = "   , '" & gstrServerName & "'  --ServerName" & Environment.NewLine & _
                      "   , '" & gstrModuleFileName & "'  --ExeName" & Environment.NewLine & _
                      "   , '" & gstrScreenResolution & "'  --ScreenResolution" & Environment.NewLine & _
                      "   , '" & strTask & "'  --RefTask" & Environment.NewLine & _
                      "   , '" & strInformation & "'  --RefInformation" & Environment.NewLine & _
                      "   , '" & strTableName & "'  --RefTableName" & Environment.NewLine & _
                      "   , '" & strKeyNName & "'  --RefKeyNName" & Environment.NewLine & _
                      "   , '" & strKeyN & "'  --RefKeyN" & Environment.NewLine & _
                      "   , '" & strSalesOrderN & "'  --RefSalesOrderN" & Environment.NewLine & _
                      "   , '" & strStepN & "'  --RefStepN" & Environment.NewLine & _
                      "   , '" & strFileName & "'  --RefFileName" & Environment.NewLine & _
                      "   , '" & strExtension & "'  --RefExtension" & Environment.NewLine & _
                      "   , '" & strClient_Id & "'  --RefClient_Id" & Environment.NewLine & _
                      "   , '" & strContact_Number & "'  --RefContact_Number" & Environment.NewLine & _
                      "   , '" & strListID & "','" & strErrStackTrace & "' ) --RefListID" & Environment.NewLine

            'Combine the strings
            strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4
            Debug.WriteLine(strTableInsert)

            'Execute the insert
            Dim TempCommand As SqlCommand
            TempCommand = cnDBPM.CreateCommand()
            TempCommand.CommandText = strTableInsert
            iErrCount += TempCommand.ExecuteNonQuery()

        Catch excep As System.Exception
            ' If Not HaveErrorErrorHandlingError Then
            'Throw excep
            'End If
            If HaveErrorErrorHandlingError Then
                Try
                    FileSystem.FileOpen(2, "C:\ErrorsLog.txt", OpenMode.Append)
                    FileSystem.PrintLine(2, "|" & strNow & _
                                         "|" & strErrStackTrace & _
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
                                         "|" & gdteVersionExpirationDate.ToString & _
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
                    FileSystem.FileClose(2)

                    ' MessageBox.Show("CRITICAL ERROR:  Error in error handling code." & Environment.NewLine & "Connection to the Server may be lost." & Environment.NewLine & "Restart the program if problems persist." & Environment.NewLine & Environment.NewLine & CStr(Information.Err().Number) & " - " & strErrNumber & "    " & excep.Message & " - " & strErrDescription & Environment.NewLine, "CRITICAL ERROR")
                    result = "ES"

                    Dim client As New SmtpClient(mailHost, mailPort)
                    Dim msg As New MailMessage(mailErrorSendFrom, mailErrorSendTo)
                    msg.Body = "Error has occurred on DBPM_Server. - Attempt to save in LOG FILE " & Environment.NewLine & Environment.NewLine & strErrDescription
                    msg.BodyEncoding = System.Text.Encoding.UTF8
                    msg.Subject = "Error on DBPM"
                    client.Port = mailPort
                    client.Credentials = New System.Net.NetworkCredential(mailErrorSendFrom, mailPassword)
                    client.EnableSsl = True
                    client.Send(msg)




                Catch exc As System.Exception
                    result = "ES"
                End Try
            End If
        End Try


        If result = "ES" Then frmMain.Close()

    End Sub




	Public Function HavePermission(ByRef strObjName As String, ByRef strSubName As String) As Boolean
		Dim result As Boolean = False
		
		'*****************************************
		'***** TURNED OFF FOR NOW ****************
		'*****************************************
        result = True
		'*****************************************
		'*****************************************
		'*****************************************


		Return result
	End Function
End Module