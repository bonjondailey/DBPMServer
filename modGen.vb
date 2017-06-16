Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports DBPM_Server.siteConstants
Module modGen

    Public cnQuickBooks As Odbc.OdbcConnection
    Public cnDBPM As SqlConnection
	Public cnMax As SqlConnection
    Public cnLogic As SqlConnection
	Public cnSQLQB As SqlConnection

	Public gstrHighestAuditTrailAlreadyLoaded As String = ""
	Public gstrHighest_CustomerN_AlreadyLoaded As String = ""
	Public gstrHighest_InvoiceN_AlreadyLoaded As String = ""

	Public gstrQBFile As String = ""

	'target amounts
	Public gstrGLTargetAmount As String = ""
	Public gstrAuditTrailTargetAmount As String = ""

	Public booQBRefreshInProgress As Boolean
	Public booQBFileIsOpen As Boolean

	Public Sub CloseConnectionQB()

		'Permission and ErrorHandling          (Auto built)
		Dim strObjName As String = "modGen" '"OBJNAME"
		Dim strSubName As String = "CloseConnectionQB" '"SUBNAME"

		'Check permission to run
		If Not HavePermission(strObjName, strSubName) Then Exit Sub

		'Error handling

        ShowUserMessage(strSubName, "Closing Quickbooks", "Closing Quickbooks")

        Try
            cnQuickBooks.Close()
            cnQuickBooks = Nothing
            booQBFileIsOpen = False
        Catch exc As System.Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "")
        End Try

    End Sub

   
    Public Sub OpenConnectionQB()

        If Not cnQuickBooks Is Nothing Then
            If cnQuickBooks.State = ConnectionState.Open Then
                booQBFileIsOpen = True
                Exit Sub
            End If
        End If

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modGen" '"OBJNAME"
        Dim strSubName As String = "OpenConnectionQB" '"SUBNAME"
        Dim qbConnectionString As String = ""

        ShowUserMessage(strSubName, "Opening Quickbooks File", "Opening Quickbooks File", True)

        Try

            cnQuickBooks = New Odbc.OdbcConnection()
            'LIVE SERVER?????????
            qbConnectionString = "DSN=QuickBooks Data;OLE DB Services=-2;"

            'DEV SERVER?????????
            'qbConnectionString = "Driver={QODBC Driver for QuickBooks};DFQ=" & drummondQuickbookPath & ";OpenMode=F;OLE DB Services=-2;uid=" & drummondQuickbookUser & ";pwd=" & drummondQuickbookPass & ";"
            cnQuickBooks.ConnectionString = qbConnectionString
            cnQuickBooks.Open()
        Catch ex As Exception
            HaveError(strObjName, strSubName, Information.Err.Number, ex.Message, Information.Err.Source, "", "")
        End Try

        If cnQuickBooks.State = ConnectionState.Open Then
            ShowUserMessage(strSubName, "Quickbooks Connection Established", , True)
            booQBFileIsOpen = True
        Else
            HaveError(strObjName, strSubName, "", "Quickbooks will not open. Check Task Manager on server for QBW32.EXE running. Force close if file is running.", "Quickbooks", "", "")
        End If

    End Sub

    Public Sub OpenConnectionDBPM()

        If Not cnDBPM Is Nothing Then
            If cnDBPM.State = ConnectionState.Open Then
                Exit Sub
            End If
        End If

        Dim strObjName As String = "modGen" '"OBJNAME"
        Dim strSubName As String = "OpenConnectionDBPM" '"SUBNAME"

        Try
            Dim myConnString As String = ""
            myConnString = gstrSQLConnectionString

            cnDBPM = New SqlConnection(myConnString)
            cnDBPM.Open()

            If cnDBPM.State = ConnectionState.Open Then
                ShowUserMessage(strSubName, "DBPM Connection Established ", , True)
                Debug.WriteLine("DBPM Connection Established")
            End If

        Catch ex As Exception
            ShowUserMessage(strSubName, "DBPM SQL Connection Failed", , True)
            HaveError(strObjName, strSubName, Information.Err.Number, ex.Message, Information.Err.Source, "", "")
        End Try


    End Sub

    Public Sub OpenConnectionMax()

        If Not cnMax Is Nothing Then
            If cnMax.State = ConnectionState.Open Then
                Exit Sub
            End If
        End If

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modGen" '"OBJNAME"
        Dim strSubName As String = "OpenConnectionMax" '"SUBNAME"

        Try
            Dim myConnString As String = ""
            myConnString = "Data Source=" & mainServerName & ";" & _
                           "Initial Catalog=" & gstrCompany & ";" & _
                           "User ID=" & sqlMaxUser & ";" & _
                           "Password=" & sqlMaxPass

            cnMax = New SqlConnection(myConnString)
            cnMax.Open()

            If cnMax.State = ConnectionState.Open Then
                ShowUserMessage(strSubName, "Max Connection Established", , True)
            Else
                ShowUserMessage(strSubName, "Max Connection Failed", , True)
            End If
        Catch ex As Exception
            HaveError(strObjName, strSubName, Information.Err.Number, ex.Message, Information.Err.Source, "", "")

        End Try

    End Sub

 
	

	
End Module