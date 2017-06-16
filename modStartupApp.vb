Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Windows.Forms
Module modStartupApp

    Public gstrUserName As String = ""
    Public gstrComputerUser As String = ""

    Public gstrModuleFileName As String = ""
    Public gstrComputerName As String = ""
    Public gstrVersionNum As String = ""
    Public gstrCompany As String = ""
    Public gstrApplicationName As String = ""
    Public gstrInstancePID As String = ""
    Public gstrScreenResolution As String = ""
    Public gstrServerName As String = ""
    Public gdteVersionExpirationDate As Date
    Dim rs1_DBPM_Login As DataSet
    Dim gstrLogin_DBPM_LoginSQL As String = ""
    Dim gstrLogin_DBPM_LoginRow As String = ""
    Dim gstrLogin_LoginN As String = ""
    Dim gstrLogin_CreatedDate As String = ""
    Dim gstrLogin_CreatedBy As String = ""
    Dim gstrLogin_CreatedOnComputer As String = ""
    Dim gstrLogin_LoginUserName As String = ""
    Dim gstrLogin_LoginPassword As String = ""
    Dim gstrLogin_AutoLogin As String = ""
    Dim gstrLogin_StartupScreen As String = ""
    Public gCustomerBalanceUpdateList As New ArrayList()
    Public gShowProcessing As Boolean = False



    <STAThread> _
    Public Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        Dim strObjName As String = "modStartupApp" '"OBJNAME"
        Dim strSubName As String = "Main" '"SUBNAME"

        modConfig.setConfiguration()

        Try

            gstrApplicationName = "DBPM_Server"

            'Set flags
            gbooUsePermissions = True
            gbooPermSelfAwarenessOn = True
            gbooEndProgram = False

            gstrUserName = Environment.UserName.ToUpper
            gstrComputerUser = gstrUserName

            gstrComputerUser = Environment.MachineName

            gstrServerName = mainServerName
            Dim gbooIsMaximizerUser As Boolean = True

            gstrServerName = mainServerName
            gbooIsMaximizerUser = True

            gbooUseErrorHandling = True

            gstrInstancePID = "2" 'HOW GET?
            gstrCompany = "DrummondPrinting"

            Application.Run(frmMain)
            Application.DoEvents()

            gstrHighestAuditTrailAlreadyLoaded = "2000010100001"
            gstrHighestAuditTrailAlreadyLoaded = "2005122300171" 'MAKE-UP LOAD 1
            gstrHighestAuditTrailAlreadyLoaded = "2006021000040" 'MAKE-UP LOAD 2

            'Start the global QB insert counter
            gintQBInsertCounter = 0

            'Reset flags
            booQBRefreshInProgress = False

        Catch exc As System.Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), exc.Message, Information.Err().Source, "", "")
        End Try


    End Sub
End Module