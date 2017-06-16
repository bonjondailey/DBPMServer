Option Strict Off
Option Explicit On
Imports System
Module modConfig
	Public mainServerName As String = ""
	Public drummondQuickbookUser, drummondQuickbookPath, drummondQuickbookPass As String
	Public sqlServerUser, sqlServerPass As String
	Public sqlMaxUser, sqlMaxPass As String
    Public gstrSQLConnectionString As String
    Public QB_Remote_Connection_Application As String

	Public Sub setConfiguration()
        mainServerName = siteConstants.GetDBServer

		drummondQuickbookPath = "\\" & mainServerName & "\Accounting\DrummondQB\DrummondPrinting.QBW"
		drummondQuickbookUser = "DBPM_Server2"
        drummondQuickbookPass = "DBPM_qb_2017"

        QB_Remote_Connection_Application = "DBPM_Server_Remote_Access"
        'drummondQuickbookPass = "ServeyUs127" 'on devserver

        sqlMaxUser = "MASTER"
		sqlMaxPass = "CONTROL"

        gstrSQLConnectionString = siteConstants.SQLConnectionString

	End Sub
End Module