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

        drummondQuickbookPath = "\\" & mainServerName & "\qbdata\DrummondPrinting.QBW"
		drummondQuickbookUser = "DBPM_Server2"
        drummondQuickbookPass = "DBPM_qb_2114Main"
        'on devserver  drummondQuickbookPass = "ServeyUs127" 

        sqlMaxUser = "MASTER"
		sqlMaxPass = "CONTROL"

        gstrSQLConnectionString = siteConstants.SQLConnectionString

        'siteConstants.ShowUserMessage("setConfiguration", "MainServerName: " & mainServerName & "---------------------" & "SQL CONNECT: " & gstrSQLConnectionString)
	End Sub
End Module