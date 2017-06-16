Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data.SqlClient
Imports System.Windows.Forms
Module modMaxInsertUserFields

	'These are AMGRUserFields insert vars for InsertIntoAMGRUserFields
	Public strMaxUF__Client_Id As String = ""
	Public strMaxUF__Contact_Number As String = ""
	Public strMaxUF__Type_Id As String = ""
	Public strMaxUF__Code_Id As String = ""
	Public strMaxUF__Last_Code_Id As String = ""
	Public strMaxUF__DateCol As String = ""
	Public strMaxUF__NumericCol As String = ""
	Public strMaxUF__AlphaNumericCol As String = ""
	Public strMaxUF__Record_Id As String = ""
	Public strMaxUF__Creator_Id As String = ""
	Public strMaxUF__Create_Date As String = ""
	Public strMaxUF__mmddDate As String = ""
	Public strMaxUF__Modified_By_Id As String = ""
	Public strMaxUF__Last_Modify_Date As String = ""


    '**************************************
    '******* 1st Code Review Complete *****
    '**************************************


	Sub InsertIntoAMGRUserFields()
		'This routine inserts data into AMGR_User_Fields_Tbl table.

        'Build the SQL string
		Dim strSQL1 As String = "INSERT INTO AMGR_User_Fields_Tbl " & Environment.NewLine &  _
		                        "   ( Client_Id " & Environment.NewLine &  _
		                        "   , Contact_Number " & Environment.NewLine &  _
		                        "   , Type_Id " & Environment.NewLine &  _
		                        "   , Code_Id " & Environment.NewLine &  _
		                        "   , Last_Code_Id " & Environment.NewLine &  _
		                        "   , DateCol " & Environment.NewLine &  _
		                        "   , NumericCol " & Environment.NewLine &  _
		                        "   , AlphaNumericCol " & Environment.NewLine &  _
		                        "   --, Record_Id " & Environment.NewLine &  _
		                        "   , Creator_Id " & Environment.NewLine &  _
		                        "   , Create_Date " & Environment.NewLine &  _
		                        "   , mmddDate " & Environment.NewLine &  _
		                        "   , Modified_By_Id " & Environment.NewLine &  _
		                        "   , Last_Modify_Date ) " & Environment.NewLine
		Dim strSQL2 As String = "VALUES " & Environment.NewLine &  _
		                        "   ( '" & strMaxUF__Client_Id & "'  --Client_Id" & Environment.NewLine &  _
		                        "   , " & strMaxUF__Contact_Number & "  --Contact_Number" & Environment.NewLine &  _
		                        "   , " & strMaxUF__Type_Id & "  --Type_Id" & Environment.NewLine &  _
		                        "   , " & strMaxUF__Code_Id & "  --Code_Id" & Environment.NewLine &  _
		                        "   , " & strMaxUF__Last_Code_Id & "  --Last_Code_Id" & Environment.NewLine &  _
		                        "   , '" & strMaxUF__DateCol & "'  --DateCol" & Environment.NewLine &  _
		                        "   , " & strMaxUF__NumericCol & "  --NumericCol" & Environment.NewLine &  _
		                        "   , '" & strMaxUF__AlphaNumericCol & "'  --AlphaNumericCol" & Environment.NewLine &  _
		                        "   --, " & strMaxUF__Record_Id & "  --Record_Id" & Environment.NewLine &  _
		                        "   , '" & strMaxUF__Creator_Id & "'  --Creator_Id" & Environment.NewLine &  _
		                        "   , '" & strMaxUF__Create_Date & "'  --Create_Date" & Environment.NewLine &  _
		                        "   , '" & strMaxUF__mmddDate & "'  --mmddDate" & Environment.NewLine &  _
		                        "   , '" & strMaxUF__Modified_By_Id & "'  --Modified_By_Id" & Environment.NewLine &  _
		                        "   , '" & strMaxUF__Last_Modify_Date & "' ) --Last_Modify_Date" & Environment.NewLine

		'Combine the strings
        Dim strTableInsert As String = strSQL1 & strSQL2

        'Execute the insert
		 SQLHelper.ExecuteSQL(cnMax, strTableInsert)

    End Sub
End Module