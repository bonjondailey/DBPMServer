Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data.SqlClient
Imports System.Windows.Forms
Module modMaxInsertClient

	'These are AMGRClient insert vars for InsertIntoAMGRClient
	Public strMaxC__Data_Machine_Id As String = ""
	Public strMaxC__Sequence_Number As String = ""
	Public strMaxC__Record_Type As String = ""
	Public strMaxC__Owner_Id As String = ""
	Public strMaxC__Private As String = ""
	Public strMaxC__Client_Id As String = ""
	Public strMaxC__Contact_Number As String = ""
	Public strMaxC__Name_Type As String = ""
	Public strMaxC__Name As String = ""
	Public strMaxC__Address_Id As String = ""
	Public strMaxC__Last_Modify_Date As String = ""
	Public strMaxC__Transfer_Date As String = ""
	Public strMaxC__Highest_Alt_Adr_Number As String = ""
	Public strMaxC__Phone_1 As String = ""
	Public strMaxC__Reverse_Phone_1 As String = ""
	Public strMaxC__Phone_1_Extension As String = ""
	Public strMaxC__Phone_2 As String = ""
	Public strMaxC__Reverse_Phone_2 As String = ""
	Public strMaxC__Phone_2_Extension As String = ""
	Public strMaxC__Phone_3 As String = ""
	Public strMaxC__Reverse_Phone_3 As String = ""
	Public strMaxC__Phone_3_Extension As String = ""
	Public strMaxC__Phone_4 As String = ""
	Public strMaxC__Reverse_Phone_4 As String = ""
	Public strMaxC__Phone_4_Extension As String = ""
	Public strMaxC__Highest_Contact_No As String = ""
	Public strMaxC__Receives_Letters As String = ""
	Public strMaxC__Use_Client_Name As String = ""
	Public strMaxC__First_Name As String = ""
	Public strMaxC__Initial As String = ""
	Public strMaxC__MrMs As String = ""
	Public strMaxC__Title As String = ""
	Public strMaxC__Salutation As String = ""
	Public strMaxC__Department As String = ""
	Public strMaxC__Firm As String = ""
	Public strMaxC__Division As String = ""
	Public strMaxC__Address_Line_1 As String = ""
	Public strMaxC__Address_Line_2 As String = ""
	Public strMaxC__City As String = ""
	Public strMaxC__State_Province As String = ""
	Public strMaxC__Country As String = ""
	Public strMaxC__Zip_Code As String = ""
	Public strMaxC__Change_Bits As String = ""
	Public strMaxC__Last_Client_Id As String = ""
	Public strMaxC__Record_Id As String = ""
	Public strMaxC__Creator_Id As String = ""
	Public strMaxC__Create_Date As String = ""
	Public strMaxC__Contact_Inherits_UDFs As String = ""
	Public strMaxC__Updated_By_Id As String = ""
	Public strMaxC__Reports_To_Contact_Number As String = ""
	Public strMaxC__Assigned_To As String = ""
	Public strMaxC__ReadPriv As String = ""
	Public strMaxC__ReadOnly_Id As String = ""
	Public strMaxC__Phone_1_Desc As String = ""
	Public strMaxC__Phone_2_Desc As String = ""
	Public strMaxC__Phone_3_Desc As String = ""
	Public strMaxC__Phone_4_Desc As String = ""
	Public strMaxC__Email_1_Desc As String = ""
	Public strMaxC__Email_2_Desc As String = ""
	Public strMaxC__Email_3_Desc As String = ""
	Public strMaxC__Lead_Status As String = ""

    '**************************************
    '******* 1st Code Review Complete *****
    '**************************************


    Sub InsertIntoAMGRClient()
        'This routine inserts data into AMGR_Client_Tbl table.



        'Build the SQL string
        Dim strSQL1 As String = "INSERT INTO AMGR_Client --_Tbl " & Environment.NewLine & _
                                "   ( Data_Machine_Id " & Environment.NewLine & _
                                "   , Sequence_Number " & Environment.NewLine & _
                                "   , Record_Type " & Environment.NewLine & _
                                "   , Owner_Id " & Environment.NewLine & _
                                "   , Private " & Environment.NewLine & _
                                "   , Client_Id " & Environment.NewLine & _
                                "   , Contact_Number " & Environment.NewLine & _
                                "   , Name_Type " & Environment.NewLine & _
                                "   , Name " & Environment.NewLine & _
                                "   , Address_Id " & Environment.NewLine & _
                                "   , Last_Modify_Date " & Environment.NewLine & _
                                "   , Transfer_Date " & Environment.NewLine & _
                                "   , Highest_Alt_Adr_Number " & Environment.NewLine & _
                                "   , Phone_1 " & Environment.NewLine & _
                                "   -- , Reverse_Phone_1 " & Environment.NewLine & _
                                "   , Phone_1_Extension " & Environment.NewLine & _
                                "   , Phone_2 " & Environment.NewLine & _
                                "   -- , Reverse_Phone_2 " & Environment.NewLine & _
                                "   , Phone_2_Extension " & Environment.NewLine
        Dim strSQL2 As String = "   , Phone_3 " & Environment.NewLine & _
                                "   -- , Reverse_Phone_3 " & Environment.NewLine & _
                                "   , Phone_3_Extension " & Environment.NewLine & _
                                "   , Phone_4 " & Environment.NewLine & _
                                "   -- , Reverse_Phone_4 " & Environment.NewLine & _
                                "   , Phone_4_Extension " & Environment.NewLine & _
                                "   , Highest_Contact_No " & Environment.NewLine & _
                                "   , Receives_Letters " & Environment.NewLine & _
                                "   , Use_Client_Name " & Environment.NewLine & _
                                "   , First_Name " & Environment.NewLine & _
                                "   , Initial " & Environment.NewLine & _
                                "   , MrMs " & Environment.NewLine & _
                                "   , Title " & Environment.NewLine & _
                                "   , Salutation " & Environment.NewLine & _
                                "   , Department " & Environment.NewLine & _
                                "   , Firm " & Environment.NewLine & _
                                "   , Division " & Environment.NewLine & _
                                "   , Address_Line_1 " & Environment.NewLine & _
                                "   , Address_Line_2 " & Environment.NewLine & _
                                "   , City " & Environment.NewLine & _
                                "   , State_Province " & Environment.NewLine
        Dim strSQL3 As String = "   , Country " & Environment.NewLine & _
                                "   , Zip_Code " & Environment.NewLine & _
                                "   -- , Change_Bits " & Environment.NewLine & _
                                "   , Last_Client_Id " & Environment.NewLine & _
                                "   , Record_Id " & Environment.NewLine & _
                                "   , Creator_Id " & Environment.NewLine & _
                                "   , Create_Date " & Environment.NewLine & _
                                "   , Contact_Inherits_UDFs " & Environment.NewLine & _
                                "   , Updated_By_Id " & Environment.NewLine & _
                                "   , Reports_To_Contact_Number " & Environment.NewLine & _
                                "   , Assigned_To " & Environment.NewLine & _
                                "   , ReadPriv " & Environment.NewLine & _
                                "   , ReadOnly_Id " & Environment.NewLine & _
                                "   , Phone_1_Desc " & Environment.NewLine & _
                                "   , Phone_2_Desc " & Environment.NewLine & _
                                "   , Phone_3_Desc " & Environment.NewLine & _
                                "   , Phone_4_Desc " & Environment.NewLine & _
                                "   , Email_1_Desc " & Environment.NewLine & _
                                "   , Email_2_Desc " & Environment.NewLine & _
                                "   , Email_3_Desc " & Environment.NewLine & _
                                "   , Lead_Status ) " & Environment.NewLine
        Dim strSQL4 As String = "VALUES " & Environment.NewLine & _
                                "   ( '" & strMaxC__Data_Machine_Id & "'  --Data_Machine_Id" & Environment.NewLine & _
                                "   , '" & strMaxC__Sequence_Number & "'  --Sequence_Number" & Environment.NewLine & _
                                "   , '" & strMaxC__Record_Type & "'  --Record_Type" & Environment.NewLine & _
                                "   , '" & strMaxC__Owner_Id & "'  --Owner_Id" & Environment.NewLine & _
                                "   , '" & strMaxC__Private & "'  --Private" & Environment.NewLine & _
                                "   , '" & strMaxC__Client_Id & "'  --Client_Id" & Environment.NewLine & _
                                "   , '" & strMaxC__Contact_Number & "'  --Contact_Number" & Environment.NewLine & _
                                "   , '" & strMaxC__Name_Type & "'  --Name_Type" & Environment.NewLine & _
                                "   , '" & strMaxC__Name & "'  --Name" & Environment.NewLine & _
                                "   , '" & strMaxC__Address_Id & "'  --Address_Id" & Environment.NewLine & _
                                "   , '" & strMaxC__Last_Modify_Date & "'  --Last_Modify_Date" & Environment.NewLine & _
                                "   , '" & strMaxC__Transfer_Date & "'  --Transfer_Date" & Environment.NewLine & _
                                "   , '" & strMaxC__Highest_Alt_Adr_Number & "'  --Highest_Alt_Adr_Number" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_1 & "'  --Phone_1" & Environment.NewLine & _
                                "   -- , '" & strMaxC__Reverse_Phone_1 & "'  --Reverse_Phone_1" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_1_Extension & "'  --Phone_1_Extension" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_2 & "'  --Phone_2" & Environment.NewLine & _
                                "   -- , '" & strMaxC__Reverse_Phone_2 & "'  --Reverse_Phone_2" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_2_Extension & "'  --Phone_2_Extension" & Environment.NewLine
        Dim strSQL5 As String = "   , '" & strMaxC__Phone_3 & "'  --Phone_3" & Environment.NewLine & _
                                "   -- , '" & strMaxC__Reverse_Phone_3 & "'  --Reverse_Phone_3" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_3_Extension & "'  --Phone_3_Extension" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_4 & "'  --Phone_4" & Environment.NewLine & _
                                "   -- , '" & strMaxC__Reverse_Phone_4 & "'  --Reverse_Phone_4" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_4_Extension & "'  --Phone_4_Extension" & Environment.NewLine & _
                                "   , '" & strMaxC__Highest_Contact_No & "'  --Highest_Contact_No" & Environment.NewLine & _
                                "   , '" & strMaxC__Receives_Letters & "'  --Receives_Letters" & Environment.NewLine & _
                                "   , '" & strMaxC__Use_Client_Name & "'  --Use_Client_Name" & Environment.NewLine & _
                                "   , '" & strMaxC__First_Name & "'  --First_Name" & Environment.NewLine & _
                                "   , '" & strMaxC__Initial & "'  --Initial" & Environment.NewLine & _
                                "   , '" & strMaxC__MrMs & "'  --MrMs" & Environment.NewLine & _
                                "   , '" & strMaxC__Title & "'  --Title" & Environment.NewLine & _
                                "   , '" & strMaxC__Salutation & "'  --Salutation" & Environment.NewLine & _
                                "   , '" & strMaxC__Department & "'  --Department" & Environment.NewLine & _
                                "   , '" & strMaxC__Firm & "'  --Firm" & Environment.NewLine & _
                                "   , '" & strMaxC__Division & "'  --Division" & Environment.NewLine & _
                                "   , '" & strMaxC__Address_Line_1 & "'  --Address_Line_1" & Environment.NewLine & _
                                "   , '" & strMaxC__Address_Line_2 & "'  --Address_Line_2" & Environment.NewLine & _
                                "   , '" & strMaxC__City & "'  --City" & Environment.NewLine & _
                                "   , '" & strMaxC__State_Province & "'  --State_Province" & Environment.NewLine
        Dim strSQL6 As String = "   , '" & strMaxC__Country & "'  --Country" & Environment.NewLine & _
                                "   , '" & strMaxC__Zip_Code & "'  --Zip_Code" & Environment.NewLine & _
                                "   -- , '" & strMaxC__Change_Bits & "'  --Change_Bits" & Environment.NewLine & _
                                "   , '" & strMaxC__Last_Client_Id & "'  --Last_Client_Id" & Environment.NewLine & _
                                "   , '" & strMaxC__Record_Id & "'  --Record_Id" & Environment.NewLine & _
                                "   , '" & strMaxC__Creator_Id & "'  --Creator_Id" & Environment.NewLine & _
                                "   , '" & strMaxC__Create_Date & "'  --Create_Date" & Environment.NewLine & _
                                "   , '" & strMaxC__Contact_Inherits_UDFs & "'  --Contact_Inherits_UDFs" & Environment.NewLine & _
                                "   , '" & strMaxC__Updated_By_Id & "'  --Updated_By_Id" & Environment.NewLine & _
                                "   , '" & strMaxC__Reports_To_Contact_Number & "'  --Reports_To_Contact_Number" & Environment.NewLine & _
                                "   , '" & strMaxC__Assigned_To & "'  --Assigned_To" & Environment.NewLine & _
                                "   , '" & strMaxC__ReadPriv & "'  --ReadPriv" & Environment.NewLine & _
                                "   , '" & strMaxC__ReadOnly_Id & "'  --ReadOnly_Id" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_1_Desc & "'  --Phone_1_Desc" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_2_Desc & "'  --Phone_2_Desc" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_3_Desc & "'  --Phone_3_Desc" & Environment.NewLine & _
                                "   , '" & strMaxC__Phone_4_Desc & "'  --Phone_4_Desc" & Environment.NewLine & _
                                "   , '" & strMaxC__Email_1_Desc & "'  --Email_1_Desc" & Environment.NewLine & _
                                "   , '" & strMaxC__Email_2_Desc & "'  --Email_2_Desc" & Environment.NewLine & _
                                "   , '" & strMaxC__Email_3_Desc & "'  --Email_3_Desc" & Environment.NewLine & _
                                "   , '" & strMaxC__Lead_Status & "' ) --Lead_Status" & Environment.NewLine


        'Combine the strings
        Dim strTableInsert As String = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
        'Execute the insert

        Try
            SQLHelper.ExecuteSQL(cnMax, strTableInsert)
        Catch ex As Exception
            HaveError("MaxInsertClient", "InsertIntoAMGRClient", CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
            Exit Try
        End Try

    End Sub
End Module