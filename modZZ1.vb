Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Globalization
Imports System.Windows.Forms

Imports DBPM_Server.siteConstants

'**************************************
'******* 1st Code Review Complete *****
'**************************************

Module modZZ1

	Public Sub InsertQBCustIntoMax() 'Except ones that just came from Max

		'Permission and ErrorHandling          (Auto built)
		Dim strObjName As String = "modZZ1" '"OBJNAME"
		Dim strSubName As String = "InsertQBCustIntoMax" '"SUBNAME"

		'Check permission to run
		If Not HavePermission(strObjName, strSubName) Then Exit Sub


		
		'FOR PART 1QBCust_ - Get records from QB_Customer
		Debug.WriteLine("List1QBCust_QB_Customer")
        Dim rs1QBCust_QB_Customer As DataSet


		Dim str1QBCust_QB_CustomerSQL, str1QBCust_QB_CustomerRow, str1QBCust_ListID, str1QBCust_TimeCreated, str1QBCust_TimeModified, str1QBCust_EditSequence, str1QBCust_Name, str1QBCust_FullName, str1QBCust_IsActive, str1QBCust_ParentRefListID, str1QBCust_ParentRefFullName, str1QBCust_Sublevel, str1QBCust_CompanyName, str1QBCust_Salutation, str1QBCust_FirstName, str1QBCust_MiddleName, str1QBCust_LastName, str1QBCust_BillAddressAddr1, str1QBCust_BillAddressAddr2, str1QBCust_BillAddressAddr3, str1QBCust_BillAddressAddr4, str1QBCust_BillAddressCity, str1QBCust_BillAddressState, str1QBCust_BillAddressPostalCode, str1QBCust_BillAddressCountry, str1QBCust_ShipAddressAddr1, str1QBCust_ShipAddressAddr2, str1QBCust_ShipAddressAddr3, str1QBCust_ShipAddressAddr4, str1QBCust_ShipAddressCity, str1QBCust_ShipAddressState, str1QBCust_ShipAddressPostalCode, str1QBCust_ShipAddressCountry, str1QBCust_Phone, str1QBCust_AltPhone, str1QBCust_Fax, str1QBCust_Email, str1QBCust_Contact, str1QBCust_AltContact, str1QBCust_CustomerTypeRefListID, str1QBCust_CustomerTypeRefFullName, str1QBCust_TermsRefListID, str1QBCust_TermsRefFullName, str1QBCust_SalesRepRefListID, str1QBCust_SalesRepRefFullName, str1QBCust_Balance, str1QBCust_TotalBalance, str1QBCust_OpenBalance, str1QBCust_OpenBalanceDate, str1QBCust_SalesTaxCodeRefListID, str1QBCust_SalesTaxCodeRefFullName, str1QBCust_ItemSalesTaxRefListID, str1QBCust_ItemSalesTaxRefFullName, str1QBCust_ResaleNumber, str1QBCust_AccountNumber, str1QBCust_CreditLimit, str1QBCust_PreferredPaymentMethodRefListID, str1QBCust_PreferredPaymentMethodRefFullName, str1QBCust_CreditCardInfoCreditCardNumber, str1QBCust_CreditCardInfoExpirationMonth, str1QBCust_CreditCardInfoExpirationYear, str1QBCust_CreditCardInfoNameOnCard, str1QBCust_CreditCardInfoCreditCardAddress, str1QBCust_CreditCardInfoCreditCardPostalCode, str1QBCust_JobStatus, str1QBCust_JobStartDate, str1QBCust_JobProjectedEndDate, str1QBCust_JobEndDate, str1QBCust_JobDesc, str1QBCust_JobTypeRefListID, str1QBCust_JobTypeRefFullName, str1QBCust_Notes, str1QBCust_PriceLevelRefListID, str1QBCust_PriceLevelRefFullName, str1QBCust_CustomFieldOther As String
		'This routine gets the 1QBCust_QB_Customer from the database according to the selection in str1QBCust_QB_CustomerSQL.
		'It then puts those 1QBCust_QB_Customer in the list box

		'FOR PART 2GetKey_ - Get Client_Id from AMGRClient
		'Debug.Print "List2GetKey_AMGRClient"
		Dim rs2GetKey_AMGRClient As DataSet
        Dim str2GetKey_AMGRClientSQL, str2GetKey_Client_Id As String
		'This routine gets the 2GetKey_AMGRClient from the database according to the selection in str2GetKey_AMGRClientSQL.
		'It then puts those 2GetKey_AMGRClient in the list box


        ShowUserMessage(strSubName, "RefreshQB: Adding QB_Customer to Max", "RefreshQB: Adding QB_Customer to Max", True)

        str1QBCust_QB_CustomerSQL = "SELECT * FROM QB_Customer WHERE ListID not in (SELECT Firm FROM AMGR_Client_Tbl WHERE Firm <> '') AND JobDesc not in (SELECT Client_Id FROM AMGR_Client_Tbl WHERE Name_Type = 'C'  AND Contact_Number = 0  AND Record_Type = 1) AND IsActive = 1 AND CustomerTypeRefFullName <> 'Local'"

        'Stop
        Debug.WriteLine(str1QBCust_QB_CustomerSQL)

        rs1QBCust_QB_Customer = SQLHelper.ExecuteDataSet(cnMax, CommandType.Text, str1QBCust_QB_CustomerSQL)

        Dim curRow As Integer = 0
        Dim rowCount As Integer = rs1QBCust_QB_Customer.Tables(0).Rows.Count

        If rowCount > 0 Then
            For Each iteration_row As DataRow In rs1QBCust_QB_Customer.Tables(0).Rows
                curRow += 1
                ShowUserMessage(strSubName, "Processing " & curRow.ToString & " of " & rowCount.ToString & " QB_Customers Into Max")

                Try

               
                'get the columns from the database
                str1QBCust_ListID = NCStr(iteration_row("ListID")).Replace("'"c, "`"c)

                str1QBCust_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                str1QBCust_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                str1QBCust_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                str1QBCust_Name = NCStr(iteration_row("Name")).Replace("'"c, "`"c)
                str1QBCust_FullName = NCStr(iteration_row("FullName")).Replace("'"c, "`"c)
                str1QBCust_IsActive = NCStr(iteration_row("IsActive")).Replace("'"c, "`"c)
                str1QBCust_ParentRefListID = NCStr(iteration_row("ParentRefListID")).Replace("'"c, "`"c)
                str1QBCust_ParentRefFullName = NCStr(iteration_row("ParentRefFullName")).Replace("'"c, "`"c)
                str1QBCust_Sublevel = NCStr(iteration_row("Sublevel")).Replace("'"c, "`"c)
                str1QBCust_CompanyName = NCStr(iteration_row("CompanyName")).Replace("'"c, "`"c)
                str1QBCust_Salutation = NCStr(iteration_row("Salutation")).Replace("'"c, "`"c)
                str1QBCust_FirstName = NCStr(iteration_row("FirstName")).Replace("'"c, "`"c)
                str1QBCust_MiddleName = NCStr(iteration_row("MiddleName")).Replace("'"c, "`"c)
                str1QBCust_LastName = NCStr(iteration_row("LastName")).Replace("'"c, "`"c)
                str1QBCust_BillAddressAddr1 = NCStr(iteration_row("BillAddressAddr1")).Replace("'"c, "`"c)
                str1QBCust_BillAddressAddr2 = NCStr(iteration_row("BillAddressAddr2")).Replace("'"c, "`"c)
                str1QBCust_BillAddressAddr3 = NCStr(iteration_row("BillAddressAddr3")).Replace("'"c, "`"c)
                str1QBCust_BillAddressAddr4 = NCStr(iteration_row("BillAddressAddr4")).Replace("'"c, "`"c)
                str1QBCust_BillAddressCity = NCStr(iteration_row("BillAddressCity")).Replace("'"c, "`"c)
                str1QBCust_BillAddressState = NCStr(iteration_row("BillAddressState")).Replace("'"c, "`"c)
                str1QBCust_BillAddressPostalCode = NCStr(iteration_row("BillAddressPostalCode")).Replace("'"c, "`"c)
                str1QBCust_BillAddressCountry = NCStr(iteration_row("BillAddressCountry")).Replace("'"c, "`"c)
                str1QBCust_ShipAddressAddr1 = NCStr(iteration_row("ShipAddressAddr1")).Replace("'"c, "`"c)
                str1QBCust_ShipAddressAddr2 = NCStr(iteration_row("ShipAddressAddr2")).Replace("'"c, "`"c)
                str1QBCust_ShipAddressAddr3 = NCStr(iteration_row("ShipAddressAddr3")).Replace("'"c, "`"c)
                str1QBCust_ShipAddressAddr4 = NCStr(iteration_row("ShipAddressAddr4")).Replace("'"c, "`"c)
                str1QBCust_ShipAddressCity = NCStr(iteration_row("ShipAddressCity")).Replace("'"c, "`"c)
                str1QBCust_ShipAddressState = NCStr(iteration_row("ShipAddressState")).Replace("'"c, "`"c)
                str1QBCust_ShipAddressPostalCode = NCStr(iteration_row("ShipAddressPostalCode")).Replace("'"c, "`"c)
                str1QBCust_ShipAddressCountry = NCStr(iteration_row("ShipAddressCountry")).Replace("'"c, "`"c)
                str1QBCust_Phone = NCStr(iteration_row("Phone")).Replace("'"c, "`"c)
                str1QBCust_AltPhone = NCStr(iteration_row("AltPhone")).Replace("'"c, "`"c)
                str1QBCust_Fax = NCStr(iteration_row("Fax")).Replace("'"c, "`"c)
                str1QBCust_Email = NCStr(iteration_row("Email")).Replace("'"c, "`"c)
                str1QBCust_Contact = NCStr(iteration_row("Contact")).Replace("'"c, "`"c)
                str1QBCust_AltContact = NCStr(iteration_row("AltContact")).Replace("'"c, "`"c)
                str1QBCust_CustomerTypeRefListID = NCStr(iteration_row("CustomerTypeRefListID")).Replace("'"c, "`"c)
                str1QBCust_CustomerTypeRefFullName = NCStr(iteration_row("CustomerTypeRefFullName")).Replace("'"c, "`"c)
                str1QBCust_TermsRefListID = NCStr(iteration_row("TermsRefListID")).Replace("'"c, "`"c)
                str1QBCust_TermsRefFullName = NCStr(iteration_row("TermsRefFullName")).Replace("'"c, "`"c)
                str1QBCust_SalesRepRefListID = NCStr(iteration_row("SalesRepRefListID")).Replace("'"c, "`"c)
                str1QBCust_SalesRepRefFullName = NCStr(iteration_row("SalesRepRefFullName")).Replace("'"c, "`"c)
                str1QBCust_Balance = NCStr(iteration_row("Balance")).Replace("'"c, "`"c)
                str1QBCust_TotalBalance = NCStr(iteration_row("TotalBalance")).Replace("'"c, "`"c)
                str1QBCust_OpenBalance = NCStr(iteration_row("OpenBalance")).Replace("'"c, "`"c)
                str1QBCust_OpenBalanceDate = NCStr(iteration_row("OpenBalanceDate")).Replace("'"c, "`"c)
                str1QBCust_SalesTaxCodeRefListID = NCStr(iteration_row("SalesTaxCodeRefListID")).Replace("'"c, "`"c)
                str1QBCust_SalesTaxCodeRefFullName = NCStr(iteration_row("SalesTaxCodeRefFullName")).Replace("'"c, "`"c)
                str1QBCust_ItemSalesTaxRefListID = NCStr(iteration_row("ItemSalesTaxRefListID")).Replace("'"c, "`"c)
                str1QBCust_ItemSalesTaxRefFullName = NCStr(iteration_row("ItemSalesTaxRefFullName")).Replace("'"c, "`"c)
                str1QBCust_ResaleNumber = NCStr(iteration_row("ResaleNumber")).Replace("'"c, "`"c)
                str1QBCust_AccountNumber = NCStr(iteration_row("AccountNumber")).Replace("'"c, "`"c)
                str1QBCust_CreditLimit = NCStr(iteration_row("CreditLimit")).Replace("'"c, "`"c)
                str1QBCust_PreferredPaymentMethodRefListID = NCStr(iteration_row("PreferredPaymentMethodRefListID")).Replace("'"c, "`"c)
                str1QBCust_PreferredPaymentMethodRefFullName = NCStr(iteration_row("PreferredPaymentMethodRefFullName")).Replace("'"c, "`"c)
                str1QBCust_CreditCardInfoCreditCardNumber = NCStr(iteration_row("CreditCardInfoCreditCardNumber")).Replace("'"c, "`"c)
                str1QBCust_CreditCardInfoExpirationMonth = NCStr(iteration_row("CreditCardInfoExpirationMonth")).Replace("'"c, "`"c)
                str1QBCust_CreditCardInfoExpirationYear = NCStr(iteration_row("CreditCardInfoExpirationYear")).Replace("'"c, "`"c)
                str1QBCust_CreditCardInfoNameOnCard = NCStr(iteration_row("CreditCardInfoNameOnCard")).Replace("'"c, "`"c)
                str1QBCust_CreditCardInfoCreditCardAddress = NCStr(iteration_row("CreditCardInfoCreditCardAddress")).Replace("'"c, "`"c)
                str1QBCust_CreditCardInfoCreditCardPostalCode = NCStr(iteration_row("CreditCardInfoCreditCardPostalCode")).Replace("'"c, "`"c)
                str1QBCust_JobStatus = NCStr(iteration_row("JobStatus")).Replace("'"c, "`"c)
                str1QBCust_JobStartDate = NCStr(iteration_row("JobStartDate")).Replace("'"c, "`"c)
                str1QBCust_JobProjectedEndDate = NCStr(iteration_row("JobProjectedEndDate")).Replace("'"c, "`"c)
                str1QBCust_JobEndDate = NCStr(iteration_row("JobEndDate")).Replace("'"c, "`"c)
                str1QBCust_JobDesc = NCStr(iteration_row("JobDesc")).Replace("'"c, "`"c)
                str1QBCust_JobTypeRefListID = NCStr(iteration_row("JobTypeRefListID")).Replace("'"c, "`"c)
                str1QBCust_JobTypeRefFullName = NCStr(iteration_row("JobTypeRefFullName")).Replace("'"c, "`"c)
                str1QBCust_Notes = NCStr(iteration_row("Notes")).Replace("'"c, "`"c)
                str1QBCust_PriceLevelRefListID = NCStr(iteration_row("PriceLevelRefListID")).Replace("'"c, "`"c)
                str1QBCust_PriceLevelRefFullName = NCStr(iteration_row("PriceLevelRefFullName")).Replace("'"c, "`"c)
                str1QBCust_CustomFieldOther = NCStr(iteration_row("CustomFieldOther")).Replace("'"c, "`"c)

                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str1QBCust_QB_CustomerRow = "" & _
                                            str1QBCust_ListID & "  | " & _
                                            str1QBCust_TimeCreated & "  | " & _
                                            str1QBCust_TimeModified & "  | " & _
                                            str1QBCust_EditSequence & "  | " & _
                                            str1QBCust_Name & "  | " & _
                                            str1QBCust_FullName & "  | " & _
                                            str1QBCust_IsActive & "  | " & _
                                            str1QBCust_ParentRefListID & "  | " & _
                                            str1QBCust_ParentRefFullName & "  | " & _
                                            str1QBCust_Sublevel & "  | " & _
                                            str1QBCust_CompanyName & "  | " & _
                                            str1QBCust_Salutation & "  | " & _
                                            str1QBCust_FirstName & "  | " & _
                                            str1QBCust_MiddleName & "  | " & _
                                            str1QBCust_LastName & "  | " & _
                                            "" & Strings.Chr(9)


                ShowUserMessage(strSubName, str1QBCust_QB_CustomerRow)

                Dim dbNumericTemp As Double
                If Not Double.TryParse(str1QBCust_AccountNumber.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture.NumberFormat, dbNumericTemp) Then str1QBCust_AccountNumber = ""

                'Public vars to fill for InsertIntoAMGR_Client
                strMaxC__Data_Machine_Id = "0" 'Auto
                strMaxC__Sequence_Number = "34" 'Auto
                strMaxC__Record_Type = "1" '1=Company  2=Individual?
                strMaxC__Owner_Id = ""
                strMaxC__Private = "0"
                strMaxC__Client_Id = ""
                strMaxC__Contact_Number = "0" 'Auto
                strMaxC__Name_Type = "C"
                strMaxC__Name = str1QBCust_Name
                strMaxC__Address_Id = "0"
                strMaxC__Last_Modify_Date = str1QBCust_TimeModified 'Auto
                strMaxC__Transfer_Date = "" ' Null
                strMaxC__Highest_Alt_Adr_Number = "0"
                strMaxC__Phone_1 = str1QBCust_Phone
                'strMaxC__Reverse_Phone_1 = ""
                strMaxC__Phone_1_Extension = ""
                strMaxC__Phone_2 = str1QBCust_Fax
                'strMaxC__Reverse_Phone_2 = ""
                strMaxC__Phone_2_Extension = ""
                strMaxC__Phone_3 = str1QBCust_AltPhone
                'strMaxC__Reverse_Phone_3 = ""
                strMaxC__Phone_3_Extension = ""
                strMaxC__Phone_4 = str1QBCust_AccountNumber
                'strMaxC__Reverse_Phone_4 = ""
                strMaxC__Phone_4_Extension = ""
                strMaxC__Highest_Contact_No = "0" 'Auto
                strMaxC__Receives_Letters = "1"
                strMaxC__Use_Client_Name = "0"
                strMaxC__First_Name = "" 'ContactInfo
                strMaxC__Initial = ""
                strMaxC__MrMs = ""
                strMaxC__Title = "BillTo" '""
                strMaxC__Salutation = ""
                strMaxC__Department = ""
                strMaxC__Firm = str1QBCust_ListID '""      
                strMaxC__Division = ""
                strMaxC__Address_Line_1 = str1QBCust_BillAddressAddr2
                strMaxC__Address_Line_2 = str1QBCust_BillAddressAddr3
                strMaxC__City = str1QBCust_BillAddressCity
                strMaxC__State_Province = str1QBCust_BillAddressState
                strMaxC__Country = str1QBCust_BillAddressCountry
                strMaxC__Zip_Code = str1QBCust_BillAddressPostalCode
                strMaxC__Change_Bits = ""
                strMaxC__Last_Client_Id = ""
                strMaxC__Record_Id = "14" 'Auto?
                strMaxC__Creator_Id = "MASTER" 'Auto
                strMaxC__Create_Date = str1QBCust_TimeCreated 'Auto
                strMaxC__Contact_Inherits_UDFs = ""
                strMaxC__Updated_By_Id = "MASTER" 'Auto
                strMaxC__Reports_To_Contact_Number = "0"
                strMaxC__Assigned_To = ""
                strMaxC__ReadPriv = "0"
                strMaxC__ReadOnly_Id = ""
                strMaxC__Phone_1_Desc = "Main"
                strMaxC__Phone_2_Desc = "Fax"
                strMaxC__Phone_3_Desc = "Cell" '"800#"
                strMaxC__Phone_4_Desc = "Acct#" 'str1QBCust_ListID
                strMaxC__Email_1_Desc = "Email"
                strMaxC__Email_2_Desc = "Email 2"
                strMaxC__Email_3_Desc = "Email 3"
                strMaxC__Lead_Status = "0" '0=Contact  1=Lead

                InsertIntoAMGRClient()

                'GET THE COMPANY KEY THAT WAS ADDED BY MAX
                'Use Client_Id added by Max to add/update the multi program customer x-ref table
                'Use Client_Id added by Max to add Contacts if there are any

                'PART 2GetKey_: Get the Client_Id from SQL


                str2GetKey_AMGRClientSQL = "SELECT Client_Id FROM AMGR_Client_Tbl WHERE Firm = '" & str1QBCust_ListID & "' AND Title = 'BillTo'"
                rs2GetKey_AMGRClient = SQLHelper.ExecuteDataSet(cnMax, CommandType.Text, str2GetKey_AMGRClientSQL)


                    If rs2GetKey_AMGRClient.Tables(0).Rows.Count > 1 Then
                        HaveError(strObjName, strSubName, "", "Attempted to INSERT company that already existed into MAX", "", "", "")
                    End If

                    If rs2GetKey_AMGRClient.Tables(0).Rows.Count = 1 Then
                        For Each iteration_row_2 As DataRow In rs2GetKey_AMGRClient.Tables(0).Rows
                            'Clear strings
                            str2GetKey_Client_Id = ""
                            'get the columns from the database
                            str2GetKey_Client_Id = NCStr(iteration_row_2("Client_Id"))
                            'Debug.Print "Client_Id = " & str2GetKey_Client_Id


                            If str1QBCust_Email.Trim() <> "" Then
                                strMaxUF__Client_Id = str2GetKey_Client_Id
                                strMaxUF__Contact_Number = "0" '(Company)
                                strMaxUF__Type_Id = "58850"
                                strMaxUF__Code_Id = "0"
                                strMaxUF__Last_Code_Id = "0"
                                strMaxUF__DateCol = Now.ToString '""
                                strMaxUF__NumericCol = "0"
                                strMaxUF__AlphaNumericCol = str1QBCust_Email '"erf@aristotle.net"
                                strMaxUF__Record_Id = "0" 'Auto I think
                                strMaxUF__Creator_Id = "MASTER"
                                strMaxUF__Create_Date = Now.ToString
                                strMaxUF__mmddDate = "" '"2525"
                                strMaxUF__Modified_By_Id = "MASTER"
                                strMaxUF__Last_Modify_Date = Now.ToString

                                InsertIntoAMGRUserFields()
                            End If

                            strMaxC__Data_Machine_Id = "0" 'Auto
                            strMaxC__Sequence_Number = "34" 'Auto
                            strMaxC__Record_Type = "3" '"1"        '1=Company  2=Individual?  3=Contact?
                            strMaxC__Owner_Id = ""
                            strMaxC__Private = "0"
                            strMaxC__Client_Id = str2GetKey_Client_Id '""
                            strMaxC__Contact_Number = "1" '"0"     'Auto
                            strMaxC__Name_Type = "I" '"C"
                            strMaxC__Name = "" 'Trim(str1QBCust_LastName) 'str1QBCust_Name
                            strMaxC__Address_Id = "0" '"60001" '"0"
                            strMaxC__Last_Modify_Date = str1QBCust_TimeModified 'Auto
                            strMaxC__Transfer_Date = "" ' Null
                            strMaxC__Highest_Alt_Adr_Number = "60000" '"0"
                            strMaxC__Phone_1 = str1QBCust_Phone '"" for client?
                            'strMaxC__Reverse_Phone_1 = ""
                            strMaxC__Phone_1_Extension = ""
                            strMaxC__Phone_2 = str1QBCust_Fax '"" for client?
                            'strMaxC__Reverse_Phone_2 = ""
                            strMaxC__Phone_2_Extension = ""
                            strMaxC__Phone_3 = str1QBCust_AltPhone '"" for client?
                            'strMaxC__Reverse_Phone_3 = ""
                            strMaxC__Phone_3_Extension = ""
                            strMaxC__Phone_4 = str1QBCust_AccountNumber
                            'strMaxC__Reverse_Phone_4 = ""
                            strMaxC__Phone_4_Extension = ""
                            strMaxC__Highest_Contact_No = "0" 'Auto
                            strMaxC__Receives_Letters = "1"
                            strMaxC__Use_Client_Name = "1" '"0"
                            strMaxC__First_Name = "" 'Trim(str1QBCust_FirstName) '""           'ContactInfo
                            strMaxC__Initial = "" 'Trim(str1QBCust_MiddleName) '""
                            strMaxC__MrMs = ""
                            strMaxC__Title = "Accounting" '"BillTo"
                            strMaxC__Salutation = ""
                            strMaxC__Department = ""
                            strMaxC__Firm = "" 'str1QBCust_ListID '""       'Are you certain?
                            strMaxC__Division = ""
                            strMaxC__Address_Line_1 = "" 'str1QBCust_BillAddressAddr2
                            strMaxC__Address_Line_2 = "" 'str1QBCust_BillAddressAddr3
                            strMaxC__City = "" 'str1QBCust_BillAddressCity
                            strMaxC__State_Province = "" 'str1QBCust_BillAddressState
                            strMaxC__Country = "" 'str1QBCust_BillAddressCountry
                            strMaxC__Zip_Code = "" 'str1QBCust_BillAddressPostalCode
                            strMaxC__Change_Bits = ""
                            strMaxC__Last_Client_Id = ""
                            strMaxC__Record_Id = "14" 'Auto key
                            strMaxC__Creator_Id = "MASTER" 'Auto
                            strMaxC__Create_Date = str1QBCust_TimeCreated 'Auto
                            strMaxC__Contact_Inherits_UDFs = ""
                            strMaxC__Updated_By_Id = "MASTER" 'Auto
                            strMaxC__Reports_To_Contact_Number = "0"
                            strMaxC__Assigned_To = ""
                            strMaxC__ReadPriv = "0"
                            strMaxC__ReadOnly_Id = ""
                            strMaxC__Phone_1_Desc = "Main"
                            strMaxC__Phone_2_Desc = "Fax"
                            strMaxC__Phone_3_Desc = "Alt"
                            strMaxC__Phone_4_Desc = "Acct#" 'str1QBCust_ListID
                            strMaxC__Email_1_Desc = "Email"
                            strMaxC__Email_2_Desc = "Email 2"
                            strMaxC__Email_3_Desc = "Email 3"
                            strMaxC__Lead_Status = "0" '0=Contact  1=Lead


                            'Insert Contact info
                            If str1QBCust_FirstName.Trim() <> "" Or str1QBCust_MiddleName.Trim() <> "" Or str1QBCust_LastName.Trim() <> "" Then
                                'Stop 'MsgBox "Found FirstMiddleLastName:  " & Trim(str1QBCust_FirstName)
                                strMaxC__MrMs = str1QBCust_Salutation
                                strMaxC__Name = str1QBCust_LastName.Trim() 'str1QBCust_Name
                                strMaxC__First_Name = str1QBCust_FirstName.Trim() '""           'ContactInfo
                                strMaxC__Initial = str1QBCust_MiddleName.Trim() '""
                                'Insert Contact
                                InsertIntoAMGRClient()
                            Else
                                If str1QBCust_Contact.Trim() <> "" Then
                                    'Stop 'MsgBox "Found A Contact:  " & Trim(str1QBCust_Contact)
                                    strMaxC__MrMs = ""
                                    strMaxC__Name = str1QBCust_Contact.Trim() 'str1QBCust_Name
                                    strMaxC__First_Name = ""
                                    strMaxC__Initial = ""
                                    'Insert Contact
                                    InsertIntoAMGRClient()
                                End If
                            End If

                            If str1QBCust_AltContact.Trim() <> "" Then
                                'Stop 'MsgBox "Found A AltContact:  " & Trim(str1QBCust_AltContact)
                                strMaxC__MrMs = ""
                                strMaxC__Name = str1QBCust_AltContact.Trim() 'str1QBCust_Name
                                strMaxC__First_Name = ""
                                strMaxC__Initial = ""
                                'Insert Contact
                                InsertIntoAMGRClient()
                            End If


                        Next iteration_row_2
                    End If

                Catch ex As Exception
                    HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                    Continue For
                End Try
            Next iteration_row

        End If

        'Update status listbox
        ShowUserMessage(strSubName, "Records Added to Max From QB", , True)

    End Sub
End Module