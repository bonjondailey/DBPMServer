Option Strict Off
Option Explicit On
Imports DBPM_Server.siteConstants
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Windows.Forms
Module modZZ0

    '**********************************
    '*** FIRST CODE REVIEW COMPLETE ***
    '**********************************



    Public Sub InsertMaxBillToIntoQB()
        Dim strAccountNumber, strAltContact, strAltPhone, strBalance, strBillAddressAddr1, strBillAddressAddr2, strBillAddressAddr3, strBillAddressAddr4, strBillAddressCity, strBillAddressCountry, strBillAddressPostalCode, strBillAddressState, strCompanyName, strContact, strCreditCardInfoCreditCardAddress, strCreditCardInfoCreditCardNumber, strCreditCardInfoCreditCardPostalCode, strCreditCardInfoExpirationMonth, strCreditCardInfoExpirationYear, strCreditCardInfoNameOnCard, strCreditLimitQ, strCustomerTypeRefListID, strCustomFieldOther As String
        Dim strEditSequence, strEmail, strErrorLine, strFax, strFirstName, strFQSaveToCache, strFullName, strIsActive, strItemSalesTaxRefFullName, strItemSalesTaxRefListID, strJobDesc, strJobEndDate, strJobProjectedEndDate, strJobStartDate, strJobStatus, strJobTypeRefFullName, strJobTypeRefListID, strLastName, strListID, strMiddleName, strName, strNotes, strOpenBalance, strOpenBalanceDate, strParentRefFullName, strParentRefListID, strPhone, strPreferredPaymentMethodRefFullName, strPreferredPaymentMethodRefListID, strPriceLevelRefFullName, strPriceLevelRefListID, strResaleNumber, strSalesRepRefFullName, strSalesRepRefListID, strSalesTaxCodeRefFullName, strSalesTaxCodeRefListID, strSalutation, strShipAddressAddr1, strShipAddressAddr2, strShipAddressAddr3, strShipAddressAddr4, strShipAddressCity, strShipAddressCountry, strShipAddressPostalCode, strShipAddressState, strState, strSublevel, strTermsRefFullName, strTermsRefListID, strTimeCreated, strTimeModified, strTotalBalance As String

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modZZ0" '"OBJNAME"
        Dim strSubName As String = "InsertMaxBillToIntoQB" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub


        'FOR PART 1MaxBillTo_ - Get records from AMGR_Client_Tbl
        Debug.WriteLine("List1MaxBillTo_AMGR_Client_Tbl")
        Dim str1MaxBillTo_AMGR_Client_TblSQL, str1MaxBillTo_Data_Machine_Id, str1MaxBillTo_Sequence_Number, str1MaxBillTo_Record_Type, str1MaxBillTo_Owner_Id, str1MaxBillTo_Private, str1MaxBillTo_Client_Id, str1MaxBillTo_Contact_Number, str1MaxBillTo_Name_Type, str1MaxBillTo_Name, str1MaxBillTo_Address_Id, str1MaxBillTo_Last_Modify_Date, str1MaxBillTo_Transfer_Date, str1MaxBillTo_Highest_Alt_Adr_Number, str1MaxBillTo_Phone_1, str1MaxBillTo_Reverse_Phone_1, str1MaxBillTo_Phone_1_Extension, str1MaxBillTo_Phone_2, str1MaxBillTo_Reverse_Phone_2, str1MaxBillTo_Phone_2_Extension, str1MaxBillTo_Phone_3, str1MaxBillTo_Reverse_Phone_3, str1MaxBillTo_Phone_3_Extension, str1MaxBillTo_Phone_4, str1MaxBillTo_Reverse_Phone_4, str1MaxBillTo_Phone_4_Extension, str1MaxBillTo_Highest_Contact_No, str1MaxBillTo_Receives_Letters, str1MaxBillTo_Use_Client_Name, str1MaxBillTo_First_Name, str1MaxBillTo_Initial, str1MaxBillTo_MrMs, str1MaxBillTo_Title, str1MaxBillTo_Salutation, str1MaxBillTo_Department, str1MaxBillTo_Firm, str1MaxBillTo_Division, str1MaxBillTo_Address_Line_1, str1MaxBillTo_Address_Line_2, str1MaxBillTo_City, str1MaxBillTo_State_Province, str1MaxBillTo_Country, str1MaxBillTo_Zip_Code, str1MaxBillTo_Last_Client_Id, str1MaxBillTo_Record_Id, str1MaxBillTo_Creator_Id, str1MaxBillTo_Create_Date, str1MaxBillTo_Contact_Inherits_UDFs, str1MaxBillTo_Updated_By_Id, str1MaxBillTo_Reports_To_Contact_Number, str1MaxBillTo_Assigned_To, str1MaxBillTo_ReadPriv, str1MaxBillTo_ReadOnly_Id, str1MaxBillTo_Phone_1_Desc, str1MaxBillTo_Phone_2_Desc, str1MaxBillTo_Phone_3_Desc, str1MaxBillTo_Phone_4_Desc, str1MaxBillTo_Email_1_Desc, str1MaxBillTo_Email_2_Desc, str1MaxBillTo_Email_3_Desc, str1MaxBillTo_Lead_Status As String
        'This routine gets the 1MaxBillTo_AMGR_Client_Tbl from the database according to the selection in str1MaxBillTo_AMGR_Client_TblSQL.
        'It then puts those 1MaxBillTo_AMGR_Client_Tbl in the list box

        str1MaxBillTo_AMGR_Client_TblSQL = "SELECT * FROM AMGR_Client_Tbl WHERE Title = 'BillTo'  AND Firm = ''  AND Name_Type = 'C'  AND Contact_Number = 0  AND Record_Type = 1"
        'this gets items that the system does not believe are in quickbooks (because Firm is empty)
        'it gets those that are the top level entry (BillTo), and Name_Type is a Company rather than individual

        Debug.WriteLine(str1MaxBillTo_AMGR_Client_TblSQL)

        ShowUserMessage(strSubName, "Inserting MAX Clients into Quickbooks", "Inserting MAX Clients into Quickbooks", True)

        Using rs1MaxBillTo_AMGR_Client_Tbl As DataSet = SQLHelper.ExecuteDataSet(cnMax, CommandType.Text, str1MaxBillTo_AMGR_Client_TblSQL)

            Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert As String
            Dim curRow As Integer = 0
            Dim rowCount As Integer = rs1MaxBillTo_AMGR_Client_Tbl.Tables(0).Rows.Count
            Dim iRecordsProcessed As Integer = 0
            Dim booErrorAddingQBCust As Boolean

            If rowCount > 0 Then

                ShowUserMessage(strSubName, "Processing  " & rowCount & "  BillTo Records", strSubName)

                For Each iteration_row As DataRow In rs1MaxBillTo_AMGR_Client_Tbl.Tables(0).Rows
                    curRow += 1
                    ShowUserMessage(strSubName, "Processing Record " & curRow.ToString & " of " & rowCount.ToString)

                    'get the columns from the database
                    str1MaxBillTo_Data_Machine_Id = NCStr(iteration_row("Data_Machine_Id")).Replace("'"c, "`"c)
                    str1MaxBillTo_Sequence_Number = NCStr(iteration_row("Sequence_Number")).Replace("'"c, "`"c)
                    str1MaxBillTo_Record_Type = NCStr(iteration_row("Record_Type")).Replace("'"c, "`"c)
                    str1MaxBillTo_Owner_Id = NCStr(iteration_row("Owner_Id")).Replace("'"c, "`"c)
                    str1MaxBillTo_Private = NCStr(iteration_row("Private")).Replace("'"c, "`"c)
                    str1MaxBillTo_Client_Id = NCStr(iteration_row("Client_Id")).Replace("'"c, "`"c)
                    str1MaxBillTo_Contact_Number = NCStr(iteration_row("Contact_Number")).Replace("'"c, "`"c)
                    str1MaxBillTo_Name_Type = NCStr(iteration_row("Name_Type")).Replace("'"c, "`"c)
                    str1MaxBillTo_Name = NCStr(iteration_row("Name")).Replace("'"c, "`"c)
                    str1MaxBillTo_Address_Id = NCStr(iteration_row("Address_Id")).Replace("'"c, "`"c)
                    str1MaxBillTo_Last_Modify_Date = NCStr(iteration_row("Last_Modify_Date")).Replace("'"c, "`"c)
                    str1MaxBillTo_Transfer_Date = NCStr(iteration_row("Transfer_Date")).Replace("'"c, "`"c)
                    str1MaxBillTo_Highest_Alt_Adr_Number = NCStr(iteration_row("Highest_Alt_Adr_Number")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_1 = NCStr(iteration_row("Phone_1")).Replace("'"c, "`"c)
                    str1MaxBillTo_Reverse_Phone_1 = NCStr(iteration_row("Reverse_Phone_1")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_1_Extension = NCStr(iteration_row("Phone_1_Extension")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_2 = NCStr(iteration_row("Phone_2")).Replace("'"c, "`"c)
                    str1MaxBillTo_Reverse_Phone_2 = NCStr(iteration_row("Reverse_Phone_2")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_2_Extension = NCStr(iteration_row("Phone_2_Extension")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_3 = NCStr(iteration_row("Phone_3")).Replace("'"c, "`"c)
                    str1MaxBillTo_Reverse_Phone_3 = NCStr(iteration_row("Reverse_Phone_3")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_3_Extension = NCStr(iteration_row("Phone_3_Extension")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_4 = NCStr(iteration_row("Phone_4")).Replace("'"c, "`"c)
                    str1MaxBillTo_Reverse_Phone_4 = NCStr(iteration_row("Reverse_Phone_4")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_4_Extension = NCStr(iteration_row("Phone_4_Extension")).Replace("'"c, "`"c)
                    str1MaxBillTo_Highest_Contact_No = NCStr(iteration_row("Highest_Contact_No")).Replace("'"c, "`"c)
                    str1MaxBillTo_Receives_Letters = NCStr(iteration_row("Receives_Letters")).Replace("'"c, "`"c)
                    str1MaxBillTo_Use_Client_Name = NCStr(iteration_row("Use_Client_Name")).Replace("'"c, "`"c)
                    str1MaxBillTo_First_Name = NCStr(iteration_row("First_Name")).Replace("'"c, "`"c)
                    str1MaxBillTo_Initial = NCStr(iteration_row("Initial")).Replace("'"c, "`"c)
                    str1MaxBillTo_MrMs = NCStr(iteration_row("MrMs")).Replace("'"c, "`"c)
                    str1MaxBillTo_Title = NCStr(iteration_row("Title")).Replace("'"c, "`"c)
                    str1MaxBillTo_Salutation = NCStr(iteration_row("Salutation")).Replace("'"c, "`"c)
                    str1MaxBillTo_Department = NCStr(iteration_row("Department")).Replace("'"c, "`"c)
                    str1MaxBillTo_Firm = NCStr(iteration_row("Firm")).Replace("'"c, "`"c)
                    str1MaxBillTo_Division = NCStr(iteration_row("Division")).Replace("'"c, "`"c)
                    str1MaxBillTo_Address_Line_1 = NCStr(iteration_row("Address_Line_1")).Replace("'"c, "`"c)
                    str1MaxBillTo_Address_Line_2 = NCStr(iteration_row("Address_Line_2")).Replace("'"c, "`"c)
                    str1MaxBillTo_City = NCStr(iteration_row("City")).Replace("'"c, "`"c)
                    str1MaxBillTo_State_Province = NCStr(iteration_row("State_Province")).Replace("'"c, "`"c)
                    str1MaxBillTo_Country = NCStr(iteration_row("Country")).Replace("'"c, "`"c)
                    str1MaxBillTo_Zip_Code = NCStr(iteration_row("Zip_Code")).Replace("'"c, "`"c)
                    'If rs1MaxBillTo_AMGR_Client_Tbl!Change_Bits <> "" Then str1MaxBillTo_Change_Bits = rs1MaxBillTo_AMGR_Client_Tbl!Change_Bits
                    str1MaxBillTo_Last_Client_Id = NCStr(iteration_row("Last_Client_Id")).Replace("'"c, "`"c)
                    str1MaxBillTo_Record_Id = NCStr(iteration_row("Record_Id")).Replace("'"c, "`"c)
                    str1MaxBillTo_Creator_Id = NCStr(iteration_row("Creator_Id")).Replace("'"c, "`"c)
                    str1MaxBillTo_Create_Date = NCStr(iteration_row("Create_Date")).Replace("'"c, "`"c)
                    str1MaxBillTo_Contact_Inherits_UDFs = NCStr(iteration_row("Contact_Inherits_UDFs")).Replace("'"c, "`"c)
                    str1MaxBillTo_Updated_By_Id = NCStr(iteration_row("Updated_By_Id")).Replace("'"c, "`"c)
                    str1MaxBillTo_Reports_To_Contact_Number = NCStr(iteration_row("Reports_To_Contact_Number")).Replace("'"c, "`"c)
                    str1MaxBillTo_Assigned_To = NCStr(iteration_row("Assigned_To")).Replace("'"c, "`"c)
                    str1MaxBillTo_ReadPriv = NCStr(iteration_row("ReadPriv")).Replace("'"c, "`"c)
                    str1MaxBillTo_ReadOnly_Id = NCStr(iteration_row("ReadOnly_Id")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_1_Desc = NCStr(iteration_row("Phone_1_Desc")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_2_Desc = NCStr(iteration_row("Phone_2_Desc")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_3_Desc = NCStr(iteration_row("Phone_3_Desc")).Replace("'"c, "`"c)
                    str1MaxBillTo_Phone_4_Desc = NCStr(iteration_row("Phone_4_Desc")).Replace("'"c, "`"c)
                    str1MaxBillTo_Email_1_Desc = NCStr(iteration_row("Email_1_Desc")).Replace("'"c, "`"c)
                    str1MaxBillTo_Email_2_Desc = NCStr(iteration_row("Email_2_Desc")).Replace("'"c, "`"c)
                    str1MaxBillTo_Email_3_Desc = NCStr(iteration_row("Email_3_Desc")).Replace("'"c, "`"c)
                    str1MaxBillTo_Lead_Status = NCStr(iteration_row("Lead_Status")).Replace("'"c, "`"c)

                    'Strip quote character out of strings
                    'Get quote characters out!
                    'Change Quote to reverse quote
                    'If KeyAscii = 39 Then KeyAscii = 96

                    strListID = ""
                    strTimeCreated = ""
                    strTimeModified = ""
                    strEditSequence = ""

                    strName = Strings.Mid(str1MaxBillTo_Name, 1, 41)
                    strFullName = strName

                    '        'strIsActive = "1"               'strInactivePurge
                    '        'strIsActive   'strInactivePurge   'REVERSED
                    '        If strInactivePurge = "1" Then strIsActive = "0"
                    '        If strInactivePurge = "0" Then strIsActive = "1"
                    strIsActive = "1"

                    'not use
                    strParentRefListID = ""
                    strParentRefFullName = ""
                    strSublevel = "0"

                    strCompanyName = Strings.Mid(str1MaxBillTo_Name, 1, 41) 'use strLongName anywhere?  -no
                    strSalutation = ""

                    strFirstName = "" '"Cust=" & strCustomerN
                    strMiddleName = "" '"SAC=" & strSalesAreaCode
                    strLastName = "" '"GL=" & strSalesAcctN

                    strBillAddressAddr1 = Strings.Mid(str1MaxBillTo_Name, 1, 41)
                    strBillAddressAddr2 = Strings.Mid(str1MaxBillTo_Address_Line_1, 1, 41)
                    strBillAddressAddr3 = Strings.Mid(str1MaxBillTo_Address_Line_2, 1, 41)
                    strBillAddressAddr4 = ""
                    strBillAddressCity = Strings.Mid(str1MaxBillTo_City, 1, 31)
                    strBillAddressState = Strings.Mid(str1MaxBillTo_State_Province, 1, 21)
                    strBillAddressPostalCode = Strings.Mid(str1MaxBillTo_Zip_Code, 1, 13)
                    strBillAddressCountry = Strings.Mid(str1MaxBillTo_Country, 1, 31)
                    strShipAddressAddr1 = ""
                    strShipAddressAddr2 = ""
                    strShipAddressAddr3 = ""
                    strShipAddressAddr4 = ""
                    strShipAddressCity = ""
                    strShipAddressState = ""
                    strShipAddressPostalCode = ""
                    strShipAddressCountry = ""

                    strPhone = Strings.Mid(str1MaxBillTo_Phone_1, 1, 21).Trim() '& " x" & s
                    strAltPhone = Strings.Mid(str1MaxBillTo_Phone_3, 1, 21).Trim()
                    strFax = Strings.Mid(str1MaxBillTo_Phone_2, 1, 21).Trim()

                    strEmail = SQLHelper.ExecuteScalerString(cnDBPM, CommandType.Text, "SELECT AlphaNumericCol FROM AMGR_User_Fields_Tbl where Type_ID = 58850 AND Client_Id = '" & str1MaxBillTo_Client_Id & "' ORDER BY Create_Date")

                    strContact = "" 'strARContact
                    strAltContact = "" '"Rep=" & strSalespersonN

                    'TODO CHECK CustomerTypeRefListID Redundants
                    strCustomerTypeRefListID = "10000-1135698580"
                    strTermsRefListID = "100000-1134763841"
                    strTermsRefFullName = "Prepay"

                    strSalesRepRefListID = "50000-1135731044"
                    strSalesRepRefFullName = "Open"

                    strState = Strings.Mid(str1MaxBillTo_State_Province, 1, 21)
                    If strState = "IA" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "IL" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "IN" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "KS" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "MI" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "MN" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "MO" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "ND" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "NE" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "OH" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "SD" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "WI" Then
                        strSalesRepRefListID = "10000-1135700185"
                        strSalesRepRefFullName = "MW"
                    End If
                    If strState = "CT" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "DC" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "DE" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "MA" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "MD" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "ME" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "NH" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "NJ" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "NY" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "PA" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "RI" Then
                        strSalesRepRefListID = "20000-1135700258"
                        strSalesRepRefFullName = "NE"
                    End If
                    If strState = "AL" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If
                    If strState = "FL" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If
                    If strState = "GA" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If
                    If strState = "KY" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If
                    If strState = "MS" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If
                    If strState = "NC" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If
                    If strState = "SC" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If
                    If strState = "TN" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If
                    If strState = "VA" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If
                    If strState = "PR" Then
                        strSalesRepRefListID = "30000-1135700292"
                        strSalesRepRefFullName = "SE"
                    End If

                    'not mess with these
                    strBalance = "0" 'no
                    strTotalBalance = "0" 'no
                    strOpenBalance = "0"

                    strOpenBalanceDate = DateTime.Now.ToString("yyyy-MM-dd") '"1999-12-12"   'strCustStartDate?


                    strSalesTaxCodeRefListID = "20000-1134607388"
                    strSalesTaxCodeRefFullName = "Non"
                    strItemSalesTaxRefListID = "20000-1134760068"
                    strItemSalesTaxRefFullName = "0"

                    strResaleNumber = "" 'strTaxExemptCert
                    strAccountNumber = Strings.Mid(str1MaxBillTo_Phone_4, 1, 21)

                    strCreditLimitQ = "0"

                    strPreferredPaymentMethodRefListID = ""
                    strPreferredPaymentMethodRefFullName = ""

                    strCreditCardInfoCreditCardNumber = ""
                    strCreditCardInfoExpirationMonth = ""
                    strCreditCardInfoExpirationYear = ""
                    strCreditCardInfoNameOnCard = ""
                    strCreditCardInfoCreditCardAddress = ""
                    strCreditCardInfoCreditCardPostalCode = ""

                    strJobStatus = "None"
                    strJobStartDate = DateTime.Now.ToString("yyyy-MM-dd") 'used Cust Start Date
                    strJobProjectedEndDate = Date.Today.AddYears(3).ToString("yyyy-MM-dd")
                    strJobEndDate = Date.Today.AddYears(3).ToString("yyyy-MM-dd")

                    strJobDesc = str1MaxBillTo_Client_Id '"GL=" & strSalesAcctN & "  Rep=" & strSalespersonN & "  Cust=" & strCustomerN

                    strJobTypeRefListID = "10000-1135734033"
                    strJobTypeRefFullName = "ASI"

                    strNotes = "Added Via Maximizer.  Needs setup by Accounting." 'strCommentN1 & vbCrLf & strCommentN2   'PUT ANY MAX INFO INTO NOTES?

                    'Ask Amberlea?  no use now.
                    strPriceLevelRefListID = ""
                    strPriceLevelRefFullName = ""

                    'no use
                    strCustomFieldOther = ""
                    strFQSaveToCache = "1" '1=save now   0=cache

                    strSQL1 = "INSERT INTO Customer " & Environment.NewLine & _
                              "   ( Name " & Environment.NewLine & _
                              "   , IsActive " & Environment.NewLine & _
                              "   , CompanyName " & Environment.NewLine & _
                              "   , Salutation " & Environment.NewLine & _
                              "   , FirstName " & Environment.NewLine & _
                              "   , MiddleName " & Environment.NewLine & _
                              "   , LastName " & Environment.NewLine & _
                              "   , BillAddressAddr1 " & Environment.NewLine & _
                              "   , BillAddressAddr2 " & Environment.NewLine & _
                              "   , BillAddressAddr3 " & Environment.NewLine & _
                              "   , BillAddressAddr4 " & Environment.NewLine & _
                              "   , BillAddressCity " & Environment.NewLine & _
                              "   , BillAddressState " & Environment.NewLine & _
                              "   , BillAddressPostalCode " & Environment.NewLine & _
                              "   , BillAddressCountry " & Environment.NewLine & _
                              "   , ShipAddressAddr1 " & Environment.NewLine
                    strSQL2 = "   , ShipAddressAddr2 " & Environment.NewLine & _
                              "   , ShipAddressAddr3 " & Environment.NewLine & _
                              "   , ShipAddressAddr4 " & Environment.NewLine & _
                              "   , ShipAddressCity " & Environment.NewLine & _
                              "   , ShipAddressState " & Environment.NewLine & _
                              "   , ShipAddressPostalCode " & Environment.NewLine & _
                              "   , ShipAddressCountry " & Environment.NewLine & _
                              "   , Phone " & Environment.NewLine & _
                              "   , AltPhone " & Environment.NewLine & _
                              "   , Fax " & Environment.NewLine & _
                              "   , Email " & Environment.NewLine & _
                              "   , Contact " & Environment.NewLine & _
                              "   , AltContact " & Environment.NewLine & _
                              "   , OpenBalance " & Environment.NewLine & _
                              "   , OpenBalanceDate " & Environment.NewLine
                    strSQL3 = "   , ResaleNumber " & Environment.NewLine & _
                              "   , AccountNumber " & Environment.NewLine & _
                              "   , CreditLimit " & Environment.NewLine & _
                              "   , JobStatus " & Environment.NewLine & _
                              "   , JobStartDate " & Environment.NewLine & _
                              "   , JobProjectedEndDate " & Environment.NewLine & _
                              "   , JobEndDate " & Environment.NewLine & _
                              "   , JobDesc " & Environment.NewLine & _
                              "   , Notes )" & Environment.NewLine
                    strSQL4 = "VALUES " & Environment.NewLine & _
                              "   ( '" & strName & "'  " & Environment.NewLine & _
                              "   , " & strIsActive & "  " & Environment.NewLine & _
                              "   , '" & strCompanyName & "'  " & Environment.NewLine & _
                              "   , '" & strSalutation & "'  " & Environment.NewLine & _
                              "   , '" & strFirstName & "'  " & Environment.NewLine & _
                              "   , '" & strMiddleName & "'  " & Environment.NewLine & _
                              "   , '" & strLastName & "'  " & Environment.NewLine & _
                              "   , '" & strBillAddressAddr1 & "'  " & Environment.NewLine & _
                              "   , '" & strBillAddressAddr2 & "'  " & Environment.NewLine & _
                              "   , '" & strBillAddressAddr3 & "'  " & Environment.NewLine & _
                              "   , '" & strBillAddressAddr4 & "'  " & Environment.NewLine & _
                              "   , '" & strBillAddressCity & "'  " & Environment.NewLine & _
                              "   , '" & strBillAddressState & "'  " & Environment.NewLine & _
                              "   , '" & strBillAddressPostalCode & "'  " & Environment.NewLine & _
                              "   , '" & strBillAddressCountry & "'  " & Environment.NewLine & _
                              "   , '" & strShipAddressAddr1 & "'  " & Environment.NewLine
                    strSQL5 = "   , '" & strShipAddressAddr2 & "'  " & Environment.NewLine & _
                              "   , '" & strShipAddressAddr3 & "'  " & Environment.NewLine & _
                              "   , '" & strShipAddressAddr4 & "'  " & Environment.NewLine & _
                              "   , '" & strShipAddressCity & "'  " & Environment.NewLine & _
                              "   , '" & strShipAddressState & "'  " & Environment.NewLine & _
                              "   , '" & strShipAddressPostalCode & "'  " & Environment.NewLine & _
                              "   , '" & strShipAddressCountry & "'  " & Environment.NewLine & _
                              "   , '" & strPhone & "'  " & Environment.NewLine & _
                              "   , '" & strAltPhone & "'  " & Environment.NewLine & _
                              "   , '" & strFax & "'  " & Environment.NewLine & _
                              "   , '" & strEmail & "'  " & Environment.NewLine & _
                              "   , '" & strContact & "'  " & Environment.NewLine & _
                              "   , '" & strAltContact & "'  " & Environment.NewLine & _
                              "   , " & strOpenBalance & "  " & Environment.NewLine & _
                              "   , {d'" & strOpenBalanceDate & "'}  " & Environment.NewLine
                    strSQL6 = "   , '" & strResaleNumber & "'  " & Environment.NewLine & _
                              "   , '" & strAccountNumber & "'  " & Environment.NewLine & _
                              "   , " & strCreditLimitQ & "  " & Environment.NewLine & _
                              "   , '" & strJobStatus & "'  " & Environment.NewLine & _
                              "   , {d'" & strJobStartDate & "'}  " & Environment.NewLine & _
                              "   , {d'" & strJobProjectedEndDate & "'}  " & Environment.NewLine & _
                              "   , {d'" & strJobEndDate & "'}  " & Environment.NewLine & _
                              "   , '" & strJobDesc & "'  " & Environment.NewLine & _
                              "   , '" & strNotes & "' ) " & Environment.NewLine

                    'Combine the strings
                    strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6


                    booErrorAddingQBCust = False
                    strErrorLine = "cnQuickBooks.Execute strTableInsert"
                    strErrorLine = "ERROR ADDING:  " & strCompanyName & "  To QB  "


                    'FOR PART 2SrcQB_ - Get records from QB_Customer
                    Debug.WriteLine("List2SrcQB_QB_Customer")


                    Try
                        Dim TempCommand As ODBCCommand
                        TempCommand = cnQuickBooks.CreateCommand()
                        TempCommand.CommandText = strTableInsert
                        TempCommand.ExecuteNonQuery() 'FIXTHIS Put error protection & resume next
                        strErrorLine = ""

                       

                        iRecordsProcessed += 1

                    Catch ex As Exception
                        ShowUserMessage(strSubName, "Tried to insert a company into Quickbooks, but failed - " & strCompanyName, "", True)
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For
                    Finally
                        'Refresh the QB Cust table so can get ListID of record that has current Max Client_Id
                        RefreshQB_Customer()

                        Dim str2SrcQB_ListID As String
                        strSQL1 = "SELECT Top 1 ListID FROM QB_Customer WHERE JObDesc = '" & str1MaxBillTo_Client_Id & "'"
                        str2SrcQB_ListID = SQLHelper.ExecuteScalerString(cnDBPM, CommandType.Text, strSQL1)

                        strSQL1 = "UPDATE  AMGR_Client SET Firm = '" & str2SrcQB_ListID & "' WHERE Client_Id = '" & str1MaxBillTo_Client_Id & "' AND  Contact_Number = 0"
                        SQLHelper.ExecuteSQL(cnMax, strSQL1)
                    End Try
                Next iteration_row

            End If
            ShowUserMessage(strSubName, iRecordsProcessed.ToString & " Records Processed from MAX to QB", , True)
        End Using
     
    End Sub
End Module