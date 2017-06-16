Option Strict Off
Option Explicit On
Imports DBPM_Server.siteConstants
Imports DBPM_Server.SQLHelper
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Windows.Forms


'**********************************
'*** FIRST CODE REVIEW COMPLETE ***
'**********************************


Module modDBPM_RefreshCustTables

    Public gstrQBMaxTimeModified_Customer As String = ""


    Public Sub RefreshQBCustomerTables()

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshCustTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQBCustomerTables" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        If frmMain.chkPauseProcessing.CheckState = CheckState.Checked Then Exit Sub

        'Set flag
        booQBRefreshInProgress = True

        ShowUserMessage(strSubName, "Starting RefreshCustomerTables", "Starting RefreshCustomerTables", True)

        RefreshQB_Customer()

        Using objSQL As New SQLHelper()
            objSQL.ExecuteSP("exec sp_TEMP_MarkCustPromoRush")
        End Using

        ShowUserMessage(strSubName, "Finished RefreshCustomerTables", "Finished RefreshCustomerTables", True)
        booQBRefreshInProgress = False

    End Sub


    Public Sub RefreshQB_Customer()
        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshCustTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQB_Customer" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub


        ''FOR PART 1MaxOfCopy_ - Get records from QB_Customer
        'Debug.Print "List1MaxOfCopy_QB_Customer"
        'Dim rs1MaxOfCopy_QB_Customer As ADODB.Recordset
        'Dim str1MaxOfCopy_QB_CustomerSQL As String
        'Dim str1MaxOfCopy_QB_CustomerRow As String
        'Dim str1MaxOfCopy_TimeModified As String
        ''This routine gets the 1MaxOfCopy_QB_Customer from the database according to the selection in str1MaxOfCopy_QB_CustomerSQL.
        ''It then puts those 1MaxOfCopy_QB_Customer in the list box

        'FOR PART 1MaxOfCopy_ - Get records from QB_Customer
        Debug.WriteLine("List1MaxOfCopy_QB_Customer")
        'This routine gets the 1MaxOfCopy_QB_Customer from the database according to the selection in str1MaxOfCopy_QB_CustomerSQL.
        'It then puts those 1MaxOfCopy_QB_Customer in the list box

        'FOR PART 2SrcQB_ - Get records from QB_Customer
        Debug.WriteLine("List2SrcQB_QB_Customer")

        Dim str2SrcQB_QB_CustomerSQL, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_FullName, str2SrcQB_IsActive, str2SrcQB_ParentRefListID, str2SrcQB_ParentRefFullName, str2SrcQB_Sublevel, str2SrcQB_CompanyName, str2SrcQB_Salutation, str2SrcQB_FirstName, str2SrcQB_MiddleName, str2SrcQB_LastName, str2SrcQB_BillAddressAddr1, str2SrcQB_BillAddressAddr2, str2SrcQB_BillAddressAddr3, str2SrcQB_BillAddressAddr4, str2SrcQB_BillAddressCity, str2SrcQB_BillAddressState, str2SrcQB_BillAddressPostalCode, str2SrcQB_BillAddressCountry, str2SrcQB_ShipAddressAddr1, str2SrcQB_ShipAddressAddr2, str2SrcQB_ShipAddressAddr3, str2SrcQB_ShipAddressAddr4, str2SrcQB_ShipAddressCity, str2SrcQB_ShipAddressState, str2SrcQB_ShipAddressPostalCode, str2SrcQB_ShipAddressCountry, str2SrcQB_Phone, str2SrcQB_AltPhone, str2SrcQB_Fax, str2SrcQB_Email, str2SrcQB_Contact, str2SrcQB_AltContact, str2SrcQB_CustomerTypeRefListID, str2SrcQB_CustomerTypeRefFullName, str2SrcQB_TermsRefListID, str2SrcQB_TermsRefFullName, str2SrcQB_SalesRepRefListID, str2SrcQB_SalesRepRefFullName, str2SrcQB_OpenBalanceDate, str2SrcQB_SalesTaxCodeRefListID, str2SrcQB_SalesTaxCodeRefFullName, str2SrcQB_ItemSalesTaxRefListID, str2SrcQB_ItemSalesTaxRefFullName, str2SrcQB_ResaleNumber, str2SrcQB_AccountNumber, str2SrcQB_CreditLimit, str2SrcQB_PreferredPaymentMethodRefListID, str2SrcQB_PreferredPaymentMethodRefFullName, str2SrcQB_CreditCardInfoCreditCardNumber, str2SrcQB_CreditCardInfoExpirationMonth, str2SrcQB_CreditCardInfoExpirationYear, str2SrcQB_CreditCardInfoNameOnCard, str2SrcQB_CreditCardInfoCreditCardAddress, str2SrcQB_CreditCardInfoCreditCardPostalCode, str2SrcQB_JobStatus, str2SrcQB_JobStartDate, str2SrcQB_JobProjectedEndDate, str2SrcQB_JobEndDate, str2SrcQB_JobDesc, str2SrcQB_JobTypeRefListID, str2SrcQB_JobTypeRefFullName, str2SrcQB_Notes, str2SrcQB_PriceLevelRefListID, str2SrcQB_PriceLevelRefFullName As String
        Dim str2SrcQB_CustomFieldOther As String = ""

        Dim str2SrcQB_Balance, str2SrcQB_TotalBalance, str2SrcQB_OpenBalance As Double
        'This routine gets the 2SrcQB_QB_Customer from the database according to the selection in str2SrcQB_QB_CustomerSQL.
        'It then puts those 2SrcQB_QB_Customer in the list box

        'FOR PART 3TestID_ - Get records from QB_Customer
        Debug.WriteLine("List3TestID_QB_Customer")
        'This routine gets the 3TestID_QB_Customer from the database according to the selection in str3TestID_QB_CustomerSQL.
        'It then puts those 3TestID_QB_Customer in the list box

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'Show what's processing
        ShowUserMessage(strSubName, "RefreshQB: Processing QB_Customer Records", "RefreshQB: Processing QB_Customer Records", True)

        If Not (cnQuickBooks.State = ConnectionState.Open) Then
            OpenConnectionQB()
        End If

     
        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QB_Customer
        str2SrcQB_QB_CustomerSQL = "SELECT * FROM Customer WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_Customer & "'}" ' ORDER BY TimeModified"
        Dim qbCommand As New OdbcCommand(str2SrcQB_QB_CustomerSQL, cnQuickBooks)

        Using rs2SrcQB_QB_Customer As OdbcDataReader = qbCommand.ExecuteReader()
            Dim iRowCount As Integer = 0
            If rs2SrcQB_QB_Customer.HasRows Then
                'Show what's processing in the listbox
                ShowUserMessage(strSubName, "Processing QB_Customer Records", "Processing QB_Customer Records")

                While rs2SrcQB_QB_Customer.Read
                    iRowCount += 1
                    ShowUserMessage(strSubName, "Processing QB_Customer Record # " & iRowCount.ToString)

                    Try



                        'get the columns from the database
                        str2SrcQB_ListID = NCStr(rs2SrcQB_QB_Customer("ListID")).Replace("'"c, "`"c)
                        str2SrcQB_TimeCreated = NCStr(rs2SrcQB_QB_Customer("TimeCreated")).Replace("'"c, "`"c)
                        str2SrcQB_TimeModified = NCStr(rs2SrcQB_QB_Customer("TimeModified")).Replace("'"c, "`"c)
                        str2SrcQB_EditSequence = NCStr(rs2SrcQB_QB_Customer("EditSequence")).Replace("'"c, "`"c)
                        str2SrcQB_Name = NCStr(rs2SrcQB_QB_Customer("Name")).Replace("'"c, "`"c)
                        str2SrcQB_FullName = NCStr(rs2SrcQB_QB_Customer("FullName")).Replace("'"c, "`"c)
                        str2SrcQB_IsActive = NCStr(rs2SrcQB_QB_Customer("IsActive")).Replace("'"c, "`"c)
                        str2SrcQB_ParentRefListID = NCStr(rs2SrcQB_QB_Customer("ParentRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ParentRefFullName = NCStr(rs2SrcQB_QB_Customer("ParentRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_Sublevel = NCStr(rs2SrcQB_QB_Customer("Sublevel")).Replace("'"c, "`"c)
                        str2SrcQB_CompanyName = NCStr(rs2SrcQB_QB_Customer("CompanyName")).Replace("'"c, "`"c)
                        str2SrcQB_Salutation = NCStr(rs2SrcQB_QB_Customer("Salutation")).Replace("'"c, "`"c)
                        str2SrcQB_FirstName = NCStr(rs2SrcQB_QB_Customer("FirstName")).Replace("'"c, "`"c)
                        str2SrcQB_MiddleName = NCStr(rs2SrcQB_QB_Customer("MiddleName")).Replace("'"c, "`"c)
                        str2SrcQB_LastName = NCStr(rs2SrcQB_QB_Customer("LastName")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr1 = NCStr(rs2SrcQB_QB_Customer("BillAddressAddr1")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr2 = NCStr(rs2SrcQB_QB_Customer("BillAddressAddr2")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr3 = NCStr(rs2SrcQB_QB_Customer("BillAddressAddr3")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr4 = NCStr(rs2SrcQB_QB_Customer("BillAddressAddr4")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressCity = NCStr(rs2SrcQB_QB_Customer("BillAddressCity")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressState = NCStr(rs2SrcQB_QB_Customer("BillAddressState")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressPostalCode = NCStr(rs2SrcQB_QB_Customer("BillAddressPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressCountry = NCStr(rs2SrcQB_QB_Customer("BillAddressCountry")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr1 = NCStr(rs2SrcQB_QB_Customer("ShipAddressAddr1")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr2 = NCStr(rs2SrcQB_QB_Customer("ShipAddressAddr2")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr3 = NCStr(rs2SrcQB_QB_Customer("ShipAddressAddr3")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr4 = NCStr(rs2SrcQB_QB_Customer("ShipAddressAddr4")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressCity = NCStr(rs2SrcQB_QB_Customer("ShipAddressCity")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressState = NCStr(rs2SrcQB_QB_Customer("ShipAddressState")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressPostalCode = NCStr(rs2SrcQB_QB_Customer("ShipAddressPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressCountry = NCStr(rs2SrcQB_QB_Customer("ShipAddressCountry")).Replace("'"c, "`"c)
                        str2SrcQB_Phone = NCStr(rs2SrcQB_QB_Customer("Phone")).Replace("'"c, "`"c)
                        str2SrcQB_AltPhone = NCStr(rs2SrcQB_QB_Customer("AltPhone")).Replace("'"c, "`"c)
                        str2SrcQB_Fax = NCStr(rs2SrcQB_QB_Customer("Fax")).Replace("'"c, "`"c)
                        str2SrcQB_Email = NCStr(rs2SrcQB_QB_Customer("Email")).Replace("'"c, "`"c)
                        str2SrcQB_Contact = NCStr(rs2SrcQB_QB_Customer("Contact")).Replace("'"c, "`"c)
                        str2SrcQB_AltContact = NCStr(rs2SrcQB_QB_Customer("AltContact")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerTypeRefListID = NCStr(rs2SrcQB_QB_Customer("CustomerTypeRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerTypeRefFullName = NCStr(rs2SrcQB_QB_Customer("CustomerTypeRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_TermsRefListID = NCStr(rs2SrcQB_QB_Customer("TermsRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_TermsRefFullName = NCStr(rs2SrcQB_QB_Customer("TermsRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_SalesRepRefListID = NCStr(rs2SrcQB_QB_Customer("SalesRepRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_SalesRepRefFullName = NCStr(rs2SrcQB_QB_Customer("SalesRepRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_Balance = NCDbl(rs2SrcQB_QB_Customer("Balance"))
                        str2SrcQB_TotalBalance = NCDbl(rs2SrcQB_QB_Customer("TotalBalance"))
                        str2SrcQB_OpenBalance = NCDbl(rs2SrcQB_QB_Customer("OpenBalance"))
                        str2SrcQB_OpenBalanceDate = NCStr(rs2SrcQB_QB_Customer("OpenBalanceDate")).Replace("'"c, "`"c)
                        str2SrcQB_SalesTaxCodeRefListID = NCStr(rs2SrcQB_QB_Customer("SalesTaxCodeRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_SalesTaxCodeRefFullName = NCStr(rs2SrcQB_QB_Customer("SalesTaxCodeRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ItemSalesTaxRefListID = NCStr(rs2SrcQB_QB_Customer("ItemSalesTaxRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ItemSalesTaxRefFullName = NCStr(rs2SrcQB_QB_Customer("ItemSalesTaxRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ResaleNumber = NCStr(rs2SrcQB_QB_Customer("ResaleNumber")).Replace("'"c, "`"c)
                        str2SrcQB_AccountNumber = NCStr(rs2SrcQB_QB_Customer("AccountNumber")).Replace("'"c, "`"c)
                        str2SrcQB_CreditLimit = NCStr(rs2SrcQB_QB_Customer("CreditLimit")).Replace("'"c, "`"c)
                        str2SrcQB_PreferredPaymentMethodRefListID = NCStr(rs2SrcQB_QB_Customer("PreferredPaymentMethodRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_PreferredPaymentMethodRefFullName = NCStr(rs2SrcQB_QB_Customer("PreferredPaymentMethodRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardInfoCreditCardNumber = NCStr(rs2SrcQB_QB_Customer("CreditCardInfoCreditCardNumber")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardInfoExpirationMonth = NCStr(rs2SrcQB_QB_Customer("CreditCardInfoExpirationMonth")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardInfoExpirationYear = NCStr(rs2SrcQB_QB_Customer("CreditCardInfoExpirationYear")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardInfoNameOnCard = NCStr(rs2SrcQB_QB_Customer("CreditCardInfoNameOnCard")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardInfoCreditCardAddress = NCStr(rs2SrcQB_QB_Customer("CreditCardInfoCreditCardAddress")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardInfoCreditCardPostalCode = NCStr(rs2SrcQB_QB_Customer("CreditCardInfoCreditCardPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_JobStatus = NCStr(rs2SrcQB_QB_Customer("JobStatus")).Replace("'"c, "`"c)
                        str2SrcQB_JobStartDate = NCStr(rs2SrcQB_QB_Customer("JobStartDate")).Replace("'"c, "`"c)
                        str2SrcQB_JobProjectedEndDate = NCStr(rs2SrcQB_QB_Customer("JobProjectedEndDate")).Replace("'"c, "`"c)
                        str2SrcQB_JobEndDate = NCStr(rs2SrcQB_QB_Customer("JobEndDate")).Replace("'"c, "`"c)
                        str2SrcQB_JobDesc = NCStr(rs2SrcQB_QB_Customer("JobDesc")).Replace("'"c, "`"c)
                        str2SrcQB_JobTypeRefListID = NCStr(rs2SrcQB_QB_Customer("JobTypeRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_JobTypeRefFullName = NCStr(rs2SrcQB_QB_Customer("JobTypeRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_Notes = NCStr(rs2SrcQB_QB_Customer("Notes")).Replace("'"c, "`"c)
                        str2SrcQB_PriceLevelRefListID = NCStr(rs2SrcQB_QB_Customer("PriceLevelRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_PriceLevelRefFullName = NCStr(rs2SrcQB_QB_Customer("PriceLevelRefFullName")).Replace("'"c, "`"c)
                        '        If rs2SrcQB_QB_Customer!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_Customer!CustomFieldOther

                        'Change flags back to binary
                        str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")

                        Dim recordCount As Integer = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(ListID) FROM QB_Customer WHERE ListID = '" & str2SrcQB_ListID & "'")

                        If recordCount = 1 Then 'record exists  -UPDATE
                            'DO UPDATE WORK:
                            Debug.WriteLine("UPDATE")

                            'Build the SQL string
                            strSQL1 = "UPDATE  " & Environment.NewLine & _
                                      "       QB_Customer " & Environment.NewLine & _
                                      "SET " & Environment.NewLine & _
                                      "       -- ListID = '" & str2SrcQB_ListID & "'" & Environment.NewLine & _
                                      "       TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                      "     , Name = '" & str2SrcQB_Name & "'" & Environment.NewLine & _
                                      "     , FullName = '" & str2SrcQB_FullName & "'" & Environment.NewLine & _
                                      "     , IsActive = '" & str2SrcQB_IsActive & "'" & Environment.NewLine & _
                                      "     , ParentRefListID = '" & str2SrcQB_ParentRefListID & "'" & Environment.NewLine & _
                                      "     , ParentRefFullName = '" & str2SrcQB_ParentRefFullName & "'" & Environment.NewLine & _
                                      "     , Sublevel = '" & str2SrcQB_Sublevel & "'" & Environment.NewLine & _
                                      "     , CompanyName = '" & str2SrcQB_CompanyName & "'" & Environment.NewLine & _
                                      "     , Salutation = '" & str2SrcQB_Salutation & "'" & Environment.NewLine & _
                                      "     , FirstName = '" & str2SrcQB_FirstName & "'" & Environment.NewLine & _
                                      "     , MiddleName = '" & str2SrcQB_MiddleName & "'" & Environment.NewLine & _
                                      "     , LastName = '" & str2SrcQB_LastName & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr1 = '" & str2SrcQB_BillAddressAddr1 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr2 = '" & str2SrcQB_BillAddressAddr2 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr3 = '" & str2SrcQB_BillAddressAddr3 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr4 = '" & str2SrcQB_BillAddressAddr4 & "'" & Environment.NewLine & _
                                      "     , BillAddressCity = '" & str2SrcQB_BillAddressCity & "'" & Environment.NewLine & _
                                      "     , BillAddressState = '" & str2SrcQB_BillAddressState & "'" & Environment.NewLine & _
                                      "     , BillAddressPostalCode = '" & str2SrcQB_BillAddressPostalCode & "'" & Environment.NewLine
                            strSQL2 = "     , BillAddressCountry = '" & str2SrcQB_BillAddressCountry & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr1 = '" & str2SrcQB_ShipAddressAddr1 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr2 = '" & str2SrcQB_ShipAddressAddr2 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr3 = '" & str2SrcQB_ShipAddressAddr3 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr4 = '" & str2SrcQB_ShipAddressAddr4 & "'" & Environment.NewLine & _
                                      "     , ShipAddressCity = '" & str2SrcQB_ShipAddressCity & "'" & Environment.NewLine & _
                                      "     , ShipAddressState = '" & str2SrcQB_ShipAddressState & "'" & Environment.NewLine & _
                                      "     , ShipAddressPostalCode = '" & str2SrcQB_ShipAddressPostalCode & "'" & Environment.NewLine & _
                                      "     , ShipAddressCountry = '" & str2SrcQB_ShipAddressCountry & "'" & Environment.NewLine & _
                                      "     , Phone = '" & str2SrcQB_Phone & "'" & Environment.NewLine & _
                                      "     , AltPhone = '" & str2SrcQB_AltPhone & "'" & Environment.NewLine & _
                                      "     , Fax = '" & str2SrcQB_Fax & "'" & Environment.NewLine & _
                                      "     , Email = '" & str2SrcQB_Email & "'" & Environment.NewLine & _
                                      "     , Contact = '" & str2SrcQB_Contact & "'" & Environment.NewLine & _
                                      "     , AltContact = '" & str2SrcQB_AltContact & "'" & Environment.NewLine & _
                                      "     , CustomerTypeRefListID = '" & str2SrcQB_CustomerTypeRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerTypeRefFullName = '" & str2SrcQB_CustomerTypeRefFullName & "'" & Environment.NewLine & _
                                      "     , TermsRefListID = '" & str2SrcQB_TermsRefListID & "'" & Environment.NewLine & _
                                      "     , TermsRefFullName = '" & str2SrcQB_TermsRefFullName & "'" & Environment.NewLine & _
                                      "     , SalesRepRefListID = '" & str2SrcQB_SalesRepRefListID & "'" & Environment.NewLine & _
                                      "     , SalesRepRefFullName = '" & str2SrcQB_SalesRepRefFullName & "'" & Environment.NewLine & _
                                      "     , Balance = '" & str2SrcQB_Balance & "'" & Environment.NewLine & _
                                      "     , TotalBalance = '" & str2SrcQB_TotalBalance & "'" & Environment.NewLine & _
                                      "     , OpenBalance = '" & str2SrcQB_OpenBalance & "'" & Environment.NewLine & _
                                      "     , OpenBalanceDate = '" & str2SrcQB_OpenBalanceDate & "'" & Environment.NewLine
                            strSQL3 = "     , SalesTaxCodeRefListID = '" & str2SrcQB_SalesTaxCodeRefListID & "'" & Environment.NewLine & _
                                      "     , SalesTaxCodeRefFullName = '" & str2SrcQB_SalesTaxCodeRefFullName & "'" & Environment.NewLine & _
                                      "     , ItemSalesTaxRefListID = '" & str2SrcQB_ItemSalesTaxRefListID & "'" & Environment.NewLine & _
                                      "     , ItemSalesTaxRefFullName = '" & str2SrcQB_ItemSalesTaxRefFullName & "'" & Environment.NewLine & _
                                      "     , ResaleNumber = '" & str2SrcQB_ResaleNumber & "'" & Environment.NewLine & _
                                      "     , AccountNumber = '" & str2SrcQB_AccountNumber & "'" & Environment.NewLine & _
                                      "     , CreditLimit = '" & str2SrcQB_CreditLimit & "'" & Environment.NewLine & _
                                      "     , PreferredPaymentMethodRefListID = '" & str2SrcQB_PreferredPaymentMethodRefListID & "'" & Environment.NewLine & _
                                      "     , PreferredPaymentMethodRefFullName = '" & str2SrcQB_PreferredPaymentMethodRefFullName & "'" & Environment.NewLine & _
                                      "     , CreditCardInfoCreditCardNumber = '" & str2SrcQB_CreditCardInfoCreditCardNumber & "'" & Environment.NewLine & _
                                      "     , CreditCardInfoExpirationMonth = '" & str2SrcQB_CreditCardInfoExpirationMonth & "'" & Environment.NewLine & _
                                      "     , CreditCardInfoExpirationYear = '" & str2SrcQB_CreditCardInfoExpirationYear & "'" & Environment.NewLine & _
                                      "     , CreditCardInfoNameOnCard = '" & str2SrcQB_CreditCardInfoNameOnCard & "'" & Environment.NewLine & _
                                      "     , CreditCardInfoCreditCardAddress = '" & str2SrcQB_CreditCardInfoCreditCardAddress & "'" & Environment.NewLine & _
                                      "     , CreditCardInfoCreditCardPostalCode = '" & str2SrcQB_CreditCardInfoCreditCardPostalCode & "'" & Environment.NewLine & _
                                      "     , JobStatus = '" & str2SrcQB_JobStatus & "'" & Environment.NewLine & _
                                      "     , JobStartDate = '" & str2SrcQB_JobStartDate & "'" & Environment.NewLine & _
                                      "     , JobProjectedEndDate = '" & str2SrcQB_JobProjectedEndDate & "'" & Environment.NewLine & _
                                      "     , JobEndDate = '" & str2SrcQB_JobEndDate & "'" & Environment.NewLine & _
                                      "     , JobDesc = '" & str2SrcQB_JobDesc & "'" & Environment.NewLine & _
                                      "     , JobTypeRefListID = '" & str2SrcQB_JobTypeRefListID & "'" & Environment.NewLine & _
                                      "     , JobTypeRefFullName = '" & str2SrcQB_JobTypeRefFullName & "'" & Environment.NewLine & _
                                      "     , Notes = '" & str2SrcQB_Notes & "'" & Environment.NewLine
                            strSQL4 = "     , PriceLevelRefListID = '" & str2SrcQB_PriceLevelRefListID & "'" & Environment.NewLine & _
                                      "     , PriceLevelRefFullName = '" & str2SrcQB_PriceLevelRefFullName & "'" & Environment.NewLine & _
                                      "     , CustomFieldOther = '" & str2SrcQB_CustomFieldOther & "'" & Environment.NewLine & _
                                      "WHERE " & Environment.NewLine & _
                                      "       ListID = '" & str2SrcQB_ListID & "'" & Environment.NewLine

                            'Combine the strings
                            strTableUpdate = strSQL1 & strSQL2 & strSQL3 & strSQL4 '& strSQL5 & strSQL6

                            'Execute the insert
                            Try
                                SQLHelper.ExecuteSQL(cnMax, strTableUpdate)
                            Catch ex As Exception
                                HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                                Continue While

                            End Try

                            strSQL1 = "UPDATE  " & Environment.NewLine & _
                                      "       AMGR_Client_Tbl " & Environment.NewLine & _
                                      "SET " & Environment.NewLine & _
                                      "       Name = '" & str2SrcQB_Name & "'" & Environment.NewLine & _
                                      "     , Phone_1 = '" & str2SrcQB_Phone & "'" & Environment.NewLine & _
                                      "     , Phone_2 = '" & str2SrcQB_Fax & "'" & Environment.NewLine
                            strSQL2 = "     , Phone_3 = '" & str2SrcQB_AltPhone & "'" & Environment.NewLine & _
                                      "     , Phone_4 = '" & str2SrcQB_AccountNumber & "'" & Environment.NewLine & _
                                      "     , Address_Line_1 = '" & str2SrcQB_BillAddressAddr2 & "'" & Environment.NewLine & _
                                      "     , Address_Line_2 = '" & str2SrcQB_BillAddressAddr3 & "'" & Environment.NewLine & _
                                      "     , City = '" & str2SrcQB_BillAddressCity & "'" & Environment.NewLine
                            strSQL3 = "     , State_Province = '" & str2SrcQB_BillAddressState & "'" & Environment.NewLine & _
                                      "     , Country = '" & str2SrcQB_BillAddressCountry & "'" & Environment.NewLine & _
                                      "     , Zip_Code = '" & str2SrcQB_BillAddressPostalCode & "'" & Environment.NewLine & _
                                      "     , Updated_By_Id = 'QUICKBOOKS'" & Environment.NewLine
                            strSQL4 = "WHERE " & Environment.NewLine & _
                                          "       Firm = '" & str2SrcQB_ListID & "'" & Environment.NewLine & _
                                          "  AND  Contact_Number = 0"

                            'Combine the strings
                            strTableUpdate = strSQL1 & strSQL2 & strSQL3 & strSQL4
                            ShowUserMessage(strSubName, "Updating " & str2SrcQB_CompanyName & " information in QB_Customer Table from Quickbooks.", , False)
                            'Execute the insert
                            Try
                                SQLHelper.ExecuteSQL(cnMax, strTableUpdate)
                            Catch ex As Exception
                                HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                                Continue While
                            End Try


                        Else
                            'record not exist  -INSERT
                            'DO INSERT WORK:
                            Debug.WriteLine("INSERT")

                            'Build the SQL string
                            strSQL1 = "INSERT INTO QB_Customer " & Environment.NewLine & _
                                      "   ( ListID " & Environment.NewLine & _
                                      ", TimeCreated " & Environment.NewLine & _
                                      ", TimeModified " & Environment.NewLine & _
                                      ", EditSequence " & Environment.NewLine & _
                                      ", Name " & Environment.NewLine & _
                                      ", FullName " & Environment.NewLine & _
                                      ", IsActive " & Environment.NewLine & _
                                      ", ParentRefListID " & Environment.NewLine & _
                                      ", ParentRefFullName " & Environment.NewLine & _
                                      ", Sublevel " & Environment.NewLine & _
                                      ", CompanyName " & Environment.NewLine & _
                                      ", Salutation " & Environment.NewLine & _
                                      ", FirstName " & Environment.NewLine & _
                                      ", MiddleName " & Environment.NewLine & _
                                      ", LastName " & Environment.NewLine & _
                                      ", BillAddressAddr1 " & Environment.NewLine & _
                                      ", BillAddressAddr2 " & Environment.NewLine & _
                                      ", BillAddressAddr3 " & Environment.NewLine & _
                                      ", BillAddressAddr4 " & Environment.NewLine & _
                                      ", BillAddressCity " & Environment.NewLine & _
                                      ", BillAddressState " & Environment.NewLine & _
                                      ", BillAddressPostalCode " & Environment.NewLine & _
                                      ", BillAddressCountry " & Environment.NewLine & _
                                      ", ShipAddressAddr1 " & Environment.NewLine
                            strSQL2 = ", ShipAddressAddr2 " & Environment.NewLine & _
                                      ", ShipAddressAddr3 " & Environment.NewLine & _
                                      ", ShipAddressAddr4 " & Environment.NewLine & _
                                      ", ShipAddressCity " & Environment.NewLine & _
                                      ", ShipAddressState " & Environment.NewLine & _
                                      ", ShipAddressPostalCode " & Environment.NewLine & _
                                      ", ShipAddressCountry " & Environment.NewLine & _
                                      ", Phone " & Environment.NewLine & _
                                      ", AltPhone " & Environment.NewLine & _
                                      ", Fax " & Environment.NewLine & _
                                      ", Email " & Environment.NewLine & _
                                      ", Contact " & Environment.NewLine & _
                                      ", AltContact " & Environment.NewLine & _
                                      ", CustomerTypeRefListID " & Environment.NewLine & _
                                      ", CustomerTypeRefFullName " & Environment.NewLine & _
                                      ", TermsRefListID " & Environment.NewLine & _
                                      ", TermsRefFullName " & Environment.NewLine & _
                                      ", SalesRepRefListID " & Environment.NewLine & _
                                      ", SalesRepRefFullName " & Environment.NewLine & _
                                      ", Balance " & Environment.NewLine & _
                                      ", TotalBalance " & Environment.NewLine & _
                                      ", OpenBalance " & Environment.NewLine & _
                                      ", OpenBalanceDate " & Environment.NewLine & _
                                      ", SalesTaxCodeRefListID " & Environment.NewLine & _
                                      ", SalesTaxCodeRefFullName " & Environment.NewLine
                            strSQL3 = ", ItemSalesTaxRefListID " & Environment.NewLine & _
                                      ", ItemSalesTaxRefFullName " & Environment.NewLine & _
                                      ", ResaleNumber " & Environment.NewLine & _
                                      ", AccountNumber " & Environment.NewLine & _
                                      ", CreditLimit " & Environment.NewLine & _
                                      ", PreferredPaymentMethodRefListID " & Environment.NewLine & _
                                      ", PreferredPaymentMethodRefFullName " & Environment.NewLine & _
                                      ", JobStatus " & Environment.NewLine & _
                                      ", JobStartDate " & Environment.NewLine & _
                                      ", JobProjectedEndDate " & Environment.NewLine & _
                                      ", JobEndDate " & Environment.NewLine & _
                                      ", JobDesc " & Environment.NewLine & _
                                      ", JobTypeRefListID " & Environment.NewLine & _
                                      ", JobTypeRefFullName " & Environment.NewLine & _
                                      ", Notes " & Environment.NewLine & _
                                      ", PriceLevelRefListID " & Environment.NewLine & _
                                      ", PriceLevelRefFullName " & Environment.NewLine & _
                                      ", CustomFieldOther )" & Environment.NewLine
                            strSQL4 = "VALUES " & Environment.NewLine & _
                                      "   ( '" & str2SrcQB_ListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_TimeCreated & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_TimeModified & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_EditSequence & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_Name & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_FullName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_IsActive & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ParentRefListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ParentRefFullName & "'  " & Environment.NewLine & _
                                      ", " & str2SrcQB_Sublevel & "  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_CompanyName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_Salutation & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_FirstName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_MiddleName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_LastName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_BillAddressAddr1 & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_BillAddressAddr2 & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_BillAddressAddr3 & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_BillAddressAddr4 & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_BillAddressCity & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_BillAddressState & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_BillAddressPostalCode & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_BillAddressCountry & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ShipAddressAddr1 & "'  " & Environment.NewLine
                            strSQL5 = ", '" & str2SrcQB_ShipAddressAddr2 & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ShipAddressAddr3 & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ShipAddressAddr4 & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ShipAddressCity & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ShipAddressState & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ShipAddressPostalCode & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ShipAddressCountry & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_Phone & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_AltPhone & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_Fax & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_Email & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_Contact & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_AltContact & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_CustomerTypeRefListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_CustomerTypeRefFullName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_TermsRefListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_TermsRefFullName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_SalesRepRefListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_SalesRepRefFullName & "'  " & Environment.NewLine & _
                                      ", " & str2SrcQB_Balance & "  " & Environment.NewLine & _
                                      ", " & str2SrcQB_TotalBalance & "  " & Environment.NewLine & _
                                      ", " & str2SrcQB_OpenBalance & "  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_OpenBalanceDate & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_SalesTaxCodeRefListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_SalesTaxCodeRefFullName & "'  " & Environment.NewLine
                            strSQL6 = ", '" & str2SrcQB_ItemSalesTaxRefListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ItemSalesTaxRefFullName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_ResaleNumber & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_AccountNumber & "'  " & Environment.NewLine & _
                                      ", " & str2SrcQB_CreditLimit & "  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_PreferredPaymentMethodRefListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_PreferredPaymentMethodRefFullName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_JobStatus & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_JobStartDate & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_JobProjectedEndDate & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_JobEndDate & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_JobDesc & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_JobTypeRefListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_JobTypeRefFullName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_Notes & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_PriceLevelRefListID & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_PriceLevelRefFullName & "'  " & Environment.NewLine & _
                                      ", '" & str2SrcQB_CustomFieldOther & "' ) " & Environment.NewLine


                            'Combine the strings
                            strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
                            ShowUserMessage(strSubName, "Inserting " & str2SrcQB_CompanyName & " into QB_Customer Table from Quickbooks.", , False)
                            'Execute the insert
                            Try
                                SQLHelper.ExecuteSQL(cnMax, strTableInsert)
                            Catch ex As Exception
                                HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                                Continue While
                            End Try


                        End If
                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue While
                    End Try
                End While

                ShowUserMessage(strSubName, "Finished processing " & iRowCount.ToString & " customer records from Quickbooks", , True)
            End If
        End Using
    End Sub

End Module