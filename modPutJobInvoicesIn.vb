Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Windows.Forms
Imports DBPM_Server.SQLHelper
Imports DBPM_Server.siteConstants

'**************************************
'******* 1st Code Review Complete *****
'**************************************



Module modPutJobInvoicesIn

    Public Sub PutJobInvoicesIn()
        Dim strSQL As String = ""

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modPutJobInvoicesIn" '"OBJNAME"
        Dim strSubName As String = "PutJobInvoicesIn" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'FOR PART Jobs_ - Get records from JOB_Header
        Debug.WriteLine("ListJobs_JOB_Header")
        Dim strJobs_JOB_HeaderSQL, strJobs_JOB_HeaderRow, strJobs_SalesOrderN, strJobs_CreatedDate, strJobs_CreatedBy, strJobs_CreatedOnComputer, strJobs_ModifiedDate, strJobs_ModifiedBy, strJobs_BillToCompanyName, strJobs_BillToCompanyKeyMaxClientID, strJobs_BillToCompanyKeyMaxContactNum, strJobs_BillToCompanyKeyQBListID, strJobs_BillToAddressName, strJobs_BillToAddressKeyMaxClientID, strJobs_BillToAddressKeyMaxContactNum, strJobs_BillToAddressKeyQBListID, strJobs_InvoiceTxnID, strJobs_CustPONum, strJobs_CustPODate, strJobs_CustPOAmt, strJobs_TypeRec, strJobs_JobNum, strJobs_QuoteNum, strJobs_TemplateNum, strJobs_Item, strJobs_ItemDesc, strJobs_ItemRefListID, strJobs_ItemRefFullName, strJobs_Imprint, strJobs_Qty, strJobs_DG, strJobs_ShipDate, strJobs_DateShipped, strJobs_AcctTerms, strJobs_AcctCreditLimit, strJobs_AcctShippingClearance, strJobs_AcctNeedBeforeShip, strJobs_CancelDate, strJobs_CancelBy, strJobs_CancelReason, strJobs_Color, strJobs_Comments, strJobs_Temp_Press, strJobs_Temp_ProofOK, strJobs_Temp_Plated, strJobs_Temp_DIDDE, strJobs_Temp_QM, strJobs_Temp_Foiled, strJobs_Temp_ShippedQty, strJobs_CSR, strJobs_OP, strJobs_TypeOrder, strJobs_PreviousJobs, strJobs_IsNewCustomer, strJobs_IsCompleted, strJobs_IsInvoiced, strJobs_IsCustom, strJobs_IsReprint, strJobs_IsPromo, strJobs_IsRush, strJobs_AcctFaxCreditApp, strJobs_ShipDateC, strJobs_ColorC, strJobs_CommentsC, strJobs_ProcessType, strJobs_ForQtyFrom, strJobs_ForQtyTo, strJobs_IsWillBeInvoiced, strJobs_IsShowOnProduction, strJobs_IsFiftyFreeBC1, strJobs_SourceCode, strJobs_IsNotForMktg, strJobs_IsRev_Item, strJobs_IsRev_Qty, strJobs_IsRev_Imprint, strJobs_IsRev_TypeOrder, strJobs_IsRev_PreviousJobs, strJobs_IsRev_ShipDate, strJobs_SetupCharge, strJobs_SetupQBItem, strJobs_RunCharge, strJobs_RunQBItem, strJobs_PriceDiscountCode, strJobs_PriceApprovedBy, strJobs_PriceComment As String

        strJobs_DateShipped = Date.Now.ToString("MM/dd/yyyy")

        'This routine gets the Jobs_JOB_Header from the database according to the selection in strJobs_JOB_HeaderSQL.
        'It then puts those Jobs_JOB_Header in the list box


        'FOR PART ShipInfo_ - Get records from JOB_Step
        Debug.WriteLine("ListShipInfo_JOB_Step")
        Dim strShipInfo_JOB_StepSQL, strShipInfo_StepCompletedDate, strShipInfo_ShipToDescription, strShipInfo_ShipToCompany, strShipInfo_ShipToPerson, strShipInfo_ShipToAddr1, strShipInfo_ShipToAddr2, strShipInfo_ShipToCity, strShipInfo_ShipToState, strShipInfo_ShipToZip, strShipInfo_ShipToPhone, strShipInfo_ShipToEmail, strShipInfo_ShipToQty, strShipInfo_ShipToCarrier, strShipInfo_ShipQtyShipped As String
        'This routine gets the ShipInfo_JOB_Step from the database according to the selection in strShipInfo_JOB_StepSQL.
        'It then puts those ShipInfo_JOB_Step in the list box



        'FOR PART Cust_ - Get records from QB_Customer
        Debug.WriteLine("ListCust_QB_Customer")
        'Dim strCust_QB_CustomerSQL, strCust_QB_CustomerRow, strCust_ListID, strCust_TimeCreated, strCust_TimeModified, strCust_EditSequence, strCust_Name, strCust_FullName, strCust_IsActive, strCust_ParentRefListID, strCust_ParentRefFullName, strCust_Sublevel, strCust_CompanyName, strCust_Salutation, strCust_FirstName, strCust_MiddleName, strCust_LastName, strCust_BillAddressAddr1, strCust_BillAddressAddr2, strCust_BillAddressAddr3, strCust_BillAddressAddr4, strCust_BillAddressCity, strCust_BillAddressState, strCust_BillAddressPostalCode, strCust_BillAddressCountry, strCust_ShipAddressAddr1, strCust_ShipAddressAddr3, strCust_ShipAddressAddr4, strCust_ShipAddressCity, strCust_ShipAddressState, strCust_ShipAddressPostalCode, strCust_ShipAddressCountry, strCust_Phone, strCust_AltPhone, strCust_Fax, strCust_Email, strCust_Contact, strCust_AltContact, strCust_CustomerTypeRefListID, strCust_CustomerTypeRefFullName, strCust_TermsRefListID, strCust_TermsRefFullName, strCust_SalesRepRefListID, strCust_SalesRepRefFullName, strCust_Balance, strCust_TotalBalance, strCust_OpenBalance, strCust_OpenBalanceDate, strCust_SalesTaxCodeRefListID, strCust_SalesTaxCodeRefFullName, strCust_ItemSalesTaxRefListID, strCust_ItemSalesTaxRefFullName, strCust_ResaleNumber, strCust_AccountNumber, strCust_CreditLimit, strCust_PreferredPaymentMethodRefListID, strCust_PreferredPaymentMethodRefFullName, strCust_CreditCardInfoCreditCardNumber, strCust_CreditCardInfoExpirationMonth, strCust_CreditCardInfoExpirationYear, strCust_CreditCardInfoNameOnCard, strCust_CreditCardInfoCreditCardAddress, strCust_CreditCardInfoCreditCardPostalCode, strCust_JobStatus, strCust_JobStartDate, strCust_JobProjectedEndDate, strCust_JobEndDate, strCust_JobDesc, strCust_JobTypeRefListID, strCust_JobTypeRefFullName, strCust_Notes, strCust_PriceLevelRefListID, strCust_PriceLevelRefFullName, strCust_CustomFieldOther As String
        Dim strCust_QB_CustomerSQL, strCust_QB_CustomerRow, strCust_ListID, strCust_TimeCreated, strCust_TimeModified, strCust_EditSequence, strCust_Name, strCust_FullName, strCust_IsActive, strCust_ParentRefListID, strCust_ParentRefFullName, strCust_Sublevel, strCust_CompanyName, strCust_Salutation, strCust_FirstName, strCust_MiddleName, strCust_LastName, strCust_BillAddressAddr1, strCust_BillAddressAddr2, strCust_BillAddressAddr3, strCust_BillAddressAddr4, strCust_BillAddressCity, strCust_BillAddressState, strCust_BillAddressPostalCode, strCust_BillAddressCountry As String
        'This routine gets the Cust_QB_Customer from the database according to the selection in strCust_QB_CustomerSQL.
        'It then puts those Cust_QB_Customer in the list box


        'FOR PART JobAddr_ - Get records from JOB_Address
        Debug.WriteLine("ListJobAddr_JOB_Address")

        'Dim strJobAddr_JOB_AddressSQL, strJobAddr_JOB_AddressRow, strJobAddr_SOAddressN, strJobAddr_SalesOrderN, strJobAddr_CreatedDate, strJobAddr_CreatedBy, strJobAddr_ModifiedDate, strJobAddr_ModifiedBy, strJobAddr_TypeAddress, strJobAddr_Name, strJobAddr_MaxClientID, strJobAddr_MaxContactNumber, strJobAddr_QBListID, strJobAddr_Address1, strJobAddr_Address2, strJobAddr_City, strJobAddr_State, strJobAddr_Zip, strJobAddr_Phone, strJobAddr_Fax, strJobAddr_Cell, strJobAddr_AcctNum As String
        Dim strJobAddr_JOB_AddressSQL, strJobAddr_JOB_AddressRow, strJobAddr_State As String
        'This routine gets the JobAddr_JOB_Address from the database according to the selection in strJobAddr_JOB_AddressSQL.
        'It then puts those JobAddr_JOB_Address in the list box


        'FOR PART Rep_ - Get records from DBPM_StateRepCSROP
        Debug.WriteLine("ListRep_DBPM_StateRepCSROP")

        Dim strRep_DBPM_StateRepCSROPSQL, strRep_DBPM_StateRepCSROPRow, strRep_RepCode As String
        'This routine gets the Rep_DBPM_StateRepCSROP from the database according to the selection in strRep_DBPM_StateRepCSROPSQL.
        'It then puts those Rep_DBPM_StateRepCSROP in the list box


        'FOR PART RepID_ - Get records from QB_SalesRep
        Debug.WriteLine("ListRepID_QB_SalesRep")
        Dim strRepID_QB_SalesRepSQL As String = "", strRepID_ListID As String = "", strRepID_Initial As String = ""
        'This routine gets the RepID_QB_SalesRep from the database according to the selection in strRepID_QB_SalesRepSQL.
        'It then puts those RepID_QB_SalesRep in the list box


        'FOR PART Terms_ - Get records from JOB_Step
        Debug.WriteLine("ListTerms_JOB_Step")

        Dim strTerms_JOB_StepSQL, strTerms_JOB_StepRow, strTerms_ARTerms As String '<--THIS IS THE ONLY ONE USED

        'This routine gets the Terms_JOB_Step from the database according to the selection in strTerms_JOB_StepSQL.
        'It then puts those Terms_JOB_Step in the list box


        'FOR PART TermsID_ - Get records from QB_Terms
        Debug.WriteLine("ListTermsID_QB_Terms")
        Dim strTermsID_QB_TermsSQL As String = "", strTermsID_ListID As String = "", strTermsID_Name As String = ""
        'This routine gets the TermsID_QB_Terms from the database according to the selection in strTermsID_QB_TermsSQL.
        'It then puts those TermsID_QB_Terms in the list box


        'FOR PART PricingInfo_ - Get records from vwT_JobPricing
        Debug.WriteLine("ListPricingInfo_vwT_JobPricing")
        Dim strPricingInfo_vwT_JobPricingSQL, strPricingInfo_vwT_JobPricingRow, strPricingInfo_Definition, strPricingInfo_Qty, strPricingInfo_SetupCharge, strPricingInfo_RunCharge, strPricingInfo_DiscountCode, strPricingInfo_ApprovedBy, strPricingInfo_SetupQBItem, strPricingInfo_RunQBItem, strPricingInfo_PriceComment, strPricingInfo_PriceDescription, strPricingInfo_Type, strPricingInfo_OrderBy, strPricingInfo_KeyNName, strPricingInfo_KeyN, strPricingInfo_StepN, strPricingInfo_SalesOrderN, strPricingInfo_OrderBy2, strPricingInfo_ComponantNum, strPricingInfo_StepNum As String
        'This routine gets the PricingInfo_vwT_JobPricing from the database according to the selection in strPricingInfo_vwT_JobPricingSQL.
        'It then puts those PricingInfo_vwT_JobPricing in the list box

        'Show what's processing

        Dim strTrackingNumber As String = ""
        Dim strShipToCarrier As String = ""

        ShowUserMessage(strSubName, "Loading Shipped Jobs Into Invoices", "Loading Shipped Jobs Into Invoices", True)

        Dim strToday As String = ""
        strToday = DateTime.Now.ToString("MM/dd/yyyy")

        Using Sql As New SQLHelper(gstrSQLConnectionString)

            '****************************************************************************************
            '***********TO INSERT NEW INVOICE INTO QB - USE BELOW SQL STATEMENT**********************
            '****************************************************************************************


            'New recordset
            strJobs_JOB_HeaderSQL = "SELECT jh.*,pp.columnprice FROM JOB_Header jh " & _
                                    "INNER JOIN Job_PrePayRequest pp ON jh.SalesOrderN = pp.SalesOrderN " & _
                                    "WHERE jh.DateShipped is not null " & _
                                    "AND jh.DateShipped > '07/25/2012'  AND jh.IsWillBeInvoiced = 1 " & _
                                    "AND jh.JobNum not in ( SELECT RefNumber FROM QB_Invoice WHERE TermsRefFullName <> 'Prepay') " & _
                                    "AND jh.IsInvoiced = 0 " & _
                                    "AND jh.BillToCompanyName not like 'Frazzled%' " & _
                                    "AND jh.BillToCompanyName not like 'Drummond%' " & _
                                    "AND jh.BillToCompanyName not like 'Drum-Line%' " & _
                                    "ORDER BY jh.DateShipped -- DESC"

            Debug.WriteLine(strJobs_JOB_HeaderSQL)

            Dim iRowCount As Integer = 0
            Using rsJobs_JOB_Header As SqlDataReader = Sql.ExecuteReader(CommandType.Text, strJobs_JOB_HeaderSQL)

                Dim bMultiShip As Boolean
                Dim strDay, strHour, strDayOfMonth, strDayOfMonthTomorrow As String
                Dim intNumDaysToAdd As Integer
                Dim strTxnDate, strSpecial, strAdditionalDesc, strSQLLoaded As String
                If rsJobs_JOB_Header.HasRows Then

                    While rsJobs_JOB_Header.Read
                        iRowCount += 1

                        'Show what's processing in the listbox
                        ShowUserMessage(strSubName, "Processing Jobs into Invoices")
                        ShowUserMessage(strSubName, "Processing Record " & iRowCount)

                        'get the columns from the database
                        Try


                            strJobs_SalesOrderN = NCStr(rsJobs_JOB_Header("SalesOrderN")).Replace("'"c, "`"c)
                            strJobs_CreatedDate = NCStr(rsJobs_JOB_Header("CreatedDate")).Replace("'"c, "`"c)
                            strJobs_CreatedBy = NCStr(rsJobs_JOB_Header("CreatedBy")).Replace("'"c, "`"c)
                            strJobs_CreatedOnComputer = NCStr(rsJobs_JOB_Header("CreatedOnComputer")).Replace("'"c, "`"c)
                            strJobs_ModifiedDate = NCStr(rsJobs_JOB_Header("ModifiedDate")).Replace("'"c, "`"c)
                            strJobs_ModifiedBy = NCStr(rsJobs_JOB_Header("ModifiedBy")).Replace("'"c, "`"c)
                            strJobs_BillToCompanyName = NCStr(rsJobs_JOB_Header("BillToCompanyName")).Replace("'"c, "`"c)
                            strJobs_BillToCompanyKeyMaxClientID = NCStr(rsJobs_JOB_Header("BillToCompanyKeyMaxClientID")).Replace("'"c, "`"c)
                            strJobs_BillToCompanyKeyMaxContactNum = NCStr(rsJobs_JOB_Header("BillToCompanyKeyMaxContactNum")).Replace("'"c, "`"c)
                            strJobs_BillToCompanyKeyQBListID = NCStr(rsJobs_JOB_Header("BillToCompanyKeyQBListID")).Replace("'"c, "`"c)
                            strJobs_BillToAddressName = NCStr(rsJobs_JOB_Header("BillToAddressName")).Replace("'"c, "`"c)
                            strJobs_BillToAddressKeyMaxClientID = NCStr(rsJobs_JOB_Header("BillToAddressKeyMaxClientID")).Replace("'"c, "`"c)
                            strJobs_BillToAddressKeyMaxContactNum = NCStr(rsJobs_JOB_Header("BillToAddressKeyMaxContactNum")).Replace("'"c, "`"c)
                            strJobs_BillToAddressKeyQBListID = NCStr(rsJobs_JOB_Header("BillToAddressKeyQBListID")).Replace("'"c, "`"c)
                            strJobs_InvoiceTxnID = NCStr(rsJobs_JOB_Header("InvoiceTxnID")).Replace("'"c, "`"c)
                            strJobs_CustPONum = NCStr(rsJobs_JOB_Header("CustPONum")).Replace("'"c, "`"c)
                            strJobs_CustPODate = NCStr(rsJobs_JOB_Header("CustPODate")).Replace("'"c, "`"c)
                            strJobs_CustPOAmt = NCStr(rsJobs_JOB_Header("CustPOAmt")).Replace("'"c, "`"c)
                            strJobs_TypeRec = NCStr(rsJobs_JOB_Header("TypeRec")).Replace("'"c, "`"c)
                            strJobs_JobNum = NCStr(rsJobs_JOB_Header("jobnum")).Replace("'"c, "`"c)
                            strJobs_QuoteNum = NCStr(rsJobs_JOB_Header("QuoteNum")).Replace("'"c, "`"c)
                            strJobs_TemplateNum = NCStr(rsJobs_JOB_Header("TemplateNum")).Replace("'"c, "`"c)
                            strJobs_Item = NCStr(rsJobs_JOB_Header("Item")).Replace("'"c, "`"c)
                            strJobs_ItemDesc = NCStr(rsJobs_JOB_Header("ItemDesc")).Replace("'"c, "`"c)
                            strJobs_ItemRefListID = NCStr(rsJobs_JOB_Header("ItemRefListID")).Replace("'"c, "`"c)
                            strJobs_ItemRefFullName = NCStr(rsJobs_JOB_Header("ItemRefFullName")).Replace("'"c, "`"c)
                            strJobs_Imprint = NCStr(rsJobs_JOB_Header("Imprint")).Replace("'"c, "`"c)
                            strJobs_Qty = NCStr(rsJobs_JOB_Header("QTY")).Replace("'"c, "`"c)
                            strJobs_DG = NCStr(rsJobs_JOB_Header("DG")).Replace("'"c, "`"c)
                            strJobs_ShipDate = NCStr(rsJobs_JOB_Header("ShipDate")).Replace("'"c, "`"c)
                            strJobs_AcctTerms = NCStr(rsJobs_JOB_Header("AcctTerms")).Replace("'"c, "`"c)
                            strJobs_AcctCreditLimit = NCStr(rsJobs_JOB_Header("AcctCreditLimit")).Replace("'"c, "`"c)
                            strJobs_AcctShippingClearance = NCStr(rsJobs_JOB_Header("AcctShippingClearance")).Replace("'"c, "`"c)
                            strJobs_AcctNeedBeforeShip = NCStr(rsJobs_JOB_Header("AcctNeedBeforeShip")).Replace("'"c, "`"c)
                            strJobs_CancelDate = NCStr(rsJobs_JOB_Header("CancelDate")).Replace("'"c, "`"c)
                            strJobs_CancelBy = NCStr(rsJobs_JOB_Header("CancelBy")).Replace("'"c, "`"c)
                            strJobs_CancelReason = NCStr(rsJobs_JOB_Header("CancelReason")).Replace("'"c, "`"c)
                            strJobs_Color = NCStr(rsJobs_JOB_Header("Color")).Replace("'"c, "`"c)
                            strJobs_Comments = NCStr(rsJobs_JOB_Header("Comments")).Replace("'"c, "`"c)
                            strJobs_Temp_Press = NCStr(rsJobs_JOB_Header("Temp_Press")).Replace("'"c, "`"c)
                            strJobs_Temp_ProofOK = NCStr(rsJobs_JOB_Header("Temp_ProofOK")).Replace("'"c, "`"c)
                            strJobs_Temp_Plated = NCStr(rsJobs_JOB_Header("Temp_Plated")).Replace("'"c, "`"c)
                            strJobs_Temp_DIDDE = NCStr(rsJobs_JOB_Header("Temp_DIDDE")).Replace("'"c, "`"c)
                            strJobs_Temp_QM = NCStr(rsJobs_JOB_Header("Temp_QM")).Replace("'"c, "`"c)
                            strJobs_Temp_Foiled = NCStr(rsJobs_JOB_Header("Temp_Foiled")).Replace("'"c, "`"c)
                            strJobs_Temp_ShippedQty = NCStr(rsJobs_JOB_Header("Temp_ShippedQty")).Replace("'"c, "`"c)
                            strJobs_CSR = NCStr(rsJobs_JOB_Header("CSR")).Replace("'"c, "`"c)
                            strJobs_OP = NCStr(rsJobs_JOB_Header("OP")).Replace("'"c, "`"c)
                            strJobs_TypeOrder = NCStr(rsJobs_JOB_Header("TypeOrder")).Replace("'"c, "`"c)
                            strJobs_PreviousJobs = NCStr(rsJobs_JOB_Header("PreviousJobs")).Replace("'"c, "`"c)
                            strJobs_IsNewCustomer = NCStr(rsJobs_JOB_Header("IsNewCustomer")).Replace("'"c, "`"c)
                            strJobs_IsCompleted = NCStr(rsJobs_JOB_Header("IsCompleted")).Replace("'"c, "`"c)
                            strJobs_IsInvoiced = NCStr(rsJobs_JOB_Header("IsInvoiced")).Replace("'"c, "`"c)
                            strJobs_IsCustom = NCStr(rsJobs_JOB_Header("IsCustom")).Replace("'"c, "`"c)
                            strJobs_IsReprint = NCStr(rsJobs_JOB_Header("IsReprint")).Replace("'"c, "`"c)
                            strJobs_IsPromo = NCStr(rsJobs_JOB_Header("IsPromo")).Replace("'"c, "`"c)
                            strJobs_IsRush = NCStr(rsJobs_JOB_Header("IsRush")).Replace("'"c, "`"c)
                            strJobs_AcctFaxCreditApp = NCStr(rsJobs_JOB_Header("AcctFaxCreditApp")).Replace("'"c, "`"c)
                            strJobs_ShipDateC = NCStr(rsJobs_JOB_Header("ShipDateC")).Replace("'"c, "`"c)
                            strJobs_ColorC = NCStr(rsJobs_JOB_Header("ColorC")).Replace("'"c, "`"c)
                            strJobs_CommentsC = NCStr(rsJobs_JOB_Header("CommentsC")).Replace("'"c, "`"c)
                            strJobs_ProcessType = NCStr(rsJobs_JOB_Header("ProcessType")).Replace("'"c, "`"c)
                            strJobs_ForQtyFrom = NCStr(rsJobs_JOB_Header("ForQtyFrom")).Replace("'"c, "`"c)
                            strJobs_ForQtyTo = NCStr(rsJobs_JOB_Header("ForQtyTo")).Replace("'"c, "`"c)
                            strJobs_IsWillBeInvoiced = NCStr(rsJobs_JOB_Header("IsWillBeInvoiced")).Replace("'"c, "`"c)
                            strJobs_IsShowOnProduction = NCStr(rsJobs_JOB_Header("IsShowOnProduction")).Replace("'"c, "`"c)
                            strJobs_IsFiftyFreeBC1 = NCStr(rsJobs_JOB_Header("IsFiftyFreeBC1")).Replace("'"c, "`"c)
                            strJobs_SourceCode = NCStr(rsJobs_JOB_Header("SourceCode")).Replace("'"c, "`"c)
                            strJobs_IsNotForMktg = NCStr(rsJobs_JOB_Header("IsNotForMktg")).Replace("'"c, "`"c)
                            strJobs_IsRev_Item = NCStr(rsJobs_JOB_Header("IsRev_Item")).Replace("'"c, "`"c)
                            strJobs_IsRev_Qty = NCStr(rsJobs_JOB_Header("IsRev_Qty")).Replace("'"c, "`"c)
                            strJobs_IsRev_Imprint = NCStr(rsJobs_JOB_Header("IsRev_Imprint")).Replace("'"c, "`"c)
                            strJobs_IsRev_TypeOrder = NCStr(rsJobs_JOB_Header("IsRev_TypeOrder")).Replace("'"c, "`"c)
                            strJobs_IsRev_PreviousJobs = NCStr(rsJobs_JOB_Header("IsRev_PreviousJobs")).Replace("'"c, "`"c)
                            strJobs_IsRev_ShipDate = NCStr(rsJobs_JOB_Header("IsRev_ShipDate")).Replace("'"c, "`"c)
                            strJobs_SetupCharge = NCStr(rsJobs_JOB_Header("SetupCharge")).Replace("'"c, "`"c)
                            strJobs_SetupQBItem = NCStr(rsJobs_JOB_Header("SetupQBItem")).Replace("'"c, "`"c)
                            strJobs_RunCharge = NCStr(rsJobs_JOB_Header("columnprice")).Replace("'"c, "`"c)
                            strJobs_RunQBItem = NCStr(rsJobs_JOB_Header("RunQBItem")).Replace("'"c, "`"c)
                            strJobs_PriceDiscountCode = NCStr(rsJobs_JOB_Header("PriceDiscountCode")).Replace("'"c, "`"c)
                            strJobs_PriceApprovedBy = NCStr(rsJobs_JOB_Header("PriceApprovedBy")).Replace("'"c, "`"c)
                            strJobs_PriceComment = NCStr(rsJobs_JOB_Header("PriceComment")).Replace("'"c, "`"c)

                            'Put the information together into a string
                            'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                            strJobs_JOB_HeaderRow = "" & _
                                                    strJobs_SalesOrderN & "  | " & _
                                                    strJobs_CreatedDate & "  | " & _
                                                    strJobs_CreatedBy & "  | " & _
                                                    strJobs_CreatedOnComputer & "  | " & _
                                                    strJobs_ModifiedDate & "  | " & _
                                                    strJobs_ModifiedBy & "  | " & _
                                                    strJobs_BillToCompanyName & "  | " & _
                                                    strJobs_BillToCompanyKeyMaxClientID & "  | " & _
                                                    strJobs_BillToCompanyKeyMaxContactNum & "  | " & _
                                                    strJobs_BillToCompanyKeyQBListID & "  | " & _
                                                    strJobs_BillToAddressName & "  | " & _
                                                    strJobs_BillToAddressKeyMaxClientID & "  | " & _
                                                    strJobs_BillToAddressKeyMaxContactNum & "  | " & _
                                                    strJobs_BillToAddressKeyQBListID & "  | " & _
                                                    strJobs_InvoiceTxnID & "  | " & _
                                                    "" & Strings.Chr(9)



                            '**********************************************************
                            '******** INSERT THIS JOB INTO INVOICE LOG ****************
                            '**********************************************************

                            Try
                                strSQL = "INSERT INTO InvoiceLog (JobNumber, BillCompany, DateShipped, ErrorCount) VALUES ('" & strJobs_JobNum & "','" & strJobs_BillToCompanyName & "','" & strJobs_DateShipped & "', 0)"
                                SQLHelper.ExecuteSQL(cnDBPM, strSQL)

                            Catch ex As Exception
                                HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                                Exit Try
                            End Try



                            'put the line in the listbox
                            ShowUserMessage(strSubName, strJobs_JOB_HeaderRow)

                            'GET THE TERMS                        
                            strTerms_JOB_StepSQL = "SELECT ARTerms FROM JOB_Step WHERE StepType = 'CreditCheck' AND SalesOrderN = '" & strJobs_SalesOrderN & "'"
                            Debug.WriteLine(strTerms_JOB_StepSQL)

                            strTerms_ARTerms = SQLHelper.ExecuteScalerString(cnDBPM, CommandType.Text, strTerms_JOB_StepSQL).Replace("'"c, "`"c)

                            'Put the information together into a string
                            'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                            strTerms_JOB_StepRow = "" & _
                                                    strTerms_ARTerms & "  | " & _
                                                    "" & Strings.Chr(9)

                            'then get an rs of the Terms and ListID of Terms

                            strTermsID_QB_TermsSQL = "SELECT ListID, Name FROM QB_Terms WHERE Name = '" & strTerms_ARTerms & "'"
                            Debug.WriteLine(strTermsID_QB_TermsSQL)

                            Using dReader As SqlDataReader = SQLHelper.ExecuteReader(cnDBPM, CommandType.Text, strTermsID_QB_TermsSQL)
                                While dReader.Read
                                    strTermsID_ListID = NCStr(dReader("ListID"), "70000-1134763425").Replace("'"c, "`"c)
                                    strTermsID_Name = NCStr(dReader("Name"), "1").Replace("'"c, "`"c)
                                End While
                            End Using


                            'GET THE BILL TO ADDRESS INFO (From JOB_Addresses?)  From QB_Customer

                            strCust_QB_CustomerSQL = "SELECT * FROM QB_Customer WHERE ListID = '" & strJobs_BillToCompanyKeyQBListID & "'"
                            Debug.WriteLine(strCust_QB_CustomerSQL)

                            Using rsCust_QB_Customer As SqlDataReader = SQLHelper.ExecuteReader(cnDBPM, CommandType.Text, strCust_QB_CustomerSQL)


                                If rsCust_QB_Customer.Read() Then

                                    'get the columns from the database
                                    strCust_ListID = NCStr(rsCust_QB_Customer("ListID")).Replace("'"c, "`"c)
                                    strCust_TimeCreated = NCStr(rsCust_QB_Customer("TimeCreated")).Replace("'"c, "`"c)
                                    strCust_TimeModified = NCStr(rsCust_QB_Customer("TimeModified")).Replace("'"c, "`"c)
                                    strCust_EditSequence = NCStr(rsCust_QB_Customer("EditSequence")).Replace("'"c, "`"c)
                                    strCust_Name = NCStr(rsCust_QB_Customer("Name")).Replace("'"c, "`"c)
                                    strCust_FullName = NCStr(rsCust_QB_Customer("FullName")).Replace("'"c, "`"c)
                                    strCust_IsActive = NCStr(rsCust_QB_Customer("IsActive")).Replace("'"c, "`"c)
                                    strCust_ParentRefListID = NCStr(rsCust_QB_Customer("ParentRefListID")).Replace("'"c, "`"c)
                                    strCust_ParentRefFullName = NCStr(rsCust_QB_Customer("ParentRefFullName")).Replace("'"c, "`"c)
                                    strCust_Sublevel = NCStr(rsCust_QB_Customer("Sublevel")).Replace("'"c, "`"c)
                                    strCust_CompanyName = NCStr(rsCust_QB_Customer("CompanyName")).Replace("'"c, "`"c)
                                    strCust_Salutation = NCStr(rsCust_QB_Customer("Salutation")).Replace("'"c, "`"c)
                                    strCust_FirstName = NCStr(rsCust_QB_Customer("FirstName")).Replace("'"c, "`"c)
                                    strCust_MiddleName = NCStr(rsCust_QB_Customer("MiddleName")).Replace("'"c, "`"c)
                                    strCust_LastName = NCStr(rsCust_QB_Customer("LastName")).Replace("'"c, "`"c)
                                    strCust_BillAddressAddr1 = NCStr(rsCust_QB_Customer("BillAddressAddr1")).Replace("'"c, "`"c)
                                    strCust_BillAddressAddr2 = NCStr(rsCust_QB_Customer("BillAddressAddr2")).Replace("'"c, "`"c)
                                    strCust_BillAddressAddr3 = NCStr(rsCust_QB_Customer("BillAddressAddr3")).Replace("'"c, "`"c)
                                    strCust_BillAddressAddr4 = NCStr(rsCust_QB_Customer("BillAddressAddr4")).Replace("'"c, "`"c)
                                    strCust_BillAddressCity = NCStr(rsCust_QB_Customer("BillAddressCity")).Replace("'"c, "`"c)
                                    strCust_BillAddressState = NCStr(rsCust_QB_Customer("BillAddressState")).Replace("'"c, "`"c)
                                    strCust_BillAddressPostalCode = NCStr(rsCust_QB_Customer("BillAddressPostalCode")).Replace("'"c, "`"c)
                                    strCust_BillAddressCountry = NCStr(rsCust_QB_Customer("BillAddressCountry")).Replace("'"c, "`"c)

                                    'Put the information together into a string
                                    'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                                    strCust_QB_CustomerRow = "" & _
                                                             strCust_ListID & "  | " & _
                                                             strCust_TimeCreated & "  | " & _
                                                             strCust_TimeModified & "  | " & _
                                                             strCust_EditSequence & "  | " & _
                                                             strCust_Name & "  | " & _
                                                             strCust_FullName & "  | " & _
                                                             strCust_IsActive & "  | " & _
                                                             strCust_ParentRefListID & "  | " & _
                                                             strCust_ParentRefFullName & "  | " & _
                                                             strCust_Sublevel & "  | " & _
                                                             strCust_CompanyName & "  | " & _
                                                             strCust_Salutation & "  | " & _
                                                             strCust_FirstName & "  | " & _
                                                             strCust_MiddleName & "  | " & _
                                                             strCust_LastName & "  | " & _
                                                             "" & Strings.Chr(9)

                                    'put the line in the listbox
                                    'Debug.Print Now & "   " & rsCust_QB_Customer.tables(0).Rows.IndexOf(iteration_row) & " of " & rsCust_QB_Customer.RecordCount
                                    ShowUserMessage(strSubName, strCust_QB_CustomerRow)




                                    'DO WORK: With each record

                                    strQBIL_BillAddressAddr1 = strCust_BillAddressAddr1 'From Cust or InvLine x4
                                    strQBIL_BillAddressAddr2 = strCust_BillAddressAddr2 'From Cust or InvLine x4
                                    strQBIL_BillAddressAddr3 = strCust_BillAddressAddr3 'From Cust or InvLine x4
                                    strQBIL_BillAddressAddr4 = strCust_BillAddressAddr4 'From Cust or InvLine x4
                                    strQBIL_BillAddressCity = strCust_BillAddressCity 'AVAILABLE for CustNum?  no cant
                                    strQBIL_BillAddressState = strCust_BillAddressState
                                    strQBIL_BillAddressPostalCode = strCust_BillAddressPostalCode
                                    strQBIL_BillAddressCountry = strCust_BillAddressCountry

                                End If


                            End Using



                            'GET THE REP & REP ListID (New table) ( DBPM_StateRepCSROP & QB_SalesRep )

                            strJobAddr_JOB_AddressSQL = "SELECT State FROM JOB_Address WHERE TypeAddress = 'SoldTo' AND SalesOrderN = '" & strJobs_SalesOrderN & "'"
                            Debug.WriteLine(strJobAddr_JOB_AddressSQL)

                            strJobAddr_State = NCStr(SQLHelper.ExecuteScaler(cnDBPM, CommandType.Text, strJobAddr_JOB_AddressSQL)).Replace("'"c, "`"c)

                            'Put the information together into a string
                            'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                            strJobAddr_JOB_AddressRow = "" & _
                                                        strJobAddr_State & "  | " & _
                                                        "" & Strings.Chr(9)


                            ShowUserMessage(strSubName, strJobAddr_JOB_AddressRow)



                            'then get an rs of DBPM_StateRepCSROP

                            strRep_DBPM_StateRepCSROPSQL = "SELECT RepCode FROM DBPM_StateRepCSROP WHERE State = '" & strJobAddr_State & "'"
                            Debug.WriteLine(strRep_DBPM_StateRepCSROPSQL)

                            strRep_RepCode = SQLHelper.ExecuteScalerString(cnDBPM, CommandType.Text, strRep_DBPM_StateRepCSROPSQL).Replace("'"c, "`"c)


                            strRep_DBPM_StateRepCSROPRow = "" & _
                                                            strRep_RepCode & "  | " & _
                                                            "" & Strings.Chr(9)

                            'put the line in the listbox
                            ShowUserMessage(strSubName, "Reps_ " & strRep_DBPM_StateRepCSROPRow)


                            'then get an rs of the Rep and ListID of Rep
                            strRepID_QB_SalesRepSQL = "SELECT ListID,Initial FROM QB_SalesRep WHERE Initial = '" & strRep_RepCode & "'"
                            Debug.WriteLine(strRepID_QB_SalesRepSQL)

                            Using rsRepID_QB_SalesRep As SqlDataReader = SQLHelper.ExecuteReader(cnDBPM, CommandType.Text, strRepID_QB_SalesRepSQL)
                                While rsRepID_QB_SalesRep.Read
                                    strRepID_ListID = NCStr(rsRepID_QB_SalesRep("ListID"), "50000-1135731044").Replace("'"c, "`"c)
                                    strRepID_Initial = NCStr(rsRepID_QB_SalesRep("Initial"), "Open").Replace("'"c, "`"c)
                                End While
                            End Using


                            'GET THE SHIP TO INFO:

                            'Get rs of Ship info
                            'New recordset
                            strShipInfo_JOB_StepSQL = "SELECT * FROM JOB_Step WHERE StepType = 'Ship' AND SalesOrderN = '" & strJobs_SalesOrderN & "'"
                            Debug.WriteLine(strShipInfo_JOB_StepSQL)

                            Using rsShipInfo_JOB_Step As DataSet = SQLHelper.ExecuteDataSet(cnDBPM, CommandType.Text, strShipInfo_JOB_StepSQL)

                                'if More than one shipment  -set QB ShipTo = "Multiple Drops"

                                bMultiShip = False
                                If rsShipInfo_JOB_Step.Tables(0).Rows.Count > 1 Then
                                    bMultiShip = True
                                End If

                                For Each iteration_row_8 As DataRow In rsShipInfo_JOB_Step.Tables(0).Rows
                                    strShipInfo_StepCompletedDate = NCStr(iteration_row_8("StepCompletedDate")).Replace("'"c, "`"c)
                                    strShipInfo_ShipToCompany = NCStr(iteration_row_8("ShipToCompany")).Replace("'"c, "`"c)
                                    strShipInfo_ShipToPerson = NCStr(iteration_row_8("ShipToPerson")).Replace("'"c, "`"c)
                                    strShipInfo_ShipToAddr1 = NCStr(iteration_row_8("ShipToAddr1")).Replace("'"c, "`"c)
                                    strShipInfo_ShipToAddr2 = NCStr(iteration_row_8("ShipToAddr2")).Replace("'"c, "`"c)
                                    strShipInfo_ShipToCity = NCStr(iteration_row_8("ShipToCity")).Replace("'"c, "`"c)
                                    strShipInfo_ShipToState = NCStr(iteration_row_8("ShipToState")).Replace("'"c, "`"c)
                                    strShipInfo_ShipToZip = NCStr(iteration_row_8("ShipToZip")).Replace("'"c, "`"c)
                                    strShipInfo_ShipToCarrier = NCStr(iteration_row_8("ShipToCarrier")).Replace("'"c, "`"c)
                                    strShipInfo_ShipQtyShipped = NCStr(iteration_row_8("ShipQtyShipped")).Replace("'"c, "`"c)


                                    If bMultiShip Then
                                        strShipInfo_ShipToDescription = ""
                                        strShipInfo_ShipToCompany = "MULTIPLE Shipments"
                                        strShipInfo_ShipToPerson = ""
                                        strShipInfo_ShipToAddr1 = ""
                                        strShipInfo_ShipToAddr2 = ""
                                        strShipInfo_ShipToCity = ""
                                        strShipInfo_ShipToState = ""
                                        strShipInfo_ShipToZip = ""
                                        strShipInfo_ShipToPhone = ""
                                        strShipInfo_ShipToEmail = ""
                                        strShipInfo_ShipToQty = ""
                                    End If

                                    ' '******************************************
                                    ' '**** MOVED HERE FROM OUTSIDE SHIPMENT LOOP
                                    ' '******************************************

                                    'GET THE ITEM INFO
                                    'get IOC info
                                    Dim hashItemOtherCharge As Hashtable = ItemOtherCharge_GetInfo(strJobs_Item) '(str7IT_ItemSaysMasterItemList)  '(strItemName)

                                    If hashItemOtherCharge("gstrIOC_Name") = "" Then
                                        'item not found in IOC so get the plain GL info again to put in (USE FORCE GL VARS)
                                        strQBIL_InvoiceLineItemRefListID = "5A0000-1136927332" 'strForceGL_ListID                  'From QB_OtherItem per InvText x2
                                        strQBIL_InvoiceLineItemRefFullName = "DP-1" 'strForceGL_Name                'From QB_OtherItem per InvText x2
                                        strQBIL_InvoiceLineDesc = "" 'str7IT_InvLineDesc 'strForceGL_SalesOrPurchaseDesc '<-- swapped description for real  'From GL Name or InvText Description
                                        strQBIL_InvoiceLineQuantity = "" 'str7IT_InvLineQty   '"1"          '1 or from InvText Qty
                                        strQBIL_InvoiceLineRate = "" 'skip so will default to Amt / Qty ?  try sample
                                        strQBIL_InvoiceLineAmount = "" 'str7IT_InvLineAmt   'strGLAmountLeft  'str5GL_Transactions                         'From GL TranAmt or InvText Amt

                                    Else
                                        'item was found in IOC so fill the vars
                                        strQBIL_InvoiceLineItemRefListID = hashItemOtherCharge("gstrIOC_ListID") 'From QB_OtherItem per InvText x2
                                        strQBIL_InvoiceLineItemRefFullName = hashItemOtherCharge("gstrIOC_Name") 'From QB_OtherItem per InvText x2
                                        strQBIL_InvoiceLineDesc = hashItemOtherCharge("gstrIOC_SalesOrPurchaseDesc") 'gstrIOC_SalesOrPurchaseDesc            'From GL Name or InvText Description
                                        strQBIL_InvoiceLineQuantity = "" 'str7IT_InvLineQty   '"1"          '1 or from InvText Qty
                                        strQBIL_InvoiceLineRate = "" 'skip so will default to Amt / Qty ?  try sample
                                        strQBIL_InvoiceLineAmount = "" 'str7IT_InvLineAmt   'strGLAmountLeft  'str5GL_Transactions    'From GL TranAmt or InvText Amt

                                    End If

                                    'FOR PART ShipMethod_ - Get records from QB_ShipMethod
                                    'This routine gets the ShipMethod_QB_ShipMethod from the database according to the selection in strShipMethod_QB_ShipMethodSQL.
                                    'It then puts those ShipMethod_QB_ShipMethod in the list box
                                    Debug.WriteLine("ListShipMethod_QB_ShipMethod")
                                    Dim strShipMethod_QB_ShipMethodSQL, strShipMethod_QB_ShipMethodRow, strShipMethod_ListID, strShipMethod_Name As String

                                    strShipMethod_QB_ShipMethodSQL = "SELECT ListID, Name FROM QB_ShipMethod WHERE Name = '" & strShipInfo_ShipToCarrier & "'"

                                    Using rsShipMethod_QB_ShipMethod As SqlDataReader = SQLHelper.ExecuteReader(cnDBPM, CommandType.Text, strShipMethod_QB_ShipMethodSQL)
                                        If rsShipMethod_QB_ShipMethod.Read Then
                                            'Clear strings
                                            strShipMethod_ListID = NCStr(rsShipMethod_QB_ShipMethod("ListID")).Replace("'"c, "`"c)
                                            strShipMethod_Name = NCStr(rsShipMethod_QB_ShipMethod("Name")).Replace("'"c, "`"c)

                                            'Put the information together into a string
                                            'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                                            strShipMethod_QB_ShipMethodRow = "" & _
                                                                             strShipMethod_ListID & "  | " & _
                                                                             strShipMethod_Name & "  | " & _
                                                                             "" & Strings.Chr(9)

                                            ShowUserMessage(strSubName, "ShipMethod_   " & strShipMethod_QB_ShipMethodRow)

                                        Else
                                            strShipMethod_ListID = "90000-1141327187" 'str3QBCust_ShipMethodRefListID                            'From Customer x2
                                            strShipMethod_Name = "Fed-Ex Ground" 'str3QBCust_ShipMethodRefFullName                          'From Customer x2
                                        End If
                                    End Using


                                    'GET THE TxnDate (One work day after ShipDate. Account for Fri to Mon etc.)
                                    strDay = DateTime.Now.ToString("ddd")
                                    strHour = DateTime.Now.ToString("HHmm")
                                    strDayOfMonth = DateTime.Now.ToString("d")
                                    strDayOfMonthTomorrow = DateTime.Now.AddDays(1).ToString("d")

                                    intNumDaysToAdd = 1

                                    If strDay = "Fri" Then intNumDaysToAdd = 3
                                    If strDay = "Sat" Then intNumDaysToAdd = 2
                                    If strDay = "Sun" Then intNumDaysToAdd = 1

                                    'Spoof DateShipped if it is empty
                                    If strJobs_DateShipped = "" Then strJobs_DateShipped = DateTime.Now.ToString("MM/dd/yyyy")

                                    'Now set the TxnDate with confidence
                                    If strShipInfo_StepCompletedDate = "" Then
                                        strTxnDate = CDate(strJobs_DateShipped).AddDays(intNumDaysToAdd).ToString("MM/dd/yyyy")
                                    Else
                                        strTxnDate = CDate(strShipInfo_StepCompletedDate).AddDays(intNumDaysToAdd).ToString("MM/dd/yyyy")
                                    End If

                                    strSpecial = ""

                                    'insert into QuickBooks

                                    'FILL VARS FOR THE INSERT!!!

                                    ''Here are the InvoiceLine vars to fill
                                    ''Fill var strings

                                    strQBIL_CustomerRefListID = strJobs_BillToCompanyKeyQBListID 'From Customer x2
                                    'strQBIL_CustomerRefFullName = str3QBCust_FullName                       'From Customer x2

                                    strQBIL_ARAccountRefListID = "C60000-1136231019" 'Forced to A/R Trade x2
                                    strQBIL_ARAccountRefFullName = "A/R-Trade" 'Forced to A/R Trade x2

                                    strQBIL_TemplateRefListID = "1E0000-1139412167" 'Created new Template.  Retrieved Refs x2
                                    strQBIL_TemplateRefFullName = ".Drummond Invoice" 'Created new Template.  Retrieved Refs x2

                                    Dim TempDate As Date
                                    strQBIL_TxnDate = IIf(DateTime.TryParse(strTxnDate, TempDate), TempDate.ToString("yyyy-MM-dd"), strTxnDate) 'InvText InvDate  -else GL PostDate

                                    strQBIL_RefNumber = strJobs_JobNum 'Inv Num

                                    strQBIL_ShipAddressAddr1 = strShipInfo_ShipToCompany 'From InvLine only x4
                                    strQBIL_ShipAddressAddr2 = strShipInfo_ShipToPerson 'From InvLine only x4
                                    strQBIL_ShipAddressAddr3 = strShipInfo_ShipToAddr1 'From InvLine only x4
                                    strQBIL_ShipAddressAddr4 = strShipInfo_ShipToAddr2 'From InvLine only x4
                                    strQBIL_ShipAddressCity = strShipInfo_ShipToCity
                                    strQBIL_ShipAddressState = strShipInfo_ShipToState
                                    strQBIL_ShipAddressPostalCode = strShipInfo_ShipToZip
                                    strQBIL_ShipAddressCountry = ""

                                    strQBIL_IsPending = "0" '0
                                    strQBIL_PONumber = strJobs_CustPONum

                                    strQBIL_TermsRefListID = NCStr(strTermsID_ListID)
                                    strQBIL_TermsRefFullName = NCStr(strTermsID_Name)

                                    strQBIL_SalesRepRefListID = NCStr(strRepID_ListID)
                                    strQBIL_SalesRepRefFullName = NCStr(strRepID_Initial)

                                    strQBIL_FOB = ""

                                    If strShipInfo_StepCompletedDate <> "" Then
                                        Dim TempDate2 As Date
                                        strQBIL_ShipDate = IIf(DateTime.TryParse(strShipInfo_StepCompletedDate, TempDate2), TempDate2.ToString("yyyy-MM-dd"), strShipInfo_StepCompletedDate) 'InvText? ShipDate  -else GL PostDate
                                    Else
                                        Dim TempDate3 As Date
                                        strQBIL_ShipDate = IIf(DateTime.TryParse(strJobs_DateShipped, TempDate3), TempDate3.ToString("yyyy-MM-dd"), strJobs_DateShipped)
                                    End If


                                    'new:  ShipMethod
                                    strQBIL_ShipMethodRefListID = strShipMethod_ListID
                                    strQBIL_ShipMethodRefFullName = strShipMethod_Name


                                    strQBIL_ItemSalesTaxRefListID = "20000-1134760068" 'forced to 0 x2
                                    strQBIL_ItemSalesTaxRefFullName = "0" 'forced to 0 x2

                                    strQBIL_Memo = "" 'str5GL_Description  'str5GL_AuditTrail                                      'AVAILABLE   -AuditTrail?
                                    strQBIL_IsToBePrinted = "1" '0

                                    strQBIL_CustomerSalesTaxCodeRefListID = "20000-1134607388" 'forced to non  x2
                                    strQBIL_CustomerSalesTaxCodeRefFullName = "Non" 'forced to non  x2

                                    'Do line one of invoice lines (Item,Qty,Imprint, etc.)
                                    strQBIL_InvoiceLineType = "ILType" '"Other?"                         'AVAILABLE? Custom?
                                    strQBIL_InvoiceLineTxnLineID = "ILTxnLineID" 'AVAILABLE 36  'From GL RecCount or InvText LineNum ?

                                    If strShipInfo_ShipQtyShipped <> "" Then
                                        strQBIL_InvoiceLineQuantity = strShipInfo_ShipQtyShipped '"1"                       '1 or from InvText Qty
                                    Else
                                        strQBIL_InvoiceLineQuantity = strJobs_Qty '"0"
                                    End If

                                    If strJobs_RunCharge = "" Then strJobs_RunCharge = "1"
                                    If strQBIL_InvoiceLineQuantity = "" Then strQBIL_InvoiceLineQuantity = "1"

                                    strQBIL_InvoiceLineRate = strJobs_RunCharge '""    'NEED TO ADD RATE    'skip so will default to Amt / Qty ?  try sample
                                    strQBIL_InvoiceLineAmount = CStr(CDbl(strJobs_RunCharge) * CDbl(strQBIL_InvoiceLineQuantity)) 'NEED TO ADD AMT   'From GL TranAmt or InvText Amt

                                    strQBIL_InvoiceLineSalesTaxCodeRefListID = "20000-1134607388" 'forced to non x2
                                    strQBIL_InvoiceLineSalesTaxCodeRefFullName = "Non" 'forced to non x2

                                    'strQBIL_InvoiceLineOverrideItemAccountRefListID = gstrIOC_SalesOrPurchaseAccountRefListID      'forced to GL via QB_OtherItem x2
                                    'strQBIL_InvoiceLineOverrideItemAccountRefFullName = gstrIOC_SalesOrPurchaseAccountRefFullName  'forced to GL via QB_OtherItem x2

                                    Dim trimSetting As New siteConstants.AdvancedTrimSettings(29, True, "...")

                                    strQBIL_CustomFieldInvoiceLineOther1 = NCStr(strJobs_Imprint, trimSetting) 'IMPRINT
                                    strQBIL_CustomFieldInvoiceLineOther2 = strSpecial 'SPECIAL

                                    strQBIL_FQSaveToCache = "1"


                                    '*******************************************************
                                    '*******************************************************
                                    'DO THE INSERT ****
                                    '*******************************************************
                                    '*******************************************************
                                    Try
                                        If Not InsertIntoQBInvoiceLineJobs() Then Continue While
                                    Catch ex As Exception
                                        HaveError(strObjName, strSubName & ":659", CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                                        Continue While
                                    End Try


                                    ' '******************************************
                                    ' '**** MOVED HERE FROM OUTSIDE SHIPMENT LOOP
                                    ' '******************************************

                                Next

                            End Using


                            'Do other invoice lines (add-ons)
                            strPricingInfo_vwT_JobPricingSQL = "SELECT * FROM vw_JobPricing2 WHERE SalesOrderN = '" & strJobs_SalesOrderN & "' AND KeyNName <> 'SalesOrderN' ORDER BY OrderBy, ComponantNum, StepNum, StepN, OrderBy2, Definition  "
                            Debug.WriteLine(strPricingInfo_vwT_JobPricingSQL)

                            Using rsPricingInfo_vwT_JobPricing As SqlDataReader = SQLHelper.ExecuteReader(cnDBPM, CommandType.Text, strPricingInfo_vwT_JobPricingSQL)


                                While rsPricingInfo_vwT_JobPricing.Read

                                    'get the columns from the database
                                    strPricingInfo_Definition = NCStr(rsPricingInfo_vwT_JobPricing("Definition")).Replace("'"c, "`"c)
                                    strPricingInfo_Qty = NCStr(rsPricingInfo_vwT_JobPricing("QTY")).Replace("'"c, "`"c)
                                    strPricingInfo_SetupCharge = NCStr(rsPricingInfo_vwT_JobPricing("SetupCharge")).Replace("'"c, "`"c)
                                    strPricingInfo_RunCharge = NCStr(rsPricingInfo_vwT_JobPricing("RunCharge")).Replace("'"c, "`"c)
                                    strPricingInfo_DiscountCode = NCStr(rsPricingInfo_vwT_JobPricing("DiscountCode")).Replace("'"c, "`"c)
                                    strPricingInfo_ApprovedBy = NCStr(rsPricingInfo_vwT_JobPricing("ApprovedBy")).Replace("'"c, "`"c)
                                    strPricingInfo_SetupQBItem = NCStr(rsPricingInfo_vwT_JobPricing("SetupQBItem")).Replace("'"c, "`"c)
                                    strPricingInfo_RunQBItem = NCStr(rsPricingInfo_vwT_JobPricing("RunQBItem")).Replace("'"c, "`"c)
                                    strPricingInfo_PriceComment = NCStr(rsPricingInfo_vwT_JobPricing("PriceComment")).Replace("'"c, "`"c)
                                    strPricingInfo_PriceDescription = NCStr(rsPricingInfo_vwT_JobPricing("PriceDescription")).Replace("'"c, "`"c)
                                    strPricingInfo_Type = NCStr(rsPricingInfo_vwT_JobPricing("Type")).Replace("'"c, "`"c)
                                    strPricingInfo_OrderBy = NCStr(rsPricingInfo_vwT_JobPricing("OrderBy")).Replace("'"c, "`"c)
                                    strPricingInfo_KeyNName = NCStr(rsPricingInfo_vwT_JobPricing("KeyNName")).Replace("'"c, "`"c)
                                    strPricingInfo_KeyN = NCStr(rsPricingInfo_vwT_JobPricing("KeyN")).Replace("'"c, "`"c)
                                    strPricingInfo_StepN = NCStr(rsPricingInfo_vwT_JobPricing("StepN")).Replace("'"c, "`"c)
                                    strPricingInfo_SalesOrderN = NCStr(rsPricingInfo_vwT_JobPricing("salesordern")).Replace("'"c, "`"c)
                                    strPricingInfo_OrderBy2 = NCStr(rsPricingInfo_vwT_JobPricing("OrderBy2")).Replace("'"c, "`"c)
                                    strPricingInfo_ComponantNum = NCStr(rsPricingInfo_vwT_JobPricing("ComponantNum")).Replace("'"c, "`"c)
                                    strPricingInfo_StepNum = NCStr(rsPricingInfo_vwT_JobPricing("StepNum")).Replace("'"c, "`"c)
                                    strTrackingNumber = NCStr(rsPricingInfo_vwT_JobPricing("TrackingNumber"))
                                    strShipToCarrier = NCStr(rsPricingInfo_vwT_JobPricing("ShipToCarrier"))

                                    'Put the information together into a string
                                    'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                                    strPricingInfo_vwT_JobPricingRow = "" & _
                                                                       strPricingInfo_Definition & "  | " & _
                                                                       strPricingInfo_SetupCharge & "  | " & _
                                                                       strPricingInfo_RunCharge & "  | " & _
                                                                       strPricingInfo_DiscountCode & "  | " & _
                                                                       strPricingInfo_ApprovedBy & "  | " & _
                                                                       strPricingInfo_SetupQBItem & "  | " & _
                                                                       strPricingInfo_RunQBItem & "  | " & _
                                                                       strPricingInfo_PriceComment & "  | " & _
                                                                       strPricingInfo_Type & "  | " & _
                                                                       strPricingInfo_OrderBy & "  | " & _
                                                                       strPricingInfo_KeyNName & "  | " & _
                                                                       strPricingInfo_KeyN & "  | " & _
                                                                       strPricingInfo_StepN & "  | " & _
                                                                       strPricingInfo_SalesOrderN & "  | " & _
                                                                       strPricingInfo_OrderBy2 & "  | " & _
                                                                       "" & Strings.Chr(9)

                                    'put the line in the listbox

                                    ShowUserMessage(strObjName, "Pricing Info: " & strPricingInfo_vwT_JobPricingRow)


                                    'strSpecial = strPricingInfo_DiscountCode & "-" & strPricingInfo_ApprovedBy

                                    If strPricingInfo_DiscountCode <> "" Then
                                        strSpecial = strPricingInfo_DiscountCode
                                        strAdditionalDesc = "  Free per " & strPricingInfo_DiscountCode
                                    ElseIf strPricingInfo_ApprovedBy <> "" Then
                                        strSpecial = strPricingInfo_ApprovedBy
                                        strAdditionalDesc = "  Free per " & strPricingInfo_ApprovedBy
                                    Else
                                        strSpecial = ""
                                        strAdditionalDesc = ""
                                    End If



                                    'insert other invoice lines (add-ons)
                                    'Once for SetupCharge, once for RunCharge

                                    'SETUP:
                                    'Setup Charge  -use setup qty
                                    'If strPricingInfo_SetupCharge <> "0" And strPricingInfo_SetupQBItem <> "" Then

                                    If (strPricingInfo_SetupCharge <> "0" And strPricingInfo_SetupQBItem <> "") _
                                        Or (strPricingInfo_SetupCharge = "0" And strPricingInfo_SetupQBItem <> "" And (strPricingInfo_ApprovedBy <> "" Or strPricingInfo_DiscountCode <> "")) _
                                        Or (strPricingInfo_SetupQBItem = "Freight") Then

                                        'GET THE ITEM INFO FOR THE ADD-ON ITEM
                                        'get IOC info
                                        Dim hashSetupCharge As Hashtable = ItemOtherCharge_GetInfo(strPricingInfo_SetupQBItem) '(strJobs_Item) '(str7IT_ItemSaysMasterItemList)  '(strItemName)
                                        If hashSetupCharge("gstrIOC_Name") = "" Then
                                            strQBIL_InvoiceLineItemRefListID = "5A0000-1136927332" 'strForceGL_ListID                  'From QB_OtherItem per InvText x2
                                            strQBIL_InvoiceLineItemRefFullName = "DP-1" 'strForceGL_Name                'From QB_OtherItem per InvText x2
                                            strQBIL_InvoiceLineDesc = strPricingInfo_SetupQBItem & " ??" 'str7IT_InvLineDesc 'strForceGL_SalesOrPurchaseDesc '<-- swapped description for real  'From GL Name or InvText Description
                                        Else
                                            'item was found in IOC so fill the vars
                                            strQBIL_InvoiceLineItemRefListID = hashSetupCharge("gstrIOC_ListID") 'From QB_OtherItem per InvText x2
                                            strQBIL_InvoiceLineItemRefFullName = hashSetupCharge("gstrIOC_Name") 'From QB_OtherItem per InvText x2
                                            strQBIL_InvoiceLineDesc = hashSetupCharge("gstrIOC_SalesOrPurchaseDesc") 'gstrIOC_SalesOrPurchaseDesc            'From GL Name or InvText Description
                                        End If


                                        'Add to the InvLine Description if zero price but has approval
                                        If strSpecial <> "" And strPricingInfo_SetupCharge = "0" Then strQBIL_InvoiceLineDesc = strQBIL_InvoiceLineDesc & strAdditionalDesc

                                        'Correct InvLine Description for Invoice using new field -strPricingInfo_PriceDescription
                                        If strPricingInfo_PriceDescription <> "" Then strQBIL_InvoiceLineDesc = strPricingInfo_PriceDescription


                                        'FILL INSERT VARS
                                        strQBIL_InvoiceLineType = "ILType" '"Other?"                         'AVAILABLE? Custom?
                                        strQBIL_InvoiceLineTxnLineID = "ILTxnLineID" 'AVAILABLE 36  'From GL RecCount or InvText LineNum ?

                                        'Defined above
                                        'strQBIL_InvoiceLineItemRefListID = gstrIOC_ListID                  'From QB_OtherItem per InvText x2
                                        'strQBIL_InvoiceLineItemRefFullName = gstrIOC_Name                'From QB_OtherItem per InvText x2
                                        'strQBIL_InvoiceLineDesc = gstrIOC_SalesOrPurchaseDesc                           'From GL Name or InvText Description

                                        'Set Quantity for Invoice using new field -strPricingInfo_Qty
                                        strQBIL_InvoiceLineQuantity = strPricingInfo_Qty '"1"
                                        strQBIL_InvoiceLineRate = strPricingInfo_SetupCharge '""    'NEED TO ADD RATE                       'skip so will default to Amt / Qty ?  try sample
                                        strQBIL_InvoiceLineAmount = CStr(CDbl(strPricingInfo_SetupCharge) * CDbl(strQBIL_InvoiceLineQuantity)) '"1"  'NEED TO ADD AMT                        'From GL TranAmt or InvText Amt

                                        strQBIL_InvoiceLineSalesTaxCodeRefListID = "20000-1134607388" 'forced to non x2
                                        strQBIL_InvoiceLineSalesTaxCodeRefFullName = "Non" 'forced to non x2

                                        'strQBIL_InvoiceLineOverrideItemAccountRefListID = gstrIOC_SalesOrPurchaseAccountRefListID      'forced to GL via QB_OtherItem x2
                                        'strQBIL_InvoiceLineOverrideItemAccountRefFullName = gstrIOC_SalesOrPurchaseAccountRefFullName  'forced to GL via QB_OtherItem x2

                                        strQBIL_CustomFieldInvoiceLineOther1 = "" 'strJobs_Imprint      'IMPRINT
                                        If strTrackingNumber <> "" Then
                                            If strShipToCarrier.Contains("fed") Then
                                                strQBIL_CustomFieldInvoiceLineOther1 = "FedEx:" & strTrackingNumber
                                            ElseIf strShipToCarrier.Contains("ups") Then
                                                strQBIL_CustomFieldInvoiceLineOther1 = "UPS:" & strTrackingNumber
                                            ElseIf strShipToCarrier.Contains("usps") Then
                                                strQBIL_CustomFieldInvoiceLineOther1 = "USPS:" & strTrackingNumber
                                            Else
                                                strQBIL_CustomFieldInvoiceLineOther1 = "Track: " & strTrackingNumber
                                            End If
                                        End If
                                        strQBIL_CustomFieldInvoiceLineOther2 = strSpecial 'SPECIAL

                                        strQBIL_FQSaveToCache = "1"

                                        'DO THE INSERT
                                        Try
                                            If Not InsertIntoQBInvoiceLineJobs() Then Continue While
                                        Catch ex As Exception
                                            HaveError(strObjName, strSubName & ":818", CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                                            Continue While
                                        End Try
                                    End If

                                    If (strPricingInfo_RunCharge <> "0" And strPricingInfo_RunQBItem <> "") Or (strPricingInfo_RunCharge = "0" And strPricingInfo_RunQBItem <> "" And (strPricingInfo_ApprovedBy <> "" Or strPricingInfo_DiscountCode <> "")) Then

                                        'GET THE ITEM INFO FOR THE ADD-ON ITEM
                                        'get IOC info
                                        Dim hashRunCharge As Hashtable = ItemOtherCharge_GetInfo(strPricingInfo_RunQBItem)

                                        '(strJobs_Item) '(str7IT_ItemSaysMasterItemList)  '(strItemName)
                                        If hashRunCharge("gstrIOC_Name") = "" Then
                                            strQBIL_InvoiceLineItemRefListID = "5A0000-1136927332" 'strForceGL_ListID                  'From QB_OtherItem per InvText x2
                                            strQBIL_InvoiceLineItemRefFullName = "DP-1" 'strForceGL_Name                'From QB_OtherItem per InvText x2
                                            strQBIL_InvoiceLineDesc = strPricingInfo_RunQBItem & " ??" 'str7IT_InvLineDesc 'strForceGL_SalesOrPurchaseDesc '<-- swapped description for real  'From GL Name or InvText Description
                                        Else
                                            'item was found in IOC so fill the vars
                                            strQBIL_InvoiceLineItemRefListID = hashRunCharge("gstrIOC_ListID") 'From QB_OtherItem per InvText x2
                                            strQBIL_InvoiceLineItemRefFullName = hashRunCharge("gstrIOC_Name") 'From QB_OtherItem per InvText x2
                                            strQBIL_InvoiceLineDesc = hashRunCharge("gstrIOC_SalesOrPurchaseDesc") 'gstrIOC_SalesOrPurchaseDesc            'From GL Name or InvText Description
                                        End If

                                        'Add to the Inv Line Description if zero price but has approval
                                        If strSpecial <> "" And strPricingInfo_RunCharge = "0" Then strQBIL_InvoiceLineDesc = strQBIL_InvoiceLineDesc & strAdditionalDesc

                                        'Correct InvLine Description for Invoice using new field -strPricingInfo_PriceDescription
                                        If strPricingInfo_PriceDescription <> "" Then strQBIL_InvoiceLineDesc = strPricingInfo_PriceDescription


                                        'FILL INSERT VARS
                                        strQBIL_InvoiceLineType = "ILType" '"Other?"                         'AVAILABLE? Custom?
                                        strQBIL_InvoiceLineTxnLineID = "ILTxnLineID" 'AVAILABLE 36  'From GL RecCount or InvText LineNum ?

                                        strQBIL_InvoiceLineRate = strPricingInfo_RunCharge 'NEED TO ADD RATE                       'skip so will default to Amt / Qty ?  try sample
                                        strQBIL_InvoiceLineAmount = CStr(CDbl(strPricingInfo_RunCharge) * CDbl(strQBIL_InvoiceLineQuantity)) 'NEED TO ADD AMT                        'From GL TranAmt or InvText Amt

                                        strQBIL_InvoiceLineSalesTaxCodeRefListID = "20000-1134607388" 'forced to non x2
                                        strQBIL_InvoiceLineSalesTaxCodeRefFullName = "Non" 'forced to non x2

                                        strQBIL_CustomFieldInvoiceLineOther1 = "" 'strJobs_Imprint      'IMPRINT
                                        strQBIL_CustomFieldInvoiceLineOther2 = strSpecial 'SPECIAL

                                        strQBIL_FQSaveToCache = "1"

                                        'DO THE INSERT
                                        Try
                                            If Not InsertIntoQBInvoiceLineJobs() Then Continue While
                                        Catch ex As Exception
                                            HaveError(strObjName, strSubName & ":867", CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                                            Continue While
                                        End Try
                                    End If

                                End While
                            End Using


                            'Do last line of invoice lines (blank)
                            strQBIL_InvoiceLineDesc = ""
                            'Do the (final) insert (blank)
                            'insert into invoice
                            strQBIL_FQSaveToCache = "0"
                            InsertIntoQBInvoiceLineComment()

                            'Mark it loaded
                            strSQLLoaded = "UPDATE JOB_Header SET IsInvoiced = 1 WHERE SalesOrderN = '" & strJobs_SalesOrderN & "'"
                            SQLHelper.ExecuteSQL(cnDBPM, strSQLLoaded)

                            strSQLLoaded = "UPDATE InvoiceLog SET DateInvoiced = '" & Now.ToString & "', ErrorCount = (SELECT Count(*) FROM InvoiceLogItem WHERE JobNumber = '" & strJobs_JobNum & "') WHERE JobNumber = '" & strJobs_JobNum & "'"
                            SQLHelper.ExecuteSQL(cnDBPM, strSQLLoaded)

                        Catch ex As Exception
                            HaveError(strObjName, strSubName & ":891", CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "", ex)
                            Continue While


                        End Try
                    End While

                End If
            End Using

            'Update status listbox
            ShowUserMessage(strSubName, iRowCount.ToString & " Records Processed", , True)

        End Using

    End Sub

End Module