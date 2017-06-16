Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Windows.Forms


'**********************
'****** NOT USED ******
'**********************


'**************************************
'******* 1st Code Review Complete *****
'**************************************


Module modQBInsertCreditMemoLine

	'These are CreditMemoLine insert vars for QBCMLine_
	Public strQBCMLine_TxnID As String = ""
	Public strQBCMLine_TimeCreated As String = ""
	Public strQBCMLine_TimeModified As String = ""
	Public strQBCMLine_EditSequence As String = ""
	Public strQBCMLine_TxnNumber As String = ""
	Public strQBCMLine_CustomerRefListID As String = ""
	Public strQBCMLine_CustomerRefFullName As String = ""
	Public strQBCMLine_ClassRefListID As String = ""
	Public strQBCMLine_ClassRefFullName As String = ""
	Public strQBCMLine_ARAccountRefListID As String = ""
	Public strQBCMLine_ARAccountRefFullName As String = ""
	Public strQBCMLine_TemplateRefListID As String = ""
	Public strQBCMLine_TemplateRefFullName As String = ""
	Public strQBCMLine_TxnDate As String = ""
	Public strQBCMLine_TxnDateMacro As String = ""
	Public strQBCMLine_RefNumber As String = ""
	Public strQBCMLine_BillAddressAddr1 As String = ""
	Public strQBCMLine_BillAddressAddr2 As String = ""
	Public strQBCMLine_BillAddressAddr3 As String = ""
	Public strQBCMLine_BillAddressAddr4 As String = ""
	Public strQBCMLine_BillAddressCity As String = ""
	Public strQBCMLine_BillAddressState As String = ""
	Public strQBCMLine_BillAddressPostalCode As String = ""
	Public strQBCMLine_BillAddressCountry As String = ""
	Public strQBCMLine_ShipAddressAddr1 As String = ""
	Public strQBCMLine_ShipAddressAddr2 As String = ""
	Public strQBCMLine_ShipAddressAddr3 As String = ""
	Public strQBCMLine_ShipAddressAddr4 As String = ""
	Public strQBCMLine_ShipAddressCity As String = ""
	Public strQBCMLine_ShipAddressState As String = ""
	Public strQBCMLine_ShipAddressPostalCode As String = ""
	Public strQBCMLine_ShipAddressCountry As String = ""
	Public strQBCMLine_IsPending As String = ""
	Public strQBCMLine_PONumber As String = ""
	Public strQBCMLine_TermsRefListID As String = ""
	Public strQBCMLine_TermsRefFullName As String = ""
	Public strQBCMLine_DueDate As String = ""
	Public strQBCMLine_SalesRepRefListID As String = ""
	Public strQBCMLine_SalesRepRefFullName As String = ""
	Public strQBCMLine_FOB As String = ""
	Public strQBCMLine_ShipDate As String = ""
	Public strQBCMLine_ShipMethodRefListID As String = ""
	Public strQBCMLine_ShipMethodRefFullName As String = ""
	Public strQBCMLine_Subtotal As String = ""
	Public strQBCMLine_ItemSalesTaxRefListID As String = ""
	Public strQBCMLine_ItemSalesTaxRefFullName As String = ""
	Public strQBCMLine_SalesTaxPercentage As String = ""
	Public strQBCMLine_SalesTaxTotal As String = ""
	Public strQBCMLine_TotalAmount As String = ""
	Public strQBCMLine_CreditRemaining As String = ""
	Public strQBCMLine_Memo As String = ""
	Public strQBCMLine_CustomerMsgRefListID As String = ""
	Public strQBCMLine_CustomerMsgRefFullName As String = ""
	Public strQBCMLine_IsToBePrinted As String = ""
	Public strQBCMLine_CustomerSalesTaxCodeRefListID As String = ""
	Public strQBCMLine_CustomerSalesTaxCodeRefFullName As String = ""
	Public strQBCMLine_CreditMemoLineType As String = ""
	Public strQBCMLine_CreditMemoLineSeqNo As String = ""
	Public strQBCMLine_CreditMemoLineGroupLineTxnLineID As String = ""
	Public strQBCMLine_CreditMemoLineGroupItemGroupRefListID As String = ""
	Public strQBCMLine_CreditMemoLineGroupItemGroupRefFullName As String = ""
	Public strQBCMLine_CreditMemoLineGroupDesc As String = ""
	Public strQBCMLine_CreditMemoLineGroupQuantity As String = ""
	Public strQBCMLine_CreditMemoLineGroupIsPrintItemsInGroup As String = ""
	Public strQBCMLine_CreditMemoLineGroupTotalAmount As String = ""
	Public strQBCMLine_CreditMemoLineGroupSeqNo As String = ""
	Public strQBCMLine_CreditMemoLineTxnLineID As String = ""
	Public strQBCMLine_CreditMemoLineItemRefListID As String = ""
	Public strQBCMLine_CreditMemoLineItemRefFullName As String = ""
	Public strQBCMLine_CreditMemoLineDesc As String = ""
	Public strQBCMLine_CreditMemoLineQuantity As String = ""
	Public strQBCMLine_CreditMemoLineRate As String = ""
	Public strQBCMLine_CreditMemoLineRatePercent As String = ""
	Public strQBCMLine_CreditMemoLinePriceLevelRefListID As String = ""
	Public strQBCMLine_CreditMemoLinePriceLevelRefFullName As String = ""
	Public strQBCMLine_CreditMemoLineClassRefListID As String = ""
	Public strQBCMLine_CreditMemoLineClassRefFullName As String = ""
	Public strQBCMLine_CreditMemoLineAmount As String = ""
	Public strQBCMLine_CreditMemoLineServiceDate As String = ""
	Public strQBCMLine_CreditMemoLineSalesTaxCodeRefListID As String = ""
	Public strQBCMLine_CreditMemoLineSalesTaxCodeRefFullName As String = ""
	Public strQBCMLine_CreditMemoLineIsTaxable As String = ""
	Public strQBCMLine_CreditMemoLineOverrideItemAccountRefListID As String = ""
	Public strQBCMLine_CreditMemoLineOverrideItemAccountRefFullName As String = ""
	Public strQBCMLine_FQSaveToCache As String = ""
	Public strQBCMLine_FQPrimaryKey As String = ""
	Public strQBCMLine_CustomFieldCreditMemoLineOther1 As String = ""
	Public strQBCMLine_CustomFieldCreditMemoLineOther2 As String = ""
	Public strQBCMLine_CustomFieldCreditMemoLineGroupOther1 As String = ""
	Public strQBCMLine_CustomFieldCreditMemoLineGroupOther2 As String = ""
	Public strQBCMLine_CustomFieldCreditMemoLineGroupLineOther1 As String = ""
	Public strQBCMLine_CustomFieldCreditMemoLineGroupLineOther2 As String = ""
	Public strQBCMLine_CustomFieldOther As String = ""
	'This routine inserts data into CreditMemoLine table.


    'UPGRADE_CHECK_TODO
    'Previously checked for dead code - leaving for now
    Sub InsertIntoQBCreditMemoLineComment()

        'This routine inserts data into CreditMemoLine table.
        'The vars are Public vars declared at top of module
        'The vars are filled in individual modules such as AR1CreditMemo()

        'On Error GoTo SubError

        'ALPHABETICAL LISTING OF ALL TABLES
        'http://www.qodbc.com/docs/html/qodbc/20/tables/table_info_all_us.asp
        '
        'QUICKBOOKS VIEW:  InvoiceLine
        'http://www.qodbc.com/docs/html/qodbc/20/tables/qbview_d_invoice_line.asp?qbviewd_id=38
        '
        'TABLE DETAIL REFERENCE:   InvoiceLine
        'http://www.qodbc.com/docs/html/qodbc/20/tables/table_detail_us.asp?details_id=38&tn_us=TRUE


        'Build the SQL string
        Dim strSQL1 As String = "INSERT INTO CreditMemoLine " & Environment.NewLine & _
                                "   ( CustomerRefListID " & Environment.NewLine & _
                                "   , CustomerRefFullName " & Environment.NewLine & _
                                "   , ARAccountRefListID " & Environment.NewLine & _
                                "   , ARAccountRefFullName " & Environment.NewLine & _
                                "   , TemplateRefListID " & Environment.NewLine & _
                                "   , TemplateRefFullName " & Environment.NewLine & _
                                "   , TxnDate " & Environment.NewLine & _
                                "   , RefNumber " & Environment.NewLine & _
                                "   , BillAddressAddr1 " & Environment.NewLine & _
                                "   , BillAddressAddr2 " & Environment.NewLine & _
                                "   , BillAddressAddr3 " & Environment.NewLine & _
                                "   , BillAddressAddr4 " & Environment.NewLine & _
                                "   , BillAddressCity " & Environment.NewLine & _
                                "   , BillAddressState " & Environment.NewLine & _
                                "   , BillAddressPostalCode " & Environment.NewLine & _
                                "   , BillAddressCountry " & Environment.NewLine
        Dim strSQL2 As String = "   , ShipAddressAddr1 " & Environment.NewLine & _
                                "   , ShipAddressAddr2 " & Environment.NewLine & _
                                "   , ShipAddressAddr3 " & Environment.NewLine & _
                                "   , ShipAddressAddr4 " & Environment.NewLine & _
                                "   , ShipAddressCity " & Environment.NewLine & _
                                "   , ShipAddressState " & Environment.NewLine & _
                                "   , ShipAddressPostalCode " & Environment.NewLine & _
                                "   , ShipAddressCountry " & Environment.NewLine & _
                                "   , IsPending " & Environment.NewLine & _
                                "   , PONumber " & Environment.NewLine & _
                                "   , TermsRefListID " & Environment.NewLine & _
                                "   , TermsRefFullName " & Environment.NewLine & _
                                "   , SalesRepRefListID " & Environment.NewLine & _
                                "   , SalesRepRefFullName " & Environment.NewLine & _
                                "   , FOB " & Environment.NewLine & _
                                "   , ShipDate " & Environment.NewLine & _
                                "   , ItemSalesTaxRefListID " & Environment.NewLine & _
                                "   , ItemSalesTaxRefFullName " & Environment.NewLine
        Dim strSQL3 As String = "   , Memo " & Environment.NewLine & _
                                "   , IsToBePrinted " & Environment.NewLine & _
                                "   , CustomerSalesTaxCodeRefListID " & Environment.NewLine & _
                                "   , CustomerSalesTaxCodeRefFullName " & Environment.NewLine & _
                                "   , CreditMemoLineDesc " & Environment.NewLine & _
                                "   , FQSaveToCache ) " & Environment.NewLine

        Dim strSQL4 As String = "VALUES " & Environment.NewLine & _
                                "   ( '" & strQBCMLine_CustomerRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CustomerRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ARAccountRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ARAccountRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_TemplateRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_TemplateRefFullName & "'  " & Environment.NewLine & _
                                "   , {d'" & strQBCMLine_TxnDate & "'}  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_RefNumber & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressAddr1 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressAddr2 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressAddr3 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressAddr4 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressCity & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressState & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressPostalCode & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressCountry & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressAddr1 & "'  " & Environment.NewLine
        Dim strSQL5 As String = "   , '" & strQBCMLine_ShipAddressAddr2 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressAddr3 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressAddr4 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressCity & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressState & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressPostalCode & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressCountry & "'  " & Environment.NewLine & _
                                "   , " & strQBCMLine_IsPending & "  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_PONumber & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_TermsRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_TermsRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_SalesRepRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_SalesRepRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_FOB & "'  " & Environment.NewLine & _
                                "   , {d'" & strQBCMLine_ShipDate & "'}  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ItemSalesTaxRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ItemSalesTaxRefFullName & "'  " & Environment.NewLine
        Dim strSQL6 As String = "   , '" & strQBCMLine_Memo & "'  " & Environment.NewLine & _
                                "   , " & strQBCMLine_IsToBePrinted & "  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CustomerSalesTaxCodeRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CustomerSalesTaxCodeRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CreditMemoLineDesc & "'  " & Environment.NewLine & _
                                "   , " & strQBCMLine_FQSaveToCache & " ) " & Environment.NewLine

        'Combine the strings
        Dim strTableInsert As String = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
        'Debug.Print strTableInsert

        'force then ^
        strQBCMLine_CreditMemoLineDesc = "-"

        'Execute the insert



        Dim TempCommand As ODBCCommand
        TempCommand = cnQuickBooks.CreateCommand()
        TempCommand.CommandText = strTableInsert
        TempCommand.ExecuteNonQuery()

        'Record the insert
        'cnlogic.Execute "INSERT INTO AAQBInsertRecord () VALUES (now,strQBCMLine_FOB)"
        'DateTime AuditTrail QBFile TypeTrans=SingARMultiGLCreditMemo

        'Increase the global QB insert counter
        gintQBInsertCounter += 1

        Exit Sub


        MessageBox.Show("<<InsertIntoQBCreditMemoLineComment>> " & Information.Err().Description, Application.ProductName)

    End Sub

    Sub InsertIntoQBCreditMemoLineConversion()

        'This routine inserts data into CreditMemoLine table.
        'The vars are Public vars declared at top of module
        'The vars are filled in individual modules such as AR1CreditMemo()

        'On Error GoTo SubError

        'ALPHABETICAL LISTING OF ALL TABLES
        'http://www.qodbc.com/docs/html/qodbc/20/tables/table_info_all_us.asp
        '
        'QUICKBOOKS VIEW:  InvoiceLine
        'http://www.qodbc.com/docs/html/qodbc/20/tables/qbview_d_invoice_line.asp?qbviewd_id=38
        '
        'TABLE DETAIL REFERENCE:   InvoiceLine
        'http://www.qodbc.com/docs/html/qodbc/20/tables/table_detail_us.asp?details_id=38&tn_us=TRUE



        'Build the SQL string
        Dim strSQL1 As String = "INSERT INTO CreditMemoLine " & Environment.NewLine & _
                                "   ( CustomerRefListID " & Environment.NewLine & _
                                "   , CustomerRefFullName " & Environment.NewLine & _
                                "   , ARAccountRefListID " & Environment.NewLine & _
                                "   , ARAccountRefFullName " & Environment.NewLine & _
                                "   , TemplateRefListID " & Environment.NewLine & _
                                "   , TemplateRefFullName " & Environment.NewLine & _
                                "   , TxnDate " & Environment.NewLine & _
                                "   , RefNumber " & Environment.NewLine & _
                                "   , BillAddressAddr1 " & Environment.NewLine & _
                                "   , BillAddressAddr2 " & Environment.NewLine & _
                                "   , BillAddressAddr3 " & Environment.NewLine & _
                                "   , BillAddressAddr4 " & Environment.NewLine & _
                                "   , BillAddressCity " & Environment.NewLine & _
                                "   , BillAddressState " & Environment.NewLine & _
                                "   , BillAddressPostalCode " & Environment.NewLine & _
                                "   , BillAddressCountry " & Environment.NewLine
        Dim strSQL2 As String = "   , ShipAddressAddr1 " & Environment.NewLine & _
                                "   , ShipAddressAddr2 " & Environment.NewLine & _
                                "   , ShipAddressAddr3 " & Environment.NewLine & _
                                "   , ShipAddressAddr4 " & Environment.NewLine & _
                                "   , ShipAddressCity " & Environment.NewLine & _
                                "   , ShipAddressState " & Environment.NewLine & _
                                "   , ShipAddressPostalCode " & Environment.NewLine & _
                                "   , ShipAddressCountry " & Environment.NewLine & _
                                "   , IsPending " & Environment.NewLine & _
                                "   , PONumber " & Environment.NewLine & _
                                "   , TermsRefListID " & Environment.NewLine & _
                                "   , TermsRefFullName " & Environment.NewLine & _
                                "   , SalesRepRefListID " & Environment.NewLine & _
                                "   , SalesRepRefFullName " & Environment.NewLine & _
                                "   , FOB " & Environment.NewLine & _
                                "   , ShipDate " & Environment.NewLine & _
                                "   , ItemSalesTaxRefListID " & Environment.NewLine & _
                                "   , ItemSalesTaxRefFullName " & Environment.NewLine
        Dim strSQL3 As String = "   , Memo " & Environment.NewLine & _
                                "   , IsToBePrinted " & Environment.NewLine & _
                                "   , CustomerSalesTaxCodeRefListID " & Environment.NewLine & _
                                "   , CustomerSalesTaxCodeRefFullName " & Environment.NewLine & _
                                "   , CreditMemoLineItemRefListID " & Environment.NewLine & _
                                "   , CreditMemoLineItemRefFullName " & Environment.NewLine & _
                                "   , CreditMemoLineDesc " & Environment.NewLine & _
                                "   , CreditMemoLineQuantity " & Environment.NewLine & _
                                "   , CreditMemoLineAmount " & Environment.NewLine & _
                                "   , CreditMemoLineSalesTaxCodeRefListID " & Environment.NewLine & _
                                "   , CreditMemoLineSalesTaxCodeRefFullName " & Environment.NewLine & _
                                "   , CreditMemoLineOverrideItemAccountRefListID " & Environment.NewLine & _
                                "   , CreditMemoLineOverrideItemAccountRefFullName " & Environment.NewLine & _
                                "   , FQSaveToCache ) " & Environment.NewLine

        Dim strSQL4 As String = "VALUES " & Environment.NewLine & _
                                "   ( '" & strQBCMLine_CustomerRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CustomerRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ARAccountRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ARAccountRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_TemplateRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_TemplateRefFullName & "'  " & Environment.NewLine & _
                                "   , {d'" & strQBCMLine_TxnDate & "'}  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_RefNumber & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressAddr1 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressAddr2 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressAddr3 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressAddr4 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressCity & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressState & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressPostalCode & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_BillAddressCountry & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressAddr1 & "'  " & Environment.NewLine
        Dim strSQL5 As String = "   , '" & strQBCMLine_ShipAddressAddr2 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressAddr3 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressAddr4 & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressCity & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressState & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressPostalCode & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ShipAddressCountry & "'  " & Environment.NewLine & _
                                "   , " & strQBCMLine_IsPending & "  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_PONumber & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_TermsRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_TermsRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_SalesRepRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_SalesRepRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_FOB & "'  " & Environment.NewLine & _
                                "   , {d'" & strQBCMLine_ShipDate & "'}  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ItemSalesTaxRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_ItemSalesTaxRefFullName & "'  " & Environment.NewLine
        Dim strSQL6 As String = "   , '" & strQBCMLine_Memo & "'  " & Environment.NewLine & _
                                "   , " & strQBCMLine_IsToBePrinted & "  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CustomerSalesTaxCodeRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CustomerSalesTaxCodeRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CreditMemoLineItemRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CreditMemoLineItemRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CreditMemoLineDesc & "'  " & Environment.NewLine & _
                                "   , " & strQBCMLine_CreditMemoLineQuantity & "  " & Environment.NewLine & _
                                "   , " & strQBCMLine_CreditMemoLineAmount & "  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CreditMemoLineSalesTaxCodeRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CreditMemoLineSalesTaxCodeRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CreditMemoLineOverrideItemAccountRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBCMLine_CreditMemoLineOverrideItemAccountRefFullName & "'  " & Environment.NewLine & _
                                "   , " & strQBCMLine_FQSaveToCache & " ) " & Environment.NewLine


        'Combine the strings
        Dim strTableInsert As String = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
        'Debug.Print strTableInsert

        'Execute the insert
        Dim TempCommand As ODBCCommand
        TempCommand = cnQuickBooks.CreateCommand()
        TempCommand.CommandText = strTableInsert
        TempCommand.ExecuteNonQuery()


        ''here's how to track the target amount:
        'gstrAuditTrailTargetAmount = CStr(CCur(gstrAuditTrailTargetAmount) + CCur(AMOUNTLOADED))
        gstrAuditTrailTargetAmount = CStr(CDec(gstrAuditTrailTargetAmount) + CDec(strQBCMLine_CreditMemoLineAmount))

        ''here's how to track the target amount:
        'gstrGLTargetAmount = CStr(CCur(gstrGLTargetAmount) + CCur(AMOUNTLOADED))
        gstrGLTargetAmount = CStr(CDec(gstrGLTargetAmount) + CDec(strQBCMLine_CreditMemoLineAmount))


        'Record the insert
        'cnlogic.Execute "INSERT INTO AAQBInsertRecord () VALUES (now,strQBCMLine_FOB)"
        'DateTime AuditTrail QBFile TypeTrans=SingARMultiGLCreditMemo

        'Increase the global QB insert counter
        gintQBInsertCounter += 1

        Exit Sub


        MessageBox.Show("<<InsertIntoQBCreditMemoLineConversion>> " & Information.Err().Description, Application.ProductName)

    End Sub


End Module