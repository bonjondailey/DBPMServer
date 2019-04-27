Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Windows.Forms

'**************************************
'******* 1st Code Review Complete *****
'**************************************



Module modQBInsertInvoiceLine

    'These are InvoiceLine insert vars for QBIL_ in modQBInsertInvoiceLine
    Public strQBIL_TxnID As String = ""
    Public strQBIL_TimeCreated As String = ""
    Public strQBIL_TimeModified As String = ""
    Public strQBIL_EditSequence As String = ""
    Public strQBIL_TxnNumber As String = ""
    Public strQBIL_CustomerRefListID As String = ""
    Public strQBIL_CustomerRefFullName As String = ""
    Public strQBIL_ClassRefListID As String = ""
    Public strQBIL_ClassRefFullName As String = ""
    Public strQBIL_ARAccountRefListID As String = ""
    Public strQBIL_ARAccountRefFullName As String = ""
    Public strQBIL_TemplateRefListID As String = ""
    Public strQBIL_TemplateRefFullName As String = ""
    Public strQBIL_TxnDate As String = ""
    Public strQBIL_TxnDateMacro As String = ""
    Public strQBIL_RefNumber As String = ""
    Public strQBIL_BillAddressAddr1 As String = ""
    Public strQBIL_BillAddressAddr2 As String = ""
    Public strQBIL_BillAddressAddr3 As String = ""
    Public strQBIL_BillAddressAddr4 As String = ""
    Public strQBIL_BillAddressCity As String = ""
    Public strQBIL_BillAddressState As String = ""
    Public strQBIL_BillAddressPostalCode As String = ""
    Public strQBIL_BillAddressCountry As String = ""
    Public strQBIL_ShipAddressAddr1 As String = ""
    Public strQBIL_ShipAddressAddr2 As String = ""
    Public strQBIL_ShipAddressAddr3 As String = ""
    Public strQBIL_ShipAddressAddr4 As String = ""
    Public strQBIL_ShipAddressCity As String = ""
    Public strQBIL_ShipAddressState As String = ""
    Public strQBIL_ShipAddressPostalCode As String = ""
    Public strQBIL_ShipAddressCountry As String = ""
    Public strQBIL_IsPending As String = ""
    Public strQBIL_IsFinanceCharge As String = ""
    Public strQBIL_PONumber As String = ""
    Public strQBIL_TermsRefListID As String = ""
    Public strQBIL_TermsRefFullName As String = ""
    Public strQBIL_DueDate As String = ""
    Public strQBIL_SalesRepRefListID As String = ""
    Public strQBIL_SalesRepRefFullName As String = ""
    Public strQBIL_FOB As String = ""
    Public strQBIL_ShipDate As String = ""
    Public strQBIL_ShipMethodRefListID As String = ""
    Public strQBIL_ShipMethodRefFullName As String = ""
    Public strQBIL_Subtotal As String = ""
    Public strQBIL_ItemSalesTaxRefListID As String = ""
    Public strQBIL_ItemSalesTaxRefFullName As String = ""
    Public strQBIL_SalesTaxPercentage As String = ""
    Public strQBIL_SalesTaxTotal As String = ""
    Public strQBIL_AppliedAmount As String = ""
    Public strQBIL_BalanceRemaining As String = ""
    Public strQBIL_Memo As String = ""
    Public strQBIL_IsPaid As String = ""
    Public strQBIL_CustomerMsgRefListID As String = ""
    Public strQBIL_CustomerMsgRefFullName As String = ""
    Public strQBIL_IsToBePrinted As String = ""
    Public strQBIL_CustomerSalesTaxCodeRefListID As String = ""
    Public strQBIL_CustomerSalesTaxCodeRefFullName As String = ""
    Public strQBIL_SuggestedDiscountAmount As String = ""
    Public strQBIL_SuggestedDiscountDate As String = ""
    Public strQBIL_InvoiceLineType As String = ""
    Public strQBIL_InvoiceLineSeqNo As String = ""
    Public strQBIL_InvoiceLineGroupTxnLineID As String = ""
    Public strQBIL_InvoiceLineGroupItemGroupRefListID As String = ""
    Public strQBIL_InvoiceLineGroupItemGroupRefFullName As String = ""
    Public strQBIL_InvoiceLineGroupDesc As String = ""
    Public strQBIL_InvoiceLineGroupQuantity As String = ""
    Public strQBIL_InvoiceLineGroupIsPrintItemsInGroup As String = ""
    Public strQBIL_InvoiceLineGroupTotalAmount As String = ""
    Public strQBIL_InvoiceLineGroupSeqNo As String = ""
    Public strQBIL_InvoiceLineTxnLineID As String = ""
    Public strQBIL_InvoiceLineItemRefListID As String = ""
    Public strQBIL_InvoiceLineItemRefFullName As String = ""
    Public strQBIL_InvoiceLineDesc As String = ""
    Public strQBIL_InvoiceLineQuantity As String = ""
    Public strQBIL_InvoiceLineRate As String = ""
    Public strQBIL_InvoiceLineRatePercent As String = ""
    Public strQBIL_InvoiceLinePriceLevelRefListID As String = ""
    Public strQBIL_InvoiceLinePriceLevelRefFullName As String = ""
    Public strQBIL_InvoiceLineClassRefListID As String = ""
    Public strQBIL_InvoiceLineClassRefFullName As String = ""
    Public strQBIL_InvoiceLineAmount As String = ""
    Public strQBIL_InvoiceLineServiceDate As String = ""
    Public strQBIL_InvoiceLineSalesTaxCodeRefListID As String = ""
    Public strQBIL_InvoiceLineSalesTaxCodeRefFullName As String = ""
    Public strQBIL_InvoiceLineOverrideItemAccountRefListID As String = ""
    Public strQBIL_InvoiceLineOverrideItemAccountRefFullName As String = ""
    Public strQBIL_FQSaveToCache As String = ""
    Public strQBIL_CustomFieldInvoiceLineOther1 As String = ""
    Public strQBIL_CustomFieldInvoiceLineOther2 As String = ""
    Public strQBIL_CustomFieldInvoiceLineAreaLocation As String = ""
    Public strQBIL_CustomFieldInvoiceLineBinNumber As String = ""
    Public strQBIL_CustomFieldInvoiceLineQtyPriceBreak As String = ""
    Public strQBIL_CustomFieldInvoiceLineShelf As String = ""
    Public strQBIL_CustomFieldInvoiceLineUnitofMeasure As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupOther1 As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupOther2 As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupAreaLocation As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupBinNumber As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupQtyPriceBreak As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupShelf As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupUnitofMeasure As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupLineOther1 As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupLineOther2 As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupLineAreaLocation As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupLineBinNumber As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupLineQtyPriceBreak As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupLineShelf As String = ""
    Public strQBIL_CustomFieldInvoiceLineGroupLineUnitofMeasure As String = ""
    Public strQBIL_CustomFieldOther As String = ""
    Public strQBIL_CustomFieldShipAccount As String = ""
    Public strQBIL_CustomFieldShipMethod As String = ""
    Public strQBIL_CustomFieldShipVendor As String = ""

    'Create the global QB insert counter
    Public gintQBInsertCounter As Integer

    Sub InsertIntoQBInvoiceLineComment()

       
        'Build the SQL string
        Dim strSQL1 As String = "INSERT INTO InvoiceLine " & Environment.NewLine & _
                                "   ( CustomerRefListID " & Environment.NewLine & _
                                "   , ARAccountRefListID " & Environment.NewLine & _
                                "   , ARAccountRefFullName " & Environment.NewLine & _
                                "   , TemplateRefListID " & Environment.NewLine & _
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
                                "   , ShipMethodRefListID " & Environment.NewLine & _
                                "   , ShipMethodRefFullName " & Environment.NewLine & _
                                "   , ItemSalesTaxRefListID " & Environment.NewLine & _
                                "   , ItemSalesTaxRefFullName " & Environment.NewLine
        Dim strSQL3 As String = "   , Memo " & Environment.NewLine & _
                                "   , IsToBePrinted " & Environment.NewLine & _
                                "   , CustomerSalesTaxCodeRefListID " & Environment.NewLine & _
                                "   , CustomerSalesTaxCodeRefFullName " & Environment.NewLine & _
                                "   , InvoiceLineDesc " & Environment.NewLine & _
                                "   , FQSaveToCache ) " & Environment.NewLine

        Dim strSQL4 As String = "VALUES " & Environment.NewLine & _
                                "   ( '" & strQBIL_CustomerRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ARAccountRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ARAccountRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_TemplateRefListID & "'  " & Environment.NewLine & _
                                "   , {d'" & strQBIL_TxnDate & "'}  " & Environment.NewLine & _
                                "   , '" & strQBIL_RefNumber & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_BillAddressAddr1 & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_BillAddressAddr2 & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_BillAddressAddr3 & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_BillAddressAddr4 & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_BillAddressCity & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_BillAddressState & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_BillAddressPostalCode & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_BillAddressCountry & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ShipAddressAddr1 & "'  " & Environment.NewLine
        Dim strSQL5 As String = "   , '" & strQBIL_ShipAddressAddr2 & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ShipAddressAddr3 & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ShipAddressAddr4 & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ShipAddressCity & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ShipAddressState & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ShipAddressPostalCode & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ShipAddressCountry & "'  " & Environment.NewLine & _
                                "   , " & strQBIL_IsPending & "  " & Environment.NewLine & _
                                "   , '" & strQBIL_PONumber & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_TermsRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_TermsRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_SalesRepRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_SalesRepRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_FOB & "'  " & Environment.NewLine & _
                                "   , {d'" & strQBIL_ShipDate & "'}  " & Environment.NewLine & _
                                "   , '" & strQBIL_ShipMethodRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ShipMethodRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ItemSalesTaxRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_ItemSalesTaxRefFullName & "'  " & Environment.NewLine
        Dim strSQL6 As String = "   , '" & strQBIL_Memo & "'  " & Environment.NewLine & _
                                "   , " & strQBIL_IsToBePrinted & "  " & Environment.NewLine & _
                                "   , '" & strQBIL_CustomerSalesTaxCodeRefListID & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_CustomerSalesTaxCodeRefFullName & "'  " & Environment.NewLine & _
                                "   , '" & strQBIL_InvoiceLineDesc & "'  " & Environment.NewLine & _
                                "   , " & strQBIL_FQSaveToCache & " ) " & Environment.NewLine


        'Combine the strings
        Dim strTableInsert As String = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
        'Debug.Print strTableInsert
        siteConstants.ShowUserMessage("InsertIntoQBInvoiceLineComment", strTableInsert, "Insert Invoice Line Comment")
        'force then ^
        strQBIL_InvoiceLineDesc = "-"

        'Execute the insert

        Try
            Dim invRecordsInserted As Integer = 0

            Dim TempCommand As ODBCCommand
            TempCommand = cnQuickBooks.CreateCommand()
            TempCommand.CommandText = strTableInsert
            invRecordsInserted = TempCommand.ExecuteNonQuery()
            Debug.WriteLine(strTableInsert)

            'Record the insert
            'cnlogic.Execute "INSERT INTO AAQBInsertRecord () VALUES (now,strQBIL_FOB)"
            'DateTime AuditTrail QBFile TypeTrans=SingARMultiGLInvoice

            'Increase the global QB insert counter
            gintQBInsertCounter += 1

            'if invRecordsInserted = 0 in the InvoiceLog table, then there is an issue with going into QB
            Dim strSQL As String = "UPDATE InvoiceLog SET recCount = recCount + " & invRecordsInserted & " WHERE JobNumber = '" & strQBIL_RefNumber & "'"
            Dim TempCommand_2 As SqlCommand
            TempCommand_2 = cnDBPM.CreateCommand()
            TempCommand_2.CommandText = strSQL
            TempCommand_2.ExecuteNonQuery()

        Catch exc As System.Exception
            HaveError("modQBInsertInvoiceLine", "InsertIntoQBInvoiceLineComment", CStr(Information.Err().Number), exc.Message & strTableInsert, Information.Err().Source, "", "")
            Exit Try
        End Try

    End Sub



    Public Function InsertIntoQBInvoiceLineJobs() As Boolean
        Dim strSQL As String
        Dim retVal As Boolean
        'This routine inserts data into InvoiceLine table.
        'The vars are Public vars declared at top of module
        'The vars are filled in individual modules such as AR1Invoice()
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6 As String
        Dim strTableInsert As String = ""

        Try


            'Build the SQL string
            strSQL1 = "INSERT INTO InvoiceLine " & Environment.NewLine & _
                      "   ( CustomerRefListID " & Environment.NewLine & _
                      "   , ARAccountRefListID " & Environment.NewLine & _
                      "   , ARAccountRefFullName " & Environment.NewLine & _
                      "   , TemplateRefListID " & Environment.NewLine & _
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
            strSQL2 = "   , ShipAddressAddr1 " & Environment.NewLine & _
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
                      "   , ShipMethodRefListID " & Environment.NewLine & _
                      "   , ShipMethodRefFullName " & Environment.NewLine & _
                      "   , ItemSalesTaxRefListID " & Environment.NewLine & _
                      "   , ItemSalesTaxRefFullName " & Environment.NewLine
            strSQL3 = "   , Memo " & Environment.NewLine & _
                      "   , IsToBePrinted " & Environment.NewLine & _
                      "   , CustomerSalesTaxCodeRefListID " & Environment.NewLine & _
                      "   , CustomerSalesTaxCodeRefFullName " & Environment.NewLine & _
                      "   , InvoiceLineItemRefListID " & Environment.NewLine & _
                      "   , InvoiceLineItemRefFullName " & Environment.NewLine & _
                      "   , InvoiceLineDesc " & Environment.NewLine & _
                      "   , InvoiceLineQuantity " & Environment.NewLine & _
                      "   , InvoiceLineRate " & Environment.NewLine & _
                      "   , InvoiceLineAmount " & Environment.NewLine & _
                      "   , InvoiceLineSalesTaxCodeRefListID " & Environment.NewLine & _
                      "   , InvoiceLineSalesTaxCodeRefFullName " & Environment.NewLine & _
                      "   , CustomFieldInvoiceLineOther1, FQSaveToCache, CustomerMsgRefListID ) " & Environment.NewLine

            strSQL4 = "VALUES " & Environment.NewLine & _
                      "   ( '" & strQBIL_CustomerRefListID & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ARAccountRefListID & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ARAccountRefFullName & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_TemplateRefListID & "'  " & Environment.NewLine & _
                      "   , {d'" & strQBIL_TxnDate & "'}  " & Environment.NewLine & _
                      "   , '" & strQBIL_RefNumber & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_BillAddressAddr1 & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_BillAddressAddr2 & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_BillAddressAddr3 & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_BillAddressAddr4 & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_BillAddressCity & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_BillAddressState & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_BillAddressPostalCode & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_BillAddressCountry & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ShipAddressAddr1 & "'  " & Environment.NewLine
            strSQL5 = "   , '" & strQBIL_ShipAddressAddr2 & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ShipAddressAddr3 & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ShipAddressAddr4 & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ShipAddressCity & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ShipAddressState & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ShipAddressPostalCode & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ShipAddressCountry & "'  " & Environment.NewLine & _
                      "   , " & strQBIL_IsPending & "  " & Environment.NewLine & _
                      "   , '" & strQBIL_PONumber & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_TermsRefListID & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_TermsRefFullName & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_SalesRepRefListID & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_SalesRepRefFullName & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_FOB & "'  " & Environment.NewLine & _
                      "   , {d'" & strQBIL_ShipDate & "'}  " & Environment.NewLine & _
                      "   , '" & strQBIL_ShipMethodRefListID & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ShipMethodRefFullName & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ItemSalesTaxRefListID & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_ItemSalesTaxRefFullName & "'  " & Environment.NewLine
            strSQL6 = "   , '" & strQBIL_Memo & "'  " & Environment.NewLine & _
                      "   , " & strQBIL_IsToBePrinted & "  " & Environment.NewLine & _
                      "   , '" & strQBIL_CustomerSalesTaxCodeRefListID & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_CustomerSalesTaxCodeRefFullName & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_InvoiceLineItemRefListID & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_InvoiceLineItemRefFullName & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_InvoiceLineDesc & "'  " & Environment.NewLine & _
                      "   , " & strQBIL_InvoiceLineQuantity & "  " & Environment.NewLine & _
                      "   , " & strQBIL_InvoiceLineRate & "  " & Environment.NewLine & _
                      "   , " & strQBIL_InvoiceLineAmount & "  " & Environment.NewLine & _
                      "   , '" & strQBIL_InvoiceLineSalesTaxCodeRefListID & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_InvoiceLineSalesTaxCodeRefFullName & "'  " & Environment.NewLine & _
                      "   , '" & strQBIL_CustomFieldInvoiceLineOther1 & "'," & strQBIL_FQSaveToCache & ", '80000021-1380317840')"


            'Combine the strings
            strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
            siteConstants.ShowUserMessage("InsertIntoQBInvoiceLineJobs", strTableInsert)
            Debug.WriteLine(strTableInsert)


            Dim iInvoicesInserted As Integer
            iInvoicesInserted = 0
            Dim TempCommand As OdbcCommand
            TempCommand = cnQuickBooks.CreateCommand()
            TempCommand.CommandText = strTableInsert
            iInvoicesInserted = TempCommand.ExecuteNonQuery()


            gintQBInsertCounter += 1
            siteConstants.ShowUserMessage("Inserting Invoice Lines Into QB", "Total Inserted So Far: " & gintQBInsertCounter.ToString)
            retVal = True

        Catch excep As System.Exception
            Dim myErrorMsg As String = ""
            myErrorMsg = ""
            myErrorMsg = excep.Message & " ------ " & strTableInsert
            If Strings.Len(myErrorMsg) >= 5000 Then
                myErrorMsg = Strings.Left(myErrorMsg, 4999)
            End If

            strSQL = "INSERT INTO InvoiceLogItem (JobNumber, ErrorMsg) VALUES ('" & strQBIL_RefNumber & "', '" & myErrorMsg & "')"
            reportInvoiceLineError(strSQL)

            HaveError("modQBInsertInvoiceLine", "InsertIntoQBInvoiceLineJobs", CStr(Information.Err().Number), excep.Message, Information.Err().Source, "", "")
            retVal = False
            Exit Try
        End Try

        Return retVal
    End Function

    Public Sub reportInvoiceLineError(sql As String)
       
        Try
            Using objSQL As New SQLHelper()
                objSQL.ExecuteSQL(sql)
            End Using

        Catch ex As Exception

        End Try
    End Sub

End Module