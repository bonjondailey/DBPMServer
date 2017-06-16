Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Windows.Forms
Imports UpgradeHelpers.Gui
Imports UpgradeHelpers.Helpers
Imports DBPM_Server.siteConstants
Imports DBPM_Server.SQLHelper

Module modDBPM_ReloadQBTables
	'To restore the See Processing checkbox after heavy processing
	Public intRestoreSeeProcessing As CheckState

    'THE FIRST SUB SEEMS TO BE THE ONLY ONE CALLED IN THIS MOD
    '**********************************
    '*** FIRST CODE REVIEW COMPLETE ***
    '**********************************


   

	Public Sub ReloadQBCustomerTableOnly()

		'Permission and ErrorHandling          (Auto built)
		Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
		Dim strSubName As String = "ReloadQBCustomerTableOnly" '"SUBNAME"

		'Check permission to run
		If Not HavePermission(strObjName, strSubName) Then Exit Sub

		'Error handling
		If gbooUseErrorHandling Then On Error GoTo ErrorFunc
		GoTo RunCode
ErrorFunc:
		If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then Resume Next Else Exit Sub
RunCode:





		'If gstrCompany = "DrummondPrinting" Then
		'    'Stop 'do you want to reload the qb tables?   NOTE: Causes DBPM Reports & CreditCheck to not work
		'ElseIf gstrCompany = "FrazzledAndBedazzled" Then
		'    Exit Sub
		'End If


		'exit if (paused, running, ...
		If frmMain.DefInstance.chkPauseProcessing.CheckState = CheckState.Checked Then Exit Sub
		If booQBRefreshInProgress Then
			MessageBox.Show("A Process Is Already Running", Application.ProductName) : Exit Sub
		End If


		'Save the state of the See Processing checkbox so it can be restored
		intRestoreSeeProcessing = frmMain.DefInstance.chkSeeProcessing.CheckState
		'Disable the See Processing checkbox for the heavy processing ahead
		frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Unchecked


		'Set flag
		booQBRefreshInProgress = True


		'Open QuickBooks  -good
		If Not booQBFileIsOpen Then

			'Open the QuickBooks file!
			If Not (cnQuickBooks.State = ConnectionState.Open) Then
				OpenConnectionQB()
			Else
				booQBFileIsOpen = True
			End If


		End If




		'Show what's processing in the listbox
		Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   Running ReloadQBCustomerTableOnly")
		'frmMain.lstConversionProgress.AddItem "" & Now & "   Running ReloadQBCustomerTableOnly"
		frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Running ReloadQBCustomerTableOnly"
		frmMain.DefInstance.lblStatus.Text = "Starting ReloadQBCustomerTableOnly"
		Application.DoEvents()

        ReloadQB_Customer()
		
		Dim TempCommand As SqlCommand
		TempCommand = cnMax.CreateCommand()
		TempCommand.CommandText = "exec sp_TEMP_MarkCustPromoRush"
		TempCommand.ExecuteNonQuery()


		'show finished
		frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Finished ReloadQBCustomerTableOnly"
		frmMain.DefInstance.lblStatus.Text = "Finished ReloadQBCustomerTableOnly"
		Application.DoEvents()


		'Reset flag
		booQBRefreshInProgress = False


		'Restore the state of the See Processing checkbox
		frmMain.DefInstance.chkSeeProcessing.CheckState = intRestoreSeeProcessing


	End Sub

	Public Sub ReloadQBTables()
        'THIS SUB DOES NOTHING - NOT USED ANYMORE 


		'Permission and ErrorHandling          (Auto built)
		Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
		Dim strSubName As String = "ReloadQBTables" '"SUBNAME"

		'Check permission to run
		If Not HavePermission(strObjName, strSubName) Then Exit Sub

		'Error handling
		


		'exit if (paused, running, ...
		If frmMain.DefInstance.chkPauseProcessing.CheckState = CheckState.Checked Then Exit Sub
		'If booQBRefreshInProgress = True Then MsgBox "A Process Is Already Running": Exit Sub
		If booQBRefreshInProgress Then Exit Sub



		'Save the state of the See Processing checkbox so it can be restored
		intRestoreSeeProcessing = frmMain.DefInstance.chkSeeProcessing.CheckState
		'Disable the See Processing checkbox for the heavy processing ahead
		frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Unchecked



		'Set flag
		booQBRefreshInProgress = True



		'Close the QuickBooks file



		'**********************************************************
		'******** CLOSE QB CONNECTION FOR REFRESH ????? ***********
		'**********************************************************

		'CloseConnectionQB





		'
		''Close QuickBooks if still open  'Kill process if necessary    'QBW32.exe
		''if xx then
		'   KillQuickBooks
		''endif
		'




		'Delete Optimizer file?     *****  'Did I negate the need for this by turning off adAsyncFetch on each module? -no

		'**********************************************************
		'************* DELETE OPTIMIZER FILE????? *****************
		'**********************************************************

		'DeleteQBOptimizerFile   'But first make a backup of the file?

		'**********************************************************
		'************* DELETE OPTIMIZER FILE????? *****************
		'**********************************************************


		GoTo SkipPoop



		'Backup the QB file
        'BackupQBFile()



		'
		'Start QuickBooks
		'Call Shell("""C:\Program Files\Intuit\QuickBooks Enterprise Solutions 6.0\QBW32Enterprise.exe""", vbMinimizedNoFocus)
		''Call Shell("""C:\Program Files\Intuit\QuickBooks Enterprise Solutions 6.0\QBW32EnterpriseAccountant.exe""", vbMinimizedNoFocus)
		'



		'Need a delay?
		Dim strNowPlus2 As String = ""
		strNowPlus2 = CStr(CDbl(DateTime.Now.ToString("HHMMss")) + 3)
		Do Until CInt(DateTime.Now.ToString("HHMMss")) > CInt(strNowPlus2)
			Application.DoEvents()
		Loop 
		Application.DoEvents()


SkipPoop:

		'If cnQuickBooks.State <> 1 Then
		OpenConnectionQB()
		'End If
		If cnDBPM.State <> ConnectionState.Open Then
			OpenConnectionDBPM()
		End If
		If cnMax.State <> ConnectionState.Open Then
			OpenConnectionMax()
		End If


		'Open QuickBooks  -good
		'If booQBFileIsOpen = False Then
		'
		'    'Open the QuickBooks file!
		'    'OpenConnectionQB
		'    If Not cnQuickBooks.State = 1 Then
		'        OpenConnectionQB
		'    Else
		'        booQBFileIsOpen = True
		'    End If
		'
		'End If



		'QB_Terms    'didnt work

		'QB_SalesRep
		'QB_ItemOtherCharge
		'QB_Item
		'QB_ShipMethod
		'
		'QB_PaymentMethod
		'QB_Account
		'QB_Vendor
		'
		'QB_StandardTerms
		'
		'
		'QB_Template
		'
		'QB_PurchaseOrderLine
		'QB_CreditMemoLine



		'Show what's processing in the listbox
		Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   Running ReloadQBTables")
		'frmMain.lstConversionProgress.AddItem "" & Now & "   Running ReloadQBTables"
		frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Running ReloadQBTables"
		frmMain.DefInstance.lblStatus.Text = "Starting ReloadQBTables"
		Application.DoEvents()


		'GetQBMaxTimeModified   'not needed for reload



		''OLD:
		''ReloadQB_Terms     'didnt work   'ERROR: Data provider or other service returned an E_FAIL status.?
		''ReloadQB_ItemOtherCharge        'ERROR: Data provider or other service returned an E_FAIL status.
		'
		''Stop
		'ReloadQB_ShipMethod     'Good
		'ReloadQB_PaymentMethod  'Good   'Needed?
		''ReloadQB_Account       'Error converting data type varchar to numeric.   'Needed?
		''ReloadQB_Vendor        'Error converting data type varchar to numeric.   'Needed?
		'ReloadQB_SalesRep
		''If gstrCompany <> "FrazzledAndBedazzled" Then
		'If gstrComputerName <> "EDSDELLXPP" Then
		'    ReloadQB_StandardTerms  'Good
		'    'ReloadQB_Item           'Good  'broke in 9.0 so far
		'End If
		'
		''CUSTOMER!
		'ReloadQB_Customer
		'
		'If frmMain.chkReloadCustOnly.Value = 0 Then
		'
		'    ReloadQB_Invoice
		'    ReloadQB_InvoiceLine    'broke in 9.0 so far?
		'
		'    'If gstrCompany <> "FrazzledAndBedazzled" Then
		'    If gstrComputerName <> "EDSDELLXPP" Then
		'        ReloadQB_ReceivePayment        'broke in 9.0 so far
		'        ReloadQB_ReceivePaymentLine    'broke in 9.0 so far
		'        ReloadQB_CreditMemo
		'        ReloadQB_CreditMemoLine
		'    End If
		'
		'End If




		''BYPASS

		'NEW
		''''''''''ReloadQB_Terms         'didnt work   'ERROR: Data provider or other service returned an E_FAIL status.?
		''''''''''ReloadQB_ItemOtherCharge        'ERROR: Data provider or other service returned an E_FAIL status.

		'UNCOMMENTJB ReloadQB_ShipMethod     'Good
		'UNCOMMENTJB ReloadQB_PaymentMethod


		'''''''''''ReloadQB_Account       'Error converting data type varchar to numeric.   'Needed?
		'''''''''''ReloadQB_Vendor        'Error converting data type varchar to numeric.   'Needed?



		'UNCOMMENTJB ReloadQB_SalesRep
		'UNCOMMENTJB ReloadQB_StandardTerms  'Good
		'UNCOMMENTJB ReloadQB_Item           'Good  'broke in 9.0 so far   'broke in 10.0 too   'fixed
		'UNCOMMENTJB ReloadQB_ItemOtherCharge  'Good  'broke in 9.0 so far   'broke in 10.0 too   'fixed




		'UNCOMMENTJB If Now > CDate("01/14/2010") Then ReloadQB_Customer

        'UPGRADE_CHECK_TODO - following was not commented (If, End If, ReloadQB_InvoiceLine())
        'If frmMain.DefInstance.chkReloadCustOnly.CheckState = CheckState.Unchecked Then

        '	'UNCOMMENTJB ReloadQB_CreditMemo

        '	'UNCOMMENTJB ReloadQB_Invoice
        '	'UNCOMMENTJB ReloadQB_ReceivePayment

        '	ReloadQB_InvoiceLine()
        '	'UNCOMMENTJB ReloadQB_ReceivePaymentLine

        '	'UNCOMMENTJB ReloadQB_CreditMemoLine

        'End If





		'InsertMaxBillToIntoQB       '*******   Where put?

		'CreditMemoLine?
		'SalesRep
		'ItemOtherCharge
		'


		'RefreshQB_InvoiceLine
		'*cnDBPM.Execute "exec sp_TEMP_MarkCustPromoRush"
		'UNCOMMENTJB cnMax.Execute "exec sp_TEMP_MarkCustPromoRush"


		'show finished
		frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Finished ReloadQBTables"
		frmMain.DefInstance.lblStatus.Text = "Finished ReloadQBTables"
		Application.DoEvents()


		'Reset flag
		booQBRefreshInProgress = False


		'Restore the state of the See Processing checkbox
		frmMain.DefInstance.chkSeeProcessing.CheckState = intRestoreSeeProcessing

	End Sub




	Public Sub ReloadQB_Customer()
		Dim rs1MaxOfCopy_QB_Customer, rs3TestID_QB_Customer As Object

		'Permission and ErrorHandling          (Auto built)
		Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
		Dim strSubName As String = "ReloadQB_Customer" '"SUBNAME"

		'Check permission to run
		If Not HavePermission(strObjName, strSubName) Then Exit Sub

		'Error handling
		If gbooUseErrorHandling Then
			'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
			UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
		End If





		'FOR PART 2SrcQB_ - Get records from QB_Customer
		Debug.WriteLine("List2SrcQB_QB_Customer")
		Dim rs2SrcQB_QB_Customer As DataSet
		Dim str2SrcQB_QB_CustomerSQL, str2SrcQB_QB_CustomerRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_FullName, str2SrcQB_IsActive, str2SrcQB_ParentRefListID, str2SrcQB_ParentRefFullName, str2SrcQB_Sublevel, str2SrcQB_CompanyName, str2SrcQB_Salutation, str2SrcQB_FirstName, str2SrcQB_MiddleName, str2SrcQB_LastName, str2SrcQB_BillAddressAddr1, str2SrcQB_BillAddressAddr2, str2SrcQB_BillAddressAddr3, str2SrcQB_BillAddressAddr4, str2SrcQB_BillAddressCity, str2SrcQB_BillAddressState, str2SrcQB_BillAddressPostalCode, str2SrcQB_BillAddressCountry, str2SrcQB_ShipAddressAddr1, str2SrcQB_ShipAddressAddr2, str2SrcQB_ShipAddressAddr3, str2SrcQB_ShipAddressAddr4, str2SrcQB_ShipAddressCity, str2SrcQB_ShipAddressState, str2SrcQB_ShipAddressPostalCode, str2SrcQB_ShipAddressCountry, str2SrcQB_Phone, str2SrcQB_AltPhone, str2SrcQB_Fax, str2SrcQB_Email, str2SrcQB_Contact, str2SrcQB_AltContact, str2SrcQB_CustomerTypeRefListID, str2SrcQB_CustomerTypeRefFullName, str2SrcQB_TermsRefListID, str2SrcQB_TermsRefFullName, str2SrcQB_SalesRepRefListID, str2SrcQB_SalesRepRefFullName, str2SrcQB_Balance, str2SrcQB_TotalBalance, str2SrcQB_OpenBalance, str2SrcQB_OpenBalanceDate, str2SrcQB_SalesTaxCodeRefListID, str2SrcQB_SalesTaxCodeRefFullName, str2SrcQB_ItemSalesTaxRefListID, str2SrcQB_ItemSalesTaxRefFullName, str2SrcQB_ResaleNumber, str2SrcQB_AccountNumber, str2SrcQB_CreditLimit, str2SrcQB_PreferredPaymentMethodRefListID, str2SrcQB_PreferredPaymentMethodRefFullName, str2SrcQB_CreditCardInfoCreditCardNumber, str2SrcQB_CreditCardInfoExpirationMonth, str2SrcQB_CreditCardInfoExpirationYear, str2SrcQB_CreditCardInfoNameOnCard, str2SrcQB_CreditCardInfoCreditCardAddress, str2SrcQB_CreditCardInfoCreditCardPostalCode, str2SrcQB_JobStatus, str2SrcQB_JobStartDate, str2SrcQB_JobProjectedEndDate, str2SrcQB_JobEndDate, str2SrcQB_JobDesc, str2SrcQB_JobTypeRefListID, str2SrcQB_JobTypeRefFullName, str2SrcQB_Notes, str2SrcQB_PriceLevelRefListID, str2SrcQB_PriceLevelRefFullName, str2SrcQB_CustomFieldOther As String
		'This routine gets the 2SrcQB_QB_Customer from the database according to the selection in str2SrcQB_QB_CustomerSQL.
		'It then puts those 2SrcQB_QB_Customer in the list box

		''FOR PART 3TestID_ - Get records from QB_Customer
		'Debug.Print "List3TestID_QB_Customer"
		'Dim rs3TestID_QB_Customer As ADODB.Recordset
		'Dim str3TestID_QB_CustomerSQL As String
		'Dim str3TestID_QB_CustomerRow As String
		'Dim str3TestID_ListID As String
		''This routine gets the 3TestID_QB_Customer from the database according to the selection in str3TestID_QB_CustomerSQL.
		''It then puts those 3TestID_QB_Customer in the list box

		'dim SQL strings
		Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


		'On Error GoTo SubError

		'frmMain.lstConversionProgress.Clear

		'Show what's processing
		frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Reloading SQL QuickBooks Mirror -Customer  "
		frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_Customer"
		Application.DoEvents()




		'Get rs from QB
		'Load table from rs

		'PART 2SrcQB_: Get the new records from Actual QB
		'Get a recordset of records from QB for QB_Customer
		rs2SrcQB_QB_Customer = New DataSet()
		str2SrcQB_QB_CustomerSQL = "SELECT * FROM Customer"
		str2SrcQB_QB_CustomerSQL = "SELECT * FROM Customer"
		str2SrcQB_QB_CustomerSQL = "SELECT * FROM Customer"
		str2SrcQB_QB_CustomerSQL = "SELECT * FROM Customer"
		'Debug.Print str2SrcQB_QB_CustomerSQL
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_CustomerSQL, cnQuickBooks)
		rs2SrcQB_QB_Customer.Tables.Clear()
		adap.Fill(rs2SrcQB_QB_Customer) ', adAsyncFetch
		If rs2SrcQB_QB_Customer.Tables(0).Rows.Count > 0 Then

			'Clear out table
			If gstrCompany = "DrummondPrinting" Then
				'*cnDBPM.Execute "DELETE FROM QB_Customer"
				Dim TempCommand As SqlCommand
				TempCommand = cnMax.CreateCommand()
				TempCommand.CommandText = "DELETE FROM QB_Customer"
				TempCommand.ExecuteNonQuery()
			ElseIf gstrCompany = "FrazzledAndBedazzled" Then 
				Dim TempCommand_2 As SqlCommand
				TempCommand_2 = cnMax.CreateCommand()
				TempCommand_2.CommandText = "DELETE FROM QB_Customer"
				TempCommand_2.ExecuteNonQuery()
			End If

			'Show what's processing in the listbox
			frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_Customer.Tables(0).Rows.Count) & "  QB_Customer  Records  ")

			For	Each iteration_row As DataRow In rs2SrcQB_QB_Customer.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Customer.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_Customer.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_Customer.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_ListID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_Name = ""
                str2SrcQB_FullName = ""
                str2SrcQB_IsActive = "1"
                str2SrcQB_ParentRefListID = ""
                str2SrcQB_ParentRefFullName = ""
                str2SrcQB_Sublevel = "0"
                str2SrcQB_CompanyName = ""
                str2SrcQB_Salutation = ""
                str2SrcQB_FirstName = ""
                str2SrcQB_MiddleName = ""
                str2SrcQB_LastName = ""
                str2SrcQB_BillAddressAddr1 = ""
                str2SrcQB_BillAddressAddr2 = ""
                str2SrcQB_BillAddressAddr3 = ""
                str2SrcQB_BillAddressAddr4 = ""
                str2SrcQB_BillAddressCity = ""
                str2SrcQB_BillAddressState = ""
                str2SrcQB_BillAddressPostalCode = ""
                str2SrcQB_BillAddressCountry = ""
                str2SrcQB_ShipAddressAddr1 = ""
                str2SrcQB_ShipAddressAddr2 = ""
                str2SrcQB_ShipAddressAddr3 = ""
                str2SrcQB_ShipAddressAddr4 = ""
                str2SrcQB_ShipAddressCity = ""
                str2SrcQB_ShipAddressState = ""
                str2SrcQB_ShipAddressPostalCode = ""
                str2SrcQB_ShipAddressCountry = ""
                str2SrcQB_Phone = ""
                str2SrcQB_AltPhone = ""
                str2SrcQB_Fax = ""
                str2SrcQB_Email = ""
                str2SrcQB_Contact = ""
                str2SrcQB_AltContact = ""
                str2SrcQB_CustomerTypeRefListID = ""
                str2SrcQB_CustomerTypeRefFullName = ""
                str2SrcQB_TermsRefListID = ""
                str2SrcQB_TermsRefFullName = ""
                str2SrcQB_SalesRepRefListID = ""
                str2SrcQB_SalesRepRefFullName = ""
                str2SrcQB_Balance = "0"
                str2SrcQB_TotalBalance = "0"
                str2SrcQB_OpenBalance = "0"
                str2SrcQB_OpenBalanceDate = ""
                str2SrcQB_SalesTaxCodeRefListID = ""
                str2SrcQB_SalesTaxCodeRefFullName = ""
                str2SrcQB_ItemSalesTaxRefListID = ""
                str2SrcQB_ItemSalesTaxRefFullName = ""
                str2SrcQB_ResaleNumber = ""
                str2SrcQB_AccountNumber = ""
                str2SrcQB_CreditLimit = "0"
                str2SrcQB_PreferredPaymentMethodRefListID = ""
                str2SrcQB_PreferredPaymentMethodRefFullName = ""
                str2SrcQB_CreditCardInfoCreditCardNumber = ""
                str2SrcQB_CreditCardInfoExpirationMonth = "0"
                str2SrcQB_CreditCardInfoExpirationYear = "0"
                str2SrcQB_CreditCardInfoNameOnCard = ""
                str2SrcQB_CreditCardInfoCreditCardAddress = ""
                str2SrcQB_CreditCardInfoCreditCardPostalCode = ""
                str2SrcQB_JobStatus = ""
                str2SrcQB_JobStartDate = ""
                str2SrcQB_JobProjectedEndDate = ""
                str2SrcQB_JobEndDate = ""
                str2SrcQB_JobDesc = ""
                str2SrcQB_JobTypeRefListID = ""
                str2SrcQB_JobTypeRefFullName = ""
                str2SrcQB_Notes = ""
                str2SrcQB_PriceLevelRefListID = ""
                str2SrcQB_PriceLevelRefFullName = ""
                str2SrcQB_CustomFieldOther = ""

                'get the columns from the database
                If iteration_row("ListID") <> "" Then str2SrcQB_ListID = iteration_row("ListID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("Name") <> "" Then str2SrcQB_Name = iteration_row("Name")
                If iteration_row("FullName") <> "" Then str2SrcQB_FullName = iteration_row("FullName")
                If iteration_row("IsActive") <> "" Then str2SrcQB_IsActive = iteration_row("IsActive")
                If iteration_row("ParentRefListID") <> "" Then str2SrcQB_ParentRefListID = iteration_row("ParentRefListID")
                If iteration_row("ParentRefFullName") <> "" Then str2SrcQB_ParentRefFullName = iteration_row("ParentRefFullName")
                If iteration_row("Sublevel") <> "" Then str2SrcQB_Sublevel = iteration_row("Sublevel")
                If iteration_row("CompanyName") <> "" Then str2SrcQB_CompanyName = iteration_row("CompanyName")
                If iteration_row("Salutation") <> "" Then str2SrcQB_Salutation = iteration_row("Salutation")
                If iteration_row("FirstName") <> "" Then str2SrcQB_FirstName = iteration_row("FirstName")
                If iteration_row("MiddleName") <> "" Then str2SrcQB_MiddleName = iteration_row("MiddleName")
                If iteration_row("LastName") <> "" Then str2SrcQB_LastName = iteration_row("LastName")
                If iteration_row("BillAddressAddr1") <> "" Then str2SrcQB_BillAddressAddr1 = iteration_row("BillAddressAddr1")
                If iteration_row("BillAddressAddr2") <> "" Then str2SrcQB_BillAddressAddr2 = iteration_row("BillAddressAddr2")
                If iteration_row("BillAddressAddr3") <> "" Then str2SrcQB_BillAddressAddr3 = iteration_row("BillAddressAddr3")
                If iteration_row("BillAddressAddr4") <> "" Then str2SrcQB_BillAddressAddr4 = iteration_row("BillAddressAddr4")
                If iteration_row("BillAddressCity") <> "" Then str2SrcQB_BillAddressCity = iteration_row("BillAddressCity")
                If iteration_row("BillAddressState") <> "" Then str2SrcQB_BillAddressState = iteration_row("BillAddressState")
                If iteration_row("BillAddressPostalCode") <> "" Then str2SrcQB_BillAddressPostalCode = iteration_row("BillAddressPostalCode")
                If iteration_row("BillAddressCountry") <> "" Then str2SrcQB_BillAddressCountry = iteration_row("BillAddressCountry")
                If iteration_row("ShipAddressAddr1") <> "" Then str2SrcQB_ShipAddressAddr1 = iteration_row("ShipAddressAddr1")
                If iteration_row("ShipAddressAddr2") <> "" Then str2SrcQB_ShipAddressAddr2 = iteration_row("ShipAddressAddr2")
                If iteration_row("ShipAddressAddr3") <> "" Then str2SrcQB_ShipAddressAddr3 = iteration_row("ShipAddressAddr3")
                If iteration_row("ShipAddressAddr4") <> "" Then str2SrcQB_ShipAddressAddr4 = iteration_row("ShipAddressAddr4")
                If iteration_row("ShipAddressCity") <> "" Then str2SrcQB_ShipAddressCity = iteration_row("ShipAddressCity")
                If iteration_row("ShipAddressState") <> "" Then str2SrcQB_ShipAddressState = iteration_row("ShipAddressState")
                If iteration_row("ShipAddressPostalCode") <> "" Then str2SrcQB_ShipAddressPostalCode = iteration_row("ShipAddressPostalCode")
                If iteration_row("ShipAddressCountry") <> "" Then str2SrcQB_ShipAddressCountry = iteration_row("ShipAddressCountry")
                If iteration_row("Phone") <> "" Then str2SrcQB_Phone = iteration_row("Phone")
                If iteration_row("AltPhone") <> "" Then str2SrcQB_AltPhone = iteration_row("AltPhone")
                If iteration_row("Fax") <> "" Then str2SrcQB_Fax = iteration_row("Fax")
                If iteration_row("Email") <> "" Then str2SrcQB_Email = iteration_row("Email")
                If iteration_row("Contact") <> "" Then str2SrcQB_Contact = iteration_row("Contact")
                If iteration_row("AltContact") <> "" Then str2SrcQB_AltContact = iteration_row("AltContact")
                If iteration_row("CustomerTypeRefListID") <> "" Then str2SrcQB_CustomerTypeRefListID = iteration_row("CustomerTypeRefListID")
                If iteration_row("CustomerTypeRefFullName") <> "" Then str2SrcQB_CustomerTypeRefFullName = iteration_row("CustomerTypeRefFullName")
                If iteration_row("TermsRefListID") <> "" Then str2SrcQB_TermsRefListID = iteration_row("TermsRefListID")
                If iteration_row("TermsRefFullName") <> "" Then str2SrcQB_TermsRefFullName = iteration_row("TermsRefFullName")
                If iteration_row("SalesRepRefListID") <> "" Then str2SrcQB_SalesRepRefListID = iteration_row("SalesRepRefListID")
                If iteration_row("SalesRepRefFullName") <> "" Then str2SrcQB_SalesRepRefFullName = iteration_row("SalesRepRefFullName")
                If iteration_row("Balance") <> "" Then str2SrcQB_Balance = iteration_row("Balance")
                If iteration_row("TotalBalance") <> "" Then str2SrcQB_TotalBalance = iteration_row("TotalBalance")
                If iteration_row("OpenBalance") <> "" Then str2SrcQB_OpenBalance = iteration_row("OpenBalance")
                If iteration_row("OpenBalanceDate") <> "" Then str2SrcQB_OpenBalanceDate = iteration_row("OpenBalanceDate")
                If iteration_row("SalesTaxCodeRefListID") <> "" Then str2SrcQB_SalesTaxCodeRefListID = iteration_row("SalesTaxCodeRefListID")
                If iteration_row("SalesTaxCodeRefFullName") <> "" Then str2SrcQB_SalesTaxCodeRefFullName = iteration_row("SalesTaxCodeRefFullName")
                If iteration_row("ItemSalesTaxRefListID") <> "" Then str2SrcQB_ItemSalesTaxRefListID = iteration_row("ItemSalesTaxRefListID")
                If iteration_row("ItemSalesTaxRefFullName") <> "" Then str2SrcQB_ItemSalesTaxRefFullName = iteration_row("ItemSalesTaxRefFullName")
                If iteration_row("ResaleNumber") <> "" Then str2SrcQB_ResaleNumber = iteration_row("ResaleNumber")
                If iteration_row("AccountNumber") <> "" Then str2SrcQB_AccountNumber = iteration_row("AccountNumber")
                If iteration_row("CreditLimit") <> "" Then str2SrcQB_CreditLimit = iteration_row("CreditLimit")
                If iteration_row("PreferredPaymentMethodRefListID") <> "" Then str2SrcQB_PreferredPaymentMethodRefListID = iteration_row("PreferredPaymentMethodRefListID")
                If iteration_row("PreferredPaymentMethodRefFullName") <> "" Then str2SrcQB_PreferredPaymentMethodRefFullName = iteration_row("PreferredPaymentMethodRefFullName")
                If iteration_row("CreditCardInfoCreditCardNumber") <> "" Then str2SrcQB_CreditCardInfoCreditCardNumber = iteration_row("CreditCardInfoCreditCardNumber")
                If iteration_row("CreditCardInfoExpirationMonth") <> "" Then str2SrcQB_CreditCardInfoExpirationMonth = iteration_row("CreditCardInfoExpirationMonth")
                If iteration_row("CreditCardInfoExpirationYear") <> "" Then str2SrcQB_CreditCardInfoExpirationYear = iteration_row("CreditCardInfoExpirationYear")
                If iteration_row("CreditCardInfoNameOnCard") <> "" Then str2SrcQB_CreditCardInfoNameOnCard = iteration_row("CreditCardInfoNameOnCard")
                If iteration_row("CreditCardInfoCreditCardAddress") <> "" Then str2SrcQB_CreditCardInfoCreditCardAddress = iteration_row("CreditCardInfoCreditCardAddress")
                If iteration_row("CreditCardInfoCreditCardPostalCode") <> "" Then str2SrcQB_CreditCardInfoCreditCardPostalCode = iteration_row("CreditCardInfoCreditCardPostalCode")
                If iteration_row("JobStatus") <> "" Then str2SrcQB_JobStatus = iteration_row("JobStatus")
                If iteration_row("JobStartDate") <> "" Then str2SrcQB_JobStartDate = iteration_row("JobStartDate")
                If iteration_row("JobProjectedEndDate") <> "" Then str2SrcQB_JobProjectedEndDate = iteration_row("JobProjectedEndDate")
                If iteration_row("JobEndDate") <> "" Then str2SrcQB_JobEndDate = iteration_row("JobEndDate")
                If iteration_row("JobDesc") <> "" Then str2SrcQB_JobDesc = iteration_row("JobDesc")
                If iteration_row("JobTypeRefListID") <> "" Then str2SrcQB_JobTypeRefListID = iteration_row("JobTypeRefListID")
                If iteration_row("JobTypeRefFullName") <> "" Then str2SrcQB_JobTypeRefFullName = iteration_row("JobTypeRefFullName")
                If iteration_row("Notes") <> "" Then str2SrcQB_Notes = iteration_row("Notes")
                If iteration_row("PriceLevelRefListID") <> "" Then str2SrcQB_PriceLevelRefListID = iteration_row("PriceLevelRefListID")
                If iteration_row("PriceLevelRefFullName") <> "" Then str2SrcQB_PriceLevelRefFullName = iteration_row("PriceLevelRefFullName")
                '        If rs2SrcQB_QB_Customer!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_Customer!CustomFieldOther

                'Strip quote character out of strings
                'Get quote characters out!
                'Change Quote to reverse quote
                'If KeyAscii = 39 Then KeyAscii = 96
                str2SrcQB_ListID = str2SrcQB_ListID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_Name = str2SrcQB_Name.Replace("'"c, "`"c)
                str2SrcQB_FullName = str2SrcQB_FullName.Replace("'"c, "`"c)
                str2SrcQB_IsActive = str2SrcQB_IsActive.Replace("'"c, "`"c)
                str2SrcQB_ParentRefListID = str2SrcQB_ParentRefListID.Replace("'"c, "`"c)
                str2SrcQB_ParentRefFullName = str2SrcQB_ParentRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Sublevel = str2SrcQB_Sublevel.Replace("'"c, "`"c)
                str2SrcQB_CompanyName = str2SrcQB_CompanyName.Replace("'"c, "`"c)
                str2SrcQB_Salutation = str2SrcQB_Salutation.Replace("'"c, "`"c)
                str2SrcQB_FirstName = str2SrcQB_FirstName.Replace("'"c, "`"c)
                str2SrcQB_MiddleName = str2SrcQB_MiddleName.Replace("'"c, "`"c)
                str2SrcQB_LastName = str2SrcQB_LastName.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr1 = str2SrcQB_BillAddressAddr1.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr2 = str2SrcQB_BillAddressAddr2.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr3 = str2SrcQB_BillAddressAddr3.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr4 = str2SrcQB_BillAddressAddr4.Replace("'"c, "`"c)
                str2SrcQB_BillAddressCity = str2SrcQB_BillAddressCity.Replace("'"c, "`"c)
                str2SrcQB_BillAddressState = str2SrcQB_BillAddressState.Replace("'"c, "`"c)
                str2SrcQB_BillAddressPostalCode = str2SrcQB_BillAddressPostalCode.Replace("'"c, "`"c)
                str2SrcQB_BillAddressCountry = str2SrcQB_BillAddressCountry.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr1 = str2SrcQB_ShipAddressAddr1.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr2 = str2SrcQB_ShipAddressAddr2.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr3 = str2SrcQB_ShipAddressAddr3.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr4 = str2SrcQB_ShipAddressAddr4.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressCity = str2SrcQB_ShipAddressCity.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressState = str2SrcQB_ShipAddressState.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressPostalCode = str2SrcQB_ShipAddressPostalCode.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressCountry = str2SrcQB_ShipAddressCountry.Replace("'"c, "`"c)
                str2SrcQB_Phone = str2SrcQB_Phone.Replace("'"c, "`"c)
                str2SrcQB_AltPhone = str2SrcQB_AltPhone.Replace("'"c, "`"c)
                str2SrcQB_Fax = str2SrcQB_Fax.Replace("'"c, "`"c)
                str2SrcQB_Email = str2SrcQB_Email.Replace("'"c, "`"c)
                str2SrcQB_Contact = str2SrcQB_Contact.Replace("'"c, "`"c)
                str2SrcQB_AltContact = str2SrcQB_AltContact.Replace("'"c, "`"c)
                str2SrcQB_CustomerTypeRefListID = str2SrcQB_CustomerTypeRefListID.Replace("'"c, "`"c)
                str2SrcQB_CustomerTypeRefFullName = str2SrcQB_CustomerTypeRefFullName.Replace("'"c, "`"c)
                str2SrcQB_TermsRefListID = str2SrcQB_TermsRefListID.Replace("'"c, "`"c)
                str2SrcQB_TermsRefFullName = str2SrcQB_TermsRefFullName.Replace("'"c, "`"c)
                str2SrcQB_SalesRepRefListID = str2SrcQB_SalesRepRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesRepRefFullName = str2SrcQB_SalesRepRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Balance = str2SrcQB_Balance.Replace("'"c, "`"c)
                str2SrcQB_TotalBalance = str2SrcQB_TotalBalance.Replace("'"c, "`"c)
                str2SrcQB_OpenBalance = str2SrcQB_OpenBalance.Replace("'"c, "`"c)
                str2SrcQB_OpenBalanceDate = str2SrcQB_OpenBalanceDate.Replace("'"c, "`"c)
                str2SrcQB_SalesTaxCodeRefListID = str2SrcQB_SalesTaxCodeRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesTaxCodeRefFullName = str2SrcQB_SalesTaxCodeRefFullName.Replace("'"c, "`"c)
                str2SrcQB_ItemSalesTaxRefListID = str2SrcQB_ItemSalesTaxRefListID.Replace("'"c, "`"c)
                str2SrcQB_ItemSalesTaxRefFullName = str2SrcQB_ItemSalesTaxRefFullName.Replace("'"c, "`"c)
                str2SrcQB_ResaleNumber = str2SrcQB_ResaleNumber.Replace("'"c, "`"c)
                str2SrcQB_AccountNumber = str2SrcQB_AccountNumber.Replace("'"c, "`"c)
                str2SrcQB_CreditLimit = str2SrcQB_CreditLimit.Replace("'"c, "`"c)
                str2SrcQB_PreferredPaymentMethodRefListID = str2SrcQB_PreferredPaymentMethodRefListID.Replace("'"c, "`"c)
                str2SrcQB_PreferredPaymentMethodRefFullName = str2SrcQB_PreferredPaymentMethodRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditCardInfoCreditCardNumber = str2SrcQB_CreditCardInfoCreditCardNumber.Replace("'"c, "`"c)
                str2SrcQB_CreditCardInfoExpirationMonth = str2SrcQB_CreditCardInfoExpirationMonth.Replace("'"c, "`"c)
                str2SrcQB_CreditCardInfoExpirationYear = str2SrcQB_CreditCardInfoExpirationYear.Replace("'"c, "`"c)
                str2SrcQB_CreditCardInfoNameOnCard = str2SrcQB_CreditCardInfoNameOnCard.Replace("'"c, "`"c)
                str2SrcQB_CreditCardInfoCreditCardAddress = str2SrcQB_CreditCardInfoCreditCardAddress.Replace("'"c, "`"c)
                str2SrcQB_CreditCardInfoCreditCardPostalCode = str2SrcQB_CreditCardInfoCreditCardPostalCode.Replace("'"c, "`"c)
                str2SrcQB_JobStatus = str2SrcQB_JobStatus.Replace("'"c, "`"c)
                str2SrcQB_JobStartDate = str2SrcQB_JobStartDate.Replace("'"c, "`"c)
                str2SrcQB_JobProjectedEndDate = str2SrcQB_JobProjectedEndDate.Replace("'"c, "`"c)
                str2SrcQB_JobEndDate = str2SrcQB_JobEndDate.Replace("'"c, "`"c)
                str2SrcQB_JobDesc = str2SrcQB_JobDesc.Replace("'"c, "`"c)
                str2SrcQB_JobTypeRefListID = str2SrcQB_JobTypeRefListID.Replace("'"c, "`"c)
                str2SrcQB_JobTypeRefFullName = str2SrcQB_JobTypeRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Notes = str2SrcQB_Notes.Replace("'"c, "`"c)
                str2SrcQB_PriceLevelRefListID = str2SrcQB_PriceLevelRefListID.Replace("'"c, "`"c)
                str2SrcQB_PriceLevelRefFullName = str2SrcQB_PriceLevelRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldOther = str2SrcQB_CustomFieldOther.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")

                '        'Change IsActive flag back to binary
                '        If str2SrcQB_IsActive = "True" Then str2SrcQB_IsActive = "1" Else str2SrcQB_IsActive = "0"
                '
                '        'Change IsActive flag back to binary
                '        If str2SrcQB_IsActive = "True" Then
                '            str2SrcQB_IsActive = "1"
                '        Else
                '            str2SrcQB_IsActive = "0"
                '        End If


                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_CustomerRow = "" & _
                                           Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_TimeCreated & "                  ", 16) & "   " & _
                                           Strings.Left(str2SrcQB_TimeModified & "                  ", 16) & "   " & _
                                           Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_Name & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_FullName & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_ParentRefListID & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_ParentRefFullName & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_Sublevel & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_CompanyName & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_Salutation & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_FirstName & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_MiddleName & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_LastName & "                  ", 18) & "   " & _
                                           "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Customer.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_Customer.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_Customer.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_CustomerRow)
                    'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                '        'New recordset
                '        Set rs3TestID_QB_Customer = New ADODB.Recordset
                '        str3TestID_QB_CustomerSQL = "SELECT TOP 100 * FROM QB_Customer"
                '        str3TestID_QB_CustomerSQL = "SELECT ListID FROM QB_Customer WHERE ListID = '" & str2SrcQB_ListID & "'"
                '        Debug.Print str3TestID_QB_CustomerSQL
                '        'rs3TestID_QB_Customer.Open str3TestID_QB_CustomerSQL, cnDBPM, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        rs3TestID_QB_Customer.Open str3TestID_QB_CustomerSQL, cnmax, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        If rs3TestID_QB_Customer.RecordCount > 1 Then Stop 'Should only be one
                '        If rs3TestID_QB_Customer.RecordCount > 0 Then  'record exists  -UPDATE
                '            'DO UPDATE WORK:
                '            Debug.Print "UPDATE"
                '
                '            'Build the SQL string
                '            'MODIFICATION REQUIRED HERE:
                '            '1)Comment the whole mess
                '            '2)Use extra strings if continuation is over 25 lines
                '            '3)Take line continuations off of last lines
                '            '4)Change squiggle to quote
                '            '5)Change  to nothing
                '            '6)Uncomment the whole mess
                '            '7)Put parenthesis into the INSERT & VALUES statements
                '            '8)Delete these MODIFICATION REQUIRED lines
                '            strSQL1 = "UPDATE  " & vbCrLf & _
                ''                      "       QB_Customer " & vbCrLf & _
                ''                      "SET " & vbCrLf & _
                ''                      "       -- ListID = '" & str2SrcQB_ListID & "'" & vbCrLf & _
                ''                      "       TimeCreated = '" & str2SrcQB_TimeCreated & "'" & vbCrLf & _
                ''                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & vbCrLf & _
                ''                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & vbCrLf & _
                ''                      "     , Name = '" & str2SrcQB_Name & "'" & vbCrLf & _
                ''                      "     , FullName = '" & str2SrcQB_FullName & "'" & vbCrLf & _
                ''                      "     , IsActive = '" & str2SrcQB_IsActive & "'" & vbCrLf & _
                ''                      "     , ParentRefListID = '" & str2SrcQB_ParentRefListID & "'" & vbCrLf & _
                ''                      "     , ParentRefFullName = '" & str2SrcQB_ParentRefFullName & "'" & vbCrLf & _
                ''                      "     , Sublevel = '" & str2SrcQB_Sublevel & "'" & vbCrLf & _
                ''                      "     , CompanyName = '" & str2SrcQB_CompanyName & "'" & vbCrLf & _
                ''                      "     , Salutation = '" & str2SrcQB_Salutation & "'" & vbCrLf & _
                ''                      "     , FirstName = '" & str2SrcQB_FirstName & "'" & vbCrLf & _
                ''                      "     , MiddleName = '" & str2SrcQB_MiddleName & "'" & vbCrLf & _
                ''                      "     , LastName = '" & str2SrcQB_LastName & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr1 = '" & str2SrcQB_BillAddressAddr1 & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr2 = '" & str2SrcQB_BillAddressAddr2 & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr3 = '" & str2SrcQB_BillAddressAddr3 & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr4 = '" & str2SrcQB_BillAddressAddr4 & "'" & vbCrLf & _
                ''                      "     , BillAddressCity = '" & str2SrcQB_BillAddressCity & "'" & vbCrLf & _
                ''                      "     , BillAddressState = '" & str2SrcQB_BillAddressState & "'" & vbCrLf & _
                ''                      "     , BillAddressPostalCode = '" & str2SrcQB_BillAddressPostalCode & "'" & vbCrLf
                '            strSQL2 = "     , BillAddressCountry = '" & str2SrcQB_BillAddressCountry & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr1 = '" & str2SrcQB_ShipAddressAddr1 & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr2 = '" & str2SrcQB_ShipAddressAddr2 & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr3 = '" & str2SrcQB_ShipAddressAddr3 & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr4 = '" & str2SrcQB_ShipAddressAddr4 & "'" & vbCrLf & _
                ''                      "     , ShipAddressCity = '" & str2SrcQB_ShipAddressCity & "'" & vbCrLf & _
                ''                      "     , ShipAddressState = '" & str2SrcQB_ShipAddressState & "'" & vbCrLf & _
                ''                      "     , ShipAddressPostalCode = '" & str2SrcQB_ShipAddressPostalCode & "'" & vbCrLf & _
                ''                      "     , ShipAddressCountry = '" & str2SrcQB_ShipAddressCountry & "'" & vbCrLf & _
                ''                      "     , Phone = '" & str2SrcQB_Phone & "'" & vbCrLf & _
                ''                      "     , AltPhone = '" & str2SrcQB_AltPhone & "'" & vbCrLf & _
                ''                      "     , Fax = '" & str2SrcQB_Fax & "'" & vbCrLf & _
                ''                      "     , Email = '" & str2SrcQB_Email & "'" & vbCrLf & _
                ''                      "     , Contact = '" & str2SrcQB_Contact & "'" & vbCrLf & _
                ''                      "     , AltContact = '" & str2SrcQB_AltContact & "'" & vbCrLf & _
                ''                      "     , CustomerTypeRefListID = '" & str2SrcQB_CustomerTypeRefListID & "'" & vbCrLf & _
                ''                      "     , CustomerTypeRefFullName = '" & str2SrcQB_CustomerTypeRefFullName & "'" & vbCrLf & _
                ''                      "     , TermsRefListID = '" & str2SrcQB_TermsRefListID & "'" & vbCrLf & _
                ''                      "     , TermsRefFullName = '" & str2SrcQB_TermsRefFullName & "'" & vbCrLf & _
                ''                      "     , SalesRepRefListID = '" & str2SrcQB_SalesRepRefListID & "'" & vbCrLf & _
                ''                      "     , SalesRepRefFullName = '" & str2SrcQB_SalesRepRefFullName & "'" & vbCrLf & _
                ''                      "     , Balance = '" & str2SrcQB_Balance & "'" & vbCrLf & _
                ''                      "     , TotalBalance = '" & str2SrcQB_TotalBalance & "'" & vbCrLf & _
                ''                      "     , OpenBalance = '" & str2SrcQB_OpenBalance & "'" & vbCrLf & _
                ''                      "     , OpenBalanceDate = '" & str2SrcQB_OpenBalanceDate & "'" & vbCrLf
                '            strSQL3 = "     , SalesTaxCodeRefListID = '" & str2SrcQB_SalesTaxCodeRefListID & "'" & vbCrLf & _
                ''                      "     , SalesTaxCodeRefFullName = '" & str2SrcQB_SalesTaxCodeRefFullName & "'" & vbCrLf & _
                ''                      "     , ItemSalesTaxRefListID = '" & str2SrcQB_ItemSalesTaxRefListID & "'" & vbCrLf & _
                ''                      "     , ItemSalesTaxRefFullName = '" & str2SrcQB_ItemSalesTaxRefFullName & "'" & vbCrLf & _
                ''                      "     , ResaleNumber = '" & str2SrcQB_ResaleNumber & "'" & vbCrLf & _
                ''                      "     , AccountNumber = '" & str2SrcQB_AccountNumber & "'" & vbCrLf & _
                ''                      "     , CreditLimit = '" & str2SrcQB_CreditLimit & "'" & vbCrLf & _
                ''                      "     , PreferredPaymentMethodRefListID = '" & str2SrcQB_PreferredPaymentMethodRefListID & "'" & vbCrLf & _
                ''                      "     , PreferredPaymentMethodRefFullName = '" & str2SrcQB_PreferredPaymentMethodRefFullName & "'" & vbCrLf & _
                ''                      "     , CreditCardInfoCreditCardNumber = '" & str2SrcQB_CreditCardInfoCreditCardNumber & "'" & vbCrLf & _
                ''                      "     , CreditCardInfoExpirationMonth = '" & str2SrcQB_CreditCardInfoExpirationMonth & "'" & vbCrLf & _
                ''                      "     , CreditCardInfoExpirationYear = '" & str2SrcQB_CreditCardInfoExpirationYear & "'" & vbCrLf & _
                ''                      "     , CreditCardInfoNameOnCard = '" & str2SrcQB_CreditCardInfoNameOnCard & "'" & vbCrLf & _
                ''                      "     , CreditCardInfoCreditCardAddress = '" & str2SrcQB_CreditCardInfoCreditCardAddress & "'" & vbCrLf & _
                ''                      "     , CreditCardInfoCreditCardPostalCode = '" & str2SrcQB_CreditCardInfoCreditCardPostalCode & "'" & vbCrLf & _
                ''                      "     , JobStatus = '" & str2SrcQB_JobStatus & "'" & vbCrLf & _
                ''                      "     , JobStartDate = '" & str2SrcQB_JobStartDate & "'" & vbCrLf & _
                ''                      "     , JobProjectedEndDate = '" & str2SrcQB_JobProjectedEndDate & "'" & vbCrLf & _
                ''                      "     , JobEndDate = '" & str2SrcQB_JobEndDate & "'" & vbCrLf & _
                ''                      "     , JobDesc = '" & str2SrcQB_JobDesc & "'" & vbCrLf & _
                ''                      "     , JobTypeRefListID = '" & str2SrcQB_JobTypeRefListID & "'" & vbCrLf & _
                ''                      "     , JobTypeRefFullName = '" & str2SrcQB_JobTypeRefFullName & "'" & vbCrLf & _
                ''                      "     , Notes = '" & str2SrcQB_Notes & "'" & vbCrLf
                '            strSQL4 = "     , PriceLevelRefListID = '" & str2SrcQB_PriceLevelRefListID & "'" & vbCrLf & _
                ''                      "     , PriceLevelRefFullName = '" & str2SrcQB_PriceLevelRefFullName & "'" & vbCrLf & _
                ''                      "     , CustomFieldOther = '" & str2SrcQB_CustomFieldOther & "'" & vbCrLf & _
                ''                      "WHERE " & vbCrLf & _
                ''                      "       ListID = '" & str2SrcQB_ListID & "'" & vbCrLf
                '
                '            'Combine the strings
                '            strTableUpdate = strSQL1 & strSQL2 & strSQL3 & strSQL4 '& strSQL5 & strSQL6
                '            'Debug.Print strTableUpdate
                '
                '            'Execute the insert
                '            '*cnDBPM.Execute strTableUpdate
                '            cnmax.Execute strTableUpdate
                '
                '
                '
                '        Else 'record not exist  -INSERT
                '            'DO INSERT WORK:
                '            Debug.Print "INSERT"

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_Customer " & Environment.NewLine & _
                          "   ( ListID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , Name " & Environment.NewLine & _
                          "   , FullName " & Environment.NewLine & _
                          "   , IsActive " & Environment.NewLine & _
                          "   , ParentRefListID " & Environment.NewLine & _
                          "   , ParentRefFullName " & Environment.NewLine & _
                          "   , Sublevel " & Environment.NewLine & _
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
                          "   , CustomerTypeRefListID " & Environment.NewLine & _
                          "   , CustomerTypeRefFullName " & Environment.NewLine & _
                          "   , TermsRefListID " & Environment.NewLine & _
                          "   , TermsRefFullName " & Environment.NewLine & _
                          "   , SalesRepRefListID " & Environment.NewLine & _
                          "   , SalesRepRefFullName " & Environment.NewLine & _
                          "   , Balance " & Environment.NewLine & _
                          "   , TotalBalance " & Environment.NewLine & _
                          "   , OpenBalance " & Environment.NewLine & _
                          "   , OpenBalanceDate " & Environment.NewLine & _
                          "   , SalesTaxCodeRefListID " & Environment.NewLine & _
                          "   , SalesTaxCodeRefFullName " & Environment.NewLine
                strSQL3 = "   , ItemSalesTaxRefListID " & Environment.NewLine & _
                          "   , ItemSalesTaxRefFullName " & Environment.NewLine & _
                          "   , ResaleNumber " & Environment.NewLine & _
                          "   , AccountNumber " & Environment.NewLine & _
                          "   , CreditLimit " & Environment.NewLine & _
                          "   , PreferredPaymentMethodRefListID " & Environment.NewLine & _
                          "   , PreferredPaymentMethodRefFullName " & Environment.NewLine & _
                          "   , CreditCardInfoCreditCardNumber " & Environment.NewLine & _
                          "   , CreditCardInfoExpirationMonth " & Environment.NewLine & _
                          "   , CreditCardInfoExpirationYear " & Environment.NewLine & _
                          "   , CreditCardInfoNameOnCard " & Environment.NewLine & _
                          "   , CreditCardInfoCreditCardAddress " & Environment.NewLine & _
                          "   , CreditCardInfoCreditCardPostalCode " & Environment.NewLine & _
                          "   , JobStatus " & Environment.NewLine & _
                          "   , JobStartDate " & Environment.NewLine & _
                          "   , JobProjectedEndDate " & Environment.NewLine & _
                          "   , JobEndDate " & Environment.NewLine & _
                          "   , JobDesc " & Environment.NewLine & _
                          "   , JobTypeRefListID " & Environment.NewLine & _
                          "   , JobTypeRefFullName " & Environment.NewLine & _
                          "   , Notes " & Environment.NewLine & _
                          "   , PriceLevelRefListID " & Environment.NewLine & _
                          "   , PriceLevelRefFullName " & Environment.NewLine & _
                          "   , CustomFieldOther )" & Environment.NewLine
                strSQL4 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_ListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_Name & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_FullName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsActive & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ParentRefListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ParentRefFullName & "'  " & Environment.NewLine & _
                          "   , " & str2SrcQB_Sublevel & "  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_CompanyName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_Salutation & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_FirstName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_MiddleName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_LastName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr1 & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr2 & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr3 & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr4 & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressCity & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressState & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressPostalCode & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressCountry & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr1 & "'  " & Environment.NewLine
                strSQL5 = "   , '" & str2SrcQB_ShipAddressAddr2 & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr3 & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr4 & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressCity & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressState & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressPostalCode & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressCountry & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_Phone & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_AltPhone & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_Fax & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_Email & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_Contact & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_AltContact & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerTypeRefListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerTypeRefFullName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_TermsRefListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_TermsRefFullName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesRepRefListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesRepRefFullName & "'  " & Environment.NewLine & _
                          "   , " & str2SrcQB_Balance & "  " & Environment.NewLine & _
                          "   , " & str2SrcQB_TotalBalance & "  " & Environment.NewLine & _
                          "   , " & str2SrcQB_OpenBalance & "  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_OpenBalanceDate & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesTaxCodeRefListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesTaxCodeRefFullName & "'  " & Environment.NewLine
                strSQL6 = "   , '" & str2SrcQB_ItemSalesTaxRefListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ItemSalesTaxRefFullName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_ResaleNumber & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_AccountNumber & "'  " & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditLimit & "  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_PreferredPaymentMethodRefListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_PreferredPaymentMethodRefFullName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardInfoCreditCardNumber & "'  " & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardInfoExpirationMonth & "  " & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardInfoExpirationYear & "  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardInfoNameOnCard & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardInfoCreditCardAddress & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardInfoCreditCardPostalCode & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_JobStatus & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_JobStartDate & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_JobProjectedEndDate & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_JobEndDate & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_JobDesc & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_JobTypeRefListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_JobTypeRefFullName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_Notes & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_PriceLevelRefListID & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_PriceLevelRefFullName & "'  " & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldOther & "' ) " & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
                'Debug.Print strTableInsert

                'Execute the insert
                If gstrCompany = "DrummondPrinting" Then
                    '*cnDBPM.Execute strTableInsert
                    Dim TempCommand_3 As SqlCommand
                    TempCommand_3 = cnMax.CreateCommand()
                    TempCommand_3.CommandText = strTableInsert
                    TempCommand_3.ExecuteNonQuery()
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    Dim TempCommand_4 As SqlCommand
                    TempCommand_4 = cnMax.CreateCommand()
                    TempCommand_4.CommandText = strTableInsert
                    TempCommand_4.ExecuteNonQuery()
                End If
                '
                '        End If
                '

            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_Customer  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '        End If
        End If


        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QB_Customer.Close()
            rs1MaxOfCopy_QB_Customer = Nothing

            rs2SrcQB_QB_Customer = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QB_Customer.Close()
            rs3TestID_QB_Customer = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_Customer>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub

    Public Sub ReloadQB_ReceivePayment()
        Dim rs1MaxOfCopy_QBTable, rs3TestID_QBTable As Object
        Dim str2SrcQB_FQSaveToCache As String = ""

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_ReceivePayment" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:




        ''FOR PART 1MaxOfCopy_ - Get records from QBTable
        'Debug.Print "List1MaxOfCopy_QBTable"
        'Dim rs1MaxOfCopy_QBTable As ADODB.Recordset
        'Dim str1MaxOfCopy_QBTableSQL As String
        'Dim str1MaxOfCopy_QBTableRow As String
        'Dim str1MaxOfCopy_TimeModified As String
        ''This routine gets the 1MaxOfCopy_QBTable from the database according to the selection in str1MaxOfCopy_QBTableSQL.
        ''It then puts those 1MaxOfCopy_QBTable in the list box

        'FOR PART 2SrcQB_ - Get records from QB_ReceivePayment
        Debug.WriteLine("List2SrcQB_QB_ReceivePayment")
        Dim rs2SrcQB_QB_ReceivePayment As DataSet
        Dim str2SrcQB_QB_ReceivePaymentSQL, str2SrcQB_QB_ReceivePaymentRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_TotalAmount, str2SrcQB_PaymentMethodRefListID, str2SrcQB_PaymentMethodRefFullName, str2SrcQB_Memo, str2SrcQB_DepositToAccountRefListID, str2SrcQB_DepositToAccountRefFullName, str2SrcQB_CreditCardTxnInfoInputCreditCardNumber, str2SrcQB_CreditCardTxnInfoInputExpirationMonth, str2SrcQB_CreditCardTxnInfoInputExpirationYear, str2SrcQB_CreditCardTxnInfoInputNameOnCard, str2SrcQB_CreditCardTxnInfoInputCreditCardAddress, str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode, str2SrcQB_CreditCardTxnInfoInputCommercialCardCode, str2SrcQB_CreditCardTxnInfoResultResultCode, str2SrcQB_CreditCardTxnInfoResultResultMessage, str2SrcQB_CreditCardTxnInfoResultCreditCardTransID, str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber, str2SrcQB_CreditCardTxnInfoResultAuthorizationCode, str2SrcQB_CreditCardTxnInfoResultAVSStreet, str2SrcQB_CreditCardTxnInfoResultAVSZip, str2SrcQB_CreditCardTxnInfoResultReconBatchID, str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode, str2SrcQB_CreditCardTxnInfoResultPaymentStatus, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp, str2SrcQB_IsAutoApply, str2SrcQB_UnusedPayment, str2SrcQB_UnusedCredits As String
        'This routine gets the 2SrcQB_QB_ReceivePayment from the database according to the selection in str2SrcQB_QB_ReceivePaymentSQL.
        'It then puts those 2SrcQB_QB_ReceivePayment in the list box

        ''FOR PART 3TestID_
        'Debug.Print "List3TestID_QBTable"
        'Dim rs3TestID_QBTable As ADODB.Recordset
        'Dim str3TestID_QBTableSQL As String
        'Dim str3TestID_QBTableRow As String
        'Dim str3TestID_ListID As String
        ''This routine gets the 3TestID_QBTable from the database according to the selection in str3TestID_QBTableSQL.
        ''It then puts those 3TestID_QBTable in the list box

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_ReceivePayment  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_ReceivePayment"
        Application.DoEvents()


        '
        ''Clear out table
        ''*cnDBPM.Execute "DELETE FROM QB_ReceivePayment"
        'cnmax.Execute "DELETE FROM QB_ReceivePayment"
        '

        'Get rs from QB
        'Load table from rs

        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        rs2SrcQB_QB_ReceivePayment = New DataSet() '*** TAKE QB_ OFF OF TABLE NAME ***
        str2SrcQB_QB_ReceivePaymentSQL = "SELECT * FROM ReceivePayment"
        'Debug.Print str2SrcQB_QB_ReceivePaymentSQL
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_ReceivePaymentSQL, cnQuickBooks)
        rs2SrcQB_QB_ReceivePayment.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_ReceivePayment) ', adAsyncFetch
        If rs2SrcQB_QB_ReceivePayment.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_ReceivePayment"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_ReceivePayment"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_ReceivePayment"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  " & CStr(rs2SrcQB_QB_ReceivePayment.Tables(0).Rows.Count) & "  QB_ReceivePayment  Records ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_ReceivePayment.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_ReceivePayment.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_ReceivePayment.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_ReceivePayment.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_TxnID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_TxnNumber = "0"
                str2SrcQB_CustomerRefListID = ""
                str2SrcQB_CustomerRefFullName = ""
                str2SrcQB_ARAccountRefListID = ""
                str2SrcQB_ARAccountRefFullName = ""
                str2SrcQB_TxnDate = ""
                str2SrcQB_TxnDateMacro = ""
                str2SrcQB_RefNumber = ""
                str2SrcQB_TotalAmount = "0"
                str2SrcQB_PaymentMethodRefListID = ""
                str2SrcQB_PaymentMethodRefFullName = ""
                str2SrcQB_Memo = ""
                str2SrcQB_DepositToAccountRefListID = ""
                str2SrcQB_DepositToAccountRefFullName = ""
                str2SrcQB_CreditCardTxnInfoInputCreditCardNumber = ""
                str2SrcQB_CreditCardTxnInfoInputExpirationMonth = "0"
                str2SrcQB_CreditCardTxnInfoInputExpirationYear = "0"
                str2SrcQB_CreditCardTxnInfoInputNameOnCard = ""
                str2SrcQB_CreditCardTxnInfoInputCreditCardAddress = ""
                str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode = ""
                str2SrcQB_CreditCardTxnInfoInputCommercialCardCode = ""
                str2SrcQB_CreditCardTxnInfoResultResultCode = "0"
                str2SrcQB_CreditCardTxnInfoResultResultMessage = ""
                str2SrcQB_CreditCardTxnInfoResultCreditCardTransID = ""
                str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber = ""
                str2SrcQB_CreditCardTxnInfoResultAuthorizationCode = ""
                str2SrcQB_CreditCardTxnInfoResultAVSStreet = ""
                str2SrcQB_CreditCardTxnInfoResultAVSZip = ""
                str2SrcQB_CreditCardTxnInfoResultReconBatchID = ""
                str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode = "0"
                str2SrcQB_CreditCardTxnInfoResultPaymentStatus = ""
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime = ""
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp = "0"
                str2SrcQB_IsAutoApply = ""
                str2SrcQB_UnusedPayment = "0"
                str2SrcQB_UnusedCredits = "0"

                'get the columns from the database
                If iteration_row("TxnID") <> "" Then str2SrcQB_TxnID = iteration_row("TxnID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("TxnNumber") <> "" Then str2SrcQB_TxnNumber = iteration_row("TxnNumber")
                If iteration_row("CustomerRefListID") <> "" Then str2SrcQB_CustomerRefListID = iteration_row("CustomerRefListID")
                If iteration_row("CustomerRefFullName") <> "" Then str2SrcQB_CustomerRefFullName = iteration_row("CustomerRefFullName")
                If iteration_row("ARAccountRefListID") <> "" Then str2SrcQB_ARAccountRefListID = iteration_row("ARAccountRefListID")
                If iteration_row("ARAccountRefFullName") <> "" Then str2SrcQB_ARAccountRefFullName = iteration_row("ARAccountRefFullName")
                If iteration_row("TxnDate") <> "" Then str2SrcQB_TxnDate = iteration_row("TxnDate")
                If iteration_row("TxnDateMacro") <> "" Then str2SrcQB_TxnDateMacro = iteration_row("TxnDateMacro")
                If iteration_row("RefNumber") <> "" Then str2SrcQB_RefNumber = iteration_row("RefNumber")
                If iteration_row("TotalAmount") <> "" Then str2SrcQB_TotalAmount = iteration_row("TotalAmount")
                If iteration_row("PaymentMethodRefListID") <> "" Then str2SrcQB_PaymentMethodRefListID = iteration_row("PaymentMethodRefListID")
                If iteration_row("PaymentMethodRefFullName") <> "" Then str2SrcQB_PaymentMethodRefFullName = iteration_row("PaymentMethodRefFullName")
                If iteration_row("Memo") <> "" Then str2SrcQB_Memo = iteration_row("Memo")
                If iteration_row("DepositToAccountRefListID") <> "" Then str2SrcQB_DepositToAccountRefListID = iteration_row("DepositToAccountRefListID")
                If iteration_row("DepositToAccountRefFullName") <> "" Then str2SrcQB_DepositToAccountRefFullName = iteration_row("DepositToAccountRefFullName")
                If iteration_row("CreditCardTxnInfoInputCreditCardNumber") <> "" Then str2SrcQB_CreditCardTxnInfoInputCreditCardNumber = iteration_row("CreditCardTxnInfoInputCreditCardNumber")
                If iteration_row("CreditCardTxnInfoInputExpirationMonth") <> "" Then str2SrcQB_CreditCardTxnInfoInputExpirationMonth = iteration_row("CreditCardTxnInfoInputExpirationMonth")
                If iteration_row("CreditCardTxnInfoInputExpirationYear") <> "" Then str2SrcQB_CreditCardTxnInfoInputExpirationYear = iteration_row("CreditCardTxnInfoInputExpirationYear")
                If iteration_row("CreditCardTxnInfoInputNameOnCard") <> "" Then str2SrcQB_CreditCardTxnInfoInputNameOnCard = iteration_row("CreditCardTxnInfoInputNameOnCard")
                If iteration_row("CreditCardTxnInfoInputCreditCardAddress") <> "" Then str2SrcQB_CreditCardTxnInfoInputCreditCardAddress = iteration_row("CreditCardTxnInfoInputCreditCardAddress")
                If iteration_row("CreditCardTxnInfoInputCreditCardPostalCode") <> "" Then str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode = iteration_row("CreditCardTxnInfoInputCreditCardPostalCode")
                If iteration_row("CreditCardTxnInfoInputCommercialCardCode") <> "" Then str2SrcQB_CreditCardTxnInfoInputCommercialCardCode = iteration_row("CreditCardTxnInfoInputCommercialCardCode")
                If iteration_row("CreditCardTxnInfoResultResultCode") <> "" Then str2SrcQB_CreditCardTxnInfoResultResultCode = iteration_row("CreditCardTxnInfoResultResultCode")
                If iteration_row("CreditCardTxnInfoResultResultMessage") <> "" Then str2SrcQB_CreditCardTxnInfoResultResultMessage = iteration_row("CreditCardTxnInfoResultResultMessage")
                If iteration_row("CreditCardTxnInfoResultCreditCardTransID") <> "" Then str2SrcQB_CreditCardTxnInfoResultCreditCardTransID = iteration_row("CreditCardTxnInfoResultCreditCardTransID")
                If iteration_row("CreditCardTxnInfoResultMerchantAccountNumber") <> "" Then str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber = iteration_row("CreditCardTxnInfoResultMerchantAccountNumber")
                If iteration_row("CreditCardTxnInfoResultAuthorizationCode") <> "" Then str2SrcQB_CreditCardTxnInfoResultAuthorizationCode = iteration_row("CreditCardTxnInfoResultAuthorizationCode")
                If iteration_row("CreditCardTxnInfoResultAVSStreet") <> "" Then str2SrcQB_CreditCardTxnInfoResultAVSStreet = iteration_row("CreditCardTxnInfoResultAVSStreet")
                If iteration_row("CreditCardTxnInfoResultAVSZip") <> "" Then str2SrcQB_CreditCardTxnInfoResultAVSZip = iteration_row("CreditCardTxnInfoResultAVSZip")
                If iteration_row("CreditCardTxnInfoResultReconBatchID") <> "" Then str2SrcQB_CreditCardTxnInfoResultReconBatchID = iteration_row("CreditCardTxnInfoResultReconBatchID")
                If iteration_row("CreditCardTxnInfoResultPaymentGroupingCode") <> "" Then str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode = iteration_row("CreditCardTxnInfoResultPaymentGroupingCode")
                If iteration_row("CreditCardTxnInfoResultPaymentStatus") <> "" Then str2SrcQB_CreditCardTxnInfoResultPaymentStatus = iteration_row("CreditCardTxnInfoResultPaymentStatus")
                If iteration_row("CreditCardTxnInfoResultTxnAuthorizationTime") <> "" Then str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime = iteration_row("CreditCardTxnInfoResultTxnAuthorizationTime")
                If iteration_row("CreditCardTxnInfoResultTxnAuthorizationStamp") <> "" Then str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp = iteration_row("CreditCardTxnInfoResultTxnAuthorizationStamp")
                If iteration_row("IsAutoApply") <> "" Then str2SrcQB_IsAutoApply = iteration_row("IsAutoApply")
                If iteration_row("UnusedPayment") <> "" Then str2SrcQB_UnusedPayment = iteration_row("UnusedPayment")
                If iteration_row("UnusedCredits") <> "" Then str2SrcQB_UnusedCredits = iteration_row("UnusedCredits")

                'Strip quote character out of strings
                'Get quote characters out!
                'Change Quote to reverse quote
                'If KeyAscii = 39 Then KeyAscii = 96
                str2SrcQB_TxnID = str2SrcQB_TxnID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_TxnNumber = str2SrcQB_TxnNumber.Replace("'"c, "`"c)
                str2SrcQB_CustomerRefListID = str2SrcQB_CustomerRefListID.Replace("'"c, "`"c)
                str2SrcQB_CustomerRefFullName = str2SrcQB_CustomerRefFullName.Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefListID = str2SrcQB_ARAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefFullName = str2SrcQB_ARAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_TxnDate = str2SrcQB_TxnDate.Replace("'"c, "`"c)
                str2SrcQB_TxnDateMacro = str2SrcQB_TxnDateMacro.Replace("'"c, "`"c)
                str2SrcQB_RefNumber = str2SrcQB_RefNumber.Replace("'"c, "`"c)
                str2SrcQB_TotalAmount = str2SrcQB_TotalAmount.Replace("'"c, "`"c)
                str2SrcQB_PaymentMethodRefListID = str2SrcQB_PaymentMethodRefListID.Replace("'"c, "`"c)
                str2SrcQB_PaymentMethodRefFullName = str2SrcQB_PaymentMethodRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Memo = str2SrcQB_Memo.Replace("'"c, "`"c)
                str2SrcQB_DepositToAccountRefListID = str2SrcQB_DepositToAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_DepositToAccountRefFullName = str2SrcQB_DepositToAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCreditCardNumber = str2SrcQB_CreditCardTxnInfoInputCreditCardNumber.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputExpirationMonth = str2SrcQB_CreditCardTxnInfoInputExpirationMonth.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputExpirationYear = str2SrcQB_CreditCardTxnInfoInputExpirationYear.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputNameOnCard = str2SrcQB_CreditCardTxnInfoInputNameOnCard.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCreditCardAddress = str2SrcQB_CreditCardTxnInfoInputCreditCardAddress.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode = str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCommercialCardCode = str2SrcQB_CreditCardTxnInfoInputCommercialCardCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultResultCode = str2SrcQB_CreditCardTxnInfoResultResultCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultResultMessage = str2SrcQB_CreditCardTxnInfoResultResultMessage.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultCreditCardTransID = str2SrcQB_CreditCardTxnInfoResultCreditCardTransID.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber = str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultAuthorizationCode = str2SrcQB_CreditCardTxnInfoResultAuthorizationCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultAVSStreet = str2SrcQB_CreditCardTxnInfoResultAVSStreet.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultAVSZip = str2SrcQB_CreditCardTxnInfoResultAVSZip.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultReconBatchID = str2SrcQB_CreditCardTxnInfoResultReconBatchID.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode = str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultPaymentStatus = str2SrcQB_CreditCardTxnInfoResultPaymentStatus.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime = str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp = str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp.Replace("'"c, "`"c)
                str2SrcQB_IsAutoApply = str2SrcQB_IsAutoApply.Replace("'"c, "`"c)
                str2SrcQB_UnusedPayment = str2SrcQB_UnusedPayment.Replace("'"c, "`"c)
                str2SrcQB_UnusedCredits = str2SrcQB_UnusedCredits.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsAutoApply = IIf(str2SrcQB_IsAutoApply = "True", "1", "0")
                str2SrcQB_FQSaveToCache = IIf(str2SrcQB_FQSaveToCache = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_ReceivePaymentRow = "" & _
                                                 Strings.Left(str2SrcQB_TxnID & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_TimeCreated & "                  ", 16) & "   " & _
                                                 Strings.Left(str2SrcQB_TimeModified & "                  ", 16) & "   " & _
                                                 Strings.Left(str2SrcQB_CustomerRefFullName & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_TxnDate & "                  ", 10) & "   " & _
                                                 Strings.Left(str2SrcQB_RefNumber & "                  ", 10) & "   " & _
                                                 Strings.Left(str2SrcQB_TotalAmount & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_PaymentMethodRefFullName & "                  ", 18) & "   " & _
                                                 "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_ReceivePayment.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_ReceivePayment.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_ReceivePayment.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_ReceivePaymentRow)
                    'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_TxnID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record

                '        'Update cust balances
                '        UpdateQBCustomerBalance (str2SrcQB_CustomerRefListID)
                '
                '        'Update RPL info
                '        'UpdateQBReceivePaymentLine (str2SrcQB_AppliedToTxnTxnID)
                '        'UpdateQBReceivePaymentLine (str2SrcQB_TxnID)
                '        'UpdateQBReceivePaymentLine (str2SrcQB_CustomerRefListID)
                '        UpdateQBReceivePaymentLine (str2SrcQB_CustomerRefFullName)
                '
                '        'Update inv info
                '        'UpdateQBInvoice (str2SrcQB_TxnID)
                '        UpdateQBInvoice (str2SrcQB_CustomerRefFullName)
                '
                '
                '
                '        'MORE WORK
                '
                '        '"       FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & vbCrLf
                '
                '        'Check to see if ListID or TxnID is in QBTable            'Yes then UPDATE record            'No then INSERT record
                '        'New recordset
                '        Set rs3TestID_QBTable = New ADODB.Recordset
                '        str3TestID_QBTableSQL = "SELECT TxnID FROM QB_ReceivePayment WHERE TxnID = '" & str2SrcQB_TxnID & "'"
                '        'str3TestID_QBTableSQL = "SELECT FQPrimaryKey FROM QB_ReceivePayment WHERE FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'"
                '        'Debug.Print str3TestID_QBTableSQL
                '        'rs3TestID_QBTable.Open str3TestID_QBTableSQL, cnDBPM, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        rs3TestID_QBTable.Open str3TestID_QBTableSQL, cnmax, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        If rs3TestID_QBTable.RecordCount > 1 Then Stop 'Should only be one
                '        If rs3TestID_QBTable.RecordCount > 0 Then  'record exists  -UPDATE
                '            'DO UPDATE WORK:
                '            Debug.Print "UPDATE"
                '
                '            'Build the SQL string
                '            strSQL1 = "UPDATE  " & vbCrLf & _
                ''                      "       QB_ReceivePayment " & vbCrLf & _
                ''                      "SET " & vbCrLf & _
                ''                      "       TxnID = '" & str2SrcQB_TxnID & "'" & vbCrLf & _
                ''                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & vbCrLf & _
                ''                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & vbCrLf & _
                ''                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & vbCrLf & _
                ''                      "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & vbCrLf & _
                ''                      "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & vbCrLf & _
                ''                      "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & vbCrLf & _
                ''                      "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & vbCrLf & _
                ''                      "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & vbCrLf & _
                ''                      "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & vbCrLf & _
                ''                      "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & vbCrLf & _
                ''                      "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & vbCrLf & _
                ''                      "     , TotalAmount = " & str2SrcQB_TotalAmount & "" & vbCrLf & _
                ''                      "     , PaymentMethodRefListID = '" & str2SrcQB_PaymentMethodRefListID & "'" & vbCrLf & _
                ''                      "     , PaymentMethodRefFullName = '" & str2SrcQB_PaymentMethodRefFullName & "'" & vbCrLf & _
                ''                      "     , Memo = '" & str2SrcQB_Memo & "'" & vbCrLf & _
                ''                      "     , DepositToAccountRefListID = '" & str2SrcQB_DepositToAccountRefListID & "'" & vbCrLf
                '            strSQL2 = "     , DepositToAccountRefFullName = '" & str2SrcQB_DepositToAccountRefFullName & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputCreditCardNumber = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardNumber & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputExpirationMonth = " & str2SrcQB_CreditCardTxnInfoInputExpirationMonth & "" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputExpirationYear = " & str2SrcQB_CreditCardTxnInfoInputExpirationYear & "" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputNameOnCard = '" & str2SrcQB_CreditCardTxnInfoInputNameOnCard & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputCreditCardAddress = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardAddress & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputCreditCardPostalCode = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputCommercialCardCode = '" & str2SrcQB_CreditCardTxnInfoInputCommercialCardCode & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultResultCode = " & str2SrcQB_CreditCardTxnInfoResultResultCode & "" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultResultMessage = '" & str2SrcQB_CreditCardTxnInfoResultResultMessage & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultCreditCardTransID = '" & str2SrcQB_CreditCardTxnInfoResultCreditCardTransID & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultMerchantAccountNumber = '" & str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultAuthorizationCode = '" & str2SrcQB_CreditCardTxnInfoResultAuthorizationCode & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultAVSStreet = '" & str2SrcQB_CreditCardTxnInfoResultAVSStreet & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultAVSZip = '" & str2SrcQB_CreditCardTxnInfoResultAVSZip & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultReconBatchID = '" & str2SrcQB_CreditCardTxnInfoResultReconBatchID & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultPaymentGroupingCode = " & str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode & "" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultPaymentStatus = '" & str2SrcQB_CreditCardTxnInfoResultPaymentStatus & "'" & vbCrLf
                '            strSQL3 = "     , CreditCardTxnInfoResultTxnAuthorizationTime = '" & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultTxnAuthorizationStamp = " & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp & "" & vbCrLf & _
                ''                      "     , IsAutoApply = '" & str2SrcQB_IsAutoApply & "'" & vbCrLf & _
                ''                      "     , UnusedPayment = " & str2SrcQB_UnusedPayment & "" & vbCrLf & _
                ''                      "     , UnusedCredits = " & str2SrcQB_UnusedCredits & "" & vbCrLf & _
                ''                      "WHERE " & vbCrLf & _
                ''                      "       TxnID = '" & str2SrcQB_TxnID & "'" & vbCrLf
                '
                '            'Combine the strings
                '            strTableUpdate = strSQL1 & strSQL2 & strSQL3 '& strSQL4 & strSQL5 & strSQL6
                '            'Debug.Print strTableUpdate
                '
                '            'Execute the insert
                '            '*cnDBPM.Execute strTableUpdate
                '            cnmax.Execute strTableUpdate
                '
                '
                '
                '
                '        Else 'record not exist  -INSERT
                '            'DO INSERT WORK:
                '            Debug.Print "INSERT"

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_ReceivePayment " & Environment.NewLine & _
                          "   ( TxnID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , TxnNumber " & Environment.NewLine & _
                          "   , CustomerRefListID " & Environment.NewLine & _
                          "   , CustomerRefFullName " & Environment.NewLine & _
                          "   , ARAccountRefListID " & Environment.NewLine & _
                          "   , ARAccountRefFullName " & Environment.NewLine & _
                          "   , TxnDate " & Environment.NewLine & _
                          "   , TxnDateMacro " & Environment.NewLine & _
                          "   , RefNumber " & Environment.NewLine & _
                          "   , TotalAmount " & Environment.NewLine & _
                          "   , PaymentMethodRefListID " & Environment.NewLine & _
                          "   , PaymentMethodRefFullName " & Environment.NewLine & _
                          "   , Memo " & Environment.NewLine & _
                          "   , DepositToAccountRefListID " & Environment.NewLine
                strSQL2 = "   , DepositToAccountRefFullName " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputCreditCardNumber " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputExpirationMonth " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputExpirationYear " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputNameOnCard " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputCreditCardAddress " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputCreditCardPostalCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputCommercialCardCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultResultCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultResultMessage " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultCreditCardTransID " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultMerchantAccountNumber " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultAuthorizationCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultAVSStreet " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultAVSZip " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultReconBatchID " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultPaymentGroupingCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultPaymentStatus " & Environment.NewLine
                strSQL3 = "   , CreditCardTxnInfoResultTxnAuthorizationTime " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultTxnAuthorizationStamp " & Environment.NewLine & _
                          "   , IsAutoApply " & Environment.NewLine & _
                          "   , UnusedPayment " & Environment.NewLine & _
                          "   , UnusedCredits ) " & Environment.NewLine
                strSQL4 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_TxnID & "'  --TxnID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , " & str2SrcQB_TxnNumber & "  --TxnNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerRefListID & "'  --CustomerRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerRefFullName & "'  --CustomerRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ARAccountRefListID & "'  --ARAccountRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ARAccountRefFullName & "'  --ARAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TxnDate & "'  --TxnDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TxnDateMacro & "'  --TxnDateMacro" & Environment.NewLine & _
                          "   , '" & str2SrcQB_RefNumber & "'  --RefNumber" & Environment.NewLine & _
                          "   , " & str2SrcQB_TotalAmount & "  --TotalAmount" & Environment.NewLine & _
                          "   , '" & str2SrcQB_PaymentMethodRefListID & "'  --PaymentMethodRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_PaymentMethodRefFullName & "'  --PaymentMethodRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Memo & "'  --Memo" & Environment.NewLine & _
                          "   , '" & str2SrcQB_DepositToAccountRefListID & "'  --DepositToAccountRefListID" & Environment.NewLine
                strSQL5 = "   , '" & str2SrcQB_DepositToAccountRefFullName & "'  --DepositToAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputCreditCardNumber & "'  --CreditCardTxnInfoInputCreditCardNumber" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoInputExpirationMonth & "  --CreditCardTxnInfoInputExpirationMonth" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoInputExpirationYear & "  --CreditCardTxnInfoInputExpirationYear" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputNameOnCard & "'  --CreditCardTxnInfoInputNameOnCard" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputCreditCardAddress & "'  --CreditCardTxnInfoInputCreditCardAddress" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode & "'  --CreditCardTxnInfoInputCreditCardPostalCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputCommercialCardCode & "'  --CreditCardTxnInfoInputCommercialCardCode" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoResultResultCode & "  --CreditCardTxnInfoResultResultCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultResultMessage & "'  --CreditCardTxnInfoResultResultMessage" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultCreditCardTransID & "'  --CreditCardTxnInfoResultCreditCardTransID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber & "'  --CreditCardTxnInfoResultMerchantAccountNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultAuthorizationCode & "'  --CreditCardTxnInfoResultAuthorizationCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultAVSStreet & "'  --CreditCardTxnInfoResultAVSStreet" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultAVSZip & "'  --CreditCardTxnInfoResultAVSZip" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultReconBatchID & "'  --CreditCardTxnInfoResultReconBatchID" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode & "  --CreditCardTxnInfoResultPaymentGroupingCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultPaymentStatus & "'  --CreditCardTxnInfoResultPaymentStatus" & Environment.NewLine
                strSQL6 = "   , '" & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime & "'  --CreditCardTxnInfoResultTxnAuthorizationTime" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp & "  --CreditCardTxnInfoResultTxnAuthorizationStamp" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsAutoApply & "'  --IsAutoApply" & Environment.NewLine & _
                          "   , " & str2SrcQB_UnusedPayment & "  --UnusedPayment" & Environment.NewLine & _
                          "   , " & str2SrcQB_UnusedCredits & " ) --FQPrimaryKey" & Environment.NewLine

                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
                'Debug.Print strTableInsert

                '            'Execute the insert
                '            '*cnDBPM.Execute strTableInsert
                '            cnMax.Execute strTableInsert

                'Execute the insert
                If gstrCompany = "DrummondPrinting" Then
                    '*cnDBPM.Execute strTableInsert
                    Dim TempCommand_3 As SqlCommand
                    TempCommand_3 = cnMax.CreateCommand()
                    TempCommand_3.CommandText = strTableInsert
                    TempCommand_3.ExecuteNonQuery()
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    Dim TempCommand_4 As SqlCommand
                    TempCommand_4 = cnMax.CreateCommand()
                    TempCommand_4.CommandText = strTableInsert
                    TempCommand_4.ExecuteNonQuery()
                End If

                '
                '        End If
                '


            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing
            'frmMain.lstConversionProgress.AddItem "" & Now & "   Processing  0  QB_ReceivePayment  Records "
            'frmMain.lblListboxStatus.Caption = "" & Now & "   Processing  QB_ReceivePayment  Records "
            'DoEvents

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '        End If
        End If



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QBTable.Close()
            rs1MaxOfCopy_QBTable = Nothing


            rs2SrcQB_QB_ReceivePayment = Nothing


            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QBTable.Close()
            rs3TestID_QBTable = Nothing



            Exit Sub


            MessageBox.Show("<<RefreshQB_ReceivePayment>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub


    Public Sub ReloadQB_ReceivePaymentLine()
        Dim rs1MaxOfCopy_QBTable, rs3TestID_QBTable As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_ReceivePaymentLine" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:




        ''FOR PART 1MaxOfCopy_ - Get records from QBTable
        'Debug.Print "List1MaxOfCopy_QBTable"
        'Dim rs1MaxOfCopy_QBTable As ADODB.Recordset
        'Dim str1MaxOfCopy_QBTableSQL As String
        'Dim str1MaxOfCopy_QBTableRow As String
        'Dim str1MaxOfCopy_TimeModified As String
        ''This routine gets the 1MaxOfCopy_QBTable from the database according to the selection in str1MaxOfCopy_QBTableSQL.
        ''It then puts those 1MaxOfCopy_QBTable in the list box

        'FOR PART 2SrcQB_ - Get records from QB_ReceivePaymentLine
        Debug.WriteLine("List2SrcQB_QB_ReceivePaymentLine")
        Dim rs2SrcQB_QB_ReceivePaymentLine As DataSet
        Dim str2SrcQB_QB_ReceivePaymentLineSQL, str2SrcQB_QB_ReceivePaymentLineRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_TotalAmount, str2SrcQB_PaymentMethodRefListID, str2SrcQB_PaymentMethodRefFullName, str2SrcQB_Memo, str2SrcQB_DepositToAccountRefListID, str2SrcQB_DepositToAccountRefFullName, str2SrcQB_CreditCardTxnInfoInputCreditCardNumber, str2SrcQB_CreditCardTxnInfoInputExpirationMonth, str2SrcQB_CreditCardTxnInfoInputExpirationYear, str2SrcQB_CreditCardTxnInfoInputNameOnCard, str2SrcQB_CreditCardTxnInfoInputCreditCardAddress, str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode, str2SrcQB_CreditCardTxnInfoInputCommercialCardCode, str2SrcQB_CreditCardTxnInfoResultResultCode, str2SrcQB_CreditCardTxnInfoResultResultMessage, str2SrcQB_CreditCardTxnInfoResultCreditCardTransID, str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber, str2SrcQB_CreditCardTxnInfoResultAuthorizationCode, str2SrcQB_CreditCardTxnInfoResultAVSStreet, str2SrcQB_CreditCardTxnInfoResultAVSZip, str2SrcQB_CreditCardTxnInfoResultReconBatchID, str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode, str2SrcQB_CreditCardTxnInfoResultPaymentStatus, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp, str2SrcQB_IsAutoApply, str2SrcQB_UnusedPayment, str2SrcQB_UnusedCredits, str2SrcQB_AppliedToTxnTxnID, str2SrcQB_AppliedToTxnPaymentAmount, str2SrcQB_AppliedToTxnTxnType, str2SrcQB_AppliedToTxnTxnDate, str2SrcQB_AppliedToTxnRefNumber, str2SrcQB_AppliedToTxnBalanceRemaining, str2SrcQB_AppliedToTxnAmount, str2SrcQB_AppliedToTxnSetCreditCreditTxnID, str2SrcQB_AppliedToTxnSetCreditAppliedAmount, str2SrcQB_AppliedToTxnDiscountAmount, str2SrcQB_AppliedToTxnDiscountAccountRefListID, str2SrcQB_AppliedToTxnDiscountAccountRefFullName, str2SrcQB_FQSaveToCache, str2SrcQB_FQPrimaryKey As String
        'This routine gets the 2SrcQB_QB_ReceivePaymentLine from the database according to the selection in str2SrcQB_QB_ReceivePaymentLineSQL.
        'It then puts those 2SrcQB_QB_ReceivePaymentLine in the list box

        ''FOR PART 3TestID_
        'Debug.Print "List3TestID_QBTable"
        'Dim rs3TestID_QBTable As ADODB.Recordset
        'Dim str3TestID_QBTableSQL As String
        'Dim str3TestID_QBTableRow As String
        'Dim str3TestID_ListID As String
        ''This routine gets the 3TestID_QBTable from the database according to the selection in str3TestID_QBTableSQL.
        ''It then puts those 3TestID_QBTable in the list box

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_ReceivePaymentLine  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_ReceivePaymentLine"
        Application.DoEvents()


        '
        ''Clear out table
        ''*cnDBPM.Execute "DELETE FROM QB_ReceivePaymentLine"
        'cnmax.Execute "DELETE FROM QB_ReceivePaymentLine"
        '

        'Get rs from QB
        'Load table from rs

        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        rs2SrcQB_QB_ReceivePaymentLine = New DataSet() '*** TAKE QB_ OFF OF TABLE NAME ***
        'str2SrcQB_QB_ReceivePaymentLineSQL = "SELECT * FROM ReceivePaymentLine WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_ReceivePaymentLineSQL = "SELECT * FROM ReceivePaymentLine WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'str2SrcQB_QB_ReceivePaymentLineSQL = "SELECT * FROM ReceivePaymentLine WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'str2SrcQB_QB_ReceivePaymentLineSQL = "SELECT * FROM ReceivePaymentLine WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_ReceivePaymentLine & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_ReceivePaymentLineSQL = "SELECT * FROM ReceivePaymentLine WHERE TxnID = '38B69-1141162298'"
        str2SrcQB_QB_ReceivePaymentLineSQL = "SELECT * FROM ReceivePaymentLine"
        'Debug.Print str2SrcQB_QB_ReceivePaymentLineSQL
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_ReceivePaymentLineSQL, cnQuickBooks)
        rs2SrcQB_QB_ReceivePaymentLine.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_ReceivePaymentLine) ', adAsyncFetch
        If rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_ReceivePaymentLine"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_ReceivePaymentLine"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_ReceivePaymentLine"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  " & CStr(rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows.Count) & "  QB_ReceivePaymentLine  Records ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_ReceivePaymentLine.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_ReceivePaymentLine.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_TxnID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_TxnNumber = "0"
                str2SrcQB_CustomerRefListID = ""
                str2SrcQB_CustomerRefFullName = ""
                str2SrcQB_ARAccountRefListID = ""
                str2SrcQB_ARAccountRefFullName = ""
                str2SrcQB_TxnDate = ""
                str2SrcQB_TxnDateMacro = ""
                str2SrcQB_RefNumber = ""
                str2SrcQB_TotalAmount = "0"
                str2SrcQB_PaymentMethodRefListID = ""
                str2SrcQB_PaymentMethodRefFullName = ""
                str2SrcQB_Memo = ""
                str2SrcQB_DepositToAccountRefListID = ""
                str2SrcQB_DepositToAccountRefFullName = ""
                str2SrcQB_CreditCardTxnInfoInputCreditCardNumber = ""
                str2SrcQB_CreditCardTxnInfoInputExpirationMonth = "0"
                str2SrcQB_CreditCardTxnInfoInputExpirationYear = "0"
                str2SrcQB_CreditCardTxnInfoInputNameOnCard = ""
                str2SrcQB_CreditCardTxnInfoInputCreditCardAddress = ""
                str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode = ""
                str2SrcQB_CreditCardTxnInfoInputCommercialCardCode = ""
                str2SrcQB_CreditCardTxnInfoResultResultCode = "0"
                str2SrcQB_CreditCardTxnInfoResultResultMessage = ""
                str2SrcQB_CreditCardTxnInfoResultCreditCardTransID = ""
                str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber = ""
                str2SrcQB_CreditCardTxnInfoResultAuthorizationCode = ""
                str2SrcQB_CreditCardTxnInfoResultAVSStreet = ""
                str2SrcQB_CreditCardTxnInfoResultAVSZip = ""
                str2SrcQB_CreditCardTxnInfoResultReconBatchID = ""
                str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode = "0"
                str2SrcQB_CreditCardTxnInfoResultPaymentStatus = ""
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime = ""
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp = "0"
                str2SrcQB_IsAutoApply = ""
                str2SrcQB_UnusedPayment = "0"
                str2SrcQB_UnusedCredits = "0"
                str2SrcQB_AppliedToTxnTxnID = ""
                str2SrcQB_AppliedToTxnPaymentAmount = "0"
                str2SrcQB_AppliedToTxnTxnType = ""
                str2SrcQB_AppliedToTxnTxnDate = ""
                str2SrcQB_AppliedToTxnRefNumber = ""
                str2SrcQB_AppliedToTxnBalanceRemaining = "0"
                str2SrcQB_AppliedToTxnAmount = "0"
                str2SrcQB_AppliedToTxnSetCreditCreditTxnID = ""
                str2SrcQB_AppliedToTxnSetCreditAppliedAmount = "0"
                str2SrcQB_AppliedToTxnDiscountAmount = "0"
                str2SrcQB_AppliedToTxnDiscountAccountRefListID = ""
                str2SrcQB_AppliedToTxnDiscountAccountRefFullName = ""
                str2SrcQB_FQSaveToCache = ""
                str2SrcQB_FQPrimaryKey = ""

                'get the columns from the database
                If iteration_row("TxnID") <> "" Then str2SrcQB_TxnID = iteration_row("TxnID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("TxnNumber") <> "" Then str2SrcQB_TxnNumber = iteration_row("TxnNumber")
                If iteration_row("CustomerRefListID") <> "" Then str2SrcQB_CustomerRefListID = iteration_row("CustomerRefListID")
                If iteration_row("CustomerRefFullName") <> "" Then str2SrcQB_CustomerRefFullName = iteration_row("CustomerRefFullName")
                If iteration_row("ARAccountRefListID") <> "" Then str2SrcQB_ARAccountRefListID = iteration_row("ARAccountRefListID")
                If iteration_row("ARAccountRefFullName") <> "" Then str2SrcQB_ARAccountRefFullName = iteration_row("ARAccountRefFullName")
                If iteration_row("TxnDate") <> "" Then str2SrcQB_TxnDate = iteration_row("TxnDate")
                If iteration_row("TxnDateMacro") <> "" Then str2SrcQB_TxnDateMacro = iteration_row("TxnDateMacro")
                If iteration_row("RefNumber") <> "" Then str2SrcQB_RefNumber = iteration_row("RefNumber")
                If iteration_row("TotalAmount") <> "" Then str2SrcQB_TotalAmount = iteration_row("TotalAmount")
                If iteration_row("PaymentMethodRefListID") <> "" Then str2SrcQB_PaymentMethodRefListID = iteration_row("PaymentMethodRefListID")
                If iteration_row("PaymentMethodRefFullName") <> "" Then str2SrcQB_PaymentMethodRefFullName = iteration_row("PaymentMethodRefFullName")
                If iteration_row("Memo") <> "" Then str2SrcQB_Memo = iteration_row("Memo")
                If iteration_row("DepositToAccountRefListID") <> "" Then str2SrcQB_DepositToAccountRefListID = iteration_row("DepositToAccountRefListID")
                If iteration_row("DepositToAccountRefFullName") <> "" Then str2SrcQB_DepositToAccountRefFullName = iteration_row("DepositToAccountRefFullName")
                If iteration_row("CreditCardTxnInfoInputCreditCardNumber") <> "" Then str2SrcQB_CreditCardTxnInfoInputCreditCardNumber = iteration_row("CreditCardTxnInfoInputCreditCardNumber")
                If iteration_row("CreditCardTxnInfoInputExpirationMonth") <> "" Then str2SrcQB_CreditCardTxnInfoInputExpirationMonth = iteration_row("CreditCardTxnInfoInputExpirationMonth")
                If iteration_row("CreditCardTxnInfoInputExpirationYear") <> "" Then str2SrcQB_CreditCardTxnInfoInputExpirationYear = iteration_row("CreditCardTxnInfoInputExpirationYear")
                If iteration_row("CreditCardTxnInfoInputNameOnCard") <> "" Then str2SrcQB_CreditCardTxnInfoInputNameOnCard = iteration_row("CreditCardTxnInfoInputNameOnCard")
                If iteration_row("CreditCardTxnInfoInputCreditCardAddress") <> "" Then str2SrcQB_CreditCardTxnInfoInputCreditCardAddress = iteration_row("CreditCardTxnInfoInputCreditCardAddress")
                If iteration_row("CreditCardTxnInfoInputCreditCardPostalCode") <> "" Then str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode = iteration_row("CreditCardTxnInfoInputCreditCardPostalCode")
                If iteration_row("CreditCardTxnInfoInputCommercialCardCode") <> "" Then str2SrcQB_CreditCardTxnInfoInputCommercialCardCode = iteration_row("CreditCardTxnInfoInputCommercialCardCode")
                If iteration_row("CreditCardTxnInfoResultResultCode") <> "" Then str2SrcQB_CreditCardTxnInfoResultResultCode = iteration_row("CreditCardTxnInfoResultResultCode")
                If iteration_row("CreditCardTxnInfoResultResultMessage") <> "" Then str2SrcQB_CreditCardTxnInfoResultResultMessage = iteration_row("CreditCardTxnInfoResultResultMessage")
                If iteration_row("CreditCardTxnInfoResultCreditCardTransID") <> "" Then str2SrcQB_CreditCardTxnInfoResultCreditCardTransID = iteration_row("CreditCardTxnInfoResultCreditCardTransID")
                If iteration_row("CreditCardTxnInfoResultMerchantAccountNumber") <> "" Then str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber = iteration_row("CreditCardTxnInfoResultMerchantAccountNumber")
                If iteration_row("CreditCardTxnInfoResultAuthorizationCode") <> "" Then str2SrcQB_CreditCardTxnInfoResultAuthorizationCode = iteration_row("CreditCardTxnInfoResultAuthorizationCode")
                If iteration_row("CreditCardTxnInfoResultAVSStreet") <> "" Then str2SrcQB_CreditCardTxnInfoResultAVSStreet = iteration_row("CreditCardTxnInfoResultAVSStreet")
                If iteration_row("CreditCardTxnInfoResultAVSZip") <> "" Then str2SrcQB_CreditCardTxnInfoResultAVSZip = iteration_row("CreditCardTxnInfoResultAVSZip")
                If iteration_row("CreditCardTxnInfoResultReconBatchID") <> "" Then str2SrcQB_CreditCardTxnInfoResultReconBatchID = iteration_row("CreditCardTxnInfoResultReconBatchID")
                If iteration_row("CreditCardTxnInfoResultPaymentGroupingCode") <> "" Then str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode = iteration_row("CreditCardTxnInfoResultPaymentGroupingCode")
                If iteration_row("CreditCardTxnInfoResultPaymentStatus") <> "" Then str2SrcQB_CreditCardTxnInfoResultPaymentStatus = iteration_row("CreditCardTxnInfoResultPaymentStatus")
                If iteration_row("CreditCardTxnInfoResultTxnAuthorizationTime") <> "" Then str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime = iteration_row("CreditCardTxnInfoResultTxnAuthorizationTime")
                If iteration_row("CreditCardTxnInfoResultTxnAuthorizationStamp") <> "" Then str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp = iteration_row("CreditCardTxnInfoResultTxnAuthorizationStamp")
                If iteration_row("IsAutoApply") <> "" Then str2SrcQB_IsAutoApply = iteration_row("IsAutoApply")
                If iteration_row("UnusedPayment") <> "" Then str2SrcQB_UnusedPayment = iteration_row("UnusedPayment")
                If iteration_row("UnusedCredits") <> "" Then str2SrcQB_UnusedCredits = iteration_row("UnusedCredits")
                If iteration_row("AppliedToTxnTxnID") <> "" Then str2SrcQB_AppliedToTxnTxnID = iteration_row("AppliedToTxnTxnID")
                If iteration_row("AppliedToTxnPaymentAmount") <> "" Then str2SrcQB_AppliedToTxnPaymentAmount = iteration_row("AppliedToTxnPaymentAmount")
                If iteration_row("AppliedToTxnTxnType") <> "" Then str2SrcQB_AppliedToTxnTxnType = iteration_row("AppliedToTxnTxnType")
                If iteration_row("AppliedToTxnTxnDate") <> "" Then str2SrcQB_AppliedToTxnTxnDate = iteration_row("AppliedToTxnTxnDate")
                If iteration_row("AppliedToTxnRefNumber") <> "" Then str2SrcQB_AppliedToTxnRefNumber = iteration_row("AppliedToTxnRefNumber")
                If iteration_row("AppliedToTxnBalanceRemaining") <> "" Then str2SrcQB_AppliedToTxnBalanceRemaining = iteration_row("AppliedToTxnBalanceRemaining")
                If iteration_row("AppliedToTxnAmount") <> "" Then str2SrcQB_AppliedToTxnAmount = iteration_row("AppliedToTxnAmount")
                If iteration_row("AppliedToTxnSetCreditCreditTxnID") <> "" Then str2SrcQB_AppliedToTxnSetCreditCreditTxnID = iteration_row("AppliedToTxnSetCreditCreditTxnID")
                If iteration_row("AppliedToTxnSetCreditAppliedAmount") <> "" Then str2SrcQB_AppliedToTxnSetCreditAppliedAmount = iteration_row("AppliedToTxnSetCreditAppliedAmount")
                If iteration_row("AppliedToTxnDiscountAmount") <> "" Then str2SrcQB_AppliedToTxnDiscountAmount = iteration_row("AppliedToTxnDiscountAmount")
                If iteration_row("AppliedToTxnDiscountAccountRefListID") <> "" Then str2SrcQB_AppliedToTxnDiscountAccountRefListID = iteration_row("AppliedToTxnDiscountAccountRefListID")
                If iteration_row("AppliedToTxnDiscountAccountRefFullName") <> "" Then str2SrcQB_AppliedToTxnDiscountAccountRefFullName = iteration_row("AppliedToTxnDiscountAccountRefFullName")
                If iteration_row("FQSaveToCache") <> "" Then str2SrcQB_FQSaveToCache = iteration_row("FQSaveToCache")
                If iteration_row("FQPrimaryKey") <> "" Then str2SrcQB_FQPrimaryKey = iteration_row("FQPrimaryKey")

                'Strip quote character out of strings
                'Get quote characters out!
                'Change Quote to reverse quote
                'If KeyAscii = 39 Then KeyAscii = 96
                str2SrcQB_TxnID = str2SrcQB_TxnID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_TxnNumber = str2SrcQB_TxnNumber.Replace("'"c, "`"c)
                str2SrcQB_CustomerRefListID = str2SrcQB_CustomerRefListID.Replace("'"c, "`"c)
                str2SrcQB_CustomerRefFullName = str2SrcQB_CustomerRefFullName.Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefListID = str2SrcQB_ARAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefFullName = str2SrcQB_ARAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_TxnDate = str2SrcQB_TxnDate.Replace("'"c, "`"c)
                str2SrcQB_TxnDateMacro = str2SrcQB_TxnDateMacro.Replace("'"c, "`"c)
                str2SrcQB_RefNumber = str2SrcQB_RefNumber.Replace("'"c, "`"c)
                str2SrcQB_TotalAmount = str2SrcQB_TotalAmount.Replace("'"c, "`"c)
                str2SrcQB_PaymentMethodRefListID = str2SrcQB_PaymentMethodRefListID.Replace("'"c, "`"c)
                str2SrcQB_PaymentMethodRefFullName = str2SrcQB_PaymentMethodRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Memo = str2SrcQB_Memo.Replace("'"c, "`"c)
                str2SrcQB_DepositToAccountRefListID = str2SrcQB_DepositToAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_DepositToAccountRefFullName = str2SrcQB_DepositToAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCreditCardNumber = str2SrcQB_CreditCardTxnInfoInputCreditCardNumber.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputExpirationMonth = str2SrcQB_CreditCardTxnInfoInputExpirationMonth.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputExpirationYear = str2SrcQB_CreditCardTxnInfoInputExpirationYear.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputNameOnCard = str2SrcQB_CreditCardTxnInfoInputNameOnCard.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCreditCardAddress = str2SrcQB_CreditCardTxnInfoInputCreditCardAddress.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode = str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoInputCommercialCardCode = str2SrcQB_CreditCardTxnInfoInputCommercialCardCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultResultCode = str2SrcQB_CreditCardTxnInfoResultResultCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultResultMessage = str2SrcQB_CreditCardTxnInfoResultResultMessage.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultCreditCardTransID = str2SrcQB_CreditCardTxnInfoResultCreditCardTransID.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber = str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultAuthorizationCode = str2SrcQB_CreditCardTxnInfoResultAuthorizationCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultAVSStreet = str2SrcQB_CreditCardTxnInfoResultAVSStreet.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultAVSZip = str2SrcQB_CreditCardTxnInfoResultAVSZip.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultReconBatchID = str2SrcQB_CreditCardTxnInfoResultReconBatchID.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode = str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultPaymentStatus = str2SrcQB_CreditCardTxnInfoResultPaymentStatus.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime = str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime.Replace("'"c, "`"c)
                str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp = str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp.Replace("'"c, "`"c)
                str2SrcQB_IsAutoApply = str2SrcQB_IsAutoApply.Replace("'"c, "`"c)
                str2SrcQB_UnusedPayment = str2SrcQB_UnusedPayment.Replace("'"c, "`"c)
                str2SrcQB_UnusedCredits = str2SrcQB_UnusedCredits.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnTxnID = str2SrcQB_AppliedToTxnTxnID.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnPaymentAmount = str2SrcQB_AppliedToTxnPaymentAmount.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnTxnType = str2SrcQB_AppliedToTxnTxnType.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnTxnDate = str2SrcQB_AppliedToTxnTxnDate.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnRefNumber = str2SrcQB_AppliedToTxnRefNumber.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnBalanceRemaining = str2SrcQB_AppliedToTxnBalanceRemaining.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnAmount = str2SrcQB_AppliedToTxnAmount.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnSetCreditCreditTxnID = str2SrcQB_AppliedToTxnSetCreditCreditTxnID.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnSetCreditAppliedAmount = str2SrcQB_AppliedToTxnSetCreditAppliedAmount.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnDiscountAmount = str2SrcQB_AppliedToTxnDiscountAmount.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnDiscountAccountRefListID = str2SrcQB_AppliedToTxnDiscountAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_AppliedToTxnDiscountAccountRefFullName = str2SrcQB_AppliedToTxnDiscountAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_FQSaveToCache = str2SrcQB_FQSaveToCache.Replace("'"c, "`"c)
                str2SrcQB_FQPrimaryKey = str2SrcQB_FQPrimaryKey.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsAutoApply = IIf(str2SrcQB_IsAutoApply = "True", "1", "0")
                str2SrcQB_FQSaveToCache = IIf(str2SrcQB_FQSaveToCache = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_ReceivePaymentLineRow = "" & _
                                                     Strings.Left(str2SrcQB_TxnID & "                  ", 18) & "   " & _
                                                     Strings.Left(str2SrcQB_TimeCreated & "                  ", 16) & "   " & _
                                                     Strings.Left(str2SrcQB_TimeModified & "                  ", 16) & "   " & _
                                                     Strings.Left(str2SrcQB_CustomerRefFullName & "                  ", 18) & "   " & _
                                                     Strings.Left(str2SrcQB_TxnDate & "                  ", 10) & "   " & _
                                                     Strings.Left(str2SrcQB_RefNumber & "                  ", 10) & "   " & _
                                                     Strings.Left(str2SrcQB_TotalAmount & "                  ", 18) & "   " & _
                                                     Strings.Left(str2SrcQB_PaymentMethodRefFullName & "                  ", 18) & "   " & _
                                                     "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_ReceivePaymentLine.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_ReceivePaymentLine.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_ReceivePaymentLineRow)
                    'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_TxnID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record

                '        'Update inv info
                '        'UpdateQBInvoice (str2SrcQB_TxnID)
                '        'UpdateQBInvoice (str2SrcQB_AppliedToTxnTxnID)
                '        UpdateQBInvoice (str2SrcQB_CustomerRefFullName)
                '
                '
                '        'Update cust balances
                '        UpdateQBCustomerBalance (str2SrcQB_CustomerRefListID)
                '
                '        'Update RPL info
                '        'UpdateQBReceivePaymentLine (str2SrcQB_AppliedToTxnTxnID)
                '        'UpdateQBReceivePaymentLine (str2SrcQB_TxnID)
                '        'UpdateQBReceivePaymentLine (str2SrcQB_CustomerRefListID)
                '        UpdateQBReceivePaymentLine (str2SrcQB_CustomerRefFullName)
                '
                '
                '        '"       FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & vbCrLf
                '
                '        'Check to see if ListID or TxnID is in QBTable            'Yes then UPDATE record            'No then INSERT record
                '        'New recordset
                '        Set rs3TestID_QBTable = New ADODB.Recordset
                '        str3TestID_QBTableSQL = "SELECT TxnID FROM QB_ReceivePaymentLine WHERE TxnID = '" & str2SrcQB_TxnID & "'"
                '        str3TestID_QBTableSQL = "SELECT FQPrimaryKey FROM QB_ReceivePaymentLine WHERE FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'"
                '        'Debug.Print str3TestID_QBTableSQL
                '        'rs3TestID_QBTable.Open str3TestID_QBTableSQL, cnDBPM, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        rs3TestID_QBTable.Open str3TestID_QBTableSQL, cnmax, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        If rs3TestID_QBTable.RecordCount > 1 Then Stop 'Should only be one
                '        If rs3TestID_QBTable.RecordCount > 0 Then  'record exists  -UPDATE
                '            'DO UPDATE WORK:
                '            Debug.Print "UPDATE"
                '
                '            'Build the SQL string
                '            strSQL1 = "UPDATE  " & vbCrLf & _
                ''                      "       QB_ReceivePaymentLine " & vbCrLf & _
                ''                      "SET " & vbCrLf & _
                ''                      "       TxnID = '" & str2SrcQB_TxnID & "'" & vbCrLf & _
                ''                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & vbCrLf & _
                ''                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & vbCrLf & _
                ''                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & vbCrLf & _
                ''                      "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & vbCrLf & _
                ''                      "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & vbCrLf & _
                ''                      "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & vbCrLf & _
                ''                      "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & vbCrLf & _
                ''                      "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & vbCrLf & _
                ''                      "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & vbCrLf & _
                ''                      "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & vbCrLf & _
                ''                      "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & vbCrLf & _
                ''                      "     , TotalAmount = " & str2SrcQB_TotalAmount & "" & vbCrLf & _
                ''                      "     , PaymentMethodRefListID = '" & str2SrcQB_PaymentMethodRefListID & "'" & vbCrLf & _
                ''                      "     , PaymentMethodRefFullName = '" & str2SrcQB_PaymentMethodRefFullName & "'" & vbCrLf & _
                ''                      "     , Memo = '" & str2SrcQB_Memo & "'" & vbCrLf & _
                ''                      "     , DepositToAccountRefListID = '" & str2SrcQB_DepositToAccountRefListID & "'" & vbCrLf
                '            strSQL2 = "     , DepositToAccountRefFullName = '" & str2SrcQB_DepositToAccountRefFullName & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputCreditCardNumber = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardNumber & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputExpirationMonth = " & str2SrcQB_CreditCardTxnInfoInputExpirationMonth & "" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputExpirationYear = " & str2SrcQB_CreditCardTxnInfoInputExpirationYear & "" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputNameOnCard = '" & str2SrcQB_CreditCardTxnInfoInputNameOnCard & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputCreditCardAddress = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardAddress & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputCreditCardPostalCode = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoInputCommercialCardCode = '" & str2SrcQB_CreditCardTxnInfoInputCommercialCardCode & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultResultCode = " & str2SrcQB_CreditCardTxnInfoResultResultCode & "" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultResultMessage = '" & str2SrcQB_CreditCardTxnInfoResultResultMessage & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultCreditCardTransID = '" & str2SrcQB_CreditCardTxnInfoResultCreditCardTransID & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultMerchantAccountNumber = '" & str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultAuthorizationCode = '" & str2SrcQB_CreditCardTxnInfoResultAuthorizationCode & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultAVSStreet = '" & str2SrcQB_CreditCardTxnInfoResultAVSStreet & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultAVSZip = '" & str2SrcQB_CreditCardTxnInfoResultAVSZip & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultReconBatchID = '" & str2SrcQB_CreditCardTxnInfoResultReconBatchID & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultPaymentGroupingCode = " & str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode & "" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultPaymentStatus = '" & str2SrcQB_CreditCardTxnInfoResultPaymentStatus & "'" & vbCrLf
                '            strSQL3 = "     , CreditCardTxnInfoResultTxnAuthorizationTime = '" & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime & "'" & vbCrLf & _
                ''                      "     , CreditCardTxnInfoResultTxnAuthorizationStamp = " & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp & "" & vbCrLf & _
                ''                      "     , IsAutoApply = '" & str2SrcQB_IsAutoApply & "'" & vbCrLf & _
                ''                      "     , UnusedPayment = " & str2SrcQB_UnusedPayment & "" & vbCrLf & _
                ''                      "     , UnusedCredits = " & str2SrcQB_UnusedCredits & "" & vbCrLf & _
                ''                      "     , AppliedToTxnTxnID = '" & str2SrcQB_AppliedToTxnTxnID & "'" & vbCrLf & _
                ''                      "     , AppliedToTxnPaymentAmount = " & str2SrcQB_AppliedToTxnPaymentAmount & "" & vbCrLf & _
                ''                      "     , AppliedToTxnTxnType = '" & str2SrcQB_AppliedToTxnTxnType & "'" & vbCrLf & _
                ''                      "     , AppliedToTxnTxnDate = '" & str2SrcQB_AppliedToTxnTxnDate & "'" & vbCrLf & _
                ''                      "     , AppliedToTxnRefNumber = '" & str2SrcQB_AppliedToTxnRefNumber & "'" & vbCrLf & _
                ''                      "     , AppliedToTxnBalanceRemaining = " & str2SrcQB_AppliedToTxnBalanceRemaining & "" & vbCrLf & _
                ''                      "     , AppliedToTxnAmount = " & str2SrcQB_AppliedToTxnAmount & "" & vbCrLf & _
                ''                      "     , AppliedToTxnSetCreditCreditTxnID = '" & str2SrcQB_AppliedToTxnSetCreditCreditTxnID & "'" & vbCrLf & _
                ''                      "     , AppliedToTxnSetCreditAppliedAmount = " & str2SrcQB_AppliedToTxnSetCreditAppliedAmount & "" & vbCrLf & _
                ''                      "     , AppliedToTxnDiscountAmount = " & str2SrcQB_AppliedToTxnDiscountAmount & "" & vbCrLf & _
                ''                      "     , AppliedToTxnDiscountAccountRefListID = '" & str2SrcQB_AppliedToTxnDiscountAccountRefListID & "'" & vbCrLf & _
                ''                      "     , AppliedToTxnDiscountAccountRefFullName = '" & str2SrcQB_AppliedToTxnDiscountAccountRefFullName & "'" & vbCrLf & _
                ''                      "     , FQSaveToCache = '" & str2SrcQB_FQSaveToCache & "'" & vbCrLf & _
                ''                      "     , FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & vbCrLf & _
                ''                      "WHERE " & vbCrLf & _
                ''                      "       FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & vbCrLf
                '
                '            'Combine the strings
                '            strTableUpdate = strSQL1 & strSQL2 & strSQL3 '& strSQL4 & strSQL5 & strSQL6
                '            'Debug.Print strTableUpdate
                '
                '            'Execute the insert
                '            '*cnDBPM.Execute strTableUpdate
                '            cnmax.Execute strTableUpdate
                '
                '
                '
                '
                '        Else 'record not exist  -INSERT
                '            'DO INSERT WORK:
                '            Debug.Print "INSERT"

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_ReceivePaymentLine " & Environment.NewLine & _
                          "   ( TxnID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , TxnNumber " & Environment.NewLine & _
                          "   , CustomerRefListID " & Environment.NewLine & _
                          "   , CustomerRefFullName " & Environment.NewLine & _
                          "   , ARAccountRefListID " & Environment.NewLine & _
                          "   , ARAccountRefFullName " & Environment.NewLine & _
                          "   , TxnDate " & Environment.NewLine & _
                          "   , TxnDateMacro " & Environment.NewLine & _
                          "   , RefNumber " & Environment.NewLine & _
                          "   , TotalAmount " & Environment.NewLine & _
                          "   , PaymentMethodRefListID " & Environment.NewLine & _
                          "   , PaymentMethodRefFullName " & Environment.NewLine & _
                          "   , Memo " & Environment.NewLine & _
                          "   , DepositToAccountRefListID " & Environment.NewLine
                strSQL2 = "   , DepositToAccountRefFullName " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputCreditCardNumber " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputExpirationMonth " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputExpirationYear " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputNameOnCard " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputCreditCardAddress " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputCreditCardPostalCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoInputCommercialCardCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultResultCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultResultMessage " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultCreditCardTransID " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultMerchantAccountNumber " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultAuthorizationCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultAVSStreet " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultAVSZip " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultReconBatchID " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultPaymentGroupingCode " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultPaymentStatus " & Environment.NewLine
                strSQL3 = "   , CreditCardTxnInfoResultTxnAuthorizationTime " & Environment.NewLine & _
                          "   , CreditCardTxnInfoResultTxnAuthorizationStamp " & Environment.NewLine & _
                          "   , IsAutoApply " & Environment.NewLine & _
                          "   , UnusedPayment " & Environment.NewLine & _
                          "   , UnusedCredits " & Environment.NewLine & _
                          "   , AppliedToTxnTxnID " & Environment.NewLine & _
                          "   , AppliedToTxnPaymentAmount " & Environment.NewLine & _
                          "   , AppliedToTxnTxnType " & Environment.NewLine & _
                          "   , AppliedToTxnTxnDate " & Environment.NewLine & _
                          "   , AppliedToTxnRefNumber " & Environment.NewLine & _
                          "   , AppliedToTxnBalanceRemaining " & Environment.NewLine & _
                          "   , AppliedToTxnAmount " & Environment.NewLine & _
                          "   , AppliedToTxnSetCreditCreditTxnID " & Environment.NewLine & _
                          "   , AppliedToTxnSetCreditAppliedAmount " & Environment.NewLine & _
                          "   , AppliedToTxnDiscountAmount " & Environment.NewLine & _
                          "   , AppliedToTxnDiscountAccountRefListID " & Environment.NewLine & _
                          "   , AppliedToTxnDiscountAccountRefFullName " & Environment.NewLine & _
                          "   , FQSaveToCache " & Environment.NewLine & _
                          "   , FQPrimaryKey ) " & Environment.NewLine
                strSQL4 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_TxnID & "'  --TxnID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , " & str2SrcQB_TxnNumber & "  --TxnNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerRefListID & "'  --CustomerRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerRefFullName & "'  --CustomerRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ARAccountRefListID & "'  --ARAccountRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ARAccountRefFullName & "'  --ARAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TxnDate & "'  --TxnDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TxnDateMacro & "'  --TxnDateMacro" & Environment.NewLine & _
                          "   , '" & str2SrcQB_RefNumber & "'  --RefNumber" & Environment.NewLine & _
                          "   , " & str2SrcQB_TotalAmount & "  --TotalAmount" & Environment.NewLine & _
                          "   , '" & str2SrcQB_PaymentMethodRefListID & "'  --PaymentMethodRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_PaymentMethodRefFullName & "'  --PaymentMethodRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Memo & "'  --Memo" & Environment.NewLine & _
                          "   , '" & str2SrcQB_DepositToAccountRefListID & "'  --DepositToAccountRefListID" & Environment.NewLine
                strSQL5 = "   , '" & str2SrcQB_DepositToAccountRefFullName & "'  --DepositToAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputCreditCardNumber & "'  --CreditCardTxnInfoInputCreditCardNumber" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoInputExpirationMonth & "  --CreditCardTxnInfoInputExpirationMonth" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoInputExpirationYear & "  --CreditCardTxnInfoInputExpirationYear" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputNameOnCard & "'  --CreditCardTxnInfoInputNameOnCard" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputCreditCardAddress & "'  --CreditCardTxnInfoInputCreditCardAddress" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode & "'  --CreditCardTxnInfoInputCreditCardPostalCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoInputCommercialCardCode & "'  --CreditCardTxnInfoInputCommercialCardCode" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoResultResultCode & "  --CreditCardTxnInfoResultResultCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultResultMessage & "'  --CreditCardTxnInfoResultResultMessage" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultCreditCardTransID & "'  --CreditCardTxnInfoResultCreditCardTransID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber & "'  --CreditCardTxnInfoResultMerchantAccountNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultAuthorizationCode & "'  --CreditCardTxnInfoResultAuthorizationCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultAVSStreet & "'  --CreditCardTxnInfoResultAVSStreet" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultAVSZip & "'  --CreditCardTxnInfoResultAVSZip" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultReconBatchID & "'  --CreditCardTxnInfoResultReconBatchID" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode & "  --CreditCardTxnInfoResultPaymentGroupingCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditCardTxnInfoResultPaymentStatus & "'  --CreditCardTxnInfoResultPaymentStatus" & Environment.NewLine
                strSQL6 = "   , '" & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime & "'  --CreditCardTxnInfoResultTxnAuthorizationTime" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp & "  --CreditCardTxnInfoResultTxnAuthorizationStamp" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsAutoApply & "'  --IsAutoApply" & Environment.NewLine & _
                          "   , " & str2SrcQB_UnusedPayment & "  --UnusedPayment" & Environment.NewLine & _
                          "   , " & str2SrcQB_UnusedCredits & "  --UnusedCredits" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AppliedToTxnTxnID & "'  --AppliedToTxnTxnID" & Environment.NewLine & _
                          "   , " & str2SrcQB_AppliedToTxnPaymentAmount & "  --AppliedToTxnPaymentAmount" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AppliedToTxnTxnType & "'  --AppliedToTxnTxnType" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AppliedToTxnTxnDate & "'  --AppliedToTxnTxnDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AppliedToTxnRefNumber & "'  --AppliedToTxnRefNumber" & Environment.NewLine & _
                          "   , " & str2SrcQB_AppliedToTxnBalanceRemaining & "  --AppliedToTxnBalanceRemaining" & Environment.NewLine & _
                          "   , " & str2SrcQB_AppliedToTxnAmount & "  --AppliedToTxnAmount" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AppliedToTxnSetCreditCreditTxnID & "'  --AppliedToTxnSetCreditCreditTxnID" & Environment.NewLine & _
                          "   , " & str2SrcQB_AppliedToTxnSetCreditAppliedAmount & "  --AppliedToTxnSetCreditAppliedAmount" & Environment.NewLine & _
                          "   , " & str2SrcQB_AppliedToTxnDiscountAmount & "  --AppliedToTxnDiscountAmount" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AppliedToTxnDiscountAccountRefListID & "'  --AppliedToTxnDiscountAccountRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AppliedToTxnDiscountAccountRefFullName & "'  --AppliedToTxnDiscountAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_FQSaveToCache & "'  --FQSaveToCache" & Environment.NewLine & _
                          "   , '" & str2SrcQB_FQPrimaryKey & "' ) --FQPrimaryKey" & Environment.NewLine

                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
                'Debug.Print strTableInsert

                '            'Execute the insert
                '            '*cnDBPM.Execute strTableInsert
                '            cnMax.Execute strTableInsert

                'Execute the insert
                If gstrCompany = "DrummondPrinting" Then
                    '*cnDBPM.Execute strTableInsert
                    Dim TempCommand_3 As SqlCommand
                    TempCommand_3 = cnMax.CreateCommand()
                    TempCommand_3.CommandText = strTableInsert
                    TempCommand_3.ExecuteNonQuery()
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    Dim TempCommand_4 As SqlCommand
                    TempCommand_4 = cnMax.CreateCommand()
                    TempCommand_4.CommandText = strTableInsert
                    TempCommand_4.ExecuteNonQuery()
                End If

                '
                '        End If
                '


            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing
            'frmMain.lstConversionProgress.AddItem "" & Now & "   Processing  0  QB_ReceivePaymentLine  Records "
            'frmMain.lblListboxStatus.Caption = "" & Now & "   Processing  QB_ReceivePaymentLine  Records "
            'DoEvents

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '        End If
        End If



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QBTable.Close()
            rs1MaxOfCopy_QBTable = Nothing


            rs2SrcQB_QB_ReceivePaymentLine = Nothing


            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QBTable.Close()
            rs3TestID_QBTable = Nothing



            Exit Sub


            MessageBox.Show("<<RefreshQB_ReceivePaymentLine>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub


    Public Sub ReloadQB_CreditMemo()
        Dim rs1MaxOfCopy_QBTable, rs3TestID_QBTable As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_CreditMemo" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:




        ''FOR PART 1MaxOfCopy_ - Get records from QBTable
        'Debug.Print "List1MaxOfCopy_QBTable"
        'Dim rs1MaxOfCopy_QBTable As ADODB.Recordset
        'Dim str1MaxOfCopy_QBTableSQL As String
        'Dim str1MaxOfCopy_QBTableRow As String
        'Dim str1MaxOfCopy_TimeModified As String
        ''This routine gets the 1MaxOfCopy_QBTable from the database according to the selection in str1MaxOfCopy_QBTableSQL.
        ''It then puts those 1MaxOfCopy_QBTable in the list box

        'FOR PART 2SrcQB_ - Get records from QB_CreditMemo
        Debug.WriteLine("List2SrcQB_QB_CreditMemo")
        Dim rs2SrcQB_QB_CreditMemo As DataSet
        Dim str2SrcQB_QB_CreditMemoSQL, str2SrcQB_QB_CreditMemoRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ClassRefListID, str2SrcQB_ClassRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TemplateRefListID, str2SrcQB_TemplateRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_BillAddressAddr1, str2SrcQB_BillAddressAddr2, str2SrcQB_BillAddressAddr3, str2SrcQB_BillAddressAddr4, str2SrcQB_BillAddressCity, str2SrcQB_BillAddressState, str2SrcQB_BillAddressPostalCode, str2SrcQB_BillAddressCountry, str2SrcQB_ShipAddressAddr1, str2SrcQB_ShipAddressAddr2, str2SrcQB_ShipAddressAddr3, str2SrcQB_ShipAddressAddr4, str2SrcQB_ShipAddressCity, str2SrcQB_ShipAddressState, str2SrcQB_ShipAddressPostalCode, str2SrcQB_ShipAddressCountry, str2SrcQB_IsPending, str2SrcQB_PONumber, str2SrcQB_TermsRefListID, str2SrcQB_TermsRefFullName, str2SrcQB_DueDate, str2SrcQB_SalesRepRefListID, str2SrcQB_SalesRepRefFullName, str2SrcQB_FOB, str2SrcQB_ShipDate, str2SrcQB_ShipMethodRefListID, str2SrcQB_ShipMethodRefFullName, str2SrcQB_Subtotal, str2SrcQB_ItemSalesTaxRefListID, str2SrcQB_ItemSalesTaxRefFullName, str2SrcQB_SalesTaxPercentage, str2SrcQB_SalesTaxTotal, str2SrcQB_TotalAmount, str2SrcQB_CreditRemaining, str2SrcQB_Memo, str2SrcQB_CustomerMsgRefListID, str2SrcQB_CustomerMsgRefFullName, str2SrcQB_IsToBePrinted, str2SrcQB_CustomerSalesTaxCodeRefListID, str2SrcQB_CustomerSalesTaxCodeRefFullName, str2SrcQB_CustomFieldOther As String
        'This routine gets the 2SrcQB_QB_CreditMemo from the database according to the selection in str2SrcQB_QB_CreditMemoSQL.
        'It then puts those 2SrcQB_QB_CreditMemo in the list box

        ''FOR PART 3TestID_
        'Debug.Print "List3TestID_QBTable"
        'Dim rs3TestID_QBTable As ADODB.Recordset
        'Dim str3TestID_QBTableSQL As String
        'Dim str3TestID_QBTableRow As String
        'Dim str3TestID_ListID As String
        ''This routine gets the 3TestID_QBTable from the database according to the selection in str3TestID_QBTableSQL.
        ''It then puts those 3TestID_QBTable in the list box

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_CreditMemo  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_CreditMemo"
        Application.DoEvents()


        '
        ''Clear out table
        ''*cnDBPM.Execute "DELETE FROM QB_CreditMemo"
        'cnmax.Execute "DELETE FROM QB_CreditMemo"
        '

        'Get rs from QB
        'Load table from rs

        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        rs2SrcQB_QB_CreditMemo = New DataSet() '*** TAKE QB_ OFF OF TABLE NAME ***
        'str2SrcQB_QB_CreditMemoSQL = "SELECT * FROM CreditMemo WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_CreditMemoSQL = "SELECT * FROM CreditMemo WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_CreditMemoSQL = "SELECT * FROM CreditMemo WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'str2SrcQB_QB_CreditMemoSQL = "SELECT * FROM CreditMemo WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_CreditMemo & "'} ORDER BY TimeModified"
        'THIS worked:  str2SrcQB_QB_CreditMemoSQL = "SELECT * FROM CreditMemo WHERE TxnID = '78E53-1145544373'"
        str2SrcQB_QB_CreditMemoSQL = "SELECT * FROM CreditMemo"
        'Debug.Print str2SrcQB_QB_CreditMemoSQL
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_CreditMemoSQL, cnQuickBooks)
        rs2SrcQB_QB_CreditMemo.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_CreditMemo) ', adAsyncFetch
        If rs2SrcQB_QB_CreditMemo.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_CreditMemo"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_CreditMemo"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_CreditMemo"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_CreditMemo.Tables(0).Rows.Count) & "  QB_CreditMemo  Records  ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_CreditMemo.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_CreditMemo.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_CreditMemo.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_CreditMemo.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_TxnID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_TxnNumber = ""
                str2SrcQB_CustomerRefListID = ""
                str2SrcQB_CustomerRefFullName = ""
                str2SrcQB_ClassRefListID = ""
                str2SrcQB_ClassRefFullName = ""
                str2SrcQB_ARAccountRefListID = ""
                str2SrcQB_ARAccountRefFullName = ""
                str2SrcQB_TemplateRefListID = ""
                str2SrcQB_TemplateRefFullName = ""
                str2SrcQB_TxnDate = ""
                str2SrcQB_TxnDateMacro = ""
                str2SrcQB_RefNumber = ""
                str2SrcQB_BillAddressAddr1 = ""
                str2SrcQB_BillAddressAddr2 = ""
                str2SrcQB_BillAddressAddr3 = ""
                str2SrcQB_BillAddressAddr4 = ""
                str2SrcQB_BillAddressCity = ""
                str2SrcQB_BillAddressState = ""
                str2SrcQB_BillAddressPostalCode = ""
                str2SrcQB_BillAddressCountry = ""
                str2SrcQB_ShipAddressAddr1 = ""
                str2SrcQB_ShipAddressAddr2 = ""
                str2SrcQB_ShipAddressAddr3 = ""
                str2SrcQB_ShipAddressAddr4 = ""
                str2SrcQB_ShipAddressCity = ""
                str2SrcQB_ShipAddressState = ""
                str2SrcQB_ShipAddressPostalCode = ""
                str2SrcQB_ShipAddressCountry = ""
                str2SrcQB_IsPending = ""
                str2SrcQB_PONumber = ""
                str2SrcQB_TermsRefListID = ""
                str2SrcQB_TermsRefFullName = ""
                str2SrcQB_DueDate = ""
                str2SrcQB_SalesRepRefListID = ""
                str2SrcQB_SalesRepRefFullName = ""
                str2SrcQB_FOB = ""
                str2SrcQB_ShipDate = ""
                str2SrcQB_ShipMethodRefListID = ""
                str2SrcQB_ShipMethodRefFullName = ""
                str2SrcQB_Subtotal = ""
                str2SrcQB_ItemSalesTaxRefListID = ""
                str2SrcQB_ItemSalesTaxRefFullName = ""
                str2SrcQB_SalesTaxPercentage = ""
                str2SrcQB_SalesTaxTotal = ""
                str2SrcQB_TotalAmount = ""
                str2SrcQB_CreditRemaining = ""
                str2SrcQB_Memo = ""
                str2SrcQB_CustomerMsgRefListID = ""
                str2SrcQB_CustomerMsgRefFullName = ""
                str2SrcQB_IsToBePrinted = ""
                str2SrcQB_CustomerSalesTaxCodeRefListID = ""
                str2SrcQB_CustomerSalesTaxCodeRefFullName = ""
                str2SrcQB_CustomFieldOther = ""

                'get the columns from the database
                If iteration_row("TxnID") <> "" Then str2SrcQB_TxnID = iteration_row("TxnID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("TxnNumber") <> "" Then str2SrcQB_TxnNumber = iteration_row("TxnNumber")
                If iteration_row("CustomerRefListID") <> "" Then str2SrcQB_CustomerRefListID = iteration_row("CustomerRefListID")
                If iteration_row("CustomerRefFullName") <> "" Then str2SrcQB_CustomerRefFullName = iteration_row("CustomerRefFullName")
                If iteration_row("ClassRefListID") <> "" Then str2SrcQB_ClassRefListID = iteration_row("ClassRefListID")
                If iteration_row("ClassRefFullName") <> "" Then str2SrcQB_ClassRefFullName = iteration_row("ClassRefFullName")
                If iteration_row("ARAccountRefListID") <> "" Then str2SrcQB_ARAccountRefListID = iteration_row("ARAccountRefListID")
                If iteration_row("ARAccountRefFullName") <> "" Then str2SrcQB_ARAccountRefFullName = iteration_row("ARAccountRefFullName")
                If iteration_row("TemplateRefListID") <> "" Then str2SrcQB_TemplateRefListID = iteration_row("TemplateRefListID")
                If iteration_row("TemplateRefFullName") <> "" Then str2SrcQB_TemplateRefFullName = iteration_row("TemplateRefFullName")
                If iteration_row("TxnDate") <> "" Then str2SrcQB_TxnDate = iteration_row("TxnDate")
                If iteration_row("TxnDateMacro") <> "" Then str2SrcQB_TxnDateMacro = iteration_row("TxnDateMacro")
                If iteration_row("RefNumber") <> "" Then str2SrcQB_RefNumber = iteration_row("RefNumber")
                If iteration_row("BillAddressAddr1") <> "" Then str2SrcQB_BillAddressAddr1 = iteration_row("BillAddressAddr1")
                If iteration_row("BillAddressAddr2") <> "" Then str2SrcQB_BillAddressAddr2 = iteration_row("BillAddressAddr2")
                If iteration_row("BillAddressAddr3") <> "" Then str2SrcQB_BillAddressAddr3 = iteration_row("BillAddressAddr3")
                If iteration_row("BillAddressAddr4") <> "" Then str2SrcQB_BillAddressAddr4 = iteration_row("BillAddressAddr4")
                If iteration_row("BillAddressCity") <> "" Then str2SrcQB_BillAddressCity = iteration_row("BillAddressCity")
                If iteration_row("BillAddressState") <> "" Then str2SrcQB_BillAddressState = iteration_row("BillAddressState")
                If iteration_row("BillAddressPostalCode") <> "" Then str2SrcQB_BillAddressPostalCode = iteration_row("BillAddressPostalCode")
                If iteration_row("BillAddressCountry") <> "" Then str2SrcQB_BillAddressCountry = iteration_row("BillAddressCountry")
                If iteration_row("ShipAddressAddr1") <> "" Then str2SrcQB_ShipAddressAddr1 = iteration_row("ShipAddressAddr1")
                If iteration_row("ShipAddressAddr2") <> "" Then str2SrcQB_ShipAddressAddr2 = iteration_row("ShipAddressAddr2")
                If iteration_row("ShipAddressAddr3") <> "" Then str2SrcQB_ShipAddressAddr3 = iteration_row("ShipAddressAddr3")
                If iteration_row("ShipAddressAddr4") <> "" Then str2SrcQB_ShipAddressAddr4 = iteration_row("ShipAddressAddr4")
                If iteration_row("ShipAddressCity") <> "" Then str2SrcQB_ShipAddressCity = iteration_row("ShipAddressCity")
                If iteration_row("ShipAddressState") <> "" Then str2SrcQB_ShipAddressState = iteration_row("ShipAddressState")
                If iteration_row("ShipAddressPostalCode") <> "" Then str2SrcQB_ShipAddressPostalCode = iteration_row("ShipAddressPostalCode")
                If iteration_row("ShipAddressCountry") <> "" Then str2SrcQB_ShipAddressCountry = iteration_row("ShipAddressCountry")
                If iteration_row("IsPending") <> "" Then str2SrcQB_IsPending = iteration_row("IsPending")
                If iteration_row("PONumber") <> "" Then str2SrcQB_PONumber = iteration_row("PONumber")
                If iteration_row("TermsRefListID") <> "" Then str2SrcQB_TermsRefListID = iteration_row("TermsRefListID")
                If iteration_row("TermsRefFullName") <> "" Then str2SrcQB_TermsRefFullName = iteration_row("TermsRefFullName")
                If iteration_row("DueDate") <> "" Then str2SrcQB_DueDate = iteration_row("DueDate")
                If iteration_row("SalesRepRefListID") <> "" Then str2SrcQB_SalesRepRefListID = iteration_row("SalesRepRefListID")
                If iteration_row("SalesRepRefFullName") <> "" Then str2SrcQB_SalesRepRefFullName = iteration_row("SalesRepRefFullName")
                If iteration_row("FOB") <> "" Then str2SrcQB_FOB = iteration_row("FOB")
                If iteration_row("ShipDate") <> "" Then str2SrcQB_ShipDate = iteration_row("ShipDate")
                If iteration_row("ShipMethodRefListID") <> "" Then str2SrcQB_ShipMethodRefListID = iteration_row("ShipMethodRefListID")
                If iteration_row("ShipMethodRefFullName") <> "" Then str2SrcQB_ShipMethodRefFullName = iteration_row("ShipMethodRefFullName")
                If iteration_row("Subtotal") <> "" Then str2SrcQB_Subtotal = iteration_row("Subtotal")
                If iteration_row("ItemSalesTaxRefListID") <> "" Then str2SrcQB_ItemSalesTaxRefListID = iteration_row("ItemSalesTaxRefListID")
                If iteration_row("ItemSalesTaxRefFullName") <> "" Then str2SrcQB_ItemSalesTaxRefFullName = iteration_row("ItemSalesTaxRefFullName")
                If iteration_row("SalesTaxPercentage") <> "" Then str2SrcQB_SalesTaxPercentage = iteration_row("SalesTaxPercentage")
                If iteration_row("SalesTaxTotal") <> "" Then str2SrcQB_SalesTaxTotal = iteration_row("SalesTaxTotal")
                If iteration_row("TotalAmount") <> "" Then str2SrcQB_TotalAmount = iteration_row("TotalAmount")
                If iteration_row("CreditRemaining") <> "" Then str2SrcQB_CreditRemaining = iteration_row("CreditRemaining")
                If iteration_row("Memo") <> "" Then str2SrcQB_Memo = iteration_row("Memo")
                If iteration_row("CustomerMsgRefListID") <> "" Then str2SrcQB_CustomerMsgRefListID = iteration_row("CustomerMsgRefListID")
                If iteration_row("CustomerMsgRefFullName") <> "" Then str2SrcQB_CustomerMsgRefFullName = iteration_row("CustomerMsgRefFullName")
                If iteration_row("IsToBePrinted") <> "" Then str2SrcQB_IsToBePrinted = iteration_row("IsToBePrinted")
                If iteration_row("CustomerSalesTaxCodeRefListID") <> "" Then str2SrcQB_CustomerSalesTaxCodeRefListID = iteration_row("CustomerSalesTaxCodeRefListID")
                If iteration_row("CustomerSalesTaxCodeRefFullName") <> "" Then str2SrcQB_CustomerSalesTaxCodeRefFullName = iteration_row("CustomerSalesTaxCodeRefFullName")
                If iteration_row("CustomFieldOther") <> "" Then str2SrcQB_CustomFieldOther = iteration_row("CustomFieldOther")

                'Strip quote character out of strings
                'Get quote characters out!
                'Change Quote to reverse quote
                'If KeyAscii = 39 Then KeyAscii = 96
                str2SrcQB_TxnID = str2SrcQB_TxnID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_TxnNumber = str2SrcQB_TxnNumber.Replace("'"c, "`"c)
                str2SrcQB_CustomerRefListID = str2SrcQB_CustomerRefListID.Replace("'"c, "`"c)
                str2SrcQB_CustomerRefFullName = str2SrcQB_CustomerRefFullName.Replace("'"c, "`"c)
                str2SrcQB_ClassRefListID = str2SrcQB_ClassRefListID.Replace("'"c, "`"c)
                str2SrcQB_ClassRefFullName = str2SrcQB_ClassRefFullName.Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefListID = str2SrcQB_ARAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefFullName = str2SrcQB_ARAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_TemplateRefListID = str2SrcQB_TemplateRefListID.Replace("'"c, "`"c)
                str2SrcQB_TemplateRefFullName = str2SrcQB_TemplateRefFullName.Replace("'"c, "`"c)
                str2SrcQB_TxnDate = str2SrcQB_TxnDate.Replace("'"c, "`"c)
                str2SrcQB_TxnDateMacro = str2SrcQB_TxnDateMacro.Replace("'"c, "`"c)
                str2SrcQB_RefNumber = str2SrcQB_RefNumber.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr1 = str2SrcQB_BillAddressAddr1.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr2 = str2SrcQB_BillAddressAddr2.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr3 = str2SrcQB_BillAddressAddr3.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr4 = str2SrcQB_BillAddressAddr4.Replace("'"c, "`"c)
                str2SrcQB_BillAddressCity = str2SrcQB_BillAddressCity.Replace("'"c, "`"c)
                str2SrcQB_BillAddressState = str2SrcQB_BillAddressState.Replace("'"c, "`"c)
                str2SrcQB_BillAddressPostalCode = str2SrcQB_BillAddressPostalCode.Replace("'"c, "`"c)
                str2SrcQB_BillAddressCountry = str2SrcQB_BillAddressCountry.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr1 = str2SrcQB_ShipAddressAddr1.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr2 = str2SrcQB_ShipAddressAddr2.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr3 = str2SrcQB_ShipAddressAddr3.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr4 = str2SrcQB_ShipAddressAddr4.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressCity = str2SrcQB_ShipAddressCity.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressState = str2SrcQB_ShipAddressState.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressPostalCode = str2SrcQB_ShipAddressPostalCode.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressCountry = str2SrcQB_ShipAddressCountry.Replace("'"c, "`"c)
                str2SrcQB_IsPending = str2SrcQB_IsPending.Replace("'"c, "`"c)
                str2SrcQB_PONumber = str2SrcQB_PONumber.Replace("'"c, "`"c)
                str2SrcQB_TermsRefListID = str2SrcQB_TermsRefListID.Replace("'"c, "`"c)
                str2SrcQB_TermsRefFullName = str2SrcQB_TermsRefFullName.Replace("'"c, "`"c)
                str2SrcQB_DueDate = str2SrcQB_DueDate.Replace("'"c, "`"c)
                str2SrcQB_SalesRepRefListID = str2SrcQB_SalesRepRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesRepRefFullName = str2SrcQB_SalesRepRefFullName.Replace("'"c, "`"c)
                str2SrcQB_FOB = str2SrcQB_FOB.Replace("'"c, "`"c)
                str2SrcQB_ShipDate = str2SrcQB_ShipDate.Replace("'"c, "`"c)
                str2SrcQB_ShipMethodRefListID = str2SrcQB_ShipMethodRefListID.Replace("'"c, "`"c)
                str2SrcQB_ShipMethodRefFullName = str2SrcQB_ShipMethodRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Subtotal = str2SrcQB_Subtotal.Replace("'"c, "`"c)
                str2SrcQB_ItemSalesTaxRefListID = str2SrcQB_ItemSalesTaxRefListID.Replace("'"c, "`"c)
                str2SrcQB_ItemSalesTaxRefFullName = str2SrcQB_ItemSalesTaxRefFullName.Replace("'"c, "`"c)
                str2SrcQB_SalesTaxPercentage = str2SrcQB_SalesTaxPercentage.Replace("'"c, "`"c)
                str2SrcQB_SalesTaxTotal = str2SrcQB_SalesTaxTotal.Replace("'"c, "`"c)
                str2SrcQB_TotalAmount = str2SrcQB_TotalAmount.Replace("'"c, "`"c)
                str2SrcQB_CreditRemaining = str2SrcQB_CreditRemaining.Replace("'"c, "`"c)
                str2SrcQB_Memo = str2SrcQB_Memo.Replace("'"c, "`"c)
                str2SrcQB_CustomerMsgRefListID = str2SrcQB_CustomerMsgRefListID.Replace("'"c, "`"c)
                str2SrcQB_CustomerMsgRefFullName = str2SrcQB_CustomerMsgRefFullName.Replace("'"c, "`"c)
                str2SrcQB_IsToBePrinted = str2SrcQB_IsToBePrinted.Replace("'"c, "`"c)
                str2SrcQB_CustomerSalesTaxCodeRefListID = str2SrcQB_CustomerSalesTaxCodeRefListID.Replace("'"c, "`"c)
                str2SrcQB_CustomerSalesTaxCodeRefFullName = str2SrcQB_CustomerSalesTaxCodeRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldOther = str2SrcQB_CustomFieldOther.Replace("'"c, "`"c)


                'Change flags back to binary
                '        If str2SrcQB_IsActive = "True" Then str2SrcQB_IsActive = "1" Else str2SrcQB_IsActive = "0"
                '        If str2SrcQB_IsPending = "True" Then str2SrcQB_IsPending = "1" Else str2SrcQB_IsPending = "0"
                '        If str2SrcQB_IsFinanceCharge = "True" Then str2SrcQB_IsFinanceCharge = "1" Else str2SrcQB_IsFinanceCharge = "0"
                '        If str2SrcQB_IsPaid = "True" Then str2SrcQB_IsPaid = "1" Else str2SrcQB_IsPaid = "0"
                '        If str2SrcQB_IsToBePrinted = "True" Then str2SrcQB_IsToBePrinted = "1" Else str2SrcQB_IsToBePrinted = "0"

                str2SrcQB_IsPending = IIf(str2SrcQB_IsPending = "True", "1", "0")
                str2SrcQB_IsToBePrinted = IIf(str2SrcQB_IsToBePrinted = "True", "1", "0")


                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_CreditMemoRow = "" & _
                                             Strings.Left(str2SrcQB_TxnID & "                  ", 18) & "   " & _
                                             Strings.Left(str2SrcQB_TimeCreated & "                  ", 16) & "   " & _
                                             Strings.Left(str2SrcQB_TimeModified & "                  ", 16) & "   " & _
                                             Strings.Left(str2SrcQB_CustomerRefFullName & "                  ", 18) & "   " & _
                                             Strings.Left(str2SrcQB_TxnDate & "                  ", 10) & "   " & _
                                             Strings.Left(str2SrcQB_RefNumber & "                  ", 10) & "   " & _
                                             Strings.Left(str2SrcQB_TxnDateMacro & "                  ", 10) & "   " & _
                                             Strings.Left(str2SrcQB_Subtotal & "                  ", 18) & "   " & _
                                             Strings.Left(str2SrcQB_TxnDateMacro & "                  ", 18) & "   " & _
                                             Strings.Left(str2SrcQB_TxnDateMacro & "                  ", 18) & "   " & _
                                             "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_CreditMemo.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_CreditMemo.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_CreditMemo.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_CreditMemoRow)
                    'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_TxnID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record

                '        'Update cust balances
                '        UpdateQBCustomerBalance (str2SrcQB_CustomerRefListID)
                '
                '        'Update inv info
                '        'UpdateQBInvoice (str2SrcQB_TxnID)
                '        UpdateQBInvoice (str2SrcQB_CustomerRefFullName)
                '
                '
                '
                '        'ADD CreditMemo LINE STUFF HERE
                '
                '
                '
                '
                '
                '
                '
                '        'Check to see if ListID is in QBTable            'Yes then UPDATE record            'No then INSERT record
                '        'New recordset
                '        Set rs3TestID_QBTable = New ADODB.Recordset
                '        str3TestID_QBTableSQL = "SELECT TxnID FROM QB_CreditMemo WHERE TxnID = '" & str2SrcQB_TxnID & "'"
                '        'Debug.Print str3TestID_QBTableSQL
                '        'rs3TestID_QBTable.Open str3TestID_QBTableSQL, cnDBPM, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        rs3TestID_QBTable.Open str3TestID_QBTableSQL, cnmax, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        If rs3TestID_QBTable.RecordCount > 1 Then Stop 'Should only be one
                '        If rs3TestID_QBTable.RecordCount > 0 Then  'record exists  -UPDATE
                '            'DO UPDATE WORK:
                '            Debug.Print "UPDATE"
                '
                '            'Build the SQL string
                '            strSQL1 = "UPDATE  " & vbCrLf & _
                ''                      "       QB_CreditMemo " & vbCrLf & _
                ''                      "SET " & vbCrLf & _
                ''                      "       TxnID = '" & str2SrcQB_TxnID & "'" & vbCrLf & _
                ''                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & vbCrLf & _
                ''                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & vbCrLf & _
                ''                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & vbCrLf & _
                ''                      "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & vbCrLf & _
                ''                      "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & vbCrLf & _
                ''                      "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & vbCrLf & _
                ''                      "     , ClassRefListID = '" & str2SrcQB_ClassRefListID & "'" & vbCrLf & _
                ''                      "     , ClassRefFullName = '" & str2SrcQB_ClassRefFullName & "'" & vbCrLf & _
                ''                      "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & vbCrLf & _
                ''                      "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & vbCrLf & _
                ''                      "     , TemplateRefListID = '" & str2SrcQB_TemplateRefListID & "'" & vbCrLf & _
                ''                      "     , TemplateRefFullName = '" & str2SrcQB_TemplateRefFullName & "'" & vbCrLf & _
                ''                      "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & vbCrLf & _
                ''                      "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & vbCrLf & _
                ''                      "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr1 = '" & str2SrcQB_BillAddressAddr1 & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr2 = '" & str2SrcQB_BillAddressAddr2 & "'" & vbCrLf
                '            strSQL2 = "     , BillAddressAddr3 = '" & str2SrcQB_BillAddressAddr3 & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr4 = '" & str2SrcQB_BillAddressAddr4 & "'" & vbCrLf & _
                ''                      "     , BillAddressCity = '" & str2SrcQB_BillAddressCity & "'" & vbCrLf & _
                ''                      "     , BillAddressState = '" & str2SrcQB_BillAddressState & "'" & vbCrLf & _
                ''                      "     , BillAddressPostalCode = '" & str2SrcQB_BillAddressPostalCode & "'" & vbCrLf & _
                ''                      "     , BillAddressCountry = '" & str2SrcQB_BillAddressCountry & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr1 = '" & str2SrcQB_ShipAddressAddr1 & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr2 = '" & str2SrcQB_ShipAddressAddr2 & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr3 = '" & str2SrcQB_ShipAddressAddr3 & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr4 = '" & str2SrcQB_ShipAddressAddr4 & "'" & vbCrLf & _
                ''                      "     , ShipAddressCity = '" & str2SrcQB_ShipAddressCity & "'" & vbCrLf & _
                ''                      "     , ShipAddressState = '" & str2SrcQB_ShipAddressState & "'" & vbCrLf & _
                ''                      "     , ShipAddressPostalCode = '" & str2SrcQB_ShipAddressPostalCode & "'" & vbCrLf & _
                ''                      "     , ShipAddressCountry = '" & str2SrcQB_ShipAddressCountry & "'" & vbCrLf & _
                ''                      "     , IsPending = '" & str2SrcQB_IsPending & "'" & vbCrLf & _
                ''                      "     , PONumber = '" & str2SrcQB_PONumber & "'" & vbCrLf & _
                ''                      "     , TermsRefListID = '" & str2SrcQB_TermsRefListID & "'" & vbCrLf & _
                ''                      "     , TermsRefFullName = '" & str2SrcQB_TermsRefFullName & "'" & vbCrLf & _
                ''                      "     , DueDate = '" & str2SrcQB_DueDate & "'" & vbCrLf & _
                ''                      "     , SalesRepRefListID = '" & str2SrcQB_SalesRepRefListID & "'" & vbCrLf & _
                ''                      "     , SalesRepRefFullName = '" & str2SrcQB_SalesRepRefFullName & "'" & vbCrLf
                '            strSQL3 = "     , FOB = '" & str2SrcQB_FOB & "'" & vbCrLf & _
                ''                      "     , ShipDate = '" & str2SrcQB_ShipDate & "'" & vbCrLf & _
                ''                      "     , ShipMethodRefListID = '" & str2SrcQB_ShipMethodRefListID & "'" & vbCrLf & _
                ''                      "     , ShipMethodRefFullName = '" & str2SrcQB_ShipMethodRefFullName & "'" & vbCrLf & _
                ''                      "     , Subtotal = " & str2SrcQB_Subtotal & "" & vbCrLf & _
                ''                      "     , ItemSalesTaxRefListID = '" & str2SrcQB_ItemSalesTaxRefListID & "'" & vbCrLf & _
                ''                      "     , ItemSalesTaxRefFullName = '" & str2SrcQB_ItemSalesTaxRefFullName & "'" & vbCrLf & _
                ''                      "     , SalesTaxPercentage = " & str2SrcQB_SalesTaxPercentage & "" & vbCrLf & _
                ''                      "     , SalesTaxTotal = " & str2SrcQB_SalesTaxTotal & "" & vbCrLf & _
                ''                      "     , TotalAmount = " & str2SrcQB_TotalAmount & "" & vbCrLf & _
                ''                      "     , CreditRemaining = " & str2SrcQB_CreditRemaining & "" & vbCrLf & _
                ''                      "     , Memo = '" & str2SrcQB_Memo & "'" & vbCrLf & _
                ''                      "     , CustomerMsgRefListID = '" & str2SrcQB_CustomerMsgRefListID & "'" & vbCrLf & _
                ''                      "     , CustomerMsgRefFullName = '" & str2SrcQB_CustomerMsgRefFullName & "'" & vbCrLf & _
                ''                      "     , IsToBePrinted = '" & str2SrcQB_IsToBePrinted & "'" & vbCrLf & _
                ''                      "     , CustomerSalesTaxCodeRefListID = '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'" & vbCrLf & _
                ''                      "     , CustomerSalesTaxCodeRefFullName = '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'" & vbCrLf & _
                ''                      "     , CustomFieldOther = '" & str2SrcQB_CustomFieldOther & "'" & vbCrLf & _
                ''                      "WHERE " & vbCrLf & _
                ''                      "       TxnID = '" & str2SrcQB_TxnID & "'" & vbCrLf
                '
                '
                '
                '            'Combine the strings
                '            strTableUpdate = strSQL1 & strSQL2 & strSQL3 '& strSQL4 & strSQL5 & strSQL6
                '            'Debug.Print strTableUpdate
                '
                '            'Execute the insert
                '            '*cnDBPM.Execute strTableUpdate
                '            cnmax.Execute strTableUpdate
                '
                '
                '
                '
                '        Else 'record not exist  -INSERT
                '            'DO INSERT WORK:
                '            Debug.Print "INSERT"

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_CreditMemo " & Environment.NewLine & _
                          "   ( TxnID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , TxnNumber " & Environment.NewLine & _
                          "   , CustomerRefListID " & Environment.NewLine & _
                          "   , CustomerRefFullName " & Environment.NewLine & _
                          "   , ClassRefListID " & Environment.NewLine & _
                          "   , ClassRefFullName " & Environment.NewLine & _
                          "   , ARAccountRefListID " & Environment.NewLine & _
                          "   , ARAccountRefFullName " & Environment.NewLine & _
                          "   , TemplateRefListID " & Environment.NewLine & _
                          "   , TemplateRefFullName " & Environment.NewLine & _
                          "   , TxnDate " & Environment.NewLine & _
                          "   , TxnDateMacro " & Environment.NewLine & _
                          "   , RefNumber " & Environment.NewLine & _
                          "   , BillAddressAddr1 " & Environment.NewLine & _
                          "   , BillAddressAddr2 " & Environment.NewLine
                strSQL2 = "   , BillAddressAddr3 " & Environment.NewLine & _
                          "   , BillAddressAddr4 " & Environment.NewLine & _
                          "   , BillAddressCity " & Environment.NewLine & _
                          "   , BillAddressState " & Environment.NewLine & _
                          "   , BillAddressPostalCode " & Environment.NewLine & _
                          "   , BillAddressCountry " & Environment.NewLine & _
                          "   , ShipAddressAddr1 " & Environment.NewLine & _
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
                          "   , DueDate " & Environment.NewLine & _
                          "   , SalesRepRefListID " & Environment.NewLine & _
                          "   , SalesRepRefFullName " & Environment.NewLine
                strSQL3 = "   , FOB " & Environment.NewLine & _
                          "   , ShipDate " & Environment.NewLine & _
                          "   , ShipMethodRefListID " & Environment.NewLine & _
                          "   , ShipMethodRefFullName " & Environment.NewLine & _
                          "   , Subtotal " & Environment.NewLine & _
                          "   , ItemSalesTaxRefListID " & Environment.NewLine & _
                          "   , ItemSalesTaxRefFullName " & Environment.NewLine & _
                          "   , SalesTaxPercentage " & Environment.NewLine & _
                          "   , SalesTaxTotal " & Environment.NewLine & _
                          "   , TotalAmount " & Environment.NewLine & _
                          "   , CreditRemaining " & Environment.NewLine & _
                          "   , Memo " & Environment.NewLine & _
                          "   , CustomerMsgRefListID " & Environment.NewLine & _
                          "   , CustomerMsgRefFullName " & Environment.NewLine & _
                          "   , IsToBePrinted " & Environment.NewLine & _
                          "   , CustomerSalesTaxCodeRefListID " & Environment.NewLine & _
                          "   , CustomerSalesTaxCodeRefFullName " & Environment.NewLine & _
                          "   , CustomFieldOther ) " & Environment.NewLine
                strSQL4 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_TxnID & "'  --TxnID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , " & str2SrcQB_TxnNumber & "  --TxnNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerRefListID & "'  --CustomerRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerRefFullName & "'  --CustomerRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ClassRefListID & "'  --ClassRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ClassRefFullName & "'  --ClassRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ARAccountRefListID & "'  --ARAccountRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ARAccountRefFullName & "'  --ARAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TemplateRefListID & "'  --TemplateRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TemplateRefFullName & "'  --TemplateRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TxnDate & "'  --TxnDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TxnDateMacro & "'  --TxnDateMacro" & Environment.NewLine & _
                          "   , '" & str2SrcQB_RefNumber & "'  --RefNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr1 & "'  --BillAddressAddr1" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr2 & "'  --BillAddressAddr2" & Environment.NewLine
                strSQL5 = "   , '" & str2SrcQB_BillAddressAddr3 & "'  --BillAddressAddr3" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr4 & "'  --BillAddressAddr4" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressCity & "'  --BillAddressCity" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressState & "'  --BillAddressState" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressPostalCode & "'  --BillAddressPostalCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressCountry & "'  --BillAddressCountry" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr1 & "'  --ShipAddressAddr1" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr2 & "'  --ShipAddressAddr2" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr3 & "'  --ShipAddressAddr3" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr4 & "'  --ShipAddressAddr4" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressCity & "'  --ShipAddressCity" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressState & "'  --ShipAddressState" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressPostalCode & "'  --ShipAddressPostalCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressCountry & "'  --ShipAddressCountry" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsPending & "'  --IsPending" & Environment.NewLine & _
                          "   , '" & str2SrcQB_PONumber & "'  --PONumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TermsRefListID & "'  --TermsRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TermsRefFullName & "'  --TermsRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_DueDate & "'  --DueDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesRepRefListID & "'  --SalesRepRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesRepRefFullName & "'  --SalesRepRefFullName" & Environment.NewLine
                strSQL6 = "   , '" & str2SrcQB_FOB & "'  --FOB" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipDate & "'  --ShipDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipMethodRefListID & "'  --ShipMethodRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipMethodRefFullName & "'  --ShipMethodRefFullName" & Environment.NewLine & _
                          "   , " & str2SrcQB_Subtotal & "  --Subtotal" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ItemSalesTaxRefListID & "'  --ItemSalesTaxRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ItemSalesTaxRefFullName & "'  --ItemSalesTaxRefFullName" & Environment.NewLine & _
                          "   , " & str2SrcQB_SalesTaxPercentage & "  --SalesTaxPercentage" & Environment.NewLine & _
                          "   , " & str2SrcQB_SalesTaxTotal & "  --SalesTaxTotal" & Environment.NewLine & _
                          "   , " & str2SrcQB_TotalAmount & "  --TotalAmount" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditRemaining & "  --CreditRemaining" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Memo & "'  --Memo" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerMsgRefListID & "'  --CustomerMsgRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerMsgRefFullName & "'  --CustomerMsgRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsToBePrinted & "'  --IsToBePrinted" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'  --CustomerSalesTaxCodeRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'  --CustomerSalesTaxCodeRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldOther & "' )  --CustomFieldOther" & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
                'Debug.Print strTableInsert

                '            'Execute the insert
                '            '*cnDBPM.Execute strTableInsert
                '            cnMax.Execute strTableInsert

                'Execute the insert
                If gstrCompany = "DrummondPrinting" Then
                    '*cnDBPM.Execute strTableInsert
                    Dim TempCommand_3 As SqlCommand
                    TempCommand_3 = cnMax.CreateCommand()
                    TempCommand_3.CommandText = strTableInsert
                    TempCommand_3.ExecuteNonQuery()
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    Dim TempCommand_4 As SqlCommand
                    TempCommand_4 = cnMax.CreateCommand()
                    TempCommand_4.CommandText = strTableInsert
                    TempCommand_4.ExecuteNonQuery()
                End If

                '
                '        End If
                '

            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)

        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_CreditMemo  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '        End If
        End If


        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QBTable.Close()
            rs1MaxOfCopy_QBTable = Nothing

            rs2SrcQB_QB_CreditMemo = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QBTable.Close()
            rs3TestID_QBTable = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_CreditMemo> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub


    Public Sub ReloadQB_CreditMemoLine()
        Dim rs1MaxOfCopy_QB_CreditMemoLine, rs3TestID_QB_CreditMemoLine As Object
        Dim strSQL7, strSQL8 As String

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_CreditMemoLine" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:




        ''FOR PART 1MaxOfCopy_ - Get records from QBTable
        'Debug.Print "List1MaxOfCopy_QBTable"
        'Dim rs1MaxOfCopy_QBTable As ADODB.Recordset
        'Dim str1MaxOfCopy_QBTableSQL As String
        'Dim str1MaxOfCopy_QBTableRow As String
        'Dim str1MaxOfCopy_TimeModified As String
        ''This routine gets the 1MaxOfCopy_QBTable from the database according to the selection in str1MaxOfCopy_QBTableSQL.
        ''It then puts those 1MaxOfCopy_QBTable in the list box

        'FOR PART 2SrcQB_ - Get records from QB_CreditMemoLine
        Debug.WriteLine("List2SrcQB_QB_CreditMemoLine")
        Dim rs2SrcQB_QB_CreditMemoLine As DataSet
        Dim str2SrcQB_QB_CreditMemoLineSQL, str2SrcQB_QB_CreditMemoLineRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ClassRefListID, str2SrcQB_ClassRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TemplateRefListID, str2SrcQB_TemplateRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_BillAddressAddr1, str2SrcQB_BillAddressAddr2, str2SrcQB_BillAddressAddr3, str2SrcQB_BillAddressAddr4, str2SrcQB_BillAddressCity, str2SrcQB_BillAddressState, str2SrcQB_BillAddressPostalCode, str2SrcQB_BillAddressCountry, str2SrcQB_ShipAddressAddr1, str2SrcQB_ShipAddressAddr2, str2SrcQB_ShipAddressAddr3, str2SrcQB_ShipAddressAddr4, str2SrcQB_ShipAddressCity, str2SrcQB_ShipAddressState, str2SrcQB_ShipAddressPostalCode, str2SrcQB_ShipAddressCountry, str2SrcQB_IsPending, str2SrcQB_PONumber, str2SrcQB_TermsRefListID, str2SrcQB_TermsRefFullName, str2SrcQB_DueDate, str2SrcQB_SalesRepRefListID, str2SrcQB_SalesRepRefFullName, str2SrcQB_FOB, str2SrcQB_ShipDate, str2SrcQB_ShipMethodRefListID, str2SrcQB_ShipMethodRefFullName, str2SrcQB_Subtotal, str2SrcQB_ItemSalesTaxRefListID, str2SrcQB_ItemSalesTaxRefFullName, str2SrcQB_SalesTaxPercentage, str2SrcQB_SalesTaxTotal, str2SrcQB_TotalAmount, str2SrcQB_CreditRemaining, str2SrcQB_Memo, str2SrcQB_CustomerMsgRefListID, str2SrcQB_CustomerMsgRefFullName, str2SrcQB_IsToBePrinted, str2SrcQB_CustomerSalesTaxCodeRefListID, str2SrcQB_CustomerSalesTaxCodeRefFullName, str2SrcQB_CreditMemoLineType, str2SrcQB_CreditMemoLineSeqNo, str2SrcQB_CreditMemoLineGroupLineTxnLineID, str2SrcQB_CreditMemoLineGroupItemGroupRefListID, str2SrcQB_CreditMemoLineGroupItemGroupRefFullName, str2SrcQB_CreditMemoLineGroupDesc, str2SrcQB_CreditMemoLineGroupQuantity, str2SrcQB_CreditMemoLineGroupIsPrintItemsInGroup, str2SrcQB_CreditMemoLineGroupTotalAmount, str2SrcQB_CreditMemoLineGroupSeqNo, str2SrcQB_CreditMemoLineTxnLineID, str2SrcQB_CreditMemoLineItemRefListID, str2SrcQB_CreditMemoLineItemRefFullName, str2SrcQB_CreditMemoLineDesc, str2SrcQB_CreditMemoLineQuantity, str2SrcQB_CreditMemoLineRate, str2SrcQB_CreditMemoLineRatePercent, str2SrcQB_CreditMemoLinePriceLevelRefListID, str2SrcQB_CreditMemoLinePriceLevelRefFullName, str2SrcQB_CreditMemoLineClassRefListID, str2SrcQB_CreditMemoLineClassRefFullName, str2SrcQB_CreditMemoLineAmount, str2SrcQB_CreditMemoLineServiceDate, str2SrcQB_CreditMemoLineSalesTaxCodeRefListID, str2SrcQB_CreditMemoLineSalesTaxCodeRefFullName, str2SrcQB_CreditMemoLineIsTaxable, str2SrcQB_CreditMemoLineOverrideItemAccountRefListID, str2SrcQB_CreditMemoLineOverrideItemAccountRefFullName, str2SrcQB_FQSaveToCache, str2SrcQB_FQPrimaryKey, str2SrcQB_CustomFieldCreditMemoLineOther1, str2SrcQB_CustomFieldCreditMemoLineOther2, str2SrcQB_CustomFieldCreditMemoLinePriceBreaks, str2SrcQB_CustomFieldCreditMemoLineGroupOther1, str2SrcQB_CustomFieldCreditMemoLineGroupOther2, str2SrcQB_CustomFieldCreditMemoLineGroupPriceBreaks, str2SrcQB_CustomFieldCreditMemoLineGroupLineOther1, str2SrcQB_CustomFieldCreditMemoLineGroupLineOther2, str2SrcQB_CustomFieldCreditMemoLineGroupLinePriceBreaks, str2SrcQB_CustomFieldOther As String
        'This routine gets the 2SrcQB_QB_CreditMemoLine from the database according to the selection in str2SrcQB_QB_CreditMemoLineSQL.
        'It then puts those 2SrcQB_QB_CreditMemoLine in the list box

        ''FOR PART 3TestID_
        'Debug.Print "List3TestID_QBTable"
        'Dim rs3TestID_QBTable As ADODB.Recordset
        'Dim str3TestID_QBTableSQL As String
        'Dim str3TestID_QBTableRow As String
        'Dim str3TestID_ListID As String
        ''This routine gets the 3TestID_QBTable from the database according to the selection in str3TestID_QBTableSQL.
        ''It then puts those 3TestID_QBTable in the list box

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_CreditMemoLine  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_CreditMemoLine"
        Application.DoEvents()


        '
        ''Clear out table
        ''*cnDBPM.Execute "DELETE FROM QB_CreditMemoLine"
        'cnmax.Execute "DELETE FROM QB_CreditMemoLine"
        '

        'Get rs from QB
        'Load table from rs

        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        rs2SrcQB_QB_CreditMemoLine = New DataSet() '*** TAKE QB_ OFF OF TABLE NAME ***
        'str2SrcQB_QB_CreditMemoLineSQL = "SELECT * FROM CreditMemoLine WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        ''str2SrcQB_QB_CreditMemoLineSQL = "SELECT * FROM CreditMemoLine WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_CreditMemoLineSQL = "SELECT * FROM CreditMemoLine WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'str2SrcQB_QB_CreditMemoLineSQL = "SELECT * FROM CreditMemoLine WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_CreditMemoLine & "'} ORDER BY TimeModified"
        'THIS worked:  str2SrcQB_QB_CreditMemoLineSQL = "SELECT * FROM CreditMemoLine WHERE TxnID = '78E53-1145544373'"
        str2SrcQB_QB_CreditMemoLineSQL = "SELECT * FROM CreditMemoLine"
        'Debug.Print str2SrcQB_QB_CreditMemoLineSQL
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_CreditMemoLineSQL, cnQuickBooks)
        rs2SrcQB_QB_CreditMemoLine.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_CreditMemoLine) ', adAsyncFetch
        If rs2SrcQB_QB_CreditMemoLine.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_CreditMemoLine"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_CreditMemoLine"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_CreditMemoLine"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_CreditMemoLine.Tables(0).Rows.Count) & "  QB_CreditMemoLine  Records  ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_CreditMemoLine.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_CreditMemoLine.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_CreditMemoLine.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_CreditMemoLine.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_TxnID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_TxnNumber = ""
                str2SrcQB_CustomerRefListID = ""
                str2SrcQB_CustomerRefFullName = ""
                str2SrcQB_ClassRefListID = ""
                str2SrcQB_ClassRefFullName = ""
                str2SrcQB_ARAccountRefListID = ""
                str2SrcQB_ARAccountRefFullName = ""
                str2SrcQB_TemplateRefListID = ""
                str2SrcQB_TemplateRefFullName = ""
                str2SrcQB_TxnDate = ""
                str2SrcQB_TxnDateMacro = ""
                str2SrcQB_RefNumber = ""
                str2SrcQB_BillAddressAddr1 = ""
                str2SrcQB_BillAddressAddr2 = ""
                str2SrcQB_BillAddressAddr3 = ""
                str2SrcQB_BillAddressAddr4 = ""
                str2SrcQB_BillAddressCity = ""
                str2SrcQB_BillAddressState = ""
                str2SrcQB_BillAddressPostalCode = ""
                str2SrcQB_BillAddressCountry = ""
                str2SrcQB_ShipAddressAddr1 = ""
                str2SrcQB_ShipAddressAddr2 = ""
                str2SrcQB_ShipAddressAddr3 = ""
                str2SrcQB_ShipAddressAddr4 = ""
                str2SrcQB_ShipAddressCity = ""
                str2SrcQB_ShipAddressState = ""
                str2SrcQB_ShipAddressPostalCode = ""
                str2SrcQB_ShipAddressCountry = ""
                str2SrcQB_IsPending = ""
                str2SrcQB_PONumber = ""
                str2SrcQB_TermsRefListID = ""
                str2SrcQB_TermsRefFullName = ""
                str2SrcQB_DueDate = ""
                str2SrcQB_SalesRepRefListID = ""
                str2SrcQB_SalesRepRefFullName = ""
                str2SrcQB_FOB = ""
                str2SrcQB_ShipDate = ""
                str2SrcQB_ShipMethodRefListID = ""
                str2SrcQB_ShipMethodRefFullName = ""
                str2SrcQB_Subtotal = ""
                str2SrcQB_ItemSalesTaxRefListID = ""
                str2SrcQB_ItemSalesTaxRefFullName = ""
                str2SrcQB_SalesTaxPercentage = ""
                str2SrcQB_SalesTaxTotal = ""
                str2SrcQB_TotalAmount = ""
                str2SrcQB_CreditRemaining = ""
                str2SrcQB_Memo = ""
                str2SrcQB_CustomerMsgRefListID = ""
                str2SrcQB_CustomerMsgRefFullName = ""
                str2SrcQB_IsToBePrinted = ""
                str2SrcQB_CustomerSalesTaxCodeRefListID = ""
                str2SrcQB_CustomerSalesTaxCodeRefFullName = ""
                str2SrcQB_CreditMemoLineType = ""
                str2SrcQB_CreditMemoLineSeqNo = ""
                str2SrcQB_CreditMemoLineGroupLineTxnLineID = ""
                str2SrcQB_CreditMemoLineGroupItemGroupRefListID = ""
                str2SrcQB_CreditMemoLineGroupItemGroupRefFullName = ""
                str2SrcQB_CreditMemoLineGroupDesc = ""
                str2SrcQB_CreditMemoLineGroupQuantity = "0"
                str2SrcQB_CreditMemoLineGroupIsPrintItemsInGroup = ""
                str2SrcQB_CreditMemoLineGroupTotalAmount = "0"
                str2SrcQB_CreditMemoLineGroupSeqNo = ""
                str2SrcQB_CreditMemoLineTxnLineID = ""
                str2SrcQB_CreditMemoLineItemRefListID = ""
                str2SrcQB_CreditMemoLineItemRefFullName = ""
                str2SrcQB_CreditMemoLineDesc = ""
                str2SrcQB_CreditMemoLineQuantity = "0"
                str2SrcQB_CreditMemoLineRate = "0"
                str2SrcQB_CreditMemoLineRatePercent = "0"
                str2SrcQB_CreditMemoLinePriceLevelRefListID = ""
                str2SrcQB_CreditMemoLinePriceLevelRefFullName = ""
                str2SrcQB_CreditMemoLineClassRefListID = ""
                str2SrcQB_CreditMemoLineClassRefFullName = ""
                str2SrcQB_CreditMemoLineAmount = "0"
                str2SrcQB_CreditMemoLineServiceDate = ""
                str2SrcQB_CreditMemoLineSalesTaxCodeRefListID = ""
                str2SrcQB_CreditMemoLineSalesTaxCodeRefFullName = ""
                str2SrcQB_CreditMemoLineIsTaxable = ""
                str2SrcQB_CreditMemoLineOverrideItemAccountRefListID = ""
                str2SrcQB_CreditMemoLineOverrideItemAccountRefFullName = ""
                str2SrcQB_FQSaveToCache = ""
                str2SrcQB_FQPrimaryKey = ""
                str2SrcQB_CustomFieldCreditMemoLineOther1 = ""
                str2SrcQB_CustomFieldCreditMemoLineOther2 = ""
                str2SrcQB_CustomFieldCreditMemoLinePriceBreaks = ""
                str2SrcQB_CustomFieldCreditMemoLineGroupOther1 = ""
                str2SrcQB_CustomFieldCreditMemoLineGroupOther2 = ""
                str2SrcQB_CustomFieldCreditMemoLineGroupPriceBreaks = ""
                str2SrcQB_CustomFieldCreditMemoLineGroupLineOther1 = ""
                str2SrcQB_CustomFieldCreditMemoLineGroupLineOther2 = ""
                str2SrcQB_CustomFieldCreditMemoLineGroupLinePriceBreaks = ""
                str2SrcQB_CustomFieldOther = ""

                'get the columns from the database
                If iteration_row("TxnID") <> "" Then str2SrcQB_TxnID = iteration_row("TxnID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("TxnNumber") <> "" Then str2SrcQB_TxnNumber = iteration_row("TxnNumber")
                If iteration_row("CustomerRefListID") <> "" Then str2SrcQB_CustomerRefListID = iteration_row("CustomerRefListID")
                If iteration_row("CustomerRefFullName") <> "" Then str2SrcQB_CustomerRefFullName = iteration_row("CustomerRefFullName")
                If iteration_row("ClassRefListID") <> "" Then str2SrcQB_ClassRefListID = iteration_row("ClassRefListID")
                If iteration_row("ClassRefFullName") <> "" Then str2SrcQB_ClassRefFullName = iteration_row("ClassRefFullName")
                If iteration_row("ARAccountRefListID") <> "" Then str2SrcQB_ARAccountRefListID = iteration_row("ARAccountRefListID")
                If iteration_row("ARAccountRefFullName") <> "" Then str2SrcQB_ARAccountRefFullName = iteration_row("ARAccountRefFullName")
                If iteration_row("TemplateRefListID") <> "" Then str2SrcQB_TemplateRefListID = iteration_row("TemplateRefListID")
                If iteration_row("TemplateRefFullName") <> "" Then str2SrcQB_TemplateRefFullName = iteration_row("TemplateRefFullName")
                If iteration_row("TxnDate") <> "" Then str2SrcQB_TxnDate = iteration_row("TxnDate")
                If iteration_row("TxnDateMacro") <> "" Then str2SrcQB_TxnDateMacro = iteration_row("TxnDateMacro")
                If iteration_row("RefNumber") <> "" Then str2SrcQB_RefNumber = iteration_row("RefNumber")
                If iteration_row("BillAddressAddr1") <> "" Then str2SrcQB_BillAddressAddr1 = iteration_row("BillAddressAddr1")
                If iteration_row("BillAddressAddr2") <> "" Then str2SrcQB_BillAddressAddr2 = iteration_row("BillAddressAddr2")
                If iteration_row("BillAddressAddr3") <> "" Then str2SrcQB_BillAddressAddr3 = iteration_row("BillAddressAddr3")
                If iteration_row("BillAddressAddr4") <> "" Then str2SrcQB_BillAddressAddr4 = iteration_row("BillAddressAddr4")
                If iteration_row("BillAddressCity") <> "" Then str2SrcQB_BillAddressCity = iteration_row("BillAddressCity")
                If iteration_row("BillAddressState") <> "" Then str2SrcQB_BillAddressState = iteration_row("BillAddressState")
                If iteration_row("BillAddressPostalCode") <> "" Then str2SrcQB_BillAddressPostalCode = iteration_row("BillAddressPostalCode")
                If iteration_row("BillAddressCountry") <> "" Then str2SrcQB_BillAddressCountry = iteration_row("BillAddressCountry")
                If iteration_row("ShipAddressAddr1") <> "" Then str2SrcQB_ShipAddressAddr1 = iteration_row("ShipAddressAddr1")
                If iteration_row("ShipAddressAddr2") <> "" Then str2SrcQB_ShipAddressAddr2 = iteration_row("ShipAddressAddr2")
                If iteration_row("ShipAddressAddr3") <> "" Then str2SrcQB_ShipAddressAddr3 = iteration_row("ShipAddressAddr3")
                If iteration_row("ShipAddressAddr4") <> "" Then str2SrcQB_ShipAddressAddr4 = iteration_row("ShipAddressAddr4")
                If iteration_row("ShipAddressCity") <> "" Then str2SrcQB_ShipAddressCity = iteration_row("ShipAddressCity")
                If iteration_row("ShipAddressState") <> "" Then str2SrcQB_ShipAddressState = iteration_row("ShipAddressState")
                If iteration_row("ShipAddressPostalCode") <> "" Then str2SrcQB_ShipAddressPostalCode = iteration_row("ShipAddressPostalCode")
                If iteration_row("ShipAddressCountry") <> "" Then str2SrcQB_ShipAddressCountry = iteration_row("ShipAddressCountry")
                If iteration_row("IsPending") <> "" Then str2SrcQB_IsPending = iteration_row("IsPending")
                If iteration_row("PONumber") <> "" Then str2SrcQB_PONumber = iteration_row("PONumber")
                If iteration_row("TermsRefListID") <> "" Then str2SrcQB_TermsRefListID = iteration_row("TermsRefListID")
                If iteration_row("TermsRefFullName") <> "" Then str2SrcQB_TermsRefFullName = iteration_row("TermsRefFullName")
                If iteration_row("DueDate") <> "" Then str2SrcQB_DueDate = iteration_row("DueDate")
                If iteration_row("SalesRepRefListID") <> "" Then str2SrcQB_SalesRepRefListID = iteration_row("SalesRepRefListID")
                If iteration_row("SalesRepRefFullName") <> "" Then str2SrcQB_SalesRepRefFullName = iteration_row("SalesRepRefFullName")
                If iteration_row("FOB") <> "" Then str2SrcQB_FOB = iteration_row("FOB")
                If iteration_row("ShipDate") <> "" Then str2SrcQB_ShipDate = iteration_row("ShipDate")
                If iteration_row("ShipMethodRefListID") <> "" Then str2SrcQB_ShipMethodRefListID = iteration_row("ShipMethodRefListID")
                If iteration_row("ShipMethodRefFullName") <> "" Then str2SrcQB_ShipMethodRefFullName = iteration_row("ShipMethodRefFullName")
                If iteration_row("Subtotal") <> "" Then str2SrcQB_Subtotal = iteration_row("Subtotal")
                If iteration_row("ItemSalesTaxRefListID") <> "" Then str2SrcQB_ItemSalesTaxRefListID = iteration_row("ItemSalesTaxRefListID")
                If iteration_row("ItemSalesTaxRefFullName") <> "" Then str2SrcQB_ItemSalesTaxRefFullName = iteration_row("ItemSalesTaxRefFullName")
                If iteration_row("SalesTaxPercentage") <> "" Then str2SrcQB_SalesTaxPercentage = iteration_row("SalesTaxPercentage")
                If iteration_row("SalesTaxTotal") <> "" Then str2SrcQB_SalesTaxTotal = iteration_row("SalesTaxTotal")
                If iteration_row("TotalAmount") <> "" Then str2SrcQB_TotalAmount = iteration_row("TotalAmount")
                If iteration_row("CreditRemaining") <> "" Then str2SrcQB_CreditRemaining = iteration_row("CreditRemaining")
                If iteration_row("Memo") <> "" Then str2SrcQB_Memo = iteration_row("Memo")
                If iteration_row("CustomerMsgRefListID") <> "" Then str2SrcQB_CustomerMsgRefListID = iteration_row("CustomerMsgRefListID")
                If iteration_row("CustomerMsgRefFullName") <> "" Then str2SrcQB_CustomerMsgRefFullName = iteration_row("CustomerMsgRefFullName")
                If iteration_row("IsToBePrinted") <> "" Then str2SrcQB_IsToBePrinted = iteration_row("IsToBePrinted")
                If iteration_row("CustomerSalesTaxCodeRefListID") <> "" Then str2SrcQB_CustomerSalesTaxCodeRefListID = iteration_row("CustomerSalesTaxCodeRefListID")
                If iteration_row("CustomerSalesTaxCodeRefFullName") <> "" Then str2SrcQB_CustomerSalesTaxCodeRefFullName = iteration_row("CustomerSalesTaxCodeRefFullName")
                If iteration_row("CreditMemoLineType") <> "" Then str2SrcQB_CreditMemoLineType = iteration_row("CreditMemoLineType")
                If iteration_row("CreditMemoLineSeqNo") <> "" Then str2SrcQB_CreditMemoLineSeqNo = iteration_row("CreditMemoLineSeqNo")
                If iteration_row("CreditMemoLineGroupLineTxnLineID") <> "" Then str2SrcQB_CreditMemoLineGroupLineTxnLineID = iteration_row("CreditMemoLineGroupLineTxnLineID")
                If iteration_row("CreditMemoLineGroupItemGroupRefListID") <> "" Then str2SrcQB_CreditMemoLineGroupItemGroupRefListID = iteration_row("CreditMemoLineGroupItemGroupRefListID")
                If iteration_row("CreditMemoLineGroupItemGroupRefFullName") <> "" Then str2SrcQB_CreditMemoLineGroupItemGroupRefFullName = iteration_row("CreditMemoLineGroupItemGroupRefFullName")
                If iteration_row("CreditMemoLineGroupDesc") <> "" Then str2SrcQB_CreditMemoLineGroupDesc = iteration_row("CreditMemoLineGroupDesc")
                If iteration_row("CreditMemoLineGroupQuantity") <> "" Then str2SrcQB_CreditMemoLineGroupQuantity = iteration_row("CreditMemoLineGroupQuantity")
                If iteration_row("CreditMemoLineGroupIsPrintItemsInGroup") <> "" Then str2SrcQB_CreditMemoLineGroupIsPrintItemsInGroup = iteration_row("CreditMemoLineGroupIsPrintItemsInGroup")
                If iteration_row("CreditMemoLineGroupTotalAmount") <> "" Then str2SrcQB_CreditMemoLineGroupTotalAmount = iteration_row("CreditMemoLineGroupTotalAmount")
                If iteration_row("CreditMemoLineGroupSeqNo") <> "" Then str2SrcQB_CreditMemoLineGroupSeqNo = iteration_row("CreditMemoLineGroupSeqNo")
                If iteration_row("CreditMemoLineTxnLineID") <> "" Then str2SrcQB_CreditMemoLineTxnLineID = iteration_row("CreditMemoLineTxnLineID")
                If iteration_row("CreditMemoLineItemRefListID") <> "" Then str2SrcQB_CreditMemoLineItemRefListID = iteration_row("CreditMemoLineItemRefListID")
                If iteration_row("CreditMemoLineItemRefFullName") <> "" Then str2SrcQB_CreditMemoLineItemRefFullName = iteration_row("CreditMemoLineItemRefFullName")
                If iteration_row("CreditMemoLineDesc") <> "" Then str2SrcQB_CreditMemoLineDesc = iteration_row("CreditMemoLineDesc")
                If iteration_row("CreditMemoLineQuantity") <> "" Then str2SrcQB_CreditMemoLineQuantity = iteration_row("CreditMemoLineQuantity")
                If iteration_row("CreditMemoLineRate") <> "" Then str2SrcQB_CreditMemoLineRate = iteration_row("CreditMemoLineRate")
                If iteration_row("CreditMemoLineRatePercent") <> "" Then str2SrcQB_CreditMemoLineRatePercent = iteration_row("CreditMemoLineRatePercent")
                If iteration_row("CreditMemoLinePriceLevelRefListID") <> "" Then str2SrcQB_CreditMemoLinePriceLevelRefListID = iteration_row("CreditMemoLinePriceLevelRefListID")
                If iteration_row("CreditMemoLinePriceLevelRefFullName") <> "" Then str2SrcQB_CreditMemoLinePriceLevelRefFullName = iteration_row("CreditMemoLinePriceLevelRefFullName")
                If iteration_row("CreditMemoLineClassRefListID") <> "" Then str2SrcQB_CreditMemoLineClassRefListID = iteration_row("CreditMemoLineClassRefListID")
                If iteration_row("CreditMemoLineClassRefFullName") <> "" Then str2SrcQB_CreditMemoLineClassRefFullName = iteration_row("CreditMemoLineClassRefFullName")
                If iteration_row("CreditMemoLineAmount") <> "" Then str2SrcQB_CreditMemoLineAmount = iteration_row("CreditMemoLineAmount")
                If iteration_row("CreditMemoLineServiceDate") <> "" Then str2SrcQB_CreditMemoLineServiceDate = iteration_row("CreditMemoLineServiceDate")
                If iteration_row("CreditMemoLineSalesTaxCodeRefListID") <> "" Then str2SrcQB_CreditMemoLineSalesTaxCodeRefListID = iteration_row("CreditMemoLineSalesTaxCodeRefListID")
                If iteration_row("CreditMemoLineSalesTaxCodeRefFullName") <> "" Then str2SrcQB_CreditMemoLineSalesTaxCodeRefFullName = iteration_row("CreditMemoLineSalesTaxCodeRefFullName")
                '        If rs2SrcQB_QB_CreditMemoLine!CreditMemoLineIsTaxable <> "" Then str2SrcQB_CreditMemoLineIsTaxable = rs2SrcQB_QB_CreditMemoLine!CreditMemoLineIsTaxable
                If iteration_row("CreditMemoLineOverrideItemAccountRefListID") <> "" Then str2SrcQB_CreditMemoLineOverrideItemAccountRefListID = iteration_row("CreditMemoLineOverrideItemAccountRefListID")
                If iteration_row("CreditMemoLineOverrideItemAccountRefFullName") <> "" Then str2SrcQB_CreditMemoLineOverrideItemAccountRefFullName = iteration_row("CreditMemoLineOverrideItemAccountRefFullName")
                If iteration_row("FQSaveToCache") <> "" Then str2SrcQB_FQSaveToCache = iteration_row("FQSaveToCache")
                If iteration_row("FQPrimaryKey") <> "" Then str2SrcQB_FQPrimaryKey = iteration_row("FQPrimaryKey")
                If iteration_row("CustomFieldCreditMemoLineOther1") <> "" Then str2SrcQB_CustomFieldCreditMemoLineOther1 = iteration_row("CustomFieldCreditMemoLineOther1")
                If iteration_row("CustomFieldCreditMemoLineOther2") <> "" Then str2SrcQB_CustomFieldCreditMemoLineOther2 = iteration_row("CustomFieldCreditMemoLineOther2")
                'If rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLinePriceBreaks <> "" Then str2SrcQB_CustomFieldCreditMemoLinePriceBreaks = rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLinePriceBreaks
                If iteration_row("CustomFieldCreditMemoLineGroupOther1") <> "" Then str2SrcQB_CustomFieldCreditMemoLineGroupOther1 = iteration_row("CustomFieldCreditMemoLineGroupOther1")
                If iteration_row("CustomFieldCreditMemoLineGroupOther2") <> "" Then str2SrcQB_CustomFieldCreditMemoLineGroupOther2 = iteration_row("CustomFieldCreditMemoLineGroupOther2")
                'If rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLineGroupPriceBreaks <> "" Then str2SrcQB_CustomFieldCreditMemoLineGroupPriceBreaks = rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLineGroupPriceBreaks
                '        If rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLineGroupLineOther1 <> "" Then str2SrcQB_CustomFieldCreditMemoLineGroupLineOther1 = rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLineGroupLineOther1
                '        If rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLineGroupLineOther2 <> "" Then str2SrcQB_CustomFieldCreditMemoLineGroupLineOther2 = rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLineGroupLineOther2
                'If rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLineGroupLinePriceBreaks <> "" Then str2SrcQB_CustomFieldCreditMemoLineGroupLinePriceBreaks = rs2SrcQB_QB_CreditMemoLine!CustomFieldCreditMemoLineGroupLinePriceBreaks
                If iteration_row("CustomFieldOther") <> "" Then str2SrcQB_CustomFieldOther = iteration_row("CustomFieldOther")
                '        If rs2SrcQB_QB_CreditMemoLine!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_CreditMemoLine!CustomFieldOther

                'Strip quote character out of strings
                str2SrcQB_TxnID = str2SrcQB_TxnID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_TxnNumber = str2SrcQB_TxnNumber.Replace("'"c, "`"c)
                str2SrcQB_CustomerRefListID = str2SrcQB_CustomerRefListID.Replace("'"c, "`"c)
                str2SrcQB_CustomerRefFullName = str2SrcQB_CustomerRefFullName.Replace("'"c, "`"c)
                str2SrcQB_ClassRefListID = str2SrcQB_ClassRefListID.Replace("'"c, "`"c)
                str2SrcQB_ClassRefFullName = str2SrcQB_ClassRefFullName.Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefListID = str2SrcQB_ARAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_ARAccountRefFullName = str2SrcQB_ARAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_TemplateRefListID = str2SrcQB_TemplateRefListID.Replace("'"c, "`"c)
                str2SrcQB_TemplateRefFullName = str2SrcQB_TemplateRefFullName.Replace("'"c, "`"c)
                str2SrcQB_TxnDate = str2SrcQB_TxnDate.Replace("'"c, "`"c)
                str2SrcQB_TxnDateMacro = str2SrcQB_TxnDateMacro.Replace("'"c, "`"c)
                str2SrcQB_RefNumber = str2SrcQB_RefNumber.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr1 = str2SrcQB_BillAddressAddr1.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr2 = str2SrcQB_BillAddressAddr2.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr3 = str2SrcQB_BillAddressAddr3.Replace("'"c, "`"c)
                str2SrcQB_BillAddressAddr4 = str2SrcQB_BillAddressAddr4.Replace("'"c, "`"c)
                str2SrcQB_BillAddressCity = str2SrcQB_BillAddressCity.Replace("'"c, "`"c)
                str2SrcQB_BillAddressState = str2SrcQB_BillAddressState.Replace("'"c, "`"c)
                str2SrcQB_BillAddressPostalCode = str2SrcQB_BillAddressPostalCode.Replace("'"c, "`"c)
                str2SrcQB_BillAddressCountry = str2SrcQB_BillAddressCountry.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr1 = str2SrcQB_ShipAddressAddr1.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr2 = str2SrcQB_ShipAddressAddr2.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr3 = str2SrcQB_ShipAddressAddr3.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressAddr4 = str2SrcQB_ShipAddressAddr4.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressCity = str2SrcQB_ShipAddressCity.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressState = str2SrcQB_ShipAddressState.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressPostalCode = str2SrcQB_ShipAddressPostalCode.Replace("'"c, "`"c)
                str2SrcQB_ShipAddressCountry = str2SrcQB_ShipAddressCountry.Replace("'"c, "`"c)
                str2SrcQB_IsPending = str2SrcQB_IsPending.Replace("'"c, "`"c)
                str2SrcQB_PONumber = str2SrcQB_PONumber.Replace("'"c, "`"c)
                str2SrcQB_TermsRefListID = str2SrcQB_TermsRefListID.Replace("'"c, "`"c)
                str2SrcQB_TermsRefFullName = str2SrcQB_TermsRefFullName.Replace("'"c, "`"c)
                str2SrcQB_DueDate = str2SrcQB_DueDate.Replace("'"c, "`"c)
                str2SrcQB_SalesRepRefListID = str2SrcQB_SalesRepRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesRepRefFullName = str2SrcQB_SalesRepRefFullName.Replace("'"c, "`"c)
                str2SrcQB_FOB = str2SrcQB_FOB.Replace("'"c, "`"c)
                str2SrcQB_ShipDate = str2SrcQB_ShipDate.Replace("'"c, "`"c)
                str2SrcQB_ShipMethodRefListID = str2SrcQB_ShipMethodRefListID.Replace("'"c, "`"c)
                str2SrcQB_ShipMethodRefFullName = str2SrcQB_ShipMethodRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Subtotal = str2SrcQB_Subtotal.Replace("'"c, "`"c)
                str2SrcQB_ItemSalesTaxRefListID = str2SrcQB_ItemSalesTaxRefListID.Replace("'"c, "`"c)
                str2SrcQB_ItemSalesTaxRefFullName = str2SrcQB_ItemSalesTaxRefFullName.Replace("'"c, "`"c)
                str2SrcQB_SalesTaxPercentage = str2SrcQB_SalesTaxPercentage.Replace("'"c, "`"c)
                str2SrcQB_SalesTaxTotal = str2SrcQB_SalesTaxTotal.Replace("'"c, "`"c)
                str2SrcQB_TotalAmount = str2SrcQB_TotalAmount.Replace("'"c, "`"c)
                str2SrcQB_CreditRemaining = str2SrcQB_CreditRemaining.Replace("'"c, "`"c)
                str2SrcQB_Memo = str2SrcQB_Memo.Replace("'"c, "`"c)
                str2SrcQB_CustomerMsgRefListID = str2SrcQB_CustomerMsgRefListID.Replace("'"c, "`"c)
                str2SrcQB_CustomerMsgRefFullName = str2SrcQB_CustomerMsgRefFullName.Replace("'"c, "`"c)
                str2SrcQB_IsToBePrinted = str2SrcQB_IsToBePrinted.Replace("'"c, "`"c)
                str2SrcQB_CustomerSalesTaxCodeRefListID = str2SrcQB_CustomerSalesTaxCodeRefListID.Replace("'"c, "`"c)
                str2SrcQB_CustomerSalesTaxCodeRefFullName = str2SrcQB_CustomerSalesTaxCodeRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineType = str2SrcQB_CreditMemoLineType.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineSeqNo = str2SrcQB_CreditMemoLineSeqNo.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineGroupLineTxnLineID = str2SrcQB_CreditMemoLineGroupLineTxnLineID.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineGroupItemGroupRefListID = str2SrcQB_CreditMemoLineGroupItemGroupRefListID.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineGroupItemGroupRefFullName = str2SrcQB_CreditMemoLineGroupItemGroupRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineGroupDesc = str2SrcQB_CreditMemoLineGroupDesc.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineGroupQuantity = str2SrcQB_CreditMemoLineGroupQuantity.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineGroupIsPrintItemsInGroup = str2SrcQB_CreditMemoLineGroupIsPrintItemsInGroup.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineGroupTotalAmount = str2SrcQB_CreditMemoLineGroupTotalAmount.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineGroupSeqNo = str2SrcQB_CreditMemoLineGroupSeqNo.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineTxnLineID = str2SrcQB_CreditMemoLineTxnLineID.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineItemRefListID = str2SrcQB_CreditMemoLineItemRefListID.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineItemRefFullName = str2SrcQB_CreditMemoLineItemRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineDesc = str2SrcQB_CreditMemoLineDesc.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineQuantity = str2SrcQB_CreditMemoLineQuantity.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineRate = str2SrcQB_CreditMemoLineRate.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineRatePercent = str2SrcQB_CreditMemoLineRatePercent.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLinePriceLevelRefListID = str2SrcQB_CreditMemoLinePriceLevelRefListID.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLinePriceLevelRefFullName = str2SrcQB_CreditMemoLinePriceLevelRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineClassRefListID = str2SrcQB_CreditMemoLineClassRefListID.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineClassRefFullName = str2SrcQB_CreditMemoLineClassRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineAmount = str2SrcQB_CreditMemoLineAmount.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineServiceDate = str2SrcQB_CreditMemoLineServiceDate.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineSalesTaxCodeRefListID = str2SrcQB_CreditMemoLineSalesTaxCodeRefListID.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineSalesTaxCodeRefFullName = str2SrcQB_CreditMemoLineSalesTaxCodeRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineIsTaxable = str2SrcQB_CreditMemoLineIsTaxable.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineOverrideItemAccountRefListID = str2SrcQB_CreditMemoLineOverrideItemAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_CreditMemoLineOverrideItemAccountRefFullName = str2SrcQB_CreditMemoLineOverrideItemAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_FQSaveToCache = str2SrcQB_FQSaveToCache.Replace("'"c, "`"c)
                str2SrcQB_FQPrimaryKey = str2SrcQB_FQPrimaryKey.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldCreditMemoLineOther1 = str2SrcQB_CustomFieldCreditMemoLineOther1.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldCreditMemoLineOther2 = str2SrcQB_CustomFieldCreditMemoLineOther2.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldCreditMemoLinePriceBreaks = str2SrcQB_CustomFieldCreditMemoLinePriceBreaks.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldCreditMemoLineGroupOther1 = str2SrcQB_CustomFieldCreditMemoLineGroupOther1.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldCreditMemoLineGroupOther2 = str2SrcQB_CustomFieldCreditMemoLineGroupOther2.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldCreditMemoLineGroupPriceBreaks = str2SrcQB_CustomFieldCreditMemoLineGroupPriceBreaks.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldCreditMemoLineGroupLineOther1 = str2SrcQB_CustomFieldCreditMemoLineGroupLineOther1.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldCreditMemoLineGroupLineOther2 = str2SrcQB_CustomFieldCreditMemoLineGroupLineOther2.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldCreditMemoLineGroupLinePriceBreaks = str2SrcQB_CustomFieldCreditMemoLineGroupLinePriceBreaks.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldOther = str2SrcQB_CustomFieldOther.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsPending = IIf(str2SrcQB_IsPending = "True", "1", "0")
                str2SrcQB_IsToBePrinted = IIf(str2SrcQB_IsToBePrinted = "True", "1", "0")
                str2SrcQB_FQSaveToCache = IIf(str2SrcQB_FQSaveToCache = "True", "1", "0")


                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_CreditMemoLineRow = "" & _
                                                 Strings.Left(str2SrcQB_TxnID & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_TxnNumber & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_CustomerRefListID & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_CustomerRefFullName & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_ClassRefListID & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_ClassRefFullName & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_ARAccountRefListID & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_ARAccountRefFullName & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_TemplateRefListID & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_TemplateRefFullName & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_TxnDate & "                  ", 18) & "   " & _
                                                 Strings.Left(str2SrcQB_TxnDateMacro & "                  ", 18) & "   " & _
                                                 "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_CreditMemoLine.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_CreditMemoLine.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_CreditMemoLine.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_CreditMemoLineRow)
                    'frmMain.lstConversionProgress.CreditMemoLineData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If






                'DO WORK: With each record

                '        'Update cust balances
                '        UpdateQBCustomerBalance (str2SrcQB_CustomerRefListID)
                '
                '        'Update inv info
                '        'UpdateQBInvoice (str2SrcQB_TxnID)
                '        UpdateQBInvoice (str2SrcQB_CustomerRefFullName)
                '
                '
                '
                '        'ADD CreditMemo LINE STUFF HERE
                '
                '
                '
                '
                '
                '
                '
                '        'Check to see if ListID is in QBTable            'Yes then UPDATE record            'No then INSERT record
                '        'New recordset
                '        Set rs3TestID_QBTable = New ADODB.Recordset
                '        str3TestID_QBTableSQL = "SELECT TxnID FROM QB_CreditMemo WHERE TxnID = '" & str2SrcQB_TxnID & "'"
                '        'Debug.Print str3TestID_QBTableSQL
                '        'rs3TestID_QBTable.Open str3TestID_QBTableSQL, cnDBPM, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        rs3TestID_QBTable.Open str3TestID_QBTableSQL, cnmax, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
                '        If rs3TestID_QBTable.RecordCount > 1 Then Stop 'Should only be one
                '        If rs3TestID_QBTable.RecordCount > 0 Then  'record exists  -UPDATE
                '            'DO UPDATE WORK:
                '            Debug.Print "UPDATE"
                '
                '            'Build the SQL string
                '            strSQL1 = "UPDATE  " & vbCrLf & _
                ''                      "       QB_CreditMemo " & vbCrLf & _
                ''                      "SET " & vbCrLf & _
                ''                      "       TxnID = '" & str2SrcQB_TxnID & "'" & vbCrLf & _
                ''                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & vbCrLf & _
                ''                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & vbCrLf & _
                ''                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & vbCrLf & _
                ''                      "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & vbCrLf & _
                ''                      "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & vbCrLf & _
                ''                      "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & vbCrLf & _
                ''                      "     , ClassRefListID = '" & str2SrcQB_ClassRefListID & "'" & vbCrLf & _
                ''                      "     , ClassRefFullName = '" & str2SrcQB_ClassRefFullName & "'" & vbCrLf & _
                ''                      "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & vbCrLf & _
                ''                      "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & vbCrLf & _
                ''                      "     , TemplateRefListID = '" & str2SrcQB_TemplateRefListID & "'" & vbCrLf & _
                ''                      "     , TemplateRefFullName = '" & str2SrcQB_TemplateRefFullName & "'" & vbCrLf & _
                ''                      "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & vbCrLf & _
                ''                      "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & vbCrLf & _
                ''                      "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr1 = '" & str2SrcQB_BillAddressAddr1 & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr2 = '" & str2SrcQB_BillAddressAddr2 & "'" & vbCrLf
                '            strSQL2 = "     , BillAddressAddr3 = '" & str2SrcQB_BillAddressAddr3 & "'" & vbCrLf & _
                ''                      "     , BillAddressAddr4 = '" & str2SrcQB_BillAddressAddr4 & "'" & vbCrLf & _
                ''                      "     , BillAddressCity = '" & str2SrcQB_BillAddressCity & "'" & vbCrLf & _
                ''                      "     , BillAddressState = '" & str2SrcQB_BillAddressState & "'" & vbCrLf & _
                ''                      "     , BillAddressPostalCode = '" & str2SrcQB_BillAddressPostalCode & "'" & vbCrLf & _
                ''                      "     , BillAddressCountry = '" & str2SrcQB_BillAddressCountry & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr1 = '" & str2SrcQB_ShipAddressAddr1 & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr2 = '" & str2SrcQB_ShipAddressAddr2 & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr3 = '" & str2SrcQB_ShipAddressAddr3 & "'" & vbCrLf & _
                ''                      "     , ShipAddressAddr4 = '" & str2SrcQB_ShipAddressAddr4 & "'" & vbCrLf & _
                ''                      "     , ShipAddressCity = '" & str2SrcQB_ShipAddressCity & "'" & vbCrLf & _
                ''                      "     , ShipAddressState = '" & str2SrcQB_ShipAddressState & "'" & vbCrLf & _
                ''                      "     , ShipAddressPostalCode = '" & str2SrcQB_ShipAddressPostalCode & "'" & vbCrLf & _
                ''                      "     , ShipAddressCountry = '" & str2SrcQB_ShipAddressCountry & "'" & vbCrLf & _
                ''                      "     , IsPending = '" & str2SrcQB_IsPending & "'" & vbCrLf & _
                ''                      "     , PONumber = '" & str2SrcQB_PONumber & "'" & vbCrLf & _
                ''                      "     , TermsRefListID = '" & str2SrcQB_TermsRefListID & "'" & vbCrLf & _
                ''                      "     , TermsRefFullName = '" & str2SrcQB_TermsRefFullName & "'" & vbCrLf & _
                ''                      "     , DueDate = '" & str2SrcQB_DueDate & "'" & vbCrLf & _
                ''                      "     , SalesRepRefListID = '" & str2SrcQB_SalesRepRefListID & "'" & vbCrLf & _
                ''                      "     , SalesRepRefFullName = '" & str2SrcQB_SalesRepRefFullName & "'" & vbCrLf
                '            strSQL3 = "     , FOB = '" & str2SrcQB_FOB & "'" & vbCrLf & _
                ''                      "     , ShipDate = '" & str2SrcQB_ShipDate & "'" & vbCrLf & _
                ''                      "     , ShipMethodRefListID = '" & str2SrcQB_ShipMethodRefListID & "'" & vbCrLf & _
                ''                      "     , ShipMethodRefFullName = '" & str2SrcQB_ShipMethodRefFullName & "'" & vbCrLf & _
                ''                      "     , Subtotal = " & str2SrcQB_Subtotal & "" & vbCrLf & _
                ''                      "     , ItemSalesTaxRefListID = '" & str2SrcQB_ItemSalesTaxRefListID & "'" & vbCrLf & _
                ''                      "     , ItemSalesTaxRefFullName = '" & str2SrcQB_ItemSalesTaxRefFullName & "'" & vbCrLf & _
                ''                      "     , SalesTaxPercentage = " & str2SrcQB_SalesTaxPercentage & "" & vbCrLf & _
                ''                      "     , SalesTaxTotal = " & str2SrcQB_SalesTaxTotal & "" & vbCrLf & _
                ''                      "     , TotalAmount = " & str2SrcQB_TotalAmount & "" & vbCrLf & _
                ''                      "     , CreditRemaining = " & str2SrcQB_CreditRemaining & "" & vbCrLf & _
                ''                      "     , Memo = '" & str2SrcQB_Memo & "'" & vbCrLf & _
                ''                      "     , CustomerMsgRefListID = '" & str2SrcQB_CustomerMsgRefListID & "'" & vbCrLf & _
                ''                      "     , CustomerMsgRefFullName = '" & str2SrcQB_CustomerMsgRefFullName & "'" & vbCrLf & _
                ''                      "     , IsToBePrinted = '" & str2SrcQB_IsToBePrinted & "'" & vbCrLf & _
                ''                      "     , CustomerSalesTaxCodeRefListID = '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'" & vbCrLf & _
                ''                      "     , CustomerSalesTaxCodeRefFullName = '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'" & vbCrLf & _
                ''                      "     , CustomFieldOther = '" & str2SrcQB_CustomFieldOther & "'" & vbCrLf & _
                ''                      "WHERE " & vbCrLf & _
                ''                      "       TxnID = '" & str2SrcQB_TxnID & "'" & vbCrLf
                '
                '
                '
                '            'Combine the strings
                '            strTableUpdate = strSQL1 & strSQL2 & strSQL3 '& strSQL4 & strSQL5 & strSQL6
                '            'Debug.Print strTableUpdate
                '
                '            'Execute the insert
                '            '*cnDBPM.Execute strTableUpdate
                '            cnmax.Execute strTableUpdate
                '
                '
                '
                '
                '        Else 'record not exist  -INSERT
                '            'DO INSERT WORK:
                '            Debug.Print "INSERT"

                'Build the SQL string
                'MODIFICATION REQUIRED HERE
                strSQL1 = "INSERT INTO QB_CreditMemoLine " & Environment.NewLine & _
                          "   ( TxnID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , TxnNumber " & Environment.NewLine & _
                          "   , CustomerRefListID " & Environment.NewLine & _
                          "   , CustomerRefFullName " & Environment.NewLine & _
                          "   , ClassRefListID " & Environment.NewLine & _
                          "   , ClassRefFullName " & Environment.NewLine & _
                          "   , ARAccountRefListID " & Environment.NewLine & _
                          "   , ARAccountRefFullName " & Environment.NewLine & _
                          "   , TemplateRefListID " & Environment.NewLine & _
                          "   , TemplateRefFullName " & Environment.NewLine & _
                          "   , TxnDate " & Environment.NewLine & _
                          "   , TxnDateMacro " & Environment.NewLine & _
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
                          "   , DueDate " & Environment.NewLine & _
                          "   , SalesRepRefListID " & Environment.NewLine & _
                          "   , SalesRepRefFullName " & Environment.NewLine & _
                          "   , FOB " & Environment.NewLine & _
                          "   , ShipDate " & Environment.NewLine & _
                          "   , ShipMethodRefListID " & Environment.NewLine & _
                          "   , ShipMethodRefFullName " & Environment.NewLine & _
                          "   , Subtotal " & Environment.NewLine & _
                          "   , ItemSalesTaxRefListID " & Environment.NewLine & _
                          "   , ItemSalesTaxRefFullName " & Environment.NewLine & _
                          "   , SalesTaxPercentage " & Environment.NewLine & _
                          "   , SalesTaxTotal " & Environment.NewLine & _
                          "   , TotalAmount " & Environment.NewLine
                strSQL3 = "   , CreditRemaining " & Environment.NewLine & _
                          "   , Memo " & Environment.NewLine & _
                          "   , CustomerMsgRefListID " & Environment.NewLine & _
                          "   , CustomerMsgRefFullName " & Environment.NewLine & _
                          "   , IsToBePrinted " & Environment.NewLine & _
                          "   , CustomerSalesTaxCodeRefListID " & Environment.NewLine & _
                          "   , CustomerSalesTaxCodeRefFullName " & Environment.NewLine & _
                          "   , CreditMemoLineType " & Environment.NewLine & _
                          "   , CreditMemoLineSeqNo " & Environment.NewLine & _
                          "   , CreditMemoLineGroupLineTxnLineID " & Environment.NewLine & _
                          "   , CreditMemoLineGroupItemGroupRefListID " & Environment.NewLine & _
                          "   , CreditMemoLineGroupItemGroupRefFullName " & Environment.NewLine & _
                          "   , CreditMemoLineGroupDesc " & Environment.NewLine & _
                          "   , CreditMemoLineGroupQuantity " & Environment.NewLine & _
                          "   , CreditMemoLineGroupIsPrintItemsInGroup " & Environment.NewLine & _
                          "   , CreditMemoLineGroupTotalAmount " & Environment.NewLine & _
                          "   , CreditMemoLineGroupSeqNo " & Environment.NewLine & _
                          "   , CreditMemoLineTxnLineID " & Environment.NewLine & _
                          "   , CreditMemoLineItemRefListID " & Environment.NewLine & _
                          "   , CreditMemoLineItemRefFullName " & Environment.NewLine & _
                          "   , CreditMemoLineDesc " & Environment.NewLine & _
                          "   , CreditMemoLineQuantity " & Environment.NewLine & _
                          "   , CreditMemoLineRate " & Environment.NewLine & _
                          "   , CreditMemoLineRatePercent " & Environment.NewLine & _
                          "   , CreditMemoLinePriceLevelRefListID " & Environment.NewLine
                strSQL4 = "   , CreditMemoLinePriceLevelRefFullName " & Environment.NewLine & _
                          "   , CreditMemoLineClassRefListID " & Environment.NewLine & _
                          "   , CreditMemoLineClassRefFullName " & Environment.NewLine & _
                          "   , CreditMemoLineAmount " & Environment.NewLine & _
                          "   , CreditMemoLineServiceDate " & Environment.NewLine & _
                          "   , CreditMemoLineSalesTaxCodeRefListID " & Environment.NewLine & _
                          "   , CreditMemoLineSalesTaxCodeRefFullName " & Environment.NewLine & _
                          "   , CreditMemoLineIsTaxable " & Environment.NewLine & _
                          "   , CreditMemoLineOverrideItemAccountRefListID " & Environment.NewLine & _
                          "   , CreditMemoLineOverrideItemAccountRefFullName " & Environment.NewLine & _
                          "   , FQSaveToCache " & Environment.NewLine & _
                          "   , FQPrimaryKey " & Environment.NewLine & _
                          "   , CustomFieldCreditMemoLineOther1 " & Environment.NewLine & _
                          "   , CustomFieldCreditMemoLineOther2 " & Environment.NewLine & _
                          "   , CustomFieldCreditMemoLinePriceBreaks " & Environment.NewLine & _
                          "   , CustomFieldCreditMemoLineGroupOther1 " & Environment.NewLine & _
                          "   , CustomFieldCreditMemoLineGroupOther2 " & Environment.NewLine & _
                          "   , CustomFieldCreditMemoLineGroupPriceBreaks " & Environment.NewLine & _
                          "   , CustomFieldCreditMemoLineGroupLineOther1 " & Environment.NewLine & _
                          "   , CustomFieldCreditMemoLineGroupLineOther2 " & Environment.NewLine & _
                          "   , CustomFieldCreditMemoLineGroupLinePriceBreaks " & Environment.NewLine & _
                          "   , CustomFieldOther ) " & Environment.NewLine


                strSQL5 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_TxnID & "'  --TxnID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , " & str2SrcQB_TxnNumber & "  --TxnNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerRefListID & "'  --CustomerRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerRefFullName & "'  --CustomerRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ClassRefListID & "'  --ClassRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ClassRefFullName & "'  --ClassRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ARAccountRefListID & "'  --ARAccountRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ARAccountRefFullName & "'  --ARAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TemplateRefListID & "'  --TemplateRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TemplateRefFullName & "'  --TemplateRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TxnDate & "'  --TxnDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TxnDateMacro & "'  --TxnDateMacro" & Environment.NewLine & _
                          "   , '" & str2SrcQB_RefNumber & "'  --RefNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr1 & "'  --BillAddressAddr1" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr2 & "'  --BillAddressAddr2" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr3 & "'  --BillAddressAddr3" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressAddr4 & "'  --BillAddressAddr4" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressCity & "'  --BillAddressCity" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressState & "'  --BillAddressState" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressPostalCode & "'  --BillAddressPostalCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BillAddressCountry & "'  --BillAddressCountry" & Environment.NewLine
                strSQL6 = "   , '" & str2SrcQB_ShipAddressAddr1 & "'  --ShipAddressAddr1" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr2 & "'  --ShipAddressAddr2" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr3 & "'  --ShipAddressAddr3" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressAddr4 & "'  --ShipAddressAddr4" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressCity & "'  --ShipAddressCity" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressState & "'  --ShipAddressState" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressPostalCode & "'  --ShipAddressPostalCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipAddressCountry & "'  --ShipAddressCountry" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsPending & "'  --IsPending" & Environment.NewLine & _
                          "   , '" & str2SrcQB_PONumber & "'  --PONumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TermsRefListID & "'  --TermsRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TermsRefFullName & "'  --TermsRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_DueDate & "'  --DueDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesRepRefListID & "'  --SalesRepRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesRepRefFullName & "'  --SalesRepRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_FOB & "'  --FOB" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipDate & "'  --ShipDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipMethodRefListID & "'  --ShipMethodRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ShipMethodRefFullName & "'  --ShipMethodRefFullName" & Environment.NewLine & _
                          "   , " & str2SrcQB_Subtotal & "  --Subtotal" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ItemSalesTaxRefListID & "'  --ItemSalesTaxRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ItemSalesTaxRefFullName & "'  --ItemSalesTaxRefFullName" & Environment.NewLine & _
                          "   , " & str2SrcQB_SalesTaxPercentage & "  --SalesTaxPercentage" & Environment.NewLine & _
                          "   , " & str2SrcQB_SalesTaxTotal & "  --SalesTaxTotal" & Environment.NewLine & _
                          "   , " & str2SrcQB_TotalAmount & "  --TotalAmount" & Environment.NewLine
                strSQL7 = "   , " & str2SrcQB_CreditRemaining & "  --CreditRemaining" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Memo & "'  --Memo" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerMsgRefListID & "'  --CustomerMsgRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerMsgRefFullName & "'  --CustomerMsgRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsToBePrinted & "'  --IsToBePrinted" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'  --CustomerSalesTaxCodeRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'  --CustomerSalesTaxCodeRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineType & "'  --CreditMemoLineType" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineSeqNo & "'  --CreditMemoLineSeqNo" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineGroupLineTxnLineID & "'  --CreditMemoLineGroupLineTxnLineID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineGroupItemGroupRefListID & "'  --CreditMemoLineGroupItemGroupRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineGroupItemGroupRefFullName & "'  --CreditMemoLineGroupItemGroupRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineGroupDesc & "'  --CreditMemoLineGroupDesc" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditMemoLineGroupQuantity & "  --CreditMemoLineGroupQuantity" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineGroupIsPrintItemsInGroup & "'  --CreditMemoLineGroupIsPrintItemsInGroup" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditMemoLineGroupTotalAmount & "  --CreditMemoLineGroupTotalAmount" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineGroupSeqNo & "'  --CreditMemoLineGroupSeqNo" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineTxnLineID & "'  --CreditMemoLineTxnLineID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineItemRefListID & "'  --CreditMemoLineItemRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineItemRefFullName & "'  --CreditMemoLineItemRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineDesc & "'  --CreditMemoLineDesc" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditMemoLineQuantity & "  --CreditMemoLineQuantity" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditMemoLineRate & "  --CreditMemoLineRate" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditMemoLineRatePercent & "  --CreditMemoLineRatePercent" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLinePriceLevelRefListID & "'  --CreditMemoLinePriceLevelRefListID" & Environment.NewLine
                strSQL8 = "   , '" & str2SrcQB_CreditMemoLinePriceLevelRefFullName & "'  --CreditMemoLinePriceLevelRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineClassRefListID & "'  --CreditMemoLineClassRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineClassRefFullName & "'  --CreditMemoLineClassRefFullName" & Environment.NewLine & _
                          "   , " & str2SrcQB_CreditMemoLineAmount & "  --CreditMemoLineAmount" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineServiceDate & "'  --CreditMemoLineServiceDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineSalesTaxCodeRefListID & "'  --CreditMemoLineSalesTaxCodeRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineSalesTaxCodeRefFullName & "'  --CreditMemoLineSalesTaxCodeRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineIsTaxable & "'  --CreditMemoLineIsTaxable" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineOverrideItemAccountRefListID & "'  --CreditMemoLineOverrideItemAccountRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditMemoLineOverrideItemAccountRefFullName & "'  --CreditMemoLineOverrideItemAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_FQSaveToCache & "'  --FQSaveToCache" & Environment.NewLine & _
                          "   , '" & str2SrcQB_FQPrimaryKey & "'  --FQPrimaryKey" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldCreditMemoLineOther1 & "'  --CustomFieldCreditMemoLineOther1" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldCreditMemoLineOther2 & "'  --CustomFieldCreditMemoLineOther2" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldCreditMemoLinePriceBreaks & "'  --CustomFieldCreditMemoLinePriceBreaks" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldCreditMemoLineGroupOther1 & "'  --CustomFieldCreditMemoLineGroupOther1" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldCreditMemoLineGroupOther2 & "'  --CustomFieldCreditMemoLineGroupOther2" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldCreditMemoLineGroupPriceBreaks & "'  --CustomFieldCreditMemoLineGroupPriceBreaks" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldCreditMemoLineGroupLineOther1 & "'  --CustomFieldCreditMemoLineGroupLineOther1" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldCreditMemoLineGroupLineOther2 & "'  --CustomFieldCreditMemoLineGroupLineOther2" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldCreditMemoLineGroupLinePriceBreaks & "'  --CustomFieldCreditMemoLineGroupLinePriceBreaks" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldOther & "' ) --CustomFieldOther" & Environment.NewLine



                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6 & strSQL7 & strSQL8
                'Stop
                'Debug.Print strTableInsert

                'Execute the insert

                If gstrCompany = "DrummondPrinting" Then
                    '*cnDBPM.Execute strTableInsert
                    Dim TempCommand_3 As SqlCommand
                    TempCommand_3 = cnMax.CreateCommand()
                    TempCommand_3.CommandText = strTableInsert
                    TempCommand_3.ExecuteNonQuery()
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    Dim TempCommand_4 As SqlCommand
                    TempCommand_4 = cnMax.CreateCommand()
                    TempCommand_4.CommandText = strTableInsert
                    TempCommand_4.ExecuteNonQuery()
                End If

                ''*cnDBPM.Execute strTableInsert
                'cnMax.Execute strTableInsert




            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_CreditMemoLine  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No CreditMemoLines found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No CreditMemoLines found with the criteria given"
            '        End If
        End If


        'Moved to main routine that called this one
        ''Run the Sub that inserts all new CreditMemoLines into maximizer all at once.
        'InsertQBCustIntoMax



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QB_CreditMemoLine.Close()
            rs1MaxOfCopy_QB_CreditMemoLine = Nothing

            rs2SrcQB_QB_CreditMemoLine = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QB_CreditMemoLine.Close()
            rs3TestID_QB_CreditMemoLine = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_CreditMemoLine>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub


    Public Sub ReloadQB_Terms()
        Dim rs1MaxOfCopy_QB_Terms, rs3TestID_QB_Terms As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_Terms" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:





        'FOR PART 2SrcQB_ - Get records from QB_Terms
        Debug.WriteLine("List2SrcQB_QB_Terms")
        Dim rs2SrcQB_QB_Terms As DataSet
        Dim str2SrcQB_QB_TermsSQL, str2SrcQB_QB_TermsRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_IsActive, str2SrcQB_DayOfMonthDue, str2SrcQB_DueNextMonthDays, str2SrcQB_DiscountDayOfMonth, str2SrcQB_DiscountPct, str2SrcQB_StdDueDays, str2SrcQB_StdDiscountDays, str2SrcQB_StdDiscountPct, str2SrcQB_Type As String
        'This routine gets the 2SrcQB_QB_Terms from the database according to the selection in str2SrcQB_QB_TermsSQL.
        'It then puts those 2SrcQB_QB_Terms in the list box


        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_Terms  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_Terms"
        Application.DoEvents()




        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QB_Terms
        'SELECT * FROM Terms WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}   --USE THIS ONE
        'Debug.Print "SELECT * FROM Terms WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}"
        'Debug.Print "SELECT * FROM Terms WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'New recordset
        rs2SrcQB_QB_Terms = New DataSet()
        'str2SrcQB_QB_TermsSQL = "SELECT TOP 100 * FROM QB_Terms"
        'str2SrcQB_QB_TermsSQL = "SELECT * FROM Terms WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_TermsSQL = "SELECT * FROM Terms WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_TermsSQL = "SELECT * FROM Terms WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_TermsSQL = "SELECT * FROM Terms WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_Terms & "'} ORDER BY TimeModified"
        str2SrcQB_QB_TermsSQL = "SELECT * FROM Terms"
        Debug.WriteLine(str2SrcQB_QB_TermsSQL)
        'rs2SrcQB_QB_Terms.Open str2SrcQB_QB_TermsSQL, cnQuickBooks, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_TermsSQL, cnQuickBooks)
        rs2SrcQB_QB_Terms.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_Terms) ', adAsyncFetch '(no Optimizer)
        If rs2SrcQB_QB_Terms.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_Terms"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_Terms"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_Terms"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_Terms.Tables(0).Rows.Count) & "  QB_Terms  Records  ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_Terms.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Terms.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_Terms.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_Terms.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_ListID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_Name = ""
                str2SrcQB_IsActive = "1"
                str2SrcQB_DayOfMonthDue = ""
                str2SrcQB_DueNextMonthDays = ""
                str2SrcQB_DiscountDayOfMonth = ""
                str2SrcQB_DiscountPct = ""
                str2SrcQB_StdDueDays = ""
                str2SrcQB_StdDiscountDays = ""
                str2SrcQB_StdDiscountPct = ""
                str2SrcQB_Type = ""

                'get the columns from the database
                If iteration_row("ListID") <> "" Then str2SrcQB_ListID = iteration_row("ListID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("Name") <> "" Then str2SrcQB_Name = iteration_row("Name")
                If iteration_row("IsActive") <> "" Then str2SrcQB_IsActive = iteration_row("IsActive")
                If iteration_row("DayOfMonthDue") <> "" Then str2SrcQB_DayOfMonthDue = iteration_row("DayOfMonthDue")
                If iteration_row("DueNextMonthDays") <> "" Then str2SrcQB_DueNextMonthDays = iteration_row("DueNextMonthDays")
                If iteration_row("DiscountDayOfMonth") <> "" Then str2SrcQB_DiscountDayOfMonth = iteration_row("DiscountDayOfMonth")
                If iteration_row("DiscountPct") <> "" Then str2SrcQB_DiscountPct = iteration_row("DiscountPct")
                If iteration_row("StdDueDays") <> "" Then str2SrcQB_StdDueDays = iteration_row("StdDueDays")
                If iteration_row("StdDiscountDays") <> "" Then str2SrcQB_StdDiscountDays = iteration_row("StdDiscountDays")
                If iteration_row("StdDiscountPct") <> "" Then str2SrcQB_StdDiscountPct = iteration_row("StdDiscountPct")
                If iteration_row("Type") <> "" Then str2SrcQB_Type = iteration_row("Type")
                '        If rs2SrcQB_QB_Terms!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_Terms!CustomFieldOther

                'Strip quote character out of strings
                str2SrcQB_ListID = str2SrcQB_ListID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_Name = str2SrcQB_Name.Replace("'"c, "`"c)
                str2SrcQB_IsActive = str2SrcQB_IsActive.Replace("'"c, "`"c)
                str2SrcQB_DayOfMonthDue = str2SrcQB_DayOfMonthDue.Replace("'"c, "`"c)
                str2SrcQB_DueNextMonthDays = str2SrcQB_DueNextMonthDays.Replace("'"c, "`"c)
                str2SrcQB_DiscountDayOfMonth = str2SrcQB_DiscountDayOfMonth.Replace("'"c, "`"c)
                str2SrcQB_DiscountPct = str2SrcQB_DiscountPct.Replace("'"c, "`"c)
                str2SrcQB_StdDueDays = str2SrcQB_StdDueDays.Replace("'"c, "`"c)
                str2SrcQB_StdDiscountDays = str2SrcQB_StdDiscountDays.Replace("'"c, "`"c)
                str2SrcQB_StdDiscountPct = str2SrcQB_StdDiscountPct.Replace("'"c, "`"c)
                str2SrcQB_Type = str2SrcQB_Type.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_TermsRow = "" & _
                                        Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_Name & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_DayOfMonthDue & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_DueNextMonthDays & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_DiscountDayOfMonth & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_DiscountPct & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_StdDueDays & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_StdDiscountDays & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_StdDiscountPct & "                  ", 18) & "   " & _
                                        Strings.Left(str2SrcQB_Type & "                  ", 18) & "   " & _
                                        "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Terms.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_Terms.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_Terms.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_TermsRow)
                    'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                'DO INSERT WORK:
                Debug.WriteLine("INSERT")

                'Build the SQL string
                'MODIFICATION REQUIRED HERE
                strSQL1 = "INSERT INTO QB_Terms " & Environment.NewLine & _
                          "   ( ListID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , Name " & Environment.NewLine & _
                          "   , IsActive " & Environment.NewLine & _
                          "   , DayOfMonthDue " & Environment.NewLine & _
                          "   , DueNextMonthDays " & Environment.NewLine & _
                          "   , DiscountDayOfMonth " & Environment.NewLine & _
                          "   , DiscountPct " & Environment.NewLine & _
                          "   , StdDueDays " & Environment.NewLine & _
                          "   , StdDiscountDays " & Environment.NewLine & _
                          "   , StdDiscountPct " & Environment.NewLine & _
                          "   , Type ) " & Environment.NewLine
                strSQL2 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_ListID & "'  --ListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Name & "'  --Name" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsActive & "'  --IsActive" & Environment.NewLine & _
                          "   , '" & str2SrcQB_DayOfMonthDue & "'  --DayOfMonthDue" & Environment.NewLine & _
                          "   , '" & str2SrcQB_DueNextMonthDays & "'  --DueNextMonthDays" & Environment.NewLine & _
                          "   , '" & str2SrcQB_DiscountDayOfMonth & "'  --DiscountDayOfMonth" & Environment.NewLine & _
                          "   , '" & str2SrcQB_DiscountPct & "'  --DiscountPct" & Environment.NewLine & _
                          "   , '" & str2SrcQB_StdDueDays & "'  --StdDueDays" & Environment.NewLine & _
                          "   , '" & str2SrcQB_StdDiscountDays & "'  --StdDiscountDays" & Environment.NewLine & _
                          "   , '" & str2SrcQB_StdDiscountPct & "'  --StdDiscountPct" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Type & "' ) --Type" & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 '& strSQL3 & strSQL4 & strSQL5 & strSQL6
                'Debug.Print strTableInsert

                'Execute the insert

                If gstrCompany = "DrummondPrinting" Then
                    '*cnDBPM.Execute strTableInsert
                    Dim TempCommand_3 As SqlCommand
                    TempCommand_3 = cnMax.CreateCommand()
                    TempCommand_3.CommandText = strTableInsert
                    TempCommand_3.ExecuteNonQuery()
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    Dim TempCommand_4 As SqlCommand
                    TempCommand_4 = cnMax.CreateCommand()
                    TempCommand_4.CommandText = strTableInsert
                    TempCommand_4.ExecuteNonQuery()
                End If

                ''*cnDBPM.Execute strTableInsert
                'cnMax.Execute strTableInsert




            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_Terms  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '        End If
        End If


        'Moved to main routine that called this one
        ''Run the Sub that inserts all new Termss into maximizer all at once.
        'InsertQBCustIntoMax



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QB_Terms.Close()
            rs1MaxOfCopy_QB_Terms = Nothing

            rs2SrcQB_QB_Terms = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QB_Terms.Close()
            rs3TestID_QB_Terms = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_Terms>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub


    Public Sub ReloadQB_SalesRep()
        Dim rs1MaxOfCopy_QB_SalesRep, rs3TestID_QB_SalesRep As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_SalesRep" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:





        'FOR PART 2SrcQB_ - Get records from QB_SalesRep
        Debug.WriteLine("List2SrcQB_QB_SalesRep")
        Dim rs2SrcQB_QB_SalesRep As DataSet
        Dim str2SrcQB_QB_SalesRepSQL, str2SrcQB_QB_SalesRepRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Initial, str2SrcQB_IsActive, str2SrcQB_SalesRepEntityRefListID, str2SrcQB_SalesRepEntityRefFullName As String
        'This routine gets the 2SrcQB_QB_SalesRep from the database according to the selection in str2SrcQB_QB_SalesRepSQL.
        'It then puts those 2SrcQB_QB_SalesRep in the list box


        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_SalesRep  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_SalesRep"
        Application.DoEvents()







        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QB_SalesRep
        'SELECT * FROM SalesRep WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}   --USE THIS ONE
        'Debug.Print "SELECT * FROM SalesRep WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}"
        'Debug.Print "SELECT * FROM SalesRep WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'New recordset
        rs2SrcQB_QB_SalesRep = New DataSet()
        'str2SrcQB_QB_SalesRepSQL = "SELECT TOP 100 * FROM QB_SalesRep"
        'str2SrcQB_QB_SalesRepSQL = "SELECT * FROM SalesRep WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_SalesRepSQL = "SELECT * FROM SalesRep WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_SalesRepSQL = "SELECT * FROM SalesRep WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_SalesRepSQL = "SELECT * FROM SalesRep WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_SalesRep & "'} ORDER BY TimeModified"
        str2SrcQB_QB_SalesRepSQL = "SELECT * FROM SalesRep"
        Debug.WriteLine(str2SrcQB_QB_SalesRepSQL)
        'rs2SrcQB_QB_SalesRep.Open str2SrcQB_QB_SalesRepSQL, cnQuickBooks, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_SalesRepSQL, cnQuickBooks)
        rs2SrcQB_QB_SalesRep.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_SalesRep) ', adAsyncFetch '(no Optimizer)
        If rs2SrcQB_QB_SalesRep.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_SalesRep"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_SalesRep"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_SalesRep"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_SalesRep.Tables(0).Rows.Count) & "  QB_SalesRep  Records  ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_SalesRep.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_SalesRep.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_SalesRep.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_SalesRep.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_ListID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_Initial = ""
                str2SrcQB_IsActive = ""
                str2SrcQB_SalesRepEntityRefListID = ""
                str2SrcQB_SalesRepEntityRefFullName = ""

                'get the columns from the database
                If iteration_row("ListID") <> "" Then str2SrcQB_ListID = iteration_row("ListID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("Initial") <> "" Then str2SrcQB_Initial = iteration_row("Initial")
                If iteration_row("IsActive") <> "" Then str2SrcQB_IsActive = iteration_row("IsActive")
                If iteration_row("SalesRepEntityRefListID") <> "" Then str2SrcQB_SalesRepEntityRefListID = iteration_row("SalesRepEntityRefListID")
                If iteration_row("SalesRepEntityRefFullName") <> "" Then str2SrcQB_SalesRepEntityRefFullName = iteration_row("SalesRepEntityRefFullName")
                '        If rs2SrcQB_QB_SalesRep!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_SalesRep!CustomFieldOther

                'Strip quote character out of strings
                str2SrcQB_ListID = str2SrcQB_ListID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_Initial = str2SrcQB_Initial.Replace("'"c, "`"c)
                str2SrcQB_IsActive = str2SrcQB_IsActive.Replace("'"c, "`"c)
                str2SrcQB_SalesRepEntityRefListID = str2SrcQB_SalesRepEntityRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesRepEntityRefFullName = str2SrcQB_SalesRepEntityRefFullName.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_SalesRepRow = "" & _
                                           Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_Initial & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_SalesRepEntityRefListID & "                  ", 18) & "   " & _
                                           Strings.Left(str2SrcQB_SalesRepEntityRefFullName & "                  ", 18) & "   " & _
                                           "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_SalesRep.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_SalesRep.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_SalesRep.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_SalesRepRow)
                    'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                'DO INSERT WORK:
                Debug.WriteLine("INSERT")

                'Build the SQL string
                'MODIFICATION REQUIRED HERE
                strSQL1 = "INSERT INTO QB_SalesRep " & Environment.NewLine & _
                          "   ( ListID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , Initial " & Environment.NewLine & _
                          "   , IsActive " & Environment.NewLine & _
                          "   , SalesRepEntityRefListID " & Environment.NewLine & _
                          "   , SalesRepEntityRefFullName ) " & Environment.NewLine
                strSQL2 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_ListID & "'  --ListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Initial & "'  --Initial" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsActive & "'  --IsActive" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesRepEntityRefListID & "'  --SalesRepEntityRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesRepEntityRefFullName & "' ) --SalesRepEntityRefFullName" & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 '& strSQL3 & strSQL4 & strSQL5 & strSQL6
                'Debug.Print strTableInsert

                'Execute the insert

                If gstrCompany = "DrummondPrinting" Then
                    '*cnDBPM.Execute strTableInsert
                    Dim TempCommand_3 As SqlCommand
                    TempCommand_3 = cnMax.CreateCommand()
                    TempCommand_3.CommandText = strTableInsert
                    TempCommand_3.ExecuteNonQuery()
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    Dim TempCommand_4 As SqlCommand
                    TempCommand_4 = cnMax.CreateCommand()
                    TempCommand_4.CommandText = strTableInsert
                    TempCommand_4.ExecuteNonQuery()
                End If

                ''*cnDBPM.Execute strTableInsert
                'cnMax.Execute strTableInsert




            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_SalesRep  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '        End If
        End If


        'Moved to main routine that called this one
        ''Run the Sub that inserts all new SalesReps into maximizer all at once.
        'InsertQBCustIntoMax



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QB_SalesRep.Close()
            rs1MaxOfCopy_QB_SalesRep = Nothing

            rs2SrcQB_QB_SalesRep = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QB_SalesRep.Close()
            rs3TestID_QB_SalesRep = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_SalesRep>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub




    Public Sub ReloadQB_ItemOtherCharge()
        Dim rs1MaxOfCopy_QB_ItemOtherCharge, rs3TestID_QB_ItemOtherCharge As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_ItemOtherCharge" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:





        'FOR PART 2SrcQB_ - Get records from QB_ItemOtherCharge
        Debug.WriteLine("List2SrcQB_QB_ItemOtherCharge")
        Dim rs2SrcQB_QB_ItemOtherCharge As DataSet
        Dim str2SrcQB_QB_ItemOtherChargeSQL, str2SrcQB_QB_ItemOtherChargeRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_FullName, str2SrcQB_IsActive, str2SrcQB_ParentRefListID, str2SrcQB_ParentRefFullName, str2SrcQB_Sublevel, str2SrcQB_SalesTaxCodeRefListID, str2SrcQB_SalesTaxCodeRefFullName, str2SrcQB_SalesOrPurchaseDesc, str2SrcQB_SalesOrPurchasePrice, str2SrcQB_SalesOrPurchasePricePercent, str2SrcQB_SalesOrPurchaseAccountRefListID, str2SrcQB_SalesOrPurchaseAccountRefFullName, str2SrcQB_SalesAndPurchaseSalesDesc, str2SrcQB_SalesAndPurchaseSalesPrice, str2SrcQB_SalesAndPurchaseIncomeAccountRefListID, str2SrcQB_SalesAndPurchaseIncomeAccountRefFullName, str2SrcQB_SalesAndPurchasePurchaseDesc, str2SrcQB_SalesAndPurchasePurchaseCost, str2SrcQB_SalesAndPurchaseExpenseAccountRefListID, str2SrcQB_SalesAndPurchaseExpenseAccountRefFullName, str2SrcQB_SalesAndPurchasePrefVendorRefListID, str2SrcQB_SalesAndPurchasePrefVendorRefFullName, str2SrcQB_CustomFieldOther1, str2SrcQB_CustomFieldOther2 As String
        'This routine gets the 2SrcQB_QB_ItemOtherCharge from the database according to the selection in str2SrcQB_QB_ItemOtherChargeSQL.
        'It then puts those 2SrcQB_QB_ItemOtherCharge in the list box


        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_ItemOtherCharge  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_ItemOtherCharge"
        Application.DoEvents()







        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QB_ItemOtherCharge
        'SELECT * FROM ItemOtherCharge WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}   --USE THIS ONE
        'Debug.Print "SELECT * FROM ItemOtherCharge WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}"
        'Debug.Print "SELECT * FROM ItemOtherCharge WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'New recordset
        rs2SrcQB_QB_ItemOtherCharge = New DataSet()
        'str2SrcQB_QB_ItemOtherChargeSQL = "SELECT TOP 100 * FROM QB_ItemOtherCharge"
        'str2SrcQB_QB_ItemOtherChargeSQL = "SELECT * FROM ItemOtherCharge WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_ItemOtherChargeSQL = "SELECT * FROM ItemOtherCharge WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_ItemOtherChargeSQL = "SELECT * FROM ItemOtherCharge WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_ItemOtherChargeSQL = "SELECT * FROM ItemOtherCharge WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_ItemOtherCharge & "'} ORDER BY TimeModified"
        str2SrcQB_QB_ItemOtherChargeSQL = "SELECT * FROM ItemOtherCharge"
        Debug.WriteLine(str2SrcQB_QB_ItemOtherChargeSQL)

        'Try this to fix E_FAIL status error
        'UPGRADE_ISSUE: (2064) ADODB.CursorLocationEnum property CursorLocationEnum.adUseServer was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
        'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_ItemOtherCharge.CursorLocation was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
        'rs2SrcQB_QB_ItemOtherCharge.cursorLocation = adUseServer

        'rs2SrcQB_QB_ItemOtherCharge.Open str2SrcQB_QB_ItemOtherChargeSQL, cnQuickBooks, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
        'rs2SrcQB_QB_ItemOtherCharge.Open str2SrcQB_QB_ItemOtherChargeSQL, cnQuickBooks, adOpenForwardOnly, adLockReadOnly ', adAsyncFetch '(no Optimizer)
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_ItemOtherChargeSQL, cnQuickBooks)
        rs2SrcQB_QB_ItemOtherCharge.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_ItemOtherCharge) '<-- According to QODBC Forum

        'Commented this after E_FAIL status error fix (above)
        Dim intRecNum As Integer
        If rs2SrcQB_QB_ItemOtherCharge.Tables(0).Rows.Count > 0 Then 'ERROR: Data provider or other service returned an E_FAIL status.

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_ItemOtherCharge"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_ItemOtherCharge"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_ItemOtherCharge"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_ItemOtherCharge.Tables(0).Rows.Count) & "  QB_ItemOtherCharge  Records  ")
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  " & s & "  QB_ItemOtherCharge  Records  "

            intRecNum = 1

            For Each iteration_row As DataRow In rs2SrcQB_QB_ItemOtherCharge.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_ItemOtherCharge.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_ItemOtherCharge.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_ItemOtherCharge.Tables(0).Rows.Count) & ""
                'frmMain.lblListboxStatus.Caption = "Processing Record " & intRecNum
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & intRecNum & " of " & CStr(rs2SrcQB_QB_ItemOtherCharge.Tables(0).Rows.Count) & ""
                intRecNum += 1
                Application.DoEvents()

                'Clear strings
                str2SrcQB_ListID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_Name = ""
                str2SrcQB_FullName = ""
                str2SrcQB_IsActive = ""
                str2SrcQB_ParentRefListID = ""
                str2SrcQB_ParentRefFullName = ""
                str2SrcQB_Sublevel = ""
                str2SrcQB_SalesTaxCodeRefListID = ""
                str2SrcQB_SalesTaxCodeRefFullName = ""
                str2SrcQB_SalesOrPurchaseDesc = ""
                str2SrcQB_SalesOrPurchasePrice = ""
                str2SrcQB_SalesOrPurchasePricePercent = ""
                str2SrcQB_SalesOrPurchaseAccountRefListID = ""
                str2SrcQB_SalesOrPurchaseAccountRefFullName = ""
                str2SrcQB_SalesAndPurchaseSalesDesc = ""
                str2SrcQB_SalesAndPurchaseSalesPrice = ""
                str2SrcQB_SalesAndPurchaseIncomeAccountRefListID = ""
                str2SrcQB_SalesAndPurchaseIncomeAccountRefFullName = ""
                str2SrcQB_SalesAndPurchasePurchaseDesc = ""
                str2SrcQB_SalesAndPurchasePurchaseCost = ""
                str2SrcQB_SalesAndPurchaseExpenseAccountRefListID = ""
                str2SrcQB_SalesAndPurchaseExpenseAccountRefFullName = ""
                str2SrcQB_SalesAndPurchasePrefVendorRefListID = ""
                str2SrcQB_SalesAndPurchasePrefVendorRefFullName = ""
                str2SrcQB_CustomFieldOther1 = ""
                str2SrcQB_CustomFieldOther2 = ""

                'get the columns from the database
                If iteration_row("ListID") <> "" Then str2SrcQB_ListID = iteration_row("ListID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("Name") <> "" Then str2SrcQB_Name = iteration_row("Name")
                If iteration_row("FullName") <> "" Then str2SrcQB_FullName = iteration_row("FullName")
                If iteration_row("IsActive") <> "" Then str2SrcQB_IsActive = iteration_row("IsActive")
                If iteration_row("ParentRefListID") <> "" Then str2SrcQB_ParentRefListID = iteration_row("ParentRefListID")
                If iteration_row("ParentRefFullName") <> "" Then str2SrcQB_ParentRefFullName = iteration_row("ParentRefFullName")
                If iteration_row("Sublevel") <> "" Then str2SrcQB_Sublevel = iteration_row("Sublevel")
                If iteration_row("SalesTaxCodeRefListID") <> "" Then str2SrcQB_SalesTaxCodeRefListID = iteration_row("SalesTaxCodeRefListID")
                If iteration_row("SalesTaxCodeRefFullName") <> "" Then str2SrcQB_SalesTaxCodeRefFullName = iteration_row("SalesTaxCodeRefFullName")
                If iteration_row("SalesOrPurchaseDesc") <> "" Then str2SrcQB_SalesOrPurchaseDesc = iteration_row("SalesOrPurchaseDesc")
                If iteration_row("SalesOrPurchasePrice") <> "" Then str2SrcQB_SalesOrPurchasePrice = iteration_row("SalesOrPurchasePrice")
                If iteration_row("SalesOrPurchasePricePercent") <> "" Then str2SrcQB_SalesOrPurchasePricePercent = iteration_row("SalesOrPurchasePricePercent")
                If iteration_row("SalesOrPurchaseAccountRefListID") <> "" Then str2SrcQB_SalesOrPurchaseAccountRefListID = iteration_row("SalesOrPurchaseAccountRefListID")
                If iteration_row("SalesOrPurchaseAccountRefFullName") <> "" Then str2SrcQB_SalesOrPurchaseAccountRefFullName = iteration_row("SalesOrPurchaseAccountRefFullName")
                If iteration_row("SalesAndPurchaseSalesDesc") <> "" Then str2SrcQB_SalesAndPurchaseSalesDesc = iteration_row("SalesAndPurchaseSalesDesc")
                If iteration_row("SalesAndPurchaseSalesPrice") <> "" Then str2SrcQB_SalesAndPurchaseSalesPrice = iteration_row("SalesAndPurchaseSalesPrice")
                If iteration_row("SalesAndPurchaseIncomeAccountRefListID") <> "" Then str2SrcQB_SalesAndPurchaseIncomeAccountRefListID = iteration_row("SalesAndPurchaseIncomeAccountRefListID")
                If iteration_row("SalesAndPurchaseIncomeAccountRefFullName") <> "" Then str2SrcQB_SalesAndPurchaseIncomeAccountRefFullName = iteration_row("SalesAndPurchaseIncomeAccountRefFullName")
                If iteration_row("SalesAndPurchasePurchaseDesc") <> "" Then str2SrcQB_SalesAndPurchasePurchaseDesc = iteration_row("SalesAndPurchasePurchaseDesc")
                If iteration_row("SalesAndPurchasePurchaseCost") <> "" Then str2SrcQB_SalesAndPurchasePurchaseCost = iteration_row("SalesAndPurchasePurchaseCost")
                If iteration_row("SalesAndPurchaseExpenseAccountRefListID") <> "" Then str2SrcQB_SalesAndPurchaseExpenseAccountRefListID = iteration_row("SalesAndPurchaseExpenseAccountRefListID")
                If iteration_row("SalesAndPurchaseExpenseAccountRefFullName") <> "" Then str2SrcQB_SalesAndPurchaseExpenseAccountRefFullName = iteration_row("SalesAndPurchaseExpenseAccountRefFullName")
                If iteration_row("SalesAndPurchasePrefVendorRefListID") <> "" Then str2SrcQB_SalesAndPurchasePrefVendorRefListID = iteration_row("SalesAndPurchasePrefVendorRefListID")
                If iteration_row("SalesAndPurchasePrefVendorRefFullName") <> "" Then str2SrcQB_SalesAndPurchasePrefVendorRefFullName = iteration_row("SalesAndPurchasePrefVendorRefFullName")
                '        If rs2SrcQB_QB_ItemOtherCharge!CustomFieldOther1 <> "" Then str2SrcQB_CustomFieldOther1 = rs2SrcQB_QB_ItemOtherCharge!CustomFieldOther1
                '        If rs2SrcQB_QB_ItemOtherCharge!CustomFieldOther2 <> "" Then str2SrcQB_CustomFieldOther2 = rs2SrcQB_QB_ItemOtherCharge!CustomFieldOther2
                '        If rs2SrcQB_QB_ItemOtherCharge!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_ItemOtherCharge!CustomFieldOther

                'Strip quote character out of strings
                str2SrcQB_ListID = str2SrcQB_ListID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_Name = str2SrcQB_Name.Replace("'"c, "`"c)
                str2SrcQB_FullName = str2SrcQB_FullName.Replace("'"c, "`"c)
                str2SrcQB_IsActive = str2SrcQB_IsActive.Replace("'"c, "`"c)
                str2SrcQB_ParentRefListID = str2SrcQB_ParentRefListID.Replace("'"c, "`"c)
                str2SrcQB_ParentRefFullName = str2SrcQB_ParentRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Sublevel = str2SrcQB_Sublevel.Replace("'"c, "`"c)
                str2SrcQB_SalesTaxCodeRefListID = str2SrcQB_SalesTaxCodeRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesTaxCodeRefFullName = str2SrcQB_SalesTaxCodeRefFullName.Replace("'"c, "`"c)
                str2SrcQB_SalesOrPurchaseDesc = str2SrcQB_SalesOrPurchaseDesc.Replace("'"c, "`"c)
                str2SrcQB_SalesOrPurchasePrice = str2SrcQB_SalesOrPurchasePrice.Replace("'"c, "`"c)
                str2SrcQB_SalesOrPurchasePricePercent = str2SrcQB_SalesOrPurchasePricePercent.Replace("'"c, "`"c)
                str2SrcQB_SalesOrPurchaseAccountRefListID = str2SrcQB_SalesOrPurchaseAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesOrPurchaseAccountRefFullName = str2SrcQB_SalesOrPurchaseAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchaseSalesDesc = str2SrcQB_SalesAndPurchaseSalesDesc.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchaseSalesPrice = str2SrcQB_SalesAndPurchaseSalesPrice.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchaseIncomeAccountRefListID = str2SrcQB_SalesAndPurchaseIncomeAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchaseIncomeAccountRefFullName = str2SrcQB_SalesAndPurchaseIncomeAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchasePurchaseDesc = str2SrcQB_SalesAndPurchasePurchaseDesc.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchasePurchaseCost = str2SrcQB_SalesAndPurchasePurchaseCost.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchaseExpenseAccountRefListID = str2SrcQB_SalesAndPurchaseExpenseAccountRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchaseExpenseAccountRefFullName = str2SrcQB_SalesAndPurchaseExpenseAccountRefFullName.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchasePrefVendorRefListID = str2SrcQB_SalesAndPurchasePrefVendorRefListID.Replace("'"c, "`"c)
                str2SrcQB_SalesAndPurchasePrefVendorRefFullName = str2SrcQB_SalesAndPurchasePrefVendorRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldOther1 = str2SrcQB_CustomFieldOther1.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldOther2 = str2SrcQB_CustomFieldOther2.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_ItemOtherChargeRow = "" & _
                                                  Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_Name & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_FullName & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_ParentRefListID & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_ParentRefFullName & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_Sublevel & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_SalesTaxCodeRefListID & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_SalesTaxCodeRefFullName & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_SalesOrPurchaseDesc & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_SalesOrPurchasePrice & "                  ", 18) & "   " & _
                                                  Strings.Left(str2SrcQB_SalesOrPurchasePricePercent & "                  ", 18) & "   " & _
                                                  "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_ItemOtherCharge.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_ItemOtherCharge.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_ItemOtherCharge.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_ItemOtherChargeRow)
                    'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                'DO INSERT WORK:
                Debug.WriteLine("INSERT")

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_ItemOtherCharge " & Environment.NewLine & _
                          "   ( ListID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , Name " & Environment.NewLine & _
                          "   , FullName " & Environment.NewLine & _
                          "   , IsActive " & Environment.NewLine & _
                          "   , ParentRefListID " & Environment.NewLine & _
                          "   , ParentRefFullName " & Environment.NewLine & _
                          "   , Sublevel " & Environment.NewLine & _
                          "   , SalesTaxCodeRefListID " & Environment.NewLine & _
                          "   , SalesTaxCodeRefFullName " & Environment.NewLine & _
                          "   , SalesOrPurchaseDesc " & Environment.NewLine & _
                          "   , SalesOrPurchasePrice " & Environment.NewLine
                strSQL2 = "   , SalesOrPurchasePricePercent " & Environment.NewLine & _
                          "   , SalesOrPurchaseAccountRefListID " & Environment.NewLine & _
                          "   , SalesOrPurchaseAccountRefFullName " & Environment.NewLine & _
                          "   , SalesAndPurchaseSalesDesc " & Environment.NewLine & _
                          "   , SalesAndPurchaseSalesPrice " & Environment.NewLine & _
                          "   , SalesAndPurchaseIncomeAccountRefListID " & Environment.NewLine & _
                          "   , SalesAndPurchaseIncomeAccountRefFullName " & Environment.NewLine & _
                          "   , SalesAndPurchasePurchaseDesc " & Environment.NewLine & _
                          "   , SalesAndPurchasePurchaseCost " & Environment.NewLine & _
                          "   , SalesAndPurchaseExpenseAccountRefListID " & Environment.NewLine & _
                          "   , SalesAndPurchaseExpenseAccountRefFullName " & Environment.NewLine & _
                          "   , SalesAndPurchasePrefVendorRefListID " & Environment.NewLine & _
                          "   , SalesAndPurchasePrefVendorRefFullName " & Environment.NewLine & _
                          "   , CustomFieldOther1 " & Environment.NewLine & _
                          "   , CustomFieldOther2 ) " & Environment.NewLine
                strSQL3 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_ListID & "'  --ListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Name & "'  --Name" & Environment.NewLine & _
                          "   , '" & str2SrcQB_FullName & "'  --FullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsActive & "'  --IsActive" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ParentRefListID & "'  --ParentRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ParentRefFullName & "'  --ParentRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Sublevel & "'  --Sublevel" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesTaxCodeRefListID & "'  --SalesTaxCodeRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesTaxCodeRefFullName & "'  --SalesTaxCodeRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesOrPurchaseDesc & "'  --SalesOrPurchaseDesc" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesOrPurchasePrice & "'  --SalesOrPurchasePrice" & Environment.NewLine
                strSQL4 = "   , '" & str2SrcQB_SalesOrPurchasePricePercent & "'  --SalesOrPurchasePricePercent" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesOrPurchaseAccountRefListID & "'  --SalesOrPurchaseAccountRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesOrPurchaseAccountRefFullName & "'  --SalesOrPurchaseAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchaseSalesDesc & "'  --SalesAndPurchaseSalesDesc" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchaseSalesPrice & "'  --SalesAndPurchaseSalesPrice" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchaseIncomeAccountRefListID & "'  --SalesAndPurchaseIncomeAccountRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchaseIncomeAccountRefFullName & "'  --SalesAndPurchaseIncomeAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchasePurchaseDesc & "'  --SalesAndPurchasePurchaseDesc" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchasePurchaseCost & "'  --SalesAndPurchasePurchaseCost" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchaseExpenseAccountRefListID & "'  --SalesAndPurchaseExpenseAccountRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchaseExpenseAccountRefFullName & "'  --SalesAndPurchaseExpenseAccountRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchasePrefVendorRefListID & "'  --SalesAndPurchasePrefVendorRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SalesAndPurchasePrefVendorRefFullName & "'  --SalesAndPurchasePrefVendorRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldOther1 & "'  --CustomFieldOther1" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldOther2 & "' ) --CustomFieldOther2" & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 '& strSQL5 & strSQL6
                Debug.WriteLine(strTableInsert)

                'Execute the insert

                If gstrCompany = "DrummondPrinting" Then
                    '*cnDBPM.Execute strTableInsert
                    Dim TempCommand_3 As SqlCommand
                    TempCommand_3 = cnMax.CreateCommand()
                    TempCommand_3.CommandText = strTableInsert
                    TempCommand_3.ExecuteNonQuery()
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    Dim TempCommand_4 As SqlCommand
                    TempCommand_4 = cnMax.CreateCommand()
                    TempCommand_4.CommandText = strTableInsert
                    TempCommand_4.ExecuteNonQuery()
                End If

                ''*cnDBPM.Execute strTableInsert
                'cnMax.Execute strTableInsert




            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_ItemOtherCharge  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No items found with the criteria given"
            '        End If
        End If


        'Moved to main routine that called this one
        ''Run the Sub that inserts all new ItemOtherCharges into maximizer all at once.
        'InsertQBCustIntoMax



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QB_ItemOtherCharge.Close()
            rs1MaxOfCopy_QB_ItemOtherCharge = Nothing

            rs2SrcQB_QB_ItemOtherCharge = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QB_ItemOtherCharge.Close()
            rs3TestID_QB_ItemOtherCharge = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_ItemOtherCharge>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub



    Public Sub ReloadQB_Item()
        Dim rs1MaxOfCopy_QB_Item, rs3TestID_QB_Item As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_Item" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:





        'FOR PART 2SrcQB_ - Get records from QB_Item
        Debug.WriteLine("List2SrcQB_QB_Item")
        Dim rs2SrcQB_QB_Item As DataSet
        Dim str2SrcQB_QB_ItemSQL, str2SrcQB_QB_ItemRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_FullName, str2SrcQB_Description, str2SrcQB_IsActive, str2SrcQB_ParentRefListID, str2SrcQB_ParentRefFullName, str2SrcQB_Sublevel, str2SrcQB_Type As String
        'This routine gets the 2SrcQB_QB_Item from the database according to the selection in str2SrcQB_QB_ItemSQL.
        'It then puts those 2SrcQB_QB_Item in the list box


        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_Item  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_Item"
        Application.DoEvents()


        rs2SrcQB_QB_Item = New DataSet()
        str2SrcQB_QB_ItemSQL = "SELECT * FROM Item"
        Debug.WriteLine(str2SrcQB_QB_ItemSQL)

        
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_ItemSQL, cnQuickBooks)
        rs2SrcQB_QB_Item.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_Item) '<-- According to QODBC Forum

        Dim intRecNum As Integer
        If rs2SrcQB_QB_Item.Tables(0).Rows.Count > 0 Then 'ERROR: Data provider or other service returned an E_FAIL status.

            SQLHelper.ExecuteSQL(cnMax, "DELETE FROM QB_Item")

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_Item.Tables(0).Rows.Count) & "  QB_Item  Records  ")
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  " & s & "  QB_Item  Records  "

            intRecNum = 1

            For Each iteration_row As DataRow In rs2SrcQB_QB_Item.Tables(0).Rows

                'frmMain.lblListboxStatus.Caption = "Processing Record " & rs2SrcQB_QB_Item.tables(0).Rows.IndexOf(iteration_row) & " of " & rs2SrcQB_QB_Item.RecordCount & ""
                'frmMain.lblListboxStatus.Caption = "Processing Record " & intRecNum
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & intRecNum & " of " & CStr(rs2SrcQB_QB_Item.Tables(0).Rows.Count) & ""
                intRecNum += 1
                Application.DoEvents()

                'Clear strings
                str2SrcQB_ListID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_FullName = ""
                str2SrcQB_Description = ""
                str2SrcQB_IsActive = ""
                str2SrcQB_ParentRefListID = ""
                str2SrcQB_ParentRefFullName = ""
                str2SrcQB_Sublevel = ""
                str2SrcQB_Type = ""

                'get the columns from the database
                str2SrcQB_ListID = NCStr(iteration_row("ListID")).Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                str2SrcQB_FullName = NCStr(iteration_row("FullName")).Replace("'"c, "`"c)
                str2SrcQB_Description = NCStr(iteration_row("Description")).Replace("'"c, "`"c)
                str2SrcQB_IsActive = NCStr(iteration_row("IsActive")).Replace("'"c, "`"c)
                str2SrcQB_ParentRefListID = NCStr(iteration_row("ParentRefListID")).Replace("'"c, "`"c)
                str2SrcQB_ParentRefFullName = NCStr(iteration_row("ParentRefFullName")).Replace("'"c, "`"c)
                str2SrcQB_Sublevel = NCStr(iteration_row("Sublevel")).Replace("'"c, "`"c)
                str2SrcQB_Type = NCStr(iteration_row("Type")).Replace("'"c, "`"c)
                '        If rs2SrcQB_QB_Item!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_Item!CustomFieldOther

                'Strip quote character out of strings
                str2SrcQB_ListID = str2SrcQB_ListID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_FullName = str2SrcQB_FullName.Replace("'"c, "`"c)
                str2SrcQB_Description = str2SrcQB_Description.Replace("'"c, "`"c)
                str2SrcQB_IsActive = str2SrcQB_IsActive.Replace("'"c, "`"c)
                str2SrcQB_ParentRefListID = str2SrcQB_ParentRefListID.Replace("'"c, "`"c)
                str2SrcQB_ParentRefFullName = str2SrcQB_ParentRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Sublevel = str2SrcQB_Sublevel.Replace("'"c, "`"c)
                str2SrcQB_Type = str2SrcQB_Type.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_ItemRow = "" & _
                                       Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_FullName & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_Description & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_ParentRefListID & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_ParentRefFullName & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_Sublevel & "                  ", 18) & "   " & _
                                       Strings.Left(str2SrcQB_Type & "                  ", 18) & "   " & _
                                       "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Item.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_Item.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_Item.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_ItemRow)
                    'frmMain.lstConversionProgress.ItemData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                'DO INSERT WORK:
                Debug.WriteLine("INSERT")

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_Item " & Environment.NewLine & _
                          "   ( ListID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , FullName " & Environment.NewLine & _
                          "   , Description " & Environment.NewLine & _
                          "   , IsActive " & Environment.NewLine & _
                          "   , ParentRefListID " & Environment.NewLine & _
                          "   , ParentRefFullName " & Environment.NewLine & _
                          "   , Sublevel " & Environment.NewLine & _
                          "   , Type ) " & Environment.NewLine
                strSQL2 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_ListID & "'  --ListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , '" & str2SrcQB_FullName & "'  --FullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Description & "'  --Description" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsActive & "'  --IsActive" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ParentRefListID & "'  --ParentRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ParentRefFullName & "'  --ParentRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Sublevel & "'  --Sublevel" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Type & "' ) --Type" & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2
               
                SQLHelper.ExecuteSQL(cnMax, strTableInsert)

            Next iteration_row

            'Run the correction script
            
            SQLHelper.ExecuteSP(cnMax, "sp_QB_LoadItems")


            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


       End If


        rs1MaxOfCopy_QB_Item = Nothing

        rs2SrcQB_QB_Item = Nothing

        rs3TestID_QB_Item = Nothing


    End Sub





    Public Sub ReloadQB_PaymentMethod()
        Dim rs1MaxOfCopy_QB_PaymentMethod, rs3TestID_QB_PaymentMethod As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_PaymentMethod" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:





        'FOR PART 2SrcQB_ - Get records from QB_PaymentMethod
        Debug.WriteLine("List2SrcQB_QB_PaymentMethod")
        Dim rs2SrcQB_QB_PaymentMethod As DataSet
        Dim str2SrcQB_QB_PaymentMethodSQL, str2SrcQB_QB_PaymentMethodRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_IsActive As String
        'This routine gets the 2SrcQB_QB_PaymentMethod from the database according to the selection in str2SrcQB_QB_PaymentMethodSQL.
        'It then puts those 2SrcQB_QB_PaymentMethod in the list box


        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_PaymentMethod  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_PaymentMethod"
        Application.DoEvents()







        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QB_PaymentMethod
        'SELECT * FROM PaymentMethod WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}   --USE THIS ONE
        'Debug.Print "SELECT * FROM PaymentMethod WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}"
        'Debug.Print "SELECT * FROM PaymentMethod WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'New recordset
        rs2SrcQB_QB_PaymentMethod = New DataSet()
        'str2SrcQB_QB_PaymentMethodSQL = "SELECT TOP 100 * FROM QB_PaymentMethod"
        'str2SrcQB_QB_PaymentMethodSQL = "SELECT * FROM PaymentMethod WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_PaymentMethodSQL = "SELECT * FROM PaymentMethod WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_PaymentMethodSQL = "SELECT * FROM PaymentMethod WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_PaymentMethodSQL = "SELECT * FROM PaymentMethod WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_PaymentMethod & "'} ORDER BY TimeModified"
        str2SrcQB_QB_PaymentMethodSQL = "SELECT * FROM PaymentMethod"
        Debug.WriteLine(str2SrcQB_QB_PaymentMethodSQL)
        'rs2SrcQB_QB_PaymentMethod.Open str2SrcQB_QB_PaymentMethodSQL, cnQuickBooks, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_PaymentMethodSQL, cnQuickBooks)
        rs2SrcQB_QB_PaymentMethod.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_PaymentMethod) ', adAsyncFetch '(no Optimizer)
        If rs2SrcQB_QB_PaymentMethod.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_PaymentMethod"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_PaymentMethod"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_PaymentMethod"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_PaymentMethod.Tables(0).Rows.Count) & "  QB_PaymentMethod  Records  ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_PaymentMethod.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_PaymentMethod.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_PaymentMethod.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_PaymentMethod.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_ListID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_Name = ""
                str2SrcQB_IsActive = ""

                'get the columns from the database
                If iteration_row("ListID") <> "" Then str2SrcQB_ListID = iteration_row("ListID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("Name") <> "" Then str2SrcQB_Name = iteration_row("Name")
                If iteration_row("IsActive") <> "" Then str2SrcQB_IsActive = iteration_row("IsActive")
                '        If rs2SrcQB_QB_PaymentMethod!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_PaymentMethod!CustomFieldOther

                'Strip quote character out of strings
                str2SrcQB_ListID = str2SrcQB_ListID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_Name = str2SrcQB_Name.Replace("'"c, "`"c)
                str2SrcQB_IsActive = str2SrcQB_IsActive.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_PaymentMethodRow = "" & _
                                                Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_Name & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                                "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_PaymentMethod.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_PaymentMethod.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_PaymentMethod.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_PaymentMethodRow)
                    'frmMain.lstConversionProgress.PaymentMethodData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                'DO INSERT WORK:
                Debug.WriteLine("INSERT")

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_PaymentMethod " & Environment.NewLine & _
                          "   ( ListID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , Name " & Environment.NewLine & _
                          "   , IsActive ) " & Environment.NewLine
                strSQL2 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_ListID & "'  --ListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Name & "'  --Name" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsActive & "' ) --IsActive" & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 '& strSQL3 & strSQL4 & strSQL5 & strSQL6
                'Debug.Print strTableInsert

                'Execute the insert
                SQLHelper.ExecuteSQL(cnMax, strTableInsert)


            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_PaymentMethod  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No PaymentMethods found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No PaymentMethods found with the criteria given"
            '        End If
        End If


        'Moved to main routine that called this one
        ''Run the Sub that inserts all new PaymentMethods into maximizer all at once.
        'InsertQBCustIntoMax



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QB_PaymentMethod.Close()
            rs1MaxOfCopy_QB_PaymentMethod = Nothing

            rs2SrcQB_QB_PaymentMethod = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QB_PaymentMethod.Close()
            rs3TestID_QB_PaymentMethod = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_PaymentMethod>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub


    Public Sub ReloadQB_Account()
        Dim rs1MaxOfCopy_QB_Account, rs3TestID_QB_Account As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_Account" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:





        'FOR PART 2SrcQB_ - Get records from QB_Account
        Debug.WriteLine("List2SrcQB_QB_Account")
        Dim rs2SrcQB_QB_Account As DataSet
        Dim str2SrcQB_QB_AccountSQL, str2SrcQB_QB_AccountRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_FullName, str2SrcQB_IsActive, str2SrcQB_ParentRefListID, str2SrcQB_ParentRefFullName, str2SrcQB_Sublevel, str2SrcQB_AccountType, str2SrcQB_SpecialAccountType, str2SrcQB_AccountNumber, str2SrcQB_BankNumber, str2SrcQB_Desc, str2SrcQB_Balance, str2SrcQB_TotalBalance, str2SrcQB_TaxLineInfoRetTaxLineID, str2SrcQB_TaxLineInfoRetTaxLineName, str2SrcQB_CashFlowClassification, str2SrcQB_OpenBalance, str2SrcQB_OpenBalanceDate As String
        'This routine gets the 2SrcQB_QB_Account from the database according to the selection in str2SrcQB_QB_AccountSQL.
        'It then puts those 2SrcQB_QB_Account in the list box


        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_Account  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_Account"
        Application.DoEvents()







        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QB_Account
        'SELECT * FROM Account WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}   --USE THIS ONE
        'Debug.Print "SELECT * FROM Account WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}"
        'Debug.Print "SELECT * FROM Account WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'New recordset
        rs2SrcQB_QB_Account = New DataSet()
        'str2SrcQB_QB_AccountSQL = "SELECT TOP 100 * FROM QB_Account"
        'str2SrcQB_QB_AccountSQL = "SELECT * FROM Account WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_AccountSQL = "SELECT * FROM Account WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_AccountSQL = "SELECT * FROM Account WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_AccountSQL = "SELECT * FROM Account WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_Account & "'} ORDER BY TimeModified"
        str2SrcQB_QB_AccountSQL = "SELECT * FROM Account"
        Debug.WriteLine(str2SrcQB_QB_AccountSQL)
        'rs2SrcQB_QB_Account.Open str2SrcQB_QB_AccountSQL, cnQuickBooks, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_AccountSQL, cnQuickBooks)
        rs2SrcQB_QB_Account.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_Account) ', adAsyncFetch '(no Optimizer)
        If rs2SrcQB_QB_Account.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_Account"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_Account"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_Account"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_Account.Tables(0).Rows.Count) & "  QB_Account  Records  ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_Account.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Account.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_Account.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_Account.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_ListID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_Name = ""
                str2SrcQB_FullName = ""
                str2SrcQB_IsActive = ""
                str2SrcQB_ParentRefListID = ""
                str2SrcQB_ParentRefFullName = ""
                str2SrcQB_Sublevel = ""
                str2SrcQB_AccountType = ""
                str2SrcQB_SpecialAccountType = ""
                str2SrcQB_AccountNumber = ""
                str2SrcQB_BankNumber = ""
                str2SrcQB_Desc = ""
                str2SrcQB_Balance = ""
                str2SrcQB_TotalBalance = ""
                str2SrcQB_TaxLineInfoRetTaxLineID = ""
                str2SrcQB_TaxLineInfoRetTaxLineName = ""
                str2SrcQB_CashFlowClassification = ""
                str2SrcQB_OpenBalance = ""
                str2SrcQB_OpenBalanceDate = ""

                'get the columns from the database
                If iteration_row("ListID") <> "" Then str2SrcQB_ListID = iteration_row("ListID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("Name") <> "" Then str2SrcQB_Name = iteration_row("Name")
                If iteration_row("FullName") <> "" Then str2SrcQB_FullName = iteration_row("FullName")
                If iteration_row("IsActive") <> "" Then str2SrcQB_IsActive = iteration_row("IsActive")
                If iteration_row("ParentRefListID") <> "" Then str2SrcQB_ParentRefListID = iteration_row("ParentRefListID")
                If iteration_row("ParentRefFullName") <> "" Then str2SrcQB_ParentRefFullName = iteration_row("ParentRefFullName")
                If iteration_row("Sublevel") <> "" Then str2SrcQB_Sublevel = iteration_row("Sublevel")
                If iteration_row("AccountType") <> "" Then str2SrcQB_AccountType = iteration_row("AccountType")
                If iteration_row("SpecialAccountType") <> "" Then str2SrcQB_SpecialAccountType = iteration_row("SpecialAccountType")
                If iteration_row("AccountNumber") <> "" Then str2SrcQB_AccountNumber = iteration_row("AccountNumber")
                If iteration_row("BankNumber") <> "" Then str2SrcQB_BankNumber = iteration_row("BankNumber")
                If iteration_row("Desc") <> "" Then str2SrcQB_Desc = iteration_row("Desc")
                If iteration_row("Balance") <> "" Then str2SrcQB_Balance = iteration_row("Balance")
                If iteration_row("TotalBalance") <> "" Then str2SrcQB_TotalBalance = iteration_row("TotalBalance")
                If iteration_row("TaxLineInfoRetTaxLineID") <> "" Then str2SrcQB_TaxLineInfoRetTaxLineID = iteration_row("TaxLineInfoRetTaxLineID")
                If iteration_row("TaxLineInfoRetTaxLineName") <> "" Then str2SrcQB_TaxLineInfoRetTaxLineName = iteration_row("TaxLineInfoRetTaxLineName")
                If iteration_row("CashFlowClassification") <> "" Then str2SrcQB_CashFlowClassification = iteration_row("CashFlowClassification")
                If iteration_row("OpenBalance") <> "" Then str2SrcQB_OpenBalance = iteration_row("OpenBalance")
                If iteration_row("OpenBalanceDate") <> "" Then str2SrcQB_OpenBalanceDate = iteration_row("OpenBalanceDate")
                '        If rs2SrcQB_QB_Account!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_Account!CustomFieldOther

                'Strip quote character out of strings
                str2SrcQB_ListID = str2SrcQB_ListID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_Name = str2SrcQB_Name.Replace("'"c, "`"c)
                str2SrcQB_FullName = str2SrcQB_FullName.Replace("'"c, "`"c)
                str2SrcQB_IsActive = str2SrcQB_IsActive.Replace("'"c, "`"c)
                str2SrcQB_ParentRefListID = str2SrcQB_ParentRefListID.Replace("'"c, "`"c)
                str2SrcQB_ParentRefFullName = str2SrcQB_ParentRefFullName.Replace("'"c, "`"c)
                str2SrcQB_Sublevel = str2SrcQB_Sublevel.Replace("'"c, "`"c)
                str2SrcQB_AccountType = str2SrcQB_AccountType.Replace("'"c, "`"c)
                str2SrcQB_SpecialAccountType = str2SrcQB_SpecialAccountType.Replace("'"c, "`"c)
                str2SrcQB_AccountNumber = str2SrcQB_AccountNumber.Replace("'"c, "`"c)
                str2SrcQB_BankNumber = str2SrcQB_BankNumber.Replace("'"c, "`"c)
                str2SrcQB_Desc = str2SrcQB_Desc.Replace("'"c, "`"c)
                str2SrcQB_Balance = str2SrcQB_Balance.Replace("'"c, "`"c)
                str2SrcQB_TotalBalance = str2SrcQB_TotalBalance.Replace("'"c, "`"c)
                str2SrcQB_TaxLineInfoRetTaxLineID = str2SrcQB_TaxLineInfoRetTaxLineID.Replace("'"c, "`"c)
                str2SrcQB_TaxLineInfoRetTaxLineName = str2SrcQB_TaxLineInfoRetTaxLineName.Replace("'"c, "`"c)
                str2SrcQB_CashFlowClassification = str2SrcQB_CashFlowClassification.Replace("'"c, "`"c)
                str2SrcQB_OpenBalance = str2SrcQB_OpenBalance.Replace("'"c, "`"c)
                str2SrcQB_OpenBalanceDate = str2SrcQB_OpenBalanceDate.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_AccountRow = "" & _
                                          Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_Name & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_FullName & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_ParentRefListID & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_ParentRefFullName & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_Sublevel & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_AccountType & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_SpecialAccountType & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_AccountNumber & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_BankNumber & "                  ", 18) & "   " & _
                                          Strings.Left(str2SrcQB_Desc & "                  ", 18) & "   " & _
                                          "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Account.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_Account.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_Account.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_AccountRow)
                    'frmMain.lstConversionProgress.AccountData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                'DO INSERT WORK:
                Debug.WriteLine("INSERT")

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_Account " & Environment.NewLine & _
                          "   ( ListID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , Name " & Environment.NewLine & _
                          "   , FullName " & Environment.NewLine & _
                          "   , IsActive " & Environment.NewLine & _
                          "   , ParentRefListID " & Environment.NewLine & _
                          "   , ParentRefFullName " & Environment.NewLine & _
                          "   , Sublevel " & Environment.NewLine & _
                          "   , AccountType " & Environment.NewLine & _
                          "   , SpecialAccountType " & Environment.NewLine & _
                          "   , AccountNumber " & Environment.NewLine & _
                          "   , BankNumber " & Environment.NewLine & _
                          "   , [Desc] " & Environment.NewLine & _
                          "   , Balance " & Environment.NewLine & _
                          "   , TotalBalance " & Environment.NewLine & _
                          "   , TaxLineInfoRetTaxLineID " & Environment.NewLine & _
                          "   , TaxLineInfoRetTaxLineName " & Environment.NewLine & _
                          "   , CashFlowClassification " & Environment.NewLine & _
                          "   , OpenBalance " & Environment.NewLine & _
                          "   , OpenBalanceDate ) " & Environment.NewLine
                strSQL2 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_ListID & "'  --ListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Name & "'  --Name" & Environment.NewLine & _
                          "   , '" & str2SrcQB_FullName & "'  --FullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsActive & "'  --IsActive" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ParentRefListID & "'  --ParentRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_ParentRefFullName & "'  --ParentRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Sublevel & "'  --Sublevel" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AccountType & "'  --AccountType" & Environment.NewLine & _
                          "   , '" & str2SrcQB_SpecialAccountType & "'  --SpecialAccountType" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AccountNumber & "'  --AccountNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_BankNumber & "'  --BankNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Desc & "'  --Desc" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Balance & "'  --Balance" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TotalBalance & "'  --TotalBalance" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TaxLineInfoRetTaxLineID & "'  --TaxLineInfoRetTaxLineID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TaxLineInfoRetTaxLineName & "'  --TaxLineInfoRetTaxLineName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CashFlowClassification & "'  --CashFlowClassification" & Environment.NewLine & _
                          "   , '" & str2SrcQB_OpenBalance & "'  --OpenBalance" & Environment.NewLine & _
                          "   , '" & str2SrcQB_OpenBalanceDate & "' ) --OpenBalanceDate" & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 '& strSQL3 & strSQL4 & strSQL5 & strSQL6
                Debug.WriteLine(strTableInsert)

                'Execute the insert

                If gstrCompany = "DrummondPrinting" Then
                    ''*cnDBPM.Execute strTableInsert
                    'cnMax.Execute strTableInsert    'Error converting data type varchar to numeric.
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    'cnMax.Execute strTableInsert    'Error converting data type varchar to numeric.
                End If

                ''*cnDBPM.Execute strTableInsert
                'cnMax.Execute strTableInsert




            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_Account  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No Accounts found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No Accounts found with the criteria given"
            '        End If
        End If


        'Moved to main routine that called this one
        ''Run the Sub that inserts all new Accounts into maximizer all at once.
        'InsertQBCustIntoMax



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QB_Account.Close()
            rs1MaxOfCopy_QB_Account = Nothing

            rs2SrcQB_QB_Account = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QB_Account.Close()
            rs3TestID_QB_Account = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_Account>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub



    Public Sub ReloadQB_Vendor()
        Dim rs1MaxOfCopy_QB_Vendor, rs3TestID_QB_Vendor As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_Vendor" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:





        'FOR PART 2SrcQB_ - Get records from QB_Vendor
        Debug.WriteLine("List2SrcQB_QB_Vendor")
        Dim rs2SrcQB_QB_Vendor As DataSet
        Dim str2SrcQB_QB_VendorSQL, str2SrcQB_QB_VendorRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_IsActive, str2SrcQB_CompanyName, str2SrcQB_Salutation, str2SrcQB_FirstName, str2SrcQB_MiddleName, str2SrcQB_LastName, str2SrcQB_VendorAddressAddr1, str2SrcQB_VendorAddressAddr2, str2SrcQB_VendorAddressAddr3, str2SrcQB_VendorAddressAddr4, str2SrcQB_VendorAddressCity, str2SrcQB_VendorAddressState, str2SrcQB_VendorAddressPostalCode, str2SrcQB_VendorAddressCountry, str2SrcQB_Phone, str2SrcQB_AltPhone, str2SrcQB_Fax, str2SrcQB_Email, str2SrcQB_Contact, str2SrcQB_AltContact, str2SrcQB_NameOnCheck, str2SrcQB_AccountNumber, str2SrcQB_Notes, str2SrcQB_VendorTypeRefListID, str2SrcQB_VendorTypeRefFullName, str2SrcQB_TermsRefListID, str2SrcQB_TermsRefFullName, str2SrcQB_CreditLimit, str2SrcQB_VendorTaxIdent, str2SrcQB_IsVendorEligibleFor1099, str2SrcQB_OpenBalance, str2SrcQB_OpenBalanceDate, str2SrcQB_Balance, str2SrcQB_CustomFieldOther As String
        'This routine gets the 2SrcQB_QB_Vendor from the database according to the selection in str2SrcQB_QB_VendorSQL.
        'It then puts those 2SrcQB_QB_Vendor in the list box


        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_Vendor  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_Vendor"
        Application.DoEvents()







        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QB_Vendor
        'SELECT * FROM Vendor WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}   --USE THIS ONE
        'Debug.Print "SELECT * FROM Vendor WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}"
        'Debug.Print "SELECT * FROM Vendor WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'New recordset
        rs2SrcQB_QB_Vendor = New DataSet()
        'str2SrcQB_QB_VendorSQL = "SELECT TOP 100 * FROM QB_Vendor"
        'str2SrcQB_QB_VendorSQL = "SELECT * FROM Vendor WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_VendorSQL = "SELECT * FROM Vendor WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_VendorSQL = "SELECT * FROM Vendor WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_VendorSQL = "SELECT * FROM Vendor WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_Vendor & "'} ORDER BY TimeModified"
        str2SrcQB_QB_VendorSQL = "SELECT * FROM Vendor"
        Debug.WriteLine(str2SrcQB_QB_VendorSQL)
        'rs2SrcQB_QB_Vendor.Open str2SrcQB_QB_VendorSQL, cnQuickBooks, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_VendorSQL, cnQuickBooks)
        rs2SrcQB_QB_Vendor.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_Vendor) ', adAsyncFetch '(no Optimizer)
        If rs2SrcQB_QB_Vendor.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_Vendor"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_Vendor"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_Vendor"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_Vendor.Tables(0).Rows.Count) & "  QB_Vendor  Records  ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_Vendor.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Vendor.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_Vendor.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_Vendor.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_ListID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_Name = ""
                str2SrcQB_IsActive = ""
                str2SrcQB_CompanyName = ""
                str2SrcQB_Salutation = ""
                str2SrcQB_FirstName = ""
                str2SrcQB_MiddleName = ""
                str2SrcQB_LastName = ""
                str2SrcQB_VendorAddressAddr1 = ""
                str2SrcQB_VendorAddressAddr2 = ""
                str2SrcQB_VendorAddressAddr3 = ""
                str2SrcQB_VendorAddressAddr4 = ""
                str2SrcQB_VendorAddressCity = ""
                str2SrcQB_VendorAddressState = ""
                str2SrcQB_VendorAddressPostalCode = ""
                str2SrcQB_VendorAddressCountry = ""
                str2SrcQB_Phone = ""
                str2SrcQB_AltPhone = ""
                str2SrcQB_Fax = ""
                str2SrcQB_Email = ""
                str2SrcQB_Contact = ""
                str2SrcQB_AltContact = ""
                str2SrcQB_NameOnCheck = ""
                str2SrcQB_AccountNumber = ""
                str2SrcQB_Notes = ""
                str2SrcQB_VendorTypeRefListID = ""
                str2SrcQB_VendorTypeRefFullName = ""
                str2SrcQB_TermsRefListID = ""
                str2SrcQB_TermsRefFullName = ""
                str2SrcQB_CreditLimit = ""
                str2SrcQB_VendorTaxIdent = ""
                str2SrcQB_IsVendorEligibleFor1099 = ""
                str2SrcQB_OpenBalance = ""
                str2SrcQB_OpenBalanceDate = ""
                str2SrcQB_Balance = ""
                str2SrcQB_CustomFieldOther = ""

                'get the columns from the database
                If iteration_row("ListID") <> "" Then str2SrcQB_ListID = iteration_row("ListID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("Name") <> "" Then str2SrcQB_Name = iteration_row("Name")
                If iteration_row("IsActive") <> "" Then str2SrcQB_IsActive = iteration_row("IsActive")
                If iteration_row("CompanyName") <> "" Then str2SrcQB_CompanyName = iteration_row("CompanyName")
                If iteration_row("Salutation") <> "" Then str2SrcQB_Salutation = iteration_row("Salutation")
                If iteration_row("FirstName") <> "" Then str2SrcQB_FirstName = iteration_row("FirstName")
                If iteration_row("MiddleName") <> "" Then str2SrcQB_MiddleName = iteration_row("MiddleName")
                If iteration_row("LastName") <> "" Then str2SrcQB_LastName = iteration_row("LastName")
                If iteration_row("VendorAddressAddr1") <> "" Then str2SrcQB_VendorAddressAddr1 = iteration_row("VendorAddressAddr1")
                If iteration_row("VendorAddressAddr2") <> "" Then str2SrcQB_VendorAddressAddr2 = iteration_row("VendorAddressAddr2")
                If iteration_row("VendorAddressAddr3") <> "" Then str2SrcQB_VendorAddressAddr3 = iteration_row("VendorAddressAddr3")
                If iteration_row("VendorAddressAddr4") <> "" Then str2SrcQB_VendorAddressAddr4 = iteration_row("VendorAddressAddr4")
                If iteration_row("VendorAddressCity") <> "" Then str2SrcQB_VendorAddressCity = iteration_row("VendorAddressCity")
                If iteration_row("VendorAddressState") <> "" Then str2SrcQB_VendorAddressState = iteration_row("VendorAddressState")
                If iteration_row("VendorAddressPostalCode") <> "" Then str2SrcQB_VendorAddressPostalCode = iteration_row("VendorAddressPostalCode")
                If iteration_row("VendorAddressCountry") <> "" Then str2SrcQB_VendorAddressCountry = iteration_row("VendorAddressCountry")
                If iteration_row("Phone") <> "" Then str2SrcQB_Phone = iteration_row("Phone")
                If iteration_row("AltPhone") <> "" Then str2SrcQB_AltPhone = iteration_row("AltPhone")
                If iteration_row("Fax") <> "" Then str2SrcQB_Fax = iteration_row("Fax")
                If iteration_row("Email") <> "" Then str2SrcQB_Email = iteration_row("Email")
                If iteration_row("Contact") <> "" Then str2SrcQB_Contact = iteration_row("Contact")
                If iteration_row("AltContact") <> "" Then str2SrcQB_AltContact = iteration_row("AltContact")
                If iteration_row("NameOnCheck") <> "" Then str2SrcQB_NameOnCheck = iteration_row("NameOnCheck")
                If iteration_row("AccountNumber") <> "" Then str2SrcQB_AccountNumber = iteration_row("AccountNumber")
                If iteration_row("Notes") <> "" Then str2SrcQB_Notes = iteration_row("Notes")
                If iteration_row("VendorTypeRefListID") <> "" Then str2SrcQB_VendorTypeRefListID = iteration_row("VendorTypeRefListID")
                If iteration_row("VendorTypeRefFullName") <> "" Then str2SrcQB_VendorTypeRefFullName = iteration_row("VendorTypeRefFullName")
                If iteration_row("TermsRefListID") <> "" Then str2SrcQB_TermsRefListID = iteration_row("TermsRefListID")
                If iteration_row("TermsRefFullName") <> "" Then str2SrcQB_TermsRefFullName = iteration_row("TermsRefFullName")
                If iteration_row("CreditLimit") <> "" Then str2SrcQB_CreditLimit = iteration_row("CreditLimit")
                If iteration_row("VendorTaxIdent") <> "" Then str2SrcQB_VendorTaxIdent = iteration_row("VendorTaxIdent")
                If iteration_row("IsVendorEligibleFor1099") <> "" Then str2SrcQB_IsVendorEligibleFor1099 = iteration_row("IsVendorEligibleFor1099")
                If iteration_row("OpenBalance") <> "" Then str2SrcQB_OpenBalance = iteration_row("OpenBalance")
                If iteration_row("OpenBalanceDate") <> "" Then str2SrcQB_OpenBalanceDate = iteration_row("OpenBalanceDate")
                If iteration_row("Balance") <> "" Then str2SrcQB_Balance = iteration_row("Balance")
                If iteration_row("CustomFieldOther") <> "" Then str2SrcQB_CustomFieldOther = iteration_row("CustomFieldOther")
                '        If rs2SrcQB_QB_Vendor!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_Vendor!CustomFieldOther

                'Strip quote character out of strings
                str2SrcQB_ListID = str2SrcQB_ListID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_Name = str2SrcQB_Name.Replace("'"c, "`"c)
                str2SrcQB_IsActive = str2SrcQB_IsActive.Replace("'"c, "`"c)
                str2SrcQB_CompanyName = str2SrcQB_CompanyName.Replace("'"c, "`"c)
                str2SrcQB_Salutation = str2SrcQB_Salutation.Replace("'"c, "`"c)
                str2SrcQB_FirstName = str2SrcQB_FirstName.Replace("'"c, "`"c)
                str2SrcQB_MiddleName = str2SrcQB_MiddleName.Replace("'"c, "`"c)
                str2SrcQB_LastName = str2SrcQB_LastName.Replace("'"c, "`"c)
                str2SrcQB_VendorAddressAddr1 = str2SrcQB_VendorAddressAddr1.Replace("'"c, "`"c)
                str2SrcQB_VendorAddressAddr2 = str2SrcQB_VendorAddressAddr2.Replace("'"c, "`"c)
                str2SrcQB_VendorAddressAddr3 = str2SrcQB_VendorAddressAddr3.Replace("'"c, "`"c)
                str2SrcQB_VendorAddressAddr4 = str2SrcQB_VendorAddressAddr4.Replace("'"c, "`"c)
                str2SrcQB_VendorAddressCity = str2SrcQB_VendorAddressCity.Replace("'"c, "`"c)
                str2SrcQB_VendorAddressState = str2SrcQB_VendorAddressState.Replace("'"c, "`"c)
                str2SrcQB_VendorAddressPostalCode = str2SrcQB_VendorAddressPostalCode.Replace("'"c, "`"c)
                str2SrcQB_VendorAddressCountry = str2SrcQB_VendorAddressCountry.Replace("'"c, "`"c)
                str2SrcQB_Phone = str2SrcQB_Phone.Replace("'"c, "`"c)
                str2SrcQB_AltPhone = str2SrcQB_AltPhone.Replace("'"c, "`"c)
                str2SrcQB_Fax = str2SrcQB_Fax.Replace("'"c, "`"c)
                str2SrcQB_Email = str2SrcQB_Email.Replace("'"c, "`"c)
                str2SrcQB_Contact = str2SrcQB_Contact.Replace("'"c, "`"c)
                str2SrcQB_AltContact = str2SrcQB_AltContact.Replace("'"c, "`"c)
                str2SrcQB_NameOnCheck = str2SrcQB_NameOnCheck.Replace("'"c, "`"c)
                str2SrcQB_AccountNumber = str2SrcQB_AccountNumber.Replace("'"c, "`"c)
                str2SrcQB_Notes = str2SrcQB_Notes.Replace("'"c, "`"c)
                str2SrcQB_VendorTypeRefListID = str2SrcQB_VendorTypeRefListID.Replace("'"c, "`"c)
                str2SrcQB_VendorTypeRefFullName = str2SrcQB_VendorTypeRefFullName.Replace("'"c, "`"c)
                str2SrcQB_TermsRefListID = str2SrcQB_TermsRefListID.Replace("'"c, "`"c)
                str2SrcQB_TermsRefFullName = str2SrcQB_TermsRefFullName.Replace("'"c, "`"c)
                str2SrcQB_CreditLimit = str2SrcQB_CreditLimit.Replace("'"c, "`"c)
                str2SrcQB_VendorTaxIdent = str2SrcQB_VendorTaxIdent.Replace("'"c, "`"c)
                str2SrcQB_IsVendorEligibleFor1099 = str2SrcQB_IsVendorEligibleFor1099.Replace("'"c, "`"c)
                str2SrcQB_OpenBalance = str2SrcQB_OpenBalance.Replace("'"c, "`"c)
                str2SrcQB_OpenBalanceDate = str2SrcQB_OpenBalanceDate.Replace("'"c, "`"c)
                str2SrcQB_Balance = str2SrcQB_Balance.Replace("'"c, "`"c)
                str2SrcQB_CustomFieldOther = str2SrcQB_CustomFieldOther.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")
                str2SrcQB_IsVendorEligibleFor1099 = IIf(str2SrcQB_IsVendorEligibleFor1099 = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_VendorRow = "" & _
                                         Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_Name & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_CompanyName & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_Salutation & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_FirstName & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_MiddleName & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_LastName & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_VendorAddressAddr1 & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_VendorAddressAddr2 & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_VendorAddressAddr3 & "                  ", 18) & "   " & _
                                         Strings.Left(str2SrcQB_VendorAddressAddr4 & "                  ", 18) & "   " & _
                                         "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Vendor.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_Vendor.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_Vendor.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_VendorRow)
                    'frmMain.lstConversionProgress.VendorData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                'DO INSERT WORK:
                Debug.WriteLine("INSERT")

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_Vendor " & Environment.NewLine & _
                          "   ( ListID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , Name " & Environment.NewLine & _
                          "   , IsActive " & Environment.NewLine & _
                          "   , CompanyName " & Environment.NewLine & _
                          "   , Salutation " & Environment.NewLine & _
                          "   , FirstName " & Environment.NewLine & _
                          "   , MiddleName " & Environment.NewLine & _
                          "   , LastName " & Environment.NewLine & _
                          "   , VendorAddressAddr1 " & Environment.NewLine & _
                          "   , VendorAddressAddr2 " & Environment.NewLine & _
                          "   , VendorAddressAddr3 " & Environment.NewLine & _
                          "   , VendorAddressAddr4 " & Environment.NewLine & _
                          "   , VendorAddressCity " & Environment.NewLine & _
                          "   , VendorAddressState " & Environment.NewLine & _
                          "   , VendorAddressPostalCode " & Environment.NewLine & _
                          "   , VendorAddressCountry " & Environment.NewLine
                strSQL2 = "   , Phone " & Environment.NewLine & _
                          "   , AltPhone " & Environment.NewLine & _
                          "   , Fax " & Environment.NewLine & _
                          "   , Email " & Environment.NewLine & _
                          "   , Contact " & Environment.NewLine & _
                          "   , AltContact " & Environment.NewLine & _
                          "   , NameOnCheck " & Environment.NewLine & _
                          "   , AccountNumber " & Environment.NewLine & _
                          "   , Notes " & Environment.NewLine & _
                          "   , VendorTypeRefListID " & Environment.NewLine & _
                          "   , VendorTypeRefFullName " & Environment.NewLine & _
                          "   , TermsRefListID " & Environment.NewLine & _
                          "   , TermsRefFullName " & Environment.NewLine & _
                          "   , CreditLimit " & Environment.NewLine & _
                          "   , VendorTaxIdent " & Environment.NewLine & _
                          "   , IsVendorEligibleFor1099 " & Environment.NewLine & _
                          "   , OpenBalance " & Environment.NewLine & _
                          "   , OpenBalanceDate " & Environment.NewLine & _
                          "   , Balance " & Environment.NewLine & _
                          "   , CustomFieldOther ) " & Environment.NewLine
                strSQL3 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_ListID & "'  --ListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Name & "'  --Name" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsActive & "'  --IsActive" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CompanyName & "'  --CompanyName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Salutation & "'  --Salutation" & Environment.NewLine & _
                          "   , '" & str2SrcQB_FirstName & "'  --FirstName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_MiddleName & "'  --MiddleName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_LastName & "'  --LastName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorAddressAddr1 & "'  --VendorAddressAddr1" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorAddressAddr2 & "'  --VendorAddressAddr2" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorAddressAddr3 & "'  --VendorAddressAddr3" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorAddressAddr4 & "'  --VendorAddressAddr4" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorAddressCity & "'  --VendorAddressCity" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorAddressState & "'  --VendorAddressState" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorAddressPostalCode & "'  --VendorAddressPostalCode" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorAddressCountry & "'  --VendorAddressCountry" & Environment.NewLine
                strSQL4 = "   , '" & str2SrcQB_Phone & "'  --Phone" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AltPhone & "'  --AltPhone" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Fax & "'  --Fax" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Email & "'  --Email" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Contact & "'  --Contact" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AltContact & "'  --AltContact" & Environment.NewLine & _
                          "   , '" & str2SrcQB_NameOnCheck & "'  --NameOnCheck" & Environment.NewLine & _
                          "   , '" & str2SrcQB_AccountNumber & "'  --AccountNumber" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Notes & "'  --Notes" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorTypeRefListID & "'  --VendorTypeRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorTypeRefFullName & "'  --VendorTypeRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TermsRefListID & "'  --TermsRefListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TermsRefFullName & "'  --TermsRefFullName" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CreditLimit & "'  --CreditLimit" & Environment.NewLine & _
                          "   , '" & str2SrcQB_VendorTaxIdent & "'  --VendorTaxIdent" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsVendorEligibleFor1099 & "'  --IsVendorEligibleFor1099" & Environment.NewLine & _
                          "   , '" & str2SrcQB_OpenBalance & "'  --OpenBalance" & Environment.NewLine & _
                          "   , '" & str2SrcQB_OpenBalanceDate & "'  --OpenBalanceDate" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Balance & "'  --Balance" & Environment.NewLine & _
                          "   , '" & str2SrcQB_CustomFieldOther & "' ) --CustomFieldOther" & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 '& strSQL5 & strSQL6
                Debug.WriteLine(strTableInsert)

                'Execute the insert

                If gstrCompany = "DrummondPrinting" Then
                    ''*cnDBPM.Execute strTableInsert
                    'cnMax.Execute strTableInsert    'Error converting data type varchar to numeric.
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    'cnMax.Execute strTableInsert    'Error converting data type varchar to numeric.
                End If

                ''*cnDBPM.Execute strTableInsert
                'cnMax.Execute strTableInsert




            Next iteration_row

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_Vendor  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No Vendors found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No Vendors found with the criteria given"
            '        End If
        End If


        'Moved to main routine that called this one
        ''Run the Sub that inserts all new Vendors into maximizer all at once.
        'InsertQBCustIntoMax



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QB_Vendor.Close()
            rs1MaxOfCopy_QB_Vendor = Nothing

            rs2SrcQB_QB_Vendor = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QB_Vendor.Close()
            rs3TestID_QB_Vendor = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_Vendor>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub



    Public Sub ReloadQB_StandardTerms()
        Dim rs1MaxOfCopy_QB_StandardTerms, rs3TestID_QB_StandardTerms As Object

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_StandardTerms" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'Error handling
        If gbooUseErrorHandling Then
            'UPGRADE_TODO: (1065) Error handling statement (On Error Goto) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("On Error Goto Label (ErrorFunc)")
        End If
        GoTo RunCode
ErrorFunc:
        If HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "") = "RN" Then
            'UPGRADE_TODO: (1065) Error handling statement (Resume Next) could not be converted. More Information: http://www.vbtonet.com/ewis/ewi1065.aspx
            UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("Resume Next Statement")
        Else
            Exit Sub
        End If
RunCode:





        'FOR PART 2SrcQB_ - Get records from QB_StandardTerms
        Debug.WriteLine("List2SrcQB_QB_StandardTerms")
        Dim rs2SrcQB_QB_StandardTerms As DataSet
        Dim str2SrcQB_QB_StandardTermsSQL, str2SrcQB_QB_StandardTermsRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_IsActive, str2SrcQB_StdDueDays, str2SrcQB_StdDiscountDays, str2SrcQB_DiscountPct As String
        'This routine gets the 2SrcQB_QB_StandardTerms from the database according to the selection in str2SrcQB_QB_StandardTermsSQL.
        'It then puts those 2SrcQB_QB_StandardTerms in the list box


        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String


        'On Error GoTo SubError

        'frmMain.lstConversionProgress.Clear

        'Show what's processing
        frmMain.DefInstance.lblListboxStatus.Text = "" & DateTimeHelper.ToString(DateTime.Now) & "   Processing  QB_StandardTerms  Records "
        frmMain.DefInstance.lblStatus.Text = "RefreshQB -Processing  QB_StandardTerms"
        Application.DoEvents()







        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QB_StandardTerms
        'SELECT * FROM StandardTerms WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}   --USE THIS ONE
        'Debug.Print "SELECT * FROM StandardTerms WHERE TimeModified > {ts '2006-04-11 13:33:02.000'}"
        'Debug.Print "SELECT * FROM StandardTerms WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'}"
        'New recordset
        rs2SrcQB_QB_StandardTerms = New DataSet()
        'str2SrcQB_QB_StandardTermsSQL = "SELECT TOP 100 * FROM QB_StandardTerms"
        'str2SrcQB_QB_StandardTermsSQL = "SELECT * FROM StandardTerms WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} ORDER BY TimeModified"
        'str2SrcQB_QB_StandardTermsSQL = "SELECT * FROM StandardTerms WHERE TimeModified >= {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_StandardTermsSQL = "SELECT * FROM StandardTerms WHERE TimeModified > {ts '" & str1MaxOfCopy_TimeModified & "'} "
        'str2SrcQB_QB_StandardTermsSQL = "SELECT * FROM StandardTerms WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_StandardTerms & "'} ORDER BY TimeModified"
        str2SrcQB_QB_StandardTermsSQL = "SELECT * FROM StandardTerms"
        Debug.WriteLine(str2SrcQB_QB_StandardTermsSQL)
        'rs2SrcQB_QB_StandardTerms.Open str2SrcQB_QB_StandardTermsSQL, cnQuickBooks, adOpenForwardOnly, adLockReadOnly, adAsyncFetch
        Dim adap As Odbc.OdbcDataAdapter = New Odbc.OdbcDataAdapter(str2SrcQB_QB_StandardTermsSQL, cnQuickBooks)
        rs2SrcQB_QB_StandardTerms.Tables.Clear()
        adap.Fill(rs2SrcQB_QB_StandardTerms) ', adAsyncFetch '(no Optimizer)
        If rs2SrcQB_QB_StandardTerms.Tables(0).Rows.Count > 0 Then

            'Clear out table
            If gstrCompany = "DrummondPrinting" Then
                '*cnDBPM.Execute "DELETE FROM QB_StandardTerms"
                Dim TempCommand As SqlCommand
                TempCommand = cnMax.CreateCommand()
                TempCommand.CommandText = "DELETE FROM QB_StandardTerms"
                TempCommand.ExecuteNonQuery()
            ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                Dim TempCommand_2 As SqlCommand
                TempCommand_2 = cnMax.CreateCommand()
                TempCommand_2.CommandText = "DELETE FROM QB_StandardTerms"
                TempCommand_2.ExecuteNonQuery()
            End If

            'Show what's processing in the listbox
            frmMain.DefInstance.lstConversionProgress.AddItem("" & DateTimeHelper.ToString(DateTime.Now) & "     Processing  " & CStr(rs2SrcQB_QB_StandardTerms.Tables(0).Rows.Count) & "  QB_StandardTerms  Records  ")

            For Each iteration_row As DataRow In rs2SrcQB_QB_StandardTerms.Tables(0).Rows

                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_StandardTerms.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                frmMain.DefInstance.lblListboxStatus.Text = "Processing Record " & rs2SrcQB_QB_StandardTerms.tables(0).Rows.IndexOf(iteration_row) & " of " & CStr(rs2SrcQB_QB_StandardTerms.Tables(0).Rows.Count) & ""
                Application.DoEvents()

                'Clear strings
                str2SrcQB_ListID = ""
                str2SrcQB_TimeCreated = ""
                str2SrcQB_TimeModified = ""
                str2SrcQB_EditSequence = ""
                str2SrcQB_Name = ""
                str2SrcQB_IsActive = ""
                str2SrcQB_StdDueDays = ""
                str2SrcQB_StdDiscountDays = ""
                str2SrcQB_DiscountPct = "0"

                'get the columns from the database
                If iteration_row("ListID") <> "" Then str2SrcQB_ListID = iteration_row("ListID")
                If iteration_row("TimeCreated") <> "" Then str2SrcQB_TimeCreated = iteration_row("TimeCreated")
                If iteration_row("TimeModified") <> "" Then str2SrcQB_TimeModified = iteration_row("TimeModified")
                If iteration_row("EditSequence") <> "" Then str2SrcQB_EditSequence = iteration_row("EditSequence")
                If iteration_row("Name") <> "" Then str2SrcQB_Name = iteration_row("Name")
                If iteration_row("IsActive") <> "" Then str2SrcQB_IsActive = iteration_row("IsActive")
                If iteration_row("StdDueDays") <> "" Then str2SrcQB_StdDueDays = iteration_row("StdDueDays")
                If iteration_row("StdDiscountDays") <> "" Then str2SrcQB_StdDiscountDays = iteration_row("StdDiscountDays")
                If iteration_row("DiscountPct") <> "" Then str2SrcQB_DiscountPct = iteration_row("DiscountPct")
                '        If rs2SrcQB_QB_StandardTerms!CustomFieldOther <> "" Then str2SrcQB_CustomFieldOther = rs2SrcQB_QB_StandardTerms!CustomFieldOther

                'Strip quote character out of strings
                str2SrcQB_ListID = str2SrcQB_ListID.Replace("'"c, "`"c)
                str2SrcQB_TimeCreated = str2SrcQB_TimeCreated.Replace("'"c, "`"c)
                str2SrcQB_TimeModified = str2SrcQB_TimeModified.Replace("'"c, "`"c)
                str2SrcQB_EditSequence = str2SrcQB_EditSequence.Replace("'"c, "`"c)
                str2SrcQB_Name = str2SrcQB_Name.Replace("'"c, "`"c)
                str2SrcQB_IsActive = str2SrcQB_IsActive.Replace("'"c, "`"c)
                str2SrcQB_StdDueDays = str2SrcQB_StdDueDays.Replace("'"c, "`"c)
                str2SrcQB_StdDiscountDays = str2SrcQB_StdDiscountDays.Replace("'"c, "`"c)
                str2SrcQB_DiscountPct = str2SrcQB_DiscountPct.Replace("'"c, "`"c)


                'Change flags back to binary
                str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")



                'Put the information together into a string
                'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                str2SrcQB_QB_StandardTermsRow = "" & _
                                                Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_Name & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_StdDueDays & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_StdDiscountDays & "                  ", 18) & "   " & _
                                                Strings.Left(str2SrcQB_DiscountPct & "                  ", 18) & "   " & _
                                                "" & Strings.Chr(9)

                'put the line in the listbox
                'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_StandardTerms.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                Debug.WriteLine(DateTimeHelper.ToString(DateTime.Now) & "   " & CStr(rs2SrcQB_QB_StandardTerms.tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_StandardTerms.Tables(0).Rows.Count))
                If frmMain.DefInstance.chkSeeProcessing.CheckState = CheckState.Checked Then
                    frmMain.DefInstance.lstConversionProgress.AddItem("2SrcQB_   " & DateTimeHelper.ToString(DateTime.Now) & "   " & str2SrcQB_QB_StandardTermsRow)
                    'frmMain.lstConversionProgress.StandardTermsData(frmMain.lstConversionProgress.NewIndex) = str2SrcQB_ListID
                    ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)
                End If




                'DO WORK: With each record


                'DO INSERT WORK:
                Debug.WriteLine("INSERT")

                'Build the SQL string
                strSQL1 = "INSERT INTO QB_StandardTerms " & Environment.NewLine & _
                          "   ( ListID " & Environment.NewLine & _
                          "   , TimeCreated " & Environment.NewLine & _
                          "   , TimeModified " & Environment.NewLine & _
                          "   , EditSequence " & Environment.NewLine & _
                          "   , Name " & Environment.NewLine & _
                          "   , IsActive " & Environment.NewLine & _
                          "   , StdDueDays " & Environment.NewLine & _
                          "   , StdDiscountDays " & Environment.NewLine & _
                          "   , DiscountPct ) " & Environment.NewLine
                strSQL2 = "VALUES " & Environment.NewLine & _
                          "   ( '" & str2SrcQB_ListID & "'  --ListID" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeCreated & "'  --TimeCreated" & Environment.NewLine & _
                          "   , '" & str2SrcQB_TimeModified & "'  --TimeModified" & Environment.NewLine & _
                          "   , '" & str2SrcQB_EditSequence & "'  --EditSequence" & Environment.NewLine & _
                          "   , '" & str2SrcQB_Name & "'  --Name" & Environment.NewLine & _
                          "   , '" & str2SrcQB_IsActive & "'  --IsActive" & Environment.NewLine & _
                          "   , '" & str2SrcQB_StdDueDays & "'  --StdDueDays" & Environment.NewLine & _
                          "   , '" & str2SrcQB_StdDiscountDays & "'  --StdDiscountDays" & Environment.NewLine & _
                          "   , '" & str2SrcQB_DiscountPct & "' ) --DiscountPct" & Environment.NewLine


                'Combine the strings
                strTableInsert = strSQL1 & strSQL2 '& strSQL3 & strSQL4 & strSQL5 & strSQL6
                'Debug.Print strTableInsert

                'Execute the insert

                If gstrCompany = "DrummondPrinting" Then
                    '*cnDBPM.Execute strTableInsert
                    Dim TempCommand_3 As SqlCommand
                    TempCommand_3 = cnMax.CreateCommand()
                    TempCommand_3.CommandText = strTableInsert
                    TempCommand_3.ExecuteNonQuery()
                ElseIf gstrCompany = "FrazzledAndBedazzled" Then
                    Dim TempCommand_4 As SqlCommand
                    TempCommand_4 = cnMax.CreateCommand()
                    TempCommand_4.CommandText = strTableInsert
                    TempCommand_4.ExecuteNonQuery()
                End If

                ''*cnDBPM.Execute strTableInsert
                'cnMax.Execute strTableInsert




            Next iteration_row

            'Run the correction script
            'cnDBPM.Execute "exec sp_QB_LoadTerms"
            Dim TempCommand_5 As SqlCommand
            TempCommand_5 = cnMax.CreateCommand()
            TempCommand_5.CommandText = "exec sp_QB_LoadTerms"
            TempCommand_5.ExecuteNonQuery()

            frmMain.DefInstance.lstConversionProgress.AddItem("")
            ListBoxHelper.SetSelected(frmMain.DefInstance.lstConversionProgress, frmMain.DefInstance.lstConversionProgress.Items.Count - 1, True)


        Else

            'Show what's NOT processing in the listbox
            'frmMain.lstConversionProgress.AddItem "" & Now & "     Processing  0  QB_StandardTerms  Records  "

            '        If frmMain.chkSeeProcessing.Value = 1 Then
            '            frmMain.lstConversionProgress.AddItem "No StandardTermss found with the criteria given"
            '            'frmMain.lstConversionProgress.AddItem txtTypeRadNum
            '            'frmMain.lstConversionProgress.AddItem "No StandardTermss found with the criteria given"
            '        End If
        End If


        'Moved to main routine that called this one
        ''Run the Sub that inserts all new StandardTermss into maximizer all at once.
        'InsertQBCustIntoMax



        'UPGRADE_TODO: (1069) Error handling statement (On Error Resume Next) was converted to a pattern that might have a different behavior. More Information: http://www.vbtonet.com/ewis/ewi1069.aspx
        Try
            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs1MaxOfCopy_QB_StandardTerms.Close()
            rs1MaxOfCopy_QB_StandardTerms = Nothing

            rs2SrcQB_QB_StandardTerms = Nothing

            'UPGRADE_TODO: (1067) Member Close is not defined in type Variant. More Information: http://www.vbtonet.com/ewis/ewi1067.aspx
            rs3TestID_QB_StandardTerms.Close()
            rs3TestID_QB_StandardTerms = Nothing


            Exit Sub


            MessageBox.Show("<<RefreshQB_StandardTerms>> " & Information.Err().Description, Application.ProductName)

        Catch exc As System.Exception
            NotUpgradedHelper.NotifyNotUpgradedElement("Resume in On-Error-Resume-Next Block")
        End Try

    End Sub
End Module