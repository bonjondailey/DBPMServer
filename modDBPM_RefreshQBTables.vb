Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Windows.Forms
Imports DBPM_Server.siteConstants

Module modDBPM_RefreshQBTables

	Public gstrQBMaxTimeModified_Invoice As String = ""
	Public gstrQBMaxTimeModified_InvoiceLine As String = ""
	Public gstrQBMaxTimeModified_ReceivePayment As String = ""
	Public gstrQBMaxTimeModified_ReceivePaymentLine As String = ""
	Public gstrQBMaxTimeModified_CreditMemo As String = ""
    Public gstrQBMaxTimeModified_Terms As String = ""


    '**********************************
    '*** FIRST CODE REVIEW COMPLETE ***
    '**********************************


	Public Sub RefreshQBTables()

		'Permission and ErrorHandling          (Auto built)
		Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
		Dim strSubName As String = "RefreshQBTables" '"SUBNAME"

		'Check permission to run
		If Not HavePermission(strObjName, strSubName) Then Exit Sub
        If frmMain.chkPauseProcessing.CheckState = CheckState.Checked Then Exit Sub

        If Not cnQuickBooks Is Nothing AndAlso cnQuickBooks.State = ConnectionState.Open Then
            booQBFileIsOpen = True
        Else
            OpenConnectionQB()
        End If
        'Set flag
        booQBRefreshInProgress = True

        'Show what's processing in the listbox
        ShowUserMessage(strSubName, "Start Running RefreshQBTables", "Start Running RefreshQBTables", True)

        GetQBMaxTimeModified() 'FIRST CODE_UPDATE COMPLETE

        InsertMaxBillToIntoQB()  'FIRST CODE_UPDATE COMPLETE
        RefreshQB_Terms()
        ReloadQB_ShipMethod(False)  'FIRST CODE_UPDATE COMPLETE
        RefreshQB_ItemOtherCharge()  'FIRST CODE_UPDATE COMPLETE
        RefreshQB_ReceivePayment() 'FIRST CODE_UPDATE COMPLETE
        RefreshQB_ReceivePaymentLine() 'FIRST CODE_UPDATE COMPLETE
        RefreshQB_CreditMemo() 'FIRST CODE_UPDATE COMPLETE
        RefreshQB_Invoice() ''FIRST CODE_UPDATE COMPLETE
        RefreshQB_InvoiceLine() ''FIRST CODE_UPDATE COMPLETE
        RefreshQB_Customer() 'FIRST CODE_UPDATE COMPLETE
        'InsertQBCustIntoMax()


        If gCustomerBalanceUpdateList.Count > 0 Then
            UpdateQBCustomerBalanceList(gCustomerBalanceUpdateList)
        End If

        'run this again...
        InsertMaxBillToIntoQB() 'FIRST CODE_UPDATE COMPLETE

        Try
            SQLHelper.ExecuteSP(cnDBPM, "sp_TEMP_MarkCustPromoRush")
        Catch ex As Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
            Exit Try
        End Try

        ShowUserMessage(strSubName, "Finished RefreshQBTables", "Finished RefreshQBTables", True)

        booQBRefreshInProgress = False


    End Sub

    Public Sub ReloadQB_ShipMethod(Optional ByVal forceShipMethodReload As Boolean = True)

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_ReloadQBTables" '"OBJNAME"
        Dim strSubName As String = "ReloadQB_ShipMethod" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

      
        'FOR PART 2SrcQB_ - Get records from QB_ShipMethod
        Debug.WriteLine("List2SrcQB_QB_ShipMethod")
        Dim rsDBPM_QBShipMethod As DataSet

        Dim str2SrcQB_QB_ShipMethodRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_IsActive As String
        'This routine gets the 2SrcQB_QB_ShipMethod from the database according to the selection in str2SrcQB_QB_ShipMethodSQL.
        'It then puts those 2SrcQB_QB_ShipMethod in the list box


        'dim SQL strings
        Dim strSQL1, strSQL2, strTableInsert As String
      
        ShowUserMessage(strObjName, "Processing  QB_ShipMethod  Records ", "RefreshQB -Processing  QB_ShipMethod", True)

        Using rs2SrcQB_QB_ShipMethod As DataSet = New DataSet()

            Dim strMaxShipMethodTimeModified As String = ""

            Dim sqlCheckMaxShipUpdateDate As String = ""
            sqlCheckMaxShipUpdateDate = "SELECT Max(TimeModified) FROM QB_ShipMethod"
            strMaxShipMethodTimeModified = SQLHelper.ExecuteScalerDate(cnMax, CommandType.Text, sqlCheckMaxShipUpdateDate, "yyyy-MM-dd HH:mm:ss.000")

            Dim curRow As Integer = 0
            Dim totalRows As Integer = 0

            Dim adap_2 As New OdbcDataAdapter("SELECT * FROM ShipMethod WHERE TimeModified > {ts '" & strMaxShipMethodTimeModified & "'}", cnQuickBooks)

            rs2SrcQB_QB_ShipMethod.Tables.Clear()
            adap_2.Fill(rs2SrcQB_QB_ShipMethod)
            totalRows = rs2SrcQB_QB_ShipMethod.Tables(0).Rows.Count

            If totalRows > 0 Then

                Dim shipList As New List(Of String)
                Using dr As SqlDataReader = SQLHelper.ExecuteReader(cnMax, CommandType.Text, "SELECT ListID from QB_ShipMethod")
                    While dr.Read
                        shipList.Add(NCStr(dr("ListID")))
                    End While
                End Using

                For Each iteration_row As DataRow In rs2SrcQB_QB_ShipMethod.Tables(0).Rows
                    Try

                   
                    'Show what's processing in the listbox
                    curRow += 1

                    ShowUserMessage(strSubName, "Processing record " & curRow.ToString & " of " & totalRows.ToString)

                    'Clear strings
                    str2SrcQB_ListID = ""
                    str2SrcQB_TimeCreated = ""
                    str2SrcQB_TimeModified = ""
                    str2SrcQB_EditSequence = ""
                    str2SrcQB_Name = ""
                    str2SrcQB_IsActive = ""

                    'get the columns from the database
                    str2SrcQB_ListID = NCStr(iteration_row("ListID"))
                    str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated"))
                    str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified"))
                    str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence"))
                    str2SrcQB_Name = NCStr(iteration_row("Name")).Replace("'"c, "`"c)
                    str2SrcQB_IsActive = NCStr(iteration_row("IsActive"))
                    str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")



                    'Put the information together into a string
                    'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                    str2SrcQB_QB_ShipMethodRow = "" & _
                                                    Strings.Left(str2SrcQB_ListID & "                  ", 18) & "   " & _
                                                    Strings.Left(str2SrcQB_TimeCreated & "                  ", 18) & "   " & _
                                                    Strings.Left(str2SrcQB_TimeModified & "                  ", 18) & "   " & _
                                                    Strings.Left(str2SrcQB_EditSequence & "                  ", 18) & "   " & _
                                                    Strings.Left(str2SrcQB_Name & "                  ", 18) & "   " & _
                                                    Strings.Left(str2SrcQB_IsActive & "                  ", 18) & "   " & _
                                                    "" & Strings.Chr(9)

                    ShowUserMessage(strSubName, "2SrcQB_" & str2SrcQB_QB_ShipMethodRow)

                    If shipList.Contains(str2SrcQB_ListID) Then
                        'exists
                        strSQL1 = "UPDATE QB_ShipMethod SET TimeCreated = '" & str2SrcQB_TimeCreated & "', TimeModified = '" & str2SrcQB_TimeModified & "', EditSequence = '" & str2SrcQB_EditSequence & "', Name = '" & str2SrcQB_Name & "', IsActive = '" & str2SrcQB_IsActive & "' WHERE ListID = '" & str2SrcQB_ListID & "'"
                        strSQL2 = ""
                    Else
                        'new one
                        strSQL1 = "INSERT INTO QB_ShipMethod " & Environment.NewLine & _
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
                    End If

                    strTableInsert = strSQL1 & strSQL2 '& strSQL3 & strSQL4 & strSQL5 & strSQL6
                    SQLHelper.ExecuteSQL(cnMax, strTableInsert)


                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For
                    End Try
                Next iteration_row

            End If
            ShowUserMessage(strObjName, "Finished Processing  QB_ShipMethod  Records ", , True)
            rsDBPM_QBShipMethod = Nothing

        End Using
    End Sub


    Public Sub RefreshQB_ItemOtherCharge()
        'FINISHED FIRST RUN_THROUGH

        Dim gstrQBMaxTimeModified_ItemOtherCharge As String = ""

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQB_ItemOtherCharge" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub


        'FOR PART 1MaxOfCopy_ - Get records from QB_ItemOtherCharge
        Debug.WriteLine("List1MaxOfCopy_QB_ItemOtherCharge")
        Dim str1MaxOfCopy_QB_ItemOtherChargeSQL, str1MaxOfCopy_TimeModified As String
        'This routine gets the 1MaxOfCopy_QB_ItemOtherCharge from the database according to the selection in str1MaxOfCopy_QB_ItemOtherChargeSQL.
        'It then puts those 1MaxOfCopy_QB_ItemOtherCharge in the list box

        'FOR PART 2SrcQB_ - Get records from QB_ItemOtherCharge
        Debug.WriteLine("List2SrcQB_QB_ItemOtherCharge")

        Dim str2SrcQB_QB_ItemOtherChargeSQL, str2SrcQB_QB_ItemOtherChargeRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_FullName, str2SrcQB_IsActive, str2SrcQB_ParentRefListID, str2SrcQB_ParentRefFullName, str2SrcQB_Sublevel, str2SrcQB_SalesTaxCodeRefListID, str2SrcQB_SalesTaxCodeRefFullName, str2SrcQB_SalesOrPurchaseDesc, str2SrcQB_SalesOrPurchasePrice, str2SrcQB_SalesOrPurchasePricePercent, str2SrcQB_SalesOrPurchaseAccountRefListID, str2SrcQB_SalesOrPurchaseAccountRefFullName, str2SrcQB_SalesAndPurchaseSalesDesc, str2SrcQB_SalesAndPurchaseSalesPrice, str2SrcQB_SalesAndPurchaseIncomeAccountRefListID, str2SrcQB_SalesAndPurchaseIncomeAccountRefFullName, str2SrcQB_SalesAndPurchasePurchaseDesc, str2SrcQB_SalesAndPurchasePurchaseCost, str2SrcQB_SalesAndPurchaseExpenseAccountRefListID, str2SrcQB_SalesAndPurchaseExpenseAccountRefFullName, str2SrcQB_SalesAndPurchasePrefVendorRefListID, str2SrcQB_SalesAndPurchasePrefVendorRefFullName As String
        'This routine gets the 2SrcQB_QB_ItemOtherCharge from the database according to the selection in str2SrcQB_QB_ItemOtherChargeSQL.
        'It then puts those 2SrcQB_QB_ItemOtherCharge in the list box

        'FOR PART 3TestID_ - Get records from QB_ItemOtherCharge
        Debug.WriteLine("List3TestID_QB_ItemOtherCharge")
        'This routine gets the 3TestID_QB_ItemOtherCharge from the database according to the selection in str3TestID_QB_ItemOtherChargeSQL.
        'It then puts those 3TestID_QB_ItemOtherCharge in the list box

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strTableInsert, strTableUpdate As String

        ShowUserMessage(strSubName, "RefreshQB: QB_ItemOtherCharge Records", "RefreshQB: QB_ItemOtherCharge Records", True)

        'ItemOtherCharge
        'PART 1MaxOfCopy_: Get the latest record from SQL
        str1MaxOfCopy_QB_ItemOtherChargeSQL = "SELECT max(TimeModified) TimeModified FROM QB_ItemOtherCharge"
        'Debug.Print str1MaxOfCopy_QB_ItemOtherChargeSQL

        str1MaxOfCopy_TimeModified = SQLHelper.ExecuteScalerDate(cnMax, CommandType.Text, str1MaxOfCopy_QB_ItemOtherChargeSQL, "yyyy-MM-dd HH:mm:ss.000")
        gstrQBMaxTimeModified_ItemOtherCharge = str1MaxOfCopy_TimeModified

        Using rs2SrcQB_QB_ItemOtherCharge As New DataSet()


            str2SrcQB_QB_ItemOtherChargeSQL = "SELECT * FROM ItemOtherCharge WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_ItemOtherCharge & "'}" ' ORDER BY TimeModified"
            Using adap_3 As New OdbcDataAdapter(str2SrcQB_QB_ItemOtherChargeSQL, cnQuickBooks)
                rs2SrcQB_QB_ItemOtherCharge.Tables.Clear()
                adap_3.Fill(rs2SrcQB_QB_ItemOtherCharge)
            End Using

            Dim intRecNum As Integer = 0
            Dim rowCount As Integer = rs2SrcQB_QB_ItemOtherCharge.Tables(0).Rows.Count
            Dim processedRecCount As Integer = 0

            If rowCount > 0 Then

                ShowUserMessage(strSubName, "Processing  " & rowCount.ToString & "  QB_ItemOtherCharge Records  ")

                For Each iteration_row As DataRow In rs2SrcQB_QB_ItemOtherCharge.Tables(0).Rows
                    intRecNum += 1

                    Try

                        ShowUserMessage(strSubName, "Processing Record " & intRecNum.ToString & " of " & rowCount.ToString & "")

                        'Clear strings
                        'TODO check default values for strings to make sure we don;t need "0" or other items
                        str2SrcQB_ListID = NCStr(iteration_row!ListID).Replace("'"c, "`"c)
                        str2SrcQB_TimeCreated = NCStr(iteration_row!TimeCreated).Replace("'"c, "`"c)
                        str2SrcQB_TimeModified = NCStr(iteration_row!TimeModified).Replace("'"c, "`"c)
                        str2SrcQB_EditSequence = NCStr(iteration_row!EditSequence).Replace("'"c, "`"c)
                        str2SrcQB_Name = NCStr(iteration_row!Name).Replace("'"c, "`"c)
                        str2SrcQB_FullName = NCStr(iteration_row!FullName).Replace("'"c, "`"c)
                        str2SrcQB_IsActive = NCStr(iteration_row!IsActive).Replace("'"c, "`"c)
                        str2SrcQB_ParentRefListID = NCStr(iteration_row!ParentRefListID).Replace("'"c, "`"c)
                        str2SrcQB_ParentRefFullName = NCStr(iteration_row!ParentRefFullName).Replace("'"c, "`"c)
                        str2SrcQB_Sublevel = NCInt(iteration_row!Sublevel)
                        str2SrcQB_SalesTaxCodeRefListID = NCStr(iteration_row!SalesTaxCodeRefListID).Replace("'"c, "`"c)
                        str2SrcQB_SalesTaxCodeRefFullName = NCStr(iteration_row!SalesTaxCodeRefFullName).Replace("'"c, "`"c)
                        str2SrcQB_SalesOrPurchaseDesc = NCStr(iteration_row!SalesOrPurchaseDesc).Replace("'"c, "`"c)
                        str2SrcQB_SalesOrPurchasePrice = NCStr(iteration_row!SalesOrPurchasePrice).Replace("'"c, "`"c)
                        str2SrcQB_SalesOrPurchasePricePercent = NCStr(iteration_row!SalesOrPurchasePricePercent).Replace("'"c, "`"c)
                        str2SrcQB_SalesOrPurchaseAccountRefListID = NCStr(iteration_row!SalesOrPurchaseAccountRefListID).Replace("'"c, "`"c)
                        str2SrcQB_SalesOrPurchaseAccountRefFullName = NCStr(iteration_row!SalesOrPurchaseAccountRefFullName).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchaseSalesDesc = NCStr(iteration_row!SalesAndPurchaseSalesDesc).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchaseSalesPrice = NCStr(iteration_row!SalesAndPurchaseSalesPrice).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchaseIncomeAccountRefListID = NCStr(iteration_row!SalesAndPurchaseIncomeAccountRefListID).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchaseIncomeAccountRefFullName = NCStr(iteration_row!SalesAndPurchaseIncomeAccountRefFullName).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchasePurchaseDesc = NCStr(iteration_row!SalesAndPurchasePurchaseDesc).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchasePurchaseCost = NCStr(iteration_row!SalesAndPurchasePurchaseCost).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchaseExpenseAccountRefListID = NCStr(iteration_row!SalesAndPurchaseExpenseAccountRefListID).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchaseExpenseAccountRefFullName = NCStr(iteration_row!SalesAndPurchaseExpenseAccountRefFullName).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchasePrefVendorRefListID = NCStr(iteration_row!SalesAndPurchasePrefVendorRefListID).Replace("'"c, "`"c)
                        str2SrcQB_SalesAndPurchasePrefVendorRefFullName = NCStr(iteration_row!SalesAndPurchasePrefVendorRefFullName).Replace("'"c, "`"c)

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

                        Dim iRowCount As Integer = 0
                        iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(ListID) FROM QB_ItemOtherCharge WHERE ListID = '" & str2SrcQB_ListID & "'")
                        'If iRowCount > 1 Then Stop 'Should only be one
                        If iRowCount = 1 Then 'record exists  -UPDATE
                            'DO UPDATE WORK:
                            Debug.WriteLine("UPDATE")

                            'Build the SQL string
                            strSQL1 = "UPDATE  " & Environment.NewLine & _
                                      "       QB_ItemOtherCharge " & Environment.NewLine & _
                                      "SET " & Environment.NewLine & _
                                      "      ListID = '" & str2SrcQB_ListID & "'" & Environment.NewLine & _
                                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                      "     , Name = '" & str2SrcQB_Name & "'" & Environment.NewLine & _
                                      "     , FullName = '" & str2SrcQB_FullName & "'" & Environment.NewLine & _
                                      "     , IsActive = '" & str2SrcQB_IsActive & "'" & Environment.NewLine & _
                                      "     , ParentRefListID = '" & str2SrcQB_ParentRefListID & "'" & Environment.NewLine & _
                                      "     , ParentRefFullName = '" & str2SrcQB_ParentRefFullName & "'" & Environment.NewLine & _
                                      "     , Sublevel = '" & str2SrcQB_Sublevel & "'" & Environment.NewLine & _
                                      "     , SalesTaxCodeRefListID = '" & str2SrcQB_SalesTaxCodeRefListID & "'" & Environment.NewLine & _
                                      "     , SalesTaxCodeRefFullName = '" & str2SrcQB_SalesTaxCodeRefFullName & "'" & Environment.NewLine & _
                                      "     , SalesOrPurchaseDesc = '" & str2SrcQB_SalesOrPurchaseDesc & "'" & Environment.NewLine & _
                                      "     , SalesOrPurchasePrice = '" & str2SrcQB_SalesOrPurchasePrice & "'" & Environment.NewLine & _
                                      "     , SalesOrPurchasePricePercent = '" & str2SrcQB_SalesOrPurchasePricePercent & "'" & Environment.NewLine & _
                                      "     , SalesOrPurchaseAccountRefListID = '" & str2SrcQB_SalesOrPurchaseAccountRefListID & "'" & Environment.NewLine & _
                                      "     , SalesOrPurchaseAccountRefFullName = '" & str2SrcQB_SalesOrPurchaseAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , SalesAndPurchaseSalesDesc = '" & str2SrcQB_SalesAndPurchaseSalesDesc & "'" & Environment.NewLine & _
                                      "     , SalesAndPurchaseSalesPrice = '" & str2SrcQB_SalesAndPurchaseSalesPrice & "'" & Environment.NewLine & _
                                      "     , SalesAndPurchaseIncomeAccountRefListID = '" & str2SrcQB_SalesAndPurchaseIncomeAccountRefListID & "'" & Environment.NewLine & _
                                      "     , SalesAndPurchaseIncomeAccountRefFullName = '" & str2SrcQB_SalesAndPurchaseIncomeAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , SalesAndPurchasePurchaseDesc = '" & str2SrcQB_SalesAndPurchasePurchaseDesc & "'" & Environment.NewLine
                            strSQL2 = "     , SalesAndPurchasePurchaseCost = '" & str2SrcQB_SalesAndPurchasePurchaseCost & "'" & Environment.NewLine & _
                                      "     , SalesAndPurchaseExpenseAccountRefListID = '" & str2SrcQB_SalesAndPurchaseExpenseAccountRefListID & "'" & Environment.NewLine & _
                                      "     , SalesAndPurchaseExpenseAccountRefFullName = '" & str2SrcQB_SalesAndPurchaseExpenseAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , SalesAndPurchasePrefVendorRefListID = '" & str2SrcQB_SalesAndPurchasePrefVendorRefListID & "'" & Environment.NewLine & _
                                      "     , SalesAndPurchasePrefVendorRefFullName = '" & str2SrcQB_SalesAndPurchasePrefVendorRefFullName & "'" & Environment.NewLine & _
                                      "WHERE " & Environment.NewLine & _
                                      "       ListID = '" & str2SrcQB_ListID & "'" & Environment.NewLine

                            'Combine the strings
                            strTableUpdate = strSQL1 & strSQL2 '& strSQL3 & strSQL4 '& strSQL5 & strSQL6

                            'Execute the update
                            SQLHelper.ExecuteSQL(cnMax, strTableUpdate)


                        Else
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
                                      "   , SalesAndPurchasePrefVendorRefFullName)" & Environment.NewLine
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
                                      "   , '" & str2SrcQB_SalesAndPurchasePrefVendorRefFullName & "') --CustomFieldOther2" & Environment.NewLine


                            strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4
                            SQLHelper.ExecuteSQL(cnMax, strTableInsert)
                            processedRecCount += 1
                        End If

                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For

                    End Try
                Next iteration_row

            End If
            ShowUserMessage(strSubName, "Finished Processing ItemOtherCharge - Processed " & processedRecCount.ToString & " of " & rowCount.ToString & " records", "Finished Processing ItemOtherCharge", True)

        End Using

    End Sub



    Public Sub GetQBMaxTimeModified()
        'FINISHED FIRST RUN_THROUGH


        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "GetQBMaxTimeModified" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        ShowUserMessage(strSubName, "RefreshQB: Get Max Time Modified", "RefreshQB: Get Max Time Modified")

        Try
            Debug.WriteLine("List1MaxOfCopy_QB_Customer")
            Dim str1MaxOfCopy_QB_CustomerSQL As String

            str1MaxOfCopy_QB_CustomerSQL = "SELECT max(TimeModified) TimeModified FROM QB_Invoice"
            gstrQBMaxTimeModified_Invoice = SQLHelper.ExecuteScalerDate(cnMax, CommandType.Text, str1MaxOfCopy_QB_CustomerSQL, "yyyy-MM-dd HH:mm:ss.000")

            str1MaxOfCopy_QB_CustomerSQL = "SELECT MAX(TimeModified) TimeModified FROM QB_InvoiceLine"
            gstrQBMaxTimeModified_InvoiceLine = SQLHelper.ExecuteScalerDate(cnMax, CommandType.Text, str1MaxOfCopy_QB_CustomerSQL, "yyyy-MM-dd HH:mm:ss.000")

            str1MaxOfCopy_QB_CustomerSQL = "SELECT max(TimeModified) TimeModified FROM QB_ReceivePayment"
            gstrQBMaxTimeModified_ReceivePayment = SQLHelper.ExecuteScalerDate(cnMax, CommandType.Text, str1MaxOfCopy_QB_CustomerSQL, "yyyy-MM-dd HH:mm:ss.000")

            str1MaxOfCopy_QB_CustomerSQL = "SELECT max(TimeModified) TimeModified FROM QB_ReceivePaymentLine"
            gstrQBMaxTimeModified_ReceivePaymentLine = SQLHelper.ExecuteScalerDate(cnMax, CommandType.Text, str1MaxOfCopy_QB_CustomerSQL, "yyyy-MM-dd HH:mm:ss.000")

            str1MaxOfCopy_QB_CustomerSQL = "SELECT max(TimeModified) TimeModified FROM QB_CreditMemo"
            gstrQBMaxTimeModified_CreditMemo = SQLHelper.ExecuteScalerDate(cnMax, CommandType.Text, str1MaxOfCopy_QB_CustomerSQL, "yyyy-MM-dd HH:mm:ss.000")

            str1MaxOfCopy_QB_CustomerSQL = "SELECT max(TimeModified) TimeModified FROM QB_Customer"
            gstrQBMaxTimeModified_Customer = SQLHelper.ExecuteScalerDate(cnMax, CommandType.Text, str1MaxOfCopy_QB_CustomerSQL, "yyyy-MM-dd HH:mm:ss")

            str1MaxOfCopy_QB_CustomerSQL = "SELECT max(TimeModified) TimeModified FROM QB_Terms"
            gstrQBMaxTimeModified_Terms = SQLHelper.ExecuteScalerDate(cnMax, CommandType.Text, str1MaxOfCopy_QB_CustomerSQL, "yyyy-MM-dd HH:mm:ss.000")
        Catch ex As Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "")
        End Try

    End Sub


    Public Sub UpdateQBReceivePaymentLine(ByRef strCustomerRefFullName As String)
        'FINISHED WITH FIRST RUN_THROUGH


        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "UpdateQBReceivePaymentLine" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        ShowUserMessage(strSubName, "Updating QB with Receive Payment Lines", strSubName, True)

        ''This routine gets the 1MaxOfCopy_QBTable from the database according to the selection in str1MaxOfCopy_QBTableSQL.
        ''It then puts those 1MaxOfCopy_QBTable in the list box

        'FOR PART 2SrcQB_ - Get records from QB_ReceivePaymentLine
        Debug.WriteLine("List2SrcQB_QB_ReceivePaymentLine")

        Dim str2SrcQB_QB_ReceivePaymentLineSQL, str2SrcQB_QB_ReceivePaymentLineRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_TotalAmount, str2SrcQB_PaymentMethodRefListID, str2SrcQB_PaymentMethodRefFullName, str2SrcQB_Memo, str2SrcQB_DepositToAccountRefListID, str2SrcQB_DepositToAccountRefFullName, str2SrcQB_CreditCardTxnInfoInputCreditCardNumber, str2SrcQB_CreditCardTxnInfoInputExpirationMonth, str2SrcQB_CreditCardTxnInfoInputExpirationYear, str2SrcQB_CreditCardTxnInfoInputNameOnCard, str2SrcQB_CreditCardTxnInfoInputCreditCardAddress, str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode, str2SrcQB_CreditCardTxnInfoInputCommercialCardCode, str2SrcQB_CreditCardTxnInfoResultResultCode, str2SrcQB_CreditCardTxnInfoResultResultMessage, str2SrcQB_CreditCardTxnInfoResultCreditCardTransID, str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber, str2SrcQB_CreditCardTxnInfoResultAuthorizationCode, str2SrcQB_CreditCardTxnInfoResultAVSStreet, str2SrcQB_CreditCardTxnInfoResultAVSZip, str2SrcQB_CreditCardTxnInfoResultReconBatchID, str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode, str2SrcQB_CreditCardTxnInfoResultPaymentStatus, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp, str2SrcQB_IsAutoApply, str2SrcQB_UnusedPayment, str2SrcQB_UnusedCredits, str2SrcQB_AppliedToTxnTxnID, str2SrcQB_AppliedToTxnPaymentAmount, str2SrcQB_AppliedToTxnTxnType, str2SrcQB_AppliedToTxnTxnDate, str2SrcQB_AppliedToTxnRefNumber, str2SrcQB_AppliedToTxnBalanceRemaining, str2SrcQB_AppliedToTxnAmount, str2SrcQB_AppliedToTxnSetCreditCreditTxnID, str2SrcQB_AppliedToTxnSetCreditAppliedAmount, str2SrcQB_AppliedToTxnDiscountAmount, str2SrcQB_AppliedToTxnDiscountAccountRefListID, str2SrcQB_AppliedToTxnDiscountAccountRefFullName, str2SrcQB_FQSaveToCache, str2SrcQB_FQPrimaryKey As String
        'This routine gets the 2SrcQB_QB_ReceivePaymentLine from the database according to the selection in str2SrcQB_QB_ReceivePaymentLineSQL.
        'It then puts those 2SrcQB_QB_ReceivePaymentLine in the list box

        'FOR PART 3TestID_
        Debug.WriteLine("List3TestID_QBTable")

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String

        Dim strNowMinusThirtyDays As String = ""
        strNowMinusThirtyDays = DateTime.Now.AddDays(-60).ToString("yyyy-MM-dd HH:mm:ss.000s")


        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        Using rs2SrcQB_QB_ReceivePaymentLine As DataSet = New DataSet()


            str2SrcQB_QB_ReceivePaymentLineSQL = "SELECT * FROM ReceivePaymentLine WHERE CustomerRefFullName = '" & strCustomerRefFullName & "' AND TimeModified > {ts '" & DateTime.Now.AddDays(-10).ToString("yyyy-MM-dd HH:mm:ss.000") & "'}"
            Debug.WriteLine(str2SrcQB_QB_ReceivePaymentLineSQL)
            Using adap As New OdbcDataAdapter(str2SrcQB_QB_ReceivePaymentLineSQL, cnQuickBooks)
                rs2SrcQB_QB_ReceivePaymentLine.Tables.Clear()
                adap.Fill(rs2SrcQB_QB_ReceivePaymentLine)
            End Using


            If rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows.Count > 10 Then
                ShowUserMessage(strSubName, "Too many RPL items found (more than 10)")
                Exit Sub
            End If

            Dim curRow As Integer = 0
            Dim rowCount As Integer = rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows.Count

            If rowCount > 0 Then


                For Each iteration_row As DataRow In rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows
                    curRow += 1

                    ShowUserMessage(strSubName, "Processing Record " & curRow.ToString & " of " & rowCount.ToString)

                    'get the columns from the database
                    str2SrcQB_TxnID = NCStr(iteration_row("TxnID")).Replace("'"c, "`"c)
                    str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                    str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                    str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                    str2SrcQB_TxnNumber = NCStr(iteration_row("TxnNumber"), "0").Replace("'"c, "`"c)
                    str2SrcQB_CustomerRefListID = NCStr(iteration_row("CustomerRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_CustomerRefFullName = NCStr(iteration_row("CustomerRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_ARAccountRefListID = NCStr(iteration_row("ARAccountRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_ARAccountRefFullName = NCStr(iteration_row("ARAccountRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_TxnDate = NCStr(iteration_row("TxnDate")).Replace("'"c, "`"c)
                    str2SrcQB_TxnDateMacro = NCStr(iteration_row("TxnDateMacro")).Replace("'"c, "`"c)
                    str2SrcQB_RefNumber = NCStr(iteration_row("RefNumber")).Replace("'"c, "`"c)
                    str2SrcQB_TotalAmount = NCStr(iteration_row("TotalAmount"), "0").Replace("'"c, "`"c)
                    str2SrcQB_PaymentMethodRefListID = NCStr(iteration_row("PaymentMethodRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_PaymentMethodRefFullName = NCStr(iteration_row("PaymentMethodRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_Memo = NCStr(iteration_row("Memo")).Replace("'"c, "`"c)
                    str2SrcQB_DepositToAccountRefListID = NCStr(iteration_row("DepositToAccountRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_DepositToAccountRefFullName = NCStr(iteration_row("DepositToAccountRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoInputCreditCardNumber = NCStr(iteration_row("CreditCardTxnInfoInputCreditCardNumber")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoInputExpirationMonth = NCStr(iteration_row("CreditCardTxnInfoInputExpirationMonth"), "0").Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoInputExpirationYear = NCStr(iteration_row("CreditCardTxnInfoInputExpirationYear"), "0").Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoInputNameOnCard = NCStr(iteration_row("CreditCardTxnInfoInputNameOnCard")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoInputCreditCardAddress = NCStr(iteration_row("CreditCardTxnInfoInputCreditCardAddress")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode = NCStr(iteration_row("CreditCardTxnInfoInputCreditCardPostalCode")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoInputCommercialCardCode = NCStr(iteration_row("CreditCardTxnInfoInputCommercialCardCode")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultResultCode = NCStr(iteration_row("CreditCardTxnInfoResultResultCode"), "0").Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultResultMessage = NCStr(iteration_row("CreditCardTxnInfoResultResultMessage")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultCreditCardTransID = NCStr(iteration_row("CreditCardTxnInfoResultCreditCardTransID")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber = NCStr(iteration_row("CreditCardTxnInfoResultMerchantAccountNumber")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultAuthorizationCode = NCStr(iteration_row("CreditCardTxnInfoResultAuthorizationCode")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultAVSStreet = NCStr(iteration_row("CreditCardTxnInfoResultAVSStreet")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultAVSZip = NCStr(iteration_row("CreditCardTxnInfoResultAVSZip")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultReconBatchID = NCStr(iteration_row("CreditCardTxnInfoResultReconBatchID")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode = NCStr(iteration_row("CreditCardTxnInfoResultPaymentGroupingCode"), "0").Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultPaymentStatus = NCStr(iteration_row("CreditCardTxnInfoResultPaymentStatus")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime = NCStr(iteration_row("CreditCardTxnInfoResultTxnAuthorizationTime")).Replace("'"c, "`"c)
                    str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp = NCStr(iteration_row("CreditCardTxnInfoResultTxnAuthorizationStamp"), "0").Replace("'"c, "`"c)
                    str2SrcQB_IsAutoApply = NCStr(iteration_row("IsAutoApply")).Replace("'"c, "`"c)
                    str2SrcQB_UnusedPayment = NCStr(iteration_row("UnusedPayment"), "0").Replace("'"c, "`"c)
                    str2SrcQB_UnusedCredits = NCStr(iteration_row("UnusedCredits"), "0").Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnTxnID = NCStr(iteration_row("AppliedToTxnTxnID")).Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnPaymentAmount = NCStr(iteration_row("AppliedToTxnPaymentAmount"), "0").Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnTxnType = NCStr(iteration_row("AppliedToTxnTxnType")).Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnTxnDate = NCStr(iteration_row("AppliedToTxnTxnDate")).Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnRefNumber = NCStr(iteration_row("AppliedToTxnRefNumber")).Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnBalanceRemaining = NCStr(iteration_row("AppliedToTxnBalanceRemaining"), "0").Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnAmount = NCStr(iteration_row("AppliedToTxnAmount"), "0").Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnSetCreditCreditTxnID = NCStr(iteration_row("AppliedToTxnSetCreditCreditTxnID")).Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnSetCreditAppliedAmount = NCStr(iteration_row("AppliedToTxnSetCreditAppliedAmount"), "0").Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnDiscountAmount = NCStr(iteration_row("AppliedToTxnDiscountAmount"), "0").Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnDiscountAccountRefListID = NCStr(iteration_row("AppliedToTxnDiscountAccountRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_AppliedToTxnDiscountAccountRefFullName = NCStr(iteration_row("AppliedToTxnDiscountAccountRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_FQSaveToCache = NCStr(iteration_row("FQSaveToCache")).Replace("'"c, "`"c)
                    str2SrcQB_FQPrimaryKey = NCStr(iteration_row("FQPrimaryKey")).Replace("'"c, "`"c)

                    'Change flags back to binary
                    str2SrcQB_IsAutoApply = IIf(str2SrcQB_IsAutoApply = "True", "1", "0")
                    str2SrcQB_FQSaveToCache = IIf(str2SrcQB_FQSaveToCache = "True", "1", "0")

                    'Put the information together into a string
                    'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                    str2SrcQB_QB_ReceivePaymentLineRow = "" & _
                                                         Strings.Left("RPL Upd" & "                  ", 18) & "   " & _
                                                         Strings.Left(str2SrcQB_TimeCreated & "                  ", 16) & "   " & _
                                                         Strings.Left(str2SrcQB_TimeModified & "                  ", 16) & "   " & _
                                                         Strings.Left(str2SrcQB_CustomerRefFullName & "                  ", 18) & "   " & _
                                                         Strings.Left(str2SrcQB_TxnDate & "                  ", 10) & "   " & _
                                                         Strings.Left(str2SrcQB_RefNumber & "                  ", 10) & "   " & _
                                                         Strings.Left(str2SrcQB_TotalAmount & "                  ", 18) & "   " & _
                                                         Strings.Left(str2SrcQB_PaymentMethodRefFullName & "                  ", 18) & "   " & _
                                                         "" & Strings.Chr(9)

                    'Left(str2SrcQB_TxnID + "                  ", 18) & "   " & _
                    '
                    'put the line in the listbox


                    ShowUserMessage(strSubName, "Processing Record " & curRow.ToString & " of " & rowCount.ToString)
                    ShowUserMessage(strSubName, str2SrcQB_QB_ReceivePaymentLineRow)

                    'Check to see if ListID or TxnID is in QBTable            'Yes then UPDATE record            'No then INSERT record
                    Dim iRowCount As Integer = 0
                    iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(FQPrimaryKey) FROM QB_ReceivePaymentLine WHERE FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'")
                    If iRowCount > 1 Then Stop 'Should only be one
                    If iRowCount = 1 Then 'record exists  -UPDATE

                        'DO UPDATE WORK:
                        Debug.WriteLine("UPDATE")

                        'Build the SQL string
                        strSQL1 = "UPDATE  " & Environment.NewLine & _
                                  "       QB_ReceivePaymentLine " & Environment.NewLine & _
                                  "SET " & Environment.NewLine & _
                                  "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine & _
                                  "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                  "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                  "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                  "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & Environment.NewLine & _
                                  "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & Environment.NewLine & _
                                  "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & Environment.NewLine & _
                                  "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & Environment.NewLine & _
                                  "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & Environment.NewLine & _
                                  "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & Environment.NewLine & _
                                  "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & Environment.NewLine & _
                                  "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & Environment.NewLine & _
                                  "     , TotalAmount = " & str2SrcQB_TotalAmount & "" & Environment.NewLine & _
                                  "     , PaymentMethodRefListID = '" & str2SrcQB_PaymentMethodRefListID & "'" & Environment.NewLine & _
                                  "     , PaymentMethodRefFullName = '" & str2SrcQB_PaymentMethodRefFullName & "'" & Environment.NewLine & _
                                  "     , Memo = '" & str2SrcQB_Memo & "'" & Environment.NewLine & _
                                  "     , DepositToAccountRefListID = '" & str2SrcQB_DepositToAccountRefListID & "'" & Environment.NewLine
                        strSQL2 = "     , DepositToAccountRefFullName = '" & str2SrcQB_DepositToAccountRefFullName & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoInputCreditCardNumber = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardNumber & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoInputExpirationMonth = " & str2SrcQB_CreditCardTxnInfoInputExpirationMonth & "" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoInputExpirationYear = " & str2SrcQB_CreditCardTxnInfoInputExpirationYear & "" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoInputNameOnCard = '" & str2SrcQB_CreditCardTxnInfoInputNameOnCard & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoInputCreditCardAddress = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardAddress & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoInputCreditCardPostalCode = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoInputCommercialCardCode = '" & str2SrcQB_CreditCardTxnInfoInputCommercialCardCode & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultResultCode = " & str2SrcQB_CreditCardTxnInfoResultResultCode & "" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultResultMessage = '" & str2SrcQB_CreditCardTxnInfoResultResultMessage & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultCreditCardTransID = '" & str2SrcQB_CreditCardTxnInfoResultCreditCardTransID & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultMerchantAccountNumber = '" & str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultAuthorizationCode = '" & str2SrcQB_CreditCardTxnInfoResultAuthorizationCode & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultAVSStreet = '" & str2SrcQB_CreditCardTxnInfoResultAVSStreet & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultAVSZip = '" & str2SrcQB_CreditCardTxnInfoResultAVSZip & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultReconBatchID = '" & str2SrcQB_CreditCardTxnInfoResultReconBatchID & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultPaymentGroupingCode = " & str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode & "" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultPaymentStatus = '" & str2SrcQB_CreditCardTxnInfoResultPaymentStatus & "'" & Environment.NewLine
                        strSQL3 = "     , CreditCardTxnInfoResultTxnAuthorizationTime = '" & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime & "'" & Environment.NewLine & _
                                  "     , CreditCardTxnInfoResultTxnAuthorizationStamp = " & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp & "" & Environment.NewLine & _
                                  "     , IsAutoApply = '" & str2SrcQB_IsAutoApply & "'" & Environment.NewLine & _
                                  "     , UnusedPayment = " & str2SrcQB_UnusedPayment & "" & Environment.NewLine & _
                                  "     , UnusedCredits = " & str2SrcQB_UnusedCredits & "" & Environment.NewLine & _
                                  "     , AppliedToTxnTxnID = '" & str2SrcQB_AppliedToTxnTxnID & "'" & Environment.NewLine & _
                                  "     , AppliedToTxnPaymentAmount = " & str2SrcQB_AppliedToTxnPaymentAmount & "" & Environment.NewLine & _
                                  "     , AppliedToTxnTxnType = '" & str2SrcQB_AppliedToTxnTxnType & "'" & Environment.NewLine & _
                                  "     , AppliedToTxnTxnDate = '" & str2SrcQB_AppliedToTxnTxnDate & "'" & Environment.NewLine & _
                                  "     , AppliedToTxnRefNumber = '" & str2SrcQB_AppliedToTxnRefNumber & "'" & Environment.NewLine & _
                                  "     , AppliedToTxnBalanceRemaining = " & str2SrcQB_AppliedToTxnBalanceRemaining & "" & Environment.NewLine & _
                                  "     , AppliedToTxnAmount = " & str2SrcQB_AppliedToTxnAmount & "" & Environment.NewLine & _
                                  "     , AppliedToTxnSetCreditCreditTxnID = '" & str2SrcQB_AppliedToTxnSetCreditCreditTxnID & "'" & Environment.NewLine & _
                                  "     , AppliedToTxnSetCreditAppliedAmount = " & str2SrcQB_AppliedToTxnSetCreditAppliedAmount & "" & Environment.NewLine & _
                                  "     , AppliedToTxnDiscountAmount = " & str2SrcQB_AppliedToTxnDiscountAmount & "" & Environment.NewLine & _
                                  "     , AppliedToTxnDiscountAccountRefListID = '" & str2SrcQB_AppliedToTxnDiscountAccountRefListID & "'" & Environment.NewLine & _
                                  "     , AppliedToTxnDiscountAccountRefFullName = '" & str2SrcQB_AppliedToTxnDiscountAccountRefFullName & "'" & Environment.NewLine & _
                                  "     , FQSaveToCache = '" & str2SrcQB_FQSaveToCache & "'" & Environment.NewLine & _
                                  "     , FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & Environment.NewLine & _
                                  "WHERE " & Environment.NewLine & _
                                  "       FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & Environment.NewLine

                        'Combine the strings
                        strTableUpdate = strSQL1 & strSQL2 & strSQL3 '& strSQL4 & strSQL5 & strSQL6
                        SQLHelper.ExecuteSQL(cnMax, strTableUpdate)

                    Else
                        'record not exist  -INSERT
                        Debug.WriteLine("INSERT")

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

                        'Execute the insert
                        SQLHelper.ExecuteSQL(cnMax, strTableInsert)

                    End If
                Next iteration_row

            Else
                ShowUserMessage(strSubName, "No receive payment lines to process")

            End If

        End Using
        ShowUserMessage(strSubName, "Finished processing receive payment lines", True)

    End Sub


    Public Sub UpdateQBInvoice(ByRef strCustomerRefFullName As String)
        'FIRST RUN THROUGH COMPLETE

        Dim str2SrcQB_IsActive As String = ""
        Dim iRowCount As Integer = 0
        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "UpdateQBInvoice" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'FOR PART 2SrcQB_ - Get records from QB_Invoice
        Debug.WriteLine("List2SrcQB_QB_Invoice")
        Dim str2SrcQB_QB_InvoiceSQL, str2SrcQB_QB_InvoiceRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ClassRefListID, str2SrcQB_ClassRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TemplateRefListID, str2SrcQB_TemplateRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_BillAddressAddr1, str2SrcQB_BillAddressAddr2, str2SrcQB_BillAddressAddr3, str2SrcQB_BillAddressAddr4, str2SrcQB_BillAddressCity, str2SrcQB_BillAddressState, str2SrcQB_BillAddressPostalCode, str2SrcQB_BillAddressCountry, str2SrcQB_ShipAddressAddr1, str2SrcQB_ShipAddressAddr2, str2SrcQB_ShipAddressAddr3, str2SrcQB_ShipAddressAddr4, str2SrcQB_ShipAddressCity, str2SrcQB_ShipAddressState, str2SrcQB_ShipAddressPostalCode, str2SrcQB_ShipAddressCountry, str2SrcQB_IsPending, str2SrcQB_IsFinanceCharge, str2SrcQB_PONumber, str2SrcQB_TermsRefListID, str2SrcQB_TermsRefFullName, str2SrcQB_DueDate, str2SrcQB_SalesRepRefListID, str2SrcQB_SalesRepRefFullName, str2SrcQB_FOB, str2SrcQB_ShipDate, str2SrcQB_ShipMethodRefListID, str2SrcQB_ShipMethodRefFullName, str2SrcQB_Subtotal, str2SrcQB_ItemSalesTaxRefListID, str2SrcQB_ItemSalesTaxRefFullName, str2SrcQB_SalesTaxPercentage, str2SrcQB_SalesTaxTotal, str2SrcQB_AppliedAmount, str2SrcQB_BalanceRemaining, str2SrcQB_Memo, str2SrcQB_IsPaid, str2SrcQB_CustomerMsgRefListID, str2SrcQB_CustomerMsgRefFullName, str2SrcQB_IsToBePrinted, str2SrcQB_CustomerSalesTaxCodeRefListID, str2SrcQB_CustomerSalesTaxCodeRefFullName, str2SrcQB_SuggestedDiscountAmount, str2SrcQB_SuggestedDiscountDate, str2SrcQB_CustomFieldOther As String
        'This routine gets the 2SrcQB_QB_Invoice from the database according to the selection in str2SrcQB_QB_InvoiceSQL.
        'It then puts those 2SrcQB_QB_Invoice in the list box

        'FOR PART 3TestID_
        Debug.WriteLine("List3TestID_QBTable")

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String

        Using rs2SrcQB_QB_Invoice As DataSet = New DataSet()

            str2SrcQB_QB_InvoiceSQL = "SELECT * FROM Invoice WHERE CustomerRefFullName = '" & strCustomerRefFullName & "' AND TimeModified > {ts '" & DateTime.Now.AddDays(-12).ToString("yyyy-MM-dd HH:mm:ss.000") & "'}"
            Debug.WriteLine(str2SrcQB_QB_InvoiceSQL)
            Using adap As New OdbcDataAdapter(str2SrcQB_QB_InvoiceSQL, cnQuickBooks)
                rs2SrcQB_QB_Invoice.Tables.Clear()
                adap.Fill(rs2SrcQB_QB_Invoice)
            End Using

            Dim curRow As Integer = 0
            Dim rowCount As Integer = rs2SrcQB_QB_Invoice.Tables(0).Rows.Count

            If rowCount > 10 Then
                ShowUserMessage(strSubName, "Too many Inv items found (more than 10)")
                Exit Sub
            End If

            If rowCount > 0 Then

                ShowUserMessage(strSubName, "Processing  " & rowCount.ToString & "  QB_Invoice  Records ", "Updating QB Invoice")

                For Each iteration_row As DataRow In rs2SrcQB_QB_Invoice.Tables(0).Rows
                    curRow += 1

                    ShowUserMessage(strSubName, "Processing Record " & curRow.ToString & " of " & rowCount)

                    'TODO - check for default values from empty string list in original code
                    'get the columns from the database
                    str2SrcQB_TxnID = NCStr(iteration_row("TxnID")).Replace("'"c, "`"c)
                    str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                    str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                    str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                    str2SrcQB_TxnNumber = NCStr(iteration_row("TxnNumber")).Replace("'"c, "`"c)
                    str2SrcQB_CustomerRefListID = NCStr(iteration_row("CustomerRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_CustomerRefFullName = NCStr(iteration_row("CustomerRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_ClassRefListID = NCStr(iteration_row("ClassRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_ClassRefFullName = NCStr(iteration_row("ClassRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_ARAccountRefListID = NCStr(iteration_row("ARAccountRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_ARAccountRefFullName = NCStr(iteration_row("ARAccountRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_TemplateRefListID = NCStr(iteration_row("TemplateRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_TemplateRefFullName = NCStr(iteration_row("TemplateRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_TxnDate = NCStr(iteration_row("TxnDate")).Replace("'"c, "`"c)
                    str2SrcQB_TxnDateMacro = NCStr(iteration_row("TxnDateMacro")).Replace("'"c, "`"c)
                    str2SrcQB_RefNumber = NCStr(iteration_row("RefNumber")).Replace("'"c, "`"c)
                    str2SrcQB_BillAddressAddr1 = NCStr(iteration_row("BillAddressAddr1")).Replace("'"c, "`"c)
                    str2SrcQB_BillAddressAddr2 = NCStr(iteration_row("BillAddressAddr2")).Replace("'"c, "`"c)
                    str2SrcQB_BillAddressAddr3 = NCStr(iteration_row("BillAddressAddr3")).Replace("'"c, "`"c)
                    str2SrcQB_BillAddressAddr4 = NCStr(iteration_row("BillAddressAddr4")).Replace("'"c, "`"c)
                    str2SrcQB_BillAddressCity = NCStr(iteration_row("BillAddressCity")).Replace("'"c, "`"c)
                    str2SrcQB_BillAddressState = NCStr(iteration_row("BillAddressState")).Replace("'"c, "`"c)
                    str2SrcQB_BillAddressPostalCode = NCStr(iteration_row("BillAddressPostalCode")).Replace("'"c, "`"c)
                    str2SrcQB_BillAddressCountry = NCStr(iteration_row("BillAddressCountry")).Replace("'"c, "`"c)
                    str2SrcQB_ShipAddressAddr1 = NCStr(iteration_row("ShipAddressAddr1")).Replace("'"c, "`"c)
                    str2SrcQB_ShipAddressAddr2 = NCStr(iteration_row("ShipAddressAddr2")).Replace("'"c, "`"c)
                    str2SrcQB_ShipAddressAddr3 = NCStr(iteration_row("ShipAddressAddr3")).Replace("'"c, "`"c)
                    str2SrcQB_ShipAddressAddr4 = NCStr(iteration_row("ShipAddressAddr4")).Replace("'"c, "`"c)
                    str2SrcQB_ShipAddressCity = NCStr(iteration_row("ShipAddressCity")).Replace("'"c, "`"c)
                    str2SrcQB_ShipAddressState = NCStr(iteration_row("ShipAddressState")).Replace("'"c, "`"c)
                    str2SrcQB_ShipAddressPostalCode = NCStr(iteration_row("ShipAddressPostalCode")).Replace("'"c, "`"c)
                    str2SrcQB_ShipAddressCountry = NCStr(iteration_row("ShipAddressCountry")).Replace("'"c, "`"c)
                    str2SrcQB_IsPending = NCStr(iteration_row("IsPending")).Replace("'"c, "`"c)
                    str2SrcQB_IsFinanceCharge = NCStr(iteration_row("IsFinanceCharge")).Replace("'"c, "`"c)
                    str2SrcQB_PONumber = NCStr(iteration_row("PONumber")).Replace("'"c, "`"c)
                    str2SrcQB_TermsRefListID = NCStr(iteration_row("TermsRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_TermsRefFullName = NCStr(iteration_row("TermsRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_DueDate = NCStr(iteration_row("DueDate")).Replace("'"c, "`"c)
                    str2SrcQB_SalesRepRefListID = NCStr(iteration_row("SalesRepRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_SalesRepRefFullName = NCStr(iteration_row("SalesRepRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_FOB = NCStr(iteration_row("FOB")).Replace("'"c, "`"c)
                    str2SrcQB_ShipDate = NCStr(iteration_row("ShipDate")).Replace("'"c, "`"c)
                    str2SrcQB_ShipMethodRefListID = NCStr(iteration_row("ShipMethodRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_ShipMethodRefFullName = NCStr(iteration_row("ShipMethodRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_Subtotal = NCDbl(iteration_row("Subtotal"))
                    str2SrcQB_ItemSalesTaxRefListID = NCStr(iteration_row("ItemSalesTaxRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_ItemSalesTaxRefFullName = NCStr(iteration_row("ItemSalesTaxRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_SalesTaxPercentage = NCStr(iteration_row("SalesTaxPercentage")).Replace("'"c, "`"c)
                    str2SrcQB_SalesTaxTotal = NCDbl(iteration_row("SalesTaxTotal"))
                    str2SrcQB_AppliedAmount = NCDbl(iteration_row("AppliedAmount"))
                    str2SrcQB_BalanceRemaining = NCDbl(iteration_row("BalanceRemaining"))
                    str2SrcQB_Memo = NCStr(iteration_row("Memo")).Replace("'"c, "`"c)
                    str2SrcQB_IsPaid = NCStr(iteration_row("IsPaid")).Replace("'"c, "`"c)
                    str2SrcQB_CustomerMsgRefListID = NCStr(iteration_row("CustomerMsgRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_CustomerMsgRefFullName = NCStr(iteration_row("CustomerMsgRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_IsToBePrinted = NCStr(iteration_row("IsToBePrinted")).Replace("'"c, "`"c)
                    str2SrcQB_CustomerSalesTaxCodeRefListID = NCStr(iteration_row("CustomerSalesTaxCodeRefListID")).Replace("'"c, "`"c)
                    str2SrcQB_CustomerSalesTaxCodeRefFullName = NCStr(iteration_row("CustomerSalesTaxCodeRefFullName")).Replace("'"c, "`"c)
                    str2SrcQB_SuggestedDiscountAmount = NCDbl(iteration_row("SuggestedDiscountAmount"))
                    str2SrcQB_SuggestedDiscountDate = NCStr(iteration_row("SuggestedDiscountDate")).Replace("'"c, "`"c)
                    str2SrcQB_CustomFieldOther = NCStr(iteration_row("CustomFieldOther")).Replace("'"c, "`"c)

                    ' TODO = CHECK TO SEE HOW THE FOLLOWING ARE COMING OUT OF THE DATABASE
                    'Change flags back to binary
                    str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")
                    str2SrcQB_IsPending = IIf(str2SrcQB_IsPending = "True", "1", "0")
                    str2SrcQB_IsFinanceCharge = IIf(str2SrcQB_IsFinanceCharge = "True", "1", "0")
                    str2SrcQB_IsPaid = IIf(str2SrcQB_IsPaid = "True", "1", "0")
                    str2SrcQB_IsToBePrinted = IIf(str2SrcQB_IsToBePrinted = "True", "1", "0")


                    'Put the information together into a string
                    'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                    str2SrcQB_QB_InvoiceRow = "" & _
                                              Strings.Left("Inv Upd" & "                  ", 18) & "   " & _
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

                    'Left(str2SrcQB_TxnID + "                  ", 18) & "   " & _
                    '
                    'put the line in the listbox
                    'UPGRADE_ISSUE: (2064) ADODB.Recordset property rs2SrcQB_QB_Invoice.tables(0).Rows.IndexOf(iteration_row) was not upgraded. More Information: http://www.vbtonet.com/ewis/ewi2064.aspx
                    Debug.WriteLine(Now.ToString & "   " & CStr(rs2SrcQB_QB_Invoice.Tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_Invoice.Tables(0).Rows.Count))

                    ShowUserMessage(strSubName, str2SrcQB_QB_InvoiceRow)

                    iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(TxnID) FROM QB_Invoice WHERE TxnID = '" & str2SrcQB_TxnID & "'")
                    If iRowCount > 1 Then Stop 'Should only be one
                    If iRowCount = 1 Then 'record exists  -UPDATE

                        'DO UPDATE WORK:
                        Debug.WriteLine("UPDATE")

                        strSQL1 = "UPDATE  " & Environment.NewLine & _
                                  "       QB_Invoice " & Environment.NewLine & _
                                  "SET " & Environment.NewLine & _
                                  "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine & _
                                  "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                  "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                  "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                  "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & Environment.NewLine & _
                                  "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & Environment.NewLine & _
                                  "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & Environment.NewLine & _
                                  "     , ClassRefListID = '" & str2SrcQB_ClassRefListID & "'" & Environment.NewLine & _
                                  "     , ClassRefFullName = '" & str2SrcQB_ClassRefFullName & "'" & Environment.NewLine & _
                                  "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & Environment.NewLine & _
                                  "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & Environment.NewLine & _
                                  "     , TemplateRefListID = '" & str2SrcQB_TemplateRefListID & "'" & Environment.NewLine & _
                                  "     , TemplateRefFullName = '" & str2SrcQB_TemplateRefFullName & "'" & Environment.NewLine & _
                                  "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & Environment.NewLine & _
                                  "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & Environment.NewLine & _
                                  "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & Environment.NewLine & _
                                  "     , BillAddressAddr1 = '" & str2SrcQB_BillAddressAddr1 & "'" & Environment.NewLine & _
                                  "     , BillAddressAddr2 = '" & str2SrcQB_BillAddressAddr2 & "'" & Environment.NewLine
                        strSQL2 = "     , BillAddressAddr3 = '" & str2SrcQB_BillAddressAddr3 & "'" & Environment.NewLine & _
                                  "     , BillAddressAddr4 = '" & str2SrcQB_BillAddressAddr4 & "'" & Environment.NewLine & _
                                  "     , BillAddressCity = '" & str2SrcQB_BillAddressCity & "'" & Environment.NewLine & _
                                  "     , BillAddressState = '" & str2SrcQB_BillAddressState & "'" & Environment.NewLine & _
                                  "     , BillAddressPostalCode = '" & str2SrcQB_BillAddressPostalCode & "'" & Environment.NewLine & _
                                  "     , BillAddressCountry = '" & str2SrcQB_BillAddressCountry & "'" & Environment.NewLine & _
                                  "     , ShipAddressAddr1 = '" & str2SrcQB_ShipAddressAddr1 & "'" & Environment.NewLine & _
                                  "     , ShipAddressAddr2 = '" & str2SrcQB_ShipAddressAddr2 & "'" & Environment.NewLine & _
                                  "     , ShipAddressAddr3 = '" & str2SrcQB_ShipAddressAddr3 & "'" & Environment.NewLine & _
                                  "     , ShipAddressAddr4 = '" & str2SrcQB_ShipAddressAddr4 & "'" & Environment.NewLine & _
                                  "     , ShipAddressCity = '" & str2SrcQB_ShipAddressCity & "'" & Environment.NewLine & _
                                  "     , ShipAddressState = '" & str2SrcQB_ShipAddressState & "'" & Environment.NewLine & _
                                  "     , ShipAddressPostalCode = '" & str2SrcQB_ShipAddressPostalCode & "'" & Environment.NewLine & _
                                  "     , ShipAddressCountry = '" & str2SrcQB_ShipAddressCountry & "'" & Environment.NewLine & _
                                  "     , IsPending = '" & str2SrcQB_IsPending & "'" & Environment.NewLine & _
                                  "     , IsFinanceCharge = '" & str2SrcQB_IsFinanceCharge & "'" & Environment.NewLine & _
                                  "     , PONumber = '" & str2SrcQB_PONumber & "'" & Environment.NewLine & _
                                  "     , TermsRefListID = '" & str2SrcQB_TermsRefListID & "'" & Environment.NewLine & _
                                  "     , TermsRefFullName = '" & str2SrcQB_TermsRefFullName & "'" & Environment.NewLine & _
                                  "     , DueDate = '" & str2SrcQB_DueDate & "'" & Environment.NewLine & _
                                  "     , SalesRepRefListID = '" & str2SrcQB_SalesRepRefListID & "'" & Environment.NewLine & _
                                  "     , SalesRepRefFullName = '" & str2SrcQB_SalesRepRefFullName & "'" & Environment.NewLine
                        strSQL3 = "     , FOB = '" & str2SrcQB_FOB & "'" & Environment.NewLine & _
                                  "     , ShipDate = '" & str2SrcQB_ShipDate & "'" & Environment.NewLine & _
                                  "     , ShipMethodRefListID = '" & str2SrcQB_ShipMethodRefListID & "'" & Environment.NewLine & _
                                  "     , ShipMethodRefFullName = '" & str2SrcQB_ShipMethodRefFullName & "'" & Environment.NewLine & _
                                  "     , Subtotal = " & str2SrcQB_Subtotal & "" & Environment.NewLine & _
                                  "     , ItemSalesTaxRefListID = '" & str2SrcQB_ItemSalesTaxRefListID & "'" & Environment.NewLine & _
                                  "     , ItemSalesTaxRefFullName = '" & str2SrcQB_ItemSalesTaxRefFullName & "'" & Environment.NewLine & _
                                  "     , SalesTaxPercentage = " & str2SrcQB_SalesTaxPercentage & "" & Environment.NewLine & _
                                  "     , SalesTaxTotal = " & str2SrcQB_SalesTaxTotal & "" & Environment.NewLine & _
                                  "     , AppliedAmount = " & str2SrcQB_AppliedAmount & "" & Environment.NewLine & _
                                  "     , BalanceRemaining = " & str2SrcQB_BalanceRemaining & "" & Environment.NewLine & _
                                  "     , Memo = '" & str2SrcQB_Memo & "'" & Environment.NewLine & _
                                  "     , IsPaid = '" & str2SrcQB_IsPaid & "'" & Environment.NewLine & _
                                  "     , CustomerMsgRefListID = '" & str2SrcQB_CustomerMsgRefListID & "'" & Environment.NewLine & _
                                  "     , CustomerMsgRefFullName = '" & str2SrcQB_CustomerMsgRefFullName & "'" & Environment.NewLine & _
                                  "     , IsToBePrinted = '" & str2SrcQB_IsToBePrinted & "'" & Environment.NewLine & _
                                  "     , CustomerSalesTaxCodeRefListID = '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'" & Environment.NewLine & _
                                  "     , CustomerSalesTaxCodeRefFullName = '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'" & Environment.NewLine & _
                                  "     , SuggestedDiscountAmount = " & str2SrcQB_SuggestedDiscountAmount & "" & Environment.NewLine & _
                                  "     , SuggestedDiscountDate = '" & str2SrcQB_SuggestedDiscountDate & "'" & Environment.NewLine & _
                                  "     , CustomFieldOther = '" & str2SrcQB_CustomFieldOther & "'" & Environment.NewLine & _
                                  "WHERE " & Environment.NewLine & _
                                  "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine

                        'Combine the strings
                        strTableUpdate = strSQL1 & strSQL2 & strSQL3 '& strSQL4 & strSQL5 & strSQL6

                        'Execute the insert
                        SQLHelper.ExecuteSQL(cnMax, strTableUpdate)

                    Else
                        'record not exist  -INSERT
                        'DO INSERT WORK:
                        Debug.WriteLine("INSERT")

                        strSQL1 = "INSERT INTO QB_Invoice " & Environment.NewLine & _
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
                                  "   , IsFinanceCharge " & Environment.NewLine & _
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
                                  "   , AppliedAmount " & Environment.NewLine & _
                                  "   , BalanceRemaining " & Environment.NewLine & _
                                  "   , Memo " & Environment.NewLine & _
                                  "   , IsPaid " & Environment.NewLine & _
                                  "   , CustomerMsgRefListID " & Environment.NewLine & _
                                  "   , CustomerMsgRefFullName " & Environment.NewLine & _
                                  "   , IsToBePrinted " & Environment.NewLine & _
                                  "   , CustomerSalesTaxCodeRefListID " & Environment.NewLine & _
                                  "   , CustomerSalesTaxCodeRefFullName " & Environment.NewLine & _
                                  "   , SuggestedDiscountAmount " & Environment.NewLine & _
                                  "   , SuggestedDiscountDate " & Environment.NewLine & _
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
                                  "   , '" & str2SrcQB_IsFinanceCharge & "'  --IsFinanceCharge" & Environment.NewLine & _
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
                                  "   , " & str2SrcQB_AppliedAmount & "  --AppliedAmount" & Environment.NewLine & _
                                  "   , " & str2SrcQB_BalanceRemaining & "  --BalanceRemaining" & Environment.NewLine & _
                                  "   , '" & str2SrcQB_Memo & "'  --Memo" & Environment.NewLine & _
                                  "   , '" & str2SrcQB_IsPaid & "'  --IsPaid" & Environment.NewLine & _
                                  "   , '" & str2SrcQB_CustomerMsgRefListID & "'  --CustomerMsgRefListID" & Environment.NewLine & _
                                  "   , '" & str2SrcQB_CustomerMsgRefFullName & "'  --CustomerMsgRefFullName" & Environment.NewLine & _
                                  "   , '" & str2SrcQB_IsToBePrinted & "'  --IsToBePrinted" & Environment.NewLine & _
                                  "   , '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'  --CustomerSalesTaxCodeRefListID" & Environment.NewLine & _
                                  "   , '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'  --CustomerSalesTaxCodeRefFullName" & Environment.NewLine & _
                                  "   , " & str2SrcQB_SuggestedDiscountAmount & "  --SuggestedDiscountAmount" & Environment.NewLine & _
                                  "   , '" & str2SrcQB_SuggestedDiscountDate & "'  --SuggestedDiscountDate" & Environment.NewLine & _
                                  "   , '" & str2SrcQB_CustomFieldOther & "' ) --CustomFieldOther" & Environment.NewLine

                        'Combine the strings
                        strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
                        'Debug.Print strTableInsert

                        'Execute the insert
                        SQLHelper.ExecuteSQL(cnMax, strTableInsert)
                    End If

                Next iteration_row

            End If
        End Using

    End Sub


    Public Sub UpdateQBCustomerBalance(ByRef strListID As String)
        'FIRST RUN_THROUGH COMPLETE

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "CustomerBalance_List" '"SUBNAME"

        If Not gCustomerBalanceUpdateList.Contains(strListID) Then
            gCustomerBalanceUpdateList.Add(strListID)
        End If



    End Sub

    Public Sub UpdateQBCustomerBalanceList(ByRef strListID As ArrayList)
        'FIRST RUN_THROUGH COMPLETE

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "UpdateQBCustomerBalance" '"SUBNAME"

        ShowUserMessage(strSubName, "Processing QB_Customer Balances", "Processing QB Customer Balances", True)

        Debug.WriteLine("List2SrcQB_QB_Customer")
        Dim str2SrcQB_QB_CustomerSQL, str2SrcQB_QB_CustomerRow As String
        Dim str2SrcQB_Balance, str2SrcQB_TotalBalance, str2SrcQB_OpenBalance, str2SrcQB_OpenBalanceDate As String
        Dim strTableUpdate As String

        Application.DoEvents()

        For Each sID As String In strListID

            Using rs2SrcQB_QB_Customer As New DataSet()

                str2SrcQB_QB_CustomerSQL = "SELECT Balance,TotalBalance,OpenBalance,OpenBalanceDate FROM Customer WHERE ListID = '" & sID & "' " 'strListID"
                Debug.WriteLine(str2SrcQB_QB_CustomerSQL)
                Using adap As New OdbcDataAdapter(str2SrcQB_QB_CustomerSQL, cnQuickBooks)
                    rs2SrcQB_QB_Customer.Tables.Clear()
                    adap.Fill(rs2SrcQB_QB_Customer)
                End Using

                Dim rowCount As Integer = rs2SrcQB_QB_Customer.Tables(0).Rows.Count
                Dim curRow As Integer = 0

                If rowCount > 0 Then

                    ShowUserMessage(strSubName, "Processing  " & rowCount.ToString & "  QB_Customer  Records")

                    For Each iteration_row As DataRow In rs2SrcQB_QB_Customer.Tables(0).Rows
                        curRow += 1
                        ShowUserMessage(strSubName, "Processing record " & curRow.ToString & " of " & rowCount.ToString)

                        Try

                            str2SrcQB_Balance = "0"
                            str2SrcQB_TotalBalance = "0"
                            str2SrcQB_OpenBalance = "0"
                            str2SrcQB_OpenBalanceDate = ""

                            str2SrcQB_Balance = NCDbl(iteration_row("Balance"))
                            str2SrcQB_TotalBalance = NCDbl(iteration_row("TotalBalance"))
                            str2SrcQB_OpenBalance = NCDbl(iteration_row("OpenBalance"))
                            str2SrcQB_OpenBalanceDate = NCStr(iteration_row("OpenBalanceDate"))

                            'Put the information together into a string
                            str2SrcQB_QB_CustomerRow = "Cust Bal   " & _
                                                        Strings.Left(str2SrcQB_Balance & "                  ", 18) & "   " & _
                                                        Strings.Left(str2SrcQB_TotalBalance & "                  ", 16) & "   " & _
                                                        Strings.Left(str2SrcQB_OpenBalance & "                  ", 16) & "   " & _
                                                        Strings.Left(str2SrcQB_OpenBalanceDate & "                  ", 18) & "   " & _
                                                        "" & Strings.Chr(9)

                            ShowUserMessage(strSubName, str2SrcQB_QB_CustomerRow)

                            'Build the SQL string
                            strTableUpdate = "UPDATE  " & Environment.NewLine & _
                                        "       QB_Customer " & Environment.NewLine & _
                                        "SET " & Environment.NewLine & _
                                        "       Balance = '" & str2SrcQB_Balance & "'" & Environment.NewLine & _
                                        "     , TotalBalance = '" & str2SrcQB_TotalBalance & "'" & Environment.NewLine & _
                                        "     , OpenBalance = '" & str2SrcQB_OpenBalance & "'" & Environment.NewLine & _
                                        "     , OpenBalanceDate = '" & str2SrcQB_OpenBalanceDate & "'" & Environment.NewLine & _
                                        "WHERE " & Environment.NewLine & _
                                        "       ListID = '" & sID & "'" & Environment.NewLine
                            SQLHelper.ExecuteSQL(cnMax, strTableUpdate)

                        Catch ex As Exception
                            HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                            Continue For
                        End Try
                    Next iteration_row

                End If
            End Using

        Next

        ShowUserMessage(strSubName, "Finished Processing QB_Customer Balances", , True)

    End Sub


    Public Sub RefreshQB_ReceivePayment()
        'FIRST RUN_THROUGH COMPLETE


        Dim str2SrcQB_FQSaveToCache As String = ""
        Dim iRowCount As Integer = 0
        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQB_ReceivePayment" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub


        'FOR PART 2SrcQB_ - Get records from QB_ReceivePayment
        Debug.WriteLine("List2SrcQB_QB_ReceivePayment")
        Dim str2SrcQB_QB_ReceivePaymentSQL, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_TotalAmount, str2SrcQB_PaymentMethodRefListID, str2SrcQB_PaymentMethodRefFullName, str2SrcQB_Memo, str2SrcQB_DepositToAccountRefListID, str2SrcQB_DepositToAccountRefFullName, str2SrcQB_IsAutoApply, str2SrcQB_UnusedPayment, str2SrcQB_UnusedCredits As String
        'This routine gets the 2SrcQB_QB_ReceivePayment from the database according to the selection in str2SrcQB_QB_ReceivePaymentSQL.
        'It then puts those 2SrcQB_QB_ReceivePayment in the list box

        'FOR PART 3TestID_
        Debug.WriteLine("List3TestID_QBTable")
        'This routine gets the 3TestID_QBTable from the database according to the selection in str3TestID_QBTableSQL.
        'It then puts those 3TestID_QBTable in the list box

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL4, strSQL5, strTableInsert, strTableUpdate As String


        'Show what's processing
        ShowUserMessage(strSubName, "Processing  QB_ReceivePayment  Records ", "Processing QB_ReceivePayment Records", True)


        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        Using rs2SrcQB_QB_ReceivePayment As DataSet = New DataSet()


            str2SrcQB_QB_ReceivePaymentSQL = "SELECT * FROM ReceivePayment WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_ReceivePayment & "'}" ' ORDER BY TimeModified"
            Debug.WriteLine(str2SrcQB_QB_ReceivePaymentSQL)
            Using adap As New OdbcDataAdapter(str2SrcQB_QB_ReceivePaymentSQL, cnQuickBooks)
                rs2SrcQB_QB_ReceivePayment.Tables.Clear()
                adap.Fill(rs2SrcQB_QB_ReceivePayment)
            End Using

            Dim curRow As Integer = 0
            Dim rowCount As Integer = rs2SrcQB_QB_ReceivePayment.Tables(0).Rows.Count

            If rowCount > 0 Then

                'Show what's processing in the listbox
                ShowUserMessage(strSubName, "Processing  " & rowCount.ToString & "  QB_ReceivePayment  Records ")

                For Each iteration_row As DataRow In rs2SrcQB_QB_ReceivePayment.Tables(0).Rows
                    curRow += 1

                    Try



                        ShowUserMessage(strSubName, "Processing " & curRow.ToString & " of " & rowCount.ToString & " QB_ReceivePayment Records")

                        'get the columns from the database
                        str2SrcQB_TxnID = NCStr(iteration_row("TxnID")).Replace("'"c, "`"c)
                        str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                        str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                        str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                        str2SrcQB_TxnNumber = NCStr(iteration_row("TxnNumber")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefListID = NCStr(iteration_row("CustomerRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefFullName = NCStr(iteration_row("CustomerRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefListID = NCStr(iteration_row("ARAccountRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefFullName = NCStr(iteration_row("ARAccountRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDate = NCStr(iteration_row("TxnDate")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDateMacro = NCStr(iteration_row("TxnDateMacro")).Replace("'"c, "`"c)
                        str2SrcQB_RefNumber = NCStr(iteration_row("RefNumber")).Replace("'"c, "`"c)
                        str2SrcQB_TotalAmount = NCStr(iteration_row("TotalAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_PaymentMethodRefListID = NCStr(iteration_row("PaymentMethodRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_PaymentMethodRefFullName = NCStr(iteration_row("PaymentMethodRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_Memo = NCStr(iteration_row("Memo")).Replace("'"c, "`"c)
                        str2SrcQB_DepositToAccountRefListID = NCStr(iteration_row("DepositToAccountRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_DepositToAccountRefFullName = NCStr(iteration_row("DepositToAccountRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_IsAutoApply = NCStr(iteration_row("IsAutoApply")).Replace("'"c, "`"c)
                        str2SrcQB_UnusedPayment = NCDbl(iteration_row("UnusedPayment"))
                        str2SrcQB_UnusedCredits = NCDbl(iteration_row("UnusedCredits"))


                        'Change flags back to binary
                        str2SrcQB_IsAutoApply = IIf(str2SrcQB_IsAutoApply = "True", "1", "0")
                        str2SrcQB_FQSaveToCache = IIf(str2SrcQB_FQSaveToCache = "True", "1", "0")

                        Debug.WriteLine(Now.ToString & "   " & CStr(rs2SrcQB_QB_ReceivePayment.Tables(0).Rows.IndexOf(iteration_row)) & " of " & CStr(rs2SrcQB_QB_ReceivePayment.Tables(0).Rows.Count))

                        'with each record....
                        UpdateQBCustomerBalance(str2SrcQB_CustomerRefListID)
                        'UpdateQBReceivePaymentLine(str2SrcQB_CustomerRefFullName)
                        'UpdateQBInvoice(str2SrcQB_CustomerRefFullName)

                        'Check to see if ListID or TxnID is in QBTable            'Yes then UPDATE record            'No then INSERT record
                        iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT TxnID FROM QB_ReceivePayment WHERE TxnID = '" & str2SrcQB_TxnID & "'")
                        If iRowCount = 1 Then 'record exists  -UPDATE
                            'DO UPDATE WORK:
                            Debug.WriteLine("UPDATE")

                            'Build the SQL string
                            strSQL1 = "UPDATE  " & Environment.NewLine & _
                                      "       QB_ReceivePayment " & Environment.NewLine & _
                                      "SET " & Environment.NewLine & _
                                      "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine & _
                                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                      "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & Environment.NewLine & _
                                      "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & Environment.NewLine & _
                                      "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & Environment.NewLine & _
                                      "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & Environment.NewLine & _
                                      "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & Environment.NewLine & _
                                      "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & Environment.NewLine & _
                                      "     , TotalAmount = " & str2SrcQB_TotalAmount & "" & Environment.NewLine & _
                                      "     , PaymentMethodRefListID = '" & str2SrcQB_PaymentMethodRefListID & "'" & Environment.NewLine & _
                                      "     , PaymentMethodRefFullName = '" & str2SrcQB_PaymentMethodRefFullName & "'" & Environment.NewLine & _
                                      "     , Memo = '" & str2SrcQB_Memo & "'" & Environment.NewLine & _
                                      "     , DepositToAccountRefListID = '" & str2SrcQB_DepositToAccountRefListID & "'" & Environment.NewLine
                            strSQL2 = "     , DepositToAccountRefFullName = '" & str2SrcQB_DepositToAccountRefFullName & "'" & Environment.NewLine & _
                                       "     , IsAutoApply = '" & str2SrcQB_IsAutoApply & "'" & Environment.NewLine & _
                                      "     , UnusedPayment = " & str2SrcQB_UnusedPayment & "" & Environment.NewLine & _
                                      "     , UnusedCredits = " & str2SrcQB_UnusedCredits & "" & Environment.NewLine & _
                                      "WHERE " & Environment.NewLine & _
                                      "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine

                            'Combine the strings
                            strTableUpdate = strSQL1 & strSQL2
                            'Debug.Print strTableUpdate

                            'Execute the insert
                            SQLHelper.ExecuteSQL(cnMax, strTableUpdate)

                        Else
                            'record not exist  -INSERT
                            'DO INSERT WORK:
                            Debug.WriteLine("INSERT")

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
                                      "   , '" & str2SrcQB_IsAutoApply & "'" & Environment.NewLine & _
                                      "   , " & str2SrcQB_UnusedPayment & Environment.NewLine & _
                                      "   , " & str2SrcQB_UnusedCredits & " ) "
                            'Combine the strings
                            strTableInsert = strSQL1 & strSQL2 & strSQL4 & strSQL5
                            'Debug.Print strTableInsert

                            'Execute the insert
                            SQLHelper.ExecuteSQL(cnMax, strTableInsert)

                        End If

                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For
                    End Try
                Next iteration_row

            End If

        End Using
        ShowUserMessage(strSubName, "Finished processing QB_ReceivePayment Records ", , True)



    End Sub


    Public Sub RefreshQB_ReceivePaymentLine()
        'FIRST RUN_THROUGH COMPLETE

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQB_ReceivePaymentLine" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'FOR PART 2SrcQB_ - Get records from QB_ReceivePaymentLine
        Debug.WriteLine("List2SrcQB_QB_ReceivePaymentLine")
        Dim str2SrcQB_QB_ReceivePaymentLineSQL, str2SrcQB_QB_ReceivePaymentLineRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_TotalAmount, str2SrcQB_PaymentMethodRefListID, str2SrcQB_PaymentMethodRefFullName, str2SrcQB_Memo, str2SrcQB_DepositToAccountRefListID, str2SrcQB_DepositToAccountRefFullName, str2SrcQB_CreditCardTxnInfoInputCreditCardNumber, str2SrcQB_CreditCardTxnInfoInputExpirationMonth, str2SrcQB_CreditCardTxnInfoInputExpirationYear, str2SrcQB_CreditCardTxnInfoInputNameOnCard, str2SrcQB_CreditCardTxnInfoInputCreditCardAddress, str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode, str2SrcQB_CreditCardTxnInfoInputCommercialCardCode, str2SrcQB_CreditCardTxnInfoResultResultCode, str2SrcQB_CreditCardTxnInfoResultResultMessage, str2SrcQB_CreditCardTxnInfoResultCreditCardTransID, str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber, str2SrcQB_CreditCardTxnInfoResultAuthorizationCode, str2SrcQB_CreditCardTxnInfoResultAVSStreet, str2SrcQB_CreditCardTxnInfoResultAVSZip, str2SrcQB_CreditCardTxnInfoResultReconBatchID, str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode, str2SrcQB_CreditCardTxnInfoResultPaymentStatus, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime, str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp, str2SrcQB_IsAutoApply, str2SrcQB_UnusedPayment, str2SrcQB_UnusedCredits, str2SrcQB_AppliedToTxnTxnID, str2SrcQB_AppliedToTxnPaymentAmount, str2SrcQB_AppliedToTxnTxnType, str2SrcQB_AppliedToTxnTxnDate, str2SrcQB_AppliedToTxnRefNumber, str2SrcQB_AppliedToTxnBalanceRemaining, str2SrcQB_AppliedToTxnAmount, str2SrcQB_AppliedToTxnSetCreditCreditTxnID, str2SrcQB_AppliedToTxnSetCreditAppliedAmount, str2SrcQB_AppliedToTxnDiscountAmount, str2SrcQB_AppliedToTxnDiscountAccountRefListID, str2SrcQB_AppliedToTxnDiscountAccountRefFullName, str2SrcQB_FQSaveToCache, str2SrcQB_FQPrimaryKey As String
        'This routine gets the 2SrcQB_QB_ReceivePaymentLine from the database according to the selection in str2SrcQB_QB_ReceivePaymentLineSQL.
        'It then puts those 2SrcQB_QB_ReceivePaymentLine in the list box

        'FOR PART 3TestID_
        Debug.WriteLine("List3TestID_QBTable")

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String

        'Show what's processing
        ShowUserMessage(strSubName, "RefreshQB: Processing QB_ReceivePaymentLine Records ", "RefreshQB: Processing  QB_ReceivePaymentLine  Records ", True)


        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        str2SrcQB_QB_ReceivePaymentLineSQL = "SELECT * FROM ReceivePaymentLine WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_ReceivePaymentLine & "'}" ' ORDER BY TimeModified"
        Using rs2SrcQB_QB_ReceivePaymentLine As DataSet = New DataSet()


            Debug.WriteLine(str2SrcQB_QB_ReceivePaymentLineSQL)

            Using adap As New OdbcDataAdapter(str2SrcQB_QB_ReceivePaymentLineSQL, cnQuickBooks)
                rs2SrcQB_QB_ReceivePaymentLine.Tables.Clear()
                adap.Fill(rs2SrcQB_QB_ReceivePaymentLine)
            End Using

            Dim curRow As Integer = 0
            Dim rowCount As Integer = rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows.Count

            If rowCount > 0 Then

                'Show what's processing in the listbox
                ShowUserMessage(strSubName, "Processing  " & rowCount.ToString & " QB_ReceivePaymentLine Records")

                For Each iteration_row As DataRow In rs2SrcQB_QB_ReceivePaymentLine.Tables(0).Rows
                    curRow += 1

                    ShowUserMessage(strSubName, "Processing Record " & curRow.ToString & " of " & rowCount.ToString)

                    Try


                        'get the columns from the database
                        str2SrcQB_TxnID = NCStr(iteration_row("TxnID")).Replace("'"c, "`"c)
                        str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                        str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                        str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                        str2SrcQB_TxnNumber = NCStr(iteration_row("TxnNumber"), "0").Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefListID = NCStr(iteration_row("CustomerRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefFullName = NCStr(iteration_row("CustomerRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefListID = NCStr(iteration_row("ARAccountRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefFullName = NCStr(iteration_row("ARAccountRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDate = NCStr(iteration_row("TxnDate")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDateMacro = NCStr(iteration_row("TxnDateMacro")).Replace("'"c, "`"c)
                        str2SrcQB_RefNumber = NCStr(iteration_row("RefNumber")).Replace("'"c, "`"c)
                        str2SrcQB_TotalAmount = NCStr(iteration_row("TotalAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_PaymentMethodRefListID = NCStr(iteration_row("PaymentMethodRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_PaymentMethodRefFullName = NCStr(iteration_row("PaymentMethodRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_Memo = NCStr(iteration_row("Memo")).Replace("'"c, "`"c)
                        str2SrcQB_DepositToAccountRefListID = NCStr(iteration_row("DepositToAccountRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_DepositToAccountRefFullName = NCStr(iteration_row("DepositToAccountRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoInputCreditCardNumber = NCStr(iteration_row("CreditCardTxnInfoInputCreditCardNumber")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoInputExpirationMonth = NCStr(iteration_row("CreditCardTxnInfoInputExpirationMonth"), "0").Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoInputExpirationYear = NCStr(iteration_row("CreditCardTxnInfoInputExpirationYear"), "0").Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoInputNameOnCard = NCStr(iteration_row("CreditCardTxnInfoInputNameOnCard")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoInputCreditCardAddress = NCStr(iteration_row("CreditCardTxnInfoInputCreditCardAddress")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode = NCStr(iteration_row("CreditCardTxnInfoInputCreditCardPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoInputCommercialCardCode = NCStr(iteration_row("CreditCardTxnInfoInputCommercialCardCode")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultResultCode = NCStr(iteration_row("CreditCardTxnInfoResultResultCode"), "0").Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultResultMessage = NCStr(iteration_row("CreditCardTxnInfoResultResultMessage")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultCreditCardTransID = NCStr(iteration_row("CreditCardTxnInfoResultCreditCardTransID")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber = NCStr(iteration_row("CreditCardTxnInfoResultMerchantAccountNumber")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultAuthorizationCode = NCStr(iteration_row("CreditCardTxnInfoResultAuthorizationCode")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultAVSStreet = NCStr(iteration_row("CreditCardTxnInfoResultAVSStreet")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultAVSZip = NCStr(iteration_row("CreditCardTxnInfoResultAVSZip")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultReconBatchID = NCStr(iteration_row("CreditCardTxnInfoResultReconBatchID")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode = NCStr(iteration_row("CreditCardTxnInfoResultPaymentGroupingCode"), "0").Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultPaymentStatus = NCStr(iteration_row("CreditCardTxnInfoResultPaymentStatus")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime = NCStr(iteration_row("CreditCardTxnInfoResultTxnAuthorizationTime")).Replace("'"c, "`"c)
                        str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp = NCStr(iteration_row("CreditCardTxnInfoResultTxnAuthorizationStamp"), "0").Replace("'"c, "`"c)
                        str2SrcQB_IsAutoApply = NCStr(iteration_row("IsAutoApply")).Replace("'"c, "`"c)
                        str2SrcQB_UnusedPayment = NCStr(iteration_row("UnusedPayment"), "0").Replace("'"c, "`"c)
                        str2SrcQB_UnusedCredits = NCStr(iteration_row("UnusedCredits"), "0").Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnTxnID = NCStr(iteration_row("AppliedToTxnTxnID")).Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnPaymentAmount = NCStr(iteration_row("AppliedToTxnPaymentAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnTxnType = NCStr(iteration_row("AppliedToTxnTxnType")).Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnTxnDate = NCStr(iteration_row("AppliedToTxnTxnDate")).Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnRefNumber = NCStr(iteration_row("AppliedToTxnRefNumber")).Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnBalanceRemaining = NCStr(iteration_row("AppliedToTxnBalanceRemaining"), "0").Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnAmount = NCStr(iteration_row("AppliedToTxnAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnSetCreditCreditTxnID = NCStr(iteration_row("AppliedToTxnSetCreditCreditTxnID")).Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnSetCreditAppliedAmount = NCStr(iteration_row("AppliedToTxnSetCreditAppliedAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnDiscountAmount = NCStr(iteration_row("AppliedToTxnDiscountAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnDiscountAccountRefListID = NCStr(iteration_row("AppliedToTxnDiscountAccountRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_AppliedToTxnDiscountAccountRefFullName = NCStr(iteration_row("AppliedToTxnDiscountAccountRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_FQSaveToCache = NCStr(iteration_row("FQSaveToCache")).Replace("'"c, "`"c)
                        str2SrcQB_FQPrimaryKey = NCStr(iteration_row("FQPrimaryKey")).Replace("'"c, "`"c)


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
                        ShowUserMessage(strSubName, str2SrcQB_QB_ReceivePaymentLineRow)

                        ' UpdateQBInvoice(str2SrcQB_CustomerRefFullName)
                        UpdateQBCustomerBalance(str2SrcQB_CustomerRefListID)
                        'UpdateQBReceivePaymentLine(str2SrcQB_CustomerRefFullName)

                        'New recordset
                        Dim iRowCount As Integer = 0
                        iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(FQPrimaryKey) FROM QB_ReceivePaymentLine WHERE FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'")
                        'If iRowCount > 1 Then Stop 'Should only be one
                        If iRowCount = 1 Then 'record exists  -UPDATE
                            'DO UPDATE WORK:
                            Debug.WriteLine("UPDATE")

                            'Build the SQL string
                            strSQL1 = "UPDATE  " & Environment.NewLine & _
                                      "       QB_ReceivePaymentLine " & Environment.NewLine & _
                                      "SET " & Environment.NewLine & _
                                      "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine & _
                                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                      "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & Environment.NewLine & _
                                      "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & Environment.NewLine & _
                                      "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & Environment.NewLine & _
                                      "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & Environment.NewLine & _
                                      "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & Environment.NewLine & _
                                      "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & Environment.NewLine & _
                                      "     , TotalAmount = " & str2SrcQB_TotalAmount & "" & Environment.NewLine & _
                                      "     , PaymentMethodRefListID = '" & str2SrcQB_PaymentMethodRefListID & "'" & Environment.NewLine & _
                                      "     , PaymentMethodRefFullName = '" & str2SrcQB_PaymentMethodRefFullName & "'" & Environment.NewLine & _
                                      "     , Memo = '" & str2SrcQB_Memo & "'" & Environment.NewLine & _
                                      "     , DepositToAccountRefListID = '" & str2SrcQB_DepositToAccountRefListID & "'" & Environment.NewLine
                            strSQL2 = "     , DepositToAccountRefFullName = '" & str2SrcQB_DepositToAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoInputCreditCardNumber = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardNumber & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoInputExpirationMonth = " & str2SrcQB_CreditCardTxnInfoInputExpirationMonth & "" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoInputExpirationYear = " & str2SrcQB_CreditCardTxnInfoInputExpirationYear & "" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoInputNameOnCard = '" & str2SrcQB_CreditCardTxnInfoInputNameOnCard & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoInputCreditCardAddress = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardAddress & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoInputCreditCardPostalCode = '" & str2SrcQB_CreditCardTxnInfoInputCreditCardPostalCode & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoInputCommercialCardCode = '" & str2SrcQB_CreditCardTxnInfoInputCommercialCardCode & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultResultCode = " & str2SrcQB_CreditCardTxnInfoResultResultCode & "" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultResultMessage = '" & str2SrcQB_CreditCardTxnInfoResultResultMessage & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultCreditCardTransID = '" & str2SrcQB_CreditCardTxnInfoResultCreditCardTransID & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultMerchantAccountNumber = '" & str2SrcQB_CreditCardTxnInfoResultMerchantAccountNumber & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultAuthorizationCode = '" & str2SrcQB_CreditCardTxnInfoResultAuthorizationCode & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultAVSStreet = '" & str2SrcQB_CreditCardTxnInfoResultAVSStreet & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultAVSZip = '" & str2SrcQB_CreditCardTxnInfoResultAVSZip & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultReconBatchID = '" & str2SrcQB_CreditCardTxnInfoResultReconBatchID & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultPaymentGroupingCode = " & str2SrcQB_CreditCardTxnInfoResultPaymentGroupingCode & "" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultPaymentStatus = '" & str2SrcQB_CreditCardTxnInfoResultPaymentStatus & "'" & Environment.NewLine
                            strSQL3 = "     , CreditCardTxnInfoResultTxnAuthorizationTime = '" & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationTime & "'" & Environment.NewLine & _
                                      "     , CreditCardTxnInfoResultTxnAuthorizationStamp = " & str2SrcQB_CreditCardTxnInfoResultTxnAuthorizationStamp & "" & Environment.NewLine & _
                                      "     , IsAutoApply = '" & str2SrcQB_IsAutoApply & "'" & Environment.NewLine & _
                                      "     , UnusedPayment = " & str2SrcQB_UnusedPayment & "" & Environment.NewLine & _
                                      "     , UnusedCredits = " & str2SrcQB_UnusedCredits & "" & Environment.NewLine & _
                                      "     , AppliedToTxnTxnID = '" & str2SrcQB_AppliedToTxnTxnID & "'" & Environment.NewLine & _
                                      "     , AppliedToTxnPaymentAmount = " & str2SrcQB_AppliedToTxnPaymentAmount & "" & Environment.NewLine & _
                                      "     , AppliedToTxnTxnType = '" & str2SrcQB_AppliedToTxnTxnType & "'" & Environment.NewLine & _
                                      "     , AppliedToTxnTxnDate = '" & str2SrcQB_AppliedToTxnTxnDate & "'" & Environment.NewLine & _
                                      "     , AppliedToTxnRefNumber = '" & str2SrcQB_AppliedToTxnRefNumber & "'" & Environment.NewLine & _
                                      "     , AppliedToTxnBalanceRemaining = " & str2SrcQB_AppliedToTxnBalanceRemaining & "" & Environment.NewLine & _
                                      "     , AppliedToTxnAmount = " & str2SrcQB_AppliedToTxnAmount & "" & Environment.NewLine & _
                                      "     , AppliedToTxnSetCreditCreditTxnID = '" & str2SrcQB_AppliedToTxnSetCreditCreditTxnID & "'" & Environment.NewLine & _
                                      "     , AppliedToTxnSetCreditAppliedAmount = " & str2SrcQB_AppliedToTxnSetCreditAppliedAmount & "" & Environment.NewLine & _
                                      "     , AppliedToTxnDiscountAmount = " & str2SrcQB_AppliedToTxnDiscountAmount & "" & Environment.NewLine & _
                                      "     , AppliedToTxnDiscountAccountRefListID = '" & str2SrcQB_AppliedToTxnDiscountAccountRefListID & "'" & Environment.NewLine & _
                                      "     , AppliedToTxnDiscountAccountRefFullName = '" & str2SrcQB_AppliedToTxnDiscountAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , FQSaveToCache = '" & str2SrcQB_FQSaveToCache & "'" & Environment.NewLine & _
                                      "     , FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & Environment.NewLine & _
                                      "WHERE " & Environment.NewLine & _
                                      "       FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & Environment.NewLine

                            'Combine the strings
                            strTableUpdate = strSQL1 & strSQL2 & strSQL3
                            'Debug.Print strTableUpdate

                            'Execute the insert
                            SQLHelper.ExecuteSQL(cnMax, strTableUpdate)



                        Else
                            'record not exist  -INSERT
                            'DO INSERT WORK:
                            Debug.WriteLine("INSERT")

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
                            strTableInsert = strSQL1 & strSQL2 & strSQL3 & NCStr(strSQL4) & NCStr(strSQL5) & NCStr(strSQL6)

                            'Debug.Print strTableInsert

                            'Execute the insert
                            SQLHelper.ExecuteSQL(cnMax, strTableInsert)

                        End If

                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For
                    End Try
                Next iteration_row

                Try
                    SQLHelper.ExecuteSP(cnMax, "sp_TEMP_FixQODBCPaidInvoices")
                Catch ex As Exception
                    HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                    Exit Try
                End Try


            End If

        End Using
        ShowUserMessage(strSubName, "RefreshQB: Finished processing QB_ReceivePaymentLine Records ", "RefreshQB: Finished processing  QB_ReceivePaymentLine  Records ", True)


    End Sub



    Public Sub RefreshQB_InvoiceLine()
        'FIRST RUN_THROUGH COMPLETE


        Dim strSQL7, strSQL8 As String

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQB_InvoiceLine" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        ''This routine gets the 1MaxOfCopy_QBTable from the database according to the selection in str1MaxOfCopy_QBTableSQL.
        ''It then puts those 1MaxOfCopy_QBTable in the list box

        'FOR PART 2SrcQB_ - Get records from QB_InvoiceLine
        Debug.WriteLine("List2SrcQB_QB_InvoiceLine")
        Dim str2SrcQB_QB_InvoiceLineSQL, str2SrcQB_QB_InvoiceLineRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ClassRefListID, str2SrcQB_ClassRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TemplateRefListID, str2SrcQB_TemplateRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_BillAddressAddr1, str2SrcQB_BillAddressAddr2, str2SrcQB_BillAddressAddr3, str2SrcQB_BillAddressAddr4, str2SrcQB_BillAddressCity, str2SrcQB_BillAddressState, str2SrcQB_BillAddressPostalCode, str2SrcQB_BillAddressCountry, str2SrcQB_ShipAddressAddr1, str2SrcQB_ShipAddressAddr2, str2SrcQB_ShipAddressAddr3, str2SrcQB_ShipAddressAddr4, str2SrcQB_ShipAddressCity, str2SrcQB_ShipAddressState, str2SrcQB_ShipAddressPostalCode, str2SrcQB_ShipAddressCountry, str2SrcQB_IsPending, str2SrcQB_IsFinanceCharge, str2SrcQB_PONumber, str2SrcQB_TermsRefListID, str2SrcQB_TermsRefFullName, str2SrcQB_DueDate, str2SrcQB_SalesRepRefListID, str2SrcQB_SalesRepRefFullName, str2SrcQB_FOB, str2SrcQB_ShipDate, str2SrcQB_ShipMethodRefListID, str2SrcQB_ShipMethodRefFullName, str2SrcQB_Subtotal, str2SrcQB_ItemSalesTaxRefListID, str2SrcQB_ItemSalesTaxRefFullName, str2SrcQB_SalesTaxPercentage, str2SrcQB_SalesTaxTotal, str2SrcQB_AppliedAmount, str2SrcQB_BalanceRemaining, str2SrcQB_Memo, str2SrcQB_IsPaid, str2SrcQB_CustomerMsgRefListID, str2SrcQB_CustomerMsgRefFullName, str2SrcQB_IsToBePrinted, str2SrcQB_CustomerSalesTaxCodeRefListID, str2SrcQB_CustomerSalesTaxCodeRefFullName, str2SrcQB_SuggestedDiscountAmount, str2SrcQB_SuggestedDiscountDate, str2SrcQB_InvoiceLineType, str2SrcQB_InvoiceLineSeqNo, str2SrcQB_InvoiceLineGroupTxnLineID, str2SrcQB_InvoiceLineGroupItemGroupRefListID, str2SrcQB_InvoiceLineGroupItemGroupRefFullName, str2SrcQB_InvoiceLineGroupDesc, str2SrcQB_InvoiceLineGroupQuantity, str2SrcQB_InvoiceLineGroupIsPrintItemsInGroup, str2SrcQB_InvoiceLineGroupTotalAmount, str2SrcQB_InvoiceLineGroupSeqNo, str2SrcQB_InvoiceLineTxnLineID, str2SrcQB_InvoiceLineItemRefListID, str2SrcQB_InvoiceLineItemRefFullName, str2SrcQB_InvoiceLineDesc, str2SrcQB_InvoiceLineQuantity, str2SrcQB_InvoiceLineRate, str2SrcQB_InvoiceLineRatePercent, str2SrcQB_InvoiceLinePriceLevelRefListID, str2SrcQB_InvoiceLinePriceLevelRefFullName, str2SrcQB_InvoiceLineClassRefListID, str2SrcQB_InvoiceLineClassRefFullName, str2SrcQB_InvoiceLineAmount, str2SrcQB_InvoiceLineServiceDate, str2SrcQB_InvoiceLineSalesTaxCodeRefListID, str2SrcQB_InvoiceLineSalesTaxCodeRefFullName, str2SrcQB_InvoiceLineOverrideItemAccountRefListID, str2SrcQB_InvoiceLineOverrideItemAccountRefFullName, str2SrcQB_FQSaveToCache, str2SrcQB_FQPrimaryKey, str2SrcQB_CustomFieldInvoiceLineOther1, str2SrcQB_CustomFieldInvoiceLineOther2, str2SrcQB_CustomFieldInvoiceLineGroupOther1, str2SrcQB_CustomFieldInvoiceLineGroupOther2, str2SrcQB_CustomFieldInvoiceLineGroupLineOther1, str2SrcQB_CustomFieldInvoiceLineGroupLineOther2, str2SrcQB_CustomFieldOther As String
        'This routine gets the 2SrcQB_QB_InvoiceLine from the database according to the selection in str2SrcQB_QB_InvoiceLineSQL.
        'It then puts those 2SrcQB_QB_InvoiceLine in the list box

        'FOR PART 3TestID_
        Debug.WriteLine("List3TestID_QBTable")

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String

        'Show what's processing

        ShowUserMessage(strSubName, "RefreshQB: Processing QB_InvoiceLine Records", "RefreshQB: Processing QB_InvoiceLine Records", True)

        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        Using rs2SrcQB_QB_InvoiceLine As DataSet = New DataSet()
            str2SrcQB_QB_InvoiceLineSQL = "SELECT * FROM InvoiceLine WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_InvoiceLine & "'}" ' ORDER BY TimeModified ASC"
            Debug.WriteLine(str2SrcQB_QB_InvoiceLineSQL)
            Using adap As New OdbcDataAdapter(str2SrcQB_QB_InvoiceLineSQL, cnQuickBooks)
                rs2SrcQB_QB_InvoiceLine.Tables.Clear()
                adap.Fill(rs2SrcQB_QB_InvoiceLine)
            End Using

            Dim curRow As Integer = 0
            Dim rowCount As Integer = rs2SrcQB_QB_InvoiceLine.Tables(0).Rows.Count
            ShowUserMessage(strSubName, "Processing  " & rowCount.ToString & "  QB_InvoiceLine  Records ")

            If rowCount > 0 Then

                For Each iteration_row As DataRow In rs2SrcQB_QB_InvoiceLine.Tables(0).Rows

                    curRow += 1
                    ShowUserMessage(strSubName, "Processing " & curRow.ToString & " of " & rowCount.ToString & " QB_InvoiceLine Records")

                    Try


                        str2SrcQB_CustomFieldInvoiceLineOther1 = ""
                        str2SrcQB_CustomFieldInvoiceLineOther2 = ""
                        str2SrcQB_CustomFieldInvoiceLineGroupOther1 = ""
                        str2SrcQB_CustomFieldInvoiceLineGroupOther2 = ""
                        str2SrcQB_CustomFieldInvoiceLineGroupLineOther1 = ""
                        str2SrcQB_CustomFieldInvoiceLineGroupLineOther2 = ""
                        str2SrcQB_CustomFieldOther = ""

                        str2SrcQB_TxnID = NCStr(iteration_row("TxnID")).Replace("'"c, "`"c)
                        str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                        str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                        str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                        str2SrcQB_TxnNumber = NCStr(iteration_row("TxnNumber"), "0").Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefListID = NCStr(iteration_row("CustomerRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefFullName = NCStr(iteration_row("CustomerRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ClassRefListID = NCStr(iteration_row("ClassRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ClassRefFullName = NCStr(iteration_row("ClassRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefListID = NCStr(iteration_row("ARAccountRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefFullName = NCStr(iteration_row("ARAccountRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_TemplateRefListID = NCStr(iteration_row("TemplateRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_TemplateRefFullName = NCStr(iteration_row("TemplateRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDate = NCStr(iteration_row("TxnDate")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDateMacro = NCStr(iteration_row("TxnDateMacro")).Replace("'"c, "`"c)
                        str2SrcQB_RefNumber = NCStr(iteration_row("RefNumber")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr1 = NCStr(iteration_row("BillAddressAddr1")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr2 = NCStr(iteration_row("BillAddressAddr2")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr3 = NCStr(iteration_row("BillAddressAddr3")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr4 = NCStr(iteration_row("BillAddressAddr4")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressCity = NCStr(iteration_row("BillAddressCity")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressState = NCStr(iteration_row("BillAddressState")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressPostalCode = NCStr(iteration_row("BillAddressPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressCountry = NCStr(iteration_row("BillAddressCountry")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr1 = NCStr(iteration_row("ShipAddressAddr1")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr2 = NCStr(iteration_row("ShipAddressAddr2")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr3 = NCStr(iteration_row("ShipAddressAddr3")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr4 = NCStr(iteration_row("ShipAddressAddr4")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressCity = NCStr(iteration_row("ShipAddressCity")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressState = NCStr(iteration_row("ShipAddressState")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressPostalCode = NCStr(iteration_row("ShipAddressPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressCountry = NCStr(iteration_row("ShipAddressCountry")).Replace("'"c, "`"c)
                        str2SrcQB_IsPending = NCStr(iteration_row("IsPending")).Replace("'"c, "`"c)
                        str2SrcQB_IsFinanceCharge = NCStr(iteration_row("IsFinanceCharge")).Replace("'"c, "`"c)
                        str2SrcQB_PONumber = NCStr(iteration_row("PONumber")).Replace("'"c, "`"c)
                        str2SrcQB_TermsRefListID = NCStr(iteration_row("TermsRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_TermsRefFullName = NCStr(iteration_row("TermsRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_DueDate = NCStr(iteration_row("DueDate")).Replace("'"c, "`"c)
                        str2SrcQB_SalesRepRefListID = NCStr(iteration_row("SalesRepRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_SalesRepRefFullName = NCStr(iteration_row("SalesRepRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_FOB = NCStr(iteration_row("FOB")).Replace("'"c, "`"c)
                        str2SrcQB_ShipDate = NCStr(iteration_row("ShipDate")).Replace("'"c, "`"c)
                        str2SrcQB_ShipMethodRefListID = NCStr(iteration_row("ShipMethodRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ShipMethodRefFullName = NCStr(iteration_row("ShipMethodRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_Subtotal = NCStr(iteration_row("Subtotal"), "0").Replace("'"c, "`"c)
                        str2SrcQB_ItemSalesTaxRefListID = NCStr(iteration_row("ItemSalesTaxRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ItemSalesTaxRefFullName = NCStr(iteration_row("ItemSalesTaxRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_SalesTaxPercentage = NCStr(iteration_row("SalesTaxPercentage"), "0").Replace("'"c, "`"c)
                        str2SrcQB_SalesTaxTotal = NCStr(iteration_row("SalesTaxTotal"), "0").Replace("'"c, "`"c)
                        str2SrcQB_AppliedAmount = NCStr(iteration_row("AppliedAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_BalanceRemaining = NCStr(iteration_row("BalanceRemaining"), "0").Replace("'"c, "`"c)
                        str2SrcQB_Memo = NCStr(iteration_row("Memo")).Replace("'"c, "`"c)
                        str2SrcQB_IsPaid = NCStr(iteration_row("IsPaid")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerMsgRefListID = NCStr(iteration_row("CustomerMsgRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerMsgRefFullName = NCStr(iteration_row("CustomerMsgRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_IsToBePrinted = NCStr(iteration_row("IsToBePrinted")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerSalesTaxCodeRefListID = NCStr(iteration_row("CustomerSalesTaxCodeRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerSalesTaxCodeRefFullName = NCStr(iteration_row("CustomerSalesTaxCodeRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_SuggestedDiscountAmount = NCStr(iteration_row("SuggestedDiscountAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_SuggestedDiscountDate = NCStr(iteration_row("SuggestedDiscountDate")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineType = NCStr(iteration_row("InvoiceLineType")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineSeqNo = NCStr(iteration_row("InvoiceLineSeqNo")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineGroupTxnLineID = NCStr(iteration_row("InvoiceLineGroupTxnLineID")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineGroupItemGroupRefListID = NCStr(iteration_row("InvoiceLineGroupItemGroupRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineGroupItemGroupRefFullName = NCStr(iteration_row("InvoiceLineGroupItemGroupRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineGroupDesc = NCStr(iteration_row("InvoiceLineGroupDesc")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineGroupQuantity = NCStr(iteration_row("InvoiceLineGroupQuantity"), "0").Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineGroupIsPrintItemsInGroup = NCStr(iteration_row("InvoiceLineGroupIsPrintItemsInGroup")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineGroupTotalAmount = NCStr(iteration_row("InvoiceLineGroupTotalAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineGroupSeqNo = NCStr(iteration_row("InvoiceLineGroupSeqNo")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineTxnLineID = NCStr(iteration_row("InvoiceLineTxnLineID")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineItemRefListID = NCStr(iteration_row("InvoiceLineItemRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineItemRefFullName = NCStr(iteration_row("InvoiceLineItemRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineDesc = NCStr(iteration_row("InvoiceLineDesc")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineQuantity = NCStr(iteration_row("InvoiceLineQuantity"), "0").Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineRate = NCStr(iteration_row("InvoiceLineRate"), "0").Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineRatePercent = NCStr(iteration_row("InvoiceLineRatePercent"), "0").Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLinePriceLevelRefListID = NCStr(iteration_row("InvoiceLinePriceLevelRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLinePriceLevelRefFullName = NCStr(iteration_row("InvoiceLinePriceLevelRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineClassRefListID = NCStr(iteration_row("InvoiceLineClassRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineClassRefFullName = NCStr(iteration_row("InvoiceLineClassRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineAmount = NCStr(iteration_row("InvoiceLineAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineServiceDate = NCStr(iteration_row("InvoiceLineServiceDate")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineSalesTaxCodeRefListID = NCStr(iteration_row("InvoiceLineSalesTaxCodeRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineSalesTaxCodeRefFullName = NCStr(iteration_row("InvoiceLineSalesTaxCodeRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineOverrideItemAccountRefListID = NCStr(iteration_row("InvoiceLineOverrideItemAccountRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_InvoiceLineOverrideItemAccountRefFullName = NCStr(iteration_row("InvoiceLineOverrideItemAccountRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_FQSaveToCache = NCStr(iteration_row("FQSaveToCache")).Replace("'"c, "`"c)
                        str2SrcQB_FQPrimaryKey = NCStr(iteration_row("FQPrimaryKey")).Replace("'"c, "`"c)
                        str2SrcQB_CustomFieldInvoiceLineOther2 = NCStr(iteration_row("CustomFieldInvoiceLineOther2")).Replace("'"c, "`"c)

                        'Change flags back to binary
                        str2SrcQB_FQSaveToCache = IIf(str2SrcQB_FQSaveToCache = "True", "1", "0")
                        str2SrcQB_IsPending = IIf(str2SrcQB_IsPending = "True", "1", "0")
                        str2SrcQB_IsFinanceCharge = IIf(str2SrcQB_IsFinanceCharge = "True", "1", "0")
                        str2SrcQB_IsPaid = IIf(str2SrcQB_IsPaid = "True", "1", "0")
                        str2SrcQB_IsToBePrinted = IIf(str2SrcQB_IsToBePrinted = "True", "1", "0")

                        'Put the information together into a string
                        'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                        str2SrcQB_QB_InvoiceLineRow = "" & _
                                                      Strings.Left(str2SrcQB_TxnID & "                  ", 18) & "   " & _
                                                      Strings.Left(str2SrcQB_TimeCreated & "                  ", 16) & "   " & _
                                                      Strings.Left(str2SrcQB_TimeModified & "                  ", 16) & "   " & _
                                                      Strings.Left(str2SrcQB_CustomerRefFullName & "                  ", 18) & "   " & _
                                                      Strings.Left(str2SrcQB_TxnDate & "                  ", 10) & "   " & _
                                                      Strings.Left(str2SrcQB_RefNumber & "                  ", 10) & "   " & _
                                                      Strings.Left(str2SrcQB_InvoiceLineItemRefFullName & "                  ", 18) & "   " & _
                                                      Strings.Left(str2SrcQB_RefNumber & "                  ", 10) & "   " & _
                                                      Strings.Left(str2SrcQB_RefNumber & "                  ", 10) & "   " & _
                                                      "" & Strings.Chr(9)

                        'put the line in the listbox
                        ShowUserMessage(strSubName, str2SrcQB_QB_InvoiceLineRow)

                        'Check to see if ListID or TxnID is in QBTable            'Yes then UPDATE record            'No then INSERT record
                        'New recordset
                        Dim iRowCount As Integer = 0
                        iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(InvoiceLineTxnLineID) FROM QB_InvoiceLine WHERE InvoiceLineTxnLineID = '" & str2SrcQB_InvoiceLineTxnLineID & "'")
                        'If iRowCount > 1 Then Stop 'Should only be one
                        If iRowCount = 1 Then 'record exists  -UPDATE
                            'DO UPDATE WORK:
                            Debug.WriteLine("UPDATE")

                            'Build the SQL string
                            strSQL1 = "UPDATE  " & Environment.NewLine & _
                                      "       QB_InvoiceLine " & Environment.NewLine & _
                                      "SET " & Environment.NewLine & _
                                      "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine & _
                                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                      "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & Environment.NewLine & _
                                      "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & Environment.NewLine & _
                                      "     , ClassRefListID = '" & str2SrcQB_ClassRefListID & "'" & Environment.NewLine & _
                                      "     , ClassRefFullName = '" & str2SrcQB_ClassRefFullName & "'" & Environment.NewLine & _
                                      "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & Environment.NewLine & _
                                      "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , TemplateRefListID = '" & str2SrcQB_TemplateRefListID & "'" & Environment.NewLine & _
                                      "     , TemplateRefFullName = '" & str2SrcQB_TemplateRefFullName & "'" & Environment.NewLine & _
                                      "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & Environment.NewLine & _
                                      "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & Environment.NewLine & _
                                      "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr1 = '" & str2SrcQB_BillAddressAddr1 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr2 = '" & str2SrcQB_BillAddressAddr2 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr3 = '" & str2SrcQB_BillAddressAddr3 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr4 = '" & str2SrcQB_BillAddressAddr4 & "'" & Environment.NewLine & _
                                      "     , BillAddressCity = '" & str2SrcQB_BillAddressCity & "'" & Environment.NewLine & _
                                      "     , BillAddressState = '" & str2SrcQB_BillAddressState & "'" & Environment.NewLine
                            strSQL2 = "     , BillAddressPostalCode = '" & str2SrcQB_BillAddressPostalCode & "'" & Environment.NewLine & _
                                      "     , BillAddressCountry = '" & str2SrcQB_BillAddressCountry & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr1 = '" & str2SrcQB_ShipAddressAddr1 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr2 = '" & str2SrcQB_ShipAddressAddr2 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr3 = '" & str2SrcQB_ShipAddressAddr3 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr4 = '" & str2SrcQB_ShipAddressAddr4 & "'" & Environment.NewLine & _
                                      "     , ShipAddressCity = '" & str2SrcQB_ShipAddressCity & "'" & Environment.NewLine & _
                                      "     , ShipAddressState = '" & str2SrcQB_ShipAddressState & "'" & Environment.NewLine & _
                                      "     , ShipAddressPostalCode = '" & str2SrcQB_ShipAddressPostalCode & "'" & Environment.NewLine & _
                                      "     , ShipAddressCountry = '" & str2SrcQB_ShipAddressCountry & "'" & Environment.NewLine & _
                                      "     , IsPending = '" & str2SrcQB_IsPending & "'" & Environment.NewLine & _
                                      "     , IsFinanceCharge = '" & str2SrcQB_IsFinanceCharge & "'" & Environment.NewLine & _
                                      "     , PONumber = '" & str2SrcQB_PONumber & "'" & Environment.NewLine & _
                                      "     , TermsRefListID = '" & str2SrcQB_TermsRefListID & "'" & Environment.NewLine & _
                                      "     , TermsRefFullName = '" & str2SrcQB_TermsRefFullName & "'" & Environment.NewLine & _
                                      "     , DueDate = '" & str2SrcQB_DueDate & "'" & Environment.NewLine & _
                                      "     , SalesRepRefListID = '" & str2SrcQB_SalesRepRefListID & "'" & Environment.NewLine & _
                                      "     , SalesRepRefFullName = '" & str2SrcQB_SalesRepRefFullName & "'" & Environment.NewLine & _
                                      "     , FOB = '" & str2SrcQB_FOB & "'" & Environment.NewLine & _
                                      "     , ShipDate = '" & str2SrcQB_ShipDate & "'" & Environment.NewLine & _
                                      "     , ShipMethodRefListID = '" & str2SrcQB_ShipMethodRefListID & "'" & Environment.NewLine & _
                                      "     , ShipMethodRefFullName = '" & str2SrcQB_ShipMethodRefFullName & "'" & Environment.NewLine & _
                                      "     , Subtotal = " & str2SrcQB_Subtotal & "" & Environment.NewLine & _
                                      "     , ItemSalesTaxRefListID = '" & str2SrcQB_ItemSalesTaxRefListID & "'" & Environment.NewLine & _
                                      "     , ItemSalesTaxRefFullName = '" & str2SrcQB_ItemSalesTaxRefFullName & "'" & Environment.NewLine
                            strSQL3 = "     , SalesTaxPercentage = " & str2SrcQB_SalesTaxPercentage & "" & Environment.NewLine & _
                                      "     , SalesTaxTotal = " & str2SrcQB_SalesTaxTotal & "" & Environment.NewLine & _
                                      "     , AppliedAmount = " & str2SrcQB_AppliedAmount & "" & Environment.NewLine & _
                                      "     , BalanceRemaining = " & str2SrcQB_BalanceRemaining & "" & Environment.NewLine & _
                                      "     , Memo = '" & str2SrcQB_Memo & "'" & Environment.NewLine & _
                                      "     , IsPaid = '" & str2SrcQB_IsPaid & "'" & Environment.NewLine & _
                                      "     , CustomerMsgRefListID = '" & str2SrcQB_CustomerMsgRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerMsgRefFullName = '" & str2SrcQB_CustomerMsgRefFullName & "'" & Environment.NewLine & _
                                      "     , IsToBePrinted = '" & str2SrcQB_IsToBePrinted & "'" & Environment.NewLine & _
                                      "     , CustomerSalesTaxCodeRefListID = '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerSalesTaxCodeRefFullName = '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'" & Environment.NewLine & _
                                      "     , SuggestedDiscountAmount = " & str2SrcQB_SuggestedDiscountAmount & "" & Environment.NewLine & _
                                      "     , SuggestedDiscountDate = '" & str2SrcQB_SuggestedDiscountDate & "'" & Environment.NewLine & _
                                      "     , InvoiceLineType = '" & str2SrcQB_InvoiceLineType & "'" & Environment.NewLine & _
                                      "     , InvoiceLineSeqNo = '" & str2SrcQB_InvoiceLineSeqNo & "'" & Environment.NewLine & _
                                      "     , InvoiceLineGroupTxnLineID = '" & str2SrcQB_InvoiceLineGroupTxnLineID & "'" & Environment.NewLine & _
                                      "     , InvoiceLineGroupItemGroupRefListID = '" & str2SrcQB_InvoiceLineGroupItemGroupRefListID & "'" & Environment.NewLine & _
                                      "     , InvoiceLineGroupItemGroupRefFullName = '" & str2SrcQB_InvoiceLineGroupItemGroupRefFullName & "'" & Environment.NewLine & _
                                      "     , InvoiceLineGroupDesc = '" & str2SrcQB_InvoiceLineGroupDesc & "'" & Environment.NewLine & _
                                      "     , InvoiceLineGroupQuantity = " & str2SrcQB_InvoiceLineGroupQuantity & "" & Environment.NewLine & _
                                      "     , InvoiceLineGroupIsPrintItemsInGroup = '" & str2SrcQB_InvoiceLineGroupIsPrintItemsInGroup & "'" & Environment.NewLine & _
                                      "     , InvoiceLineGroupTotalAmount = " & str2SrcQB_InvoiceLineGroupTotalAmount & "" & Environment.NewLine & _
                                      "     , InvoiceLineGroupSeqNo = '" & str2SrcQB_InvoiceLineGroupSeqNo & "'" & Environment.NewLine & _
                                      "     , InvoiceLineTxnLineID = '" & str2SrcQB_InvoiceLineTxnLineID & "'" & Environment.NewLine & _
                                      "     , InvoiceLineItemRefListID = '" & str2SrcQB_InvoiceLineItemRefListID & "'" & Environment.NewLine
                            strSQL4 = "     , InvoiceLineItemRefFullName = '" & str2SrcQB_InvoiceLineItemRefFullName & "'" & Environment.NewLine & _
                                      "     , InvoiceLineDesc = '" & str2SrcQB_InvoiceLineDesc & "'" & Environment.NewLine & _
                                      "     , InvoiceLineQuantity = " & str2SrcQB_InvoiceLineQuantity & "" & Environment.NewLine & _
                                      "     , InvoiceLineRate = " & str2SrcQB_InvoiceLineRate & "" & Environment.NewLine & _
                                      "     , InvoiceLineRatePercent = " & str2SrcQB_InvoiceLineRatePercent & "" & Environment.NewLine & _
                                      "     , InvoiceLinePriceLevelRefListID = '" & str2SrcQB_InvoiceLinePriceLevelRefListID & "'" & Environment.NewLine & _
                                      "     , InvoiceLinePriceLevelRefFullName = '" & str2SrcQB_InvoiceLinePriceLevelRefFullName & "'" & Environment.NewLine & _
                                      "     , InvoiceLineClassRefListID = '" & str2SrcQB_InvoiceLineClassRefListID & "'" & Environment.NewLine & _
                                      "     , InvoiceLineClassRefFullName = '" & str2SrcQB_InvoiceLineClassRefFullName & "'" & Environment.NewLine & _
                                      "     , InvoiceLineAmount = " & str2SrcQB_InvoiceLineAmount & "" & Environment.NewLine & _
                                      "     , InvoiceLineServiceDate = '" & str2SrcQB_InvoiceLineServiceDate & "'" & Environment.NewLine & _
                                      "     , InvoiceLineSalesTaxCodeRefListID = '" & str2SrcQB_InvoiceLineSalesTaxCodeRefListID & "'" & Environment.NewLine & _
                                      "     , InvoiceLineSalesTaxCodeRefFullName = '" & str2SrcQB_InvoiceLineSalesTaxCodeRefFullName & "'" & Environment.NewLine & _
                                      "     , InvoiceLineOverrideItemAccountRefListID = '" & str2SrcQB_InvoiceLineOverrideItemAccountRefListID & "'" & Environment.NewLine & _
                                      "     , InvoiceLineOverrideItemAccountRefFullName = '" & str2SrcQB_InvoiceLineOverrideItemAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , FQSaveToCache = '" & str2SrcQB_FQSaveToCache & "'" & Environment.NewLine & _
                                      "     , FQPrimaryKey = '" & str2SrcQB_FQPrimaryKey & "'" & Environment.NewLine & _
                                      "     , CustomFieldInvoiceLineOther1 = '" & str2SrcQB_CustomFieldInvoiceLineOther1 & "'" & Environment.NewLine & _
                                      "     , CustomFieldInvoiceLineOther2 = '" & str2SrcQB_CustomFieldInvoiceLineOther2 & "'" & Environment.NewLine & _
                                      "     , CustomFieldInvoiceLineGroupOther1 = '" & str2SrcQB_CustomFieldInvoiceLineGroupOther1 & "'" & Environment.NewLine & _
                                      "     , CustomFieldInvoiceLineGroupOther2 = '" & str2SrcQB_CustomFieldInvoiceLineGroupOther2 & "'" & Environment.NewLine & _
                                      "     , CustomFieldInvoiceLineGroupLineOther1 = '" & str2SrcQB_CustomFieldInvoiceLineGroupLineOther1 & "'" & Environment.NewLine & _
                                      "     , CustomFieldInvoiceLineGroupLineOther2 = '" & str2SrcQB_CustomFieldInvoiceLineGroupLineOther2 & "'" & Environment.NewLine & _
                                      "     , CustomFieldOther = '" & str2SrcQB_CustomFieldOther & "'" & Environment.NewLine
                            strSQL5 = "WHERE " & Environment.NewLine & _
                                      "       InvoiceLineTxnLineID = '" & str2SrcQB_InvoiceLineTxnLineID & "'" & Environment.NewLine

                            'Combine the strings
                            strTableUpdate = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 '& strSQL6
                            'Debug.Print strTableUpdate

                            'Execute the insert
                            '*cnDBPM.Execute strTableUpdate
                            SQLHelper.ExecuteSQL(cnMax, strTableUpdate)


                        Else
                            'record not exist  -INSERT
                            'DO INSERT WORK:
                            Debug.WriteLine("INSERT")

                            'Build the SQL string
                            strSQL1 = "INSERT INTO QB_InvoiceLine " & Environment.NewLine & _
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
                                      "   , IsFinanceCharge " & Environment.NewLine & _
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
                                      "   , SalesTaxTotal " & Environment.NewLine
                            strSQL3 = "   , AppliedAmount " & Environment.NewLine & _
                                      "   , BalanceRemaining " & Environment.NewLine & _
                                      "   , Memo " & Environment.NewLine & _
                                      "   , IsPaid " & Environment.NewLine & _
                                      "   , CustomerMsgRefListID " & Environment.NewLine & _
                                      "   , CustomerMsgRefFullName " & Environment.NewLine & _
                                      "   , IsToBePrinted " & Environment.NewLine & _
                                      "   , CustomerSalesTaxCodeRefListID " & Environment.NewLine & _
                                      "   , CustomerSalesTaxCodeRefFullName " & Environment.NewLine & _
                                      "   , SuggestedDiscountAmount " & Environment.NewLine & _
                                      "   , SuggestedDiscountDate " & Environment.NewLine & _
                                      "   , InvoiceLineType " & Environment.NewLine & _
                                      "   , InvoiceLineSeqNo " & Environment.NewLine & _
                                      "   , InvoiceLineGroupTxnLineID " & Environment.NewLine & _
                                      "   , InvoiceLineGroupItemGroupRefListID " & Environment.NewLine & _
                                      "   , InvoiceLineGroupItemGroupRefFullName " & Environment.NewLine & _
                                      "   , InvoiceLineGroupDesc " & Environment.NewLine & _
                                      "   , InvoiceLineGroupQuantity " & Environment.NewLine & _
                                      "   , InvoiceLineGroupIsPrintItemsInGroup " & Environment.NewLine & _
                                      "   , InvoiceLineGroupTotalAmount " & Environment.NewLine & _
                                      "   , InvoiceLineGroupSeqNo " & Environment.NewLine & _
                                      "   , InvoiceLineTxnLineID " & Environment.NewLine & _
                                      "   , InvoiceLineItemRefListID " & Environment.NewLine & _
                                      "   , InvoiceLineItemRefFullName " & Environment.NewLine & _
                                      "   , InvoiceLineDesc " & Environment.NewLine
                            strSQL4 = "   , InvoiceLineQuantity " & Environment.NewLine & _
                                      "   , InvoiceLineRate " & Environment.NewLine & _
                                      "   , InvoiceLineRatePercent " & Environment.NewLine & _
                                      "   , InvoiceLinePriceLevelRefListID " & Environment.NewLine & _
                                      "   , InvoiceLinePriceLevelRefFullName " & Environment.NewLine & _
                                      "   , InvoiceLineClassRefListID " & Environment.NewLine & _
                                      "   , InvoiceLineClassRefFullName " & Environment.NewLine & _
                                      "   , InvoiceLineAmount " & Environment.NewLine & _
                                      "   , InvoiceLineServiceDate " & Environment.NewLine & _
                                      "   , InvoiceLineSalesTaxCodeRefListID " & Environment.NewLine & _
                                      "   , InvoiceLineSalesTaxCodeRefFullName " & Environment.NewLine & _
                                      "   , InvoiceLineOverrideItemAccountRefListID " & Environment.NewLine & _
                                      "   , InvoiceLineOverrideItemAccountRefFullName " & Environment.NewLine & _
                                      "   , FQSaveToCache " & Environment.NewLine & _
                                      "   , FQPrimaryKey " & Environment.NewLine & _
                                      "   , CustomFieldInvoiceLineOther1 " & Environment.NewLine & _
                                      "   , CustomFieldInvoiceLineOther2 " & Environment.NewLine & _
                                      "   , CustomFieldInvoiceLineGroupOther1 " & Environment.NewLine & _
                                      "   , CustomFieldInvoiceLineGroupOther2 " & Environment.NewLine & _
                                      "   , CustomFieldInvoiceLineGroupLineOther1 " & Environment.NewLine & _
                                      "   , CustomFieldInvoiceLineGroupLineOther2 " & Environment.NewLine & _
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
                                      "   , '" & str2SrcQB_IsFinanceCharge & "'  --IsFinanceCharge" & Environment.NewLine & _
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
                                      "   , " & str2SrcQB_SalesTaxTotal & "  --SalesTaxTotal" & Environment.NewLine
                            strSQL7 = "   , " & str2SrcQB_AppliedAmount & "  --AppliedAmount" & Environment.NewLine & _
                                      "   , " & str2SrcQB_BalanceRemaining & "  --BalanceRemaining" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_Memo & "'  --Memo" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_IsPaid & "'  --IsPaid" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomerMsgRefListID & "'  --CustomerMsgRefListID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomerMsgRefFullName & "'  --CustomerMsgRefFullName" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_IsToBePrinted & "'  --IsToBePrinted" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'  --CustomerSalesTaxCodeRefListID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'  --CustomerSalesTaxCodeRefFullName" & Environment.NewLine & _
                                      "   , " & str2SrcQB_SuggestedDiscountAmount & "  --SuggestedDiscountAmount" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_SuggestedDiscountDate & "'  --SuggestedDiscountDate" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineType & "'  --InvoiceLineType" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineSeqNo & "'  --InvoiceLineSeqNo" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineGroupTxnLineID & "'  --InvoiceLineGroupTxnLineID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineGroupItemGroupRefListID & "'  --InvoiceLineGroupItemGroupRefListID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineGroupItemGroupRefFullName & "'  --InvoiceLineGroupItemGroupRefFullName" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineGroupDesc & "'  --InvoiceLineGroupDesc" & Environment.NewLine & _
                                      "   , " & str2SrcQB_InvoiceLineGroupQuantity & "  --InvoiceLineGroupQuantity" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineGroupIsPrintItemsInGroup & "'  --InvoiceLineGroupIsPrintItemsInGroup" & Environment.NewLine & _
                                      "   , " & str2SrcQB_InvoiceLineGroupTotalAmount & "  --InvoiceLineGroupTotalAmount" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineGroupSeqNo & "'  --InvoiceLineGroupSeqNo" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineTxnLineID & "'  --InvoiceLineTxnLineID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineItemRefListID & "'  --InvoiceLineItemRefListID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineItemRefFullName & "'  --InvoiceLineItemRefFullName" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineDesc & "'  --InvoiceLineDesc" & Environment.NewLine
                            strSQL8 = "   , " & str2SrcQB_InvoiceLineQuantity & "  --InvoiceLineQuantity" & Environment.NewLine & _
                                      "   , " & str2SrcQB_InvoiceLineRate & "  --InvoiceLineRate" & Environment.NewLine & _
                                      "   , " & str2SrcQB_InvoiceLineRatePercent & "  --InvoiceLineRatePercent" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLinePriceLevelRefListID & "'  --InvoiceLinePriceLevelRefListID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLinePriceLevelRefFullName & "'  --InvoiceLinePriceLevelRefFullName" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineClassRefListID & "'  --InvoiceLineClassRefListID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineClassRefFullName & "'  --InvoiceLineClassRefFullName" & Environment.NewLine & _
                                      "   , " & str2SrcQB_InvoiceLineAmount & "  --InvoiceLineAmount" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineServiceDate & "'  --InvoiceLineServiceDate" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineSalesTaxCodeRefListID & "'  --InvoiceLineSalesTaxCodeRefListID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineSalesTaxCodeRefFullName & "'  --InvoiceLineSalesTaxCodeRefFullName" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineOverrideItemAccountRefListID & "'  --InvoiceLineOverrideItemAccountRefListID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_InvoiceLineOverrideItemAccountRefFullName & "'  --InvoiceLineOverrideItemAccountRefFullName" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_FQSaveToCache & "'  --FQSaveToCache" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_FQPrimaryKey & "'  --FQPrimaryKey" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomFieldInvoiceLineOther1 & "'  --CustomFieldInvoiceLineOther1" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomFieldInvoiceLineOther2 & "'  --CustomFieldInvoiceLineOther2" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomFieldInvoiceLineGroupOther1 & "'  --CustomFieldInvoiceLineGroupOther1" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomFieldInvoiceLineGroupOther2 & "'  --CustomFieldInvoiceLineGroupOther2" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomFieldInvoiceLineGroupLineOther1 & "'  --CustomFieldInvoiceLineGroupLineOther1" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomFieldInvoiceLineGroupLineOther2 & "'  --CustomFieldInvoiceLineGroupLineOther2" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomFieldOther & "' ) --CustomFieldOther" & Environment.NewLine

                            'Combine the strings
                            strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6 & strSQL7 & strSQL8

                            SQLHelper.ExecuteSQL(cnMax, strTableInsert)

                        End If

                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For
                    End Try

                Next iteration_row
                ShowUserMessage(strSubName, "Finished Processing QB_InvoiceLine Records", True)

            End If
        End Using


    End Sub


    Public Sub RefreshQB_Invoice()
        'First RUN_THROUGH COMPLETE
        Dim str2SrcQB_IsActive As String = ""

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQB_Invoice" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        ShowUserMessage(strSubName, "RefreshQB: Processing QB_Invoice Records", "RefreshQB: Processing QB_Invoice Records", True)

        'FOR PART 2SrcQB_ - Get records from QB_Invoice
        Debug.WriteLine("List2SrcQB_QB_Invoice")
        Dim str2SrcQB_QB_InvoiceSQL, str2SrcQB_QB_InvoiceRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ClassRefListID, str2SrcQB_ClassRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TemplateRefListID, str2SrcQB_TemplateRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_BillAddressAddr1, str2SrcQB_BillAddressAddr2, str2SrcQB_BillAddressAddr3, str2SrcQB_BillAddressAddr4, str2SrcQB_BillAddressCity, str2SrcQB_BillAddressState, str2SrcQB_BillAddressPostalCode, str2SrcQB_BillAddressCountry, str2SrcQB_ShipAddressAddr1, str2SrcQB_ShipAddressAddr2, str2SrcQB_ShipAddressAddr3, str2SrcQB_ShipAddressAddr4, str2SrcQB_ShipAddressCity, str2SrcQB_ShipAddressState, str2SrcQB_ShipAddressPostalCode, str2SrcQB_ShipAddressCountry, str2SrcQB_IsPending, str2SrcQB_IsFinanceCharge, str2SrcQB_PONumber, str2SrcQB_TermsRefListID, str2SrcQB_TermsRefFullName, str2SrcQB_DueDate, str2SrcQB_SalesRepRefListID, str2SrcQB_SalesRepRefFullName, str2SrcQB_FOB, str2SrcQB_ShipDate, str2SrcQB_ShipMethodRefListID, str2SrcQB_ShipMethodRefFullName, str2SrcQB_Subtotal, str2SrcQB_ItemSalesTaxRefListID, str2SrcQB_ItemSalesTaxRefFullName, str2SrcQB_SalesTaxPercentage, str2SrcQB_SalesTaxTotal, str2SrcQB_AppliedAmount, str2SrcQB_BalanceRemaining, str2SrcQB_Memo, str2SrcQB_IsPaid, str2SrcQB_CustomerMsgRefListID, str2SrcQB_CustomerMsgRefFullName, str2SrcQB_IsToBePrinted, str2SrcQB_CustomerSalesTaxCodeRefListID, str2SrcQB_CustomerSalesTaxCodeRefFullName, str2SrcQB_SuggestedDiscountAmount, str2SrcQB_SuggestedDiscountDate, str2SrcQB_CustomFieldOther As String
        'This routine gets the 2SrcQB_QB_Invoice from the database according to the selection in str2SrcQB_QB_InvoiceSQL.
        'It then puts those 2SrcQB_QB_Invoice in the list box

        'FOR PART 3TestID_
        Debug.WriteLine("List3TestID_QBTable")
        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String
        strTableInsert = ""

        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        Using rs2SrcQB_QB_Invoice As DataSet = New DataSet() '*** TAKE QB_ OFF OF TABLE NAME ***
            str2SrcQB_QB_InvoiceSQL = "SELECT * FROM Invoice WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_Invoice & "'}" ' ORDER BY TimeModified"

            Debug.WriteLine(str2SrcQB_QB_InvoiceSQL)
            Using adap As New OdbcDataAdapter(str2SrcQB_QB_InvoiceSQL, cnQuickBooks)
                rs2SrcQB_QB_Invoice.Tables.Clear()
                adap.Fill(rs2SrcQB_QB_Invoice)
            End Using

            Dim curRow As Integer = 0
            Dim rowCount As Integer = rs2SrcQB_QB_Invoice.Tables(0).Rows.Count

            If rowCount > 0 Then

                'Show what's processing in the listbox
                ShowUserMessage(strSubName, "Processing  " & rowCount.ToString & "  QB_Invoice  Records")

                For Each iteration_row As DataRow In rs2SrcQB_QB_Invoice.Tables(0).Rows
                    curRow += 1
                    ShowUserMessage(strSubName, "Processing " & curRow.ToString & " of " & rowCount.ToString & " QB_Invoice Records")
                    strSQL1 = ""
                    strSQL2 = ""
                    strSQL3 = ""
                    strSQL4 = ""
                    strSQL5 = ""
                    strSQL6 = ""

                    Try

                        'get the columns from the database
                        str2SrcQB_TxnID = NCStr(iteration_row("TxnID")).Replace("'"c, "`"c)
                        str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                        str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                        str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                        str2SrcQB_TxnNumber = NCStr(iteration_row("TxnNumber"), "0").Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefListID = NCStr(iteration_row("CustomerRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefFullName = NCStr(iteration_row("CustomerRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ClassRefListID = NCStr(iteration_row("ClassRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ClassRefFullName = NCStr(iteration_row("ClassRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefListID = NCStr(iteration_row("ARAccountRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefFullName = NCStr(iteration_row("ARAccountRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_TemplateRefListID = NCStr(iteration_row("TemplateRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_TemplateRefFullName = NCStr(iteration_row("TemplateRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDate = NCStr(iteration_row("TxnDate")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDateMacro = NCStr(iteration_row("TxnDateMacro")).Replace("'"c, "`"c)
                        str2SrcQB_RefNumber = NCStr(iteration_row("RefNumber")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr1 = NCStr(iteration_row("BillAddressAddr1")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr2 = NCStr(iteration_row("BillAddressAddr2")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr3 = NCStr(iteration_row("BillAddressAddr3")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr4 = NCStr(iteration_row("BillAddressAddr4")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressCity = NCStr(iteration_row("BillAddressCity")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressState = NCStr(iteration_row("BillAddressState")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressPostalCode = NCStr(iteration_row("BillAddressPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressCountry = NCStr(iteration_row("BillAddressCountry")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr1 = NCStr(iteration_row("ShipAddressAddr1")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr2 = NCStr(iteration_row("ShipAddressAddr2")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr3 = NCStr(iteration_row("ShipAddressAddr3")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr4 = NCStr(iteration_row("ShipAddressAddr4")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressCity = NCStr(iteration_row("ShipAddressCity")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressState = NCStr(iteration_row("ShipAddressState")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressPostalCode = NCStr(iteration_row("ShipAddressPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressCountry = NCStr(iteration_row("ShipAddressCountry")).Replace("'"c, "`"c)
                        str2SrcQB_IsPending = NCStr(iteration_row("IsPending")).Replace("'"c, "`"c)
                        str2SrcQB_IsFinanceCharge = NCStr(iteration_row("IsFinanceCharge")).Replace("'"c, "`"c)
                        str2SrcQB_PONumber = NCStr(iteration_row("PONumber")).Replace("'"c, "`"c)
                        str2SrcQB_TermsRefListID = NCStr(iteration_row("TermsRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_TermsRefFullName = NCStr(iteration_row("TermsRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_DueDate = NCStr(iteration_row("DueDate")).Replace("'"c, "`"c)
                        str2SrcQB_SalesRepRefListID = NCStr(iteration_row("SalesRepRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_SalesRepRefFullName = NCStr(iteration_row("SalesRepRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_FOB = NCStr(iteration_row("FOB")).Replace("'"c, "`"c)
                        str2SrcQB_ShipDate = NCStr(iteration_row("ShipDate")).Replace("'"c, "`"c)
                        str2SrcQB_ShipMethodRefListID = NCStr(iteration_row("ShipMethodRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ShipMethodRefFullName = NCStr(iteration_row("ShipMethodRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_Subtotal = NCStr(iteration_row("Subtotal"), "0").Replace("'"c, "`"c)
                        str2SrcQB_ItemSalesTaxRefListID = NCStr(iteration_row("ItemSalesTaxRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ItemSalesTaxRefFullName = NCStr(iteration_row("ItemSalesTaxRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_SalesTaxPercentage = NCStr(iteration_row("SalesTaxPercentage"), "0").Replace("'"c, "`"c)
                        str2SrcQB_SalesTaxTotal = NCStr(iteration_row("SalesTaxTotal"), "0").Replace("'"c, "`"c)
                        str2SrcQB_AppliedAmount = NCStr(iteration_row("AppliedAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_BalanceRemaining = NCStr(iteration_row("BalanceRemaining"), "0").Replace("'"c, "`"c)
                        str2SrcQB_Memo = NCStr(iteration_row("Memo")).Replace("'"c, "`"c)
                        str2SrcQB_IsPaid = NCStr(iteration_row("IsPaid")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerMsgRefListID = NCStr(iteration_row("CustomerMsgRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerMsgRefFullName = NCStr(iteration_row("CustomerMsgRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_IsToBePrinted = NCStr(iteration_row("IsToBePrinted")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerSalesTaxCodeRefListID = NCStr(iteration_row("CustomerSalesTaxCodeRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerSalesTaxCodeRefFullName = NCStr(iteration_row("CustomerSalesTaxCodeRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_SuggestedDiscountAmount = NCStr(iteration_row("SuggestedDiscountAmount"), "0").Replace("'"c, "`"c)
                        str2SrcQB_SuggestedDiscountDate = NCStr(iteration_row("SuggestedDiscountDate")).Replace("'"c, "`"c)
                        str2SrcQB_CustomFieldOther = NCStr(iteration_row("CustomFieldOther")).Replace("'"c, "`"c)

                        'Change flags back to binary
                        str2SrcQB_IsActive = IIf(str2SrcQB_IsActive = "True", "1", "0")
                        str2SrcQB_IsPending = IIf(str2SrcQB_IsPending = "True", "1", "0")
                        str2SrcQB_IsFinanceCharge = IIf(str2SrcQB_IsFinanceCharge = "True", "1", "0")
                        str2SrcQB_IsPaid = IIf(str2SrcQB_IsPaid = "True", "1", "0")
                        str2SrcQB_IsToBePrinted = IIf(str2SrcQB_IsToBePrinted = "True", "1", "0")

                        'Put the information together into a string
                        'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                        str2SrcQB_QB_InvoiceRow = "" & _
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

                        ShowUserMessage(strSubName, str2SrcQB_QB_InvoiceRow)

                        UpdateQBCustomerBalance(str2SrcQB_CustomerRefListID)
                        'UpdateQBReceivePaymentLine(str2SrcQB_CustomerRefFullName)

                        Dim iRowCount As Integer = 0
                        iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(TxnID) FROM QB_Invoice WHERE TxnID = '" & str2SrcQB_TxnID & "'")

                        If iRowCount = 1 Then 'record exists  -UPDATE

                            'DO UPDATE WORK:
                            Debug.WriteLine("UPDATE")


                            strSQL1 = "UPDATE  " & Environment.NewLine & _
                                      "       QB_Invoice " & Environment.NewLine & _
                                      "SET " & Environment.NewLine & _
                                      "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine & _
                                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                      "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & Environment.NewLine & _
                                      "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & Environment.NewLine & _
                                      "     , ClassRefListID = '" & str2SrcQB_ClassRefListID & "'" & Environment.NewLine & _
                                      "     , ClassRefFullName = '" & str2SrcQB_ClassRefFullName & "'" & Environment.NewLine & _
                                      "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & Environment.NewLine & _
                                      "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , TemplateRefListID = '" & str2SrcQB_TemplateRefListID & "'" & Environment.NewLine & _
                                      "     , TemplateRefFullName = '" & str2SrcQB_TemplateRefFullName & "'" & Environment.NewLine & _
                                      "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & Environment.NewLine & _
                                      "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & Environment.NewLine & _
                                      "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr1 = '" & str2SrcQB_BillAddressAddr1 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr2 = '" & str2SrcQB_BillAddressAddr2 & "'" & Environment.NewLine
                            strSQL2 = "     , BillAddressAddr3 = '" & str2SrcQB_BillAddressAddr3 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr4 = '" & str2SrcQB_BillAddressAddr4 & "'" & Environment.NewLine & _
                                      "     , BillAddressCity = '" & str2SrcQB_BillAddressCity & "'" & Environment.NewLine & _
                                      "     , BillAddressState = '" & str2SrcQB_BillAddressState & "'" & Environment.NewLine & _
                                      "     , BillAddressPostalCode = '" & str2SrcQB_BillAddressPostalCode & "'" & Environment.NewLine & _
                                      "     , BillAddressCountry = '" & str2SrcQB_BillAddressCountry & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr1 = '" & str2SrcQB_ShipAddressAddr1 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr2 = '" & str2SrcQB_ShipAddressAddr2 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr3 = '" & str2SrcQB_ShipAddressAddr3 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr4 = '" & str2SrcQB_ShipAddressAddr4 & "'" & Environment.NewLine & _
                                      "     , ShipAddressCity = '" & str2SrcQB_ShipAddressCity & "'" & Environment.NewLine & _
                                      "     , ShipAddressState = '" & str2SrcQB_ShipAddressState & "'" & Environment.NewLine & _
                                      "     , ShipAddressPostalCode = '" & str2SrcQB_ShipAddressPostalCode & "'" & Environment.NewLine & _
                                      "     , ShipAddressCountry = '" & str2SrcQB_ShipAddressCountry & "'" & Environment.NewLine & _
                                      "     , IsPending = '" & str2SrcQB_IsPending & "'" & Environment.NewLine & _
                                      "     , IsFinanceCharge = '" & str2SrcQB_IsFinanceCharge & "'" & Environment.NewLine & _
                                      "     , PONumber = '" & str2SrcQB_PONumber & "'" & Environment.NewLine & _
                                      "     , TermsRefListID = '" & str2SrcQB_TermsRefListID & "'" & Environment.NewLine & _
                                      "     , TermsRefFullName = '" & str2SrcQB_TermsRefFullName & "'" & Environment.NewLine & _
                                      "     , DueDate = '" & str2SrcQB_DueDate & "'" & Environment.NewLine & _
                                      "     , SalesRepRefListID = '" & str2SrcQB_SalesRepRefListID & "'" & Environment.NewLine & _
                                      "     , SalesRepRefFullName = '" & str2SrcQB_SalesRepRefFullName & "'" & Environment.NewLine
                            strSQL3 = "     , FOB = '" & str2SrcQB_FOB & "'" & Environment.NewLine & _
                                      "     , ShipDate = '" & str2SrcQB_ShipDate & "'" & Environment.NewLine & _
                                      "     , ShipMethodRefListID = '" & str2SrcQB_ShipMethodRefListID & "'" & Environment.NewLine & _
                                      "     , ShipMethodRefFullName = '" & str2SrcQB_ShipMethodRefFullName & "'" & Environment.NewLine & _
                                      "     , Subtotal = " & str2SrcQB_Subtotal & "" & Environment.NewLine & _
                                      "     , ItemSalesTaxRefListID = '" & str2SrcQB_ItemSalesTaxRefListID & "'" & Environment.NewLine & _
                                      "     , ItemSalesTaxRefFullName = '" & str2SrcQB_ItemSalesTaxRefFullName & "'" & Environment.NewLine & _
                                      "     , SalesTaxPercentage = " & str2SrcQB_SalesTaxPercentage & "" & Environment.NewLine & _
                                      "     , SalesTaxTotal = " & str2SrcQB_SalesTaxTotal & "" & Environment.NewLine & _
                                      "     , AppliedAmount = " & str2SrcQB_AppliedAmount & "" & Environment.NewLine & _
                                      "     , BalanceRemaining = " & str2SrcQB_BalanceRemaining & "" & Environment.NewLine & _
                                      "     , Memo = '" & str2SrcQB_Memo & "'" & Environment.NewLine & _
                                      "     , IsPaid = '" & str2SrcQB_IsPaid & "'" & Environment.NewLine & _
                                      "     , CustomerMsgRefListID = '" & str2SrcQB_CustomerMsgRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerMsgRefFullName = '" & str2SrcQB_CustomerMsgRefFullName & "'" & Environment.NewLine & _
                                      "     , IsToBePrinted = '" & str2SrcQB_IsToBePrinted & "'" & Environment.NewLine & _
                                      "     , CustomerSalesTaxCodeRefListID = '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerSalesTaxCodeRefFullName = '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'" & Environment.NewLine & _
                                      "     , SuggestedDiscountAmount = " & str2SrcQB_SuggestedDiscountAmount & "" & Environment.NewLine & _
                                      "     , SuggestedDiscountDate = '" & str2SrcQB_SuggestedDiscountDate & "'" & Environment.NewLine & _
                                      "     , CustomFieldOther = '" & str2SrcQB_CustomFieldOther & "'" & Environment.NewLine & _
                                      "WHERE " & Environment.NewLine & _
                                      "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine

                            'Combine the strings
                            strTableUpdate = strSQL1 & strSQL2 & strSQL3
                            SQLHelper.ExecuteSQL(cnMax, strTableUpdate)

                        Else
                            'record not exist  -INSERT
                            Debug.WriteLine("INSERT")

                            strSQL1 = "INSERT INTO QB_Invoice " & Environment.NewLine & _
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
                                      "   , IsFinanceCharge " & Environment.NewLine & _
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
                                      "   , AppliedAmount " & Environment.NewLine & _
                                      "   , BalanceRemaining " & Environment.NewLine & _
                                      "   , Memo " & Environment.NewLine & _
                                      "   , IsPaid " & Environment.NewLine & _
                                      "   , CustomerMsgRefListID " & Environment.NewLine & _
                                      "   , CustomerMsgRefFullName " & Environment.NewLine & _
                                      "   , IsToBePrinted " & Environment.NewLine & _
                                      "   , CustomerSalesTaxCodeRefListID " & Environment.NewLine & _
                                      "   , CustomerSalesTaxCodeRefFullName " & Environment.NewLine & _
                                      "   , SuggestedDiscountAmount " & Environment.NewLine & _
                                      "   , SuggestedDiscountDate " & Environment.NewLine & _
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
                                      "   , '" & str2SrcQB_IsFinanceCharge & "'  --IsFinanceCharge" & Environment.NewLine & _
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
                                      "   , " & str2SrcQB_AppliedAmount & "  --AppliedAmount" & Environment.NewLine & _
                                      "   , " & str2SrcQB_BalanceRemaining & "  --BalanceRemaining" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_Memo & "'  --Memo" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_IsPaid & "'  --IsPaid" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomerMsgRefListID & "'  --CustomerMsgRefListID" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_CustomerMsgRefFullName & "'  --CustomerMsgRefFullName" & Environment.NewLine & _
                                      "   , '" & str2SrcQB_IsToBePrinted & "'  --IsToBePrinted" & Environment.NewLine & _
                                          "   , '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'  --CustomerSalesTaxCodeRefListID" & Environment.NewLine & _
                                          "   , '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'  --CustomerSalesTaxCodeRefFullName" & Environment.NewLine & _
                                          "   , " & str2SrcQB_SuggestedDiscountAmount & "  --SuggestedDiscountAmount" & Environment.NewLine & _
                                          "   , '" & str2SrcQB_SuggestedDiscountDate & "'  --SuggestedDiscountDate" & Environment.NewLine & _
                                          "   , '" & str2SrcQB_CustomFieldOther & "' ) --CustomFieldOther" & Environment.NewLine

                            'Combine the strings
                            strTableInsert = strSQL1 & strSQL2 & strSQL3 & strSQL4 & strSQL5 & strSQL6
                            SQLHelper.ExecuteSQL(cnMax, strTableInsert)

                        End If

                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For
                    End Try

                Next iteration_row

                ShowUserMessage(strSubName, "RefreshQB: Finished Processing QB_Invoice Records", "RefreshQB: Finished Processing QB_Invoice Records", True)

            End If
        End Using
    End Sub

    Public Sub RefreshQB_CreditMemo()
        'FIRST RUN_THROUGH COMPLETE


        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQB_CreditMemo" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        'FOR PART 2SrcQB_ - Get records from QB_CreditMemo
        Debug.WriteLine("List2SrcQB_QB_CreditMemo")
        Dim str2SrcQB_QB_CreditMemoSQL, str2SrcQB_QB_CreditMemoRow, str2SrcQB_TxnID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_TxnNumber, str2SrcQB_CustomerRefListID, str2SrcQB_CustomerRefFullName, str2SrcQB_ClassRefListID, str2SrcQB_ClassRefFullName, str2SrcQB_ARAccountRefListID, str2SrcQB_ARAccountRefFullName, str2SrcQB_TemplateRefListID, str2SrcQB_TemplateRefFullName, str2SrcQB_TxnDate, str2SrcQB_TxnDateMacro, str2SrcQB_RefNumber, str2SrcQB_BillAddressAddr1, str2SrcQB_BillAddressAddr2, str2SrcQB_BillAddressAddr3, str2SrcQB_BillAddressAddr4, str2SrcQB_BillAddressCity, str2SrcQB_BillAddressState, str2SrcQB_BillAddressPostalCode, str2SrcQB_BillAddressCountry, str2SrcQB_ShipAddressAddr1, str2SrcQB_ShipAddressAddr2, str2SrcQB_ShipAddressAddr3, str2SrcQB_ShipAddressAddr4, str2SrcQB_ShipAddressCity, str2SrcQB_ShipAddressState, str2SrcQB_ShipAddressPostalCode, str2SrcQB_ShipAddressCountry, str2SrcQB_IsPending, str2SrcQB_PONumber, str2SrcQB_TermsRefListID, str2SrcQB_TermsRefFullName, str2SrcQB_DueDate, str2SrcQB_SalesRepRefListID, str2SrcQB_SalesRepRefFullName, str2SrcQB_FOB, str2SrcQB_ShipDate, str2SrcQB_ShipMethodRefListID, str2SrcQB_ShipMethodRefFullName, str2SrcQB_Subtotal, str2SrcQB_ItemSalesTaxRefListID, str2SrcQB_ItemSalesTaxRefFullName, str2SrcQB_SalesTaxPercentage, str2SrcQB_SalesTaxTotal, str2SrcQB_TotalAmount, str2SrcQB_CreditRemaining, str2SrcQB_Memo, str2SrcQB_CustomerMsgRefListID, str2SrcQB_CustomerMsgRefFullName, str2SrcQB_IsToBePrinted, str2SrcQB_CustomerSalesTaxCodeRefListID, str2SrcQB_CustomerSalesTaxCodeRefFullName, str2SrcQB_CustomFieldOther As String
        'This routine gets the 2SrcQB_QB_CreditMemo from the database according to the selection in str2SrcQB_QB_CreditMemoSQL.
        'It then puts those 2SrcQB_QB_CreditMemo in the list box

        'FOR PART 3TestID_
        Debug.WriteLine("List3TestID_QBTable")

        'dim SQL strings
        Dim strSQL1, strSQL2, strSQL3, strSQL4, strSQL5, strSQL6, strTableInsert, strTableUpdate As String

        'Show what's processing
        ShowUserMessage(strSubName, "RefreshQB: Processing QB_CreditMemo", "RefreshQB: Processing QB_CreditMemo", True)

        'PART 2SrcQB_: Get the new records from Actual QB
        'Get a recordset of records from QB that are newer than QBTable
        Using rs2SrcQB_QB_CreditMemo As DataSet = New DataSet()

            str2SrcQB_QB_CreditMemoSQL = "SELECT * FROM CreditMemo WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_CreditMemo & "'}" ' ORDER BY TimeModified"

            Debug.WriteLine(str2SrcQB_QB_CreditMemoSQL)
            Using adap As New OdbcDataAdapter(str2SrcQB_QB_CreditMemoSQL, cnQuickBooks)
                rs2SrcQB_QB_CreditMemo.Tables.Clear()
                adap.Fill(rs2SrcQB_QB_CreditMemo)
            End Using

            Dim curRow As Integer = 0
            Dim rowCount As Integer = rs2SrcQB_QB_CreditMemo.Tables(0).Rows.Count

            If rowCount > 0 Then

                'Show what's processing in the listbox
                ShowUserMessage(strSubName, "Processing " & rowCount.ToString & " QB_CreditMemo Records")

                For Each iteration_row As DataRow In rs2SrcQB_QB_CreditMemo.Tables(0).Rows
                    curRow += 1
                    ShowUserMessage(strSubName, "Processing " & curRow.ToString & " of " & rowCount.ToString & " QB_CreditMemo Records")

                    Try

                        'get the columns from the database
                        str2SrcQB_TxnID = NCStr(iteration_row("TxnID")).Replace("'"c, "`"c)
                        str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                        str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                        str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                        str2SrcQB_TxnNumber = NCStr(iteration_row("TxnNumber")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefListID = NCStr(iteration_row("CustomerRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerRefFullName = NCStr(iteration_row("CustomerRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ClassRefListID = NCStr(iteration_row("ClassRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ClassRefFullName = NCStr(iteration_row("ClassRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefListID = NCStr(iteration_row("ARAccountRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ARAccountRefFullName = NCStr(iteration_row("ARAccountRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_TemplateRefListID = NCStr(iteration_row("TemplateRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_TemplateRefFullName = NCStr(iteration_row("TemplateRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDate = NCStr(iteration_row("TxnDate")).Replace("'"c, "`"c)
                        str2SrcQB_TxnDateMacro = NCStr(iteration_row("TxnDateMacro")).Replace("'"c, "`"c)
                        str2SrcQB_RefNumber = NCStr(iteration_row("RefNumber")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr1 = NCStr(iteration_row("BillAddressAddr1")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr2 = NCStr(iteration_row("BillAddressAddr2")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr3 = NCStr(iteration_row("BillAddressAddr3")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressAddr4 = NCStr(iteration_row("BillAddressAddr4")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressCity = NCStr(iteration_row("BillAddressCity")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressState = NCStr(iteration_row("BillAddressState")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressPostalCode = NCStr(iteration_row("BillAddressPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_BillAddressCountry = NCStr(iteration_row("BillAddressCountry")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr1 = NCStr(iteration_row("ShipAddressAddr1")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr2 = NCStr(iteration_row("ShipAddressAddr2")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr3 = NCStr(iteration_row("ShipAddressAddr3")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressAddr4 = NCStr(iteration_row("ShipAddressAddr4")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressCity = NCStr(iteration_row("ShipAddressCity")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressState = NCStr(iteration_row("ShipAddressState")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressPostalCode = NCStr(iteration_row("ShipAddressPostalCode")).Replace("'"c, "`"c)
                        str2SrcQB_ShipAddressCountry = NCStr(iteration_row("ShipAddressCountry")).Replace("'"c, "`"c)
                        str2SrcQB_IsPending = NCStr(iteration_row("IsPending")).Replace("'"c, "`"c)
                        str2SrcQB_PONumber = NCStr(iteration_row("PONumber")).Replace("'"c, "`"c)
                        str2SrcQB_TermsRefListID = NCStr(iteration_row("TermsRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_TermsRefFullName = NCStr(iteration_row("TermsRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_DueDate = NCStr(iteration_row("DueDate")).Replace("'"c, "`"c)
                        str2SrcQB_SalesRepRefListID = NCStr(iteration_row("SalesRepRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_SalesRepRefFullName = NCStr(iteration_row("SalesRepRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_FOB = NCStr(iteration_row("FOB")).Replace("'"c, "`"c)
                        str2SrcQB_ShipDate = NCStr(iteration_row("ShipDate")).Replace("'"c, "`"c)
                        str2SrcQB_ShipMethodRefListID = NCStr(iteration_row("ShipMethodRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ShipMethodRefFullName = NCStr(iteration_row("ShipMethodRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_Subtotal = NCStr(iteration_row("Subtotal")).Replace("'"c, "`"c)
                        str2SrcQB_ItemSalesTaxRefListID = NCStr(iteration_row("ItemSalesTaxRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_ItemSalesTaxRefFullName = NCStr(iteration_row("ItemSalesTaxRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_SalesTaxPercentage = NCStr(iteration_row("SalesTaxPercentage")).Replace("'"c, "`"c)
                        str2SrcQB_SalesTaxTotal = NCStr(iteration_row("SalesTaxTotal")).Replace("'"c, "`"c)
                        str2SrcQB_TotalAmount = NCStr(iteration_row("TotalAmount")).Replace("'"c, "`"c)
                        str2SrcQB_CreditRemaining = NCStr(iteration_row("CreditRemaining")).Replace("'"c, "`"c)
                        str2SrcQB_Memo = NCStr(iteration_row("Memo")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerMsgRefListID = NCStr(iteration_row("CustomerMsgRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerMsgRefFullName = NCStr(iteration_row("CustomerMsgRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_IsToBePrinted = NCStr(iteration_row("IsToBePrinted")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerSalesTaxCodeRefListID = NCStr(iteration_row("CustomerSalesTaxCodeRefListID")).Replace("'"c, "`"c)
                        str2SrcQB_CustomerSalesTaxCodeRefFullName = NCStr(iteration_row("CustomerSalesTaxCodeRefFullName")).Replace("'"c, "`"c)
                        str2SrcQB_CustomFieldOther = NCStr(iteration_row("CustomFieldOther")).Replace("'"c, "`"c)


                        str2SrcQB_IsPending = IIf(str2SrcQB_IsPending = "True", "1", "0")
                        str2SrcQB_IsToBePrinted = IIf(str2SrcQB_IsToBePrinted = "True", "1", "0")


                        'Put the information together into a string
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
                        ShowUserMessage(strSubName, str2SrcQB_QB_CreditMemoRow)

                        UpdateQBCustomerBalance(str2SrcQB_CustomerRefListID)
                        'UpdateQBInvoice(str2SrcQB_CustomerRefFullName)

                        Dim iRowCount As Integer = 0
                        iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(TxnID) FROM QB_CreditMemo WHERE TxnID = '" & str2SrcQB_TxnID & "'")
                        'If iRowCount > 1 Then Stop 'Should only be one
                        If iRowCount = 1 Then 'record exists  -UPDATE


                            'DO UPDATE WORK:
                            Debug.WriteLine("UPDATE")

                            'Build the SQL string
                            strSQL1 = "UPDATE  " & Environment.NewLine & _
                                      "       QB_CreditMemo " & Environment.NewLine & _
                                      "SET " & Environment.NewLine & _
                                      "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine & _
                                      "     , TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                      "     , TxnNumber = " & str2SrcQB_TxnNumber & "" & Environment.NewLine & _
                                      "     , CustomerRefListID = '" & str2SrcQB_CustomerRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerRefFullName = '" & str2SrcQB_CustomerRefFullName & "'" & Environment.NewLine & _
                                      "     , ClassRefListID = '" & str2SrcQB_ClassRefListID & "'" & Environment.NewLine & _
                                      "     , ClassRefFullName = '" & str2SrcQB_ClassRefFullName & "'" & Environment.NewLine & _
                                      "     , ARAccountRefListID = '" & str2SrcQB_ARAccountRefListID & "'" & Environment.NewLine & _
                                      "     , ARAccountRefFullName = '" & str2SrcQB_ARAccountRefFullName & "'" & Environment.NewLine & _
                                      "     , TemplateRefListID = '" & str2SrcQB_TemplateRefListID & "'" & Environment.NewLine & _
                                      "     , TemplateRefFullName = '" & str2SrcQB_TemplateRefFullName & "'" & Environment.NewLine & _
                                      "     , TxnDate = '" & str2SrcQB_TxnDate & "'" & Environment.NewLine & _
                                      "     , TxnDateMacro = '" & str2SrcQB_TxnDateMacro & "'" & Environment.NewLine & _
                                      "     , RefNumber = '" & str2SrcQB_RefNumber & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr1 = '" & str2SrcQB_BillAddressAddr1 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr2 = '" & str2SrcQB_BillAddressAddr2 & "'" & Environment.NewLine
                            strSQL2 = "     , BillAddressAddr3 = '" & str2SrcQB_BillAddressAddr3 & "'" & Environment.NewLine & _
                                      "     , BillAddressAddr4 = '" & str2SrcQB_BillAddressAddr4 & "'" & Environment.NewLine & _
                                      "     , BillAddressCity = '" & str2SrcQB_BillAddressCity & "'" & Environment.NewLine & _
                                      "     , BillAddressState = '" & str2SrcQB_BillAddressState & "'" & Environment.NewLine & _
                                      "     , BillAddressPostalCode = '" & str2SrcQB_BillAddressPostalCode & "'" & Environment.NewLine & _
                                      "     , BillAddressCountry = '" & str2SrcQB_BillAddressCountry & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr1 = '" & str2SrcQB_ShipAddressAddr1 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr2 = '" & str2SrcQB_ShipAddressAddr2 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr3 = '" & str2SrcQB_ShipAddressAddr3 & "'" & Environment.NewLine & _
                                      "     , ShipAddressAddr4 = '" & str2SrcQB_ShipAddressAddr4 & "'" & Environment.NewLine & _
                                      "     , ShipAddressCity = '" & str2SrcQB_ShipAddressCity & "'" & Environment.NewLine & _
                                      "     , ShipAddressState = '" & str2SrcQB_ShipAddressState & "'" & Environment.NewLine & _
                                      "     , ShipAddressPostalCode = '" & str2SrcQB_ShipAddressPostalCode & "'" & Environment.NewLine & _
                                      "     , ShipAddressCountry = '" & str2SrcQB_ShipAddressCountry & "'" & Environment.NewLine & _
                                      "     , IsPending = '" & str2SrcQB_IsPending & "'" & Environment.NewLine & _
                                      "     , PONumber = '" & str2SrcQB_PONumber & "'" & Environment.NewLine & _
                                      "     , TermsRefListID = '" & str2SrcQB_TermsRefListID & "'" & Environment.NewLine & _
                                      "     , TermsRefFullName = '" & str2SrcQB_TermsRefFullName & "'" & Environment.NewLine & _
                                      "     , DueDate = '" & str2SrcQB_DueDate & "'" & Environment.NewLine & _
                                      "     , SalesRepRefListID = '" & str2SrcQB_SalesRepRefListID & "'" & Environment.NewLine & _
                                      "     , SalesRepRefFullName = '" & str2SrcQB_SalesRepRefFullName & "'" & Environment.NewLine
                            strSQL3 = "     , FOB = '" & str2SrcQB_FOB & "'" & Environment.NewLine & _
                                      "     , ShipDate = '" & str2SrcQB_ShipDate & "'" & Environment.NewLine & _
                                      "     , ShipMethodRefListID = '" & str2SrcQB_ShipMethodRefListID & "'" & Environment.NewLine & _
                                      "     , ShipMethodRefFullName = '" & str2SrcQB_ShipMethodRefFullName & "'" & Environment.NewLine & _
                                      "     , Subtotal = " & str2SrcQB_Subtotal & "" & Environment.NewLine & _
                                      "     , ItemSalesTaxRefListID = '" & str2SrcQB_ItemSalesTaxRefListID & "'" & Environment.NewLine & _
                                      "     , ItemSalesTaxRefFullName = '" & str2SrcQB_ItemSalesTaxRefFullName & "'" & Environment.NewLine & _
                                      "     , SalesTaxPercentage = " & str2SrcQB_SalesTaxPercentage & "" & Environment.NewLine & _
                                      "     , SalesTaxTotal = " & str2SrcQB_SalesTaxTotal & "" & Environment.NewLine & _
                                      "     , TotalAmount = " & str2SrcQB_TotalAmount & "" & Environment.NewLine & _
                                      "     , CreditRemaining = " & str2SrcQB_CreditRemaining & "" & Environment.NewLine & _
                                      "     , Memo = '" & str2SrcQB_Memo & "'" & Environment.NewLine & _
                                      "     , CustomerMsgRefListID = '" & str2SrcQB_CustomerMsgRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerMsgRefFullName = '" & str2SrcQB_CustomerMsgRefFullName & "'" & Environment.NewLine & _
                                      "     , IsToBePrinted = '" & str2SrcQB_IsToBePrinted & "'" & Environment.NewLine & _
                                      "     , CustomerSalesTaxCodeRefListID = '" & str2SrcQB_CustomerSalesTaxCodeRefListID & "'" & Environment.NewLine & _
                                      "     , CustomerSalesTaxCodeRefFullName = '" & str2SrcQB_CustomerSalesTaxCodeRefFullName & "'" & Environment.NewLine & _
                                      "     , CustomFieldOther = '" & str2SrcQB_CustomFieldOther & "'" & Environment.NewLine & _
                                      "WHERE " & Environment.NewLine & _
                                      "       TxnID = '" & str2SrcQB_TxnID & "'" & Environment.NewLine



                            'Combine the strings
                            strTableUpdate = strSQL1 & strSQL2 & strSQL3
                            SQLHelper.ExecuteSQL(cnMax, strTableUpdate)

                        Else
                            'record not exist  -INSERT
                            'DO INSERT WORK:
                            Debug.WriteLine("INSERT")

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
                            SQLHelper.ExecuteSQL(cnMax, strTableInsert)
                        End If
                    Catch ex As Exception
                        HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                        Continue For
                    End Try

                Next iteration_row


            End If
        End Using

        ShowUserMessage(strSubName, "RefreshQB: Finished Processing QB_CreditMemo", "RefreshQB: Finished Processing QB_CreditMemo", True)

    End Sub


    Public Sub RefreshQBTablesOnceDaily()

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQBTablesOnceDaily" '"SUBNAME"

        'Check permission to run
        If Not HavePermission(strObjName, strSubName) Then Exit Sub

        If frmMain.chkPauseProcessing.CheckState = CheckState.Checked Then Exit Sub

        'Open the QuickBooks file
        If Not booQBFileIsOpen Then
            If Not (cnQuickBooks.State = ConnectionState.Open) Then
                OpenConnectionQB()
            Else
                booQBFileIsOpen = True
            End If
        End If

        'Set flag
        booQBRefreshInProgress = True

        ShowUserMessage(strSubName, "Processing RefreshQBTables: Once Daily", "Processing RefreshQBTables: Once Daily", True)

        GetQBMaxTimeModified()

        RefreshQB_Customer()
        InsertMaxBillToIntoQB()

        RefreshQB_Invoice()
        RefreshQB_InvoiceLine()

        If gstrComputerName <> "EDSDELLXPP" Then
            RefreshQB_ReceivePayment()
            RefreshQB_ReceivePaymentLine()
            RefreshQB_CreditMemo()
        End If

        Dim TempCommand As SqlCommand
        TempCommand = cnMax.CreateCommand()
        TempCommand.CommandText = "sp_TEMP_MarkCustPromoRush"
        TempCommand.ExecuteNonQuery()


        ShowUserMessage(strSubName, "Finished RefreshQBTables: OnceDaily", "Finished RefreshQBTables: OnceDaily", True)

        'Reset flag
        booQBRefreshInProgress = False


    End Sub



    Public Sub RefreshQB_Terms()

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modDBPM_RefreshQBTables" '"OBJNAME"
        Dim strSubName As String = "RefreshQB_Terms" '"SUBNAME"

        Dim str2SrcQB_QB_TermsSQL, str2SrcQB_QB_TermsRow, str2SrcQB_ListID, str2SrcQB_TimeCreated, str2SrcQB_TimeModified, str2SrcQB_EditSequence, str2SrcQB_Name, str2SrcQB_IsActive, str2SrcQB_DayOfMonthDue, str2SrcQB_DueNextMonthDays, str2SrcQB_DiscountDayOfMonth, str2SrcQB_DiscountPct, str2SrcQB_StdDueDays, str2SrcQB_StdDiscountDays, str2SrcQB_StdDiscountPct, str2SrcQB_Type As String

        Dim strSQL1, strSQL2, strTableInsert, strTableUpdate As String

        ShowUserMessage(strSubName, "RefreshQB: Processing QB_Terms", "RefreshQB: Processing QB_Terms", True)

        Using rs2SrcQB_QB_Terms As DataSet = New DataSet()
            str2SrcQB_QB_TermsSQL = "SELECT * FROM Terms WHERE TimeModified > {ts '" & gstrQBMaxTimeModified_Terms & "'}" ' ORDER BY TimeModified"

            Try
                Using adap_3 As New OdbcDataAdapter(str2SrcQB_QB_TermsSQL, cnQuickBooks)
                    rs2SrcQB_QB_Terms.Tables.Clear()
                    adap_3.Fill(rs2SrcQB_QB_Terms) ', adAsyncFetch '(no Optimizer)
                End Using
            Catch ex As Exception
                HaveError(strObjName, strSubName, CStr(Information.Err().Number), ex.Message, Information.Err().Source, "", "")
                Exit Sub
            End Try

            Dim curRow As Integer = 0
            Dim rowCount As Integer = rs2SrcQB_QB_Terms.Tables(0).Rows.Count

            If rowCount > 0 Then

                'Show what's processing in the listbox
                ShowUserMessage(strSubName, "Processing " & rowCount.ToString & " QB_Terms Records")

                For Each iteration_row As DataRow In rs2SrcQB_QB_Terms.Tables(0).Rows
                    Try

                        curRow += 1
                        ShowUserMessage(strSubName, "Processing " & curRow.ToString & " of " & rowCount.ToString & " QB_Terms Records")

                        'get the columns from the database
                        str2SrcQB_ListID = NCStr(iteration_row("ListID")).Replace("'"c, "`"c)
                        str2SrcQB_TimeCreated = NCStr(iteration_row("TimeCreated")).Replace("'"c, "`"c)
                        str2SrcQB_TimeModified = NCStr(iteration_row("TimeModified")).Replace("'"c, "`"c)
                        str2SrcQB_EditSequence = NCStr(iteration_row("EditSequence")).Replace("'"c, "`"c)
                        str2SrcQB_Name = NCStr(iteration_row("Name")).Replace("'"c, "`"c)
                        str2SrcQB_IsActive = NCStr(iteration_row("IsActive"), 1).Replace("'"c, "`"c)
                        str2SrcQB_DayOfMonthDue = NCStr(iteration_row("DayOfMonthDue")).Replace("'"c, "`"c)
                        str2SrcQB_DueNextMonthDays = NCStr(iteration_row("DueNextMonthDays")).Replace("'"c, "`"c)
                        str2SrcQB_DiscountDayOfMonth = NCStr(iteration_row("DiscountDayOfMonth")).Replace("'"c, "`"c)
                        str2SrcQB_DiscountPct = NCDbl(iteration_row("DiscountPct"))
                        str2SrcQB_StdDueDays = NCStr(iteration_row("StdDueDays")).Replace("'"c, "`"c)
                        str2SrcQB_StdDiscountDays = NCStr(iteration_row("StdDiscountDays")).Replace("'"c, "`"c)
                        str2SrcQB_StdDiscountPct = NCStr(iteration_row("StdDiscountPct")).Replace("'"c, "`"c)
                        str2SrcQB_Type = NCStr(iteration_row("Type")).Replace("'"c, "`"c)

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
                        ShowUserMessage(strSubName, str2SrcQB_QB_TermsRow)

                        'Check to see if ListID is in QB_Terms
                        'Yes then UPDATE record
                        'No then INSERT record
                        '"SELECT ListID FROM QB_Terms WHERE ListID = '" & str2SrcQB_ListID & "'"
                        Dim iRowCount As Integer = 0
                        iRowCount = SQLHelper.ExecuteScalerInt(cnMax, CommandType.Text, "SELECT Count(ListID) FROM QB_Terms WHERE ListID = '" & str2SrcQB_ListID & "'")
                        If iRowCount > 0 Then 'record exists  -UPDATE

                            'DO UPDATE WORK:
                            Debug.WriteLine("UPDATE")

                            'Build the SQL string
                            strSQL1 = "UPDATE  " & Environment.NewLine & _
                                      "       QB_Terms " & Environment.NewLine & _
                                      "SET " & Environment.NewLine & _
                                      "       TimeCreated = '" & str2SrcQB_TimeCreated & "'" & Environment.NewLine & _
                                      "     , TimeModified = '" & str2SrcQB_TimeModified & "'" & Environment.NewLine & _
                                      "     , EditSequence = '" & str2SrcQB_EditSequence & "'" & Environment.NewLine & _
                                      "     , Name = '" & str2SrcQB_Name & "'" & Environment.NewLine & _
                                      "     , IsActive = '" & str2SrcQB_IsActive & "'" & Environment.NewLine & _
                                      "     , DayOfMonthDue = '" & str2SrcQB_DayOfMonthDue & "'" & Environment.NewLine & _
                                      "     , DueNextMonthDays = '" & str2SrcQB_DueNextMonthDays & "'" & Environment.NewLine & _
                                      "     , DiscountDayOfMonth = '" & str2SrcQB_DiscountDayOfMonth & "'" & Environment.NewLine & _
                                      "     , DiscountPct = '" & str2SrcQB_DiscountPct & "'" & Environment.NewLine & _
                                      "     , StdDueDays = '" & str2SrcQB_StdDueDays & "'" & Environment.NewLine & _
                                      "     , StdDiscountDays = '" & str2SrcQB_StdDiscountDays & "'" & Environment.NewLine & _
                                      "     , StdDiscountPct = '" & str2SrcQB_StdDiscountPct & "'" & Environment.NewLine & _
                                      "     , Type = '" & str2SrcQB_Type & "'" & Environment.NewLine & _
                                      "WHERE " & Environment.NewLine & _
                                      "       ListID = '" & str2SrcQB_ListID & "'" & Environment.NewLine

                            strTableUpdate = strSQL1
                            SQLHelper.ExecuteSQL(cnMax, strTableUpdate)

                        Else
                            'record not exist  -INSERT
                            Debug.WriteLine("INSERT")

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
                                      "   , '" & str2SrcQB_Type & "' )  --Type" & Environment.NewLine


                            'Combine the strings
                            strTableInsert = strSQL1 & strSQL2
                            SQLHelper.ExecuteSQL(cnMax, strTableInsert)

                        End If
                    Catch ex As Exception
                        HaveError(strObjName, strSubName, Information.Err.Number.ToString, ex.Message, Information.Err.Source, "", "", ex)
                        Continue For
                    End Try

                Next iteration_row


            End If
            ShowUserMessage(strSubName, "Finished Processing Refresh QB_Terms", "Finished Processing Refresh QB_Terms", True)
        End Using
    End Sub
   
End Module