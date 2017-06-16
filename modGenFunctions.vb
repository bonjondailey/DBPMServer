Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports DBPM_Server.siteConstants

Module modGenFunctions
    '**********************************
    '*** FIRST CODE REVIEW COMPLETE ***
    '**********************************
	
	'This routine gets the IOC_QB_ItemOtherCharge from the database according to the selection in gstrIOC_QB_ItemOtherChargeSQL.
	'It then puts those IOC_QB_ItemOtherCharge in the list box


    Public Function ItemOtherCharge_GetInfo(ByRef strItemName As String) As Hashtable

        'Permission and ErrorHandling          (Auto built)
        Dim strObjName As String = "modGenFunctions" '"OBJNAME"
        Dim strSubName As String = "ItemOtherCharge_GetInfo" '"SUBNAME"


        Dim gstrIOC_QB_ItemOtherChargeSQL As String = ""
        Dim gstrIOC_QB_ItemOtherChargeRow As String = ""
        Dim gstrIOC_ListID As String = ""
        Dim gstrIOC_Name As String = ""
        Dim gstrIOC_SalesTaxCodeRefListID As String = ""
        Dim gstrIOC_SalesTaxCodeRefFullName As String = ""
        Dim gstrIOC_SalesOrPurchaseDesc As String = ""
        Dim gstrIOC_SalesOrPurchaseAccountRefListID As String = ""
        Dim gstrIOC_SalesOrPurchaseAccountRefFullName As String = ""
        Dim hashItemOtherCharge As New Hashtable()

        Dim strSQL1 As String = "SELECT ListID , Name , SalesTaxCodeRefListID , SalesTaxCodeRefFullName , SalesOrPurchaseDesc , SalesOrPurchaseAccountRefListID , SalesOrPurchaseAccountRefFullName" & Environment.NewLine & _
                  "FROM  QB_ItemOtherCharge WHERE Name = '" & strItemName & "'"

        Using Sql As New SQLHelper(gstrSQLConnectionString)

            ShowUserMessage(strSubName, "Processing QB_ItemOtherCharge Information", "ItemOtherCharge", True)
            Using rsIOC_QB_ItemOtherCharge As SqlDataReader = Sql.ExecuteReader(CommandType.Text, strSQL1)
                If rsIOC_QB_ItemOtherCharge.Read Then

                    hashItemOtherCharge("gstrIOC_ListID") = NCStr(rsIOC_QB_ItemOtherCharge("ListID")).Replace("'"c, "`"c).Replace(":"c, ";"c)
                    hashItemOtherCharge("gstrIOC_Name") = NCStr(rsIOC_QB_ItemOtherCharge("Name")).Replace("'"c, "`"c).Replace(":"c, ";"c)
                    hashItemOtherCharge("gstrIOC_SalesTaxCodeRefListID") = NCStr(rsIOC_QB_ItemOtherCharge("SalesTaxCodeRefListID")).Replace("'"c, "`"c).Replace(":"c, ";"c)
                    hashItemOtherCharge("gstrIOC_SalesTaxCodeRefFullName") = NCStr(rsIOC_QB_ItemOtherCharge("SalesTaxCodeRefFullName")).Replace("'"c, "`"c).Replace(":"c, ";"c)
                    hashItemOtherCharge("gstrIOC_SalesOrPurchaseDesc") = NCStr(rsIOC_QB_ItemOtherCharge("SalesOrPurchaseDesc")).Replace("'"c, "`"c).Replace(":"c, ";"c)
                    hashItemOtherCharge("gstrIOC_SalesOrPurchaseAccountRefListID") = NCStr(rsIOC_QB_ItemOtherCharge("SalesOrPurchaseAccountRefListID")).Replace("'"c, "`"c).Replace(":"c, ";"c)
                    hashItemOtherCharge("gstrIOC_SalesOrPurchaseAccountRefFullName") = NCStr(rsIOC_QB_ItemOtherCharge("SalesOrPurchaseAccountRefFullName")).Replace("'"c, "`"c).Replace(":"c, ";"c)

                    'Put the information together into a string
                    'strEmpName = Trim(strEmpLast) & ", " & Trim(strEmpFirst) '& " " & Trim(strEmpMI) & ".  " & Trim(strEmpSuffix)
                    gstrIOC_QB_ItemOtherChargeRow = "" & _
                                                    Strings.Left(gstrIOC_ListID & "                  ", 18) & "   " & _
                                                    Strings.Left(gstrIOC_Name & "                  ", 18) & "   " & _
                                                    Strings.Left(gstrIOC_SalesTaxCodeRefListID & "                  ", 18) & "   " & _
                                                    Strings.Left(gstrIOC_SalesTaxCodeRefFullName & "                  ", 18) & "   " & _
                                                    Strings.Left(gstrIOC_SalesOrPurchaseDesc & "                  ", 18) & "   " & _
                                                    Strings.Left(gstrIOC_SalesOrPurchaseAccountRefListID & "                  ", 18) & "   " & _
                                                    Strings.Left(gstrIOC_SalesOrPurchaseAccountRefFullName & "                  ", 18) & "   " & _
                                                    "" & Strings.Chr(9)

                    'put the line in the listbox
                    ShowUserMessage(strSubName, gstrIOC_QB_ItemOtherChargeRow)

                End If
            End Using
        End Using

        Return hashItemOtherCharge

    End Function
End Module