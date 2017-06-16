Imports DBPM_Server.siteConstants
Imports System.Data.SqlClient



Public Class SQLHelper
    Inherits System.ComponentModel.Component

    Public conn As SqlConnection


#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()
        conn = New SqlConnection(SQLConnectionString)
        conn.Open()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal sqlConnString As String)
        MyBase.New()
        conn = New SqlConnection(sqlConnString)
        conn.Open()
        InitializeComponent()
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            conn.Dispose()
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region
    Function GetDataReader(ByVal sql As String) As SqlDataReader
        Dim command As New SqlCommand(sql, conn)
        Dim dr As SqlDataReader
        dr = command.ExecuteReader
        command.Dispose()
        Return dr
    End Function

    Function GetDataSet(ByVal sql As String, ByVal tableName As String) As DataSet
        Dim da As New SqlDataAdapter(sql, conn)
        Dim ds As New DataSet
        ds.DataSetName = "DataSet" & tableName
        da.Fill(ds, tableName)
        da.Dispose()
        Return ds
    End Function

    Function GetDataItem(ByVal sql As String) As String
        Dim strResult As String
        Try
            Using command As New SqlCommand(sql, conn)
                strResult = FixDBNulls(command.ExecuteScalar)
            End Using

        Catch ex As Exception
            strResult = ""
        End Try
        Return strResult
    End Function

    Function GetDataItemBoolean(ByVal sql As String) As Boolean
        Dim bResult As Boolean
        Using command As New SqlCommand(sql, conn)
            bResult = NCBool(command.ExecuteScalar)
        End Using
        Return bResult
    End Function

    Function getDataItemFloat(ByVal sql As String, Optional ByVal returnForNull As Integer = -9999) As Double
        Dim myResult As Double
        Try
            Using command As New SqlCommand(sql, conn)
                myResult = NCDbl(command.ExecuteScalar)
            End Using
        Catch ex As Exception
            myResult = returnForNull
        End Try
        Return myResult
    End Function

    Function GetDataItemInt(ByVal sql As String) As Integer
        Dim intResult As Integer
        Using command As New SqlCommand(sql, conn)
            Try
                intResult = NCInt(command.ExecuteScalar)
            Catch ex As System.Exception
                intResult = -9999
            End Try
        End Using

        Return intResult
    End Function

    Function getDataItemGuid(ByVal sql As String) As Guid
        Dim myGuid As Guid
        Using command As New SqlCommand(sql, conn)
            myGuid = NCGuid(command.ExecuteScalar)
        End Using
        Return myGuid
    End Function

    'Function getDataItemBlob(ByVal sql As String) As BitArray
    '    Dim myBlob As BitArray
    '    Using command As New SqlCommand(sql, conn)
    '        myBlob = (command.ExecuteScalar)
    '    End Using
    '    Return myBlob
    'End Function

    Sub updateRS(ByVal sql As String)
        Using command As New SqlCommand(sql, conn)
            command.ExecuteNonQuery()
        End Using
    End Sub

    Sub updateRS(ByVal sql As String, ByRef RecordsAffected As Integer)
        Dim command As New SqlCommand(sql, conn)
        RecordsAffected = command.ExecuteNonQuery()
        command.Dispose()
        command = Nothing
    End Sub

    Sub updateRSParameters(ByVal sql As String, ByVal cmdParameters() As SqlParameter)
        Dim command As New SqlCommand(sql, conn)
        attachParameters(command, cmdParameters)
        command.ExecuteNonQuery()
        command.Dispose()
        command = Nothing
    End Sub

    Sub updateRSParameters2(ByVal sql As String, ByVal foo As SqlParameterCollection)
        Dim command As New SqlCommand(sql, conn)
        attachParameters2(command, foo)
        command.ExecuteNonQuery()
        command.Dispose()
        command = Nothing
    End Sub

    Sub attachParameters2(ByVal command As SqlCommand, ByVal params As SqlParameterCollection)
        Dim p As SqlParameter
        For Each p In params
            command.Parameters.Add(p)
        Next
    End Sub


    Sub attachParameters(ByVal command As SqlCommand, ByVal cmdParamenters() As SqlParameter)
        If Not cmdParamenters Is Nothing Then
            Dim p As SqlParameter
            For Each p In cmdParamenters
                command.Parameters.Add(p)
            Next
        End If
    End Sub

#Region "ExecuteSQL"
    Public Sub ExecuteSQL(ByVal SQL As String)
        ExecuteSQL(conn, SQL)
    End Sub

    Public Sub ExecuteSQL(ByVal SQL As String, ByVal ParamArray Parameters() As SqlParameter)
        ExecuteSQL(conn, SQL, Parameters)
    End Sub

    Public Shared Sub ExecuteSQL(ByVal Conn As SqlConnection, ByVal SQL As String)
        Using cmd As New SqlCommand(SQL, Conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Shared Sub ExecuteSQL(ByVal Conn As SqlConnection, ByVal trans As SqlTransaction, ByVal SQL As String)
        Using cmd As New SqlCommand
            With cmd
                .Connection = Conn
                .Transaction = trans
                .CommandText = SQL
                .CommandType = CommandType.Text
                .ExecuteNonQuery()
            End With
        End Using
    End Sub

    Public Shared Sub ExecuteSQL(ByVal Conn As SqlConnection, ByVal SQL As String, ByVal ParamArray Parameters() As SqlParameter)
        Using cmd As New SqlCommand(SQL, Conn)
            With cmd
                .Parameters.AddRange(Parameters)
                .ExecuteNonQuery()
            End With
        End Using
    End Sub

    Public Shared Sub ExecuteSQL(ByVal Conn As SqlConnection, ByVal trans As SqlTransaction, ByVal SQL As String, ByVal ParamArray Parameters() As SqlParameter)
        Using cmd As New SqlCommand()
            With cmd
                .Connection = Conn
                .Transaction = trans
                .CommandText = SQL
                .CommandType = CommandType.Text
                .Parameters.AddRange(Parameters)
                .ExecuteNonQuery()
            End With
        End Using
    End Sub
    Public Shared Function ExecuteSQLReturnAffected(ByVal Conn As SqlConnection, ByVal Trans As SqlTransaction, ByVal SQL As String, ByVal ParamArray Parameters() As SqlParameter) As Integer
        Dim returnCount As Integer = 0
        Using cmd As New SqlCommand()
            With cmd
                .Connection = Conn
                .Transaction = Trans
                .CommandText = SQL
                .CommandType = CommandType.Text
                .Parameters.AddRange(Parameters)
                returnCount = .ExecuteNonQuery()
            End With
        End Using

        Return returnCount
    End Function

#End Region

#Region "ExecuteSP"
    Public Function ExecuteSP(ByVal StoredProcedure As String, ByVal ParamArray Parameters() As SqlParameter) As Integer
        Return ExecuteSP(conn, StoredProcedure, Parameters)
    End Function

    Public Shared Function ExecuteSP(ByVal Conn As SqlConnection, ByVal StoredProcedure As String, _
                                ByVal ParamArray Parameters() As SqlParameter) As Integer

        Dim ReturnParam As New SqlParameter("@RETURNVALUE", SqlDbType.Int)
        ReturnParam.Direction = ParameterDirection.ReturnValue

        Using cmd As New SqlCommand(StoredProcedure, Conn)
            With cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add(ReturnParam) 'Return Value Parameter
                .Parameters.AddRange(Parameters)
                .ExecuteNonQuery()

                Return NCInt(ReturnParam.Value)
            End With
        End Using
    End Function

    Public Shared Function ExecuteSP(ByVal Conn As SqlConnection, ByVal trans As SqlTransaction, ByVal StoredProcedure As String, _
                                ByVal ParamArray Parameters() As SqlParameter) As Integer

        Dim ReturnParam As New SqlParameter("@RETURNVALUE", SqlDbType.Int)
        ReturnParam.Direction = ParameterDirection.ReturnValue

        Using cmd As New SqlCommand
            With cmd
                .Connection = Conn
                .Transaction = trans
                .CommandText = StoredProcedure
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add(ReturnParam) 'Return Value Parameter
                .Parameters.AddRange(Parameters)
                .ExecuteNonQuery()

                Return NCInt(ReturnParam.Value)
            End With
        End Using
    End Function

#End Region

#Region "ExecuteReader"
    Public Function ExecuteReader(ByVal CommandType As System.Data.CommandType, ByVal CommandText As String, _
                                  ByVal ParamArray Parameters() As SqlParameter) As SqlDataReader
        Return ExecuteReader(conn, CommandType, CommandText, Parameters)
    End Function

    Public Shared Function ExecuteReader(ByVal Conn As SqlConnection, ByVal CommandType As System.Data.CommandType, _
                                         ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As SqlDataReader

        Using cmd As New SqlCommand(CommandText, Conn)
            With cmd
                .CommandType = CommandType
                .Parameters.AddRange(Parameters)
                Return .ExecuteReader()
            End With
        End Using
    End Function

    Public Shared Function ExecuteReader(ByVal Conn As SqlConnection, ByVal trans As SqlTransaction, ByVal CommandType As System.Data.CommandType, _
                                          ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As SqlDataReader

        Using cmd As New SqlCommand(CommandText, Conn)
            With cmd
                .Transaction = trans
                .CommandType = CommandType
                .Parameters.AddRange(Parameters)
                Return .ExecuteReader()
            End With
        End Using
    End Function

#End Region

#Region "ExecuteScaler Functions"

    Public Shared Function ExecuteScalerDate(ByVal Conn As SqlConnection, ByVal CommandType As System.Data.CommandType, _
                                             ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As String

        Dim dtResult As String
        Using cmd As New SqlCommand
            Try
                With cmd
                    .Connection = Conn
                    .CommandType = CommandType
                    .CommandText = CommandText
                    .Parameters.AddRange(Parameters)
                End With
                dtResult = cmd.ExecuteScalar.ToString
            Catch ex As System.Exception
                dtResult = ""
            End Try
        End Using

        Return dtResult
    End Function


    Public Shared Function ExecuteScalerDate(ByVal Conn As SqlConnection, ByVal CommandType As System.Data.CommandType, ByVal CommandText As String, _
                                             dateStringFormat As String, ByVal ParamArray Parameters() As SqlParameter) As String

        Dim dtResult As String
        Using cmd As New SqlCommand
            Try
                With cmd
                    .Connection = Conn
                    .CommandType = CommandType
                    .CommandText = CommandText
                    .Parameters.AddRange(Parameters)
                End With
                dtResult = CType(cmd.ExecuteScalar, Date).ToString(dateStringFormat)
            Catch ex As System.Exception
                dtResult = ""
            End Try
        End Using

        Return dtResult
    End Function

    Public Shared Function ExecuteScalerInt(ByVal Conn As SqlConnection, ByVal CommandType As System.Data.CommandType, _
                                            ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As Integer

        Dim intResult As Integer
        Using cmd As New SqlCommand
            Try
                With cmd
                    .Connection = Conn
                    .CommandType = CommandType
                    .CommandText = CommandText
                    .Parameters.AddRange(Parameters)
                End With
                intResult = CType(cmd.ExecuteScalar, Integer)
            Catch ex As System.Exception
                intResult = -9999
            End Try
        End Using

        Return intResult
    End Function

    Public Shared Function ExecuteScalerInt(ByVal Conn As SqlConnection, ByVal Trans As SqlTransaction, ByVal CommandType As System.Data.CommandType, _
                                            ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As Integer

        Dim intResult As Integer
        Using cmd As New SqlCommand
            Try
                With cmd
                    .Connection = Conn
                    .Transaction = Trans
                    .CommandType = CommandType
                    .CommandText = CommandText
                    .Parameters.AddRange(Parameters)
                End With
                intResult = CType(cmd.ExecuteScalar, Integer)
            Catch ex As System.Exception
                intResult = -9999
            End Try
        End Using

        Return intResult
    End Function

    Public Shared Function ExecuteScalerString(ByVal Conn As SqlConnection, ByVal CommandType As System.Data.CommandType, _
                                            ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As String

        Dim sResult As String
        Using cmd As New SqlCommand
            Try
                With cmd
                    .Connection = Conn
                    .CommandType = CommandType
                    .CommandText = CommandText
                    .Parameters.AddRange(Parameters)
                End With
                sResult = NCStr(cmd.ExecuteScalar)
            Catch ex As System.Exception
                sResult = ""
            End Try
        End Using

        Return sResult
    End Function

    Public Shared Function ExecuteScalerString(ByVal Conn As SqlConnection, ByVal trans As SqlTransaction, ByVal CommandType As System.Data.CommandType, _
                                            ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As String

        Dim sResult As String
        Using cmd As New SqlCommand
            Try
                With cmd
                    .Connection = Conn
                    .Transaction = trans
                    .CommandType = CommandType
                    .CommandText = CommandText
                    .Parameters.AddRange(Parameters)
                End With
                sResult = NCStr(cmd.ExecuteScalar)
            Catch ex As System.Exception
                sResult = ""
            End Try
        End Using

        Return sResult
    End Function


    Public Shared Function ExecuteScalerFloat(ByVal Conn As SqlConnection, ByVal CommandType As System.Data.CommandType, _
                                            ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As Double

        Dim dblResult As Double
        Using cmd As New SqlCommand
            Try
                With cmd
                    .Connection = Conn
                    .CommandType = CommandType
                    .CommandText = CommandText
                    .Parameters.AddRange(Parameters)
                End With
                dblResult = CType(cmd.ExecuteScalar, Double)
            Catch ex As System.Exception
                dblResult = -9999
            End Try
        End Using

        Return dblResult
    End Function

    Public Function ExecuteScaler(ByVal CommandType As System.Data.CommandType, ByVal CommandText As String, _
                                  ByVal ParamArray Parameters() As SqlParameter) As Object
        Return ExecuteScaler(conn, CommandType, CommandText, Parameters)
    End Function

    Public Shared Function ExecuteScaler(ByVal Conn As SqlConnection, ByVal CommandType As System.Data.CommandType, _
                                         ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As Object

        Using da As New SqlDataAdapter
            Using cmd As New SqlCommand(CommandText, Conn)
                With cmd
                    .CommandType = CommandType
                    .Parameters.AddRange(Parameters)
                    Return .ExecuteScalar()
                End With
            End Using
        End Using
    End Function


    Public Shared Function ExecuteScaler(ByVal Conn As SqlConnection, ByVal Trans As SqlTransaction, ByVal CommandType As System.Data.CommandType, _
                                         ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As Object

        Using da As New SqlDataAdapter
            Using cmd As New SqlCommand()
                With cmd
                    .Connection = Conn
                    .Transaction = Trans
                    .CommandText = CommandText
                    .CommandType = CommandType
                    .Parameters.AddRange(Parameters)
                    Return .ExecuteScalar()
                End With
            End Using
        End Using
    End Function

#End Region

#Region "ExecuteDataSet"
    Public Function ExecuteDataSet(ByVal CommandType As System.Data.CommandType, ByVal CommandText As String, ByVal dataSet As Data.DataSet, ByVal tableName As String, _
                                  ByVal ParamArray Parameters() As SqlParameter) As DataSet
        Return ExecuteDataSet(conn, CommandType, CommandText, dataSet, tableName, Parameters)
    End Function

    Public Shared Function ExecuteDataSet(ByVal Conn As SqlConnection, ByVal CommandType As System.Data.CommandType, _
                                         ByVal CommandText As String, ByVal dataSet As DataSet, ByVal tableName As String, ByVal ParamArray Parameters() As SqlParameter) As DataSet

        Using da As New SqlDataAdapter
            Using cmd As New SqlCommand(CommandText, Conn)
                With cmd
                    .CommandType = CommandType
                    .Parameters.AddRange(Parameters)
                End With

                da.SelectCommand = cmd
                da.Fill(dataSet, tableName)
            End Using
        End Using

        Return dataSet
    End Function
    Public Function ExecuteDataSet(ByVal CommandType As System.Data.CommandType, ByVal CommandText As String, _
                                  ByVal ParamArray Parameters() As SqlParameter) As DataSet
        Return ExecuteDataSet(conn, CommandType, CommandText, Parameters)
    End Function

    Public Shared Function ExecuteDataSet(ByVal Conn As SqlConnection, ByVal CommandType As System.Data.CommandType, _
                                         ByVal CommandText As String, ByVal ParamArray Parameters() As SqlParameter) As DataSet

        Dim ds As New DataSet

        Using da As New SqlDataAdapter
            Using cmd As New SqlCommand(CommandText, Conn)
                With cmd
                    .CommandType = CommandType
                    .Parameters.AddRange(Parameters)
                End With

                da.SelectCommand = cmd
                da.Fill(ds)
            End Using
        End Using

        Return ds
    End Function
#End Region

#Region "Debug Functions"
    ''' <summary>
    ''' Return a string representation of the object for use when debugging.
    ''' </summary>
    ''' <param name="cmd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DEBUG_ToString(ByVal cmd As SqlCommand) As String
        Dim strResults As String

        Try
            With cmd
                strResults = "CommandText:" & cmd.CommandText & Environment.NewLine
                strResults &= "CommandType:" & cmd.CommandType & Environment.NewLine

                For Each param As SqlParameter In .Parameters
                    strResults &= "(Parameter) " & param.ParameterName & ":" & NCStr(param.Value)
                Next

            End With
        Catch ex As Exception
            strResults = ex.Message
        End Try

        Return strResults
    End Function
#End Region

    Public Shared Function GetConnection() As SqlConnection
        Dim oConn As New SqlConnection(SQLConnectionString)

        oConn.Open()

        Return oConn
    End Function

    ''' <summary>
    ''' For use with drop down list that would be nice to have an empty first row
    ''' to be used as a default
    ''' </summary>
    ''' <param name="ds"></param>
    ''' <remarks></remarks>
    Public Shared Sub PadDatasetWithBlankFirstRow(ByRef ds As DataSet)
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Dim oRow As DataRow = ds.Tables(0).NewRow()
            ds.Tables(0).Rows.InsertAt(ds.Tables(0).NewRow(), 0)
        End If
    End Sub
End Class



