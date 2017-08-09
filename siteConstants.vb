Imports Microsoft.VisualBasic.Interaction

Public Class siteConstants
    Public Shared Function SQLConnectionString() As String

        Dim myConnString As String

        Dim DataSource As String = GetDBServer()
        Dim InitialCatalog As String = GetDBName()
        Dim UserID As String = GetDBUser()
        Dim Password As String = GetDBPass()

        myConnString = "Data Source=" & DataSource & ";" & _
                        "Initial Catalog=" & InitialCatalog & ";" & _
                        "User ID=" & UserID & ";" & _
                        "Password=" & Password

        Return myConnString

    End Function

    Public Shared mailHost As String = "smtp.office365.com"
    Public Shared mailPort As String = "587"
    Public Shared mailErrorSendTo As String = GetErrorSendToEmail()
    Public Shared mailErrorSendFrom As String = "jt@drum-line.com"
    Public Shared mailPassword As String = "dbpm_2114"


    Public Shared Function GetDBServer() As String
        Select Case Environment.MachineName
            Case "SERVER06"
                Return "SERVER06"
            Case "DLDBPM"
                Return "DLDBPM\MSSQLSERVER2012"
            Case Else
                Return "DEVSERVER"
        End Select
    End Function

    Public Shared Function GetErrorSendToEmail() As String
        Select Case Environment.MachineName
            Case "SERVER06", "DLDBPM"
                Return "abarnes@drum-line.com"
            Case Else
                Return "jdbailey@gmail.com"
        End Select
    End Function

    Private Shared Function GetDBName() As String
        Return "DrummondPrinting"
    End Function

    Private Shared Function GetDBUser() As String
        Return "DBPMUsers"
    End Function

    Private Shared Function GetDBPass() As String
        Return "thisthingisgreat"
    End Function


    Public Shared ReadOnly ACCEPTED_FILES As String = "|.mpg|.mov|.wmv|.mp3|.m4p|.eps|.pdf|.doc|.xls|.txt|.zip|.sit|.wpd|.rtf|.ppt|.jpg|.jpeg|.tif|.tiff|"

    Public Shared Function showSize(ByVal fileSize As Double) As String
        If getKB(fileSize) > 1000 Then
            Return CType(getMB(fileSize), Integer).ToString & " MB"
        Else
            Return CType(getKB(fileSize), Integer).ToString & " KB"
        End If
    End Function

    Public Shared Function getMB(ByVal fileSize As Double) As Double
        Return fileSize / 1024000
    End Function

    Public Shared Function getKB(ByVal fileSize As Double) As Double
        Return fileSize / 1024
    End Function

    Public Shared Function getSecurityQuestion(ByVal item As Integer) As String
        Dim secList As ArrayList
        secList = getSecurityQuestionList()
        Return secList(item).ToString
    End Function

    Public Shared Function getSecurityQuestionList() As ArrayList
        Dim secList As New ArrayList
        secList.Add("What is your Father's middle name?")
        secList.Add("What is your Mother's maiden name?")
        secList.Add("What is your favorite pet's name?")
        secList.Add("In what city were you born?")
        Return secList

    End Function


    Public Shared Function padZero(ByVal s As String) As String
        If s.Length = 1 Then
            Return "0" & s
        Else
            Return s
        End If
    End Function


    Public Shared Function FormatQuery(ByVal s As String, Optional ByVal allowedChar As String = "") As Object
        If s.Trim = "" Then
            Return DBNull.Value
        ElseIf Not s Is DBNull.Value Then
            Dim badChars() As String
            badChars = {"%"c, "'"c, "+"c, "&"c, ","c, "="c, "!"c}
            For Each item As String In badChars
                If Not item.Equals(allowedChar) Then
                    s.Replace(item, "%")
                End If
            Next
            s = s.Replace("""", "%")
        End If
        Return s
    End Function

    Public Shared Function FormatSQL(ByVal s As Object) As Object
        If s Is Nothing Then
            Return DBNull.Value
        ElseIf s.Equals(DBNull.Value) Then
            Return DBNull.Value
        ElseIf s.ToString.Trim.Length = 0 Then
            Return DBNull.Value
        Else
            Return s
        End If
    End Function

    Public Shared Function FixDBNulls(ByVal s As Object, Optional ByVal sReturn As String = "") As String
        If s Is System.DBNull.Value OrElse s Is Nothing Then
            Return sReturn
        Else
            Return s.ToString
        End If
    End Function


    Public Shared Function getNullGuid() As Guid
        Return StringToGuid("00000000-0000-0000-0000-000000000000")
    End Function

    Public Shared Function FixDBNullGuid(ByVal s As Object) As Object
        If s Is Nothing OrElse s Is DBNull.Value Then
            Return StringToGuid("00000000-0000-0000-0000-000000000000")
        Else
            Return s
        End If
    End Function

    Public Shared Function isNullGuid(ByVal g As Guid) As Boolean
        If g.ToString = "00000000-0000-0000-0000-000000000000" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function GuidToString(ByVal s As Guid) As String
        Return s.ToString
    End Function

    Public Shared Function StringToGuid(ByVal s As String) As Guid
        Dim myGUID As New Guid(s)
        Return myGUID
    End Function

    Public Shared Function StripGuid(ByVal s As Object) As String
        Dim strResult As String = ""
        If s.GetType.ToString = "System.Guid" Then
            strResult = GuidToString(CType(s, Guid))
        Else
            strResult = CStr(s)
        End If

        strResult = strResult.Replace("{", "_").Replace("}", "").Replace("-", "")
        Return strResult
    End Function

    Public Shared Function myCBool(ByVal s As String) As Boolean
        Select Case s
            Case "1"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Public Shared Function myCBoolOut(ByVal s As Boolean) As String
        Select Case s
            Case True
                Return "1"
            Case Else
                Return "0"
        End Select
    End Function

    Public Shared Function checkFileUpload(ByVal imageType As String, ByVal Extension As String) As Boolean
        If imageType = "UNKNOWN" AndAlso ACCEPTED_FILES.ToString.IndexOf("|" & Extension.ToLower & "|") = 0 Then
            Return False
        Else
            Return True
        End If
    End Function



#Region "Conversion Functions"
    Public Structure AdvancedTrimSettings
        Public MaxLength As Integer
        Public ShowElipsis As Boolean
        Public Elipsis As String
        Public ForceBreak As Boolean
        Public BreakCharacter As String

        ''' <summary>
        ''' Truncate the string at _MaxLength characters.  If the 
        ''' </summary>
        ''' <param name="_MaxLength"></param>
        ''' <param name="_ShowElipses"></param>
        ''' <param name="_ReplaceElipsis"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal _MaxLength As Integer, Optional ByVal _ShowElipses As Boolean = True, Optional ByVal _ReplaceElipsis As String = "...")
            MaxLength = _MaxLength
            ShowElipsis = _ShowElipses
            Elipsis = _ReplaceElipsis
            ForceBreak = False
            BreakCharacter = ""
        End Sub


        ''' <summary>
        ''' Force wrap at LineLength characres using the LineBreakString as a line break.
        ''' </summary>
        ''' <param name="WrapLength"></param>
        ''' <param name="LineBreakString"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal WrapLength As Integer, ByVal LineBreakString As String)
            MaxLength = WrapLength
            ShowElipsis = False
            Elipsis = ""
            ForceBreak = True
            BreakCharacter = LineBreakString
        End Sub

    End Structure

    Public Shared Function NCStr(ByVal Value As Object, Optional ByVal DefaultValue As String = "") As String
        Dim Results As String = DefaultValue
        Try
            If Value IsNot Nothing AndAlso Not (TypeOf (Value) Is System.DBNull) Then
                Results = Value.ToString()
            End If
        Catch ex As Exception
            Results = DefaultValue
        End Try

        Return Results
    End Function

    Public Shared Function NCStr(ByVal Value As Object, ByVal TrimSettings As AdvancedTrimSettings, Optional ByVal DefaultValue As String = "") As String
        Dim Results As String = NCStr(Value, DefaultValue)

        If Results.Length > TrimSettings.MaxLength Then
            'Determine if the string should be truncated or wrapped
            If TrimSettings.ForceBreak Then
                'Wrap the string
                Dim sb As New System.Text.StringBuilder
                Dim StartPos As Integer = 0
                Dim LineLength As Integer = 0

                While StartPos < Results.Length
                    LineLength = NCInt(IIf(StartPos + TrimSettings.MaxLength > Results.Length, Results.Length - StartPos, TrimSettings.MaxLength))

                    sb.Append(Results.Substring(StartPos, LineLength) & TrimSettings.BreakCharacter)

                    StartPos += LineLength
                End While

                Results = NCStr(sb.ToString())
            Else
                'Truncate the string
                Results = Results.Substring(0, TrimSettings.MaxLength - NCInt(IIf(TrimSettings.ShowElipsis, TrimSettings.Elipsis.Length, 0)))
                Results = Results.Remove(Results.LastIndexOf(" "))

                If TrimSettings.ShowElipsis Then
                    Results &= TrimSettings.Elipsis
                End If
            End If

        End If

        Return Results
    End Function

    Public Shared Function NCInt(ByVal Value As Object, Optional ByVal DefaultValue As Integer = 0) As Integer
        Dim Results As Integer = DefaultValue, intResult As Integer
        Try
            If Value IsNot Nothing AndAlso Integer.TryParse(Value.ToString, intResult) Then
                Results = intResult
            End If
        Catch ex As Exception
            Results = DefaultValue
        End Try

        Return Results
    End Function

    Public Shared Function NCDbl(ByVal Value As Object, Optional ByVal DefaultValue As Double = 0) As Double
        Dim Results As Double = DefaultValue
        Try
            If Value IsNot Nothing Then
                Results = CType(Value, Double)
            End If
        Catch ex As Exception
            Results = DefaultValue
        End Try

        Return Results
    End Function

    Public Shared Function NCBool(ByVal Value As Object, Optional ByVal DefaultValue As Boolean = False) As Boolean
        Dim Results As Boolean = DefaultValue, intResult As Integer
        Try
            If Value IsNot Nothing AndAlso Not (TypeOf (Value) Is System.DBNull) Then
                If TypeOf (Value) Is System.String Then
                    'Check for a number
                    If Integer.TryParse(Value.ToString, intResult) Then
                        Results = (Not (intResult = 0))
                    Else
                        Results = Boolean.Parse(Value.ToString)
                    End If

                Else
                    Results = CType(Value, Boolean)
                End If
            End If
        Catch ex As Exception
            Results = DefaultValue
        End Try

        Return Results
    End Function

    Public Shared Function NCGuid(ByVal Value As Object) As Guid
        Dim Results As Guid = StringToGuid("00000000-0000-0000-0000-000000000000")

        Try
            If Value IsNot Nothing AndAlso Not (TypeOf (Value) Is System.DBNull) Then
                If TypeOf (Value) Is String Then
                    Results = StringToGuid(Value.ToString)
                Else
                    Results = CType(Value, Guid)
                End If
            End If
        Catch ex As Exception
            Results = StringToGuid("00000000-0000-0000-0000-000000000000")
        End Try

        Return Results
    End Function

    Public Shared Function NCDate(ByVal Value As Object) As Date
        Dim Results As Date = Date.Now

        Try
            If Value IsNot Nothing AndAlso Not (TypeOf (Value) Is System.DBNull) Then
                If TypeOf (Value) Is String Then
                    Results = Date.Parse(Value.ToString())
                Else
                    Results = CType(Value, Date)
                End If
            End If
        Catch ex As Exception
            Results = Date.Now
        End Try

        Return Results
    End Function


#End Region

    Public Shared Sub ShowUserMessage(ByVal functionName As String, ByVal sMessage As String, Optional ByVal TopWindowMessage As String = "", Optional ByVal showProcessing As Boolean = False)
        Dim myMsg As String = Date.Now.ToString & " : " & functionName & " -- " & sMessage
        Debug.WriteLine(myMsg)

        If showProcessing OrElse gShowProcessing Then
            frmMain.lstConversionProgress.Items.Add(myMsg)
            frmMain.lstConversionProgress.Items.Add("")
            frmMain.lstConversionProgress.SelectedIndex = frmMain.lstConversionProgress.Items.Count - 1
        End If

        If myMsg.Contains("process") Then frmMain.lblStatus.Text = myMsg

        If Not TopWindowMessage = "" Then
            frmMain.lblStatus.Text = TopWindowMessage
        End If

        If sMessage.Length > 300 Then
            frmMain.txtOutput.Text += Environment.NewLine & sMessage
        End If

        Application.DoEvents()
    End Sub



End Class
