Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic
Imports System
Imports System.Diagnostics
Imports UpgradeHelpers.Gui
Imports UpgradeHelpers.Helpers
Partial Friend Class frmChooseCompany
	Inherits System.Windows.Forms.Form
	'SELECT eValue FROM MaConfig WHERE eKey = 'DB_NAME'
	'
	'
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly().EntryPoint <> Nothing AndAlso System.Reflection.Assembly.GetExecutingAssembly().EntryPoint.DeclaringType = Me.GetType() Then
						m_vb6FormDefInstance = Me
					End If

				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		ReLoadForm(False)
	End Sub



	Private Sub cmdOpenCompany_Click(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles cmdOpenCompany.Click

		Dim strObjName As String = "frmChooseCompany" '"OBJNAME"
		Dim strSubName As String = "cmdOpenCompany_Click" '"SUBNAME"

        Try
            OpenConnectionDBPM()
            OpenConnectionMax()

            frmMain.Text = "Main - " & gstrApplicationName & " - " & gstrUserName & " - " & gstrCompany
            frmMain.Show()
            Me.Close()
            
        Catch exc As System.Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), exc.Message, Information.Err().Source, "", "")
        End Try


	End Sub


	Private Sub frmChooseCompany_Load(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles MyBase.Load

		Dim strObjName As String = "frmChooseCompany" '"OBJNAME"
		Dim strSubName As String = "Form_Load" '"SUBNAME"

		Try 
			'TEMP BYPASS
				gstrCompany = "DrummondPrinting"
				cmdOpenCompany_Click(cmdOpenCompany, New EventArgs())

		Catch exc As System.Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), exc.Message, Information.Err().Source, "", "")
		End Try


	End Sub

	Private Sub lstChooseCompany_SelectedIndexChanged(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles lstChooseCompany.SelectedIndexChanged
		Dim booFillingListbox As Boolean

		Dim strObjName As String = "frmChooseCompany" '"OBJNAME"
		Dim strSubName As String = "lstChooseCompany_Click" '"SUBNAME"

		If Not HavePermission(strObjName, strSubName) Then Exit Sub

		cmdOpenCompany.Enabled = True

        If booFillingListbox Then Exit Sub

		For lngA As Integer = 0 To lstChooseCompany.Items.Count - 1
			If ListBoxHelper.GetSelected(lstChooseCompany, lngA) Then
				gstrCompany = lstChooseCompany.Text.Trim()
				Debug.WriteLine(lstChooseCompany.Text)
				Debug.WriteLine(gstrCompany)
			End If
		Next lngA

	End Sub

	Private Sub lstChooseCompany_DoubleClick(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles lstChooseCompany.DoubleClick

		Dim strObjName As String = "frmChooseCompany" '"OBJNAME"
		Dim strSubName As String = "lstChooseCompany_DblClick" '"SUBNAME"

		If Not HavePermission(strObjName, strSubName) Then Exit Sub

        Try

        Catch ex As Exception
            HaveError(strObjName, strSubName, CStr(Information.Err().Number), Information.Err().Description, Information.Err().Source, "", "")


        End Try

        lstChooseCompany_SelectedIndexChanged(lstChooseCompany, New EventArgs())
		cmdOpenCompany_Click(cmdOpenCompany, New EventArgs())

	End Sub
	Private Sub frmChooseCompany_Closed(ByVal eventSender As Object, ByVal eventArgs As EventArgs) Handles MyBase.Closed
	End Sub
End Class