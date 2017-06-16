Imports System
Imports System.Runtime.InteropServices
Namespace UnsafeNative
	<System.Security.SuppressUnmanagedCodeSecurity> _
	 Public Module Structures
		<StructLayout(LayoutKind.Sequential)> _
		 _
		Public Structure PROCESS_INFORMATION
			Dim hProcess As Integer
			Dim hThread As Integer
			Dim dwProcessID As Integer
			Dim dwThreadID As Integer
		End Structure
		<StructLayout(LayoutKind.Sequential)> _
		 _
		Public Structure STARTUPINFO
			Dim cb As Integer
			Dim lpReserved As String
			Dim lpDesktop As String
			Dim lpTitle As String
			Dim dwX As Integer
			Dim dwY As Integer
			Dim dwXSize As Integer
			Dim dwYSize As Integer
			Dim dwXCountChars As Integer
			Dim dwYCountChars As Integer
			Dim dwFillAttribute As Integer
			Dim dwFlags As Integer
			Dim wShowWindow As Short
			Dim cbReserved2 As Short
			Dim lpReserved2 As Integer
			Dim hStdInput As Integer
			Dim hStdOutput As Integer
			Dim hStdError As Integer
			Private Shared Sub InitStruct(ByRef result As STARTUPINFO, ByVal init As Boolean)
				If init Then
					result.lpReserved = String.Empty
					result.lpDesktop = String.Empty
					result.lpTitle = String.Empty
				End If
			End Sub
			Public Shared Function CreateInstance() As STARTUPINFO
				Dim result As New STARTUPINFO()
				InitStruct(result, True)
				Return result
			End Function
			Public Function Clone() As STARTUPINFO
				Dim result As STARTUPINFO = Me
				InitStruct(result, False)
				Return result
			End Function
		End Structure
	End Module
End Namespace