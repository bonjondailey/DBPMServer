Imports System
Imports System.Runtime.InteropServices
Namespace SafeNative
	Public Module advapi32
		Public Function GetUserName(ByRef lpBuffer As String, ByRef nSize As Integer) As Integer
			Return DBPM_Server_4_28_2015Support.UnsafeNative.advapi32.GetUserName(lpBuffer, nSize)
		End Function
	End Module
End Namespace