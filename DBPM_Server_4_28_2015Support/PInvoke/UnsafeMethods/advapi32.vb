Imports System
Imports System.Runtime.InteropServices
Namespace UnsafeNative
	<System.Security.SuppressUnmanagedCodeSecurity> _
	 Public Module advapi32
		Public Declare Function GetUserName Lib "advapi32.dll"  Alias "GetUserNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	End Module
End Namespace