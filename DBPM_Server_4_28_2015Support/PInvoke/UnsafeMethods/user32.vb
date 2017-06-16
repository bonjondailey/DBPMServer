Imports System
Imports System.Runtime.InteropServices
Namespace UnsafeNative
	<System.Security.SuppressUnmanagedCodeSecurity> _
	 Public Module user32
		Public Declare Function LockWorkStation Lib "user32" () As Boolean
	End Module
End Namespace