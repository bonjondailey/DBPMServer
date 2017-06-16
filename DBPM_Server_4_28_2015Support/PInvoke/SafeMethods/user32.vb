Imports System
Imports System.Runtime.InteropServices
Namespace SafeNative
	Public Module user32
		Public Function LockWorkStation() As Boolean
			Return DBPM_Server_4_28_2015Support.UnsafeNative.user32.LockWorkStation()
		End Function
	End Module
End Namespace