Imports System
Imports System.Runtime.InteropServices
Namespace SafeNative
	Public Module kernel32
		Public Function CloseHandle(ByVal hObject As Integer) As Integer
			Return DBPM_Server_4_28_2015Support.UnsafeNative.kernel32.CloseHandle(hObject)
		End Function
		Public Function CreateProcessA(ByRef lpApplicationName As String, ByRef lpCommandLine As String, ByVal lpProcessAttributes As Integer, ByVal lpThreadAttributes As Integer, ByVal bInheritHandles As Integer, ByVal dwCreationFlags As Integer, ByVal lpEnvironment As Integer, ByRef lpCurrentDirectory As String, ByRef lpStartupInfo As DBPM_Server_4_28_2015Support.UnsafeNative.Structures.STARTUPINFO, ByRef lpProcessInformation As DBPM_Server_4_28_2015Support.UnsafeNative.Structures.PROCESS_INFORMATION) As Integer
			Return DBPM_Server_4_28_2015Support.UnsafeNative.kernel32.CreateProcessA(lpApplicationName, lpCommandLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, lpStartupInfo, lpProcessInformation)
		End Function
		Public Function GetComputerName(ByRef lpBuffer As String, ByRef nSize As Integer) As Integer
			Return DBPM_Server_4_28_2015Support.UnsafeNative.kernel32.GetComputerName(lpBuffer, nSize)
		End Function
		Public Function GetExitCodeProcess(ByVal hProcess As Integer, ByRef lpExitCode As Integer) As Integer
			Return DBPM_Server_4_28_2015Support.UnsafeNative.kernel32.GetExitCodeProcess(hProcess, lpExitCode)
		End Function
		Public Function WaitForSingleObject(ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
			Return DBPM_Server_4_28_2015Support.UnsafeNative.kernel32.WaitForSingleObject(hHandle, dwMilliseconds)
		End Function
	End Module
End Namespace