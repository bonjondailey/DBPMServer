Imports System
Imports System.Runtime.InteropServices
Namespace UnsafeNative
	<System.Security.SuppressUnmanagedCodeSecurity> _
	 Public Module kernel32
		Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
		Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Integer, ByRef lpExitCode As Integer) As Integer
		Declare Function GetModuleFileName Lib "kernel32"  Alias "GetModuleFileNameA"(ByVal hModule As Integer, ByVal lpFileName As String, ByVal nSize As Integer) As Integer
		Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
		'UPGRADE_TODO: (1050) Structure STARTUPINFO may require marshalling attributes to be passed as an argument in this Declare statement. More Information: http://www.vbtonet.com/ewis/ewi1050.aspx
		Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Integer, ByVal lpThreadAttributes As Integer, ByVal bInheritHandles As Integer, ByVal dwCreationFlags As Integer, ByVal lpEnvironment As Integer, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As DBPM_Server_4_28_2015Support.UnsafeNative.Structures.STARTUPINFO, ByRef lpProcessInformation As DBPM_Server_4_28_2015Support.UnsafeNative.Structures.PROCESS_INFORMATION) As Integer
		Declare Function GetComputerName Lib "kernel32.dll"  Alias "GetComputerNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
	End Module
End Namespace