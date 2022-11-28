Option Strict Off
Option Explicit On
Module DIA_Common
	
    Public db As DAO.Database
	Public rs1, rs, rs2 As dao.Recordset
	Public sHoldDB As String
	Public sDBDir As String
	
	Declare Function GetUserName Lib "advapi32.dll"  Alias "GetUserNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
End Module