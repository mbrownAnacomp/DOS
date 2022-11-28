Option Strict Off
Option Explicit On
Module Common
    Public sHoldVMDB, sHoldMLDB, sMLDBDir, sVMDBDir, sHoldODBC(3), sODBC(3) As String
    Public db As DAO.Database
	Public rsML, rsVM As dao.Recordset
    Public sHoldDRP, sDRPDir As String
    Public bWinAuth As Boolean = True
    Public bHoldWinAuth As Boolean
	Public fs As Scripting.FileSystemObject
    Public rsNOACheck As DAO.Recordset
    Public rs1, rs, rs2 As DAO.Recordset
    Public sHoldDB As String
    Public sDBDir As String
    Public sNOADir, sHoldNOA As String
    Public sPS50NOADir, sHoldPS50NOA As String
    Public sOtherNOADir, sHoldOtherNOA As String

	Declare Function GetUserName Lib "advapi32.dll"  Alias "GetUserNameA"(ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
End Module