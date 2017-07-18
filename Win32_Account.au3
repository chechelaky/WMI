; #UDF# =========================================================================================================================
; Description ...: lista os usuários do Windows
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/Win32_Account.au3
; Source.........: https://msdn.microsoft.com/en-us/library/aa394061(v=vs.85).aspx
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

#include <String.au3>
ConsoleWrite(wmi_test() & @LF)

Func wmi_test()
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\cimv2")
	Local $arrName = StringSplit('Name', ',', 2)

	Local $colItems = $objWMIService.ExecQuery("SELECT Name FROM Win32_Account WHERE SIDType=1", "WQL", 0x10 + 0x20)
	Local $Result = ''
	If IsObj($colItems) Then
		ConsoleWrite('[Win32_Account]' & @LF)
		For $objItem In $colItems
			For $jj = 0 To UBound($arrName) - 1
				ConsoleWrite('[' & $arrName[$jj] & _StringRepeat('.', 30 - StringLen($arrName[$jj])) & ']=[' & Execute('$objItem.' & $arrName[$jj]) & ']' & @LF)
			Next
		Next
		Return StringMid($Result, 1, StringLen($Result) - 1)
	Else
		Return False
	EndIf
EndFunc   ;==>wmi_test
