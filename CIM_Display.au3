; #UDF# =========================================================================================================================
; Description ...: coleta o serial da motherboard
;                  Em alguns modelos de motherboard pode retornar vazio, não é erro, simplesmente a informação não foi gravada
;                  logo, não tem número de série
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/AutoIt/blob/master/wmi/CIM_Display.au3
; Source.........: https://msdn.microsoft.com/en-us/library/aa387258(v=vs.85).aspx
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

#include <String.au3>
#include <Array.au3>
ConsoleWrite(wmi_test() & @LF)

Func wmi_test()
	Local $objWMIService = ObjGet("WINMGMTS:\\" & @ComputerName & "\ROOT\cimv2")
	Local $arrName = StringSplit("Availability,Caption,ConfigManagerErrorCode,ConfigManagerUserConfig,CreationClassName,Description,DeviceID,ErrorCleared,ErrorDescription,InstallDate,IsLocked,LastErrorCode,Name,PNPDeviceID,PowerManagementCapabilities,PowerManagementSupported,Status,StatusInfo,SystemCreationClassName,SystemName", ",", 2)

	Local $colItems = $objWMIService.ExecQuery("SELECT * FROM CIM_Display", "WQL", 0x10 + 0x20)
	Local $Result = '', $data
	If IsObj($colItems) Then
		ConsoleWrite('[CIM_Display]' & @LF)
		For $objItem In $colItems
			For $jj = 0 To UBound($arrName) - 1
				$data = Execute('$objItem.' & $arrName[$jj])
				If IsArray($data) Then $data = _ArrayToString($data, ",")
				ConsoleWrite('[' & $arrName[$jj] & _StringRepeat('.', 30 - StringLen($arrName[$jj])) & ']=[' & $data & ']' & @LF)
			Next
		Next
		Return StringMid($Result, 1, StringLen($Result) - 1)
	Else
		Return False
	EndIf
EndFunc   ;==>wmi_test



