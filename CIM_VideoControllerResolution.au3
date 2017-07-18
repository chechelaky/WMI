; #UDF# =========================================================================================================================
; Description ...:
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/CIM_VideoControllerResolution.au3
;                  http://stackoverflow.com/questions/7967699/get-screen-resolution-using-wmi-powershell-in-windows-7
; Source.........: https://msdn.microsoft.com/en-us/library/aa388669(v=vs.85).aspx
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

#include <String.au3>
#include <Array.au3>
ConsoleWrite(wmi_test() & @LF)

Func wmi_test()
	Local $objWMIService = ObjGet("WINMGMTS:\\" & @ComputerName & "\ROOT\cimv2")
	Local $arrName = StringSplit("Caption,Description,HorizontalResolution,MaxRefreshRate,MinRefreshRate,NumberOfColors,RefreshRate,ScanMode,SettingID,VerticalResolution", ",", 2)

	Local $colItems = $objWMIService.ExecQuery("SELECT * FROM CIM_VideoControllerResolution", "WQL", 0x10 + 0x20) ;  Where DeviceID=""DesktopMonitor1""
	Local $Result = '', $data
	If IsObj($colItems) Then
		ConsoleWrite('[CIM_VideoControllerResolution]' & @LF)
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



