; #UDF# =========================================================================================================================
; Description ...: coleta informações de dispositivos físicos de entrada/saída (cd-rom, pendrive, hdd, etc)
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/store_devices.au3
; Source.........: https://msdn.microsoft.com/en-us/library/aa394132(v=vs.85).aspx
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

ConsoleWrite(_HddSerial() & @LF)

Func _HddSerial()
	;	Adapted function
	;	06/09/2011
	;	http://www.autoitscript.com/forum/topic/69501-drivegetserial-report-different-hdd-serial-no/page__hl__drive++serial
	;	http://www.codeproject.com/KB/cs/hard_disk_serialno.aspx
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\cimv2")
	Local $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive", "WQL", 0x10 + 0x20)
	Local $Counter = 0
	Local $Result = ""
	If IsObj($colItems) Then
		For $objItem In $colItems
			If StringInStr($objItem.Name, "\\.\PHYSICALDRIVE" & $Counter) >= 0 Then $Result &= $Counter & ":" & $objItem.SerialNumber & @LF
;~ 			If StringInStr($objItem.Description, "\\.\PHYSICALDRIVE" & $Counter) >= 0 Then $Result &= $Counter & ":" & $objItem.Description & @LF
			If StringInStr($objItem.Model, "\\.\PHYSICALDRIVE" & $Counter) >= 0 Then $Result &= $Counter & ":" & $objItem.Model & @LF
;~ 			If StringInStr($objItem.DeviceID, "\\.\PHYSICALDRIVE" & $Counter) >= 0 Then $Result &= $Counter & ":" & $objItem.DeviceID & @LF
			$Counter += 1
		Next
		Return StringMid($Result, 1, StringLen($Result) - 1)
	Else
		Return False
	EndIf
EndFunc   ;==>_HddSerial