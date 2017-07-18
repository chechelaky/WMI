; #UDF# =========================================================================================================================
; Description ...: coleta informações de hard disk, mac addres e virtual machine
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/hdd_mac_vm_01.au3
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

;	http://www.autoitscript.com/forum/topic/69501-drivegetserial-report-different-hdd-serial-no/page__hl__drive++serial
;	http://www.codeproject.com/KB/cs/hard_disk_serialno.aspx
#include <File.au3>

ConsoleWrite("Hard Disk Drive Serial[" & _HddSerial() & "]" & @LF)
ConsoleWrite("MacAdress[" & _Mac() & "]" & @LF)
ConsoleWrite("Virtual Machine[" & _HddCheck() & "]" & @LF)

Func _HddSerial()
	;	Função adaptada
	;	Luismar Chechelaky
	;	06/09/2011
	;	http://www.autoitscript.com/forum/topic/69501-drivegetserial-report-different-hdd-serial-no/page__hl__drive++serial
	;	http://www.codeproject.com/KB/cs/hard_disk_serialno.aspx
	Local $objWMIService =  ObjGet("winmgmts:\\" & @ComputerName  & "\root\cimv2")
	Local $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive", "WQL", 0x10 + 0x20)
	Local $Counter = 0
	Local $Result = ""
	If IsObj($colItems) then
		For $objItem In $colItems
			If StringInStr($objItem.Name,"\\.\PHYSICALDRIVE" & $Counter)>=0 Then $Result &= $Counter & ":" & $objItem.SerialNumber
			$Counter +=1
		Next
		Return StringMid($Result,1,StringLen($Result)-1)
	Else
		Return False
	Endif
EndFunc

Func _Mac()
	;	06/09/2011
	;	http://www.autoitscript.com/forum/topic/93183-wmi-ip-data/

	Local $objWMIService =  ObjGet("winmgmts:\\" & @ComputerName  & "\root\cimv2")
	Local $colAdapters = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	If IsObj($colAdapters) then
		For $objAdapter In $colAdapters
			Return $objAdapter.MACAddress
		Next
	Else
		Return False
	Endif
EndFunc

Func _HddCheck()
	;	11/09/2011
	;	http://www.autoitscript.com/forum/topic/131876-detect-current-os/page__p__918467#entry918467
	Local $oWMIService = ObjGet("winmgmts:\\" & @ComputerName  & "\root\cimv2"), $VM, $Serial
	Local $oColItems = $oWMIService.ExecQuery("Select * From Win32_ComputerSystemProduct", "WQL", 0x30)
	If IsObj($oColItems) Then
		For $oObjectItem In $oColItems
			If StringRegExp($oObjectItem.Name, "(?i)Virtual Machine|VMware Virtual Platform|VirtualBox|VMWare|Virtual PC") Then
				$VM = 1
				Dim $szDrive, $szDir, $szFName, $szExt
				$Drive = _PathSplit(@ScriptFullPath, $szDrive, $szDir, $szFName, $szExt)
				$Serial = DriveGetSerial($Drive[1])
			Else
				$VM = 0
				Local $colItems = $oWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive", "WQL", 0x10 + 0x20)
				Local $Result = ""
				If IsObj($colItems) then
					For $objItem In $colItems
						If StringInStr($objItem.Name,"\\.\PHYSICALDRIVE0") >= 0 Then $Result = $objItem.SerialNumber
					Next
					$Serial = StringMid($Result,1,StringLen($Result)-1)
				Else
					$Serial = -1
				Endif
			EndIf
		Next
	EndIf
	;Return SetError(1, 0, 0)
	Return $VM & ":" & $Serial
EndFunc   ;==>_HddCheck