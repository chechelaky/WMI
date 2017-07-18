; #UDF# =========================================================================================================================
; Description ...: coleta diversas informações interessantes do processador
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/processador_01.au3
; Source.........: https://msdn.microsoft.com/en-us/library/aa394373(v=vs.85).aspx
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

#include <Array.au3>
#include <String.au3>

Func page_win32_processor()
	; verson 0.1
	Local $oo = ObjCreate('Scripting.Dictionary')
	Local $objWMIService = ObjGet('winmgmts:\\' & @ComputerName & '\root\cimv2')
	;Local $arrName = StringSplit('AddressWidth,Architecture,Availability,Caption,ConfigManagerErrorCode,ConfigManagerUserConfig,CpuStatus,CreationClassName,CurrentClockSpeed,CurrentVoltage,DataWidth,Description,DeviceI,ErrorCleared,ErrorDescription,ExtClock,Family,InstallDate,L2CacheSize,L2CacheSpeed,L3CacheSize,L3CacheSpeed,LastErrorCode,Level,LoadPercentage,Manufacturer,MaxClockSpeed,Name,NumberOfCres,NumberOfLogicalProcessors,OtherFamilyDescription,PNPDeviceID,PowerManagementCapabilities,PowerManagementSupported,ProcessorId,ProcessorType,Revision,Role,SocketDesignation,StatusSatusInfo,Stepping,SystemCreationClassName,SystemName,UniqueId,UpgradeMethod,Version,VoltageCaps', ',', 2)
;~ 	Local $arrName = StringSplit('Manufacturer,NumberOfLogicalProcessors,Name,Description,MaxClockSpeed,L2CacheSize,L3CacheSize,CreationClassName,SystemCreationClassName,Revision', ',', 2)
	Local $string = 'AddressWidth,Architecture,Caption,CreationClassName,L2CacheSize,L3CacheSize,Manufacturer,MaxClockSpeed,Name,NumberOfCores,NumberOfLogicalProcessors,ProcessorId,Revision,SystemCreationClassName'
	Local $arrName = StringSplit($string, ',', 2)
	Local $colItems = $objWMIService.ExecQuery('SELECT ' & _ArrayToString($arrName, ',') & ' FROM Win32_Processor', 'WQL', 0x10 + 0x20)
	If IsObj($colItems) Then
		For $objItem In $colItems
			For $jj = 0 To UBound($arrName) - 1
				$oo.Add($arrName[$jj], Execute('$objItem.' & $arrName[$jj]))
			Next
		Next
		Return $oo
	Else
		Return SetError(0, 0, 0)
	EndIf
EndFunc   ;==>page_win32_processor

;http://msdn.microsoft.com/en-us/library/aa394199(v=vs.85).aspx
#include <String.au3>
ConsoleWrite(wmi_test() & @LF)

Func wmi_test()
	Local $objWMIService = ObjGet('winmgmts:\\' & @ComputerName & '\root\cimv2')
	Local $string = 'AddressWidth,Architecture,Caption,CreationClassName,L2CacheSize,L3CacheSize,Manufacturer,MaxClockSpeed,Name,NumberOfCores,NumberOfLogicalProcessors,ProcessorId,Revision,SystemCreationClassName'
	Local $arrName = StringSplit($string, ',', 2)

	Local $colItems = $objWMIService.ExecQuery('SELECT * FROM Win32_Processor', 'WQL', 0x10 + 0x20)
	Local $Result = ''
	If IsObj($colItems) Then
		ConsoleWrite('[Win32_Processor]' & @LF)
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



