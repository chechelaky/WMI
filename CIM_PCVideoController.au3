; #UDF# =========================================================================================================================
; Description ...: coleta informações da placa de vídeo
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/CIM_PCVideoController.au3
; Source.........: https://msdn.microsoft.com/en-us/library/aa387956(v=vs.85).aspx
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

#include <Array.au3>
#include <String.au3>
ConsoleWrite(wmi_test() & @LF)

Func wmi_test()
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\cimv2")
	Local $arrName = StringSplit('AcceleratorCapabilities,AdapterCompatibility,AdapterDACType,AdapterRAM,Availability,CapabilityDescriptions,Caption,ColorTableEntries,ConfigManagerErrorCode,ConfigManagerUserConfig,CreationClassName,CurrentBitsPerPixel,CurrentHorizontalResolution,CurrentNumberOfColors,CurrentNumberOfColumns,CurrentNumberOfRows,CurrentRefreshRate,CurrentScanMode,CurrentVerticalResolution,Description,DeviceID,DeviceSpecificPens,DitherType,DriverDate,DriverVersion,ErrorCleared,ErrorDescription,ICMIntent,ICMMethod,InfFilename,InfSection,InstallDate,InstalledDisplayDrivers,LastErrorCode,MaxMemorySupported,MaxNumberControlled,MaxRefreshRate,MinRefreshRate,Monochrome,Name,NumberOfColorPlanes,NumberOfVideoPages,PNPDeviceID,PowerManagementCapabilities,PowerManagementSupported,ProtocolSupported,ReservedSystemPaletteEntries,SpecificationVersion,Status,StatusInfo,SystemCreationClassName,SystemName,SystemPaletteEntries,TimeOfLastReset,VideoArchitecture,VideoMemoryType,VideoMode,VideoModeDescription,VideoProcessor', ',', 2)

	Local $colItems = $objWMIService.ExecQuery('SELECT * FROM CIM_PCVideoController', 'WQL', 0x10 + 0x20)
	Local $Result = ''
	Local $data
	If IsObj($colItems) Then
		ConsoleWrite('[CIM_PCVideoController]' & @LF)
		For $objItem In $colItems
			For $jj = 0 To UBound($arrName) - 1
				$data = Execute('$objItem.' & $arrName[$jj])
				If IsArray($data) Then $data = _ArrayToString(Execute('$objItem.' & $arrName[$jj]), "|")
				ConsoleWrite('[' & $arrName[$jj] & _StringRepeat('.', 30 - StringLen($arrName[$jj])) & ']=[' & $data & ']' & @LF)
			Next
		Next
		Return StringMid($Result, 1, StringLen($Result) - 1)
	Else
		Return False
	EndIf
EndFunc   ;==>wmi_test

