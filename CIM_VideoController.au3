; #UDF# =========================================================================================================================
; Description ...: coleta informações da placa de vídeo
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/CIM_VideoController.au3
;~                 http://stackoverflow.com/questions/7967699/get-screen-resolution-using-wmi-powershell-in-windows-7
; Source.........: https://msdn.microsoft.com/en-us/library/aa388668(v=vs.85).aspx
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

#include <String.au3>
#include <Array.au3>
ConsoleWrite(wmi_test() & @LF)

Func wmi_test()
	Local $objWMIService = ObjGet("WINMGMTS:\\" & @ComputerName & "\ROOT\cimv2")
	Local $arrName = StringSplit("AcceleratorCapabilities,Availability,CapabilityDescriptions,Caption,ConfigManagerErrorCode,ConfigManagerUserConfig,CreationClassName,CurrentBitsPerPixel,CurrentHorizontalResolution,CurrentNumberOfColors,CurrentNumberOfColumns,CurrentNumberOfRows,CurrentRefreshRate,CurrentScanMode,CurrentVerticalResolution,Description,DeviceID,ErrorCleared,ErrorDescription,InstallDate,LastErrorCode,MaxMemorySupported,MaxNumberControlled,MaxRefreshRate,MinRefreshRate,Name,NumberOfVideoPages,PNPDeviceID,PowerManagementCapabilities,PowerManagementSupported,ProtocolSupported,Status,StatusInfo,SystemCreationClassName,SystemName,TimeOfLastReset,VideoMemoryType,VideoProcessor", ",", 2)

	Local $colItems = $objWMIService.ExecQuery("SELECT * FROM CIM_VideoController", "WQL", 0x10 + 0x20) ;  Where DeviceID=""DesktopMonitor1""
	Local $Result = '', $data
	If IsObj($colItems) Then
		ConsoleWrite('[CIM_VideoController]' & @LF)
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

