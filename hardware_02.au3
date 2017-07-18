; #UDF# =========================================================================================================================
; Description ...: coleta diversas informações interessantes do hardware II
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/hardware_02.au3
; Source.........: ;http://msdn.microsoft.com/en-us/library/aa394199(v=vs.85).aspx
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

#include <String.au3>
ConsoleWrite(wmi_test() & @LF)

Func wmi_test()
	Local $objWMIService = ObjGet("winmgmts:\\.\root\cimv2")
	;Local $arrName = StringSplit('Caption,ConfigOptions,CreationClassName,Depth,Description,Height,HostingBoard,HotSwappable,InstallDate,Manufacturer,Model,Name,OtherIdentifyingInfo,PartNumber,PoweredOn,Product,Removable,Replaceable,RequirementsDescription,RequiresDaughterBoard,SerialNumber,SKU,SlotLayout,SpecialRequirements,Status,Tag,Version,Weight,Width', ',', 2)
	Local $arrName = StringSplit('Caption,CreationClassName,Manufacturer,Model,Name,OtherIdentifyingInfo,PartNumber,Product,SerialNumber,SKU,Tag,Version', ',', 2)

	Local $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_BaseBoard", "WQL", 0x10 + 0x20)
	Local $Result = ''
	If IsObj($colItems) Then
		ConsoleWrite('[Win32_BaseBoard]' & @LF)
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
