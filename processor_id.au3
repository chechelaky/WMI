; #UDF# =========================================================================================================================
; Description ...: coleta processor id do processador (uns dizem que é o serial)
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/processor_id.au3
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

Func GetProcessorId()
	$objWMIService = ObjGet("winmgmts:\\localhost\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_Processor", "WQL",0x10+0x20)
	If IsObj($colItems) Then
		For $objItem In $colItems
			Local $PROC_ID = $objItem.ProcessorId
		Next
			Return $PROC_ID
	Else
		Return 0
	EndIf
EndFunc

MsgBox(0,"",GetProcessorId())