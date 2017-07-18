; #UDF# =========================================================================================================================
; Description ...: coleta o serial da motherboard
;                  Em alguns modelos de motherboard pode retornar vazio, não é erro, simplesmente a informação não foi gravada
;                  logo, não tem número de série
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMi/motherboard_02.au3
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

ConsoleWrite(_BiosSerial() & @LF)

Func _BiosSerial()
	Local $sAns
	$WMI = ObjGet("WinMgmts:")
	$objs = $WMI.InstancesOf("Win32_BaseBoard")
	For $obj In $objs
		$sAns = $sAns & $obj.SerialNumber
		If $sAns < $objs.Count Then $sAns = $sAns
	Next
	Return $sAns
EndFunc   ;==>_BiosSerial