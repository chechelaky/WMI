; #UDF# =========================================================================================================================
; Description ...: coleta ip e mac address
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/ip_macaddres.au3
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems = ""
$strComputer = "localhost"

$Output = ""
$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
$colItems = $objWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True', "WQL", _
		$wbemFlagReturnImmediately + $wbemFlagForwardOnly)

If IsObj($colItems) Then
	For $objItem In $colItems
		$strIPAddress = $objItem.IPAddress(0)
		$Output = $Output & "IPAddress: " & $strIPAddress & @CRLF
		$Output = $Output & "MACAddress: " & $objItem.MACAddress & @CRLF
		If MsgBox(1, "WMI Output", $Output) = 2 Then ExitLoop
		$Output = ""
	Next
Else
	MsgBox(0, "WMI Output", "No WMI Objects Found for class: " & "Win32_NetworkAdapterConfiguration")
EndIf
