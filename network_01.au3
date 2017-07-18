; #UDF# =========================================================================================================================
; Description ...: coleta ip, macaddres e nome do adaptador de rede
; Exemplo........:
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/network_01.au3
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

#include <Array.au3>

Global $ConnInfo[1][3]
$ConnInfo[0][0] = 0
;_ArrayDisplay($ConnInfo)

$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems = ""
$strComputer = "localhost"

$Output = ""
$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
$colItems = $objWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True', "WQL", _
		$wbemFlagReturnImmediately + $wbemFlagForwardOnly)


If IsObj($colItems) Then
	$i = 1
	For $objItem In $colItems
		ReDim $ConnInfo[$i + 1][3]
		$ConnInfo[$i][0] = $objItem.IPAddress(0)
		$ConnInfo[$i][1] = $objItem.MACAddress
		$i = $i + 1
		$ConnInfo[0][0] += 1
	Next
EndIf

;~ _ArrayDisplay($ConnInfo)

For $i = 1 To $ConnInfo[0][0]

	$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
	$colItems = $objWMIService.ExecQuery('SELECT * FROM Win32_NetworkAdapter WHERE MACAddress = "' & $ConnInfo[$i][1] & '"', "WQL", _
			$wbemFlagReturnImmediately + $wbemFlagForwardOnly)

	If IsObj($colItems) Then
		For $objItem In $colItems
			$ConnInfo[$i][2] = $objItem.NetConnectionID
		Next
	EndIf

Next

_ArrayDisplay($ConnInfo)
