; #UDF# =========================================================================================================================
; Description ...: Colega alguns dados da motherboard: serial number (as vases o serial é vazio), modelo e versão da motherboard
; Author ........: Luigi (Luismar Chechelaky)
; Link ..........: https://github.com/chechelaky/WMI/motherboard_01.au3
; AutoIt version.: 3.3.14.2
; ===============================================================================================================================

#include <file.au3>
Global $itdid, $objWMIService, $colItems, $sWMIService, $sName, $sModel, $uuItem, $objSWbemObject, $strName, $strVersion, $strWMIQuery, $objItem, $uiDitem
;$itdid = RegRead ( "HKEY_LOCAL_MACHINE\SOFTWARE\Intel\PIcon\AMTData", "System UUID" )
$sWMIService = "winmgmts:\\" & @ComputerName & "\root\CIMV2"
$objWMIService = ObjGet($sWMIService)
If IsObj($objWMIService) Then
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_BaseBoard")
	If IsObj($colItems) Then
		For $oItem In $colItems
			$sName = $oItem.Product
			$sModel = $oItem.SerialNumber
		Next
	EndIf
	$uuItem = $objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
	If IsObj($uuItem) Then
		For $objSWbemObject In $uuItem
			$strIdentifyingNumber = $objSWbemObject.IdentifyingNumber
			$strName = $objSWbemObject.Name
			$strVersion = $objSWbemObject.Version
		Next
	EndIf
	$strWMIQuery = ":Win32_ComputerSystemProduct.IdentifyingNumber='" & $strIdentifyingNumber & "',Name='" & $strName & "',Version='" & $strVersion & Chr(39)
	$uiDitem = ObjGet($sWMIService & $strWMIQuery)
	If IsObj($uiDitem) Then
		For $objItem In $uiDitem.Properties_
			$itdid = $objItem.value
		Next
;~ 		_FileWriteLog(@ScriptDir & "\ITD.log", $sModel & " " & $sName & " " & $itdid)
		MsgBox(64, "Win32_BaseBoard Item", "SN = " & $sModel & @CRLF & "MB = " & $sName & @CRLF & "ITD ID = " & $itdid & @CRLF & "Log file Updated")
	EndIf
EndIf
