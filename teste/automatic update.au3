#cs -----------------------------------------------------------------------------------------------------------------
	Author:
	Colyn Via
	Script Function:
	Query local device WMI info and perform a SQL lookup to check for firmware updates. The SQL database contains
	device models, current firmware version, and the path to the update package. This script looks for the firmware
	version based on the device model, and then copies the installer to the host drive. The script then performs a
	silent install of the firmware update all the while preventing an accidental shutdown. This script has a verbose
	log which gets copied to a network location if there's a problem. If it's successful the log writes a single
	line to a master log in the same network location reporting it's success. Once the update is complete, the
	installer file is deleted, as well as the local log file.
	Version Notes:
	1.0.03 - Adjusted final write to local log so that the $FirmwareDBResult is written in the log.
	1.0.02 - Added processwait() to script to prevent premature script completion
	1.0.01 - Added values to BIOSString dump in addition to names
	1.0.00 - Initial complete version
#ce -----------------------------------------------------------------------------------------------------------------
#NoTrayIcon
#include
#include
; Unless directed otherwise, autoit will do an internet lookup for the driver to use. Later in the code
; the _SQLite_Startup UDF directs the program to get the driver locally. This is why it's important to
; include the driver in the package.
FileInstall("SQLite3.dll", @ScriptDir & "", 1)
Global $FWversion, $Product, $hDB, $hQuery, $aRow, $UpgradeDBResult, $FWversionDBResult, $CheckCurrent
Global $CheckDB, $UpgradeEXE, $SQLQuery1, $SQLQuery2, $PID
Global Const $Timestamp = @MON & "/" & @MDAY & "/" & @YEAR & " " & @HOUR & ":" & @MIN & ":" & @SEC & " "
Global Const $Log = "C:AMTFirmwareLog_" & @ComputerName & ".txt"
Global Const $GlobalLog = 'redactedvProFirmware ResponsesSuccessSucceeded.txt'
Global Const $FirmwareDB = 'redactedvProFirmware.db'
Global Const $GlobalResponse = 'redactedvProFirmware Responses'
; This script was designed to be verbose and log most events. This code creates the log and adds the first line.
FileWriteLine($Log, $Timestamp & "Hello! My name is " & @ComputerName & ". " & @OSVersion & " " & @OSArch & " is how I roll.")
FileWriteLine($Log, $Timestamp & "I'm running script 1.0.03.")
FWversion()
; This function runs a WMI query to get the current AMT firmware version on the device.
Func FWversion()
	$objIntelME = ObjGet("winmgmts:" & "." & "rootIntel_ME")
	; An error catch must be present here, as AutoIt does not mute COM errors naturally. Without this catch,
	; an ugly sounding error will appear to the user. IsObj() is not enough, as it cannot handle a return.
	If @error Then
		FileWriteLine($Log, $Timestamp & "Intel_ME is not a valid NAMESPACE.")
		FileCopy($Log, $GlobalResponse)
		FileDelete($Log)
		Exit
	EndIf
	$colIntelME = $objIntelME.ExecQuery("SELECT * FROM ME_System")
	; The "Else" of this statement reports back which devices couldn't determine their firmware version.
	If IsObj($colIntelME) Then
		For $objMEsystem In $colIntelME
			$FWversion = $objMEsystem.fwversion
		Next
	Else
		FileWriteLine($Log, $Timestamp & "ME_System query failed.")
		FileCopy($Log, $GlobalResponse)
		FileDelete($Log)
		Exit
	EndIf
	FileWriteLine($Log, $Timestamp & "Local host vPro firmware lookup result = " & $FWversion)
	ConsoleWrite("Local host vPro firmware lookup result = " & $FWversion & @LF)
EndFunc   ;==>FWversion
Product()
; This function runs a WMI query to get the current model information for the device.
Func Product()
	; Since these variables are used often in general coding practice, they are Local'd instead of
	; declared global. When the code finishes this function, it will drop these variables.
	Local $x, $y, $prodArray, $aRefine
	$objWMI = ObjGet("winmgmts:" & @ComputerName & "rootHPInstrumentedBIOS")
	If @error Then
		FileWriteLine($Log, $Timestamp & "rootHPInstrumentedBIOS is not a valid NAMESPACE.")
		FileCopy($Log, $GlobalResponse)
		FileDelete($Log)
		Exit
	EndIf
	$colItems = $objWMI.ExecQuery("SELECT * FROM HPBIOS_BIOSString")
	; The "Else" of this block accounts for the possibility that this code might run on a non HP device.
	If IsObj($colItems) Then
		FileWriteLine($Log, $Timestamp & "Attempting to grab hardware info.")
		For $objItems In $colItems
			If $objItems.Name = "Notebook Model" Then
				$Product = $objItems.Value
				FileWriteLine($Log, $Timestamp & "Raw value for " & $objItems.Name & " is " & $Product)
				$prodArray = StringSplit($Product, " ")
				; BIOS revisions change slightly how the information is presented by WMI. Sometimes it's
				; simply the Model number, other times it can be a string of words. For simplicity and
				; uniformity, this code strips out just the Model number from the string. This code is
				; looking for a word that starts with a number. If HP begins using Model Numbers that
				; start with letters, this code will need to be reworked.
				For $x = 1 To $prodArray[0]
					If StringIsAlpha($prodArray[$x]) Then
					Else
						; As it turns out, HP in its infinite wisdom decided to, in some cases, put both
						; product identifiers in the SAME CIM component value. This confuses the scrip and
						; it ends up recording the wrong value. To correct for this, and as the model
						; number always starts with a number (currently), I've directed the script to check
						; the first character in the string it's wanting to use to see if it's a number
						; IsInt means "is it an integer", which in this case, the value is not. AutoIt
						; sees it as a string, so Number() is used instead.
						FileWriteLine($Log, $Timestamp & "I'm considering using " & $prodArray[$x] & ".")
						$aRefine = StringSplit($prodArray[$x], "")
						If Number($aRefine[1]) Then
							FileWriteLine($Log, $Timestamp & "It turns out that " & $prodArray[$x] & " starts with " & $aRefine[1] & _
									" and must therefore be what I'm looking for. Right?")
							$Product = ($prodArray[$x])
						EndIf
					EndIf
				Next
				ExitLoop
			ElseIf $objItems.Name = "Product Name" Then
				$Product = $objItems.Value
				FileWriteLine($Log, $Timestamp & "Raw value for " & $objItems.Name & " is " & $Product)
				$prodArray = StringSplit($Product, " ")
				For $x = 1 To $prodArray[0]
					If StringIsAlpha($prodArray[$x]) Then
					Else
						FileWriteLine($Log, $Timestamp & "I'm considering using " & $prodArray[$x] & ".")
						$aRefine = StringSplit($prodArray[$x], "")
						If Number($aRefine[1]) Then
							FileWriteLine($Log, $Timestamp & "It turns out that " & $prodArray[$x] & " starts with " & $aRefine[1] & _
									" and must therefore be what I'm looking for. Right?")
							$Product = ($prodArray[$x])
						EndIf
					EndIf
				Next
				ExitLoop
			EndIf
		Next
		FileWriteLine($Log, $Timestamp & "Finished attempts to grab hardware info.")
		; In the event that "Notebook Model" is not the value to read, it helps to see what should be added.
		; The following code runs a dump of the contents of HPBIOS_BIOSstring and copies the log to the log directory.
		; Later the results of the log can be used to determine what other $objItems.Name might exist.
		If $Product = "" Then
			FileWriteLine($Log, $Timestamp & "Query (SELECT * FROM HPBIOS_BIOSString) SUCCEEDED! However unable to " & _
					"resolve object.name of known variables on " & @ComputerName & ".")
			FileWriteLine($Log, "*************** Beginning Dump of HPBIOS_BIOSstring ***************")
			For $objItems In $colItems
				FileWriteLine($Log, "*** " & $objItems.name & " = " & $objItems.Value)
			Next
			FileWriteLine($Log, "*************** Done Exporting HPBIOS_BIOSstring ***************")
			Sleep(5)
			FileCopy($Log, $GlobalResponse)
			FileDelete($Log)
			Exit
		EndIf
	Else
		FileWriteLine($Log, $Timestamp & "rootHPInstrumentedBIOS is not a valid NAMESPACE on " & @ComputerName & ".")
	EndIf
	ConsoleWrite("WMI query for Model produced " & $Product & @LF)
EndFunc   ;==>Product
If $FWversion = "" Then Exit
If $Product = "" Then Exit
SQL()
; This function names the SQL queries and initializes the SQL engine.
Func SQL()
	; SQL Queries
	$SQLQuery1 = "SELECT " & '"' & "Update" & '"' & " FROM Lookup WHERE Product=" & '"' & $Product & '"' & ';'
	$SQLQuery2 = "SELECT " & '"' & "FWVersion" & '"' & " FROM Lookup WHERE Product=" & '"' & $Product & '"' & ';'
	; Check the DB and open it if the script can find it
	FileWriteLine($Log, $Timestamp & "Checking to see if " & '"' & $FirmwareDB & '"' & " exists...")
	If FileExists($FirmwareDB) Then
		FileWriteLine($Log, $Timestamp & "Found it!")
	Else
		FileWriteLine($Log, $Timestamp & "I can't see it!")
		FileCopy($Log, $GlobalResponse)
		FileDelete($Log)
		Exit
	EndIf
	_SQLite_Startup("", "", 0)
	_SQLite_Open($FirmwareDB)
EndFunc   ;==>SQL
Eligibility()
; This function determines the devices eligibility to upgrade, and figures out where to get the update from.
Func Eligibility()
	; Since these variables are used often in general coding practice, they are Local'd instead of
	; declared global. When the code finishes this function, it will drop these variables.
	Local $x, $y, $z, $Proceed, $diff, $exe
	_SQLite_Query(-1, $SQLQuery1, $hQuery)
	If _SQLite_FetchData($hQuery, $aRow) = $SQLITE_OK Then
	Else
		ConsoleWrite("Query Failed")
		FileWriteLine($Log, $Timestamp & $SQLQuery1 & " FAILED.")
		FileCopy($Log, $GlobalResponse)
		FileDelete($Log)
		_SQLite_QueryFinalize($hQuery)
		_SQLite_Close()
		_SQLite_Shutdown()
		Exit
	EndIf
	$UpgradeDBResult = $aRow[0]
	FileWriteLine($Log, $Timestamp & $SQLQuery1 & " SUCCEEDED.")
	_SQLite_QueryFinalize($hQuery)
	ConsoleWrite($UpgradeDBResult & @LF)
	; Need to get just the name of the update executable for use later.
	$exe = StringSplit($UpgradeDBResult, "")
	$UpgradeEXE = $exe[($exe[0])]
	FileWriteLine($Log, $Timestamp & "The executable needed later is " & $UpgradeEXE & ".")
	; Testing purposes - Making my host appear to be a 6550b
;~ $SQLQuery2 = "SELECT " & '"' & "FWVersion" & '"' & " FROM Lookup WHERE Product=" & '"' & "6550b" & '"' & ';'
	_SQLite_Query(-1, $SQLQuery2, $hQuery)
	If _SQLite_FetchData($hQuery, $aRow) = $SQLITE_OK Then
	Else
		ConsoleWrite("Query Failed")
		FileWriteLine($Log, $Timestamp & $SQLQuery2 & " FAILED.")
		FileCopy($Log, $GlobalResponse)
		FileDelete($Log)
		_SQLite_QueryFinalize($hQuery)
		_SQLite_Close()
		_SQLite_Shutdown()
		Exit
	EndIf
	$FWversionDBResult = $aRow[0]
	FileWriteLine($Log, $Timestamp & $SQLQuery2 & " SUCCEEDED.")
	_SQLite_QueryFinalize($hQuery)
	ConsoleWrite($FWversionDBResult & @LF)
	FileWriteLine($Log, $Timestamp & "Query result was " & $FWversionDBResult)
	; Be sure to close down SQL when it's not needed.
	_SQLite_Close()
	_SQLite_Shutdown()
	; Testing purposes - Make my firmware appear older
;~ $FWversion = "6.2.0.1022"
	FileWriteLine($Log, $Timestamp & "Comparing firmware versions.")
	; So as not to run the risk of destroying the data, the data is copied to new
	; variables before be begin to break it apart into arrays.
	$CheckCurrent = StringSplit($FWversion, ".")
	$CheckDB = StringSplit($FWversionDBResult, ".")
	FileWriteLine($Log, $Timestamp & "Splitting firmware strings.")
	; This code checks to make sure the two arrays are equal in terms of size.
	; If they are not, the code will generate an error. This code also manually
	; adjusts either array so that it always matches the larger of the two.
	If $CheckDB[0] > $CheckCurrent[0] Then
		$diff = ($CheckDB[0] - $CheckCurrent[0])
		For $z = 1 To $diff
			_ArrayAdd($CheckCurrent, 0)
		Next
	EndIf
	If $CheckDB[0] < $CheckCurrent[0] Then
		$diff = ($CheckCurrent[0] - $CheckDB[0])
		For $z = 1 To $diff
			_ArrayAdd($CheckDB, 0)
		Next
	EndIf
	FileWriteLine($Log, $Timestamp & "Finished comparing stringsplit array sizes, continuing to determine if upgrade is needed.")
	; This part is checking to see if the firmware update indicated by the SQL query
	; is newer than the current version on the device. If the firmware is current or
	; better, the program deletes the log file and closes.
	Do
		For $x = 1 To $CheckDB[0]
			If $CheckDB[$x] > $CheckCurrent[$x] Then
				$Proceed = "y"
				ExitLoop
			ElseIf $CheckDB[$x] < $CheckCurrent[$x] Then
				$Proceed = "n"
				ExitLoop
			ElseIf $CheckDB[$x] = $CheckCurrent[$x] And $x = $CheckDB[0] Then
				$Proceed = "n"
				ExitLoop
			EndIf
		Next
	Until $Proceed <> ""
	If $Proceed = "n" Then
		FileDelete($Log)
		Exit
	EndIf
	FileWriteLine($Log, $Timestamp & "Finished comparing firmware versions. Proceeding to UpdateFirmware()")
EndFunc   ;==>Eligibility
; $UpgradeDBResult is the path to the firmware update as retrieved by the Eligibility() function
UpdateFirmware()
; This function pulls down the update and executes it. It also writes to the $GlobalLog.
Func UpdateFirmware()
	FileWriteLine($Log, $Timestamp & "Attempting to copy " & '"' & $UpgradeDBResult & '"' & " to local drive.")
	If FileExists($UpgradeDBResult) Then
		FileCopy($UpgradeDBResult, @HomeDrive & "")
	Else
		FileWriteLine($Log, $Timestamp & "FileExists() check failed. I quit!")
		FileCopy($Log, $GlobalResponse)
		FileDelete($Log)
		Exit
	EndIf
	FileWriteLine($Log, $Timestamp & "Run command issued for " & @HomeDrive & "" & $UpgradeEXE)
	; The -s switch is for silent install
	Run(@HomeDrive & "" & $UpgradeEXE & " -s", "", @SW_HIDE)
	; Prevent shutdown during firmware update process
	ProcessWait($UpgradeEXE, 60)
	Do
		Run("shutdown /a", @SystemDir, @SW_HIDE)
		Sleep(10)
		$PID = ProcessExists($UpgradeEXE)
	Until $PID = 0
	FileWriteLine($GlobalLog, $Timestamp & @ComputerName & " launched " & $UpgradeEXE & " to achieve " & $FWversionDBResult & ".")
EndFunc   ;==>UpdateFirmware
FileDelete(@HomeDrive & "" & $UpgradeEXE)
FileDelete($Log)
Exit
