#include <StringConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
Opt("WinTitleMatchMode", 2)

Global $sAccountNum, $sTraceNum, $sTranCode, $sAmount, $sCustID, $sCustName, $sABA, _
	  $sCompanyName, $sEntryDesc, $sEffEntryDateMonth, $sEffEntryDateDay, $sEffEntryDateYear, $sSec, $sDiscData, $sCompanyID, $sReturnCode, $iIndex, $sFile, $sReportDate

#Region GUI
$Form1 = GUICreate("RegE Returns", 308, 138, 192, 124)
$Label1 = GUICtrlCreateLabel("Return Code: R", 16, 24, 74, 17)
$Label2 = GUICtrlCreateLabel("Trace Number:", 16, 62, 75, 17)
$Label3 = GUICtrlCreateLabel("Report Date:", 16, 100, 65, 17)
$Label4 = GUICtrlCreateLabel("Done! Please verify", 176, 112, 115, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
GUICtrlSetColor(-1, 0xFF0000)
GUICtrlSetState(-1, $GUI_HIDE)
$Input1 = GUICtrlCreateInput("", 92, 22, 33, 20, BitOR($GUI_SS_DEFAULT_INPUT,$ES_NUMBER))
GUICtrlSetLimit(-1, 2)
$Input2 = GUICtrlCreateInput("", 96, 60, 105, 20)
GUICtrlSetLimit(-1, 15)
$Input3 = GUICtrlCreateInput("", 88, 97, 65, 20)
GUICtrlSetLimit(-1, 10)
GUICtrlSetTip(-1, "M-D-YYYY")
$Checkbox1 = GUICtrlCreateCheckbox("Skip Bank Num", 130, 24, 90, 17)
$Button1 = GUICtrlCreateButton("Start", 232, 24, 57, 33)
$Button2 = GUICtrlCreateButton("Stop", 231, 73, 57, 33)
GUISetState(@SW_SHOW)
#EndRegion GUI

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Form1
		Case $Form1
		Case $Form1
		Case $Form1
		Case $Label1
		Case $Label2
		Case $Input1
		Case $Input2
		Case $Input3
			If StringLen(GUICtrlRead($Input3)) <> 0 And StringLen(GUICtrlRead($Input3)) < 8 Or StringInStr(GUICtrlRead($Input3), "/") Then
				MsgBox($MB_ICONWARNING, "Error", "Report Date is in an incorrect format. Please enter in M-D-YYYY.")
				GUICtrlSetState($Input3, $GUI_FOCUS)
			EndIf
		Case $Button1
				ClearVariables()
				SearchFile()
				ReadFirstLine()
				ReadSecondLine()
				CleanupVariables()
				MainframeDataEntry()
				GUICtrlSetData($Input2, "")
				GUICtrlSetState($Label4, $GUI_SHOW)
				WinActivate("RegE Returns")
				Sleep(3000)
				GUICtrlSetState($Label4, $GUI_HIDE)
		Case $Button2
			Exit
		Case $Label3
	EndSwitch
WEnd

Func ClearVariables()
$sAccountNum = ""
$sTraceNum = ""
$sTranCode = ""
$sAmount = ""
$sCustID = ""
$sCustName = ""
$sABA = ""
$sCompanyName = ""
$sEntryDesc = ""
$sEffEntryDateMonth = ""
$sEffEntryDateDay = ""
$sEffEntryDateYear = ""
$sSec = ""
$sDiscData = ""
$sCompanyID = ""
$sReturnCode = ""
$iIndex = ""
$sFile = ""
$sReportDate = ""
EndFunc ;==>ClearVariables

Func SearchFile()
$sReportDate = GUICtrlRead($Input3)
Local $sSearch = FileFindFirstFile("ACH ENTRIES LIST*" & $sReportDate & "*.txt")
$sFile = FileFindNextFile($sSearch)
If @error Then
	MsgBox($MB_ICONWARNING, "Error", "No file found for the date entered!")
	Exit
EndIf
$sTextToFind = GUICtrlRead($Input2)
$aArray = FileReadToArray($sFile)
$iStart = 0

While 1
	$iIndex = _ArraySearch($aArray, $sTextToFind, $iStart, "", "", 1)
	If @error = 6 Then
		MsgBox($MB_ICONWARNING, "Error", "Transaction not found!")
		Exit
	EndIf
	If $iIndex = -1 Then Exit
    $iStart = $iIndex + 1
    If $iStart > $aArray[0] Then ExitLoop
WEnd

FileClose($sFile)
EndFunc ;==>SearchFile

Func ReadFirstLine()
Local $sLine1 = FileReadLine($sFile, $iIndex + 1)
Local $SplitLine1 = StringSplit($sLine1, "")

For $i = 1 To 17
$sAccountNum = $sAccountNum & $SplitLine1[$i]
Next
For $i = 21 To 35
$sTraceNum = $sTraceNum & $SplitLine1[$i]
Next
For $i = 37 To 39
$sTranCode = $sTranCode & $SplitLine1[$i]
Next
For $i = 43 To 56
$sAmount = $sAmount & $SplitLine1[$i]
Next
For $i = 58 To 72
$sCustID = $sCustID & $SplitLine1[$i]
Next
For $i = 74 To 94
$sCustName = $sCustName & $SplitLine1[$i]
Next
For $i = 96 To 104
$sABA = $sABA & $SplitLine1[$i]
Next
For $i = 106 To 121
$sCompanyName = $sCompanyName & $SplitLine1[$i]
Next
For $i = 123 To UBound($SplitLine1) - 1
$sEntryDesc = $sEntryDesc & $SplitLine1[$i]
Next
EndFunc ;==>ReadFirstLine

Func ReadSecondLine()
Local $sLine2 = FileReadLine($sFile, $iIndex + 2)
Local $SplitLine2 = StringSplit($sLine2, "")

For $i = 24 To 25
$sEffEntryDateMonth = $sEffEntryDateMonth & $SplitLine2[$i]
Next
For $i = 27 To 28
$sEffEntryDateDay = $sEffEntryDateDay & $SplitLine2[$i]
Next
For $i = 30 To 31
$sEffEntryDateYear = $sEffEntryDateYear & $SplitLine2[$i]
Next
For $i = 37 To 39
$sSec = $sSec & $SplitLine2[$i]
Next
For $i = 48 To 49
$sDiscData = $sDiscData & $SplitLine2[$i]
Next
For $i = 108 To UBound($SplitLine2) - 1
$sCompanyID = $sCompanyID & $SplitLine2[$i]
Next
EndFunc ;==>ReadSecondLine

Func CleanupVariables()
$sABA = StringStripWS($sABA, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sAccountNum = StringStripWS($sAccountNum, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sAmount = StringRegExpReplace($sAmount, "[^0-9]", "")
$sCompanyID = StringStripWS($sCompanyID, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sCompanyName = StringStripWS($sCompanyName, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sCustID = StringStripWS($sCustID, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sCustName = StringStripWS($sCustName, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sDiscData = StringStripWS($sDiscData, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sEffEntryDateDay = StringStripWS($sEffEntryDateDay, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sEffEntryDateMonth = StringStripWS($sEffEntryDateMonth, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sEffEntryDateYear = StringStripWS($sEffEntryDateYear, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sEntryDesc = StringStripWS($sEntryDesc, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sSec = StringStripWS($sSec, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sTraceNum = StringStripWS($sTraceNum, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
$sTranCode = StringStripWS($sTranCode, $STR_STRIPLEADING +  $STR_STRIPTRAILING)
EndFunc ;==>CleanupVariables

Func MainframeDataEntry()
WinActivate("BlueZone Mainframe")
WinWaitActive("BlueZone Mainframe")
Send("{HOME}")
Send("{TAB 2}")
If GUICtrlRead($Checkbox1) = $GUI_UNCHECKED Then Send("493ACH")
If StringLen($sAccountNum) < 17 Then
	Send($sAccountNum)
	Send("{TAB}")
Else
	Send($sAccountNum)
EndIf
Send($sTraceNum)
Send($sEffEntryDateYear)
Send($sEffEntryDateMonth)
Send($sEffEntryDateDay)
Send($sTranCode - 1)
Send($sSec)
Send($sAmount)
Send("{TAB}")
If StringLen($sDiscData) < 2 Then
	Send($sDiscData)
	Send("{TAB}")
Else
	Send($sDiscData)
EndIf
If StringLen($sCustID) < 15 Then
	Send($sCustID)
	Send("{TAB}")
Else
	Send($sCustID)
EndIf
If StringLen($sCustName) < 22 Then
	Send($sCustName)
	Send("{TAB}")
Else
	Send($sCustName)
EndIf
Send($sABA)
If StringLen($sCompanyName) < 16 Then
	Send($sCompanyName)
	Send("{TAB}")
Else
	Send($sCompanyName)
EndIf
If StringLen($sCompanyID) < 10 Then
	Send($sCompanyID)
	Send("{TAB}")
Else
	Send($sCompanyID)
EndIf
If StringLen($sEntryDesc) < 10 Then
	Send($sEntryDesc)
	Send("{TAB}")
Else
	Send($sEntryDesc)
EndIf
$sReturnCode = GUICtrlRead($Input1)
Send($sReturnCode)
EndFunc ;==>MainframeDataEntry
Exit