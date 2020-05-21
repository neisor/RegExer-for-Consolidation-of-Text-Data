#SingleInstance force

SetWorkingDir %A_ScriptDir%

version = Customer Non-SAP - Org.Remarks Validator (v1.9)

Gui, Add, Edit, x11 y32 w340 h340 vOldOrgRemarks, 
Gui, Add, Edit, x491 y32 w340 h340 vNewOrgRemarksTextField, 
Gui, Add, Button, vConvertButton gConvertOrgRemarks x371 y162 w100 h30 , VALIDATE!
Gui, Add, Button, gHelp x409 y285 w25 h25 , ?
Gui, Add, Button, vCopyNewOrgRemarksToClipboardButton gCopyNewOrgRemarksToClipboardButton x371 y230 w100 h30 , Copy to clipboard
Gui, Add, Text, x11 y12 w340 h20 , Paste the original Org.Remarks here:
Gui, Add, Text, x491 y12 w340 h20 , Validated Org.Remarks:
Gui, Add, Text, x371 y205 w100 h20 center, ========>
Gui, Add, Text, x371 y330 w100 h60 center, Created by `nAntonio Raffaele Iannaccone
Gui, Show, x180 y135 h388 w840, %version%
Return


Help:
MsgBox,0,%version%,The Org.Remarks should have this structure:`n`nCustomer naming convention: `nLandscape: `nCategory: `nStatus:`nCommon/Local:`nOPCO:`nDR startup PRIO:`nSupplier:`nFunction support party:`nFunctional Management/Owner:`nTAM:`n###`n`nThe Validator adds automatically N/A to the end of the line if it's empty or if the line is missing it automatically ads the entire line with N/A at the end. `nPlease, add all of the other information which are not found in this Org.Remarks structure after the ### (in the comments section).

;Main Function
ConvertOrgRemarks:
#IfWinActive, Customer Non-SAP - Org.Remarks Validator
~Enter::
#IfWinActive
;empty the NewOrgRemarks text field
GuiControl,, NewOrgRemarksTextField,

;make myString empty
myString =

GuiControlGet, OutputOldOrgRemarksFROMgui,, OldOrgRemarks

;check if old Org.Remarks are pasted into the OldOrgRemarks text field, and if not, give a MsgBox
if(OutputOldOrgRemarksFROMgui = ""){
MsgBox,0,%version%, Original Org.Remarks text field is empty.`n`nPlease, paste the Org.Remarks you want to validate into the Original Org.Remarks text field.
return
}

myString := OutputOldOrgRemarksFROMgui

;Check if there is a new line at the end of the entire pasted Org.Remarks and if not, add a new line
StringRight, CheckIfThereIsASpaceAtTheEndOfTheString, myString, 1
if (CheckIfThereIsASpaceAtTheEndOfTheString = "`n"){
;do nothing
}
else{
myString = %myString%`n`n
}

;Check for multiple spaces and replace them
ReplaceSpaces := RegExReplace(myString, "`r`n`r`n", "`r`n")
myString = %ReplaceSpaces%

ReplaceSpaces2 := RegExReplace(myString, "`n`n", "`n")
myString = %ReplaceSpaces2%

ReplaceSpaces3 := RegExReplace(myString, "`r`n`r`n`r`n", "`r`n")
myString = %ReplaceSpaces3%

ReplaceSpaces4 := RegExReplace(myString, "`n`n`n", "`n`n")
myString = %ReplaceSpaces4%

;Check for 1 uvodzovka " and replace it with nothing
Replace1UvodzovkaWithNothing := RegExReplace(myString, """", "")
myString = %Replace1UvodzovkaWithNothing%

;Check for 6 TABs \t\t\t and replace it with space
Replace6TABswithSpaceString := RegExReplace(myString, "`t`t`t`t`t`t", " ")
myString = %Replace6TABswithSpaceString%

;Check for 5 TABs \t\t\t and replace it with space
Replace5TABswithSpaceString := RegExReplace(myString, "`t`t`t`t`t", " ")
myString = %Replace5TABswithSpaceString%

;Check for 4 TABs \t\t\t and replace it with space
Replace4TABswithSpaceString := RegExReplace(myString, "`t`t`t`t", " ")
myString = %Replace4TABswithSpaceString%

;Check for 3 TABs \t\t\t and replace it with space
Replace3TABswithSpaceString := RegExReplace(myString, "`t`t`t", " ")
myString = %Replace3TABswithSpaceString%

;Check for 2 TABs \t\t and replace it with space
Replace2TABswithSpaceString := RegExReplace(myString, "`t`t", " ")
myString = %Replace2TABswithSpaceString%

;Check for TAB \t and replace it with space
ReplaceTABwithSpaceString := RegExReplace(myString, "`t", " ")
myString = %ReplaceTABwithSpaceString%

;Check for 7 spaces and replace them with one space
Replace7SpacesWithOneSpaceString := RegExReplace(myString, "       ", " ")
myString = %Replace7SpacesWithOneSpaceString%

;Check for 6 spaces and replace them with one space
Replace6SpacesWithOneSpaceString := RegExReplace(myString, "      ", " ")
myString = %Replace6SpacesWithOneSpaceString%

;Check for 5 spaces and replace them with one space
Replace5SpacesWithOneSpaceString := RegExReplace(myString, "     ", " ")
myString = %Replace5SpacesWithOneSpaceString%

;Check for 4 spaces and replace them with one space
Replace4SpacesWithOneSpaceString := RegExReplace(myString, "    ", " ")
myString = %Replace4SpacesWithOneSpaceString%

;Check for 3 spaces and replace them with one space
Replace3SpacesWithOneSpaceString := RegExReplace(myString, "   ", " ")
myString = %Replace3SpacesWithOneSpaceString%

;Check for 2 spaces and replace them with one space
Replace2SpaceswithOneSpaceString := RegExReplace(myString, "  ", " ")
myString = %Replace2SpaceswithOneSpaceString%

;Check for N\A and replace it with N/A
myString := StrReplace(myString, "N\A", "N/A")

FindcustNamingConvention := RegExMatch(myString,"mU)(Customer naming convention:.+)\n", result_FindcustNamingConvention)
;check if the result is found in the text, if not add N/A
if (FindcustNamingConvention = 0){
FindcustNamingConvention := 
FindcustNamingConvention := RegExMatch(myString,"mU)(Customer naming convention:.+)\r", result_FindcustNamingConvention)
if (FindcustNamingConvention = 0){
result_FindcustNamingConvention = Customer naming convention: N/A`n
}
}
if (RegExMatch(result_FindcustNamingConvention,"(Customer naming convention: 0|Customer naming convention:0|Customer naming convention:`n|Customer naming convention: `n)", zero_result_FindcustNamingConvention) = 1){
result_FindcustNamingConvention = Customer naming convention: N/A`n
}
if(RegExMatch(result_FindcustNamingConvention,"\n") = 0){
	result_FindcustNamingConvention = %result_FindcustNamingConvention%`n
}

FindcustLandscape := RegExMatch(myString,"mU)(Landscape:.+)\n", result_FindcustLandscape)
;check if the result is found in the text, if not add N/A
if (FindcustLandscape = 0){
FindcustLandscape := 
FindcustLandscape := RegExMatch(myString,"mU)(Landscape:.+)\r", result_FindcustLandscape)
if (FindcustLandscape = 0){
result_FindcustLandscape = Landscape: N/A`n
}
}
if (RegExMatch(result_FindcustLandscape,"(Landscape: 0|Landscape:0|Landscape:`n|Landscape: `n)", zero_result_FindcustLandscape) = 1){
result_FindcustLandscape = Landscape: N/A`n
}
if(RegExMatch(result_FindcustLandscape,"\n") = 0){
	result_FindcustLandscape = %result_FindcustLandscape%`n
}


FindcustCategory := RegExMatch(myString,"mU)(Category:.+)\n", result_FindcustCategory)
;MsgBox % result_FindcustCategory
;check if the result is found in the text, if not add N/A
if (FindcustCategory = 0){
FindcustCategory := 
FindcustCategory := RegExMatch(myString,"mU)(Category:.+)\r", result_FindcustCategory)
if (FindcustCategory = 0){
result_FindcustCategory = Category: N/A`n
}
}
if (RegExMatch(result_FindcustCategory,"(Category: 0|Category:0|Category:`n|Category: `n)", zero_result_FindcustCategory) = 1){
result_FindcustCategory = Category: N/A`n
}
if(RegExMatch(result_FindcustCategory,"\n") = 0){
	result_FindcustCategory = %result_FindcustCategory%`n
}


FindcustStatus := RegExMatch(myString,"mU)(Status:.+)\n", result_FindcustStatus)
;MsgBox % result_FindcustStatus
;check if the result is found in the text, if not add N/A
if (FindcustStatus = 0){
FindcustStatus := 
FindcustStatus := RegExMatch(myString,"mU)(Status:.+)\r", result_FindcustStatus)
if (FindcustStatus = 0){
result_FindcustStatus = Status: N/A`n
}
}
if (RegExMatch(result_FindcustStatus,"(Status: 0|Status:0|Status:`n|Status: `n)", zero_result_FindcustStatus) = 1){
result_FindcustStatus = Status: N/A`n
}
if(RegExMatch(result_FindcustStatus,"\n") = 0){
	result_FindcustStatus = %result_FindcustStatus%`n
}



FindcustCommonLocal := RegExMatch(myString,"mU)(Common/Local:.+)\n", result_FindcustCommonLocal)
;MsgBox % result_FindcustCommonLocal
;check if the result is found in the text, if not add N/A
if (FindcustCommonLocal = 0){
FindcustCommonLocal := 
FindcustCommonLocal := RegExMatch(myString,"mU)(Common/Local:.+)\r", result_FindcustCommonLocal)
if (FindcustCommonLocal = 0){
result_FindcustCommonLocal = Common/Local: N/A`n
}
}
if (RegExMatch(result_FindcustCommonLocal,"(Common/Local: 0|Common/Local:0|Common/Local:`n|Common/Local: `n)", zero_result_FindcustCommonLocal) = 1){
result_FindcustCommonLocal = Common/Local: N/A`n
}
if(RegExMatch(result_FindcustCommonLocal,"\n") = 0){
	result_FindcustCommonLocal = %result_FindcustCommonLocal%`n
}


FindcustOPCO := RegExMatch(myString,"mU)(OPCO:.+)\n", result_FindcustOPCO)
;check if the result is found in the text, if not add N/A
if (FindcustOPCO = 0){
FindcustOPCO := 
FindcustOPCO := RegExMatch(myString,"mU)(OPCO:.+)\r", result_FindcustOPCO)
if (FindcustOPCO = 0){
result_FindcustOPCO = OPCO: N/A`n
}
}
if (RegExMatch(result_FindcustOPCO,"(OPCO: 0|OPCO:0|OPCO:`n|OPCO: `n)", zero_result_FindcustOPCO) = 1){
result_FindcustOPCO = OPCO: N/A`n
}
if(RegExMatch(result_FindcustOPCO,"\n") = 0){
	result_FindcustOPCO = %result_FindcustOPCO%`n
}



FindcustDRStartupPRIO := RegExMatch(myString,"mU)(DR startup PRIO:.+)\n", result_FindcustDRStartupPRIO)
;MsgBox % result_FindcustDRStartupPRIO
;check if the result is found in the text, if not add N/A
if (FindcustDRStartupPRIO = 0){
FindcustDRStartupPRIO := 
FindcustDRStartupPRIO := RegExMatch(myString,"mU)(DR startup PRIO:.+)\r", result_FindcustDRStartupPRIO)
if (FindcustDRStartupPRIO = 0){
result_FindcustDRStartupPRIO = DR startup PRIO: N/A`n
}
}
if (RegExMatch(result_FindcustDRStartupPRIO,"(DR startup PRIO: 0|DR startup PRIO:0|DR startup PRIO:`n|DR startup PRIO: `n)", zero_result_FindcustDRStartupPRIO) = 1){
result_FindcustDRStartupPRIO = DR startup PRIO: N/A`n
}
if(RegExMatch(result_FindcustDRStartupPRIO,"\n") = 0){
	result_FindcustDRStartupPRIO = %result_FindcustDRStartupPRIO%`n
}



FindcustSupplier := RegExMatch(myString,"mU)(Supplier:.+)\n", result_FindcustSupplier)
;MsgBox % result_FindcustSupplier
;check if the result is found in the text, if not add N/A
if (FindcustSupplier = 0){
FindcustSupplier := 
FindcustSupplier := RegExMatch(myString,"mU)(Supplier:.+)\r", result_FindcustSupplier)
if (FindcustSupplier = 0){
result_FindcustSupplier = Supplier: N/A`n
}
}
if (RegExMatch(result_FindcustSupplier,"(Supplier: 0|Supplier:0|Supplier:`n|Supplier: `n)", zero_result_FindcustSupplier) = 1){
result_FindcustSupplier = Supplier: N/A`n
}
if(RegExMatch(result_FindcustSupplier,"\n") = 0){
	result_FindcustSupplier = %result_FindcustSupplier%`n
}

FindcustFunctionSupportParty :=
FindcustFunctionSupportParty := RegExMatch(myString,"mU)(Function support party:.+)\n", result_FindcustFunctionSupportParty)
;MsgBox % result_FindcustFunctionSupportParty
;check if the result is found in the text, if not add N/A
if (FindcustFunctionSupportParty = 0){
FindcustFunctionSupportParty := 
FindcustFunctionSupportParty := RegExMatch(myString,"mU)(Function support party:.+)\r", result_FindcustFunctionSupportParty)
if (FindcustFunctionSupportParty = 0){
result_FindcustFunctionSupportParty = Function support party: N/A`n
found_FindcustFunctionSupportParty = 0 ; to be able to find Functional Support Party:
}
}
if (RegExMatch(result_FindcustFunctionSupportParty,"(Function support party: 0|Function support party:0|Function support party:`n|Function support party: `n)", zero_result_FindcustFunctionSupportParty) = 1){
result_FindcustFunctionSupportParty = Function support party: N/A`n
found_FindcustFunctionSupportParty = 0
}
if(RegExMatch(result_FindcustFunctionSupportParty,"\n") = 0){
	result_FindcustFunctionSupportParty = %result_FindcustFunctionSupportParty%`n
}

;MsgBox,,, %result_FindcustFunctionSupportParty%


FindcustFunctionalManagementOwner := RegExMatch(myString,"mU)(Functional Management/Owner:.+)\n", result_FindcustFunctionalManagementOwner)
;MsgBox % result_FindcustFunctionalManagementOwner
;check if the result is found in the text, if not add N/A
if (FindcustFunctionalManagementOwner = 0){
FindcustFunctionalManagementOwner := 
FindcustFunctionalManagementOwner := RegExMatch(myString,"mU)(Functional Management/Owner:.+)\r", result_FindcustFunctionalManagementOwner)
if (FindcustFunctionalManagementOwner = 0){
result_FindcustFunctionalManagementOwner = Functional Management/Owner: N/A`n
}
}
if (RegExMatch(result_FindcustFunctionalManagementOwner,"(Functional Management/Owner: 0|Functional Management/Owner:0|Functional Management/Owner:`n|Functional Management/Owner: `n)", zero_result_FindcustFunctionalManagementOwner) = 1){
result_FindcustFunctionSupportParty = Function support party: N/A`n
}
if(RegExMatch(result_FindcustFunctionalManagementOwner,"\n") = 0){
	result_FindcustFunctionalManagementOwner = %result_FindcustFunctionalManagementOwner%`n
}


FindcustTAM := RegExMatch(myString,"mU)(TAM:.+)\n", result_FindcustTAM)
;MsgBox % result_FindcustTAM
;check if the result is found in the text, if not add N/A
if (FindcustTAM = 0){
FindcustTAM := 
FindcustTAM := RegExMatch(myString,"mU)(TAM:.+)\r", result_FindcustTAM)
if (FindcustTAM = 0){
result_FindcustTAM = TAM: N/A`n
}
}
if (RegExMatch(result_FindcustTAM,"(TAM: 0|TAM:0|TAM: Basic / Standard / N/A|TAM:`n|TAM: `n)", zero_result_FindcustTAM) = 1){
result_FindcustTAM = TAM: N/A`n
}
if(RegExMatch(result_FindcustTAM,"\n") = 0){
	result_FindcustTAM = %result_FindcustTAM%`n
}


FindcustCommentsHashtag := RegExMatch(myString,"(###)(?<=###)(?s)(.*$)", result_FindcustCommentsHashtag)
;add a new line after ###
if (RegExMatch(result_FindcustCommentsHashtag, "###`n", result_temporaryVariable )){
;do nothing
}
else {
result_FindcustCommentsHashtagAddNewLinesAfterHasthags := RegExReplace(result_FindcustCommentsHashtag, "###", "###`n")
result_FindcustCommentsHashtag = %result_FindcustCommentsHashtagAddNewLinesAfterHasthags%
}
;check if the result is found in the text, if not add N/A
if (FindcustCommentsHashtag = 0){
result_FindcustCommentsHashtag = ###`n
}

;Check for Contact person (not a part of the Org.Remarks, but will paste it after comments (###))
FindcustContactPerson := RegExMatch(myString,"mU)(Contact person:.+)\n", result_FindcustContactPerson)

;Append all the results one to another to match the right Org.Remarks format

newOrgRemarks = %result_FindcustNamingConvention%%result_FindcustLandscape%%result_FindcustCategory%%result_FindcustStatus%%result_FindcustCommonLocal%%result_FindcustOPCO%%result_FindcustDRStartupPRIO%%result_FindcustSupplier%%result_FindcustFunctionSupportParty%%result_FindcustFunctionalManagementOwner%%result_FindcustTAM%%result_FindcustCommentsHashtag%%result_FindcustContactPerson%


;Paste new/converted/fixed Org.Remarks into the text field on the right
GuiControl,, NewOrgRemarksTextField, %newOrgRemarks%


return

CopyNewOrgRemarksToClipboardButton:
;Empty clipboard
Clipboard =

GuiControl, focus, NewOrgRemarksTextField
Send, ^a
Send, ^c
ClipWait, 0.1
;Check if new Org.Remarks have been copied to the clipboard, of not give a MsgBox
if ErrorLevel{
MsgBox, 0,%version%, You did not validate the Org.Remarks.`n`nValidated Org.Remarks text field is empty.
return
}
else{
Send, {Down}
}

return

;ensures complete closure of the script when X is pressed in the window
GuiClose:
ExitApp

