
    ;======= GLOBAL VARIABLES =======
	
	
	
	;=== ProScript Variables ===

; ProScript Search Box 
global proScriptSearchBoxX1 := -1534
global proScriptSearchBoxY1 := -8
global proScriptSearchBoxX2 := -282
global proScriptSearchBoxY2 := 1064




	;=== Siebel Various Variables ===
	
; Patient Account Search Box
global siebelPatientSearchBoxX1 := 45
global siebelPatientSearchBoxY1 := 628
global siebelPatientSearchBoxX2 := 71
global siebelPatientSearchBoxY2 := 661


; Duplicate Search Box Coordinates
global siebelDuplicateBoxPosX1 := 1260
global siebelDuplicateBoxPosY1 := 700
global siebelDuplicateBoxPosX2 := 1912
global siebelDuplicateBoxPosY2 := 1077


; Auto GAS Selected Button Search Box
global siebelAutoGasSelectedButtonBoxPosX1 := 1567
global siebelAutoGasSelectedButtonBoxPosY1 := 532
global siebelAutoGasSelectedButtonBoxPosX2 := 1874
global siebelAutoGasSelectedButtonBoxPosY2 := 672



	
	;=== Siebel Activity & Template Window Variables ===

; Template Window Seach Box
global siebelTemplateBoxPosX1 := 528
global siebelTemplateBoxPosY1 := 363
global siebelTemplateBoxPosX2 := 1395
global siebelTemplateBoxPosY2 := 776


; Activity Tab Search Box 
global siebelActivityTabSearchBoxX1 := 960
global siebelActivityTabSearchBoxY1 := 520
global siebelActivityTabSearchBoxX2 := 1884
global siebelActivityTabSearchBoxY2 := 622


; Activity Tab Duplicate Search Box
global siebelActivityTabDuplicateBoxPosX1 := 84
global siebelActivityTabDuplicateBoxPosY1 := 367
global siebelActivityTabDuplicateBoxPosX2 := 1897
global siebelActivityTabDuplicateBoxPosY2 := 518


; Activity New Button Search Box
global siebelActivityNewButtonSearchBoxX1 := 565
global siebelActivityNewButtonSearchBoxY1 := 570
global siebelActivityNewButtonSearchBoxX2 := 613
global siebelActivityNewButtonSearchBoxY2 := 593


; Newly Created Activity Search Box
global siebelNewlyCreatedActivitySearchBoxX1 := 492
global siebelNewlyCreatedActivitySearchBoxY1 := 569
global siebelNewlyCreatedActivitySearchBoxX2 := 1265
global siebelNewlyCreatedActivitySearchBoxY2 :=650




	;=== Siebel Safety Functions Variables ===

; Siebel Check If Current window Is Correct AP Lines Search Box
global siebelCheckIfCurrentWindowIsCorrectAPLinesBoxX1 := 490
global siebelCheckIfCurrentWindowIsCorrectAPLinesBoxY1 := 518
global siebelCheckIfCurrentWindowIsCorrectAPLinesBoxX2 := 983
global siebelCheckIfCurrentWindowIsCorrectAPLinesBoxY2 := 574


; Siebel Check If Current window Is Correct Contacts Search Box
global siebelCheckIfCurrentWindowIsCorrectContactsBoxX1 := 6
global siebelCheckIfCurrentWindowIsCorrectContactsBoxY1 := 114
global siebelCheckIfCurrentWindowIsCorrectContactsBoxX2 := 426
global siebelCheckIfCurrentWindowIsCorrectContactsBoxY2 := 156


; Siebel Check If Current window Is Correct Stock & Check Search Box
global siebelCheckIfCurrentWindowIsCorrectStockCheckBoxX1 := 6
global siebelCheckIfCurrentWindowIsCorrectStockCheckBoxY1 := 114
global siebelCheckIfCurrentWindowIsCorrectStockCheckBoxX2 := 888
global siebelCheckIfCurrentWindowIsCorrectStockCheckBoxY2 := 176









#UseHook
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen

^p::
if RegExMatch(Clipboard, "(\QEPS\E.{6}\-.{6}\-.{6})", OutputVar)
{
Msgbox 445544
Return
}
else
{
Msgbox l675757
Return
}
Return





















































SafePaste() 
{
    ; A way of pasting that only returns control when the paste is complete
    ; by jeeswg
    ; See https://www.autohotkey.com/boards/viewtopic.php?p=271514&sid=f898e28c59efcb6871c1dff403e663dd#p271517
    ; the point of this is that with a simple Ctrl + v, you don't know when the pasting is complete,
    ; so if you immediately reload the Clipboard, the new text may end up getting pasted...
ControlGetFocus, vCtlClassNN, A
ControlGet, hCtl, Hwnd,, % vCtlClassNN, A
SendMessage, 0x302,,,, % "ahk_id " hCtl ;WM_PASTE := 0x302
}











  ; ==================================================================== EXITS SCRIPT DOCUMENT ==========================================================

; This script exits the AHK application



#UseHook
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen

#x::ExitApp  ; Win+X
return





	; ==================================================================== RELOAD WHOLE SCRIPT DOCUMENT ==============================================================

; This script reloads the whole script document in case any of the scripts gets stuck

#UseHook
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen

F12::
Reload

Return









; === IMPROVEMENTS !!!!!!!!!!!!!!!!!! === 
; https://www.autohotkey.com/boards/viewtopic.php?t=6413

; #NoEnv
; #MaxHotkeysPerInterval 99000000
; #HotkeyInterval 99000000
; #KeyHistory 0
; ListLines Off
; Process, Priority, , A
; SetBatchLines, -1
; SetKeyDelay, -1, -1
; SetMouseDelay, -1
; SetDefaultMouseSpeed, 0
; SetWinDelay, -1
; SetControlDelay, -1
; SendMode Input
; DllCall("ntdll\ZwSetTimerResolution","Int",5000,"Int",1,"Int*",MyCurrentTimerResolution) ;setting the Windows Timer Resolution to 0.5ms, THIS IS A GLOBAL CHANGE


















removeCommentsField()
{
SetDefaultMouseSpeed, 0
MouseMove 879, 337
MouseClick
MouseClick
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SetDefaultMouseSpeed, 2
}



removeDescriptionField()
{
SetDefaultMouseSpeed, 0
MouseMove 882, 275
MouseClick
MouseClick
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SetDefaultMouseSpeed, 2
}





setStatusDone()
{
SetDefaultMouseSpeed, 0
MouseMove 935, 380
Sleep 50
MouseClick
Sleep 50
MouseMove 936, 409
Sleep 50
MouseClick
SetDefaultMouseSpeed, 2
}



setTypeOpenedInError()
{
SetDefaultMouseSpeed, 0
MouseMove 715, 279
Sleep 50
MouseClick
Sleep 50
SendInput {Raw}Opened in Error
Sleep 50
SendInput {Enter}
SetDefaultMouseSpeed, 2
}























#UseHook
SetBatchLines, -1
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen

^Delete::
BlockInput, MouseMove
Clipboard := Clipboard
Loop, 50
{
Sleep 50
var++
Gui,help:Add, Text,     , %var%
Gui,help:+toolwindow
Gui,help:Show
sleep 100
Gui, help: Destroy







siebelCheckPerscriptionFieldIfPXIsPastedIn()




















}

BlockInput, MouseMoveOff
MsgBox Meczyk wygrany, kurczak podany
Return ; Script Run Finished




















































































	;============================================== FUNCTIONS =====================================================


	; ====================== PROSCRIPT FUNCTIONS ========================


lookForNHSOnPX()
{
WinActivate, ahk_class TscShellContainerClass ; Opens up ProScript by its class
Loop, 5
{
global proScriptSearchBoxX1 
global proScriptSearchBoxY1 
global proScriptSearchBoxX2 
global proScriptSearchBoxY2 
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *100 %A_ScriptDir%\Images\NHS_Number_Green.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 110
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 0.2
Sleep 100
SetDefaultMouseSpeed, 2
if RegExMatch(Clipboard, "(\d{10})", OutputVar)
{
Return
}
else
{
Sleep 100
}
}
else if (ErrorLevel != 0)
{
Break
}
}
Loop, 5
{
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *100 %A_ScriptDir%\Images\NHS_Number_Purple.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 110
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 0.2
Sleep 100
SetDefaultMouseSpeed, 2
if RegExMatch(Clipboard, "(\d{10})", OutputVar)
{
Return
}
else
{
Sleep 100
}
}
else if (ErrorLevel != 0)
{
Break
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, "Script Failed - Please try again or copy NHS Manually. Press F12 to Exit the error message"
SetDefaultMouseSpeed, 2
Reload
}




checkForMedicationOnPX()
{
WinActivate, ahk_class TscShellContainerClass ; Opens up ProScript by its class
Loop, 5
{
global proScriptSearchBoxX1 
global proScriptSearchBoxY1 
global proScriptSearchBoxX2 
global proScriptSearchBoxY2 
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *100 %A_ScriptDir%\Images\Prescribed_Medication_Green.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 15
VarPosY := OutputVarY + 25
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 50
SetDefaultMouseSpeed, 2
if (Clipboard = "")
{
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 200
}
else
{
Return
}
}
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *100 %A_ScriptDir%\Images\Prescribed_Medication_Purple.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 15
VarPosY := OutputVarY + 25
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 50
SetDefaultMouseSpeed, 2
if (Clipboard = "")
{
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 200
}
else
{
Return
}
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, Could not copy the product! Press F12 to exit
Reload
}




selectsPXNumberInProScriptAndCopyIt()
{
WinActivate, ahk_class TscShellContainerClass ; Opens up ProScript by its class
Loop, 5
{
global proScriptSearchBoxX1 
global proScriptSearchBoxY1 
global proScriptSearchBoxX2 
global proScriptSearchBoxY2 
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *170 %A_ScriptDir%\Images\Script_Id.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 130
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 50
SetDefaultMouseSpeed, 2
if RegExMatch(Clipboard, "(.{6}-.{6}-.{6})", OutputVar)
{
Return
}
else 
{
Sleep 100
}
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, "Could not find the PX Number. Press F12 to Exit the error message"
BlockInput, MouseMoveOff
Reload
}




selectsPXNumberInProScriptAndCopyItForDuplicate()
{
WinActivate, ahk_class TscShellContainerClass ; Opens up ProScript by its class
Loop, 5
{
global proScriptSearchBoxX1 
global proScriptSearchBoxY1 
global proScriptSearchBoxX2 
global proScriptSearchBoxY2 
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *170 %A_ScriptDir%\Images\Script_Id.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 130
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 50
SetDefaultMouseSpeed, 2
Return
}
else
{
Sleep 100
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MsgBox, 4096, Error, "Could not find the PX Number. Press F12 to Exit the error message"
BlockInput, MouseMoveOff
Reload
}




checkIfPxIsScrolled()
{
WinActivate, ahk_class TscShellContainerClass ; Opens up ProScript by its class
Loop
{
global proScriptSearchBoxX1 
global proScriptSearchBoxY1 
global proScriptSearchBoxX2 
global proScriptSearchBoxY2 
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *25 %A_ScriptDir%\Images\PX_Scroll_Up_Arrow.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 7
VarPosY := OutputVarY + 19
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
Sleep 50
MouseClick
Sleep 50
MouseClick
Sleep 50
MouseClick
Sleep 50
Return
}
else if (ErrorLevel != 0)
{
;
}
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *25 %A_ScriptDir%\Images\PX_Scroll_Up_Arrow_2.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 7
VarPosY := OutputVarY + 19
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
Sleep 50
MouseClick
Sleep 50
MouseClick
Sleep 50
MouseClick
Sleep 50
Return
}
else if (ErrorLevel != 0)
{
Break
}
}
}




checkForExpiryDateOnPXFirstLine()
{
WinActivate, ahk_class TscShellContainerClass ; Opens up ProScript by its class
Loop, 2
{
global proScriptSearchBoxX1 
global proScriptSearchBoxY1 
global proScriptSearchBoxX2 
global proScriptSearchBoxY2 
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *100 %A_ScriptDir%\Images\PX_Expiry_Date_Green.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 120
VarPosY := OutputVarY + 20
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 50
SetDefaultMouseSpeed, 2
if (Clipboard = "")
{
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 200
}
else
{
Return
}
}
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *100 %A_ScriptDir%\Images\PX_Expiry_Date_Purple.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 120
VarPosY := OutputVarY + 20
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 50
SetDefaultMouseSpeed, 2
if (Clipboard = "")
{
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 200
}
else
{
Return
}
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, Could not find the PX! Press F12 to exit
Reload
}




checkForExpiryDateOnPXSecondLine()
{
WinActivate, ahk_class TscShellContainerClass ; Opens up ProScript by its class
Loop, 2
{
global proScriptSearchBoxX1 
global proScriptSearchBoxY1 
global proScriptSearchBoxX2 
global proScriptSearchBoxY2 
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *100 %A_ScriptDir%\Images\PX_Expiry_Date_Green.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 120
VarPosY := OutputVarY + 37
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 50
SetDefaultMouseSpeed, 2
if (Clipboard = "")
{
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 200
}
else
{
Return
}
}
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY1%, %proScriptSearchBoxY2%, *100 %A_ScriptDir%\Images\PX_Expiry_Date_Purple.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 120
VarPosY := OutputVarY + 37
MouseMove %VarPosX%, %VarPosY%
Sleep 200
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 50
SetDefaultMouseSpeed, 2
if (Clipboard = "")
{
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput {Raw}c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 200
}
else
{
Return
}
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, Could not find the PX! Press F12 to exit
Reload
}
















	; ====================== SIEBEL FUNCTIONS ========================


	; === Activity Coordinates ===


siebelActivityTab()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop, 10
{
global siebelActivityTabSearchBoxX1 
global siebelActivityTabSearchBoxY1 
global siebelActivityTabSearchBoxX2 
global siebelActivityTabSearchBoxY2
ImageSearch OutputVarX, OutputVarY, %siebelActivityTabSearchBoxX1%, %siebelActivityTabSearchBoxY1%, %siebelActivityTabSearchBoxX2%, %siebelActivityTabSearchBoxY2%, *60 %A_ScriptDir%\Images\Activities_Tab.png
if (ErrorLevel = 0)
{
VarPosX := OutputVarX + 30
VarPosY := OutputVarY + 10
SetDefaultMouseSpeed, 0
MouseMove %VarPosX%, %VarPosY%
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
sleep 50
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
SetDefaultMouseSpeed, 2
MsgBox, 4096, Error, "Script Failed - Could not find Activity tab. Press F12 to Exit the error message"
Reload
}




siebelActivityNewButton()
{
Loop, 10
{
global siebelActivityNewButtonSearchBoxX1 
global siebelActivityNewButtonSearchBoxY1 
global siebelActivityNewButtonSearchBoxX2 
global siebelActivityNewButtonSearchBoxY2
ImageSearch OutputVarX, OutputVarY, %siebelActivityNewButtonSearchBoxX1%, %siebelActivityNewButtonSearchBoxY1%, %siebelActivityNewButtonSearchBoxX2%, %siebelActivityNewButtonSearchBoxY2%, *100, *TransBlack %A_ScriptDir%\Images\New_Activity_Button.png
if (ErrorLevel = 0)
{
VarPosX := OutputVarX + 15
VarPosY := OutputVarY + 5
SetDefaultMouseSpeed, 0
MouseMove %VarPosX%, %VarPosY%
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
sleep 50
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
SetDefaultMouseSpeed, 2
MsgBox, 4096, Error, "Script Failed - Could not find New Activity button. Press F12 to Exit the error message"
Reload
}




siebelActivityWhiteSpace()
{
SetDefaultMouseSpeed, 0

	MouseMove 776, 483 ; Moves the cursor away from activities field after creating new activity

MouseClick
Sleep 50
MouseClick
Sleep 50
MouseClick
SetDefaultMouseSpeed, 2
}




clickOnNewlyCreatedActivity()
{
Loop, 20
{
global siebelNewlyCreatedActivitySearchBoxX1 
global siebelNewlyCreatedActivitySearchBoxY1 
global siebelNewlyCreatedActivitySearchBoxX2 
global siebelNewlyCreatedActivitySearchBoxY2
ImageSearch OutputVarX, OutputVarY, %siebelNewlyCreatedActivitySearchBoxX1%, %siebelNewlyCreatedActivitySearchBoxY1%, %siebelNewlyCreatedActivitySearchBoxX2%, %siebelNewlyCreatedActivitySearchBoxY2%, *50, *TransBlack %A_ScriptDir%\Images\New_Activity_Other.png
if (ErrorLevel = 0)
{
VarPosX := OutputVarX + 15
VarPosY := OutputVarY + 5
SetDefaultMouseSpeed, 0
MouseMove %VarPosX%, %VarPosY%
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
sleep 50
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
SetDefaultMouseSpeed, 2
MsgBox, 4096, Error, "Script Failed - Could not find new activity created. Press F12 to Exit the error message"
Reload
}




siebelActivityDescriptionField()
{
SetDefaultMouseSpeed, 0

	MouseMove 872, 268 ; Moves to description field in New Activity
	
SetDefaultMouseSpeed, 2
}




siebelActivityCommentField()
{
SetDefaultMouseSpeed, 0

	MouseMove 865, 336 ; Moves to comments field in New Activity

SetDefaultMouseSpeed, 2
}




siebelSiebelOrderNoField()
{
SetDefaultMouseSpeed, 0

	MouseMove 1376, 404 ; Clicks on Siebel Order No Field in Shortfall Activity and pastes in Siebel Order Number

Sleep 100
MouseClick
Sleep 50
MouseClick
Sleep 50
Send {Ctrl Down}
SendInput {Raw}v
Send {Ctrl Up}
ClipWait, 1
Sleep 100
MouseClick
SendInput ^{End}
SendInput +{Home}
Sleep 50
Send {Ctrl Down}
SendInput {Raw}c
Send {Ctrl Up}
ClipWait, 1
Sleep 50
SetDefaultMouseSpeed, 2
Loop, 5
{
if RegExMatch(Clipboard, "(\d{1}-\d{10})", OutputVar)
{
Return
}
else
{
MouseClick
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput {backspace}
Sleep 50
Send {Ctrl Down}
SendInput {Raw}v
Send {Ctrl Up}
ClipWait, 1
Sleep 50
MouseClick
SendInput ^{End}
SendInput +{Home}
Sleep 50
Send {Ctrl Down}
SendInput {Raw}c
Send {Ctrl Up}
ClipWait, 1
Sleep 50
}
}
}




siebelSourceDownArrow()
{
SetDefaultMouseSpeed, 0

	MouseMove 715, 327 ; Moves to source field down arrow in a newly created Activity and clicks it
	
MouseClick
Sleep 100
SetDefaultMouseSpeed, 2
}




siebelSelectSourceMED()
{
SendInput {Raw}Medication
SendInput {Enter}
}




siebelSelectSourceDUP()
{
SendInput {Raw}Dupliate
SendInput {Enter}
}




siebelSourcePXService()
{
SendInput {Raw}Px Services
SendInput {Enter}
}




siebelSelectStatusDownArrow()
{
SetDefaultMouseSpeed, 0

	MouseMove 939, 383 ; Moves to status field down arrow in a newly created Activity and clicks it
	
MouseClick
SetDefaultMouseSpeed, 2
}




siebelSelectStatusDone()
{
SendInput {Raw}Done
SendInput {Enter}
}




siebelSubTypeDownArrow()
{
SetDefaultMouseSpeed, 0

	MouseMove 714, 302 ; Moves over sub-type drop arrow in Shortfall activity and clicks it

MouseClick
SetDefaultMouseSpeed, 2
}
















	; === Template Window Coordinates ===	


checkForTemplateWindow()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop, 50
{
global siebelTemplateBoxPosX1
global siebelTemplateBoxPosY1
global siebelTemplateBoxPosX2
global siebelTemplateBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelTemplateBoxPosX1%, %siebelTemplateBoxPosY1%, %siebelTemplateBoxPosX2%, %siebelTemplateBoxPosY2%, *150 %A_ScriptDir%\Images\Siebel_Template_Window.png
if (ErrorLevel = 0)
{
Return
}
else
{
Sleep 50
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, "Cannot find the Template Window. Press F12 to Exit the error message"
SetDefaultMouseSpeed, 2
Reload
}




checkIfTemplateWindowIsScrolled()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop, 1
{
global siebelTemplateBoxPosX1
global siebelTemplateBoxPosY1
global siebelTemplateBoxPosX2
global siebelTemplateBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelTemplateBoxPosX1%, %siebelTemplateBoxPosY1%, %siebelTemplateBoxPosX2%, %siebelTemplateBoxPosY2%, *25, %A_ScriptDir%\Images\Siebel_Template_Window_Scroll_Arrow.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 7
VarPosY := OutputVarY + 43
MouseMove %VarPosX%, %VarPosY%
MouseClick
Sleep 50
MouseClick
Sleep 50
MouseClick
Sleep 50
Return
}
else if (ErrorLevel != 0)
{
Sleep 50
}
}
}




searchForEpsShortfallTemplate()
{
Loop, 5
{
global siebelTemplateBoxPosX1
global siebelTemplateBoxPosY1
global siebelTemplateBoxPosX2
global siebelTemplateBoxPosY2  
ImageSearch OutputVarX, OutputVarY, %siebelTemplateBoxPosX1%, %siebelTemplateBoxPosY1%, %siebelTemplateBoxPosX2%, %siebelTemplateBoxPosY2%, *150 %A_ScriptDir%\Images\EPS_Shortfall_Template_White.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 400
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
Sleep 50
}
}
Loop, 5
{
ImageSearch OutputVarX, OutputVarY, %siebelTemplateBoxPosX1%, %siebelTemplateBoxPosY1%, %siebelTemplateBoxPosX2%, %siebelTemplateBoxPosY2%, *150 %A_ScriptDir%\Images\EPS_Shortfall_Template_Yellow.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 400
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
Sleep 50
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, "Cannot find the EPS Shortfall Template. Press F12 to Exit the error message"
SetDefaultMouseSpeed, 2
Reload
}




searchForEpsReturnTemplate()
{
Loop, 5
{
global siebelTemplateBoxPosX1
global siebelTemplateBoxPosY1
global siebelTemplateBoxPosX2
global siebelTemplateBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelTemplateBoxPosX1%, %siebelTemplateBoxPosY1%, %siebelTemplateBoxPosX2%, %siebelTemplateBoxPosY2%, *150 %A_ScriptDir%\Images\EPS_Return_Template_Yellow.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 400
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
Sleep 50
}
}
Loop
{
ImageSearch OutputVarX, OutputVarY, %siebelTemplateBoxPosX1%, %siebelTemplateBoxPosY1%, %siebelTemplateBoxPosX2%, %siebelTemplateBoxPosY2%, *150 %A_ScriptDir%\Images\EPS_Return_Template_White.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 400
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
Return
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, "Cannot find the EPS Return Template. Press F12 to Exit the error message"
SetDefaultMouseSpeed, 2
Reload
}




searchForEpsPxServiceReturnTemplate()
{
Loop, 5
{
global siebelTemplateBoxPosX1
global siebelTemplateBoxPosY1
global siebelTemplateBoxPosX2
global siebelTemplateBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelTemplateBoxPosX1%, %siebelTemplateBoxPosY1%, %siebelTemplateBoxPosX2%, %siebelTemplateBoxPosY2%, *150 %A_ScriptDir%\Images\EPS_Return_PXService_Template_White.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 400
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
Sleep 50
}
}
Loop, 5
{
ImageSearch OutputVarX, OutputVarY, %siebelTemplateBoxPosX1%, %siebelTemplateBoxPosY1%, %siebelTemplateBoxPosX2%, %siebelTemplateBoxPosY2%, *150 %A_ScriptDir%\Images\EPS_Return_PXService_Template_Yellow.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 400
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
Sleep 50
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, "Cannot find the EPS Return PX Service Template. Press F12 to Exit the error message"
SetDefaultMouseSpeed, 2
Reload
}





clickOkButtonInTemplateWindow()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop, 10
{
global siebelTemplateBoxPosX1
global siebelTemplateBoxPosY1
global siebelTemplateBoxPosX2
global siebelTemplateBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelTemplateBoxPosX1%, %siebelTemplateBoxPosY1%, %siebelTemplateBoxPosX2%, %siebelTemplateBoxPosY2%, *120 %A_ScriptDir%\Images\Ok_Button_Template_Window.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 15
VarPosY := OutputVarY + 10
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
MouseClick
Sleep 50
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
Sleep 50
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, "Cannot find OK button in template window. Press F12 to Exit the error message"
SetDefaultMouseSpeed, 2
Reload
}
















	; === Duplicate Coordinates ===


siebelClickOnNewInNotes()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
SetDefaultMouseSpeed, 0

	MouseMove 1422, 899 ; Moves over New button in Active Notes on AP Order Lines
	
MouseClick ; Clicks on New to create a New Active Note
SetDefaultMouseSpeed, 2
}




siebelActivityTabDuplicate()
{
Loop, 5
{
global siebelActivityTabDuplicateBoxPosX1
global siebelActivityTabDuplicateBoxPosY1
global siebelActivityTabDuplicateBoxPosX2
global siebelActivityTabDuplicateBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelActivityTabDuplicateBoxPosX1%, %siebelActivityTabDuplicateBoxPosY1%, %siebelActivityTabDuplicateBoxPosX2%, %siebelActivityTabDuplicateBoxPosY2%, *60 %A_ScriptDir%\Images\Activities_Tab.png
if (ErrorLevel = 0)
{
VarPosX := OutputVarX + 30
VarPosY := OutputVarY + 10
SetDefaultMouseSpeed, 0
MouseMove %VarPosX%, %VarPosY%
MouseClick
SetDefaultMouseSpeed, 2
Return
}
else if (ErrorLevel != 0)
{
sleep 50
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
SetDefaultMouseSpeed, 2
MsgBox, 4096, Error, "Script Failed - Could not find Activity tab. Press F12 to Exit the error message"
Reload
}




siebelFlagContactTick()
{
Loop, 5
{
global siebelDuplicateBoxPosX1
global siebelDuplicateBoxPosY1
global siebelDuplicateBoxPosX2
global siebelDuplicateBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelDuplicateBoxPosX1%, %siebelDuplicateBoxPosY1%, %siebelDuplicateBoxPosX2%, %siebelDuplicateBoxPosY2%, *120 %A_ScriptDir%\Images\Flag_Contact_Label.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 14
VarPosY := OutputVarY + 30
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
Sleep 50
Break
}
else
{
Sleep 50
}
}
Loop, 2
{
ImageSearch OutputVarX, OutputVarY, %siebelDuplicateBoxPosX1%, %siebelDuplicateBoxPosY1%, %siebelDuplicateBoxPosX2%, %siebelDuplicateBoxPosY2%, *120 %A_ScriptDir%\Images\Flag_Contact_Tick_Empty.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 7
VarPosY := OutputVarY + 7
MouseMove %VarPosX%, %VarPosY%
Sleep 100
MouseClick
Sleep 100
Break
}
else if (ErrorLevel != 0)
{
Sleep 50
}
}
Loop
{
ImageSearch OutputVarX, OutputVarY, %siebelDuplicateBoxPosX1%, %siebelDuplicateBoxPosY1%, %siebelDuplicateBoxPosX2%, %siebelDuplicateBoxPosY2%, *120 %A_ScriptDir%\Images\Flag_Contact_Ticked.png
if (ErrorLevel = 0)
{
Return
}
else
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 7
VarPosY := OutputVarY + 7
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
Return
}
}
}




siebelDescriptionFieldInActiveNotes()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop, 5
{
global siebelDuplicateBoxPosX1
global siebelDuplicateBoxPosY1
global siebelDuplicateBoxPosX2
global siebelDuplicateBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelDuplicateBoxPosX1%, %siebelDuplicateBoxPosY1%, %siebelDuplicateBoxPosX2%, %siebelDuplicateBoxPosY2%, *120 %A_ScriptDir%\Images\Description_Field_Active_Notes.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 12
VarPosY := OutputVarY + 30
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
MouseClick
Sleep 100
Break
}
else
{
Sleep 50
}
}
}




siebelTypeFieldInActiveNotes()
{
Loop, 5
{
global siebelDuplicateBoxPosX1
global siebelDuplicateBoxPosY1
global siebelDuplicateBoxPosX2
global siebelDuplicateBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelDuplicateBoxPosX1%, %siebelDuplicateBoxPosY1%, %siebelDuplicateBoxPosX2%, %siebelDuplicateBoxPosY2%, *100 %A_ScriptDir%\Images\Type_Field_Active_Notes.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 6
VarPosY := OutputVarY + 28
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
Sleep 50
Break
}
else
{
Sleep 50
}
}
Loop, 5
{
ImageSearch OutputVarX, OutputVarY, %siebelDuplicateBoxPosX1%, %siebelDuplicateBoxPosY1%, %siebelDuplicateBoxPosX2%, %siebelDuplicateBoxPosY2%, *100 %A_ScriptDir%\Images\Type_Field_Active_Notes_Down_Arrow.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 6
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
Sleep 50
SendInput {Raw}Duplicate Prescription Alert
Sleep 50
SendInput {Enter}
Sleep 50
Break
}
else
{
Sleep 50
}
}
}

















	; === GAS Related (Buttons, PX Field) Coordinates ===


siebelPerscriptionNoField()
{
SetDefaultMouseSpeed, 0

	MouseMove 672, 241 ; Perscription No Field coordinates in Siebel, after GASing the order
	
SetDefaultMouseSpeed, 2
}




siebelPerscriptionNoFieldClearAndPastePX()
{
MouseClick
Sleep, 50
MouseClick
Sleep 50
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
Sleep 50
SendInput {Raw}EPS
Send {Ctrl Down}
SendInput {Raw}v
Send {Ctrl Up}
ClipWait 1
Sleep 200
}




siebelCheckPerscriptionFieldIfPXIsPastedIn()
{
global proScriptSearchBoxX1 
global proScriptSearchBoxY1 
global proScriptSearchBoxX2 
global proScriptSearchBoxY2 
MouseClick
SendInput ^{End}
SendInput +{Home}
Sleep 50
Send {Ctrl Down}
SendInput c
Send {Ctrl Up}
ClipWait, 1
Sleep 100
Loop, 5
{
if RegExMatch(Clipboard, "(\QEPS\E.{6}\-.{6}\-.{6})", OutputVar)
{
Return
}
else
{
ImageSearch OutputVarX, OutputVarY, %proScriptSearchBoxX1%, %proScriptSearchBoxY1%, %proScriptSearchBoxX2%, %proScriptSearchBoxY2%, *150 %A_ScriptDir%\Images\Script_Id.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 130
VarPosY := OutputVarY + 6
MouseMove %VarPosX%, %VarPosY%
Sleep 50
MouseClick
Sleep 50
SendInput {Ctrl Down}
SendInput a
SendInput {Ctrl Up}
Sleep 50
SendInput {Ctrl Down}
SendInput c
SendInput {Ctrl Up}
ClipWait, 1
Sleep 50
MouseMove 672, 241 ; Perscription No Field coordinates in Siebel, after GASSing the order   
MouseClick
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
SendInput {Enter}
Sleep 50
MouseClick
Sleep 50
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {backspace}
SendInput {backspace}
Sleep 50
SendInput {Raw}EPS
Send {Ctrl Down}
SendInput v
Send {Ctrl Up}
ClipWait 1
Sleep 100
MouseClick
SendInput ^{End}
SendInput +{Home}
Sleep 50
Send {Ctrl Down}
SendInput c
Send {Ctrl Up}
ClipWait, 1
Sleep 100
}
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
SetDefaultMouseSpeed, 2
MsgBox, 4096, Error, "Script Failed - Please copy and paste the PX number manually. Press F12 to Exit the error message"
Reload
}




siebelAutoGASAllButton()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
SetDefaultMouseSpeed, 0

	MouseMove 1609, 600 ; "Auto GAS All" button Coordinates (Siebel)

MouseClick

SetDefaultMouseSpeed, 2
}




siebelAutoGASSelectedButton()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
SetDefaultMouseSpeed, 0
MouseMove 1790, 483 ; Moves over to whitespace to click and activate Auto GAS Selected button
MouseClick
MouseClick
Loop, 30
{
global siebelAutoGasSelectedButtonBoxPosX1
global siebelAutoGasSelectedButtonBoxPosY1
global siebelAutoGasSelectedButtonBoxPosX2
global siebelAutoGasSelectedButtonBoxPosY2 
ImageSearch OutputVarX, OutputVarY, %siebelAutoGasSelectedButtonBoxPosX1%, %siebelAutoGasSelectedButtonBoxPosY1%, %siebelAutoGasSelectedButtonBoxPosX2%, %siebelAutoGasSelectedButtonBoxPosY2%, *150 %A_ScriptDir%\Images\Auto_GAS_Selected_Button.png
if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
VarPosX := OutputVarX + 45
VarPosY := OutputVarY + 5
MouseMove %VarPosX%, %VarPosY%
MouseClick
MouseClick
Return
}
else if (ErrorLevel != 0)
{
sleep 50
}
}
SetDefaultMouseSpeed, 2
BlockInput, MouseMoveOff
MsgBox, 4096, Error, Could not find the Auto GAS Selected button. Please run the script one more time. Press F12 to Exit the error message.
Reload
}
















	; === NHS Copy & Paste Coordinates ===	


checkForPatientNameToAppear()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop, 100
{
global siebelPatientSearchBoxX1 
global siebelPatientSearchBoxY1 
global siebelPatientSearchBoxX2
global siebelPatientSearchBoxY2
PixelSearch, OutputVarX, OutputVarY, %siebelPatientSearchBoxX1%, %siebelPatientSearchBoxY1%, %siebelPatientSearchBoxX2%, %siebelPatientSearchBoxY2%, 6666CC, 210, Fast ; Last Name Field Coordinates (Search Box - Siebel)
if (ErrorLevel != 0)
{
Sleep 20
}
else if (ErrorLevel = 0)
{
SetDefaultMouseSpeed, 0
MouseMove %OutputVarX%, %OutputVarY%
MouseClick
SetDefaultMouseSpeed, 2
Return
}
}
Return
}




siebelClickReset()
{
SetDefaultMouseSpeed, 0

	MouseMove 72, 557 ; Reset Button in Search Side Tab in Siebel

MouseClick

SetDefaultMouseSpeed, 2
}




siebelPasteNHSNumber()
{
SetDefaultMouseSpeed, 0

	MouseMove 91, 439 ; NHS Number Field coordinates in Search Side Tab in Siebel

MouseClick
MouseClick
Sleep 50
SendInput {Backspace}
SendInput {Backspace}
Sleep 50
SendInput ^{End}
SendInput +{Home}
SendInput {Backspace}
SendInput {Backspace}
SendInput ^{End}
SendInput +{Home}
SendInput {Backspace}
SendInput {Backspace}
Sleep 50
SendInput {Ctrl Down}
SendInput v
SendInput {Ctrl Up}
Sleep 50
SendInput {Enter}
Sleep 50
SetDefaultMouseSpeed, 2
}
















	; === Stock & Check Coordinates ===	


siebelStockAndCheckOrderStatusChanges()
{
SetDefaultMouseSpeed, 0
OrderStatusX := 752 ; X coordinate of the "Order Status" down arrow in new Order Window
OrderStatusY := 295 ; Y coordinate of the "Order Status" down arrow in new Order Window
MouseMove %OrderStatusX%, %OrderStatusY% ; "Order Status" downarrow coordinates in new Order window
MouseClick
Sleep 50
SendInput {Raw}Pending
Sleep 50
SendInput {Enter}
MouseClick
Sleep 50
MouseMove 851, 704 ; Stock & Check button Coordinates in  new Order Window
MouseClick
SendInput {Enter}
SendInput {Enter}
SendInput {Enter}
MouseMove %OrderStatusX%, %OrderStatusY% ; "Order Status" downarrow coordinates in new Order window
MouseClick
Sleep 50
SendInput {Raw}Open
Sleep 50
SendInput {Enter}
MouseClick
Sleep 50
MouseMove %OrderStatusX%, %OrderStatusY% ; "Order Status" downarrow coordinates in new Order window
MouseClick
Sleep 50
SendInput {Raw}Awaiting Payment
Sleep 50
SendInput {Enter}
MouseClick
Sleep 50
SetDefaultMouseSpeed, 2
}




siebelSourceAndOriginFields()
{
SetDefaultMouseSpeed, 0
MouseMove 752, 477 ; "Origin" downarrow field Coordinates (Siebel)
MouseClick
Sleep 50
SendInput {Raw}Prescription - EPS
Sleep 50
SendInput {Enter}
MouseMove 751, 373 ; "Source" downarrow field Coordinates (Siebel)
MouseClick
Sleep 50
SendInput {Raw}Professional Contact
SendInput {Enter}
SetDefaultMouseSpeed, 2
}




siebelAddOosFlier()
{
MouseMove 591, 702 ; New product button coordinates
MouseClick
checkForProgressBar()
SendInput {Raw}380NLTC034
Sleep 50
SendInput {Tab}
Sleep 50
SendInput {Tab}
Sleep 500
SendInput {Raw}1
Sleep 50
}
















	; ====================== WAIT FUNCTIONS ========================


checkForProgressBar()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop, 100
{
	PixelSearch, OutputVarX, OutputVarY, 1299, 1027, 1383, 1035, 000080, 150, Fast ; Loop looks for progress bar to appear before moving forward

if (ErrorLevel != 0)
Sleep 50
else
break
}
Loop, 100
{
	PixelSearch, OutputVarX, OutputVarY, 1299, 1027, 1383, 1035, 000080, 150, Fast ; Loop looks for progress bar to disappear before moving forward

if (ErrorLevel = 0)
Sleep 50
else
Return
}
}




checkForWindowChangeGas()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
	PixelGetColor, color, 691, 526 ; "Team" field Coordinates (Siebel - gray color sample) - Loop looks for white color to appear at these coordinates before moving forward
While color = 0xEEEEEE
{
	PixelGetColor, color, 691, 526 
Sleep 10
}
}
















	; ====================== SAFETY FUNCTIONS ========================


checkIfScreenIsScrolledToTop()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop
{
PixelGetColor, color, 1908, 190
if (color = 0xFFFFFF)
{
SetDefaultMouseSpeed, 0
	MouseMove 1907, 178 ; Scrolls up to the top of the window
MouseClick
MouseClick
MouseClick
MouseClick
Sleep 1000
SetDefaultMouseSpeed, 2
return
}
else
{
break
}
}
}




checkIfSiebelOrderNoIsCopied()
{
if RegExMatch(Clipboard, "(\d{1}-\d{10})", OutputVar)
{
;
}
else
{
MsgBox, 4096, Error, "Script Failed - Please copy the Siebel Order number first. Press F12 to Exit the error message"
Reload
}
}




checkIfCurrentWindowIsCorrect()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop
{
global siebelCheckIfCurrentWindowIsCorrectAPLinesBoxX1 
global siebelCheckIfCurrentWindowIsCorrectAPLinesBoxY1 
global siebelCheckIfCurrentWindowIsCorrectAPLinesBoxX2 
global siebelCheckIfCurrentWindowIsCorrectAPLinesBoxY2 
global siebelCheckIfCurrentWindowIsCorrectContactsBoxX1 
global siebelCheckIfCurrentWindowIsCorrectContactsBoxY1 
global siebelCheckIfCurrentWindowIsCorrectContactsBoxX2 
global siebelCheckIfCurrentWindowIsCorrectContactsBoxY2 
ImageSearch OutputVarX, OutputVarY, %siebelCheckIfCurrentWindowIsCorrectAPLinesBoxX1%, %siebelCheckIfCurrentWindowIsCorrectAPLinesBoxY1%, %siebelCheckIfCurrentWindowIsCorrectAPLinesBoxX2%, %siebelCheckIfCurrentWindowIsCorrectAPLinesBoxY2%, *70 %A_ScriptDir%\Images\AP_Order_Lines.png
if (ErrorLevel = 0)
{
;
}
else if (ErrorLevel != 0)
{
BlockInput, MouseMoveOff
MsgBox, 4096, Error, You can only run this script from AP Order Lines Tab. Press F12 to Exit the error message.
Reload
}
ImageSearch OutputVarX, OutputVarY, %siebelCheckIfCurrentWindowIsCorrectContactsBoxX1%, %siebelCheckIfCurrentWindowIsCorrectContactsBoxY1%, %siebelCheckIfCurrentWindowIsCorrectContactsBoxX2%, %siebelCheckIfCurrentWindowIsCorrectContactsBoxY2%, *70 %A_ScriptDir%\Images\Contacts.png
if (ErrorLevel = 0)
{
Break
}
else if (ErrorLevel != 0)
{
BlockInput, MouseMoveOff
MsgBox, 4096, Error, You can only run this script from AP Order Lines Tab. Press F12 to Exit the error message.
Reload
}
}
}




checkIfCurrentWindowIsCorrectForStockAndCheck()
{
WinActivate, ahk_class Transparent Windows Client ; Opens up Siebel by its class
Loop
{
global siebelCheckIfCurrentWindowIsCorrectStockCheckBoxX1 
global siebelCheckIfCurrentWindowIsCorrectStockCheckBoxY1 
global siebelCheckIfCurrentWindowIsCorrectStockCheckBoxX2 
global siebelCheckIfCurrentWindowIsCorrectStockCheckBoxY2 
ImageSearch OutputVarX, OutputVarY, %siebelCheckIfCurrentWindowIsCorrectStockCheckBoxX1%, %siebelCheckIfCurrentWindowIsCorrectStockCheckBoxY1%, %siebelCheckIfCurrentWindowIsCorrectStockCheckBoxX2%, %siebelCheckIfCurrentWindowIsCorrectStockCheckBoxY2%, *20 %A_ScriptDir%\Images\Orders.png
if (ErrorLevel = 0)
{
Break
}
else (ErrorLevel != 0)
{
BlockInput, MouseMoveOff
MsgBox, 4096, Error, You can only run this script while Stock Checking the order. Press F12 to Exit the error message.
Reload
}
}
}




checkIfStickyNotesAreRunning()
{
Loop
{
if WinExist("ahk_class ApplicationFrameWindow")
Return
else
MsgBox, 4096, Error, Please open Sticky Notes app before running the script! Press F12 to exit
Reload
}
}




copyStickyNotes()
{
Loop, 5
{
if (Clipboard = "")
{
Sleep 100
WinActivate, ahk_class ApplicationFrameWindow ; Opens up Sticky Notes app by its class
Sleep 100
Send {Ctrl Down}
SendInput {Raw}c
Send {Ctrl Up}
ClipWait, 1
Sleep 50
}
else
{
Return
}
}
SetDefaultMouseSpeed, 0
MouseMove 370, 484
MouseClick
MsgBox, 4096, Error, Could not find the selected sticky notes! Press F12 to exit
Reload
}
















	; ====================== MISC FUNCTIONS ========================


clickOnPXAfterGAS()
{
SetDefaultMouseSpeed, 0
	MouseMove -910, 275 ; General PX coordinates (So that the script selects the PX at the end)
	MouseClick
SetDefaultMouseSpeed, 2
}




clearClipboard()
{
Clipboard :=
return
}




keyFix()
{
GetKeyState, state, Shift
if (state = "D")
{
SendInput {Shift down}
SendInput {Shift up}
}
else
{
;
}

GetKeyState, state, Ctrl
if (state = "D")
{
SendInput {Ctrl down}
SendInput {Ctrl up}
}
else
{
;
}

GetKeyState, state, Alt
if (state = "D")
{
SendInput {Alt down}
SendInput {Alt up}
}
else
{
;
}
}








; Scripts created by PLBASIO & PLTOTO