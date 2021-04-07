; ***************************************************************
; Small program to copy changes in the Dev environment to the 
; Master directory for use by end-users
; ***************************************************************

#SingleInstance force ; Run only one instance and ignore update dialogue
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.


#include %A_ScriptDir%\Library\Private variables.ahk
#include %A_ScriptDir%\Library\Basic functions.ahk

MasterQSPath := Settings.MasterDirectory . "Quick Spiritum.ahk"
DevDirectory := Settings.MasterDirectory . "Dev\QuickSpiritum\"
DevQSPath := DevDirectory . "Quick Spiritum.ahk"
newLine := "`r`n"


; GUI to inform dev transfer happening
Gui, MBGUI:Font, s12
Gui, MBGUI:Color, %dialogueColour%
Gui, MBGUI:Add, Text, ym w300 H60 hwndMessagePtr, % "Transferring QS from Dev now..."
ControlGetPos, x, y, w, h, , ahk_id %MessagePtr%
h := h + 50
;Gui, MBGUI:Add, Button, x350 y%h% default,  &OK
Gui, MBGUI:Show,, %Title%
Gui, MBGUI:+AlwaysOnTop



; Copy across library updates
FileCopy, % DevDirectory . "Library\*.*", % Settings.MasterDirectory . "Library",  1
FileCopyDir, % DevDirectory . "Library", % Settings.MasterDirectory . "Library", 1 

; Copy across signature updates
FileCopy, % DevDirectory . "Signatures\*.*", % Settings.MasterDirectory . "Signatures",  1

; Copy across template updates
FileCopy, % DevDirectory . "Templates\*.*", % Settings.MasterDirectory . "Templates",  1

; Copy across graphics updates
FileCopy, % DevDirectory . "Graphics\*.*", % Settings.MasterDirectory . "Graphics",  1


FileRead, QSDevRead, % DevQSPath
lines := StrSplit(QSDevRead, newLine)
MasterVersionNumber := Trim(SubStr(lines[6], 25))
MasterVersionNumberInt := MasterVersionNumber + 1
lines[6] := "CurrentVersionNumber := " . MasterVersionNumberInt


; Dev file
fileDelete, % DevQSPath
fileObj := FileOpen(DevQSPath, "w")

Loop, % lines.MaxIndex()
{
	if (A_Index != lines.MaxIndex())
		fileObj.Write(lines[A_index] . newLine)
}

fileObj.Close()


; Master file
fileDelete, % MasterQSPath
fileObj := FileOpen(MasterQSPath, "w")

Loop, % lines.MaxIndex()
{
	if instr(lines[A_index], "Developing := True")
	{
		fileObj.Write("Developing := False" . newLine)
	}
	else if instr(lines[A_index], "emailTest := True")
	{
		fileObj.Write("emailTest := False" . newLine)
	}
	else
	{
		if (A_Index != lines.MaxIndex())
			fileObj.Write(lines[A_index] . newLine)
	}
}

fileObj.Close()

Gui, MBGUI:Destroy

Msgbox, % "Masterfile update from development folder complete - version: " . MasterVersionNumberInt

ExitApp

