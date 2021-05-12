; ******************************************************************************
; Quick Spiritum (QS) main code - copies selected hospital number 
; (MRN) or asks for this via GUI and then searches for patient
; in selected program or creates a request; If you see !!!, this is a note to me 
; that I need to fix / improve something at this point

CurrentVersionNumber := 101


#Warn useUnsetLocal			; Warn if local variable not initialised
#Warn useUnsetGlobal		; Warn if global variable not initialised
#SingleInstance force 		; Run only one instance and ignore update dialogue
#NoEnv  					; Recommended for performance and compatibility with future AutoHotkey releases.
#Persistent 				; keep running


; ********
; Settings
; ********

SendMode Input  			; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%	; Ensures a consistent starting directory.
SetKeyDelay, 10, 10
SetDefaultMouseSpeed, 0
UpdateCheck := True			; Will set to False in the future if running QS without a master file on a shared drive
Developing := True			; True turns off coping to desktop instead of running (Update Master File.ahk turns this off)
emailTest := True			; True sends emails / requests to developer only ('Update Master File.ahk' turns this off when copying from Dev to Master location)
pythonEXE := True			; Using EXE or .py version of python script
resetSettingsFile := False	; If new settings in settings file added then set to True to reset the end-users settings and add new ones.




; *********
; Variables
; *********

CurrentDirectory := A_ScriptDir . "\"

; The 'Private varialbles' file holds trust specific information 
; and keys. Make sure you initialise this to you local situation
#include %A_ScriptDir%\Library\Private variables.ahk

SettingsPath := CurrentDirectory . "settings.ini"
CurrentProgramPath := A_ScriptFullPath
CurrentProgramModfiedDate := ""
FileGetTime, CurrentProgramModfiedDate, % CurrentProgramPath
FormatTime, CurrentProgramModfiedDate, % CurrentProgramModfiedDate, dd/MM/yy
InstallINIPath := CurrentDirectory . "install.ini"
GraphicsPath := CurrentDirectory . "Graphics\"
TrayTipTitle := "Quick Spiritum"

if Developing
{
	Settings.RequestsFolder := settings.MasterDirectory . "Dev\Requests\"
	TrayTipTitle .= " DEV"
}

; Get or store settings in file "settings.ini"
if (!resetSettingsFile and FileExist(SettingsPath))
{
	Loop, Read, %SettingsPath%
	{
		LineSplit := StrSplit(A_LoopReadLine, A_Tab)
		Settings[LineSplit[1]] := LineSplit[2]
	}
}
else
{
	FileObj := FileOpen(SettingsPath, "w")
	for Key, label in settings
	{
		FileObj.WriteLine(Key . A_Tab . label . "`r")
	}
	FileObj.Close()
}


; Info messages
Menu, Tray, Icon, % GraphicsPath . "QS Logo.ico"
Menu, Tray, Tip, % "Quick Silver is running`nPress 'Shift+ESC' to close"




; Need info from settings file before setting these variables
pythonEXEPath := CurrentDirectory . "Library\GovNotifyMMF.exe"
DevPath := Settings.MasterDirectory . "Dev\QuickSpiritum\" . A_ScriptName
MasterPath := Settings.MasterDirectory . A_ScriptName
LogDirectory := Settings.MasterDirectory . "Dev\"
LogPath := LogDirectory . "QSLog.txt"

Desktop := A_Desktop . "\"
DesktopSubFolder := A_Desktop . "\QS" 	; need \ removed to work later
DesktopProgramPath := A_Desktop . "\QS\" . A_ScriptName
StartupShortcutName := "Quick Spiritum Shortcut.lnk"
DesktopShortcutName := "Quick Spiritum.lnk"

; For setttings GUI
MasterDirectoryTemp := ""
RequestsFolderTemp := ""
TemplatesFolderTemp := ""
LocListLocationzTemp := ""
IFZoomTemp := ""


; MRN = hospital number (7 numeric digits long at Gloucestershire Hospitals NHS trust)
; Lenghth can be changed in the Private Variables file
MRN := ""
MRNinput := ""
MRNGetPassed := False
InfoBoxTitle := "Quick Spiritum Info"
dialogueColour := "00FFFF"
SettingsTitle := "Settings"
CurrentlyRunning := False
MRNRequestTitle := "MRN Request"
Paused := False
Files := ""
S := 5 ; Spacing of GUIs



; clinicianDetails in private variables file
currentUser := getClinicianDetails(clinicianUsername, A_UserName, clinicianActualName)
consultantList := ""
PFT_consultantList := Username1 . "|"
doctorList := "" 

createClinicianLists()

createClinicianLists()
{
	global
	local clinicianType := ""
	local realname := ""
	local currentUserPostionInListConsultant := ""
	local PFT_currentUserPostionInListConsultant := ""
	local currentUserPostionInListDoctor := ""
	
	Loop, % clinicianDetails.MaxIndex()
	{
		clinicianType := clinicianDetails[A_index][clinicianPosition]
		realname := clinicianDetails[A_index][clinicianActualName]
		
		if (clinicianType = "consultant")
		{
			consultantList .= realname . "|"
			PFT_consultantList .= realname . "|"
			doctorList .= realname . "|"
		}
		else
		{
			doctorList .= realname . "|"
		}
	}


	; 0  if not found
	currentUserPostionInListConsultant := InStr(consultantList, currentUser)
	PFT_currentUserPostionInListConsultant := InStr(PFT_consultantList, currentUser)
	currentUserPostionInListDoctor := InStr(doctorList, currentUser)

		
	; Set consultant's list
	if (CurrentUserPostionInListConsultant != 0)
		consultantList := SubStr(consultantList, 1, CurrentUserPostionInListConsultant + StrLen(CurrentUser)) . "|" . SubStr(consultantList, CurrentUserPostionInListConsultant + StrLen(CurrentUser) + 1)
	
	; Set consultants and SAS list for PFTs
	if (PFT_CurrentUserPostionInListConsultant != 0)
		PFT_consultantList := SubStr(PFT_consultantList, 1, PFT_CurrentUserPostionInListConsultant + StrLen(CurrentUser)) . "|" . SubStr(PFT_consultantList, PFT_CurrentUserPostionInListConsultant + StrLen(CurrentUser) + 1)
			
	; Set doctor (and PA) list
	if (CurrentUserPostionInListDoctor != 0)
		doctorList := SubStr(doctorList, 1, CurrentUserPostionInListDoctor + StrLen(CurrentUser)) . "|" . SubStr(doctorList, CurrentUserPostionInListDoctor + StrLen(CurrentUser) + 1)
	
	return
}




getClinicianDetails(searchBy, searchValue, returnValue)
{
	global
	

	Loop, % clinicianDetails.MaxIndex()
	{
		if (clinicianDetails[A_index][searchBy] == searchValue)
		{
			return clinicianDetails[A_index][returnValue]
		}
	}
	
	return False
}




; Check if running on a Citrix machine
if (InStr(GetComputerNameEx(), "-CX-") > 0)
	CitrixSession := True
else
	CitrixSession := False




; *********
; Libraries
; *********

#include %A_ScriptDir%\Library\Basic functions.ahk
#include %A_ScriptDir%\Library\Main.ahk
#include %A_ScriptDir%\Library\PFT.ahk
#include %A_ScriptDir%\Library\Requests.ahk
#include %A_ScriptDir%\Library\TrayIcon.ahk
#include %A_ScriptDir%\Library\InfoFlex.ahk
#include %A_ScriptDir%\Library\Internet Explorer.ahk
#include %A_ScriptDir%\Library\Trakcare.ahk
#include %A_ScriptDir%\Library\Noxturnal.ahk
#include %A_ScriptDir%\Library\PACS.ahk
#include %A_ScriptDir%\Library\DigiDictate.ahk
#include %A_ScriptDir%\Library\ICE.ahk
#include %A_ScriptDir%\Library\Sunrise.ahk
#include %A_ScriptDir%\Library\Spiritum.ahk
#include %A_ScriptDir%\Library\Email.ahk
#include %A_ScriptDir%\Library\VBA_AHK_IPC.ahk
#include %A_ScriptDir%\Library\QIP.ahk
#include %A_ScriptDir%\Library\Sleep service PPV.ahk
#include %A_ScriptDir%\Library\PIV.ahk
#include %A_ScriptDir%\Library\Vis2.ahk
#include %A_ScriptDir%\Library\Gdip_All.ahk


; Read through files in the startup folder
/*
Loop %A_startup%\*.*
	Files1 := Files1 . "`n" . A_LoopFileName

MsgBox, % Files1
*/

;FileDelete, % A_startup . "\" . StartupShortcutName ; Used to delete the startup process (if needed)


; Start up functions including copying master if new version 
; present, copy master to desktop if master file used and also 
; places a short cut in the OS Startup folder
;FileDelete, % A_Desktop . "\" . A_ScriptName ; !!! will need to delete eventually

if !Developing
{
	if fileExist(InstallINIPath) ; Designed so that an installation version of AHK can be run outside of the master directory
	{
		if (A_Args[1] = "update")
			TrayTip, %TrayTipTitle%, % "Updating. Please wait for 'running' message before trying to use QS again..."
		else
			TrayTip, %TrayTipTitle%, % "Copying master Quick Spiritum program to Desktop"
		
		FileCreateDir, %DesktopSubFolder%
		FileSetAttrib, +H, %DesktopSubFolder%
		FileCopy, %CurrentProgramPath%, %DesktopProgramPath%, 1     ; 1 = overwrite
		
		FileCreateDir, %DesktopSubFolder%\Graphics
		FileCopy, %GraphicsPath%*.*, %DesktopSubFolder%\Graphics,  1
		
		FileCreateDir, %DesktopSubFolder%\Templates
		FileCopy, % Settings.MasterDirectory . "Templates\*.*", %DesktopSubFolder%\Templates,  1
		
		FileCreateDir, %DesktopSubFolder%\Signatures
		FileCopy, % Settings.MasterDirectory . "Signatures\*.*", %DesktopSubFolder%\Signatures,  1
				
		FileCreateDir, %DesktopSubFolder%\Library
		FileCopy, % Settings.MasterDirectory . "Library\*.*", %DesktopSubFolder%\Library,  1
		FileCopyDir, % Settings.MasterDirectory . "Library", %DesktopSubFolder%\Library, 1 	
		
		FileCreateShortcut, %DesktopProgramPath%, %Desktop%%DesktopShortcutName%, %DesktopSubFolder%,, Quick Spiritum Shortcut, %DesktopSubFolder%\Graphics\QS logo.ico
		TrayTip, %TrayTipTitle%, % "Transfer complete"
		ExitApp
	}
	
	checkForUpdate()

	; To check if QS startup shortcut in Windows Startup folder, 
	; if then not create this.
	Shortcutfound := False

	Loop %A_startup%\*.*
	{
		Files := Files . "`n" . A_LoopFileName

		if (A_LoopFileName == StartupShortcutName)
		{
			Shortcutfound := True
		}
	}

	if (Shortcutfound == False AND CurrentProgramPath != DevPath)
	{
		FileCreateShortcut, %A_ScriptFullPath%, %A_startup%\%StartupShortcutName%, %A_scriptDir%,, Quick Spiritum Shortcut
	}
}


;msgbox, FileCreateShortcut, %A_ScriptFullPath%, %A_startup%\%StartupShortcutName%, %A_scriptDir%,, Quick Spiritum Shortcut
;FileCreateShortcut, %A_ScriptFullPath%, %A_Desktop%\%StartupShortcutName%, %A_scriptDir%,, Quick Spiritum Shortcut
/*
Loop %A_startup%\*.*
{
	Files2 := Files2 . "`n" . A_LoopFileName
}

MsgBox, % Files2
;FileDelete, %A_startup%\%ShortcutName%
*/




; Turn off Imprivata (automatic credentials entering software). 
; QS takes over this functionality.
timedFunction()
setTimer, timedFunction, 300000 ; 300000 = run every 5 minutes



; Check for update
setTimer, checkForUpDate, 3600000 ;86400000 ; 86400000 millisec = 1 day, 3600000 = 1 hour

checkForUpdate()
{
	global
	local lineM := ""
	local MasterVersionNumber := ""

	; to check for master AHK update
	if (UpdateCheck == True AND CurrentProgramPath != MasterPath)
	{
		if FileExist(MasterPath)
		{
			Loop, read, % MasterPath
			{
				if (inStr(A_LoopReadLine, "CurrentVersionNumber") == 1)
				{
					MasterVersionNumber := Trim(SubStr(A_LoopReadLine, 25))

					if (MasterVersionNumber != CurrentVersionNumber)
					{					
						RunWait, % MasterPath . " update"
						Reload
					}
					
					break
				}
			}
			/*
			FileReadLine, lineM, %MasterPath%, 6
			MasterVersionNumber := Trim(SubStr(lineM, 25))

			if (MasterVersionNumber != CurrentVersionNumber)
			{					
				RunWait, % MasterPath . " update"
				Reload
			}
			*/
		}
	}
	
	return True
}




TrayTip, %TrayTipTitle%, Running

; Start up the VBA to QS IPC function
V2AMessages := new MemoryMappedFile_IPC()

; Array (FIFO) with all of the hotkeys available
; Spaces removed later
HotKeys := FiFoArray("F1" , "Info Box" 
	   ,"F2" , "Sunrise"
	   ,"F3" , "Trakcare" 
	   ,"F4" , "InfoFlex"
	   ,"F5" , "PACS"
	   ,"F6" , "PAS"
	   ,"F7" , "DigiDictate"
	   ,"F8" , "Lung Function Tests"
	   ,"F9" , "Spiritum"
	   ,"F10" , "Noxturnal"
	   ;,"F12" , "NHS Mail"			; recently NHS mail was changed to Office 365. I have not bothered trying to automate this update yet. Anyone interested?
	   ,"^F2" , "Sunrise Open Only"
	   ,"^F3" , "Trakcare Open Only"
	   ,"^F4" , "InfoFlex Open Only"
	   ,"^F5" , "PACS Open Only"
	   ,"+F2" , "ICE"
	   ,"+F3" , "PET-CT Request"
	   ,"+F4" , "Lung Function Request"
	   ,"+F5" , "Bronchoscopy Request"
	   ,"+F6" , "Healthy Lifestyles Gloucestershire Referral"
	   ,"+F7" , "Sleepstation Referral"
	   ,"+F8", "Send Patient Information Videos"
	   ,"+F11", "Sleep Questionnaire"
	   ,"+F12", "Quality Improvement Program")

; Set the hotkeys
For index, Key in HotKeys[]
{
	Label := StrReplace(HotKeys[Key], " ", "")
	Hotkey, %Key%, %Label%, On
}




; ************
; Dev material
; ************

if False ;Developing
{
	Gui +LastFound
	hWnd := WinExist()
	SRTMessage := ""

	DllCall("RegisterShellHookWindow", UInt, hWnd)
	MsgNum := DllCall("RegisterWindowMessage", Str, "SHELLHOOK")
	OnMessage(MsgNum, "ShellMessage")
}

if False
{
	; Start up OCR class
	OCRcontrol := new Vis2()
}


ShellMessage(wParam,lParam) 
{
  global
  local TimerAppeared := False
  
  
  If (wParam = 1)  ;  HSHELL_WINDOWCREATED := 1
  {
    NewID := lParam

	WinGetTitle, WinTitle, ahk_id %NewID%
	OnMessage(MsgNum, "ShellMessage", 0)

	if inStr(winTitle, "Allscripts Gateway")
	{
		msgbox, % "Timer script to run now for'n" . "Title: " . winTitle
/*
		Loop 10
		{
			if GraphicWait("SunriseAboutToLogOff", 200,,, False)
			{
				if GraphicClick("SunriseAboutToLogOff", 132, 47,,, False)
				{
					TimerAppeared := True
					SRTMessage := SRTMessage . "Caught timer dialogue at: " . Today . "`n"
					Msgbox,, Sunrise timer closer, % SRTMessage, 3
					Menu, Tray, Tip, % SRTMessage
				}
			}
			else if (TimerAppeared == True)
			{
				break
			}
			
			sleep 200
		}
*/
	}
	else
		msgbox, % "Title: " winTitle
		
	OnMessage(MsgNum, "ShellMessage")
  }
  
  return True
}











; *******************
; End of Dev material
; *******************

return ; Place AHK into permanent state




; ******************
; Hot key functions
; ******************

InfoBox:
InfoBox()
return




InfoFlexOpenOnly:
if runningStatus()
	return
MRN := ""
InfoFlexSearch(MRN)
runningStatus("done")
LogUpdate("InfoFlex (open only)")
return




InfoFlex:
if runningStatus()
	return
if !GrabMRN("InfoFlex")
	return
InfoFlexSearch(MRN)
runningStatus("done")
LogUpdate("Infoflex")
return



TrakcareOpenOnly:
if runningStatus()
	return
MRN := ""
TrakcareSearch(MRN)
runningStatus("done")
LogUpdate("Trakcare (open only)")
return




Trakcare:
if runningStatus()
	return
if !GrabMRN("Trakcare")
	return
TrakcareSearch(MRN)
runningStatus("done")
LogUpdate("Trakcare")
return




PACSOpenOnly:
if runningStatus()
	return
MRN := ""
PACSSearch(MRN)
runningStatus("done")
LogUpdate("PACS (open only)")
return




PACS:
if runningStatus()
	return
if !GrabMRN("PACS")
	return
PACSSearch(MRN)
runningStatus("done")
LogUpdate("PACS")
return




PAS:
if runningStatus()
	return
if !GrabMRN("PAS (via Trakcare)")
	return
if TrakcareSearch(MRN)
	PASSubSearch(MRN)
runningStatus("done")
LogUpdate("PAS")
return




ICE:
if runningStatus()
	return
if !GrabMRN("ICE")
	return
ICESearch(MRN)
runningStatus("done")
LogUpdate("ICE")
return




DigiDictate:
if runningStatus()
	return
if !GrabMRN("DigiDictate")
	return
DigidictateSearch(MRN)
runningStatus("done")
LogUpdate("Digidictate")
return




Noxturnal:
if runningStatus()
	return
if !GrabMRN("Noxturnal")
	return
NoxturnalSearch(MRN)
runningStatus("done")
LogUpdate("Noxturnal")
return




LungFunctionTests:
if runningStatus()
	return
if !GrabMRN("Lung function results")
	return
PFTSearch(MRN)
runningStatus("done")
LogUpdate("Lung function requests")
return




SunriseOpenOnly:
if runningStatus()
	return
MRN := ""
SunriseSearch(MRN)
runningStatus("done")
LogUpdate("Sunrise (open only)")
return




Sunrise: 
if runningStatus()
	return
if !GrabMRN("Sunrise")
	return
SunriseSearch(MRN)
runningStatus("done")
LogUpdate("Sunrise")
return




Spiritum:
if runningStatus()
	return
SpiritumStart()
runningStatus("done")
LogUpdate("Spiritum")
return




NHSMail:
if runningStatus()
	return
NHSMailOpen()
runningStatus("done")
LogUpdate("NHS mail")
return




PET-CTrequest:
RRCreateAndSend("PET_CT")
return




LungFunctionRequest:
RRCreateAndSend("PFT")
return




BronchoscopyRequest:
RRCreateAndSend("bronchoscopy")
return




HealthyLifestylesGloucestershireReferral:
RRCreateAndSend("HLSG")
return




SleepStationReferral:
RRCreateAndSend("Sleepstation")
return




SleepQuestionnaire:
sleepQuestionnaire(MRN1)
return




SendPatientInformationVideos:
if !registeredUsername()
	return
if runningStatus()
	return
if !GrabMRN("patient information videos")
	return

;MRN := MRN1
		
CloseProgressBar()
PBMainString := "Patient information video:`n"
CreateProgressBar()

if GetPatientDetailsTrakcare(MRN)
{
	UpdateProgressBar(100,, PBMainString . "done")
	CloseProgressBar()
	GOV_UK_Notify()
}
CloseProgressBar()
runningStatus("done")
return




QualityImprovementProgram:
QIP_main()
return




ESC::
SetTimer, timedFunction, Delete
;Acc_UnhookWinEvent(pCallback) ; !!!
SRCatchFunctionRunning := False

if !Paused
{
	Paused := True

	For index, Key in HotKeys[]
	{
		Label := StrReplace(HotKeys[Key], " ", "")
		Hotkey, %Key%, %Label%, Off
	}

	TrayTip, %TrayTipTitle%, Paused
}
else
{
	TrayTip, %TrayTipTitle%, restarting...
	;Acc_UnhookWinEvent(pCallback) ; !!!
	;SRCatchFunctionRunning := False ; !!!
	Sleep 1000
	Reload
	return
}
return




; Pressing shift and escape closes down QS
+ESC::
;Acc_UnhookWinEvent(pCallback) ; !!!
TrayTip, %TrayTipTitle%, Closing down...
Sleep 1000
ExitApp
return




; ************************
; Functions specific to QS
; ************************

InfoBox()
{
	global
	local Column1 := "Open and search:"
	local Column2 := "Open without search (Ctrl + ...):`n"
	local Column3 := "Requests (Shift + ...):`n"
	local NumberOfHotKeys := 2
	local OKYValue := 0


	if InStr(CloseGUIs(), "InfoBox")
		return True

	For index, Key in HotKeys[]
	{
		Label := HotKeys[Key]
		
		if InStr(Key, "^") = 1
		{
			Key := StrReplace(Key, "^", "")
			Label := StrReplace(Label, " Open Only", "")
			Column2 := Column2 . "`n" . Key . ": " . Label
		}
		else if InStr(Key, "+") = 1
		{
			Key := StrReplace(Key, "+", "")
			Column3 := Column3 . "`n" . Key . ": " . Label
		}
		else
		{
			Column1 := Column1 . "`n" . Key . ": " . Label
			NumberOfHotKeys := NumberOfHotKeys + 1	
		}
	}



	OKYValue := NumberOfHotKeys * 23

	Gui, InfoBoxGUI:Color, %dialogueColour%, %dialogueColour%
	Gui, InfoBoxGUI:Font, s12
	Gui, InfoBoxGUI:Add, Text, x50 y10, %Column1%
	Gui, InfoBoxGUI:Add, Text, x240 y10, %Column2%
	Gui, InfoBoxGUI:Add, Text, x490 y10, %Column3%
	Gui, InfoBoxGUI:Add, Text, x200 y%OKYValue%, % "ESC: pause/restart,     Shift + ESC: end program."
	versionYvalue := OKYValue + 20
	Gui, InfoBoxGUI:Add, Text, x200 y%versionYvalue%, % "Date last updated: " . CurrentProgramModfiedDate . ", Version: " . CurrentVersionNumber
	
	if Developing
	{
		versionYvalue += 20
		OKYValue += 20
		Gui, InfoBoxGUI:Add, Text, x200 y%versionYvalue%, % "DEVELOPMENT VERSION!"
	}
	
	Gui, InfoBoxGUI:Add, Button, x580 y%OKYValue% default, &Close
	Gui, InfoBoxGUI:Add, Button, x650 y%OKYValue% default, &Settings
	Gui, InfoBoxGUI:+AlwaysOnTop -MinimizeBox
	Gui, InfoBoxGUI:Show,, %InfoBoxTitle%
	return True
}




CloseGUIs()
{
	global
	local returnString := ""


	if WinExist(InfoBoxTitle)
	{
		Gui, InfoBoxGUI:Destroy
		returnString := returnString . "-InfoBox"
	}
	
	if WinExist(SettingsTitle)
	{
		Gui, SettingsGUI:Destroy
		returnString := returnString . "-Settings"
	}
	
	if WinExist(MRNRequestTitle)
	{
		Gui, MRNGUI:Destroy
		returnString := returnString . "-MRNRequest"
	}

	return returnString
}




InfoBoxGUIButtonClose:
InfoBoxGUIguiClose:
Gui, InfoBoxGUI:Destroy
return




InfoBoxGUIButtonSettings:
Gui, InfoBoxGUI:Destroy
SettingsGUI()
return




CloseInfoBox()
{
	CloseGUIs()
	return
}




runningStatus(status := "")
{
	global
	
	; Frees up other Hotkeys to run
	if (status = "done")
	{
		CurrentlyRunning := False
		;turnOffImprivata()
		return False
	}
	
	if (CurrentlyRunning == True)
	{
		return True
	}
	else
	{
		CurrentlyRunning := True
		CloseInfoBox()
		return False
	}
}




GrabMRN(system)
{
	global
	
	MRNInput := ""
	MRN := ""

	clipboard := ""
	sleep 200
	send, ^c
	sleep 200

	MRNinput := StrReplace(clipboard, " ", "")
	
	if !MRNCheck()
	{
		/*
		if False ; Holding off using OCR for now. Around 80-90% accuracy
		{
			MRNinput := OCRcontrol.OCR()
			MRNinput := StrReplace(MRNinput, "O", "0")
			MRNinput := StrReplace(MRNinput, "Q", "0")
		}
		*/
		AskForMRN(system)
	}
	else
	{
		MRNGetPassed := True
	}
	
	if (MRNGetPassed == False)
		CurrentlyRunning := False
		
	return MRNGetPassed
}




MRNCheck()
{
	global
	
	if MRNInput is digit
	{
		if (strlen(MRNInput) = MRNLength)
		{
			MRN := MRNinput
			return True
		}
	}

	; Got to this point if the MRN entered is either not 7 (at our trust) characters long, or contains non-digits!
	MRN := ""
	MRNInput := ""
	return False
}




AskForMRN(system)
{
	global

	Gui, MRNGUI:Font, s12
	Gui, MRNGUI:Color, %dialogueColour%
	Gui, MRNGUI:Add, Text, x10 y10, Please provide an MRN for '%system%':
	Gui, MRNGUI:Add, Text, x10 y40, MRN:
	Gui, MRNGUI:Add, Edit, x60 y40 vMRNInput, % MRNinput
	;Gui, MRNGUI:Add, Text, x10 y70, % "(If OCR used, please check correct MRN found)"
	Gui, MRNGUI:Add, Button, x200 y100 default,  &OK
	Gui, MRNGUI:Add, Button, x240 y100,  &Cancel
	Gui, MRNGUI:Show,, % MRNRequestTitle
	Gui, MRNGUI:+AlwaysOnTop
	ControlSend, Edit1, {Right}, % MRNRequestTitle ; unhighlight MRNInput field
	WinWaitClose, %MRNRequestTitle%
	return True
}




MRNGUIButtonOK:

Gui, Submit  ; Save the input from the user to each control's associated variable.
Gui, Destroy ; Need destroy to be able to use MRNInput more than once (strange)!

if !MRNCheck()
{
	MB("Invalid MRN entered!")
	MRNGetPassed := False
}
else
{
	MRNGetPassed := True
}

return




MRNGUIButtonCancel:
MRNGUIGuiClose:

Gui, Submit  ; Save the input from the user to each control's associated variable.
Gui, Destroy ; Need destroy to be able to use MRNInput more than once (strange)!
MRNGetPassed := False
return




SettingsGUI()
{
	global
	local Height := 70
	
	
	Gui, settingsGUI:Font, s12
	Gui, settingsGUI:Color, %dialogueColour%
	Gui, settingsGUI:Add, Text, x10 y10, Master directory:
	Gui, settingsGUI:Add, Edit, x200 y10 W400 H%Height% vMasterDirectoryTemp, % settings.MasterDirectory
	
	Gui, settingsGUI:Add, Text, x10 y90, Requests folder:
	Gui, settingsGUI:Add, Edit, x200 y90 W400 H%Height% vRequestsFolderTemp, % settings.RequestsFolder
	
	Gui, settingsGUI:Add, Text, x10 y170, Templates folder:
	Gui, settingsGUI:Add, Edit, x200 y170 W400 H%Height% vTemplatesFolderTemp, % settings.TemplatesFolder
	
	Gui, settingsGUI:Add, Text, x10 y250, Trakcare Location choice:
	Gui, settingsGUI:Add, Edit, x200 y250 W400 H%Height% vLocListLocationzTemp, % settings.LocListLocationz
	
	if settings["IFZoom"]
		checked := "Checked"
	else
		checked := ""
	
	Gui, settingsGUI:Add, CheckBox, x5 y330 vIFZoomTemp +Right %checked%, % "InfoFlex zoom into letter:   " 
	
	Gui, settingsGUI:Add, Button, x480 y380 default,  &OK
	Gui, settingsGUI:Add, Button, x535 y380, &Cancel
	
	Gui, settingsGUI:Show,, SettingsTitle
	Gui, settingsGUI:+AlwaysOnTop
	WinWaitClose, SettingsTitle
	return True
}




settingsGUIButtonOK:

errorString := ""
pass := False
Gui, settingsGUI:Submit, Nohide

if !FileExist(MasterDirectoryTemp)
	errorString := errorString . "- The path for master directory does not exist`n"
	
if !FileExist(RequestsFolderTemp)
	errorString := errorString . "- The path for request directory does not exist`n"
	
if !FileExist(TemplatesFolderTemp)
	errorString := errorString . "- The path for templates directory does not exist`n"

if LocListLocationzTemp is number
	if (LocListLocationzTemp > 0 AND LocListLocationzTemp < 15)
		pass := True
		
if (pass == False)
	errorString := errorString . "- The Trakcare location choice needs to be between 1 and 16 (exclusive)`n"

pass := False

if (errorString != "")
{
	MB("Please correct the below errors and submit again:`n" . errorString)
	return
}

Gui, settingsGUI:Destroy

settings.MasterDirectory := MasterDirectoryTemp
settings.RequestsFolder := RequestsFolderTemp
settings.TemplatesFolder := TemplatesFolderTemp
settings.LocListLocationz := LocListLocationzTemp 
settings.IFZoom := IFZoomTemp

FileObj := FileOpen(SettingsPath, "w")

for Key, label in settings
{
	FileObj.WriteLine(Key . A_Tab . label . "`r")
}

FileObj.Close()
return




settingsGUIClose:
settingsGUIGuiClose:
settingsGUIButtonCancel:
Gui, settingsGUI:Destroy
return




LogUpdate(message)
{
	global
	local Today := ""
	FormatTime, Today,, dd/MM/yy - HH:mm
	
	
	if !fileExist(LogDirectory)
		return False
	
	if !fileExist(LogPath)
	{
		FileAppend, Quick Spiritum Log`n, % LogPath
		sleep 200
	}
	
	FileAppend, % Today . ": " . getClinicianDetails(clinicianUsername, A_UserName, clinicianActualName) . " (" . A_UserName . ") - " . message . "`n", % LogPath
	return True
}




registeredUsername()
{
	global 
	

	if (getClinicianDetails(clinicianUsername, A_UserName, clinicianActualName) == False)
	{
		msgBox, % "You are not registered to use the request/referrals functionality, please contact the superuser to rectify."
		return False
	}	
	return True
}

