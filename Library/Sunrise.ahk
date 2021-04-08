; ***************************************************************************
; Sunrise library - electronic patient record (EPR) system
; 
; Notes:
; Sunrise works via a citirx session. Hence, only "pictures" are presented 
; on the screen and only graphic recognition, rather than direct control
; manipulation (for the most part) is possible to interact with this program.
; Hence the use of graphic finding functions.
; ***************************************************************************


; *********
; Variables
; *********

; Different executions are needed depending on if the user is on a Citrix machine or normal desktop
if (CitrixSession == True)
{
	SunriseEXEArray := ["""C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe"" -launch -reg ""Software\Microsoft\Windows\CurrentVersion\Uninstall\ght-436d6fbf@@GHT-CX-DDC.SCM02 - Allscripts"" -desktopShortcut"
						,"""C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe"" -launch -reg ""Software\Microsoft\Windows\CurrentVersion\Uninstall\ght-436d6fbf2@@GHT-CX-DDC.SCM02 - Allscripts"" -desktopShortcut"]    
}
else
{
	SunriseEXEArray := ["""C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe"" -launch -reg ""Software\Microsoft\Windows\CurrentVersion\Uninstall\ghtweb-c0cf2b92@@GHT-CX-DDC.SCM02 - Allscripts"" -desktopShortcut"]    
}	  

SunriseLoginTitle := "Sunrise EPR"
SunriseMainTitle := "Allscripts Gateway"
SunriseMainTitleMinimised := "My Applications"
SunriseFindPatientTitle := "Find Patient"
SunriseShowVisits := "Show Visits"
SunriseConnecting := "Connecting..."
SunriseHWND := ""
SunrisePID := ""
SunriseWarningMessage := "Warning Message"
SunriseErrorMessage := "Error Message"
SunriseClinicalManager := "Sunrise Clinical Manager"
SunriseXAMessage := "SunriseXA"


; To catch Sunrise timer dialogue box
pCallback := ""
SRTMessage := ""
SRCatchFunctionRunning := False




; *********
; Functions
; *********

SunriseSearch(MRN)
{
	global
	local system := "Sunrise"
	local StartSuccessful := False
	local LoginPass := False
	
	SetKeyDelay, 20, 20
	

	Loop, 5
	{
		try
		{
			if (GetCredentials(system, True) == False)
			{
				return False
			}
		
			SetTitleMatchMode, 2
			
			; The minimised window has a different title, hence the below if statement
			if WinExist(SunriseMainTitleMinimised)
			{
				WinMaximize, %SunriseMainTitleMinimised%
			}

			SetTitleMatchMode, 1

			
			if WinExist(SunriseLoginTitle) OR !WinExist(SunriseMainTitle)
			{
				if !WinExist(SunriseLoginTitle)
				{
					; Start Sunrise if not already running
					; As Sunrise is slow to start, loops help out here
					; Also Sunrise sometimes needs different arguements to start properly
					SubtitleMessage("Starting up Sunrise...")
					
					Loop, 5
					{
						Loop, % SunriseEXEArray.MaxIndex()
						{
							run, % SunriseEXEArray[A_index]

							Loop, 25
							{
								if winExist(SunriseConnecting)
								{
									StartSuccessful := True
									break
								}
								else if winExist(SunriseLoginTitle)
								{
									StartSuccessful := True
									break
								}
								
								sleep 200
							}
							
							if StartSuccessful
								break
						}
						
						if StartSuccessful
							break
					}
					
					if !StartSuccessful
					{
						SubtitleClose()
						MB("Sunrise did not seem to start. Please try again or manually start Sunrise from the desktop icon")
						return False
					}
				}
				
				; Input credentials
				WinWait, %SunriseLoginTitle%
				WinActivate, %SunriseLoginTitle%
				SubtitleClose()

				if !GraphicWait("SunriseUserName", 60000,,, "NoFail")
				{
					MB("Sunrise did not seem to start. Please try again or manually start Sunrise from the desktop icon")
					return False
				}

				GlobalInputLockSet(True)
				WinActivate, % SunriseLoginTitle
				GraphicClick("SunriseUserName", 100, 5,,, "SoftFail", True)
				ControlSend,, %username%, %SunriseLoginTitle%
				ControlSend,, {tab}, %SunriseLoginTitle%
				ControlSend,, %password%, %SunriseLoginTitle%
				ControlSend,, {enter}, %SunriseLoginTitle%
				GlobalInputLockSet(False)

				; Check if credentials worked
				Loop 50
				{
					if !winExist(SunriseLoginTitle)
					{
						LoginPass := True
						break
					}
					
					; Catch a warning for password update
					if winExist(SunriseWarningMessage)
					{
						if GraphicWait("SunriseCloseWarningMessage", 5000,, 20, "NoFail")
							GraphicClick("SunriseCloseWarningMessage", 5, 5,,, "SoftFail")
					}

					; !!! Need to confirm what this is for when I have time
					if WinExist(SunriseErrorMessage)
					{
						break
					}

					sleep 200
				}
				
				if !LoginPass
				{
					MB("Loggin error. Stopping automation")
					DeleteCredentials(system)
					return False
				}
			}
			else
			{
				; Close all extra Sunrise windows
				if WinExist(SunriseFindPatientTitle)
				{
					WinClose, %SunriseFindPatientTitle%
				}

				if WinExist(SunriseShowVisits)
				{
					WinClose, %SunriseShowVisits%
				}
			}
			
			Loop, 300 ; 300 * 200 = 1 minute
			{
				if WinExist(SunriseMainTitle)
				{
					; Click on "find patient" icon
					WinActivate, %SunriseMainTitle%
					break
				}
				else if winExist(SunriseClinicalManager)
				{
					; need to write this functionality - click on close/ok, but happens infrequently
					return False
				}
				
				sleep 200
			}
			
			; !!! Start timer catch function (need to finalise this work)
			;SunriseTimerCatch()
			
			
			; If not searching for a patient but only opening Sunrise, then complete run here
			if (MRN = "")
			{
				return True
			}
			
			GraphicWait("SunriseFindPatient", 30000,,, "SoftFail")
			sleep 200
			GraphicClick("SunriseFindPatient", 5, 5,,, "SoftFail")

			; Input MRN number and search
			WinWait, %SunriseFindPatientTitle%
			WinActivate, %SunriseFindPatientTitle%
			GraphicWait("SunriseID", 20000,,, "SoftFail")
			GlobalInputLockSet(True)
			GraphicClick("SunriseID", 100, 10,,, "SoftFail")
			Send, %MRN% ; controlSend did not work here
			GraphicClick("SunriseSearchButton", 0, 0,,, "SoftFail")
			GlobalInputLockSet(False)

			Loop 100
			{
				if WinExist(SunriseXAMessage)
				{
					MB("Patient with MRN" . MRN . " was not found on Sunrise")
					return False
				}

				if GraphicWait("SunriseTableName", 200,,, "NoFail")
				{
					break
				} 

				; No sleep needed as included in above function
			}
			
			GraphicWait("SunriseTableName", 20000,,, "SoftFail")
			GraphicClick("SunriseTableName", 5, 25,,, "SoftFail", True)
			WinWait, %SunriseShowVisits%
			WinActivate, %SunriseShowVisits%
			GraphicWait("SunriseTableAdmit", 20000,,, "SoftFail")
			GlobalInputLockSet(True)
			GraphicClick("SunriseTableAdmitUpArrow", 5, 5,,, "SoftFail", False)
			GraphicClick("SunriseTableAdmit", 5, 25,,, "SoftFail", True)
			GlobalInputLockSet(False)
			
			return True
		}
		catch, err
		{
			SubtitleClose()
			GlobalInputLockSet(False)
			
			; Close everything Sunrise related
			GroupAdd, SRGroup, % SunriseLoginTitle
			GroupAdd, SRGroup, % SunriseMainTitle
			GroupAdd, SRGroup, % SunriseMainTitleMinimised
			GroupAdd, SRGroup, % SunriseFindPatientTitle
			GroupAdd, SRGroup, % SunriseShowVisits
			GroupAdd, SRGroup, % SunriseClinicalManager
			GroupAdd, SRGroup, % SunriseWarningMessage
			GroupAdd, SRGroup, % SunriseErrorMessage
			GroupAdd, SRGroup, % SunriseXAMessage
			WinClose, ahk_group SRGroup
				
			sleep 1000
			
			if (err.message != "SoftFail" OR A_index >= 4)
			{
				errorhandler(err, "Sunrise", "Sunrise")
				return False
			}
			else
			{
				LogUpdate("Caught error with Sunrise (soft error) [error counter: " . A_index . "]")
			}
		}
	}
	
	return True
}



