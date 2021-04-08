; ************
; PACS library
; ************

; *********
; Variables
; *********

PACSEXE := """C:\Program Files (x86)\Philips\IntelliSpace PACS Enterprise\4.4\IntelliSpacePACSEnterprise.exe"""
PACSTitle := "Philips IntelliSpace PACS Enterprise"
PACSMRNBox := "Edit4"



; *********
; Functions
; *********

PACSSearch(MRN)
{
	Global
	system:= "PACS"
	LoginPass := False
	
	SetKeyDelay, 5, 5
	
	
	try	
	{
		if (!WinExist(PACSTitle) OR !ControlPresent("Button2", PACSTitle) == True)
		{
			if (GetCredentials(system) == False)
			{
				return False
			}

			if !WinExist(PACSTitle)
			{
				try
				{
					SubtitleMessage("Starting up PACS...")
					run, %PACSEXE%
				}
				catch
				{
					SubtitleClose()
					MB("Issue with starting PACS. Stopping automation",, "DevInform")
					return False
				}
				WinWait, %PACSTitle%
			}

			SubtitleClose()
			if !GraphicWait("PACSTriangle", 60000,,, False)
			{
				MB("It appears PACS did not start up correctly! Please try again.")
				return False
			}
			GlobalInputLockSet(True)
			ControlSetText, Edit1, %username%, %PACSTitle%
			ControlSetText, Edit2, %password%, %PACSTitle%
			PostClick(5, 5, "Button1", PACSTitle)
			sleep 1000 ; need otherwise previous login error might be read off screen
			GlobalInputLockSet(False)
			
			Loop 100
			{
				if ControlPresent("&OK", PACSTitle)
				{
					sleep 200
					DeleteCredentials(system)
					PostClick(5, 5, "&OK", PACSTitle)
					SubtitleClose()
					MB("Password update noted. Stopping PACS automation. Please update PACS password and try QS automation again")
					return False
				}
				
				if !GraphicWait("PACSTriangle", 200,,, False)
				{
					LoginPass := True
					break
				}
				
				if GraphicWait("PACSLoginFail", 200,,, False)
				{
					LoginPass := False
					break
				}

				; No sleep needed as incoorporated in GraphicWait
			}

			if (LoginPass == False)
			{
				SubtitleClose()
				MB("Credentials error. Stopping automation. Please try again")
				DeleteCredentials(system)
				return False
			}
		}
		else
		{
			PostClick(5, 5, "Patient Lookup", PACSTitle)
			PostClick(5, 5, "clear all", PACSTitle)
		}

		WinActivate, % PACSTitle
		
		if (MRN = "")
		{
			return True
		}
		
		if !GraphicWait("PACSMRN", 1000,,, False)
		{
			Loop 10
			{
				if !GraphicWait("PACSMRN", 500,,, False)
				{
					WinRestore, % PACSTitle
					sleep 500
					WinMaximize, % PACSTitle
					sleep 500
					break
				}
			}
		}
		
		GlobalInputLockSet(True)
		WinActivate, %PACSTitle%
		GraphicClick("PACSMRN", 5, 25)
		sleep 200
		ControlSetText, %PACSMRNBox%, %MRN%, %PACSTitle%
		PostClick(5, 5, "search", PACSTitle)
		GlobalInputLockSet(False)

		GraphicWait("PACSPlus", 2000,,, False)
		GraphicClick("PACSPlus", 2, 2,,, False)
	}
	catch, err
	{
		SubtitleClose()
		errorhandler(err, "PACS", "PACS")
		return False
	}
	
	return True
}



