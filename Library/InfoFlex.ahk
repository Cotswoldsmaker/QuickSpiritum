; ****************
; InfoFlex library
; ****************

; ***********
; Variables
; ***********

InfoFlexEXE := "C:\Program Files (x86)\CIMS\InfoFlex5\Progs\InfoFlex5.exe"
InfoFlexLogonMessage := "InfoFlex Logon Message"
InfoFlexLoginTitle := "InfoFlex Login"
InfoFlexTitle := "InfoFlex v5 Data Entry"
InfoFlexSelectModule := "Select Module"
InfoFlexSubjectSearch := "Infoflex Subject Search"
InfoFlexDataEntry := "InfoFlex Data Entry"
InfoFlexUnlockTitle := "Unlock InfoFlex v5"




; *********
; Functions
; *********

SelectDataView(View)
{
	global
	local text := ""
	local ScrollTotal := 10
	
	SetTitleMatchMode, 1
	
	
	; !!! Perhaps allowing a no fail here is lazy. Could perhaps do with improving
	if !GraphicWait("InfoFlexDataView", 10000,, 30, "NoFail") ; !!! Need higher variance perhaps, has failed here
		return False
		
	ControlGetText, text, Edit1, % InfoFlexTitle

	if (text = View)
	{
		return True
	}
	else
	{
		Lockset(True)

		ControlGetPos, xctrl, yctrl,,, Edit1, %InfoFlexTitle%
		MouseMove, xctrl + 5, yctrl + 5
		sleep 200
		ControlClick, Edit1, % InfoFlexTitle
		
		; Rather complicated way of handling dialogue boxes that have diff named buttons
		ControlClick, Edit1, %InfoFlexTitle%, ,WheelUp
		sleep 200
		
		if WinExist(InfoFlexDataEntry)
		{
			try
				PostClick(5, 5, "Button1", InfoFlexDataEntry)
				
			try ; needed for situation when no subfolders in 'SAP - Sleep apnoea' folder
				PostClick(5, 5, "SSCommandWndClass3", InfoFlexDataEntry)
		}
		
		sleep 200
		
		Loop %ScrollTotal%
		{
			ControlClick, Edit1, % InfoFlexTitle,,WheelUp
		}

		Loop %ScrollTotal%
		{
			  ControlGetText, text, Edit1, % InfoFlexTitle
			  
			  if (text = View)
			  {      
					Lockset(False)
					return True
			  }

			  ControlClick, Edit1, % InfoFlexTitle,, WheelDown
		}
	}

	Lockset(False)
	return False
}




; Note: InfoFlex needs to move the mouse over items (except buttons) to interact consistently with them
InfoFlexSearch(MRN)
{
	global
	local system := "InfoFlex"
	local PatientFound := False
	FormatTime, Today, ,dd/MM/yyyy
	
	SetTitleMatchMode, 1
	SetKeyDelay, 10, 10


	try
	{
		; Handle if InfoFlex is at Select Module dialogue
		if WinExist(InfoFlexSelectModule)
		{
			keyPress("SSCommandWndClass1", "enter", InfoFlexSelectModule)
			WinWait, % InfoFlexTitle
		}
		
		; Start InfoFlex if not already running or start to input credentials on Login screen
		if !WinExist(InfoFlexTitle) or ControlPresent("ThunderRT6TextBox2", InfoFlexLoginTitle) == True
		{	
			if !GetCredentials(system)
			{
				return False
			}
		
			if !WinExist(InfoFlexTitle) AND !WinExist(InfoFlexLoginTitle)
			{
				Try
				{
					SubtitleMessage("Starting up InfoFlex...")
					Run, % InfoFlexEXE
				}
				Catch
				{
					SubtitleClose()
					MB("InfoFlex did not start correctly. Stopping automation.",, "DevInform")
					return False
				}

				WinWait, % InfoFlexLoginTitle
				SubtitleClose()
			}

			WinActivate, %InfoFlexLoginTitle%
			GlobalInputLockSet(True)
			ControlSetText, ThunderRT6TextBox2, % username, % InfoFlexLoginTitle
			ControlSetText, ThunderRT6TextBox1, % password, % InfoFlexLoginTitle
			Send, {enter}
			GlobalInputLockSet(False)

			; Handle incorrect username or password
			Loop 50
			{
				if WinExist(InfoFlexLogonMessage)
				{
					try
					{
						ControlClick, OK, % InfoFlexLogonMessage
						MB("Credentials error. Stopping automation. Please try again")
						DeleteCredentials(system)
						return False
					}
					catch
					{
						ControlClick, &No, % InfoFlexLogonMessage
					}
				}
			
				if WinExist(InfoFlexSelectModule)
				{
					break
				}

				sleep 200
			}

			; Select data entry
			keyPress("SSCommandWndClass1", "enter", InfoFlexSelectModule)

			if (MRN = "")
			{
				return True
			}
			
			; Open patient search		
			WinActivate, %InfoFlexTitle%
			GraphicWait("InfoFlexFindPatient", 600000,, 10)

			Loop 10
			{
				GraphicClick("InfoFlexFindPatient", 0, 0,, 10, False)
			
				if WinExist(InfoFlexSubjectSearch)
				{
					break
				}	

				sleep 200
			}

		}
		else
		{
			WinActivate, % InfoFlexTitle
			
			; InfoFlex is locked
			if WinExist(InfoFlexUnlockTitle)
			{
				if !GetCredentials(system)
				{
					return False
				}
				ControlSetText, ThunderRT6TextBox2, % password, % InfoFlexUnlockTitle	
				PostClick(5, 5, "Unlock", InfoFlexUnlockTitle)
			}
			
			if WinExist(InfoFlexSubjectSearch)
			{
				WinClose, % InfoFlexSubjectSearch
				sleep 200
			}
			
			if (MRN = "")
			{
				return True
			}
			
			; Open patient search
			GraphicWait("InfoFlexFindPatient", 60000,, 5)
			GraphicClick("InfoFlexFindPatient", 0, 0,, 5)

			Loop 10
			{
				; Wait for if the save dialogue appears
				if WinExist(InfoFlexDataEntry)
				{
					try
						PostClick(5, 5, "Button1", InfoFlexDataEntry)
						
					try ; needed for situation when no subfolders in 'SAP - Sleep apnoea' folder
						PostClick(5, 5, "SSCommandWndClass3", InfoFlexDataEntry)
				}
			
				if WinExist(InfoFlexSubjectSearch)
				{
					break
				}

				sleep 200		
			}
		}

		;  Search for patient
		WinWait, % InfoFlexSubjectSearch
		WinActivate, % InfoFlexSubjectSearch

		ControlSetText, ThunderRT6TextBox4, % MRN, % InfoFlexSubjectSearch
		ControlSetText, ThunderRT6TextBox5,, % InfoFlexSubjectSearch
		ControlSetText, ThunderRT6TextBox6,, % InfoFlexSubjectSearch
		ControlSetText, ThunderRT6TextBox3,, % InfoFlexSubjectSearch
		ControlSetText, ThunderRT6TextBox7,, % InfoFlexSubjectSearch
		keyPressWindow("enter", InfoFlexSubjectSearch) ; ControlClick cannot do this step


		; Check if patient found or not
		; Cannot use "WinGet, Number, count, %InfoFlexSubjectSearch%" to see if "No patient found" dialogue is shown! Always shows 1 even if 2 windows with same name. Cannot even test if "&OK" button is present (returns 0)
		loop 10
		{
			if !WinExist(InfoFlexSubjectSearch) 
			{
				PatientFound := True
				break
			}

			sleep 200
		}

		if(PatientFound == False)
		{
			MB("Patient could not be found on InfoFlex! Please check patient's MRN")
			return False
		}
		
		WinActivate, % InfoFlexTitle
		
		if IF_ExitBeforeSelectAndZoom
			return True
		
		; Select dataview
		if !SelectDataView("CL - Respiratory")
			return False

		; Click 'no' if save dialogue appears again
		Loop 3
		{
			if WinExist(InfoFlexDataEntry)
			{
				try
					PostClick(5, 5, "Button1", InfoFlexDataEntry)
					
				try ; needed for situation when no subfolders in 'SAP - Sleep apnoea' folder
					PostClick(5, 5, "SSCommandWndClass3", InfoFlexDataEntry)
			}

			sleep 200
		}
	}
	catch, err
	{
		errorhandler(err, "InfoFlex", "InfoFlex")
		return False
	}
	

	; No real issue if any of the below functions fail as only clicking on last letters
	; and zooming in. Hence no catch statement
	try
	{
		; Click on 'Patient Demographics' tree node
		ControlGetPos, xctrl, yctrl,,, SSTreeWndClass3, % InfoFlexTitle
		MouseMove, xctrl + 29, yctrl + 5
		MouseClick
		keyPress("SSTreeWndClass3", "Down 40", InfoFlexTitle)
		keyPress("SSTreeWndClass3", "Space", InfoFlexTitle)


		; Zoom in
		if Settings["IFZoom"]
		{
			; Next function is to test if there are any letters
			if GraphicWait("InfoFlexFindInDocument", 50,, 5, False)
			{
				if GraphicWait("InfoFlexFindInDocument", 10000,, 5, False)
				{
					GlobalInputLockSet(True)
					GraphicClick("InfoFlexFindInDocument", 100, 5,, 5, False)
					GraphicWait("InfoFlexPageWidth", 1000,, 20, False)
					GraphicClick("InfoFlexPageWidth", 5, 5,, 50, False)
					GlobalInputLockSet(False)
				}
			}
		}
		

		; Click 'no' if save dialogue appears again
		Loop 3
		{
			if WinExist(InfoFlexDataEntry)
			{
				; Press cancel button
				ControlClick, Button2, % InfoFlexDataEntry
				ControlClick, SSCommandWndClass3, % InfoFlexDataEntry ; !!! review what this is
				break
			}

			sleep 200
		}
	}
	
	return True
}




; Only opens InfoFlex and searches for patient. Does not open letters
InfoFlexStartupAndSearch(MRN)
{
	global
	local system := "InfoFlex"
	local PatientFound := False
	FormatTime, Today, ,dd/MM/yyyy
	
	SetTitleMatchMode, 1
	SetKeyDelay, 10, 10


	try
	{
		; Handle if InfoFlex is at Select Module dialogue
		if WinExist(InfoFlexSelectModule)
		{
			keyPress("SSCommandWndClass1", "enter", InfoFlexSelectModule)
			WinWait, % InfoFlexTitle
		}

		; Start InfoFlex if not already running or start to input credentials on Login screen
		if !WinExist(InfoFlexTitle) or ControlPresent("ThunderRT6TextBox2", InfoFlexLoginTitle) == True
		{	
			if !GetCredentials(system)
			{
				return False
			}

			UpdateProgressBar(10,, PBMainString . "Opening InfoFlex...")
		
			if !WinExist(InfoFlexTitle) AND !WinExist(InfoFlexLoginTitle)
			{
				; Start InfoFlex
				Try
				{
					Run, % InfoFlexEXE
				}
				Catch err
				{
					ErrorMessage := "Line: " . err.Line . ", Message: " . err.Message . ", What: " . err.What . ", Extra: " . err.what . ", File: " . err.File
					MB("InfoFlex did not start correctly [" . ErrorMessage . "]. Stopping automation.",, "DevInform")
					return False
				}

				WinWait, %InfoFlexLoginTitle%
			}
		
			UpdateProgressBar(10,, PBMainString . "Entering credentials...")
		
			WinActivate, % InfoFlexLoginTitle
			GlobalInputLockSet(True)
			ControlSetText, ThunderRT6TextBox2, % username, % InfoFlexLoginTitle
			ControlSetText, ThunderRT6TextBox1, % password, % InfoFlexLoginTitle
			Send, {enter}
			GlobalInputLockSet(False)

			; Handle incorrect username or password
			Loop 50
			{
				if WinExist(InfoFlexLogonMessage)
				{
					MB("Loggin error. Stopping automation")
					DeleteCredentials(system)
					ControlClick, OK, %InfoFlexLogonMessage%
					return False
				}
			
				if WinExist(InfoFlexSelectModule)
				{
					break
				}

				sleep 200
			}

			UpdateProgressBar(10,, PBMainString . "Selecting data entry...")

			; Select data entry
			keyPress("SSCommandWndClass1", "enter", InfoFlexSelectModule)

			UpdateProgressBar(10,, PBMainString . "Searching for patient...")

			; Open patient search		
			WinActivate, % InfoFlexTitle
			GraphicWait("InfoFlexFindPatient", 600000,, 10)

			Loop 10
			{
				GraphicClick("InfoFlexFindPatient", 0, 0,, 10, False)
			
				if WinExist(InfoFlexSubjectSearch)
				{
					break
				}	

				sleep 200
			}

		}
		else
		{
			if WinExist(InfoFlexSubjectSearch)
			{
				WinClose, % InfoFlexSubjectSearch
				sleep 200
			}

			
			UpdateProgressBar(40, "absolute", PBMainString . "Searching for patient...")


			; If already on the patient demographics page, exit this subfunction (no need to search for patient again)
			if (GrabText("ThunderRT6TextBox25", InfoFlexTitle) = MRN)
			{
				return True
			}
			
			; Open patient search
			WinActivate, % InfoFlexTitle
			GraphicWait("InfoFlexFindPatient", 600000,, 5)
			GraphicClick("InfoFlexFindPatient", 0, 0,, 5)

			Loop 10
			{
				; Wait for if the save dialogue appears
				if WinExist(InfoFlexDataEntry)
				{
					PostClick(5, 5, "Button1", InfoFlexDataEntry)
					break
				}
			
				if WinExist(InfoFlexSubjectSearch)
				{
					break
				}

				sleep 200		
			}
		}

		;  Search for patient
		WinWait, % InfoFlexSubjectSearch
		WinActivate, % InfoFlexSubjectSearch
		ControlSetText, ThunderRT6TextBox4,, % InfoFlexSubjectSearch
		ControlSetText, ThunderRT6TextBox4, % MRN, % InfoFlexSubjectSearch
		ControlSetText, ThunderRT6TextBox6,, % InfoFlexSubjectSearch
		ControlSetText, ThunderRT6TextBox3,, % InfoFlexSubjectSearch
		keyPressWindow("enter", InfoFlexSubjectSearch) ; ControlClick cannot do this step


		; Check if patient found or not
		; Cannot use "WinGet, Number, count, %InfoFlexSubjectSearch%" to see if "No patient found" dialogue is shown! Always shows 1 even if 2 windows with same name. Cannot even test if "&OK" button is present (returns 0)
		loop 10
		{
			if !WinExist(InfoFlexSubjectSearch) 
			{
				PatientFound := True
				break
			}

			sleep 200
		}

		if(PatientFound == False)
		{
			MB("Patient could not be found on InfoFlex! Please check patient's MRN")
			return False
		}

		UpdateProgressBar(10,, PBMainString . "Retrieving patient and GP demographics...")

		; Click on 'Patient Demographics' tree node
		ControlGetPos, xctrl, yctrl,,, SSTreeWndClass3, % InfoFlexTitle
		MouseMove, xctrl + 29, yctrl + 5
		MouseClick 


		; Click 'no' if save dialogue appears again
		Loop 3
		{
			if WinExist(InfoFlexDataEntry)
			{
				; Press cancel button
				ControlClick, Button2, %InfoFlexDataEntry%
				ControlClick, SSCommandWndClass3, %InfoFlexDataEntry%
				break
			}

			sleep 200
		}
	}
	catch, err
	{
		GlobalInputLockSet(False)
		errorhandler(err, "InfoFlex (startup and search)", "InfoFlex")
		return False
	}
	
	return True
}




; No longer used at Gloucester. Now using Trakcare to get patient demographics
GetPatientDetailsInfoFlex(MRN)
{
	global

	try
	{
		if !InfoFlexStartupAndSearch(MRN)
		{
			return False
		}
		
		;clearPatientDetails()
		PatientDetails := {}
		
		PatientDetails.name := GrabText("ThunderRT6TextBox28", InfoFlexTitle) . " " . GrabText("ThunderRT6TextBox29", InfoFlexTitle)
		PatientDetails.DOB := GrabText("ThunderRT6TextBox32", InfoFlexTitle)
		PatientDetails.gender := GrabText("ThunderRT6TextBox34", InfoFlexTitle)
		PatientDetails.address1 := GrabText("ThunderRT6TextBox36", InfoFlexTitle)
		PatientDetails.address2 := GrabText("ThunderRT6TextBox37", InfoFlexTitle)
		PatientDetails.address3 := GrabText("ThunderRT6TextBox38", InfoFlexTitle)
		PatientDetails.address4 := GrabText("ThunderRT6TextBox39", InfoFlexTitle)
		PatientDetails.postCode := GrabText("ThunderRT6TextBox40", InfoFlexTitle)
		PatientDetails.TeleNo := GrabText("ThunderRT6TextBox41", InfoFlexTitle)
		PatientDetails.MobNo := GrabText("ThunderRT6TextBox42", InfoFlexTitle)
		PatientDetails.NHSNumber := GrabText("ThunderRT6TextBox26", InfoFlexTitle)

		PatientDetails.GPName := GrabText("ThunderRT6TextBox45", InfoFlexTitle) . ", " . GrabText("ThunderRT6TextBox47", InfoFlexTitle)
		PatientDetails.GPAddress := GrabText("ThunderRT6TextBox48", InfoFlexTitle) . ", " . GrabText("ThunderRT6TextBox49", InfoFlexTitle) . ", " . GrabText("ThunderRT6TextBox50", InfoFlexTitle) . ", " . GrabText("ThunderRT6TextBox51", InfoFlexTitle)
		PatientDetails.GPPostCode := GrabText("ThunderRT6TextBox53", InfoFlexTitle)
	}
	catch, err
	{
		errorhandler(err, "InfoFlex (Demographics and GP details)", "InfoFlex")
		return False
	}

	if (PatientDetails.name == "")
	{
		MB("Error with patient demographics and GP details retrieval")
		return False
	}
	
	return True
}



