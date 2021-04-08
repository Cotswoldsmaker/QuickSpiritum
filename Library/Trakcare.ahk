; **********************
; Trakcare / PAS library
; **********************

; *********
; Variables
; *********

TrakcareAddress := "https://live.trakcare.glos.nhs.uk/trakcare/csp/logon.csp?LANGID=1"
TrakcareTitle := "ENGL-PRD-2018"
TC := "" ; IE class




; *********
; Functions
; *********

TrakcareSearch(MRN, detailsGrab := False)
{
	Global TC, TrakcareAddress, TrakcareTitle, username, password, Settings
	Global PBMainString


	Loop, 5		; Allow 5 errors before failing function
	{
		system := "Trakcare"
		TrakCareFound := False
		ElementFound := False
		Results := ""
		IE_PAS := ""
		TCLocked := False
		

		try
		{
			if !GetCredentials(system)
			{
				return False
			}
			
			if !IEWinExist(TrakcareTitle)
			{
				if !TrakcareStart(detailsGrab)
					return False
			}
			else ; if already running
			{
				; Grab already running session
				TC := new IE_control(,TrakcareTitle,,, "grab")

				if detailsGrab
				{
					WinGet, hWnd, ID, %TrakcareTitle%
					WinMinimize, ahk_id %hWnd%
				}
				
				; Determine if on Location selection page
				if TC.elementPresent(, "ID", "LocListLocationz1")
				{
					Goto, LoginLocationTC
				}

				; Determine if on Login page
				if TC.elementPresent(,, "USERNAME")
				{
					Goto, LoginTC
				}
				
				; Determine if session is locked
				TCLocked := TC.HTMLFind("TRAK_hidden","NameInnerHTML", "cMessage", "This session is locked")
				
				; Determine if session has expired if not locked
				if !TCLocked
					TCLocked := TC.HTMLFind(, "OuterHTML", "This session has expired. Please refresh")
				
				if (TCLocked == True)
				{
					TC.quit
					sleep 1000
					
					if !TrakcareStart(detailsGrab)
						return False
				}

				GoTo, ContinueTC		
			}

LoginTC:

			TC.busy()

			if !TC.elementPresent(,, "USERNAME")
			{
				Goto, LoginLocationTC
			}

			if detailsGrab 
				UpdateProgressBar(10, "absolute", PBMainString . "Logging in to Trakcare...")
			
			TC.set(,, "USERNAME", username)
			TC.set(,, "PASSWORD", password)
			TC.click(,, "Logon")
			
			; Check if login was successful: should be able to find LocListLocationz1 (searching for USERNAME as name returns true on the login location page)
			if !TC.elementPresent(, "ID", "LocListLocationz1")
			{
				MB("Loggin error. Stopping automation")
				WinClose, % TrakcareTitle
				DeleteCredentials(system)
				return False
			}
			
			if detailsGrab
				UpdateProgressBar(20, "absolute", PBMainString . "Trakcare loggin passed")

LoginLocationTC:

			TC.click(, "ID", "LocListLocationz" . Settings.LocListLocationz)

ContinueTC:

			if (MRN = "")
			{
				return True
			}
			
			if detailsGrab
				UpdateProgressBar(30, "absolute", PBMainString . "Searching for patient in Trakcare")
		
			; Favourite Patients
			TC.click("eprmenu", "ID", "MainMenuItemAnchor50367")

			; Patient Enquiry
			TC.click("eprmenu", "ID", "MainMenuItemAnchor50583")
			TC.set("TRAK_main",, "RegistrationNo", MRN)
			; Click OK if for example a deceased message comes up
			TC.clickOK()
			TC.click("TRAK_main", "ID", "find1")

			
			if TC.HTMLFind("TRAK_main", "TagNameInnerHTML", "h1", "Patient List")
			{
				if !TC.elementPresent("TRAK_main", "ID", "RegistrationNoz1")
				{
					MB("Patient was not found on Trakcare. Stopping Trakcare automation")
					return False
				}
				else
				{
					; Click on the first patient if a patient list is created
					TC.click("TRAK_main", "ID", "RegistrationNoz1")
				}
			}
			
			return True
		}
		catch, err
		{
			; Basically close everything and start again
			try
				TC.quitAll()
				
			sleep 200
				
			if (err.message != "IE fail" OR A_index >= 5)
			{
				errorhandler(err, "Trakcare", "Internet Explorer / Trakcare")
				return False
			}
			else
			{
				LogUpdate("Caught error trying to grab COM to run Trakcare [error counter: " . A_index . "]")
			}
		}
	}
	
	return True
}




TrakcareStart(detailsGrab)
{
	global
	
	Loop, 5 
	{
		try
		{
			; start up a new class and IE session
			TC := new IE_control(TrakcareAddress, TrakcareTitle, True)
			WinWait, % TrakcareTitle
			
			if detailsGrab
			{
				WinMaximize, % TrakcareTitle
				sleep 200
				WinMinimize, % TrakcareTitle ; minimise if grabbing patient demographics and GP details
			}
			else
			{
				WinActivate, % TrakcareTitle
			}
			
			return True
		}
		Catch, err
		{
			; Basically close everything and start again
			try
				TC.quit
				
			sleep 200
			
			if (err.message != "IE fail" OR A_index >= 4)
			{
				throw Exception("IE Start up fail", -1)
			}
			else
			{
				LogUpdate("Caught error trying to grab COM to startup Trakcare [error counter: " . A_index . "]")
			}
		}
	}
	
	return True
}




PASSubSearch(MRN)
{
	Global TC, TrakcareTitle
	
	Loop, 5	; Allow 5 errors before failing
	{
		system := "Trakcare"
		TrakCareFound := False
		ElementFound := False
		Results := ""
		
		try
		{
			if (TC == "")
			{
				Msgbox, Error connection to Trakcare for PAS functionality!
				Return False
			}

			TC.busy()

			If !TC.elementPresent("TRAK_main", "ID", "PAADMTypez1")
			{
				MB("No in/outpatient episodes to click on to then subsequently open PAS from. Stopping automation!")
				return True
			}

			TC.click("TRAK_main", "ID", "PAADMTypez1")
			; To click on OK on the dialogue message that may appear
			TC.clickOK()
			TC.click("eprmenu", "ID", "MainMenuItemAnchor50767")
			sleep 200
			TC.click("TRAK_main", "TagNameTitle", "img", "Historical Data Viewer")
			return True
		}
		catch, err
		{
			; Basically close everything and start again
			try
				TC.quit
				
			sleep 200
				
			if (err.message != "IE fail" OR A_index >= 5)
			{
				errorhandler(err, "PAS via Trakcare", "Internet Explorer / Trakcare")
				return False
			}
			else
			{
				LogUpdate("Caught error, trying to grab COM to run PAS (via Trakcare) again [counter: " . A_index . "]")
			}
		}
	}
	
	return True
}




GetPatientDetailsTrakcare(MRN)
{
	global TC, PatientDetails, PBMainString
	
	TC_details := ""
	
	try
	{		
		if (PatientDetails.MRN == MRN)
		{
			return True
		}
			
		if !TrakcareSearch(MRN, True)
		{
			return False
		}
		
		UpdateProgressBar(40, "absolute", PBMainString . "Grabbing patient demographics and GP details")
		PatientDetails := {}
		TC.busy()

		if !TC.elementPresent("TRAK_main", "ID", "MRNz1")
		{
			CloseProgressBar()
			MB("Patient does not have any IP/OP episodes. Cannot get patient demographics and GP details via automation for this patient")
			return False
		}

		; An MRSA message can pop up next, so needs closing
		TC.clickOK()
		TC.click("TRAK_main", "ID", "MRNz1")
		; Best way to grab the new window that opens !!! may have to improve incase some other window is made active at same time
		WinGet, hWnd, ID, A
		WinMinimize, ahk_id %hWnd%
		IE_details := WBGet("ahk_id " hWnd)
		TC_details := new IE_control(, TrakcareTitle,, IE_details) ; need title incase function has error and needs to close TC
		
		PatientDetails.MRN := TC_details.getValue(, "TagNameID", "label", "RegistrationNumber")
		PatientDetails.name := TC_details.getValue(, "ID", "PAPERName2") . " " . TC_details.getValue(, "ID", "PAPERName")
		PatientDetails.DOB := TC_details.getValue(, "ID", "PAPERDob")
		PatientDetails.gender := TC_details.getValue(, "ID", "CTSEXDesc")
		PatientDetails.address1 := TC_details.getValue(, "ID", "PAPERStNameLine1")
		PatientDetails.address2 := TC_details.getValue(, "ID", "PAPERForeignAddress")
		PatientDetails.address3 := TC_details.getValue(, "ID", "CTCITDesc")
		PatientDetails.address4 := TC_details.getValue(, "ID", "PROVDesc")
		PatientDetails.postCode := TC_details.getValue(, "ID", "CTZIPCode")
		PatientDetails.TeleNo := TC_details.getValue(, "ID", "PAPERTelH")
		PatientDetails.MobNo := TC_details.getValue(, "ID", "PAPERMobPhone")
		PatientDetails.NHSNumber := TC_details.getValue(, "ID", "PAPERID")
		PatientDetails.email := TC_details.getValue(, "ID", "PAPEREmail")

		PatientDetails.GPName := TC_details.getValue(, "ID", "REFDDesc")
		PatientDetails.GPAddress := TC_details.getValue(, "ID", "CLNAddress1")
		PatientDetails.GPPostCode := ""
	}
	catch, err
	{
		CloseProgressBar()
		errorhandler(err, "Trakcare (Demographics and GP details)", "Trakcare")
		TC_details.quit()
		return False
	}

	
	if (PatientDetails.name == "")
	{
		CloseProgressBar()
		ErrorHandler(, "Trakcare", "Internet Explorer / Trakcare", "Error with missing patient name whilst getting patient demographics and GP details.")
		TC_details.quit()
		return False
	}
	else if (PatientDetails.MRN != MRN)
	{
		CloseProgressBar()
		ErrorHandler(, "Trakcare", "Internet Explorer / Trakcare", "Error with MRN mismatch whilst getting demographics and GP details.")
		TC_details.quit()
		return False
	}
	
	; Testing function - False stops this running
	if False
	{

		PatientString := ""
		
		for key, value in PatientDetails
		{
			PatientString .= key . " - " . value . "`n"
		}
		msgbox, % PatientString
	}
	
	TC_details.quit()
	return True
}



