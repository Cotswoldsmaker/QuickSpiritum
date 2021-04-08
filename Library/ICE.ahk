; ***********
; ICE library
; ***********

; ***********
; Variables
; ***********

ICE := ""
ICETitle := "Sunquest Ice Desktop"
ICEAddress := "https://ice.glos.nhs.uk/icedesktop/"
ICELoginTitle := "Sunquest Login"




; *********
; Functions
; *********

ICESearch(MRN)
{
	Global ICE, ICETitle, ICEAddress, ICELoginTitle, username, password
	
	Loop, 5 	; Give 5 chances for error to happen and then restart ICE	
	{
		system := "ICE"
		ICEFound := False
		ElementFound := False
		Results := ""
		PatientFound := False
		ICETitlePresent := False
		counter := 0
		
		try
		{
			; Close 1st screen on starting up ICE if present (complicated process as this screen has a windows title but not an IE.LocationName)
			if WinExist(ICETitle)
			{
				ICETitlePresent := True
			}

			For IE in ComObjCreate("Shell.Application").Windows ; for each open window
			{
				If InStr(IE.FullName, "iexplore.exe") && InStr(IE.LocationName, ICETitle) ; check if it's an ie window
				{
					ICEFound := True
					break
				}
			}

			if !ICEFound AND ICETitlePresent
			{
				WinClose, % ICETitle
				Sleep 200
			}

			
			; Close login screen if present before starting		
			if WinExist(ICELoginTitle) ; IEWinExist did not work here
			{
				WinClose, % ICELoginTitle
				Sleep 200
			}



			ICETitlePresent := False
			ICEFound := False

			if !IEWinExist(ICETitle) AND !IEWinExist(ICELoginTitle)
			{
				if !GetCredentials(system)
				{
					return False
				}
		
				; create a InternetExplorerMedium instance
				ICE := new IE_control(ICEAddress, ICETitle)
				ICE.click(, "TagNameSRC", "img", "images/USER.jpg")
				ICE.quit()
				
				; Wait until the login window is present, fail after 10 seconds
				Loop, 200
				{
					if WinExist(ICELoginTitle)
					{
						break
					}
					
					sleep 50
				}
			}
			


			if WinExist(ICELoginTitle)
			{
				ICE := new IE_control(,ICELoginTitle,,, "grab")
				ICE.set(,, "txtName", username)
				ICE.set(,, "txtPassword", password)
				ICE.click(,, "btnSubmit")

				ICE.click(, "ID", "btnExpireWarning")
				; !!! Need to check if still need this
				;if ICE.elementPresent(, "ID", "btnExpireWarning")
				;{
				;	MB("Password update needed")
				;	return False
				;}
				
				if ICE.elementPresent(,, "txtName")
				{
					MB("Loggin error. Stopping automation")
					WinClose, %ICELoginTitle%
					DeleteCredentials(system)
					return False
				}	
			}
			else ; already logged in
			{			
				ICE := new IE_control(,ICETitle,,, "grab")

				; Take page back to search screen
				ICE.click(, "TagNameSRC", "img", "images/group.gif")
			}

			WinActivate, %ICETitle%
			ICE.busy()

			
			while !ICE.elementPresent("Right",, "PatientSearch1$txtSearch") AND counter < 10
			{
				sleep 200
				counter := counter + 1
				
				if ICE.HTMLFind(, "OuterHTML","Your password has expired. Please enter a new one now.")
				{
					MB("Password update detected. Stopping ICE automation. Please update your password and try ICE automation again")
					DeleteCredentials(system)
					return False
				}
			}

			if (counter == 10)
			{
				throw Exception("MRN search field load fail", -1)
			}

			ICE.set("Right",, "PatientSearch1$txtSearch", "MRN" . MRN)
			ICE.click("Right",, "PatientSearch1_rdoSrchHosp")
			ICE.click("Right",, "PatientSearch1$btnSearch")

			if !ICE.click("Right", "TagNameclassname", "TR", " even")
			{
				MB("Patient not found on ICE. Stopping automation!")
				return True
			}
			
			ICE.click(, "TagNameSRC","img", "images/toolbar/icons/ToolbarItem18.gif")

			return True
		}
		Catch, err
		{
			; Basically close everything and start again
			try
				ICE.quitAll()
			
			sleep 200
			
			GroupAdd, ICEGroup, % ICETitle
			GroupAdd, ICEGroup, % ICELoginTitle
			WinClose, ahk_group ICEGroup
				
			if (err.message != "IE fail" OR A_index >= 5)
			{
				errorhandler(err, "ICE", "Internet Explorer / ICE")
				return False
			}
			else
			{
				LogUpdate("Caught error trying to grab COM to run ICE [error counter: " . A_index . "]")
			}
		}
	}

	return False
}



; From IPC, click on requests and fill in clinical details
ICERequests(rawData)
{
	global ICE, PBMainString, secretaryExt
	
	CloseProgressBar()
	PBMainString := "ICE requests:`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Starting…")


	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	CT_head := DS[2]
	MRI_head := DS[3]
	ECHO := DS[4]
	CT_guided_biopsy := DS[5]
	clinician := DS[6]
	clinicalDetails := DS[7]


	UpdateProgressBar(10, "absolute", PBMainString . "Searching for patient...")	
	ICESearch(MRN)
	UpdateProgressBar(10, "absolute", PBMainString . "Navigating...")
	IE := ICE.getPointer()
	ICEControl := new ICEControl(IE)
	
	if CT_head
	{
		ICEControl.testsWaitAndClick("CT head")
		UpdateProgressBar(15, "absolute", PBMainString . "CT-head selected")
	}
	
	if MRI_head
	{
		ICEControl.testsWaitAndClick("MRI Examinations of the Head")
		ICEControl.testsWaitAndClick("MRI head")
		UpdateProgressBar(30, "absolute", PBMainString . "MRI-head selected")
	}
	
	if ECHO
	{
		ICEControl.testsPageClick("ECHO")
		ICEControl.testsWaitAndClick("Trans-Thoracic Echocardiogram")
		UpdateProgressBar(45, "absolute", PBMainString . "ECHO selected")
	}
	
	if CT_guided_biopsy
	{
		ICEControl.testsPageClick("Search")
		ICEControl.search("CT Guided biopsy")
		UpdateProgressBar(55, "absolute", PBMainString . "Please complete 'Rules -- Web page Dialogue' to continue")
		ICEControl.testsWaitAndClick("CT Guided biopsy", True)
		UpdateProgressBar(65, "absolute", PBMainString . "CT guided biopsy selected")
	}
	

	ICEControl.testsContinue()
	
	UpdateProgressBar(75, "absolute", PBMainString . "Updating final details")
	ICEControl.bleep(secretaryExt)
	ICEControl.priority("Two-Week Wait (RAD)")
	ICEControl.requestingDoctor(clinician)
	ICEControl.location("GRH Medical OPD")
	ICEControl.globalClinicalDetails(clinicalDetails)
		
	UpdateProgressBar(100, "absolute", "Please continue to submit requests on ICE manually")
	sleep 4000
	CloseProgressBar()
	return "Pass"
}




; Addition IE control but just for ICE
class ICEControl
{
	; Class variable initialisation
	IE := ""
	
	
	
	
	; initialise the class and IE session
    __New(IE)
    {
		this.IE := IE
		this.busy()
    }
	
	
	
	
	; Get the IE pointer
	getPointer()
	{
		return this.IE
	}
	
	
	
	
	busy()
	{
		sleepAmount := 50

		Loop, 5
		{
			try
			{
				while this.IE.busy || this.IE.ReadyState != 4
					sleep %sleepAmount%

				return 2000
			}
			catch
			{
				sleep 200
			}
		}

		throw Exception("IE fail", -1)
	}




	testsPageClick(page)
	{
		Loop, 1000
		{
			try
			{
				Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.frames("ifPages").document.GetElementsByTagName("TD")
				Results := Results[0].GetElementsByTagName("TD")
				
				Loop % Results.length
				{
					if (InStr(Results[A_index-1].outerHTML, page))
					{
						Results[A_index-1].click()
						this.busy()
						return True
					}
				}
			}
			
			sleep 50
		}
		
		return False	
	}
	
	
	
	
	testsWaitAndClick(Test, waitIndefinitely := False)
	{
		counter := 0
		
		Loop
		{
			try
			{
				Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.frames("ifTests").document.GetElementsByTagName("TD")
				
				Loop % Results.length
				{
					if (InStr(Results[A_index-1].outerHTML, Test))
					{
						Results[A_index-1].click()
						this.busy()
						return True
					}
				}
			}
			
			if (!waitIndefinitely AND counter >= 1000)
				break
			
			sleep 50
		}
		
		return False
	}
	
	
	
	
	search(value)
	{
		found := False
		
		; Enter serach value
		Loop, 1000
		{
			try
			{
				Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.GetElementsByTagName("INPUT")
				
				Loop % Results.length
				{
					if (InStr(Results[A_index-1].id, "txtSearch"))
					{
						Results[A_index-1].value := value
						this.busy()
						found := True
						break
					}
				}
				
				if found	
					break
			}
			
			sleep 50
		}
		
		
		; Press 'Search button'
		Loop, 1000
		{
			try
			{
				Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.GetElementsByTagName("INPUT")
				
				Loop % Results.length
				{
					if (InStr(Results[A_index-1].id, "btnSearch"))
					{
						Results[A_index-1].click()
						this.busy()
						return True
					}
				}
			}
			
			sleep 50
		}
		
		return False
	}
	
	
	
	
	testsContinue()
	{
		Loop, 1000
		{
			try
			{
				Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.frames("ifPages").document.GetElementsByTagName("Button")
				
				Loop % Results.length
				{
					if (InStr(Results[A_index-1].outerHTML, "Continue<BR>with<BR>request..."))
					{
						Results[A_index-1].click()
						this.busy()
						return True
					}
				}
			}
			
			sleep 50
		}
		
		return False	
	}
	
	
	
	
	bleep(bleep)
	{
		Loop, 1000
		{
			try
			{
				Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.GetElementsByTagName("INPUT")
				
				Loop % Results.length
				{
					if (InStr(Results[A_index-1].id, "txtBleep"))
					{
						Results[A_index-1].value := bleep
						this.busy()
						return True
					}
				}
			}
			
			sleep 50
		}
		
		return False
	}




	priority(Priority)
	{
		value := 0

		if (Priority == "Two-Week Wait (RAD)")
		{
			value := 6
		}
	
		Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.GetElementsByTagName("SELECT")
		Loop % Results.length
		{
			if (InStr(Results[A_index-1].name, "Priority"))
			{
				Results[A_index-1].value := value
				this.busy()
				; Don't use return here as want to get all priority boxes
			}
		}

		return False
	}




	requestingDoctor(doctor)
	{
		Loop, 1000
		{
			try
			{
				Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.GetElementsByTagName("SELECT")
				
				Loop % Results.length
				{
					if (InStr(Results[A_index-1].id, "ddlReqPhy"))
					{
						SuperIndex := A_index
						ResultsO := Results[A_index-1].document.GetElementsByTagName("OPTION")
						
						Loop % ResultsO.length
						{
							if (InStr(ResultsO[A_index-1].innerHTML, doctor))
							{
								Results[SuperIndex-1].value := ResultsO[A_index-1].value
								return True
							}
						}
					
						return True
					}
				}
			}
			
			sleep 50
		}
		
		return False
	}
	
	
	
	
	location(location)
	{
		Loop, 1000
		{
			try
			{
				Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.GetElementsByTagName("SELECT")
				
				Loop % Results.length
				{
					if (InStr(Results[A_index-1].id, "ddlLocation"))
					{
						SuperIndex := A_index
						ResultsO := Results[A_index-1].document.GetElementsByTagName("OPTION")
						
						Loop % ResultsO.length
						{
							if (InStr(ResultsO[A_index-1].innerHTML, location))
							{
								Results[SuperIndex-1].value := ResultsO[A_index-1].value
								return True
							}
						}

						return True
					}
				}
			}
			
			sleep 50
		}
		
		return False
	}
	
	
	
	
	globalClinicalDetails(details)
	{
		Loop, 1000
		{
			try
			{
				Results := this.IE.document.parentWindow.frames("iframeMain").document.frames("appFrame").document.GetElementsByTagName("TEXTAREA")
				
				Loop % Results.length
				{
					if (InStr(Results[A_index-1].id, "txtGCD"))
					{
						Results[A_index-1].value := details
						this.busy()
						return True
					}
				}
			}
			
			sleep 50
		}
		
		return False
	}
}



