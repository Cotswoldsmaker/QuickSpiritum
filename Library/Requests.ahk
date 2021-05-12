; ***********************************************************************
; Requests and referrals via
; ***********************************************************************


; *********
; Variables
; *********

PatientDetails := {}
requestList := ["PET_CT", "PFT", "bronchoscopy"]
referralList := ["HLSG", "Sleepstation"]
RR_suffix := ""
ccEmail := ""
RRDetails := {}




; PET-CT request variables
;PET_CT_template_path := Settings.TemplatesFolder . "PET_CT_request_template.docx"
PET_CT_latestRequestPath := ""	
PET_CT_consultantName := ""
PET_CT_doctorName := ""
PET_CT_previousImagingType := ""
PET_CT_previousImagingDate := ""
PET_CT_clinicalInformation := ""
PET_CT_diabetic := ""
PET_CT_emailLungCancerCoordinators := 1 	; 1 = checked, 0 = unchecked
PET_CTButtonPressed := ""




; Lung function test request variables
PFT_templatePath := Settings.TemplatesFolder . "PFT_request_template.docx"
PET_latestRequestPath := ""
PFT_consultantName := ""
PFT_doctorName := ""
PFT_clinicalInformation := ""
PFT_buttonPressed := ""
PFT_speciality := ""
PFT_location := ""
PFT_wardText := ""
PFT_ward := ""
PFT_urgency := ""
PFT_clinicDateText := ""
PFT_clinicDate := ""

PFT_tests := FiFoArray("Withhold bronchodilators prior", "holdBronchodilators"
		 , "Spirometry / flow loop" , "spirometry"
		 , "Gas transfer (TLCO & KCO)", "gasTransfer"
		 , "Static lung volumes", "staticLungVolumes"
		 , "Exhaled nitric oxide (FENO)", "FENO"
		 , "Capillary blood gases on air", "CBG"
		 , "Sitting / supine spirometry", "sittingStanding"
		 , "Bronchodilator response", "bronchodilatorResponse"
		 , "Osmohale (Mannitol)", "osmohale"
		 , "Hypoxic challenge (fit to fly)", "fitToFly"
		 , "Multi-channel sleep study", "SS"
		 , "CPAP trial", "CPAP"
		 , "Mouth pressures", "mouthPressures"
		 , "Overnight pulse oximetry", "overnightPulseOx"
		 , "6 week occupational asthma study", "occupationalAsthmaStudy"
		 , "Home NIV", "NIV")

holdBronchodilators := 1
spirometry := 0
gasTransfer := 0
staticLungVolumes := 0
FENO := 0
CBG := 0
sittingStanding := 0
bronchodilatorResponse := 0
osmohale := 0
fitToFly := 0
SS := 0
CPAP := 0
mouthPressures := 0
overnightPulseOx := 0
occupationalAsthmaStudy := 0
NIV := 0
NIVMode := 0
NIVIPAP := 0
NIVEPAP := 0




; Bronchoscopy request variables
bronch_templatePath := Settings.TemplatesFolder . "Bronchoscopy_request_template.docx"
bronch_latestRequestPath := ""
bronch_procedure := ""
bronch_location := ""
bronch_primaryDiagnosis := ""
bronch_significantComorbidities := ""
bronch_preferredBronchoscopyDate := ""
bronch_consentCompleted := ""
bronch_anticoagulantAntiplatelets := ""
bronch_aim := ""
bronch_imaging := ""
bronch_consultantName := ""
bronch_doctorName := ""
bronch_emailCoordinators := 1 	; 1 = checked, 0 = unchecked




;HLSG
HLSG_templatePath := Settings.TemplatesFolder . "HLSG_request_template.docx"
HLSG_latestRequestPath := ""
HLSG_doctorName := ""
HLSG_ethnicity := ""
HLSG_Comorbidities := ""
HLSG_buttonPressed := ""
HLSG_referralTypes := FiFoArray("Smoking cessation", "HLSG_smoking"
								, "Alcohol services" , "HLSG_alcohol",
								, "Weight management", "HLSG_weight",
								, "Physical activity", "HLSG_physicalActivity")

HLSG_smoking := 0
HLSG_alcohol := 0
HLSG_weight := 0
HLSG_physicalActivity := 0




;SleepStation
Sleepstation_template_path := Settings.TemplatesFolder . "Sleepstation_referral_template.docx"
SleepStation_latestRequestPath := ""
Sleepstation_buttonPressed := ""




; Test email
if emailTest
{
	PET_CT_Email := TestEmail
	PET_CT_Email_cc := ""
	PFT_Email := TestEmail
	bronchoscopy_Email := TestEmail
	bronchoscopy_Email_cc := ""
	HLSG_Email := TestEmail
	SleepStation_Email := TestEmail
}




; *********
; Functions
; *********



RRCreateAndSend(RRtypeTemp)
{
	global
	RRDetails := {}
	RRtype := RRtypeTemp
	ccEmail := ""
	RRMessageAddSuffix()
	local RRFunctionCR := "create_" . RRtype . "_" . RR_suffix
	local index, element
	
	
	if !registeredUsername()
		return False

	if runningStatus()
		return False
		
	if !GrabMRN(RRMessage)
		return False
		

	CloseProgressBar()
	PBMainString := RRMessage . ":`n"
	CreateProgressBar()
	UpdateProgressBar(0, "absolute", PBMainString . "Starting...")
	sleep 2000

	if GetPatientDetailsTrakcare(MRN)
	{
		CloseProgressBar()

		if RRGetExtraInfo()
		{
			CreateProgressBar()
			UpdateProgressBar(60, "absolute", PBMainString . "Creating " . RR_suffix . "...")

			if RRCreatePDF() ;if %RRFunctionCR%()
			{
				UpdateProgressBar(20,, PBMainString . "Emailing referral...")
				RREmail()
			}
		}
	}
	
	CloseProgressBar()
	runningStatus("done")
	LogUpdate("Sleepstation request")
	return True
}




RRMessageAddSuffix()
{
	global
	RRMessage := RRtype
	
	
	; Add request or referral to end
	for index, element in requestList
	{
		if (RRMessage == element)
		{
			RR_suffix := "request"
			RRMessage .= " request"
			break
		}
	}
	
	if (RRMessage == RRtype)
	{
		for index, element in referralList
		{
			if (RRMessage == element)
			{
				RR_suffix := "referral"
				RRMessage .= " referral"
				break
			}
		}
		
		if (RRMessage == RRtype)
		{
			RR_suffix := ""
			throw Exception("'RRtype' has not been matched up with either a request or referral!", -1)
		}
	}
}




RRGetExtraInfo()
{
	global

	if (%RRtype%_GUI() != "OK")
	{
		return False
	}

	if !FileExist(RRdetails.signature)
	{
		MB("No signature found for selected doctor. Ending request/referral")
		return False
	}

	return True
}




RRCreatePDF()
{
	global
	local dateToday := ""
	FormatTime, dateToday,, dd/MM/yy
	local signatureWidth := signatureHeight * GraphicsDimensions(RRdetails.signature)
	
	Loop, 5		; Allow 5 errors before failing function
	{
		RRLatestSavePath := Settings.RequestsFolder . "\" . RRMessage . " MRN" . MRN . "_"

		try
		{
			Loop
			{
				RRLatestSavePathTemp := RRLatestSavePath . A_Index . ".pdf"

				if(FileExist(RRLatestSavePathTemp) == "")
				{
					RRLatestSavePath := RRLatestSavePathTemp
					break
				}
			}

			oWord := ComObjCreate("Word.Application")
			oWord.Visible := False
			oWord.Documents.Open(Settings.TemplatesFolder . RRType . "_" . RR_suffix . "_template.docx")


			for key, value in PatientDetails
				try
					WriteAtBookmark(key, value)


			for key, value in RRdetails
			{
				if (instr(key, "signature") == 1)
				{
					WriteAtBookmark(key, RRdetails.doctorName)
					oWord.ActiveDocument.Bookmarks(key).Select
					picObj := oWord.ActiveDocument.InLineShapes.AddPicture(value)
					picObj.Height := signatureHeight 
					picObj.Width := signatureWidth
				}
				else if (InStr(key, "CB_") == 1)
				{
					InsertCheckboxAtBookmark(SubStr(key,4), value)
				}
				else if (key != "doctorName")
				{
					WriteAtBookmark(key, value)
				}
			}

			try WriteAtBookmark("dateToday", dateToday)
			
			UpdateProgressBar(80, "absolute", PBMainString . "Saving to requests folder")
			oWord.ActiveDocument.SaveAs(RRLatestSavePath, 17) ; 17 is PDF format
			oWord.ActiveDocument.close(False)
			
			return True
		
		}
		catch, err
		{
			try
				oWord.ActiveDocument.close(False)
				
			if (A_index >= 5)
			{
				errorhandler(err, RRMessage, "Microsoft Word")
				return False
			}			
			else
			{
				LogUpdate("Error during " . RRMessage . "[counter: " . A_index . "]")
			}
		}
	}
	
	return False
}




RREmail()
{
	global 
	local RREmail := RRtype . "_Email"
	
	if (CitrixSession == True)
	{
		UpdateProgressBar(100, "absolute", PBMainString . "Currently cannot send via a Citrix session. Request / referral can be found in the Request folder (see Settings under F1)", True)
		;NHSMailOpen()
		;Email_NHS_Mail(%RREmail%, "", RRMessage, "Please find attached a " . RRMessage, %RRLatestRequestPath%)
	}
	else
	{
		if EmailOutlook(%RREmail%, ccEmail, RRMessage, "Please find attached a " . RRMessage, RRLatestSavePath)
		{
			UpdateProgressBar(100, "absolute", PBMainString . "Complete")
			
			Loop , 15
			{
				if !WinExist(PBTitle)
					break
					
				sleep 200
			}
		}	
	}
	return
}




; PET-CT request via IPC
send_PET_CT(rawData)
{
	global
	RRDetails := {}
	RRType := "PET_CT"
	local Pass := False
	local DS := ""
	local emailAddress := ""
	local cc_addresses := ""
	
	CloseProgressBar()
	PBMainString := "PET-CT request:`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Starting…")

	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	RRDetails.consultantName := DS[2]
	RRDetails.doctorName := DS[3]
	RRDetails.signature := CurrentDirectory . "Signatures\" . DS[3] . ".jpg"
	RRDetails.MDTDate := DS[4]
	RRDetails.previousImagingType := DS[5]
	RRDetails.previousImagingDate := DS[6]
	RRDetails.clinicalInformation := DS[7]
	RRDetails.diabetic := DS[8]

	
	if (DS[9] = "testEmail")
	{
		emailAddress := TestEmail
		PET_CT_emailLungCancerCoordinators := 0
	}
	else
	{
		emailAddress := PET_CT_Email
		PET_CT_emailLungCancerCoordinators := 1
	}


	if GetPatientDetailsTrakcare(MRN)
	{
		UpdateProgressBar(60, "absolute", PBMainString . "Creating request…")

		if RRCreatePDF ;if Create_PET_CT_request()
		{
			UpdateProgressBar(20,, PBMainString . "Emailing PET-CT request…")
			
			if (PET_CT_emailLungCancerCoordinators == 1)
				cc_addresses := PET_CT_Email_cc
			else
				cc_addresses := ""
			
			if (CitrixSession == True)
			{
				MB("Current cannot send via a Citrix session. Please manually locate your request in the requests folder in the master directory and send via the NHSmail web app")
				;NHSMailOpen()
				;Email_NHS_Mail(emailAddress, cc_addresses, "PET-CT request", "Please find attached a PET-CT request", PET_CT_latestRequestPath)
				; pass := True
			}
			else
			{
				if EmailOutlook(emailAddress, cc_addresses, "PET-CT request", "Please find attached a PET-CT request", PET_CT_latestRequestPath)
				{
					pass := True
				}
			}
		}
	}

	CloseProgressBar()
	
	if pass
	{
		return "Pass"
	}
	else
	{
		return "Fail"
	}
}




PET_CT_GUI()
{
	global
	local GUIStart := 237
	local GUIWidth := 400
	local yAdditive := 10
	local GUI_name := "PET_CT_GUI"
	local H := 0 ; Height of field
	
	Gui, %GUI_name%:Font, s12
	Gui, %GUI_name%:Color, %dialogueColour%
	
	H := 40
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Consultant:
	Gui, %GUI_name%:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_consultantName, %consultantList%
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_doctorName, %doctorList%
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Previous imaging type:
	Gui, %GUI_name%:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_previousImagingType, CT chest||CT body|MRI
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Previous imaging date:
	Gui, %GUI_name%:Add, DateTime, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_previousImagingDate
	yAdditive += H + S
	
	H := 300
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Clinical information:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W600 H%H% vPET_CT_clinicalInformation
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Diabetic:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_diabetic, Unknown||No|Yes
	yAdditive += H + S
	
	Gui, %GUI_name%:Add, CheckBox, x5 y%yAdditive%  vPET_CT_emailLungCancerCoordinators +Right Checked, Email lung cancer coordinators: 

	Gui, %GUI_name%:Add, Button, x800 y%yAdditive% default gPET_CT_OK, &OK
	Gui, %GUI_name%:Add, Button, x840 y%yAdditive% gPET_CT_close,  &Cancel
	Gui, %GUI_name%:Show,, PET-CT request details
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, PET-CT request details
	return PET_CT_ButtonPressed
}




PET_CT_OK()
{
	global
	errorString := ""
	pass := False
	PET_CT_ButtonPressed := "OK"
	Gui PET_CT_GUI:Submit, NoHide

	if (PET_CT_consultantName == "")
		errorString := errorString . "- No consultant name was chosen.`n"
		
	if (PET_CT_doctorName == "")
		errorString := errorString . "- No requesting doctor was chosen.`n"

	if (PET_CT_clinicalInformation == "")
		errorString := errorString . "- No clinical information was given.`n"

	if (errorString != "")
	{
		MB("Please correct the below errors and submit again:`n" . errorString)
		return
	}

	Gui, PET_CT_GUI:Destroy
	RRDetails.consultantName := PET_CT_consultantName
	RRDetails.doctorName := PET_CT_doctorName
	RRDetails.signature := CurrentDirectory . "Signatures\" . PET_CT_doctorName . ".jpg"
	Yesterday += -1, Days
	FormatTime, Yesterday, %Yesterday%, dd/MM/yy
	RRDetails.MDTDate := Yesterday
	RRDetails.previousImagingType := PET_CT_previousImagingType
	FormatTime, PET_CT_previousImagingDate, %PET_CT_previousImagingDate%, dd/MM/yy
	RRDetails.previousImagingDate := PET_CT_previousImagingDate
	RRDetails.previousImagingDate := PET_CT_previousImagingDate
	
	RRDetails.clinicalInformation := PET_CT_clinicalInformation
	RRDetails.diabetic := PET_CT_diabetic


	if (PET_CT_emailLungCancerCoordinators == 1)
		ccEmail := PET_CT_Email_cc
	else
		ccEmail := ""

	return
}




PET_CT_GUIGuiClose()
{
	PET_CT_close()
}
PET_CT_close()
{
	PET_CT_ButtonPressed := "close"
	Gui, PET_CT_GUI:Destroy
	return
}




; PFT request via IPC
send_PFT(rawData)
{
	global
	RRType := "PFT"
	local Pass := False
	local DS := ""
	local emailAddress := ""


	CloseProgressBar()
	PBMainString := "Lung function request:`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Starting…")
	
	
	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	RRdetails.consultantName := DS[2]
	RRdetails.doctorName := DS[3]
	RRdetails.signature := CurrentDirectory . "Signatures\" . DS[3] . ".jpg"
	RRdetails.speciality := DS[4]
	RRdetails.location := DS[5]
	RRdetails.ward := DS[6]
	RRdetails.urgency := DS[7]
	RRdetails.clinicDate := DS[8]
	RRdetails.clinicalDetails := DS[9]
	RRdetails.holdBronchodilators := DS[10]
	RRdetails.spirometry := DS[11]
	RRdetails.gasTransfer := DS[12]
	RRdetails.staticLungVolumes := DS[13]
	RRdetails.FENO := DS[14]
	RRdetails.CBG := DS[15]
	RRdetails.sittingStanding := DS[16]
	RRdetails.bronchodilatorResponse := DS[17]
	RRdetails.osmohale := DS[18]
	RRdetails.fitToFly := DS[10]
	RRdetails.SS := DS[20]
	RRdetails.CPAP := DS[21]
	RRdetails.mouthPressures := DS[22]
	RRdetails.overnightPulseOx := DS[23]
	RRdetails.occupationalAsthmaStudy := DS[24]
	RRdetails.NIV := DS[25]
	RRdetails.NIVMode := DS[26]
	RRdetails.NIVIPAP := DS[27]
	RRdetails.NIVEPAP := DS[28]
	RRdetails.salbutamolCheckBox := DS[29]
	RRdetails.ipratropiumCheckBox := DS[30]
	
	
	if (DS[32] = "testEmail")
	{
		emailAddress := TestEmail
	}
	else
	{
		emailAddress := PFT_Email
	}


	if GetPatientDetailsTrakcare(MRN)
	{
		UpdateProgressBar(60, "absolute", PBMainString . "Creating request…")

		if RRCreatePDF()
		{
			UpdateProgressBar(20,, PBMainString . "Emailing lung function request…")
			
			if (CitrixSession == True)
			{
				MB("Current cannot send via a Citrix session. Please manually locate your request in the requests folder in the master directory and send via the NHSmail web app")
				;NHSMailOpen()
				;Email_NHS_Mail(emailAddress,, "Lung function request", "Please find attached a lung function request", PFT_latestRequestPath)
				; pass := True
			}
			else
			{
				if EmailOutlook(emailAddress,, "Lung function request", "Please find attached a lung function request", PFT_latestRequestPath)
				{
					pass := True
				}
			}
		}
	}

	CloseProgressBar()
	
	if pass
	{
		return "Pass"
	}
	else
	{
		return "Fail"
	}
}




PFT_GUI()
{
	global
	local GUIStart := 160
	local GUIWidth := 300
	local yAdditiveStart := 540
	local yAdditive := yAdditiveStart
	local checkBoxStart := 10
	local checkBoxWidth := 250
	local clinicIn6weeks += 42, Days
	local GUI_name := "PFT_GUI"
	local H := 0 ; Height of field
	local tickBoxSpacing := 20
	local yAdditive := 10

	
	Gui, %GUI_name%:Font, s12
	Gui, %GUI_name%:Color, %dialogueColour%
	
	H := 40
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Consultant / SAS:
	Gui, %GUI_name%:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_consultantName, %PFT_consultantList%
	yAdditive += H + S
		
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_doctorName, %doctorList%
	yAdditive += H + S
		
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Speciality:
	Gui, %GUI_name%:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_speciality, Respiratory||Oncology|Cardiology|Gastroenterology|
	yAdditive += H + S
		
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Location:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_location gPFT_location_click, Outpatient||Inpatient|Tetbury|Private patient
	
	H := 30
	Gui, %GUI_name%:Add, Text, x500 y%yAdditive% vPFT_wardText, Ward:
	Gui, %GUI_name%:Add, Edit, x600 y%yAdditive% W150 vPFT_ward
	yAdditive += H + S
	
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Urgency:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_urgency gPFT_urgency_click, Routine (4-6 weeks)||Urgent (< 2 weeks)|Prior to next outpatient appointment
	
		
	H := 30
	Gui, %GUI_name%:Add, Text, x500 y%yAdditive% vPFT_clinicDateText, Clinic date:
	Gui, %GUI_name%:Add, DateTime, x600 y%yAdditive% W150 vPFT_clinicDate choose%clinicIn6weeks%, dd/MM/yy
	yAdditive += H + S

	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Clinical information:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W500 H300 vPFT_clinicalDetails

	yAdditiveStart := yAdditive + 310
	yAdditive := yAdditiveStart
	
	; Insert tests
	For index, key in PFT_tests[]
	{
		variable := PFT_tests[key]

		if (PFT_tests[key] = "NIV")
		{
			Gui, %GUI_name%:Add, Checkbox, x%checkBoxStart% y%yAdditive% v%variable% +Right w%checkBoxWidth% gPFT_NIV_click, %key%:
		}
		else if (PFT_tests[key] = "bronchodilatorResponse")
		{
			Gui, %GUI_name%:Add, Checkbox, x%checkBoxStart% y%yAdditive% v%variable% +Right w%checkBoxWidth% gPFT_bronchodilatorResponse_click, %key%:
		}
		else
		{
			Gui, %GUI_name%:Add, Checkbox, x%checkBoxStart% y%yAdditive% v%variable% +Right w%checkBoxWidth%, %key%:
		}

		yAdditive += tickBoxSpacing
		
		if (index == 8)
		{
			checkBoxStart := 270
			checkBoxWidth := 300
			yAdditive := yAdditiveStart
		}
	}
		
	H := tickBoxSpacing
	Gui, %GUI_name%:Add, Checkbox, x10 y%yAdditive% vsalbutamolCheckBox +Right w250, Prescribe salbutamol:
	yAdditive += H ; no S here!
	
	H := tickBoxSpacing
	Gui, %GUI_name%:Add, Checkbox, x10 y%yAdditive% vipratropiumCheckBox +Right w250, Prescribe ipratropium bromide:
	yAdditive += H + S
	
	yAdditive = 590
	
	H := 30
	Gui, %GUI_name%:Add, Text, x600 y%yAdditive% vNIVShowHide1, NIV mode:
	Gui, %GUI_name%:Add, Edit, x700 y%yAdditive% W150 vNIVMode
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x600 y%yAdditive% vNIVShowHide2, IPAP:
	Gui, %GUI_name%:Add, Edit, x700 y%yAdditive% W150 vNIVIPAP
	yAdditive += H + S
		
	H := 30
	Gui, %GUI_name%:Add, Text, x600 y%yAdditive% vNIVShowHide3, EPAP:
	Gui, %GUI_name%:Add, Edit, x700 y%yAdditive% W150 vNIVEPAP
	yAdditive += H + 30
	
	Gui, %GUI_name%:Add, Button, x720 y%yAdditive% default gPFT_OK,  &OK
	Gui, %GUI_name%:Add, Button, x780 y%yAdditive% gPFT_close, &Cancel

	; Initially hide certain controls
	GuiControl, %GUI_name%:Hide, PFT_wardText
	GuiControl, %GUI_name%:Hide, PFT_ward
	GuiControl, %GUI_name%:Hide, PFT_clinicDateText
	GuiControl, %GUI_name%:Hide, PFT_clinicDate
	GuiControl, %GUI_name%:Hide, NIVShowHide1
	GuiControl, %GUI_name%:Hide, NIVShowHide2
	GuiControl, %GUI_name%:Hide, NIVShowHide3
	GuiControl, %GUI_name%:Hide, NIVMode
	GuiControl, %GUI_name%:Hide, NIVIPAP
	GuiControl, %GUI_name%:Hide, NIVEPAP
	GuiControl, %GUI_name%:Hide, salbutamolCheckBox
	GuiControl, %GUI_name%:Hide, ipratropiumCheckBox

	Gui, %GUI_name%:Show,, Lung function request details
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, Lung function request details
	return PFT_ButtonPressed
}




PFT_OK()
{
	global
	PFT_ButtonPressed := "OK"
	local dateToday := ""
	FormatTime, dateToday,, dd/MM/yy
	local errorString := ""
	local pass := False
	local atLeastOneInvestigationChecked := False


	Gui PFT_GUI:Submit, Nohide


	if (PFT_consultantName == "")
		errorString := errorString . "- No consultant name was chosen.`n"


	if (PFT_doctorName == "")
		errorString := errorString . "- No requesting doctor was chosen.`n"


	if (PFT_location = "Inpatient")
		if (PFT_ward = "")
			errorString := errorString . "- You have not specified a ward for this inpatient.`n"


	if (PFT_clinicalDetails == "")
		errorString := errorString . "- No clinical details have been entered.`n"


	; tests
	For index, key in PFT_tests[]
	{
		variable := PFT_tests[key]

		if (%variable% == 1)
		{
			atLeastOneInvestigationChecked := True
			break
		}
	}

	if !atLeastOneInvestigationChecked
		errorString := errorString . "- No investigation was selected.`n"	


	if (NIV == 1)
	{
		if (NIVMode == "")
			errorString := errorString . "- NIV mode is missing.`n"

		if NIVIPAP is number
			if (NIVIPAP > 0 AND NIVIPAP < 50)
				pass := True
		
		if (pass == False)
			errorString := errorString . "- The IPAP needs to be between 0 and 50 (exclusive).`n"

		pass := False

		if NIVEPAP is number
			if (NIVEPAP > 0 AND NIVEPAP < NIVIPAP)
				pass := True
		if (pass == False)
			errorString := errorString . "- The EPAP needs to be between 0 and 50 (exclusive) and less than the IPAP.`n"
	}

	if (errorString != "")
	{
		MB("Please correct the below errors and submit again:`n" . errorString)
		return
	}

		
	Gui, PFT_GUI:Destroy
	RRDetails.consultantName := PFT_consultantName
	RRDetails.addressForReport := PFT_consultantName
	RRDetails.doctorName := PFT_doctorName
	RRDetails.signature := CurrentDirectory . "Signatures\" . PFT_doctorName . ".jpg"
	RRDetails.speciality := PFT_speciality
	RRDetails.clinicalDetails := PFT_clinicalDetails

	
	if (PFT_location = "Outpatient")
	{
		RRDetails.CB_outpatient := 1
		RRDetails.CB_inpatient := 0
		RRDetails.CB_privatePatient := 0
	}
	else if (PFT_location = "Inpatient")
	{
		RRDetails.CB_outpatient := 0
		RRDetails.CB_inpatient := 1
		RRDetails.CB_privatePatient := 0
		RRDetails.ward := PFT_ward
	}
	else if (PFT_location = "Private patient")
	{
		RRDetails.CB_outpatient := 0
		RRDetails.CB_inpatient := 0
		RRDetails.CB_privatePatient := 1
	}
	else
	{
		RRDetails.outpatient := PFT_location
		RRDetails.CB_inpatient := 0
		RRDetails.CB_privatePatient := 0
	}


	if (PFT_urgency = "Routine (4-6 weeks)")
	{
		RRDetails.CB_routine := 1
		RRDetails.CB_urgent := 0
		RRDetails.CB_priorToNextOPA := 0
	}
	else if (PFT_urgency = "Urgent (< 2 weeks)")
	{
		RRDetails.CB_routine := 0
		RRDetails.CB_urgent := 1
		RRDetails.CB_priorToNextOPA := 0
	}
	else if (PFT_urgency = "Prior to next outpatient appointment")
	{
		RRDetails.CB_routine := 0
		RRDetails.CB_urgent := 0
		RRDetails.CB_priorToNextOPA := 1
		FormatTime, PFT_clinicDate, %PFT_clinicDate%, dd/MM/yy
		RRDetails.clinicDate := PFT_clinicDate
	}


	; I am not going to change these at present
	RRDetails.CB_patientTransport := 0
	RRDetails.CB_interpreter := 0
	RRDetails.CB_learningDifficulties := 0


	; I am not going to change these at present
	RRDetails.CB_DnV := 0
	RRDetails.CB_haemoptysis := 0
	RRDetails.CB_pneumothorax := 0
	RRDetails.CB_aneurysm := 0
	RRDetails.CB_tuberculosis := 0
	RRDetails.CB_MI := 0
	RRDetails.CB_stroke := 0
	RRDetails.CB_PE := 0
	RRDetails.CB_surgery := 0


	RRDetails.CB_holdBronchodilators := holdBronchodilators

	RRDetails.CB_spirometry := spirometry
	RRDetails.CB_gasTransfer := gasTransfer
	RRDetails.CB_staticLungVolumes := staticLungVolumes
	RRDetails.CB_FENO := FENO
	RRDetails.CB_CBG := CBG
	RRDetails.CB_sittingStanding := sittingStanding
	RRDetails.CB_bronchodilatorResponse := bronchodilatorResponse


	RRDetails.CB_osmohale := osmohale
	RRDetails.CB_fitToFly := fitToFly
	RRDetails.CB_SS := SS
	RRDetails.CB_CPAP := CPAP
	RRDetails.CB_mouthPressures := mouthPressures
	RRDetails.CB_overnightPulseOx := overnightPulseOx
	RRDetails.CB_occupationalAsthmaStudy := occupationalAsthmaStudy
	RRDetails.CB_NIV := NIV


	if (NIV == 1)
	{
		RRDetails.NIVMode := NIVMode
		; need "" before number to not cause an error
		RRDetails.NIVIPAP := "" NIVIPAP
		RRDetails.NIVEPAP := "" NIVEPAP
	}


	if (salbutamolCheckBox == 1)
	{
		RRDetails.CB_salbutamolCheckBox := 1
		RRDetails.signatureSalbutamol := RRDetails.signature
		RRDetails.dateSalbutamol := dateToday
	}


	if (ipratropiumCheckBox == 1)
	{
		RRDetails.CB_ipratropiumCheckBox := 1
		RRDetails.signatureIpratropium := RRDetails.signature
		RRDetails.dateIpratropium := dateToday
	}


	if (osmohale == 1)
	{
		RRDetails.CB_mannitolCheckBox := 1
		RRDetails.signatureMannitol := RRDetails.signature
		RRDetails.dateMannitol := dateToday
	}


	if (fitToFly == 1)
	{
		RRDetails.CB_oxygenCheckBox := 1
		RRDetails.signatureOxygen := RRDetails.signature
		RRDetails.dateOxygen := dateToday
	}

	return
}




PFT_GUIGuiClose()
{
	PFT_close()
}
PFT_close()
{
	global
	PFT_buttonPressed := "close"
	Gui, PFT_GUI:Destroy
	return
}




PFT_NIV_click(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global


	Gui PFT_GUI:Submit, NoHide

	if (NIV == 1)
	{
		GuiControl, PFT_GUI:Show, NIVShowHide1
		GuiControl, PFT_GUI:Show, NIVShowHide2
		GuiControl, PFT_GUI:Show, NIVShowHide3
		GuiControl, PFT_GUI:Show, NIVMode
		GuiControl, PFT_GUI:Show, NIVIPAP
		GuiControl, PFT_GUI:Show, NIVEPAP
	}
	else
	{
		GuiControl, PFT_GUI:Hide, NIVShowHide1
		GuiControl, PFT_GUI:Hide, NIVShowHide2
		GuiControl, PFT_GUI:Hide, NIVShowHide3
		GuiControl, PFT_GUI:Hide, NIVMode
		GuiControl, PFT_GUI:Hide, NIVIPAP
		GuiControl, PFT_GUI:Hide, NIVEPAP
	}

	return True
}




PFT_bronchodilatorResponse_click(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global
	
	
	Gui PFT_GUI:Submit, NoHide

	if (bronchodilatorResponse == 1)
	{
		GuiControl, PFT_GUI: , salbutamolCheckBox, 1
		GuiControl, PFT_GUI:Show, salbutamolCheckBox
		GuiControl, PFT_GUI:Show, ipratropiumCheckBox
	}
	else
	{
		GuiControl, PFT_GUI: , salbutamolCheckBox, 0
		GuiControl, PFT_GUI: , ipratropiumCheckBox, 0
		GuiControl, PFT_GUI:Hide, salbutamolCheckBox
		GuiControl, PFT_GUI:Hide, ipratropiumCheckBox
	}

	return True
}




PFT_location_click(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global
	
	
	Gui PFT_GUI:Submit, NoHide

	if (PFT_location = "Inpatient")
	{
		GuiControl, PFT_GUI:Show, PFT_wardText
		GuiControl, PFT_GUI:Show, PFT_ward
	}
	else
	{
		GuiControl, PFT_GUI:Hide, PFT_wardText
		GuiControl, PFT_GUI:Hide, PFT_ward
	}

	return True
}




PFT_urgency_click(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global
	
	
	Gui PFT_GUI:Submit, NoHide

	if (PFT_urgency = "Prior to next outpatient appointment")
	{
		GuiControl, PFT_GUI:Show, PFT_clinicDateText
		GuiControl, PFT_GUI:Show, PFT_clinicDate
	}
	else
	{
		GuiControl, PFT_GUI:Hide, PFT_clinicDateText
		GuiControl, PFT_GUI:Hide, PFT_clinicDate
	}

	return True
}




; Bronchoscopy request via IPC
send_bronchoscopy(rawData)
{
	global
	RRType := "bronchoscopy"
	local Pass := False
	local DS := ""
	local emailAddress := ""
	local cc_addresses := ""
	
	
	CloseProgressBar()
	PBMainString := "Bronchoscopy request:`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Starting…")

	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	RRDetails.procedure  := DS[2]
	RRDetails.location  := DS[3]
	RRDetails.primaryDiagnosis :=  DS[4]
	RRDetails.significantComorbidities :=  DS[5]
	RRDetails.preferredBronchoscopyDate := DS[6]
	RRDetails.consentCompleted := DS[7]
	RRDetails.imaging := DS[8]
	RRDetails.anticoagulantAntiplatelets :=DS[9]
	RRDetails.aim := DS[10]
	RRDetails.consultantName :=DS[11]
	RRDetails.doctorName := DS[12]
	RRDetails.signature := CurrentDirectory . "Signatures\" . .  DS[12] . ".jpg"

	if (DS[13] = "testEmail")
	{
		emailAddress := TestEmail
		bronch_emailCoordinators := 0
	}
	else
	{
		emailAddress := bronch_Email
		bronch_emailCoordinators := 1
	}


	if GetPatientDetailsTrakcare(MRN)
	{
		UpdateProgressBar(60, "absolute", PBMainString . "Creating request…")

		if RRCreatePDF()
		{
			UpdateProgressBar(20,, PBMainString . "Emailing bronchoscopy request…")
			
			if (bronch_emailCoordinators == 1)
				cc_addresses := bronch_Email_cc
			else
				cc_addresses := ""
			
			if (CitrixSession == True)
			{
				MB("Current cannot send via a Citrix session. Please manually locate your request in the requests folder in the master directory and send via the NHSmail web app")
				;NHSMailOpen()
				;Email_NHS_Mail(emailAddress, cc_addresses, "Bronchoscopy request", "Please find attached a bronchoscopy request", bronch_latestRequestPath)
				; pass := True
			}
			else
			{
				if EmailOutlook(emailAddress, cc_addresses, "Bronchoscopy request", "Please find attached a bronchoscopy request", bronch_latestRequestPath)
				{
					pass := True
				}
			}
		}
	}

	CloseProgressBar()
	
	if pass
		return "Pass"
	else
		return "Fail"
}




bronchoscopy_GUI()
{
	global
	local GUIStart := 250
	local GUIWidth := 500
	local H := 0 ; Height of field
	local yAdditiveStart := 10
	local yAdditive := yAdditiveStart
	local checkBoxStart := 10
	local checkBoxWidth := 250
	local clinicIn6weeks += 42, Days
	FormatTime, clinicIn6weeks, %clinicIn6weeks%, dd/MM/yy
	local GUI_name := "bronchoscopy_GUI"

		
	Gui, %GUI_name%:Font, s12
	Gui, %GUI_name%:Color, %dialogueColour%
	
	H := 40
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Procedure:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vbronch_procedure, Bronchoscopy|EBUS|Thoracoscopy
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Patient's current location:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vbronch_location, Outpatient||Inpatient
	yAdditive += H + S
	
	H := 70
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Primary diagnosis:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% H%H% W%GUIWidth% vbronch_primaryDiagnosis
	yAdditive += H + S
	
	H := 70
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Significant co-morbidities:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% H%H% W%GUIWidth% vbronch_significantComorbidities
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive% , Preferred bronchoscopy date:
	Gui, %GUI_name%:Add, DateTime, x%GUIStart% y%yAdditive% W%GUIWidth% vbronch_preferredBronchoscopyDate
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Consent completed:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vbronch_consentCompleted, No||Yes|
	yAdditive += H + S
	
	H := 70
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Imaging (modality and result):
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% H%H% W%GUIWidth% vbronch_imaging
	yAdditive += H + S
	
	H := 70
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Anticoagulant / anti-platelets:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% H%H% W%GUIWidth% vbronch_anticoagulantAntiplatelets
	yAdditive += H + S
	
	H := 70
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Aim of bronchoscopy / sampling:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% H%H% W%GUIWidth% vbronch_aim
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Consultant:
	Gui, %GUI_name%:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vbronch_consultantName, %consultantList%
	yAdditive += H + S
		
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vbronch_doctorName, %doctorList%
	yAdditive += H + S
		
	H := 30
	Gui, %GUI_name%:Add, CheckBox, x5 y%yAdditive%  vbronch_emailCoordinators +Right Checked, Email coordinators (April, Lisa + Natasha): 
	yAdditive += H + S
	
	Gui, %GUI_name%:Add, Button, x720 y%yAdditive% default gbronchoscopy_OK,  &OK
	Gui, %GUI_name%:Add, Button, x780 y%yAdditive% gbronchoscopy_close, &Cancel
	

	Gui, %GUI_name%:Show,, Bronchoscopy request details
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, Bronchoscopy request details
	return bronch_buttonPressed
}




bronchoscopy_OK()
{
	global
	bronch_buttonPressed := "OK"
	local errorString := ""
	local pass := False


	Gui bronchoscopy_GUI:Submit, Nohide
	local fieldsArr := [bronch_procedure, bronch_location, bronch_primaryDiagnosis, bronch_significantComorbidities, bronch_preferredBronchoscopyDate, bronch_consentCompleted, bronch_imaging, bronch_anticoagulantAntiplatelets,  bronch_aim, bronch_consultantName]
	

	if (bronch_consultantName == "")
		errorString := errorString . "- No consultant name was chosen.`n"

	if (bronch_doctorName == "")
		errorString := errorString . "- No requesting doctor was chosen.`n"
		
	for index, key in fieldsArr
	{
		;msgbox, % index . " - " . key
		if (key = "")
		{
			errorString := errorString . "- Empty field(s).`n"
			break
		}
	}

		
	if (errorString != "")
	{
		MB("Please correct the below errors and submit again:`n" . errorString)
		return
	}


	Gui, bronchoscopy_GUI:Destroy

	RRDetails.procedure  := bronch_procedure
	RRDetails.location  := bronch_location
	RRDetails.primaryDiagnosis := bronch_primaryDiagnosis
	RRDetails.significantComorbidities := bronch_significantComorbidities
	FormatTime, bronch_preferredBronchoscopyDate, %bronch_preferredBronchoscopyDate%, dd/MM/yy
	RRDetails.preferredBronchoscopyDate := bronch_preferredBronchoscopyDate
	RRDetails.consentCompleted := bronch_consentCompleted
	RRDetails.imaging := bronch_imaging
	RRDetails.anticoagulantAntiplatelets := bronch_anticoagulantAntiplatelets
	RRDetails.aim := bronch_aim
	RRDetails.consultantName := bronch_consultantName
	RRDetails.doctorName := bronch_doctorName
	RRDetails.signature := CurrentDirectory . "Signatures\" .  bronch_doctorName . ".jpg"
	return
}




bronchoscopy_GUIGuiClose()
{
	bronchoscopy_Close()
}
bronchoscopy_Close()
{
	global
	bronch_buttonPressed := "close"
	Gui, bronchoscopy_GUI:Destroy
	return
}




; HLSG request via IPC
send_HLSG(rawData)
{
	global
	RRType := "HLSG"
	local Pass := False
	local DS := ""
	local emailAddress := ""


	CloseProgressBar()
	PBMainString := "Healthy Lifestyle Gloucestershire request:`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Starting…")
	
	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	RRDetails.Comorbidities := DS[2]
	RRDetails.ethnicity := DS[3]
	RRDetails.Comorbidities := DS[4]
	RRDetails.smoking  := DS[5]
	RRDetails.alcohol  := DS[6]
	RRDetails.weight  := DS[7]
	RRDetails.physicalActivity  := DS[8]
	RRDetails.doctorName := DS[9]
	RRDetails.signature := CurrentDirectory .  DS[9] . ".jpg"

	if (DS[10] = "testEmail")
	{
		emailAddress := TestEmail
	}
	else
	{
		emailAddress := HLSG_Email
	}


	if GetPatientDetailsTrakcare(MRN)
	{
		UpdateProgressBar(60, "absolute", PBMainString . "Creating request…")

		if RRCreatePDF()
		{
			UpdateProgressBar(20,, PBMainString . "Emailing Healthy Lifestyle Gloucestershire request…")
			
			if (CitrixSession == True)
			{
				MB("Current cannot send via a Citrix session. Please manually locate your request in the requests folder in the master directory and send via the NHSmail web app")
				;NHSMailOpen()
				;Email_NHS_Mail(emailAddress,, "Healthy Lifestyle Gloucestershire request", "Please find attached a Healthy Lifestyle Gloucestershire request", HLSG_latestRequestPath)
				; pass := True
			}
			else
			{
				if EmailOutlook(emailAddress,, "Healthy Lifestyle Gloucestershire request", "Please find attached a Healthy Lifestyle Gloucestershire request", HLSG_latestRequestPath)
				{
					pass := True
				}
			}
		}
	}

	CloseProgressBar()
	
	if pass
		return "Pass"
	else
		return "Fail"
}




HLSG_GUI()
{
	global
	local GUIStart := 200
	local GUIWidth := 300
	local H := 0 ; Height of field
	local yAdditiveStart := 10
	local yAdditive := yAdditiveStart
	local checkBoxStart := 10
	local checkBoxWidth := 180
	local Title := "Healthy Lifestyles Gloucestershire referral details"
	local GUI_name := "HLSG_GUI"

		
	Gui, %GUI_name%:Font, s12
	Gui, %GUI_name%:Color, %dialogueColour%
	
	H := 40
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vHLSG_doctorName, %doctorList%
	yAdditive += H + S
		
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Ethnicity:
	Gui, %GUI_name%:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vHLSG_ethnicity, White||Black|Asian|; I appreciate this is a short ethnicity list. Will have to check best list of groups to use
	yAdditive += H + S
		
	H := 70
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Co-morbidities:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% H%H% W%GUIWidth% vHLSG_Comorbidities
	yAdditive += H + S

		
	; insert referral types
	For index, key in HLSG_referralTypes[]
	{
		variable := HLSG_referralTypes[key]
		Gui, %GUI_name%:Add, Checkbox, x%checkBoxStart% y%yAdditive% v%variable% +Right w%checkBoxWidth%, %key%:
		yAdditive += 20
	}


	Gui, %GUI_name%:Add, Button, x400 y%yAdditive% default gHLSG_OK,  &OK
	Gui, %GUI_name%:Add, Button, x450 y%yAdditive% gHLSG_close, &Cancel

	Gui, %GUI_name%:Show,, %Title%
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, %Title%
	return HLSG_buttonPressed
}




HLSG_OK()
{
	global
	HLSG_buttonPressed := "OK"
	local errorString := ""
	local pass := False
	local atLeastOneReferralTypeTicked := False
	
	
	Gui HLSG_GUI:Submit, Nohide
	

	if (HLSG_doctorName == "")
		errorString := errorString . "- No requesting doctor was chosen.`n"
		
		
	; referral types
	For index, key in HLSG_referralTypes[]
	{
		variable := HLSG_referralTypes[key]

		if (%variable% == 1)
		{
			atLeastOneReferralTypeTicked := True
			break
		}
	}


	if !atLeastOneReferralTypeTicked
		errorString := errorString . "- No investigation was selected.`n"	


		
	if (errorString != "")
	{
		MB("Please correct the below errors and submit again:`n" . errorString)
		return
	}


	Gui, HLSG_GUI:Destroy
	RRDetails.comorbidities := HLSG_Comorbidities
	RRDetails.ethnicity := HLSG_ethnicity
	RRDetails.doctorName := HLSG_doctorName
	RRDetails.signature := CurrentDirectory . "Signatures\" .  HLSG_doctorName . ".jpg"
	RRDetails.referrerDetails := getClinicianDetails(clinicianActualName, RRDetails.doctorName, clinicianTitle) . " " . RRDetails.doctorName . ", " . DeptAddress
	
	if (PatientDetails.gender = "Male")
	{
		RRDetails.CB_male := 1
		RRDetails.CB_female := 0
		RRDetails.CB_transgender := 0
	}
	else if (PatientDetails.gender = "Female")
	{
		RRDetails.CB_male := 0
		RRDetails.CB_female := 1
		RRDetails.CB_transgender := 0
	}
	else
	{
		RRDetails.CB_male := 0
		RRDetails.CB_female := 0
		RRDetails.CB_transgender := 0
	}
	
	RRDetails.patientAddress := PatientDetails.address1 . ", " . PatientDetails.address2 . ", " . PatientDetails.address3 . ", " . PatientDetails.address4 . ", " . PatientDetails.postCode
	
	RRDetails.telephoneNr := PatientDetails.MobNo . ", " . PatientDetails.TeleNo
	
	
	if (HLSG_smoking == 1)
		RRDetails.CB_smoking := 1
	else
		RRDetails.CB_smoking := 0
	
	if (HLSG_alcohol == 1)
		RRDetails.CB_alcohol := 1
	else
		RRDetails.CB_alcohol := 0

	if (HLSG_weight == 1)
		RRDetails.CB_weight := 1
	else
		RRDetails.CB_weight := 0

	if (HLSG_physicalActivity == 1)
		RRDetails.CB_physicalActivity := 1
	else
		RRDetails.CB_physicalActivity := 0	
	
	return
}




HLSG_GUIGuiClose()
{
	HLSG_Close()
}
HLSG_Close()
{
	global
	HLSG_buttonPressed := "close"
	Gui, HLSG_GUI:Destroy
	return
}




sleepstation_GUI()
{
	global
	local Title := "Sleepstation referral details"
	local GUIStart := 200
	local GUIWidth := 300
	local H := 0 ; Height of field
	local yAdditiveStart := 10
	local yAdditive := yAdditiveStart
	local checkBoxStart := 10
	local checkBoxWidth := 180
	local GUI_name := "sleepstation_GUI" 
		
	Gui, %GUI_name%:Font, s12
	Gui, %GUI_name%:Color, %dialogueColour%
	
	H := 40
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vSleepstationDoctorName, %doctorList%
	yAdditive += H + S
		
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Doctor email:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vSleepstationDoctorEmail, % getClinicianDetails(clinicianUsername, A_UserName, clinicianEmail)
	yAdditive += H + S
		
	H := 70
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Clinical details:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% H%H% W%GUIWidth% vSleepstationClinicalDetails, % "Insomnia "
	yAdditive += H + S

	Gui, %GUI_name%:Add, Button, x400 y%yAdditive% default gsleepstation_OK,  &OK
	Gui, %GUI_name%:Add, Button, x450 y%yAdditive% gsleepstation_Close, &Cancel

	Gui, %GUI_name%:Show,, %Title%
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, %Title%
	return Sleepstation_buttonPressed
}




sleepstation_OK()
{
	global
	
	local errorString := ""
		
	Gui sleepstation_GUI:Submit, Nohide
	Sleepstation_buttonPressed := "OK"


	if (SleepstationDoctorName == "")
		errorString := errorString . "- No requesting doctor was chosen.`n"
		
	if (SleepstationDoctorEmail == "" or !instr(SleepstationDoctorEmail, "@"))
		errorString := errorString . "- No requesting doctor email provided.`n"
		
	if (SleepstationClinicalDetails = "" or SleepstationClinicalDetails == "Insomnia ")
		errorString := errorString . "- No clinical details were provided.`n"
		
	if (errorString != "")
	{
		MB("Please correct the below errors and submit again:`n" . errorString)
		return
	}

	Gui, sleepstation_GUI:Destroy
	RRDetails.doctorName := SleepstationDoctorName
	RRDetails.signature := CurrentDirectory . "Signatures\" .  SleepstationDoctorName . ".jpg"
	RRDetails.doctorEmail := SleepstationDoctorEmail
	RRDetails.clinicalDetails := SleepstationClinicalDetails
	RRDetails.patientAddress := PatientDetails.address1 . ", " . PatientDetails.address2 . ", " . PatientDetails.address3 . ", " . PatientDetails.address4 . ", " . PatientDetails.postCode
	RRDetails.telephoneNr := PatientDetails.MobNo . ", " . PatientDetails.TeleNo
	
	return
}




sleepstation_GUIGuiClose()
{
	sleepstation_close()
}
sleepstation_close()
{
	global
	Sleepstation_buttonPressed := "close"
	Gui, sleepstation_GUI:Destroy
	return
}



