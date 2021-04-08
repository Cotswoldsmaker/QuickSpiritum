; ***********************************************************************
; Requests via Quick Spiritum (mind out for GOTO statements with returns)
; ***********************************************************************

; PET-CT request via IPC
send_PET_CT(rawData)
{
	global PBMainString, PET_CT_details, PET_CT_Email, PET_CT_Email_cc, PET_CT_latestRequestPath
	global CitrixSession, TestEmail, CurrentDirectory, CurrentDirectory

	Pass := False

	CloseProgressBar()
	PBMainString := "PET-CT request:`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Starting…")

	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	PET_CT_details.consultantName := DS[2]
	PET_CT_details.doctorName := DS[3]
	PET_CT_details.signature := CurrentDirectory . "Signatures\" . DS[3] . ".jpg"
	PET_CT_details.MDTDate := DS[4]
	PET_CT_details.previousImagingType := DS[5]
	PET_CT_details.previousImagingDate := DS[6]
	PET_CT_details.clinicalInformation := DS[7]
	PET_CT_details.diabetic := DS[8]

	
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

		if CreatePET_CT_request(MRN)
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




GetPET_CT_ExtraInfo(MRN)
{
	global PET_CT_details
	
	PET_CT_details := {}

	if !(PET_CT_GUI(MRN) = "OK")
	{
		return False
	}
	
	if !FileExist(PET_CT_details.signature)
	{
		MB("No signature found for selected doctor. Ending PET-CT request")
		return False
	}

	return True
}




PET_CT_GUI(MRN)
{
	global PET_CT_details, PET_CT_consultantName, PET_CT_doctorName, PET_CT_previousImagingType 
	global PET_CT_previousImagingDate, PET_CT_clinicalInformation, PET_CT_diabetic, PET_CTButtonPressed
	global PET_CT_emailLungCancerCoordinators, patientDetails, consultantListConstant, doctorListConstant
	global dialogueColour
	
	GUIStart := 237
	GUIWidth := 400
	CurrentUser := ConvertUsername(A_UserName)
	CurrentUserPostionInListConsultant := InStr(consultantListConstant, CurrentUser)  ; 0  if not found
	CurrentUserPostionInListDoctor := InStr(doctorListConstant, CurrentUser)  ; 0  if not found
	yAdditive := 10
	
	
	if (CurrentUserPostionInListConsultant == 0)
		consultantList := consultantListConstant
	else
		consultantList := SubStr(consultantListConstant, 1, CurrentUserPostionInListConsultant + StrLen(CurrentUser)) . "|" . SubStr(consultantListConstant, CurrentUserPostionInListConsultant + StrLen(CurrentUser) + 1)
		
	
	if (CurrentUserPostionInListDoctor == 0)
		doctorList := doctorListConstant
	else
		doctorList := SubStr(doctorListConstant, 1, CurrentUserPostionInListDoctor + StrLen(CurrentUser)) . "|" . SubStr(doctorListConstant, CurrentUserPostionInListDoctor + StrLen(CurrentUser) + 1)
		

	Gui, PET_CT:Font, s12
	Gui, PET_CT:Color, %dialogueColour%
	Gui, PET_CT:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive := yAdditive + 35
	
	Gui, PET_CT:Add, Text, x10 y%yAdditive%, Consultant:
	Gui, PET_CT:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_consultantName, %consultantList%
	yAdditive := yAdditive + 35
	
	Gui, PET_CT:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, PET_CT:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_doctorName, %doctorList%
	yAdditive := yAdditive + 35
	
	Gui, PET_CT:Add, Text, x10 y%yAdditive%, Previous imaging type:
	Gui, PET_CT:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_previousImagingType, CT chest||CT body|MRI
	yAdditive := yAdditive + 35
	
	Gui, PET_CT:Add, Text, x10 y%yAdditive%, Previous imaging date:
	Gui, PET_CT:Add, DateTime, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_previousImagingDate
	yAdditive := yAdditive + 35
	
	Gui, PET_CT:Add, Text, x10 y%yAdditive%, Clinical information:
	Gui, PET_CT:Add, Edit, x%GUIStart% y%yAdditive% W600 H300 vPET_CT_clinicalInformation
	yAdditive := yAdditive + 310
	
	Gui, PET_CT:Add, Text, x10 y%yAdditive%, Diabetic:
	Gui, PET_CT:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPET_CT_diabetic, Unknown||No|Yes
	yAdditive := yAdditive + 45
	
	Gui, PET_CT:Add, CheckBox, x5 y%yAdditive%  vPET_CT_emailLungCancerCoordinators +Right Checked, Email lung cancer coordinators: 

	Gui, PET_CT:Add, Button, x800 y%yAdditive% default,  &OK
	Gui, PET_CT:Add, Button, x840 y%yAdditive%,  &Cancel
	Gui, PET_CT:Show,, PET-CT request details
	Gui, PET_CT:+AlwaysOnTop
	WinWaitClose, PET-CT request details
	return PET_CTButtonPressed
}




PET_CTButtonOK:

errorString := ""
pass := False
PET_CTButtonPressed := "OK"
Gui PET_CT:Submit, NoHide

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

Gui, PET_CT:Destroy
PET_CT_details.consultantName := PET_CT_consultantName
PET_CT_details.doctorName := PET_CT_doctorName
PET_CT_details.signature := CurrentDirectory . "Signatures\" . PET_CT_doctorName . ".jpg"
Yesterday += -1, Days
FormatTime, Yesterday, %Yesterday%, dd/MM/yy
PET_CT_details.MDTDate := Yesterday
PET_CT_details.previousImagingType := PET_CT_previousImagingType
PET_CT_details.previousImagingDate := PET_CT_previousImagingDate
PET_CT_details.clinicalInformation := PET_CT_clinicalInformation
PET_CT_details.diabetic := PET_CT_diabetic
return




PET_CTClose:
PET_CTButtonCancel:
PET_CTGuiClose:
PET_CTButtonPressed := "close"
Gui, PET_CT:Destroy
return




CreatePET_CT_request(MRN)
{
	global Settings, PET_CT_template_path, oWord, PatientDetails, PET_CT_details
	global PET_CT_latestRequestPath, signatureHeight
	global PBMainString
	
	Loop, 5		; Allow 5 errors before failing function
	{
		SavePath := Settings.RequestsFolder . "PET CT request MRN" . MRN . "_"
		FormatTime, Today,, dd/MM/yy
		signatureWidth := signatureHeight * GraphicsDimensions(PET_CT_details.signature)
		imagingDateFormated := PET_CT_details.previousImagingDate
		FormatTime, imagingDateFormated, %imagingDateFormated%, dd/MM/yy
		PET_CT_details.previousImagingDate := imagingDateFormated
		
		try
		{
			Loop
			{
				SavePathTemp := SavePath . A_Index . ".pdf"

				if(FileExist(SavePathTemp) == "")
				{
					SavePath := SavePathTemp
					PET_CT_latestRequestPath := SavePath
					break
				}
			}

			oWord := ComObjCreate("Word.Application")
			oWord.Visible := False
			oWord.Documents.Open(PET_CT_template_path)


			for key, value in PatientDetails
			{
				if (key != "MRN" and key != "email")
					WriteAtBookmark(key, value)
			}


			for key, value in PET_CT_details
			{
				if (key = "signature")
				{
					WriteAtBookmark("signature", PET_CT_details.doctorName)
					oWord.ActiveDocument.Bookmarks("signature").Select
					picObj := oWord.ActiveDocument.InLineShapes.AddPicture(value)
					picObj.Height := signatureHeight 
					picObj.Width := signatureWidth
				}
				else if (key = "doctorName")
				{
					; Nothing
				}
				else
				{
					WriteAtBookmark(key, value)
				}
			}

			WriteAtBookmark("referralDate", Today)
			
			UpdateProgressBar(80, "absolute", PBMainString . "Saving referral to requests folder")
			
			oWord.ActiveDocument.SaveAs(SavePath, 17) ; 17 is PDF format
			oWord.ActiveDocument.close(False)
			
			return True
		}
		catch, err
		{
			try
				oWord.ActiveDocument.close(False)
				
			if (A_index >= 5)
			{
				errorhandler(err, "PET-CT request", "Microsoft Word")
				return False
			}			
			else
			{
				LogUpdate("Error during PET-CT request [counter: " . A_index . "]")
			}
		}
	}
	
	return True
}




; PFT request via IPC
send_PFT(rawData)
{
	global PFT_details, CitrixSession, TestEmail, PBMainString
	global PFT_Email, PFT_latestRequestPath, CurrentDirectory

	Pass := False

	CloseProgressBar()
	PBMainString := "Lung function request:`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Starting…")
	
	
	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	PFT_details.consultantName := DS[2]
	PFT_details.doctorName := DS[3]
	PFT_details.signature := CurrentDirectory . "Signatures\" . DS[3] . ".jpg"
	PFT_details.speciality := DS[4]
	PFT_details.location := DS[5]
	PFT_details.ward := DS[6]
	PFT_details.urgency := DS[7]
	PFT_details.clinicDate := DS[8]
	PFT_details.clinicalDetails := DS[9]
	PFT_details.holdBronchodilators := DS[10]
	PFT_details.spirometry := DS[11]
	PFT_details.gasTransfer := DS[12]
	PFT_details.staticLungVolumes := DS[13]
	PFT_details.FENO := DS[14]
	PFT_details.CBG := DS[15]
	PFT_details.sittingStanding := DS[16]
	PFT_details.bronchodilatorResponse := DS[17]
	PFT_details.osmohale := DS[18]
	PFT_details.fitToFly := DS[10]
	PFT_details.SS := DS[20]
	PFT_details.CPAP := DS[21]
	PFT_details.mouthPressures := DS[22]
	PFT_details.overnightPulseOx := DS[23]
	PFT_details.occupationalAsthmaStudy := DS[24]
	PFT_details.NIV := DS[25]
	PFT_details.NIVMode := DS[26]
	PFT_details.NIVIPAP := DS[27]
	PFT_details.NIVEPAP := DS[28]
	PFT_details.salbutamolCheckBox := DS[29]
	PFT_details.ipratropiumCheckBox := DS[30]
	
	
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

		if Create_PFT_request(MRN)
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




get_PFT_extraInfo(MRN)
{
	global PFT_details
	
	PFT_details := {}

	if !(PFT_GUI(MRN) = "OK")
	{
		return False
	}
	
	if !FileExist(PFT_details.signature)
	{
		MB("No signature found for requesting doctor. Ending lung function request")
		return False
	}

	return True
}




PFT_GUI(MRN)
{
	global

	local GUIStart := 160
	local GUIWidth := 300
	local CurrentUser := ConvertUsername(A_UserName)
	local consultantList := Username1 . "|" . consultantListConstant
	local CurrentUserPostionInListConsultant := InStr(consultantList, CurrentUser)  ; 0  if not found
	local CurrentUserPostionInListDoctor := InStr(doctorListConstant, CurrentUser)  ; 0  if not found

	local yAdditiveStart := 540
	local yAdditive := yAdditiveStart
	local checkBoxStart := 10
	local checkBoxWidth := 250
	local clinicIn6weeks += 42, Days
	yAdditive := 10

	
	if (CurrentUserPostionInListConsultant > 0)
		consultantList := SubStr(consultantList, 1, CurrentUserPostionInListConsultant + StrLen(CurrentUser)) . "|" . SubStr(consultantList, CurrentUserPostionInListConsultant + StrLen(CurrentUser) + 1)
		
	
	if (CurrentUserPostionInListDoctor == 0)
		doctorList := doctorListConstant
	else
		doctorList := SubStr(doctorListConstant, 1, CurrentUserPostionInListDoctor + StrLen(CurrentUser)) . "|" . SubStr(doctorListConstant, CurrentUserPostionInListDoctor + StrLen(CurrentUser) + 1)
		
	
	Gui, PFTReq:Font, s12
	Gui, PFTReq:Color, %dialogueColour%
	Gui, PFTReq:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive := yAdditive + 35
	
	Gui, PFTReq:Add, Text, x10 y%yAdditive%, Consultant / SAS:
	Gui, PFTReq:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_consultantName, %consultantList%
	yAdditive := yAdditive + 35
	
	Gui, PFTReq:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, PFTReq:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_doctorName, %doctorList%
	yAdditive := yAdditive + 35
	
	Gui, PFTReq:Add, Text, x10 y%yAdditive%, Speciality:
	Gui, PFTReq:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_speciality, Respiratory||Oncology|Cardiology|Gastroenterology|
	yAdditive := yAdditive + 35
	
	Gui, PFTReq:Add, Text, x10 y%yAdditive%, Location:
	Gui, PFTReq:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_location gPFT_location_click, Outpatient||Inpatient|Tetbury|Private patient

	Gui, PFTReq:Add, Text, x500 y%yAdditive% vPFT_wardText, Ward:
	Gui, PFTReq:Add, Edit, x600 y%yAdditive% W150 vPFT_ward
	yAdditive := yAdditive + 35
	
	Gui, PFTReq:Add, Text, x10 y%yAdditive%, Urgency:
	Gui, PFTReq:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vPFT_urgency gPFT_urgency_click, Routine (4-6 weeks)||Urgent (< 2 weeks)|Prior to next outpatient appointment
	
	Gui, PFTReq:Add, Text, x500 y%yAdditive% vPFT_clinicDateText, Clinic date:
	Gui, PFTReq:Add, DateTime, x600 y%yAdditive% W150 vPFT_clinicDate choose%clinicIn6weeks%, dd/MM/yy
	yAdditive := yAdditive + 35

	Gui, PFTReq:Add, Text, x10 y%yAdditive%, Clinical information:
	Gui, PFTReq:Add, Edit, x%GUIStart% y%yAdditive% W500 H300 vPFT_clinicalDetails

	yAdditiveStart := yAdditive + 310
	yAdditive := yAdditiveStart
	
	; Insert tests
	For index, key in PFT_tests[]
	{
		variable := PFT_tests[key]

		if (PFT_tests[key] = "NIV")
		{
			Gui, PFTReq:Add, Checkbox, x%checkBoxStart% y%yAdditive% v%variable% +Right w%checkBoxWidth% gPFT_NIV_click, %key%:
		}
		else if (PFT_tests[key] = "bronchodilatorResponse")
		{
			Gui, PFTReq:Add, Checkbox, x%checkBoxStart% y%yAdditive% v%variable% +Right w%checkBoxWidth% gPFT_bronchodilatorResponse_click, %key%:
		}
		else
		{
			Gui, PFTReq:Add, Checkbox, x%checkBoxStart% y%yAdditive% v%variable% +Right w%checkBoxWidth%, %key%:
		}

		yAdditive += 20
		
		if (index == 8)
		{
			checkBoxStart := 270
			checkBoxWidth := 300
			yAdditive := yAdditiveStart
		}
	}
	
	Gui, PFTReq:Add, Checkbox, x10 y%yAdditive% vsalbutamolCheckBox +Right w250, Prescribe salbutamol:
	yAdditive := yAdditive + 20
	
	Gui, PFTReq:Add, Checkbox, x10 y%yAdditive% vipratropiumCheckBox +Right w250, Prescribe ipratropium bromide:
	yAdditive := 550
	
	Gui, PFTReq:Add, Text, x600 y%yAdditive% vNIVShowHide1, NIV mode:
	Gui, PFTReq:Add, Edit, x700 y%yAdditive% W150 vNIVMode
	yAdditive := yAdditive + 35
	
	Gui, PFTReq:Add, Text, x600 y%yAdditive% vNIVShowHide2, IPAP:
	Gui, PFTReq:Add, Edit, x700 y%yAdditive% W150 vNIVIPAP
	yAdditive := yAdditive + 35
	
	Gui, PFTReq:Add, Text, x600 y%yAdditive% vNIVShowHide3, EPAP:
	Gui, PFTReq:Add, Edit, x700 y%yAdditive% W150 vNIVEPAP
	yAdditive := yAdditive + 35
	
	Gui, PFTReq:Add, Button, x720 y%yAdditive% default,  &OK
	Gui, PFTReq:Add, Button, x780 y%yAdditive%, &Cancel

	; Initially hide certain controls
	GuiControl, PFTReq:Hide, PFT_wardText
	GuiControl, PFTReq:Hide, PFT_ward
	GuiControl, PFTReq:Hide, PFT_clinicDateText
	GuiControl, PFTReq:Hide, PFT_clinicDate
	GuiControl, PFTReq:Hide, NIVShowHide1
	GuiControl, PFTReq:Hide, NIVShowHide2
	GuiControl, PFTReq:Hide, NIVShowHide3
	GuiControl, PFTReq:Hide, NIVMode
	GuiControl, PFTReq:Hide, NIVIPAP
	GuiControl, PFTReq:Hide, NIVEPAP
	GuiControl, PFTReq:Hide, salbutamolCheckBox
	GuiControl, PFTReq:Hide, ipratropiumCheckBox

	Gui, PFTReq:Show,, Lung function request details
	Gui, PFTReq:+AlwaysOnTop
	WinWaitClose, Lung function request details
	return PFT_buttonPressed
}




PFTReqButtonOK:

errorString := ""
pass := False
atLeastOneInvestigationChecked := False
PFT_buttonPressed := "OK"
Gui PFTReq:Submit, Nohide


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

	
Gui, PFTReq:Destroy
PFT_details.consultantName := PFT_consultantName
PFT_details.doctorName := PFT_doctorName
PFT_details.signature := CurrentDirectory . "Signatures\" . PFT_doctorName . ".jpg"
PFT_details.speciality := PFT_speciality
PFT_details.location := PFT_location
PFT_details.ward := PFT_ward
PFT_details.urgency := PFT_urgency
PFT_details.clinicDate := PFT_clinicDate
PFT_details.clinicalDetails := PFT_clinicalDetails
PFT_details.holdBronchodilators := holdBronchodilators
PFT_details.spirometry := spirometry
PFT_details.gasTransfer := gasTransfer
PFT_details.staticLungVolumes := staticLungVolumes
PFT_details.FENO := FENO
PFT_details.CBG := CBG
PFT_details.sittingStanding := sittingStanding
PFT_details.bronchodilatorResponse := bronchodilatorResponse
PFT_details.osmohale := osmohale
PFT_details.fitToFly := fitToFly
PFT_details.SS := SS
PFT_details.CPAP := CPAP
PFT_details.mouthPressures := mouthPressures
PFT_details.overnightPulseOx := overnightPulseOx
PFT_details.occupationalAsthmaStudy := occupationalAsthmaStudy
PFT_details.NIV := NIV
PFT_details.NIVMode := NIVMode
PFT_details.NIVIPAP := NIVIPAP
PFT_details.NIVEPAP := NIVEPAP
PFT_details.salbutamolCheckBox := salbutamolCheckBox
PFT_details.ipratropiumCheckBox := ipratropiumCheckBox
return




PFTReqClose:
PFTReqButtonCancel:
PFTReqGuiClose:
PFT_buttonPressed := "close"
Gui, PFTReq:Destroy
return




PFT_NIV_click(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global NIV

	Gui PFTReq:Submit, NoHide

	if (NIV == 1)
	{
		GuiControl, PFTReq:Show, NIVShowHide1
		GuiControl, PFTReq:Show, NIVShowHide2
		GuiControl, PFTReq:Show, NIVShowHide3
		GuiControl, PFTReq:Show, NIVMode
		GuiControl, PFTReq:Show, NIVIPAP
		GuiControl, PFTReq:Show, NIVEPAP
	}
	else
	{
		GuiControl, PFTReq:Hide, NIVShowHide1
		GuiControl, PFTReq:Hide, NIVShowHide2
		GuiControl, PFTReq:Hide, NIVShowHide3
		GuiControl, PFTReq:Hide, NIVMode
		GuiControl, PFTReq:Hide, NIVIPAP
		GuiControl, PFTReq:Hide, NIVEPAP
	}

	return True
}




PFT_bronchodilatorResponse_click(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global bronchodilatorResponse

	Gui PFTReq:Submit, NoHide

	if (bronchodilatorResponse == 1)
	{
		GuiControl, PFTReq: , salbutamolCheckBox, 1
		GuiControl, PFTReq:Show, salbutamolCheckBox
		GuiControl, PFTReq:Show, ipratropiumCheckBox
	}
	else
	{
		GuiControl, PFTReq: , salbutamolCheckBox, 0
		GuiControl, PFTReq: , ipratropiumCheckBox, 0
		GuiControl, PFTReq:Hide, salbutamolCheckBox
		GuiControl, PFTReq:Hide, ipratropiumCheckBox
	}

	return True
}




PFT_location_click(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global PFT_location

	Gui PFTReq:Submit, NoHide

	if (PFT_location = "Inpatient")
	{
		GuiControl, PFTReq:Show, PFT_wardText
		GuiControl, PFTReq:Show, PFT_ward
	}
	else
	{
		GuiControl, PFTReq:Hide, PFT_wardText
		GuiControl, PFTReq:Hide, PFT_ward
	}

	return True
}




PFT_urgency_click(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global PFT_urgency

	Gui PFTReq:Submit, NoHide

	if (PFT_urgency = "Prior to next outpatient appointment")
	{
		GuiControl, PFTReq:Show, PFT_clinicDateText
		GuiControl, PFTReq:Show, PFT_clinicDate
	}
	else
	{
		GuiControl, PFTReq:Hide, PFT_clinicDateText
		GuiControl, PFTReq:Hide, PFT_clinicDate
	}

	return True
}




Create_PFT_request(MRN)
{
	Global Settings, PFT_templatePath, oWord, PatientDetails, PFT_details, PFT_latestRequestPath, settings, signatureHeight
	Global PBMainString
	
	Loop, 5
	{
		SavePath := Settings.RequestsFolder . "PFT request MRN" . MRN . "_"
		FormatTime, Today,, dd/MM/yy
		Yesterday += -1, Days
		FormatTime, Yesterday, %Yesterday%, dd/MM/yy
		clinicDateFormated := PFT_details.clinicDate
		FormatTime, clinicDateFormated, %clinicDateFormated%, dd/MM/yy
		signatureWidth := signatureHeight * GraphicsDimensions(PFT_details.signature)

		try
		{
			Loop
			{
				SavePathTemp := SavePath . A_Index . ".pdf"

				if(FileExist(SavePathTemp) == "")
				{
					SavePath := SavePathTemp
					PFT_latestRequestPath := SavePathTemp
					break
				}
			}

			oWord := ComObjCreate("Word.Application")
			oWord.Visible := False
			oWord.Documents.Open(PFT_templatePath)


			WriteAtBookmark("consultantName", PFT_details.consultantName)
			WriteAtBookmark("speciality", PFT_details.speciality)
			WriteAtBookmark("date", Today)

			WriteAtBookmark("name", PatientDetails.name)
			WriteAtBookmark("MRN", MRN)
			WriteAtBookmark("DOB", PatientDetails.DOB)

			if (PFT_details.location = "Outpatient")
			{
				InsertCheckboxAtBookmark("outpatient", 1)
				InsertCheckboxAtBookmark("inpatient", 0)
				InsertCheckboxAtBookmark("privatePatient", 0)
			}
			else if (PFT_details.location = "Inpatient")
			{
				InsertCheckboxAtBookmark("outpatient", 0)
				InsertCheckboxAtBookmark("inpatient", 1)
				WriteAtBookmark("ward", PFT_details.ward)
				InsertCheckboxAtBookmark("privatePatient", 0)
			}
			else if (PFT_details.location = "Private patient")
			{
				InsertCheckboxAtBookmark("outpatient", 0)
				InsertCheckboxAtBookmark("inpatient", 0)
				InsertCheckboxAtBookmark("privatePatient", 1)
			}


			if (PFT_details.urgency = "Routine (4-6 weeks)")
			{
				InsertCheckboxAtBookmark("routine", 1)
				InsertCheckboxAtBookmark("urgent", 0)
				InsertCheckboxAtBookmark("priorToNextOPA", 0)
			}
			else if (PFT_details.urgency = "Urgent (< 2 weeks)")
			{
				InsertCheckboxAtBookmark("routine", 0)
				InsertCheckboxAtBookmark("urgent", 1)
				InsertCheckboxAtBookmark("priorToNextOPA", 0)
			}
			else if (PFT_details.urgency = "Prior to next outpatient appointment")
			{
				InsertCheckboxAtBookmark("routine", 0)
				InsertCheckboxAtBookmark("urgent", 0)
				InsertCheckboxAtBookmark("priorToNextOPA", 1)
				WriteAtBookmark("clinicDate", clinicDateFormated)
			}

			; I am not going to change these at present
			InsertCheckboxAtBookmark("patientTransport", 0)
			InsertCheckboxAtBookmark("interpreter", 0)
			InsertCheckboxAtBookmark("learningDifficulties", 0)

			WriteAtBookmark("clinicalDetails", PFT_details.clinicalDetails)

			; I am not going to change these at present
			InsertCheckboxAtBookmark("DnV", 0)
			InsertCheckboxAtBookmark("haemoptysis", 0)
			InsertCheckboxAtBookmark("pneumothorax", 0)
			InsertCheckboxAtBookmark("aneurysm", 0)
			InsertCheckboxAtBookmark("tuberculosis", 0)
			InsertCheckboxAtBookmark("MI", 0)
			InsertCheckboxAtBookmark("stroke", 0)
			InsertCheckboxAtBookmark("PE", 0)
			InsertCheckboxAtBookmark("surgery", 0)
			;WriteAtBookmark("details", "details of surgery")

			InsertCheckboxAtBookmark("holdBronchodilators", PFT_details.holdBronchodilators)

			InsertCheckboxAtBookmark("spirometry", PFT_details.spirometry)
			InsertCheckboxAtBookmark("gasTransfer", PFT_details.gasTransfer)
			InsertCheckboxAtBookmark("staticLungVolumes", PFT_details.staticLungVolumes)
			InsertCheckboxAtBookmark("FENO", PFT_details.FENO)
			InsertCheckboxAtBookmark("CBG", PFT_details.CBG)
			InsertCheckboxAtBookmark("sittingStanding", PFT_details.sittingStanding)
			InsertCheckboxAtBookmark("bronchodilatorResponse", PFT_details.bronchodilatorResponse)

			InsertCheckboxAtBookmark("osmohale", PFT_details.osmohale)
			InsertCheckboxAtBookmark("fitToFly", PFT_details.fitToFly)
			InsertCheckboxAtBookmark("SS", PFT_details.SS)
			InsertCheckboxAtBookmark("CPAP", PFT_details.CPAP)
			InsertCheckboxAtBookmark("mouthPressures", PFT_details.mouthPressures)
			InsertCheckboxAtBookmark("overnightPulseOx", PFT_details.overnightPulseOx)
			InsertCheckboxAtBookmark("occupationalAsthmaStudy", PFT_details.occupationalAsthmaStudy)
			InsertCheckboxAtBookmark("NIV", PFT_details.NIV)

			if (PFT_details.NIV == 1)
			{
				WriteAtBookmark("NIVMode", PFT_details.NIVMode)
				WriteAtBookmark("NIVIPAP", "" PFT_details.NIVIPAP) ; need "" before number to not cause an error
				WriteAtBookmark("NIVEPAP", "" PFT_details.NIVEPAP) ; need "" before number to not cause an error
			}

			if (PFT_details.salbutamolCheckBox == 1)
			{
				InsertCheckboxAtBookmark("salbutamolCheckBox", 1)
				WriteAtBookmark("salbutamolSignature", Today)
				oWord.ActiveDocument.Bookmarks("salbutamolSignature").Select
				picObj := oWord.ActiveDocument.InLineShapes.AddPicture(PFT_details.signature)
				picObj.Height := signatureHeight 
				picObj.Width := signatureWidth
			}


			if (PFT_details.ipratropiumCheckBox == 1)
			{
				InsertCheckboxAtBookmark("ipratropiumCheckBox", 1)
				WriteAtBookmark("ipratropiumSignature", Today)
				oWord.ActiveDocument.Bookmarks("ipratropiumSignature").Select
				picObj := oWord.ActiveDocument.InLineShapes.AddPicture(PFT_details.signature)
				picObj.Height := signatureHeight 
				picObj.Width := signatureWidth
			}


			if (PFT_details.osmohale == 1)
			{
				InsertCheckboxAtBookmark("mannitolCheckBox", 1)
				WriteAtBookmark("mannitolSignature", Today)
				oWord.ActiveDocument.Bookmarks("mannitolSignature").Select
				picObj := oWord.ActiveDocument.InLineShapes.AddPicture(PFT_details.signature)
				picObj.Height := signatureHeight 
				picObj.Width := signatureWidth
			}


			if (PFT_details.fitToFly == 1)
			{
				InsertCheckboxAtBookmark("oxygenCheckBox", 1)
				WriteAtBookmark("oxygenSignature", Today)
				oWord.ActiveDocument.Bookmarks("oxygenSignature").Select
				picObj := oWord.ActiveDocument.InLineShapes.AddPicture(PFT_details.signature)
				picObj.Height := signatureHeight 
				picObj.Width := signatureWidth
			}
			
			WriteAtBookmark("signature", PFT_details.doctorName)
			oWord.ActiveDocument.Bookmarks("signature").Select
			picObj := oWord.ActiveDocument.InLineShapes.AddPicture(PFT_details.signature)
			picObj.Height := signatureHeight 
			picObj.Width := signatureWidth

			WriteAtBookmark("addressForReport", PFT_details.consultantName)

			UpdateProgressBar(80, "absolute", PBMainString . "Saving referral to requests folder")
					
			oWord.ActiveDocument.SaveAs(SavePath, 17) ; 17 is PDF format
			oWord.ActiveDocument.close(False)
			return True
		}
		catch, err
		{
			try
				oWord.ActiveDocument.close(False)
				
			if (A_index >= 5)
			{
				errorhandler(err, "lung function request", "Microsoft Word")
				return False
			}
			else
			{
				LogUpdate("Error during lung function request [counter: " . A_index . "]")
			}
		}
	}
	
	return True
}




; Bronchoscopy request via IPC
send_bronchoscopy(rawData)
{
	global Bronch_details, Bronch_emailCoordinators
	global CitrixSession, TestEmail, PBMainString
	global Bronch_Email, Bronch_Email_cc, Bronch_latestRequestPath, CurrentDirectory

	Pass := False

	CloseProgressBar()
	PBMainString := "Bronchoscopy request:`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Starting…")

	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	Bronch_details.procedure  := DS[2]
	Bronch_details.location  := DS[3]
	Bronch_details.primaryDiagnosis :=  DS[4]
	Bronch_details.significantComorbidities :=  DS[5]
	Bronch_details.preferredBronchoscopyDate := DS[6]
	Bronch_details.consentCompleted := DS[7]
	Bronch_details.imaging := DS[8]
	Bronch_details.anticoagulantAntiplatelets :=DS[9]
	Bronch_details.aim := DS[10]
	Bronch_details.consultantName :=DS[11]
	Bronch_details.doctorName := DS[12]
	Bronch_details.signature := CurrentDirectory . "Signatures\" . .  DS[12] . ".jpg"

	if (DS[13] = "testEmail")
	{
		emailAddress := TestEmail
		Bronch_emailCoordinators := 0
	}
	else
	{
		emailAddress := Bronch_Email
		Bronch_emailCoordinators := 1
	}


	if GetPatientDetailsTrakcare(MRN)
	{
		UpdateProgressBar(60, "absolute", PBMainString . "Creating request…")

		if Create_bronchoscopy_request(MRN)
		{
			UpdateProgressBar(20,, PBMainString . "Emailing bronchoscopy request…")
			
			if (Bronch_emailCoordinators == 1)
				cc_addresses := Bronch_Email_cc
			else
				cc_addresses := ""
			
			if (CitrixSession == True)
			{
				MB("Current cannot send via a Citrix session. Please manually locate your request in the requests folder in the master directory and send via the NHSmail web app")
				;NHSMailOpen()
				;Email_NHS_Mail(emailAddress, cc_addresses, "Bronchoscopy request", "Please find attached a bronchoscopy request", Bronch_latestRequestPath)
				; pass := True
			}
			else
			{
				if EmailOutlook(emailAddress, cc_addresses, "Bronchoscopy request", "Please find attached a bronchoscopy request", Bronch_latestRequestPath)
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




get_bronchoscopy_extraInfo(MRN)
{
	global
	Bronch_details := {}
	
	if !(bronchoscopy_GUI(MRN) = "OK")
		return False
	
	if !FileExist(Bronch_details.signature)
	{
		MB("No signature found for the requesting doctor. Ending bronchoscopy request")
		return False
	}

	return True
}




bronchoscopy_GUI(MRN)
{
	global dialogueColour, patientDetails, Bronch_details, consultantListConstant, doctorListConstant, Bronch_buttonPressed
	global Bronch_procedure, Bronch_location, Bronch_primaryDiagnosis, Bronch_significantComorbidities, Bronch_preferredBronchoscopyDate, Bronch_consentCompleted
	global Bronch_anticoagulantAntiplatelets, Bronch_aim, Bronch_imaging, Bronch_consultantName, Bronch_doctorName
	global Bronch_emailCoordinators

	GUIStart := 250
	GUIWidth := 500
	CurrentUser := ConvertUsername(A_UserName) ;StrReplace(A_UserName, ".", " ")
	CurrentUserPostionInListConsultant := InStr(consultantListConstant, CurrentUser)  ; 0  if not found
	CurrentUserPostionInListDoctor := InStr(doctorListConstant, CurrentUser)  ; 0  if not found
	yAdditiveStart := 10
	yAdditive := yAdditiveStart
	checkBoxStart := 10
	checkBoxWidth := 250
	clinicIn6weeks += 42, Days
	FormatTime, clinicIn6weeks, %clinicIn6weeks%, dd/MM/yy


	if (CurrentUserPostionInListConsultant == 0)
		consultantList := consultantListConstant
	else
		consultantList := SubStr(consultantListConstant, 1, CurrentUserPostionInListConsultant + StrLen(CurrentUser)) . "|" . SubStr(consultantListConstant, CurrentUserPostionInListConsultant + StrLen(CurrentUser) + 1)
		
	
	if (CurrentUserPostionInListDoctor == 0)
		doctorList := doctorListConstant
	else
		doctorList := SubStr(doctorListConstant, 1, CurrentUserPostionInListDoctor + StrLen(CurrentUser)) . "|" . SubStr(doctorListConstant, CurrentUserPostionInListDoctor + StrLen(CurrentUser) + 1)
		
		
	Gui, BronchReq:Font, s12
	Gui, BronchReq:Color, %dialogueColour%
	Gui, BronchReq:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive := yAdditive + 40

	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Procedure:
	Gui, BronchReq:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vBronch_procedure, EBUS||Bronchoscopy|Thoracoscopy
	yAdditive := yAdditive + 35

	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Patient's current location:
	Gui, BronchReq:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vBronch_location, Outpatient||Inpatient
	yAdditive := yAdditive + 35

	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Primary diagnosis:
	Gui, BronchReq:Add, Edit, x%GUIStart% y%yAdditive% H70 W%GUIWidth% vBronch_primaryDiagnosis
	yAdditive := yAdditive + 77

	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Significant co-morbidities:
	Gui, BronchReq:Add, Edit, x%GUIStart% y%yAdditive% H70 W%GUIWidth% vBronch_significantComorbidities
	yAdditive := yAdditive + 77

	Gui, BronchReq:Add, Text, x10 y%yAdditive% , Preferred bronchoscopy date:
	Gui, BronchReq:Add, DateTime, x%GUIStart% y%yAdditive% W%GUIWidth% vBronch_preferredBronchoscopyDate
	yAdditive := yAdditive + 35

	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Consent completed:
	Gui, BronchReq:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vBronch_consentCompleted, No||Yes|
	yAdditive := yAdditive + 35

	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Imaging (modality and result):
	Gui, BronchReq:Add, Edit, x%GUIStart% y%yAdditive% H70 W%GUIWidth% vBronch_imaging
	yAdditive := yAdditive + 77

	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Anticoagulant / anti-platelets:
	Gui, BronchReq:Add, Edit, x%GUIStart% y%yAdditive% H70 W%GUIWidth% vBronch_anticoagulantAntiplatelets
	yAdditive := yAdditive + 77

	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Aim of bronchoscopy / sampling:
	Gui, BronchReq:Add, Edit, x%GUIStart% y%yAdditive% H70 W%GUIWidth% vBronch_aim
	yAdditive := yAdditive + 77

	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Consultant:
	Gui, BronchReq:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vBronch_consultantName, %consultantList%
	yAdditive := yAdditive + 40
	
	Gui, BronchReq:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, BronchReq:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vBronch_doctorName, %doctorList%
	yAdditive := yAdditive + 40
	
	Gui, BronchReq:Add, CheckBox, x5 y%yAdditive%  vBronch_emailCoordinators +Right Checked, Email coordinators (April, Lisa + Natasha): 
	yAdditive += 40
	
	Gui, BronchReq:Add, Button, x720 y%yAdditive% default,  &OK
	Gui, BronchReq:Add, Button, x780 y%yAdditive%, &Cancel
	

	Gui, BronchReq:Show,, Bronchoscopy request details
	Gui, BronchReq:+AlwaysOnTop
	WinWaitClose, Bronchoscopy request details
	return Bronch_buttonPressed
}




BronchReqButtonOK:

errorString := ""
pass := False
Bronch_buttonPressed := "OK"
Gui BronchReq:Submit, Nohide

fieldsArr := [Bronch_procedure, Bronch_location, Bronch_primaryDiagnosis, Bronch_significantComorbidities, Bronch_preferredBronchoscopyDate, Bronch_consentCompleted, Bronch_imaging, Bronch_anticoagulantAntiplatelets,  Bronch_aim, Bronch_consultantName]

if (Bronch_consultantName == "")
	errorString := errorString . "- No consultant name was chosen.`n"

if (Bronch_doctorName == "")
	errorString := errorString . "- No requesting doctor was chosen.`n"
	
for index, key in fieldsArr
	if (key = "")
	{
		errorString := errorString . "- Empty field(s).`n"
		break
	}

	
if (errorString != "")
{
	MB("Please correct the below errors and submit again:`n" . errorString)
	return
}


Gui, BronchReq:Destroy

Bronch_details.procedure  := Bronch_procedure
Bronch_details.location  := Bronch_location
Bronch_details.primaryDiagnosis := Bronch_primaryDiagnosis
Bronch_details.significantComorbidities := Bronch_significantComorbidities
FormatTime, preferredDateFormated, %Bronch_preferredBronchoscopyDate%, dd/MM/yy
Bronch_details.preferredBronchoscopyDate := preferredDateFormated
Bronch_details.consentCompleted := Bronch_consentCompleted
Bronch_details.imaging := Bronch_imaging
Bronch_details.anticoagulantAntiplatelets := Bronch_anticoagulantAntiplatelets
Bronch_details.aim := Bronch_aim
Bronch_details.consultantName := Bronch_consultantName
Bronch_details.doctorName := Bronch_doctorName
Bronch_details.signature := CurrentDirectory . "Signatures\" .  Bronch_doctorName . ".jpg"
return




BronchReqClose:
BronchReqButtonCancel:
BronchReqGuiClose:
Bronch_buttonPressed := "close"
Gui, BronchReq:Destroy
return




Create_bronchoscopy_request(MRN)
{
	Global Settings, Bronch_templatePath, oWord, PatientDetails, Bronch_details
	Global Bronch_latestRequestPath, signatureHeight, PBMainString

	Loop, 5
	{
		SavePath := Settings.RequestsFolder . "Bronchoscopy request MRN" . MRN . "_"
		FormatTime, Today,, dd/MM/yy
		signatureWidth := signatureHeight * GraphicsDimensions(Bronch_details.signature)

		try
		{
			Loop
			{
				SavePathTemp := SavePath . A_Index . ".pdf"

				if(FileExist(SavePathTemp) == "")
				{
					SavePath := SavePathTemp
					Bronch_latestRequestPath := SavePathTemp
					break
				}
			}

			oWord := ComObjCreate("Word.Application")
			oWord.Visible := False
			oWord.Documents.Open(Bronch_templatePath)


			WriteAtBookmark("name", PatientDetails.name)
			WriteAtBookmark("MRN", MRN)
			WriteAtBookmark("DOB", PatientDetails.DOB)


			WriteAtBookmark("procedure", Bronch_details.procedure)
			WriteAtBookmark("location", Bronch_details.location)
			WriteAtBookmark("primaryDiagnosis", Bronch_details.primaryDiagnosis)
			WriteAtBookmark("significantComorbidities", Bronch_details.significantComorbidities)
			WriteAtBookmark("preferredBronchoscopyDate", Bronch_details.preferredBronchoscopyDate)
			WriteAtBookmark("consentCompleted", Bronch_details.consentCompleted)
			WriteAtBookmark("imaging", Bronch_details.imaging)
			WriteAtBookmark("anticoagulantAntiplatelets", Bronch_details.anticoagulantAntiplatelets)
			WriteAtBookmark("aim", Bronch_details.aim)

			WriteAtBookmark("consultantName", Bronch_details.consultantName)
			WriteAtBookmark("doctorName", Today)
			WriteAtBookmark("doctorName", Bronch_details.doctorName . ", ")
			oWord.ActiveDocument.Bookmarks("doctorName").Select

			picObj := oWord.ActiveDocument.InLineShapes.AddPicture(Bronch_details.signature)
			picObj.Height := signatureHeight 
			picObj.Width := signatureWidth

			UpdateProgressBar(80, "absolute", PBMainString . "Saving referral to requests folder")
					
			oWord.ActiveDocument.SaveAs(SavePath, 17) ; 17 is PDF format
			oWord.ActiveDocument.close(False)
			return True
		}
		catch, err
		{
			try
				oWord.ActiveDocument.close(False)
				
			if (A_index >= 5)
			{
				errorhandler(err, "lung bronchoscopy request", "Microsoft Word")
				return False
			}
			else
			{
				LogUpdate("Error during bronchoscopy request [counter: " . A_index . "]")
			}
		}
	}
	
	return True
}




; HLSG request via IPC
send_HLSG(rawData)
{
	global HLSG_details, CitrixSession, TestEmail, PBMainString
	global HLSG_Email, HLSG_latestRequestPath, CurrentDirectory

	Pass := False

	CloseProgressBar()
	PBMainString := "Healthy Lifestyle Gloucestershire request:`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Starting…")
	
	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	HLSG_details.Comorbidities := DS[2]
	HLSG_details.ethnicity := DS[3]
	HLSG_details.Comorbidities := DS[4]
	HLSG_details.smoking  := DS[5]
	HLSG_details.alcohol  := DS[6]
	HLSG_details.weight  := DS[7]
	HLSG_details.physicalActivity  := DS[8]
	HLSG_details.doctorName := DS[9]
	HLSG_details.signature := CurrentDirectory .  DS[9] . ".jpg"

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

		if Create_HLSG_request(MRN)
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




get_HLSG_extraInfo(MRN)
{
	global
	HLSG_details := {}


	if !(HLSG_GUI(MRN) = "OK")
		return False
	
	if !FileExist(HLSG_details.signature)
	{
		MB("No signature found for the requesting doctor. Ending bronchoscopy request")
		return False
	}

	return True
}




HLSG_GUI(MRN)
{
	global dialogueColour, patientDetails, HLSG_details, doctorListConstant, HLSG_buttonPressed
	global HLSG_doctorName, HLSG_buttonPressed, HLSG_referralTypes
	global HLSG_ethnicity, HLSG_Comorbidities

	GUIStart := 200
	GUIWidth := 300
	CurrentUser := StrReplace(A_UserName, ".", " ")
	CurrentUserPostionInListDoctor := InStr(doctorListConstant, CurrentUser)  ; 0  if not found
	yAdditiveStart := 10
	yAdditive := yAdditiveStart
	checkBoxStart := 10
	checkBoxWidth := 180
	Title := "Healthy Lifestyles Gloucestershire referral details"


	if (CurrentUserPostionInListDoctor == 0)
		doctorList := doctorListConstant
	else
		doctorList := SubStr(doctorListConstant, 1, CurrentUserPostionInListDoctor + StrLen(CurrentUser)) . "|" . SubStr(doctorListConstant, CurrentUserPostionInListDoctor + StrLen(CurrentUser) + 1)
		
		
	Gui, HLSGRef:Font, s12
	Gui, HLSGRef:Color, %dialogueColour%
	Gui, HLSGRef:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive := yAdditive + 40

	Gui, HLSGRef:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, HLSGRef:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vHLSG_doctorName, %doctorList%
	yAdditive := yAdditive + 40
	
	Gui, HLSGRef:Add, Text, x10 y%yAdditive%, Ethnicity:
	Gui, HLSGRef:Add, ComboBox, x%GUIStart% y%yAdditive% W%GUIWidth% vHLSG_ethnicity, White||Black|Asian|; I appreciate this is a short ethnicity list. Will have to check best list of groups to use
	yAdditive := yAdditive + 40
	
	Gui, HLSGRef:Add, Text, x10 y%yAdditive%, Co-morbidities:
	Gui, HLSGRef:Add, Edit, x%GUIStart% y%yAdditive% H70 W%GUIWidth% vHLSG_Comorbidities
	yAdditive := yAdditive + 77


	; insert referral types
	For index, key in HLSG_referralTypes[]
	{
		variable := HLSG_referralTypes[key]
		Gui, HLSGRef:Add, Checkbox, x%checkBoxStart% y%yAdditive% v%variable% +Right w%checkBoxWidth%, %key%:
		yAdditive += 20
	}


	Gui, HLSGRef:Add, Button, x400 y%yAdditive% default,  &OK
	Gui, HLSGRef:Add, Button, x450 y%yAdditive%, &Cancel

	Gui, HLSGRef:Show,, %Title%
	Gui, HLSGRef:+AlwaysOnTop
	WinWaitClose, %Title%
	return HLSG_buttonPressed
}




HLSGRefButtonOK:

errorString := ""
pass := False
HLSG_buttonPressed := "OK"
Gui HLSGRef:Submit, Nohide
atLeastOneReferralTypeTicked := False


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


Gui, HLSGRef:Destroy
HLSG_details.Comorbidities := HLSG_Comorbidities
HLSG_details.ethnicity := HLSG_ethnicity
HLSG_details.smoking  := HLSG_smoking
HLSG_details.alcohol  := HLSG_alcohol
HLSG_details.weight  := HLSG_weight
HLSG_details.physicalActivity  := HLSG_physicalActivity
HLSG_details.doctorName := HLSG_doctorName
HLSG_details.signature := CurrentDirectory . "Signatures\" .  HLSG_doctorName . ".jpg"
return




HLSGRefClose:
HLSGRefButtonCancel:
HLSGRefGuiClose:
HLSG_buttonPressed := "close"
Gui, HLSGRef:Destroy
return




Create_HLSG_request(MRN)
{
	Global Settings, HLSG_templatePath, oWord, PatientDetails, HLSG_details, HLSG_latestRequestPath, signatureHeight
	Global DeptAddress
	
	Loop, 5
	{
		SavePath := Settings.RequestsFolder . "HLSG request MRN" . MRN . "_"
		FormatTime, Today,, dd/MM/yy
		signatureWidth := signatureHeight * GraphicsDimensions(HLSG_details.signature)

		try
		{
			Loop
			{
				SavePathTemp := SavePath . A_Index . ".pdf"

				if(FileExist(SavePathTemp) == "")
				{
					SavePath := SavePathTemp
					HLSG_latestRequestPath := SavePathTemp
					break
				}
			}

			oWord := ComObjCreate("Word.Application")
			oWord.Visible := False
			oWord.Documents.Open(HLSG_templatePath)
			WriteAtBookmark("referrerDetails", DeptAddress)
			WriteAtBookmark("referrerDetails", "Dr " . HLSG_details.doctorName . ", ")
			
			WriteAtBookmark("name", PatientDetails.name)
			WriteAtBookmark("DOB", PatientDetails.DOB)
			WriteAtBookmark("ethnicity", HLSG_details.ethnicity)
			
			if (PatientDetails.gender = "Male")
			{
				InsertCheckboxAtBookmark("male", 1)
				InsertCheckboxAtBookmark("female", 0)
				InsertCheckboxAtBookmark("transgender", 0)
			}
			else if (PatientDetails.gender = "Female")
			{
				InsertCheckboxAtBookmark("male", 0)
				InsertCheckboxAtBookmark("female", 1)
				InsertCheckboxAtBookmark("transgender", 0)
			}
			else
			{
				InsertCheckboxAtBookmark("male", 0)
				InsertCheckboxAtBookmark("female", 0)
				InsertCheckboxAtBookmark("transgender", 0)
			}
			
			WriteAtBookmark("address", PatientDetails.postCode)
			WriteAtBookmark("address", PatientDetails.address4 . ", ")
			WriteAtBookmark("address", PatientDetails.address3 . ", ")
			WriteAtBookmark("address", PatientDetails.address2 . ", ")
			WriteAtBookmark("address", PatientDetails.address1 . ", ")
			
			WriteAtBookmark("telephoneNr", ", " . PatientDetails.MobNo)
			WriteAtBookmark("telephoneNr", PatientDetails.TeleNo)
			
			WriteAtBookmark("issues", HLSG_details.Comorbidities)
			
			WriteAtBookmark("referrerName", "Dr " . HLSG_details.doctorName)
			oWord.ActiveDocument.Bookmarks("referrerName").Select
			picObj := oWord.ActiveDocument.InLineShapes.AddPicture(HLSG_details.signature)
			picObj.Height := signatureHeight 
			picObj.Width := signatureWidth
			WriteAtBookmark("date", Today)
			
			if (HLSG_details.smoking == 1)
				InsertCheckboxAtBookmark("smoking", 1)
			else
				InsertCheckboxAtBookmark("smoking", 0)
			
			if (HLSG_details.alcohol == 1)
				InsertCheckboxAtBookmark("alcohol", 1)
			else
				InsertCheckboxAtBookmark("alcohol", 0)

			if (HLSG_details.weight == 1)
				InsertCheckboxAtBookmark("weight", 1)
			else
				InsertCheckboxAtBookmark("weight", 0)

			if (HLSG_details.physicalActivity == 1)
				InsertCheckboxAtBookmark("physicalActivity", 1)
			else
				InsertCheckboxAtBookmark("physicalActivity", 0)			
			
			UpdateProgressBar(80, "absolute", PBMainString . "Saving referral to requests folder")
					
			oWord.ActiveDocument.SaveAs(SavePath, 17) ; 17 is PDF format
			oWord.ActiveDocument.close(False)
			return True
		}
		catch, err
		{
			try
				oWord.ActiveDocument.close(False)
				
			if (A_index >= 5)
			{
				errorhandler(err, "HLSG request", "Microsoft Word")
				return False
			}
			else
			{
				LogUpdate("Error during HLSG request [counter: " . A_index . "]")
			}
		}
	}
	
	return False
}




GOV_UK_Notify()
{
	global dialogueColour, GUKN_type
	
	title := "Notification type"
	GUIStart := 60
	GUIWidth := 100
	yAdditive := 10
	
	Gui, GUKN_main:Font, s12
	Gui, GUKN_main:Color, % dialogueColour
	Gui, GUKN_main:Add, Text, x10 y%yAdditive%, Type:
	Gui, GUKN_main:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_type, SMS||email|letter
	
	yAdditive := yAdditive + 50
	
	Gui, GUKN_main:Add, Button, x150 y%yAdditive% default gGUKN_mainOK,  &OK
	Gui, GUKN_main:Add, Button, x200 y%yAdditive% gGUKN_mainClose,  &Cancel
	Gui, GUKN_main:Show,, % title
	Gui, GUKN_main:+AlwaysOnTop
	WinWaitClose, % title
	return
}



; First use of functions rather than Goto statements here. Will have to incoorporate this through the code to remove GoTo statements with their associated return statement
GUKN_mainOK()
{
	global GUKN_type
	
	Gui, GUKN_main:Submit, NoHide
	Gui, GUKN_main:Destroy
	
	if (GUKN_type == "SMS")
	{
		GUKN_SMS()
	}
	else if (GUKN_type == "email")
	{
		GUKN_email()
	}
	else if (GUKN_type == "letter")
	{
		GUKN_letter()
	}
	
	return
}




GUKN_mainButtonCancel()
{
	GUKN_mainClose()
}
GUKN_mainGuiClose()
{
	GUKN_mainClose()
}
GUKN_mainClose()
{
	Gui, GUKN_main:Destroy
	return
}




GUKN_SMS()
{
	global dialogueColour, GUKN_mobileNumber, GUKN_SMS_message, mobile1
	
	title := "Send SMS message"
	GUIStart := 120
	GUIWidth := 300
	yAdditive := 10
	
	Gui, GUKN_SMS:Font, s12
	Gui, GUKN_SMS:Color, % dialogueColour
	Gui, GUKN_SMS:Add, Text, x10 y%yAdditive%, Mobile number:
	Gui, GUKN_SMS:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_mobileNumber, % mobile1
	yAdditive := yAdditive + 50
	
	Gui, GUKN_SMS:Add, Text, x10 y%yAdditive%, Message:
	Gui, GUKN_SMS:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H100 vGUKN_SMS_message, % "Test Message, change as needed"
	yAdditive := yAdditive + 120
	
	Gui, GUKN_SMS:Add, Button, x300 y%yAdditive% default gGUKN_SMSOK,  &OK
	Gui, GUKN_SMS:Add, Button, x350 y%yAdditive% gGUKN_SMSClose,  &Cancel
	Gui, GUKN_SMS:Show,, % title
	Gui, GUKN_SMS:+AlwaysOnTop
	WinWaitClose, % title

	return
}




GUKN_SMSOK()
{
	global CurrentDirectory, dialogueColour, API_key
	global SMS_template_ID, GUKN_mobileNumber
	global GUKN_SMS_message
	
	pass := False
	
	Gui, GUKN_SMS:Submit, NoHide
	Gui, GUKN_SMS:Destroy
	
	if !MobileNumberCheck(GUKN_mobileNumber)
		return "Fail"

	MobileNumber := "+44" . SubStr(GUKN_mobileNumber, 2)
	variables := "SMS;" . API_key . ";" . SMS_template_ID . ";" .  MobileNumber . ";" .  GUKN_SMS_message
	sendMessageGUKN("SMS message", variables)
	return
}




GUKN_SMSButtonCancel()
{
	GUKN_SMSClose()
}
GUKN_SMSGuiClose()
{
	GUKN_SMSClose()
}
GUKN_SMSClose()
{
	Gui, GUKN_SMS:Destroy
	return
}




GUKN_email()
{
	global dialogueColour, GUKN_email_address, GUKN_email_subject, GUKN_email_body
	
	title := "Send email"
	GUIStart := 120
	GUIWidth := 300
	yAdditive := 10
	
	Gui, GUKN_email:Font, s12
	Gui, GUKN_email:Color, % dialogueColour
	Gui, GUKN_email:Add, Text, x10 y%yAdditive%, email:
	Gui, GUKN_email:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_email_address, % "mark.allan.bailey@gmail.com"
	yAdditive := yAdditive + 50
	
	Gui, GUKN_email:Add, Text, x10 y%yAdditive%, % "Subject:"
	Gui, GUKN_email:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H100 vGUKN_email_subject, % "Test subject, change as needed"
	yAdditive := yAdditive + 120
	
	Gui, GUKN_email:Add, Text, x10 y%yAdditive%, % "Body:"
	Gui, GUKN_email:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H100 vGUKN_email_body, % "Test body, change as needed"
	yAdditive := yAdditive + 120
	
	Gui, GUKN_email:Add, Button, x200 y%yAdditive% default gGUKN_emailOK,  &OK
	Gui, GUKN_email:Add, Button, x250 y%yAdditive% gGUKN_emailClose,  &Cancel
	Gui, GUKN_email:Show,, % title
	Gui, GUKN_email:+AlwaysOnTop
	WinWaitClose, % title

	return
}




GUKN_emailOK()
{
	global CurrentDirectory, dialogueColour, API_key
	global email_template_ID
	global GUKN_email_address, GUKN_email_subject, GUKN_email_body
	
	Gui, GUKN_email:Submit, NoHide
	Gui, GUKN_email:Destroy
	
	if (not InStr(GUKN_email_address, "@"))
	{
		msgbox, % "Error with email address, please try again!"
		return False
	}

	variables := "email;" . API_key . ";" . email_template_ID . ";" .  GUKN_email_address . ";" .  GUKN_email_subject . ";" . GUKN_email_body
	sendMessageGUKN("email", variables)
	return True
}




GUKN_emailButtonCancel()
{
	GUKN_emailClose()
}
GUKN_emailGuiClose()
{
	GUKN_emailClose()
}
GUKN_emailClose()
{
	Gui, GUKN_email:Destroy
	return
}




GUKN_letter()
{
	global 
	
	local title := "Send letter"
	local GUIStart := 120
	local GUIWidth := 600
	local yAdditive := 10
	
	Gui, GUKN_letter:Font, s12
	Gui, GUKN_letter:Color, % dialogueColour
	Gui, GUKN_letter:Add, Text, x10 y%yAdditive%, To:
	Gui, GUKN_letter:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_letter_to, % Username1
	yAdditive := yAdditive + 50
	
	Gui, GUKN_letter:Add, Text, x10 y%yAdditive%, Address:
	Gui, GUKN_letter:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H100 vGUKN_letter_address, % address1
	yAdditive := yAdditive + 120
	
	Gui, GUKN_letter:Add, Text, x10 y%yAdditive%, % "From:"
	Gui, GUKN_letter:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H30 vGUKN_letter_from, % clinician1
	yAdditive := yAdditive + 40
	
	Gui, GUKN_letter:Add, Text, x10 y%yAdditive%, % "Header:"
	Gui, GUKN_letter:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H100 vGUKN_letter_header, % "Test header, change as needed"
	yAdditive := yAdditive + 120
	
	Gui, GUKN_letter:Add, Text, x10 y%yAdditive%, % "Letter body:"
	Gui, GUKN_letter:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H100 vGUKN_letter_body, % "Test body, change as needed"
	yAdditive := yAdditive + 120
	
	Gui, GUKN_letter:Add, Button, x200 y%yAdditive% default gGUKN_letterOK,  &OK
	Gui, GUKN_letter:Add, Button, x250 y%yAdditive% gGUKN_letterClose,  &Cancel
	Gui, GUKN_letter:Show,, % title
	Gui, GUKN_letter:+AlwaysOnTop
	WinWaitClose, % title
	return
}




GUKN_letterOK()
{
	global CurrentDirectory, dialogueColour, API_key
	global letter_template_ID
	global GUKN_letter_to, GUKN_letter_address
	global GUKN_letter_from, GUKN_letter_header, GUKN_letter_body
	
	Gui, GUKN_letter:Submit, NoHide
	Gui, GUKN_letter:Destroy
	
	addressConverted := GUKN_letter_to . "`n" . GUKN_letter_address
	addressConverted := strReplace(addressConverted, "`n", "//")
	variables := "letter;" . API_key . ";" . letter_template_ID . ";" . addressConverted . ";" . GUKN_letter_from . ";" .  GUKN_letter_header . ";" . GUKN_letter_body
	sendMessageGUKN("Letter", variables)
	return
}




GUKN_letterButtonCancel()
{
	GUKN_letterClose()
}
GUKN_letterGuiClose()
{
	GUKN_letterClose()
}
GUKN_letterClose()
{
	Gui, GUKN_letter:Destroy
	return
}




sendMessageGUKN(method, variables, MMF := False)
{
	global pythonEXE, pythonEXEPath
	
	outcome := "timeOut"
	
	if MMF
	{
		CloseProgressBar()
		PBMainString := "Sending " . method . ":`n"
		CreateProgressBar()
		UpdateProgressBar(5, "absolute", PBMainString . ".")
	}

	pythonMessages := new MemoryMappedFile_IPC("AHK_2_Python_IPC", False)
	pythonMessages.send("GOV_UK_notify", variables)
	UpdateProgressBar(10, "absolute", PBMainString . "..")
	
	if pythonEXE
		Run, % pythonEXEPath,, Hide
	
	UpdateProgressBar(60, "absolute", PBMainString . "...")
	
	Loop 600		; 600 x 100 = 1 min
	{
		returnResult := pythonMessages.read()
		returnResultSplit := strSplit(returnResult, "|")
		if (returnResultSplit[1] = "SPRead")
		{
			outcome := returnResultSplit[2]
			break
		}
		
		sleep 100
	}
	
	UpdateProgressBar(100, "absolute", PBMainString . "Message sent via " . method)
	sleep 2000
	CloseProgressBar()
	
	if (instr(outcome, "pass") = 1)
	{
		;msgbox, % method . " sent " . outcome
		return True
	}
	else if (outCome = "timeOut")
	{
		msgbox, % "The GOV.uk Notification program (written in python) did not start and hence the " . method . " has not been sent!"
		return False
	}
	else if (outcome = "Wrong initial arguement")
	{
		msgbox, % "Wrong initial arguement provided"
		return False
	}
	else if (outcome = "fail")
	{
		msgbox, % method . " - request failed!"
		return False
	}
	else if (outcome = "wrong method")
	{
		msgbox, % method . " - wrong method provided to python script"
		return False
	}
	else
	{
		msgbox, % "AHK error with " . method . " function [" . outcome . "]!"
		return False
	}	
	
	return False
}




send_GUKN(rawData)
{
	global API_key, SMS_template_ID, email_template_ID, letter_template_ID, PatientDetails
	
	pass := False
	mobileNumber := ""
	
	
	DS := StrSplit(rawData, ";")
	MRN := DS[1]
	method := DS[2]
	dataRetrieval := DS[3]
	
	CloseProgressBar()
	PBMainString := "Sending " . method . ":`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . "Creating request...")
	
	if (method = "SMS")
	{
		if (dataRetrieval = "VBA")
			mobileNumber := DS[4]
		else if (dataRetrieval = "Trakcare")
		{
			GetPatientDetailsTrakcare(MRN)
			mobileNumber := PatientDetails.MobNo
		}
		else
		{
			msgbox, % "Wrong dataRetrieval arguement provided!"
			GoTo Fail
		}
		
		if !MobileNumberCheck(mobileNumber)
			GoTo Fail
		
		mobileNumber := "+44" . SubStr(mobileNumber, 2)
		
		GUKN_SMS_message := DS[5]
		variables := "SMS;" . API_key . ";" . SMS_template_ID . ";" .  MobileNumber . ";" .  GUKN_SMS_message
		UpdateProgressBar(55, "absolute", PBMainString . "Sending request to GOV.uk Notify...")
		sendMessageGUKN("SMS message", variables, True)
	}
	else if (method = "email")
	{
		if (dataRetrieval = "VBA")
		{
			GUKN_email_address := DS[4]
		}
		else if (dataRetrieval = "Trakcare")
		{
			GetPatientDetailsTrakcare(MRN)
			GUKN_email_address := PatientDetails.email
		}
		else
		{
			msgbox, % "Wrong dataRetrieval arguement provided!"
			GoTo Fail
		}

		GUKN_email_subject := DS[5]
		GUKN_email_body := DS[6]
		
		if (not InStr(GUKN_email_address, "@"))
		{
			msgbox, % "Error with email address, please try again!"
			GoTo Fail
		}

		variables := "email;" . API_key . ";" . email_template_ID . ";" .  GUKN_email_address . ";" .  GUKN_email_subject . ";" . GUKN_email_body
		UpdateProgressBar(55, "absolute", PBMainString . "Sending request to GOV.uk Notify...")
		sendMessageGUKN("email", variables, True)
	}
	else if (method = "letter")
	{
		GUKN_letter_to := DS[4]
		GUKN_letter_address := DS[5]
		addressConverted := GUKN_letter_to . "`n" . GUKN_letter_address
		addressConverted := strReplace(addressConverted, "`n", "//")
		GUKN_letter_from := DS[6]
		GUKN_letter_header := DS[7]
		GUKN_letter_body := DS[8]
		variables := "letter;" . API_key . ";" . letter_template_ID . ";" . addressConverted . ";" . GUKN_letter_from . ";" .  GUKN_letter_header . ";" . GUKN_letter_body
		UpdateProgressBar(55, "absolute", PBMainString . "Sending request to GOV.uk Notify...")
		sendMessageGUKN("Letter", variables, True)
	}
	else
	{
		msgbox, % "Wrong method send for sending messages from Access to AHK!"
		GoTo Fail
	}
	
	CloseProgressBar()
	return "Pass"
	
Fail:
	CloseProgressBar()
	return "Fail"
}




MobileNumberCheck(number)
{
	number := StrReplace(number, " ", "")
	
	if (StrLen(number) == 11)
		if (SubStr(number,1,2) == "07")
			if (not number ~= "[^0-9]")
				return True
	
	; Return false if above fails
	msgbox, % "Error with mobile number. Please try again"
	return False

}




referralRequestCreateAndSend(RRtype)
{
	global

	local RRMessage := RRMessageAddSuffix(RRtype)
	local RRFunctionCR := "create_" . RRtype . "_request"
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

		if RRGetExtraInfo(MRN, RRtype)
		{
			CreateProgressBar()
			UpdateProgressBar(60, "absolute", PBMainString . "Creating referral...")

			if %RRFunctionCR%(MRN)
			{
				UpdateProgressBar(20,, PBMainString . "Emailing referral...")
				requestEmail(RRtype)
			}
		}
	}
	CloseProgressBar()
	runningStatus("done")
	LogUpdate("Sleepstation request")
	return True
}




RRMessageAddSuffix(RRType)
{
	global
	
	local RRMessage := RRtype
	
	; Add request or referral to end
	for index, element in requestArray
	{
		if (RRMessage == element)
		{
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
				RRMessage .= " referral"
				break
			}
		}
		
		if (RRMessage == RRtype)
		{
			throw Exception("'RRtype' has not been matched up with either a request or referral!", -1)
		}
	}
	
	return RRMessage
}




requestEmail(RRtype)
{
	global 
	
	local RRMessage := RRMessageAddSuffix(RRtype)

	RRLatestRequestPath := RRtype . "_latestRequestPath"
	RREmail := RRtype . "_Email"
	
	if (CitrixSession == True)
	{
		MB("Currently cannot send via a Citrix session. Stopping automation")
		;NHSMailOpen()
		;Email_NHS_Mail(%RREmail%, "", RRMessage, "Please find attached a " . RRMessage, %RRLatestRequestPath%)
	}
	else
	{
		if EmailOutlook(%RREmail%,, RRMessage, "Please find attached a " . RRMessage, %RRLatestRequestPath%)
		{
			UpdateProgressBar(20,, PBMainString . "Complete.`n")
			
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




RRGetExtraInfo(MRN, RRtype)
{
	global 

	%RRtype%_details := {}

	if (%RRtype%_GUI(MRN) != "OK")
	{
		return False
	}

	if !FileExist(%RRtype%_details.signature)
	{
		MB("No signature found for selected doctor. Ending request/referral")
		return False
	}

	return True
}




Sleepstation_GUI(MRN)
{
	global
	
	local Title := "Sleepstation referral details"
	local GUIStart := 200
	local GUIWidth := 300
	local CurrentUser := StrReplace(A_UserName, ".", " ")
	local CurrentUserPostionInListDoctor := InStr(doctorListConstant, CurrentUser)  ; 0  if not found
	local yAdditiveStart := 10
	local yAdditive := yAdditiveStart
	local checkBoxStart := 10
	local checkBoxWidth := 180
	local doctorList


	if (CurrentUserPostionInListDoctor == 0)
		doctorList := doctorListConstant
	else
		doctorList := SubStr(doctorListConstant, 1, CurrentUserPostionInListDoctor + StrLen(CurrentUser)) . "|" . SubStr(doctorListConstant, CurrentUserPostionInListDoctor + StrLen(CurrentUser) + 1)
		
		
	Gui, SSRef:Font, s12
	Gui, SSRef:Color, %dialogueColour%
	Gui, SSRef:Add, Text, x10 y%yAdditive%, % "Please fill in the below details for " . patientDetails.name . ", MRN" . MRN
	yAdditive := yAdditive + 40

	Gui, SSRef:Add, Text, x10 y%yAdditive%, Requesting doctor:
	Gui, SSRef:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vSleepstationDoctorName, %doctorList%
	yAdditive := yAdditive + 40
	
	Gui, SSRef:Add, Text, x10 y%yAdditive%, Doctor email:
	Gui, SSRef:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vSleepstationDoctorEmail, % "@nhs.net"
	yAdditive := yAdditive + 40
	
	Gui, SSRef:Add, Text, x10 y%yAdditive%, Clinical details:
	Gui, SSRef:Add, Edit, x%GUIStart% y%yAdditive% H70 W%GUIWidth% vSleepstationClinicalDetails, % "Insomnia "
	yAdditive := yAdditive + 77

	Gui, SSRef:Add, Button, x400 y%yAdditive% default gSSRefOK,  &OK
	Gui, SSRef:Add, Button, x450 y%yAdditive% gSSRefClose, &Cancel

	Gui, SSRef:Show,, %Title%
	Gui, SSRef:+AlwaysOnTop
	WinWaitClose, %Title%
	return Sleepstation_buttonPressed
}




SSRefOK()
{
	global
	
	local errorString := ""
		
	Gui SSRef:Submit, Nohide
	Sleepstation_buttonPressed := "OK"


	if (SleepstationDoctorName == "")
		errorString := errorString . "- No requesting doctor was chosen.`n"
		
	if (SleepstationDoctorEmail == "" or SleepstationDoctorEmail == "@nhs.net" or !instr(SleepstationDoctorEmail, "@"))
		errorString := errorString . "- No requesting doctor email provided.`n"
		
	if (SleepstationClinicalDetails = "" or SleepstationClinicalDetails == "Insomnia ")
		errorString := errorString . "- No clinical details were provided.`n"
		
	if (errorString != "")
	{
		MB("Please correct the below errors and submit again:`n" . errorString)
		return
	}

	SleepStation_details.doctorName := SleepstationDoctorName
	SleepStation_details.signature := CurrentDirectory . "Signatures\" .  SleepstationDoctorName . ".jpg"
	SleepStation_details.doctorEmail := SleepstationDoctorEmail
	SleepStation_details.clinicalDetails := SleepstationClinicalDetails
	Gui, SSRef:Destroy
	return
}




SSRefGuiClose()
{
	SSRefClose()
}
SSRefClose()
{
	global
	
	Sleepstation_buttonPressed := "close"
	Gui, SSRef:Destroy
	return
}




create_Sleepstation_request()
{
	global
	
	Loop, 5
	{
		local SavePath := Settings.RequestsFolder . "Sleepstation request MRN" . MRN . "_"
		local Today
		FormatTime, Today,, dd/MM/yy
		local signatureWidth := signatureHeight * GraphicsDimensions(SleepStation_details.signature)

		try
		{
			Loop
			{
				SavePathTemp := SavePath . A_Index . ".pdf"

				if(FileExist(SavePathTemp) == "")
				{
					SavePath := SavePathTemp
					Sleepstation_latestRequestPath := SavePathTemp
					break
				}
			}

			oWord := ComObjCreate("Word.Application")   ; create MS Word object
			oWord.Visible := False
			oWord.Documents.Open(Sleepstation_template_path)
			
			WriteAtBookmark("Date", Today)
			oWord.ActiveDocument.Bookmarks("Doctor").Select
			picObj := oWord.ActiveDocument.InLineShapes.AddPicture(Sleepstation_details.signature)
			picObj.Height := signatureHeight 
			picObj.Width := signatureWidth
			WriteAtBookmark("Doctor", "Dr " . Sleepstation_details.doctorName . ", ")
			WriteAtBookmark("DoctorEmail", Sleepstation_details.doctorEmail)
			
			WriteAtBookmark("Name", PatientDetails.name)
			WriteAtBookmark("DOB", PatientDetails.DOB)
			WriteAtBookmark("NHSNumber", PatientDetails.NHSNumber)
			WriteAtBookmark("address", PatientDetails.postCode)
			WriteAtBookmark("address", PatientDetails.address4 . ", ")
			WriteAtBookmark("address", PatientDetails.address3 . ", ")
			WriteAtBookmark("address", PatientDetails.address2 . ", ")
			WriteAtBookmark("address", PatientDetails.address1 . ", ")
			WriteAtBookmark("telephoneNumber", PatientDetails.MobNo . ", " . PatientDetails.TeleNo)
			WriteAtBookmark("email", PatientDetails.email)
			WriteAtBookmark("ClinicalDetails", Sleepstation_details.ClinicalDetails)
			
			UpdateProgressBar(80, "absolute", PBMainString . "Saving referral to requests folder")		
			oWord.ActiveDocument.SaveAs(SavePath, 17) ; 17 is PDF format
			oWord.ActiveDocument.close(False)
			return True
		}
		catch, err
		{
			try
				oWord.ActiveDocument.close(False)
				
			if (A_index >= 5)
			{
				errorhandler(err, "HLSG request", "Microsoft Word")
				return False
			}
			else
			{
				LogUpdate("Error during HLSG request [counter: " . A_index . "]")
			}
		}
	}
	
	return False
}



