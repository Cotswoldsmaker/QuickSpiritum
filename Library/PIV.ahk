; ***********************************************************************
; Patient information vidoe functionality. Uses GOV.uk Notify to send
; messages
; ***********************************************************************


; *********
; Variables
; *********

PIV_Excel_path := "S:\Thoracic\Spiritum\Patient Information Video Messages.xlsx"
PIV_messages_array := {}

GUKN_type := ""

GUKN_SMS_messages := {}
GUKN_SMS_message_type := ""
GUKN_SMS_message_types := ""
GUKN_mobileNumber := ""

GUKN_email_messages := {}
GUKN_email_address := ""
GUKN_email_subject := ""
GUKN_email_body := ""

GUKN_letter_messages := {}
GUKN_letter_to := ""
GUKN_letter_address := ""
GUKN_letter_from := ""
GUKN_letter_header := ""
GUKN_letter_body := ""

PIV_read_Excel()





; *********
; Functions
; *********

PIV_read_Excel()
{
	global
	local XL := ComObjCreate("Excel.Application")
	local title := ""
	local SMS_message := ""
	local email_message := ""
	local letter_message := ""
	local URL := ""
	
	
	XL.Workbooks.Open(PIV_Excel_path)

	For c In xl.ActiveSheet.UsedRange.cells
		PIV_messages_array[c.address(0,0)] := c.Value

	XL.quit
	XL := ""
	
	
	Loop
	{
		title := PIV_messages_array["b" . A_index]
		SMS_message := PIV_messages_array["c" . A_index]
		email_message := PIV_messages_array["d" . A_index]
		letter_message := PIV_messages_array["e" . A_index]
		URL := PIV_messages_array["f" . A_index]
		
		if (title == "")
			break
			
		if (A_index > 1)
		{
			GUKN_message_types .= title . "|"
			GUKN_SMS_messages[title] := SMS_message . " " . URL
			GUKN_email_messages[title] := email_message . " " . URL
			GUKN_letter_messages[title] := letter_message . " " . URL
		}
	}
	
	return
}




GOV_UK_Notify()
{
	global
	local title := "Notification type"
	local GUIStart := 60
	local GUIWidth := 100
	local H := 0 ; Height of field
	local yAdditive := 10
	local GUI_name := "GUKN_main"
	
	Gui, %GUI_name%:Font, s12
	Gui, %GUI_name%:Color, % dialogueColour
	
	H := 40
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % patientDetails.name . ", MRN" . MRN . ", " . patientDetails.DOB
	yAdditive += H + S
		
	H := 45
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Type:
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_type, SMS||email|letter
	yAdditive += H + S
	
	Gui, %GUI_name%:Add, Button, x150 y%yAdditive% default gGUKN_mainOK,  &OK
	Gui, %GUI_name%:Add, Button, x200 y%yAdditive% gGUKN_mainClose,  &Cancel
	Gui, %GUI_name%:Show,, % title
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, % title
	return
}




GUKN_mainOK()
{
	global
	
	
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
	global
	local title := "Send SMS message"
	local GUIStart := 130
	local GUIWidth := 500
	local H := 0 ; Height of field
	local yAdditive := 10
	local GUI_name := "GUKN_SMS"
	

	Gui, %GUI_name%:Font, s12
	Gui, %GUI_name%:Color, % dialogueColour
	
	H := 40
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % patientDetails.name . ", MRN" . MRN . ", " . patientDetails.DOB
	yAdditive := yAdditive + 40
		
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Mobile number:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_mobileNumber, % PatientDetails.MobNo ;mobile1
	yAdditive := yAdditive + 35
		
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Type:"
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_SMS_message_type gGUKN_SMS_message_type_clicked, % GUKN_message_types
	yAdditive := yAdditive + 35
		
	H := 150
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Message:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H%H% vGUKN_SMS_message
	yAdditive := yAdditive + 170
	
	Gui, %GUI_name%:Add, Button, x500 y%yAdditive% default gGUKN_SMSOK,  &OK
	Gui, %GUI_name%:Add, Button, x550 y%yAdditive% gGUKN_SMSClose,  &Cancel
	Gui, %GUI_name%:Show,, % title
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, % title

	return
}




GUKN_SMS_message_type_clicked()
{
	global
	local GUI_name := "GUKN_SMS"
	
	Gui %GUI_name%:Submit, NoHide
	GuiControl, %GUI_name%: , GUKN_SMS_message, % GUKN_SMS_messages[GUKN_SMS_message_type]
	return
}




GUKN_SMSOK()
{
	global
	local pass := False
	local variables := ""
	
	Gui, GUKN_SMS:Submit, NoHide
	Gui, GUKN_SMS:Destroy
	
	if !MobileNumberCheck(GUKN_mobileNumber)
	{
		MB("Error with mobile number. Please try again")
		return "Fail"
	}

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
	global
	local title := "Send email"
	local GUIStart := 120
	local GUIWidth := 500
	local H := 0 ; Height of field
	local yAdditive := 10
	local GUI_name := "GUKN_email"


	Gui, %GUI_name%:Font, s12
	Gui, %GUI_name%:Color, % dialogueColour
	
	H := 40 ; height not actually used here
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % patientDetails.name . ", MRN" . MRN . ", " . patientDetails.DOB
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Email:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H%H% vGUKN_email_address, % PatientDetails.email
	yAdditive += H + S
	
	H := 30 ; height not actually used here
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Type:"
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_email_type gGUKN_email_type_clicked, % GUKN_message_types
	yAdditive += H + S
	
	H := 100
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Subject:"
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H%H% vGUKN_email_subject, % "Message from GHNHSFT Respiratory Department"
	yAdditive += H + S
	
	H := 200
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Body:"
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H%H% vGUKN_email_body
	yAdditive += H + S
	
	Gui, %GUI_name%:Add, Button, x500 y%yAdditive% default gGUKN_emailOK,  &OK
	Gui, %GUI_name%:Add, Button, x550 y%yAdditive% gGUKN_emailClose,  &Cancel
	Gui, %GUI_name%:Show,, % title
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, % title
	return
}




GUKN_email_type_clicked()
{
	global
	local GUI_name := "GUKN_email"
	
	
	Gui %GUI_name%:Submit, NoHide
	GuiControl, %GUI_name%: , GUKN_email_body, % GUKN_email_messages[GUKN_email_type]
	return
}



GUKN_emailOK()
{
	global
	local variables := ""
	
	
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
	local GUIWidth := 500
	local H := 0 ; Height of field
	local yAdditive := 10
	local GUI_name := "GUKN_letter"
	local address := ""

		
	Loop, 4
	{
		line := "address" . A_index
		
		if (PatientDetails[line] != "")
			address .= PatientDetails[line] . "`r`n"
		else
			break
	}
	
	address .= PatientDetails.postCode
		
	Gui, %GUI_name%:Font, s12
	Gui, %GUI_name%:Color, % dialogueColour
	
	H := 40
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive% H%H%, % patientDetails.name . ", MRN" . MRN . ", " . patientDetails.DOB
	yAdditive += H + S
	
	H := 30 ; not actually used height here
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Type:"
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_letter_type gGUKN_letter_type_clicked, % GUKN_message_types
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, To:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H%H% vGUKN_letter_to, % PatientDetails.name
	yAdditive += H + S
	
	H := 120
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, Address:
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H%H% vGUKN_letter_address, % address
	yAdditive += H + S

	H := 30 ; not actually used in height here
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "From:"
	Gui, %GUI_name%:Add, DropDownList, x%GUIStart% y%yAdditive% W%GUIWidth% vGUKN_letter_from, % doctorList
	yAdditive += H + S
	
	H := 30
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Header:"
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H%H% vGUKN_letter_header, % "GHNHST Respiratory Department"
	yAdditive += H + S
	
	H := 110
	Gui, %GUI_name%:Add, Text, x10 y%yAdditive%, % "Letter body:"
	Gui, %GUI_name%:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% H%H% vGUKN_letter_body
	yAdditive += H + S
	
	Gui, %GUI_name%:Add, Button, x500 y%yAdditive% default gGUKN_letterOK,  &OK
	Gui, %GUI_name%:Add, Button, x550 y%yAdditive% gGUKN_letterClose,  &Cancel
	Gui, %GUI_name%:Show,, % title
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, % title
	return
}




GUKN_letter_type_clicked()
{
	global
	local GUI_name := "GUKN_letter"
	
	
	Gui %GUI_name%:Submit, NoHide
	GuiControl, %GUI_name%: , GUKN_letter_body, % GUKN_letter_messages[GUKN_letter_type]
	return
}




GUKN_letterOK()
{
	global
	local addressConverted := ""
	local variables := ""
	
	
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




sendMessageGUKN(method, variables, WaitToClose := True)
{
	global
	local outcome := "timeOut"
	local pythonMessages := ""
	local returnResult := ""
	local returnResultSplit := ""
	

	CloseProgressBar()
	PBMainString := "Sending " . method . ":`n"
	CreateProgressBar()
	UpdateProgressBar(5, "absolute", PBMainString . ".")


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
	
	UpdateProgressBar(100, "absolute", PBMainString . "....")
	
	if (instr(outcome, "pass") = 1)
	{
		UpdateProgressBar(,, PBMainString . method . " sent", WaitToClose)
		return True
	}
	else if (outCome = "timeOut")
	{
		UpdateProgressBar(,, PBMainString . "The GOV.uk Notification program (written in python) did not start and hence the " . method . " has not been sent!", WaitToClose)
		return False
	}
	else if (outcome = "Wrong initial arguement")
	{
		UpdateProgressBar(,, PBMainString . "Wrong initial arguement provided", WaitToClose)
		return False
	}
	else if (outcome = "fail")
	{
		UpdateProgressBar(,, PBMainString . method . " - request failed!", WaitToClose)
		return False
	}
	else if (outcome = "wrong method")
	{
		UpdateProgressBar(,, PBMainString . method . " - wrong method provided to python script", WaitToClose)
		return False
	}
	else
	{
		UpdateProgressBar(,, PBMainString . "AHK error with " . method . " function [" . outcome . "]!", WaitToClose)
		return False
	}	
	
	CloseProgressBar()
	return False
}




; IPC method to send PIVs
send_GUKN(rawData)
{
	global
	local pass := False
	local mobileNumber := ""
	local DS := ""
	local method := ""
	local dataRetrieval := ""
	local mobileNumber := ""
	local GUKN_SMS_message := ""
	local variables := ""
	local GUKN_email_address := ""
	local GUKN_email_subject := ""
	local GUKN_email_body := ""
	local GUKN_letter_to := ""
	local GUKN_letter_address := ""
	local addressConverted := ""
	local addressConverted := ""
	local GUKN_letter_from := ""
	local GUKN_letter_header := ""
	local GUKN_letter_body := ""
	
	
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
		{
			MB("Error with mobile number. Please try again")
			GoTo Fail
		}
		
		mobileNumber := "+44" . SubStr(mobileNumber, 2)
		
		GUKN_SMS_message := DS[5]
		variables := "SMS;" . API_key . ";" . SMS_template_ID . ";" .  MobileNumber . ";" .  GUKN_SMS_message
		UpdateProgressBar(55, "absolute", PBMainString . "Sending request to GOV.uk Notify...")
		sendMessageGUKN("SMS message", variables, False)
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
		sendMessageGUKN("email", variables, False)
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
		sendMessageGUKN("Letter", variables, False)
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
	
	return False

}



