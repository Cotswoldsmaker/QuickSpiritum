;*****************************************
; Sleep servie personalised patient videos
;*****************************************

; *********
; Variables
; *********

youtubeURL := "https://www.youtube.com/watch_videos?video_ids="
introID := "9pjIiY2Ekqg"
obeseID := "ClYjSv93d3k"
morbidlyObeseID := "lx84USt7Oxo"
smokingID := "kNXWi97Fax8"
insomniaID := "7JnItBSmEro"
sleepHygieneID := "UG08RTvvGvs"
alcoholID := "WU5tD2CUvNE"
sedativesID := "qcrbUoW7qzs"
DVLAID := "vXo7Z6Wpbm8"




; *********
; Functions
; *********

sleepQuestionnaire(MRNTemp)
{
	global
	MRN := MRNTemp
	sleepQuestionnaire_GUI()
	return True
}


sleepQuestionnaire_GUI()
{
	global
	local Title := "Patient sleep questionnaire"
	local LabelWidth := 450
	local LabelWidth := 450
	local FieldStart := 500
	local FieldWidth := 300
	local yAdditiveStart := 10
	local yAdditive := yAdditiveStart
	local checkBoxStart := 10
	local checkBoxWidth := 500 

	Gui, SQues:Font, s12
	Gui, SQues:Color, %dialogueColour%
	Gui, SQues:Add, Text, x10 y%yAdditive%, % "Please fill in the below details to help us determine how best to treat your disturbed sleep."
	yAdditive := yAdditive + 40
	
	Gui, SQues:Add, Text, x10 y%yAdditive%, % "Full name:"
	Gui, SQues:Add, Edit, x%FieldStart% y%yAdditive% W%FieldWidth% vSQuesFullName
	yAdditive := yAdditive + 45
	
	Gui, SQues:Add, Text, x10 y%yAdditive% W%LabelWidth%, % "Mobile number (if happy to receive results via SMS messages on a smart phone):"
	Gui, SQues:Add, Edit, x%FieldStart% y%yAdditive% W%FieldWidth% vSQuesMobileNumber
	yAdditive := yAdditive + 45
	
	Gui, SQues:Add, Text, x10 y%yAdditive% W%LabelWidth%, % "What time on average do you go to bed (please enter in 24-hour format)?"
	Gui, SQues:Add, Edit, x%FieldStart% y%yAdditive% W%FieldWidth% vSQuesBedTime
	yAdditive := yAdditive + 45
	
	Gui, SQues:Add, Text, x10 y%yAdditive% W%LabelWidth%, % "What time on average do you normally wake up (please enter in 24-hour format)?"
	Gui, SQues:Add, Edit, x%FieldStart% y%yAdditive% W%FieldWidth% vSQuesWakeupTime
	yAdditive := yAdditive + 45
	
	Gui, SQues:Add, Checkbox, x%checkBoxStart% y%yAdditive% vSQuesInsomnia +Right w%checkBoxWidth%, % "Does it take longer than an hour to fall asleep, or do you wake up in the night and find it hard to fall back to sleep?"
	yAdditive += 65
	
	Gui, SQues:Add, Checkbox, x%checkBoxStart% y%yAdditive% vSQuesSmokerVaper +Right w%checkBoxWidth%, % "Have you smoked or vaped in the last month?"
	yAdditive := yAdditive + 45
	
	Gui, SQues:Add, Text, x10 y%yAdditive%, % "How many units of alcohol do you drink in a week?"
	Gui, SQues:Add, DropDownList, x%FieldStart% y%yAdditive% W%FieldWidth% vSQuesEtOH, % "0-5|5-10|10-15|15-30|30+"
	yAdditive := yAdditive + 45
	
	Gui, SQues:Add, Checkbox, x%checkBoxStart% y%yAdditive% vSQuesDriver +Right w%checkBoxWidth%, % "Do you hold a driving licence?"
	yAdditive := yAdditive + 45
	
	Gui, SQues:Add, Checkbox, x%checkBoxStart% y%yAdditive% vSQuesDSleepHygiene +Right w%checkBoxWidth%, % "Do you watch TV before bed or drink caffeinated drinks 6 hours before bed?"
	yAdditive := yAdditive +50
	
	Gui, SQues:Add, Text, x10 y%yAdditive% W%LabelWidth%, % "Please list the names (not doses or times) of all the medications you take:"
	Gui, SQues:Add, Edit, x%FieldStart% y%yAdditive% W%FieldWidth% vSQuesMedications
	yAdditive := yAdditive + 45
	
	Gui, SQues:Add, Button, x700 y%yAdditive% default gSQuesOK,  &OK
	Gui, SQues:Add, Button, x750 y%yAdditive% gSQuesClose, &Cancel

	Gui, SQues:Show,, %Title%
	Gui, SQues:+AlwaysOnTop
	WinWaitClose, %Title%
	return SQues_OKButtonPressed
}




SQuesOK()
{
	global
	local errorString := ""
		
	Gui SQues:Submit, Nohide
	SQues_OKButtonPressed := "yes"
	
	; !!! code to check all fields filled in (or not)
		
	Gui, SQues:Destroy
	sleepStudy_GUI()
	return
}




SQuesGuiClose()
{
	SQuesClose()
}
SQuesClose()
{
	global
	
	Sleepstation_buttonPressed := "close"
	Gui, SQues:Destroy
	return
}





sleepStudy_GUI()
{
	global
	local Title := "Sleep study results"
	local GUIStart := 500
	local GUIWidth := 300
	local yAdditiveStart := 10
	local yAdditive := yAdditiveStart
	local checkBoxStart := 10
	local checkBoxWidth := 500

	Gui, SS:Font, s12
	Gui, SS:Color, %dialogueColour%
	Gui, SS:Add, Text, x10 y%yAdditive%, % "Please fill in results from the patient's sleep study"
	yAdditive := yAdditive + 40
	
	Gui, SS:Add, Text, x10 y%yAdditive%, % "Height (cm):"
	Gui, SS:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vSSheight
	yAdditive := yAdditive + 40
	
	Gui, SS:Add, Text, x10 y%yAdditive%, % "Weight (kg)"
	Gui, SS:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vSSWeight
	yAdditive := yAdditive + 40
	
	Gui, SS:Add, Text, x10 y%yAdditive%, % "Epworth sleepines score"
	Gui, SS:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vSSESS
	yAdditive := yAdditive + 40
	
	Gui, SS:Add, Text, x10 y%yAdditive%, % "AHI"
	Gui, SS:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vSSAHI
	yAdditive := yAdditive + 40
	
	Gui, SS:Add, Text, x10 y%yAdditive%, % "3% ODI"
	Gui, SS:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vSS3ODI
	yAdditive := yAdditive + 40
	
	Gui, SS:Add, Button, x700 y%yAdditive% default gSSOK,  &OK
	Gui, SS:Add, Button, x750 y%yAdditive% gSSClose, &Cancel

	Gui, SS:Show,, %Title%
	Gui, SS:+AlwaysOnTop
	WinWaitClose, %Title%
	return SS_OKButtonPressed
}




SSOK()
{
	global
	local errorString := ""
	local fullURL := ""
	local message := ""
	local timeInBed := ""
	local BMI := ""
	
	Gui SS:Submit, Nohide
	SS_OKButtonPressed := "yes"
	
	; !!! code to check all fields filled in (or not)
	
	Gui, SS:Destroy
	
	fullURL := youtubeURL . introID

/*
*introID := "9pjIiY2Ekqg"
*obeseID := "ClYjSv93d3k"
*morbidlyObeseID := "lx84USt7Oxo"
*smokingID := "kNXWi97Fax8"
*insomniaID := "7JnItBSmEro"
*sleepHygieneID := "UG08RTvvGvs"
*alcoholID := "WU5tD2CUvNE"
*sedativesID := "qcrbUoW7qzs"
*DVLAID := "vXo7Z6Wpbm8"
*/

	if (timeDiff(SQuesBedTime, SQuesWakeupTime) < 6)
		fullURL .= "," . sleepHygieneID
	
	if SQuesInsomnia
		fullURL .= "," . insomniaID

	if SQuesSmokerVaper
		fullURL .= "," . smokingID
	
	if (SQuesEtOH > 20)
		fullURL .= "," . alcoholID
		
	if inStr(SQuesMedications, "bisoprolol")
		fullURL .= "," . sedativesID
	
	SSheight := SSHeight / 100
	
	BMI :=	SSWeight / (SSheight * SSheight)
	msgbox, % BMI

	if (BMI >= 30 AND BMI < 40)
		fullURL .= "," . obeseID
	else if (BMI < 40)
		fullURL .= "," . morbidlyObeseID
	
	if SQuesDriver
		fullURL .= "," . DVLAID
		
	
	message := "TEST TEST ONLY! Please click the below link to watch a personalised information video following your sleep study: " . fullURL
	msgbox, % message
	;clipboard := fullURL
	
	if !MobileNumberCheck(SQuesMobileNumber)
		return False
	
	MobileNumber := "+44" . SubStr(SQuesMobileNumber, 2)
	variables := "SMS;" . API_key . ";" . SMS_template_ID . ";" .  MobileNumber . ";" .  message
	sendMessageGUKN("SMS message", variables)
	msgbox, % "Message sent"
	return
}




SSGuiClose()
{
	SSClose()
}
SSClose()
{
	global
	
	Sleepstation_buttonPressed := "close"
	Gui, SS:Destroy
	return
}



