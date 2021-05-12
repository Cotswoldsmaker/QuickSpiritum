; *********************************
; Quality improvement project - QIP
; *********************************

; *********
; Variables
; *********

timerStart := 0
MRN := ""
MRN_GUI := ""
elapsedTime := ""
QIP_log_path := Desktop . "QIP_log.txt"
QIP_infoButtonPressed := False
QIP_answer := ""
QIPX := ""
QIPY := ""
QIP_title := ""
counter := 0
IF_ExitBeforeSelectAndZoom := False


; *********
; Functions
; *********	

QIP_main()
{
	global
	
	counter := 0
	
	; Set emails to email TestEmail only during this QIP run
	if !emailTest
	{
		local PET_CT_EmailStore := PET_CT_Email
		PET_CT_Email := TestEmail
		local PET_CT_Email_ccStore := PET_CT_Email_cc
		PET_CT_Email_cc := ""
		local PFT_EmailStore := PFT_Email
		PFT_Email := TestEmail
		local Bronch_EmailStore := Bronch_Email
		Bronch_Email := TestEmail
		local Bronch_Email_ccStore := Bronch_Email_cc
		Bronch_Email_cc := ""
		local HLSG_EmailStore := HLSG_Email
		HLSG_Email := TestEmail
		local SleepStation_EmailStore := SleepStation_Email
		SleepStation_Email := TestEmail
	}
	
	local username := getClinicianDetails(clinicianUsername, A_UserName, clinicianActualName) . " (" . A_UserName . ")"
	local result
	local title := "Quick Spiritum - Quality Improvement Project (QS-QIP)"
	local message := username . " started on "
	
	if CitrixSession
		message .= "Citrix machine."
	else
		message .= "Desktop."
	
	QIP_log_update(message)
	
	message := "Many thanks for agreeing to help with this quality improvement project to research "
			 . "how Quick Spiritum (QS) can speed up your daily clinical work flow. This should only take 20-30 minutes "
			 . "of your time. You will be asked to undertake some routine clinical tasks, first manually "
			 . "and then with automation via Quick Spiritum.`n`n"
			 . "Tasks will include finding:`n`n"
			 . "- A patient's telephone number via Sunrise`n"
			 . "- Clinician's name for the most recent clinical episode on Trakcare`n"
			 . "- Most recent letter on InfoFlex`n"
			 . "- Date of last chest X-ray via PACS`n"
			 . "- Last serum potassium level via PAS`n"
			 . "- Date of last lung function test`n"
			 . "- Date of last chest X-ray via ICE`n`n"
			 . "You will also be asked to send requests for the below (but only using QS, not manually):`n`n"
			 . "- PET-CT`n"
			 . "- Lung function`n"
			 . "- Bronchoscopy`n`n"
			 . "The timings and results you enter will be emailed to " . username1 . " when you have completed the tasks.`n`n"
			 . "If you make a mistake or feel a task took longer than it would normally you can redo that individual task.`n`n"
			 . "It is best to run this program on a dual screen computer and then place the QS dialogue box on the less active screen.`n`n"
			 . "Please make sure QS already has your credentials for the different clinical systems (if you find credentials are missing during a task, then please complete the credentials, allow QS to finish its task, click OK on the QI window and then repeat the task).`n`n"
			 . "If you are not too familiar with QS's functionality, it would be beneficial to trying playing around with the F-keys (F1 for help) to get a feel of its functionality before running this QIP program.`n`n"
			 . "Click OK when you are happy to start."

	if !QIP_info(title, message)
		GoTo CleanUp
	
	
	; Manual tasks
	if !QIP_info(title, "You will now be asked to undertake some tasks manually.",, 450)
		GoTo CleanUp
			
	runningStatus() ; Stop the ability to use QS to do searches during manual part

	if !QIP_task("Manually", "Please close all programs and then manually find the patient's telephone number via Sunrise. You will see this in the top section of the Sunrise window once you search for the patient and click on an admission episode (under 'Phone and Email').", "Sunrise startup", MRN1)
		goto CleanUp
		
	if !QIP_task("Manually", "Keeping Sunrise open, manually find the patient's telephone number via Sunrise.", "Sunrise opened", MRN2)
		goto CleanUp
		
	if !QIP_task("Manually", "Please close all programs and then manually find the clinican's name for the last clinical episode on Trakcare.", "Trakcare startup", MRN3)
		goto CleanUp
		
	if !QIP_task("Manually", "Keeping Trakcare open, manually find the clinician's name for the last clinical episode on Trakcare.", "Trakcare opened", MRN4)
		goto CleanUp
		
	if !QIP_task("Manually", "Please close all programs and then manually find the date for the last letter via InfoFlex. If there is no letter then please enter 'none' as a result.", "InfoFlex startup", MRN5)
		goto CleanUp
		
	if !QIP_task("Manually", "Keeping InfoFlex open, manually find the date of the last letter via InfoFlex. If there is no letter then please enter 'none' as a result.", "InfoFlex opened", MRN6)
		goto CleanUp
		
	if !QIP_task("Manually", "Please close all programs and then manually find the date for the last chest X-ray via PACS.", "PACS startup", MRN7)
		goto CleanUp
		
	if !QIP_task("Manually", "Keeping PACS open, manually find the date of the last chest X-ray via PACS", "PACS opened", MRN8)
		goto CleanUp
		
	if !QIP_task("Manually", "Please close all programs and then manually find the date for the latest potassium (K+) via PAS through Trakcare.", "PAS startup", MRN9)
		goto CleanUp
		
	if !QIP_task("Manually", "Close PAS, but keep Trakcare open. Manually find the date of the latest potassium (K+) via PACS through Trakcare.", "PAS opened", MRN10)
		goto CleanUp
		
	if !QIP_task("Manually", "Please close all programs and then manually find the date for the last lung function test.", "Lung function test", MRN11)
		goto CleanUp
		
	if !QIP_task("Manually", "Please close all programs and then manually find the date for the last chest X-ray via ICE.", "ICE startup", MRN12)
		goto CleanUp
		
	if !QIP_task("Manually", "Keeping ICE open, manually find the date of the last chest X-ray via ICE.", "ICE opened", MRN13)
		goto CleanUp
		
	runningStatus("done")
	
	
	; Automated tasks
	if !QIP_info(title, "You will now be asked to undertake some tasks using automation via Quick Spiritum.`n`n"
					  . "Please make sure Quick Spiritum already has your credentials for the different clinical systems (if you find credentials are missing during a task, then please complete the credentials, allow QS to finish its task, click OK on the QI window and then repeat the task).`n`n"
					  . "Remember, you can highlight the MRN on the screen and then press the corresponding F-key to start QS automation. You can press F1 at any time to see a list of function keys.")
		GoTo CleanUp
		
	if !QIP_task("Automated", "Please close all programs and then use Quick Spiritum to find the patient's telephone number via Sunrise. You will see this at the top of the Sunrise screen (under 'Phone and Email').", "Sunrise startup", MRN14)
		goto CleanUp
		
	if !QIP_task("Automated", "Keeping Sunrise open, use Quick Spiritum to find the patient's telephone number via Sunrise.", "Sunrise opened", MRN15)
		goto CleanUp
		
	if !QIP_task("Automated", "Please close all programs and then use Quick Spiritum to find the clinician's name for the last clinical episode on Trakcare.", "Trakcare startup", MRN16)
		goto CleanUp
		
	if !QIP_task("Automated", "Keeping Trakcare open, use Quick Spiritum to find the clinician's name for the last clinical episode on Trakcare.", "Trakcare opened", MRN17)
		goto CleanUp
	
	IF_ExitBeforeSelectAndZoom := True
	
	if !QIP_task("Automated", "Please close all programs and then use Quick Spiritum to find the date for the last letter via InfoFlex. If there is no letter then please enter 'none' as a result.", "InfoFlex startup", MRN18)
		goto CleanUp
		
	if !QIP_task("Automated", "Keeping InfoFlex open, use Quick Spiritum to find the date of the last letter last via InfoFlex. If there is no letter then please enter 'none' as a result.", "InfoFlex opened", MRN19)
		goto CleanUp
	
	if !QIP_task("Automated", "Please close all programs and then use Quick Spiritum to find the date for the last chest X-ray via PACS.", "PACS startup", MRN20)
		goto CleanUp
		
	if !QIP_task("Automated", "Keeping PACS open, use Quick Spiritum to find the date of the last chest X-ray via PACS.", "PACS opened", MRN21)
		goto CleanUp
			
	if !QIP_task("Automated", "Please close all programs and then use Quick Spiritum to find the date for the latest serum potassium (K+) via PAS through Trakcare (please note Quick Spiritum does not control the PAS program, so that last part of the search is done manually).", "PAS startup", MRN22)
		goto CleanUp
		
	if !QIP_task("Automated", "Close PAS, but keep Trakcare open. Use Quick Spiritum to find the date of the latest serum potassium (K+) via PACS through Trakcare.", "PAS opened", MRN23)
		goto CleanUp
		
	if !QIP_task("Automated", "Please close all programs and then use Quick Spiritum to find the date for the last lung function test.", "Lung function test", MRN24)
		goto CleanUp
		
	if !QIP_task("Automated", "Please close all programs and then use Quick Spiritum to find the date for the last chest X-ray via ICE.", "ICE startup", MRN25)
		goto CleanUp
	
	if !QIP_task("Automated", "Keeping ICE open, use Quick Spiritum to find the date of the last chest X-ray via ICE.", "ICE opened", MRN26)
		goto CleanUp	
	

	; Requests
	if !QIP_info(title, "You will finally be asked to undertake some requests using Quick Spiritum only.`n`n" 
					  . "All requests will only be emailed to " . username1 . ". They will NOT be sent to actual request inboxes.",, 600)
		GoTo CleanUp
	
	if !QIP_task("Automated", "PET-CT", "PET-CT request", MRN27, True)
		goto CleanUp
		
	if !QIP_task("Automated", "lung function", "Lung function request", MRN28, True)
		goto CleanUp
		
	if !QIP_task("Automated", "bronchoscopy", "Bronchoscopy request", MRN29, True)
		goto CleanUp
		
	
	; Finished
	title := "Quick Spiritum QIP Complete"
	message := "You have now finished. Many thanks again for taking the time to help with this quality improvement project.`n`n"
			 . "If you are on a desktop computer, the results will now be emailed across. If on a Citrix computer, please find the QIP_log.txt file on your desktop and email this to " . userEmail1 . " via the NHSmail web app.`n`n"
			 . "Regards`n`n" . username1
			
	QIP_info(title, message, False, 500)
	
	if !CitrixSession
		EmailOutlook(TestEmail,, "Quick Spiritum QIP Results for " . username, "QIP timings for " . username . " attached", QIP_log_path)
	
	QIP_log_update("Completed successfully")
	
CleanUp:
	if !emailTest
	{
		PET_CT_Email := PET_CT_EmailStore
		PET_CT_Email_cc := PET_CT_Email_ccStore
		PFT_Email := PFT_EmailStore
		Bronch_Email := Bronch_EmailStore
		Bronch_Email_cc := Bronch_Email_ccStore
		HLSG_Email := HLSG_EmailStore
		SleepStation_Email := SleepStation_EmailStore
	}
	
	runningStatus("done")
	IF_ExitBeforeSelectAndZoom := False
	
	return
}




QIP_task(ManualAutomated, task, program, MRNStore, Request := False)
{
	global
	counter += 1
	local title := "QIP - " . ManualAutomated
	local message
	local messageRepeat := "Are you happy with this run? If not press 'Repeat'."
	local InfoResult := ""
	
	
	if Request
		message := counter . ". Please close all programs, including Trakcare (which may be minimised), except Outlook and then use Quick Spiritum to create a " . task . " request, using made up clinical details. Click OK to show the MRN and start the timer. When complete, press OK again (no need to enter anything in the 'Search result')."
	else
		message := counter . ". " . task . " Click OK to show the MRN and start the timer. When found, enter the result in the 'Search result' field and press OK again."
	
	while True
	{
		MRN := MRNStore
		result := QIP_sub(title, message)

		if (result == "cancelled")
			return False
			
		QIP_log_update(counter . "," . ManualAutomated . "," . program . ",MRN" . MRN . "," . result . "," . elapsedTime)
	
		InfoResult := QIP_info(title, messageRepeat, "repeat", 350)
		
		if (InfoResult == "OK")
			return True
		else if !InfoResult
			return False
	}
		
	return True
}




QIP_info(title, message, secondButton := "cancel", width := 800 )
{
	global
	
	timerStart := 0
	QIP_infoButtonPressed := False
	QIP_title := title
	
	local GUIStart := 200
	local GUIWidth := 300
	local yAdditive := 10


	Gui, QIP_info:Font, s12
	Gui, QIP_info:Color, %dialogueColour%
	Gui, QIP_info:Add, Text, x10 y%yAdditive% w%width% hwndMessagePtr, % message
	ControlGetPos, x, y, w, h, , ahk_id %MessagePtr%
	yAdditive := yAdditive + h + 10
	
	widthTemp := width - 100
	Gui, QIP_info:Add, Button, x%widthTemp% y%yAdditive% default gQIP_infoOK,  &OK
	
	widthTemp := width - 50
	
	if (secondButton == "cancel")
		Gui, QIP_info:Add, Button, x%widthTemp% y%yAdditive% gQIP_infoClose, &Cancel
	else if (secondButton == "repeat")
		Gui, QIP_info:Add, Button, x%widthTemp% y%yAdditive% gQIP_repeat, &Repeat
	
	if !QIPX
		Gui, QIP_info:Show,, % title
	else
		Gui, QIP_info:Show, x%QIPX% y%QIPY%, % title
	
	Gui, QIP_info:+AlwaysOnTop
	WinWaitClose, %title%
	return QIP_infoButtonPressed
}




QIP_infoOK()
{
	global
	
	QIP_infoButtonPressed := "OK"
	QIP_infoClose()
}
QIP_repeat()
{
	global
	
	QIP_infoButtonPressed := "repeat"
	QIP_infoClose()
}
QIP_infoGuiClose()
{
	QIP_infoClose()
}
QIP_infoClose()
{
	global
	
	WinGetPos, QIPX, QIPY,,, % QIP_title
	Gui, QIP_info:Destroy
	return
}




QIP_sub(title, message)
{
	global
	
	timerStart := 0
	QIP_title := title
	
	local GUIStart := 200
	local GUIWidth := 300
	local yAdditive := 10


	Gui, QIP:Font, s12
	Gui, QIP:Color, %dialogueColour%
	Gui, QIP:Add, Text, x10 y%yAdditive% w500 hwndMessagePtr, % message
	
	ControlGetPos, x, y, w, h, , ahk_id %MessagePtr%
	yAdditive := yAdditive + h + 15
	
	Gui, QIP:Add, Text, x10 y%yAdditive%, % "MRN to search for:"
	Gui, QIP:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vMRN_GUI
	yAdditive := yAdditive + 40
	
	Gui, QIP:Add, Text, x10 y%yAdditive%, Search result:
	Gui, QIP:Add, Edit, x%GUIStart% y%yAdditive% W%GUIWidth% vQIP_answer
	yAdditive := yAdditive + 77

	Gui, QIP:Add, Button, x400 y%yAdditive% default gQIPOK,  &OK
	Gui, QIP:Add, Button, x450 y%yAdditive% gQIPClose, &Cancel

	Gui, QIP:Show, x%QIPX% y%QIPY%, % title
	Gui, QIP:+AlwaysOnTop
	WinWaitClose, %Title%
	return QIP_answer
}




QIPOK()
{
	global
	
	local errorString := ""
		
	Gui QIP:Submit, Nohide
	
	if !timerStart
	{
		GuiControl, QIP:, MRN_GUI, % MRN
		timerStart := A_TickCount
		return
	}
	else
	{
		elapsedTime := format("{:.3f}", (A_TickCount - timerStart) / 1000)
	}
	
	WinGetPos, QIPX, QIPY,,, % QIP_title
	Gui, QIP:Destroy
}




QIPGuiClose()
{
	QIPClose()
}
QIPClose()
{
	global 
	
	Gui, QIP:Destroy
	QIP_answer := "cancelled"
	return
}




QIP_log_update(message)
{
	global QIP_log_path
	
	FormatTime, Today,, dd/MM/yy - HH:mm
	
	if !fileExist(QIP_log_path)
	{
		FileAppend, Quick Spiritum QIP Log`n, % QIP_log_path
		FileAppend, % getClinicianDetails(clinicianUsername, A_UserName, clinicianActualName) . " (" . A_UserName . ")`n", % QIP_log_path
		sleep 200
	}
	
	FileAppend, % Today . "," . message . "`n", % QIP_log_path
	return True
}



