; *************
; Email library
; *************

; *********
; Variables
; *********

NHSMailAddress := "https://email.nhs.net/owa"
NHSMailLoginTitle := "Sign In"
NHSMailMainTitle := "Outlook Web App"


; *********
; Functions
; *********

; turn testMail variable on or off via IPC
TestEmailSet(rawData)
{
	global testEmail
	
	testEmail := rawData

	return "pass"
}




; Automation by IPC to send email via Outlook
send_Email(rawData)
{
	global PBMainString, TestEmail
	
	returnValue := False
	
	CloseProgressBar()
	PBMainString := "Sending email:`n"
	CreateProgressBar()
	UpdateProgressBar(45, "absolute", PBMainString . "Sending…")

	DS := StrSplit(rawData, ";")
	toTemp := DS[1]
	cc := DS[2]
	title := DS[3]
	body := DS[4]
	attachmentmentPath := DS[5]
	
	if (DS[6] = "testEmail")
	{
		emailAddress := TestEmail
	}
	else
	{
		emailAddress := toTemp
	}
	
	returnValue := EmailOutlook(emailAddress, cc, title, body, attachmentPath)
	UpdateProgressBar(100, "absolute", PBMainString . "Sent")
	sleep 1000
	CloseProgressBar()
	
	if returnValue
		return "pass"
	else
		return "fail"
}




; Send email via Outlook
EmailOutlook(To := "", CC := "", Subject := "", Body := "", attachmentPath := "")
{
	SetTitleMatchMode, 2		; 2 = contains anywhere
	
	Process, Exist, OUTLOOK.EXE

	if !Errorlevel
	{
		Run, outlook.exe
		WinWait, % "Microsoft Outlook"
	}
	
	Loop 10
	{
		try
		{
			MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
			Emails:= StrSplit(To, ",")
			
			For key, email in Emails
			{
				MailItem.Recipients.Add(email)
			}

			if !(CC = "")
			{
				MailItem.cc := CC
			}

			if !(attachmentPath = "")
			{
				MailItem.attachments.add(attachmentPath)
			}

			MailItem.Subject := Subject
			MailItem.body := Body
			MailItem.send
			return True
		}
		catch, err
		{
			sleep 1000
			;errorhandler(err, "Outlook", "Outlook")
			;return False
		}
	}
	
	return False
}




; Open up NHS mail (obsolete), now no longer used as NHS mail is within Office 365!
NHSMailOpen()
{
	SetTitleMatchMode, 2           ; 2 = contains

	Global NHSMailAddress, NHSMailLoginTitle, NHSMailMainTitle, username, password

	ErrorCounter := 0

	Loop, 5 ; allow 5 errors before fail
	{
		system := "NHSMail"
		NHSMailLoginFound := False
		NHSMailMainFound := False
		NHSMailFound := False
		ElementFound := False
		Results := ""
		IE := ""
		LoginPass := False

		try
		{
			; need to reset these incase loop runs more than once
			NHSMailLoginFound := False
			NHSMailMainFound := False
			NHSMailLoginFound := False
			LoginPass := False

			For IE in ComObjCreate("Shell.Application").Windows
			{
				if InStr(IE.FullName, "iexplore.exe")
				{
					if InStr(IE.LocationName, NHSMailLoginTitle)
					{
						NHSMailLoginFound := True
					}
					else if InStr(IE.LocationName, NHSMailMainTitle)
					{
						NHSMailMainFound := True
					}
				}
			}


			if !NHSMailMainFound
			{
				if !GetCredentials(system)
				{
					return False
				}


				if !NHSMailLoginFound
				{
					IE := ComObjCreate("{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")
					IE.Visible := True
					IE.Navigate(NHSMailAddress)
					; Need next two lines as otherwise fails. Possibly as IE disconnects when URL address changes
					Loop, 10
					{
						if winExist(NHSMailLoginTitle)
							break

						if winExist(NHSMailMainTitle)
						{
							winActivate, % NHSMailMainTitle
							return True
						}

						sleep 200
					}

					;winWait, % NHSMailLoginTitle
					IE := GrabOpenIESession(NHSMailLoginTitle)
					IEBusy(IE)
				}


				WinMaximize, % NHSMailLoginTitle
				IEBusy(IE)

				IESet(,, "UserName", username . "@nhs.net", IE)
				IESet(,, "Password", password, IE)
				IEClick(, "ID", "submitButton",, IE)

				Loop 50
				{
					if IEHTMLFind(, "TagNameInnerHTML", "span", "Search mail and people", IE)
					{
						LoginPass := True
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
				WinActivate, % NHSMailMainTitle
			}

			return True
		}
		catch, err
		{
			if (err.message != "IE fail" OR ErrorCounter == 4)
			{
				try
					IE.quit

				WinClose, % NHSMailLoginTitle
				WinClose, % NHSMailMainTitle
				errorhandler(err, "NHS mail", "Internet Explorer / NHS mail")
				return False
			}
			else
			{
				; Basically close everything andn start again
				try
					IE.quit

				WinClose, % NHSMailLoginTitle
				WinClose, % NHSMailMainTitle
				ErrorCounter := ErrorCounter + 1
				SoundPlay, S:\Thoracic\Spiritum\Dev\bellplate-corner4.mp3
				MsgBox,,, % "Caught error, trying to grab COM to run NHS mail again [error counter: " . ErrorCounter . "]", 3
			}
		}
	}

	return False
}




; Send emails via NHS mail (obsolete). Never fully finished this work!
Email_NHS_Mail(To, CC, Subject, Body, attachmentPath)
{
	SetTitleMatchMode, 2           ; 2 = contains

	Global NHSMailAddress, NHSMailLoginTitle, NHSMailMainTitle, username, password

	system := "NHSMail"
	NHSMailLoginFound := False
	NHSMailMainFound := False
	NHSMailFound := False
	ElementFound := False
	Results := ""
	IE := ""

	try
	{

		; Needed a loop here as wb.busy and wb.ReadyState (or at least one of them) caused run time errors
		Loop 150
		{
			For IE in ComObjCreate("Shell.Application").Windows
			{
				If InStr(IE.FullName, "iexplore.exe") && InStr(IE.LocationName, NHSMailMainTitle)
				{
					NHSMailFound := True
					break
				}
			}

			if (NHSMailFound == True)
			{
				break
			}

			sleep 200
		}

		if (NHSMailFound == False)
		{
			MB("Error connecting to NHS mail! Stopping automation",, "DevInform")
			return False
		}

		WinMaximize, %NHSMailLoginTitle%
		IEBusy(IE)
		Results := IE.document.getElementsByTagname("span")

		Loop % Results.length
		{
			;msgbox, % Results[A_index-1].innerHTML
			if InStr(Results[A_index-1].innerHTML, "New Mail")
			{
				;msgbox, hit
				Results[A_index-1].click()
				break
			}
		}

		IEBusy(IE)
		sleep 5000
		Results := IE.document.getElementsByTagname("button")

		Loop % Results.length
		{
			;msgbox, % Results[A_index-1].innerHTML

			if InStr(Results[A_index-1].innerHTML, ">INSERT<")
			{
				;msgbox, hit
				;msgbox, % Results[A_index-1].innerHTML
				Results[A_index-1].click()
				break
			}

		}

		IEBusy(IE)
		sleep 2000
		Results := IE.document.getElementsByTagname("span")

		Loop % Results.length
		{
			;msgbox, % Results[A_index-1].innerHTML
			if InStr(Results[A_index-1].innerHTML, "Insert attachment")
			{
				;msgbox, hit
				Results[A_index-1].click()
				break
			}

		}

		;HTMLText := IE.document.documentElement.outerHTML
		;HTML2Notepad(HTMLText)

	}
	catch err
	{
		errorhandler(err, "NHS mail (sending)", "Internet explorer / NHS mail")
		;ErrorMessage := "Line: " . err.Line . ", Message: " . err.Message . ", What: " . err.What . ", Extra: " . err.what . ", File: " . err.File
		;MB("Problem with sending email via NHS mail [" . ErrorMessage . "]. Stopping automation",, "DevInform")
		return False
	}

	return True
}



