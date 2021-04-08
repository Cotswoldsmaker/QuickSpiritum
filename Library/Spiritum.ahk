; **********************************
; Starting Spiritum database library
; **********************************

; *********
; Variables
; *********

SpiritumTitle := "Spiritum"
SpiritumSecurityNotice := "Microsoft Access Security Notice"
SpiritumUpdateNotice := "Front end update"
SpiritumLoginPage := "Login Page"




; *********
; Functions
; *********

SpiritumStart()
{
	SetKeyDelay, 1, 1	; 0, 0 will not work

	Global SpiritumTitle, SpiritumSecurityNotice, SpiritumUpdateNotice, SpiritumLoginPage, username, password

	system := "Spiritum"
	LoginSuccessful := False
	DatabaseRunning := False
	DatabaseOpen := False

	try
	{
		if (GetCredentials(system, True) == False)
		{
			return False
		}

		for name, obj in GetActiveObjects()
		{
			if InStr(name, "Spiritum.accde")
			{
				DatabaseOpen := True
				break
			}
		}


		; Start Spiritum if not already running
		if !DatabaseOpen or WinExist(SpiritumLoginPage)
		{
			if WinExist(SpiritumLoginPage)
				GoTo, LoginSp

			Try
				Run, %A_Desktop%\SpiritumFolder\Spiritum.accde
			Catch
			{
				MB("Spiritum does not seem to be installed. Please install before trying again")
				return False
			}

			; Click OK on security warning if present
			Loop, 20
			{
				If WinExist(SpiritumSecurityNotice)
				{
					PostClick(202, 213, "NetUIHWND1", SpiritumSecurityNotice)
					break
				}

					sleep 200
			}


			; Click on Update OK button if present
			Loop, 10
			{
				If WinExist(SpiritumUpdateNotice)
				{
					ControlGet, chwnd, Hwnd,,, %SpiritumUpdateNotice%
						ControlSend,, {enter}, ahk_id %chwnd%
				
					; Wait for update to complete
					Loop, 60
					{
						If WinExist(SpiritumSecurityNotice)
						{
							PostClick(202, 213, "NetUIHWND1", SpiritumSecurityNotice)
							break
						}

							sleep 1000
					}
				}
				else if WinExist(SpiritumLoginPage)
				{
					break
				}
					
				sleep 200
			}

			WinWait, %SpiritumLoginPage%

	LoginSp:

			sleep 200
			GlobalInputLockSet(True)
			write("OKttbx1", username, SpiritumLoginPage, 0)
			write("OKttbx1", password, SpiritumLoginPage, 0)
			ControlSend, ahk_parent, {enter}, ahk_class OFormPopup
			GlobalInputLockSet(False)

			Loop 20
			{
				if !WinExist(SpiritumLoginPage)
				{
					LoginSuccessful := True
					break
				}

				if WinExist("Microsoft Access")
				{
					break
				}
				
				sleep 200
			}
			
			if !LoginSuccessful
			{
				MB("Loggin error. Stopping automation")
				DeleteCredentials(system)
				return False
			}
		}
		else
		{
			MB("Spiritum is already open")
		}
	}
	catch, err
	{
		errorhandler(err, "Spiritum", "Microsoft Access / Spiritum")
		return False
	}

	return True
}



