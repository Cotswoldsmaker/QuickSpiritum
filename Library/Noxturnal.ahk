; *****************
; Noxturnal library
; *****************


; *****************
; Variables !!! Need to see if can remove some of these variables
; *****************
NoxturnalEXE := """C:\Program Files\Nox Medical\Noxturnal\Noxturnal.exe"""
NoxturnalTitle := "Noxturnal"
NoxSearchBox :=  "WindowsForms10.EDIT.app.0.2386859_r30_ad11"
NoxFirstSelection := "WindowsForms10.STATIC.app.0.2386859_r30_ad123"
NoxViewTabs := "WindowsForms10.Window.8.app.0.2386859_r30_ad142"
NoxLibrary := "WindowsForms10.Window.8.app.0.2386859_r30_ad156"
NoxAddButton := "WindowsForms10.BUTTON.app.0.2386859_r30_ad11"




; *********
; Functions
; *********

; !!! Does not always click on the right patient. Need to work on this.
NoxturnalSearch(MRN)
{
	SetTitleMatchMode, 2	; 2 = contains
	SetKeyDelay, 10, 10

	global NoxturnalEXE, NoxturnalTitle, NoxSearchBox, NoxFirstSelection, NoxLibrary, NoxAddButton

	try
	{
		if !WinExist(NoxturnalTitle)
		{
			; Run Noxturnal
			try
			{
				SubtitleMessage("Starting up Noxturnal...")
				Run, %NoxturnalEXE%
			}
			catch
			{
				SubtitleClose()
				MB("Error starting Noxturnal! This is likely as it is not installed. Ending automation")
				return False
			}
				
			; Wait for search box
			Loop 500
			{
				if GraphicWait("NoxLibrarySelected", 200, , 50, False)
				{
					break
				}

				if GraphicWait("NoxLibraryUnselected", 200,, 50, False)
				{
					GraphicClick("NoxLibraryUnselected", 10, 10,, 50)
					sleep 500
				}

				; no need for sleep here as already included in above functions
			}

			ControlWait(NoxSearchBox, NoxturnalTitle)
			SetTitleMatchMode, 1     ; 1 = starts with
			SubtitleClose()
		}
		else
		{
			SetTitleMatchMode, 2           ; 2 = contains
			WinActivate, %NoxturnalTitle%
			WinWaitActive, %NoxturnalTitle%
			SetTitleMatchMode, 1     ; Starts with

			if !WinExist(NoxturnalTitle)
			{
				GlobalInputLockSet(True)
				Send, {alt}f{down 2}{enter}
				GlobalInputLockSet(False)
			}

			; Check that on Library tab
			Loop 500
			{
				if GraphicWait("NoxLibrarySelected", 200, , 50, False)
				{
					break
				}

				if GraphicWait("NoxLibraryUnselected", 200,, 50, False)
				{
					GraphicClick("NoxLibraryUnselected", 10, 10,, 50)
					sleep 500
				}

				; no need for sleep here as already included in above functions
			}

			ControlWait(NoxSearchBox, NoxturnalTitle)
		}

		write(NoxSearchBox, MRN, NoxturnalTitle)
		sleep 200
		ControlGetPos, xctrl, yctrl,,, %NoxSearchBox%, %NoxturnalTitle%
		GlobalInputLockSet(True)
		MouseMove, xctrl, yctrl + 60
		Click, 2
		GlobalInputLockSet(False)
	}
	catch, err
	{
		GlobalInputLockSet(False)
		SubtitleClose()
		errorhandler(err, "Noxturnal", "Noxturnal")
		return False
	}
	
	return True
}



