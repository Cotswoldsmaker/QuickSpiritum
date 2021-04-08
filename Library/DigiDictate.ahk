; *******************
; DigiDictate library
; *******************

; *********
; Variables
; *********

DigiDictateEXE := """C:\Program Files (x86)\Crescendo\DigiScribe-XL\DigiDictate-IP.exe"""
DigiDictateTitle := "DigiDictate-IP"
DigiDictateBar := "WindowsForms10.Window.8.app.0.378734a19"
DigiDictateNewDictationTitle := "New Dictation"
DigiDictateHospitalNo := "Edit2"
DigiDictateDate := "Edit4"




; *********
; Functions
; *********

DigiDictateSearch(MRN)
{
	global
	local NewDictationWindowOpen := 2
	local Today
	FormatTime, Today, ,dd/MM/yyyy
	local LoopBreak := False
	
	SetKeyDelay, 10, 10
	

	try
	{
		if !WinExist(DigiDictateTitle)
		{
			try
			{
				SubtitleMessage("Starting up DigiDictate...")
				Run, %DigiDictateEXE%
			}
			catch
			{
				SubtitleClose()
				MB("Error starting DigiDictate! This is likely as it is not installed. Ending automation")
				return False
			}

			; Wait for WindowForms control to be visible. Need this as control ID changes slightly each time DigiDitate starts up (last digit)
			Loop 300
			{
				WinGet, cList, ControlList, %DigiDictateTitle%

				Loop, parse, cList, \`n,`r
				{
					if (InStr(A_Loopfield, "WindowsForms10.Window.8.app.0.378734a") > 0)
					{
						LoopBreak := True
						break
					}
				}
				sleep 200
			
				if (LoopBreak == True)
				{
					break
				}
			}

			sleep 5000
		}
		else
		{
			WinActivate, %DigiDictateTitle%
		}
		
		SubtitleClose()

		Loop 2
		{
			ControlClick, x39 y96, %DigiDictateTitle%

			Loop 50
			{
				if WinExist(DigiDictateNewDictationTitle)
				{
					NewDictationWindowOpen := 0
					break
				}	
				sleep 200
			}

			if (NewDictationWindowOpen == 0)
			{
				break
			}
			else if (NewDictationWindowOpen == 2)
			{
				NewDictationWindowOpen := 1
			}
			else
			{
				MB("A new dictation window did not open! Ending automation",, "DevInform")
				return False
			}
		}
		
		sleep 200

		write(DigiDictateHospitalNo, MRN, DigiDictateNewDictationTitle)
		write(DigiDictateDate, Today, DigiDictateNewDictationTitle)
		PostClick(5, 5, "OK", DigiDictateNewDictationTitle)
	}
	catch, err
	{
		SubtitleClose()
		errorhandler(err, "DigiDictate", "DigiDictate")
		return False
	}
	
	return True
}



