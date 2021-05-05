; ****************************************************************
; Main library for AHK
; ****************************************************************



; *********
; Variables
; *********

; Progress bar
PBTitle := "Progress"
PBMainString := ""
PBCounter := 0
PBPercentage := 0
PBText := ""


; general
ptr := A_PtrSize ? "ptr" : "Uint" 	; determine if 32 or 64 system when using Windows DLL API
GlobalInputLock := False
username := "" 
password := ""
cmdkeyTitle := "C:\Windows\SYSTEM32\cmdkey.exe"
hwndMyPic := ""
picture  := ""
oWord := ""
checkedBox := Chr(0xFE)
uncheckedBox := Chr(0x6F)	
signatureHeight := 30
MBMessage := ""
ErrCounterGlobal := 0
ImprivataRunning := False
vis2Obj := ""




; *********
; Functions
; *********

writeFAST(Control, Message, WinTitle, SleepAmount = 500)
{
      Blank := ""


      Lockset(True)
      ControlSetText, %Control%, %Blank%, %WinTitle%
      ControlSend, %Control%, %Message%, %WinTitle%
      Lockset(False)
      sleep SleepAmount
      return True
}




write(Control, Message, WinTitle, SleepAmount = 500)
{
      Lockset(True)
      ControlSetText, %Control%,, %WinTitle%		; Set blank
      ControlSend, %Control%, %Message%, %WinTitle%
      Lockset(False)
      sleep SleepAmount

      Lockset(True)
      keyPress(Control, "tab", WinTitle)
      Lockset(False)
      return True
}




keyPress(Control, key, WinTitle)
{
      Lockset(True)
	  ControlSend, %Control%, {%key%}, %WinTitle%
      Lockset(False)
      return True
}




keyPressWindow(Key, WinTitle)
{
      Lockset(True)
      ControlGet, chwnd, Hwnd,,, %WinTitle%
      ControlSend,, {%Key%}, ahk_id %chwnd%
      Lockset(False)
      return True
}




ControlWait(Control, WinTitle)
{
      x := ""
      y := ""


      Loop 1000
      {
            ControlGetPos, x, y,,, %Control%, %WinTitle%
			
            if (x <> "")
            {
                  return True
            }
            sleep 200
      }
	  
      return False
}




ControlPresent(Control, WinTitle)
{
	x := ""
	y := ""


	ControlGetPos, x, y,,, %Control%, %WinTitle%

	if (x <> "")
	{
		return True
	}
	else
	{
		return False	
	}
}




PostClick(x, y, control, title) 
{
      lParam := x & 0xFFFF | (y & 0xFFFF) << 16


      Lockset(True)
      PostMessage, 0x201, 1, %lParam%, %control%, %title% ;WM_LBUTTONDOWN 
      PostMessage, 0x202, 0, %lParam%, %control%, %title% ;WM_LBUTTONUP
      Lockset(False)

      return True
}




; Locks/unlocks keystrokes and mouse clicks
ToggleLock(cmd)
{
	SetFormat, IntegerFast, Hex
	Count := 0

	If (cmd == True)
	{
		Loop 0x1FF  ; Blocks keystrokes
		{
			if (A_Index != 1) ; skips ESC key
			{
				Hotkey, sc%A_Index%, DUMMY, On
			}
		}

		Loop 0x7    ; Blocks mouse clicks
		{
 			Hotkey, *vk%A_Index%, DUMMY, On
		}
	}
	Else
	{
		Loop 0x1FF  ; Unblocks keystrokes
		{
			Hotkey, sc%A_Index%, DUMMY, Off
		}

		Loop 0x7    ; Unblocks mouse clicks
		{
			Hotkey, *vk%A_Index%, DUMMY, Off
		}
	}

	SetFormat, IntegerFast, D ; D = Decimal
	return True
}




DUMMY()
{
	return
}




GlobalInputLockSet(set)
{
      global GlobalInputLock
	  
      if (set = True)
      {
            globalInputLock := True
            ToggleLock(True)
      }
      else
      {
            globalInputLock := False
            ToggleLock(False)
      }
	  
      return True
}




Lockset(set)
{
      global GlobalInputLock

      if (set = True)
      {
            ToggleLock(True)
      }
      else if ( GlobalInputLock == False AND set == False)
      {
            ToggleLock(False)
      }
	  
      return True
}




; Wait for a variant of a graphic. Don't include file type ending or version number in GraphicName
GraphicWait(GraphicName, WaitTime = 10000, type := "bmp", variance := 50, responseType := "HardFail")
{
	coordMode, Mouse, Window
	
	global GraphicsPath
	
	FoundX := ""
	FoundY := ""      
	SleepAmount := 50
	StartTime := A_TickCount
	FileVersions := 0
	
	If (WaitTime < SleepAmount)
		WaitTime := SleepAmount
	

	; Get number of versions of the graphic. If none available then cancel
	Loop
	{
		if !FileExist(GraphicsPath . GraphicName . A_Index . "." . type)
		{
			break
		}
		
		FileVersions := A_Index
	}
	
	if (FileVersions == 0)
	{
		throw Exception("No files found with " . type . " name prefix '" . GraphicName . "'", -1)
	}
	

	Loop
	{
		Loop, % FileVersions
		{
			fileLocation :=  GraphicsPath . GraphicName . A_Index . "." . type
			ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *%variance% %fileLocation%

			if(foundX != "")
			{
				return True
			}
		}

		if (A_TickCount - StartTime) > WaitTime
		{
			break
		}
		
		sleep SleepAmount
	}
	
	if (responseType == True or responseType == "HardFail")
	{
		throw Exception("Graphic fail", -1)
	}
	else if (responseType == "SoftFail")
	{
		throw Exception("SoftFail", 0)
	}
	else ; No exception thrown
	{
		return False
	}
}




; Click on a variant of graphic. Don't include file type ending or version number in GraphicName
GraphicClick(GraphicName, Xoffset, Yoffset, type := "bmp", variance := 50, responseType := "HardFail", doubleClick := False)
{
	coordMode, Mouse, Window
	
	global GraphicsPath
	
	FoundX := 0
	FoundY := 0
	newX := 0
	newY := 0   
	FileVersions := 0

	; Get number of versions of the graphic. If none available then cancel
	Loop
	{
		if !FileExist(GraphicsPath . GraphicName . A_Index . "." . type)
		{
			break
		}
		
		FileVersions := A_Index
	}
	
	if (FileVersions == 0)
	{
		throw Exception("No files found with " . type . " name prefix '" . GraphicName . "'", -1)
	}


	Loop, % FileVersions
	{
		fileLocation :=  GraphicsPath . GraphicName . A_Index . "." . type
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *%variance% %fileLocation%
	
		if(foundX != "")
		{
			newX := FoundX + Xoffset
			newY := FoundY + Yoffset
			Click, %newX%, %newY%
		
			if doubleClick
			{
				sleep 200
				Click, %newX%, %newY%
			}
			
			return True
		}
	}
	
	if (responseType == True or responseType == "HardFail")
	{
		throw Exception("Graphic fail", -1)
	}
	else if (responseType == "SoftFail")
	{
		throw Exception("SoftFail", 0)
	}
	else ; No exception thrown
	{
		return False
	}
}




; Move to a variant of graphic. Don't include file type ending or version number in GraphicName
GraphicMove(GraphicName, Xoffset, Yoffset, type := "bmp", variance := 50, responseType := "HardFail")
{
	coordMode, Mouse, Window
		
	Global GraphicsPath
	
	FoundX := 0
	FoundY := 0
	newX := 0
	newY := 0   
	FileVersions := 0

	; Get number of versions of the graphic. If none available then cancel
	Loop
	{
		if !FileExist(GraphicsPath . GraphicName . A_Index . "." . type)
		{
			break
		}
		
		FileVersions := A_Index
	}
	
	if (FileVersions == 0)
	{
		throw Exception("No files found with " . type . " name prefix '" . GraphicName . "'", -1)
	}


	Loop, % FileVersions
	{
		fileLocation :=  GraphicsPath . GraphicName . A_Index . "." . type
		ImageSearch, FoundX, FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, *%variance% %fileLocation%
		
		if(foundX != "")
		{
			newX := FoundX + Xoffset
			newY := FoundY + Yoffset
			Mousemove, %newX%, %newY%
			msgbox, to move
			return True
		}
	}
	
	if (responseType == True or responseType == "HardFail")
	{
		throw Exception("Graphic fail", -1)
	}
	else if (responseType == "SoftFail")
	{
		throw Exception("SoftFail", 0)
	}
	else ; No exception thrown
	{
		return False
	}
}




; Uses Windows Credentials to store credentials
SetCredentials(system, sendMethod := False)
{
	global
	local TempPassword := ""
	
	
	SetCredentialsGUI(system)

	if (username != "" AND password != "")
	{
		RunWait, cmdkey /generic:%system% /user:%username% /pass:%password%,, Hide
		
		; This does not change the saved credentials, just what is offered up to 'Send' 
		; functions if credentials are set at the same time as using them
		if sendMethod
		{
			Loop, Parse, password
			{
				if (A_loopfield == "{")
					TempPassword := TempPassword . "{{}"
				else if (A_loopfield == "}")
					TempPassword := TempPassword . "{}}"
				else if (A_loopfield == "!")
					TempPassword := TempPassword . "{!}"
				else if (A_loopfield == "#")
					TempPassword := TempPassword . "{#}"
				else if (A_loopfield == "+")
					TempPassword := TempPassword . "{+}"
				else if (A_loopfield == "^")
					TempPassword := TempPassword . "{^}"
				else
					TempPassword := TempPassword . A_loopfield
			}
			password := TempPassword
		}
		
		return True
	}
	else
	{
		MB("No value was entered for either username and/or password!")
		return False
	}

	return False
}




SetCredentialsGUI(system)
{
	global
	local GUI_name := "crendentials_GUI" 
	
	Gui, %GUI_name%:Add, Text, ym, No credentials on file for %system%. Please complete these below:
	Gui, %GUI_name%:Add, Text, x10 y30, Username:
	Gui, %GUI_name%:Add, Edit, x70 y30 vusername
	Gui, %GUI_name%:Add, Text, x10 y60, Password:
	Gui, %GUI_name%:Add, Edit, x70 y60 vpassword Password
	Gui, %GUI_name%:Add, Button, x230 y90 default gcrendentials_OK,  &OK
	Gui, %GUI_name%:Add, Button, x270 y90 gcrendentials_close,  &Cancel
	Gui, %GUI_name%:Show,, Credentials needed
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, Credentials needed
	return True
}




crendentials_OK()
{
	Gui, Crendentials_GUI:Submit
	Gui, Crendentials_GUI:Destroy
	return
}




crendentials_GUIGuiClose()
{
	crendentials_close()
}
crendentials_close()
{
	Gui, crendentials_GUI:Destroy
	return
}





GetCredentials(system, sendMethod := False) 
{ 
	global username, password
	
	TempPassword := ""
 	pCred := 0 
	credentialBlobSizeOffset := 16 + 2*A_PtrSize 
	pCredentialBlobOffset := 16 + 3*A_PtrSize 
	userNameOffset := 24 + 6*A_PtrSize 


	ret := DllCall("ADVAPI32\CredReadW", "WStr", system, "UInt", 1, "UInt", 0, "Ptr*", pCred, "Int") 

	if (ErrorLevel != 0) 
	{  
		return False
	} 
 
	; This is triggered if no credentials already set.
	if (ret != 1) 
	{ 
		if SetCredentials(system, SendMethod)
		{
			return True
		}
		else
		{
			return False
		}
	}  

	credentialBlobSize := NumGet(pCred + credentialBlobSizeOffset, "UInt") 
	pCredentialBlob := NumGet(pCred + pCredentialBlobOffset, "Ptr") 
	pUserName := NumGet(pCred + userNameOffset, "Ptr")
	password := StrGet(pCredentialBlob, credentialBlobSize / 2, "UTF-16") 
	userName := StrGet(pUserName, , "UTF-16" ) 
	DllCall("ADVAPI32\CredFree", "Ptr", pCred) 
 
	if (ErrorLevel != 0) 
	{ 
		MB("DllCall error invoking CredFree: " . ErrorLevel)
		return False
	} 

	if SendMethod
	{
		PasswordAlter()
	}

	return True
} 




PasswordAlter()
{ 
	global
	local TempPassword := ""

	Loop, Parse, password
	{
		if (A_loopfield == "{")
			TempPassword := TempPassword . "{{}"
		else if (A_loopfield == "}")
			TempPassword := TempPassword . "{}}"
		else if (A_loopfield == "!")
			TempPassword := TempPassword . "{!}"
		else if (A_loopfield == "#")
			TempPassword := TempPassword . "{#}"
		else if (A_loopfield == "+")
			TempPassword := TempPassword . "{+}"
		else if (A_loopfield == "^")
			TempPassword := TempPassword . "{^}"
		else
			TempPassword := TempPassword . A_loopfield
	}
	
	password := TempPassword
	return
}




DeleteCredentials(system)
{
	global cmdkeyTitle

	RunWait, cmdkey /delete:%system%,, Hide
	return True
}




CreateProgressBar()
{
	global
	PBPercentage := 0
	local GUI_name := "progressBar_GUI" 
	
	
	Gui, %GUI_name%:Font, s12
	GUI, %GUI_name%:Add, Progress, vPBCounter H50 W270
	Gui, %GUI_name%:Add, Text, x130 y60 W50 vPBPercentage, 0`%
	Gui, %GUI_name%:Add, Text, x15 y80  H270 W270 +Wrap vPBText, start
	GUI, %GUI_name%:Show, H270 W300, %PBTitle%
	Gui, %GUI_name%:+AlwaysOnTop
	return True
}




progressBar_GUIGuiClose()
{
	Gui, progressBar_GUI:Destroy
	return
}




CloseProgressBar()
{
	global PBTitle, PBCounter, PBPercentage, PBText

	if WinExist(PBTitle)
	{
		WinClose, %PBTitle%
		sleep 200
		PBCounter := 0 
		PBPercentage := 0 
		PBText := ""
		return True
	}

	return False
}




UpdateProgressBar(amount := 10, mode := "additive", text := "", waitForClosure := False)
{
	global
	local GUI_name := "progressBar_GUI"

	if PBPercentage < 100
	{
		if (mode = "additive")
		{
			GuiControl, %GUI_name%: , PBCounter, +%amount%
			PBPercentage := PBPercentage + amount
			GuiControl, %GUI_name%: , PBPercentage, %PBPercentage%`%
		}
		else if (mode = "absolute")
		{
			GuiControl, %GUI_name%: , PBCounter, %amount%
			PBPercentage := amount
			GuiControl, %GUI_name%: , PBPercentage, %PBPercentage%`%
		}
	}

	if !(text = "")
	{
		GuiControl, %GUI_name%: , PBText, %text%
	}
	
	if waitForClosure
	{
		Loop, 50
		{
			if !WinExist(PBTitle)
				break
				
			sleep 200
		}
	}

	return True
}




GraphicsDimensions(Path)
{
	global
	local x, y, w, h

	Gui, PicDimensions:Add, Picture, vpicture hwndMyPic, %Path%
	ControlGetPos, x, y, w, h, , ahk_id %MyPic%
	Gui, PicDimensions:Destroy
	return Picture w / h
}




GetComputerNameEx(COMPUTER_NAME_FORMAT := 0)                               
{
    DllCall("GetComputerNameEx", "UInt", COMPUTER_NAME_FORMAT, "Ptr", 0, "UInt*", size)
    VarSetCapacity(buf, size * (A_IsUnicode ? 2 : 1), 0)

    if !(DllCall("GetComputerNameEx", "UInt", COMPUTER_NAME_FORMAT, "Ptr", &buf, "UInt*", size))
        return "*" A_LastError

    return StrGet(&buf, size, "UTF-16")
}




MB(Message, Title := "Quick Spiritum", method := "message", wait := True)
{	
	global
	MBMessage := Message
	local GUI_name := "MB_GUI" 
	
	SetTitleMatchMode, 3
	

	if (method = "message")
	{
		Gui, %GUI_name%:Font, s12
		Gui, %GUI_name%:Color, %dialogueColour%
		Gui, %GUI_name%:Add, Text, ym w300 hwndMessagePtr, %Message%
		ControlGetPos, x, y, w, h, , ahk_id %MessagePtr%
		h := h + 50
		Gui, %GUI_name%:Add, Button, x350 y%h% default gMB_close,  &OK
		Gui, %GUI_name%:Show,, %Title%
		Gui, %GUI_name%:+AlwaysOnTop
		
		if wait
			WinWaitClose, %Title%
	}
	else if (method = "DevInform")
	{
		Gui, %GUI_name%:Font, s12
		Gui, %GUI_name%:Color, %dialogueColour%
		Gui, %GUI_name%:Add, Text, ym w1000 hwndMessagePtr, % Message . "`n`nAre you happy to send the above error message to " . username1 . "?"
		ControlGetPos, x, y, w, h, , ahk_id %MessagePtr%
		h := h + 50
		Gui, %GUI_name%:Add, Button, x900 y%h% default gMB_Yes,  &Yes
		Gui, %GUI_name%:Add, Button, x950 y%h% gMB_close,  &No
		Gui, %GUI_name%:Show,, Error - %Title%
		Gui, %GUI_name%:+AlwaysOnTop
		
		if wait
			WinWaitClose, Error - %Title%
	}
	
	return True
}



MB_Yes()
{
	global 
	DetectHiddenWindows Off
	Gui, MB_GUI:Destroy
	WinGet windows, List
		
	Loop %windows%
	{
		id := windows%A_Index%
		WinGetTitle wt, ahk_id %id%
		
		if (wt != "")
			WindowsList .= wt . "`n"
	}

	EmailOutlook(userEmail1,, "Quick Spiritum error", "An error has occured with quick Spiritum:`n`n" . MBMessage . "`n`nOpen windows include:`n`n" . WindowsList) ;, ScreenGrabPath)
	return
}




MB_GUIGuiClose()
{
	MB_close()
}
MB_close()
{
	Gui, MB_GUI:Destroy
	return
}




ErrorHandler(err := "", Program := "", Systems := "", message := "")
{
	global CurrentProgramPath, CurrentVersionNumber
	
	GlobalInputLockSet(False)
	
	if (err = "")
	{
		LogUpdate("Error: program - " . Program . ", message - " . message)
		MB("Error with automating " . Program . ". Please close all " . Systems . " sessions and try again [QS version number: " . CurrentVersionNumber . "]`n`nMessage: " . message,, "DevInform")
	}
	else
	{
		ErrorMessage := "Line: " . err.Line . ", Message: " . err.Message . ", What: " . err.What . ", Extra: " . err.what . ", File: " . err.File
		LineNumber := err.Line - 11
		Code := ""
		line := ""
		FileName := err.File
	
		Loop 20, 
		{
			FileReadLine, line, %FileName%, % LineNumber + A_Index
			if (A_Index = 11)
			{
				Code := Code . "`n*--->" . line			
			}
			else
			{
				Code := Code . "`n*" . line
			}
		}
	
		LogUpdate("Error: program - " . Program . ", error message - " . ErrorMessage)
		MB("Error with automating " . Program . ". Please close all " . Systems . " sessions and try again [" . ErrorMessage . ", QS version number: " . CurrentVersionNumber . "]`n`nCode: " . Code,, "DevInform")
	}
	
	return True
}




; lexikos - https://www.autohotkey.com/boards/viewtopic.php?f=6&t=6494 - 22/02/2015
GetActiveObjects(Prefix:="", CaseSensitive:=false) {
    objects := {}
    DllCall("ole32\CoGetMalloc", "uint", 1, "ptr*", malloc) ; malloc: IMalloc
    DllCall("ole32\CreateBindCtx", "uint", 0, "ptr*", bindCtx) ; bindCtx: IBindCtx
    DllCall(NumGet(NumGet(bindCtx+0)+8*A_PtrSize), "ptr", bindCtx, "ptr*", rot) ; rot: IRunningObjectTable
    DllCall(NumGet(NumGet(rot+0)+9*A_PtrSize), "ptr", rot, "ptr*", enum) ; enum: IEnumMoniker

    while DllCall(NumGet(NumGet(enum+0)+3*A_PtrSize), "ptr", enum, "uint", 1, "ptr*", mon, "ptr", 0) = 0 ; mon: IMoniker
    {
        DllCall(NumGet(NumGet(mon+0)+20*A_PtrSize), "ptr", mon, "ptr", bindCtx, "ptr", 0, "ptr*", pname) ; GetDisplayName
        name := StrGet(pname, "UTF-16")
        DllCall(NumGet(NumGet(malloc+0)+5*A_PtrSize), "ptr", malloc, "ptr", pname) ; Free

        if InStr(name, Prefix, CaseSensitive) = 1 
	{
            DllCall(NumGet(NumGet(rot+0)+6*A_PtrSize), "ptr", rot, "ptr", mon, "ptr*", punk) ; GetObject
            ; Wrap the pointer as IDispatch if available, otherwise as IUnknown.
            if (pdsp := ComObjQuery(punk, "{00020400-0000-0000-C000-000000000046}"))
                obj := ComObject(9, pdsp, 1), ObjRelease(punk)
            else
                obj := ComObject(13, punk, 1)
            ; Store it in the return array by suffix.
            objects[SubStr(name, StrLen(Prefix) + 1)] := obj
        }

        ObjRelease(mon)
    }

    ObjRelease(enum)
    ObjRelease(rot)
    ObjRelease(bindCtx)
    ObjRelease(malloc)
    return objects
}




GrabText(control, window)
{
	ControlGetText, text, %control%, %window%
	return text
}




; Writing at book mark in word document
WriteAtBookmark(Bookmark, ToType)
{
	Global oWord
	try
	{
		oWord.ActiveDocument.Bookmarks(Bookmark).Select
		oWord.selection.TypeText(ToType)
		return true
	}
	catch
	{
		throw Exception("Write at bookmark fail: " Bookmark . " - " . ToType, -1)
	}
}




; In word documents
InsertCheckboxAtBookmark(Bookmark, status)
{
	Global oWord, checkedBox, uncheckedBox

	oWord.ActiveDocument.Bookmarks(Bookmark).Select
	oWord.Selection.Font.Name := "Wingdings"

	if (status == 1) ; checked
	{
		oWord.selection.TypeText(checkedBox)
	}
	else if (status == 0) ; unchecked
	{
		oWord.selection.TypeText(uncheckedBox)
	}
	else
	{
		return False
	}

	return True
}




timedFunction()
{
	turnOffImprivata()
	;SunriseTimerCatch()
	return
}




; !!! Work in progress
/*
SunriseTimerCatch()
{
	global SunriseMainTitle, SunriseHWND, SunrisePID, pCallback, SRCatchFunctionRunning
	
	EVENT_OBJECT_CREATE := "0x8000"
	
	if winExist(SunriseMainTitle)
	{
		if (SRCatchFunctionRunning == False)
		{
			WinGet, SunriseHWND, ID, %SunriseMainTitle%
			WinGet, SunrisePID, PID, %SunriseMainTitle%
			pCallback := RegisterCallback("WinEventProc")
			Acc_SetWinEventHook(EVENT_OBJECT_CREATE, EVENT_OBJECT_CREATE, pCallBack)
			SRCatchFunctionRunning := True
		}
	}
	else
	{
		if (SRCatchFunctionRunning == True)
		{
			Acc_UnhookWinEvent(pCallback)
			SRCatchFunctionRunning := False
		}
	}
	
	return
}




WinEventProc(hHook, event, ChildHwnd, idObject, idChild, eventThread, eventTime) 
{
	global SunriseHWND, SunrisePID, SunriseMainTitle, SRTMessage
	
	SetTitleMatchMode, 2
	msgbox, % SunrisePID . " - " . ChildHwnd
	;if (idChild == SunrisePID)
	if (DllCall("GetParent", "Ptr", ChildHwnd) = SunriseHWND)
	{
		WinGetTitle, WinTitle, ahk_id %ChildHwnd%
		
		if (WinTitle = "Find Patient")
		{
			msgbox, child present
		}
	}
*/
/*
	try
	{
		TimerAppeared := False
		Acc := Acc_ObjectFromEvent(_idChild_, ChildHwnd, idObject, idChild)
		FormatTime, Today,, HH:mm
		
		;msgbox, caught
		
		if (DllCall("GetParent", "Ptr", ChildHwnd) = SunriseHWND) ;Acc.accRole(0)=9 and 
		{
			WinGetTitle, WinTitle, ahk_id %ChildHwnd%
			msgbox, child function
			if inStr(WinTitle, SunriseMainTitle)
			{
				Msgbox,, abc, % SRTMessage, 3
				
				Loop 10
				{
					if GraphicWait("SunriseAboutToLogOff", 200,,, False)
					{
						if GraphicClick("SunriseAboutToLogOff", 132, 47,,, False)
						{
							TimerAppeared := True
							SRTMessage := SRTMessage . "Caught timer dialogue at: " . Today . "`n"
							Msgbox,, Sunrise timer closer, % SRTMessage, 3
							Menu, Tray, Tip, % SRTMessage
						}
					}
					else if (TimerAppeared == True)
					{
						break
					}
					
					sleep 200
				}
				
				Msgbox, Hopefully closed
			}
		}
	}
	catch
	{
		Acc_UnhookWinEvent(pCallback)
		MsgBox, Error ? closed
		ExitApp
	}
*/

/*
}




GetClassNN(Chwnd, Whwnd) {
	global _GetClassNN := {}
	_GetClassNN.Hwnd := Chwnd
	Detect := A_DetectHiddenWindows
	WinGetClass, Class, ahk_id %Chwnd%
	_GetClassNN.Class := Class
	DetectHiddenWindows, On
	EnumAddress := RegisterCallback("GetClassNN_EnumChildProc")
	DllCall("EnumChildWindows", "uint",Whwnd, "uint",EnumAddress)
	DetectHiddenWindows, %Detect%
	return, _GetClassNN.ClassNN, _GetClassNN:=""
}




GetClassNN_EnumChildProc(hwnd, lparam) {
	static N
	global _GetClassNN
	WinGetClass, Class, ahk_id %hwnd%
	if _GetClassNN.Class == Class
		N++
	return _GetClassNN.Hwnd==hwnd? (0, _GetClassNN.ClassNN:=_GetClassNN.Class N, N:=0):1
}


*/




; Turn off Imprivata from entering crendentials - otherwise this interfers with QS
turnOffImprivata()
{
	DetectHiddenWindows, % (Setting_A_DetectHiddenWindows := A_DetectHiddenWindows) ? "On" :
	
	TI := TrayIcon_GetInfo("ISXAgent.exe")
	tooltip := TI[1].tooltip
	HWND := TI[1].hwnd

	if !inStr(tooltip, "suspended")
	{
		; Change running/paused state
		PostMessage, 0x111, 223, 0,, ahk_id %HWND%
	}

	DetectHiddenWindows, %Setting_A_DetectHiddenWindows%
	return
}




; !!! Need to condense this and potentially remove the OCR part of this
SubtitleMessage(message := "")
{
	global vis2Obj
	
	vis2Obj := IsObject(obj) ? obj : {}
	vis2Obj.EXITCODE := 0 ; 0 = in progress, -1 = escape, 1 = success
	vis2Obj.selectMode := "Quick"
	vis2Obj.area := new Vis2.Graphics.Area("Vis2_Aries", "0x7FDDDDDD")
	vis2Obj.image := new Vis2.Graphics.Image("Vis2_Kitsune").Hide()
	vis2Obj.subtitle := new Vis2.Graphics.Subtitle("Vis2_Hermes")

	vis2Obj.style1_back := {"x":"center", "y":"83%", "padding":"1.35%", "color":"DD000000", "radius":8}
	vis2Obj.style1_text := {"q":4, "size":"2.23%", "font":"Arial", "z":"Arial Narrow", "justify":"left", "color":"White"}
	vis2Obj.style2_back := {"x":"center", "y":"83%", "padding":"1.35%", "color":"FF88EAB6", "radius":8}
	vis2Obj.style2_text := {"q":4, "size":"2.23%", "font":"Arial", "z":"Arial Narrow", "justify":"left", "color":"Black"}
	vis2Obj.subtitle.render(message, vis2Obj.style1_back, vis2Obj.style1_text)

	return
}




SubtitleClose()
{
	global vis2Obj
	vis2Obj.subtitle.destroy()
	return
}



; Code acknowledgement:
; jebb - https://autohotkey.com/board/topic/39129-help-calculating-the-difference-between-two-times/page-2 - 29/03/2009
timeDiff(startTime, FinishTime) 
{
  starthour := SubStr(startTime, 1, 2)
  startminute := SubStr(startTime, 4, 2)
  Start := 20000101 starthour startminute "00"

  finishhour := SubStr(finishTime, 1, 2)
  finishminute := SubStr(finishTime, 4, 2)
  finish := 20000101 finishhour finishminute "00"
  finish -= start, seconds

  T = 20000101000000 
  T += finish,Seconds 
  FormatTime FormdTHH, %T%, HH.mm
  return, FormdTHH
}




; Checks if in format HH:MM
checkifTime(timeStr)
{
	length := StrLen(timeStr)
	
	if (length < 3 and length > 5)
		return False
		
	SCPosition := InStr(timeStr, ":")
	
	if (SCPosition == 0 or SCPosition == 1 or SCPosition == length)
		return False
		
	H := SubStr(timeStr, 1, SCPosition - 1)
	if H is not digit
		return False
	
	M := SubStr(timeStr, SCPosition + 1)
	if M is not digit
		return False

	if (H >= 0 and H <= 23)
		if (M >= 0 and M <= 59)
				return True
	
	return False
}



