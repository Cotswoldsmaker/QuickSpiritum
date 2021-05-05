; **************************************
; Lung function requests (view) library
; **************************************


; *********
; Variables
; *********

PFTFileList := []
PFTStoredHWNDs := []




; *********
; Functions
; *********

PFTSearch(MRN)
{
	global
	PFTFileList := []
	local FilesCount := 0
	
	try
	{
		SubtitleMessage("Searching for lung function test(s)...")
		
		if !FileExist(PFTPath)
		{
			SubtitleClose()
			MB("You currently do not have access to the Medisoft lung function folder! Stopping search")
			return False
		}
		
		
		Loop, Files, %PFTPath%*.PDF
		{
			if InStr(A_LoopFileName, MRN) > 0
			{
				PFTFileList.Push(A_LoopFileName)
				FilesCount += 1
			}
		}
		
		SubtitleClose()
		
		if (FilesCount == 0)
		{
			MB("No lung function results found for MRN" . MRN)
		}
		else
		{
			ShowPFTs(MRN)
		}
	}
	catch, err
	{
		SubtitleClose()
		errorhandler(err, "Lung Function Tests retrieval", "PDF")
		return False
	}
	
	return True
}




ShowPFTs(MRN)
{
	global
	StoredHWND := 0
	PFTStoredHWNDs := []
	local FilesCount := 0

	local StoredHWNDDec := 0 

	local GUI_name := "ShowPFTs_GUI" 

	for index, value in PFTFileList
	{
		Path := PFTPath . A_LoopField		
		Gui, %GUI_name%:Add, Text, cBlue gOpenPFT hwndStoredHWND, %value%
		StoredHWNDDec := StoredHWND + 0
		PFTStoredHWNDs.Push(StoredHWNDDec + 0)
		FilesCount += 1
	}
	
	yValue := 30 * FilesCount
	Gui, %GUI_name%:Add, Button, x240 y%yValue% gShowPFTs_close,  &Cancel
	Gui, %GUI_name%:Show,, PFTs for MRN%MRN%
	Gui, %GUI_name%:+AlwaysOnTop
	WinWaitClose, PFTs for MRN%MRN%
	return True
}


ShowPFTs_GUIguiClose()
{
	ShowPFTs_close()
}
ShowPFTs_close()
{
	global
	Gui, ShowPFTs_GUI:Destroy
	return
}




OpenPFT(CtrlHwnd, GuiEvent, EventLevel, ErrLevel :="")
{
	global
	local FullPath := ""


	For index, value in PFTStoredHWNDs
	{
		if (value == CtrlHwnd)
		{
			FullPath := PFTPath . PFTFileList[index]
			run, %FullPath%
			break
		}
	}
	
	Gui, ShowPFTs_GUI:Destroy
	return True
}

