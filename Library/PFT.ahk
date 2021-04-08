; Lung function requests (view) library (mind out for GOTO statements with returns)


; Variables in header file


PFTSearch(MRN)
{
	; Test MRNs 1268241, 0669612
	
	global PFTPath, PFTFileList
	PFTFileList := []
	FilesCount := 0
	
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
	global MRNinput, PFTPath, PFTStoredHWNDs, PFTFileList

	FilesCount := 0
	StoredHWND := 0
	StoredHWNDDec := 0 
	PFTStoredHWNDs := []

	for index, value in PFTFileList
	{
		Path := PFTPath . A_LoopField		
		Gui, PFTGUI:Add, Text, cBlue gOpenPFT hwndStoredHWND, %value%
		StoredHWNDDec := StoredHWND + 0
		PFTStoredHWNDs.Push(StoredHWNDDec + 0)
		FilesCount += 1
	}
	
	yValue := 30 * FilesCount
	Gui, PFTGUI:Add, Button, x240 y%yValue%,  &Cancel
	Gui, PFTGUI:Show,, PFTs for MRN%MRN%
	Gui, PFTGUI:+AlwaysOnTop
	WinWaitClose, PFTs for MRN%MRN%§
	return True
}




PFTGUIClose:
PFTGUIButtonCancel:
PFTGUIGuiClose:
Gui, PFTGUI:Destroy
return




OpenPFT(CtrlHwnd, GuiEvent, EventLevel, ErrLevel :="")
{
	global PFTPath, PFTFileList, PFTStoredHWNDs

	FullPath := ""

	For index, value in PFTStoredHWNDs
	{
		if (value == CtrlHwnd)
		{
			FullPath := PFTPath . PFTFileList[index]
			run, %FullPath%
			break
		}
	}
	
	Gui, PFTGUI:Destroy
	return True
}

