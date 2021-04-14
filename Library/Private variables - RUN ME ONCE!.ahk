; **************************************************************************************
; Private variables - RUN ME ONCE! - run this program once to create a Private variables
; Folder outside of the git folder (in the folder before the \QuickSpiritum\ folder, 
; so secret / trust data is not copied on to GitHub
; **************************************************************************************

CurrentProgramPath := A_ScriptFullPath
PV_name := "Private variables.ahk"
QSPreviousFolderPosition := ""
PrivateVariablesLinkPath := A_ScriptDir . "\" . PV_name
PrivateVariablesMainPath := ""
newLine := "`r`n"
writing := False
created := False


QSPreviousFolderPosition := inStr(CurrentProgramPath, "\QuickSpiritum") 
PrivateVariablesMainPath := SubStr(CurrentProgramPath, 1, QSPreviousFolderPosition) . PV_name

if !fileExist(PrivateVariablesLinkPath)
{
	created := True
	fileObj := FileOpen(PrivateVariablesLinkPath, "w")
	fileObj.Write("#include " . PrivateVariablesMainPath)
}

if !fileExist(PrivateVariablesMainPath)
{
	created := True
	FileRead, FullFileContents, % CurrentProgramPath
	lines := StrSplit(FullFileContents, newLine)
	
	fileObj := FileOpen(PrivateVariablesMainPath, "w")
	
	For key, value in lines
	{
		if (inStr(value, "; END COPY") = 1)
		{
			break
		}
		
		if writing
			fileObj.Write(value . newLine)
		
		if (inStr(value, "; START COPY") = 1)
		{
			writing := True
		}
	}
}

if created
	msgbox, % "Created QS 'Private variables.ahk' link file and/or main file"
else
	msgbox, % "Link file and main 'Private variables.ahk' files have already been created"

/*
; START COPY
; *********************************************************************
; Place variables here that should not be known to public (eg API keys)
; *********************************************************************


; Length of MRN (hospital number)
MRNLength := 7


TestEmail := "emailAddress"


; Make sure you put a backslash on the end of paths to the below folders
settings := {}
settings.MasterDirectory := "S:\[path]\"
settings.RequestsFolder := settings.MasterDirectory . "Requests\"	; Where to place request PDFs. In Dev mode these are placed in the dev folder
settings.TemplatesFolder := CurrentDirectory . "Templates\"
settings.LocListLocationz := "1"
settings.IFZoom := "1"												; InfoFlex zoom in on letters (1) or not (0)


; Hospital / surgery address
DeptAddress := "1st line, 2nd line, City, Postcode"


; Could perhaps make this doctor name variable more elegant with a dictionary of sorts !!!
; D = doctor, C = consultant
DrConversion := FiFoArray("username1", "DrealName1"
						 ,"username2", "CRealName2")


; PET-CT					
PET_CT_Email := "emailAddress"
PET_CT_Email_cc :=  "emailAddress1; emailAddress2; emailAddress3"


; Lung function tests
PFT_Email := "emailAddress"


; Bronchoscopy
Bronch_Email :=  "emailAddress"
Bronch_Email_cc :=  "emailAddress1; emailAddress2; emailAddress3"


; Healthy life style Gloucestershire
HLSG_Email := "emailAddress"


; Sleepstation
SleepStation_Email := "emailAddress"


;GOV.uk Notify
API_key := "API key"
SMS_template_ID := "template1"
email_template_ID := "template2"
letter_template_ID := "template3"


; Quality Improvement Project
MRN1 := "0000000"
MRN2 := "0000000"
MRN3 := "0000000"
MRN4 := "0000000"
MRN5 := "0000000"
MRN6 := "0000000"
MRN7 := "0000000"
MRN8 := "0000000"
MRN9 := "0000000"
MRN10 := "0000000"
MRN11 := "0000000" ; PFT
MRN12 := "0000000"
MRN13 := "0000000"
MRN14 := "0000000"
MRN15 := "0000000"
MRN16 := "0000000"
MRN17 := "0000000"
MRN18 := "0000000"
MRN19 := "0000000"
MRN20 := "0000000"
MRN21 := "0000000"
MRN22 := "0000000"  
MRN23 := "0000000"
MRN24 := "0000000" ; PFT
MRN25 := "0000000"
MRN26 := "0000000"
MRN27 := "0000000"
MRN28 := "0000000"
MRN29 := "0000000"


secretaryExt := 1234 


PFTPath := "S:\[path]\"


clinician1 := "Dr ..."
username1 := "full name"
userEmail1 := "emailAddress"
mobile1 := "mobile number"
address1 := "1st line, 2nd line, City, Postcode"
; END COPY
*/