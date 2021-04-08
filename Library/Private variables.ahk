; ***************************************************************
; Place variables here that should not be known to public (eg 
; API keys)
; ***************************************************************

; Length of MRN (hospital number)
MRNLength := 7


TestEmail := "test@nhs.net" ; 


; Make sure you put a backslash on the end of paths to the below folders
settings := {}
settings.MasterDirectory := "S:\testFolder\"
settings.RequestsFolder := settings.MasterDirectory . "Requests\"	; Where to place request PDFs. In Dev mode these are placed in the \dev\Requests\ folder
settings.TemplatesFolder := CurrentDirectory . "Templates\"
settings.LocListLocationz := "1"
settings.IFZoom := "1"												; InfoFlex zoom in on letters (1) or not (0)


; Hospital / surgery address
DeptAddress := "first address line, second address line, postcode"


; Could perhaps make this doctor name variable more elegant with a dictionary of sorts !!!
; prefix D = doctor, C = consultant
DrConversion := FiFoArray("username1", "DRealName1"
							,"username2", "CRealName2")


; PET-CT					
PET_CT_Email := "testPET_CT@nhs.net"
PET_CT_Email_cc :=  "test1@nhs.net; test2@nhs.net; test3@nhs.net"


; Lung function tests
PFT_Email := "testPFT@nhs.net"


; Bronchoscopy
Bronch_Email :=  "testBronchoscopy@nhs.net"
Bronch_Email_cc :=  "test1@nhs.net; test2@nhs.net; test3@nhs.net"


; Healthy life style Gloucestershire
HLSG_Email := "testHLSG@nhs.net"


; Sleepstation
SleepStation_Email := "test_SleepStation@nhs.net"


;GOV.uk Notify
API_key := "secret key"
SMS_template_ID := "template_ID1"
email_template_ID := "template_ID2"
letter_template_ID := "template_ID3"


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


PFTPath := "S:\testPath\"


clinician1 := "Dr John Smith"
username1 := "John Smith"
userEmail1 := "test1@nhs.net"
mobile1 := "07123456789"
address1 := "first address line`r`nsecond address line`r`npostcode"