; ***********************************************************
; Requests header (mind out for GOTO statements with returns)
; ***********************************************************


PatientDetails := {}
requestList := ["PET_CT", "PFT", "bronchoscopy"]
referralList := ["HLSG", "Sleepstation"]


; PET-CT request variables
PET_CT_template_path := Settings.TemplatesFolder . "PET_CT_request_template.docx"
PET_CT_latestRequestPath := ""	
PET_CT_details := {} ; Initialise dictionary
PET_CT_consultantName := ""
PET_CT_doctorName := ""
PET_CT_previousImagingType := ""
PET_CT_previousImagingDate := ""
PET_CT_clinicalInformation := ""
PET_CT_diabetic := ""
PET_CT_emailLungCancerCoordinators := 1 	; 1 = checked, 0 = unchecked
PET_CTButtonPressed := ""




; Lung function test request variables
PFT_templatePath := Settings.TemplatesFolder . "PFT_request_template.docx"
PET_latestRequestPath := ""
PFT_details := {} ; Initialise dictionary
PFT_consultantName := ""
PFT_doctorName := ""
PFT_clinicalInformation := ""
PFT_buttonPressed := ""
PFT_speciality := ""
PFT_location := ""
PFT_wardText := ""
PFT_ward := ""
PFT_urgency := ""
PFT_clinicDateText := ""
PFT_clinicDate := ""

PFT_tests := FiFoArray("Withhold bronchodilators prior", "holdBronchodilators"
		 , "Spirometry / flow loop" , "spirometry"
		 , "Gas transfer (TLCO & KCO)", "gasTransfer"
		 , "Static lung volumes", "staticLungVolumes"
		 , "Exhaled nitric oxide (FENO)", "FENO"
		 , "Capillary blood gases on air", "CBG"
		 , "Sitting / supine spirometry", "sittingStanding"
		 , "Bronchodilator response", "bronchodilatorResponse"
		 , "Osmohale (Mannitol)", "osmohale"
		 , "Hypoxic challenge (fit to fly)", "fitToFly"
		 , "Multi-channel sleep study", "SS"
		 , "CPAP trial", "CPAP"
		 , "Mouth pressures", "mouthPressures"
		 , "Overnight pulse oximetry", "overnightPulseOx"
		 , "6 week occupational asthma study", "occupationalAsthmaStudy"
		 , "Home NIV", "NIV")

holdBronchodilators := 1
spirometry := 0
gasTransfer := 0
staticLungVolumes := 0
FENO := 0
CBG := 0
sittingStanding := 0
bronchodilatorResponse := 0
osmohale := 0
fitToFly := 0
SS := 0
CPAP := 0
mouthPressures := 0
overnightPulseOx := 0
occupationalAsthmaStudy := 0
NIV := 0
NIVMode := 0
NIVIPAP := 0
NIVEPAP := 0




; Bronchoscopy request variables
Bronch_templatePath := Settings.TemplatesFolder . "Bronchoscopy_request_template.docx"
Bronch_details := {} ; Initialise dictionary
Bronch_latestRequestPath := ""
Bronch_procedure := ""
Bronch_location := ""
Bronch_primaryDiagnosis := ""
Bronch_significantComorbidities := ""
Bronch_preferredBronchoscopyDate := ""
Bronch_consentCompleted := ""
Bronch_anticoagulantAntiplatelets := ""
Bronch_aim := ""
Bronch_imaging := ""
Bronch_consultantName := ""
Bronch_doctorName := ""
Bronch_emailCoordinators := 1 	; 1 = checked, 0 = unchecked




;HLSG
HLSG_templatePath := Settings.TemplatesFolder . "HLSG_request_template.docx"
HLSG_latestRequestPath := ""
HLSG_doctorName := ""
HLSG_details := {}
HLSG_ethnicity := ""
HLSG_Comorbidities := ""
HLSG_buttonPressed := ""
HLSG_referralTypes := FiFoArray("Smoking cessation", "HLSG_smoking"
								, "Alcohol services" , "HLSG_alcohol",
								, "Weight management", "HLSG_weight",
								, "Physical activity", "HLSG_physicalActivity")

HLSG_smoking := 0
HLSG_alcohol := 0
HLSG_weight := 0
HLSG_physicalActivity := 0




;SleepStation
Sleepstation_template_path := Settings.TemplatesFolder . "Sleepstation_referral_template.docx"
SleepStation_latestRequestPath := ""
SleepStation_details := {}
Sleepstation_buttonPressed := ""




;GOV.uk Notify
GUKN_type := ""
GUKN_mobileNumber := ""

GUKN_email_address := ""
GUKN_email_subject := ""
GUKN_email_body := ""

GUKN_letter_to := ""
GUKN_letter_address := ""
GUKN_letter_from := ""
GUKN_letter_header := ""
GUKN_letter_body := ""



; Test email
if emailTest
{
	PET_CT_Email := TestEmail
	PET_CT_Email_cc := ""
	PFT_Email := TestEmail
	Bronch_Email := TestEmail
	Bronch_Email_cc := ""
	HLSG_Email := TestEmail
	SleepStation_Email := TestEmail
}


