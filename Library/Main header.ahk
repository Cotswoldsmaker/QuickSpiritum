; *******************************************************
; Main header (mind out for GOTO statements with returns)
; !!! I need to merge the main header and main files
; *******************************************************

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