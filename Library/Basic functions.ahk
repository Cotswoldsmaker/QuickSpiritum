; *********************************************************************************
; Some basic functions required by both Quick Spiritum and 'Update Master File.ahk'
; *********************************************************************************

; FiFo functions from HotKeyIT, posted 27/12/2010, https://autohotkey.com/board/topic/61792-ahk-l-for-loop-in-order-of-key-value-pair-creation/
FiFoArray(p*)
{
	array:=Object("",Object("",Object()),"base",Object("__Set","FiFoArray_","__Get","FiFoArray_"))
	for k,v in p
		If Mod(A_Index,2)
			_:=v
		else
			array[_]:=v
	Return array
}




FiFoArray_(o,k="",v="~empty~")
{
	if k=
		Return o["",""]
	else If v=~empty~
		Return o["",k]
	Return o["",k]:=v , o["",""].Insert(k)
}