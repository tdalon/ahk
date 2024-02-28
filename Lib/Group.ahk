; Reference: https://www.autohotkey.com/boards/viewtopic.php?t=60272

/* 
;Uses a combined list & array Groups index,for ease of use & efficiency,respectively.
; Example:
Group_Add("Browser", "ahk_exe iexplore.exe","ahk_exe chrome.exe","ahk_exe sidekick.exe","ahk_exe firefox.exe","ahk_exe palemoon.exe","ahk_exe waterfox.exe","ahk_exe vivaldi.exe","ahk_exe msedge.exe","ahk_exe opera.exe")
Order: last in the list is checked first

;create new group and make it a 'real' group...
GroupAdd("browsers", "ahk_exe chrome.exe", "ahk_exe firefox.exe")
GroupTranspose("browsers")
SetTimer, listGroups, 300
;toogle group member
x::IsInGroup("mediaPlayers","vlc.exe") ? GroupDelete("mediaPlayers", "ahk_exe vlc.exe") : GroupAdd("mediaPlayers", "ahk_exe vlc.exe")

;(de)activate a group... of course group members can't be removed from the transposed 'real' group,only from the custom group they were inherited from...
a::GroupActivate % GroupTranspose("browsers")
d::GroupDeactivate % GroupTranspose("browsers")

#If GroupActive("mediaPlayers")
g::MsgBox 0x40040,, MEDIA PLAYER GROUP ACTIVE
#If

#IfWinActive ahk_group browsers	;Custom Group transformed to real group...
g::MsgBox 0x40040,, BROWSER GROUP ACTIVE
#IfWinActive

listGroups:
ToolTip % A_Groups "`n`n" GroupActive("mediaPlayers") "`n" GroupActive("browsers")
Return
 */
;====================================================================================================

Group_Add(groupName,groupMembers*){
	Global A_Groups,A_GroupsArr
	( !InStr(A_Groups, groupName ",") ? (A_Groups .= A_Groups ? "`n" groupName "," : groupName ",") : "" )	;initialise group if it doesn't exist...
	For i,groupMember in groupMembers
		( !InStr(A_Groups, groupMember) ? (A_Groups := StrReplace(A_Groups, groupName ",", groupName "," groupMember ",")) : )	;append to or create new group...
	A_Groups := RegExReplace(RegExReplace(A_Groups, "(^|\R)\K\s+"), "\R+\R", "`r`n")	;clean up groups to remove any possible blank lines
	,ArrayFromList(A_GroupsArr,A_Groups)	;rebuild group as array for most efficient cycling through groups...
}

Group_Delete(groupName, groupMember:=""){
	Global A_Groups,A_GroupsArr
	For i,group in StrSplit(A_Groups,"`n")
		( groupMember && group && InStr(A_Groups,groupMember) && groupName = StrSplit(group,",")[1] ? (A_Groups:=StrReplace(A_Groups,group,StrReplace(group,groupMember ","))) : !groupMember && groupName = StrSplit(group,",")[1] ? (A_Groups:=StrReplace(A_Groups,group))  )	;remove group member from group & update group in A_Groups
	A_Groups := RegExReplace(RegExReplace(A_Groups, "(^|\R)\K\s+"), "\R+\R", "`r`n")	;clean up groups to remove any possible blank lines
	,ArrayFromList(A_GroupsArr,A_Groups)	;rebuild group as array for most efficient cycling through groups...
}

ArrayFromList(ByRef larray, ByRef list, listDelim := "`n", lineDelim:=","){
	larray := []
	Loop, Parse, list, % listDelim
		larray.Push(StrSplit(A_LoopField,lineDelim))
}

;Functions below are subject to a performance overhead & hence use A_GroupsArr...as they are repeatedly called...
Group_Active(groupName){
	Global A_GroupsArr
	For i,group in A_GroupsArr
		If (group.1 = groupName)
			For iG,groupMember in group
					If (iG > 1 && groupMember && (firstMatchId := WinActive(groupMember)))
						Return firstMatchId ;Return group.1 "," firstMatchId
	return 0 ; compatible with WinActive function			
}

Group_Transpose(groupName){	;makes this custom group,a 'real' group,some use cases....
	Global A_GroupsArr
	For i,group in A_GroupsArr
		If (group.1 = groupName)
			For iG,groupMember in group
				If (iG > 1 && groupMember)
					GroupAdd, % group.1, % groupMember
	Return groupName
}

IsInGroup(groupName, groupMember){
	Global A_Groups
	Loop, Parse, A_Groups, `n
		If (StrSplit(A_LoopField,",")[1] = groupName && InStr(A_LoopField,groupMember))
			Return True
}

;====================================================================================================