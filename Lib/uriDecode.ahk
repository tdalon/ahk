; -------------------------------------------------------------------------------------------------------------------
; https://www.autohotkey.com/boards/viewtopic.php?t=5520
uriDecode(Uri, Encoding := "UTF-8") { ; Encoding must be either "UTF-16" or "CP0"
   If (Encoding <> "UTF-8") && (Encoding <> "CP0")
      Encoding := "UTF-8"
   Split := StrSplit(Uri)
   Length := Split.MaxIndex()
   VarSetCapacity(AStr, Length, 0)
   Index := 1
   Addr := &AStr
   While (Index <= Length) {
      If (Split[Index] <> "%")
         Addr := NumPut(Asc(Split[Index]), Addr + 0, "UChar")
      Else
         Addr := NumPut("0x" . Split[++Index] . Split[++Index], Addr + 0, "UChar")
      Index++
   }
   ; TD: encode %2f into urlsep
   sDecoded := StrGet(&AStr, Encoding)
   sDecoded := StrReplace(sDecoded,"%2F","/")
   return sDecoded
}


; Test 
; https://tenantname.sharepoint.com/:p:/r/teams/team_10000778/Shared%20Documents/Explore/New%20Work%20Style%20%E2%80%93%20O365%20-%20Why%20using%20Teams.pptx?d=we1a512b97ed844fc92dd5a1d028ef827&csf=1&e=crdehv
