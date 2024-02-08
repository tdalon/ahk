
/*  https://www.autohotkey.com/boards/viewtopic.php?t=75317
PathX() splits a filepath and returns filepath parts as an associative array with following keys:
  1) Drive
  2) Dir
  3) Fname (Filename without extension)
  4) Ext
  5) Folder (Concatenation of 1 & 2)
  6) File (Concatenation of 3 & 4)
  7) Full (Concatenation of 1,2,3 & 4) 
*/

PathX(S, P*) {                                ; PathX v0.67 by SKAN on D34U/D35I @ tiny.cc/pathx
Local K,V,N,  U:={},   T:=Format(A_IsUnicode ? "{1:260}" : "{1:520}", ""),   dr:=di:=fn:=ex := T

  For K,V in P
      N := StrSplit(V,":",,2),  K := SubStr(N[1],1,2),  U[K] := N[2]

  DllCall("GetFullPathName", "Str",Trim(S,Chr(34)), "UInt",260, "Str",T, "Ptr",0)
  DllCall("msvcrt\_wsplitpath", "WStr",T, "WStr",dr, "WStr",di, "WStr",fn, "WStr",ex)

Return { "Drive"  : dr := u.dr ? u.dr : dr 
       , "Dir"    : di := (u.dp) (u.di="" ? di : u.di) (u.ds) 
       , "Fname"  : fn := (u.fp) (u.fn="" ? fn : u.fn) (u.fs)     
       , "Ext"    : ex := u._e!="" ? (ex ? ex : u._e) : u.ex="" ? ex : u.ex        
       , "Folder" : (dr) (di)
       , "File"   : (fn) (ex)
       , "Full"   : (u.pp) (dr) (di) (fn) (ex) (u.ps) }
}