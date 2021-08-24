SM_Solve(expression, ahk=false){
static fchars := "e- e+ **- ** **+ ^- ^+ + - * / \" , rchars := "#< #> ^< ^> ^> ^< ^> ¢ ¤ ¥ ¦ ¦"
if expression=
return
if expression is alpha
{
temp2 := %expression%
return %temp2%
}
else if expression is number
{
if !Instr(expression, "e")
return expression
}
StringReplace,expression,expression,%A_space%,,All
StringReplace,expression,expression,%A_tab%,,All
expression := SM_Fixexpression(expression)
while b_pos := RegexMatch(expression, "i)[\+\-\*\\\/\^]\(")
{
b_count := {"(": 1, ")": 0}
b_temp := Substr(expression, b_pos+2)
loop, parse, b_temp
{
b_count[A_LoopField] += 1
if b_count["("] = b_count[")"]
{
end_pos := A_index
break
}
}
expression := Substr(expression, 1, b_pos) SM_Solve( Substr(expression, b_pos+2, end_pos-1) ) Substr(expression, end_pos+b_pos+2)
}
expression := SM_FixExpression(expression)
loop,
{
if !(Instr(expression, "(")){
expression := SM_PowerReplace(expression, fchars, rchars, "All")
reserve .= expression
break
}
temp := Substr(expression, 1, Instr(expression, "("))
temp := SM_PowerReplace(temp, fchars, rchars, "All")
temp2 := SubStr(expression, Instr(expression, "(") + 1, Instr(expression, ")") - Instr(expression, "("))
reserve .= temp . temp2
expression := Substr(expression,Instr(expression, ")")+ 1)
}
expression := reserve
loop, parse, expression,¢¤¥¦
{
if RegExMatch(A_LoopField, "iU)^[a-z0-9_]+\(.*\)$")
{
fname := Substr(A_LoopField, 1, Instr(A_loopfield,"(") - 1)
ffeed := Substr(A_loopfield, Instr(A_loopfield, "(") + 1, Instr(A_loopfield, ")") - Instr(A_loopfield, "(") - 1)
loop, parse, ffeed,`,
{
StringReplace,feed,A_loopfield,",,All
feed%A_index% := SM_Solve(feed)
totalfeeds := A_index
}
if fname = SM_toExp
outExp := 1
, number := feed1
else if totalfeeds = 1
number := %fname%(feed1)
else if totalfeeds = 2
number := %fname%(feed1, feed2)
else if totalfeeds = 3
number := %fname%(feed1, feed2, feed3)
else if totalfeeds = 4
number := %fname%(feed1, feed2, feed3, feed4)
function := 1
}
else
number := A_LoopField , function := 0
if (char != "") {
if (!function) {
while match_pos := RegExMatch(number, "iU)%.*%", output_var)
output_var := Substr(output_var, 2 , -1)
, number := Substr(number, 1, match_pos-1)   SM_Solve(Instr(output_var, "(") ? output_var : %output_var%)   Substr(number, match_pos+Strlen(output_var)+2)
if Instr(number, "#") or Instr(number, "^")
number := SM_PowerReplace(number, "#< #> ^> ^<", "e- e ^ ^-", "All")
if ( p := Instr(number, "c") ) or ( p := p + Instr(number, "p") )
term_n := Substr(number, 1, p-1) , term_r := Substr(number,p+1)
, number := SM_Solve( term_n "!/" term_r "!" ( Instr(number, "c") ? "/(" term_n "-" term_r ")!" : "" ) )
else if Instr(number, "^")
number := SM_Pow( SM_Solve( SubStr(number, 1, posofpow := Instr(number, "^")-1 ) )   ,   SM_Solve( Substr(number, posofpow+2) ) )
else if Instr(number, "!")
number := SM_fact( SM_Solve( Substr(number, 1, -1) ) )
else if Instr(number, "e")
number := SM_fromExp( number )
}
if (Ahk){
if char = ¢
solved := solved + (number)
else if char = ¤
solved := solved - (number)
else if char = ¦
{
if !number
return
solved := solved / (number)
}
else if char = ¥
solved := solved * (number)
}else{
if char = ¢
solved := SM_Add(solved, number)
else if char = ¤
solved := SM_Add(solved,"-" . number)
else if char = ¦
{
if !number
return
solved := SM_Divide(solved, number)
}
else if char = ¥
solved := SM_Multiply(solved, number)
}
}
if solved =
solved := number
char := Substr(expression, Strlen(A_loopfield) + 1,1)
expression := Substr(expression, Strlen(A_LoopField) + 2)
}
return, outExp ? SM_ToExp( solved ) : SM_Prefect( solved )
}
SM_Add(number1, number2, prefect=true){
IfInString,number2,--
count := 2
else IfInString,number2,-
count := 1
else
count := 0
IfInString,number1,-
count+=1
n1 := number1
n2 := number2
StringReplace,number1,number1,-,,All
StringReplace,number2,number2,-,,All
dec1 := Instr(number1,".") ? StrLen(number1) - InStr(number1, ".") : 0
dec2 := Instr(number2,".") ? StrLen(number2) - InStr(number2, ".") : 0
if (dec1 > dec2){
dec := dec1
loop,% (dec1 - dec2)
number2 .= "0"
}
else if (dec2 > dec1){
dec := dec2
loop,% (dec2 - dec1)
number1 .= "0"
}
else
dec := dec1
StringReplace,number1,number1,.
StringReplace,number2,number2,.
if (Strlen(number1) >= StrLen(number2)){
loop,% (Strlen(number1) - strlen(number2))
number2 := "0" . number2
}
else
loop,% (Strlen(number2) - strlen(number1))
number1 := "0" . number1
n := strlen(number1)
if count not in 1,3
{
loop,
{
digit := SubStr(number1,1 - A_Index, 1) + SubStr(number2, 1 - A_index, 1) + (carry ? 1 : 0)
if (A_index == n){
sum := digit . sum
break
}
if (digit > 9){
carry := true
digit := SubStr(digit, 0, 1)
}
else
carry := false
sum := digit . sum
}
if (Instr(n2,"-") and Instr(n1, "-"))
sum := "-" . sum
}
elsE
{
numbercompare := SM_Greater(number1, number2, true)
if !(numbercompare){
mid := number2
number2 := number1
number1 := mid
}
loop,
{
digit := SubStr(number1,1 - A_Index, 1) - SubStr(number2, 1 - A_index, 1) + (borrow ? -1 : 0)
if (A_index == n){
StringReplace,digit,digit,-
sum := digit . sum
break
}
if Instr(digit, "-")
borrow:= true , digit := 10 + digit
else
borrow := false
sum := digit sum
}
If InStr(n2,"--"){
if (numbercompare)
sum := "-" . sum
}else If InStr(n2,"-"){
if !(numbercompare)
sum := "-" . sum
}else IfInString,n1,-
if (numbercompare)
sum := "-" . sum
}
if ((sum == "-"))
sum := 0
If (dec)
if (sum)
sum := SubStr(sum,1,StrLen(sum) - dec) . "." . SubStr(sum,1 - dec)
return, Prefect ? SM_Prefect(sum) : sum
}
SM_Multiply(number1, number2){
positive := true
if Instr(number2, "-")
positive := false
if Instr(number1, "-")
positive := !positive
number1 := Substr(number1, Instr(number1, "-") ? 2 : 1)
number2 := Substr(number2, Instr(number2, "-") ? 2 : 1)
dec := InStr(number1,".") ? StrLen(number1) - InStr(number1, ".") : 0
If n2dotpos := Instr(number2, ".")
dec := dec + StrLen(number2) - n2dotpos
StringReplace,number1,number1,.
StringReplace,number2,number2,.
loop,% Strlen(number2)
number2temp .= Substr(number2, 1-A_Index, 1)
number2 := number2temp
product := "0"
Loop,parse,number2
{
row := "0"
zeros := ""
if (A_loopfield)
loop,% (A_loopfield)
row := SM_Add(row, number1, 0)
else
loop,% (Strlen(number1) - 1)
row .= "0"
loop,% (A_index - 1)
zeros .= "0"
row .= zeros
product := SM_Add(product, row, false)
}
if (dec){
product := SubStr(product,1,StrLen(product) - dec) . "." . SubStr(product,1 - dec)
product := SM_Prefect(product)
}
if !(positive)
product := "-" . product
return, product
}
SM_Divide(number1, number2, length=10){
positive := true
if (Instr(number2, "-"))
positive := false
if (Instr(number1, "-"))
positive := !positive
StringReplace,number1,number1,-
StringReplace,number2,number2,-
number1 := SM_Prefect(number1) , number2 := SM_Prefect(number2)
if !number2
return
dec := 0
if Instr(number1, ".")
dec := - (Strlen(number1) - Instr(number1, "."))
if Instr(number2, ".")
dec := Strlen(number2) - Instr(number2, ".") + dec + 0
StringReplace,number1,number1,.
StringReplace,number2,number2,.
number1 := Ltrim(number1, "0") , number2 := Ltrim(number2, "0")
decimal := dec , num1 := number1 , num2 := number2
n1 := Strlen(number1) , n2 := StrLen(number2)
loop,% n2 + length
number1 .= "0"
coveredlength := 0 , dec := dec - n2 - length , takeone := false , n1f := n1 + n2 + length
while(number1 != "")
{
times := 0 , below := "" , lendivide := 0 , n1fromleft := (takeone) ? Substr(number1, 1, n2+1) : Substr(number1, 1, n2)
if SM_Greater(n1fromleft, number2, true)
{
todivide := n1fromleft
loop, 10
{
num2temp%A_index% := SM_Multiply(number2, A_index)
if !(SM_Greater(todivide, num2temp%A_index%, true)){
lendivide := (takeone) ? n2 + 1 : n2
times := A_index - 1 , below := num2temp%times%
break
}
}
res .= zeroes_r
}
else
{
todivide := SubStr(number1, 1, n2+1)
loop, 10
{
num2temp%A_index% := SM_Multiply(number2, A_index)
if !(SM_Greater(todivide, num2temp%A_index%, true)){
lendivide := n2 + 1
times := A_index - 1 , below := num2temp%times%
break
}
}
if (coveredlength != 0)
res .= zeroes_r "0"
}
res .= times , coveredlength+=(lendivide - Strlen(remainder))
remainder := SM_Add(todivide, "-" below)
if remainder = 0
remainder := ""
number1 := remainder . Substr(number1, lendivide + 1)
if SM_Greater("0", remainder, true)
{
zeroes_k := ""
loop,% Strlen(number1)
zeroes_k .= "0"
if (number1 == zeroes_k){
StringTrimRight,number1,number1,1
number1 := "1" . number1
res := SM_Multiply(res, number1)
break
}
}
if times = 0
break
zeroes_r := "" , takeone := false
if (remainder == "") {
loop,
if (Instr(number1, "0") == 1)
zeroes_r .= "0" , number1 := Substr(number1, 2) , coveredlength+=1
else
break
}
if (Strlen(remainder) == n2)
takeone := true
else
loop,% n2 - StrLen(remainder) - 1
zeroes_r .= "0"
}
if (dec < 0)
{
oldformat := A_formatfloat
SetFormat,float,0.16e
Divi := Substr(num1,1,15) / Substr(num2,1,15)
decimal := decimal + Strlen(Substr(num1,16)) - Strlen(Substr(num2,16))
if (Instr(divi,"-"))
decimal := decimal - Substr(divi,-1) + 1
else
decimal := decimal + Substr(divi,-1) + 1
if (decimal > 0)
res := Substr(res, 1, decimal) . "." . Substr(res, decimal + 1)
else if (decimal < 0){
loop,% Abs(decimal)
zeroes_e .= "0"
res := "0." zeroes_e res
}
else
res := "0." res
SetFormat,float,%oldformat%
}
else
{
num := "1"
loop,% dec
num .= "0"
res := SM_Multiply(SM_Prefect(res), num)
}
return, ( (positive) ? "" : "-" ) . SM_Round(SM_Prefect(res), decimal < 0 ? Abs(decimal)+length : length)
}
SM_UniquePmt(series, ID="", Delimiter=","){
if Instr(series, Delimiter)
loop, parse, series,%Delimiter%
item%A_index% := A_LoopField , last := lastbk := A_Index
else{
loop, parse, series
item%A_index% := A_loopfield
last := lastbk := Strlen(series) , Delimiter := ""
}
if (ID == "")
{
fact := SM_fact(last)
loop,% fact
toreturn .= SM_UniquePmt(series, A_index) . "`n"
return, Rtrim(toreturn, "`n")
}
posfactor := (SM_Mod(ID, last) == "0") ? last : SM_Mod(ID, last)
incfactor := (SM_Mod(ID, last) == "0") ? SM_Floor(SM_Divide(ID,last)) : SM_Floor(SM_Divide(ID,last)) + 1
loop,% last
{
tempmod := SM_Mod(posfactor + incfactor - 1, last)
posfactor := (tempmod == "0") ? last : tempmod
res .= item%posfactor% . Delimiter , item%posfactor% := ""
loop,% lastbk
if (item%A_index% == "")
plus1 := A_index + 1 , item%A_index% := item%plus1% , item%plus1% := ""
last-=1
if (posfactor > last)
posfactor := 1
}
return, Rtrim(res, Delimiter)
}
SM_Greater(number1, number2, trueforequal=false){
IfInString,number2,-
IfNotInString,number1,-
return, true
IfInString,number1,-
IfNotInString,number2,-
return, false
if (Instr(number1, "-") and Instr(number2, "-"))
bothminus := true
number1 := SM_Prefect(number1) , number2 := SM_Prefect(number2)
dec1 := (Instr(number1,".")) ? ( StrLen(number1) - InStr(number1, ".") ) : (0)
dec2 := (Instr(number2,".")) ? ( StrLen(number2) - InStr(number2, ".") ) : (0)
if (dec1 > dec2)
loop,% (dec1 - dec2)
number2 .= "0"
else if (dec2 > dec1)
loop,% (dec2 - dec1)
number1 .= "0"
StringReplace,number1,number1,.
StringReplace,number2,number2,.
if (Strlen(number1) > Strlen(number2))
return,% (bothminus) ? (false) : (true)
else if (Strlen(number2) > Strlen(number1))
return,% (bothminus) ? (true) : (false)
else
{
stop := StrLen(number1)
loop,
{
if (SubStr(number1,A_Index, 1) > Substr(number2,A_index, 1))
return bothminus ? 0 : 1
else if (Substr(number2,A_index, 1) > SubStr(number1,A_Index, 1))
return bothminus ? 1 : 0
if (a_index == stop)
return, (trueforequal) ? 1 : 0
}
}
}
SM_Prefect(number){
number .= ""
number := RTrim(number, "+-")
if number=
return 0
if Instr(number, "-")
number := Substr(number, 2) , negative := true
if Instr(number, "."){
number := Trim(number, "0")
if (Substr(number,1,1) == ".")
number := "0" number
if (Substr(number, 0) == ".")
number := Substr(number, 1, -1)
return,% (negative) ? ("-" . number) : (number)
}
else
{
if Trim(number, "0")
return negative ? ("-" . Ltrim(number, "0")) : (Ltrim(number, "0"))
else
return 0
}
}
SM_Mod(dividend, divisor){
positive := true
if Instr(divisor, "-")
positive := false
if (Instr(dividend, "-"))
positive := !positive
dividend := Substr(dividend, Instr(dividend, "-") ? 2 : 1) , divisor := Substr(divisor, Instr(divisor, "-") ? 2 : 1) , Remainder := dividend
if SM_Greater(dividend, divisor, true){
div := SM_Divide(dividend, divisor)
div := Instr(div, ".") ? SubStr(div, 1, Instr(div, ".") - 1) : 0
if ( div == "0" )
Remainder := 0
else
Remainder := SM_Add(dividend, "-" SM_Multiply(Divisor, div))
}
return, ( (Positive or Remainder=0) ? "" : "-" ) . Remainder
}
SM_Exp(number, decimals=""){
return SM_ToExp(number, decimals)
}
SM_ToExp(number, decimals=""){
if (dec_pos := Instr(number, "."))
{
number := SM_Prefect(number) , number := Substr(number, Instr(number, "0")=1 ? 2 : 1)
Loop, parse, number
{
if A_loopfield > 0
break
tempnum .= A_LoopField
}
number := Substr(number, Strlen(tempnum)+1) , power := dec_pos-Strlen(tempnum)-2
number2 := Substr(number, 2)
StringReplace,number2,number2,.
number := Substr(number, 1, 1) "." number2
decimals := ( decimals="" or decimals>Strlen(number2) ) ? Strlen(number2) : decimals
return SM_Round(number, decimals) "e" power
}
else
{
number := SM_Prefect(number) , decimals := ( decimals="" or decimals>Strlen(Substr(number,2)) ) ? Strlen(Substr(number,2)) : decimals
return SM_Round( Substr(number, 1, 1) "." Substr(number, 2), decimals ) "e" Strlen(number)-1
}
}
SM_FromExp(expnum){
if !Instr(expnum, "e")
return expnum
n1 := Substr(expnum, 1, t := Instr(expnum, "e")-1) , n2 := Substr(expnum, t+2)
return SM_ShiftDecimal(n1, n2)
}
SM_Round(number, decimals){
if Instr(number,".")
{
nofdecimals := StrLen(number) - ( Instr(number, ".") = 0 ? Strlen(number) : Instr(number, ".") )
if (nofdecimals > decimals)
{
secdigit := Substr(number, Instr(number,".")+decimals+1, 1)
if secdigit >= 5
loop,% decimals-1
zeroes .= "0"
number := SM_Add(Substr(number, 1, Instr(number, ".")+decimals), (secdigit >= 5) ? "0." zeroes "1" : "0")
}
else
{
loop,% decimals - nofdecimals
zeroes .= "0"
number .= zeroes
}
return, Rtrim(number, ".")
}
else
return, number
}
SM_Floor(number){
number := SM_Prefect(number)
if Instr(number, "-")
if Instr(number,".")
return, SM_Add(Substr(number, 1, Instr(number, ".") - 1), -1)
else
return, number
else
return, Instr(number, ".") ? Substr(number, 1, Instr(number, ".") - 1) : number
}
SM_Ceil(number){
number := SM_Prefect(number)
if Instr(number, "-")
{
if Instr(number,".")
return, Substr(number, 1, Instr(number, ".") - 1)
else
return, number
}
else
return, Instr(number, ".") ? SM_Add( Substr(number, 1, Instr(number, ".") - 1), 1) : number
}
SM_fact(N){
res := 1 , k := 1 , carry := 0
N -= 1
loop % N
{
StringSplit, l_times, res
index := l_times0
k := A_index + 1
Loop %index%
{
digit := k * l_times%index% + carry
if ( digit > 9 )
{
carry := RegExReplace(digit, "(.*)(\d)", "$1")
digit := RegExReplace(digit, "(.*)(\d)", "$2")
}
else
carry := 0
r := digit r
index --
}
if ( carry != 0 )
final := carry r
else
final := r
res := final
digit := index := final := r =
r := ""
carry := 0
}
return final ? final : 1
}
SM_Pow(number, power){
if (power < 1)
{
if !power
return 1
if power Between -1 and 1
return number ** power
power := -power , I := Floor(power) , D := Mod(power, 1)
if Instr(number, "-") && D
return
D_part := number ** D
I_part := SM_Pow(number, I)
return SM_Prefect( SM_Divide(1, SM_Multiply(I_part, D_part)) )
}
else
{
if power > 6
{
sqrt_c := Floor(Sqrt(power))
x_c := SM_Iterate(number, sqrt_c) , loopc := Floor(power/sqrt_c)
x_c_loop := SM_iterate(x_c, loopc) , remPow := power - (sqrt_c*loopc)
x_remPow := SM_iterate(number, remPow)
return SM_Multiply(x_c_loop, x_remPow)
}
else x_7_pow7 := 1
a := 1
loop % Mod(power, 7)
a := SM_Multiply(number, a)
return SM_Multiply(x_7_pow7, a)
}
}
SM_e(N, auto=1){
static e := 2.71828182845905 , d := 14
if (N > 5) and auto
e := SM_Round("2.71828182845905", (F := d-N+5)>2 ? F : 2)
return SM_Pow(e, N)
}
SM_Number2Base(N, base=16){
baseLen:=base<10 ? SM_Ceil((10/base)*Strlen(N)) : Strlen(N)
if SM_checkformat(N) && SM_Checkformat(base**(baseLen-1))
loop % baseLen
D:=Floor(N/(T:=base**(baseLen-A_index))),H.=!D?0:(D>9?Chr(D+87):D),N:=N-D*T
else
loop % baseLen
D:=SM_Floor( SM_Divide(N , T:= SM_Pow(base, baselen-A_index)) ) , H.=!D?0:(D>9?Chr(D+87):D) , N:=SM_Add( N, "-" SM_Multiply(D,T) )
return Ltrim(H,"0")
}
SM_Base2Number(H, base=16){
StringLower, H, H
S:=Strlen(H),N:=0
loop,parse,H
N := SM_Add(  N, SM_Multiply( (A_LoopField*1="")?Asc(A_LoopField)-87:A_LoopField , SM_Pow(base, S-A_index) )  )
return N
}
SM_Checkformat(n){
static ahk_ct := 9223372036854775807
if n < 0
return 0
if ( ahk_ct > n+0 )
return 1
}
SM_ShiftDecimal(number, dec_shift=0){
if Instr(number, "-")
number := Substr(number,2) , minus := 1
dec_pos := Instr(number, ".") , numlen := StrLen(number)
loop % Abs(dec_shift)
zeroes .= "0"
if !dec_pos
number .= ".0"
number := dec_shift>0 ? number zeroes : zeroes number
dec_pos := Instr(number, ".")
StringReplace, number, number, % "."
number := Substr(number, 1, dec_pos+dec_shift-1) "." Substr(number, dec_pos+dec_shift)
return ( minus ? "-" : "" ) SM_Prefect(number)
}
SM_Iterate(number, times){
x := 1
loop % times
x := SM_Multiply(x, number)
return x
}
SM_PowerReplace(input, find, replace, options="All"){
StringSplit, rep, replace, % A_space
loop, parse, find, % A_Space
StringReplace, input, input, % A_LoopField, % rep%A_index%, % options
return Input
}
SM_FixExpression(expression){
expression := Rtrim(expression, "+-=*/\^")
StringReplace,expression,expression,--,+,All
StringReplace,expression,expression,-+,-,All
StringReplace,expression,expression,+-,-,All
StringReplace,expression,expression,++,+,All
if expression=
return
if Substr(expression, 1, 1) == "-"
expression := "0" expression
else expression := "+" expression
loop,
{
if Instr(expression, "*-"){
fromleft := Substr(expression, 1, Instr(expression, "*-"))
StringGetPos,posplus,fromleft,+,R
StringGetPos,posminus,fromleft,-,R
if (posplus > posminus)
fromleft := Substr(fromleft, 1, posplus) "-" Substr(fromleft, posplus + 2)
else
fromleft := Substr(fromleft, 1, posminus) "+" Substr(fromleft, posminus + 2)
expression := fromleft . Substr(expression, Instr(expression, "*-") + 2)
}else if Instr(expression, "/-"){
fromleft := Substr(expression, 1, Instr(expression, "/-"))
StringGetPos,posplus,fromleft,+,R
StringGetPos,posminus,fromleft,-,R
if (posplus > posminus)
fromleft := Substr(fromleft, 1, posplus) "-" Substr(fromleft, posplus + 2)
else
fromleft := Substr(fromleft, 1, posminus) "+" Substr(fromleft, posminus + 2)
expression := fromleft . Substr(expression, Instr(expression, "/-") + 2)
}else if Instr(expression, "\-"){
fromleft := Substr(expression, 1, Instr(expression, "\-"))
StringGetPos,posplus,fromleft,+,R
StringGetPos,posminus,fromleft,-,R
if (posplus > posminus)
fromleft := Substr(fromleft, 1, posplus) "-" Substr(fromleft, posplus + 2)
else
fromleft := Substr(fromleft, 1, posminus) "+" Substr(fromleft, posminus + 2)
expression := fromleft . Substr(expression, Instr(expression, "\-") + 2)
}else if Instr(expression, "x-"){
fromleft := Substr(expression, 1, Instr(expression, "x-"))
StringGetPos,posplus,fromleft,+,R
StringGetPos,posminus,fromleft,-,R
if (posplus > posminus)
fromleft := Substr(fromleft, 1, posplus) "-" Substr(fromleft, posplus + 2)
else
fromleft := Substr(fromleft, 1, posminus) "+" Substr(fromleft, posminus + 2)
expression := fromleft . Substr(expression, Instr(expression, "x-") + 2)
}else
break
}
StringReplace,expression,expression,--,+,All
StringReplace,expression,expression,-+,-,All
StringReplace,expression,expression,+-,-,All
StringReplace,expression,expression,++,+,All
return, expression
}
