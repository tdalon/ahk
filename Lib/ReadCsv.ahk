ReadCsv(CsvFile,InProp,InVal,OutProp){
;    OutProp := ReadCsv(FilePath,InProp,InVal,OutProp)

RowMatch := 0
Loop, read, %CsvFile%
{
    RowCount := A_Index
    Loop, parse, A_LoopReadLine, CSV
    {
        If (RowCount = 1) { ; first row= header 
            If (A_LoopField == InProp) 
                InCol := A_Index   

            If (A_LoopField == OutProp) 
                OutCol := A_Index
                           
            If (!OutCol OR !InCol) 
                Continue
            Else
                break   
        }

        If (A_Index = InCol) {
            If (A_LoopField == InVal) {
                RowMatch := RowCount
                If (InCol > OutCol)
                    return OutVal
            }
        }
           
        If (A_Index = OutCol) {
            OutVal := A_LoopField
            If (RowMatch>0) ; not empty
                return OutVal
        }    
}
}
} ; end of function