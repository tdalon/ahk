b64Encode(string)
; ref: https://github.com/jNizM/AHK_Scripts/blob/master/src/encoding_decoding/base64.ahk
{
    VarSetCapacity(bin, StrPut(string, "UTF-8")) && len := StrPut(string, &bin, "UTF-8") - 1 
    if !(DllCall("crypt32\CryptBinaryToString", "ptr", &bin, "uint", len, "uint", 0x40000001, "ptr", 0, "uint*", size))
        throw Exception("CryptBinaryToString failed", -1)
    VarSetCapacity(buf, size << 1, 0)
    if !(DllCall("crypt32\CryptBinaryToString", "ptr", &bin, "uint", len, "uint", 0x40000001, "ptr", &buf, "uint*", size))
        throw Exception("CryptBinaryToString failed", -1)
    return StrGet(&buf)
}

; bug: remove linebreaks fixed