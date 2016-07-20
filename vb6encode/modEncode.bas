Attribute VB_Name = "modEncode"
Option Explicit
Private Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_ACP As Long = 0
Private Const CP_UTF8 As Long = 65001
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Private m_AES As New AES

Private Const C_ERROR_NUMBER As Long = 10000
Private Const G_ModName As String = "modEncode.bas"

'*************************************************************
'Procedure:    Public Method UrlDecode_Utf8
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       Str
'Returns:
'Remarks:
'*************************************************************
Function UrlDecode_Utf8(ByVal Str As String) As String
    Dim b, ub   '中文字的Unicode码(2字节)
    Dim UtfB    'Utf-8单个字节
    Dim UtfB1, UtfB2, UtfB3 'Utf-8码的三个字节
    Dim i As Integer, n As Integer, s As String
    n = 0
    ub = 0
    For i = 1 To Len(Str)
        b = Mid(Str, i, 1)
        Select Case b
            Case "+"
                s = s & " "
            Case "%"
                ub = Mid(Str, i + 1, 2)
                UtfB = CInt("&H" & ub)
                If UtfB < 128 Then
                    i = i + 2
                    s = s & ChrW(UtfB)
                Else
                    UtfB1 = (UtfB And &HF) * &H1000   '取第1个Utf-8字节的二进制后4位
                    UtfB2 = (CInt("&H" & Mid(Str, i + 4, 2)) And &H3F) * &H40      '取第2个Utf-8字节的二进制后6位
                    UtfB3 = CInt("&H" & Mid(Str, i + 7, 2)) And &H3F      '取第3个Utf-8字节的二进制后6位
                    s = s & ChrW(UtfB1 Or UtfB2 Or UtfB3)
                    i = i + 8
                End If
            Case Else    'Ascii码
                s = s & b
        End Select
    Next
    UrlDecode_Utf8 = s
End Function

'*************************************************************
'Procedure:    Public Method UrlEncode_Utf8
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       Str
'Returns:
'Remarks:
'*************************************************************
Public Function UrlEncode_Utf8(ByVal Str As String) As String
    Dim wch, uch, szRet As String
    Dim x
    Dim nAsc, nAsc2, nAsc3
    If Str = "" Then
        UrlEncode_Utf8 = Str
        Exit Function
    End If
    For x = 1 To Len(Str)
        wch = Mid(Str, x, 1)
        nAsc = AscW(wch)
        
        If (nAsc >= 48 And nAsc <= 57) Or (nAsc >= 65 And nAsc <= 90) Or (nAsc >= 97 And nAsc <= 122) Or nAsc = 42 Or nAsc = 45 Or nAsc = 46 Or nAsc = 64 Or nAsc = 95 Then
            ''48 to 57代表0~9;65 to 90代表A~Z;97 to 122代表a~z
            ''42代表*;46代表.;64代表@;45代表-;95代表_
            szRet = szRet & wch
        ElseIf nAsc = 32 Then ''空格转成+
            szRet = szRet & "+"
        ElseIf nAsc < 128 And nAsc > 0 Then ''低于128的Ascii转成1个字节
            szRet = szRet & "%" & Right("00" & Hex(nAsc), 2)
        Else
            If nAsc < 0 Then nAsc = nAsc + 65536
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    UrlEncode_Utf8 = szRet
End Function

'*************************************************************
'Procedure:    Public Method URLEncode
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       strURL
'Returns:
'Remarks:
'*************************************************************
Function URLEncode(strURL As String) As String
    Dim i As Integer
    Dim tempStr As String
    Dim nAsc As Integer
    For i = 1 To Len(strURL)
        nAsc = Asc(Mid(strURL, i, 1))
        If nAsc < 0 Then
            tempStr = "%" & Right(CStr(Hex(nAsc)), 2)
            tempStr = "%" & Left(CStr(Hex(nAsc)), Len(CStr(Hex(nAsc))) - 2) & tempStr
            URLEncode = URLEncode & tempStr
        ElseIf (nAsc >= 48 And nAsc <= 57) Or (nAsc >= 65 And nAsc <= 90) Or (nAsc >= 97 And nAsc <= 122) Or nAsc = 42 Or nAsc = 45 Or nAsc = 46 Or nAsc = 64 Or nAsc = 95 Then
            ''48 to 57代表0~9;65 to 90代表A~Z;97 to 122代表a~z
            ''42代表*;46代表.;64代表@;45代表-;95代表_
            URLEncode = URLEncode & Mid(strURL, i, 1)
        ElseIf nAsc = 32 Then ''空格转成+
            URLEncode = URLEncode & "+"
        Else
            URLEncode = URLEncode & "%" & Right("00" & Hex(nAsc), 2)
        End If
    Next
End Function

'*************************************************************
'Procedure:    Public Method URLDecode
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       strURL
'Returns:
'Remarks:
'*************************************************************
Function URLDecode(strURL As String) As String
    Dim i As Integer

    If InStr(strURL, "%") = 0 Then
        URLDecode = strURL
        Exit Function
    End If

    For i = 1 To Len(strURL)
        If Mid(strURL, i, 1) = "%" Then
            If Val("&H" & Mid(strURL, i + 1, 2)) > 127 Then
                URLDecode = URLDecode & Chr(Val("&H" & Mid(strURL, i + 1, 2) & Mid(strURL, i + 4, 2)))
                i = i + 5
            Else
                URLDecode = URLDecode & Chr(Val("&H" & Mid(strURL, i + 1, 2)))
                i = i + 2
            End If
        Else
            URLDecode = URLDecode & Mid(strURL, i, 1)
        End If
    Next
End Function

'*************************************************************
'Procedure:    Public Method Base64Decode
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       B64
'Returns:
'Remarks:
'*************************************************************
Function Base64Decode(B64 As String) As Byte()                                  'Base64 解码
On Error GoTo over                                                             '排错
    Dim OutStr() As Byte, i As Long, j As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    If InStr(1, B64, "=") <> 0 Then B64 = Left(B64, InStr(1, B64, "=") - 1)     '判断Base64真实长度,除去补位
    Dim Length As Long, mods As Long
    mods = Len(B64) Mod 4
    Length = Len(B64) - mods
    ReDim OutStr(Length / 4 * 3 - 1 + Switch(mods = 0, 0, mods = 2, 1, mods = 3, 2))
    For i = 1 To Length Step 4
        Dim buf(3) As Byte
        For j = 0 To 3
            buf(j) = InStr(1, B64_CHAR_DICT, Mid(B64, i + j, 1)) - 1            '根据字符的位置取得索引值
        Next
        OutStr((i - 1) / 4 * 3) = buf(0) * &H4 + (buf(1) And &H30) / &H10
        OutStr((i - 1) / 4 * 3 + 1) = (buf(1) And &HF) * &H10 + (buf(2) And &H3C) / &H4
        OutStr((i - 1) / 4 * 3 + 2) = (buf(2) And &H3) * &H40 + buf(3)
    Next
    If mods = 2 Then
        OutStr(Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, Length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 2, 1)) - 1) And &H30) / 16
    ElseIf mods = 3 Then
        OutStr(Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, Length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 2, 1)) - 1) And &H30) / 16
        OutStr(Length / 4 * 3 + 1) = ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 2, 1)) - 1) And &HF) * &H10 + ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 3, 1)) - 1) And &H3C) / &H4
    End If
    Base64Decode = OutStr                                                       '读取解码结果
over:
End Function

'*************************************************************
'Procedure:    Public Method Base64Encode
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       Str(
'Returns:
'Remarks:
'*************************************************************
Function Base64Encode(Str() As Byte) As String                                  'Base64 编码
    On Error GoTo over                                                          '排错
    Dim buf() As Byte, Length As Long, mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    mods = (UBound(Str) + 1) Mod 3   '除以3的余数
    Length = UBound(Str) + 1 - mods
    ReDim buf(Length / 3 * 4 + IIf(mods <> 0, 4, 0) - 1)
    Dim i As Long
    For i = 0 To Length - 1 Step 3
        buf(i / 3 * 4) = (Str(i) And &HFC) / &H4
        buf(i / 3 * 4 + 1) = (Str(i) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
        buf(i / 3 * 4 + 2) = (Str(i + 1) And &HF) * &H4 + (Str(i + 2) And &HC0) / &H40
        buf(i / 3 * 4 + 3) = Str(i + 2) And &H3F
    Next
    If mods = 1 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10
        buf(Length / 3 * 4 + 2) = 64
        buf(Length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10 + (Str(Length + 1) And &HF0) / &H10
        buf(Length / 3 * 4 + 2) = (Str(Length + 1) And &HF) * &H4
        buf(Length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(buf)
        Base64Encode = Base64Encode + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
    Next
over:
End Function

'*************************************************************
'Procedure:    Public Method ToUTF8Bytes
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       sData
'Returns:
'Remarks:
'*************************************************************
'字符转 UTF8
Public Function ToUTF8Bytes(ByVal sData As String) As Byte()
    Dim aRetn() As Byte
    Dim nSize As Long
    nSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sData), -1, 0, 0, 0, 0) - 1
    If nSize = 0 Then Exit Function
    ReDim aRetn(0 To nSize - 1) As Byte
    WideCharToMultiByte CP_UTF8, 0, StrPtr(sData), -1, VarPtr(aRetn(0)), nSize, 0, 0
    ToUTF8Bytes = aRetn
    Erase aRetn
End Function

'*************************************************************
'Procedure:    Public Method FromUTF8Bytes
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       sData
'Returns:
'Remarks:
'*************************************************************
'' UTF8 转字符
Public Function FromUTF8Bytes(ByVal sData As String) As Byte()
    Dim aRetn() As Byte
    Dim nSize As Long
    nSize = MultiByteToWideChar(CP_UTF8, 0, StrPtr(sData), -1, 0, 0) - 1
    If nSize = 0 Then Exit Function
    ReDim aRetn(0 To 2 * nSize - 1) As Byte
    MultiByteToWideChar CP_UTF8, 0, StrPtr(sData), -1, VarPtr(aRetn(0)), nSize
    FromUTF8Bytes = aRetn
    Erase aRetn
End Function

'*************************************************************
'Procedure:    Public Method BytesToHex
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       bytB(
'Returns:
'Remarks:
'*************************************************************
Public Function BytesToHex(bytB() As Byte) As String
    Dim strTmp As String, i As Long
    For i = 0 To UBound(bytB)
        strTmp = strTmp & " " & Hex(bytB(i))
    Next
    BytesToHex = strTmp
End Function

'*************************************************************
'Procedure:    Public Method utf8AESBase64
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       source
'       key
'Returns:
'Remarks:
'*************************************************************
Public Function utf8AESBase64(source As String, key As String) As String
    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long
    
    KeyBits = 128
    BlockBits = 128
    
    pass = ToUTF8Bytes(key)
    ReDim Preserve pass(31)
    
    plaintext = ToUTF8Bytes(source)

    m_AES.SetCipherKey pass, KeyBits
    m_AES.ArrayEncrypt plaintext, ciphertext, 0

    utf8AESBase64 = Base64Encode(ciphertext)

End Function

'*************************************************************
'Procedure:    Public Method utf8AESBase64dec
'Description:
'Created:      2016-7-19 by lizhipeng
'Parameters:
'       cipherStr
'       key
'Returns:
'Remarks:
'*************************************************************
Public Function utf8AESBase64dec(cipherStr As String, key As String) As String
On Error GoTo ErrHandler

    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    KeyBits = 128
    BlockBits = 128

    pass = ToUTF8Bytes(key)
    ReDim Preserve pass(31)

    ciphertext = Base64Decode(cipherStr)

    m_AES.SetCipherKey pass, KeyBits
    m_AES.ArrayDecrypt plaintext, ciphertext, 0

    utf8AESBase64dec = FromUTF8Bytes(plaintext)

Exit Function
ErrHandler:
    Err.Raise C_ERROR_NUMBER, G_ModName & "-" & "utf8AESBase64dec", Err.source & vbCrLf & Err.Description
End Function

