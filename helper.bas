Attribute VB_Name = "helper"

Option Explicit

Function bytes2string(bytes() As Byte) As String
    Dim i As Integer
    On Error GoTo hell
    bytes2string = ""
    
    For i = LBound(bytes) To UBound(bytes)
        bytes2string = bytes2string + Chr(bytes(i))
    Next
    Exit Function
hell:
    bytes2string = "-"
End Function

Function bytes2stringHuman(bytes() As Byte, Optional bWithWhitechars = False) As String
    Dim i As Integer
    Dim c As String
    Dim b As Byte
    
    On Error GoTo hell
    'bytes2string = ""
    
    For i = LBound(bytes) To UBound(bytes)
        b = Chr(bytes(i))
        
        If Not bWithWhitechars Then
            Select Case b
            Case Chr(10), Chr(13)     'cr,lf
                b = "."
            End Select
        End If
        Select Case b
        Case " " To Chr(255)
        Case Chr(10), Chr(13)     'cr,lf
            'nop
        Case Else
            b = "."
        End Select
        bytes2stringHuman = bytes2stringHuman + b
    Next
    Exit Function
hell:
    bytes2stringHuman = "-"
End Function
Function strip0s(bytes() As Byte) As Byte()
    Dim i As Integer
    Dim res() As Byte
    On Error GoTo hell
    
    For i = LBound(bytes) To UBound(bytes)
        If bytes(i) > 0 Then
            ReDim res(helper.fUbound(res) + 1)
            res(fUbound(res)) = bytes(i)
        End If
    Next
    strip0s = res
    Exit Function
    
hell:
    strip0s = res
End Function
Function bytes2hexString(bytes() As Byte) As String
    Dim i As Integer
    On Error GoTo hell
    
    For i = LBound(bytes) To UBound(bytes)
        If (bytes(i) < &H10) Then bytes2hexString = bytes2hexString + "0"
        bytes2hexString = bytes2hexString + Hex(bytes(i)) + " "
    Next
    bytes2hexString = Trim(bytes2hexString)
    Exit Function
hell:
    bytes2hexString = "-"
End Function

Function string2hex(ByVal text As String) As String
    string2hex = ""
    Dim ch As Integer
    Dim i As Integer
    
    For i = 1 To Len(text)
        ch = Asc(Mid(text, i, 1))
        If ch < &H10 Then string2hex = string2hex + "0"
        string2hex = string2hex + Hex(ch) + " "
    Next
    
    Trim (string2hex)
End Function

Function hex2string(text) As String
    Dim i As Integer
    Dim ones() As String
    
    ones = Split(Trim(text), " ")
    
    For i = LBound(ones) To UBound(ones)
        hex2string = hex2string + Chr("&h" + ones(i))
    Next
End Function

Function hexString2bytes(text) As Byte()
    Dim i As Integer
    Dim ones() As String
    On Error GoTo hell
    Dim bytes() As Byte
    
    While (InStr(text, "  "))   '//remove double spaces
        text = Replace(text, "  ", " ")
    Wend
    
    ones = Split(Trim(text), " ")
    ReDim bytes(UBound(ones))
    
    For i = LBound(ones) To UBound(ones)
       bytes(i) = Val("&h" + ones(i))
    Next
    
    hexString2bytes = bytes
    Exit Function
hell:
    ReDim hexString2bytes(1)
End Function


Function binString2bytes(text) As Byte()
    Dim i As Integer
    Dim ones() As String
    Dim bytes() As Byte
    
    On Error GoTo hell
    
    'ones = Split(Trim(text), " ")
    'ReDim bytes(UBound(ones))
    
    'For i = LBound(ones) To UBound(ones)
    '   bytes = ones(i)
    'Next
    
    ReDim bytes(Len(text) - 1)
    
    For i = LBound(bytes) To UBound(bytes)
       bytes(i) = Asc(Mid(text, 1 + i, 1) + Chr(0))
    Next
    
    binString2bytes = bytes
    Exit Function
hell:
    ReDim binString2bytes(1)
End Function


Function fitPattern(filterPattern As String, text As String) As Boolean
    '? = any char
    On Error GoTo hell:
    
    Dim i As Integer
    
    fitPattern = True ' recessive optimistic
    
    filterPattern = LCase(filterPattern)
    text = LCase(text)
    
    For i = 1 To Len(filterPattern)
        If Mid(filterPattern, i, 1) = "*" Then Exit Function
        
        If i > Len(text) Then
            fitPattern = False
            Exit Function
        End If
        
        If Mid(filterPattern, i, 1) <> "?" Then
            If Mid(filterPattern, i, 1) <> Mid(text, i, 1) Then
                fitPattern = False
                Exit Function
            End If
        End If
    Next i

    'no "*" found, so rest is important
    If Len(text) > Len(filterPattern) Then fitPattern = False
    
    Exit Function
hell:
    fitPattern = False
End Function

Function hex4(value As Variant) As String
On Error GoTo hell:
    hex4 = Hex(value)
    While Len(hex4) < 4
        hex4 = "0" + hex4
    Wend
    Exit Function
hell:
    hex4 = "----"
End Function

Function hexN(ByVal value As Variant, ByVal leng As Integer) As String
    Dim v2 As Variant
On Error GoTo hell:
    'problems with values > 0x7ffffff
    hexN = ""
    'leng = leng / 2
    If value > &H7FFFFFFF Then
        v2 = value
        'v2 = value Mod &H100   'geht ned:overflow
        value = Int(value / &H100)       'remember this
        v2 = v2 - (value * &H100)   'mod = lowest byte
        leng = leng - 2
        hexN = hexN + hex2(v2 And &HFF)
    End If
    While leng > 0
        hexN = hex2(value And &HFF) + hexN
        value = value / &H100
        leng = leng - 2
    Wend
    hexN = "&h" + hexN
    Exit Function
hell:
    hexN = "-"
End Function

Function hex2(value As Variant) As String
On Error GoTo hell:
    hex2 = Hex(Int(value))
    While Len(hex2) < 2
        hex2 = "0" + hex2
    Wend
    Exit Function
hell:
    hex2 = "--"
End Function


Function midBytes(bytes() As Byte, pos As Integer, Optional leng As Integer = 32767) As Byte()
    Dim res() As Byte
    
    If leng < 1 Then Exit Function
    If fUbound(bytes) < leng - pos Then leng = fUbound(bytes) - pos
    If fUbound(bytes) < 0 Then Exit Function
    
    ReDim res(leng - 1)
    
    While leng > 0
        leng = leng - 1
        res(leng) = bytes(pos + leng)
    Wend
    midBytes = res
End Function

Function addArray(a1() As Byte, a2() As Byte) As Byte()
    Dim b() As Byte
    Dim i As Integer
    Dim oldi As Integer
    
    
    If (fUbound(a2)) < 0 Then
        addArray = a1
        Exit Function
    End If
    If fUbound(a1) < 0 Then
        'ReDim a1(0)
        addArray = a2
        Exit Function
    End If
    ReDim b(fUbound(a1) - LBound(a1) + fUbound(a2) - LBound(a2) + 1)
    
    If (UBound(a1) - LBound(a1)) > 0 Then
   
        For i = 0 To UBound(a1) - LBound(a1)
            b(i) = a1(LBound(a1) + i)
        Next
    End If
    oldi = i
    For i = 0 To UBound(a2) - LBound(a2)
        b(i + oldi) = a2(LBound(a2) + i)
    Next
    addArray = b
End Function

Function arrayExist(arr As Variant) As Boolean
    On Error GoTo hell
    If UBound(arr) <> -1 Then arrayExist = True
    Exit Function
hell:
    arrayExist = False
End Function

Function fUbound(arr As Variant) As Integer
    On Error GoTo hell
    fUbound = UBound(arr)
    Exit Function
hell:
    fUbound = -1
End Function

Function getBytesFromHex4(htext As String) As Variant
    Dim bytes(3) As Byte
    Dim pos As Integer
'    If Left(htext, 1) = "&" Then
'        htext = Mid(htext, 2)
'    End If
    
'    If Left(htext, 1) = "h" Then
'        htext = Mid(htext, 2)
'    End If
    
    bytes(0) = Val("&h" + Right(htext, 2))
    pos = Len(htext) - 3
    If Len(htext) > 3 Then bytes(1) = Val("&h" + Mid(htext, pos, 2))
    pos = pos - 2
    If Len(htext) > 5 Then bytes(2) = Val("&h" + Mid(htext, pos, 2))
    pos = pos - 2
    If Len(htext) > 7 Then bytes(3) = Val("&h" + Mid(htext, pos, 2))
    
    
    getBytesFromHex4 = bytes
End Function
Function getU32FromBytes(bytes() As Byte) As Variant
    Dim lMul As Variant
    Dim max As Integer
    Dim pos As Integer
    
    lMul = 1
    max = UBound(bytes) - LBound(bytes)
    If max > 3 Then max = 3
    For pos = LBound(bytes) To LBound(bytes) + max
        getU32FromBytes = getU32FromBytes + bytes(pos) * lMul
        If pos < UBound(bytes) Then lMul = lMul * &H100&
    Next pos
End Function

Function getU16FromBytes(bytes() As Byte) As Variant
    Dim max As Integer
    
    max = UBound(bytes) - LBound(bytes)
    If max < 1 Then Exit Function
    
    getU16FromBytes = bytes(LBound(bytes))
    getU16FromBytes = getU16FromBytes Or 256 * bytes(LBound(bytes) + 1)
End Function

Function LongToUShort(Unsigned As Long) As Integer
    LongToUShort = CInt(Unsigned - &H10000)
End Function

Function UShortToLong(Unsigned As Integer) As Long
    UShortToLong = Unsigned
    If UShortToLong < 0 Then
        UShortToLong = UShortToLong + &H10000
    End If
End Function

Function TextToLines(text As String) As String()
    Dim res() As String
    res = Split(text, vbCrLf)
    TextToLines = res
End Function
