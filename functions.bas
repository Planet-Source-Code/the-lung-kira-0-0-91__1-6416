Attribute VB_Name = "Functions"
Option Explicit


'Only getdata contains a propriatary variable otherwise all are portable

Public Function Asc2Bin(strAscii As String) As String
    Dim strBin As String
    Dim byteLoop As Byte
    Dim tmpByte As Byte
    
    tmpByte = 128 'Set it up
    For byteLoop = 1 To 8 'Cycle through
        If Asc(strAscii) And tmpByte Then
            strBin = strBin & "1"
        Else
            strBin = strBin & "0"
        End If
        
        tmpByte = tmpByte / 2
    Next byteLoop
    
    Asc2Bin = strBin 'Send it back out
End Function

Public Function CharSet(lngCharSet As Long) As String
    Select Case lngCharSet
        Case ANSI_CHARSET: CharSet = "ANSI"
        Case BALTIC_CHARSET: CharSet = "BALTIC"
        Case CHINESEBIG5_CHARSET: CharSet = "CHINESEBIG5"
        Case DEFAULT_CHARSET: CharSet = "DEFAULT"
        Case EASTEUROPE_CHARSET: CharSet = "EASTEUROPE"
        Case GB2312_CHARSET: CharSet = "GB2312"
        Case GREEK_CHARSET: CharSet = "GREEK"
        Case HANGUL_CHARSET: CharSet = "HANGUL"
        Case MAC_CHARSET: CharSet = "MAC"
        Case OEM_CHARSET: CharSet = "OEM"
        Case RUSSIAN_CHARSET: CharSet = "RUSSIAN"
        Case SHIFTJIS_CHARSET: CharSet = "SHIFTJIS"
        Case SYMBOL_CHARSET: CharSet = "SYMBOL"
        Case TURKISH_CHARSET: CharSet = "TURKISH"
        Case HEBREW_CHARSET: CharSet = "HEBREW"
        Case ARABIC_CHARSET: CharSet = "ARABIC"
        Case THAI_CHARSET: CharSet = "THAI"
    End Select
End Function

'Thank ms samples for this one
Public Function CLargeInt(Lo As Long, Hi As Long) As Double
'This function converts the LARGE_INTEGER data type to a double
    
    Dim dblLo As Double, dblHi As Double 'Makes them a double precision number
    
    If Lo < 0 Then
        dblLo = 2 ^ 32 + Lo
    Else
        dblLo = Lo
    End If
    If Hi < 0 Then
        dblHi = 2 ^ 32 + Hi
    Else
        dblHi = Hi
    End If

    CLargeInt = dblLo + dblHi * 2 ^ 32 'Gives out the 32bit dblLo + dblHi as Large Integer
End Function

Public Function Fix_Dir(strDir As String) As String
    If strDir = "" Then Exit Function 'If no data then exit function
    If Len(strDir) < 2 Then Exit Function 'Do not want to have it possibly remove the last char
    
    If Right$(strDir, 1) = "\" Then 'If extra \
        strDir = Left$(strDir, Len(strDir) - 1)
    End If
    
    Fix_Dir = strDir 'Sends back out
End Function

Public Function Fix_NullTermStr(strData As String) As String
    If strData = "" Then Exit Function 'If no data then exit function
    
    If InStr(1, strData, Chr(0)) = 0 Then
        Exit Function 'If thier is no null exit
    Else
        Fix_NullTermStr = Left$(strData, InStr(1, strData, Chr(0)) - 1) '-1 for removing null also
    End If
End Function

Public Function Get_CpuDesc(strTFMS As String) As String
    'Much easier than branching a billion times
    Dim cpu_desc As String
    
    Select Case strTFMS
        Case "0401": cpu_desc = "Intel 486 DX"
        'Case "0402": cpu_desc = "AMD 4x86 & 5x86"
        Case "0411": cpu_desc = "Intel 486 DX"
        Case "0421": cpu_desc = "Intel 486 SX"
        Case "0431": cpu_desc = "Intel 487 / DX2 / DX2 OverDrive"
        Case "0441": cpu_desc = "Intel 486 SL"
        Case "0451": cpu_desc = "Intel SX2"
        Case "0471": cpu_desc = "Write-Back Enhanced Intel DX2"
        Case "0481": cpu_desc = "Intel DX4"
        Case "0481": cpu_desc = "Intel DX4 / DX4 OverDrive"
        Case "0491": cpu_desc = "Write-Back Enhanced Intel DX4"
        Case "0502": cpu_desc = "AMD K5 - Model 0"
        Case "0512": cpu_desc = "AMD K5 - Model 1"
        Case "0513": cpu_desc = "Intel Pentium"
        Case "0522": cpu_desc = "AMD K5 - Model 2"
        Case "0523": cpu_desc = "Intel Pentium"
        Case "0532": cpu_desc = "AMD K5 - Model 3"
        Case "0562": cpu_desc = "AMD K6 - Model 6"
        Case "0572": cpu_desc = "AMD K6 - Model 7"
        Case "0582": cpu_desc = "AMD K6-2 - Model 8"
        Case "0592": cpu_desc = "AMD K6-III - Model 9"
        Case "0612": cpu_desc = "AMD Athlon - Model 1"
        Case "0613": cpu_desc = "Intel Pentium Pro"
        Case "0622": cpu_desc = "AMD Athlon - Model 2"
        
        Case "1523": cpu_desc = "Intel Pentium OverDrive For Pentium"
        Case "1533": cpu_desc = "Intel Pentium OverDrive For 486"
        Case "1543": cpu_desc = "Intel Pentium OverDrive For Pentium"
        Case "1553": cpu_desc = "Intel Pentium OverDrive For 486"
        Case "1581": cpu_desc = "Intel DX4 OverDrive"
        Case "1633": cpu_desc = "Intel Pentium Pro OverDrive For Pentium Pro"
        
        Case Else: cpu_desc = "Unknown"
    End Select
    
    Get_CpuDesc = cpu_desc 'Send back out
End Function

Public Function Get_Data(pathFile As String, lngLines As Long, aryData() As String)
    ReDim Preserve aryData(lngLines) 'Resizes array without destroying
    Dim tmpLong As Long
    Dim tmpString As String
    
    If lngLines = 0 Then 'If no limit then
        Open pathFile For Input As #1
            Do While Not EOF(1) 'Loop until end of file
                Line Input #1, tmpString 'Read line into variable
                tmpLong = tmpLong + 1 'Increment
                
                ReDim Preserve aryData(tmpLong)
                aryData(tmpLong) = Trim(tmpString)
            Loop
        Close #1
        
        WindowKillerNum = tmpLong 'Propriatary
        Exit Function 'Exit out
    End If
    
    Open pathFile For Input As #1
        Do While Not EOF(1) 'Loop until end of file
            Line Input #1, tmpString 'Read line into variable
            tmpLong = tmpLong + 1 'Increment
            
            aryData(tmpLong) = Trim(tmpString)

            If tmpLong = lngLines Then Exit Do 'Exit loop after retriving lines
        Loop
    Close #1
End Function

Function Hex2Bin(strHex As String) As String
    Dim tmpInt As Integer
    Dim tmpString As String
    
    For tmpInt = 0 To 255 'Cycle through
        If Hex(tmpInt) = UCase(strHex) Then 'Gotta be one of them
            Hex2Bin = Asc2Bin(Chr(tmpInt)) 'Send it back out
            Exit For 'Exit here
        End If
    Next tmpInt
End Function

Public Function LangIdent(strCode As String) As String
    If strCode = "" Then Exit Function 'If no data then exit function

    Select Case strCode
        Case "00000000": LangIdent = "Language Neutral"
        Case "00000400": LangIdent = "Process Default Language"
        Case "00000436": LangIdent = "Afrikaans"
        Case "0000041c": LangIdent = "Albanian"
        Case "00000401": LangIdent = "Arabic (Saudi Arabia)"
        Case "00000801": LangIdent = "Arabic (Iraq)"
        Case "00000c01": LangIdent = "Arabic (Egypt)"
        Case "00001001": LangIdent = "Arabic (Libya)"
        Case "00001401": LangIdent = "Arabic (Algeria)"
        Case "00001801": LangIdent = "Arabic (Morocco)"
        Case "00001c01": LangIdent = "Arabic (Tunisia)"
        Case "00002001": LangIdent = "Arabic (Oman)"
        Case "00002401": LangIdent = "Arabic (Yemen)"
        Case "00002801": LangIdent = "Arabic (Syria)"
        Case "00002c01": LangIdent = "Arabic (Jordan)"
        Case "00003001": LangIdent = "Arabic (Lebanon)"
        Case "00003401": LangIdent = "Arabic (Kuwait)"
        Case "00003801": LangIdent = "Arabic (U.A.E.)"
        Case "00003c01": LangIdent = "Arabic (Bahrain)"
        Case "00004001": LangIdent = "Arabic (Qatar)"
        Case "0000042b": LangIdent = "Armenian"
        Case "0000044d": LangIdent = "Assamese"
        Case "0000042c": LangIdent = "Azeri (Latin)"
        Case "0000082c": LangIdent = "Azeri (Cyrillic)"
        Case "0000042d": LangIdent = "Basque"
        Case "00000423": LangIdent = "Belarussian"
        Case "00000445": LangIdent = "Bengali"
        Case "00000402": LangIdent = "Bulgarian"
        Case "00000455": LangIdent = "Burmese"
        Case "00000403": LangIdent = "Catalan"
        Case "00000404": LangIdent = "Chinese (Taiwan)"
        Case "00000804": LangIdent = "Chinese (PRC)"
        Case "00000c04": LangIdent = "Chinese (Hong Kong SAR, PRC)"
        Case "00001004": LangIdent = "Chinese (Singapore)"
        Case "00001404": LangIdent = "Chinese (Macau SAR)"
        Case "0000041a": LangIdent = "Croatian"
        Case "00000405": LangIdent = "Czech"
        Case "00000406": LangIdent = "Danish"
        Case "00000413": LangIdent = "Dutch (Netherlands)"
        Case "00000813": LangIdent = "Dutch (Belgium)"
        Case "00000409": LangIdent = "English (United States)"
        Case "00000809": LangIdent = "English (United Kingdom)"
        Case "00000c09": LangIdent = "English (Australian)"
        Case "00001009": LangIdent = "English (Canadian)"
        Case "00001409": LangIdent = "English (New Zealand)"
        Case "00001809": LangIdent = "English (Ireland)"
        Case "00001c09": LangIdent = "English (South Africa)"
        Case "00002009": LangIdent = "English (Jamaica)"
        Case "00002409": LangIdent = "English (Caribbean)"
        Case "00002809": LangIdent = "English (Belize)"
        Case "00002c09": LangIdent = "English (Trinidad)"
        Case "00003009": LangIdent = "English (Zimbabwe)"
        Case "00003409": LangIdent = "English (Philippines)"
        Case "00000425": LangIdent = "Estonian"
        Case "00000438": LangIdent = "Faeroese"
        Case "00000429": LangIdent = "Farsi"
        Case "0000040b": LangIdent = "Finnish"
        Case "0000040c": LangIdent = "French (Standard)"
        Case "0000080c": LangIdent = "French (Belgian)"
        Case "00000c0c": LangIdent = "French (Canadian)"
        Case "0000100c": LangIdent = "French (Switzerland)"
        Case "0000140c": LangIdent = "French (Luxembourg)"
        Case "0000180c": LangIdent = "French (Monaco)"
        Case "0000043c": LangIdent = "Gaelic - Scotland"
        Case "00000437": LangIdent = "Georgian"
        Case "00000407": LangIdent = "German (Standard)"
        Case "00000807": LangIdent = "German (Switzerland)"
        Case "00000c07": LangIdent = "German (Austria)"
        Case "00001007": LangIdent = "German (Luxembourg)"
        Case "00001407": LangIdent = "German (Liechtenstein)"
        Case "00000408": LangIdent = "Greek"
        Case "00000447": LangIdent = "Gujarati"
        Case "0000040d": LangIdent = "Hebrew"
        Case "00000439": LangIdent = "Hindi"
        Case "0000040e": LangIdent = "Hungarian"
        Case "0000040f": LangIdent = "Icelandic"
        Case "00000421": LangIdent = "Indonesian"
        Case "00000410": LangIdent = "Italian (Standard)"
        Case "00000810": LangIdent = "Italian (Switzerland)"
        Case "00000411": LangIdent = "Japanese"
        Case "0000044b": LangIdent = "Kannada"
        Case "00000860": LangIdent = "Kashmiri (India)"
        Case "0000043f": LangIdent = "Kazakh"
        Case "00000457": LangIdent = "Konkani"
        Case "00000412": LangIdent = "Korean"
        Case "00000812": LangIdent = "Korean (Johab)"
        Case "00000426": LangIdent = "Latvian"
        Case "00000427": LangIdent = "Lithuanian"
        Case "00000827": LangIdent = "Lithuanian (Classic)"
        Case "0000042f": LangIdent = "Macedonian"
        Case "0000043e": LangIdent = "Malay (Malaysian)"
        Case "0000083e": LangIdent = "Malay (Brunei Darussalam)"
        Case "0000044c": LangIdent = "Malayalam"
        Case "0000043a": LangIdent = "Maltese"
        Case "00000458": LangIdent = "Manipuri"
        Case "0000044e": LangIdent = "Marathi"
        Case "00000861": LangIdent = "Nepali (India)"
        Case "00000414": LangIdent = "Norwegian (Bokmal)"
        Case "00000814": LangIdent = "Norwegian (Nynorsk)"
        Case "00000448": LangIdent = "Oriya"
        Case "00000415": LangIdent = "Polish"
        Case "00000416": LangIdent = "Portuguese (Brazil)"
        Case "00000816": LangIdent = "Portuguese (Standard)"
        Case "00000446": LangIdent = "Punjabi"
        Case "00000417": LangIdent = "Raeto-Romance"
        Case "00000418": LangIdent = "Romanian"
        Case "00000818": LangIdent = "Romanian - Moldova"
        Case "00000419": LangIdent = "Russian"
        Case "00000819": LangIdent = "Russian - Moldova"
        Case "0000044f": LangIdent = "Sanskrit"
        Case "00000c1a": LangIdent = "Serbian (Cyrillic)"
        Case "0000081a": LangIdent = "Serbian (Latin)"
        Case "00000459": LangIdent = "Sindhi"
        Case "0000041b": LangIdent = "Slovak"
        Case "00000424": LangIdent = "Slovenian"
        Case "0000042e": LangIdent = "Sorbian"
        Case "0000040a": LangIdent = "Spanish (Traditional Sort)"
        Case "0000080a": LangIdent = "Spanish (Mexican)"
        Case "00000c0a": LangIdent = "Spanish (Modern Sort)"
        Case "0000100a": LangIdent = "Spanish (Guatemala)"
        Case "0000140a": LangIdent = "Spanish (Costa Rica)"
        Case "0000180a": LangIdent = "Spanish (Panama)"
        Case "00001c0a": LangIdent = "Spanish (Dominican Republic)"
        Case "0000200a": LangIdent = "Spanish (Venezuela)"
        Case "0000240a": LangIdent = "Spanish (Colombia)"
        Case "0000280a": LangIdent = "Spanish (Peru)"
        Case "00002c0a": LangIdent = "Spanish (Argentina)"
        Case "0000300a": LangIdent = "Spanish (Ecuador)"
        Case "0000340a": LangIdent = "Spanish (Chile)"
        Case "0000380a": LangIdent = "Spanish (Uruguay)"
        Case "00003c0a": LangIdent = "Spanish (Paraguay)"
        Case "0000400a": LangIdent = "Spanish (Bolivia)"
        Case "0000440a": LangIdent = "Spanish (El Salvador)"
        Case "0000480a": LangIdent = "Spanish (Honduras)"
        Case "00004c0a": LangIdent = "Spanish (Nicaragua)"
        Case "0000500a": LangIdent = "Spanish (Puerto Rico)"
        Case "00000430": LangIdent = "Sutu"
        Case "00000441": LangIdent = "Swahili (Kenya)"
        Case "0000041d": LangIdent = "Swedish"
        Case "0000081d": LangIdent = "Swedish (Finland)"
        Case "00000449": LangIdent = "Tamil"
        Case "00000444": LangIdent = "Tatar (Tatarstan)"
        Case "0000044a": LangIdent = "Telugu"
        Case "0000041e": LangIdent = "Thai"
        Case "00000431": LangIdent = "Tsonga"
        Case "0000041f": LangIdent = "Turkish"
        Case "00000422": LangIdent = "Ukrainian"
        Case "00000420": LangIdent = "Urdu (Pakistan)"
        Case "00000820": LangIdent = "Urdu (India)"
        Case "00000443": LangIdent = "Uzbek (Latin)"
        Case "00000843": LangIdent = "Uzbek (Cyrillic)"
        Case "0000042a": LangIdent = "Vietnamese"
        Case "00000434": LangIdent = "Xhosa"
        Case "0000043d": LangIdent = "Yiddish"
        Case "00000435": LangIdent = "Zulu"
    End Select
End Function

Public Function MpgTagGenre(byteCode As Byte) As String
    Select Case byteCode
        Case 0: MpgTagGenre = "Blues"
        Case 1: MpgTagGenre = "Classic Rock"
        Case 2: MpgTagGenre = "Country"
        Case 3: MpgTagGenre = "Dance"
        Case 4: MpgTagGenre = "Disco"
        Case 5: MpgTagGenre = "Funk"
        Case 6: MpgTagGenre = "Grunge"
        Case 7: MpgTagGenre = "Hip-Hop"
        Case 8: MpgTagGenre = "Jazz"
        Case 9: MpgTagGenre = "Metal"
        Case 10: MpgTagGenre = "New Age"
        Case 11: MpgTagGenre = "Oldies"
        Case 12: MpgTagGenre = "Other"
        Case 13: MpgTagGenre = "Pop"
        Case 14: MpgTagGenre = "R&B"
        Case 15: MpgTagGenre = "Rap"
        Case 16: MpgTagGenre = "Reggae"
        Case 17: MpgTagGenre = "Rock"
        Case 18: MpgTagGenre = "Techno"
        Case 19: MpgTagGenre = "Industrial"
        Case 20: MpgTagGenre = "Alternative"
        Case 21: MpgTagGenre = "Ska"
        Case 22: MpgTagGenre = "Death Metal"
        Case 23: MpgTagGenre = "Pranks"
        Case 24: MpgTagGenre = "Soundtrack"
        Case 25: MpgTagGenre = "Euro-Techno"
        Case 26: MpgTagGenre = "Ambient"
        Case 27: MpgTagGenre = "Trip-Hop"
        Case 28: MpgTagGenre = "Vocal"
        Case 29: MpgTagGenre = "Jazz+Funk"
        Case 30: MpgTagGenre = "Fusion"
        Case 31: MpgTagGenre = "Trance"
        Case 32: MpgTagGenre = "Classical"
        Case 33: MpgTagGenre = "Instrumental"
        Case 34: MpgTagGenre = "Acid"
        Case 35: MpgTagGenre = "House"
        Case 36: MpgTagGenre = "Game"
        Case 37: MpgTagGenre = "Sound Clip"
        Case 38: MpgTagGenre = "Gospel"
        Case 39: MpgTagGenre = "Noise"
        Case 40: MpgTagGenre = "AlternRock"
        Case 41: MpgTagGenre = "Bass"
        Case 42: MpgTagGenre = "Soul"
        Case 43: MpgTagGenre = "Punk"
        Case 44: MpgTagGenre = "Space"
        Case 45: MpgTagGenre = "Meditative"
        Case 46: MpgTagGenre = "Instrumental Pop"
        Case 47: MpgTagGenre = "Instrumental Rock"
        Case 48: MpgTagGenre = "Ethnic"
        Case 49: MpgTagGenre = "Gothic"
        Case 50: MpgTagGenre = "Darkwave"
        Case 51: MpgTagGenre = "Techno-Industrial"
        Case 52: MpgTagGenre = "Electronic"
        Case 53: MpgTagGenre = "Pop-Folk"
        Case 54: MpgTagGenre = "Eurodance"
        Case 55: MpgTagGenre = "Dream"
        Case 56: MpgTagGenre = "Southern Rock"
        Case 57: MpgTagGenre = "Comedy"
        Case 58: MpgTagGenre = "Cult"
        Case 59: MpgTagGenre = "Gangsta"
        Case 60: MpgTagGenre = "Top 40"
        Case 61: MpgTagGenre = "Christian Rap"
        Case 62: MpgTagGenre = "Pop/Funk"
        Case 63: MpgTagGenre = "Jungle"
        Case 64: MpgTagGenre = "Native American"
        Case 65: MpgTagGenre = "Cabaret"
        Case 66: MpgTagGenre = "New Wave"
        Case 67: MpgTagGenre = "Psychadelic"
        Case 68: MpgTagGenre = "Rave"
        Case 69: MpgTagGenre = "Showtunes"
        Case 70: MpgTagGenre = "Trailer"
        Case 71: MpgTagGenre = "Lo-Fi"
        Case 72: MpgTagGenre = "Tribal"
        Case 73: MpgTagGenre = "Acid Punk"
        Case 74: MpgTagGenre = "Acid Jazz"
        Case 75: MpgTagGenre = "Polka"
        Case 76: MpgTagGenre = "Retro"
        Case 77: MpgTagGenre = "Musical"
        Case 78: MpgTagGenre = "Rock & Roll"
        Case 79: MpgTagGenre = "Hard Rock"
        Case 80: MpgTagGenre = "Unknown"
    End Select
End Function

'Not my function
Public Function QuickSort(varData As Variant, Low As Long, Hi As Long)
    'Note: I start my arrays with one and not zero.

    If Not IsArray(varData) Then Exit Function 'If not array exit function

    Dim lngTmpLow As Long
    Dim lngTmpHi As Long
    Dim lngTmpMid As Long
    Dim varTempVal As Variant
    Dim varTmpHold As Variant
  
    lngTmpLow = Low
    lngTmpHi = Hi

    If Hi <= Low Then Exit Function 'If nothing to sort exit function
    lngTmpMid = (Low + Hi) \ 2 'Find the middle to start comparing values
    
    'Move the item in the middle of the array to the temp holding area as a point of ref while sorting. Changes each time a recursive call is made to this routine.
    varTempVal = varData(lngTmpMid)
    
    Do While (lngTmpLow <= lngTmpHi) 'Loop until they meet in the middle
        'Loop as long the array data element is less than data in the temporary holding area and the temporary low value is less than maximum number of array elements.
        Do While (varData(lngTmpLow) < varTempVal And lngTmpLow < Hi)
            lngTmpLow = lngTmpLow + 1
        Loop

        'Loop as long the data in the temporary holding area is less than array data element and the temporary high value is greater than minimum number of array elements.
        Do While (varTempVal < varData(lngTmpHi) And lngTmpHi > Low)
            lngTmpHi = lngTmpHi - 1
        Loop

        'if the temp low end is less than or equal
        'to the temp high end, then swap places
        If (lngTmpLow <= lngTmpHi) Then
            varTmpHold = varData(lngTmpLow)          ' Move the Low value to Temp Hold
            varData(lngTmpLow) = varData(lngTmpHi)     ' Move the high value to the low
            varData(lngTmpHi) = varTmpHold           ' move the Temp Hod to the High
            lngTmpLow = lngTmpLow + 1              ' Increment the temp low counter
            lngTmpHi = lngTmpHi - 1                ' Dcrement the temp high counter
        End If
    Loop
          
    'If the minimum number of elements in the array is less than the temp high end, then make a recursive call to this routine.
    If (Low < lngTmpHi) Then
        QuickSort varData, Low, lngTmpHi
    End If
          
    'If the temp low end is less than the maximum number of elements in the array, then make a recursive call to this routine.
    If (lngTmpLow < Hi) Then
        QuickSort varData, lngTmpLow, Hi
    End If
End Function

Public Function Read_MpgInfo(strFileName As String, MpgInfo As MpgInfo)
    Dim tmpString As String
    Dim tmpBin As String
    
    Open strFileName For Binary As #1 'Opens it for binary
        tmpString = Space(4) '4 bytes
        Get #1, , tmpString 'Dumps contents of file to string
    Close #1
    
    'Setup the binary
    tmpBin = tmpBin & Asc2Bin(Mid$(tmpString, 1, 1))
    tmpBin = tmpBin & Asc2Bin(Mid$(tmpString, 2, 1))
    tmpBin = tmpBin & Asc2Bin(Mid$(tmpString, 3, 1))
    tmpBin = tmpBin & Asc2Bin(Mid$(tmpString, 4, 1))

    With MpgInfo
        'Select Case Mid(tmpBin, 1, 11) 'Sync
        Select Case Mid$(tmpBin, 12, 2)
            Case "00": .Version = "2.5"
            Case "01": 'Reserved
            Case "10": .Version = "2"
            Case "11": .Version = "1"
        End Select
        Select Case Mid$(tmpBin, 14, 2)
            'Case "00": 'Reserved
            Case "01": .Layer = 3
            Case "10": .Layer = 2
            Case "11": .Layer = 1
        End Select
        Select Case Mid$(tmpBin, 16, 1)
            Case "0": .Error_Protection = 1
            Case "1": .Error_Protection = 0
        End Select
        Select Case Mid$(tmpBin, 17, 4)
            Case "0000" '1
                .Bitrate_Index = 0
            Case "0001" '2
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 32
                    If .Layer = 2 Then .Bitrate_Index = 32
                    If .Layer = 3 Then .Bitrate_Index = 32
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 32
                    If .Layer = 2 Then .Bitrate_Index = 32
                    If .Layer = 3 Then .Bitrate_Index = 8
                End If
            Case "0010" '3
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 64
                    If .Layer = 2 Then .Bitrate_Index = 48
                    If .Layer = 3 Then .Bitrate_Index = 40
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 64
                    If .Layer = 2 Then .Bitrate_Index = 48
                    If .Layer = 3 Then .Bitrate_Index = 16
                End If
            Case "0011" '4
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 96
                    If .Layer = 2 Then .Bitrate_Index = 56
                    If .Layer = 3 Then .Bitrate_Index = 48
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 96
                    If .Layer = 2 Then .Bitrate_Index = 56
                    If .Layer = 3 Then .Bitrate_Index = 24
                End If
            Case "0100" '5
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 128
                    If .Layer = 2 Then .Bitrate_Index = 64
                    If .Layer = 3 Then .Bitrate_Index = 56
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 128
                    If .Layer = 2 Then .Bitrate_Index = 64
                    If .Layer = 3 Then .Bitrate_Index = 32
                End If
            Case "0101" '6
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 160
                    If .Layer = 2 Then .Bitrate_Index = 80
                    If .Layer = 3 Then .Bitrate_Index = 64
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 160
                    If .Layer = 2 Then .Bitrate_Index = 80
                    If .Layer = 3 Then .Bitrate_Index = 64
                    If .Version = "2.5" Then .Bitrate_Index = 40
                End If
            Case "0110" '7
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 192
                    If .Layer = 2 Then .Bitrate_Index = 96
                    If .Layer = 3 Then .Bitrate_Index = 80
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 192
                    If .Layer = 2 Then .Bitrate_Index = 96
                    If .Layer = 3 Then .Bitrate_Index = 80
                    If .Version = "2.5" Then .Bitrate_Index = 48
                End If
            Case "0111" '8
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 224
                    If .Layer = 2 Then .Bitrate_Index = 112
                    If .Layer = 3 Then .Bitrate_Index = 96
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 224
                    If .Layer = 2 Then .Bitrate_Index = 112
                    If .Layer = 3 Then .Bitrate_Index = 56
                End If
            Case "1000" '9
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 256
                    If .Layer = 2 Then .Bitrate_Index = 128
                    If .Layer = 3 Then .Bitrate_Index = 112
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 256
                    If .Layer = 2 Then .Bitrate_Index = 128
                    If .Layer = 3 Then .Bitrate_Index = 64
                End If
            Case "1001" '10
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 288
                    If .Layer = 2 Then .Bitrate_Index = 160
                    If .Layer = 3 Then .Bitrate_Index = 128
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 288
                    If .Layer = 2 Then .Bitrate_Index = 160
                    If .Layer = 3 Then .Bitrate_Index = 128
                    If .Version = "2.5" Then .Bitrate_Index = 80
                End If
            Case "1010" '11
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 320
                    If .Layer = 2 Then .Bitrate_Index = 192
                    If .Layer = 3 Then .Bitrate_Index = 160
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 320
                    If .Layer = 2 Then .Bitrate_Index = 192
                    If .Layer = 3 Then .Bitrate_Index = 160
                    If .Version = "2.5" Then .Bitrate_Index = 96
                End If
            Case "1011" '12
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 352
                    If .Layer = 2 Then .Bitrate_Index = 224
                    If .Layer = 3 Then .Bitrate_Index = 192
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 352
                    If .Layer = 2 Then .Bitrate_Index = 224
                    If .Layer = 3 Then .Bitrate_Index = 112
                End If
            Case "1100" '13
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 384
                    If .Layer = 2 Then .Bitrate_Index = 256
                    If .Layer = 3 Then .Bitrate_Index = 224
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 384
                    If .Layer = 2 Then .Bitrate_Index = 256
                    If .Layer = 3 Then .Bitrate_Index = 128
                End If
            Case "1101" '14
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 416
                    If .Layer = 2 Then .Bitrate_Index = 320
                    If .Layer = 3 Then .Bitrate_Index = 256
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 416
                    If .Layer = 2 Then .Bitrate_Index = 320
                    If .Layer = 3 Then .Bitrate_Index = 256
                    If .Version = "2.5" Then .Bitrate_Index = 144
                End If
            Case "1110" '15
                If .Version = 1 Then
                    If .Layer = 1 Then .Bitrate_Index = 448
                    If .Layer = 2 Then .Bitrate_Index = 384
                    If .Layer = 3 Then .Bitrate_Index = 320
                Else '2
                    If .Layer = 1 Then .Bitrate_Index = 448
                    If .Layer = 2 Then .Bitrate_Index = 384
                    If .Layer = 3 Then .Bitrate_Index = 320
                    If .Version = "2.5" Then .Bitrate_Index = 160
                End If
            'Case "1111" '16
        End Select
        Select Case Mid$(tmpBin, 21, 2)
            Case "00"
                If .Version = "1" Then .Sampling_Freq = 44100
                If .Version = "2" Then .Sampling_Freq = 22050
                If .Version = "2.5" Then .Sampling_Freq = 11025
            Case "01"
                If .Version = "1" Then .Sampling_Freq = 48000
                If .Version = "2" Then .Sampling_Freq = 24000
                If .Version = "2.5" Then .Sampling_Freq = 12000
            Case "10"
                If .Version = "1" Then .Sampling_Freq = 32000
                If .Version = "2" Then .Sampling_Freq = 16000
                If .Version = "2.5" Then .Sampling_Freq = 8000
            'Case "11" 'Reserved
        End Select
        'Select Case Mid$(tmpBin, 23, 1) 'Padding
        'Select Case Mid$(tmpBin, 24, 1) 'Ignored
        Select Case Mid$(tmpBin, 25, 2)
            Case "00": .Mode = "Stereo"
            Case "01": .Mode = "Joint Stereo"
            Case "10": .Mode = "Dual Channel"
            Case "11": .Mode = "Single Channel"
        End Select
        'Select Case Mid$(tmpBin, 27, 2) 'Used with "joint stereo" mode
        Select Case Mid$(tmpBin, 29, 1)
            Case "0": .Copyright = 0
            Case "1": .Copyright = 1
        End Select
        Select Case Mid$(tmpBin, 30, 1)
            Case "0": .Original = 0
            Case "1": .Original = 1
        End Select
        Select Case Mid$(tmpBin, 31, 2)
            Case "00": .Emphasis = "None"
            Case "01": .Emphasis = "50/15 Microsec"
            Case "10": .Emphasis = "Reserved"
            Case "11": .Emphasis = "CCITT J. 17"
        End Select
    End With
End Function

Public Function Read_MpgTag(strFileName As String, MpgTag As MpgTag)
    Dim tmpString As String
    
    Open strFileName For Binary As #1 'Opens it for binary
        tmpString = Space(128) '128 bytes
        Get #1, (LOF(1) - 127), tmpString 'Dumps contents of file to string
    Close #1
    
    If Left(tmpString, 3) = "TAG" Then 'Tag was found
        'Fills tag info
        MpgTag.Tag = True
        MpgTag.Title = Mid(tmpString, 4, 30)
        MpgTag.Artist = Mid(tmpString, 34, 30)
        MpgTag.Album = Mid(tmpString, 64, 30)
        MpgTag.Year = Mid(tmpString, 94, 4)
        MpgTag.Comments = Mid(tmpString, 98, 30)
        MpgTag.Genre = Asc(Mid(tmpString, 128, 1))
    Else 'No tag was found
        MpgTag.Tag = False
        Exit Function
    End If
End Function

Public Function Rem_NonStd_Chr(strData As String) As String
    If strData = "" Then Exit Function 'If no data then exit function
    
    Dim tmpInt As Integer
    
    'Removes nonstandard characters not allowed in
    For tmpInt = 0 To 32 'Defines nonstandard characters that are not allowed
        strData = Replace(strData, Chr(tmpInt), "", 1, -1) 'Removes them from string
    Next tmpInt
    For tmpInt = 42 To 44 'Defines nonstandard characters that are not allowed
        strData = Replace(strData, Chr(tmpInt), "", 1, -1) 'Removes them from string
    Next tmpInt
    For tmpInt = 58 To 63 'Defines nonstandard characters that are not allowed
        strData = Replace(strData, Chr(tmpInt), "", 1, -1) 'Removes them from string
    Next tmpInt
    For tmpInt = 91 To 93 'Defines nonstandard characters that are not allowed
        strData = Replace(strData, Chr(tmpInt), "", 1, -1) 'Removes them from string
    Next tmpInt
    For tmpInt = 128 To 255 'Defines nonstandard characters that are not allowed
        strData = Replace(strData, Chr(tmpInt), "", 1, -1) 'Removes them from string
    Next tmpInt
    
    'Removes extra characters not in sequence
    strData = Replace(strData, Chr(34), "", 1, -1)
    strData = Replace(strData, Chr(47), "", 1, -1)
    strData = Replace(strData, Chr(96), "", 1, -1)
    strData = Replace(strData, Chr(124), "", 1, -1)
    
    Rem_NonStd_Chr = strData 'Sends it back out
End Function

Public Function Rem_NonFat_Chr(strData As String) As String
    If strData = "" Then Exit Function 'If no data then exit function

    'Removes nonfat characters not allowed in
    strData = Replace(strData, "*", "", 1, -1)
    strData = Replace(strData, "?", "", 1, -1)
    strData = Replace(strData, "/", "", 1, -1)
    strData = Replace(strData, "\", "", 1, -1)
    strData = Replace(strData, "|", "", 1, -1)
    strData = Replace(strData, ".", "", 1, -1)
    strData = Replace(strData, ",", "", 1, -1)
    strData = Replace(strData, ";", "", 1, -1)
    strData = Replace(strData, ":", "", 1, -1)
    strData = Replace(strData, "+", "", 1, -1)
    strData = Replace(strData, "=", "", 1, -1)
    strData = Replace(strData, " ", "", 1, -1)
    strData = Replace(strData, "[", "", 1, -1)
    strData = Replace(strData, "]", "", 1, -1)
    strData = Replace(strData, "(", "", 1, -1)
    strData = Replace(strData, ")", "", 1, -1)
    strData = Replace(strData, "&", "", 1, -1)
    strData = Replace(strData, "^", "", 1, -1)
    strData = Replace(strData, "<", "", 1, -1)
    strData = Replace(strData, ">", "", 1, -1)
    strData = Replace(strData, Chr(34), "", 1, -1)
    
    Rem_NonFat_Chr = strData 'Sends it back out
End Function

Public Function Unicode_Padd(strData As String) As String
    If strData = "" Then Exit Function 'If no data then exit function
    If Len(strData) < 2 Then Exit Function 'Can not pad in between 1 char
    
    'Dont declare variables until checked for data
    Dim lenData As Long
    Dim tmpString As String
    
    For lenData = 1 To Len(strData) 'Cycle through the data
        tmpString = tmpString & Mid$(strData, lenData, 1) & Chr(0) 'Add nulls between every character
    Next lenData
    
    Unicode_Padd = Left(tmpString, Len(tmpString) - 1) 'Sends it back without trailing null
End Function

Public Function Write_MpgTag(strFileName As String, MpgTag As MpgTag)
    Dim tmpString As String
    
    'Combine structure to string
    With MpgTag
        tmpString = "TAG"
        tmpString = tmpString & .Title
        tmpString = tmpString & .Artist
        tmpString = tmpString & .Album
        tmpString = tmpString & .Year
        tmpString = tmpString & .Comments
        tmpString = tmpString & Chr(.Genre)
    End With
    
    Open strFileName For Binary As #1
        Put #1, (LOF(1) - 127), tmpString
    Close #1
End Function
