VERSION 5.00
Begin VB.Form frmCmos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cmos Contents"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "frmCmos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCmos 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmCmos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If Not Dir(Dirs.AppPath & "\cmos.exe") <> "" Then
        MsgBox "Missing cmos.exe", vbExclamation, "Error"
        Exit Sub
    Else
        If ShellExecute(Me.hwnd, "open", "cmos.exe", "", App.Path & "\", SW_HIDE) <= 32 Then
            Errors.Errors apiError, "ShellExecute"
        End If
    End If
    
    Dim tmpLong As Long
    tmpLong = GetTickCount
    Do While Not Dir(Dirs.AppPath & "\cmos.dat") <> ""
        If tmpLong + 5000 < GetTickCount Then Exit Do
        DoEvents
    Loop
    
    Dim tmpInt As Integer
    Dim tmpString As String
    
    If Not Dir(Dirs.AppPath & "\cmos.dat") <> "" Then
        MsgBox "Missing cmos.dat", vbExclamation, "Error"
        Exit Sub
    End If
    
    Open Dirs.AppPath & "\cmos.dat" For Input As #1
        Do While Not EOF(1)
        Input #1, tmpString 'Line dump
        
        Select Case tmpInt
            Case 0: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Seconds" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 1: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Second Alarm" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 2: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Minutes" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 3: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Minute Alarm" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 4: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Hours" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 5: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Hour Alarm" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 6: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Day of Week" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 7: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Day of Month" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 8: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Month" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 9: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Year" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 10 'A
                tmpString = StrReverse(Hex2Bin(tmpString)) 'High bit most signifigant (bit 7 in 0-7)
                
                lstCmos.AddItem ""
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Status Register A" & Space(40), 40) & Right$("00" & tmpString, 2)
                
                'Interrupt select rate
                If Mid$(tmpString, 1, 4) = "0000" Then
                    lstCmos.AddItem Left$("0-3" & Space(6), 6) & Left$("Interrupt Selection Rate" & Space(40), 40) & "None"
                ElseIf Mid$(tmpString, 1, 4) = "0011" Then
                    lstCmos.AddItem Left$("0-3" & Space(6), 6) & Left$("Interrupt Selection Rate" & Space(40), 40) & "122 microseconds"
                ElseIf Mid$(tmpString, 1, 4) = "1111" Then
                    lstCmos.AddItem Left$("0-3" & Space(6), 6) & Left$("Interrupt Selection Rate" & Space(40), 40) & "500 milliseconds"
                ElseIf Mid$(tmpString, 1, 4) = "0110" Then
                    lstCmos.AddItem Left$("0-3" & Space(6), 6) & Left$("Interrupt Selection Rate" & Space(40), 40) & "976.562 microseconds"
                Else
                    lstCmos.AddItem Left$("0-3" & Space(6), 6) & Left$("Interrupt Selection Rate" & Space(40), 40) & ""
                End If
                
                '22 stage divider
                If Mid$(tmpString, 5, 3) = "010" Then
                    lstCmos.AddItem Left$("4-6" & Space(6), 6) & Left$("22 Stage Divider" & Space(40), 40) & "32768hz"
                Else
                    lstCmos.AddItem Left$("4-6" & Space(6), 6) & Left$("22 Stage Divider" & Space(40), 40) & ""
                End If
                
                lstCmos.AddItem Left$("7" & Space(6), 6) & Left$("Time Update Cycle In Progress" & Space(40), 40) & CBool(Mid$(tmpString, 8, 1))
            Case 11 'B
                tmpString = StrReverse(Hex2Bin(tmpString))
                
                lstCmos.AddItem ""
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Status Register B" & Space(40), 40) & Right$("00" & tmpString, 2)
                
                lstCmos.AddItem Left$("0" & Space(6), 6) & Left$("Daylight Savings" & Space(40), 40) & CBool(Mid$(tmpString, 1, 1))
                lstCmos.AddItem Left$("1" & Space(6), 6) & Left$("24 Hour Mode" & Space(40), 40) & CBool(Mid$(tmpString, 2, 1))
                
                'Data mode
                If Mid$(tmpString, 3, 1) = "0" Then
                    lstCmos.AddItem Left$("2" & Space(6), 6) & Left$("Data Mode" & Space(40), 40) & "BCD"
                ElseIf Mid$(tmpString, 3, 1) = "1" Then
                    lstCmos.AddItem Left$("2" & Space(6), 6) & Left$("Data Mode" & Space(40), 40) & "Binary"
                Else
                    lstCmos.AddItem Left$("2" & Space(6), 6) & Left$("Data Mode" & Space(40), 40) & ""
                End If
                
                lstCmos.AddItem Left$("3" & Space(6), 6) & Left$("Square Wave Output" & Space(40), 40) & CBool(Mid$(tmpString, 4, 1))
                lstCmos.AddItem Left$("4" & Space(6), 6) & Left$("Update Ended Interrupt" & Space(40), 40) & CBool(Mid$(tmpString, 5, 1))
                lstCmos.AddItem Left$("5" & Space(6), 6) & Left$("Alarm Interrupt" & Space(40), 40) & CBool(Mid$(tmpString, 6, 1))
                lstCmos.AddItem Left$("6" & Space(6), 6) & Left$("Periodic Interrupt" & Space(40), 40) & CBool(Mid$(tmpString, 7, 1))
                lstCmos.AddItem Left$("7" & Space(6), 6) & Left$("Clock Setting By Freezing Updates" & Space(40), 40) & CBool(Mid$(tmpString, 8, 1))
            Case 12 'C
                tmpString = StrReverse(Hex2Bin(tmpString))
                
                lstCmos.AddItem ""
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Status Register C" & Space(40), 40) & Right$("00" & tmpString, 2)
                
                lstCmos.AddItem Left$("0-3" & Space(6), 6) & Left$("Unused" & Space(40), 40)
                lstCmos.AddItem Left$("4" & Space(6), 6) & Left$("Update Ended Interrupt" & Space(40), 40) & CBool(Mid$(tmpString, 5, 1))
                lstCmos.AddItem Left$("5" & Space(6), 6) & Left$("Alarm Interrupt" & Space(40), 40) & CBool(Mid$(tmpString, 6, 1))
                lstCmos.AddItem Left$("6" & Space(6), 6) & Left$("Periodic Interrupt" & Space(40), 40) & CBool(Mid$(tmpString, 7, 1))
                lstCmos.AddItem Left$("7" & Space(6), 6) & Left$("Interrupt Request" & Space(40), 40) & CBool(Mid$(tmpString, 8, 1))
            Case 13 'D
                tmpString = StrReverse(Hex2Bin(tmpString))
                
                lstCmos.AddItem ""
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Status Register D" & Space(40), 40) & Right$("00" & tmpString, 2)
                
                lstCmos.AddItem Left$("0-6" & Space(6), 6) & Left$("Unused" & Space(40), 40)
                lstCmos.AddItem Left$("7" & Space(6), 6) & Left$("Valid Ram" & Space(40), 40) & CBool(Mid$(tmpString, 8, 1))
                
                lstCmos.AddItem ""
            Case 14 'Diagnostic status byte
                tmpString = StrReverse(Hex2Bin(tmpString))
                
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Diagnostic Status Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
                
                lstCmos.AddItem Left$("0" & Space(6), 6) & Left$("Read Adaptor ID Timeout" & Space(40), 40) & CBool(Mid$(tmpString, 1, 1))
                lstCmos.AddItem Left$("1" & Space(6), 6) & Left$("Adaptor Mismatch Configuration" & Space(40), 40) & CBool(Mid$(tmpString, 2, 1))
                lstCmos.AddItem Left$("2" & Space(6), 6) & Left$("Invalid Time" & Space(40), 40) & CBool(Mid$(tmpString, 3, 1))
                lstCmos.AddItem Left$("3" & Space(6), 6) & Left$("Controller Or Drive Failed Init" & Space(40), 40) & CBool(Mid$(tmpString, 4, 1))
                lstCmos.AddItem Left$("4" & Space(6), 6) & Left$("Memory Size Error" & Space(40), 40) & CBool(Mid$(tmpString, 5, 1))
                lstCmos.AddItem Left$("5" & Space(6), 6) & Left$("Incorrect Equipment Configuration" & Space(40), 40) & CBool(Mid$(tmpString, 6, 1))
                lstCmos.AddItem Left$("6" & Space(6), 6) & Left$("Incorrect Checksum" & Space(40), 40) & CBool(Mid$(tmpString, 7, 1))
                lstCmos.AddItem Left$("7" & Space(6), 6) & Left$("Clock Lost Power" & Space(40), 40) & CBool(Mid$(tmpString, 8, 1))
                
                lstCmos.AddItem ""
            'Case 15 'Fh
            Case 16 'Floppy Drive Type
                If tmpString = "00" Then
                    lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Floppy Drive Type" & Space(40), 40) & "No Drive"
                ElseIf tmpString = "01" Then
                    lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Floppy Drive Type" & Space(40), 40) & "360 KB 5.25 Drive"
                ElseIf tmpString = "02" Then
                    lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Floppy Drive Type" & Space(40), 40) & "1.2 MB 5.25 Drive"
                ElseIf tmpString = "03" Then
                    lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Floppy Drive Type" & Space(40), 40) & "720 KB 3.5 Drive"
                ElseIf tmpString = "04" Then
                    lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Floppy Drive Type" & Space(40), 40) & "1.44 MB 3.5 Drive"
                ElseIf tmpString = "05" Then
                    lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Floppy Drive Type" & Space(40), 40) & "2.88 MB 3.5 drive"
                Else
                    lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Floppy Drive Type" & Space(40), 40) & "Unused"
                End If
            'Case 17 11h
            'Case 18 12h
            'Case 19 13h
            Case 20 'Equipment Byte
                tmpString = StrReverse(Hex2Bin(tmpString))
                
                lstCmos.AddItem ""
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Equipment Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
                
                lstCmos.AddItem Left$("0" & Space(6), 6) & Left$("Floppy Drive Installed" & Space(40), 40) & CBool(Mid$(tmpString, 1, 1))
                lstCmos.AddItem Left$("1" & Space(6), 6) & Left$("Math Coprocessor Installed" & Space(40), 40) & CBool(Mid$(tmpString, 2, 1))
                lstCmos.AddItem Left$("2" & Space(6), 6) & Left$("Keyboard Enabled" & Space(40), 40) & CBool(Mid$(tmpString, 3, 1))
                lstCmos.AddItem Left$("3" & Space(6), 6) & Left$("Display Enabled" & Space(40), 40) & CBool(Mid$(tmpString, 4, 1))
                
                'Monitor Type
                If Mid$(tmpString, 5, 2) = "00" Then
                    lstCmos.AddItem Left$("4-5" & Space(6), 6) & Left$("Monitor Type" & Space(40), 40) & "Not CGA or MDA (EGA or VGA)"
                ElseIf Mid$(tmpString, 5, 2) = "01" Then
                    lstCmos.AddItem Left$("4-5" & Space(6), 6) & Left$("Monitor Type" & Space(40), 40) & "40x25 CGA"
                ElseIf Mid$(tmpString, 5, 2) = "10" Then
                    lstCmos.AddItem Left$("4-5" & Space(6), 6) & Left$("Monitor Type" & Space(40), 40) & "80x25 CGA"
                ElseIf Mid$(tmpString, 5, 2) = "11" Then
                    lstCmos.AddItem Left$("4-5" & Space(6), 6) & Left$("Monitor Type" & Space(40), 40) & "MDA (Monochrome)"
                End If
                
                'Number of floppy drives
                If Mid$(tmpString, 7, 2) = "00" Then
                    lstCmos.AddItem Left$("6-7" & Space(6), 6) & Left$("Number Of Floppy Drives" & Space(40), 40) & "1"
                ElseIf Mid$(tmpString, 7, 2) = "01" Then
                    lstCmos.AddItem Left$("6-7" & Space(6), 6) & Left$("Number Of Floppy Drives" & Space(40), 40) & "2"
                ElseIf Mid$(tmpString, 7, 2) = "10" Then
                    lstCmos.AddItem Left$("6-7" & Space(6), 6) & Left$("Number Of Floppy Drives" & Space(40), 40) & "3?"
                ElseIf Mid$(tmpString, 7, 2) = "11" Then
                    lstCmos.AddItem Left$("6-7" & Space(6), 6) & Left$("Number Of Floppy Drives" & Space(40), 40) & "4?"
                End If
                
                lstCmos.AddItem ""
            Case 21 '15h
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Base Memory In Kb - Low Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 22 '16h
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Base Memory In Kb - High Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
                lstCmos.AddItem Space(6) & Left$("Base Memory In Kb - Total" & Space(40), 40) & CLng("&H" & Right$("00" & tmpString, 2) & Right$(lstCmos.List(61), 2))
                lstCmos.AddItem ""
            Case 23 '17h
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Extended Memory In Kb - Low Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 24 '18h
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Extended Memory In Kb - High Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
                lstCmos.AddItem Space(6) & Left$("Extended Memory In Kb - Total" & Space(40), 40) & CLng("&H" & Right$("00" & tmpString, 2) & Right$(lstCmos.List(65), 2))
                lstCmos.AddItem ""
            'Case 25 '19h
            'Case 26 '1Ah
            'Case 27 '1Bh
            'Case 28 '1Ch
            'Case 29 '1Dh
            'Case 30 '1Eh
            'Case 31 '1Fh
            'Case 32 '20h
            'Case 33 '21h
            'Case 34 '22h
            'Case 35 '23h
            'Case 36 '24h
            'Case 37 '25h
            'Case 38 '26h
            'Case 39 '27h
            'Case 40 '28h
            'Case 41 '29h
            'Case 42 '2Ah
            'Case 43 '2Bh
            'Case 44 '2Ch
            'Case 45 '2Dh
            Case 46: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Standard CMOS Checksum - High Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 47: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Standard CMOS Checksum - Low Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 48 '30h
                lstCmos.AddItem ""
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Extended Memory In Kb - Low Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case 49 '31h
                lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Extended Memory In Kb - High Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
                lstCmos.AddItem Space(6) & Left$("Extended Memory In Kb - Total" & Space(40), 40) & CLng("&H" & Right$("00" & tmpString, 2) & Right$(lstCmos.List(93), 2))
                lstCmos.AddItem ""
            Case 50: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Century Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
            'Case 51 '33h
            'Case 52 '34h
            'Case 53 '35h
            'Case 54 '36h
            Case 55: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Date Century Byte" & Space(40), 40) & Right$("00" & tmpString, 2)
            Case Else: lstCmos.AddItem Left$(Hex(tmpInt) & "h" & Space(6), 6) & Left$("Raw Data" & Space(40), 40) & Right$("00" & tmpString, 2)
        End Select
        
        tmpInt = tmpInt + 1
        Loop
    Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Kill Dirs.AppPath & "\cmos.dat"
    Unload Me
End Sub
