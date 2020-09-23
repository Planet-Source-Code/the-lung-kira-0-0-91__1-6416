Attribute VB_Name = "commdlg"
Option Explicit


Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSE_COLOR) As Long
Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSE_FONT) As Long
Public Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


    Public Const CC_RGBINIT = &H1
    Public Const CC_FULLOPEN = &H2
    Public Const CC_PREVENTFULLOPEN = &H4
    Public Const CC_SHOWHELP = &H8
    Public Const CC_ENABLEHOOK = &H10
    Public Const CC_ENABLETEMPLATE = &H20
    Public Const CC_ENABLETEMPLATEHANDLE = &H40
    Public Const CC_SOLIDCOLOR = &H80
    Public Const CC_ANYCOLOR = &H100
    
    Public Const CF_SCREENFONTS = &H1
    Public Const CF_PRINTERFONTS = &H2
    Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
    Public Const CF_SHOWHELP = &H4
    Public Const CF_ENABLEHOOK = &H8
    Public Const CF_ENABLETEMPLATE = &H10
    Public Const CF_ENABLETEMPLATEHANDLE = &H20
    Public Const CF_INITTOLOGFONTSTRUCT = &H40
    Public Const CF_USESTYLE = &H80
    Public Const CF_EFFECTS = &H100
    Public Const CF_APPLY = &H200
    Public Const CF_ANSIONLY = &H400
    Public Const CF_SCRIPTSONLY = CF_ANSIONLY
    Public Const CF_NOVECTORFONTS = &H800
    Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
    Public Const CF_NOSIMULATIONS = &H1000
    Public Const CF_LIMITSIZE = &H2000
    Public Const CF_FIXEDPITCHONLY = &H4000
    Public Const CF_WYSIWYG = &H8000                'must also have CF_SCREENFONTS & CF_PRINTERFONTS
    Public Const CF_FORCEFONTEXIST = &H10000
    Public Const CF_SCALABLEONLY = &H20000
    Public Const CF_TTONLY = &H40000
    Public Const CF_NOFACESEL = &H80000
    Public Const CF_NOSTYLESEL = &H100000
    Public Const CF_NOSIZESEL = &H200000
    Public Const CF_SELECTSCRIPT = &H400000
    Public Const CF_NOSCRIPTSEL = &H800000
    Public Const CF_NOVERTFONTS = &H1000000

    Public Const SIMULATED_FONTTYPE = &H8000
    Public Const PRINTER_FONTTYPE = &H4000
    Public Const SCREEN_FONTTYPE = &H2000
    Public Const BOLD_FONTTYPE = &H100
    Public Const ITALIC_FONTTYPE = &H200
    Public Const REGULAR_FONTTYPE = &H400

    Public Const OFN_READONLY = &H1
    Public Const OFN_OVERWRITEPROMPT = &H2
    Public Const OFN_HIDEREADONLY = &H4
    Public Const OFN_NOCHANGEDIR = &H8
    Public Const OFN_SHOWHELP = &H10
    Public Const OFN_ENABLEHOOK = &H20
    Public Const OFN_ENABLETEMPLATE = &H40
    Public Const OFN_ENABLETEMPLATEHANDLE = &H80
    Public Const OFN_NOVALIDATE = &H100
    Public Const OFN_ALLOWMULTISELECT = &H200
    Public Const OFN_EXTENSIONDIFFERENT = &H400
    Public Const OFN_PATHMUSTEXIST = &H800
    Public Const OFN_FILEMUSTEXIST = &H1000
    Public Const OFN_CREATEPROMPT = &H2000
    Public Const OFN_SHAREAWARE = &H4000
    Public Const OFN_NOREADONLYRETURN = &H8000
    Public Const OFN_NOTESTFILECREATE = &H10000
    Public Const OFN_NONETWORKBUTTON = &H20000
    Public Const OFN_NOLONGNAMES = &H40000              'force no long names for 4.x modules
    Public Const OFN_EXPLORER = &H80000                 'new look commdlg
    Public Const OFN_NODEREFERENCELINKS = &H100000
    Public Const OFN_LONGNAMES = &H200000               'force long names for 3.x modules
    Public Const OFN_ENABLEINCLUDENOTIFY = &H400000     'send include message to callback
    Public Const OFN_ENABLESIZING = &H800000
    Public Const OFN_DONTADDTORECENT = &H2000000
    Public Const OFN_FORCESHOWHIDDEN = &H10000000       'Show All files including System and hidden files
    Public Const OFN_EX_NOPLACESBAR = &H1

    Public CHOOSE_COLOR As CHOOSE_COLOR
    Public Type CHOOSE_COLOR 'Name conflict
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As String
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type

    Public CHOOSE_FONT As CHOOSE_FONT
    Public Type CHOOSE_FONT
        lStructSize As Long
        hwndOwner As Long           'caller's window handle
        hdc As Long                 'printer DC/IC or NULL
        lpLogFont As Long
        iPointSize As Long          '10 * size in points of selected font
        flags As Long               'enum. type flags
        rgbColors As Long           'returned text color
        lCustData As Long           'data passed to hook fn.
        lpfnHook As Long            'ptr. to hook function
        lpTemplateName As String    'custom template name
        hInstance As Long           'instance handle of.EXE that contains cust. dlg. template
        lpszStyle As String         'return the style field here must be LF_FACESIZE or bigger
        nFontType As Integer        'same value reported to the EnumFonts call back with the extra FONTTYPE_ bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long            'minimum pt size allowed &
        nSizeMax As Long            'max pt size allowed if CF_LIMITSIZE is used
    End Type

    Public OPENFILENAME As OPENFILENAME
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        'pvReserved  As String
        'dwReserved As Long
        'FlagsEx As Long
    End Type

Public Function ChooseAColor(lngHwnd As Long, lngStartColor As Long, lngRetColor As Long)
    'Dim lngRetColor As Long
    'ChooseAColor Me.hwnd, 0, lngRetColor
    
    Dim CustomColors() As Byte
    ReDim CustomColors(0 To 15) As Byte
    
    With CHOOSE_COLOR
        .flags = CC_ANYCOLOR Or CC_FULLOPEN Or CC_RGBINIT
        .hwndOwner = lngHwnd
        .lpCustColors = StrConv(CustomColors, vbUnicode)
        .lStructSize = Len(CHOOSE_COLOR)
        .rgbResult = lngStartColor
    End With
    
    If ChooseColor(CHOOSE_COLOR) = 0 Then
        CommDlgError
    Else
        lngRetColor = CHOOSE_COLOR.rgbResult
    End If
End Function

Public Function ChooseAFont(lngHwnd As Long, intPoint As Integer, rgbColor As Long, strStyle As String, intFontType As Integer)
    With CHOOSE_FONT
        .flags = CF_FORCEFONTEXIST Or CF_USESTYLE Or CF_BOTH
        .hwndOwner = lngHwnd
        .lStructSize = Len(CHOOSE_FONT)
    End With
    
    If ChooseFont(CHOOSE_FONT) = 0 Then
        CommDlgError
    Else
        With CHOOSE_FONT
            intPoint = .iPointSize
            rgbColor = .rgbColors
            strStyle = .lpszStyle
            intFontType = .nFontType
        End With
    End If
End Function

Public Function GetOpenName(lngHwnd As Long, strWindowTitle As String, bufFileName As String)
    bufFileName = Space(2048) & Chr(0) 'Limit = 2048

    With OPENFILENAME
        .flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_DONTADDTORECENT
        .hwndOwner = lngHwnd 'Gives owner
        .lpstrFile = bufFileName 'Null term string
        .lpstrFilter = "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
        .lpstrTitle = strWindowTitle 'Window title
        .lStructSize = Len(OPENFILENAME) 'Fill size of structure
        .nFilterIndex = 2
        .nMaxFile = Len(bufFileName)
    End With
    
    If GetOpenFileName(OPENFILENAME) = 0 Then 'If error or cancel
        bufFileName = "" 'Return nothing
        CommDlgError
    Else
        bufFileName = Fix_NullTermStr(OPENFILENAME.lpstrFile) 'Send it back out
    End If
End Function

Public Function GetSaveName(lngHwnd As Long, strWindowTitle As String, bufFileName As String)
    bufFileName = Space(2048) & Chr(0) 'Limit = 2048

    With OPENFILENAME
        .flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_DONTADDTORECENT
        .hwndOwner = lngHwnd 'Gives owner
        .lpstrFile = bufFileName 'Null term string
        .lpstrFilter = "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
        .lpstrTitle = strWindowTitle 'Window title
        .lStructSize = Len(OPENFILENAME) 'Fill size of structure
        .nFilterIndex = 2
        .nMaxFile = Len(bufFileName)
    End With
    
    If GetSaveFileName(OPENFILENAME) = 0 Then 'If error or cancel
        bufFileName = "" 'Return nothing
        CommDlgError
    Else
        bufFileName = Fix_NullTermStr(OPENFILENAME.lpstrFile) 'Send it back out
    End If
End Function

Private Function CommDlgError()
    Dim errDescription As String
    
    Select Case CommDlgExtendedError 'Select error
        'All
        Case CDERR_DIALOGFAILURE: errDescription = "The common dialog box procedure's call to the DialogBox function failed."
        Case CDERR_FINDRESFAILURE: errDescription = "The common dialog box procedure failed to find a specified resource."
        Case CDERR_INITIALIZATION: errDescription = "The common dialog box procedure failed during initialization."
        Case CDERR_LOADRESFAILURE: errDescription = "The common dialog box procedure failed to load a specified resource."
        Case CDERR_LOADSTRFAILURE: errDescription = "The common dialog box procedure failed to load a specified string."
        Case CDERR_LOCKRESFAILURE: errDescription = "The common dialog box procedure failed to lock a specified resource."
        Case CDERR_MEMALLOCFAILURE: errDescription = "The common dialog box procedure was unable to allocate memory for internal structures."
        Case CDERR_MEMLOCKFAILURE: errDescription = "The common dialog box procedure was unable to lock the memory associated with a handle."
        Case CDERR_NOHINSTANCE: errDescription = "The ENABLETEMPLATE flag was specified in the Flags member of a structure for the corresponding common dialog box, but the application failed to provide a corresponding instance handle."
        Case CDERR_NOHOOK: errDescription = "The ENABLEHOOK flag was specified in the Flags member of a structure for the corresponding common dialog box, but the application failed to provide a pointer to a corresponding hook function."
        Case CDERR_NOTEMPLATE: errDescription = "The ENABLETEMPLATE flag was specified in the Flags member of a structure for the corresponding common dialog box, but the application failed to provide a corresponding template."
        Case CDERR_REGISTERMSGFAIL: errDescription = "The RegisterWindowMessage function returned an error value when it was called by the common dialog box procedure."
        Case CDERR_STRUCTSIZE: errDescription = "The lStructSize member of a structure for the corresponding common dialog box is invalid."
        
        Case CFERR_MAXLESSTHANMIN: errDescription = "The size specified in the nSizeMax member of the CHOOSEFONT structure is less than the size specified in the nSizeMin member."
        Case CFERR_NOFONTS: errDescription = "No fonts exist."
        
        Case FNERR_BUFFERTOOSMALL: errDescription = "The buffer for a filename is too small."
        Case FNERR_INVALIDFILENAME: errDescription = "A filename is invalid."
        Case FNERR_SUBCLASSFAILURE: errDescription = "An attempt to subclass a list box failed because insufficient memory was available."
        
        Case FRERR_BUFFERLENGTHZERO: errDescription = "A member in a structure for the corresponding common dialog box points to an invalid buffer."
        
        'Print dialog
        Case PDERR_CREATEICFAILURE: errDescription = "The PrintDlg function failed when it attempted to create an information context."
        Case PDERR_DEFAULTDIFFERENT: errDescription = "An application called the PrintDlg function with the DN_DEFAULTPRN flag specified in the wDefault member of the DEVNAMES structure, but the printer described by the other structure members did not match the current default printer."
        Case PDERR_DNDMMISMATCH: errDescription = "The data in the DEVMODE and DEVNAMES structures describes two different printers."
        Case PDERR_GETDEVMODEFAIL: errDescription = "The printer driver failed to initialize a DEVMODE structure."
        Case PDERR_INITFAILURE: errDescription = "The PrintDlg function failed during initialization, and there is no more specific extended error code to describe the failure. This is the generic default error code for the function."
        Case PDERR_LOADDRVFAILURE: errDescription = "The PrintDlg function failed to load the device driver for the specified printer."
        Case PDERR_NODEFAULTPRN: errDescription = "A default printer does not exist."
        Case PDERR_NODEVICES: errDescription = "No printer drivers were found."
        Case PDERR_PARSEFAILURE: errDescription = "The PrintDlg function failed to parse the strings in the [devices] section of the WIN.INI file."
        Case PDERR_PRINTERNOTFOUND: errDescription = "The [devices] section of the WIN.INI file did not contain an entry for the requested printer."
        Case PDERR_RETDEFFAILURE: errDescription = "The PD_RETURNDEFAULT flag was specified in the Flags member of the PRINTDLG structure, but the hDevMode or hDevNames member was nonzero."
        Case PDERR_SETUPFAILURE: errDescription = "The PrintDlg function failed to load the required resources."
    End Select
    
    'Do not send blank message
    If errDescription <> "" Then
        If errMsg = True Then
            MsgBox errDescription, vbExclamation, "Error"
        End If
    End If
End Function
