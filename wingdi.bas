Attribute VB_Name = "wingdi"
'wingdi.h

Option Explicit


'Declared Functions

Public Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long


    'Constants

    'size of a device name string
    Public Const CCHDEVICENAME = 32

    'size of a form name string
    Public Const CCHFORMNAME = 32

    Public Const ANSI_CHARSET = 0
    Public Const DEFAULT_CHARSET = 1
    Public Const SYMBOL_CHARSET = 2
    Public Const SHIFTJIS_CHARSET = 128
    Public Const HANGEUL_CHARSET = 129
    Public Const HANGUL_CHARSET = 129
    Public Const GB2312_CHARSET = 134
    Public Const CHINESEBIG5_CHARSET = 136
    Public Const OEM_CHARSET = 255

    Public Const JOHAB_CHARSET = 130
    Public Const HEBREW_CHARSET = 177
    Public Const ARABIC_CHARSET = 178
    Public Const GREEK_CHARSET = 161
    Public Const TURKISH_CHARSET = 162
    Public Const VIETNAMESE_CHARSET = 163
    Public Const THAI_CHARSET = 222
    Public Const EASTEUROPE_CHARSET = 238
    Public Const RUSSIAN_CHARSET = 204
    
    Public Const MAC_CHARSET = 77
    Public Const BALTIC_CHARSET = 186

    'field selection bits
    Public Const DM_ORIENTATION = &H1
    Public Const DM_PAPERSIZE = &H2
    Public Const DM_PAPERLENGTH = &H4
    Public Const DM_PAPERWIDTH = &H8
    Public Const DM_SCALE = &H10
    
    Public Const DM_POSITION = &H20
    Public Const DM_NUP = &H40

    Public Const DM_COPIES = &H100
    Public Const DM_DEFAULTSOURCE = &H200
    Public Const DM_PRINTQUALITY = &H400
    Public Const DM_COLOR = &H800
    Public Const DM_DUPLEX = &H1000
    Public Const DM_YRESOLUTION = &H2000
    Public Const DM_TTOPTION = &H4000
    Public Const DM_COLLATE = &H8000
    Public Const DM_FORMNAME = &H10000
    Public Const DM_LOGPIXELS = &H20000
    Public Const DM_BITSPERPEL = &H40000
    Public Const DM_PELSWIDTH = &H80000
    Public Const DM_PELSHEIGHT = &H100000
    Public Const DM_DISPLAYFLAGS = &H200000
    Public Const DM_DISPLAYFREQUENCY = &H400000

    Public Const DM_ICMMETHOD = &H800000
    Public Const DM_ICMINTENT = &H1000000
    Public Const DM_MEDIATYPE = &H2000000
    Public Const DM_DITHERTYPE = &H4000000
    Public Const DM_PANNINGWIDTH = &H8000000
    Public Const DM_PANNINGHEIGHT = &H10000000

    'orientation selections
    Public Const DMORIENT_PORTRAIT = 1
    Public Const DMORIENT_LANDSCAPE = 2

    'paper selections
    Public Const DMPAPER_LETTER = 1                'Letter 8 1/2 x 11 in
    Public Const DMPAPER_FIRST = DMPAPER_LETTER
    Public Const DMPAPER_LETTERSMALL = 2           'Letter Small 8 1/2 x 11 in
    Public Const DMPAPER_TABLOID = 3               'Tabloid 11 x 17 in
    Public Const DMPAPER_LEDGER = 4                'Ledger 17 x 11 in
    Public Const DMPAPER_LEGAL = 5                 'Legal 8 1/2 x 14 in
    Public Const DMPAPER_STATEMENT = 6             'Statement 5 1/2 x 8 1/2 in
    Public Const DMPAPER_EXECUTIVE = 7             'Executive 7 1/4 x 10 1/2 in
    Public Const DMPAPER_A3 = 8                    'A3 297 x 420 mm
    Public Const DMPAPER_A4 = 9                    'A4 210 x 297 mm
    Public Const DMPAPER_A4SMALL = 10              'A4 Small 210 x 297 mm
    Public Const DMPAPER_A5 = 11                   'A5 148 x 210 mm
    Public Const DMPAPER_B4 = 12                   'B4 (JIS) 250 x 354
    Public Const DMPAPER_B5 = 13                   'B5 (JIS) 182 x 257 mm
    Public Const DMPAPER_FOLIO = 14                'Folio 8 1/2 x 13 in
    Public Const DMPAPER_QUARTO = 15               'Quarto 215 x 275 mm
    Public Const DMPAPER_10X14 = 16                '1&H14 in
    Public Const DMPAPER_11X17 = 17                '11x17 in
    Public Const DMPAPER_NOTE = 18                 'Note 8 1/2 x 11 in
    Public Const DMPAPER_ENV_9 = 19                'Envelope #9 3 7/8 x 8 7/8
    Public Const DMPAPER_ENV_10 = 20               'Envelope #10 4 1/8 x 9 1/2
    Public Const DMPAPER_ENV_11 = 21               'Envelope #11 4 1/2 x 10 3/8
    Public Const DMPAPER_ENV_12 = 22               'Envelope #12 4 \276 x 11
    Public Const DMPAPER_ENV_14 = 23               'Envelope #14 5 x 11 1/2
    Public Const DMPAPER_CSHEET = 24               'C size sheet
    Public Const DMPAPER_DSHEET = 25               'D size sheet
    Public Const DMPAPER_ESHEET = 26               'E size sheet
    Public Const DMPAPER_ENV_DL = 27               'Envelope DL 110 x 220mm
    Public Const DMPAPER_ENV_C5 = 28               'Envelope C5 162 x 229 mm
    Public Const DMPAPER_ENV_C3 = 29               'Envelope C3  324 x 458 mm
    Public Const DMPAPER_ENV_C4 = 30               'Envelope C4  229 x 324 mm
    Public Const DMPAPER_ENV_C6 = 31               'Envelope C6  114 x 162 mm
    Public Const DMPAPER_ENV_C65 = 32              'Envelope C65 114 x 229 mm
    Public Const DMPAPER_ENV_B4 = 33               'Envelope B4  250 x 353 mm
    Public Const DMPAPER_ENV_B5 = 34               'Envelope B5  176 x 250 mm
    Public Const DMPAPER_ENV_B6 = 35               'Envelope B6  176 x 125 mm
    Public Const DMPAPER_ENV_ITALY = 36            'Envelope 110 x 230 mm
    Public Const DMPAPER_ENV_MONARCH = 37          'Envelope Monarch 3.875 x 7.5 in
    Public Const DMPAPER_ENV_PERSONAL = 38         '6 3/4 Envelope 3 5/8 x 6 1/2 in
    Public Const DMPAPER_FANFOLD_US = 39           'US Std Fanfold 14 7/8 x 11 in
    Public Const DMPAPER_FANFOLD_STD_GERMAN = 40   'German Std Fanfold 8 1/2 x 12 in
    Public Const DMPAPER_FANFOLD_LGL_GERMAN = 41   'German Legal Fanfold 8 1/2 x 13 in
    
    Public Const DMPAPER_ISO_B4 = 42               'B4 (ISO) 250 x 353 mm
    Public Const DMPAPER_JAPANESE_POSTCARD = 43    'Japanese Postcard 100 x 148 mm
    Public Const DMPAPER_9X11 = 44                 '9 x 11 in
    Public Const DMPAPER_10X11 = 45                '10 x 11 in
    Public Const DMPAPER_15X11 = 46                '15 x 11 in
    Public Const DMPAPER_ENV_INVITE = 47           'Envelope Invite 220 x 220 mm
    Public Const DMPAPER_RESERVED_48 = 48          'RESERVED--DO NOT USE
    Public Const DMPAPER_RESERVED_49 = 49          'RESERVED--DO NOT USE
    Public Const DMPAPER_LETTER_EXTRA = 50         'Letter Extra 9 \275 x 12 in
    Public Const DMPAPER_LEGAL_EXTRA = 51          'Legal Extra 9 \275 x 15 in
    Public Const DMPAPER_TABLOID_EXTRA = 52        'Tabloid Extra 11.69 x 18 in
    Public Const DMPAPER_A4_EXTRA = 53             'A4 Extra 9.27 x 12.69 in
    Public Const DMPAPER_LETTER_TRANSVERSE = 54    'Letter Transverse 8 \275 x 11 in
    Public Const DMPAPER_A4_TRANSVERSE = 55        'A4 Transverse 210 x 297 mm
    Public Const DMPAPER_LETTER_EXTRA_TRANSVERSE = 56 'Letter Extra Transverse 9\275 x 12 in
    Public Const DMPAPER_A_PLUS = 57               'SuperA/SuperA/A4 227 x 356 mm
    Public Const DMPAPER_B_PLUS = 58               'SuperB/SuperB/A3 305 x 487 mm
    Public Const DMPAPER_LETTER_PLUS = 59          'Letter Plus 8.5 x 12.69 in
    Public Const DMPAPER_A4_PLUS = 60              'A4 Plus 210 x 330 mm
    Public Const DMPAPER_A5_TRANSVERSE = 61        'A5 Transverse 148 x 210 mm
    Public Const DMPAPER_B5_TRANSVERSE = 62        'B5 (JIS) Transverse 182 x 257 mm
    Public Const DMPAPER_A3_EXTRA = 63             'A3 Extra 322 x 445 mm
    Public Const DMPAPER_A5_EXTRA = 64             'A5 Extra 174 x 235 mm
    Public Const DMPAPER_B5_EXTRA = 65             'B5 (ISO) Extra 201 x 276 mm
    Public Const DMPAPER_A2 = 66                   'A2 420 x 594 mm
    Public Const DMPAPER_A3_TRANSVERSE = 67        'A3 Transverse 297 x 420 mm
    Public Const DMPAPER_A3_EXTRA_TRANSVERSE = 68  'A3 Extra Transverse 322 x 445 mm

    Public Const DMPAPER_DBL_JAPANESE_POSTCARD = 69 'Japanese Double Postcard 200 x 148 mm
    Public Const DMPAPER_A6 = 70                   'A6 105 x 148 mm
    Public Const DMPAPER_JENV_KAKU2 = 71           'Japanese Envelope Kaku #2
    Public Const DMPAPER_JENV_KAKU3 = 72           'Japanese Envelope Kaku #3
    Public Const DMPAPER_JENV_CHOU3 = 73           'Japanese Envelope Chou #3
    Public Const DMPAPER_JENV_CHOU4 = 74           'Japanese Envelope Chou #4
    Public Const DMPAPER_LETTER_ROTATED = 75       'Letter Rotated 11 x 8 1/2 11 in
    Public Const DMPAPER_A3_ROTATED = 76           'A3 Rotated 420 x 297 mm
    Public Const DMPAPER_A4_ROTATED = 77           'A4 Rotated 297 x 210 mm
    Public Const DMPAPER_A5_ROTATED = 78           'A5 Rotated 210 x 148 mm
    Public Const DMPAPER_B4_JIS_ROTATED = 79       'B4 (JIS) Rotated 364 x 257 mm
    Public Const DMPAPER_B5_JIS_ROTATED = 80       'B5 (JIS) Rotated 257 x 182 mm
    Public Const DMPAPER_JAPANESE_POSTCARD_ROTATED = 81 'Japanese Postcard Rotated 148 x 100 mm
    Public Const DMPAPER_DBL_JAPANESE_POSTCARD_ROTATED = 82 ' Double Japanese Postcard Rotated 148 x 200 mm
    Public Const DMPAPER_A6_ROTATED = 83           'A6 Rotated 148 x 105 mm
    Public Const DMPAPER_JENV_KAKU2_ROTATED = 84   'Japanese Envelope Kaku #2 Rotated
    Public Const DMPAPER_JENV_KAKU3_ROTATED = 85   'Japanese Envelope Kaku #3 Rotated
    Public Const DMPAPER_JENV_CHOU3_ROTATED = 86   'Japanese Envelope Chou #3 Rotated
    Public Const DMPAPER_JENV_CHOU4_ROTATED = 87   'Japanese Envelope Chou #4 Rotated
    Public Const DMPAPER_B6_JIS = 88               'B6 (JIS) 128 x 182 mm
    Public Const DMPAPER_B6_JIS_ROTATED = 89       'B6 (JIS) Rotated 182 x 128 mm
    Public Const DMPAPER_12X11 = 90                '12 x 11 in
    Public Const DMPAPER_JENV_YOU4 = 91            'Japanese Envelope You #4
    Public Const DMPAPER_JENV_YOU4_ROTATED = 92    'Japanese Envelope You #4 Rotated
    Public Const DMPAPER_P16K = 93                 'PRC 16K 146 x 215 mm
    Public Const DMPAPER_P32K = 94                 'PRC 32K 97 x 151 mm
    Public Const DMPAPER_P32KBIG = 95              'PRC 32K(Big) 97 x 151 mm
    Public Const DMPAPER_PENV_1 = 96               'PRC Envelope #1 102 x 165 mm
    Public Const DMPAPER_PENV_2 = 97               'PRC Envelope #2 102 x 176 mm
    Public Const DMPAPER_PENV_3 = 98               'PRC Envelope #3 125 x 176 mm
    Public Const DMPAPER_PENV_4 = 99               'PRC Envelope #4 110 x 208 mm
    Public Const DMPAPER_PENV_5 = 100              'PRC Envelope #5 110 x 220 mm
    Public Const DMPAPER_PENV_6 = 101              'PRC Envelope #6 120 x 230 mm
    Public Const DMPAPER_PENV_7 = 102              'PRC Envelope #7 160 x 230 mm
    Public Const DMPAPER_PENV_8 = 103              'PRC Envelope #8 120 x 309 mm
    Public Const DMPAPER_PENV_9 = 104              'PRC Envelope #9 229 x 324 mm
    Public Const DMPAPER_PENV_10 = 105             'PRC Envelope #10 324 x 458 mm
    Public Const DMPAPER_P16K_ROTATED = 106        'PRC 16K Rotated
    Public Const DMPAPER_P32K_ROTATED = 107        'PRC 32K Rotated
    Public Const DMPAPER_P32KBIG_ROTATED = 108     'PRC 32K(Big) Rotated
    Public Const DMPAPER_PENV_1_ROTATED = 109      'PRC Envelope #1 Rotated 165 x 102 mm
    Public Const DMPAPER_PENV_2_ROTATED = 110      'PRC Envelope #2 Rotated 176 x 102 mm
    Public Const DMPAPER_PENV_3_ROTATED = 111      'PRC Envelope #3 Rotated 176 x 125 mm
    Public Const DMPAPER_PENV_4_ROTATED = 112      'PRC Envelope #4 Rotated 208 x 110 mm
    Public Const DMPAPER_PENV_5_ROTATED = 113      'PRC Envelope #5 Rotated 220 x 110 mm
    Public Const DMPAPER_PENV_6_ROTATED = 114      'PRC Envelope #6 Rotated 230 x 120 mm
    Public Const DMPAPER_PENV_7_ROTATED = 115      'PRC Envelope #7 Rotated 230 x 160 mm
    Public Const DMPAPER_PENV_8_ROTATED = 116      'PRC Envelope #8 Rotated 309 x 120 mm
    Public Const DMPAPER_PENV_9_ROTATED = 117      'PRC Envelope #9 Rotated 324 x 229 mm
    Public Const DMPAPER_PENV_10_ROTATED = 118     'PRC Envelope #10 Rotated 458 x 324 mm

    'If (WINVER >= &H0500)
    'Public Const DMPAPER_LAST = DMPAPER_PENV_10_ROTATED
    '#elseif (WINVER >= &H0400)
    Public Const DMPAPER_LAST = DMPAPER_A3_EXTRA_TRANSVERSE
    'Else
    'Public Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN

    Public Const DMPAPER_USER = 256

    'bin selections
    Public Const DMBIN_UPPER = 1
    Public Const DMBIN_FIRST = DMBIN_UPPER
    Public Const DMBIN_ONLYONE = 1
    Public Const DMBIN_LOWER = 2
    Public Const DMBIN_MIDDLE = 3
    Public Const DMBIN_MANUAL = 4
    Public Const DMBIN_ENVELOPE = 5
    Public Const DMBIN_ENVMANUAL = 6
    Public Const DMBIN_AUTO = 7
    Public Const DMBIN_TRACTOR = 8
    Public Const DMBIN_SMALLFMT = 9
    Public Const DMBIN_LARGEFMT = 10
    Public Const DMBIN_LARGECAPACITY = 11
    Public Const DMBIN_CASSETTE = 14
    Public Const DMBIN_FORMSOURCE = 15
    Public Const DMBIN_LAST = DMBIN_FORMSOURCE

    Public Const DMBIN_USER = 256    'device specific bins start here

    'print qualities
    Public Const DMRES_DRAFT = (-1)
    Public Const DMRES_LOW = (-2)
    Public Const DMRES_MEDIUM = (-3)
    Public Const DMRES_HIGH = (-4)

    'color enable/disable for color printers
    Public Const DMCOLOR_MONOCHROME = 1
    Public Const DMCOLOR_COLOR = 2

    'duplex enable
    Public Const DMDUP_SIMPLEX = 1
    Public Const DMDUP_VERTICAL = 2
    Public Const DMDUP_HORIZONTAL = 3

    'TrueType options
    Public Const DMTT_BITMAP = 1           'print TT fonts as graphics
    Public Const DMTT_DOWNLOAD = 2         'download TT fonts as soft fonts
    Public Const DMTT_SUBDEV = 3           'substitute device fonts for TT fonts
    Public Const DMTT_DOWNLOAD_OUTLINE = 4 'download TT fonts as outline soft fonts

    'Collation selections
    Public Const DMCOLLATE_FALSE = 0
    Public Const DMCOLLATE_TRUE = 1

    'DEVMODE dmDisplayFlags flags
    Public Const DM_GRAYSCALE = &H1              'This flag is no longer valid
    Public Const DM_INTERLACED = &H2             'This flag is no longer valid
    Public Const DMDISPLAYFLAGS_TEXTMODE = &H4

    'dmNup , multiple logical page per physical page options
    Public Const DMNUP_SYSTEM = 1
    Public Const DMNUP_ONEUP = 2

    'ICM methods
    Public Const DMICMMETHOD_NONE = 1      'ICM disabled
    Public Const DMICMMETHOD_SYSTEM = 2    'ICM handled by system
    Public Const DMICMMETHOD_DRIVER = 3    'ICM handled by driver
    Public Const DMICMMETHOD_DEVICE = 4    'ICM handled by device

    Public Const DMICMMETHOD_USER = 256    'Device-specific methods start here

    'ICM Intents
    Public Const DMICM_SATURATE = 1            'Maximize color saturation
    Public Const DMICM_CONTRAST = 2            'Maximize color contrast
    Public Const DMICM_COLORIMETRIC = 3        'Use specific color metric
    Public Const DMICM_ABS_COLORIMETRIC = 4    'Use specific color metric

    Public Const DMICM_USER = 256          'Device-specific intents start here

    'Media types
    Public Const DMMEDIA_STANDARD = 1        'Standard paper
    Public Const DMMEDIA_TRANSPARENCY = 2    'Transparency
    Public Const DMMEDIA_GLOSSY = 3          'Glossy paper

    Public Const DMMEDIA_USER = 256          'Device-specific media start here

    'Dither types
    Public Const DMDITHER_NONE = 1              'No dithering
    Public Const DMDITHER_COARSE = 2            'Dither with a coarse brush
    Public Const DMDITHER_FINE = 3              'Dither with a fine brush
    Public Const DMDITHER_LINEART = 4           'LineArt dithering
    Public Const DMDITHER_ERRORDIFFUSION = 5    'LineArt dithering
    Public Const DMDITHER_RESERVED6 = 6         'LineArt dithering
    Public Const DMDITHER_RESERVED7 = 7         'LineArt dithering
    Public Const DMDITHER_RESERVED8 = 8         'LineArt dithering
    Public Const DMDITHER_RESERVED9 = 9         'LineArt dithering
    Public Const DMDITHER_GRAYSCALE = 10        'Device does grayscaling

    Public Const DMDITHER_USER = 256    'Device-specific dithers start here

    Public Const RASTER_FONTTYPE = &H1
    Public Const DEVICE_FONTTYPE = &H2
    Public Const TRUETYPE_FONTTYPE = &H4

    Public Const LF_FACESIZE = 32
    Public Const LF_FULLFACESIZE = 64

    Public Const PFD_DOUBLEBUFFER = &H1
    Public Const PFD_STEREO = &H2
    Public Const PFD_DRAW_TO_WINDOW = &H4
    Public Const PFD_DRAW_TO_BITMAP = &H8
    Public Const PFD_SUPPORT_GDI = &H10
    Public Const PFD_SUPPORT_OPENGL = &H20
    Public Const PFD_GENERIC_FORMAT = &H40
    Public Const PFD_NEED_PALETTE = &H80
    Public Const PFD_NEED_SYSTEM_PALETTE = &H100
    Public Const PFD_SWAP_EXCHANGE = &H200
    Public Const PFD_SWAP_COPY = &H400
    Public Const PFD_SWAP_LAYER_BUFFERS = &H800
    Public Const PFD_GENERIC_ACCELERATED = &H1000
    Public Const PFD_SUPPORT_DIRECTDRAW = &H2000
    
    Public Const PFD_DEPTH_DONTCARE = &H20000000
    Public Const PFD_DOUBLEBUFFER_DONTCARE = &H40000000
    Public Const PFD_STEREO_DONTCARE = &H80000000


    'Types

    Public DEVMODE As DEVMODE
    Public Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
    End Type

    Public LOGFONT As LOGFONT
    Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To LF_FACESIZE) As Byte
    End Type
