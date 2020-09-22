Attribute VB_Name = "PICPRINT1"
Option Explicit

' Public constants
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const NOTSRCCOPY = &H330008
Public Const NOTSRCERASE = &H1100A6
Public Const MERGECOPY = &HC000CA
Public Const MERGEPAINT = &HBB0226
Public Const PATCOPY = &HF00021
Public Const PATPAINT = &HFB0A09
Public Const PATINVERT = &H5A0049
Public Const DSTINVERT = &H550009
Public Const BLACKNESS = &H42&
Public Const WHITENESS = &HFF0062
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_SCALE = &H10&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_PRINTQUALITY = &H400&
Public Const DM_COLOR = &H800&
Public Const DM_DUPLEX = &H1000&
Public Const DM_YRESOLUTION = &H2000&
Public Const DM_TTOPTION = &H4000&
Public Const DMORIENT_PORTRAIT = 1
Public Const DMORIENT_LANDSCAPE = 2
Public Const DMPAPER_LETTER = 1
Public Const DMPAPER_LETTERSMALL = 2
Public Const DMPAPER_TABLOID = 3
Public Const DMPAPER_LEDGER = 4
Public Const DMPAPER_LEGAL = 5
Public Const DMPAPER_STATEMENT = 6
Public Const DMPAPER_EXECUTIVE = 7
Public Const DMPAPER_A3 = 8
Public Const DMPAPER_A4 = 9
Public Const DMPAPER_A4SMALL = 10
Public Const DMPAPER_A5 = 11
Public Const DMPAPER_B4 = 12
Public Const DMPAPER_B5 = 13
Public Const DMPAPER_FOLIO = 14
Public Const DMPAPER_QUARTO = 15
Public Const DMPAPER_10X14 = 16
Public Const DMPAPER_11X17 = 17
Public Const DMPAPER_NOTE = 18
Public Const DMPAPER_ENV_9 = 19
Public Const DMPAPER_ENV_10 = 20
Public Const DMPAPER_ENV_11 = 21
Public Const DMPAPER_ENV_12 = 22
Public Const DMPAPER_ENV_14 = 23
Public Const DMPAPER_CSHEET = 24
Public Const DMPAPER_DSHEET = 25
Public Const DMPAPER_ESHEET = 26
Public Const DMPAPER_ENV_DL = 27
Public Const DMPAPER_ENV_C5 = 28
Public Const DMPAPER_ENV_C3 = 29
Public Const DMPAPER_ENV_C4 = 30
Public Const DMPAPER_ENV_C6 = 31
Public Const DMPAPER_ENV_C65 = 32
Public Const DMPAPER_ENV_B4 = 33
Public Const DMPAPER_ENV_B5 = 34
Public Const DMPAPER_ENV_B6 = 35
Public Const DMPAPER_ENV_ITALY = 36
Public Const DMPAPER_ENV_MONARCH = 37
Public Const DMPAPER_ENV_PERSONAL = 38
Public Const DMPAPER_FANFOLD_US = 39
Public Const DMPAPER_FANFOLD_STD_GERMAN = 40
Public Const DMPAPER_FANFOLD_LGL_GERMAN = 41
Public Const DMPAPER_USER = 256
Public Const DMBIN_UPPER = 1
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
Public Const DMBIN_USER = 256
Public Const DMRES_DRAFT = -1
Public Const DMRES_LOW = -2
Public Const DMRES_MEDIUM = -3
Public Const DMRES_HIGH = -4
Public Const DMCOLOR_MONOCHROME = 1
Public Const DMCOLOR_COLOR = 2
Public Const DMDUP_SIMPLEX = 1
Public Const DMDUP_VERTICAL = 2
Public Const DMDUP_HORIZONTAL = 3
Public Const DMTT_BITMAP = 1
Public Const DMTT_DOWNLOAD = 2
Public Const DMTT_SUBDEV = 3
Public Const DM_UPDATE = 1
Public Const DM_COPY = 2
Public Const DM_PROMPT = 4
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = 8
Public Const DM_IN_PROMPT = 4
Public Const DM_OUT_BUFFER = 2
Public Const DM_OUT_DEFAULT = 1
Public Const DC_FIELDS = 1
Public Const DC_PAPERS = 2
Public Const DC_PAPERSIZE = 3
Public Const DC_MINEXTENT = 4
Public Const DC_MAXEXTENT = 5
Public Const DC_BINS = 6
Public Const DC_DUPLEX = 7
Public Const DC_SIZE = 8
Public Const DC_EXTRA = 9
Public Const DC_VERSION = 10
Public Const DC_DRIVER = 11
Public Const DC_BINNAMES = 12
Public Const DC_ENUMRESOLUTIONS = 13
Public Const DC_FILEDEPENDENCIES = 14
Public Const DC_TRUETYPE = 15
Public Const DC_PAPERNAMES = 16
Public Const DC_ORIENTATION = 17
Public Const DC_COPIES = 18
Public Const DCTT_BITMAP = &H1&
Public Const DCTT_DOWNLOAD = &H2&
Public Const DCTT_SUBDEV = &H4&
Public Const SP_NOTREPORTED = &H4000
Public Const SP_ERROR = (-1)
Public Const SP_APPABORT = (-2)
Public Const SP_USERABORT = (-3)
Public Const SP_OUTOFDISK = (-4)
Public Const SP_OUTOFMEMORY = (-5)
Public Const PR_JOBSTATUS = &H0
Public Const DRIVERVERSION = 0
Public Const TECHNOLOGY = 2
Public Const HORZSIZE = 4
Public Const VERTSIZE = 6
Public Const HORZRES = 8
Public Const VERTRES = 10
Public Const BITSPIXEL = 12
Public Const PLANES = 14
Public Const NUMBRUSHES = 16
Public Const NUMPENS = 18
Public Const NUMMARKERS = 20
Public Const NUMFONTS = 22
Public Const NUMCOLORS = 24
Public Const PDEVICESIZE = 26
Public Const CURVECAPS = 28
Public Const LINECAPS = 30
Public Const POLYGONALCAPS = 32
Public Const TEXTCAPS = 34
Public Const CLIPCAPS = 36
Public Const RASTERCAPS = 38
Public Const ASPECTX = 40
Public Const ASPECTY = 42
Public Const ASPECTXY = 44
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const SIZEPALETTE = 104
Public Const NUMRESERVED = 106
Public Const COLORRES = 108
Public Const RC_BITBLT = 1
Public Const RC_BANDING = 2
Public Const RC_SCALING = 4
Public Const RC_BITMAP64 = 8
Public Const RC_GDI20_OUTPUT = &H10
Public Const RC_DI_BITMAP = &H80
Public Const RC_PALETTE = &H100
Public Const RC_DIBTODEV = &H200
Public Const RC_BIGFONT = &H400
Public Const RC_STRETCHBLT = &H800
Public Const RC_FLOODFILL = &H1000
Public Const RC_STRETCHDIB = &H2000
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MODIFY = &H80
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_LOWER = GMEM_NOT_BANKED
Public Const DIB_RGB_COLORS = 0
Public Const DIB_PAL_COLORS = 1
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' Public variables
Public AbortPrinting%
Public UseHourglass%

'  size of a device name string
Public Const CCHDEVICENAME = 32

'  size of a form name string
Public Const CCHFORMNAME = 32

Public Const BI_RGB = 0&

Type POINTAPI
        x As Long
        y As Long
End Type

Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME '32
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
        dmFormName As String * CCHFORMNAME '32
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
            dmICMMethod As Long        ' // Windows 95 only
    dmICMIntent As Long        ' // Windows 95 only
    dmMediaType As Long        ' // Windows 95 only
    dmDitherType As Long       ' // Windows 95 only
    dmReserved1 As Long        ' // Windows 95 only
    dmReserved2 As Long        ' // Windows 95 only (146 Bytes)
End Type

Type PRINTER_DEFAULTS
        pDatatype As String
        pDevMode As Long
        DesiredAccess As Long
End Type

Type DOCINFO
        cbSize As Long
        lpszDocName As String
        lpszOutput As String
End Type

Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

' BITMAPINFO for this example is for 16 color bitmap
Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(256) As RGBQUAD
End Type

Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Declare Function OpenPrinterBynum Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Declare Function ResetPrinter Lib "winspool.drv" Alias "ResetPrinterA" (ByVal hPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, ByVal fMode As Long) As Long
Declare Function AdvancedDocumentProperties Lib "winspool.drv" Alias "AdvancedDocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As DEVMODE, ByVal pDevModeInput As Long) As Long
Declare Function ConnectToPrinterDlg Lib "winspool.drv" (ByVal hwnd As Long, ByVal flags As Long) As Long
Declare Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long
Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, ByVal lpDevMode As Long) As Long
Declare Function CreateDCBynum Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndDocAPI Lib "gdi32" Alias "EndDoc" (ByVal hdc As Long) As Long
Declare Function AbortDoc Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SetAbortProc Lib "gdi32" (ByVal hdc As Long, ByVal lpAbortProc As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


Public Function Callback1_AbortProc(ByVal hPr As Long, ByVal code As Long) As Long
    ' We must allow events to take place, otherwise the
    ' user button press on the abortform form will never
    ' be detected!
    DoEvents
    If code = SP_OUTOFDISK Or AbortPrinting% Then
        Callback1_AbortProc = 0
        Exit Function
    End If
    Callback1_AbortProc = -1
End Function

