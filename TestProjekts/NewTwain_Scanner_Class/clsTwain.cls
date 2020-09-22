VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTwain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const OFFSET_4 = 4294967296#
      Private Const MAXINT_4 = 2147483647
      Private Const OFFSET_2 = 65536
      Private Const MAXINT_2 = 32767

Private Type TW_ENUMERATION
 ItemType As Integer
 NumItems(3) As Byte 'TW_UINT32 'How many items in ItemList
 CurrentIndex(3) As Byte 'TW_UINT32 'Current value is in ItemList(CurrentIndex)
 DefaultIndex(3) As Byte 'TW_UINT32 'Powerup value is in ItemList(DefaultIndex)
End Type


Public Enum TWAIN_CLASS_COLORTYPE
    Bitmap_BW_1bit = 0
    Bitmap_GREY_8bit = 1
    Bitmap_RGB_24bit = 2
    Bitmap_PAL_8bit = 3
End Enum

Private Enum TWAIN_MDL_ITEMYPE
    INT16 = 1                           ' TW_INT16      short
    UINT16 = 4                          ' TW_UINT16     unsigned short
    BOOL = 6                            ' TW_BOOL       unsigned short
    FIX32 = 7                           ' TW_FIX32      structure
End Enum

Private Enum TWAIN_MDL_CAPABILITY
    XFERCOUNT = 1                       ' CAP_XFERCOUNT
    PIXELTYPE = 257                     ' ICAP_PIXELTYPE
    INDICATORS = 4107                   ' CAP_INDICATORS
    UICONTROLLABLE = 4110               ' CAP_UICONTROLLABLE
    PHYSICALWIDTH = 4369                ' ICAP_PSYSICALWIDTH
    PHYSICALHEIGHT = 4370               ' ICAP_PSYSICALHEIGHT
    XRESOLUTION = 4376                  ' ICAP_XRESOLUTION
    YRESOLUTION = 4377                  ' ICAP_YRESOLUTION
    BITDEPTH = 4395                     ' ICAP_BITDEPTH
End Enum

Private Type TW_FIX32
    Whole   As Integer                          ' TW_INT16
    Frac    As Integer                          ' TW_UINT16
End Type

Private Type TW_USERINTERFACE
    ShowUI   As Integer                         ' TW_BOOL
    ModalUI  As Integer                         ' TW_BOOL
    hParent  As Long                            ' TW_HANDLE
End Type

Private Type TW_PENDINGXFERS
    Count       As Integer                      ' TW_UINT16
    Reserved1   As Integer                      ' TW_UINT32
    Reserved2   As Integer
End Type

Private Type TW_EVENT
    pEvent      As Long                         ' TW_MEMREF
    TWMessage   As Integer                      ' TW_UINT16
End Type

Private Type TW_FRAME
    Left     As TW_FIX32                        ' TW_FIX32
    Top      As TW_FIX32                        ' TW_FIX32
    Right    As TW_FIX32                        ' TW_FIX32
    Bottom   As TW_FIX32                        ' TW_FIX32
End Type

Private Type TW_IMAGELAYOUT
    Frame            As TW_FRAME                ' TW_FRAME
    DocumentNumber   As Long                    ' TW_UINT32
    PageNumber       As Long                    ' TW_UINT32
    FrameNumber      As Long                    ' TW_UINT32
End Type

Private Type TW_ONEVALUE
    ItemType As Integer                         ' TW_UINT16
    Item1    As Integer                         ' TW_UINT32
    Item2    As Integer
End Type

Private Type TW_CAPABILITY
    Cap          As Integer                     ' TW_UINT16
    ConType      As Integer                     ' TW_UINT16
    hContainer   As Long                        ' TW_HANDLE
End Type

Private Type TW_VERSION
 MajorNum As Integer
 MinorNum As Integer
 Language As Integer
 Country As Integer
 Info(33) As Byte
End Type

Private Type TW_IDENTITY
 Id(3) As Byte 'TW_UINT32
 Version As TW_VERSION
 ProtocolMajor As Integer
 ProtocolMinor As Integer
 SupportedGroups(3) As Byte 'TW_UINT32
 Manufacturer(33) As Byte
 ProductFamily(33) As Byte
 ProductName(33) As Byte
End Type

Private Type ForBytes
    Bytearray(3) As Byte
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSG
    hwnd    As Long
    message As Long
    wParam  As Long
    lParam  As Long
    time    As Long
    pt      As POINTAPI
End Type

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
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

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO2
        bmiHeader As BITMAPINFOHEADER
        bmiColors(0 To 1) As RGBQUAD
End Type

Private Type BITMAPINFO8
        bmiHeader As BITMAPINFOHEADER
        bmiColors(0 To 255) As RGBQUAD
End Type

Private Const DG_CONTROL = &H1&
Private Const DG_IMAGE = &H2&

Private Const DAT_CAPABILITY = 1
Private Const DAT_EVENT = 2
Private Const DAT_IDENTITY = 3
Private Const DAT_PARENT = 4
Private Const DAT_PENDINGXFERS = 5
Private Const DAT_USERINTERFACE = 9
Private Const DAT_IMAGELAYOUT = 258
Private Const DAT_IMAGENATIVEXFER = 260

Private Const MSG_GET = 1
Private Const MSG_SET = 6
Private Const MSG_XFERREADY = 257
Private Const MSG_CLOSEDSREQ = 258
Private Const MSG_OPENDSM = 769
Private Const MSG_CLOSEDSM = 770
Private Const MSG_OPENDS = 1025
Private Const MSG_CLOSEDS = 1026
Private Const MSG_USERSELECT = 1027
Private Const MSG_DISABLEDS = 1281
Private Const MSG_ENABLEDS = 1282
Private Const MSG_PROCESSEVENT = 1537
Private Const MSG_ENDXFER = 1793

Private Const TWLG_ENG = 2
Private Const TWCY_USA = 1

Private Const TWRC_CHECKSTATUS = 2
Private Const TWRC_DSEVENT = 4
Private Const TWRC_NOTDSEVENT = 5
Private Const TWRC_XFERDONE = 6

Private Const TWON_PROTOCOLMAJOR = 1
Private Const TWON_ARRAY = 3
Private Const TWON_ENUMERATION = 4
Private Const TWON_ONEVALUE = 5
Private Const TWON_RANGE = 6
Private Const TWON_PROTOCOLMINOR = 6

Private Const GHND = 66

Private Declare Function DSM_Entry Lib "Twain_32.dll" (ByRef pOrigin As Any, ByRef pDest As Any, ByVal DG As Long, ByVal DAT As Integer, ByVal MSG As Integer, ByRef pData As Any) As Integer
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal pDest As Long, ByVal Length As Long)
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateIC Lib "gdi32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSource As Long, ByVal Length As Long)
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (ByRef lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32.dll" (ByRef lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (ByRef lpMsg As MSG) As Long
Private Declare Function DrawDibOpen Lib "msvfw32" () As Long
Private Declare Function DrawDibDraw Lib "msvfw32" (ByVal DDC As Long, ByVal DC As Long, ByVal XDst As Long, ByVal YDst As Long, ByVal WDst As Long, ByVal HDst As Long, pBInfo As BITMAPINFOHEADER, pP As Any, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WSrc As Long, ByVal HSrc As Long, ByVal DDFlags As Long) As Long
Private Declare Function DrawDibClose Lib "msvfw32" (ByVal DDC As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemoryVar Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private m_ColorType As TWAIN_CLASS_COLORTYPE
Private m_lHndMsgWin As Long
Private m_tAppID As TW_IDENTITY
Private m_tSrcID As TW_IDENTITY

Public Function Scan_with_Twain(ByVal ColorType As TWAIN_CLASS_COLORTYPE, ByVal ShowIndicators As Boolean, Display As Object, Optional Resolution As Single = 100, Optional ScanLeft As Single = 0, Optional ScanWidth As Single = -1, Optional ScanTop As Single = 0, Optional ScanHeight As Single = -1, Optional HasUI As Boolean = False) As Long
    Dim lRtn As Long
    Dim lTmp As Long
    Dim blTwainOpen As Boolean
    Dim lhDib As Long
    
    On Local Error GoTo ErrPlace

    '-------------------------------
    ' Open Twain Data Source Manager
    '-------------------------------
    lRtn = OpenTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = True

    '-----------------------
    ' Open Twain Data Source
    '-----------------------
    lRtn = OpenTwainDS()
    If lRtn Then GoTo ErrPlace
        blTwainOpen = True
        '-----------------------------------------------------------
    ' Set all important attributes of the image and the transfer
    '-----------------------------------------------------------
    If HasUI = False Then

    '----------------------------------------------------------------------
    ' Set image size and position
    ' If sngImageRight or sngImageBottom is 0 put physical width and height
    '   of the scanner into these values
    '----------------------------------------------------------------------
    If (ScanWidth = -1) Or (ScanHeight = 0) Then
        lRtn = TwainGetOneValue(PHYSICALWIDTH, ScanWidth)
        If lRtn Then GoTo ErrPlace
        lRtn = TwainGetOneValue(PHYSICALHEIGHT, ScanHeight)
        If lRtn Then GoTo ErrPlace
    End If

    lRtn = SetImageSize(ScanLeft, ScanTop, ScanWidth, ScanHeight)
    If lRtn Then GoTo ErrPlace

    '-----------------------------------------------
    ' Set the image resolution in DPI - both X and Y
    '-----------------------------------------------
    lRtn = TwainSetOneValue(XRESOLUTION, FIX32, Resolution)
    If lRtn Then GoTo ErrPlace

    lRtn = TwainSetOneValue(YRESOLUTION, FIX32, Resolution)
    If lRtn Then GoTo ErrPlace

    '--------------------------
    ' Set the image colour type
    '--------------------------
    lRtn = TwainSetOneValue(PIXELTYPE, UINT16, ColorType)
    If lRtn Then GoTo ErrPlace
    
    '----------------------------------------------------------------
    ' If the colour type is fullcolour, set the bitdepth of the image
    '   - 24 bits, 32 bits, ...
    '----------------------------------------------------------------
    If ColorType = Bitmap_RGB_24bit Then lRtn = TwainSetOneValue(BITDEPTH, UINT16, 24)

    '---------------------------------------------------
    ' Set number of images you want to transfer (just 1)
    '---------------------------------------------------
    lRtn = TwainSetOneValue(XFERCOUNT, INT16, 1)
    If lRtn Then GoTo ErrPlace

    End If

    '----------------------------------------------------
    ' TRANSFER the image with UI disabled or enabled.
    '   If successful, lhDIB is filled with handle to DIB
    '----------------------------------------------------
    lRtn = TwainTransfer(HasUI, lhDib)
    If lRtn Then GoTo ErrPlace
    '------------------
    ' Close Data Source
    '------------------
    lRtn = CloseTwainDS()
    If lRtn Then GoTo ErrPlace
    
    '--------------------------
    ' Close Data Source Manager
    '--------------------------
    lRtn = CloseTwainDSM()
    If lRtn Then GoTo ErrPlace
    blTwainOpen = False
 DisplayDib lhDib, Display, ColorType
Scan_with_Twain = 0
Exit Function
ErrPlace:
    If lhDib Then lRtn = GlobalFree(lhDib)
    If blTwainOpen Then lRtn = CloseTwainDS(): lRtn = CloseTwainDSM()
    Scan_with_Twain = 1

End Function

Private Function OpenTwainDSM() As Long
    
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '----------------------------------------------------
    ' Create window that will receive all TWAIN messages
    ' Message loop can be found in TwainTransfer function
    '----------------------------------------------------
    m_lHndMsgWin = CreateWindowEx(0&, "#32770", "TWAIN_MSG_WINDOW", 0&, _
                                  10&, 10&, 150&, 50&, 0&, 0&, 0&, 0&)
    If m_lHndMsgWin = 0 Then GoTo ErrPlace
    
    '------------------------------------------------------------
    ' Introduce yourself to TWAIN - MajorNum, MinorNum, Language,
    ' Country, Manufacturer, ProductFamily, ProductName, etc.
    '------------------------------------------------------------
    Call ZeroMemory(VarPtr(m_tAppID), Len(m_tAppID))
    With m_tAppID
        .Version.MajorNum = 1
        .Version.Language = TWLG_ENG
        .Version.Country = TWCY_USA
        .ProtocolMajor = TWON_PROTOCOLMAJOR
        .ProtocolMinor = TWON_PROTOCOLMINOR
        SetUnsignedLong .SupportedGroups, DG_CONTROL Or DG_IMAGE
    End With
    
    Call CopyMemory(VarPtr(m_tAppID.Manufacturer(1)), _
                    StrPtr(StrConv("LMik", vbFromUnicode)), _
                    Len("LMik"))
    Call CopyMemory(VarPtr(m_tAppID.ProductFamily(1)), _
                    StrPtr(StrConv("VB Module", vbFromUnicode)), _
                    Len("VB Module"))
    Call CopyMemory(VarPtr(m_tAppID.ProductName(1)), _
                    StrPtr(StrConv("VB Module for TWAIN", vbFromUnicode)), _
                    Len("VB Module for TWAIN"))
    
    '--------------------------------------
    ' Open Data Source Manager
    '   DG_CONTROL, DAT_PARENT, MSG_OPENDSM
    '--------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_PARENT, MSG_OPENDSM, _
                     m_lHndMsgWin)
    If iRtn <> 0 Then GoTo ErrPlace
    
    OpenTwainDSM = 0
    Exit Function
    
ErrPlace:
    OpenTwainDSM = 1
End Function

Private Function OpenTwainDS() As Long

    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    Call ZeroMemory(VarPtr(m_tSrcID), Len(m_tSrcID))
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, MSG_OPENDS, _
                     m_tSrcID)
    If iRtn <> 0 Then GoTo ErrPlace

    OpenTwainDS = 0
    Exit Function
    
ErrPlace:
    OpenTwainDS = 1
End Function

Public Property Get Get_ScreenColorMode() As Long
    Dim lwndHandle As Long
    lwndHandle = CreateIC("DISPLAY", vbNullString, vbNullString, 0&)
    Get_ScreenColorMode = GetDeviceCaps(lwndHandle, 12)
End Property

Private Function TwainGetOneValue(ByVal Cap As TWAIN_MDL_CAPABILITY, _
                                  ByRef Item As Variant) As Long

    '-----------------------------------------------------------------------
    ' There are four types of containers that TWAIN defines for capabilities
    ' (TW_ONEVALUE, TW_ARRAY, TW_RANGE and TW_ENUMERATION)
    ' This module deals with one of them only - TW_ONEVALUE (single value)
    ' To get some capability you have to fill TW_ONEVALUE fields and use
    '   the triplet DG_CONTROL DAT_CAPABILITY MSG_GET
    ' The macros that convert some data types are used here as well
    '-----------------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    
    Dim tCapability As TW_CAPABILITY
    Dim tOneValue As TW_ONEVALUE
    Dim tFix32 As TW_FIX32
    Dim lpOneValue As Long
    Dim lRtn As Long
    Dim iRtn As Integer
    
    tCapability.ConType = TWON_ONEVALUE
    tCapability.Cap = Cap
    
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_GET, _
                     tCapability)
    If iRtn <> 0 Then GoTo ErrPlace
    
    lpOneValue = GlobalLock(tCapability.hContainer)
    Call CopyMemory(VarPtr(tOneValue), lpOneValue, Len(tOneValue))
    lRtn = GlobalUnlock(tCapability.hContainer)
    lRtn = GlobalFree(tCapability.hContainer)

    Select Case tOneValue.ItemType
        Case INT16
            Item = tOneValue.Item1
        Case UINT16, BOOL
            Item = FromUnsignedShort(tOneValue.Item1)
        Case FIX32
            Call CopyMemory(VarPtr(tFix32), VarPtr(tOneValue.Item1), 4&)
            Item = Fix32ToFloat(tFix32)
    End Select
    
    TwainGetOneValue = 0
    Exit Function

ErrPlace:
    TwainGetOneValue = 1
End Function

Private Function FromUnsignedShort(ByRef iSrc As Integer) As Long

    '------------------------------------------------------------------------
    ' Gets the 2-byte unsigned number from VB Integer data type
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns unsigned 2-byte value (in VB Long type)
    '------------------------------------------------------------------------
    
    Dim lTmp As Long
    
    Call CopyMemory(VarPtr(lTmp), VarPtr(iSrc), 2&)
    
    ' Another way
    'lTmp = IIf(iSrc < 0, iSrc + 65536, iSrc)
    
    FromUnsignedShort = lTmp

End Function

Private Function Fix32ToFloat(ByRef tFix32 As TW_FIX32) As Single
        
    '----------------------------------------------------------------
    ' Converts TWAIN TW_FIX32 data structure into VB Single data type
    ' (needed for communicating with TWAIN)
    '
    ' Function returns floating-point number in VB Single data type
    '----------------------------------------------------------------
        
    Dim sngTmp As Single
    
    sngTmp = tFix32.Whole + CSng(FromUnsignedShort(tFix32.Frac) / 65536)
    Fix32ToFloat = sngTmp

End Function

Private Function SetImageSize(ByRef sngLeft As Single, _
                              ByRef sngTop As Single, _
                              ByRef sngRight As Single, _
                              ByRef sngBottom As Single) As Long

    Dim tImageLayout As TW_IMAGELAYOUT
    Dim lRtn As Long
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '-------------------------------------------------------------------
    ' Set the size of the image - in default units
    '   DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET
    '
    ' If you do not select any units the INCHES are selected as default.
    ' The values of Single type are converted into TWAIN TW_FIX32.
    '-------------------------------------------------------------------
    lRtn = FloatToFix32(sngLeft, tImageLayout.Frame.Left)
    If lRtn Then GoTo ErrPlace
    
    lRtn = FloatToFix32(sngTop, tImageLayout.Frame.Top)
    If lRtn Then GoTo ErrPlace
    
    lRtn = FloatToFix32(sngRight, tImageLayout.Frame.Right)
    If lRtn Then GoTo ErrPlace
    
    lRtn = FloatToFix32(sngBottom, tImageLayout.Frame.Bottom)
    If lRtn Then GoTo ErrPlace
    
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET, _
                     tImageLayout)
    If (iRtn <> 0) And (iRtn <> TWRC_CHECKSTATUS) Then GoTo ErrPlace
    
    SetImageSize = 0
    Exit Function
    
ErrPlace:
    SetImageSize = 1
End Function

Private Function FloatToFix32(ByRef sngSrc As Single, _
                              ByRef tFix32 As TW_FIX32) As Long
    
    '----------------------------------------------------------------
    ' Converts VB Single data type into TWAIN TW_FIX32 data structure
    ' (needed for communicating with TWAIN)
    '
    ' Function returns 0 is OK, 1 if an error occurs
    '----------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    
    tFix32.Whole = CInt(Fix(sngSrc))
    Call ToUnsignedShort(CLng(sngSrc * 65536) And 65535, tFix32.Frac)
    FloatToFix32 = 0
    Exit Function

ErrPlace:
    FloatToFix32 = 1
End Function

Private Function ToUnsignedShort(ByRef lSrc As Long, _
                                 ByRef iDst As Integer) As Long
    
    '------------------------------------------------------------------------
    ' Sets number ranging from 0 to 65535 into 2-byte VB Integer
    ' (useful for communicating with other dll that uses unsigned data types)
    '
    ' Function returns 0 is OK, 1 if an error occurs
    '------------------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    
    If (lSrc < 0) Or (lSrc > 65535) Then GoTo ErrPlace
    
    Call CopyMemory(VarPtr(iDst), VarPtr(lSrc), 2&)
    
    ' Another way
    'iDst = IIf(lSrc > 32767, lSrc - 65536, lSrc)
    
    ToUnsignedShort = 0
    Exit Function
    
ErrPlace:
    ToUnsignedShort = 1
End Function

Private Function CloseTwainDS() As Long
    
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '----------------------------------------
    ' Close Data Source
    '   DG_CONTROL, DAT_IDENTITY, MSG_CLOSEDS
    '----------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_IDENTITY, _
                     MSG_CLOSEDS, m_tSrcID)
    If iRtn <> 0 Then GoTo ErrPlace
    
    CloseTwainDS = 0
    Exit Function

ErrPlace:
    CloseTwainDS = 1
End Function

Private Function CloseTwainDSM() As Long
    
    Dim lRtn As Long
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace

    '---------------------------------------
    ' Close Data Source Manager
    '   DG_CONTROL, DAT_PARENT, MSG_CLOSEDSM
    '---------------------------------------
    iRtn = DSM_Entry(m_tAppID, ByVal 0&, DG_CONTROL, DAT_PARENT, MSG_CLOSEDSM, _
                     m_lHndMsgWin)
    If iRtn <> 0 Then
        lRtn = DestroyWindow(m_lHndMsgWin)
        GoTo ErrPlace
    End If

    '---------------------------
    ' Destroy the message window
    '---------------------------
    lRtn = DestroyWindow(m_lHndMsgWin)
    If lRtn = 0 Then GoTo ErrPlace
    
    CloseTwainDSM = 0
    Exit Function

ErrPlace:
    CloseTwainDSM = 1
End Function

Private Function TwainSetOneValue(ByVal Cap As TWAIN_MDL_CAPABILITY, _
                                  ByVal ItemType As TWAIN_MDL_ITEMYPE, _
                                  ByRef Item As Variant) As Long

    '-----------------------------------------------------------------------
    ' There are four types of containers that TWAIN defines for capabilities
    ' (TW_ONEVALUE, TW_ARRAY, TW_RANGE and TW_ENUMERATION)
    ' This module deals with one of them only - TW_ONEVALUE (single value)
    ' To set some capability you have to fill TW_ONEVALUE fields and use
    '   the triplet DG_CONTROL DAT_CAPABILITY MSG_SET
    ' The macros that convert some data types are used here as well
    '-----------------------------------------------------------------------
    On Local Error GoTo ErrPlace
    
    Dim tCapability As TW_CAPABILITY
    Dim tOneValue As TW_ONEVALUE
    Dim lhOneValue As Long
    Dim lpOneValue As Long
    Dim lRtn As Long
    Dim iRtn As Integer
    Dim tFix32 As TW_FIX32
    Dim iTmp As Integer

    tCapability.ConType = TWON_ONEVALUE
    tCapability.Cap = Cap
    
    tOneValue.ItemType = ItemType
    
    Select Case ItemType
        Case INT16
            tOneValue.Item1 = CInt(Item)
        Case UINT16, BOOL
            If ToUnsignedShort(CLng(Item), iTmp) Then GoTo ErrPlace
            Call CopyMemory(VarPtr(tOneValue.Item1), VarPtr(iTmp), 2&)
        Case FIX32
            If FloatToFix32(CSng(Item), tFix32) Then GoTo ErrPlace
            Call CopyMemory(VarPtr(tOneValue.Item1), VarPtr(tFix32), 4&)
    End Select

    lhOneValue = GlobalAlloc(GHND, Len(tOneValue))
    lpOneValue = GlobalLock(lhOneValue)
    Call CopyMemory(lpOneValue, VarPtr(tOneValue), Len(tOneValue))
    lRtn = GlobalUnlock(lhOneValue)
    tCapability.hContainer = lhOneValue

    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_SET, _
                     tCapability)
    If iRtn <> 0 Then
        lRtn = GlobalFree(lhOneValue)
        GoTo ErrPlace
    End If
    lRtn = GlobalFree(lhOneValue)

    TwainSetOneValue = 0
    Exit Function

ErrPlace:
    TwainSetOneValue = 1
End Function

Private Function TwainTransfer(ByRef blShowUI As Boolean, _
                               ByRef lDIBHandle As Long) As Long
                               
    Dim tUI As TW_USERINTERFACE
    Dim tPending As TW_PENDINGXFERS
    Dim lhDib As Long
    Dim tEvent As TW_EVENT
    Dim tMSG As MSG
    Dim lRtn As Long
    Dim iRtn As Integer
    
    On Local Error GoTo ErrPlace
    
    '---------------------------------------------
    ' Set tUI.ShowUI to 1 (show UI) or 0 (hide UI)
    '---------------------------------------------
    With tUI
        .ShowUI = IIf(blShowUI = True, 1, 0)
        .ModalUI = 1
        .hParent = m_lHndMsgWin
    End With
    
    '----------------------------------------------
    ' Enable Data Source User Interface
    '   DG_CONTROL, DAT_USERINTERFACE, MSG_ENABLEDS
    '----------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                     MSG_ENABLEDS, tUI)
    If iRtn <> 0 Then GoTo ErrPlace
    
    '-----------------------------------------------------------------
    ' Process events in the message loop
    '   DG_CONTROL, DAT_EVENT, MSG_PROCESSEVENT
    '
    ' There are two messages we are interested in in this message loop
    '   - MSG_XFERREADY - the data source is ready to transfer
    '   - MSG_CLOSEDSREQ - the data source requests to close itself
    '-----------------------------------------------------------------
    While GetMessage(tMSG, 0&, 0&, 0&)
        Call ZeroMemory(VarPtr(tEvent), Len(tEvent))
        tEvent.pEvent = VarPtr(tMSG)
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_EVENT, _
                         MSG_PROCESSEVENT, tEvent)
        Select Case tEvent.TWMessage
            Case MSG_XFERREADY
                GoTo MSGGET
            Case MSG_CLOSEDSREQ
                GoTo MSGDISABLEDS
        End Select
        lRtn = TranslateMessage(tMSG)
        lRtn = DispatchMessage(tMSG)
    Wend
    
MSGGET:
    '----------------------------------------------------
    ' Start transfer
    '   DG_IMAGE, DAT_IMAGENATIVEXFER, MSG_GET
    '
    ' If transfer is successful you get the handle to DIB
    '----------------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_IMAGE, DAT_IMAGENATIVEXFER, _
                     MSG_GET, lhDib)
    If iRtn <> TWRC_XFERDONE Then
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, _
                         MSG_ENDXFER, tPending)
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                         MSG_DISABLEDS, tUI)
        GoTo ErrPlace
    End If
    
    '--------------------------------------------
    ' End transfer
    '   DG_CONTROL, DAT_PENDINGXFERS, MSG_ENDXFER
    '--------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_PENDINGXFERS, _
                     MSG_ENDXFER, tPending)
    If iRtn <> 0 Then
        iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                         MSG_DISABLEDS, tUI)
        GoTo ErrPlace
    End If
    
MSGDISABLEDS:
    '-----------------------------------------------
    ' Disable Data Source
    '   DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS
    '-----------------------------------------------
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_USERINTERFACE, _
                     MSG_DISABLEDS, tUI)
    If iRtn <> 0 Then GoTo ErrPlace

    lDIBHandle = lhDib
    TwainTransfer = 0
    Exit Function
    
ErrPlace:
    If lhDib Then lRtn = GlobalFree(lhDib)
    lDIBHandle = 0
    TwainTransfer = 1
End Function


Private Function DisplayDib(hDib As Long, Display As Object, Coltype As TWAIN_CLASS_COLORTYPE)
Dim tbih As BITMAPINFOHEADER
Dim BIH2 As BITMAPINFO2
Dim BIH8 As BITMAPINFO8
Dim Dibarray() As Byte
Dim DrawDC As Long
Dim Back As Long
Dim Palarray() As Byte
Dim TWidth As Long
Dim THeight As Long

GetDib hDib, tbih, Dibarray, Palarray, Coltype
TWidth = tbih.biWidth
THeight = tbih.biHeight
Select Case Coltype
Case Bitmap_RGB_24bit
      DrawDC = DrawDibOpen()
      Back = DrawDibDraw(DrawDC, Display.hdc, 0, 0, TWidth, THeight, tbih, Dibarray(0), 0, 0, TWidth, THeight, 0)
      Back = DrawDibClose(DrawDC)
Case Bitmap_BW_1bit
CopyMemoryVar ByVal VarPtr(BIH2.bmiHeader), ByVal VarPtr(tbih), 40
CopyMemoryVar BIH2.bmiColors(0), Palarray(0), 8
      Back = StretchDIBits(Display.hdc, 0, 0, TWidth, THeight, 0, 0, TWidth, THeight, Dibarray(0), BIH2, 0, vbSrcCopy)
Case Bitmap_GREY_8bit, Bitmap_PAL_8bit
CopyMemoryVar ByVal VarPtr(BIH8.bmiHeader), ByVal VarPtr(tbih), 40
CopyMemoryVar BIH8.bmiColors(0), Palarray(0), 256 * 4
      Back = StretchDIBits(Form1.Picture1.hdc, 0, 0, TWidth, THeight, 0, 0, TWidth, THeight, Dibarray(1), BIH8, 0, vbSrcCopy)

End Select
Display.Refresh
DoEvents
End Function


Private Function GetDib(lhDib As Long, tbih As BITMAPINFOHEADER, bDibits() As Byte, Palarray() As Byte, Coltype As TWAIN_CLASS_COLORTYPE)
    Dim lpDIB As Long
    Dim lDibSize As Long
    Dim iFileNum As Integer
    Dim Imagesize As Long
    DoEvents
    Sleep 100
    On Local Error GoTo ErrPlace
    lpDIB = GlobalLock(lhDib)
    Call CopyMemory(VarPtr(tbih), lpDIB, Len(tbih))
    Imagesize = tbih.biSizeImage

    Select Case Coltype
Case Bitmap_RGB_24bit
    ReDim bDibits(Imagesize - 1)
    Call CopyMemory(VarPtr(bDibits(0)), lpDIB + 40, tbih.biSizeImage)
Case Bitmap_BW_1bit
    ReDim Palarray(7) As Byte
    ReDim bDibits(Imagesize - 1)
    Call CopyMemory(VarPtr(Palarray(0)), lpDIB + 40, 8)
    Call CopyMemory(VarPtr(bDibits(0)), lpDIB + 48, Imagesize)
Case Bitmap_PAL_8bit, Bitmap_GREY_8bit
    ReDim Palarray(0 To 1023)
    ReDim bDibits(1 To Imagesize)
    Call CopyMemory(VarPtr(Palarray(0)), lpDIB + 40, 1024)
    Call CopyMemory(VarPtr(bDibits(1)), lpDIB + 1064, Imagesize)
  End Select

    GlobalUnlock lhDib
    If GlobalFree(lhDib) = 0 Then lhDib = 0
Sleep 100
DoEvents
ErrPlace:
End Function


Public Property Get BitdephtAll_Twain() As String
OpenTwainDSM
OpenTwainDS
BitdephtAll_Twain = TwainGetEnumeration(BITDEPTH)
CloseTwainDS
CloseTwainDSM
End Property

Private Function TwainGetEnumeration(ByVal Cap As TWAIN_MDL_CAPABILITY) As String

    '-----------------------------------------------------------------------
    ' There are four types of containers that TWAIN defines for capabilities
    ' (TW_ONEVALUE, TW_ARRAY, TW_RANGE and TW_ENUMERATION)
    ' This module deals with one of them only - TW_ONEVALUE (single value)
    ' To get some capability you have to fill TW_ONEVALUE fields and use
    '   the triplet DG_CONTROL DAT_CAPABILITY MSG_GET
    ' The macros that convert some data types are used here as well
    '-----------------------------------------------------------------------
    
    On Local Error GoTo ErrPlace
    Dim ItemListInt() As Integer
    Dim ItemListLong() As Long
    Dim tCapability As TW_CAPABILITY
    Dim tEnumValue As TW_ENUMERATION
    Dim tFix32 As TW_FIX32
    Dim lpEnumValue As Long
    Dim lRtn As Long
    Dim iRtn As Integer
    Dim Anzahl As Long
    Dim i As Long
    Dim Test As String
    tCapability.ConType = TWON_ENUMERATION
    tCapability.Cap = Cap
    
    iRtn = DSM_Entry(m_tAppID, m_tSrcID, DG_CONTROL, DAT_CAPABILITY, MSG_GET, _
                     tCapability)
    If iRtn <> 0 Then GoTo ErrPlace
    lpEnumValue = GlobalLock(tCapability.hContainer)
    Call CopyMemory(VarPtr(tEnumValue), lpEnumValue, Len(tEnumValue))
    Test = "NumItems:" & CStr(GetUnsignedLong(tEnumValue.NumItems)) & " CurrentIndex:" & CStr(GetUnsignedLong(tEnumValue.CurrentIndex)) & " DefaultIndex:" & CStr(GetUnsignedLong(tEnumValue.DefaultIndex))
    Anzahl = GetUnsignedLong(tEnumValue.NumItems)
        Select Case tEnumValue.ItemType
        Case INT16
            ReDim ItemListInt(Anzahl - 1)
                Call CopyMemory(VarPtr(ItemListInt(0)), lpEnumValue + Len(tEnumValue), Anzahl * 2)
                Test = Test & " "
                For i = 0 To Anzahl - 1
    Test = Test & ItemListInt(i) & "/"
    Next i
            'Item = tEnumValue.Item1
        Case UINT16, BOOL
                    ReDim ItemListInt(Anzahl - 1)
                Call CopyMemory(VarPtr(ItemListInt(0)), lpEnumValue + Len(tEnumValue), Anzahl * 2)
                Test = Test & " ItemList:"
                For i = 0 To Anzahl - 1
    Test = Test & ItemListInt(i) & "/"
Next i
            'Item = FromUnsignedShort(tEnumValue.Item1)
        Case FIX32
            'Call CopyMemory(VarPtr(tFix32), VarPtr(tEnumValue.Item1), 4&)
            'Item = Fix32ToFloat(tFix32)
    End Select

    lRtn = GlobalUnlock(tCapability.hContainer)
    lRtn = GlobalFree(tCapability.hContainer)

    
    TwainGetEnumeration = Test
    Exit Function

ErrPlace:
    TwainGetEnumeration = Test
End Function

Private Sub SetUnsignedLong(LongArr() As Byte, Number As Double)
Dim Testlong As Long
Testlong = UnsignedToLong(Number)
CopyMemoryVar LongArr(0), ByVal VarPtr(Testlong), 4
End Sub

Private Sub SetSignedLong(LongArr() As Byte, Number As Long)
CopyMemoryVar LongArr(0), ByVal VarPtr(Number), 4
End Sub

Private Function GetUnsignedLong(LongArr() As Byte) As Double
Dim Testlong As Long
CopyMemoryVar ByVal VarPtr(Testlong), LongArr(0), 4
GetUnsignedLong = LongToUnsigned(Testlong)
End Function

Private Function GetSignedLong(LongArr() As Byte) As Long
Dim Testlong As Long
CopyMemoryVar ByVal VarPtr(Testlong), LongArr(0), 4
GetUnsignedLong = Testlong
End Function


      Function UnsignedToLong(Value As Double) As Long
        If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
        If Value <= MAXINT_4 Then
          UnsignedToLong = Value
        Else
          UnsignedToLong = Value - OFFSET_4
        End If
      End Function

      Function LongToUnsigned(Value As Long) As Double
        If Value < 0 Then
          LongToUnsigned = Value + OFFSET_4
        Else
          LongToUnsigned = Value
        End If
      End Function

      Function UnsignedToInteger(Value As Long) As Integer
        If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
        If Value <= MAXINT_2 Then
          UnsignedToInteger = Value
        Else
          UnsignedToInteger = Value - OFFSET_2
        End If
      End Function

      Function IntegerToUnsigned(Value As Integer) As Long
        If Value < 0 Then
          IntegerToUnsigned = Value + OFFSET_2
        Else
          IntegerToUnsigned = Value
        End If
      End Function


