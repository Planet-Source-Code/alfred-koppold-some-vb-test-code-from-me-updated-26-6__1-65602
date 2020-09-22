VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontViewDLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Schriftarten wählen.
Private m_Canceled As Boolean
Private OwnerVAR As Long
Private BoldVAR As Boolean
Private ItalicVAR As Boolean
Private UnderlineVAR As Boolean
Private StrikeOutVAR As Boolean
Private ColorVAR As Long
Private FontNameVAR As String
Private FontSizeVAR As Long
Private CancelErrorVAR As Boolean
Private Const FW_NORMAL = 400
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_ROMAN = 16

Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const REGULAR_FONTTYPE = &H400

Private Const FW_BOLD = 700

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Const LF_FACESIZE = 32


Private Type LOGFONT
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
    lfFaceName As String * 31
End Type

Private Type FONTSTRUC
    lStructSize As Long
    hWnd As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Declare Function ChooseFont _
                         Lib "comdlg32.dll" Alias "ChooseFontA" ( _
                         pChoosefont As FONTSTRUC) As Long

Private Declare Function GlobalAlloc _
                         Lib "kernel32" ( _
                         ByVal wFlags As Long, _
                         ByVal dwBytes As Long) As Long

Private Declare Function GlobalLock _
                         Lib "kernel32" ( _
                         ByVal hMem As Long) As Long

Private Declare Sub CopyMemory _
                    Lib "kernel32" Alias "RtlMoveMemory" ( _
                    hpvDest As Any, _
                    hpvSource As Any, _
                    ByVal cbCopy As Long)

Public Property Get CancelError() As Boolean
 CancelError = CancelErrorVAR
End Property

Public Property Let CancelError(Wert As Boolean)
 CancelErrorVAR = Wert
End Property



Public Property Get Italic() As Boolean
 Italic = ItalicVAR
End Property

Public Property Get StrikeOut() As Boolean
 StrikeOut = StrikeOutVAR
End Property


Public Property Get Underline() As Boolean
 Underline = UnderlineVAR
End Property



Public Property Get Bold() As Boolean
 Bold = BoldVAR
End Property



Public Function Show(Optional FontColor As Long = vbBlack, Optional Fontheight As Long = 10, Optional SchriftName As String = "Arial", Optional FontItalic As Boolean = False, Optional FontBold As Boolean = False, Optional FontUnderline As Boolean = False, Optional FontStrikeout As Boolean = False) As Boolean

Dim LogFontVAR As LOGFONT, FontStrucVAR As FONTSTRUC
Dim lLogFontAddress As Long, lMemHandle As Long
    Dim hMem As Long
    Dim pMem As Long
ColorVAR = FontColor
FontSizeVAR = Fontheight
FontNameVAR = SchriftName
ItalicVAR = FontItalic
BoldVAR = FontBold
UnderlineVAR = FontUnderline
StrikeOutVAR = FontStrikeout

    LogFontVAR.lfUnderline = FontUnderline
    LogFontVAR.lfStrikeOut = FontStrikeout
    LogFontVAR.lfItalic = FontItalic
    Select Case BoldVAR
    Case False
    LogFontVAR.lfWeight = FW_NORMAL 'Standard
    Case True
    LogFontVAR.lfWeight = FW_BOLD  ' Bold
    End Select
    LogFontVAR.lfEscapement = 0  ' angle between baseline and escapement vector
    LogFontVAR.lfOrientation = 0  ' angle between baseline and orientation vector
    LogFontVAR.lfWidth = 0 ' determine default width
    LogFontVAR.lfCharSet = DEFAULT_CHARSET  ' use default character set
    LogFontVAR.lfOutPrecision = OUT_DEFAULT_PRECIS  ' default precision mapping
    LogFontVAR.lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
    LogFontVAR.lfQuality = DEFAULT_QUALITY  ' default quality setting
    LogFontVAR.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN  ' default pitch, proportional with serifs
    LogFontVAR.lfFaceName = FontName & vbNullChar  ' string must be null-terminated
    LogFontVAR.lfHeight = Fontsize * -20 / Screen.TwipsPerPixelY 'Fontsize: hier 10
 lMemHandle = GlobalAlloc(GHND, Len(LogFontVAR))
 If lMemHandle = 0 Then Exit Function
 lLogFontAddress = GlobalLock(lMemHandle)
 If lLogFontAddress = 0 Then Exit Function
 
 
 CopyMemory ByVal lLogFontAddress, LogFontVAR, Len(LogFontVAR)
 
    FontStrucVAR.lStructSize = Len(FontStrucVAR)  ' size of structure
    FontStrucVAR.lpLogFont = lLogFontAddress   ' pointer to LOGFONT memory block buffer
    FontStrucVAR.iPointSize = 120  ' 12 point font (in units of 1/10 point)
    FontStrucVAR.Flags = CF_SCREENFONTS Or CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
    FontStrucVAR.rgbColors = FontColor 'RGB(0, 0, 0)  ' black
    FontStrucVAR.nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
    FontStrucVAR.nSizeMin = 6  ' minimum point size
    FontStrucVAR.nSizeMax = 72  ' maximum point size
With LogFontVAR
 If ChooseFont(FontStrucVAR) = 1 Then
  CopyMemory LogFontVAR, ByVal lLogFontAddress, Len(LogFontVAR)
  
  If .lfWeight >= FW_BOLD Then BoldVAR = True Else BoldVAR = False
  If .lfItalic <> 0 Then ItalicVAR = True Else ItalicVAR = False

  If .lfUnderline = 1 Then UnderlineVAR = True Else UnderlineVAR = False
        
  If .lfStrikeOut = 1 Then StrikeOutVAR = True Else StrikeOutVAR = False
  
  ColorVAR = FontStrucVAR.rgbColors
  
  FontNameVAR = .lfFaceName
  
  FontSizeVAR = CLng(FontStrucVAR.iPointSize / 10)
        
  If Len(FontNameVAR) > 0 Then Show = True
 Else
  Show = False
  m_Canceled = True
  If CancelError Then
   Err.Clear
   Err.Raise 18
  End If
 End If
End With
End Function
Public Property Get Fontsize() As Long
 Fontsize = FontSizeVAR
End Property



Public Property Get FontName() As String
 FontName = FontNameVAR
End Property

Public Property Get Owner() As Long
 Owner = OwnerVAR
End Property

Public Property Get Color() As Long
 Color = ColorVAR
End Property

Public Property Let Owner(Wert As Long)
 OwnerVAR = Wert
End Property


Public Property Get Canceled() As Boolean
Canceled = m_Canceled
End Property

