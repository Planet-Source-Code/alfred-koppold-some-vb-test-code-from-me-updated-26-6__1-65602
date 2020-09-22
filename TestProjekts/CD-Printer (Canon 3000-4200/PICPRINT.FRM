VERSION 5.00
Begin VB.Form Picprint 
   Caption         =   "Picture Print Demo"
   ClientHeight    =   4620
   ClientLeft      =   2916
   ClientTop       =   1848
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   4620
   ScaleWidth      =   8460
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Text            =   "Teststring"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Text            =   "3"
      Top             =   480
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mit Vorschau"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Value           =   1  'Aktiviert
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2904
      Left            =   0
      Picture         =   "PICPRINT.frx":0000
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   0
      Width           =   3864
   End
   Begin VB.CommandButton CmdPrintAPI 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Print to CD"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Font"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Skalierung der Grafik"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Picprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
        (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount _
        As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Const DT_BOTTOM = &H8
Const DT_CENTER = &H1
Const DT_LEFT = &H0
Const DT_RIGHT = &H2
Const DT_TOP = &H0
Const DT_VCENTER = &H4
Const DT_WORDBREAK = &H10

Const DT_CALCRECT = &H400
Const DT_EDITCONTROL = &H2000
Const DT_NOCLIP = &H100

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private DevmodeAlt As DEVMODE
Private PrivateAlt() As Byte
Private Gesichert As Boolean

Private Sub cboFont_Change()
Text2.Fontname = cboFont.List(cboFont.ListIndex)
End Sub

Private Sub cboFont_Click()
Text2.Fontname = cboFont.List(cboFont.ListIndex)

End Sub

Private Sub cboFont_Scroll()
Text2.Fontname = cboFont.List(cboFont.ListIndex)
End Sub

Private Sub cboFont_Validate(Cancel As Boolean)
Text2.Fontname = cboFont.List(cboFont.ListIndex)
End Sub

Private Sub CmdPrintAPI_Click()
On Error GoTo Wiederherstellen
    Dim DeviceName$
    Dim dm As DEVMODE, dmout As DEVMODE
    Dim bufsize&
    Dim dmInBuf() As Byte
    Dim dmOutBuf() As Byte
    Dim prhdc&
    Dim dinfo As DOCINFO
    Dim docname$
    Dim oldcursor&
    Dim hPrinter&
    Dim res&, di&
    Dim Textrect As RECT
    Dim TextToPrint As String

    ChangeToCDorBack False, Check1
    hPrinter = OpenDefaultPrinter(DeviceName$)
        
    ' Get a copy of the DEVMODE structure for this printer
    ' First find out how big the DEVMODE structure is
    bufsize& = DocumentProperties(hwnd, hPrinter, DeviceName$, 0, 0, 0)
    
    ' Allocate buffers of that size
    ReDim dmInBuf(bufsize&)
    ReDim dmOutBuf(bufsize&)
    
    ' Get the output DEVMODE structure
    res = DocumentProperties(hwnd, hPrinter, DeviceName$, ByVal VarPtr(dmOutBuf(0)), ByVal VarPtr(dmInBuf(0)), DM_OUT_BUFFER)

    ' Copy the data buffer into the DEVMODE structure
    ' CopyMemory dmOutBuf(0), dmout, Len(dmout)
    ' Set the orientation, and set the dmField flag so that
    ' the function will know that it is valid.
    
    'dmout.dmOrientation = DMORIENT_LANDSCAPE
    'dmout.dmFields = dm.dmFields Or DM_ORIENTATION
    ' Now copy the data back to the buffer
    'CopyMemory dmout, dmOutBuf(0), Len(dmout)
    ' We now have need DC to the default printer
    ' This DC is also initialized to landscape mode
    prhdc = CreateDCBynum("WINSPOOL", DeviceName$, vbNullString, ByVal VarPtr(dmOutBuf(0)))
    If prhdc = 0 Then GoTo cleanup2
    ' The DOCINFO structure is the information that the
    ' print manager will show. This also gives you the
    ' opportunity of dumping output to a file.
    docname$ = "Sample Document"
    dinfo.cbSize = Len(dinfo)
    dinfo.lpszDocName = docname$
    dinfo.lpszOutput = vbNullString
    
    ' We set up the abort procdure here
    AbortPrinting% = 0
    di = SetAbortProc(prhdc, AddressOf Callback1_AbortProc)
    
    ' And show the abort form which will be system modal
    AbortForm.Show
    Call BringWindowToTop(AbortForm.hwnd)
    AbortForm.Refresh
    
    Enabled = False ' Disable the main form
   
    ' The usual print sequence here
    di = StartDoc(prhdc, dinfo)
    di = StartPage(prhdc)
    
    PrintBitmap prhdc, 950, 2400, Text1.Text
    Textrect.Left = 1000
    Textrect.Right = 6000
    Textrect.Top = 4000
    Textrect.Bottom = 6000
    TextToPrint = "Testtext:" & vbCrLf & "Dies ist ein Test!"
    DrawText prhdc, TextToPrint, Len(TextToPrint), Textrect, DT_LEFT Or DT_WORDBREAK
    ' The system will spend a long time in the EndPage
    ' function, but it will periodically call the Abort
    ' procedure which in turn triggers the Callback1
    ' AbortProc event.
    drawCircularText prhdc, "Dies ist ein CD-Printtest", 0, 16000, 16000, 16000, 360, cboFont.List(cboFont.ListIndex), 180, 130, 2050

    di = EndPage(prhdc)
    If di >= 0 Then di = EndDocAPI(prhdc)
    Unload AbortForm
    Enabled = True
    
cleanup2:
    If prhdc <> 0 Then di = DeleteDC(prhdc)
    If hPrinter <> 0 Then Call ClosePrinter(hPrinter)
Wiederherstellen:
ChangeToCDorBack True
End Sub

'   This function retrieves the definition of the default
'   printer on this system
'
Private Function GetDefPrinter$()
    Dim def$
    Dim di&

    def$ = String$(128, 0)
    di = GetProfileString("WINDOWS", "DEVICE", "", def$, 127)
    GetDefPrinter$ = Mid(def$, 1, InStr(def, Chr(0)) - 1)

End Function

'   This function returns the driver module name
'
Private Function GetDeviceDriver$(dev$)
    Dim firstpos%, nextpos%
    firstpos% = InStr(dev$, ",")
    nextpos% = InStr(firstpos% + 1, dev$, ",")
    GetDeviceDriver$ = Mid$(dev$, firstpos% + 1, nextpos% - firstpos% - 1)
End Function

'   Retrieves the name portion of a device string
'
Private Function GetDeviceName$(dev$)
    Dim npos%
    npos% = InStr(dev$, ",")
    GetDeviceName$ = Left$(dev$, npos% - 1)
End Function

'   Returns the output destination for the specified device
'
Private Function GetDeviceOutput$(dev$)
    Dim firstpos%, nextpos%
    firstpos% = InStr(dev$, ",")
    nextpos% = InStr(firstpos% + 1, dev$, ",")
    GetDeviceOutput$ = Mid$(dev$, nextpos% + 1)
End Function


Private Sub Combo1_Chan()
End Sub

Private Sub Form_Load()
Dim x As Long
Dim Startfont As Long
Gesichert = False
    For x = 0 To Printer.FontCount
        cboFont.AddItem Printer.Fonts(x)
        If Printer.Fonts(x) = "Arial" Then Startfont = x
    Next x
    cboFont.RemoveItem (0)
    If Startfont = 0 Then
    cboFont.ListIndex = 0
    Else
    cboFont.ListIndex = Startfont - 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Gesichert = True Then
ChangeToCDorBack True
End If
End Sub


'   Prints the bitmap in the picture1 control to the
'   printer context specified.
'
Private Sub PrintBitmap(hdc&, PointX As Long, PointY As Long, Skalierung As Long)
    Dim bi As BITMAPINFO
    Dim dctemp&, dctemp2&
    Dim msg$
    Dim bufsize&
    Dim bm As BITMAP
    Dim ghnd&
    Dim gptr&
    Dim xpix&, ypix&
    Dim doscale&
    Dim uy&, ux&
    Dim di&

    ' Create a temporary memory DC and select into it
    ' the background picture of the picture1 control.
    dctemp& = CreateCompatibleDC(Picture1.hdc)
    
    ' Get the size of the picture bitmap
    di = GetObjectAPI(Picture1.Picture, Len(bm), bm)

    ' Can this printer handle the DIB?
    If (GetDeviceCaps(hdc, RASTERCAPS)) And RC_DIBTODEV = 0 Then
        msg$ = "This device does not support DIB's" + vbCrLf + "See source code for further info"
        MsgBox msg$, 0, "No DIB support"
    End If

    ' Fill the BITMAPINFO for the desired DIB
    bi.bmiHeader.biSize = Len(bi.bmiHeader)
    bi.bmiHeader.biWidth = bm.bmWidth
    bi.bmiHeader.biHeight = bm.bmHeight
    bi.bmiHeader.biPlanes = 1
    ' Set to 24 here to create a 24 bit DIB
    ' Set to 8 here to create an 8 bit DIB
    bi.bmiHeader.biBitCount = 4
    bi.bmiHeader.biCompression = BI_RGB
    ' Now calculate the data buffer size needed
    bufsize& = bi.bmiHeader.biWidth

    ' Figure out the number of bytes based on the
    ' number of pixels in each byte. In this case we
    ' really don't need all this code because this example
    ' always uses a 16 color DIB, but the code is shown
    ' here for your future reference
    Select Case bi.bmiHeader.biBitCount
        Case 1
            bufsize& = (bufsize& + 7) / 8
        Case 4
            bufsize& = (bufsize& + 1) / 2
        Case 24
            bufsize& = bufsize& * 3
    End Select
    ' And make sure it aligns on a long boundary
    bufsize& = ((bufsize& + 3) / 4) * 4
    ' And multiply by the # of scan lines
    bufsize& = bufsize& * bi.bmiHeader.biHeight

    ' Now allocate a buffer to hold the data
    ' We use the global memory pool because this buffer
    ' could easily be above 64k bytes.
    ghnd = GlobalAlloc(GMEM_MOVEABLE, bufsize&)
    gptr& = GlobalLock&(ghnd)

    di = GetDIBits(dctemp, Picture1.Picture, 0, bm.bmHeight, ByVal gptr&, bi, DIB_RGB_COLORS)
    'di = SetDIBitsToDevice(hdc, PointX, PointY, bm.bmWidth, bm.bmHeight, 0, 0, 0, bm.bmHeight, ByVal gptr&, bi, DIB_RGB_COLORS)
    
    ' Now see if we can also print a scaled version
    xpix = GetDeviceCaps(hdc, HORZRES)
    ' We subtract off the size of the bitmap already
    ' printed, plus some extra space
    ypix = GetDeviceCaps(hdc, VERTRES) - (bm.bmHeight + 50)

    ' Find out the largest multiplier we can use and still
    ' fit on the page
    doscale = xpix / bm.bmWidth
    If (ypix / bm.bmHeight < doscale) Then doscale = ypix / bm.bmHeight
    If doscale > 1 Then
    If doscale >= Skalierung Then
        doscale = Skalierung
        Else
        doscale = doscale 'größte Möglichkeit
        End If
    Else
    doscale = 1 'normal nicht
    End If
        ux = bm.bmWidth * doscale
        uy = bm.bmHeight * doscale
        ' Now how this is offset a bit so that we don't
        ' print over the 1:1 scaled bitmap
        di = StretchDIBits(hdc, PointX, PointY, ux, uy, 0, 0, bm.bmWidth, bm.bmHeight, ByVal gptr&, bi, DIB_RGB_COLORS, SRCCOPY)
    ' Dump the global memory block
    di = GlobalUnlock(ghnd)
    di = GlobalFree(ghnd)
    di = DeleteDC(dctemp)

End Sub

' Shows information about the current device mode
'
Private Sub ShowDevMode(dm As DEVMODE)
    Dim crlf$
    Dim a$

    crlf$ = Chr$(13) + Chr$(10)
    a$ = "Device name = " + Mid(dm.dmDeviceName, 1, InStr(dm.dmDeviceName, Chr(0)) - 1) + crlf$
    a$ = a$ + "Devmode Version: " + Hex$(dm.dmSpecVersion) + ", Driver version: " + Hex$(dm.dmDriverVersion) + crlf$
    a$ = a$ + "Orientation: "
    If dm.dmOrientation = DMORIENT_PORTRAIT Then a$ = a$ + "Portrait" Else a$ = a$ + "Landscape"
    a$ = a$ + crlf$
    a$ = a$ + "Field mask = " + Hex$(dm.dmFields) + crlf$
    a$ = a$ + "Copies = " + Str$(dm.dmCopies) + crlf$
    If dm.dmFields And DM_YRESOLUTION <> 0 Then
        a$ = a$ + "X,Y resolution = " + Str$(dm.dmPrintQuality) + "," + Str$(dm.dmYResolution) + crlf$
    End If
    MsgBox a$, 0, "Devmode structure"
End Sub

Public Function OpenDefaultPrinter(Optional DeviceName) As Long
    Dim dev$, devname$, devoutput$
    Dim hPrinter&, res&
    Dim pdefs As PRINTER_DEFAULTS
    
    pdefs.pDatatype = vbNullString
    pdefs.pDevMode = 0
    pdefs.DesiredAccess = PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE

    
    dev$ = GetDefPrinter$() ' Get default printer info
    
    If dev$ = "" Then Exit Function
    devname$ = GetDeviceName$(dev$)
    devoutput$ = GetDeviceOutput$(dev$)
    
    If Not IsMissing(DeviceName) Then
        DeviceName = devname$
    End If
    
    ' You can use OpenPrinterBynum to pass a zero as the
    ' third parameter, but you won't have full access to
    ' edit the printer properties
    res& = OpenPrinter(devname$, hPrinter, pdefs)
    If res <> 0 Then OpenDefaultPrinter = hPrinter
End Function

Private Sub ChangeToCDorBack(Optional Back As Boolean = False, Optional MitVorschau As Boolean = False)
    Dim PrinterHandle As Long
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim Formname As String * 32
    Dim MyDevMode As DEVMODE
    Dim Result As Long
    Dim Needed As Long
    Dim PrivateNeu() As Byte
    Dim pFullDevMode As Long
    Dim pi2_buffer() As Long     'This is a block of memory for the Printer_Info_2 structure
        'If you need to use the Printer_Info_2 User Defined Type, the
        '  definition of Printer_Info_2 in the API viewer is incorrect.
        '  pDevMode and pSecurityDescriptor should be defined As Long.
    Formname = "CD-R-Fach B" & Chr(0) & Chr(30) & Chr(5) & Chr(100) & Chr(0) & Chr(1) & Chr(0) & Chr(7) & Chr(1) & Chr(253) & Chr(255) & Chr(2) & Chr(0) & Chr(1) & Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(0) & Chr(0)
    PrinterName = Printer.DeviceName
    If PrinterName = "" Then
        Exit Sub
    End If
    
    pd.pDatatype = vbNullString
    pd.pDevMode = 0&
    'Printer_Access_All is required for NT security
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    
    Result = OpenPrinter(PrinterName, PrinterHandle, pd)
    
    'The first call to GetPrinter gets the size, in bytes, of the buffer needed.
    'This value is divided by 4 since each element of pi2_buffer is a long.
    Result = GetPrinter(PrinterHandle, 2, ByVal 0&, 0, Needed)
    ReDim pi2_buffer((Needed \ 4))
    Result = GetPrinter(PrinterHandle, 2, pi2_buffer(0), Needed, Needed)
    
    'The seventh element of pi2_buffer is a Pointer to a block of memory
    '  which contains the full DevMode (including the PRIVATE portion).
    pFullDevMode = pi2_buffer(7)
    
    'Copy the Public portion of FullDevMode into our DevMode structure
    If Gesichert = False Then
    Gesichert = True 'zur Wiederherstellung der alten Einstellungen
    Call CopyMemory(DevmodeAlt, ByVal pFullDevMode, Len(DevmodeAlt))
    End If
Select Case Back
Case False 'Auf CD umstellen
    Call CopyMemory(MyDevMode, ByVal pFullDevMode, Len(MyDevMode))
    ReDim PrivateAlt(MyDevMode.dmDriverExtra - 1) 'zur Wiederherstellung der alten Eigenschaften
    ReDim PrivateNeu(MyDevMode.dmDriverExtra - 1)
    Call CopyMemory(PrivateAlt(0), ByVal pFullDevMode + 146 + 2, MyDevMode.dmDriverExtra) '2 Bytes dazwischen
    Call CopyMemory(PrivateNeu(0), PrivateAlt(0), MyDevMode.dmDriverExtra)

        'Make desired changes
        Select Case MyDevMode.dmSpecVersion
        Case 1024 '0004 - Win 98 / NT
        MyDevMode.dmDuplex = 1
    MyDevMode.dmFields = DM_DUPLEX Or DM_ORIENTATION Or DM_PAPERLENGTH Or DM_PAPERSIZE Or DM_PAPERWIDTH Or DM_DEFAULTSOURCE
    MyDevMode.dmOrientation = 1
    MyDevMode.dmPaperLength = 2387
    MyDevMode.dmPaperSize = 304
    MyDevMode.dmPaperWidth = 1310
    MyDevMode.dmDefaultSource = 260
    PrivateNeu(48) = 31
    PrivateNeu(50) = 4
    PrivateNeu(54) = 5
    PrivateNeu(57) = 5
    PrivateNeu(72) = 48
    PrivateNeu(73) = 1
    PrivateNeu(74) = 30
    PrivateNeu(75) = 5
    PrivateNeu(76) = 83
    PrivateNeu(77) = 9
    PrivateNeu(78) = 48
    PrivateNeu(79) = 1
    PrivateNeu(80) = 30
    PrivateNeu(81) = 5
    PrivateNeu(82) = 83
    PrivateNeu(83) = 9
    Select Case MitVorschau
    Case False
    PrivateNeu(535) = 0 'ohne Vorschau
    Case True
    PrivateNeu(535) = 129 'mit Vorschau
    End Select
    PrivateNeu(588) = 0
    PrivateNeu(590) = 104

        Case 1025 '0104 - XP / 2000
    MyDevMode.dmDuplex = 1
    MyDevMode.dmFields = DM_DUPLEX Or DM_ORIENTATION Or DM_PAPERLENGTH Or DM_PAPERSIZE Or DM_PAPERWIDTH Or DM_DEFAULTSOURCE
    MyDevMode.dmOrientation = 1
    MyDevMode.dmPaperLength = 2387
    MyDevMode.dmPaperSize = 129
    MyDevMode.dmPaperWidth = 1310
    MyDevMode.dmDefaultSource = 263
    MyDevMode.dmMediaType = 287
    MyDevMode.dmFields = 59039491
    MyDevMode.dmFormName = Formname
    'PrivateNeu(48) = 44
    'PrivateNeu(49) = 51
    'PrivateNeu(52) = 67
    'PrivateNeu(53) = 93
    'PrivateNeu(64) = 132
    'PrivateNeu(65) = 48
    'PrivateNeu(68) = 35
    'PrivateNeu(69) = 90
    'PrivateNeu(80) = 132
    'PrivateNeu(81) = 48
    'PrivateNeu(84) = 35
    'PrivateNeu(85) = 90
    'PrivateNeu(88) = 44
    'PrivateNeu(89) = 51
    'PrivateNeu(92) = 67
    'PrivateNeu(93) = 93
    'PrivateNeu(112) = 132
    'PrivateNeu(113) = 48
    'PrivateNeu(116) = 35
    'PrivateNeu(117) = 90
    'PrivateNeu(152) = 132
    'PrivateNeu(153) = 48
    'PrivateNeu(156) = 35
    'PrivateNeu(157) = 90
    Select Case MitVorschau
    Case False
    PrivateNeu(904) = 0 'ohne Vorschau
    PrivateNeu(908) = 0
    Case True
    PrivateNeu(904) = 1 'mit Vorschau
    PrivateNeu(908) = 1
    End Select
    
    End Select
    
    'Copy our DevMode structure back into FullDevMode
    Call CopyMemory(ByVal pFullDevMode, MyDevMode, Len(MyDevMode))
    Call CopyMemory(ByVal pFullDevMode + 146 + 2, PrivateNeu(0), MyDevMode.dmDriverExtra)

Case True 'ursprüngliche Einstellungen wiederherstellen
    Call CopyMemory(ByVal pFullDevMode, DevmodeAlt, Len(DevmodeAlt))
    Call CopyMemory(ByVal pFullDevMode + 146 + 2, PrivateAlt(0), MyDevMode.dmDriverExtra)
    Gesichert = False
End Select
    'Copy our changes to "the PUBLIC portion of the DevMode" into "the PRIVATE portion of the DevMode"
    Result = DocumentProperties(hwnd, PrinterHandle, PrinterName, ByVal pFullDevMode, ByVal pFullDevMode, DM_IN_BUFFER Or DM_OUT_BUFFER)
    
    'Update the printer's default properties (to verify, go to the Printer folder
    '  and check the properties for the printer)
    Result = SetPrinter(PrinterHandle, 2, pi2_buffer(0), 0&)
    
    Call ClosePrinter(PrinterHandle)
    
    'Note: Once "Set Printer = " is executed, anywhere in the code, after that point
    '      changes made with SetPrinter will ONLY affect the system-wide printer  --
    '      -- the changes will NOT affect the VB printer object.
    '      Therefore, it may be necessary to reset the printer object's parameters to
    '      those chosen in the devmode.
    Dim p As Printer
    For Each p In Printers
        If p.DeviceName = PrinterName Then
            Set Printer = p
            Exit For
        End If
    Next p
End Sub

