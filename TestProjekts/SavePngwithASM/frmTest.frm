VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4368
   ClientLeft      =   1200
   ClientTop       =   1668
   ClientWidth     =   6384
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmTest.frx":0000
   ScaleHeight     =   4368
   ScaleWidth      =   6384
   WindowState     =   2  'Maximiert
   Begin VB.CommandButton Command3 
      Caption         =   "Settings"
      Height          =   372
      Left            =   3480
      TabIndex        =   7
      Top             =   240
      Width           =   972
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      ItemData        =   "frmTest.frx":0152
      Left            =   2280
      List            =   "frmTest.frx":015C
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   240
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "Compress Algorythm"
      Height          =   852
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1692
      Begin VB.OptionButton Option2 
         Caption         =   "VB (slow)"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1452
      End
      Begin VB.OptionButton Option1 
         Caption         =   "zLib-Dll (fast)"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Picture"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1332
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   1800
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   120
      Picture         =   "frmTest.frx":0170
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   1
      Top             =   1440
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Picture in png-File"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Private Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type BITMAPINFOHEADER '40 bytes
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
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&
Private pixels() As Byte
Dim r As Integer
Dim g As Integer
Dim b As Integer



Private Sub Command1_Click()
Dim Alphabyte As Byte
Dim Transarray() As Byte
Dim BitsPP As Long
Dim Filename As String
Dim a As New SavePNG
Dim AlphaArray() As Byte
CommonDialog1.InitDir = App.Path
CommonDialog1.Filename = "Test.png"
CommonDialog1.Filter = "*.png|*.png"
CommonDialog1.ShowSave
Filename = CommonDialog1.Filename
If Filename <> "" Then
ChDir App.Path
Command1.Enabled = False
Command2.Enabled = False
Form1.Refresh
If Combo1.ListIndex = 0 Then
BitsPP = 24
Else
BitsPP = 8
End If
Select Case BitsPP
Case 8
If frm8Bit.chktrans8 Then
a.HasTrans = True
a.SetTransBytes8 TransArray8
End If
Case 24
If frm24Bit.chkTransparent.Value = 1 Then
a.HasTrans = True
a.Transparent24 = frm24Bit.pbTransparentColor.BackColor
End If
If frm24Bit.chkBkgd = 1 Then
a.HasbkgdColor = True
a.bkgdColorLong = frm24Bit.picbkgd.BackColor
End If
If frm24Bit.chkAlphablend = 1 Then
ReDim AlphaArray(Picture1.ScaleHeight * Picture1.ScaleWidth - 1)
Alphabyte = frm24Bit.txtAlphablend
FillMemory AlphaArray(0), UBound(AlphaArray) + 1, Alphabyte
a.HasAlpha = True
a.SetAlphaBytes AlphaArray
End If
End Select
a.SavePNGinFile Filename, Picture1, BitsPP, Option2
Command1.Enabled = True
Command2.Enabled = True
MsgBox "ready"

End If
End Sub

Private Sub Command2_Click()
Dim Filename As String
CommonDialog1.InitDir = App.Path

CommonDialog1.Filter = "*.bmp *.gif *.jpg *.ico|*.bmp; *.gif; *.jpg; *.ico"
CommonDialog1.ShowOpen
Filename = CommonDialog1.Filename
If Filename <> "" Then
Picture1.Picture = LoadPicture(Filename)
FillcolorArray Picture1, pixels
End If

End Sub


Private Sub Command3_Click()
Select Case Combo1.ListIndex
Case 0
frm24Bit.Show
Case 1
frm8Bit.Show
End Select
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
FillcolorArray Picture1, pixels

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Form1.MousePointer = 0
frm24Bit.pbColorunderCursor.BackColor = frm24Bit.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If frm24Bit.chkTransparent = 1 Then
Form1.MousePointer = 99
r = pixels(3, x, y)
g = pixels(2, x, y)
b = pixels(1, x, y)
frm24Bit.pbColorunderCursor.BackColor = RGB(r, g, b)
frm24Bit.Show
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If frm24Bit.chkTransparent = 1 Then
r = pixels(3, x, y)
g = pixels(2, x, y)
b = pixels(1, x, y)

frm24Bit.pbTransparentColor.BackColor = RGB(r, g, b)
End If
End Sub
Private Sub FillcolorArray(ByVal picColor As PictureBox, pixels() As Byte)
Dim bitmap_info As BITMAPINFO
Dim bytes_per_scanLine As Integer
Dim pad_per_scanLine As Integer
Dim x As Integer
Dim y As Integer
Dim ave_color As Byte

    With bitmap_info.bmiHeader
        .biSize = 40
        .biWidth = picColor.ScaleWidth
        .biHeight = -picColor.ScaleHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        bytes_per_scanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
        pad_per_scanLine = bytes_per_scanLine - (((.biWidth * .biBitCount) + 7) \ 8)
        .biSizeImage = bytes_per_scanLine * Abs(.biHeight)
    End With

    ReDim pixels(1 To 4, picColor.ScaleWidth - 1, picColor.ScaleHeight - 1)
    GetDIBits picColor.hdc, picColor.Image, _
        0, picColor.ScaleHeight, pixels(1, 0, 0), _
        bitmap_info, DIB_RGB_COLORS

End Sub

