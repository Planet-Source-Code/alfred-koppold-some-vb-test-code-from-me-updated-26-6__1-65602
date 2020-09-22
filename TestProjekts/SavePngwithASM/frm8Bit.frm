VERSION 5.00
Begin VB.Form frm8Bit 
   AutoRedraw      =   -1  'True
   Caption         =   "8 Bits Per Pixel"
   ClientHeight    =   4368
   ClientLeft      =   1776
   ClientTop       =   2280
   ClientWidth     =   7212
   LinkTopic       =   "Form2"
   ScaleHeight     =   4368
   ScaleWidth      =   7212
   Begin VB.CheckBox chktrans8 
      Caption         =   "Have Transparence"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   1692
   End
   Begin VB.Label lblG 
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   500
   End
   Begin VB.Label lblB 
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   500
   End
   Begin VB.Label lblR 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   500
   End
   Begin VB.Label lblIndex 
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   500
   End
End
Attribute VB_Name = "frm8Bit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type RGBTRIPLE
        rgbtBlue As Byte
        rgbtGreen As Byte
        rgbtRed As Byte
End Type
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type ZPOINTS
Point1 As POINTAPI
Point2 As POINTAPI
End Type
Dim Zeichenpoints() As ZPOINTS
Dim AnzPalPLine As Long
Dim BPalette() As RGBTRIPLE
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)



Private Sub Form_Load()
ReDim TransArray8(255)
FillMemory TransArray8(0), 256, 255
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim anfangx As Long
Dim endx As Long
Dim anfangy As Long
Dim endy As Long
Dim AnzLines As Long
Dim Linenumber As Long
Dim Test As Long
Dim i As Long
Dim xKas As Long
Dim yKas As Long
Dim x1 As Long
Dim y1 As Long
Dim p0 As Long
Dim p100 As Long
Dim prozp As Long
Dim yLine As Long
Dim KasNr As Long
Dim Back As String
Dim Teil As Long
anfangx = 700
anfangy = 200
AnzLines = 256 \ AnzPalPLine
If 256 Mod AnzPalPLine <> 0 Then AnzLines = AnzLines + 1
endx = Zeichenpoints(AnzPalPLine - 1).Point2.X
endy = Zeichenpoints(255).Point2.Y
If X > anfangx And Y > anfangy And X < endx And Y < endy Then
For i = 0 To AnzPalPLine - 1 'x
Test = -1
If X < Zeichenpoints(i).Point2.X Then
If X > Zeichenpoints(i).Point1.X Then
Test = i
xKas = i
Exit For
End If
End If
Next i

If Test <> -1 Then
For i = 0 To AnzLines - 1 'y
Test = -1
If Y < Zeichenpoints(i * AnzPalPLine).Point2.Y Then
If Y > Zeichenpoints(i * AnzPalPLine).Point1.Y Then
Test = i
yKas = i
Exit For
End If
End If
Next i

End If
If Test <> -1 Then
KasNr = (yKas * AnzPalPLine) + xKas
If KasNr < 256 Then
    Line (Zeichenpoints(KasNr).Point1.X, Zeichenpoints(KasNr).Point1.Y)-(Zeichenpoints(KasNr).Point2.X, Zeichenpoints(KasNr).Point2.Y), vbRed, B
Back = InputBox("Set Transparence for Palettennumber " & KasNr, "Please give in a Number from 0 to 255", "255")
    Line (Zeichenpoints(KasNr).Point1.X, Zeichenpoints(KasNr).Point1.Y)-(Zeichenpoints(KasNr).Point2.X, Zeichenpoints(KasNr).Point2.Y), vbBlack, B
If IsNumeric(Back) Then
If Back >= 0 And Back <= 255 Then
TransArray8(KasNr) = CByte(Back)
   'x1 = Zeichenpoints(KasNr).Point1.X
   'y1 = Zeichenpoints(KasNr).Point1.Y
'Me.Line (x1, y1)-(x1, y1 - 60)

End If
End If
End If
End If
End If
Zeichne
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim anfangx As Long
Dim endx As Long
Dim anfangy As Long
Dim endy As Long
Dim AnzLines As Long
Dim Linenumber As Long
Dim Test As Long
Dim i As Long
Dim xKas As Long
Dim yKas As Long
Dim yLine As Long
Dim KasNr As Long
anfangx = 700
anfangy = 200

AnzLines = 256 \ AnzPalPLine
If 256 Mod AnzPalPLine <> 0 Then AnzLines = AnzLines + 1
endx = Zeichenpoints(AnzPalPLine - 1).Point2.X
endy = Zeichenpoints(255).Point2.Y
If X > anfangx And Y > anfangy And X < endx And Y < endy Then
For i = 0 To AnzPalPLine - 1 'x
Test = -1
If X < Zeichenpoints(i).Point2.X Then
If X > Zeichenpoints(i).Point1.X Then
Test = i
xKas = i
Exit For
End If
End If
Next i

If Test <> -1 Then
For i = 0 To AnzLines - 1 'y
Test = -1
If Y < Zeichenpoints(i * AnzPalPLine).Point2.Y Then
If Y > Zeichenpoints(i * AnzPalPLine).Point1.Y Then
Test = i
yKas = i
Exit For
End If
End If
Next i

End If
If Test <> -1 Then
KasNr = (yKas * AnzPalPLine) + xKas
If KasNr < 256 Then
lblIndex.Caption = "I: " & KasNr
lblR.Caption = "R: " & BPalette(KasNr).rgbtRed
lblG.Caption = "G: " & BPalette(KasNr).rgbtGreen
lblB.Caption = "B: " & BPalette(KasNr).rgbtBlue
Else
lblIndex.Caption = ""
lblR.Caption = ""
lblG.Caption = ""
lblB.Caption = ""
End If
Else
lblIndex.Caption = ""
lblR.Caption = ""
lblG.Caption = ""
lblB.Caption = ""
End If
Else
lblIndex.Caption = ""
lblR.Caption = ""
lblG.Caption = ""
lblB.Caption = ""
End If
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Form_Paint()
Zeichne
End Sub

Private Sub Form_Resize()
Zeichne
End Sub

Private Sub lblB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblIndex.Caption = ""
lblR.Caption = ""
lblG.Caption = ""
lblB.Caption = ""
End Sub

Private Sub lblG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblIndex.Caption = ""
lblR.Caption = ""
lblG.Caption = ""
lblB.Caption = ""
End Sub

Private Sub lblIndex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblIndex.Caption = ""
lblR.Caption = ""
lblG.Caption = ""
lblB.Caption = ""
End Sub

Private Sub lblR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblIndex.Caption = ""
lblR.Caption = ""
lblG.Caption = ""
lblB.Caption = ""
End Sub

Private Sub Zeichne()
Dim Teil As Single
Dim a As New SavePNG
Dim p0 As Long
Dim p100 As Long
Dim prozp As Long
Dim Zähler As Long
Dim x1 As Long
Dim y1 As Long
   Dim CX As Long
   Dim cy As Long
   Dim dx As Long
   Dim dy As Long
   Dim i As Long
   Dim BMA() As Byte
   Dim Übergabe As Single
Dim PALA() As Byte
Me.Cls
AnzPalPLine = 20
ReDim Zeichenpoints(255)
a.GetBitmapData Form1.Picture1, 8, BMA, PALA
ReDim BPalette(255)
   For i = 0 To 255
   BPalette(i).rgbtRed = PALA(Zähler)
   BPalette(i).rgbtGreen = PALA(Zähler + 1)
   BPalette(i).rgbtBlue = PALA(Zähler + 2)
   Zähler = Zähler + 3
   Next i
   Zähler = 0
   DrawWidth = 1   ' DrawWidth setzen.
   CX = 700
   cy = 200
   dx = 950
   dy = 400
   For i = 1 To 256
   Zeichenpoints(i - 1).Point1.X = CX
   Zeichenpoints(i - 1).Point1.Y = cy
   Zeichenpoints(i - 1).Point2.X = dx
   Zeichenpoints(i - 1).Point2.Y = dy
    Line (CX, cy)-(dx, dy), RGB(BPalette(i - 1).rgbtRed, BPalette(i - 1).rgbtGreen, BPalette(i - 1).rgbtBlue), BF
    Line (CX, cy)-(dx, dy), vbBlack, B

Zähler = Zähler + 3
CX = CX + 300
dx = dx + 300
If i Mod AnzPalPLine = 0 Then
cy = cy + 300
dy = dy + 300
CX = 700
dx = 950
End If
   Next i
For i = 0 To 255
p0 = Zeichenpoints(i).Point1.X
p100 = Zeichenpoints(i).Point2.X
prozp = p100 - p0
If TransArray8(i) > 0 Then
   Teil = 255 / TransArray8(i)
   Übergabe = prozp / Teil
   x1 = CLng(Übergabe) + p0
   Else
   x1 = p0
   End If
   
   y1 = Zeichenpoints(i).Point1.Y
Me.Line (x1, y1)-(x1, y1 - 60)
Next i

End Sub
