VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   684
   ClientTop       =   1980
   ClientWidth     =   8232
   LinkTopic       =   "Form1"
   ScaleHeight     =   515
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   686
   WindowState     =   2  'Maximiert
   Begin VB.CommandButton Command2 
      Caption         =   "SaveCursor"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   3840
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      ItemData        =   "Form1.frx":0000
      Left            =   7560
      List            =   "Form1.frx":0002
      TabIndex        =   5
      Text            =   "Select File"
      Top             =   600
      Width           =   3132
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame framePalette 
      Caption         =   "Palette"
      Height          =   7335
      Left            =   7560
      TabIndex        =   2
      Top             =   960
      Width           =   2775
      Begin VB.PictureBox picReverse 
         BackColor       =   &H00C0E0FF&
         Height          =   250
         Left            =   1320
         ScaleHeight     =   204
         ScaleWidth      =   204
         TabIndex        =   9
         Top             =   240
         Width           =   250
      End
      Begin VB.PictureBox picTrans 
         BackColor       =   &H80000001&
         Height          =   250
         Left            =   120
         ScaleHeight     =   204
         ScaleWidth      =   204
         TabIndex        =   8
         Top             =   240
         Width           =   250
      End
      Begin VB.PictureBox Picture4 
         Height          =   250
         Index           =   0
         Left            =   120
         ScaleHeight     =   204
         ScaleWidth      =   204
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.Label Label3 
         Caption         =   "Reverse"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Transparent"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblRGB 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   6840
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      DrawStyle       =   2  'Punkt
      Height          =   2055
      Left            =   120
      ScaleHeight     =   171
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "open *.ico *.cur"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Height          =   252
      Left            =   7560
      TabIndex        =   4
      Top             =   120
      Width           =   3372
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Farbengeladen As Boolean
Public farbanzahlalt As Long
Public Geladen As Boolean
Public oldcolor As Integer
Public Malfarbe As Long
Public Palnummer As Long
Public Transparent As Boolean
Public Reverse As Boolean

Private Sub Combo1_Click()
Picture4(oldcolor).BorderStyle = 1
lblRGB.Caption = ""
Nummer = 0
If Geladen = True Then
If Combo1.ListIndex <> Nummer Then
Nummer = Combo1.ListIndex
End If
End If
Timer1.Interval = 10
End Sub

Private Sub Command1_Click()
Picture4(oldcolor).BorderStyle = 1
lblRGB.Caption = ""
Geladen = False
Combo1.Clear
Dim Filename As String
Dim i As Long
CommonDialog1.Filename = ""
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "*.ico; *.cur|*.ico;*.cur"
CommonDialog1.ShowOpen
Filename = CommonDialog1.Filename
DoEvents
If Filename = "" Then Exit Sub
OpenFile Filename
For i = 0 To Dateiinhalt(0).Grafikmenge - 1
Combo1.AddItem Dateiinhalt(i).Farbenanzahl & " Farben " & "(" & Dateiinhalt(i).BreitePixel & "*" & Dateiinhalt(i).HöhePixel & ")"
Next i
Combo1.ListIndex = 0
Zeichnen
Geladen = True
End Sub

Private Sub Farbenladen()
lblRGB.Caption = ""
Dim i As Long
Dim Farbanzahl As Long
Dim Zeile As Long
Dim Reihe As Long
Dim Reihenanzahl As Long

Select Case Dateiinhalt(Nummer).Farbenanzahl
Case 256
Reihenanzahl = 10
Case Else
Reihenanzahl = 5
End Select

If Dateiinhalt(Nummer).Farbenanzahl > 260 And farbanzahlalt < 260 Then
Picture4(0).Visible = False
For i = 1 To farbanzahlalt
Unload Picture4(i)
Next i
farbanzahlalt = 0
End If

If Dateiinhalt(Nummer).Farbenanzahl < 260 Then
Farbanzahl = UBound(Paletten(Nummer).Palett)

If Farbengeladen = True Then
If Farbanzahl <> farbanzahlalt Then
For i = 1 To farbanzahlalt
Unload Picture4(i)
Next i
End If
End If
Picture4(0).BackColor = RGB(Paletten(Nummer).Palett(0).R, Paletten(Nummer).Palett(0).G, Paletten(Nummer).Palett(0).b)
Zeile = 1
Reihe = 1
For i = 1 To Farbanzahl
If Farbanzahl <> farbanzahlalt Then
Load Picture4(i)
'Wohin zeichnen
Picture4(i).Left = Picture4(i - 1).Left + Picture4(0).Width
Picture4(i).Top = Picture4(i - 1).Top
If Reihe = Reihenanzahl Then
Picture4(i).Left = Picture4(0).Left
Picture4(i).Top = Picture4(i - 1).Top + Picture4(0).Height
Reihe = 0
End If
End If
Picture4(i).BackColor = RGB(Paletten(Nummer).Palett(i).R, Paletten(Nummer).Palett(i).G, Paletten(Nummer).Palett(i).b)
Reihe = Reihe + 1
Next i
framePalette.Visible = False
For i = 0 To Farbanzahl
Picture4(i).Visible = True
Next i

framePalette.Visible = True
End If
Farbengeladen = True
farbanzahlalt = Farbanzahl

End Sub



Private Sub Command2_Click()
CommonDialog1.Filename = "Test.cur"
CommonDialog1.InitDir = App.Path
CommonDialog1.DefaultExt = ".cur"
CommonDialog1.Filter = "*.cur"
CommonDialog1.ShowSave
SaveCursor CommonDialog1.Filename
End Sub

Private Sub picReverse_Click()
Reverse = True
Transparent = False
End Sub

Private Sub picTrans_Click()
Transparent = True
Reverse = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Transparent = False And Reverse = False Then
ZeichneKasten Picture1, GetKästchennummer(X, Y).X, GetKästchennummer(X, Y).Y, Malfarbe
Select Case Dateiinhalt(Nummer).Farbenanzahl
Case Is < 257
ChangeArray GetKästchennummer(X, Y).X, GetKästchennummer(X, Y).Y, Palnummer
Case Else ' 24 Bit
ChangeArray GetKästchennummer(X, Y).X, GetKästchennummer(X, Y).Y, Malfarbe
End Select
End If
If Transparent = True Then
ZeichneKasten Picture1, GetKästchennummer(X, Y).X, GetKästchennummer(X, Y).Y, &H80000001
Select Case Dateiinhalt(Nummer).Farbenanzahl
Case Is < 257
ChangeArray GetKästchennummer(X, Y).X, GetKästchennummer(X, Y).Y, Palnummer, True
Case Else ' 24 Bit
'ChangeArray GetKästchennummer(X, Y).X, GetKästchennummer(X, Y).Y, Malfarbe
End Select
End If
If Reverse = True Then
ZeichneKasten Picture1, GetKästchennummer(X, Y).X, GetKästchennummer(X, Y).Y, &HC0E0FF
Select Case Dateiinhalt(Nummer).Farbenanzahl
Case Is < 257
ChangeArray GetKästchennummer(X, Y).X, GetKästchennummer(X, Y).Y, Palnummer, False, True
Case Else ' 24 Bit
'ChangeArray GetKästchennummer(X, Y).X, GetKästchennummer(X, Y).Y, Malfarbe
End Select
End If

End Sub

Private Sub Zeichnen()
Dim i As Long
Grundzeichnen Picture1, Dateiinhalt(Nummer).BreitePixel, Dateiinhalt(Nummer).HöhePixel, 10, 10
Select Case Dateiinhalt(Nummer).Farbenanzahl
Case 2
Zeichne1Bit (Nummer)
Case 3
Zeichne3Bit (Nummer)
Case 16
Zeichne4Bit (Nummer)
Case 256
Zeichne8Bit (Nummer)
Case 16777216
Zeichne24Bit (Nummer)
End Select
If Dateiinhalt(Nummer).Type = "Cursor" Then
Label2.Caption = "Hotspot: " & Dateiinhalt(Nummer).CursorXHotspot & ", " & Dateiinhalt(Nummer).CursoryHotspot
Else
Label2.Caption = ""
End If
Farbenladen

End Sub

Private Sub Picture4_Click(Index As Integer)
Dim rr, Gr, br As Long
Dim rest As Long
Dim mcolor As Long
Dim blue As Integer
Dim red As Integer
Dim green As Integer

Reverse = False
Transparent = False
Picture4(oldcolor).BorderStyle = 1
mcolor = Picture4(Index).BackColor
rr = 1: Gr = 256: br = 65536

rest = mcolor \ br
blue = rest
mcolor = mcolor Mod br

If blue < 0 Then blue = 0

rest = mcolor \ Gr
green = rest
mcolor = mcolor Mod Gr

If green < 0 Then green = 0

rest = mcolor \ rr
red = rest
mcolor = mcolor Mod rr

If red < 0 Then red = 0
lblRGB.Caption = "RGB: " & red & " " & green & " " & blue
Malfarbe = Picture4(Index).BackColor
Palnummer = Index
Picture4(Index).BorderStyle = 0
oldcolor = Index
End Sub

Private Sub Timer1_Timer()
Zeichnen
Timer1.Interval = 0
End Sub
