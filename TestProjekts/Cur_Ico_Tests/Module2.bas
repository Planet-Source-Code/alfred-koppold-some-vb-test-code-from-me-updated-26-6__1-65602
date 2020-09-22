Attribute VB_Name = "Module1"
Option Explicit
Public Farbengeladen As Boolean
Public farbanzahlalt As Long
Public Sub Farbenladen(PicObj As PictureBox)
Dim i As Long
Dim Farbanzahl As Long
Dim Zeile As Long
Dim Reihe As Long

Farbanzahl = UBound(Paletten(0).Palett)
If Farbengeladen = True Then
For i = 1 To farbanzahlalt
Unload PicObj(i)
Next i
End If
PicObj(0).BackColor = RGB(Paletten(0).Palett(0).R, Paletten(0).Palett(0).G, Paletten(0).Palett(0).b)
Zeile = 1
Reihe = 1
For i = 1 To Farbanzahl
Load PicObj(i)
'Wohin zeichnen

PicObj(i).Left = PicObj(i - 1).Left + PicObj(0).Width
PicObj(i).Top = PicObj(i - 1).Top
If Reihe = 5 Then
PicObj(i).Left = PicObj(0).Left
PicObj(i).Top = PicObj(i - 1).Top + PicObj(0).Height
Reihe = 0
End If
PicObj(i).BackColor = RGB(Paletten(0).Palett(i).R, Paletten(0).Palett(i).G, Paletten(0).Palett(i).b)
PicObj(i).Visible = True
Reihe = Reihe + 1
Next i
Farbengeladen = True
farbanzahlalt = Farbanzahl
End Sub

