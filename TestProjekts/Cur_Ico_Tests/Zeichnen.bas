Attribute VB_Name = "Zeichnen"
Option Explicit

Public SWMaske() As Byte
Public FarbeMaske() As Byte
Public TempSW() As Byte
Public TempFarbe() As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Sub Zeichne1Bit(Index As Long)
Dim aByte As Byte
Dim xByte As Byte
Dim KXNummer As Long
Dim Reihe As Long
Dim XB As Boolean
Dim AB As Boolean
Dim Farbe As Integer
Dim AFarbe As String
Dim XFarbe As String
Dim bisher As Long
Dim Höhe As Long

Höhe = Dateiinhalt(Index).HöhePixel
AB = WandleBytes1(BitmapdatasA(Index).Arrays, SWMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel)
XB = WandleBytes1(BitmapdatasX(Index).Arrays, FarbeMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel)
For Reihe = 1 To Dateiinhalt(Index).HöhePixel
For KXNummer = 1 To Dateiinhalt(Index).BreitePixel
bisher = (Reihe - 1) * Dateiinhalt(Index).BreitePixel

aByte = SWMaske(KXNummer + bisher - 1)
xByte = FarbeMaske(KXNummer + bisher - 1)
Select Case xByte
Case "0"
Select Case aByte
Case "0"
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, vbBlack
'Farbe = schwarz
Case "1"
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, &H80000001
'Farbe = Transparent
End Select

Case "1"
Select Case aByte
Case "0"
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, vbWhite
'Farbe = weiß
Case "1"
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, &HC0E0FF
'Farbe = reverse
End Select
End Select
Next KXNummer
Next Reihe

'In Arbeitsfelder laden!!
ReDim TempFarbe(UBound(FarbeMaske))
ReDim TempSW(UBound(SWMaske))
CopyMemory TempFarbe(0), FarbeMaske(0), UBound(FarbeMaske) + 1
CopyMemory TempSW(0), SWMaske(0), UBound(SWMaske) + 1

End Sub

Public Sub Zeichne4Bit(Index As Long)
Dim aByte As Byte
Dim xByte As Byte
Dim KXNummer As Long
Dim Reihe As Long
Dim XB As Boolean
Dim AB As Boolean
Dim Farbe As Integer
Dim AFarbe As String
Dim XFarbe As String
Dim bisher As Long
Dim Pal(15) As Long
Dim black As Integer
Dim white As Integer
Dim Test As Long
Dim i As Long
Dim Höhe As Long

Höhe = Dateiinhalt(Index).HöhePixel
For i = 0 To 15
Pal(i) = RGB(Paletten(Index).Palett(i).R, Paletten(Index).Palett(i).G, Paletten(Index).Palett(i).b)
Next i

If Pal(0) = vbBlack Then
black = 0
Else
For i = 0 To 15
Test = Pal(i)
If Test = vbBlack Then black = i
Next i
End If
If Pal(15) = vbWhite Then
white = 15
Else
For i = 0 To 15
Test = Pal(i)
If Test = vbWhite Then white = i
Next i

End If

AB = WandleBytes1(BitmapdatasA(Index).Arrays, SWMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel) 'schwarzweiß
XB = WandleBytes4(BitmapdatasX(Index).Arrays, FarbeMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel)
For Reihe = 1 To Dateiinhalt(Index).HöhePixel
For KXNummer = 1 To Dateiinhalt(Index).BreitePixel
bisher = (Reihe - 1) * Dateiinhalt(Index).BreitePixel
aByte = SWMaske(KXNummer + bisher - 1)
xByte = FarbeMaske(KXNummer + bisher - 1)
Select Case xByte
Case black
Select Case aByte
Case 0
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, Pal(black)
'Farbe = schwarz
Case 1
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, &H80000001
'Farbe = Transparent
End Select

Case white ' weiß ??
Select Case aByte
Case 0
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, Pal(white)

'Farbe = weiß
Case 1
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, vbYellow

'Farbe = reverse
End Select
Case Else
Select Case aByte
Case 0
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, Pal(xByte)
Case 1
End Select


End Select

Next KXNummer
Next Reihe

'In Arbeitsfelder laden!!
ReDim TempFarbe(UBound(FarbeMaske))
ReDim TempSW(UBound(SWMaske))
CopyMemory TempFarbe(0), FarbeMaske(0), UBound(FarbeMaske) + 1
CopyMemory TempSW(0), SWMaske(0), UBound(SWMaske) + 1

End Sub

Public Sub Zeichne8Bit(Index As Long)
Dim aByte As Byte
Dim xByte As Byte
Dim KXNummer As Long
Dim Reihe As Long
Dim XB As Boolean
Dim AB As Boolean
Dim Farbe As Integer
Dim AFarbe As String
Dim XFarbe As String
Dim bisher As Long
Dim Pal(255) As Long
Dim black As Integer
Dim white As Integer
Dim Test As Long
Dim i As Long
Dim Höhe As Long

Höhe = Dateiinhalt(Index).HöhePixel
For i = 0 To 255
Pal(i) = RGB(Paletten(Index).Palett(i).R, Paletten(Index).Palett(i).G, Paletten(Index).Palett(i).b)
Next i

If Pal(0) = vbBlack Then
black = 0
Else
For i = 0 To 255
Test = Pal(i)
If Test = vbBlack Then black = i
Next i
End If
If Pal(255) = vbWhite Then
white = 255
Else
For i = 0 To 255
Test = Pal(i)
If Test = vbWhite Then white = i
Next i

End If

AB = WandleBytes1(BitmapdatasA(Index).Arrays, SWMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel) 'schwarzweiß
XB = WandleBytes8(BitmapdatasX(Index).Arrays, FarbeMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel)
For Reihe = 1 To Dateiinhalt(Index).HöhePixel
For KXNummer = 1 To Dateiinhalt(Index).BreitePixel
bisher = (Reihe - 1) * Dateiinhalt(Index).BreitePixel
aByte = SWMaske(KXNummer + bisher - 1)
xByte = FarbeMaske(KXNummer + bisher - 1)
Select Case xByte
Case black
Select Case aByte
Case 0
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, Pal(black)
'Farbe = schwarz
Case 1
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, &H80000001
'Farbe = Transparent
End Select

Case white ' weiß ??
Select Case aByte
Case 0
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, Pal(white)

'Farbe = weiß
Case 1
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, vbYellow

'Farbe = reverse
End Select
Case Else
Select Case aByte
Case 0
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, Pal(xByte)
Case 1
End Select

End Select

Next KXNummer
Next Reihe

'In Arbeitsfelder laden!!
ReDim TempFarbe(UBound(FarbeMaske))
ReDim TempSW(UBound(SWMaske))
CopyMemory TempFarbe(0), FarbeMaske(0), UBound(FarbeMaske) + 1
CopyMemory TempSW(0), SWMaske(0), UBound(SWMaske) + 1

End Sub
Public Sub Zeichne24Bit(Index As Long)
Dim aByte As Byte
Dim xByte As Long
Dim KXNummer As Long
Dim Reihe As Long
Dim XB As Boolean
Dim AB As Boolean
Dim Farbe As Integer
Dim AFarbe As String
Dim XFarbe As String
Dim bisher As Long
Dim Pal(255) As Long
Dim black As Integer
Dim white As Integer
Dim Test As Long
Dim i As Long
Dim Höhe As Long
Dim Dreier As Long
Dim wobinich As Long

Höhe = Dateiinhalt(Index).HöhePixel


AB = WandleBytes1(BitmapdatasA(Index).Arrays, SWMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel) 'schwarzweiß
XB = WandleBytes24(BitmapdatasX(Index).Arrays, FarbeMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel)
For Reihe = 1 To Dateiinhalt(Index).HöhePixel
For KXNummer = 1 To Dateiinhalt(Index).BreitePixel
bisher = (Reihe - 1) * Dateiinhalt(Index).BreitePixel
aByte = SWMaske(KXNummer + bisher - 1)
'xByte = FarbeMaske(KXNummer + bisher - 1)
wobinich = (KXNummer + bisher - 1) * 3
xByte = RGB(FarbeMaske(wobinich + 2), FarbeMaske(wobinich + 1), FarbeMaske(wobinich)) 'Anordnung  BGR!!!

Select Case xByte
Case 0 'FarbeMaske = schwarz
Select Case aByte
Case 0
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, 0
'Farbe = schwarz
Case 1
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, &H80000001
'Farbe = Transparent
End Select

Case vbWhite '  Farbe = weiß ??
Select Case aByte
Case 0
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, vbWhite

'Farbe = weiß
Case 1
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, vbYellow

'Farbe = reverse
End Select
Case Else
Select Case aByte
Case 0
ZeichneKasten Form1.Picture1, KXNummer, Höhe - Reihe + 1, xByte
Case 1
End Select

End Select
Dreier = Dreier + 1
Next KXNummer
Next Reihe
End Sub
