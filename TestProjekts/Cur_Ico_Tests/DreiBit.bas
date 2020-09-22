Attribute VB_Name = "DreiBit"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Sub Zeichne3Bit(Index As Long)
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
For i = 0 To 7
Pal(i) = RGB(Paletten(Index).Palett(i).R, Paletten(Index).Palett(i).G, Paletten(Index).Palett(i).b)
Next i

If Pal(0) = vbBlack Then
black = 0
Else
For i = 0 To 7
Test = Pal(i)
If Test = vbBlack Then black = i
Next i
End If
If Pal(7) = vbWhite Then
white = 7
Else
For i = 0 To 7
Test = Pal(i)
If Test = vbWhite Then white = i
Next i

End If

AB = WandleBytes1(BitmapdatasA(Index).Arrays, SWMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel) 'schwarzweiß
XB = WandleBytes3(BitmapdatasX(Index).Arrays, FarbeMaske, Dateiinhalt(Index).BreitePixel, Dateiinhalt(Index).HöhePixel)
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
End Sub
Public Function WandleBytes3(Bytefeld() As Byte, Umgewandelt() As Byte, BMPBreite As Long, BMPHöhe As Long) As Boolean
Dim Numb As Long
Dim i As Long
Dim Gr As Long
Dim Übergabe() As Byte
Dim BMPBytesBreite As Long

WandleBytes3 = False
Numb = 0
Gr = (((UBound(Bytefeld) - LBound(Bytefeld)) + 1) * 8) - 1
ReDim Umgewandelt(Gr)
For i = LBound(Bytefeld) To UBound(Bytefeld)
Numb = i * 8
Select Case GetByte(Bytefeld(i), 1)
Case 0
Umgewandelt(Numb) = 0
Case 1
Umgewandelt(Numb) = 1
End Select
Select Case GetByte(Bytefeld(i), 2)
Case 0
Umgewandelt(1 + Numb) = 0
Case 1
Umgewandelt(1 + Numb) = 1
End Select
Select Case GetByte(Bytefeld(i), 3)
Case 0
Umgewandelt(2 + Numb) = 0
Case 1
Umgewandelt(2 + Numb) = 1
End Select
Select Case GetByte(Bytefeld(i), 4)
Case 0
Umgewandelt(3 + Numb) = 0
Case 1
Umgewandelt(3 + Numb) = 1
End Select
Select Case GetByte(Bytefeld(i), 5)
Case 0
Umgewandelt(4 + Numb) = 0
Case 1
Umgewandelt(4 + Numb) = 1
End Select
Select Case GetByte(Bytefeld(i), 6)
Case 0
Umgewandelt(5 + Numb) = 0
Case 1
Umgewandelt(5 + Numb) = 1
End Select
Select Case GetByte(Bytefeld(i), 7)
Case 0
Umgewandelt(6 + Numb) = 0
Case 1
Umgewandelt(6 + Numb) = 1
End Select
Select Case GetByte(Bytefeld(i), 8)
Case 0
Umgewandelt(7 + Numb) = 0
Case 1
Umgewandelt(7 + Numb) = 1
End Select
Next i

Umwandeln3bit Umgewandelt
BMPBytesBreite = BerechneBMPBytes(BMPBreite, 1)
If BMPBytesBreite <> (BMPBreite \ 8) Then
If UBound(Umgewandelt) - LBound(Umgewandelt) <> BMPBreite * BMPHöhe - 1 Then
ReDim Übergabe((BMPBreite * BMPHöhe) - 1)
For i = 0 To BMPHöhe - 1
CopyMemory Übergabe(i * BMPBreite), Umgewandelt(i * (BMPBytesBreite * 8)), BMPBreite
Next i
ReDim Umgewandelt(UBound(Übergabe))
CopyMemory Umgewandelt(0), Übergabe(0), BMPBreite * BMPHöhe
End If
End If
WandleBytes3 = True
End Function


Private Sub Umwandeln3bit(Bytefeld() As Byte)
Dim i As Long
Dim Dreib(2) As Byte
Dim Übergabe() As Byte
Dim Wo As Long
Wo = 0
ReDim Übergabe(((UBound(Bytefeld) + 1) \ 3) - 1)
For i = 0 To UBound(Bytefeld) Step 3
CopyMemory Dreib(0), Bytefeld(i), 3
Übergabe(Wo) = Wandle3Bit(Dreib)
Wo = Wo + 1
Next i

ReDim Bytefeld(UBound(Übergabe))
CopyMemory Bytefeld(0), Übergabe(0), UBound(Übergabe) + 1
End Sub
Private Function Wandle3Bit(Dreibit() As Byte) As Byte
Dim Gewandelt As Byte

Gewandelt = 0

If Dreibit(0) = 1 Then Gewandelt = Gewandelt + 4
If Dreibit(1) = 1 Then Gewandelt = Gewandelt + 2
If Dreibit(2) = 1 Then Gewandelt = Gewandelt + 1
Wandle3Bit = Gewandelt

End Function
Private Sub Make3Bit(ByteArrayAlt() As Byte, Rückgabe() As Byte)
Dim i As Long
Dim Bytestring As String
Dim DreiBits As String

For i = 0 To UBound(ByteArrayAlt)
DreiBits = Mache3String(i)
Bytestring = Bytestring & DreiBits
Next i
FillByteArray Bytestring, Rückgabe
End Sub
Private Function Mache3String(EinByte As Byte) As String
Select Case EinByte
Case 0
Mache3String = "000"
Case 1
Mache3String = "001"
Case 2
Mache3String = "010"
Case 3
Mache3String = "011"
Case 4
Mache3String = "100"
Case 5
Mache3String = "101"
Case 6
Mache3String = "110"
Case 7
Mache3String = "111"
End Select

End Function
Private Sub FillByteArray(Bitstring As String, Rück() As Byte)
Dim i As Long
Dim größe As Long
Dim BitstringAcht As String

größe = Len(Bitstring \ 8)
ReDim Rück(größe - 1)
For i = 1 To Len(Bitstring) Step 8 'Im 8erSchritt in Bytes umwandeln
BitstringAcht = Mid(i, Bitstring, 8)
Nummer = WandleinZahl(BitstringAcht)
Next i
End Sub
Private Function WandleinZahl(Nummer As String) As Long
Dim newn As Long
newn = 0
If Mid(Nummer, 1, 1) = 1 Then newn = newn + 1 'oder umgekehrt
If Mid(Nummer, 2, 1) = 1 Then newn = newn + 2
If Mid(Nummer, 3, 1) = 1 Then newn = newn + 4
If Mid(Nummer, 4, 1) = 1 Then newn = newn + 8
If Mid(Nummer, 5, 1) = 1 Then newn = newn + 16
If Mid(Nummer, 6, 1) = 1 Then newn = newn + 32
If Mid(Nummer, 7, 1) = 1 Then newn = newn + 64
If Mid(Nummer, 8, 1) = 1 Then newn = newn + 128
WandleinZahl = newn

End Function

