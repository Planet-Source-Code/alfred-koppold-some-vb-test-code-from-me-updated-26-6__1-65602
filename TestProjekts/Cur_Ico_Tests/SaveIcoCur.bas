Attribute VB_Name = "SaveIcoCur"
Option Explicit

Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long) As Long

Public Function SaveIcon()

End Function

Public Function SaveCursor(Weg As String) As Boolean
Dim ca As CursorDir
Dim cb As CursorDirEntry
Dim bmi As BITMAPINFOHEADER
Dim XorM() As Byte 'Farbe
Dim AndM() As Byte 'Schwarzweiß

Dim GrImage As Long
Dim GrGesamt As Long
Select Case Dateiinhalt(Nummer).Farbenanzahl
Case 2
BitsToByte TempFarbe, XorM, 1
BitsToByte TempSW, AndM, 1
Case 16
BitsToByte TempFarbe, XorM, 4
BitsToByte TempSW, AndM, 1
Case 256
BitsToByte TempFarbe, XorM, 8
BitsToByte TempSW, AndM, 1
Case 16777216
'BitsToByte TempFarbe, XorM, 24 noch nicht im temp-Speicher
BitsToByte FarbeMaske, XorM, 24
BitsToByte SWMaske, AndM, 1

End Select

Select Case Dateiinhalt(Nummer).Farbenanzahl
Case 16777216
GrGesamt = 40 + (UBound(XorM) + 1) + (UBound(AndM) + 1)  'keine Palette
GrImage = (UBound(XorM) + 1)
Case Else
GrGesamt = 40 + (Dateiinhalt(Nummer).Farbenanzahl * 4) + (UBound(XorM) + 1) + (UBound(AndM) + 1)
GrImage = (UBound(XorM) + 1)
End Select

ca.cdCount = 1
ca.cdReserved = 0
ca.cdType = 2 'cursor
If Dateiinhalt(Nummer).Farbenanzahl < 255 Then
cb.bColorCount = Dateiinhalt(Nummer).Farbenanzahl
Else
cb.bColorCount = 0 'Nur 1 Bit steht zur Verfügung (nicht unbedingt notwendig)
End If
cb.bHeight = Dateiinhalt(Nummer).HöhePixel
cb.bReserved = 0
cb.bWidth = Dateiinhalt(Nummer).BreitePixel

cb.dwBytesinRes = GrGesamt 'GrGesamt 'Komplettgröße
cb.dwImageOffset = 22 ' Cursordir + Cursordirentry (+ bei mehreren vorige Cursor)
cb.wXHotspot = Dateiinhalt(Nummer).CursorXHotspot
cb.wYHotspot = Dateiinhalt(Nummer).CursoryHotspot
Select Case Dateiinhalt(Nummer).Farbenanzahl
Case 2
bmi.biBitCount = 1
Case 8
bmi.biBitCount = 3
Case 16
bmi.biBitCount = 4
Case 256
bmi.biBitCount = 8
Case 16777216
bmi.biBitCount = 24
End Select

bmi.biClrImportant = 0
bmi.biClrUsed = 0
bmi.biCompression = 0
bmi.biHeight = Dateiinhalt(Nummer).HöhePixel * 2
bmi.biPlanes = 1
bmi.biSize = 40
bmi.biSizeImage = GrImage 'GrImage 'Gr + 1 'Achtung
bmi.biWidth = Dateiinhalt(Nummer).BreitePixel
bmi.biXPelsPerMeter = 0
bmi.biYPelsPerMeter = 0

Open Weg For Binary Access Write As #1
Put #1, , ca '0
Put #1, , cb '7
Put #1, , bmi '23

If Dateiinhalt(Nummer).Farbenanzahl < 260 Then
Put #1, , Paletten(Nummer).Palett '63 'bis 256 Farben Palette
End If

Put #1, , XorM  '71
Put #1, , AndM
Close #1

End Function
