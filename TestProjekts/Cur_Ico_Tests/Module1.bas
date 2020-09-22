Attribute VB_Name = "IconCursor"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Übergabe() As Byte
Public Type Weg
Colour As String
SW As String
End Type
Public PictureWays() As Weg
Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER
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
Type BitmapArray
Arrays() As Byte
End Type
Type RGB
b As Byte
G As Byte
R As Byte
Reserved As Byte
End Type

Type Palette
Palett() As RGB
End Type
Public Type CursorDir '6 Bytes
cdReserved As Integer ' nicht benutzt 0
cdType As Integer 'für Cursor 2
cdCount As Integer 'Anzahl Cursor
End Type

Public Type CursorDirEntry '16 Byte
bWidth As Byte 'Breite in Byte
bHeight As Byte 'Höhe in Byte
bColorCount As Byte 'Anzahl der benutzten Farben (2 für Schwarzweiß)
bReserved As Byte 'nicht bunutzt 0
wXHotspot As Integer 'MausHotspot X
wYHotspot As Integer 'MausHotspot Y
dwBytesinRes As Long 'Cursorgröße in Bytes
dwImageOffset As Long ' Offset des Cursors vom Dateibeginn
End Type

Public Type IconDir '6Byte
idReserved As Integer
idType As Integer
idCount As Integer
End Type

Public Type IconDirEntry '16 Byte
bWidth As Byte
bHeight As Byte
bColorCount As Byte
bReserved As Byte
wPlanes As Integer
wBitcount As Integer
dwBytesinRes As Long
dwImageOffset As Long
End Type
Public Type FileDir '6Byte
idReserved As Integer
idType As Integer
idCount As Integer
End Type
Public Type Inhalt
Grafikmenge As Long
BreitePixel As Long
HöhePixel As Long
Farbenanzahl As Long
Type As String
CursorXHotspot As Long
CursoryHotspot As Long
End Type
Public Paletten() As Palette
Public Dateiinhalt() As Inhalt
Public BitmapdatasX() As BitmapArray
Public BitmapdatasA() As BitmapArray

Public Function OpenFile(Filename As String)
Dim filenummer As Long
Dim Filedescr As FileDir
filenummer = FreeFile
Open Filename For Binary As filenummer
Get filenummer, , Filedescr
Close filenummer
Select Case Filedescr.idType
Case 1
'Icon
OpenIcon (Filename)
Case 2
'Cursor
OpenCursor (Filename)
End Select
End Function

Private Function OpenCursor(Filename As String)
Dim BMFilehead As BITMAPFILEHEADER
Dim BMpB As Long
Dim Tempname As String
Dim tempnumber As Long
Dim i As Long
Dim Cursorfirst As CursorDir
Dim Cursorsec() As CursorDirEntry
Dim filenummer As Long
Dim Palettesw As Palette
Dim Bitmhead() As BITMAPINFOHEADER
Dim Berechnung As Long
Dim BWGröße As Long
Dim Realhöhe As Long
Dim KleinerCursor As Boolean
Dim BerechnungRight As Long
Dim Breitemus
ReDim Paletten(0)
ReDim Dateiinhalt(0)
ReDim BitmapdatasX(0)
ReDim BitmapdatasA(0)
Dim Farbenanzahl As Long

filenummer = FreeFile
Open Filename For Binary As filenummer
Get filenummer, 1, Cursorfirst
ReDim Dateiinhalt(Cursorfirst.cdCount - 1)
ReDim PictureWays(Cursorfirst.cdCount - 1)
For i = 0 To Cursorfirst.cdCount - 1 ' 0 nach (Menge Cursor - 1)
KleinerCursor = False
ReDim Preserve Cursorsec(i)
Get filenummer, 7 + (i * 16), Cursorsec(i)
'Fehlerkorrekturen
ReDim Bitmhead(i)
Get filenummer, Cursorsec(i).dwImageOffset + 1, Bitmhead(i)
If Cursorsec(i).bHeight <> Bitmhead(i).biHeight \ 2 Then
Cursorsec(i).bHeight = Bitmhead(i).biHeight \ 2
End If
If Cursorsec(i).bWidth <> Bitmhead(i).biWidth Then
Cursorsec(i).bWidth = Bitmhead(i).biWidth
End If
Select Case Bitmhead(i).biBitCount
Case 1
Cursorsec(i).bColorCount = 2
Farbenanzahl = 2
Case 3
Cursorsec(i).bColorCount = 8
Farbenanzahl = 8
Case 4
Cursorsec(i).bColorCount = 16
Farbenanzahl = 16
Case 8
Cursorsec(i).bColorCount = 255 '???
Farbenanzahl = 256
Case 24
Cursorsec(i).bColorCount = 0
Farbenanzahl = 16777216
Case 0
Cursorsec(i).bColorCount = 2
End Select




'Berechnen der Grafikgröße
BMpB = BerechneBMPBytes(Bitmhead(i).biWidth, CLng(Bitmhead(i).biBitCount))
Select Case Bitmhead(i).biBitCount
Case Is > 8 ' also 24 bit
BMpB = BMpB \ (Bitmhead(i).biBitCount \ 8)
Case Else
BMpB = BMpB * (8 \ Bitmhead(i).biBitCount) 'Berechnung der nötigen Breite
End Select
Berechnung = (Bitmhead(i).biHeight \ 2) * BMpB * Bitmhead(i).biBitCount \ 8 'BMPb = Breite einer Zeile - normalerweise gleich Bitmaphead(i).biwidth
BMpB = BerechneBMPBytes(Bitmhead(i).biWidth, 1) 'Berechnen der Bytes pro BMPZeile
BMpB = BMpB * 8 ' * 8 da 1 Bit = 1 Pixel bei BW
BWGröße = BMpB * Bitmhead(i).biHeight \ 8 \ 2
Realhöhe = Bitmhead(i).biHeight \ 2

If Bitmhead(i).biBitCount < 24 Then
'2 Bilder da 1.Bild farbig und 2.Bild schwarzweiß



ReDim Preserve Paletten(i)
ReDim Paletten(i).Palett(Farbenanzahl - 1)
Get filenummer, Cursorsec(i).dwImageOffset + 1 + 40, Paletten(i).Palett


ReDim Preserve BitmapdatasX(i)
ReDim BitmapdatasX(i).Arrays(Berechnung - 1)
Get filenummer, , BitmapdatasX(i).Arrays
ReDim Preserve BitmapdatasA(i)
ReDim BitmapdatasA(i).Arrays(BWGröße - 1)
Get filenummer, , BitmapdatasA(i).Arrays
Dateiinhalt(i).BreitePixel = Cursorsec(i).bWidth
Dateiinhalt(i).HöhePixel = Realhöhe
Dateiinhalt(i).Farbenanzahl = Farbenanzahl
Dateiinhalt(i).Type = "Cursor"
Dateiinhalt(i).CursorXHotspot = Cursorsec(i).wXHotspot
Dateiinhalt(i).CursoryHotspot = Cursorsec(i).wYHotspot
Dateiinhalt(i).Grafikmenge = Cursorfirst.cdCount
End If
If Bitmhead(i).biBitCount = 24 Then
ReDim Preserve BitmapdatasX(i)
ReDim BitmapdatasX(i).Arrays(Berechnung - 1)
Get filenummer, , BitmapdatasX(i).Arrays
ReDim Preserve BitmapdatasA(i)
ReDim BitmapdatasA(i).Arrays(BWGröße - 1)
Get filenummer, , BitmapdatasA(i).Arrays
Dateiinhalt(i).Type = "Cursor"
Dateiinhalt(i).Farbenanzahl = Farbenanzahl
Dateiinhalt(i).HöhePixel = Realhöhe
Dateiinhalt(i).BreitePixel = Cursorsec(i).bWidth
Dateiinhalt(i).Grafikmenge = Cursorfirst.cdCount

End If

Next i
Close filenummer
'Form1.Text1.Text = Form1.Text1.Text & Cursorfirst.cdCount & " Cursor(s) in der Datei"
End Function


Private Sub OpenIcon(Filename As String)
Dim BMFilehead As BITMAPFILEHEADER
Dim Tempname As String
Dim tempnumber As Long
Dim i As Long
Dim IconFirst As IconDir
Dim Iconsec() As CursorDirEntry
Dim filenummer As Long
Dim Bitmhead() As BITMAPINFOHEADER
Dim Berechnung As Long
Dim BWGröße As Long
Dim Realhöhe As Long
Dim KleinesIcon As Boolean
Dim BerechnungRight As Long
Dim Palettesw As Palette
Dim BMpB As Long
Dim Farbenanzahl As Long

ReDim Paletten(0)
ReDim Dateiinhalt(0)
ReDim BitmapdatasX(0)
ReDim BitmapdatasA(0)

filenummer = FreeFile
Open Filename For Binary As filenummer
Get filenummer, 1, IconFirst
ReDim Dateiinhalt(IconFirst.idCount - 1)
ReDim PictureWays(IconFirst.idCount - 1)
For i = 0 To IconFirst.idCount - 1 ' 0 nach (Menge Cursor - 1)
KleinesIcon = False
ReDim Preserve Iconsec(i)
Get filenummer, 7 + (i * 16), Iconsec(i)
'Fehlerkorrekturen
ReDim Bitmhead(i)
Get filenummer, Iconsec(i).dwImageOffset + 1, Bitmhead(i)
If Iconsec(i).bHeight <> Bitmhead(i).biHeight \ 2 Then
Iconsec(i).bHeight = Bitmhead(i).biHeight \ 2
End If
If Iconsec(i).bWidth <> Bitmhead(i).biWidth Then
Iconsec(i).bWidth = Bitmhead(i).biWidth
End If
Select Case Bitmhead(i).biBitCount
Case 1
Iconsec(i).bColorCount = 2
Farbenanzahl = 2
Case 3
Iconsec(i).bColorCount = 8
Farbenanzahl = 8
Case 4
Iconsec(i).bColorCount = 16
Farbenanzahl = 16
Case 8
Iconsec(i).bColorCount = 0 ' 256 Farben
Farbenanzahl = 256
Case 24
Iconsec(i).bColorCount = 0 '16777216 Farben
Farbenanzahl = 16777216
Case 0
Iconsec(i).bColorCount = 2
Farbenanzahl = 2
End Select

'Berechnen der Grafikgröße
BMpB = BerechneBMPBytes(Bitmhead(i).biWidth, CLng(Bitmhead(i).biBitCount))
Select Case Bitmhead(i).biBitCount
Case Is > 8 ' also 24 bit
BMpB = BMpB \ (Bitmhead(i).biBitCount \ 8)
Case Else
BMpB = BMpB * (8 \ Bitmhead(i).biBitCount) 'Berechnung der nötigen Breite
End Select
Berechnung = (Bitmhead(i).biHeight \ 2) * BMpB * Bitmhead(i).biBitCount \ 8 'BMPb = Breite einer Zeile - normalerweise gleich Bitmaphead(i).biwidth
BMpB = BerechneBMPBytes(Bitmhead(i).biWidth, 1) 'Berechnen der Bytes pro BMPZeile
BMpB = BMpB * 8 ' * 8 da 1 Bit = 1 Pixel bei BW
BWGröße = BMpB * Bitmhead(i).biHeight \ 8 \ 2
Realhöhe = Bitmhead(i).biHeight \ 2

If Bitmhead(i).biBitCount < 24 Then
ReDim Preserve Paletten(i)
ReDim Paletten(i).Palett(Farbenanzahl - 1)
Get filenummer, Iconsec(i).dwImageOffset + 1 + 40, Paletten(i).Palett



'Breitemus = Bitmhead(i).biWidth \ (8 \ Bitmhead(i).biBitCount) \ 4
'If Bitmhead(i).biWidth < 32 Then
'BWGröße = 32 * (Bitmhead(i).biHeight \ 2) \ 8
'Berechnung = Bitmhead(i).biSizeImage - BWGröße
'End If
'2 Bilder da 1.Bild farbig und 2.Bild schwarzweiß
ReDim Preserve BitmapdatasX(i)
ReDim BitmapdatasX(i).Arrays(Berechnung - 1)
Get filenummer, , BitmapdatasX(i).Arrays
ReDim Preserve BitmapdatasA(i)
ReDim BitmapdatasA(i).Arrays(BWGröße - 1)
Get filenummer, , BitmapdatasA(i).Arrays
Dateiinhalt(i).Type = "Icon"
Dateiinhalt(i).Farbenanzahl = Farbenanzahl
Dateiinhalt(i).HöhePixel = Realhöhe
Dateiinhalt(i).BreitePixel = Iconsec(i).bWidth
Dateiinhalt(i).Grafikmenge = IconFirst.idCount
End If

If Bitmhead(i).biBitCount = 24 Then
ReDim Preserve BitmapdatasX(i)
ReDim BitmapdatasX(i).Arrays(Berechnung - 1)
Get filenummer, , BitmapdatasX(i).Arrays
ReDim Preserve BitmapdatasA(i)
ReDim BitmapdatasA(i).Arrays(BWGröße - 1)
Get filenummer, , BitmapdatasA(i).Arrays
Dateiinhalt(i).Type = "Icon"
Dateiinhalt(i).Farbenanzahl = Farbenanzahl
Dateiinhalt(i).HöhePixel = Realhöhe
Dateiinhalt(i).BreitePixel = Iconsec(i).bWidth
Dateiinhalt(i).Grafikmenge = IconFirst.idCount

End If

Next i
Close filenummer
'Form1.Text1.Text = Form1.Text1.Text & IconFirst.idCount & " Icon(s) in der Datei"

End Sub
