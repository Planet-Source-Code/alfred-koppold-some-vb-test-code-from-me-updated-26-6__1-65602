Attribute VB_Name = "TempArray"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Public Sub ChangeArray(XNumber As Long, yNumber As Long, Color As Long, Optional Transparent As Boolean = False, Optional Reverse As Boolean = False)
Dim Höhe As Long
Dim Breite As Long
Dim ArrayNumber As Long

Höhe = Dateiinhalt(Nummer).HöhePixel
Breite = Dateiinhalt(Nummer).BreitePixel
'TempFarbe()
'TempSW()
ArrayNumber = ((Höhe - yNumber) * Breite) + XNumber - 1
TempFarbe(ArrayNumber) = Color '= Palettennummer
TempSW(ArrayNumber) = 0
Select Case Transparent
Case True
TempFarbe(ArrayNumber) = 0
TempSW(ArrayNumber) = 1
Case False
'TempSW(ArrayNumber) = 0
End Select
Select Case Reverse
Case True
TempFarbe(ArrayNumber) = 1
TempSW(ArrayNumber) = 1
Case False
'TempSW(ArrayNumber) = 0
End Select
End Sub
Public Sub BitsToByte(Grundarray() As Byte, FertigesArray() As Byte, BPP As Long)

'ReDim FertigesArray((UBound(Grundarray) + 1) / BPP)
Select Case BPP
Case 1
Mache1BitBMP Grundarray, FertigesArray
Case 3
Case 4
Mache4BitBMP Grundarray, FertigesArray
Case 8
Mache8BitBMP Grundarray, FertigesArray
Case 24
Mache24BitBMP Grundarray, FertigesArray

End Select
End Sub
Private Sub Mache1BitBMP(Grund() As Byte, Fertig() As Byte)
Dim Bytezahl As Long
Dim Byteinhalt As Long
Dim i As Long
Dim BitmBreite As Long
Dim Überg() As Byte
Dim br As Long
Dim hoehe As Long
hoehe = Dateiinhalt(Nummer).HöhePixel
br = Dateiinhalt(Nummer).BreitePixel
Bytezahl = 0

'Herrichten für BMP
BitmBreite = BerechneBMPBytes(br, 1)
If BitmBreite * 8 <> br Then
ReDim Überg(((BitmBreite * 8) * hoehe) - 1)
For i = 0 To Dateiinhalt(Nummer).HöhePixel - 1
CopyMemory Überg(i * BitmBreite * 8), Grund(i * br), br
Next i
ReDim Grund(UBound(Überg))
CopyMemory Grund(0), Überg(0), UBound(Grund) + 1
End If

ReDim Fertig(((UBound(Grund) + 1) \ 8) - 1)

For i = 0 To UBound(Grund) Step 8
If Grund(i) = 1 Then Byteinhalt = Byteinhalt + 128
If Grund(i + 1) = 1 Then Byteinhalt = Byteinhalt + 64
If Grund(i + 2) = 1 Then Byteinhalt = Byteinhalt + 32
If Grund(i + 3) = 1 Then Byteinhalt = Byteinhalt + 16
If Grund(i + 4) = 1 Then Byteinhalt = Byteinhalt + 8
If Grund(i + 5) = 1 Then Byteinhalt = Byteinhalt + 4
If Grund(i + 6) = 1 Then Byteinhalt = Byteinhalt + 2
If Grund(i + 7) = 1 Then Byteinhalt = Byteinhalt + 1
Fertig(Bytezahl) = Byteinhalt
Byteinhalt = 0
Bytezahl = Bytezahl + 1
Next i
End Sub

Private Sub Mache4BitBMP(Grund() As Byte, Fertig() As Byte)
Dim Bytezahl As Long
Dim Byteinhalt As Long
Dim i As Long
Dim BitmBreite As Long
Dim Überg() As Byte
Dim br As Long
Dim hoehe As Long
hoehe = Dateiinhalt(Nummer).HöhePixel
br = Dateiinhalt(Nummer).BreitePixel
Bytezahl = 0

'Herrichten für BMP
BitmBreite = BerechneBMPBytes(br, 4)
If BitmBreite * 2 <> br Then
ReDim Überg(((BitmBreite * 2) * hoehe) - 1)
For i = 0 To Dateiinhalt(Nummer).HöhePixel - 1
CopyMemory Überg(i * BitmBreite * 2), Grund(i * br), br
Next i
ReDim Grund(UBound(Überg))
CopyMemory Grund(0), Überg(0), UBound(Grund) + 1
End If

ReDim Fertig(((UBound(Grund) + 1) \ 2) - 1)

For i = 0 To UBound(Grund) Step 8
Byteinhalt = Byteinhalt + Grund(i) * 16 'Achtung Byte ist gedreht!!!
Byteinhalt = Byteinhalt + Grund(i + 1)
Fertig(Bytezahl) = Byteinhalt
Byteinhalt = 0
Bytezahl = Bytezahl + 1
Byteinhalt = Byteinhalt + Grund(i + 2) * 16
Byteinhalt = Byteinhalt + Grund(i + 3)
Fertig(Bytezahl) = Byteinhalt
Byteinhalt = 0
Bytezahl = Bytezahl + 1
Byteinhalt = Byteinhalt + Grund(i + 4) * 16
Byteinhalt = Byteinhalt + Grund(i + 5)
Fertig(Bytezahl) = Byteinhalt
Byteinhalt = 0
Bytezahl = Bytezahl + 1
Byteinhalt = Byteinhalt + Grund(i + 6) * 16
Byteinhalt = Byteinhalt + Grund(i + 7)
Fertig(Bytezahl) = Byteinhalt
Byteinhalt = 0
Bytezahl = Bytezahl + 1
Next i
End Sub
Private Sub Mache8BitBMP(Grund() As Byte, Fertig() As Byte)
Dim Bytezahl As Long
Dim Byteinhalt As Long
Dim i As Long
Dim BitmBreite As Long
Dim Überg() As Byte
Dim br As Long
Dim hoehe As Long
hoehe = Dateiinhalt(Nummer).HöhePixel
br = Dateiinhalt(Nummer).BreitePixel
Bytezahl = 0

'Herrichten für BMP
BitmBreite = BerechneBMPBytes(br, 8)
If BitmBreite <> br Then
ReDim Überg(((BitmBreite) * hoehe) - 1)
For i = 0 To Dateiinhalt(Nummer).HöhePixel - 1
CopyMemory Überg(i * BitmBreite), Grund(i * br), br
Next i
ReDim Fertig(UBound(Überg))
CopyMemory Fertig(0), Überg(0), UBound(Überg) + 1
Else
ReDim Fertig(UBound(Grund))
CopyMemory Fertig(0), Grund(0), UBound(Grund) + 1

End If

End Sub

Private Sub Mache24BitBMP(Grund() As Byte, Fertig() As Byte)
Dim Bytezahl As Long
Dim Byteinhalt As Long
Dim i As Long
Dim BitmBreite As Long
Dim Überg() As Byte
Dim br As Long
Dim hoehe As Long
hoehe = Dateiinhalt(Nummer).HöhePixel
br = Dateiinhalt(Nummer).BreitePixel
Bytezahl = 0

'Herrichten für BMP
BitmBreite = BerechneBMPBytes(br, 24)
If BitmBreite \ 3 <> br Then
ReDim Überg(((BitmBreite * 3) * hoehe) - 1)
For i = 0 To Dateiinhalt(Nummer).HöhePixel - 1
CopyMemory Überg(i * BitmBreite), Grund(i * br * 3), br * 3
Next i
ReDim Fertig(UBound(Überg))
CopyMemory Fertig(0), Überg(0), UBound(Überg) + 1
Else
ReDim Fertig(UBound(Grund))
CopyMemory Fertig(0), Grund(0), UBound(Grund) + 1

End If

End Sub

