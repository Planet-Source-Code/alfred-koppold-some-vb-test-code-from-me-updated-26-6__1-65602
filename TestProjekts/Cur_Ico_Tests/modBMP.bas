Attribute VB_Name = "modBMP"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Function BerechneBMPBytes(BMPBreite As Long, BPP As Long) As Long
Dim Breite As Long
Dim BreiteBMP
Dim PixelprByte As Long

PixelprByte = BMPBreite * BPP
If PixelprByte Mod (8) <> 0 Then
PixelprByte = ((PixelprByte \ 8) + 1) * 8 ' erster Fehler falls nicht durch 8 teilbar (keine genauen bytes)
End If

BreiteBMP = PixelprByte \ 8
If BreiteBMP Mod (4) <> 0 Then
BreiteBMP = ((BreiteBMP \ 4) + 1) * 4 ' zweiter Fehler Bytes pro Zeile nicht durch 4 teilbar siehe BMP-Format
End If

If BreiteBMP < 4 Then
BreiteBMP = 4 ' dritter Fehler falls Zeile kleiner als 4 Bytes (Mindestgröße)
End If
BerechneBMPBytes = BreiteBMP
End Function
Public Function Wandle4Bits(ForBits As String) As Byte
Select Case ForBits
Case "0000"
Wandle4Bits = 0
Case "0001"
Wandle4Bits = 1
Case "0010"
Wandle4Bits = 2
Case "0011"
Wandle4Bits = 3
Case "0100"
Wandle4Bits = 4
Case "0101"
Wandle4Bits = 5
Case "0110"
Wandle4Bits = 6
Case "0111"
Wandle4Bits = 7
Case "1000"
Wandle4Bits = 8
Case "1001"
Wandle4Bits = 9
Case "1010"
Wandle4Bits = 10
Case "1011"
Wandle4Bits = 11
Case "1100"
Wandle4Bits = 12
Case "1101"
Wandle4Bits = 13
Case "1110"
Wandle4Bits = 14
Case "1111"
Wandle4Bits = 15
End Select
End Function
Public Function WandleBytes4(Bytefeld() As Byte, Umgewandelt() As Byte, BMPBreite As Long, BMPHöhe As Long) As Boolean
Dim Test As String
Dim i As Long
Dim Größe As Long
Dim Wo As Long
Dim Übergabe() As Byte
Dim BMPBytesBreite As Long

WandleBytes4 = False
Wo = 0
Größe = ((UBound(Bytefeld) - LBound(Bytefeld) + 1) * 2) - 1
ReDim Umgewandelt(Größe) '(BMPBreite * BMPHöhe - 1)
For i = LBound(Bytefeld) To UBound(Bytefeld)
Test = ""
Select Case GetByte(Bytefeld(i), 1)
Case 0
Test = Test & "0"
Case 1
Test = Test & "1"
End Select
Select Case GetByte(Bytefeld(i), 2)
Case 0
Test = Test & "0"
Case 1
Test = Test & "1"
End Select
Select Case GetByte(Bytefeld(i), 3)
Case 0
Test = Test & "0"
Case 1
Test = Test & "1"
End Select
Select Case GetByte(Bytefeld(i), 4)
Case 0
Test = Test & "0"
Case 1
Test = Test & "1"
End Select
Umgewandelt(Wo) = Wandle4Bits(Test)
Wo = Wo + 1
Test = ""


Select Case GetByte(Bytefeld(i), 5)
Case 0
Test = Test & "0"
Case 1
Test = Test & "1"
End Select
Select Case GetByte(Bytefeld(i), 6)
Case 0
Test = Test & "0"
Case 1
Test = Test & "1"
End Select
Select Case GetByte(Bytefeld(i), 7)
Case 0
Test = Test & "0"
Case 1
Test = Test & "1"
End Select
Select Case GetByte(Bytefeld(i), 8)
Case 0
Test = Test & "0"
Case 1
Test = Test & "1"
End Select
Umgewandelt(Wo) = Wandle4Bits(Test)
Wo = Wo + 1
Test = ""
Next i
BMPBytesBreite = BerechneBMPBytes(BMPBreite, 4)
If BMPBytesBreite <> (BMPBreite \ 2) Then
'oder If UBound(Umgewandelt) - LBound(Umgewandelt) <> BMPBreite * BMPHöhe - 1 Then
ReDim Übergabe((BMPBreite * BMPHöhe) - 1)
For i = 0 To BMPHöhe - 1
CopyMemory Übergabe(i * BMPBreite), Umgewandelt(i * (BMPBytesBreite * 2)), BMPBreite
Next i
ReDim Umgewandelt(UBound(Übergabe))
CopyMemory Umgewandelt(0), Übergabe(0), BMPBreite * BMPHöhe
End If

WandleBytes4 = True
End Function

Private Function GetByte(Bytes As Byte, Position As Long) As Integer
GetByte = 0
Select Case Position
Case 1
If Bytes And 128 Then GetByte = 1
Case 2
If Bytes And 64 Then GetByte = 1
Case 3
If Bytes And 32 Then GetByte = 1
Case 4
If Bytes And 16 Then GetByte = 1
Case 5
If Bytes And 8 Then GetByte = 1
Case 6
If Bytes And 4 Then GetByte = 1
Case 7
If Bytes And 2 Then GetByte = 1
Case 8
If Bytes And 1 Then GetByte = 1
End Select

End Function

Public Function WandleBytes1(Bytefeld() As Byte, Umgewandelt() As Byte, BMPBreite As Long, BMPHöhe As Long) As Boolean
Dim Numb As Long
Dim i As Long
Dim Gr As Long
Dim Übergabe() As Byte
Dim BMPBytesBreite As Long

WandleBytes1 = False
Numb = 0
Gr = (((UBound(Bytefeld) - LBound(Bytefeld)) + 1) * 8) - 1
ReDim Umgewandelt(Gr)
For i = LBound(Bytefeld) To UBound(Bytefeld)
Numb = i * 8
Select Case GetByte(Bytefeld(i), 1)
Case 0
Umgewandelt(Numb) = 0
'Test = Test & "0"
Case 1
Umgewandelt(Numb) = 1
'Test = Test & "1"
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
WandleBytes1 = True
End Function



Public Function WandleBytes8(Bytefeld() As Byte, Umgewandelt() As Byte, BMPBreite As Long, BMPHöhe As Long) As Boolean
Dim Test As String
Dim i As Long
Dim Größe As Long
Dim Wo As Long
Dim Übergabe() As Byte
Dim BMPBytesBreite As Long
Dim Testa As Boolean
Testa = True
WandleBytes8 = False
ReDim Umgewandelt(UBound(Bytefeld))
CopyMemory Umgewandelt(0), Bytefeld(0), UBound(Bytefeld) + 1
BMPBytesBreite = BerechneBMPBytes(BMPBreite, 8)
If Testa = True Then
If BMPBytesBreite <> (BMPBreite) Then
'oder If UBound(Umgewandelt) - LBound(Umgewandelt) <> BMPBreite * BMPHöhe - 1 Then
ReDim Übergabe((BMPBreite * BMPHöhe) - 1)
For i = 0 To BMPHöhe - 1
CopyMemory Übergabe(i * BMPBreite), Bytefeld(i * BMPBytesBreite), BMPBreite
Next i
ReDim Umgewandelt(UBound(Übergabe))
CopyMemory Umgewandelt(0), Übergabe(0), BMPBreite * BMPHöhe
End If
End If
WandleBytes8 = True
End Function

Public Function WandleBytes24(Bytefeld() As Byte, Umgewandelt() As Byte, BMPBreite As Long, BMPHöhe As Long) As Boolean
Dim Test As String
Dim i As Long
Dim Größe As Long
Dim Wo As Long
Dim Übergabe() As Byte
Dim BMPBytesBreite As Long
Dim Testa As Boolean
Testa = True
WandleBytes24 = False
ReDim Umgewandelt(UBound(Bytefeld))
CopyMemory Umgewandelt(0), Bytefeld(0), UBound(Bytefeld) + 1
BMPBytesBreite = BerechneBMPBytes(BMPBreite, 24)
If Testa = True Then
If BMPBytesBreite <> (BMPBreite * 3) Then
'oder If UBound(Umgewandelt) - LBound(Umgewandelt) <> BMPBreite * BMPHöhe - 1 Then
ReDim Übergabe((BMPBreite * BMPHöhe) - 1)
For i = 0 To BMPHöhe - 1
CopyMemory Übergabe(i * (BMPBreite * 3)), Bytefeld(i * BMPBytesBreite), BMPBreite * 3
Next i
ReDim Umgewandelt(UBound(Übergabe))
CopyMemory Umgewandelt(0), Übergabe(0), BMPBreite * BMPHöhe
End If
End If
WandleBytes24 = True
End Function

