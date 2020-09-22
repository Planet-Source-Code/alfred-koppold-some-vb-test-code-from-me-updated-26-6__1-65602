Attribute VB_Name = "Grids"
Option Explicit
Dim Linewidthbreite As Long
Dim Linewidthhoehe As Long
Dim Kastenfarbe As Integer
Dim Grundbild() As Byte
Dim Grundstring As String
Dim Maske1 As String
Dim Maske2 As String
Dim Kästchenhöhe As Long
Dim Kästchenbreite As Long
Public Type Kästchennummer
X As Long
Y As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Public Sub ZeichneKasten(PicObj As Object, x1 As Long, y1 As Long, Farbe As Long)
Dim xw As Long
Dim yw As Long
Dim x2 As Long
Dim y2 As Long

xw = x1 - 1
yw = y1 - 1
xw = xw * Kästchenbreite
yw = yw * Kästchenhöhe

xw = xw + Linewidthbreite
yw = yw + Linewidthhoehe
y2 = yw + Kästchenhöhe - Linewidthhoehe
x2 = xw + Kästchenbreite - Linewidthbreite
PicObj.Line (xw, yw)-(x2 - Linewidthbreite, y2 - Linewidthhoehe), Farbe, BF

End Sub


Public Sub Grundzeichnen(PicObj As Object, KinWidth As Long, KinHeight As Long, KB As Long, KH As Long)
Dim breit As Long
Dim Hoch As Long
Kästchenhöhe = KH
Kästchenbreite = KB
PicObj.Cls
PicObj.BackColor = &H80000001
Linewidthhoehe = PicObj.DrawWidth '
Linewidthbreite = PicObj.DrawWidth '
PicObj.Width = KinWidth * KB + Linewidthbreite
PicObj.Height = KinHeight * KH + Linewidthhoehe
'Außenkästchen zeichnen
PicObj.Line (0, 0)-(PicObj.Width - Linewidthbreite, PicObj.Height - Linewidthhoehe), &H80000010, B

'Kästchenlinien zeichnen
For Hoch = KH To (KH * KinHeight - KH) Step KH
PicObj.Line (Hoch, 0)-(Hoch, PicObj.Height), &H80000010
Next Hoch
For breit = KB To (KB * KinWidth - KB) Step KB
PicObj.Line (0, breit)-(PicObj.Width, breit), &H80000010
Next breit
End Sub







Public Function GetKästchennummer(X As Single, Y As Single) As Kästchennummer
GetKästchennummer.X = (X \ Kästchenbreite) + 1
GetKästchennummer.Y = (Y \ Kästchenhöhe) + 1
End Function
