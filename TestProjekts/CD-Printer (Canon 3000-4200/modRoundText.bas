Attribute VB_Name = "modRoundText"
'------------------------------------------------------------
' Project Name: Project1
' Module Name: modRoundText
' Date: 07/05/2001
' Time: 12.29
' Revision:
' Author: NDV Software
'------------------------------------------------------------
' ****************************************************************************************************
' Copyright © 1990 - 2001 NDV Software,
' All rights are reserved, ndv@interfree.it
' ****************************************************************************************************
Option Explicit
Global Const PIGRECO = 3.141592654

Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFacename As String * 33
End Type
Public Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

'------------------------------------------------------------
' Name: drawCircularText
' Desc: Draw a circle/arc text
' Type: Public
' Parameters:
'    Obj As Object              Destination Object of the printing (Picture o Printer object)
'    Testo As String            Text to print
'    TextStartAngle As Single   Starting Angle of the text
'    Raggio As Single           Radius of the circle/arc on that is printed the text
'    CX As Integer              Center X coord of the circle
'    CY As Integer              Center Y coord of the circle
'    TextSector As Single       Sector of circle to fill with text. Good value are between 0 and 360.
'                               0 -> All the text is printed on the same point
'                               180 -> the text make a semicircular beginning from the starting angle
'                               360 -> the text make a circle beginning from the starting angle
'    All the other graphics parameters like font,color,font propertyes (bold,italic,...) can be
'    setted on the Obj Object before to pass it to the drawCircularText procedure.
'
' Date: lunedì 7 maggio 2001
' Time: 12.25
' Author: NDV Software
' Revision:
'------------------------------------------------------------
Public Sub drawCircularText(hdc As Long, Testo As String, TextStartAngle As Single, Raggio As Single, CX As Integer, CY As Integer, TextSector As Single, Fontname As String, Fontsize As Long, XBegin As Long, YBegin As Long)
  On Error GoTo Errore
  Dim F As LOGFONT
  Dim hPrevFont As Long
  Dim hFont As Long
  Dim I As Integer
  Dim Passo As Single
  Dim x As Long
  Dim y As Long
  Dim oldmode As Long
  oldmode = GetBkMode(hdc)
  SetBkMode hdc, 1
  Passo = TextSector / Len(Testo)   'Angular Step
    
  For I = 1 To Len(Testo)
    F.lfEscapement = 10 * TextStartAngle - (10 * Passo * (I - 1)) 'rotation angle, in tenths (x10)
    F.lfFacename = Fontname & Chr(0)
    F.lfHeight = Fontsize
    hFont = CreateFont(Fontsize, 0, 10 * TextStartAngle - (10 * Passo * (I - 1)), 0, 100, 0, 0, 0, 136, 0, 0, 2, 0, Fontname)
    'hFont = CreateFontIndirect(F)
    hPrevFont = SelectObject(hdc, hFont)
    x = CX + Raggio * Sin((-180 + TextStartAngle - (I - 1) * Passo) * PIGRECO / 180)
    y = CY + Raggio * Cos((-180 + TextStartAngle - (I - 1) * Passo) * PIGRECO / 180)
    
    x = x \ Screen.TwipsPerPixelX
    y = y \ Screen.TwipsPerPixelY
    x = x + XBegin
    y = y + YBegin
TextOut hdc, x, y, Mid(Testo, I, 1), 1
    hFont = SelectObject(hdc, hPrevFont)
    DeleteObject hFont
  Next I
  SetBkMode hdc, oldmode
  Exit Sub
Errore:
  SetBkMode hdc, oldmode
  Exit Sub
End Sub
