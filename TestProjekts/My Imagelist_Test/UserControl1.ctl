VERSION 5.00
Begin VB.UserControl ImageList 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00C0C0C0&
   MouseIcon       =   "UserControl1.ctx":0000
   PropertyPages   =   "UserControl1.ctx":030A
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "UserControl1.ctx":0337
End
Attribute VB_Name = "ImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Pic As StdPicture
'Standard-Eigenschaftswerte:
Const m_def_hImageList = 0
Const m_def_ImageHeight = 0
Const m_def_ImageWidth = 0
Const m_def_UseMaskColor = True
Const m_def_Backcolor = &H80000005
Const m_def_Maskcolor = &HC0C0C0
'Eigenschaftsvariablen:
Dim m_hImageList As OLE_HANDLE
Dim m_ImageHeight As Integer
Dim m_ImageWidth As Integer
Dim m_UseMaskColor As Boolean
Dim m_Backcolor As Long
Dim m_MaskColor As Long
Public ListImages As ListImages

Private Sub UserControl_Initialize()
Set Pic = UserControl.MouseIcon
Anzahl = 0
End Sub

Private Sub UserControl_Paint()
Dim tR As RECT
   tR.Right = 38
   tR.Bottom = 38
   DrawEdge UserControl.hdc, tR, 5, 15
UserControl.PaintPicture Pic, 3, 3
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 38 * Screen.TwipsPerPixelX
UserControl.Width = 38 * Screen.TwipsPerPixelY

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Legt die Hintergrundfarbe zum Anzeigen von Text und Grafiken in einem Objekt fest oder gibt diese zurück."
Attribute BackColor.VB_HelpID = 12
    BackColor = m_Backcolor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_Backcolor = New_BackColor
    PropertyChanged "BackColor"
    BackColorIntern = m_Backcolor
End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Gibt einen Wert zurück oder legt einen Wert fest, der bestimmt, ob die in grafischen Abbildungsliste-Operationen verwendete Farbe transparent ist."
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
    MaskColorIntern = m_MaskColor
End Property

Public Property Get hImageList() As OLE_HANDLE
Attribute hImageList.VB_MemberFlags = "400"
    hImageList = m_hImageList
End Property

Public Property Let hImageList(ByVal New_hImageList As OLE_HANDLE)
    m_hImageList = New_hImageList
    PropertyChanged "hImageList"
End Property

Public Property Get ImageHeight() As Integer
Attribute ImageHeight.VB_Description = "Legt die Höhe eine ListImage-Objekts fest oder gibt sie zurück."
    ImageHeight = m_ImageHeight
End Property

Public Property Let ImageHeight(ByVal New_ImageHeight As Integer)
Dim ahwnd As Long
    If Anzahl > 0 And New_ImageHeight <> m_ImageHeight Then
    MsgBox EMsg
    SendKeys " {BS}~"
    Else
    m_ImageHeight = New_ImageHeight
    PropertyChanged "ImageHeight"
    End If
End Property

Public Property Get ImageWidth() As Integer
Attribute ImageWidth.VB_Description = "Legt die Breite eine ListImage-Objekts in einem Abbildungsliste-Steuerelement fest oder gibt sie zurück."
    ImageWidth = m_ImageWidth
End Property

Public Property Let ImageWidth(ByVal New_ImageWidth As Integer)
    If Anzahl > 0 And New_ImageWidth <> m_ImageWidth Then
    MsgBox EMsg
    SendKeys " {BS}~"
    Else
    m_ImageWidth = New_ImageWidth
    PropertyChanged "ImageWidth"
    End If
End Property

Public Function Overlay(Key1, Key2) As Variant

End Function

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Gibt einen Wert zurück oder legt einen Wert fest, der bestimmt, ob das Abbildungsliste-Steuerelement die MaskColor-Eigenschaft verwendet."
    UseMaskColor = m_UseMaskColor
End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)
If m_UseMaskColor <> New_UseMaskColor Then
    m_UseMaskColor = New_UseMaskColor
    PropertyChanged "UseMaskColor"
End If
End Property

Private Sub UserControl_InitProperties()
    m_hImageList = m_def_hImageList
    m_ImageHeight = m_def_ImageHeight
    m_ImageWidth = m_def_ImageWidth
    m_UseMaskColor = m_def_UseMaskColor
    m_MaskColor = m_def_Maskcolor
    m_Backcolor = m_def_Backcolor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim FileAr() As Byte
Dim i As Long
Dim Def(0) As Byte
Dim k As String
Dim t As String
Dim DStr As String
    m_Backcolor = PropBag.ReadProperty("BackColor", &H80000005)
    m_MaskColor = PropBag.ReadProperty("MaskColor", &HC0C0C0)
    m_hImageList = PropBag.ReadProperty("hImageList", m_def_hImageList)
    m_ImageHeight = PropBag.ReadProperty("ImageHeight", m_def_ImageHeight)
    m_ImageWidth = PropBag.ReadProperty("ImageWidth", m_def_ImageWidth)
    m_UseMaskColor = PropBag.ReadProperty("UseMaskColor", m_def_UseMaskColor)
    BackColorIntern = m_Backcolor
    MaskColorIntern = m_MaskColor
    Anzahl = PropBag.ReadProperty("Anzahl", 0)
    ReDim ImgArr(0)
    ReDim KeyArr(0)
    ReDim TagArr(0)
    If Anzahl > 0 Then
    ReDim ImgArr(Anzahl)
    ReDim KeyArr(Anzahl)
    ReDim TagArr(Anzahl)
    For i = 1 To Anzahl
    FileAr = PropBag.ReadProperty(i, Def)
    k = PropBag.ReadProperty(i & "k", DStr)
    t = PropBag.ReadProperty(i & "t", DStr)
    DoEvents
    PP.WriteProperty i - 1 & "i", FileAr, Def
    TagArr(i) = t
    KeyArr(i) = k
    DoEvents
    ImgArr(i) = i - 1
    Next i
    End If
    Set ListImages = New ListImages
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Long
Dim a() As Byte
Dim Def(0) As Byte
Dim Name As Long
Dim DStr As String
Dim k As String
Dim t As String

    Call PropBag.WriteProperty("BackColor", m_Backcolor, &H80000005)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, &HC0C0C0)
    Call PropBag.WriteProperty("hImageList", m_hImageList, m_def_hImageList)
    Call PropBag.WriteProperty("ImageHeight", m_ImageHeight, m_def_ImageHeight)
    Call PropBag.WriteProperty("ImageWidth", m_ImageWidth, m_def_ImageWidth)
    Call PropBag.WriteProperty("UseMaskColor", m_UseMaskColor, m_def_UseMaskColor)
    Call PropBag.WriteProperty("Anzahl", Anzahl, 0)
    If Anzahl > 0 Then
    For i = 1 To Anzahl
    Name = ImgArr(i)
    a = PP.ReadProperty(Name & "i", Def)
    k = KeyArr(i)
    t = TagArr(i)
    Call PropBag.WriteProperty(i, a, Def)
    Call PropBag.WriteProperty(i & "k", k, DStr)
    Call PropBag.WriteProperty(i & "t", t, DStr)
    Next i
    End If
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Attribute AboutBox.VB_MemberFlags = "40"
Dim x As Long
Dim y As Long
x = (Screen.Width / Screen.TwipsPerPixelX / 2) - 263
y = (Screen.Height / Screen.TwipsPerPixelY / 2) - 67
ShowAboutBox UserControl.hwnd, "Info zum Abbildungsliste-Steuerelement", "ActiveX-Steuerelement Abbildungsliste (Imagelist)," & vbCrLf & "Version 1.0" & vbCrLf & vbCrLf & "Copyright © ALKO 2006", "", UserControl.MouseIcon.handle, x, y
End Sub

