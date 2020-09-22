VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColumnHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum Stil
BildundText = 0
NurBild = 1
NurText = 2
End Enum
Public Caption As String
Private m_Bold As Boolean
Private m_FontName As String
Private m_Italic As Boolean
Private m_FontUnderline As Boolean
Private m_FontStrikeout As Boolean
Private m_FontSize As Long
Private m_Picture As Long
Private m_FontColor As Long
Private m_Headerstil As Stil
Private m_Left As Long         'in Pixels
Private m_Width As Long        'In Pixels
Private m_Alignment As AlignmentConstants
Private m_ControlContaining As OLListe
Private m_ColumnWidthIsEditable As Boolean
Friend Property Set Parent(New_Parent As OLListe)
    Set m_ControlContaining = New_Parent
End Property

Friend Property Get ScaleWidth() As Long
    ScaleWidth = m_Width
End Property
Friend Property Let ScaleWidth(New_ScaleWitdh As Long)
    m_Width = New_ScaleWitdh
End Property

Friend Property Get ColLeft() As Long
    ColLeft = m_Left
End Property
Friend Property Let ColLeft(New_ColLeft As Long)
    m_Left = New_ColLeft
End Property

Public Property Get Width() As Long
    Width = m_Width * Screen.TwipsPerPixelX
End Property
Public Property Let Width(New_Width As Long)
    m_Width = New_Width / Screen.TwipsPerPixelX
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.AdjustColumnWidths
    End If
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property
Public Property Let Alignment(New_Alignment As AlignmentConstants)
    m_Alignment = New_Alignment
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get Bold() As Boolean
    Bold = m_Bold
End Property
Public Property Let Bold(New_Bold As Boolean)
    m_Bold = New_Bold
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property
Public Property Get FontName() As String
    FontName = m_FontName
End Property
Public Property Let FontName(New_FontName As String)
    m_FontName = New_FontName
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get Italic() As Boolean
    Italic = m_Italic
End Property
Public Property Let Italic(New_Italic As Boolean)
    m_Italic = New_Italic
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property


Public Property Get Picture() As Long
Picture = m_Picture
End Property

Public Property Let Picture(ByVal New_Picture As Long)
    m_Picture = New_Picture
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If

End Property

Public Property Get Headerstil() As Stil
Headerstil = m_Headerstil
End Property

Public Property Let Headerstil(ByVal New_Headerstil As Stil)
m_Headerstil = New_Headerstil
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get FontColor() As Long
    FontColor = m_FontColor
End Property
Public Property Let FontColor(New_FontColor As Long)
    m_FontColor = New_FontColor
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property
Public Property Get Fontsize() As Long
    Fontsize = m_FontSize
End Property
Public Property Let Fontsize(New_FontSize As Long)
    m_FontSize = New_FontSize
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property
Public Property Get FontUnderline() As Long
    FontUnderline = m_FontUnderline
End Property
Public Property Let FontUnderline(New_FontUnderline As Long)
    m_FontUnderline = New_FontUnderline
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property
Public Property Get FontStrikeout() As Long
    FontStrikeout = m_FontStrikeout
End Property
Public Property Let FontStrikeout(New_FontStrikeout As Long)
    m_FontStrikeout = New_FontStrikeout
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get ColumnWidthIsEditable() As Boolean
    ColumnWidthIsEditable = m_ColumnWidthIsEditable
End Property
Public Property Let ColumnWidthIsEditable(New_ColumnWidthIsEditable As Boolean)
    m_ColumnWidthIsEditable = New_ColumnWidthIsEditable
End Property

