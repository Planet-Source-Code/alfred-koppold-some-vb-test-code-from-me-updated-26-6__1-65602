VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRowOfList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_ColumnText As String
Private m_Bold As Boolean
Private m_Italic As Boolean
Private m_BackColor As Long
Private m_ForeColor As Long

Private m_DetailText As String
Private m_ShowDetails As Boolean

Private m_ControlContaining As OLListe
Private m_RowInView As Boolean
Private m_RowTop As Long

Friend Property Get ParentControl() As OLListe
    ParentControl = m_ControlContaining
End Property
Friend Property Set ParentControl(New_Parent As OLListe)
    Set m_ControlContaining = New_Parent
End Property

Friend Property Get RowTop() As Long
    RowTop = m_RowTop
End Property
Friend Property Let RowTop(New_RowTop As Long)
    m_RowTop = New_RowTop
End Property

Friend Property Get RowInView() As Boolean
    RowInView = m_RowInView
End Property
Friend Property Let RowInView(New_RowInView As Boolean)
    m_RowInView = New_RowInView
End Property

Public Property Get ColumnText() As String
    ColumnText = m_ColumnText
End Property
Public Property Let ColumnText(New_sValue As String)
    m_ColumnText = New_sValue
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get Bold() As Boolean
    Bold = m_Bold
End Property
Public Property Let Bold(New_bValue As Boolean)
    m_Bold = New_bValue
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get Italic() As Boolean
    Italic = m_Italic
End Property
Public Property Let Italic(New_bValue As Boolean)
    m_Italic = New_bValue
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property
Public Property Let BackColor(New_lValue As Long)
    m_BackColor = New_lValue
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(New_lValue As Long)
    m_ForeColor = New_lValue
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get DetailText() As String
    DetailText = m_DetailText
End Property
Public Property Let DetailText(New_sValue As String)
    m_DetailText = New_sValue
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

Public Property Get ShowDetails() As Boolean
    ShowDetails = m_ShowDetails
End Property
Public Property Let ShowDetails(New_bValue As Boolean)
    m_ShowDetails = New_bValue
    If Not m_ControlContaining Is Nothing Then
        m_ControlContaining.Refresh
    End If
End Property

