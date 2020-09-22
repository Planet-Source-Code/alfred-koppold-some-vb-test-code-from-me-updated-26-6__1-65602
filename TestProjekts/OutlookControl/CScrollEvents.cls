VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScrollEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public hwnd As Long

Public Event ScrollLine(ByVal lLines As Long)
Public Event ScrollPage(ByVal lPages As Long)
Public Event ScrollPos(ByVal lPos As Long)
Public Event ScrollTrack(ByVal lPos As Long)
Public Event Scroll()
Public Event count(ByVal count As Long)

Public Sub TriggerScrollLine(ByVal lLines As Long)
    RaiseEvent ScrollLine(lLines)
End Sub
Public Sub TriggerScrollPages(ByVal lPages As Long)
    RaiseEvent ScrollPage(lPages)
End Sub
Public Sub TriggerScrollPos(ByVal lPos As Long)
    RaiseEvent ScrollPos(lPos)
End Sub
Public Sub TriggerScrollTrack(ByVal lPos As Long)
    RaiseEvent ScrollTrack(lPos)
End Sub
Public Sub TriggerScroll()
    RaiseEvent Scroll
End Sub
Public Sub TriggerCount(ByVal count As Long)
    RaiseEvent count(count)
End Sub

