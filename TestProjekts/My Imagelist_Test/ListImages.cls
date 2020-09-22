VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ListCollection As Collection
Private ListID As Long
Private Nr As Long

Public Function Add(Optional Index, Optional Key, Optional Picture) As ListImage
Dim ListNew As New ListImage
Dim Testtype As Long
Testtype = VarType(Index)
If Testtype <> vbError Then ListNew.Index = Index
Testtype = VarType(Key)
If Testtype <> vbError Then ListNew.Key = Key
Set ListNew.Picture = Picture
ListCollection.Add ListNew, Key
Set Add = ListNew
Set ListNew = Nothing
End Function

Public Sub Clear()
Dim i As Long

For i = 1 To ListCollection.Count
ListCollection.Remove i
Next i
End Sub

Public Property Get Count() As Integer
Count = ListCollection.Count
End Property

Public Property Get Item(Index) As ListImage
Set Item = ListCollection(Index)
End Property

Public Sub Remove(Index)
ListCollection.Remove Index
End Sub

Private Sub Class_Initialize()
Nr = 0
Set ListCollection = New Collection
End Sub

Public Property Get NewEnum() As Object
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
Set NewEnum = ListCollection.[_NewEnum]
End Property

Private Sub Class_Terminate()
Set ListCollection = Nothing
End Sub


