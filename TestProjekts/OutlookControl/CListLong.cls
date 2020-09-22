VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Anzahl Knoten, um die das Array bei ReDim vergrößert wird
Private m_lReDimAmount As Long

' aktuelle Größe der Liste
Private m_lCount As Long
' das Array mit den Daten
Private mData() As LongNode
' Index auf den aktuellen Knoten
Private mlCurrent As Long
' der eingestellte Filter
Private msFilter As String

' Auflistung mit freien Indizes im Array
Private mcolFreeIndex As Collection

' initiale Größe des internen Array und Anzahl, um die vergrößert wird festlegen
Public Sub InitList(ByVal lInitialAmount As Long, Optional ByVal lSubsequentAmount As Long = 0)
Attribute InitList.VB_Description = "Initialisierung der Liste. Es läßt sich eine initiale Größe und ein Wert für die weitere Vergrößerung festlegen."
    If lSubsequentAmount = 0 Then
        m_lReDimAmount = 1
    Else
        m_lReDimAmount = Abs(lSubsequentAmount)
    End If
    ReDim mData(0 To lInitialAmount) As LongNode
End Sub

' Filter-Muster festlegen
Public Property Get Filter() As String
Attribute Filter.VB_Description = "Setzt eine Filtereigenschaft oder gibt diese zurück. Das Setzen der Eigenschaft macht das erste Element der Liste, dessen Key-Eigenschaft dem Filter entspricht, zum aktuellen Element."
    Filter = msFilter
End Property
Public Property Let Filter(New_Filter As String)
    msFilter = New_Filter
    MoveFirst
End Property

'Prüfung auf das Ende der Liste
Public Property Get EOL() As Boolean
Attribute EOL.VB_Description = "Gibt True zurück, falls es kein aktuelles Element gibt."
    EOL = CBool(mlCurrent = 0)
End Property

'auf den ersten Knoten der Liste springen
Public Sub MoveFirst()
Attribute MoveFirst.VB_Description = "Macht das erste Element zum aktuellen Element."
    mlCurrent = mData(0).Next
    Do While Not mData(mlCurrent).Key Like msFilter
        If mlCurrent = 0 Then Exit Do
        mlCurrent = mData(mlCurrent).Next
    Loop
End Sub

'auf den letzen Knoten der Liste springen
Public Sub MoveLast()
Attribute MoveLast.VB_Description = "Macht das letzte Element zum aktuellen Element."
    mlCurrent = mData(0).Prev
    Do While Not mData(mlCurrent).Key Like msFilter
        If mlCurrent = 0 Then Exit Do
        mlCurrent = mData(mlCurrent).Prev
    Loop
End Sub

'auf den nächsten Knoten der Liste springen
Public Sub MoveNext()
    If mlCurrent = 0 Then Exit Sub
    mlCurrent = mData(mlCurrent).Next
    Do While Not mData(mlCurrent).Key Like msFilter
        If mlCurrent = 0 Then Exit Do
        mlCurrent = mData(mlCurrent).Next
    Loop
End Sub

'auf den vorigen Knoten der Liste springen
Public Sub MovePrev()
Attribute MovePrev.VB_Description = "Macht das vorige Element zum aktuellen Element."
    If mlCurrent = 0 Then Exit Sub
    mlCurrent = mData(mlCurrent).Prev
    Do While Not mData(mlCurrent).Key Like msFilter
        If mlCurrent = 0 Then Exit Do
        mlCurrent = mData(mlCurrent).Prev
    Loop
End Sub
Public Sub Move(lIndex As Long)
    Dim i As Long
    
    If lIndex < 1 Or lIndex > m_lCount Then Err.Raise 9, "LinkedList", "Out of list"
    mlCurrent = mData(0).Next
    For i = 2 To lIndex
        mlCurrent = mData(mlCurrent).Next
    Next i
End Sub

Public Property Get Count() As Long
Attribute Count.VB_Description = "Anzahl der Elemente in der Liste. (Diese Eigenschaft ändert sich NICHT durch die Anwendung eines Filters!)"
    Count = m_lCount
End Property

Public Sub Add(lValue As Long, Optional ByVal sKey As String = "")
Attribute Add.VB_Description = "Fügt der Liste (am Ende) ein neues Element hinzu."
    Dim i As Long
    
    If m_lReDimAmount = 0 Then
        'not yet initialized, so do it now
        InitList 100, 5
    End If
    
    i = NextFreeIndex
        
    m_lCount = m_lCount + 1
    mData(i).Data = lValue
    mData(i).Key = sKey
    mData(i).Next = 0
    mData(i).Prev = mData(0).Prev       'the old LastNode
    mData(mData(i).Prev).Next = i
    mData(0).Prev = i
End Sub

Public Sub Insert(lValue As Long, Optional ByVal sKey As String = "")
Attribute Insert.VB_Description = "Einfügen eines neuen Elementes vor dem aktuellen Element."

    Dim i As Long
    
    If m_lReDimAmount = 0 Then
        'not yet initialized, so do it now
        InitList 100, 5
    End If
    
    i = NextFreeIndex
    
    m_lCount = m_lCount + 1
    mData(i).Data = lValue
    mData(i).Key = sKey
    If mlCurrent = 0 Then
        'insert as the first Node:
        mData(i).Prev = 0
        mData(i).Next = mData(0).Next
        mData(0).Next = i
        If mData(i).Next <> 0 Then mData(mData(i).Next).Prev = i
        If mData(0).Prev = 0 Then mData(0).Prev = i
    Else
        mData(i).Next = mlCurrent
        mData(i).Prev = mData(mlCurrent).Prev
        mData(mData(i).Prev).Next = i
        mData(mlCurrent).Prev = i
    End If
    
End Sub

Public Sub Remove()
Attribute Remove.VB_Description = "Löscht das aktuelle Element. Das neue aktuelle Element wird das nächste in der Liste."
    Dim i As Long
    'remove the current node
    'der neue Nachfolger meines Vorgängers ist mein Nachfolger
    i = mlCurrent
    mData(mData(i).Prev).Next = mData(i).Next
    'der neue Vorgänger meines Nachfolgers ist mein Vorgänger
    mData(mData(i).Next).Prev = mData(i).Prev
    mlCurrent = mData(i).Next
    mData(i).Prev = 0
    mData(i).Next = 0
    mData(i).Data = 0
    m_lCount = m_lCount - 1
    mcolFreeIndex.Add i
End Sub

Public Sub RemoveAll()
    MoveFirst
    While Count > 0
        Remove
    Wend
End Sub

Public Property Get Data() As Long
Attribute Data.VB_Description = "Wert des aktuellen Elementes."
Attribute Data.VB_UserMemId = 0
    If mlCurrent = 0 Then Err.Raise 9, "LinkedList", "Out of list"
    Data = mData(mlCurrent).Data
End Property
Public Property Let Data(New_lValue As Long)
    If mlCurrent = 0 Then Err.Raise 9, "LinkedList", "Out of list"
    mData(mlCurrent).Data = New_lValue
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "String, der für das Filtern von Elementen eingesetzt werden kann."
    If mlCurrent = 0 Then Err.Raise 9, "LinkedList", "Out of list"
    Key = mData(mlCurrent).Key
End Property
Public Property Let Key(New_Key As String)
    If mlCurrent = 0 Then Err.Raise 9, "LinkedList", "Out of list"
    mData(mlCurrent).Key = New_Key
End Property

Public Property Get Item(lIndex As Long) As Long
Attribute Item.VB_Description = "Gibt einen Wert anhand des Indizes zurück."
    Dim i As Long
    Dim Idx As Long
    If lIndex < 0 Then Err.Raise 9
    If lIndex > m_lCount Then Err.Raise 9
    
    If lIndex < m_lCount / 2 Then
        'vorwärts suchen ...
        Idx = mData(0).Next
        For i = 1 To lIndex - 1
            Idx = mData(Idx).Next
        Next i
        Item = mData(Idx).Data
    Else
        'rückwärts suchen ...
        Idx = mData(0).Prev
        For i = 1 To m_lCount - lIndex
            Idx = mData(Idx).Prev
        Next i
        Item = mData(Idx).Data
    End If
End Property
Public Property Let Item(lIndex As Long, New_lValue As Long)
    Dim i As Long
    Dim Idx As Long
    If lIndex < 0 Then Err.Raise 9
    If lIndex > m_lCount Then Err.Raise 9
    
    If lIndex < m_lCount / 2 Then
        'vorwärts suchen ...
        Idx = mData(0).Next
        For i = 1 To lIndex - 1
            Idx = mData(Idx).Next
        Next i
        mData(Idx).Data = New_lValue
    Else
        'rückwärts suchen ...
        Idx = mData(0).Prev
        For i = 1 To m_lCount - lIndex
            Idx = mData(Idx).Prev
        Next i
        mData(Idx).Data = New_lValue
    End If
End Property

Private Function SearchDown(lIdxAb As Long, sKey As String) As Long
    If lIdxAb = 0 Then
        SearchDown = lIdxAb
        Exit Function
    End If
    If mData(lIdxAb).Key <= sKey Then
        SearchDown = lIdxAb
        Exit Function
    End If
    SearchDown = SearchDown(mData(lIdxAb).Prev, sKey)
End Function
Private Function SearchUp(lIdxAb As Long, sKey As String) As Long
    If lIdxAb = 0 Then
        SearchUp = lIdxAb
        Exit Function
    End If
    If mData(lIdxAb).Key >= sKey Then
        SearchUp = lIdxAb
        Exit Function
    End If
    SearchUp = SearchUp(mData(lIdxAb).Next, sKey)
End Function


Public Sub Sort()
Attribute Sort.VB_Description = "Sortieren der Liste nach den Keys."
    Dim iIdx As Long
    Dim iNext As Long
    Dim iExchange As Long
    Dim bChanged As Boolean
    
    Dim iTemp As Long
    
    If m_lCount <= 1 Then Exit Sub
        
    iIdx = mData(0).Next        'the first element
    iNext = mData(iIdx).Next    'the second element
    Do
        bChanged = False
        If mData(iIdx).Prev <> 0 Then
            'search downward the first element that is smaller or equal
            iExchange = SearchDown(mData(iIdx).Prev, mData(iIdx).Key)
            If iExchange <> mData(iIdx).Prev Then
                'remove myself
                mData(mData(iIdx).Prev).Next = iNext
                mData(iNext).Prev = mData(iIdx).Prev
                
                'insert myself behind iExchange
                mData(iIdx).Next = mData(iExchange).Next
                mData(mData(iIdx).Next).Prev = iIdx
                mData(iExchange).Next = iIdx
                mData(iIdx).Prev = iExchange
                
                'Flag the exchange
                bChanged = True
            End If
        End If
        
        If Not bChanged Then
            'search upward the first element that is larger or equal!
            If iNext <> 0 Then
                iExchange = SearchUp(iNext, mData(iIdx).Key)
                If iExchange <> iNext Then
                    'remove myself
                    mData(mData(iIdx).Prev).Next = iNext
                    mData(iNext).Prev = mData(iIdx).Prev
                    
                    'insert myself before iExchange
                    mData(mData(iExchange).Prev).Next = iIdx
                    mData(iIdx).Next = iExchange
                    mData(iIdx).Prev = mData(iExchange).Prev
                    mData(iExchange).Prev = iIdx
                End If
            End If
        End If
        
        iIdx = iNext
        iNext = mData(iIdx).Next
        
    Loop Until iIdx = 0
        
End Sub

Private Sub EnlargeArray()
    Dim i As Long
    Dim j As Long
    
    i = UBound(mData) + m_lReDimAmount
    ReDim Preserve mData(0 To i) As LongNode
    For j = UBound(mData) - m_lReDimAmount + 1 To UBound(mData)
        mcolFreeIndex.Add j
    Next j
End Sub

Private Function NextFreeIndex() As Long
    If mcolFreeIndex.Count = 0 Then
        EnlargeArray
    End If
    NextFreeIndex = mcolFreeIndex(1)
    mcolFreeIndex.Remove 1
End Function

Private Sub Class_Initialize()
    Set mcolFreeIndex = New Collection
    ReDim mData(0 To 0) As LongNode
    m_lReDimAmount = 0
    mlCurrent = 0
    m_lCount = 0
    msFilter = "*"
End Sub

Private Sub Class_Terminate()
    Erase mData
    Set mcolFreeIndex = Nothing
    m_lReDimAmount = 0
    mlCurrent = 0
    m_lCount = 0
End Sub

