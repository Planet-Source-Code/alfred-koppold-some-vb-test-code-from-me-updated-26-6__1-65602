VERSION 5.00
Begin VB.UserControl OLListe 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   EditAtDesignTime=   -1  'True
   MouseIcon       =   "OLListe.ctx":0000
   PropertyPages   =   "OLListe.ctx":0152
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   680
   ToolboxBitmap   =   "OLListe.ctx":0195
End
Attribute VB_Name = "OLListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private Looking As Boolean
Private active As Boolean
Private Ending As Boolean

Enum Zeiger
Default = 0
Kreuz1 = 1
Kreuz2 = 2
Kreuz3 = 3
End Enum
'Standardwerte der Eigenschaften
Const m_def_BackColor = vbWindowBackground
Const m_def_ForeColor = vbWindowText
Private a As StdPicture

Private Const MINCOLWIDTH = 16
Private ToLittle As Boolean
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'für die Scrollbar:
Private WithEvents Scrollbar As CScrollEvents
Attribute Scrollbar.VB_VarHelpID = -1
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Const SB_VERT = 1
Private Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Const ESB_ENABLE_BOTH = &H0
Private Const ESB_DISABLE_BOTH = &H3
Private Declare Function InvertRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'für den Mauszeiger:
Private Declare Function CreateCursor Lib "user32" (ByVal hInstance As Long, ByVal nXhotspot As Long, ByVal nYhotspot As Long, ByVal nWidth As Long, ByVal nHeight As Long, lpANDbitPlane As Any, lpXORbitPlane As Any) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetClassWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const GCW_HCURSOR = (-12)
Private Const IDC_SIZEALL = 32646&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_ARROW = 32512&
Private SysCursHandle As Long
Private Curs1Handle As Long
Private Curs2Handle As Long
Const Mask1 = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFEFFFFFFFEFFFFFFFEFFFFFFFEFFFFFFBEFDFFFF3EFCFFFE02807FFC02803FFE02807FFF3EFCFFFFBEFDFFFFFEFFFFFFFEFFFFFFFEFFFFFFFEFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF"
Const Mask2 = "0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"
Const Mask1a = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFC7FFFFFFC7FFFFFFC7FFFFFFC7FFFFFBC7DFFFF1C78FFFE00007FFC00003FF800001FFC00003FFE00007FFF1C78FFFFBC7DFFFFFC7FFFFFFC7FFFFFFC7FFFFFFC7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF"
Const Mask2a = "00000000000000000000000000000000000380000002800000028000000280000042820000A28500013EFC80020280400402802002028040013EFC8000A2850000428200000280000002800000028000000380000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"
Const Mask1b = "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF81FFFFFF81FFFFFF81FFFFFF81FFFFFD81BFFFF8811FFFF0000FFFE00007FFF0000FFFF8811FFFFD81BFFFFF81FFFFFF81FFFFFF81FFFFFF81FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF"
Const Mask2b = "00000000000000000000000000000000000000000000000000000000000000000007E0000005A0000005A0000005A0000025A4000055AA00009DB90001018080009DB9000055AA000025A4000005A0000005A0000005A0000007E000000000000000000000000000000000000000000000000000000000000000000000000000"
'Standard-Eigenschaftswerte:
Const m_def_Mauszeiger = 0
'Eigenschaftsvariablen:
Dim m_Mauszeiger As Zeiger


'Eigenschaftsvariablen:
Private m_BackColor As Long
Private m_ForeColor As Long
Private m_AllowColumnResize As Boolean
Private m_HilightSelectedRow As Boolean
Private m_HideSelection As Boolean
Private m_Columns As Integer
Private m_ColumnHeaders() As CColumnHeader
Private m_SelectedRow As Long
Private m_Columnstil As Stil
Private m_FirstRow As Long

'die Liste mit den Zeilen
Private mlistRows As CListObject

'interne Variablen
Private miResizeCol As Integer
Private miResizePos As Integer
Private m_HasFocus As Boolean
Private Colgedrückt As Single
Private RectColumn As RECT
Private Rectold1 As RECT
Private Rectold2 As RECT
Private angefangen As Boolean
Private IconuPic As Long
'Ereignisdeklarationen:
Event SelectionChanged()
Event Click()
Event DblClick(Cancel As Boolean)
Event ColumnClick(ByVal Column As Integer, ByVal Button As Integer, ByVal Shift As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'Hintergrundfarbe des Controls
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    UserControl.Refresh
    PropertyChanged "BackColor"
End Property

'Vordergrundfarbe des Controls
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    UserControl.Refresh
    PropertyChanged "ForeColor"
End Property

'ist das Control enabled?
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    UserControl.Refresh
    PropertyChanged "Enabled"
End Property

'die Schriftart des Controls
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    With UserControl
    Set .Font = New_Font
    'Fett und Kursiv sind nicht zulässig, dies muß zeilenweise angegeben werden
    .Font.Bold = False
    .Font.Italic = False
    .Refresh
    End With
    PropertyChanged "Font"
End Property

'Control ohne Rahmen
Public Property Get Borderless() As Boolean
    Borderless = CBool(UserControl.BorderStyle = 0)
End Property
Public Property Let Borderless(ByVal New_Borderless As Boolean)
    With UserControl
    .BorderStyle = IIf(New_Borderless, 0, 1)
    'eine Änderung am BorderStyle resettet die Hintergrundfarbe,
    ' sie muß hier deshalb neu gesetzt werden
    .BackColor = m_BackColor
    .Refresh
    End With
    PropertyChanged "Borderless"
End Property

'Control nicht in 3D sondern flach
Public Property Get Flat() As Boolean
    Flat = CBool(UserControl.Appearance = 0)
End Property
Public Property Let Flat(ByVal New_Flat As Boolean)
    With UserControl
    .Appearance = IIf(New_Flat, 0, 1)
    .BackColor = m_BackColor
    .Refresh
    End With
    PropertyChanged "Flat"
End Property

'darf der Benutzer die Spaltenbreite ändern?
Public Property Get AllowColumnResize() As Boolean
    AllowColumnResize = m_AllowColumnResize
End Property
Public Property Let AllowColumnResize(ByVal New_AllowColumnResize As Boolean)
    m_AllowColumnResize = New_AllowColumnResize
    PropertyChanged "AllowColumnResize"
End Property

'soll die selektierte Zeile mit einer anderen Hintergrundfarbe
'gekennzeichnet werden?
Public Property Get HilightSelectedRow() As Boolean
    HilightSelectedRow = m_HilightSelectedRow
End Property
Public Property Let HilightSelectedRow(New_bValue As Boolean)
    m_HilightSelectedRow = New_bValue
    UserControl.Refresh
    PropertyChanged "HilightSelectedRow"
End Property

'soll die Markierung der selektierten Zeile erhalten bleiben,
' wenn der Fokus das Control verläßt?
Public Property Get HideSelection() As Boolean
    HideSelection = m_HideSelection
End Property
Public Property Let HideSelection(New_bValue As Boolean)
    m_HideSelection = New_bValue
    UserControl.Refresh
    PropertyChanged "HideSelection"
End Property

'Neuzeichnen des Controls erzwingen
Public Sub Refresh()
Attribute Refresh.VB_Description = "Erzwingt ein vollständiges Neuzeichnen eines Objekts."
    UserControl.Refresh
End Sub

'Anzahl der Spalten im Control
Public Property Get Columns() As Integer
Attribute Columns.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Columns = m_Columns
End Property
Public Property Let Columns(New_Columns As Integer)
    Dim i As Integer
    Dim llColLeft As Long
    
    'mindestens eine Spalte !
    If New_Columns < 1 Then New_Columns = 1
    'nicht mehr benötigte Spaltenköpfe löschen
    For i = New_Columns + 1 To m_Columns
        Set m_ColumnHeaders(i) = Nothing
    Next i
    'Array neu dimensionieren
    ReDim Preserve m_ColumnHeaders(1 To New_Columns) As CColumnHeader
    llColLeft = 0
    Select Case New_Columns
    Case Is >= m_Columns
    'alle Spalten gleich verteilen
    For i = 1 To m_Columns
        m_ColumnHeaders(i).ScaleWidth = UserControl.ScaleWidth / New_Columns
        m_ColumnHeaders(i).ColLeft = llColLeft
        llColLeft = llColLeft + m_ColumnHeaders(i).ScaleWidth
    Next i
    'neue Spalten erzeugen und ebenfalls gleich verteilen
    For i = m_Columns + 1 To New_Columns
        Set m_ColumnHeaders(i) = New CColumnHeader
        m_ColumnHeaders(i).Caption = "Column" & i
        m_ColumnHeaders(i).Alignment = vbLeftJustify
        m_ColumnHeaders(i).ScaleWidth = UserControl.ScaleWidth / New_Columns
        m_ColumnHeaders(i).ColLeft = llColLeft
        llColLeft = llColLeft + m_ColumnHeaders(i).ScaleWidth
        Set m_ColumnHeaders(i).Parent = Me
    Next i
    Case Else
        For i = 1 To New_Columns
        m_ColumnHeaders(i).ScaleWidth = UserControl.ScaleWidth / New_Columns
        m_ColumnHeaders(i).ColLeft = llColLeft
        llColLeft = llColLeft + m_ColumnHeaders(i).ScaleWidth
    Next i

    End Select
    m_Columns = New_Columns
    UserControl.Refresh
    PropertyChanged "Columns"
End Property

'Spaltenbezeichnung zu einer Spalte liefern
Public Property Get ColumnHeaders(Index As Integer) As CColumnHeader
    If Index < 1 Or Index > m_Columns Then Err.Raise 9  'Index out of range
    Set ColumnHeaders = m_ColumnHeaders(Index)
End Property

'Anzahl der Zeilen
Public Property Get Rows() As Long
    Rows = mlistRows.Count
End Property

'die aktuell selektierte Spalte
Public Property Get SelectedRow() As Long
    SelectedRow = m_SelectedRow
End Property
Public Property Let SelectedRow(New_SelectedRow As Long)
    m_SelectedRow = New_SelectedRow
    UserControl.Refresh
End Property

'sortieren der Spalten nach deren Sortierschlüssel
Public Sub Sort()
    mlistRows.Sort
    m_SelectedRow = 1
    UserControl.Refresh
    RaiseEvent SelectionChanged
End Sub

'bei allen Zeilen den Detail-Bereich anzeigen oder ausblenden
Public Sub ShowAllDetails(bShow As Boolean)
    Dim lobjRow As CRowOfList
    
    mlistRows.MoveFirst
    While Not mlistRows.EOL
        Set lobjRow = mlistRows.Data
        If lobjRow.DetailText <> "" Then
            lobjRow.ShowDetails = bShow
        End If
        mlistRows.MoveNext
    Wend
End Sub

'eine neue Zeile hinzufügen
Public Sub AddRow(ColumnText As String, Bold As Boolean, Italic As Boolean, DetailText As String, ShowDetails As Boolean, Optional SortKey As String = "")
    Dim lobjRow As CRowOfList
    
    'Zeile erzeugen
    Set lobjRow = New CRowOfList
    'Eigenschaften übernehmen
    With lobjRow
    .ColumnText = ColumnText
    .Bold = Bold
    .Italic = Italic
    .DetailText = DetailText
    .ShowDetails = ShowDetails
    .BackColor = m_BackColor
    .ForeColor = m_ForeColor
    'Verweis auf das Control setzen
    Set .ParentControl = Me
    End With
    'am Ende anfügen
    mlistRows.Add lobjRow, SortKey

    'Scrollbar anpassen
    SetScrollRange UserControl.hWnd, SB_VERT, 1, mlistRows.Count, True
    If mlistRows.Count > 1 Then
        EnableScrollBar UserControl.hWnd, SB_VERT, ESB_ENABLE_BOTH
    End If
    '... und neuzeichen nicht vergessen
    UserControl.Refresh
    
End Sub

'den Inhalt einer Zeile liefern
Public Property Get Row(Index As Long) As CRowOfList
    If Index < 1 Or Index > mlistRows.Count Then Err.Raise 9
    Set Row = mlistRows.Item(Index)
End Property

'liefern der Spaltennummer zur Mausposition
Public Property Get ColContaining(x As Single) As Integer
    Dim i As Long
    
    If x < m_ColumnHeaders(1).ColLeft Then
        ColContaining = 0
        Exit Property
    End If
    For i = 1 To m_Columns
        If i < m_Columns Then
            If x < m_ColumnHeaders(i + 1).ColLeft Then Exit For
        Else
            If x < UserControl.ScaleWidth Then Exit For
        End If
    Next i
    ColContaining = i
    
End Property

'liefern der Zeilennummer zur Mausposition
Public Property Get RowContaining(y As Single) As Long
    Dim i As Long
    Dim lTop As Long
    Dim lHeight As Long
    Dim lobjRow As CRowOfList
    
    With UserControl
    lTop = .TextHeight("any text") + 4
    If y <= lTop Then
        RowContaining = 0
        Exit Property
    End If
    If mlistRows.Count = 0 Then
        RowContaining = 0
        Exit Function
    End If
    For i = m_FirstRow To mlistRows.Count
        Set lobjRow = mlistRows.Item(i)
        'Gesamthöhe der Zeile bestimmen
        lHeight = .TextHeight("Text") + 4
        'Gesamthöhe erhöht sich, wenn die Details gezeigt werden
        If lobjRow.ShowDetails Then
            lHeight = lHeight + 4 + modListe.CountLines(lobjRow.DetailText) * .TextHeight("Text")
        End If
        If y <= lTop + lHeight Then
            RowContaining = i
            Exit Property
        End If
        lTop = lTop + lHeight
    Next i
    RowContaining = mlistRows.Count
    End With
End Property

'ist die Zeile sichtbar?
Public Property Get RowIsVisible(Index As Long) As Boolean
    Dim lobjRow As CRowOfList
    
    If Index < 1 Or Index > mlistRows.Count Then Err.Raise 9
    Set lobjRow = mlistRows.Item(Index)
    RowIsVisible = lobjRow.RowInView
End Property

'Sichtbarkeit der Zeile sicherstellen
Public Sub EnsureVisible(RowIndex As Long)
    Dim lobjRow As CRowOfList
    
    If RowIndex < 1 Or RowIndex > mlistRows.Count Then Err.Raise 9 'Index out of range
    Set lobjRow = mlistRows.Item(RowIndex)
    If lobjRow.RowInView Then Exit Sub 'Zeile ist bereits sichtbar
    If lobjRow.RowTop = 0 Then
        'die Zeile ist oberhalb der ersten Zeile,
        ' also diese Zeile zur ersten Zeile machen
        m_FirstRow = RowIndex
        SetScrollPos UserControl.hWnd, SB_VERT, m_FirstRow, True
        UserControl_Paint
    Else
        'solange nach oben scrollen (erste Zeile ändern),
        'bis die Zeile sichtbar ist.
        Do While Not lobjRow.RowInView
            If m_FirstRow = RowIndex Then Exit Do
            m_FirstRow = m_FirstRow + 1
            SetScrollPos UserControl.hWnd, SB_VERT, m_FirstRow, True
            UserControl_Paint
        Loop
    End If

End Sub

' der User hat irgendwie Scrolling ausgelöst (und ist jetzt fertig)
Private Sub ScrollBar_Scroll()
    Dim lNewFirstRow As Long
        
    'neue erste Zeile aus der Scrollbar bestimmen
    lNewFirstRow = GetScrollPos(UserControl.hWnd, SB_VERT)
    'Scrollbar aktualisieren
    SetScrollPos UserControl.hWnd, SB_VERT, lNewFirstRow, True
    'neue erste Zeile setzen
    If lNewFirstRow <> m_FirstRow Then
        m_FirstRow = lNewFirstRow
        UserControl_Paint
    End If
End Sub

' der User hat um eine oder mehrere Zeilen gescrollt
' (Klick auf die Buttons mit den Pfeilen)
Private Sub ScrollBar_ScrollLine(ByVal lLines As Long)
    Dim lNewFirstRow As Long
        
    'neue erste Zeile bestimmen
    lNewFirstRow = m_FirstRow + lLines
    If lNewFirstRow < 1 Then lNewFirstRow = 1
    If lNewFirstRow > mlistRows.Count Then lNewFirstRow = mlistRows.Count
    'Scrollbar aktualisieren
    SetScrollPos UserControl.hWnd, SB_VERT, lNewFirstRow, True
    'neue erste Zeile setzen
    If m_FirstRow <> lNewFirstRow Then
        m_FirstRow = lNewFirstRow
        UserControl_Paint
    End If
End Sub

' der User hat um eine oder mehrere Seiten gescrollt
' (Klick auf den Hintergrund der Scrollbar)
Private Sub ScrollBar_ScrollPage(ByVal lPages As Long)
    Dim lNewFirstRow As Long
        
    'neue erste Zeile bestimmen
    lNewFirstRow = m_FirstRow + (lPages * 3)
    If lNewFirstRow < 1 Then lNewFirstRow = 1
    If lNewFirstRow > mlistRows.Count Then lNewFirstRow = mlistRows.Count
    'Scrollbar aktualisieren
    SetScrollPos UserControl.hWnd, SB_VERT, lNewFirstRow, True
    'neue erste Zeile setzen
    If m_FirstRow <> lNewFirstRow Then
        m_FirstRow = lNewFirstRow
        UserControl_Paint
    End If
End Sub

'User hat den Scrollbalken bewegt
Private Sub ScrollBar_ScrollPos(ByVal lPos As Long)
    'Scrollbar aktualisieren
    SetScrollPos UserControl.hWnd, SB_VERT, lPos, True
    'neue erste Zeile setzen
    If m_FirstRow <> lPos Then
        m_FirstRow = lPos
        UserControl_Paint
    End If
End Sub

'USer bewegt gerade den Scrollbalken
Private Sub ScrollBar_ScrollTrack(ByVal lPos As Long)
    
    'man könnte hier auch die Zeilennummer als Tooltip anzeigen
    'sollte man auch, wenn das Zeichnen zu lange dauert
    
    'wir ändern hier direkt die erste Zeile
    If m_FirstRow <> lPos Then
        m_FirstRow = lPos
        UserControl_Paint
    End If
End Sub

'User hat einen Doppelklick ausgelöst
Private Sub Usercontrol_DblClick()
    Dim objRow As CRowOfList
    Dim bCancel As Boolean

    'dem Entwickler das Event mitteilen und die Möglichkeit geben,
    'die Standardbehandlung abzubrechen
    RaiseEvent DblClick(bCancel)
    
    If Not bCancel Then
        'ansonsten für die aktuelle Zeile ...
        Set objRow = mlistRows.Item(m_SelectedRow)
        If objRow.DetailText <> "" Then
            ' ... die Details anzeigen oder verbergen
            objRow.ShowDetails = Not objRow.ShowDetails
            UserControl_Paint
        End If
    End If
End Sub

'das Control hat den Fokus erhalten
Private Sub UserControl_EnterFocus()
    'Status umsetzen (für Fokusrechteck oder Farbgebung)
    m_HasFocus = True
    'neu zeichnen
    UserControl_Paint
End Sub

'das Control hat den Fokus abgegeben
Private Sub UserControl_ExitFocus()
    'Status umsetzen (für Fokusrechteck oder Farbgebung)
    m_HasFocus = False
    'neu zeichnen
    UserControl_Paint
End Sub

'der User hat eine Taste niedergedrückt
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objRow As CRowOfList
    
    'dem Entwickler die Möglichkeit geben, dieses Ereignis selbst zu bearbeiten
    ' und die weitere Verarbeitung zu unterdrücken oder zu ändern
    RaiseEvent KeyDown(KeyCode, Shift)
    Select Case KeyCode
        Case vbKeyReturn
            'die Details der selektierten Zeile anzeigen/verstecken
            Set objRow = mlistRows.Item(m_SelectedRow)
            If objRow.DetailText <> "" Then
                objRow.ShowDetails = Not objRow.ShowDetails
                UserControl_Paint
            End If
            'Reaktion auf diese Taste ist erfolgt,
            'also weitere (Standard-)Verarbeitung unterdrücken
            KeyCode = 0
            Exit Sub
        Case vbKeyUp
            'die Zeile oberhalb der aktuell selektierten Zeile selekieren
            If m_SelectedRow > 1 Then
                m_SelectedRow = m_SelectedRow - 1
                UserControl_Paint
                RaiseEvent SelectionChanged
            End If
            'falls Zeile nicht sichtbar, dann anzeigen
            If Not RowIsVisible(m_SelectedRow) Then EnsureVisible m_SelectedRow
            'Reaktion auf diese Taste ist erfolgt,
            'also weitere (Standard-)Verarbeitung unterdrücken
            KeyCode = 0
            Exit Sub
        Case vbKeyDown
            'die Zeile unterhalb der aktuell selektierten Zeile selektieren
            If m_SelectedRow < mlistRows.Count Then
                m_SelectedRow = m_SelectedRow + 1
                UserControl_Paint
                RaiseEvent SelectionChanged
            Else
                'zu ende Scrollen erlauben, bis selected row die first row ist!
                If m_SelectedRow > m_FirstRow Then
                    m_FirstRow = m_FirstRow + 1
                    SetScrollPos UserControl.hWnd, SB_VERT, m_FirstRow, True
                    UserControl_Paint
                End If
            End If
            'falls Zeile nicht sichtbar, dann anzeigen
            If Not RowIsVisible(m_SelectedRow) Then EnsureVisible m_SelectedRow
            'Reaktion auf diese Taste ist erfolgt,
            'also weitere (Standard-)Verarbeitung unterdrücken
            KeyCode = 0
            Exit Sub
    End Select
End Sub

'der User hat eine Ascii-Taste gedrückt (und losgelassen)
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    'Event weiterleiten
    RaiseEvent KeyPress(KeyAscii)
End Sub

'der User hat eine Taste losgelassen
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    'Event weiterleiten
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'der User hat eine Maustaste niedergedrückt
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim RectColumn As RECT
    'Event an Entwickler weiterleiten
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    If Button = vbLeftButton Then
        If y < UserControl.TextHeight("Text") + 4 Then
            'falls auf Spaltenköpfen
            If Not Ambient.UserMode Or m_AllowColumnResize Then
                'falls Umgebung (Form) im Entwicklungsmodus
                'oder Resizing erlaubt
                For i = 2 To m_Columns
                    If Abs(x - m_ColumnHeaders(i).ColLeft) <= 2 Then
                        'falls auf Spaltengrenze
                        'Spaltenbreitenänderung beginnen
                        miResizeCol = i - 1
                        miResizePos = x
                    'Mauszeiger ändern
                    'UserControl.MousePointer = vbSizeWE
                    Select Case m_Mauszeiger
                    Case 0
                        Curs1Handle = LoadCursor(ByVal 0&, IDC_SIZEWE)
                        Case Else
                        Curs1Handle = CreateNewCursor
                        End Select
    'Set the form's mouse cursor
    'SysCursHandle = SetClassWord(UserControl.hWnd, GCW_HCURSOR, Curs1Handle)
    SysCursHandle = SetCursor(Curs1Handle)
                        'Hier ändern
                            If miResizeCol <> 0 Then
                            drawLines (1)
                            End If

                        Exit For
                    End If
                Next i
                    If miResizeCol = 0 Then

                        i = ColContaining(x)
                    If i > 0 And i <= m_Columns Then
                    'Maustaste gedrückt
                    Colgedrückt = i
                    'UserControl_Paint
                    DrawColumnHeader (i)
                    End If
                    End If
            End If
        Else
            'nicht im Spaltenkopf, also auf irgendeiner Zeile
            If y > UserControl.ScaleHeight Then Exit Sub
            If x > 0 And x < UserControl.ScaleWidth Then
                'angeklickte Zeile bestimmen
                i = RowContaining(y)
                If i > 0 And i <= mlistRows.Count Then
                    'angeklickte Zeile selektieren
                    m_SelectedRow = i
                    UserControl_Paint
                    RaiseEvent SelectionChanged
                    EnsureVisible m_SelectedRow
                End If
            End If
        End If
    End If
    
End Sub

'der User hat die Maus bewegt
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim lbDoResize As Boolean
    Dim WidthWithScroll As Long
    Dim Übergabe As Long
    'ReleaseCapture

    WidthWithScroll = UserControl.Width / Screen.TwipsPerPixelX
    'If active = False Then
    'If x > 0 And x < WidthWithScroll And y > 0 And y < UserControl.ScaleHeight Then ' Aktivieren
        'Subclassing aktivieren
    'Set Scrollbar = HookWindow(UserControl.hWnd)
    'active = True
    'End If
    'End If
    'If active = True Then
    'If x < 0 Or x > WidthWithScroll Or y < 0 Or y > UserControl.ScaleHeight Then 'deaktivieren
    'Subclassing beenden
    'UnHookWindow UserControl.hWnd
     'active = False
    'End If
    'End If
    'falls Spaltenbreitenänderung aktiv ist
    If miResizeCol <> 0 Then
        lbDoResize = False
        If x > m_ColumnHeaders(miResizeCol).ColLeft + MINCOLWIDTH Then ToLittle = False

    If x < m_ColumnHeaders(miResizeCol).ColLeft + MINCOLWIDTH Then
    If ToLittle = False Then
    x = m_ColumnHeaders(miResizeCol).ColLeft + MINCOLWIDTH + 1
ToLittle = True
End If
End If
        'Breitenänderung nur bis zur Mindestbreite
        If x > m_ColumnHeaders(miResizeCol).ColLeft + MINCOLWIDTH Then 'wenn größer als Mindestbreite
            If miResizeCol + 1 < m_Columns Then 'wenn nicht letzte Spalte
                If x < m_ColumnHeaders(miResizeCol + 2).ColLeft - MINCOLWIDTH Then lbDoResize = True
            Else
                If x < UserControl.ScaleWidth - MINCOLWIDTH Then lbDoResize = True 'wenn letzte Spalte
            End If
        End If
        'falls Breitenänderung ok
        If lbDoResize Then
            'dann Position (für die Linie) merken
            miResizePos = x
                                    'Hier ändern
                            If miResizeCol <> 0 Then
                            drawLines (0)
                            End If

            'UserControl_Paint
        End If
        Exit Sub
    End If
    
    'falls auf den Spaltenköpfen
    If y < UserControl.TextHeight("Text") + 4 Then
    If y < 0 Then
    If Colgedrückt <> -1 Then
    Übergabe = Colgedrückt
    Colgedrückt = -1
    DrawColumnHeader (Übergabe)
    'UserControl_Paint
    End If
    End If
        'UserControl.MousePointer = vbDefault
                                Curs1Handle = LoadCursor(ByVal 0&, IDC_ARROW)
    'Set the form's mouse cursor
    'SysCursHandle = SetClassWord(UserControl.hWnd, GCW_HCURSOR, Curs1Handle)
SysCursHandle = SetCursor(Curs1Handle)
        If (Not Ambient.UserMode) Or m_AllowColumnResize Then
        Select Case Colgedrückt
        Case -1
        Case Else
                    i = ColContaining(x)
                    Übergabe = Colgedrückt
                    If i <> Colgedrückt Then
                    Colgedrückt = -1
                    DrawColumnHeader (Übergabe)
                    'UserControl_Paint
                    End If
                    End Select
            For i = 2 To m_Columns
                'falls auf Spaltengrenze
                If m_ColumnHeaders(i - 1).ColumnWidthIsEditable = True Then
                If Abs(x - m_ColumnHeaders(i).ColLeft) <= 2 Then
                
                    'Mauszeiger ändern
                    'UserControl.MousePointer = vbSizeWE
         Select Case m_Mauszeiger
                    Case 0
                        Curs1Handle = LoadCursor(ByVal 0&, IDC_SIZEWE)
                        Case Else
                        Curs1Handle = CreateNewCursor
                        End Select
    'Set the form's mouse cursor
    'SysCursHandle = SetClassWord(UserControl.hWnd, GCW_HCURSOR, Curs1Handle)
    SysCursHandle = SetCursor(Curs1Handle)

                    Exit For
                End If
                End If
            Next i
        End If
    
    Else
    If Colgedrückt <> -1 Then
    Übergabe = Colgedrückt
    Colgedrückt = -1
    DrawColumnHeader (Übergabe)
    'UserControl_Paint
    End If
        'UserControl.MousePointer = vbDefault
                                        Curs1Handle = LoadCursor(ByVal 0&, IDC_ARROW)
    'Set the form's mouse cursor
    'SysCursHandle = SetClassWord(UserControl.hWnd, GCW_HCURSOR, Curs1Handle)
    SysCursHandle = SetCursor(Curs1Handle)

    End If
    DoEvents
    'If active = True Then SetCapture UserControl.hWnd

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
        Dim RectColumn As RECT
Dim Übergabe As Long
    ToLittle = True
    'Event an den Entwickler weiterleiten
    RaiseEvent MouseUp(Button, Shift, x, y)
    
    'falls Spaltenbreitenänderung aktiv ist
    If miResizeCol <> 0 Then
        'neue Breite der geänderten Spalte (links vom Cursor)
        m_ColumnHeaders(miResizeCol).ScaleWidth = miResizePos - m_ColumnHeaders(miResizeCol).ColLeft
        'neue Breite der folgenden Spalte (rechts vom Cursor)
        m_ColumnHeaders(miResizeCol + 1).ColLeft = miResizePos
        If miResizeCol + 1 < m_Columns Then
            m_ColumnHeaders(miResizeCol + 1).ScaleWidth = m_ColumnHeaders(miResizeCol + 2).ColLeft - m_ColumnHeaders(miResizeCol + 1).ColLeft
        Else
            m_ColumnHeaders(miResizeCol + 1).ScaleWidth = UserControl.ScaleWidth - m_ColumnHeaders(miResizeCol + 1).ColLeft
        End If
            PropertyChanged "Columns"

        'Spaltenbreitenänderung beendet
        miResizeCol = 0
        miResizePos = 0
        UserControl_Paint
    Else
        'falls in den Spaltenköpfen
        If y < UserControl.TextHeight("Text") + 4 Then
            'angeklickte Spalte bestimmen
            i = ColContaining(x)
            If i > 0 And i <= m_Columns Then
            If Colgedrückt <> -1 Then
            Übergabe = Colgedrückt
                Colgedrückt = -1
                DrawColumnHeader (Übergabe)
                'UserControl_Paint
                End If
                'Event auslösen
                RaiseEvent ColumnClick(i, Button, Shift)
            End If
        End If
    End If
    
End Sub

'Zeichnen des Controls
Private Sub UserControl_Paint()
angefangen = False
    Dim i As Long
    Dim loRow As CRowOfList
    Dim llTop As Long
    Dim bSelected As Boolean
    Dim DTStyle As Long
    Dim RectLine As RECT
    Dim hBrush As Long
    Dim utBrush As LOGBRUSH
    'Bitmaps zeichnen
    Dim a As StdPicture
    Dim b As Long
    Dim rc As RECT
'Bitmap in Speicher laden
Set a = LoadResPicture(101, vbResBitmap)
b = LoadBitmapIntoMemory(a)


    With UserControl
    .BackColor = m_BackColor
    .ForeColor = m_ForeColor
    .Cls
    
    RectColumn.Top = 0
    RectColumn.Bottom = 4 + .TextHeight("Text")
    For i = 1 To m_Columns
    IconuPic = b
DrawColumnHeader i


        .Font.Bold = False
        .Font.Italic = False
    Next i
                
    llTop = RectColumn.Bottom
    If m_FirstRow > 0 Then
        For i = 1 To mlistRows.Count
            Set loRow = mlistRows.Item(i)
            If i < m_FirstRow Then
                loRow.RowInView = False
                loRow.RowTop = 0
            Else
                If llTop < UserControl.ScaleHeight Then
                    loRow.RowTop = llTop
                    bSelected = (i = m_SelectedRow)
                    DrawRow .hDC, loRow, bSelected, llTop
                    If llTop <= UserControl.ScaleHeight Then
                        loRow.RowInView = True
                    Else
                        loRow.RowInView = False
                    End If
                Else
                    loRow.RowInView = False
                    loRow.RowTop = UserControl.ScaleHeight + 1
                End If
            End If
        Next i
    End If
    
    
    End With
    

'BitBlt(UserControl.hdc, 0, 0, 16, a.Height, b, 16 * 9, 0, SRCCOPY)

End Sub

Private Sub DrawRow(myHDC As Long, objRow As CRowOfList, ByVal bSelected As Boolean, ByRef lTop As Long)
    
    Dim lsColumns() As String
    Dim lsOut As String
    Dim RectColumn As RECT
    Dim rectFocus As RECT
    Dim i As Long
    Dim lBackColor As Long
    Dim lOldBackColor As Long
    Dim lForeColor As Long
    Dim lOldForeColor As Long
    Dim lBackMode As Long
    Dim lbBrush As LOGBRUSH
    Dim hBrush As Long
    Dim DTStyle As Long
    Dim lLines As Long
    Dim rc As RECT
    Dim rec As RECT
    Dim Bild As Long
    Bild = 4
    With UserControl
    RectColumn.Top = lTop
    RectColumn.Bottom = lTop + .TextHeight("Text") + 4
    RectColumn.Left = 0
    RectColumn.Right = UserControl.ScaleWidth
    
    If bSelected And m_HilightSelectedRow Then
        If m_HasFocus Then
            lBackColor = GetSysColor(COLOR_HIGHLIGHT)
            lForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
        ElseIf Not m_HideSelection Then
            lBackColor = GetSysColor(COLOR_BTNFACE)
            lForeColor = GetSysColor(COLOR_BTNTEXT)
        End If
    Else
        If objRow.BackColor < 0 Then
            lBackColor = GetSysColor(objRow.BackColor And &H7FFFFFFF)
        Else
            lBackColor = objRow.BackColor
        End If
        If objRow.ForeColor < 0 Then
            lForeColor = GetSysColor(objRow.ForeColor And &H7FFFFFFF)
        Else
            lForeColor = objRow.ForeColor
        End If
    End If
    
    lbBrush.lbColor = lBackColor
    lbBrush.lbStyle = BS_SOLID
    hBrush = CreateBrushIndirect(lbBrush)
    FillRect myHDC, RectColumn, hBrush
    lOldBackColor = SetBkColor(myHDC, lBackColor)
    lBackMode = SetBkMode(myHDC, OPAQUE)
    lOldForeColor = .ForeColor
    .ForeColor = lForeColor
    
    lsColumns = Split(objRow.ColumnText, "|")
    For i = 1 To m_Columns
        If i - 1 + LBound(lsColumns) <= UBound(lsColumns) Then
            lsOut = lsColumns(LBound(lsColumns) + i - 1)
        Else
            lsOut = ""
        End If
        If lsOut <> "" Then
        rec.Left = 0
        rec.Top = lTop
            RectColumn.Left = m_ColumnHeaders(i).ColLeft + 2 '+ 16
            If i < m_Columns Then
                RectColumn.Right = m_ColumnHeaders(i + 1).ColLeft - 4
            Else
                RectColumn.Right = UserControl.ScaleWidth - 4
            End If
            DTStyle = DT_SINGLELINE + DT_VCENTER + DT_WORD_ELLIPSIS
            Select Case m_ColumnHeaders(i).Alignment
                Case vbRightJustify
                    DTStyle = DTStyle + DT_RIGHT
                Case vbCenter
                    DTStyle = DTStyle + DT_CENTER
                Case Else   'vbLeftJustify
                    DTStyle = DTStyle + DT_LEFT
            End Select
            .Font.Bold = objRow.Bold Or m_ColumnHeaders(i).Bold
            .Font.Italic = objRow.Italic Or m_ColumnHeaders(i).Italic
            DrawText myHDC, lsOut, Len(lsOut), RectColumn, DTStyle
        End If
    Next i
    
    If bSelected And Not m_HilightSelectedRow And (m_HasFocus Or Not m_HideSelection) Then
        rectFocus.Left = 1
        rectFocus.Right = UserControl.ScaleWidth - 1
        rectFocus.Top = RectColumn.Top + 1
        rectFocus.Bottom = RectColumn.Bottom - 1
        DrawFocusRect myHDC, rectFocus
    End If
    
    .Font.Bold = False
    .Font.Italic = False
    
    lTop = RectColumn.Bottom
    If objRow.ShowDetails Then
    Bild = 5
        lLines = CountLines(objRow.DetailText)
        RectColumn.Top = lTop
        RectColumn.Bottom = lTop + 4 + lLines * .TextHeight("Text")
        RectColumn.Left = 0
        RectColumn.Right = UserControl.ScaleWidth
        FillRect myHDC, RectColumn, hBrush
        RectColumn.Left = MINCOLWIDTH
        RectColumn.Right = RectColumn.Right - 4
        DTStyle = DT_LEFT + DT_TOP + DT_EXPANDTABS
        DrawText myHDC, objRow.DetailText, Len(objRow.DetailText), RectColumn, DTStyle
        lTop = RectColumn.Bottom
    End If
    
    SetBkColor myHDC, lOldBackColor
    .ForeColor = lOldForeColor
    SetBkMode myHDC, lBackMode
    DeleteObject hBrush
    End With
    
          With rc
   .Left = Bild * 16 'm_ColumnHeaders(Col).Picture * 16
   .Top = 0
   .Right = (Bild + 1) * 16 '(m_ColumnHeaders(Col).Picture + 1) * 16 'Picture1.ScaleWidth
   .Bottom = 15 'Picture1.ScaleHeight
  End With

 TransparentBlt UserControl.hDC, UserControl.hDC, IconuPic, rc, rec.Left, rec.Top, RGB(255, 0, 255)

End Sub

'UserControl wird gerade erzeugt (noch nicht plaziert)
Private Sub UserControl_Initialize()
    Colgedrückt = -1
    'Liste mit den Zeilendaten erzeugen
    Set mlistRows = New CListObject
    'Scrollbar erzeugen und initialisieren
    ShowScrollBar UserControl.hWnd, SB_VERT, True
    SetScrollRange UserControl.hWnd, SB_VERT, 1, mlistRows.Count + 1, False
    m_FirstRow = 1
    SetScrollPos UserControl.hWnd, SB_VERT, m_FirstRow, True
    EnableScrollBar UserControl.hWnd, SB_VERT, ESB_DISABLE_BOTH
    ShowScrollBar UserControl.hWnd, SB_VERT, True
    'Set Scrollbar = HookWindow(UserControl.hWnd)

End Sub

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    Dim i As Integer
    
    'Eigenschaften initialisieren
    m_ColumnHeaderStil = m_def_ColumnHeaderStil
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    Set UserControl.Font = Ambient.Font
    m_AllowColumnResize = False
    m_HilightSelectedRow = False
    m_HideSelection = False
    m_Columns = 1
    MakeHeaders
    m_FirstRow = 1
        m_Mauszeiger = m_def_Mauszeiger

End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim i As Integer
    Dim llColLeft As Long
    m_Mauszeiger = PropBag.ReadProperty("Mauszeiger", m_def_Mauszeiger)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.Appearance = IIf(PropBag.ReadProperty("Flat", False), 0, 1)
    UserControl.BorderStyle = IIf(PropBag.ReadProperty("Borderless", False), 0, 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_AllowColumnResize = PropBag.ReadProperty("AllowColumnResize", False)
    m_HilightSelectedRow = PropBag.ReadProperty("HilightSelectedRow", False)
    m_HideSelection = PropBag.ReadProperty("HideSelection", False)
    m_Columns = PropBag.ReadProperty("Columns", 1)
    ReDim m_ColumnHeaders(1 To m_Columns) As CColumnHeader
    llColLeft = 0
    For i = 1 To m_Columns
        Set m_ColumnHeaders(i) = New CColumnHeader
        m_ColumnHeaders(i).Caption = PropBag.ReadProperty("CHCaption" & i, "Column" & i)
        m_ColumnHeaders(i).FontName = PropBag.ReadProperty("CHFontName" & i, UserControl.Font.Name)
        m_ColumnHeaders(i).Bold = PropBag.ReadProperty("CHFontBold" & i, UserControl.Font.Bold)
        m_ColumnHeaders(i).FontUnderline = PropBag.ReadProperty("CHFontUnderline" & i, UserControl.Font.Underline)
        m_ColumnHeaders(i).FontStrikeout = PropBag.ReadProperty("CHFontStrikeout" & i, UserControl.Font.Strikethrough)
        m_ColumnHeaders(i).Fontsize = PropBag.ReadProperty("CHFontSize" & i, UserControl.Font.Size)
        m_ColumnHeaders(i).Italic = PropBag.ReadProperty("CHFontItalic" & i, UserControl.Font.Italic)
        m_ColumnHeaders(i).FontColor = PropBag.ReadProperty("CHFontColor" & i, UserControl.ForeColor)
        m_ColumnHeaders(i).Picture = PropBag.ReadProperty("CHPicture" & i, "0")
        m_ColumnHeaders(i).Headerstil = PropBag.ReadProperty("CHStil" & i, "0")
        m_ColumnHeaders(i).Alignment = PropBag.ReadProperty("CHAlign" & i, vbLeftJustify)
        m_ColumnHeaders(i).ScaleWidth = PropBag.ReadProperty("CHWidth" & i, (UserControl.ScaleWidth / m_Columns))
        m_ColumnHeaders(i).ColLeft = llColLeft
        llColLeft = llColLeft + m_ColumnHeaders(i).ScaleWidth
        Set m_ColumnHeaders(i).Parent = Me
    Next i
    
    m_FirstRow = 1

End Sub

Private Sub UserControl_Resize()
    If m_Columns > 0 Then AdjustColumnWidths
    'Paint Ereignis wird nicht unbedingt bei jedem Resizing
    ' ausgelöst, also hier besser manuell auslösen
    UserControl.Refresh
End Sub

Friend Sub AdjustColumnWidths()
    Dim lWidth As Long
    Dim i As Integer
    Dim j As Integer
    Dim llColLeft As Long
    
    'Spaltenbreiten anpassen
    With UserControl
    '1. Alle ColLeft properties neu setzen:
    llColLeft = 0
    For i = 1 To m_Columns
        m_ColumnHeaders(i).ColLeft = llColLeft
        llColLeft = llColLeft + m_ColumnHeaders(i).ScaleWidth
    Next i
    
    With m_ColumnHeaders(m_Columns)
    If .ColLeft + .ScaleWidth < UserControl.ScaleWidth Then
        '2. falls letzte Spalte zu schmal ist, dann diese vergrößern.
        .ScaleWidth = UserControl.ScaleWidth - .ColLeft
        Exit Sub
    ElseIf .ColLeft + .ScaleWidth > UserControl.ScaleWidth Then
        '3. falls letzte Spalte zu breit ist, dann erst diese, evtl. alle Spalten verkürzen...
        If UserControl.ScaleWidth - .ColLeft > MINCOLWIDTH Then
            'nur diese Spalte ändern:
            .ScaleWidth = UserControl.ScaleWidth - .ColLeft
            'fertig:
            Exit Sub
        Else
            'diese Spalte auf den minimalwert setzen
            .ScaleWidth = MINCOLWIDTH
        End If
    End If
    End With
    
    'Ändern der letzten Spalte hat nicht gereicht, also
    'von hinten beginnend alle Spalten verkürzen...
    For i = m_Columns - 1 To 1 Step -1
        'wenn der hinter überragende Teil:
        lWidth = m_ColumnHeaders(m_Columns).ScaleWidth - (.ScaleWidth - m_ColumnHeaders(m_Columns).ColLeft)
        'von der aktuellen Breite abgezogen wird, muß minimalwert übrig bleiben...
        If m_ColumnHeaders(i).ScaleWidth - lWidth > MINCOLWIDTH Then
            'anpassen dieser Spalte reicht:
            m_ColumnHeaders(i).ScaleWidth = m_ColumnHeaders(i).ScaleWidth - lWidth
            'linke Spaltenkante weiterrechnen ....
            llColLeft = m_ColumnHeaders(i).ColLeft + m_ColumnHeaders(i).ScaleWidth
            For j = i + 1 To m_Columns
                m_ColumnHeaders(j).ColLeft = llColLeft
                llColLeft = llColLeft + m_ColumnHeaders(j).ScaleWidth
            Next j
            'und fertig:
            Exit For
        Else
            'anpassen dieser spalte reicht nicht,
            'also zunächst auf minimum festlegen
            m_ColumnHeaders(i).ScaleWidth = MINCOLWIDTH
            'linke Spaltenkante weiterrechnen ....
            llColLeft = m_ColumnHeaders(i).ColLeft + m_ColumnHeaders(i).ScaleWidth
            For j = i + 1 To m_Columns
                m_ColumnHeaders(j).ColLeft = llColLeft
                llColLeft = llColLeft + m_ColumnHeaders(j).ScaleWidth
            Next j
            'und nächste Spalte behandeln
        End If
    Next i
    End With
    
    PropertyChanged "Columns"
End Sub

'Usercontrol wird beendet (letztes Event vor dem Zerstören)
Private Sub UserControl_Terminate()
'Mauszeiger aufräumen
If Curs1Handle <> 0 Then
    'SetClassWord UserControl.hWnd, GCW_HCURSOR, SysCursHandle
    SysCursHandle = SetCursor(SysCursHandle)

    'Clean up
    DestroyCursor Curs1Handle
End If
    'Daten löschen
    Set mlistRows = Nothing
    'Scrollbar deaktivieren
    Set Scrollbar = Nothing
    'If active = True Then
    'Subclassing beenden
    'UnHookWindow UserControl.hWnd
    'active = False
    'End If
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim i As Integer
    PropBag.WriteProperty "Mauszeiger", m_Mauszeiger, m_def_Mauszeiger
    PropBag.WriteProperty "Borderless", CBool(UserControl.BorderStyle = 0), False
    PropBag.WriteProperty "Flat", CBool(UserControl.Appearance = 0), False
    PropBag.WriteProperty "BackColor", m_BackColor, m_def_BackColor
    PropBag.WriteProperty "ForeColor", m_ForeColor, m_def_ForeColor
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True
    PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
    PropBag.WriteProperty "AllowColumnResize", m_AllowColumnResize, False
    PropBag.WriteProperty "HilightSelectedRow", m_HilightSelectedRow, False
    PropBag.WriteProperty "HideSelection", m_HideSelection, False
    PropBag.WriteProperty "Columns", m_Columns, 1
    For i = 1 To m_Columns
        PropBag.WriteProperty "CHCaption" & i, m_ColumnHeaders(i).Caption, "Column" & i
        PropBag.WriteProperty "CHFontName" & i, m_ColumnHeaders(i).FontName, UserControl.Font.Name
        PropBag.WriteProperty "CHFontBold" & i, m_ColumnHeaders(i).Bold, UserControl.Font.Bold
        PropBag.WriteProperty "CHFontUnderline" & i, m_ColumnHeaders(i).FontUnderline, UserControl.Font.Underline
        PropBag.WriteProperty "CHFontStrikeout" & i, m_ColumnHeaders(i).FontStrikeout, UserControl.Font.Strikethrough
        PropBag.WriteProperty "CHFontSize" & i, m_ColumnHeaders(i).Fontsize, UserControl.Font.Size
        PropBag.WriteProperty "CHFontItalic" & i, m_ColumnHeaders(i).Italic, UserControl.Font.Italic
        PropBag.WriteProperty "CHFontColor" & i, m_ColumnHeaders(i).FontColor, UserControl.ForeColor
        PropBag.WriteProperty "CHAlign" & i, m_ColumnHeaders(i).Alignment, vbLeftJustify
        PropBag.WriteProperty "CHWidth" & i, m_ColumnHeaders(i).ScaleWidth, UserControl.ScaleWidth / m_Columns
        PropBag.WriteProperty "CHPicture" & i, m_ColumnHeaders(i).Picture, "0"
        PropBag.WriteProperty "CHStil" & i, m_ColumnHeaders(i).Headerstil, "0"
    Next i
End Sub


Private Sub MakeHeaders()
Dim i As Long
    'Spaltenköpfe erzeugen
    ReDim m_ColumnHeaders(1 To m_Columns) As CColumnHeader
    For i = 1 To m_Columns
        Set m_ColumnHeaders(i) = New CColumnHeader
        m_ColumnHeaders(i).Caption = "Column" & i
        m_ColumnHeaders(i).Alignment = vbLeftJustify
        m_ColumnHeaders(i).Width = (UserControl.Width - 120) / m_Columns
        m_ColumnHeaders(i).Headerstil = BildundText
        Set m_ColumnHeaders(i).Parent = Me
    Next i

End Sub




Public Property Get ColumnCaption(Nummer As Long) As String
ColumnCaption = m_ColumnHeaders(Nummer).Caption
End Property

Public Property Let ColumnCaption(Nummer As Long, ByVal vNewValue As String)
m_ColumnHeaders(Nummer).Caption = vNewValue
PropertyChanged "ColumnCaption"
Me.Refresh
End Property
Public Property Get ColumnFontname(Nummer As Long) As String
If m_ColumnHeaders(Nummer).FontName <> "" Then
ColumnFontname = m_ColumnHeaders(Nummer).FontName
Else
ColumnFontname = UserControl.Font.Name
End If
End Property

Public Property Let ColumnFontname(Nummer As Long, ByVal vNewValue As String)

m_ColumnHeaders(Nummer).FontName = vNewValue
PropertyChanged "ColumnFontName"
Me.Refresh
End Property
Public Property Get ColumnBold(Nummer As Long) As Boolean
If m_ColumnHeaders(Nummer).FontName <> "" Then
ColumnBold = m_ColumnHeaders(Nummer).Bold
Else
ColumnBold = UserControl.Font.Bold
End If
End Property

Public Property Let ColumnBold(Nummer As Long, ByVal vNewValue As Boolean)

m_ColumnHeaders(Nummer).Bold = vNewValue
PropertyChanged "ColumnBold"
Me.Refresh
End Property


Public Property Get ColumnItalic(Nummer As Long) As Boolean
If m_ColumnHeaders(Nummer).FontName <> "" Then
ColumnItalic = m_ColumnHeaders(Nummer).Italic
Else
ColumnItalic = UserControl.Font.Italic
End If
End Property

Public Property Let ColumnItalic(Nummer As Long, ByVal vNewValue As Boolean)

m_ColumnHeaders(Nummer).Italic = vNewValue
PropertyChanged "ColumnItalic"
Me.Refresh
End Property
Public Property Get ColumnUnderline(Nummer As Long) As Boolean
If m_ColumnHeaders(Nummer).FontName <> "" Then
ColumnUnderline = m_ColumnHeaders(Nummer).FontUnderline
Else
ColumnUnderline = UserControl.Font.Underline
End If
End Property

Public Property Let ColumnUnderline(Nummer As Long, ByVal vNewValue As Boolean)

m_ColumnHeaders(Nummer).FontUnderline = vNewValue
PropertyChanged "ColumnUnderline"
Me.Refresh
End Property
Public Property Get ColumnStrikeout(Nummer As Long) As Boolean
If m_ColumnHeaders(Nummer).FontName <> "" Then
ColumnStrikeout = m_ColumnHeaders(Nummer).FontStrikeout
Else
ColumnStrikeout = UserControl.Font.Strikethrough
End If
End Property

Public Property Let ColumnStrikeout(Nummer As Long, ByVal vNewValue As Boolean)

m_ColumnHeaders(Nummer).FontStrikeout = vNewValue
PropertyChanged "ColumnStrikeout"
Me.Refresh
End Property

Public Property Get ColumnFontColor(Nummer As Long) As Long
If m_ColumnHeaders(Nummer).FontName <> "" Then
ColumnFontColor = m_ColumnHeaders(Nummer).FontColor
Else
ColumnFontColor = UserControl.ForeColor
End If
End Property

Public Property Let ColumnFontColor(Nummer As Long, ByVal vNewValue As Long)

m_ColumnHeaders(Nummer).FontColor = vNewValue
PropertyChanged "ColumnFontColor"
Me.Refresh
End Property
Public Property Get ColumnFontSize(Nummer As Long) As Long
If m_ColumnHeaders(Nummer).FontName <> "" Then
ColumnFontSize = m_ColumnHeaders(Nummer).Fontsize
Else
ColumnFontSize = UserControl.Font.Size
End If
End Property

Public Property Let ColumnFontSize(Nummer As Long, ByVal vNewValue As Long)

m_ColumnHeaders(Nummer).Fontsize = vNewValue
PropertyChanged "ColumnFontSize"
Me.Refresh
End Property

Public Property Get ColumnPicture(Nummer As Long) As Long
ColumnPicture = m_ColumnHeaders(Nummer).Picture
End Property

Public Property Let ColumnPicture(Nummer As Long, ByVal vNewValue As Long)
m_ColumnHeaders(Nummer).Picture = vNewValue
PropertyChanged "ColumnPicture"
Me.Refresh

End Property
Public Property Get Columnstil(Nummer As Long) As Stil
Columnstil = m_ColumnHeaders(Nummer).Headerstil
End Property

Public Property Let Columnstil(Nummer As Long, ByVal vNewValue As Stil)
m_ColumnHeaders(Nummer).Headerstil = vNewValue
PropertyChanged "ColumnStil"
Me.Refresh

End Property

Private Sub DrawColumnHeader(Col As Long)
Dim Stiel As Long
    Dim i As Long
    Dim loRow As CRowOfList
    Dim llTop As Long
    Dim bSelected As Boolean
    Dim DTStyle As Long
    Dim RectLine As RECT
    Dim hBrush As Long
    Dim utBrush As LOGBRUSH
    Dim Oldfontname As String
    Dim OldUnderline As Boolean
    Dim OldFontsize As Long
    Dim OldBold As Boolean
    Dim OldItalic As Boolean
    Dim OldColor As Long
    Dim OldStrikeout As Boolean
    Dim WidthFromText As Long
    Dim Abstand As Long
    
    'Bitmaps zeichnen
    Dim b As Long
    Dim rc As RECT
    Dim Übergabe As RECT
        Oldfontname = UserControl.Font.Name
        OldBold = UserControl.Font.Bold
        OldItalic = UserControl.Font.Italic
        OldColor = UserControl.ForeColor
        OldUnderline = UserControl.Font.Underline
        OldFontsize = UserControl.Font.Size
        OldStrikeout = UserControl.Font.Strikethrough
    Select Case Colgedrückt
    Case Col
    Stiel = BDR_SUNKENINNER
    Case Else
    Stiel = BDR_RAISEDINNER
    End Select
    
    Select Case Flat
    Case True
    Stiel = BDR_RAISEDINNER
    End Select
        RectColumn.Left = m_ColumnHeaders(Col).ColLeft
        RectColumn.Right = RectColumn.Left + m_ColumnHeaders(Col).ScaleWidth
        DrawEdge UserControl.hDC, RectColumn, Stiel, BF_RECT + BF_MIDDLE
Übergabe.Bottom = RectColumn.Bottom
Übergabe.Left = RectColumn.Left
Übergabe.Right = RectColumn.Right
Übergabe.Top = RectColumn.Top

With UserControl
        RectColumn.Left = RectColumn.Left + 2
        RectColumn.Right = RectColumn.Right - 4
        DTStyle = DT_SINGLELINE + DT_VCENTER + DT_WORD_ELLIPSIS
        Select Case m_ColumnHeaders(Col).Alignment
            Case vbRightJustify
                DTStyle = DTStyle + DT_RIGHT
            Case vbCenter
                DTStyle = DTStyle + DT_CENTER
            Case Else   'vbLeftJustify
                DTStyle = DTStyle + DT_LEFT
        End Select
        If m_ColumnHeaders(Col).FontName <> "" Then
        .Font.Name = m_ColumnHeaders(Col).FontName
        .Font.Bold = m_ColumnHeaders(Col).Bold
        .Font.Italic = m_ColumnHeaders(Col).Italic
        .ForeColor = m_ColumnHeaders(Col).FontColor
        .Font.Underline = m_ColumnHeaders(Col).FontUnderline
        If Font.Size <> 0 Then
        .Font.Size = m_ColumnHeaders(Col).Fontsize
        End If
        .Font.Strikethrough = m_ColumnHeaders(Col).FontStrikeout
        Else
        'bleibt alles normal
        End If
        RectColumn.Left = RectColumn.Left '+ 16 'wegen Bild
        If m_ColumnHeaders(Col).Headerstil = BildundText Or m_ColumnHeaders(Col).Headerstil = NurText Then
        DrawText .hDC, m_ColumnHeaders(Col).Caption, Len(m_ColumnHeaders(Col).Caption), RectColumn, DTStyle
          WidthFromText = CLng(UserControl.TextWidth(m_ColumnHeaders(Col).Caption))
          End If
          With rc
   .Left = m_ColumnHeaders(Col).Picture * 16
   .Top = 0
   .Right = (m_ColumnHeaders(Col).Picture + 1) * 16 'Picture1.ScaleWidth
   .Bottom = 15 'Picture1.ScaleHeight
  End With
    
      If m_ColumnHeaders(Col).Headerstil = BildundText Then Abstand = 16

  If m_ColumnHeaders(Col).Headerstil = BildundText Or m_ColumnHeaders(Col).Headerstil = NurBild Then

 TransparentBlt UserControl.hDC, UserControl.hDC, IconuPic, rc, RectColumn.Left + WidthFromText + Abstand, RectColumn.Top, RGB(255, 0, 255)
End If
End With
UserControl.Font.Name = Oldfontname
UserControl.Font.Bold = OldBold
UserControl.Font.Italic = OldItalic
UserControl.ForeColor = OldColor
UserControl.Font.Size = OldFontsize
UserControl.Font.Underline = OldUnderline
UserControl.Font.Strikethrough = OldStrikeout
Select Case Flat
Case True
If Colgedrückt = Col Then InvertRect UserControl.hDC, Übergabe
End Select
End Sub

Private Sub drawLines(Modes As Long)
        'die Linien zeigen !!
            Dim RectLine As RECT
            If angefangen = True Then
            'InvertRect UserControl.hdc, Rectold1
            InvertRect UserControl.hDC, Rectold2
            End If
If Modes = 1 Then
        RectLine.Top = 0
        RectLine.Bottom = UserControl.ScaleHeight
        
        RectLine.Left = m_ColumnHeaders(miResizeCol).ColLeft
        RectLine.Right = RectLine.Left + 1
Rectold1.Top = RectLine.Top
Rectold1.Bottom = RectLine.Bottom
Rectold1.Left = RectLine.Left
Rectold1.Right = RectLine.Right
InvertRect UserControl.hDC, RectLine
End If
RectLine.Top = 0
RectLine.Bottom = UserControl.ScaleHeight
RectLine.Left = miResizePos
RectLine.Right = RectLine.Left + 1
InvertRect UserControl.hDC, RectLine
Rectold2.Top = RectLine.Top
Rectold2.Bottom = RectLine.Bottom
Rectold2.Left = RectLine.Left
Rectold2.Right = RectLine.Right
angefangen = True
End Sub
Public Function CreateNewCursor() As Long
Dim MaskX As String
Dim MaskA As String
Dim HotX As Long
Dim HotY As Long
    Dim andbits() As Byte  ' stores the AND mask
    Dim xorbits() As Byte  ' stores the XOR mask
Dim c As Long
    ReDim andbits(0 To 127)
    ReDim xorbits(0 To 127)
Select Case m_Mauszeiger
Case 1
MaskX = Mask1
MaskA = Mask2
HotX = 15
HotY = 5
Case 2
MaskX = Mask1a
MaskA = Mask2a
HotX = 15
HotY = 12
Case 3
MaskX = Mask1b
MaskA = Mask2b
HotX = 15
HotY = 15
End Select
    For c = 0 To 127
        andbits(c) = Val("&H" & Mid(MaskX, 2 * c + 1, 2))
        xorbits(c) = Val("&H" & Mid(MaskA, 2 * c + 1, 2))
    Next c
    CreateNewCursor = CreateCursor(App.hInstance, HotX, HotY, 32, 32, andbits(0), xorbits(0))
End Function
Public Property Let Mauszeiger(ByVal New_Mauszeiger As Zeiger)
    m_Mauszeiger = New_Mauszeiger
    PropertyChanged "Mauszeiger"
End Property
Public Property Get Mauszeiger() As Zeiger
    Mauszeiger = m_Mauszeiger
End Property

Public Property Get ColumnWidth(Nummer As Long) As Long
ColumnWidth = m_ColumnHeaders(Nummer).Width
End Property

Public Property Let ColumnWidth(Nummer As Long, ByVal vNewValue As Long)
m_ColumnHeaders(Nummer).Width = vNewValue
PropertyChanged "ColumnWidth"
Me.Refresh
End Property

Public Property Get ColumnWidthIsEditable(Nummer As Long) As Boolean
ColumnWidthIsEditable = m_ColumnHeaders(Nummer).Width
End Property

Public Property Let ColumnWidthIsEditable(Nummer As Long, ByVal vNewValue As Boolean)
m_ColumnHeaders(Nummer).ColumnWidthIsEditable = vNewValue
PropertyChanged "ColumnWidthIsEditable"
Me.Refresh
End Property

