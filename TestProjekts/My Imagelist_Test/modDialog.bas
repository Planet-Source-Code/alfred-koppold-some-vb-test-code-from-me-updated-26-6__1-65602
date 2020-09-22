Attribute VB_Name = "modDialog"
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Private strfileName As OPENFILENAME

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    End Type
    
Public Function GetFilename() As String
Attribute GetFilename.VB_MemberFlags = "40"
    Dim lngReturnValue As Long
    Dim intRest As Integer
    strfileName.lpstrFile = ""
    strfileName.lpstrTitle = "Bild auswählen"
    strfileName.hInstance = App.hInstance
    strfileName.lpstrFile = Chr(0) & Space(259)
    strfileName.nMaxFile = 260
    strfileName.flags = &H4
    strfileName.lStructSize = Len(strfileName)
    strfileName.lpstrFilter = "Alle Bilddateien" & Chr(0) & "*.bmp; *.dib; *.ico; *.cur; *.gif; *.jpg" & Chr(0) & "Bitmaps (*.bmp;*.dib)" & Chr(0) & "*.bmp; *.dib" & Chr(0) & "Symbol-/Cursordateien (*.ico;*.cur)" & Chr(0) & "*.ico; *.cur" & Chr(0) & "GIF-Dateien (*.gif)" & Chr(0) & "*.gif" & Chr(0) & "JPEG-Dateien (*.jpg)" & Chr(0) & "*.jpg" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*" & Chr(0) & Chr(0)
    lngReturnValue = GetOpenFileName(strfileName)
    Select Case lngReturnValue
    Case 1
    GetFilename = Left(strfileName.lpstrFile, InStr(strfileName.lpstrFile, Chr(0)) - 1)
    Case 0
    GetFilename = ""
    End Select
End Function
