VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4356
   ClientLeft      =   -276
   ClientTop       =   1296
   ClientWidth     =   9648
   LinkTopic       =   "Form1"
   ScaleHeight     =   4356
   ScaleWidth      =   9648
   WindowState     =   2  'Maximiert
   Begin RichTextLib.RichTextBox Text1 
      Height          =   6132
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   6732
      _ExtentX        =   11875
      _ExtentY        =   10816
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   2520
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      Filter          =   "*.tlb|*.tlb"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2052
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Filename As String
Dim Besch As String
Dim i As Long
Dim z As Long
Dim y As Long
Dim Fast As New cSpeedString
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowOpen
Filename = CommonDialog1.Filename
OpenSLTG Filename
Fast.Append "Filename: " & Descr.TypelibFilename & vbCrLf
Fast.Append "UUID: " & Descr.TypelibGUID & vbCrLf
Fast.Append "Version: " & Descr.TypelibVersion & vbCrLf
If Descr.Helpstring <> "" Then
Fast.Append "Helpstring: " & Descr.Helpstring & vbCrLf
End If
If Descr.Helpfilename <> "" Then
Fast.Append "Helpfilename: " & Descr.Helpfilename & vbCrLf
End If
If Descr.Helpcontext <> "" And Descr.Helpcontext <> 0 Then
Fast.Append "Helpcontext: " & Hex(Descr.Helpcontext) & vbCrLf
End If

Fast.Append "Library: " & Descr.TypelibRealname & vbCrLf
If TypelibDescription.NrOfImpLibs > 0 Then
For z = 1 To TypelibDescription.NrOfImpLibs
Fast.Append "ImportlibGuid: " & Inhalt(0).ImportedLib(z).GUID & vbCrLf
If Inhalt(0).ImportedLib(z).Libname <> "" Then
Fast.Append "ImportlibFileName: " & Inhalt(0).ImportedLib(z).Libname & vbCrLf
Else
Fast.Append "ImportlibFileName: " & Inhalt(0).ImportedLib(z).Pfad & vbCrLf
End If
If Inhalt(0).ImportedLib(z).Name <> "" Then
Fast.Append "ImportlibName: " & Inhalt(0).ImportedLib(z).Name & vbCrLf
End If
Next z
End If
For i = 0 To UBound(Inhalt)
Fast.Append Inhalt(i).InhaltTyp & ": " & Inhalt(i).Infoname & vbCrLf
If Inhalt(i).GUIDString <> "00000000-0000-0000-0000-000000000000" Then
Fast.Append "UUID: " & Inhalt(i).GUIDString & vbCrLf
End If
If Inhalt(i).Version <> "0.0" And Inhalt(i).Version <> "" Then
Fast.Append "Version: " & Inhalt(i).Version & vbCrLf
End If
If Inhalt(i).Helpcontext <> 0 Then
Fast.Append "Helpcontext: " & Hex(Inhalt(i).Helpcontext) & vbCrLf
End If
If Inhalt(i).Attribute.dual = 1 Then Fast.Append "dual" & vbCrLf
If Inhalt(i).Attribute.hidden = 1 Then Fast.Append "hidden" & vbCrLf
If Inhalt(i).Attribute.nonextensible = 1 Then Fast.Append "nonextensible" & vbCrLf
If Inhalt(i).Attribute.oleautomation = 1 Then Fast.Append "(oleautomation)" & vbCrLf
If Inhalt(i).Attribute.restricted = 1 Then Fast.Append "restricted" & vbCrLf
Select Case Inhalt(i).InhaltTyp
Case "Module"
If Inhalt(i).InfoDllName <> "" Then
Fast.Append "DllName: " & Inhalt(i).InfoDllName & vbCrLf
Else
Fast.Append "DllName:  Error in Typelib! No DllName for this module!" & vbCrLf
End If
For z = 0 To Inhalt(i).NumberOfConst - 1
Fast.Append "Const: " & Inhalt(i).ConstTable(z).ConstTyp & " " & Inhalt(i).ConstTable(z).ConstName & " = " & Inhalt(i).ConstTable(z).ConstValue & vbCrLf
Next z
For z = 0 To Inhalt(i).NumberofMethods - 1
Fast.Append "dispid: " & Inhalt(i).MethTable(z).ID & vbCrLf
If Inhalt(i).MethTable(z).Helpcontext <> "" Then
Fast.Append "Helpcontext: " & Inhalt(i).MethTable(z).Helpcontext & vbCrLf
End If
Fast.Append Inhalt(i).MethTable(z).MethodenTyp & ": " & Inhalt(i).MethTable(z).MethodenName
If Inhalt(i).MethTable(z).MethodenName <> Inhalt(i).MethTable(z).OrginalName Then
Fast.Append " Alias " & Inhalt(i).MethTable(z).OrginalName
End If
Fast.Append "("
If Inhalt(i).MethTable(z).NrArguments = 0 Then
Fast.Append ")"
Else
For y = 0 To Inhalt(i).MethTable(z).NrArguments - 1
Select Case Inhalt(i).MethTable(z).Argumente(y).ByValOrByRef
Case 1
Fast.Append "ByVal "
Case 0
Fast.Append "ByRef "
End Select
Fast.Append Inhalt(i).MethTable(z).Argumente(y).Namen & " as " & Inhalt(i).MethTable(z).Argumente(y).Typ
If y < Inhalt(i).MethTable(z).NrArguments - 1 Then Fast.Append ", "
Next y
Fast.Append ")"
End If
Select Case Inhalt(i).MethTable(z).MethodenTyp
Case "Function", "Property Get"
Fast.Append " as " & Inhalt(i).MethTable(z).RückgabeTyp & vbCrLf
Case "Sub", "Property Let", "Property Set"
Fast.Append vbCrLf
End Select
Next z
Case "Type"
For z = 0 To Inhalt(i).TypeTable(0).NrArguments - 1
Fast.Append Inhalt(i).TypeTable(0).Argumente(z).Namen
If Inhalt(i).TypeTable(0).Argumente(z).Array <> "" Then
Fast.Append Inhalt(i).TypeTable(0).Argumente(z).Array
End If
Fast.Append " as " & Inhalt(i).TypeTable(0).Argumente(z).Typ & vbCrLf
Next z
Case "Enum"
Fast.Append vbCrLf
If Inhalt(i).GUIDString <> "" Then
Fast.Append "uuid: " & Inhalt(i).GUIDString & vbCrLf
End If
If Inhalt(i).Version <> "" Then
Fast.Append "Version: " & Inhalt(i).Version & vbCrLf
End If
If Inhalt(i).Helpstring <> "" Then
Fast.Append "Helpstring: " & Inhalt(i).Helpstring & vbCrLf
End If
If Inhalt(i).Helpcontext <> "" Then
Fast.Append "Helpcontext: " & Inhalt(i).Helpcontext & vbCrLf
End If
Fast.Append "Public Enum " & Inhalt(i).Infoname & vbCrLf
For z = 0 To Inhalt(i).TypeTable(0).NrArguments - 1
Fast.Append Inhalt(i).TypeTable(0).Argumente(z).Namen & " = "
Fast.Append Inhalt(i).TypeTable(0).Argumente(z).Wert & vbCrLf
Next z
Fast.Append "End Enum" & vbCrLf
Case "Interface"
Fast.Append "Version: " & Inhalt(i).Version & vbCrLf
For z = 0 To Inhalt(i).NumberofMethods - 1
Fast.Append "dispid: " & Inhalt(i).MethTable(z).ID & vbCrLf
If Inhalt(i).MethTable(z).Helpcontext <> "" Then
Fast.Append "Helpcontext: " & Inhalt(i).MethTable(z).Helpcontext & vbCrLf
End If
Fast.Append Inhalt(i).MethTable(z).MethodenTyp & ": " & Inhalt(i).MethTable(z).MethodenName & "("
If Inhalt(i).MethTable(z).NrArguments = 0 Then
Fast.Append ")"
Else
For y = 0 To Inhalt(i).MethTable(z).NrArguments - 1
Select Case Inhalt(i).MethTable(z).Argumente(y).ByValOrByRef
Case 1
Fast.Append "ByVal "
Case 0
Fast.Append "ByRef "
End Select
Fast.Append Inhalt(i).MethTable(z).Argumente(y).Namen & " as " & Inhalt(i).MethTable(z).Argumente(y).Typ
If y < Inhalt(i).MethTable(z).NrArguments - 1 Then Fast.Append ", "
Next y
Fast.Append ")"
End If
Select Case Inhalt(i).MethTable(z).MethodenTyp
Case "Function", "Property Get"
Fast.Append " as " & Inhalt(i).MethTable(z).RückgabeTyp & vbCrLf
Case "Sub", "Property Let", "Property Set"
Fast.Append vbCrLf
End Select
Next z
End Select
Next i
Text1.Text = Fast.Data
End Sub

