VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmtest 
   Caption         =   "Form1"
   ClientHeight    =   6384
   ClientLeft      =   1320
   ClientTop       =   1392
   ClientWidth     =   6396
   LinkTopic       =   "Form1"
   ScaleHeight     =   6384
   ScaleWidth      =   6396
   WindowState     =   2  'Maximiert
   Begin RichTextLib.RichTextBox Text1 
      Height          =   5652
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   8772
      _ExtentX        =   15473
      _ExtentY        =   9970
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmtest.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "get"
      Height          =   732
      Left            =   240
      TabIndex        =   0
      Top             =   6000
      Width           =   2412
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   1440
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      Filter          =   "*.tlb; *.dll; *.ocx; *.exe; *.olb|*.tlb; *.dll; *.ocx;*.exe; *.olb"
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Filename As String
Dim Descrstring As String
Dim Testlong As Long
Dim i As Long
Dim z As Long
Dim m As Long
Dim TypelibNr As Long
Dim Stringclass As New cSpeedString

Testlong = 0
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowOpen
Filename = CommonDialog1.Filename
Testlong = TestFile(Filename)
Select Case Testlong
Case 0
Stringclass.Append "No MSFT-Typelib"
GoTo TypeExit
Case 1 'tlb
Testlong = LoadDirect(Filename)
Case 2 'exe...
Testlong = GetTypelibFromMZ(Filename)
End Select
If Testlong > 0 Then
TypelibNr = 0
If Testlong > 1 Then
TypelibNr = InputBox("More than 1 Typelib in this File!!", "Please give Number between 0 and " & CStr(Testlong - 1) & "!")
End If
 
Testlong = OpenTypelibMSFT(TypelibNr, Filename)
If Testlong = 0 Then
Stringclass.Append "No MSFT-Typelib"
GoTo TypeExit
End If
Else 'error
Stringclass.Append "No MSFT-Typelib"
GoTo TypeExit
End If

If TypelibDescription.TypelibFilename = "" Then
Stringclass.Append "Registry-Filename: <could not determine filename>" & vbCrLf
Stringclass.Append "Real Filename: " & Mid(TypelibDescription.Filename, InStrRev(TypelibDescription.Filename, "\") + 1) & vbCrLf
Else
Stringclass.Append "Registry-Filename: " & TypelibDescription.TypelibFilename & vbCrLf
Stringclass.Append "Real Filename: " & Mid(TypelibDescription.Filename, InStrRev(TypelibDescription.Filename, "\") + 1) & vbCrLf
End If
Stringclass.Append "Library: " & TypelibDescription.TypelibRealname & vbCrLf
Stringclass.Append "uuid: " & TypelibDescription.TypelibGUID & vbCrLf
Stringclass.Append "Version: " & TypelibDescription.TypelibVersion & vbCrLf
Stringclass.Append "Helpstring: " & TypelibDescription.Helpstring & vbCrLf
If TypelibDescription.Helpfilename <> "" Then
Stringclass.Append "Helpfile: " & TypelibDescription.Helpfilename & vbCrLf
Stringclass.Append "Helpcontext: " & TypelibDescription.Helpcontext & vbCrLf
End If
For i = 1 To TypelibDescription.NrOfImpLibs
Stringclass.Append "ImportlibGuid: " & Inhalt(0).ImportedLib(i).GUID & vbCrLf
If Inhalt(0).ImportedLib(i).Pfad <> "" Then
Stringclass.Append "ImportlibFileName: " & Inhalt(0).ImportedLib(i).Pfad & vbCrLf
Else
Stringclass.Append "ImportlibFileName: " & Inhalt(0).ImportedLib(i).Libname & vbCrLf
End If
If Inhalt(0).ImportedLib(i).Name <> "" Then
Stringclass.Append "ImportlibName: " & Inhalt(0).ImportedLib(i).Name & vbCrLf
End If
Next i
For i = 0 To TypelibDescription.NrOfTypeInfos - 1
Select Case Inhalt(i).InhaltTyp
Case "Type"
Stringclass.Append vbCrLf
If Inhalt(i).GUIDString <> "" Then
Stringclass.Append "uuid: " & Inhalt(i).GUIDString & vbCrLf
End If
If Inhalt(i).Version <> "" Then
Stringclass.Append "Version: " & Inhalt(i).Version & vbCrLf
End If
If Inhalt(i).Helpstring <> "" Then
Stringclass.Append "Helpstring: " & Inhalt(i).Helpstring & vbCrLf
End If
If Inhalt(i).Helpcontext <> "" Then
Stringclass.Append "Helpcontext: " & Inhalt(i).Helpcontext & vbCrLf
End If
Stringclass.Append "Public Type " & Inhalt(i).Infoname & vbCrLf
For z = 0 To Inhalt(i).TypeTable(0).NrArguments - 1
If Inhalt(i).TypeTable(0).Argumente(z).Helpstring <> "" Then
Stringclass.Append "Helpstring : " & Inhalt(i).TypeTable(0).Argumente(z).Helpstring & vbCrLf
End If
Stringclass.Append Inhalt(i).TypeTable(0).Argumente(z).Namen
If Inhalt(i).TypeTable(0).Argumente(z).Array <> "" Then
Stringclass.Append Inhalt(i).TypeTable(0).Argumente(z).Array
End If
Stringclass.Append " as "
Stringclass.Append Inhalt(i).TypeTable(0).Argumente(z).Typ & vbCrLf
Next z
Stringclass.Append "End Type" & vbCrLf
Case "Enum"
Stringclass.Append vbCrLf
If Inhalt(i).GUIDString <> "" Then
Stringclass.Append "uuid: " & Inhalt(i).GUIDString & vbCrLf
End If
If Inhalt(i).Version <> "" Then
Stringclass.Append "Version: " & Inhalt(i).Version & vbCrLf
End If
If Inhalt(i).Helpstring <> "" Then
Stringclass.Append "Helpstring: " & Inhalt(i).Helpstring & vbCrLf
End If
If Inhalt(i).Helpcontext <> "" Then
Stringclass.Append "Helpcontext: " & Inhalt(i).Helpcontext & vbCrLf
End If
Stringclass.Append "Public Enum " & Inhalt(i).Infoname & vbCrLf
For z = 0 To Inhalt(i).TypeTable(0).NrArguments - 1
Stringclass.Append Inhalt(i).TypeTable(0).Argumente(z).Namen & " = "
Stringclass.Append Inhalt(i).TypeTable(0).Argumente(z).Wert & vbCrLf
Next z
Stringclass.Append "End Enum" & vbCrLf
Case Else
Stringclass.Append vbCrLf & "Infoname: " & Inhalt(i).Infoname & vbCrLf
Stringclass.Append "Type: " & Inhalt(i).InhaltTypString & vbCrLf
If Inhalt(i).InhaltTypString = "module" Then
Stringclass.Append "DllName: " & Inhalt(i).InfoDllName & vbCrLf
End If
If Inhalt(i).Version <> "" Then
Stringclass.Append "Version: " & Inhalt(i).Version & vbCrLf
End If
Stringclass.Append "Nummer: " & i & vbCrLf 'Nur zur Hilfe
If Inhalt(i).Attribute.dual = 1 Then Stringclass.Append "dual" & vbCrLf
If Inhalt(i).Attribute.hidden = 1 Then Stringclass.Append "hidden" & vbCrLf
If Inhalt(i).Attribute.nonextensible = 1 Then Stringclass.Append "nonextensible" & vbCrLf
If Inhalt(i).Attribute.oleautomation = 1 Then Stringclass.Append "(oleautomation)" & vbCrLf
If Inhalt(i).Attribute.restricted = 1 Then Stringclass.Append "restricted" & vbCrLf
If Inhalt(i).Attribute.noncreatable_control = 1 Then Stringclass.Append "noncreatable" & vbCrLf & "control" & vbCrLf
If Inhalt(i).Helpstring <> "" Then
Stringclass.Append "Helpstring: " & Inhalt(i).Helpstring & vbCrLf
End If
If Inhalt(i).Helpcontext <> "" Then
Stringclass.Append "Helpcontext: " & Inhalt(i).Helpcontext & vbCrLf
End If
If Inhalt(i).GUIDString <> "" Then
Stringclass.Append "uuid: " & Inhalt(i).GUIDString & vbCrLf
End If
For z = 1 To Inhalt(i).NumberofMethods
If Inhalt(i).MethTable(z - 1).Helpcontext <> "" Then
Stringclass.Append "Helpcontext: " & Inhalt(i).MethTable(z - 1).Helpcontext & vbCrLf
End If
If Inhalt(i).MethTable(z - 1).Helpstring <> "" Then
Stringclass.Append "Helpstring: " & Inhalt(i).MethTable(z - 1).Helpstring & vbCrLf
End If
If Inhalt(i).MethTable(z - 1).ID <> "" Then
Stringclass.Append "(ID: " & Inhalt(i).MethTable(z - 1).ID & ")" & vbCrLf
End If
If Inhalt(i).MethTable(z - 1).IsNotVisible = True Then
Stringclass.Append "hidden" & vbCrLf
End If

If Inhalt(i).MethTable(z - 1).IsStandard = True Then
Stringclass.Append "This is the Standard (Property, Function...)" & vbCrLf
End If
If Inhalt(i).MethTable(z - 1).MethodenTyp = "Variable" Or Inhalt(i).MethTable(z - 1).MethodenTyp = "Constant" Then
If Inhalt(i).MethTable(z - 1).MethodenTyp = "Variable" Then
Stringclass.Append "Public " & Inhalt(i).MethTable(z - 1).MethodenName & " as " & Inhalt(i).MethTable(z - 1).RückgabeTyp & vbCrLf
End If
If Inhalt(i).MethTable(z - 1).MethodenTyp = "Constant" Then
Stringclass.Append "Public Const " & Inhalt(i).MethTable(z - 1).MethodenName & " as " & Inhalt(i).MethTable(z - 1).RückgabeTyp & " = " & StringTest(Inhalt(i).MethTable(z - 1).ConstantValue) & vbCrLf
End If
Else
Stringclass.Append "Public " & Inhalt(i).MethTable(z - 1).MethodenTyp & " " & Inhalt(i).MethTable(z - 1).MethodenName & " "
If Inhalt(i).MethTable(z - 1).MethodenName <> Inhalt(i).MethTable(z - 1).OrginalName And Inhalt(i).MethTable(z - 1).OrginalName <> "" Then
Stringclass.Append " Alias " & Inhalt(i).MethTable(z - 1).OrginalName
End If
If Inhalt(i).MethTable(z - 1).NrArguments > 0 Then
Stringclass.Append " ("
For m = 0 To Inhalt(i).MethTable(z - 1).NrArguments - 1
If m > 0 Then Stringclass.Append ", "
If Inhalt(i).MethTable(z - 1).Argumente(m).IsOptional = 1 Then
Stringclass.Append "Optional "
End If
Stringclass.Append Inhalt(i).MethTable(z - 1).Argumente(m).Namen & " as " & Inhalt(i).MethTable(z - 1).Argumente(m).Typ
If Inhalt(i).MethTable(z - 1).Argumente(m).HasOptionalString = 1 Then
Stringclass.Append " = " & Inhalt(i).MethTable(z - 1).Argumente(m).OptionalString
End If

Next m
Stringclass.Append ")"
If Inhalt(i).MethTable(z - 1).RückgabeTyp <> "" Then Stringclass.Append " as " & Inhalt(i).MethTable(z - 1).RückgabeTyp
Stringclass.Append vbCrLf
Else
Stringclass.Append "()"
If Inhalt(i).MethTable(z - 1).RückgabeTyp <> "" Then Stringclass.Append " as " & Inhalt(i).MethTable(z - 1).RückgabeTyp
Stringclass.Append vbCrLf
End If
End If
Next z
End Select
Next i
Text1.Text = ""

TypeExit:
DoEvents
Text1.Text = Stringclass.Data
ClearStdModule
End Sub

