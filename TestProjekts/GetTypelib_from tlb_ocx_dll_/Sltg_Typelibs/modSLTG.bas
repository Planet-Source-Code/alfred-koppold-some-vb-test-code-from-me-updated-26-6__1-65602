Attribute VB_Name = "modSLTG"
Option Explicit
Public Descr As TYPELIB_DESCRIPTION
Public Type SLTG_HEADER
Magic(3) As Byte
nrOffFileBlks As Integer
res06 As Integer
res08 As Integer
first_blk As Integer
res0c As Long
res10 As Long
res14 As Long
res18 As Long
res1c As Long
res20 As Long
End Type

Public Type SLTG_BLKENTRY
len As Long
index_string As Integer
next As Integer
End Type

Public Type SLTG_Magic
res00 As Byte
CompObj_magic(7) As Byte
dir_magic(3) As Byte
End Type

Public Type SLTG_INDEX
Strings(10) As Byte
End Type

Public Type SLTG_EnumItem
Magic As Integer
next As Integer
Name As Integer
Value As Integer
res08 As Integer
memid As Long
Helpcontext As Integer
Helpstring As Integer
End Type

Public Type SLTG_TYPEINFOHEADER
Magic As Integer
href_table As Long
res06 As Long
elem_table As Long
res0e As Long
major_version As Integer
minor_version As Integer
res16 As Long
typeflags1 As Byte
typeflags2 As Byte
typeflags3 As Byte
typekind As Byte
res1e As Long
End Type

Public Type SLTG_TypeInfoTail
cFuncs As Integer
cVars As Integer
cImplTypes As Integer
res06 As Integer
res08 As Integer
res0a As Integer
res0c As Integer
res0e As Integer
res10 As Integer
res12 As Integer
tdescalias As Integer
res16 As Integer
res18 As Integer
res1a As Integer
res1c As Integer
res1e As Integer
cbSizeInstance As Integer
cbAlignment As Integer
res24 As Integer
res26 As Integer
cbSizeVft As Integer
res2a As Integer
res2c As Integer
res2e As Integer
res30 As Integer
res32 As Integer
res34 As Integer
End Type

Public Type SLTG_LIBBLK_1
Magic As Integer
res02 As Integer
Name As Integer
End Type

Public Type SLTG_LIBBLK_2
Helpcontext As Long
syskind As Integer
lcid As Integer
res12 As Long
libflags As Integer
maj_vers As Integer
min_vers As Integer
uuid(15) As Byte
End Type

Public Type SLTG_OTHERTYPEINFO_1
small_no As Integer
index_name As String
other_name As String
End Type

Public Type SLTG_OTHERTYPEINFO_2
res1a As Integer
name_offs As Integer
moreBytes As Integer
End Type

Public Type SLTG_OTHERTYPEINFO_3
res20 As Integer
Helpcontext As Long
res26 As Integer
uuid(15) As Byte
End Type

Public Type SLTG_RECORDITEM
Magic As Byte
Typepos As Byte
next As Integer
Name As Integer
byte_offs As Integer
type As Integer
memid As Long
Helpcontext As Integer
Helpstring As Integer
End Type

Public Type SLTG_FUNCTION
Magic As Byte
inv As Byte
next As Integer
Name As Integer
End Type

Public Type SLTG_FUNCTION2
Helpcontext As Integer
Helpstring As Integer
arg_off As Integer
nacc As Byte
retnextopt As Byte
rettype As Integer
vtblpos As Integer
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private filear() As Byte
Private NameTableOffset As Long
Private ImportLibString() As String
Private Importlibanz As Long
Private vtblBegin As Long

Public Function OpenSLTG(Filename As String) As Long
Dim filenr As Long
Dim y As Long
Dim Test As String
Dim Good As Boolean
Dim Stand As Long
Dim EnumItem As SLTG_EnumItem
Dim Testing As Boolean
Dim Teststring As String
Dim Header As SLTG_HEADER
Dim BlkEntry() As SLTG_BLKENTRY
Dim BlockEntryUnsorted() As SLTG_BLKENTRY
Dim Anzahl As Long
Dim MagicStand As Long
Dim Testlong As Long
Dim Magic As SLTG_Magic
Dim Charstrings() As SLTG_INDEX
Dim pad9(8) As Byte
Dim Infoheader As SLTG_TYPEINFOHEADER
Dim Nr As Long
Dim Offset As Long
Dim Ende As Boolean
Dim i As Long
Dim Tail As SLTG_TypeInfoTail
Dim z As Long
Dim AnzahlParams As Byte
Dim Testint As Integer
Dim Zwischenstand As Long
Dim Recorditem As SLTG_RECORDITEM
Dim Functions As SLTG_FUNCTION
Dim Functions2 As SLTG_FUNCTION2
Dim Testbyte As Byte
Dim AndTest As Byte
Dim StandMember As Long
Dim byteArray() As Byte
Dim Typbyte As Byte
Dim Argstand As Long
Dim Grundstand As Long
Dim aStand As Long
Dim ID As Long
Dim HasBack As Boolean
Dim Number As Long
Dim Testint2 As Integer
Dim SortedNumber As Long
Dim Typ As Integer
Dim ArrayAnz As Integer
Dim Testbyte2 As Byte
Dim Size As Integer
Dim Teststand As Long
Dim IDBack As Long
Dim EndLast As Long
Dim Ending As Boolean
Dim Argname As String
ReDim ImportLibString(0)
Dim cbextra As Long
Importlibanz = 0
ClearStdModule
filenr = FreeFile
Open Filename For Binary As filenr
ReDim filear(LOF(filenr) - 1)
Get filenr, 1, filear
CopyMemory ByVal VarPtr(Header), filear(0), 36
Stand = 36
'Test = StrConv(Header.Magic, vbUnicode) 'SLTG#
Anzahl = Header.nrOffFileBlks - 1

ReDim BlkEntry(Anzahl - 1)
ReDim BlockEntryUnsorted(Anzahl - 1)
CopyMemory ByVal VarPtr(BlockEntryUnsorted(0)), filear(Stand), Anzahl * 8
Number = Header.first_blk
Do While Ending = False 'sortieren
BlkEntry(SortedNumber).index_string = BlockEntryUnsorted(Number - 1).index_string
BlkEntry(SortedNumber).len = BlockEntryUnsorted(Number - 1).len
BlkEntry(SortedNumber).next = BlockEntryUnsorted(Number - 1).next
Number = BlkEntry(SortedNumber).next
SortedNumber = SortedNumber + 1
If Number = 0 Or Number = &HFF Then Ending = True
Loop

Stand = Stand + Anzahl * 8
MagicStand = Stand
CopyMemory ByVal VarPtr(Magic), filear(Stand), 13
Stand = Stand + 13
ReDim Charstrings(Anzahl - 2)
CopyMemory ByVal VarPtr(Charstrings(0)), filear(Stand), 11 * (Anzahl - 1)
Test = Charstrings(0).Strings
Stand = Stand + (11 * (Anzahl - 1))
CopyMemory pad9(0), filear(Stand), 9
Stand = Stand + 9
'main (last) Block
Nr = Header.first_blk
Offset = 0
ReDim Aufteilung(Anzahl - 1)
Do While Ende = False
If BlockEntryUnsorted(Nr - 1).next = 0 Then
Ende = True
Exit Do
Else
Offset = Offset + BlockEntryUnsorted(Nr - 1).len
Nr = BlockEntryUnsorted(Nr - 1).next
End If
Loop

GetMainBlock Stand + Offset, Header.nrOffFileBlks

'normal
For i = 0 To Anzahl - 2
CopyMemory ByVal VarPtr(Infoheader.Magic), filear(Stand), 2
CopyMemory ByVal VarPtr(Infoheader.href_table), filear(Stand + 2), 4
CopyMemory ByVal VarPtr(Infoheader.res06), filear(Stand + 6), 4
CopyMemory ByVal VarPtr(Infoheader.elem_table), filear(Stand + 10), 4
CopyMemory ByVal VarPtr(Infoheader.res0e), filear(Stand + 14), 4
CopyMemory ByVal VarPtr(Infoheader.major_version), filear(Stand + 18), 2
CopyMemory ByVal VarPtr(Infoheader.minor_version), filear(Stand + 20), 2
CopyMemory ByVal VarPtr(Infoheader.res16), filear(Stand + 22), 4
CopyMemory ByVal VarPtr(Infoheader.typeflags1), filear(Stand + 26), 1
CopyMemory ByVal VarPtr(Infoheader.typeflags2), filear(Stand + 27), 1
CopyMemory ByVal VarPtr(Infoheader.typeflags3), filear(Stand + 28), 1
CopyMemory ByVal VarPtr(Infoheader.typekind), filear(Stand + 29), 4
CopyMemory ByVal VarPtr(Infoheader.res16), filear(Stand + 33), 4

AndTest = Infoheader.typeflags1
Testbyte = AndTest And 8
If Testbyte = 8 Then 'appobject predeclid
End If
Testbyte = AndTest And 16
If Testbyte = 16 Then 'cancreate
End If
Testbyte = AndTest And 32
If Testbyte = 32 Then 'licensed
End If
Testbyte = AndTest And 64
If Testbyte = 64 Then 'predeclid
End If
Testbyte = AndTest And 128
If Testbyte = 128 Then 'hidden
Inhalt(i).Attribute.hidden = 1
End If
AndTest = Infoheader.typeflags2
Testbyte = AndTest And 2
If Testbyte = 2 Then 'dual
Inhalt(i).Attribute.dual = 1
End If
Testbyte = AndTest And 4
If Testbyte = 4 Then 'nonextensible
Inhalt(i).Attribute.nonextensible = 1
End If
Testbyte = AndTest And 8
If Testbyte = 8 Then 'oleautomation
Inhalt(i).Attribute.oleautomation = 1
End If
Testbyte = AndTest And 16
If Testbyte = 16 Then 'restricted
Inhalt(i).Attribute.restricted = 1
End If

'Stand = Stand + 34
Zwischenstand = Stand
CopyMemory ByVal VarPtr(cbextra), filear(Stand + 6), 4

Select Case cbextra
Case -1, 65535
aStand = Stand + BlkEntry(i).len
Case Else
aStand = Stand + cbextra
End Select

Stand = Stand + BlkEntry(i).len

CopyMemory ByVal VarPtr(Tail), filear(aStand - 54), 54
EndLast = Stand - 54

CopyMemory ByVal VarPtr(IDBack), filear(Stand - 30), 4
Select Case Infoheader.typekind
Case 0
Inhalt(i).InhaltTyp = "Enum"
Inhalt(i).NumberOfTypes = Tail.cVars
ReDim Inhalt(i).TypeTable(Tail.cVars - 1)
ReDim Inhalt(i).TypeTable(0).Argumente(Tail.cVars - 1)
Zwischenstand = Zwischenstand + Infoheader.elem_table 'Memberhead
StandMember = Zwischenstand
Zwischenstand = Zwischenstand + 9 'Beginn Recorditems
aStand = Zwischenstand
Inhalt(i).TypeTable(0).NrArguments = Tail.cVars
For z = 0 To Tail.cVars - 1
CopyMemory ByVal VarPtr(EnumItem), filear(aStand), 18
Inhalt(i).TypeTable(0).Argumente(z).Namen = GetName_SLTG(IntegerToUnsigned(EnumItem.Name) + NameTableOffset)
CopyMemory ByVal VarPtr(Testlong), filear(Zwischenstand + EnumItem.Value), 4
Inhalt(i).TypeTable(0).Argumente(z).Wert = Testlong
If EnumItem.next <> &HFFFF Then
aStand = Zwischenstand + EnumItem.next
End If
Next z
Case 1
Inhalt(i).InhaltTyp = "Type"
Inhalt(i).NumberOfTypes = Tail.cVars
ReDim Inhalt(i).TypeTable(0)
ReDim Inhalt(i).TypeTable(0).Argumente(Tail.cVars - 1)
Zwischenstand = Zwischenstand + Infoheader.elem_table 'Memberhead
StandMember = Zwischenstand
Zwischenstand = Zwischenstand + 9 'Beginn Recorditems
Inhalt(i).TypeTable(0).NrArguments = Tail.cVars
For z = 0 To Tail.cVars - 1
CopyMemory ByVal VarPtr(Recorditem), filear(Zwischenstand), 18
Zwischenstand = Zwischenstand + 18
Inhalt(i).TypeTable(0).Argumente(z).Namen = GetName_SLTG(IntegerToUnsigned(Recorditem.Name) + NameTableOffset)
If Recorditem.Typepos = 2 Then
Inhalt(i).TypeTable(0).Argumente(z).Typ = MakeTypestring(Recorditem.type)
Else
GetTypeEx Recorditem.type + StandMember, Typ, ArrayAnz
Inhalt(i).TypeTable(0).Argumente(z).Array = "(" & ArrayAnz - 1 & ")"
Inhalt(i).TypeTable(0).Argumente(z).Typ = MakeTypestring(Typ)
End If
If z < Tail.cVars - 1 Then 'not last
Testing = False
Do While Testing = False
CopyMemory ByVal VarPtr(Testbyte), filear(Zwischenstand), 1
If Testbyte = 10 Then
Testing = True
Else
Zwischenstand = Zwischenstand + 1
End If
Loop
End If
Next z
Case 2
Inhalt(i).InhaltTyp = "Module"
Zwischenstand = Zwischenstand + Infoheader.elem_table 'Memberhead
Zwischenstand = Zwischenstand + 9 'Beginn Items
ReDim byteArray(1)
CopyMemory byteArray(0), filear(aStand), 2
Teststring = StrConv(byteArray, vbUnicode)
If Teststring = "ME" Then
vtblBegin = aStand + 37 '1. Funktion (0 bei vtblpos , then astand + vtablpos)

CopyMemory ByVal VarPtr(Testint), filear(aStand + 23), 2
Teststring = GetName_SLTG(IntegerToUnsigned(Testint - 1) + NameTableOffset)
Inhalt(i).InfoDllName = Teststring
'If there is no Function then Error (no Dll!!)
End If
aStand = Zwischenstand
Argstand = Zwischenstand
Grundstand = Zwischenstand
If Tail.cFuncs > 0 Then
ReDim Inhalt(i).MethTable(Tail.cFuncs - 1)
Inhalt(i).NumberofMethods = Tail.cFuncs
End If
If Tail.cVars > 0 Then
ReDim Inhalt(i).ConstTable(Tail.cVars - 1)
Inhalt(i).NumberOfConst = Tail.cVars
Inhalt(i).NumberOfVariables = Tail.cVars
End If
Testlong = DecodeModule(Zwischenstand, Tail.cVars, Tail.cFuncs, i, EndLast, Grundstand, IDBack)
Case 3
Inhalt(i).InhaltTyp = "Interface"
Inhalt(i).NumberofMethods = Tail.cFuncs
Inhalt(i).Version = Infoheader.major_version & "." & Infoheader.minor_version
If Tail.cFuncs > 0 Then
ReDim Inhalt(i).MethTable(Tail.cFuncs - 1)
Teststand = Zwischenstand + Infoheader.href_table
GetWithMagic Teststand
Zwischenstand = Zwischenstand + Infoheader.elem_table 'Begin Memberhead
Zwischenstand = Zwischenstand + 9 'Beginn ImpInfo or Function
Grundstand = Zwischenstand
CopyMemory ByVal VarPtr(Testbyte), filear(Zwischenstand), 1
If Testbyte = &H4A Then
CopyMemory ByVal VarPtr(Size), filear(Zwischenstand + 2), 2
Zwischenstand = Zwischenstand + 22 'Importinfo
End If
For z = 0 To Tail.cFuncs - 1
HasBack = False
CopyMemory ByVal VarPtr(Functions), filear(Zwischenstand), 6
CopyMemory ByVal VarPtr(Functions2), filear(Zwischenstand + 10), 12
Testbyte = Functions.inv And &H10
If Testbyte = &H10 Then Inhalt(i).MethTable(z).MethodenTyp = "Sub or Function"
Testbyte = Functions.inv And &H20
If Testbyte = &H20 Then Inhalt(i).MethTable(z).MethodenTyp = "Property Get"
Testbyte = Functions.inv And &H40
If Testbyte = &H40 Then Inhalt(i).MethTable(z).MethodenTyp = "Property Let"
Testbyte = Functions.inv And &H80
If Testbyte = &H80 Then Inhalt(i).MethTable(z).MethodenTyp = "Property Set"
Select Case Functions2.Helpcontext
Case &HFFFE
'No Helpcontext
Case Else

Testlong = Functions2.Helpcontext And 1
If Size = -1 Then Size = EndLast - Grundstand
Size = Size - 4 'abzüglich Long

If Testlong = 1 Then
Testint = Functions2.Helpcontext And 2
Select Case Testint
Case 2
Testlong = (Functions2.Helpcontext - 3) / 4
Testlong = IDBack - Testlong
Case Else
Testlong = (Functions2.Helpcontext - 1) / 4
Testlong = IDBack + Testlong
End Select
Inhalt(i).MethTable(z).Helpcontext = Hex(Testlong)
Else 'Nicht 1
If Functions2.Helpcontext > 0 And Functions2.Helpcontext < Size Then
CopyMemory ByVal VarPtr(ID), filear(Grundstand + Functions2.Helpcontext), 4
Inhalt(i).MethTable(z).Helpcontext = Hex(ID)
Else
Debug.Print "Error ?"
End If
End If
End Select

AnzahlParams = ShiftRight(Functions2.nacc, 3)
Testbyte = Functions2.nacc And 8
If Testbyte = 8 Then 'restricted
End If
Testbyte = Functions2.nacc And 16
If Testbyte = 16 Then 'Function
End If
Testbyte = Functions2.nacc And 32
If Testbyte = 32 Then 'PropGet
End If
Testbyte = Functions2.nacc And 64
If Testbyte = 64 Then 'Propput
End If
Testbyte = Functions2.nacc And 128
If Testbyte = 128 Then 'PropPutRet
End If
If Inhalt(i).MethTable(z).MethodenTyp = "Sub or Function" Then Inhalt(i).MethTable(z).MethodenTyp = "Sub"
Testbyte = Functions2.retnextopt And 128
If Testbyte = 128 Then
'Direct
Inhalt(i).MethTable(z).RückgabeTyp = MakeTypestring(Functions2.rettype)
If Inhalt(i).MethTable(z).RückgabeTyp <> "" Then
If Inhalt(i).MethTable(z).MethodenTyp = "Sub or Function" Or Inhalt(i).MethTable(z).MethodenTyp = "Sub" Then Inhalt(i).MethTable(z).MethodenTyp = "Function"
HasBack = True 'Rückgabe
End If
Else
'Offset
End If

'hier anfangen
Argstand = Grundstand + Functions2.arg_off
If AnzahlParams > 0 Then
ReDim Preserve Inhalt(i).MethTable(z).Argumente(AnzahlParams - 1)
End If
For y = 0 To AnzahlParams - 1
CopyMemory ByVal VarPtr(Testint), filear(Argstand), 2
Select Case Inhalt(i).MethTable(z).MethodenTyp
Case "Sub", "Function", "Sub or Function"
Inhalt(i).MethTable(z).Argumente(y).Namen = GetName_SLTG(IntegerToUnsigned(Testint) + NameTableOffset - 1)
Case Else
Inhalt(i).MethTable(z).Argumente(y).Namen = "vNewValue" 'Name nicht gespeichert
End Select
Good = False
Do While Good = False
Select Case Left(Inhalt(i).MethTable(z).Argumente(y).Namen, 1)
Case Chr(255), Chr(192)
Inhalt(i).MethTable(z).Argumente(y).Namen = Mid(Inhalt(i).MethTable(z).Argumente(y).Namen, 2) 'Why??
Case Else
Good = True
End Select
Loop
CopyMemory ByVal VarPtr(Testbyte), filear(Argstand + 3), 1
CopyMemory ByVal VarPtr(Typbyte), filear(Argstand + 2), 1

AndTest = Testbyte And 2
If AndTest = 2 Then
Inhalt(i).MethTable(z).Argumente(y).ByValOrByRef = 0
Else
Inhalt(i).MethTable(z).Argumente(y).ByValOrByRef = 1
End If
AndTest = Testbyte And 64
If AndTest = 64 Then
If Inhalt(i).MethTable(z).MethodenTyp = "Sub or Function" Or Inhalt(i).MethTable(z).MethodenTyp = "Sub" Then Inhalt(i).MethTable(z).MethodenTyp = "Function"
Inhalt(i).MethTable(z).RückgabeTyp = MakeTypestring(CInt(Typbyte And Not 128))
AnzahlParams = AnzahlParams - 1
Else
Inhalt(i).MethTable(z).Argumente(y).Typ = MakeTypestring(CInt(Typbyte And Not 128))
End If
Argstand = Argstand + 4
Next y

Inhalt(i).MethTable(z).NrArguments = AnzahlParams

CopyMemory ByVal VarPtr(ID), filear(Zwischenstand + 6), 4
Inhalt(i).MethTable(z).ID = Hex(ID)
Zwischenstand = Grundstand + Functions.next
Inhalt(i).MethTable(z).MethodenName = GetName_SLTG(IntegerToUnsigned(Functions.Name) + NameTableOffset)
Next z
End If

Case 4
Inhalt(i).InhaltTyp = "Dispatch"
Case 5
Inhalt(i).InhaltTyp = "Coclass"
Case 6
Inhalt(i).InhaltTyp = "Alias"
Case 7
Inhalt(i).InhaltTyp = "Union"
Case 8
Inhalt(i).InhaltTyp = "Max"
End Select

Next i
Close filenr
MakeImportarray
End Function

Public Function GetMainBlock(Stand As Long, nrOfFileBlks As Integer) As Long
Dim First As SLTG_LIBBLK_1
Dim Sec As SLTG_LIBBLK_2
Dim res06 As String
Dim Helpstring As String
Dim NewStand As Long
Dim Hilfsint As Integer
Dim Hilfslong As Long
Dim i As Long
Dim OthertypeInfo1() As SLTG_OTHERTYPEINFO_1
Dim OthertypeInfo2() As SLTG_OTHERTYPEINFO_2
Dim OthertypeInfo3() As SLTG_OTHERTYPEINFO_3
Dim Name As String


NewStand = Stand
CopyMemory ByVal VarPtr(First), filear(NewStand), 6
NewStand = NewStand + 6
res06 = GetSLTG_Name(NewStand)
Helpstring = GetSLTG_Name(NewStand)
Descr.Helpstring = Helpstring
Descr.Helpfilename = GetSLTG_Name(NewStand)
CopyMemory ByVal VarPtr(Sec.Helpcontext), filear(NewStand), 4
Descr.Helpcontext = Sec.Helpcontext
CopyMemory ByVal VarPtr(Sec.syskind), filear(NewStand + 4), 4
CopyMemory ByVal VarPtr(Sec.res12), filear(NewStand + 8), 4
CopyMemory ByVal VarPtr(Sec.libflags), filear(NewStand + 12), 22
Descr.TypelibGUID = MakeGuidString(Sec.uuid)
Descr.TypelibVersion = Sec.maj_vers & "." & Sec.min_vers
NewStand = NewStand + 34
NewStand = NewStand + &H40 'FF..
If nrOfFileBlks > 2 Then
ReDim OthertypeInfo1(nrOfFileBlks - 3)
ReDim OthertypeInfo2(nrOfFileBlks - 3)
ReDim OthertypeInfo3(nrOfFileBlks - 3)

For i = 0 To nrOfFileBlks - 3
CopyMemory ByVal VarPtr(Hilfsint), filear(NewStand), 2
OthertypeInfo1(i).small_no = Hilfsint
NewStand = NewStand + 2
OthertypeInfo1(i).index_name = GetSLTG_Name(NewStand)
OthertypeInfo1(i).other_name = GetSLTG_Name(NewStand)
CopyMemory ByVal VarPtr(OthertypeInfo2(i)), filear(NewStand), 6
CopyMemory ByVal VarPtr(OthertypeInfo3(i).Helpcontext), filear(NewStand + 6 + OthertypeInfo2(i).moreBytes + 2), 4
CopyMemory ByVal VarPtr(OthertypeInfo3(i).uuid(0)), filear(NewStand + 6 + OthertypeInfo2(i).moreBytes + 8), 16

NewStand = NewStand + 30
NewStand = NewStand + OthertypeInfo2(i).moreBytes
Next i
End If

CopyMemory ByVal VarPtr(Hilfsint), filear(NewStand), 2
NewStand = NewStand + 2
CopyMemory ByVal VarPtr(Hilfslong), filear(NewStand), 4
NameTableOffset = Hilfslong + &H216 + Stand
NewStand = NewStand + 4
Descr.TypelibRealname = GetName_SLTG(First.Name + NameTableOffset)
Descr.TypelibFilename = GetLibNameInRegistry(Descr.TypelibGUID, Descr.TypelibVersion)
ReDim Inhalt(nrOfFileBlks - 3)
For i = 0 To nrOfFileBlks - 3
Inhalt(i).Helpcontext = OthertypeInfo3(i).Helpcontext
Inhalt(i).Infoname = GetName_SLTG(IntegerToUnsigned(OthertypeInfo2(i).name_offs) + NameTableOffset)
Inhalt(i).GUIDString = MakeGuidString(OthertypeInfo3(i).uuid)
Next i
End Function

Public Function GetSLTG_Name(NewStand As Long) As String
Dim Size As Integer
Dim StringArray() As Byte
On Error GoTo es
Dim Str As String
CopyMemory ByVal VarPtr(Size), filear(NewStand), 2
NewStand = NewStand + 2
If Size <> -1 Then
ReDim StringArray(Size - 1)
CopyMemory StringArray(0), filear(NewStand), Size
NewStand = NewStand + Size
Str = StrConv(StringArray, vbUnicode)
GetSLTG_Name = Str
Else
'Nothing
End If
Exit Function
es:
MsgBox "fehler"
End Function

Public Function GetName_SLTG(NameStand As Long) As String

Dim StandBeg As Long
Dim StandReal As Long
Dim Ende As Boolean
Dim LenIs As Long
Dim Namearray() As Byte
Dim Name As String
StandBeg = NameStand + 2
StandReal = StandBeg
Do While Ende = False
If filear(StandReal) = 0 Then
Ende = True
Else
StandReal = StandReal + 1
End If
Loop
LenIs = StandReal - StandBeg
If LenIs > 0 Then
ReDim Namearray(LenIs - 1)
CopyMemory Namearray(0), filear(StandBeg), LenIs
Name = StrConv(Namearray, vbUnicode)
GetName_SLTG = Name
End If
End Function

Public Function GetWithMagic(Stand As Long)
Dim MagicByte As Byte
Dim Neustand As Long
Dim Ending As Boolean
Dim Hilfsint As Integer
Dim Hilfsstand As Long
Dim Size As Integer

Neustand = Stand
Do While Ending = False
CopyMemory ByVal VarPtr(MagicByte), filear(Neustand), 1
Select Case MagicByte
Case &HDF 'Refinfo
Neustand = GetRefInfo(Neustand)
Case &H4A 'Importinfo
CopyMemory ByVal VarPtr(Size), filear(Neustand + 2), 2
Hilfsstand = Neustand
Neustand = Neustand + 22
Case &H4C 'Function
CopyMemory ByVal VarPtr(Hilfsint), filear(Neustand + 2), 2
If Hilfsint <> -1 Then
Neustand = Hilfsstand + Hilfsint
Else
Ending = True
End If
Case Else
Ending = True
End Select
Loop
End Function

Public Function GetRefInfo(Stand As Long) As Long
Dim Number As Long
Dim names As String
Dim OffsetHex As String
Dim HexLen As Long
Dim Externoffset As Long
Dim ExternName As String
Dim Testbyte As Byte
Dim i As Long
Dim Anzahl As Long

CopyMemory ByVal VarPtr(Testbyte), filear(Stand + 1), 1
Select Case Testbyte
Case 0 'Refinfo
CopyMemory ByVal VarPtr(Number), filear(Stand + 68), 4 '8 Bytes per name
Anzahl = Number / 8

For i = 1 To Anzahl
names = GetSLTG_Name(Stand + 79 + Number)
HexLen = InStr(names, "*#")
OffsetHex = Mid(names, 4, HexLen - 4)
ReDim Preserve ImportLibString(Importlibanz)
If OffsetHex = "ffff" Then
ImportLibString(Importlibanz) = names
Else
Externoffset = Val("&H" & OffsetHex)
ExternName = GetName_SLTG(Externoffset + NameTableOffset)
ImportLibString(Importlibanz) = ExternName
Inhalt(0).HasImpLibs = True
'TypelibDescription.NrOfImpLibs = TypelibDescription.NrOfImpLibs + 1
End If
Importlibanz = Importlibanz + 1
Stand = Stand + Len(names) + 2
Next i

GetRefInfo = Stand + 79 + Number
Case Else '(5?)
GetRefInfo = Stand + 10
End Select
End Function

Public Function MakeImportarray()
On Error Resume Next
Dim i As Long
Dim z As Long
Dim a As Long
Dim Gefunden As Boolean
Dim Begin As Long
Dim Ending As Long
Dim Version As String
Dim Name As String
Dim ImpName As String

For i = 0 To Importlibanz - 1
Gefunden = False
    If Left(ImportLibString(i), 3) <> "*\R" Then
        For a = 0 To i
            If ImportLibString(a) = ImportLibString(i) And a <> i Then
            Gefunden = True
            End If
        Next a
        If Gefunden = False Then
        z = z + 1
        ReDim Preserve Inhalt(0).ImportedLib(z)
        Begin = InStr(ImportLibString(i), "{") + 1
        Ending = InStr(ImportLibString(i), "}")
        Inhalt(0).ImportedLib(z).GUID = Mid(ImportLibString(i), Begin, Ending - Begin)
        Begin = InStr(Ending, ImportLibString(i), "#") + 1
        Ending = InStr(Begin, ImportLibString(i), "#")
        Version = Mid(ImportLibString(i), Begin, Ending - Begin)
        Begin = InStr(Ending + 1, ImportLibString(i), "#") + 1
        Ending = InStr(Begin, ImportLibString(i), "#")
        Name = Mid(ImportLibString(i), Begin, Ending - Begin)
        Inhalt(0).ImportedLib(z).Libname = Name
        Name = GetLibNameInRegistry(Inhalt(0).ImportedLib(z).GUID, Version, ImpName)
        Inhalt(0).ImportedLib(z).Name = Name
        Inhalt(0).ImportedLib(z).Libname = ImpName
        Else
        End If
    End If
Next i
TypelibDescription.NrOfImpLibs = z
End Function

Private Function GetTypeEx(Stand As Long, Typ As Integer, ArrayAnz As Integer) As String
Dim Testint As Integer
Dim Anzahl As Integer
Dim Typint As Integer

CopyMemory ByVal VarPtr(Testint), filear(Stand + 9), 2
Select Case Testint
Case &H1C 'VT_Carray
CopyMemory ByVal VarPtr(Anzahl), filear(Stand + 1), 2
CopyMemory ByVal VarPtr(Typint), filear(Stand + 13), 2
Typ = Typint
ArrayAnz = Anzahl
End Select

End Function

Public Function TestFC(Stand As Long) As Long
Dim Bytetest As Long
CopyMemory ByVal VarPtr(Bytetest), filear(Stand), 1
If Bytetest = 10 Then TestFC = 10
End Function

Private Function DebugFunction(Stand As Long, z As Long, i As Long, EndLast As Long, Grundstand As Long, IDBack As Long, NextF As Long) As Long
Dim HasBack As Boolean
Dim Functions As SLTG_FUNCTION
Dim Functions2 As SLTG_FUNCTION2
Dim Testint As Integer
Dim Testbyte As Byte
Dim Testlong As Long
Dim Size As Long
Dim ID As Long
Dim AnzahlParams As Integer
Dim Argstand As Long
Dim y As Long
Dim Good As Boolean
Dim Typbyte As Byte
Dim AndTest As Byte
HasBack = False
Dim Zwischenstand As Long
Zwischenstand = Stand
CopyMemory ByVal VarPtr(Functions), filear(Zwischenstand), 6
CopyMemory ByVal VarPtr(Functions2), filear(Zwischenstand + 10), 12
If Functions.next <> -1 Then
NextF = CLng(Functions.next) + Grundstand
Else
NextF = -1
End If
Inhalt(i).MethTable(z).MethodenName = GetName_SLTG(IntegerToUnsigned(Functions.Name) + NameTableOffset)
CopyMemory ByVal VarPtr(Testbyte), filear(Functions2.vtblpos + vtblBegin + 1), 1
If Testbyte = 255 Then 'why
Inhalt(i).MethTable(z).OrginalName = GetName_SLTG(Functions2.vtblpos + vtblBegin)
Else
Inhalt(i).MethTable(z).OrginalName = GetName_SLTG(Functions2.vtblpos + vtblBegin - 2)
End If
Testbyte = Functions.inv And &H10
If Testbyte = &H10 Then Inhalt(i).MethTable(z).MethodenTyp = "Sub or Function"
Testbyte = Functions.inv And &H20
If Testbyte = &H20 Then Inhalt(i).MethTable(z).MethodenTyp = "Property Get"
Testbyte = Functions.inv And &H40
If Testbyte = &H40 Then Inhalt(i).MethTable(z).MethodenTyp = "Property Let"
Testbyte = Functions.inv And &H80
If Testbyte = &H80 Then Inhalt(i).MethTable(z).MethodenTyp = "Property Set"
Select Case Functions2.Helpcontext
Case &HFFFE
'No Helpcontext
Case Else

Testlong = Functions2.Helpcontext And 1
If Size = -1 Then Size = EndLast - Grundstand
Size = Size - 4 'abzüglich Long

If Testlong = 1 Then
Testint = Functions2.Helpcontext And 2
Select Case Testint
Case 2
Testlong = (Functions2.Helpcontext - 3) / 4
Testlong = IDBack - Testlong
Case Else
Testlong = (Functions2.Helpcontext - 1) / 4
Testlong = IDBack + Testlong
End Select
Inhalt(i).MethTable(z).Helpcontext = Hex(Testlong)
Else 'Nicht 1
If Functions2.Helpcontext > 0 And Functions2.Helpcontext < Size Then
CopyMemory ByVal VarPtr(ID), filear(Grundstand + Functions2.Helpcontext), 4
Inhalt(i).MethTable(z).Helpcontext = Hex(ID)
Else
Debug.Print "Error ?"
End If
End If
End Select

AnzahlParams = ShiftRight(Functions2.nacc, 3)
Testbyte = Functions2.nacc And 8
If Testbyte = 8 Then 'restricted
End If
Testbyte = Functions2.nacc And 16
If Testbyte = 16 Then 'Function
End If
Testbyte = Functions2.nacc And 32
If Testbyte = 32 Then 'PropGet
End If
Testbyte = Functions2.nacc And 64
If Testbyte = 64 Then 'Propput
End If
Testbyte = Functions2.nacc And 128
If Testbyte = 128 Then 'PropPutRet
End If
If Inhalt(i).MethTable(z).MethodenTyp = "Sub or Function" Then Inhalt(i).MethTable(z).MethodenTyp = "Sub"
Testbyte = Functions2.retnextopt And 128
If Testbyte = 128 Then
'Direct
Inhalt(i).MethTable(z).RückgabeTyp = MakeTypestring(Functions2.rettype)
If Inhalt(i).MethTable(z).RückgabeTyp <> "" Then
If Inhalt(i).MethTable(z).MethodenTyp = "Sub or Function" Or Inhalt(i).MethTable(z).MethodenTyp = "Sub" Then Inhalt(i).MethTable(z).MethodenTyp = "Function"
HasBack = True 'Rückgabe
End If
Else
'Offset
End If

'hier anfangen
If Functions2.arg_off <> -1 Then
Argstand = Grundstand + Functions2.arg_off
Else
Argstand = Stand + 24 'Without Args
End If
If AnzahlParams > 0 Then
ReDim Preserve Inhalt(i).MethTable(z).Argumente(AnzahlParams - 1)
End If
For y = 0 To AnzahlParams - 1
CopyMemory ByVal VarPtr(Testint), filear(Argstand), 2
Select Case Inhalt(i).MethTable(z).MethodenTyp
Case "Sub", "Function", "Sub or Function"
Inhalt(i).MethTable(z).Argumente(y).Namen = GetName_SLTG(IntegerToUnsigned(Testint) + NameTableOffset - 1)
Case Else
Inhalt(i).MethTable(z).Argumente(y).Namen = "vNewValue" 'Name nicht gespeichert
End Select
Good = False
Do While Good = False
Select Case Left(Inhalt(i).MethTable(z).Argumente(y).Namen, 1)
Case Chr(255), Chr(192)
Inhalt(i).MethTable(z).Argumente(y).Namen = Mid(Inhalt(i).MethTable(z).Argumente(y).Namen, 2) 'Why??
Case Else
Good = True
End Select
Loop
CopyMemory ByVal VarPtr(Testbyte), filear(Argstand + 3), 1
CopyMemory ByVal VarPtr(Typbyte), filear(Argstand + 2), 1

AndTest = Testbyte And 2
If AndTest = 2 Then
Inhalt(i).MethTable(z).Argumente(y).ByValOrByRef = 0
Else
Inhalt(i).MethTable(z).Argumente(y).ByValOrByRef = 1
End If
AndTest = Testbyte And 64
If AndTest = 64 Then
If Inhalt(i).MethTable(z).MethodenTyp = "Sub or Function" Or Inhalt(i).MethTable(z).MethodenTyp = "Sub" Then Inhalt(i).MethTable(z).MethodenTyp = "Function"
Inhalt(i).MethTable(z).RückgabeTyp = MakeTypestring(CInt(Typbyte And Not 128))
AnzahlParams = AnzahlParams - 1
Else
Inhalt(i).MethTable(z).Argumente(y).Typ = MakeTypestring(CInt(Typbyte And Not 128))
End If
Argstand = Argstand + 4
Next y

Inhalt(i).MethTable(z).NrArguments = AnzahlParams

CopyMemory ByVal VarPtr(ID), filear(Zwischenstand + 6), 4
Inhalt(i).MethTable(z).ID = Hex(ID)
'Hier rein Constanten!!
If Functions.next <> -1 Then
Zwischenstand = Grundstand + Functions.next
Else
'Debug.Print 2
End If
If Argstand <> NextF And NextF <> -1 Then
Debug.Print "Fehler??"
'Stand = NextF
Stand = Argstand
Else
Stand = Argstand
End If
End Function

Private Function DebugConst(Stand As Long, z As Long, i As Long, EndLast As Long, Grundstand As Long, IDBack As Long, NextC As Long)
Dim Recorditem As SLTG_RECORDITEM
Dim Zwischenstand As Long
Dim Testint As Integer
Dim Testint2 As Integer
Dim Testlong As Long
Dim Test As String
Dim byteArray() As Byte

Zwischenstand = Stand

CopyMemory ByVal VarPtr(Recorditem), filear(Zwischenstand), 18
Inhalt(i).ConstTable(z).ConstName = GetName_SLTG(IntegerToUnsigned(Recorditem.Name) + NameTableOffset)

If Recorditem.Typepos = 2 Then
Inhalt(i).ConstTable(z).ConstTyp = MakeTypestring(Recorditem.type)
Else
Select Case Recorditem.type
Case 66
Inhalt(i).ConstTable(z).ConstTyp = "Short"
Testint2 = Recorditem.Typepos And 8
Select Case Testint2
Case 8
Testint = Recorditem.byte_offs
Case Else
CopyMemory ByVal VarPtr(Testint), filear(Recorditem.byte_offs + Grundstand), 2
End Select
Inhalt(i).ConstTable(z).ConstValue = CStr(Testint)

Case 67
Inhalt(i).ConstTable(z).ConstTyp = "Long"
CopyMemory ByVal VarPtr(Testlong), filear(Recorditem.byte_offs + Grundstand), 4
Inhalt(i).ConstTable(z).ConstValue = CStr(Testlong)

Case 86
Inhalt(i).ConstTable(z).ConstTyp = "Integer"
Testint2 = Recorditem.Typepos And 8
Select Case Testint2
Case 8
Testint = Recorditem.byte_offs
Case Else
CopyMemory ByVal VarPtr(Testint), filear(Recorditem.byte_offs + Grundstand), 2
End Select
Inhalt(i).ConstTable(z).ConstValue = CStr(Testint)

Case 94
Inhalt(i).ConstTable(z).ConstTyp = "String"
CopyMemory ByVal VarPtr(Testint), filear(Recorditem.byte_offs + Grundstand), 2
If Testint > 0 Then
ReDim byteArray(Testint - 1)
CopyMemory byteArray(0), filear(Recorditem.byte_offs + Grundstand + 2), Testint
Test = StrConv(byteArray, vbUnicode)
Else
Test = ""
End If
Inhalt(i).ConstTable(z).ConstValue = StringTest(Test)
Case Else
Debug.Print "Unknown: " & Recorditem.type
End Select
End If
If Recorditem.next = -1 Then
NextC = -1
Else
NextC = Grundstand + Recorditem.next
End If
If NextC = -1 Then
Stand = Stand + 22
Else
Stand = NextC
End If
End Function

Private Function DecodeModule(Stand As Long, NumberConst As Integer, NumberFuncs As Integer, Interfacenumber As Long, EndLast As Long, Grundstand As Long, IDBack As Long) As Long
Dim Ende As Boolean
Dim NewStand As Long
Dim Testlong As Long
Dim Typlong As Long
Dim AktFunc As Long
Dim AktConst As Long
Dim Moduletyp As Long
Dim StandAkt As Long
Dim AnzAll As Long
Dim NextF As Long
Dim NextC As Long

NewStand = Stand
AnzAll = NumberConst + NumberFuncs
If AnzAll = 0 Then Ende = True
Do While Ende = False
Typlong = TestFC(NewStand)
If NumberConst > 0 And NumberFuncs > 0 Then
If NumberConst > AktConst And NumberFuncs > AktFunc Then Moduletyp = 2
If NumberConst = AktConst And NumberFuncs > AktFunc Then Moduletyp = 1
If NumberConst > AktConst And NumberFuncs = AktFunc Then Moduletyp = 0
End If
If NumberConst > 0 And NumberFuncs = 0 Then Moduletyp = 0
If NumberConst = 0 And NumberFuncs > 0 Then Moduletyp = 1
If NextF > 0 And Typlong = 0 Then
If NextF <> NewStand Then
NewStand = NextF 'Functions have sometimes more Bytes! why?
End If
End If
Select Case Moduletyp
Case 2
Select Case Typlong
Case 10 'Const
AktConst = AktConst + 1
Testlong = DebugConst(NewStand, AktConst - 1, Interfacenumber, EndLast, Grundstand, IDBack, NextC)
Case 0 'Function
AktFunc = AktFunc + 1
Testlong = DebugFunction(NewStand, AktFunc - 1, Interfacenumber, EndLast, Grundstand, IDBack, NextF)
End Select

Case 0
AktConst = AktConst + 1
Testlong = DebugConst(NewStand, AktConst - 1, Interfacenumber, EndLast, Grundstand, IDBack, NextC)
NewStand = NextC

Case 1
AktFunc = AktFunc + 1
Testlong = DebugFunction(NewStand, AktFunc - 1, Interfacenumber, EndLast, Grundstand, IDBack, NextF)
NewStand = NextF
End Select
If AktConst + AktFunc = AnzAll Then Ende = True
Loop

End Function
