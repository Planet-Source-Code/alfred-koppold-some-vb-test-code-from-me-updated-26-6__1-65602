Attribute VB_Name = "tlbMSFT"
Option Explicit
Private Type MSFT_HEADER
magic1 As Long
magic2 As Long
posguid As Long
lcid As Long
lcid2 As Long
varflags As Long
Version As Integer
versionunder As Integer
Flags As Long
nrtypeinfos As Long
Helpstring As Long
helpstringcontext As Long
Helpcontext As Long
nametablecount As Long
nametablechars As Long
NameOffset As Long
helpfile As Long
CustomDataOffset As Long
res44 As Long
res48 As Long
dispatchpos As Long
res50 As Long
End Type

Public Type MSFT_PSEG
Offset As Long
Length As Long
res08 As Long
res0c As Long
End Type

Private Type MSFT_TYPEINFOBASE
typekind As Integer
typenumber As Integer
memoffset As Long
res2 As Long
res3 As Long
res4 As Long
res5 As Long
cFuncs As Integer
cVars As Integer
res7 As Long
res8 As Long
res9 As Long
resA As Long
posguid As Long
Flags As Long
NameOffset As Long
Version As Integer
versionunder As Integer
docstringoffs As Long
helpstringcontext As Long
Helpcontext As Long
oCustData As Long
cImplTypes As Integer
cbSizeVft As Integer
Size As Long
Datatype1 As Long
Datatype2 As Long
res18 As Long
res19 As Long
End Type

Private Type MSFT_NAMEINTRO
unk00 As Long
unk10 As Long
Namelen As Byte
unk20 As Byte
unk30 As Integer
End Type

Private Type MSFT_IMPINFO
res0 As Long
oImpFile As Long
oGuid As Long
End Type

Private Type IMPORTS
GUID As String
ImportedName As String
ImportedLib As String
Number As Long
End Type

Private Type GUID_INHALT
GUIDString(15) As Byte
unk10 As Long
unk14 As Long
End Type

Public Type IMPORT_FILE
OffsetGUID As Long
res1 As Long
res2 As Long
LenStringSec As Integer
End Type

Private Type TYPELIB_TYPEDESCR
Type1 As Integer
Type2a As Byte
Type2b As Byte
Type3 As Integer
Type4 As Integer
End Type

Public Type PARAMETER_INFO
Datatype As Integer
Datatype1 As Byte
Datatype2 As Byte
oName As Long
Flags As Long
End Type

Public Type TYPELIB_TYPE
Recsize As Integer
Number As Integer
Datatype As Integer
unknown As Integer
unknown2 As Long
unknown3 As Integer
NameOffset As Integer
unknown4 As Long
unknown5 As Long
unknown6 As Long
End Type

Public Type ARRAY_TYPE
Type As Integer
res1 As Integer
res2 As Integer
res3 As Integer
Numbers As Long
End Type

Public Type ARRAY_TYPE2
Type As Integer
res1 As Integer
Number As Long
End Type

Public Type TYPELIB_CODE
Recsize As Integer
Number As Integer
Calling1 As Integer
Calling2 As Integer
Flags As Long
vTableOffset As Integer
res3 As Integer
'fkccic As Long 'enum
EnumWert As Integer
EnumNumber As Integer
nrargs As Integer
nroargs As Integer
End Type

Private NrOfImpLibs As Long
Private OffsetOld As Long
Private OldName As String
Private Optionalstand As Long
Private Import() As IMPORTS
Public TypeArray() As String
Private startsegdir(14) As MSFT_PSEG
Private Typeinfobase() As MSFT_TYPEINFOBASE

Public Function OpenTypelibMSFT(Number As Long, Filename As String) As Long
Dim Tlheader As MSFT_HEADER
Dim Helplong As Long
Dim Stand As Long
Dim Offsetypeinfos() As Long
Dim StandGUID As Long
Dim GUIDString As String
Dim Interfacename As String
Dim Interfacenumber As Long
Dim InfoHelpstring As String
Dim Hilfslong As Long
Dim i As Long
Dim StandToName As Long
'On Error GoTo ErrorExit
On Error Resume Next
ClearModul
CopyMemory ByVal VarPtr(Tlheader), Typelibarray(Number).Bytes(0), 84
CopyMemory ByVal VarPtr(Helplong), Typelibarray(Number).Bytes(84), 4
Stand = 84
If Helplong <> 0 Then
Stand = Stand + 4 'Helpfile?
End If
ReDim Offsetypeinfos(Tlheader.nrtypeinfos - 1)
CopyMemory ByVal VarPtr(Offsetypeinfos(0)), Typelibarray(Number).Bytes(Stand), Tlheader.nrtypeinfos * 4
Stand = Stand + Tlheader.nrtypeinfos * 4
CopyMemory ByVal VarPtr(startsegdir(0)), Typelibarray(Number).Bytes(Stand), 240
Stand = Stand + 240
ReDim Typeinfobase(Tlheader.nrtypeinfos - 1)
CopyMemory ByVal VarPtr(Typeinfobase(0)), Typelibarray(Number).Bytes(Stand), Tlheader.nrtypeinfos * 100
Stand = Stand + Tlheader.nrtypeinfos * 100
TypelibDescription.TypelibRealname = GetMSFTName(startsegdir(7).Offset, Number)
GetImportInfo Number
If Tlheader.helpfile <> -1 Then
TypelibDescription.Helpfilename = GetString(Tlheader.helpfile + startsegdir(8).Offset, Number)
TypelibDescription.Helpcontext = Hex(Tlheader.Helpcontext)
Do While Len(TypelibDescription.Helpcontext) < 8
TypelibDescription.Helpcontext = "0" & TypelibDescription.Helpcontext
Loop
TypelibDescription.Helpcontext = "&H" & TypelibDescription.Helpcontext
End If
TypelibDescription.TypelibVersion = Tlheader.Version & "." & Tlheader.versionunder
If Tlheader.Helpstring <> -1 Then
TypelibDescription.Helpstring = GetString(Tlheader.Helpstring + startsegdir(8).Offset, Number)
End If
'evtl für Types aber vorher'Decodenametable startsegdir(7).Offset + 1, Filenumber, startsegdir(7).length, Interfacenumber, Typeinfobase(i).NameOffset, i
GetType startsegdir(9).Offset, startsegdir(9).Length, Number, startsegdir(7).Offset
Select Case Tlheader.posguid
Case 0, -1
StandGUID = startsegdir(5).Offset
Case Else
StandGUID = startsegdir(5).Offset + Tlheader.posguid
End Select
TypelibDescription.TypelibGUID = GetGuid(StandGUID, Number, -1)

ReDim Preserve Inhalt(Tlheader.nrtypeinfos - 1)
For i = 0 To Tlheader.nrtypeinfos - 1
If Typeinfobase(i).Helpcontext <> 0 Then Inhalt(i).Helpcontext = Hex(Typeinfobase(i).Helpcontext)
Inhalt(i).Infoname = GetMSFTName(startsegdir(7).Offset + Typeinfobase(i).NameOffset, Number, Inhalt(i).Infonumber)
'Attribute
Hilfslong = Typeinfobase(i).Flags And 16
If Hilfslong = 16 Then Inhalt(i).Attribute.hidden = 1
Hilfslong = Typeinfobase(i).Flags And 32
If Hilfslong = 32 Then Inhalt(i).Attribute.noncreatable_control = 1
Hilfslong = Typeinfobase(i).Flags And 64
If Hilfslong = 64 Then Inhalt(i).Attribute.dual = 1
Hilfslong = Typeinfobase(i).Flags And 128
If Hilfslong = 128 Then Inhalt(i).Attribute.nonextensible = 1
Hilfslong = Typeinfobase(i).Flags And 256
If Hilfslong = 256 Then Inhalt(i).Attribute.oleautomation = 1
Hilfslong = Typeinfobase(i).Flags And 512
If Hilfslong = 512 Then Inhalt(i).Attribute.restricted = 1
Inhalt(i).Version = Typeinfobase(i).Version & "." & Typeinfobase(i).versionunder
If Inhalt(i).Version = "0.0" Then Inhalt(i).Version = ""
If Typeinfobase(i).docstringoffs > -1 Then
InfoHelpstring = GetString(Typeinfobase(i).docstringoffs + startsegdir(8).Offset, Number)
Inhalt(i).Helpstring = InfoHelpstring
End If
Inhalt(i).VorgabeNr = Typeinfobase(i).typekind
Inhalt(i).InhaltTypString = GetInfoType(Inhalt(i).VorgabeNr, Typeinfobase(i).Datatype1 + startsegdir(8).Offset, i)
Inhalt(i).TypNumber = Typeinfobase(i).typenumber
Inhalt(i).NumberofMethods = Typeinfobase(i).cFuncs
Inhalt(i).NumberOfVariables = Typeinfobase(i).cVars
If Inhalt(i).NumberofMethods > 0 Then
ReDim Inhalt(i).MethTable(Typeinfobase(i).cFuncs - 1)
End If
If Typeinfobase(i).posguid > 0 Then 'kein -1 = leer
GUIDString = GetGuid(startsegdir(5).Offset + Typeinfobase(i).posguid, Number, i)
End If
If Typeinfobase(i).memoffset < UBound(Typelibarray(Number).Bytes) Then
StandToName = DecodeCode(Typeinfobase(i).memoffset, Number, i, startsegdir(7).Offset, Typeinfobase(i).res2, Typeinfobase(i).Flags, startsegdir(11).Offset, startsegdir(8).Offset, startsegdir(9).Offset, startsegdir(10).Offset)
GetToName StandToName, Number, i, startsegdir(7).Offset
End If
Next i
TypelibDescription.TypelibFilename = GetLibNameInRegistry(TypelibDescription.TypelibGUID, TypelibDescription.TypelibVersion)
TypelibDescription.Filename = Filename
TypelibDescription.NrOfTypeInfos = Tlheader.nrtypeinfos
TypelibDescription.NrOfImpLibs = NrOfImpLibs
OpenTypelibMSFT = 1
Exit Function
ErrorExit:
End Function

Private Function GetType(Offset As Long, TeilLen As Long, Number As Long, Stringtableoffset As Long) As Long
Dim i As Long
Dim z As Long
Dim TypeDescr() As TYPELIB_TYPEDESCR
Dim Testlong As Long
Dim extint As Long
Dim DimAnzahl As Long
Dim stringname As String

On Error Resume Next

If TeilLen > 0 Then
DimAnzahl = TeilLen \ 8
ReDim TypeDescr(DimAnzahl - 1)
CopyMemory ByVal VarPtr(TypeDescr(0)), Typelibarray(Number).Bytes(Offset), 8 * DimAnzahl
ReDim TypeArray(DimAnzahl - 1)
Do While i < DimAnzahl
Select Case TypeDescr(i).Type1
Case &H1A, &H1B, &H1C
TypeArray(i) = MakeTypestring(TypeDescr(i).Type3) 'only for TypeTable
i = i + 1
Case &H1D
CopyMemory ByVal VarPtr(Testlong), ByVal VarPtr(TypeDescr(i).Type3), 2
CopyMemory ByVal VarPtr(Testlong) + 2, ByVal VarPtr(TypeDescr(i).Type4), 2
extint = Testlong And 1
Select Case extint
Case 1 'Extern 24 Bytes
For z = 0 To UBound(Import)
If Import(z).Number = Testlong Then
Exit For
End If
Next z
Select Case TypeDescr(i).Type2a
Case &HFF
TypeArray(i + 1) = Import(z).ImportedName
TypeArray(i) = Import(z).ImportedName
i = i + 2
Case Else
i = i + 1
Select Case Import(z).ImportedName
Case ""
TypeArray(i) = "Extern Type" 'Import(z).ImportedName
Case Else
TypeArray(i) = Import(z).ImportedName
End Select
i = i + 1
End Select
Case 0 'Intern 16/24 Bytes
Select Case Typeinfobase(Testlong / 100).typekind
Case 8481, 4385, 4257, 2145 'Type...
TypeArray(i + 1) = GetMSFTName(Stringtableoffset + Typeinfobase(Testlong / 100).NameOffset, Number)
i = i + 2
Case Else 'Object...
stringname = GetMSFTName(Stringtableoffset + Typeinfobase(Testlong / 100).NameOffset, Number)
If Left(stringname, 1) = "_" Then stringname = Mid(stringname, 2)
TypeArray(i + 2) = stringname
i = i + 3
End Select
Case Else
Debug.Print "Error in  GetType"
End Select
Case Else
i = i + 1
Debug.Print Hex(TypeDescr(i).Type1)
End Select
Loop
End If
End Function

Private Function GetTypeOld(Offset As Long, TeilLen As Long, Number As Long, Stringtableoffset As Long) As Long
Dim i As Long
Dim z As Long
Dim TypeDescr() As TYPELIB_TYPEDESCR
Dim Testlong As Long
Dim extint As Long
Dim DimAnzahl As Long
Dim stringname As String

On Error Resume Next

If TeilLen > 0 Then
DimAnzahl = TeilLen \ 8
ReDim TypeDescr(DimAnzahl - 1)
CopyMemory ByVal VarPtr(TypeDescr(0)), Typelibarray(Number).Bytes(Offset), 8 * DimAnzahl
ReDim TypeArray(DimAnzahl - 1)
Do While i < DimAnzahl
Select Case TypeDescr(i).Type2a
Case 0
TypeArray(i + 1) = "Unknown Type"
i = i + 2
Case &HFF
CopyMemory ByVal VarPtr(Testlong), ByVal VarPtr(TypeDescr(i).Type3), 2
CopyMemory ByVal VarPtr(Testlong) + 2, ByVal VarPtr(TypeDescr(i).Type4), 2
extint = Testlong And 1
Select Case extint
Case 1 'Extern 24 Bytes
For z = 0 To UBound(Import)
If Import(z).Number = Testlong Then
TypeArray(i + 2) = Import(z).ImportedName
Exit For
End If
Next z
i = i + 3
Case 0 'Intern 16/24 Bytes
Select Case Typeinfobase(Testlong / 100).typekind
Case 8481, 4385, 4257, 2145 'Type...
TypeArray(i + 1) = GetMSFTName(Stringtableoffset + Typeinfobase(Testlong / 100).NameOffset, Number)
i = i + 2
Case Else 'Object...
stringname = GetMSFTName(Stringtableoffset + Typeinfobase(Testlong / 100).NameOffset, Number)
If Left(stringname, 1) = "_" Then stringname = Mid(stringname, 2)
TypeArray(i + 2) = stringname
End Select
i = i + 3
Case Else
Debug.Print "Error in  GetType"
End Select
Case Else
TypeArray(i) = MakeTypestring(TypeDescr(i).Type3) 'only for TypeTable
i = i + 1
End Select
Loop
End If
End Function

Private Function GetString(Offs As Long, Number As Long) As String
Dim Namelen As Integer
Dim NameArray() As Byte

CopyMemory ByVal VarPtr(Namelen), Typelibarray(Number).Bytes(Offs), 2
ReDim NameArray(Namelen - 1)
CopyMemory NameArray(0), Typelibarray(Number).Bytes(Offs + 2), Namelen
GetString = StrConv(NameArray, vbUnicode)
End Function

Private Sub GetImportInfo(Number As Long)
Dim ImpInfo As MSFT_IMPINFO
Dim GUIDString As String
Dim GuidName As String
Dim Filename As String
Dim Anz As Long
Dim i As Long
Dim Stand As Long
Dim Guidnumber As Long
Dim LenImpTable As Long

LenImpTable = startsegdir(1).Length
If LenImpTable > 0 Then
Stand = startsegdir(1).Offset
Anz = LenImpTable / 12
ReDim Import(Anz - 1)
For i = 0 To Anz - 1
CopyMemory ByVal VarPtr(ImpInfo), Typelibarray(Number).Bytes(Stand), 12
Stand = Stand + 12
GUIDString = GetGuid(startsegdir(5).Offset + ImpInfo.oGuid, Number, -1, Guidnumber)
Filename = AddImportFile(startsegdir(2).Offset + ImpInfo.oImpFile, startsegdir(5).Offset, Number)
GuidName = NameFromGUID(GUIDString, "Interface")
If GuidName = "" Then
'GuidName =
'Get Extern Name
'from Filename (dll, tlb... then Name from the GUID (GUIDString)
End If
Import(i).GUID = GUIDString
Import(i).ImportedLib = Filename
Import(i).ImportedName = GuidName
Import(i).Number = Guidnumber '1, 13, 25...
Next i
End If
End Sub

Public Function AddImportFile(FileOffset As Long, Guidoffset As Long, Number As Long) As String
Dim Guids As String
Dim Name As String
Dim Name1 As String
Dim Namelen As Long
Dim ImportFile As IMPORT_FILE
Dim NameArray() As Byte
Dim ImpName As String
CopyMemory ByVal VarPtr(ImportFile), Typelibarray(Number).Bytes(FileOffset), 14
If ImportFile.OffsetGUID <> OffsetOld Then
OffsetOld = ImportFile.OffsetGUID
Guids = GetGuid(ImportFile.OffsetGUID + Guidoffset, Number, -1)
Name = GetLibNameInRegistry(Guids, , ImpName)

Namelen = (ImportFile.LenStringSec - 1)
If Namelen Mod 4 = 0 Then
Namelen = Namelen / 4
ReDim NameArray(Namelen - 1)
CopyMemory NameArray(0), Typelibarray(Number).Bytes(FileOffset + 14), Namelen
Name1 = StrConv(NameArray, vbUnicode)
End If
If Name = "" Then Name = "Lib not found"
NrOfImpLibs = NrOfImpLibs + 1
ReDim Preserve Inhalt(0).ImportedLib(NrOfImpLibs)
Inhalt(0).ImportedLib(NrOfImpLibs).GUID = Guids
Inhalt(0).ImportedLib(NrOfImpLibs).Pfad = Name
Inhalt(0).ImportedLib(NrOfImpLibs).Libname = Name1
Inhalt(0).ImportedLib(NrOfImpLibs).Name = ImpName
AddImportFile = Name
OldName = Name
Else
AddImportFile = OldName
End If
End Function

Private Function GetGuid(Stand As Long, Number As Long, Nummer As Long, Optional Guidnumber As Long) As String
Dim GuidInhalt As GUID_INHALT
Dim GUIDString As String

CopyMemory ByVal VarPtr(GuidInhalt), Typelibarray(Number).Bytes(Stand), 24
Guidnumber = GuidInhalt.unk10
GUIDString = MakeGuidString(GuidInhalt.GUIDString)
If Nummer <> -1 Then
Inhalt(Nummer).GUIDString = GUIDString
End If
GetGuid = GUIDString
End Function

Private Function GetMSFTName(Startpos As Long, Number As Long, Optional NameNumber As Long) As String
Dim NameIntro As MSFT_NAMEINTRO
Dim NameByte() As Byte
Dim Name As String
If Startpos < 0 Or Startpos > UBound(Typelibarray(Number).Bytes) Then
Debug.Print "Error in GetMSFTName!!"
Exit Function
End If
CopyMemory ByVal VarPtr(NameIntro), Typelibarray(Number).Bytes(Startpos), 12
ReDim NameByte(NameIntro.Namelen - 1)
CopyMemory NameByte(0), Typelibarray(Number).Bytes(Startpos + 12), NameIntro.Namelen
Name = StrConv(NameByte, vbUnicode)
If InStr(Name, Chr(0)) = 0 Then
GetMSFTName = Name
NameNumber = NameIntro.unk00
Else
Debug.Print "Error in GetMSFTName!!"
End If
End Function


Private Function NameFromGUID(GUIDString As String, RegName As String) As String
Dim Back As String
Dim readystring As String
readystring = "{" & GUIDString & "}"
Back = RegGetKeyValue(&H80000000, RegName & "\" & readystring, "")
If Back <> "" Then
Back = Left(Back, InStr(Back, Chr(0)) - 1)
NameFromGUID = Back
End If
End Function


Public Sub ClearModul()
Optionalstand = 0
NrOfImpLibs = 0
OffsetOld = 0
ReDim Inhalt(0) 'Begin
ReDim TypeArray(0)
ReDim Import(0)
OldName = ""
ReDim Typeinfobase(0)
End Sub

Public Function DecodeCode(CodeOffset As Long, Number As Long, Nummer As Long, OffsetNameTable As Long, AnzElements As Long, Flags As Long, CustdataOffset As Long, Stringoffset As Long, Typeoffset As Long, ArrayDescrOffset As Long) As Long
Dim LongArray() As Long
Dim Intarray() As Integer
Dim Testlong As Long
Dim Testlong1 As Long
Dim Testlong2 As Long
Dim Testint As Integer
Dim Helplong As Long
Dim Codeart As Long
Dim Testname As String
Dim twoLoop As Long
Dim GrLong As Long
Dim Firstlong() As Long
Dim LastLong() As Long
Dim AnzahlArgumente As Long
Dim ArgNameOffset() As Long
Dim ArgumententypArray() As PARAMETER_INFO
Dim Art As String
Dim Typear As TYPELIB_TYPE
Dim ArrayType As ARRAY_TYPE
Dim ArrayType2 As ARRAY_TYPE2
Dim Code As TYPELIB_CODE
Dim GesLen As Long
Dim Test As Long
Dim EndOffset As Long
Dim i As Long
Dim Firststdcall(1) As Long
Dim Testdouble As Double
Dim StandNummer As Long
Dim TestNumMethods As Long
Dim Testlen As Integer
Dim allsub As Boolean
Dim Stand As Long
Dim more As Long
Dim Testoff As Long
Dim Typestand As Long
Dim Methodenstand As Long
Dim Helpstring As String
Dim StandBegin As Long
Dim Bytear() As Byte
Dim IsConst As Long
Dim Helpint As Integer
Dim OptAuf As Long
Dim OptHelpAuf As Long
CopyMemory ByVal VarPtr(GesLen), Typelibarray(Number).Bytes(CodeOffset), 4
EndOffset = CodeOffset + GesLen + 4

DecodeCode = EndOffset
Stand = CodeOffset + 4
StandBegin = Stand
If Inhalt(Nummer).NumberOfVariables > 0 Or Inhalt(Nummer).NumberofMethods > 0 Or Inhalt(Nummer).NumberOfTypes > 0 Then
Do While Stand < EndOffset
CopyMemory ByVal VarPtr(Code), Typelibarray(Number).Bytes(Stand), 24
Stand = Stand + Code.Recsize
Select Case Code.Calling1
Case 24, 25
Case Else
allsub = True 'only subs and funcs (no events)
End Select
TestNumMethods = TestNumMethods + 1 'For Enums..(+ NumVariables)
Loop
End If

If Inhalt(Nummer).NumberOfVariables > 0 Then
Stand = CodeOffset + 4

If TestNumMethods <> Inhalt(Nummer).NumberofMethods Then
ReDim Preserve Inhalt(Nummer).MethTable(TestNumMethods - 1)
Inhalt(Nummer).NumberofMethods = TestNumMethods
End If

End If
Stand = CodeOffset + 4
StandBegin = Stand

Select Case Inhalt(Nummer).VorgabeNr
Case 8481, 4385, 4257, 2145, 8737 'Type hier Fehler??? nur 20
AnzahlArgumente = TestAnz(Number, CodeOffset, EndOffset)
Inhalt(Nummer).InhaltTyp = "Type"
ReDim Preserve Inhalt(Nummer).TypeTable(0)
ReDim Preserve Inhalt(Nummer).TypeTable(0).Argumente(AnzahlArgumente - 1)
Inhalt(Nummer).TypeTable(0).NrArguments = AnzahlArgumente
i = 0
Do While Stand < EndOffset
CopyMemory ByVal VarPtr(Code), Typelibarray(Number).Bytes(Stand), 24

Testlong = IntegerToUnsigned(Code.Calling2)
Testlong = Testlong And 32768
If Testlong = 32768 Then 'normal
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Typ = MakeTypestring(Code.Calling1)
Else 'Array oder sonstiges
ReDim Intarray(3)
CopyMemory ByVal VarPtr(Intarray(0)), Typelibarray(Number).Bytes(Typeoffset + Code.Calling1), 8
Select Case Intarray(0)
Case 27
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Array = "()"
Select Case Intarray(3)
Case 0
CopyMemory ByVal VarPtr(ArrayType2), Typelibarray(Number).Bytes(Typeoffset + Code.Calling1), 8
CopyMemory ByVal VarPtr(Testlong1), Typelibarray(Number).Bytes(Typeoffset + ArrayType2.Number + 4), 4
For twoLoop = 0 To UBound(Inhalt)
If Testlong1 = Inhalt(twoLoop).Infonumber Then
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Typ = Inhalt(twoLoop).Infoname
Exit For
End If
Next twoLoop
Case Else
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Typ = MakeTypestring(Intarray(2))
End Select
Case 28
Testlong = Intarray(2)
CopyMemory ByVal VarPtr(ArrayType), Typelibarray(Number).Bytes(ArrayDescrOffset + Testlong), 12 ' Achtung!! Neues Array machen
Testlong2 = IntegerToUnsigned(ArrayType.res1)
Testlong2 = Testlong2 And 32768
If Testlong2 <> 32768 Then Testlong2 = 0
Select Case Testlong2 'Intarray(3)'Fehler!!!!!!
Case 0
CopyMemory ByVal VarPtr(ArrayType2), Typelibarray(Number).Bytes(Typeoffset + Code.Calling1), 8
CopyMemory ByVal VarPtr(ArrayType2), Typelibarray(Number).Bytes(Typeoffset + ArrayType2.Number), 8
Testlong1 = ArrayType2.Number
If ArrayType2.Type = 28 Then
CopyMemory ByVal VarPtr(Testlong1), Typelibarray(Number).Bytes(Typeoffset + Testlong1 + 4), 4
End If
For twoLoop = 0 To UBound(Inhalt)
If Testlong1 = Inhalt(twoLoop).Infonumber Then
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Typ = Inhalt(twoLoop).Infoname
Exit For
End If
Next twoLoop
Case Else
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Typ = MakeTypestring(ArrayType.Type)
End Select
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Array = "(" & ArrayType.Numbers - 1 & ")"
Case 29
CopyMemory ByVal VarPtr(Testlong), Typelibarray(Number).Bytes(Typeoffset + Code.Calling1 + 4), 4
For twoLoop = 0 To UBound(Inhalt)
If Testlong = Import(twoLoop).Number Then
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Typ = Import(twoLoop).ImportedName
Exit For
End If
Next twoLoop
Case 26
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Typ = TypeArray(Code.Calling1 / 8) 'Import(Code.Calling1 / 16 + 1).ImportedName
Case Else
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Typ = "Any"
End Select
End If
Helplong = Stand
If Code.Recsize > 20 Then
GetOptional Number, Stand + 20, Stringoffset, Code.Recsize + 4, Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Helpcontext, Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Helpstring, Inhalt(Nummer).TypeTable(Typestand).Argumente(i).ConstantValue, 0
End If

Stand = Stand + Code.Recsize
i = i + 1
Loop
'Stand = Helplong
If AnzahlArgumente > 1 Then
Select Case Inhalt(Nummer).VorgabeNr
Case 4385, 8481 '????
'Stand = Stand + Code.Recsize
If Code.Recsize = 20 Then
ReDim Firstlong(AnzahlArgumente - 1)
Else
ReDim Firstlong(AnzahlArgumente - 2)
End If
Case Else
ReDim Firstlong(AnzahlArgumente - 2)
End Select
Else
If Inhalt(Nummer).VorgabeNr = 4385 And AnzahlArgumente = 1 Then ReDim Firstlong(0)
End If
ReDim ArgNameOffset(AnzahlArgumente - 1)
ReDim LastLong(AnzahlArgumente - 1)
If Code.Recsize = 28 Or Code.Recsize = 20 Then
CopyMemory ByVal VarPtr(GrLong), Typelibarray(Number).Bytes(Stand), 4
Stand = Stand + 4
End If
If AnzahlArgumente > 1 Then
CopyMemory ByVal VarPtr(Firstlong(0)), Typelibarray(Number).Bytes(Stand), (UBound(Firstlong) + 1) * 4
Stand = Stand + (UBound(Firstlong) + 1) * 4
End If
If Inhalt(Nummer).VorgabeNr = 4385 And AnzahlArgumente = 1 And Code.Recsize = 20 Then
CopyMemory ByVal VarPtr(Firstlong(0)), Typelibarray(Number).Bytes(Stand), (UBound(Firstlong) + 1) * 4
Stand = Stand + (UBound(Firstlong) + 1) * 4
End If
CopyMemory ByVal VarPtr(ArgNameOffset(0)), Typelibarray(Number).Bytes(Stand), AnzahlArgumente * 4
Stand = Stand + AnzahlArgumente * 4
CopyMemory ByVal VarPtr(LastLong(0)), Typelibarray(Number).Bytes(Stand), AnzahlArgumente * 4
Stand = Stand + AnzahlArgumente * 4
For i = 0 To AnzahlArgumente - 1
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Namen = GetMSFTName(ArgNameOffset(i) + OffsetNameTable, Number)
Next i
Case 8480, 8736 'Enum
AnzahlArgumente = TestAnz(Number, CodeOffset, EndOffset)
Inhalt(Nummer).InhaltTyp = "Enum"
ReDim Preserve Inhalt(Nummer).TypeTable(0)
ReDim Preserve Inhalt(Nummer).TypeTable(0).Argumente(AnzahlArgumente - 1)
Inhalt(Nummer).TypeTable(0).NrArguments = AnzahlArgumente
i = 0
Do While Stand < EndOffset
CopyMemory ByVal VarPtr(Code), Typelibarray(Number).Bytes(Stand), 24
If Code.EnumNumber <> 0 Then '= -29696 Then '(Hex 008C)
CopyMemory ByVal VarPtr(Testlong), Typelibarray(Number).Bytes(Stand + 16), 4
Testdouble = LongToUnsigned(Testlong)
If Testdouble >= 2348810240# Then
Testdouble = Testdouble - 2348810240#
Else
Testdouble = 0 'Error
End If
If Testdouble <= 2147483647 Then
Testlong = CLng(Testdouble)
Else
Testlong = 0 'Error
End If
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Wert = Testlong
Else
CopyMemory ByVal VarPtr(Testlong), Typelibarray(Number).Bytes(CustdataOffset + Code.EnumWert + 2), 4 'Integer Nummer, Long Wert, Integer Ende
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Wert = Testlong
End If
Stand = Stand + Code.Recsize
i = i + 1
Loop
If AnzahlArgumente > 1 Then
ReDim Firstlong(AnzahlArgumente - 2)
End If
ReDim ArgNameOffset(AnzahlArgumente - 1)
ReDim LastLong(AnzahlArgumente - 1)
'If Code.RecSize = 28 Or Code.RecSize = 20 Then 'unnötig aber Größe unterschiedlich!!
CopyMemory ByVal VarPtr(GrLong), Typelibarray(Number).Bytes(Stand), 4
Stand = Stand + 4
'End If
If AnzahlArgumente > 1 Then
CopyMemory ByVal VarPtr(Firstlong(0)), Typelibarray(Number).Bytes(Stand), 4 * (UBound(Firstlong) + 1)
Stand = Stand + ((UBound(Firstlong) + 1) * 4)
End If
CopyMemory ByVal VarPtr(ArgNameOffset(0)), Typelibarray(Number).Bytes(Stand), 4 * AnzahlArgumente
Stand = Stand + (4 * AnzahlArgumente)
CopyMemory ByVal VarPtr(LastLong(0)), Typelibarray(Number).Bytes(Stand), 4 * AnzahlArgumente
Stand = Stand + (4 * AnzahlArgumente)
For i = 0 To AnzahlArgumente - 1
Inhalt(Nummer).TypeTable(Typestand).Argumente(i).Namen = GetMSFTName(ArgNameOffset(i) + OffsetNameTable, Number)
Next i
Case Else
If Inhalt(Nummer).NumberofMethods > 0 Then
Methodenstand = 0
Do While Stand < EndOffset
If Methodenstand < Inhalt(Nummer).NumberofMethods - Inhalt(Nummer).NumberOfVariables Then
CopyMemory ByVal VarPtr(Code), Typelibarray(Number).Bytes(Stand), 24
Helpint = Code.EnumWert And 4096
'bei Code.vTableOffset orginalFunktionsname
'MsgBox Code.vTableOffset
If Helpint = 4096 Then
OptAuf = Code.nrargs * 4
Else
OptAuf = 0
End If
If Code.Recsize > 24 Then
OptHelpAuf = GetOptional(Number, Stand + 24, Stringoffset, Code.Recsize - (Code.nrargs * 12) - OptAuf, Inhalt(Nummer).MethTable(Methodenstand).Helpcontext, Inhalt(Nummer).MethTable(Methodenstand).Helpstring, Inhalt(Nummer).MethTable(Methodenstand).ConstantValue, 0, , , , Inhalt(Nummer).MethTable(Methodenstand).OrginalName, Code.EnumWert)
End If
Stand = Stand + Code.Recsize - (Code.nrargs * 3 * 4)

Codeart = Testart(Code.EnumWert, Inhalt(Nummer).VorgabeNr)
Select Case Codeart
Case 0
Art = "Type"
Case 1
Art = "Sub"
Case 2
Art = "Funktion"
Case 3
Art = "Property Get"
Case 4
Art = "Property Let"
Case 5
Art = "Property Set"
Case 6
If allsub = False Then ' sometimes error
Art = "Event"
Else
Art = "Sub"
End If
End Select
Select Case Code.Calling1
Case 24, 25
Case Else
Art = "stdcall"
allsub = True
End Select
If Art <> "Fehler" Then
Inhalt(Nummer).MethTable(Methodenstand).MethodenTyp = Art
Inhalt(Nummer).MethTable(Methodenstand).NrArguments = Code.nrargs
If Code.Flags And 64 = 64 Then
Inhalt(Nummer).MethTable(Methodenstand).IsNotVisible = True
Else
If Code.Flags And 512 = 512 Then Inhalt(Nummer).MethTable(Methodenstand).IsStandard = True
End If
If Code.nrargs > 0 Then
If Art = "stdcall" And Inhalt(Nummer).VorgabeNr = 8483 And Code.nrargs > 2 Then
CopyMemory ByVal VarPtr(Firststdcall(0)), Typelibarray(Number).Bytes(Stand), 8 'Achtung
Stand = Stand + 8
End If
ReDim ArgumententypArray(Code.nrargs - 1)
ReDim Inhalt(Nummer).MethTable(Methodenstand).Argumente(Code.nrargs - 1)
Optionalstand = Code.Recsize - 24 - (Code.nrargs * 12)
Testlong = Code.Recsize - (Code.nrargs * 12)
Optionalstand = StandBegin + 24 + (OptHelpAuf * 4)
Stand = StandBegin + Testlong 'Code.RecSize  '????
'If Optionalstand > 0 Then Optionalstand = Stand - Optionalstand
CopyMemory ByVal VarPtr(ArgumententypArray(0)), Typelibarray(Number).Bytes(Stand), Code.nrargs * 12
Stand = Stand + (Code.nrargs * 12)
Select Case Art
Case "Sub", "Event"
For i = 0 To Code.nrargs - 1
If ArgumententypArray(i).Datatype >= 0 Then
If ArgumententypArray(i).Datatype1 <> 0 Then
Select Case ArgumententypArray(i).Datatype2
Case 128
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ = MakeTypestring(ArgumententypArray(i).Datatype)
Case Else
End Select
Else
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ = TypeArray(ArgumententypArray(i).Datatype \ 8)
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).ByValOrByRef = 1
End If
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Namen = GetMSFTName(ArgumententypArray(i).oName + OffsetNameTable, Number)
End If
TestArgflags ArgumententypArray(i).Flags, Nummer, Methodenstand, i, Number, CustdataOffset, Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ
Next i
Case "Property Let", "Property Set"
For i = 0 To Code.nrargs - 1 'immer 1
Select Case ArgumententypArray(i).Datatype
Case 0
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ = TypeArray(ArgumententypArray(i).Datatype \ 8)
Case Else
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ = MakeTypestring(ArgumententypArray(i).Datatype)
End Select
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Namen = "vNewValue"
TestArgflags ArgumententypArray(i).Flags, Nummer, Methodenstand, i, Number, CustdataOffset, Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ
Next i

Case "Funktion", "Property Get"
For i = 0 To Code.nrargs - 1
If i < (Code.nrargs - 1) Then
If ArgumententypArray(i).Datatype >= 0 Then
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ = TypeArray(ArgumententypArray(i).Datatype \ 8)
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Namen = GetMSFTName(ArgumententypArray(i).oName + OffsetNameTable, Number)
End If
Else
If ArgumententypArray(i).Datatype >= 0 Then
Inhalt(Nummer).MethTable(Methodenstand).RückgabeTyp = TypeArray(ArgumententypArray(i).Datatype \ 8)
Inhalt(Nummer).MethTable(Methodenstand).NrArguments = Inhalt(Nummer).MethTable(Methodenstand).NrArguments - 1
End If
End If
TestArgflags ArgumententypArray(i).Flags, Nummer, Methodenstand, i, Number, CustdataOffset, Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ
Next i
Case "stdcall"
Inhalt(Nummer).MethTable(Methodenstand).MethodenTyp = "Funktion"
Inhalt(Nummer).MethTable(Methodenstand).RückgabeTyp = MakeTypestring(Code.Calling1)
Inhalt(Nummer).MethTable(Methodenstand).NrArguments = Code.nrargs
For i = 0 To Code.nrargs - 1
Select Case ArgumententypArray(i).Datatype2
Case 0
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ = MakeTypestring(CInt(ArgumententypArray(i).Flags))
Case Else
Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ = MakeTypestring(ArgumententypArray(i).Datatype)
End Select

Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Namen = GetMSFTName(ArgumententypArray(i).oName + OffsetNameTable, Number)
TestArgflags ArgumententypArray(i).Flags, Nummer, Methodenstand, i, Number, CustdataOffset, Inhalt(Nummer).MethTable(Methodenstand).Argumente(i).Typ
Next i
End Select
Else
If Art = "stdcall" Then
Art = "Funktion"
Inhalt(Nummer).MethTable(Methodenstand).MethodenTyp = Art
End If
End If
End If
Else
Inhalt(Nummer).MethTable(Methodenstand).MethodenTyp = "Variable"
CopyMemory ByVal VarPtr(Code), Typelibarray(Number).Bytes(Stand), 24
If Code.Recsize > 20 Then
GetOptional Number, Stand + 20, Stringoffset, Code.Recsize + 4, Inhalt(Nummer).MethTable(Methodenstand).Helpcontext, Inhalt(Nummer).MethTable(Methodenstand).Helpstring, Inhalt(Nummer).MethTable(Methodenstand).ConstantValue, 0, CLng(Code.EnumWert), IsConst, Code.EnumNumber   'hier evtl Fehler (EnumNumber)
'Const only 20!!
If IsConst = 1 Then Inhalt(Nummer).MethTable(Methodenstand).MethodenTyp = "Constant"
End If
Stand = Stand + Code.Recsize
Inhalt(Nummer).MethTable(Methodenstand).RückgabeTyp = MakeTypestring(Code.Calling1)
End If
Methodenstand = Methodenstand + 1
StandBegin = StandBegin + Code.Recsize
Loop
End If
End Select
End Function

Public Function TestAnz(Number As Long, Beginning As Long, Ending As Long) As Long
Dim LenPiece As Integer
Dim Anza As Long
Dim Startpos As Long
Startpos = Beginning + 4
Do While Startpos < Ending
CopyMemory ByVal VarPtr(LenPiece), Typelibarray(Number).Bytes(Startpos), 2
Startpos = Startpos + LenPiece
Anza = Anza + 1
Loop
TestAnz = Anza
End Function

Public Function Testart(Number As Integer, TN As Long) As Long
Dim Test As Long
Dim Test1 As Long
Dim Test2 As Long
Dim Test3 As Long
Dim Test4 As Long
Dim Test5 As Long
Dim Backg As Long

Test = Number And 1036
If Test = 1036 Then
Select Case TN
Case 8484
Backg = 6 'Event
Case 8500
Backg = 1 'Sub
End Select
End If
Test1 = Number And 1089
If Test1 = 1089 Then Backg = 5 'Set Prop
Test2 = Number And 17425
If Test2 = 17425 Then Backg = 3 'Get Prop
Test3 = Number And 1057
If Test3 = 1057 Then Backg = 4 'Prop Let
Test4 = Number And 1033
If Test4 = 1033 Then Backg = 1 'Sub
Test5 = Number And 17417
If Test5 = 17417 Then Backg = 2 'Funktion
If TN = 8481 Then Backg = "Type"
If TN = 8480 Then Backg = "Enum"
Testart = Backg
End Function

Public Function TestArgflags(Flag As Long, Nummer As Long, Methodenstand As Long, argnr As Long, Number As Long, CustdataOffset As Long, Typ As String)
Dim Testlong As Long
Dim TypInt As Integer
Dim Byte4Array(3) As Byte
Dim Testdouble As Double
Dim AufLong As Long
Dim Hilfsboolean As Boolean
Dim Hilfsstring As String

Testlong = Flag And 1 'in
If Testlong = 1 Then Inhalt(Nummer).MethTable(Methodenstand).Argumente(argnr).IsIn = 1
Testlong = Flag And 2 'out
If Testlong = 2 Then Inhalt(Nummer).MethTable(Methodenstand).Argumente(argnr).IsOut = 1
Testlong = Flag And 4 'lcid
If Testlong = 4 Then Inhalt(Nummer).MethTable(Methodenstand).Argumente(argnr).Islcid = 1
Testlong = Flag And 8 'retval
If Testlong = 8 Then Inhalt(Nummer).MethTable(Methodenstand).Argumente(argnr).Isretval = 1
Testlong = Flag And 16 'optional
If Testlong = 16 Then Inhalt(Nummer).MethTable(Methodenstand).Argumente(argnr).IsOptional = 1
Testlong = Flag And 32 'Has optional Wert
If Testlong = 32 Then
Inhalt(Nummer).MethTable(Methodenstand).Argumente(argnr).HasOptionalString = 1
AufLong = argnr * 4
CopyMemory Byte4Array(0), Typelibarray(Number).Bytes(Optionalstand + AufLong), 4
If Byte4Array(3) <> 0 Then 'direct
Byte4Array(3) = 0
Testlong = 0
CopyMemory ByVal VarPtr(Testlong), Byte4Array(0), 4
    Select Case Typ
    Case "Boolean"
    If Testlong = 0 Then
    Hilfsboolean = False
    Else
    Hilfsboolean = True
    End If
    Inhalt(Nummer).MethTable(Methodenstand).Argumente(argnr).OptionalString = CStr(Hilfsboolean)
    Case Else
    Inhalt(Nummer).MethTable(Methodenstand).Argumente(argnr).OptionalString = CStr(Testlong)
    End Select
    
Else 'Not direct
Testlong = 0
CopyMemory ByVal VarPtr(Testlong), Byte4Array(0), 4
Hilfsstring = GetWert(Testlong + CustdataOffset, Number)
Inhalt(Nummer).MethTable(Methodenstand).Argumente(argnr).OptionalString = Hilfsstring
End If
End If
End Function

Public Sub GetToName(Offset As Long, Number As Long, Nummer As Long, Startoffset As Long)
Dim IDs() As Long
Dim Offsets() As Long
Dim i As Long
Dim z As Long
Dim NumberMethods As Long
Dim NumberArguments As Long
Dim MName As String

NumberMethods = Inhalt(Nummer).NumberofMethods
If NumberMethods > 0 Then
ReDim IDs(NumberMethods - 1)
ReDim Offsets(NumberMethods - 1)
CopyMemory ByVal VarPtr(IDs(0)), Typelibarray(Number).Bytes(Offset), 4 * NumberMethods
CopyMemory ByVal VarPtr(Offsets(0)), Typelibarray(Number).Bytes(Offset + 4 * NumberMethods), 4 * NumberMethods

For i = 0 To NumberMethods - 1
MName = GetMSFTName(Startoffset + Offsets(i), Number)
Inhalt(Nummer).MethTable(i).MethodenName = MName
If Left(MName, 1) = "_" Then Inhalt(Nummer).MethTable(i).IsNotVisible = True
Inhalt(Nummer).MethTable(i).ID = Hex(IDs(i))
If IDs(i) = 0 Then Inhalt(Nummer).MethTable(i).IsStandard = True
Next i
Else
'ReDim Beginn(2 - 1)
'ReDim Offsets(2 - 1)
'Get Filenumber, Offset, Beginn
'Get Filenumber, , Offsets

'For i = 0 To 2 - 1
'MName = GetName(Startoffset + Offsets(i), Filenumber)
'MName = MName
'Next i
End If
End Sub

Private Function GetInfoType(Number As Long, DT1Begin As Long, Optional iNum As Long) As String
Dim Teststring As String
Select Case Number
Case 2338, 2594
GetInfoType = "module"
Teststring = GetString(DT1Begin, iNum)
Inhalt(iNum).InfoDllName = Teststring
Case 8483, 8739
GetInfoType = "interface"
Case 8484, 8740
GetInfoType = "dispinterface"
Case 8485, 8741
GetInfoType = "coclass"
Case 8486, 4390
GetInfoType = "C-Types"
Case 8742, 4646, 16934
GetInfoType = "Typedef"
Case 8500, 8756
GetInfoType = "dispinterface and interface"
Case Else
GetInfoType = "????? " & "Nr: " & Number
End Select

End Function


Private Function GetOptional(Number As Long, Stand As Long, Stringoffset As Long, Recsize As Integer, Helpcontext As String, Helpstring As String, ConstantValue As String, withoutHelpContext As Long, Optional WertOffset As Long, Optional IsConst As Long = 0, Optional EnumNumber As Integer, Optional OrgName As String, Optional FKCCIC As Integer) As Long
Dim Longtest As Long
Dim Testlong() As Long
Dim Testint As Integer
Dim Andint As Integer
Dim Bytear() As Byte
Dim Anz As Long
Dim ActNumber As Long
Dim ActLen As Long
On Error GoTo ErrorExit
Anz = (Recsize - 24) / 4
GetOptional = Anz
If Anz = 0 Then Exit Function 'No Optional (more only for Arguments)
ReDim Testlong(Anz - 1)
CopyMemory ByVal VarPtr(Testlong(0)), Typelibarray(Number).Bytes(Stand), 4 * Anz
If withoutHelpContext = 0 Then
If Testlong(0) <> -1 And Testlong(0) <> 0 Then
Helpcontext = Hex(Testlong(0))
End If
ActNumber = 1
Anz = Anz - 1
End If
If Anz >= 1 Then
If Testlong(ActNumber) <> -1 Then 'Has Helpstring
CopyMemory ByVal VarPtr(Testint), Typelibarray(Number).Bytes(Stringoffset + Testlong(ActNumber)), 2 'Len
ReDim Bytear(Testint - 1)
CopyMemory Bytear(0), Typelibarray(Number).Bytes(Stringoffset + Testlong(ActNumber) + 2), Testint 'Len
Helpstring = StrConv(Bytear, vbUnicode)
End If
ActNumber = ActNumber + 1
If Anz >= 2 Then 'oEntry?
If Testlong(ActNumber) <> -1 Then 'Has OEntry
Andint = FKCCIC And 8192
If Andint = 8192 Then 'is Numeric
CopyMemory ByVal VarPtr(Longtest), Typelibarray(Number).Bytes(Stand + 8), 4
OrgName = "#" & CStr(Longtest)
Else
CopyMemory ByVal VarPtr(Testint), Typelibarray(Number).Bytes(Stringoffset + Testlong(ActNumber)), 2 'Len
If Testint > 0 Then 'Nothing
ReDim Bytear(Testint - 1)
CopyMemory Bytear(0), Typelibarray(Number).Bytes(Stringoffset + Testlong(ActNumber) + 2), Testint 'Len
OrgName = StrConv(Bytear, vbUnicode)
End If
End If
End If
ActNumber = ActNumber + 1
If Anz >= 3 Then
ActNumber = ActNumber + 1
If Anz >= 4 Then 'Const Value
If Testlong(ActNumber) <> -1 Then 'Has Const
IsConst = 1
Select Case EnumNumber
Case 0
ConstantValue = GetWert(startsegdir(11).Offset + WertOffset, Number)
Case Else '88
ConstantValue = CStr(WertOffset)
End Select
'?IrgeneineNummer = Testlong(ActNumber)
End If
ActNumber = ActNumber + 1
If Anz >= 5 Then 'HelpStringContext
ActNumber = ActNumber + 1
End If
End If
End If
End If
End If
Exit Function
ErrorExit:
Debug.Print "Error in GetOptional"
End Function

Public Function GetWert(Offset As Long, Number As Long) As String
Dim TypInt As Integer
Dim Bytear() As Byte
Dim Hilfslong As Long
Dim Hilfsstring As String
Dim Hilfsdate As Date
Dim Hilfsdouble As Double
Dim Hilfssingle As Single
Dim Hilfscurrency As Currency
Dim Hilfsboolean As Boolean
Dim WertLen As Long
On Error GoTo Ex
CopyMemory ByVal VarPtr(TypInt), Typelibarray(Number).Bytes(Offset), 2
Select Case TypInt
Case 2 'Integer
Case 3 'Long
CopyMemory ByVal VarPtr(Hilfslong), Typelibarray(Number).Bytes(Offset + 2), 4
Hilfsstring = CStr(Hilfslong)
Case 4 'Single
CopyMemory ByVal VarPtr(Hilfssingle), Typelibarray(Number).Bytes(Offset + 2), 4
Hilfsstring = CStr(Hilfssingle)
Case 5 'Double
CopyMemory ByVal VarPtr(Hilfsdouble), Typelibarray(Number).Bytes(Offset + 2), 8
Hilfsstring = CStr(Hilfsdouble)
Case 6 'Currency
CopyMemory ByVal VarPtr(Hilfscurrency), Typelibarray(Number).Bytes(Offset + 2), 8
Hilfsstring = CStr(Hilfscurrency)
Case 7 'Date
CopyMemory ByVal VarPtr(Hilfsdate), Typelibarray(Number).Bytes(Offset + 2), 8
Hilfsstring = CStr(Hilfsdate)
Case 8 'String
CopyMemory ByVal VarPtr(WertLen), Typelibarray(Number).Bytes(Offset + 2), 4
If WertLen <> 0 And WertLen <> -1 Then
ReDim Bytear(WertLen - 1)
CopyMemory Bytear(0), Typelibarray(Number).Bytes(Offset + 2 + 4), WertLen
Hilfsstring = StrConv(Bytear, vbUnicode)
Else
Hilfsstring = vbNullString
End If
Case 9 'Object
Case 11 'Boolean
Case 12 'Variant
Case 13 'Unknown
Case 17 'Byte
Case 30 'String
CopyMemory ByVal VarPtr(WertLen), Typelibarray(Number).Bytes(Offset + 2), 4
ReDim Bytear(WertLen - 1)
CopyMemory Bytear(0), Typelibarray(Number).Bytes(Offset + 2 + 4), WertLen
Hilfsstring = StrConv(Bytear, vbUnicode)
Case Else 'Unknown Type
End Select
GetWert = Hilfsstring
Exit Function
Ex:
Debug.Print "Error in GetWert"
End Function
