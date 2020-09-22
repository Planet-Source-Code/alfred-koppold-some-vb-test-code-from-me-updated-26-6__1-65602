Attribute VB_Name = "tlbStandard"
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ROpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Private Declare Function RQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Public Type byteArray
Bytes() As Byte
End Type

Public Type Attributs
dual As Long '64
hidden As Long '16
nonextensible As Long '128
oleautomation As Long '256
restricted As Long '512
noncreatable_control As Long '32 '???
End Type

Public Type ARGUMENTS
Namen As String
Typ As String
Wert As Long
ByValOrByRef As Long
Array As String
IsIn As Long
IsOut As Long
Islcid As Long
Isretval As Long
IsOptional As Long
HasOptionalString As Long
OptionalString As String
Helpstring As String
Helpcontext As String
ConstantValue As String
End Type

Public Type Methodes
MethodenName As String
OrginalName As String
NameOffset As Long
MethodenTyp As String
NrArguments As Long
Argumente() As ARGUMENTS
RückgabeTyp As String
Helpstring As String
Helpcontext As String
ConstantValue As String
ID As String
IsNotVisible As Boolean
IsStandard As Boolean
End Type

Public Type Consts
ConstName As String
ConstTyp As String
ConstValue As String
End Type

Public Type Types
TypeName As String
NameOffset As Long
NrArguments As Long
Argumente() As ARGUMENTS
End Type

Public Type IMPORTED_LIBS
Name As String
Libname As String
GUID As String
Pfad As String
End Type


Public Type TYPELIB_INHALT
Infoname As String
Infonumber As Long
MethTable() As Methodes
TypeTable() As Types
ConstTable() As Consts
GUIDString As String
NumberofMethods As Long
NumberOfVariables As Long
NumberOfConst As Long
NumberOfTypes As Long
InhaltTyp As String
VorgabeNr As Long
InhaltTypString As String
TypNumber As Long
Helpstring As String
Helpcontext As String
Version As String
Attribute As Attributs
ImportedLib() As IMPORTED_LIBS
HasImpLibs As Boolean
InfoDllName As String
End Type


Public Type TYPELIB_DESCRIPTION
TypelibFilename As String
TypelibRealname As String
Filename As String
TypelibGUID As String
TypelibVersion As String
Helpstring As String
Helpfilename As String
Helpcontext As String
NrOfImpLibs As Long
NrOfTypeInfos As Long
End Type

Public TypelibDescription As TYPELIB_DESCRIPTION
Public Typelibname() As String
Public Typelibarray() As byteArray
Public Typelibanz As Long
Private Type IMAGEDOSHEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type

Private Type IMAGEDATADIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Private Type IMAGEFILEHEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Private Type IMAGEOPTIONALHEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Reserved1 As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(1 To 16) As IMAGEDATADIRECTORY
End Type

Private Type IMAGESECTIONHEADER
    NameSec As String * 8
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
End Type

Private Type IMAGERESOURCEDIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    NumberOfNamedEntries As Integer
    NumberOfIdEntries As Integer
End Type

Private Type IMAGERESOURCEDIRECTORYENTRY
    Name As Long
    OffsetToData As Long
End Type

Private Type IMAGERESOURCEDATAENTRY
    OffsetToData As Long
    Size As Long
    CodePage As Long
    Reserved As Long
End Type

Private Type IMAGERESOURCEDIRSTRINGU
    Length As Integer
    NameString(64) As Byte
End Type

Private Const mask = 2147483647

Private Doshead As IMAGEDOSHEADER
Private Filehead As IMAGEFILEHEADER
Private ImgOphead As IMAGEOPTIONALHEADER
Private SectionHead As IMAGESECTIONHEADER
Private RootResDir As IMAGERESOURCEDIRECTORY
Private TypeIrde() As IMAGERESOURCEDIRECTORYENTRY
Private Const REG_SZ = 1
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767
Public Inhalt() As TYPELIB_INHALT

Public Function TestFile(Filename As String) As Long 'Is tlb in File??
Dim FileHandle As Long
Dim Teststring As String
Dim Testlong  As Long

On Error GoTo ErrorDescr

    If Filename = "" Then
    TestFile = 0
    Exit Function
    End If
FileHandle = FreeFile
Open Filename For Binary As FileHandle
    If LOF(FileHandle) < 10 Then GoTo ErrorDescr
Teststring = Space(2)
Get FileHandle, 1, Teststring
    Select Case Teststring
        Case "MZ" 'exe..?
        Get FileHandle, 1, Doshead
        Get FileHandle, Doshead.e_lfanew + 1, Teststring
        If Teststring <> "PE" Then GoTo ErrorDescr
        TestFile = 2 'exe..
        Case "MS" 'tlb?
        Teststring = Space(4)
        Get FileHandle, 1, Teststring
        If Teststring <> "MSFT" Then
        GoTo ErrorDescr
        Else
        TestFile = 1 'tlb
        End If
        Case Else
        GoTo ErrorDescr
        End Select
Close FileHandle
Exit Function
ErrorDescr:
TestFile = 0
Close FileHandle
End Function

Public Function GetTypelibFromMZ(Filename As String, Optional Name As String) As Long
Dim FileHandle As Long
Dim i As Long
Dim counter As Long
Dim RSRCsechead As Long
Dim ResType As Long
Dim IrdeOff As Long
Dim IrdeOffset As Long
Dim ResDir1 As IMAGERESOURCEDIRECTORY
Dim ResDir2() As IMAGERESOURCEDIRECTORY
Dim ResIrde1() As IMAGERESOURCEDIRECTORYENTRY
Dim ResIrde2() As IMAGERESOURCEDIRECTORYENTRY
Dim IRDAE() As IMAGERESOURCEDATAENTRY
Dim NewLoop As Long
Dim NewOffset As Long
Dim Restypename As String
Dim NewFileName As String
Dim NewFileNumber As Long
Typelibanz = 0
On Error GoTo ErrorDescr

FileHandle = FreeFile

Open Filename For Binary As FileHandle
Get FileHandle, 1, Doshead
Get FileHandle, Doshead.e_lfanew + 5, Filehead
Get FileHandle, Doshead.e_lfanew + 25, ImgOphead
Get FileHandle, Doshead.e_lfanew + 249, SectionHead
       Do While i < Filehead.NumberOfSections
                If Mid(SectionHead.NameSec, 1, 5) = ".rsrc" Then
                    Exit Do
                End If
                RSRCsechead = RSRCsechead + 40
                Get FileHandle, Doshead.e_lfanew + 249 + RSRCsechead, SectionHead
                i = i + 1
        Loop
Get FileHandle, SectionHead.PointerToRawData + 1, RootResDir
ReDim TypeIrde(1 To RootResDir.NumberOfIdEntries + RootResDir.NumberOfNamedEntries)
For ResType = 1 To RootResDir.NumberOfIdEntries + RootResDir.NumberOfNamedEntries
IrdeOff = 0
    Get FileHandle, SectionHead.PointerToRawData + 1 + Len(RootResDir) + IrdeOffset, TypeIrde(ResType)
    Get FileHandle, (mask + TypeIrde(ResType).OffsetToData + 2) + SectionHead.PointerToRawData, ResDir1
    ReDim ResIrde1(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
    ReDim ResDir2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
    ReDim ResIrde2(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)
    ReDim IRDAE(1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries)

    For i = 1 To ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries
        Get FileHandle, (mask + TypeIrde(ResType).OffsetToData + 2) + SectionHead.PointerToRawData + Len(ResDir1) + IrdeOff, ResIrde1(i)
        Get FileHandle, mask + ResIrde1(i).OffsetToData + 2 + SectionHead.PointerToRawData, ResDir2(i)
            If (ResDir1.NumberOfIdEntries + ResDir1.NumberOfNamedEntries) < (ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries) Then
            ReDim ResIrde2(1 To ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries)
            ReDim IRDAE(1 To ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries)
            counter = 1
        Else
            counter = i
        End If
        If (ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries) > 1 Then
            counter = 1
        End If
    NewOffset = 0
            For NewLoop = 0 To ((ResDir2(i).NumberOfIdEntries + ResDir2(i).NumberOfNamedEntries) - 1)
                Get FileHandle, mask + ResIrde1(i).OffsetToData + 2 + SectionHead.PointerToRawData + Len(ResDir1) + NewOffset, ResIrde2(NewLoop + counter)
                Get FileHandle, ResIrde2(NewLoop + counter).OffsetToData + SectionHead.PointerToRawData + 1, IRDAE(NewLoop + counter)
                Select Case TypeIrde(ResType).Name
                Case Is < 0
                    Restypename = GetResourceName(TypeIrde(ResType).Name, FileHandle)
                    If Restypename = "TYPELIB" Then
                    NewFileName = GetResourceName(ResIrde1(i).Name, FileHandle)
                    ReDim Preserve Typelibname(Typelibanz)
                    Typelibname(Typelibanz) = NewFileName
                    ReDim Preserve Typelibarray(Typelibanz)
                    ReDim Typelibarray(Typelibanz).Bytes(0) 'zur Sicherheit
                    Typelibanz = Typelibanz + 1
                    If Name = NewFileName Or Name = "" Then
                    NewFileName = App.Path & "\" & NewFileName & ".tlb"
                     ReDim Typelibarray(Typelibanz - 1).Bytes(IRDAE(NewLoop + counter).Size - 1)
                    Get FileHandle, IRDAE(NewLoop + counter).OffsetToData - ImgOphead.DataDirectory(3).VirtualAddress + SectionHead.PointerToRawData + 1, Typelibarray(Typelibanz - 1).Bytes
                    NewFileNumber = FreeFile
                    Open NewFileName For Binary As NewFileNumber
                    Put NewFileNumber, 1, Typelibarray(Typelibanz - 1).Bytes
                    Close NewFileNumber
                    DoEvents
                    End If
                    End If
                End Select
                NewOffset = NewOffset + Len(ResIrde2(NewLoop + counter))
            Next NewLoop
        IrdeOff = IrdeOff + Len(ResIrde2(i))
    Next i
IrdeOffset = IrdeOffset + Len(TypeIrde(ResType))
Next ResType
Close FileHandle
GetTypelibFromMZ = Typelibanz
Exit Function
ErrorDescr:
GetTypelibFromMZ = 0
Close FileHandle
End Function

Private Function GetResourceName(OffsetN As Long, OpenResFile As Long) As Byte()
  Dim tmpstr As IMAGERESOURCEDIRSTRINGU
   Dim stra() As Byte
   Dim Offset As Long
   Offset = OffsetN
If Offset < 0 Then
  Offset = mask + Offset + SectionHead.PointerToRawData + 2
  Get OpenResFile, Offset, tmpstr
  ReDim stra(tmpstr.Length * 2)
  Get OpenResFile, Offset + 2, stra
  GetResourceName = Mid(stra, 1, tmpstr.Length)
Else
    GetResourceName = CStr(Offset)
End If
End Function

Public Function LoadDirect(Filename As String) As Long
Dim FileHandle As Long
On Error GoTo ErrDescr
Typelibanz = 1
ReDim Typelibarray(0)
FileHandle = FreeFile
Open Filename For Binary As FileHandle
ReDim Typelibarray(0).Bytes(LOF(FileHandle) - 1)
Get FileHandle, 1, Typelibarray(0).Bytes
LoadDirect = 1
Exit Function
ErrDescr:
LoadDirect = 0
End Function

Public Sub ClearStdModule()
ReDim Typelibname(0)
ReDim Typelibarray(0)
ReDim Inhalt(0)
Typelibanz = 0
TypelibDescription.NrOfImpLibs = 0
TypelibDescription.NrOfTypeInfos = 0
End Sub

Public Function RegGetKeyValue(hcHKEY As Long, strSubKey As String, strValueName As String) As String
    Dim lngReturn As Long
    Dim lngKeyHandle As Long
    Dim strBuffer As String * 255
    lngReturn = ROpenKey(hcHKEY, strSubKey, lngKeyHandle)
    If lngReturn <> 0 Then GoTo ErrorHandler 'if there was an error go to the error handler
    lngReturn = RQueryValueEx(lngKeyHandle, strValueName, 0, REG_SZ, ByVal strBuffer, Len(strBuffer))
    If lngReturn <> 0 Then GoTo ErrorHandler 'if there was an error go to the error handler
    lngReturn = RCloseKey(lngKeyHandle)
    If lngReturn <> 0 Then GoTo ErrorHandler 'if there was an error go to the error handler
    RegGetKeyValue = strBuffer
    Exit Function
ErrorHandler:
    RegGetKeyValue = ""
End Function

Public Function GetKeyInfo(ByVal section As Long, ByVal key_name As String, ByVal indent As Integer, Optional Anweisung As Long) As String
Dim subkeys As Collection
Dim subkey_values As Collection
Dim subkey_num As Integer
Dim subkey_name As String
Dim subkey_value As String
Dim Length As Long
Dim hKey As Long
Dim txt As String
    Set subkeys = New Collection
    Set subkey_values = New Collection
    If ROpenKey(section, key_name, hKey) <> 0 Then
        Exit Function
    End If
    subkey_num = 0
    Do
        ' Enumerate subkeys until we get an error.
        Length = 256
        subkey_name = Space$(Length)
        If RegEnumKey(hKey, subkey_num, subkey_name, Length) <> 0 Then Exit Do
        subkey_num = subkey_num + 1
        
        subkey_name = Left$(subkey_name, InStr(subkey_name, Chr$(0)) - 1)
        subkeys.Add subkey_name
    
        ' Get the subkey's value.
        Length = 256
        subkey_value = Space$(Length)

        If RQueryValueEx(hKey, subkey_name, 0, REG_SZ, subkey_value, Length) <> 0 Then
            subkey_values.Add "Error"
        Else
            ' Remove the trailing null character.
            subkey_value = Left$(subkey_value, Length - 1)
            subkey_values.Add subkey_value
        End If
    Loop
    
    ' Close the key.
    If RCloseKey(hKey) <> 0 Then
         'Error
    End If
Select Case Anweisung
Case 0
    GetKeyInfo = subkeys(1) 'First
Case 1
    GetKeyInfo = subkeys(subkeys.Count) 'Last
Case 2
    For subkey_num = 1 To subkeys.Count
    If subkeys(subkey_num) <> "FLAGS" And subkeys(subkey_num) <> "HELPDIR" Then
        GetKeyInfo = subkeys(subkey_num) 'special
    End If
    Next subkey_num
 End Select
End Function

Public Function GetLibNameInRegistry(GUID As String, Optional Version As String = "-1", Optional Valuestring As String) As String
On Error Resume Next
Dim readystring As String
Dim UnderKey As String
Dim TestVal As String
Dim Underkey2 As String
Dim Name As String
readystring = "{" & GUID & "}"
If Version = "-1" Then
UnderKey = GetKeyInfo(&H80000000, "TypeLib\" & readystring, 0, 1)
Else
UnderKey = Version
End If
readystring = readystring & "\" & UnderKey
UnderKey = GetKeyInfo(&H80000000, "TypeLib\" & readystring, 0, 2)
TestVal = RegGetKeyValue(&H80000000, "TypeLib" & "\" & readystring, "")
TestVal = Left$(TestVal, InStr(TestVal, Chr$(0)) - 1)
Valuestring = TestVal
readystring = readystring & "\" & UnderKey & "\" & "win32"
Name = RegGetKeyValue(&H80000000, "TypeLib" & "\" & readystring, "")
Name = Left$(Name, InStr(Name, Chr$(0)) - 1)
GetLibNameInRegistry = Name
End Function

Public Function UnsignedToLong(Value As Double) As Long
 If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
  If Value <= MAXINT_4 Then
     UnsignedToLong = Value
   Else
     UnsignedToLong = Value - OFFSET_4
End If
End Function

Public Function LongToUnsigned(Value As Long) As Double
If Value < 0 Then
    LongToUnsigned = Value + OFFSET_4
Else
    LongToUnsigned = Value
End If
End Function

Public Function UnsignedToInteger(Value As Long) As Integer
If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
  If Value <= MAXINT_2 Then
    UnsignedToInteger = Value
  Else
    UnsignedToInteger = Value - OFFSET_2
End If
End Function
Public Function IntegerToUnsigned(Value As Integer) As Long
If Value < 0 Then
  IntegerToUnsigned = Value + OFFSET_2
Else
 IntegerToUnsigned = Value
End If
End Function


Public Function StringTest(InString As String) As String
Dim Test As String
Dim i As Long
Dim ErsString As String
StringTest = InString
For i = 0 To 31
Select Case i
Case 0 To 6, 14 To 31
ErsString = Hex(i)
If ErsString <> "0" Then If Len(ErsString) = 1 Then ErsString = "0" & ErsString
ErsString = "\x" & ErsString
Case 7 'Bell
ErsString = "\a"
Case 8 'Back
ErsString = "\b"
Case 9 'Tab
ErsString = "\t"
Case 10 'vblf
ErsString = "\n"
Case 11 'VericalTab
ErsString = "\v"
Case 12 'Formfeed
ErsString = "\f"
Case 13 'vbcr
ErsString = "\r"
End Select
StringTest = Replace(StringTest, Chr(i), ErsString)
Next i

StringTest = Chr(34) & StringTest & Chr(34)
End Function

Public Function MakeGuidString(byteArray() As Byte) As String
Dim GUIDString As String
Dim Teilstring As String

GUIDString = ""
Dim i As Long
For i = 0 To 3
GUIDString = GUIDString & TestHex(Hex(byteArray(3 - i)))
Next i
GUIDString = GUIDString & "-"
For i = 0 To 1
Teilstring = TestHex(Hex(byteArray(5 - i)))
GUIDString = GUIDString & Teilstring
Next i
GUIDString = GUIDString & "-"
For i = 0 To 1
Teilstring = TestHex(Hex(byteArray(7 - i)))
GUIDString = GUIDString & Teilstring
Next i
GUIDString = GUIDString & "-"
For i = 8 To 9
Teilstring = TestHex(Hex(byteArray(i)))
GUIDString = GUIDString & Teilstring
Next i
GUIDString = GUIDString & "-"
For i = 10 To 15
Teilstring = TestHex(Hex(byteArray(i)))
GUIDString = GUIDString & Teilstring
Next i
MakeGuidString = GUIDString
End Function

Public Function TestHex(Teststring As String) As String
If Len(Teststring) Mod 2 Then
TestHex = "0" & Teststring
Else
TestHex = Teststring
End If
End Function

Public Function MakeTypestring(Number As Integer) As String
    'VT_EMPTY = 0,
    'VT_NULL = 1,
    'VT_I2 = 2,
    'VT_I4 = 3,
    'VT_R4 = 4,
    'VT_R8 = 5,
    'VT_CY = 6,
    'VT_DATE = 7,
    'VT_BSTR = 8,
    'VT_DISPATCH = 9,
    'VT_ERROR = 10,
    'VT_BOOL = 11,
    'VT_VARIANT = 12,
    'VT_UNKNOWN = 13,
    'VT_DECIMAL = 14,
    'VT_I1 = 16,
    'VT_UI1 = 17,
    'VT_UI2 = 18,
    'VT_UI4 = 19,
    'VT_I8 = 20,
    'VT_UI8 = 21,
    'VT_INT = 22,
    'VT_UINT = 23,
    'VT_VOID = 24,
    'VT_HRESULT = 25,
    'VT_PTR = 26,
    'VT_SAFEARRAY = 27,
    'VT_CARRAY = 28,
    'VT_USERDEFINED = 29,
    'VT_LPSTR = 30,
    'VT_LPWSTR = 31,
    'VT_RECORD = 36,
    'VT_FILETIME = 64,
    'VT_BLOB = 65,
    'VT_STREAM = 66,
    'VT_STORAGE = 67,
    'VT_STREAMED_OBJECT = 68,
    'VT_STORED_OBJECT = 69,
    'VT_BLOB_OBJECT = 70,
    'VT_CF = 71,
    'VT_CLSID = 72,
    'VT_BSTR_BLOB = 0xfff,
    'VT_VECTOR = 0x1000,
    'VT_ARRAY = 0x2000,
    'VT_BYREF = 0x4000,
    'VT_RESERVED = 0x8000,
    'VT_ILLEGAL = 0xffff,
    'VT_ILLEGALMASKED = 0xfff,
    'VT_TYPEMASK = 0xfff
Select Case Number
Case 1
MakeTypestring = "All (Variant)"
Case 2
MakeTypestring = "Integer"
Case 3, 22
MakeTypestring = "Long"
Case 4
MakeTypestring = "Single"
Case 5
MakeTypestring = "Double"
Case 6
MakeTypestring = "Currency"
Case 7
MakeTypestring = "Date"
Case 8
MakeTypestring = "String"
Case 9
MakeTypestring = "Object"
Case 11
MakeTypestring = "Boolean"
Case 12
MakeTypestring = "Variant"
Case 13
MakeTypestring = "Unknown"
Case 17
MakeTypestring = "Byte"
Case 18
MakeTypestring = "Short [unsigned]"
Case 19
MakeTypestring = "Long [unsigned]"
Case 23
MakeTypestring = "Integer [unsigned]"
Case 30
MakeTypestring = "String"
Case &H99
'Nothing
Case 24
MakeTypestring = "Long" 'void
Case &H19
'Nothing
Case Else
MakeTypestring = "Unknown Type-Nr: " & Number
End Select
End Function


Public Function ShiftRight(ByVal Value As Byte, ByVal ShiftCount As Byte) As Long
Const conMaxLong As Byte = &HFF
Dim BytePower2 As Byte
  Select Case ShiftCount
  Case 0&:  ShiftRight = Value
  Case 1& To 7
  BytePower2 = 2 ^ ShiftCount
    ShiftRight = (Value And conMaxLong - BytePower2 + 1&) \ BytePower2
  Case Else
  ShiftRight = 0
  End Select
End Function
