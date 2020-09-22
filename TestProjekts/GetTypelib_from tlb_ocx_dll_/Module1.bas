Attribute VB_Name = "Module1"
Option Explicit
Public Type MSFT_NAMEINTRO
NameNumber As Long
unk10 As Long
Namelen As Byte
unk20 As Byte
unk30 As Integer
End Type

Public Type Nametable
Name As String
Number As Long
End Type

Public NameTesttable() As Nametable
Public Function FillNameArray(Startpos As Long, Number As Long, Leng As Long, Interfacenumber As Long, Interfacenamestand As Long, Nummer As Long) As Long
Dim NameIntro As MSFT_NAMEINTRO
Dim Name As String
Dim Stand As Long
Dim Bytear() As Byte
Dim NrOfNames As Long
Stand = Startpos
ReDim NameTesttable(0)
Do While Stand < Startpos + Leng
CopyMemory ByVal VarPtr(NameIntro), Typelibarray(Number).Bytes(Stand), 12
Stand = Stand + 12
ReDim Bytear(NameIntro.Namelen - 1)
CopyMemory Bytear(0), Typelibarray(Number).Bytes(Stand), NameIntro.Namelen
Name = StrConv(Bytear, vbUnicode)
ReDim Preserve NameTesttable(NrOfNames)
NameTesttable(NrOfNames).Number = NameIntro.NameNumber
NameTesttable(NrOfNames).Name = Name
Stand = Stand + TestFourBound(NameIntro.Namelen)
NrOfNames = NrOfNames + 1
Loop
End Function

Private Function TestFourBound(Number As Byte) As Integer
Dim Test As Integer
If Number Mod 4 = 0 Then
Test = Number
Else
Test = Number \ 4
Test = Test + 1
Test = Test * 4
End If
TestFourBound = Test
End Function

