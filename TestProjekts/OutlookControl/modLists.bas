Attribute VB_Name = "modLists"
Option Explicit

Public Type VariantNode
    Prev As Long
    Next As Long
    Key  As String
    Data As Variant
End Type

Public Type ObjectNode
    Prev As Long
    Next As Long
    Key  As String
    Data As Object
End Type

Public Type LongNode
    Prev As Long
    Next As Long
    Key  As String
    Data As Long
End Type

Public Type StringNode
    Prev As Long
    Next As Long
    Key  As String
    Data As String
End Type
