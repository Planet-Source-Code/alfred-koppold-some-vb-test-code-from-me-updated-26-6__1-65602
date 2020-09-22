VERSION 5.00
Begin VB.Form AbortForm 
   Caption         =   "Abort Printing"
   ClientHeight    =   1695
   ClientLeft      =   2235
   ClientTop       =   2340
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   3795
   Begin VB.CommandButton CmdAbort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Press to Abort Print Job"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "AbortForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Copyright © 1997 by Desaware Inc. All Rights Reserved.

Private Sub CmdAbort_Click()
    AbortPrinting = -1
End Sub


