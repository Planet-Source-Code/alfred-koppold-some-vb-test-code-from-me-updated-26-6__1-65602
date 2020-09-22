VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6384
   ClientLeft      =   1248
   ClientTop       =   1392
   ClientWidth     =   6384
   LinkTopic       =   "Form1"
   ScaleHeight     =   6384
   ScaleWidth      =   6384
   WindowState     =   2  'Maximiert
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6132
      Left            =   960
      ScaleHeight     =   6084
      ScaleWidth      =   8364
      TabIndex        =   1
      Top             =   1560
      Width           =   8412
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan with Twain"
      Height          =   372
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   1572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim a As New clsTwain
a.Scan_with_Twain Bitmap_GREY_8bit, False, Picture1, 100, , , , , False
End Sub

