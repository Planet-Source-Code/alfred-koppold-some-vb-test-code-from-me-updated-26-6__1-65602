VERSION 5.00
Object = "*\AProjekt1.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1104
   ClientLeft      =   2808
   ClientTop       =   3528
   ClientWidth     =   6384
   LinkTopic       =   "Form1"
   ScaleHeight     =   1104
   ScaleWidth      =   6384
   Begin VB.PictureBox Picture2 
      Height          =   492
      Left            =   3720
      ScaleHeight     =   444
      ScaleWidth      =   444
      TabIndex        =   1
      Top             =   240
      Width           =   492
   End
   Begin VB.PictureBox Picture1 
      Height          =   612
      Left            =   2520
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   564
      ScaleWidth      =   564
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
   Begin ComLib.ImageList ImageList1 
      Left            =   1440
      Top             =   480
      _ExtentX        =   804
      _ExtentY        =   804
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Picture = normale Größe
'draw = ImageList Größe
Private Sub Form_Load()
ImageList1.ListImages.Add , , Picture1.Picture
Picture2.Picture = ImageList1.ListImages.Item(1).Picture
End Sub
