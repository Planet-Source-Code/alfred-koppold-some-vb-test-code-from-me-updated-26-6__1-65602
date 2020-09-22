VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm24Bit 
   Caption         =   "24 Bits Per Pixel"
   ClientHeight    =   4368
   ClientLeft      =   1320
   ClientTop       =   1392
   ClientWidth     =   6384
   LinkTopic       =   "Form2"
   ScaleHeight     =   4368
   ScaleWidth      =   6384
   Begin VB.CommandButton Command1 
      Caption         =   "Do it"
      Height          =   612
      Left            =   4200
      TabIndex        =   13
      Top             =   2640
      Width           =   1692
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   1440
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transparent Color (click to set)"
      Height          =   972
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   3012
      Begin VB.PictureBox pbTransparentColor 
         Height          =   132
         Left            =   2040
         ScaleHeight     =   84
         ScaleWidth      =   804
         TabIndex        =   10
         Top             =   600
         Width           =   852
      End
      Begin VB.CheckBox chkTransparent 
         Caption         =   "Set Transparent Color"
         Height          =   192
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1932
      End
      Begin VB.PictureBox pbColorunderCursor 
         Height          =   132
         Left            =   2040
         ScaleHeight     =   84
         ScaleWidth      =   804
         TabIndex        =   8
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "Color under Cursor"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1812
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "BkgdColor"
      Height          =   492
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2892
      Begin VB.PictureBox picbkgd 
         BackColor       =   &H00FFFFFF&
         Height          =   132
         Left            =   1560
         ScaleHeight     =   84
         ScaleWidth      =   804
         TabIndex        =   6
         Top             =   240
         Width           =   852
      End
      Begin VB.CheckBox chkBkgd 
         Caption         =   "Set BkgdColor"
         Height          =   192
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Alphablend"
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3492
      Begin VB.TextBox txtAlphablend 
         Height          =   288
         Left            =   2760
         TabIndex        =   2
         Text            =   "100"
         Top             =   480
         Width           =   612
      End
      Begin VB.CheckBox chkAlphablend 
         BackColor       =   &H8000000A&
         Caption         =   "with Alphablend"
         Height          =   132
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3012
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         Caption         =   "Alphablend (Number from 0 to 255)"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2532
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Only Transparent or Alphablend (not and)"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   3252
   End
End
Attribute VB_Name = "frm24Bit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAlphablend_Click()
Select Case chkAlphablend
Case 1
chkTransparent = 0
End Select
End Sub

Private Sub chkTransparent_Click()
Select Case chkTransparent
Case 1
chkAlphablend = 0
End Select
End Sub

Private Sub Command1_Click()
Me.Visible = False
End Sub

Private Sub picbkgd_Click()
CommonDialog1.ShowColor
picbkgd.BackColor = CommonDialog1.Color
End Sub

