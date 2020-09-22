VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   2052
   ClientTop       =   1680
   ClientWidth     =   6384
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   6384
   Begin VB.Frame Frame2 
      Caption         =   "Icons"
      Height          =   1452
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   4812
      Begin VB.OptionButton Option2 
         Caption         =   "Direct from File"
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   4332
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ressource (work only in exe!!!)"
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   4332
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Form1.Picture1"
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   4332
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Form1.Icon"
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   4332
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   432
      Left            =   9480
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stil"
      Height          =   1692
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4812
      Begin VB.OptionButton Option1 
         Caption         =   "3 - without Closebutton"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2772
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2 - Closebutton disabled"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2652
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1-Closebutton enabled"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   3732
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AboutBox"
      Height          =   612
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   1572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Stil As Long
Dim IconNr As Long

Private Sub Command1_Click()
Dim x As Long
Dim y As Long
Dim Picture As StdPicture
Dim Iconhandle As Long
Dim Headertext As String
Dim Message As String

Headertext = "Info to the Aboutbox-Projekt"
Message = "Infodialog-Project (Aboutbox) from 2006," & vbCrLf & "Version 1.0" & vbCrLf & vbCrLf & "Copyright © ALKO"
Select Case IconNr
Case 0
Iconhandle = Form1.icon.Handle
Case 1
Iconhandle = Picture1.Picture.Handle
Case 2
Iconhandle = LoadResPicture(101, vbResIcon)
Case 3
Set Picture = LoadPicture(App.Path & "\Yellfish.ico")
Iconhandle = Picture.Handle
End Select
AboutBox Me.hWnd, Headertext, Message, Iconhandle, True, Stil
Set Picture = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
Stil = Index
End Sub

Private Sub Option2_Click(Index As Integer)
IconNr = Index
End Sub
