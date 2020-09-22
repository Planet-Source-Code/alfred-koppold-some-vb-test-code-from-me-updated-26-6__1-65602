VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   ClientHeight    =   6384
   ClientLeft      =   2388
   ClientTop       =   1968
   ClientWidth     =   6384
   LinkTopic       =   "Form1"
   ScaleHeight     =   6384
   ScaleWidth      =   6384
   Begin VB.Frame Frame4 
      Caption         =   "Picstyle"
      Height          =   1692
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   2532
      Begin VB.OptionButton optPicBoarder 
         Caption         =   "Boarder bump"
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   2052
      End
      Begin VB.OptionButton optPicBoarder 
         Caption         =   "Boarder etched"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   2052
      End
      Begin VB.OptionButton optPicBoarder 
         Caption         =   "Boarder raised"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2052
      End
      Begin VB.OptionButton optPicBoarder 
         Caption         =   "Boarder sunken"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
         Width           =   2052
      End
      Begin VB.OptionButton optPicBoarder 
         Caption         =   "no Boarder"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2052
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Stil"
      Height          =   972
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   2532
      Begin VB.CheckBox chkFlat 
         Caption         =   "Flat"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   2052
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Position"
      Height          =   1212
      Left            =   3240
      TabIndex        =   8
      Top             =   4920
      Width           =   2532
      Begin VB.CheckBox chkMid 
         Caption         =   "Mid Of Screen"
         Height          =   372
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   1692
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   972
      Left            =   3480
      TabIndex        =   7
      Top             =   3600
      Width           =   2172
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stil"
      Height          =   2052
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2532
      Begin VB.CheckBox chkBackGround 
         Caption         =   "with Background-Picture"
         Height          =   252
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   2172
      End
      Begin VB.CheckBox chkPos 
         Caption         =   "Pic on Bottom"
         Height          =   192
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2052
      End
      Begin VB.CheckBox chkPos 
         Caption         =   "Pic on Top"
         Height          =   192
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2052
      End
      Begin VB.CheckBox chkPos 
         Caption         =   "Pic right"
         Height          =   192
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   2052
      End
      Begin VB.CheckBox chkPos 
         Caption         =   "Pic left"
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2052
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Standard Open"
      Height          =   852
      Left            =   3240
      TabIndex        =   1
      Top             =   2400
      Width           =   2292
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Open"
      Height          =   852
      Left            =   3240
      TabIndex        =   0
      Top             =   1080
      Width           =   2292
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   600
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Image bg1 
      Height          =   1152
      Left            =   10080
      Picture         =   "Form1.frx":0000
      Top             =   480
      Visible         =   0   'False
      Width           =   1152
   End
   Begin VB.Image bg 
      Height          =   480
      Left            =   10560
      Picture         =   "Form1.frx":0B82
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Picstyle As Long
'Stil
'1 = MidofScreen
'2 = Button (Not Flat)

Private Sub Check1_Click()

End Sub

'PicStyle
'0 = NoBoarder
'1 = Sunken
'2 = Raised
'4 = Etched
'8 = Bump

Private Sub cmdButton_Click()
Dim Stil As Long
Dim Images() As StdPicture
ReDim Images(4)
If chkMid Then Stil = Stil Or 1
If chkFlat = 0 Then Stil = Stil Or 2
If chkPos(0) Then
Set Images(0) = LoadPicture(App.Path & "\test.bmp")
End If
If chkPos(1) Then
Set Images(1) = LoadPicture(App.Path & "\Test.bmp")
End If
If chkPos(2) Then
Set Images(2) = LoadPicture(App.Path & "\Test1.bmp")
End If
If chkPos(3) Then
Set Images(3) = LoadPicture(App.Path & "\Mngabout.bmp")
End If
If chkBackGround Then
Set Images(4) = bg.Picture
End If
CommonDialog1.Flags = &H4
CommonDialog1.InitDir = "c:\"
OwnOpen Form1.Hwnd, Images, CommonDialog1, Stil, Picstyle
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
End Sub

Private Sub Command2_Click()
CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects Or cdlCFApply
CommonDialog1.ShowFont
End Sub

Private Sub Form_Load()
Picstyle = 1
End Sub

Private Sub optPicBoarder_Click(Index As Integer)
Select Case Index
Case 0, 1, 2
Picstyle = Index
Case 3
Picstyle = 4
Case 4
Picstyle = 8
End Select
End Sub
