VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7248
   ClientLeft      =   2172
   ClientTop       =   1536
   ClientWidth     =   6396
   LinkTopic       =   "Form1"
   ScaleHeight     =   7248
   ScaleWidth      =   6396
   WindowState     =   2  'Maximiert
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   13695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   13695
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   120
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1137
      TabIndex        =   1
      Top             =   120
      Width           =   13695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   1320
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      Filter          =   "*.ttf|*.ttf"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Fontfile"
      Height          =   492
      Left            =   4080
      TabIndex        =   0
      Top             =   6600
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Text to write"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowOpen
Main CommonDialog1.Filename, Text1.Text, Text2.Text
End Sub

Private Sub Form_Load()
Text1 = "The quick brown fox jumps over the lazy dog."
Text2 = "1234567890"
End Sub
