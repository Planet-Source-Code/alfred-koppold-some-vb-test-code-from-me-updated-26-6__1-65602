VERSION 5.00
Object = "*\AOLListeOCX.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4752
   ClientLeft      =   1536
   ClientTop       =   1668
   ClientWidth     =   7404
   LinkTopic       =   "Form1"
   ScaleHeight     =   4752
   ScaleWidth      =   7404
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   612
      Left            =   2280
      TabIndex        =   1
      Top             =   3360
      Width           =   1212
   End
   Begin OLListeOCX.OLListe OLListe1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12510
      _ExtentY        =   3831
      Mauszeiger      =   2
      Flat            =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowColumnResize=   -1  'True
      Columns         =   8
      CHCaption1      =   ""
      CHFontName1     =   ""
      CHFontSize1     =   0
      CHFontColor1    =   0
      CHPicture1      =   4
      CHStil1         =   1
      CHCaption2      =   ""
      CHFontName2     =   ""
      CHFontSize2     =   0
      CHFontColor2    =   0
      CHPicture2      =   9
      CHStil2         =   1
      CHCaption3      =   ""
      CHFontName3     =   ""
      CHFontSize3     =   0
      CHFontColor3    =   0
      CHPicture3      =   6
      CHStil3         =   1
      CHCaption4      =   ""
      CHFontName4     =   ""
      CHFontSize4     =   0
      CHFontColor4    =   0
      CHPicture4      =   2
      CHStil4         =   1
      CHCaption5      =   ""
      CHFontName5     =   ""
      CHFontSize5     =   0
      CHFontColor5    =   0
      CHPicture5      =   3
      CHStil5         =   1
      CHCaption6      =   "Von"
      CHFontName6     =   ""
      CHFontSize6     =   0
      CHFontColor6    =   0
      CHCaption7      =   "Betreff"
      CHFontName7     =   ""
      CHFontSize7     =   0
      CHFontColor7    =   0
      CHCaption8      =   "Erhalten"
      CHFontName8     =   ""
      CHFontSize8     =   0
      CHFontColor8    =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim RestWidth As Long
    With OLListe1
    .ColumnWidth(1) = 200
    .ColumnWidthIsEditable(6) = True
    .ColumnWidthIsEditable(7) = True
    .ColumnWidthIsEditable(8) = True
    .ColumnWidth(2) = 200
    .ColumnWidth(3) = 200
    .ColumnWidth(4) = 200
    .ColumnWidth(5) = 200
    RestWidth = .Width - 1000
    .ColumnWidth(6) = RestWidth / 3
    .ColumnWidth(7) = RestWidth / 3
    .ColumnWidth(8) = RestWidth / 3

    .AddRow "|||||Manfred Müller|Dies ist ein Test|So 1.10.970 12:03", False, False, "Dies ist ein etwas längerer Text" & vbCrLf & "den man übrigens auch umbrechen kann," & vbCrLf & "so dass mehrere Zeilen dargestellt werden", False
    .AddRow "|||||Hans Maier|Dies ist auch ein Test|Mo 2.10.970 12:03", False, False, "Dieser Text kann ein- und ausgeblendet werden indem man den Eintrag in der Liste doppelt mit der Maus anklickt.", False
    .AddRow "|||||Gerhard Main|Noch ein Test|Di 3.10.970 12:03", False, False, "Wie man an dem vorigen Eintrag sehen kann, wird dieser Text nicht umgebrochen!", False
    .AddRow "|||||Roland Test|Und noch ein Test|Mi 4.10.970 12:03", False, False, "Hierzu müßte ein klein wenig mehr programmiert werden!", False
    .AddRow "|||||Richard Cain|Dies ist auch noch ein Test|Do 5.10.970 12:03", False, False, "Versuchen Sie es doch einfach selbst!", False
    .AddRow "|||||Norbert Seilen|Der letzte Test|Fr 6.10.970 12:03", False, False, "Auch dieser Eintrag hat einen Text.", False
    End With
    
    
End Sub

