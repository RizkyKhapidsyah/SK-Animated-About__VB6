VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About MyApplication"
   ClientHeight    =   2955
   ClientLeft      =   2445
   ClientTop       =   2955
   ClientWidth     =   4425
   ControlBox      =   0   'False
   Icon            =   "AboutMW.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2955
   ScaleWidth      =   4425
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   570
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   120
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "your info goes here"
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   120
      TabIndex        =   1
      Top             =   2190
      Width           =   4215
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   3600
      Picture         =   "AboutMW.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   480
      Picture         =   "AboutMW.frx":1194
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "your font and text go here"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   -2000
      Picture         =   "AboutMW.frx":1A5E
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Image1.Left = -2000
Form4.Image1.Visible = False
impos = -2000
gone = 0
runner = 0
Unload Form4
End Sub

Private Sub Form_Load()
DoEvents
Form4.Timer1.Interval = 10
impos = -2000
Form4.Image1.Visible = True
gone = 0
Timer1.Enabled = True
End Sub


Private Sub Image1_Click()
runner = 0
impos = -2000
gone = 0
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If impos < 4460 And runner = 0 Then impos = impos + 150
If impos < 4460 And runner = 0 Then Form4.Image1.Left = impos
If impos >= 4460 Then gone = gone + 1
If gone = 1 Then GoSub tiny
If gone = 1 Then Form4.Image1.Height = 255
If gone = 1 Then Form4.Image1.Width = 495
If gone = 1 Then Form4.Image1.Left = 1440
If gone = 2 Then Form4.Image1.Height = 375
If gone = 2 Then Form4.Image1.Width = 615
If gone = 3 Then Form4.Image1.Height = 495
If gone = 3 Then Form4.Image1.Width = 735
If gone = 4 Then Form4.Image1.Height = 615
If gone = 4 Then Form4.Image1.Width = 855
If gone = 5 Then Form4.Image1.Height = 735
If gone = 5 Then Form4.Image1.Width = 975
If gone = 6 Then Form4.Image1.Height = 855
If gone = 6 Then Form4.Image1.Width = 1095
If gone = 7 Then Form4.Image1.Height = 975
If gone = 7 Then Form4.Image1.Width = 1215
If gone = 8 Then Form4.Image1.Height = 1095
If gone = 8 Then Form4.Image1.Width = 1335
If gone = 9 Then Form4.Image1.Height = 1215
If gone = 9 Then Form4.Image1.Width = 1455
If gone > 1 Then gone = gone + 1
If gone = 30 Then Form4.Timer1.Enabled = False
If gone = 30 Then Command1.Visible = True
GoTo endit
tiny:
runner = runner + 1
Form4.Image1.Height = 60
Form4.Image1.Width = 60
impos = impos - 150
Form4.Image1.Left = impos
If impos <= 1440 Then gone = 2
If gone = 2 Then Return
endit:




End Sub


