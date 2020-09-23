VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shape Flasher"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   2895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Interval        =   5
      Left            =   600
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   1560
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   1080
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   240
      Shape           =   3  'Circle
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   120
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer3.Enabled = True
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
MsgBox "Simple Shape Flasher created by Bobby Carter! Please vote on PSC :)", vbExclamation, "Shape Flasher"
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Command3.Visible = True
End Sub

Private Sub Timer1_Timer()
Shape1.Visible = True
Timer2.Enabled = True
Command1.Visible = False
End Sub

Private Sub Timer2_Timer()
Shape1.Visible = False
Shape2.Visible = True
End Sub

Private Sub Timer3_Timer()
Shape1.Visible = False
Shape2.Visible = True
Timer2.Enabled = False
Timer1.Enabled = False
Command2.Visible = False
Command3.Visible = True
End Sub
