VERSION 5.00
Begin VB.Form Credits 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6660
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   375
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4455
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter"
      Height          =   375
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4005
      Width           =   1455
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   -10
      X2              =   6680
      Y1              =   5265
      Y2              =   5265
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   3105
      X2              =   3105
      Y1              =   3645
      Y2              =   5265
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   -10
      X2              =   6680
      Y1              =   3620
      Y2              =   3620
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   -10
      X2              =   6680
      Y1              =   875
      Y2              =   875
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4832
      X2              =   6660
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   4830
      X2              =   4830
      Y1              =   900
      Y2              =   3600
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "qtgeo"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4860
      TabIndex        =   5
      Top             =   3330
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Y3K Network Page: http://www.y3knetwork.com"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1530
      TabIndex        =   3
      Top             =   630
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Credits.frx":0000
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   0
      TabIndex        =   4
      Top             =   3645
      Width           =   3075
   End
   Begin VB.Image Image2 
      Height          =   2700
      Left            =   0
      Picture         =   "Credits.frx":011F
      Top             =   900
      Width           =   4800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Credits.frx":34D6
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   5310
      Width           =   6630
   End
   Begin VB.Image Image1 
      Height          =   2370
      Left            =   4860
      Picture         =   "Credits.frx":358F
      Top             =   900
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   2565
      Picture         =   "Credits.frx":4481
      Top             =   0
      Width           =   1515
   End
End
Attribute VB_Name = "Credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Credits And StartUp Screen

'Start Main Program
Private Sub Command1_Click()
 TAURUS.Show
 Unload Me
End Sub
'Exit
Private Sub Command2_Click()
 End
End Sub

