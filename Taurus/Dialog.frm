VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exit"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No"
      Height          =   375
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   270
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Yes"
      Height          =   375
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "DO YOU REALLY WANT TO EXIT?"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   3030
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Exit Dialog

'Position
Private Sub Form_Load()
 Me.Left = TAURUS.Left + TAURUS.Width / 2 - Me.Width / 2
 Me.Top = TAURUS.Top + TAURUS.Height / 2 - Me.Height / 2
End Sub
'End
Private Sub Command1_Click()
 On Error Resume Next
 'Delete Database And End Program
 Kill App.Path & "\Taurus.tdb"
 Kill App.Path & "\Edit.tcf"
 Kill App.Path & "\Prog.tcf"
 Kill App.Path & "\Commands.tcf"
 Kill App.Path & "\Loops.tcf"
 Unload Prog
 Buffer = TAURUS.Left & "|" & TAURUS.Top
 'Save TAURUS Position To Registry
 SaveSetting App.Title, "SavePos", TAURUS.Caption, Buffer
 End
End Sub
'Cancel
Private Sub Command2_Click()
 TAURUS.Enabled = True
 Unload Me
End Sub
