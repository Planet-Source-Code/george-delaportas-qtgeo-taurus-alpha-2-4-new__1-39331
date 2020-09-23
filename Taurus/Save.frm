VERSION 5.00
Begin VB.Form TLoadSave 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save Source File"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   585
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   45
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "C:\"
      Top             =   225
      Width           =   4605
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   855
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   585
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "WHERE DO YOU WANT TO SAVE THE SOURCE FILE?"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4650
   End
End
Attribute VB_Name = "TLoadSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Load And Save

Option Explicit
'Dim Area
Dim File As String
'Position
Private Sub Form_Load()
 Me.Left = TAURUS.Left + TAURUS.Width / 2 - Me.Width / 2
 Me.Top = TAURUS.Top + TAURUS.Height / 2 - Me.Height / 2
End Sub
'Load/Save Source File
Private Sub Command1_Click()
 'First Main Check
 If Text1.Text = "C:\" Or Text1.Text = "" Then GoTo EndIt:
 On Error GoTo EndIt:
 'Close All Files
 Close
 'Main Check
 If Me.Caption = "Save Source File" Then
  'Save Source Code
  Open Text1.Text & ".tef" For Output As #1
   Print #1, TAURUS.Text1.Text
  Close #1
 Else
  'Clear Previous File
  TAURUS.Text1.Text = ""
  'Load Source Code
  Open Text1.Text & ".tef" For Input As #2
   While Not EOF(2)
    Line Input #2, File
    TAURUS.Text1.Text = TAURUS.Text1.Text + File
    TAURUS.Text1.Text = TAURUS.Text1.Text & vbCrLf
   Wend
  Close #2
 End If
 'Save File Path
 TAURUS.FPath.Caption = Text1.Text
 'Unload This Form
 TAURUS.Enabled = True
 TAURUS.Text2.Text = "OK"
 Unload Me
 GoTo OK:
EndIt:
 'Unload This Form With Error
 TAURUS.Enabled = True
 TAURUS.Text2.Text = "Error During Loading Or Saving File."
 Unload Me
OK:
End Sub
'Cancel
Private Sub Command2_Click()
 'Unload This Form
 TAURUS.Enabled = True
 Unload Me
End Sub
'Text
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 'Enter Key Activation
 If KeyCode = vbKeyReturn Then
  'First Main Check
  If Text1.Text = "C:\" Or Text1.Text = "" Then GoTo EndIt:
  On Error GoTo EndIt:
  'Cloes All Files
  Close
  'Main Check
  If Me.Caption = "Save Source File" Then
   'Save Source Code
   Open Text1.Text & ".tef" For Output As #1
    Print #1, TAURUS.Text1.Text
   Close #1
  Else
   'Clear Previous File
   TAURUS.Text1.Text = ""
   'Load Source Code
   Open Text1.Text & ".tef" For Input As #2
    While Not EOF(2)
     Line Input #2, File
     TAURUS.Text1.Text = TAURUS.Text1.Text + File
     TAURUS.Text1.Text = TAURUS.Text1.Text & vbCrLf
    Wend
   Close #2
  End If
  'Save File Path
  TAURUS.FPath.Caption = Text1.Text
  'Unload This Form
  TAURUS.Enabled = True
  TAURUS.Text2.Text = "OK"
  Unload Me
  GoTo OK:
EndIt:
  'Unload This Form With Error
  TAURUS.Enabled = True
  TAURUS.Text2.Text = "Error During Loading Or Saving File."
  Unload Me
 End If
OK:
End Sub
