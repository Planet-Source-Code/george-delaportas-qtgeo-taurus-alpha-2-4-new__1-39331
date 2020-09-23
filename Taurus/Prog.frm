VERSION 5.00
Begin VB.Form Prog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label LShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image NoMouse 
      Height          =   480
      Left            =   0
      Picture         =   "Prog.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Prog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program

'Keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyEscape Then
  Unload Me
 End If
End Sub
'Text Warp
Private Sub LShow_Change(Index As Integer)
 If (LShow(Index).Left + LShow(Index).Width) >= 1000 Then
  'Exit Program
  TAURUS.Text2.Text = "Warning: You Have Reached The Limit Of This Screen.The Last Label Will Not Be Drawn."
  Unload Prog
 End If
End Sub
'End
Private Sub Form_Unload(Cancel As Integer)
 'Visible Mouse
 ShowCursor (1)
End Sub
