VERSION 5.00
Begin VB.Form TAURUS 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAURUS"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6300
   Tag             =   "0"
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Help"
      Height          =   375
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4860
      Width           =   4430
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Make"
      Height          =   375
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   4430
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   375
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4005
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load"
      Height          =   375
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4005
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   0
      Top             =   270
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3705
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   225
      Width           =   6315
   End
   Begin VB.TextBox Text2 
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
      Height          =   1050
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5490
      Width           =   6315
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go"
      Height          =   375
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4005
      Width           =   1455
   End
   Begin VB.Label FPath 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   270
      TabIndex        =   10
      Top             =   4140
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   4680
      TabIndex        =   9
      Top             =   5265
      Width           =   1590
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Lines:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5480
      TabIndex        =   6
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESS VIEWER"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   5265
      Width           =   6360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5910
      TabIndex        =   2
      Top             =   0
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "EDITOR"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6360
   End
End
Attribute VB_Name = "TAURUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                 'TAURUS
                          'Programming Language
                                'Coded By
                         'qtgeo(George Delaportas)
                        'http://www.y3knetwork.com
                            'Greek Programmer

'===========================================================================
'About (Read This First Please)

'This Is Actually An Interpreter Of VB.This Gives The User-Programmer
'Much More Time To Think Of How His Program Should Work Corectly Neither
'Than Writting Endless Code!
'Also Its Actually A ShortCut To Big Routines Such As DirectX So
'With Only One Line Of Code The Programmer Can Directly Initialize And Run
'DirectX Routines (Draw Cubes,Draw Triangles,Draw Sphers,Draw 3D Text
'And Other Things).
'He Can Also Load Sounds-Music In No Time!
'Also Programmers Of Taurus May Add Some New Commands To Taurus DataBase.
'Please Inform Me Of Any Changes You Do
'By Sending Us An e-Mail At: questionteamgeo@hotmail.com
'If You Wanna Be A Member Of Y3K Network You May Enter
'In Our ICQ: 107825161 And Then We Will Register You In (If We See That
'You Are Compatible With Our Philosophy And Ofcourse You Are An Active Member)
'Or You Can Enter Our Chat Room In <<Undernet Server>> At: #y3knetwork
'As A Member Of Y3K Network You Can Fully ReProgram Any Of Our Projects
'And In CoOperation Such As Linux Programmers Do We Can Create
'Optimized Programs So As To Help Users Write Easier Code
'Or Working Easier With Their Programs.
'Note That We Also Work On Linux Programs So If AnyOne Of You Feels Really
'Good On Linux Sytems We Would Appreciate To Work With Us.
'Please Report Any BUGS...

'Member Of...
'Planet Source Code (http://www.planetsourcecode.com)
'Programmers Heaven (http://www.programmersheaven.com)
'Toxicity (http://www.toxicity.gr)|Personal Page (http://virax.toxicity.gr)
'Insomnia Programs (http://www.insomniaprogs.tk)|Independent Programmers

'Happy Hacking!
'===========================================================================

'Set Options

'Need To Dim
Option Explicit
'Same Upper And Lower Case Texts
Option Compare Text

'WARNING!
'DO NOT CHANGE THE DIMMED VARIABLES UNLESS YOU REALLY KNOW WHAT YOU ARE DOING!
'IT IS NOT RECOMMENDED TO CHANGE THEM BECAUSE THEY ARE DESCRIBING EACH
'ACTION IN TAURUS!

'Dim Section

'COMMENTS...
'============
'TAURUS Interpreter Is Very Optimized.To Understand Why,Remember That All
'The Commands Are In The Best Position For VB To Search Them.This Means For
'Example That <Run\> Command Has The Highest Priority Of All.
'In Two Words The Whole System Is Using An Intelligent Mechanism Which
'Profits The User So As To Has The Best Results!
'But The Position Of The Commands During The Search Will Be The Best If You
'All Give Me Some Statistics And Tell Me What Kind Of Commands You Would Give
'Better Priority So As To Make The Interpreter Realistic To The Needs Of
'Programmers-Users.
'============

'Integers
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim Z As Integer
Dim M As Integer
Dim D As Integer
Dim N As Integer
Dim V As Integer
Dim U As Integer
Dim C As Integer
Dim H As Integer
Dim FOFL As Integer
Dim LOFL As Integer
Dim OtherL As Integer
Dim Looper As Integer
Dim Times As Integer
Dim FTime As Integer
Dim LTime As Integer
Dim Used As Integer
Dim Founds As Integer
Dim VarsToDel As Integer
Dim Sp As Integer
Dim nTime As Integer
Dim Clr As Integer
Dim Pos As Integer
Dim IntShow As Integer
Dim RndNum As Integer
Dim Absol As Integer
Dim SPos(MaxLines) As Integer
Dim Lnh(MaxLines) As Integer
Dim Num(MaxLines) As Integer
Dim RUNTrace As Integer
Dim EXITrace As Integer
Dim HELPTrace As Integer
'Strings
Dim Smb As String
Dim File As String
Dim Com As String
Dim EndCom As String
Dim Obj As String
Dim Obj2 As String
Dim Obj3 As String
Dim Check As String
Dim Stg As String
Dim Stg2 As String
Dim Search(0 To 99) As String
Dim YCom(MaxLines) As String
Dim Found(MaxLines) As String
Dim Pinax(MaxLines) As String
Dim SCode(MaxLines) As String
Dim Vars(MaxLines) As String
Dim FVars(MaxLines) As String
Dim Txt As String
Dim Txt2 As String
Dim Txt3 As String
'Long
Dim Common As Long
Dim PVal As Long
'Boolean
Dim Ready As Boolean
'Dynamic Array
Dim Vals() As String

'REMEMBER!
'============
'Variables <I>,<J>,<K> And <Z> Are Counters.DO NOT TOUCH <K> BECAUSE YOU WILL
'HAVE PROBLEMS WITH THE COMMANDS.Variables I And J Are Used In
'Looping Statements.More General From All Is <I> Which Is Used In Nested
'Loops Or In Main Loops.<J> Is Used For Loops That Compares Variables
'(So I Does) But Is Less General Than <I>.Lastly <Z> Is A Helping Variable.

'The Way That TAURUS Is Using Variables Is Kinda Strange Than You Might
'Used To Use Them.The Main Difference Is That Variables Are Prosseced From
'Left To Right And Not From Right To Left!!!
'For Example When You Use <Put\> Command The Left Goes To Right.
'(Put\X\Y\) [X->Y,Not Y->X]

'<MaxLines> Is A Constant That <<Tells>> TAURUS The Maximum Number Of Lines
'That Can Be Used By The User.(<MaxLines> Is Inside The Module)
'<Looper> Is A Loop Inspector For Loops.
'The Default Value For All Seted Variables In TAURUS Is Zero(0).
'============

'SOME OF THE DECLARATIONS ARE DECLARED IN A MODULE SO AS TO BE PUBLIC!

'Known BUGS...
'Nothing Yet.

'Begin
Private Sub Form_Load()
 On Error Resume Next
 'Load Last Position Of Form (X,Y) From Registry
 Buffer = GetSetting(App.Title, "SavePos", Me.Caption)
 Pos = InStr(Buffer, "|")
 Me.Left = Left(Buffer, Pos - 1)
 Me.Top = Right(Buffer, Len(Buffer) - Pos)
 'Coding Symbol
 Smb = "¥"
 'Initial Values
 U = 0
 Sp = 1
 nTime = 0
 Founds = 0
 VarsToDel = -1
 IntShow = 0
 'Default Path
 FPath.Caption = App.Path & "\Prog"
 
 'List Of KeyWords
 '=======================
 Search(0) = "run" 'Run Program
 Search(1) = "set" 'Dim Variables
 Search(2) = "put" 'Put Something Into A Variable
 Search(3) = "loop" 'Loop
 Search(4) = "cs" 'Clear Screen
 Search(5) = "show" 'Print Strings Or Variables
 Search(6) = "color" 'Change Screen Color (Back Color)
 Search(7) = "mouse" 'Mouse On/Off
 Search(8) = "rnd" 'Random Function
 Search(9) = "absolute" 'Absolute Function
 Search(10) = "sin" 'Sin
 Search(11) = "cos" 'Cos
 Search(12) = "div" 'Division
 Search(13) = "mod" 'Modification
 Search(14) = "add" 'Addition
 Search(15) = "sub" 'Substruction
 Search(16) = "mul" 'Multiplication
 Search(17) = "sdiv" 'Simple Division
 Search(18) = "sqr" 'Sqrt
 Search(19) = "pow" 'Power
 Search(20) = "key" 'Trace KeyBoard Keys
 Search(21) = "=" 'Equal To
 Search(22) = "&" 'Enterprise <<And>>
 Search(23) = ">" 'Bigger Than
 Search(24) = "<" 'Smaller Than
 Search(25) = "tlen" 'Length Of String
 Search(26) = "end" 'End Loop
 Search(27) = "button" 'Draw Simple Buttons
 Search(28) = "sbutton" 'Draw Special Buttons
 Search(29) = "mouse ico" 'Mouse Icon
 Search(30) = "mouse pic" 'Mouse Picture
 Search(31) = "find" 'Find Word
 Search(32) = "scr" 'Screen Dimensions,Centered:Boolean
 Search(33) = "fade" 'Fade FX
 Search(34) = "mkey" 'Trace Mouse Keys
 Search(35) = "circle" 'Draw Circles
 Search(36) = "square" 'Draw Squares
 Search(37) = "array" 'Make Arrays
 Search(38) = "repeat" 'Repeat Functions
 Search(39) = "in" 'Input
 Search(40) = "out" 'Output
 Search(41) = "draw" 'Draw Pictures
 Search(42) = "play snd" 'Play Sounds
 Search(43) = "play mov" 'Play Movies
 Search(44) = "move" 'Move Any Picture
 Search(45) = "point" 'Print Points On Screen
 Search(46) = "stop" 'Pause For A Single Time
 Search(47) = "fill color" 'Fill In Buttons And Windows Or Screen
 Search(48) = "link" 'GoTo A Label
 Search(49) = "execute" 'Execute A Program
 Search(50) = "search" 'Search Anything
 Search(51) = "timer" 'Timer
 Search(52) = "list" 'Make A New List
 Search(53) = "window" 'Draw A Window
 Search(54) = "@" 'Label
 Search(55) = "line" 'Draw Line
 Search(56) = "dir" 'Open/Close Directoy
 Search(57) = "send" 'Send File To Directory
 Search(58) = "del" 'Delete File
 Search(59) = "type" 'Make User Type
 Search(60) = "end type" 'End User Type
 Search(61) = "printer" 'Print To Printer
 Search(62) = "modem" 'Modem Connect
 Search(63) = "reset" 'Reset All Variables
 Search(64) = "counter" 'Counter For Loops
 Search(65) = "help" 'Help
 Search(66) = "exit" 'Exit Program

'This Is The Place Where You Can Load Your Commands
'Search(67) = "Blank" '
'Search(68) = "Blank" '
'Search(69) = "Blank" '
'Search(70) = "Blank" '
'Search(71) = "Blank" '
'Search(72) = "Blank" '
'Search(73) = "Blank" '
'Search(74) = "Blank" '
'Search(75) = "Blank" '
'Search(76) = "Blank" '
'Search(77) = "Blank" '
'Search(78) = "Blank" '
'Search(79) = "Blank" '
'Search(80) = "Blank" '
'Search(81) = "Blank" '
'Search(82) = "Blank" '
'Search(83) = "Blank" '
'Search(84) = "Blank" '
'Search(85) = "Blank" '
'Search(86) = "Blank" '
'Search(87) = "Blank" '
'Search(88) = "Blank" '
'Search(89) = "Blank" '
'Search(90) = "Blank" '
'Search(91) = "Blank" '
'Search(92) = "Blank" '
'Search(93) = "Blank" '
'Search(94) = "Blank" '
'Search(95) = "Blank" '
'Search(96) = "Blank" '
'Search(97) = "Blank" '
'Search(98) = "Blank" '
'Search(99) = "Blank" '
 
 '=======================
 'Make TAURUS DataBase
 Open App.Path & "\Taurus.tdb" For Output As #1
  For I = 0 To UBound(Search)
   Print #1, Search(I)
  Next I
 Close #1
End Sub
'Debug And Run Program
Private Sub Command1_Click()
 
 'This Is The Section Where The Main Algorithm Works.
 'Firstly It Closes All Open Files And Then It Resets All Variables
 'Used Until Now.Then It Gives Them Their Initial Values.After All That
 'It Comes The Time To Save The Edited Code.So It Opens The Code As OutPut
 'And Puts A Symbol In THe End Of Every Line.This Symbol(¥) Gives TAURUS
 'The Oportunity To Undertand When A Command With Its Arguments Stops And
 'Starts A New One.

Begin:
 'Close All Open Files
 Close
 'Reset Found
 For I = 0 To Founds
  Found(I) = ""
 Next I
 'Reset Variables
 For I = 0 To VarsToDel
  Vars(I) = ""
 Next I
 'Clear All Variables
 On Error Resume Next
 For N = 0 To VarsToDel
  Kill App.Path & "\" & Vars(N) & ".set"
 Next N
 'Clear All Loaded Objects In Prog Form And Clear IntShow
 For I = 0 To IntShow
  Unload Prog.LShow(I)
  IntShow = 0
 Next I
 'Delete Previous Loops,Commands
 Kill App.Path & "\Loops.tcf"
 Kill App.Path & "\Commands.tcf"
 'Initialize Variables
 I = 0
 J = 0
 K = 0
 Z = 0
 C = 0
 D = 0
 Founds = 0
 nTime = 0
 FOFL = 0
 LOFL = 0
 OtherL = -1
 If H = "" Then
  H = -1
 Else
  VarsToDel = H
  H = -1
 End If
 HELPTrace = 0
 RUNTrace = 0
 EXITrace = 0
 Looper = 0
 Pos = 0
 RndNum = 0
 Absol = 0
 Com = ""
 EndCom = ""
 Obj = ""
 Obj2 = ""
 Obj3 = ""
 Clr = 0
 Err.Number = 0
 Ready = False
 'Debug Code And Run
 Label9.Caption = "Checking Code..."
 Timer1.Enabled = True
 'Check If Path Exists
 If FPath.Caption <> "" Then
  'Save Edit File To Disk
  Open App.Path & "\Edit.tcf" For Output As #1
   Print #1, Text1.Text; vbCrLf
  Close #1
  'Delete Previous Appends
  Open "Prog.tcf" For Output As #1
   Print #1, ""
  Close #1
  'Open And Make Simple Edit File To Code
  Open App.Path & "\Edit.tcf" For Input As #2
   While Not EOF(2)
    Input #2, File
    Open "Prog.tcf" For Append As #3
     Stg = Right(File, 1)
     If Stg = "\" Then
      Print #3, File & Smb
     Else
      Print #3, File
     End If
    Close #3
   Wend
  Close #2
  
  'This Is The Section Where Code Is Ready To Be Recognized By Taurus.
  'This Is The Main Program.It Is Really Hard To Explain How It Works
  'But I Will Do My Best.
  'Firstly You Should See The Code Below.
  'Now Watch Carefully.
  '<\> Back Slash Is The Symbol That Programmer Should Use So As To
  'Make Clear To TAURUS That The Word Which Is Using Is A Command.
  'WithOut <\> The Interpeter Does Not Understand That <Run> For Example
  'Is A Command And Just Recognises That As A Comment!
  'BE CAREFUL!.This Does Not Means That Any Word With <\> Is A Command!
  'For Example <Cool\> Is Not A Command And Will Be A Comment For TAURUS.
  '<Pos> Is An Integer That Gets The Position Of <¥> In Any String.When It
  'Finds It Pos Is Getting Bigger By One(Pos<-Pos+1).If <Pos> Is Zero(0)
  'Then This Means That This Line Is A Comment.In Any Other Situation(Not 0<)
  'TAURUS Goes On.Now That We Know This Is A Command We Can Go On Next Check.
  'Note That <Pos> Also Gets The Position Of <\> In Any String.In Our
  'Case This String Is Any Command.Now That We Know That There Is A Slash
  'We Can Check For This Command.
  'The Loop (For...Next) Checks If The Command Is One Of Those In Database.
  'If Yes Then It Saves It In An Array.
  'After All If No Commands Found A Message Says To User That No Commands
  'Were Found.
  
  'Open And Read Code
  Open "Prog.tcf" For Input As #2
Again:
   'Reset Pos
   Pos = 0
   While Not EOF(2)
    Input #2, Stg
    Pos = InStr(Stg, Smb)
    If Pos > 0 Then
     Stg2 = Left(Stg, Pos - 1)
     Pos = InStr(Stg2, "\")
     Com = Left(Stg, Abs(Pos - 1))
     YCom(J) = Right(Stg2, Abs(Len(Stg2) - Len(Com) - 1))
     'Find Commands
     For I = 0 To UBound(Search)
      If Com = Search(I) And Com <> "" Then
       C = 1
       Found(J) = Com
       J = J + 1
       GoTo Again:
      End If
     Next I
    Else
     GoTo Again:
    End If
   Wend
   'Save Found Commands
   Founds = J
   'Check For Suitable Action
   If C = 0 Then
    GoTo EndIt:
   Else
    GoTo Coms:
   End If
  Close #2
 Else
  Text2.Text = "No Commands Found"
  Label9.Caption = "Interpretation...FAILED"
  Timer1.Enabled = True
  GoTo EndIt:
 End If
'Loops
Looping:
'Choose Step
If FTime = Null Or LTime = Null Then Err.Number = 0: GoTo EndIt:
If FTime > LTime Then
 '-
 Sp = -1
Else
 '+
 Sp = 1
End If
'Loop Start
For Times = FTime To LTime Step Sp
 'Initialize Variables
 I = 0
 J = 0
 K = 0
 Z = 0
 D = 0
 RndNum = 0
 Absol = 0
 Com = ""
 Obj = ""
 Obj2 = ""
 Obj3 = ""
 Clr = 0
 Ready = False
 
 'Here Is Where Its Command Is Checked So That It Can Be Fully Recognized
 'By TAURUS System.
 
 'Commands
Coms:
 'Close All Open Files
 Close
 For K = nTime To Founds  'nTime Is A Very Important Variable For Loops
  'Previous Values
  PVal = 0
  'Others
  Used = 0
  M = -1
  Common = 0
  On Error GoTo EndIt:
   '/////////////////////////////////////////
   If Found(K) = "run" And RUNTrace = 0 Then
    If Right(YCom(K), 1) = "\" Then Err.Number = 13: Obj = "Run": GoTo EndIt:
    'Run Program
    Prog.Show
    RUNTrace = 1
   '/////////////////////////////////////////
   ElseIf Found(K) = "set" Then
    If Right(YCom(K), 1) = "" Then: Err.Number = 13: Obj = "Set": GoTo EndIt:
    If Asc(YCom(K)) >= Asc("A") And Asc(YCom(K)) <= Asc("Z") Or Asc(YCom(K)) >= Asc("a") And Asc(YCom(K)) <= Asc("z") Then
     H = H + 1
     Vars(H) = Left(YCom(K), Len(YCom(K)) - 1)
     Open App.Path & "\" & Vars(H) & ".set" For Output As #1
      Print #1, "0"
     Close #1
     VarsToDel = H
    Else
     Err.Number = 3:  GoTo EndIt:
    End If
   '/////////////////////////////////////////
   ElseIf Found(K) = "put" Then
    If Right(YCom(K), 1) = "" Then: Err.Number = 13: Obj = "Put": GoTo EndIt:
    'Extra Initial Values
    D = 0
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) > 2 Or UBound(Vals) < 2 Then Err.Number = 13: Obj = "Put": GoTo EndIt:
    On Error Resume Next
    If H >= 0 Then
     If Vals(1) = "" Then Err.Number = 8: GoTo EndIt:
     'Check If Any Variables Exist
     For J = Used To H
      If Vars(J) = Vals(1) Then
       For I = 0 To H
        If Vars(I) = Vals(0) Then
         'Load Value Or String From This Variable
         Open App.Path & "\" & Vals(0) & ".set" For Input As #2
          Input #2, Vals(0)
          M = 0
          D = 0
         Close #2
         GoTo FillVar:
        Else
FillVar:
         If D = 0 Then
          'Save Value Or String To This Variable
          Open App.Path & "\" & Vals(1) & ".set" For Output As #1
           Print #1, Vals(0)
           M = 0
           D = -1
          Close #1
         End If
        End If
       Next I
       Used = Used + 1
       Exit For
      End If
     Next J
     If M = -1 Then Err.Number = 4: GoTo EndIt:
    Else
     Err.Number = 4
     GoTo EndIt:
    End If
   '/////////////////////////////////////////
   ElseIf Found(K) = "loop" Then
    'Extra Initial Values
    Looper = 1
    Used = 0
    M = -1
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) > 2 Or UBound(Vals) < 2 Then Err.Number = 13: Obj = "Loop": GoTo EndIt:
    If Vals(0) = "" Or Vals(1) = "" Then
     Err.Number = 10: GoTo EndIt:
    Else
     'Find Start,End And The Commands That Will Be Looped
     For I = K + 1 To Founds
      If Found(I) <> "end" And Found(I) <> "" Then
       Open App.Path & "\Loops.tcf" For Append As #3
        Print #3, Found(I)
       Close #3
      ElseIf Found(I) <> "" Then
       nTime = K + 1
       EndCom = "end"
       'Check If Variables Exist.If Not Then Load The Given Values
       For J = Used To H
        For Z = 0 To 1
         If Vars(J) = Vals(Z) Then
          'Load Values From Variables
          Open App.Path & "\" & Vals(Z) & ".set" For Input As #2
           Input #2, Vals(Z)
          Close #2
          FTime = Vals(0)
          LTime = Vals(1)
          Used = Used + 1
          M = 0
          Exit For
         End If
        Next Z
       Next J
       If M < 0 Then
        FTime = Vals(0)
        LTime = Vals(1)
       End If
      End If
     Next I
     'Keep In The First Value (FOFL) And The Last Value Of First Loop (LOFL)
     If FOFL = 0 Then FOFL = Vals(0)
     If LOFL = 0 Then LOFL = Vals(1)
     'Tracer For More Than One Loops
     OtherL = OtherL + 1
     If EndCom <> "end" Then
      'Loop Without End
      Err.Number = 12
      GoTo EndIt:
     Else
      GoTo Looping:
     End If
    End If
   '/////////////////////////////////////////
   ElseIf Found(K) = "cs" Then
    If Right(YCom(K), 1) = "\" Then: Err.Number = 13: Obj = "CS (Clear Screen)": GoTo EndIt:
    On Error GoTo EndIt:
    'Clear Screen
    For I = 1 To IntShow: Unload Prog.LShow(I): Next I
    IntShow = 0
    Prog.BackColor = QBColor(0)
    Prog.Cls
   '/////////////////////////////////////////
   ElseIf Found(K) = "show" Then
    'New LShow
    IntShow = IntShow + 1
    Load Prog.LShow(IntShow)
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 4 Or UBound(Vals) > 4 Then: Err.Number = 13: Obj = "Show": GoTo EndIt:
    If Asc(Vals(1)) >= Asc("0") And Asc(Vals(1)) <= Asc("9") And Asc(Vals(2)) >= Asc("0") And Asc(Vals(2)) <= Asc("9") And Asc(Vals(3)) >= Asc("0") And Asc(Vals(3)) <= Asc("9") Then
     Prog.LShow(IntShow).Left = Vals(1)
     Prog.LShow(IntShow).Top = Vals(2)
     If (Vals(3)) >= 0 And Vals(3) <= 15 Then
      Prog.LShow(IntShow).ForeColor = QBColor(Vals(3))
      Prog.LShow(IntShow).Caption = Vals(0)
     Else
      Err.Number = 5
      GoTo EndIt:
     End If
     For I = Used To H
      If Vars(I) = Vals(0) Then
       Open App.Path & "\" & Vars(I) & ".set" For Input As #2
        Input #2, Vals(0)
        Prog.LShow(IntShow).Caption = Vals(0)
       Close #2
       Used = Used + 1
      Else
       Prog.LShow(IntShow).Caption = Vals(0)
      End If
     Next I
     Prog.LShow(IntShow).Visible = True
    ElseIf (Asc(Vals(0)) >= Asc("A") And Asc(Vals(0)) <= Asc("Z") Or Asc(Vals(0)) >= Asc("a") And Asc(Vals(0)) <= Asc("z")) Or (Asc(Vals(1)) >= Asc("A") And Asc(Vals(1)) <= Asc("Z")) Or (Asc(Vals(1)) >= Asc("a") And Asc(Vals(1)) <= Asc("z")) Or (Asc(Vals(2)) >= Asc("A") And Asc(Vals(2)) <= Asc("Z")) Or (Asc(Vals(2)) >= Asc("a") And Asc(Vals(2)) <= Asc("z")) Or ((Asc(Vals(3)) >= Asc("A") And Asc(Vals(3)) <= Asc("Z")) Or (Asc(Vals(3)) >= Asc("a") And Asc(Vals(3)) <= Asc("z"))) And H >= 0 Then
     'Check If Numbers Exist
     For I = 0 To UBound(Vals)
Nums:
      If Vals(I) <> "" Then
       If Asc(Vals(I)) >= Asc("0") And Asc(Vals(I)) <= Asc("9") Then
        If I = 1 Then
         Prog.LShow(IntShow).Left = Vals(1)
         M = 0
        ElseIf I = 2 Then
         Prog.LShow(IntShow).Top = Vals(2)
         M = 0
        Else
         Prog.LShow(IntShow).ForeColor = QBColor(Vals(3))
         M = 0
        End If
       Else
        For J = Used To H Step -1
         If Vars(J) = Vals(I) Then
          If I < UBound(Vals) Then
           Used = Used + 1
           I = I + 1
           GoTo Nums:
          Else
           Used = Used + 1
           GoTo Variables:
          End If
         End If
        Next J
        If Vars(J + 1) = Vals(I + 1) Then
         Used = Used + 1
         I = I + 1
         GoTo Nums:
        'Else
         'Err.Number = 4
         'GoTo EndIt:
        End If
       End If
      End If
     Next I
Variables:
     'Check If Variables Exist
     For J = 0 To H
      If Vars(J) = Vals(0) Then
       M = M + 1
       FVars(M) = Vals(0)
      End If
      If Vars(J) = Vals(1) Then
       M = M + 1
       FVars(M) = Vals(1)
      End If
      If Vars(J) = Vals(2) Then
       M = M + 1
       FVars(M) = Vals(2)
      End If
      If Vars(J) = Vals(3) Then
       M = M + 1
       FVars(M) = Vals(3)
      End If
     Next J
     If M >= 0 Then
      For I = 0 To M
       On Error Resume Next
       Open App.Path & "\" & FVars(I) & ".set" For Input As #2
        On Error Resume Next
        If FVars(I) = Vals(3) Then
         Input #2, FVars(I)
         If FVars(I) >= 0 And FVars(I) <= 15 Then
          Prog.LShow(IntShow).ForeColor = QBColor(FVars(I))
         Else
          Err.Number = 5
          GoTo EndIt:
         End If
        End If
        If FVars(I) = Vals(0) Then
          Input #2, FVars(I)
          Prog.LShow(IntShow).Caption = FVars(I)
        End If
        If I > 0 And I < M And Vals(I + 1) = Vals(I) Then
         Input #2, FVars(I)
         If FVars(I) <> "" Then
          Prog.LShow(IntShow).Left = FVars(I)
          Prog.LShow(IntShow).Top = FVars(I)
         End If
        End If
        If FVars(I + 1) = Vals(1) Then
         Input #2, FVars(I)
         Prog.LShow(IntShow).Left = FVars(I)
        End If
        If FVars(I) = Vals(2) Then
         Input #2, FVars(I)
         Prog.LShow(IntShow).Top = FVars(I)
        End If
       Close #2
      Next I
      If Prog.LShow(IntShow).Caption = "" Then
       Prog.LShow(IntShow).Caption = Vals(0)
      End If
      Prog.LShow(IntShow).Visible = True
     Else
      Err.Number = 4
      GoTo EndIt:
     End If
    Else
     Err.Number = 5
     GoTo EndIt:
    End If
   '/////////////////////////////////////////
   ElseIf Found(K) = "color" Then
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 1 Or UBound(Vals) > 1 Then: Err.Number = 13: Obj = "Color": GoTo EndIt:
    On Error Resume Next
    If Asc(Vals(0)) >= Asc("0") And Asc(Vals(0)) <= Asc("9") Then
     GoTo ColorOK:
    Else
     If Vals(0) = "" Then
      Err.Number = 13
      Obj = "Color"
      GoTo EndIt:
     Else
ColorOK:
      If Asc(Vals(0)) >= Asc("A") And Asc(Vals(0)) <= Asc("Z") Or Asc(Vals(0)) >= Asc("a") And Asc(Vals(0)) <= Asc("z") And H >= 0 Then
       'Check If This Variable Exists
       For J = Used To H
        If Vars(J) = Vals(0) Then
         'Load Values From Variables
         Open App.Path & "\" & Vals(0) & ".set" For Input As #2
          Input #2, Clr
          If Clr <= 15 And Clr >= 0 Then
           Prog.BackColor = QBColor(Clr)
           M = 0
          Else
           Err.Number = 2
           Obj = "Color"
           Obj2 = "0"
           Obj3 = "15"
           GoTo EndIt:
          End If
         Close #2
         Used = Used + 1
         Exit For
        End If
       Next J
       If M <> 0 Then Err.Number = 4: GoTo EndIt:
      ElseIf Vals(0) <= 15 And Vals(0) >= 0 Then
       Prog.BackColor = QBColor(Vals(0))
      ElseIf H = -1 And Asc(Vals(0)) > Asc("9") Then
       Err.Number = 4
       GoTo EndIt:
      Else
       Err.Number = 2
       Obj = "Color"
       Obj2 = "0"
       Obj3 = "15"
       GoTo EndIt:
      End If
     End If
    End If
   '/////////////////////////////////////////
   ElseIf Found(K) = "mouse" Then
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 1 Or UBound(Vals) > 1 Then: Err.Number = 13: Obj = "Mouse": GoTo EndIt:
    If Vals(0) = "" Then Err.Number = 15: Obj = "Mouse": GoTo EndIt:
    If Asc(Vals(0)) >= Asc("A") And Asc(Vals(0)) <= Asc("Z") Or Asc(Vals(0)) >= Asc("a") And Asc(Vals(0)) <= Asc("z") Then
     'First Checks
     If Vals(0) = "on" Then
      'Change No Mouse To Mouse
      Prog.MousePointer = 0
      ShowCursor (1)
     ElseIf Vals(0) = "off" Then
      'Change Mouse To No Mouse
      Prog.MousePointer = 99
      Prog.MouseIcon = Prog.NoMouse.Picture
      ShowCursor (0)
     ElseIf H >= 0 Then
      'Check If This Variable Exists
      For J = Used To H
       If Vars(J) = Vals(0) Then
        'Load Values From Variables
        Open App.Path & "\" & Vals(0) & ".set" For Input As #2
         Input #2, Vals(0)
         'Same Checks
         If Vals(0) = "on" Then
          'Change No Mouse To Mouse
           Prog.MousePointer = 0
           ShowCursor (1)
         ElseIf Vals(0) = "off" Then
          'Change Mouse To No Mouse
          Prog.MousePointer = 99
          Prog.MouseIcon = Prog.NoMouse.Picture
          ShowCursor (0)
         Else
          Err.Number = 7
          GoTo EndIt:
         End If
        Close #2
        Used = Used + 1
        Exit For
       ElseIf J = H Then
        Err.Number = 7
        GoTo EndIt:
       End If
      Next J
     Else
      Err.Number = 4
      GoTo EndIt:
     End If
    Else
     Err.Number = 4
     GoTo EndIt:
    End If
   '/////////////////////////////////////////
   ElseIf Found(K) = "rnd" Then
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 1 Or UBound(Vals) > 1 Then Err.Number = 13: Obj = "Rnd": GoTo EndIt:
    If Vals(0) = "" Then Err.Number = 15: Obj = "Rnd": GoTo EndIt:
    If Asc(Vals(0)) >= Asc("A") And Asc(Vals(0)) <= Asc("Z") Or Asc(Vals(0)) >= Asc("a") And Asc(Vals(0)) <= Asc("z") Or Asc(Vals(0)) = Asc("-") And H >= 0 Then
     'Check If This Variable Exists
     If Left(Vals(0), 1) = "-" Then
      For J = Used To H
       If Vars(J) = Right(Vals(0), Len(Vals(0)) - 1) Then
        'Load Values From Variables
        Open App.Path & "\" & Vars(J) & ".set" For Input As #2
         Input #2, Vals(0)
         Randomize Timer
         RndNum = Int(Rnd * Abs(Vals(0)))
         M = 0
         Exit For
        Close #2
       End If
      Next J
      If M = -1 Then
       Err.Number = 4
       GoTo EndIt:
      End If
     Else
      For I = 0 To H
       If Vars(I) = Vals(0) Then
        'Load Values From Variables
        Open App.Path & "\" & Vals(0) & ".set" For Input As #2
         Input #2, Vals(0)
         Randomize Timer
         RndNum = Int(Rnd * Vals(0))
         M = 0
        Close #2
        Exit For
       End If
      Next I
      Used = Used + 1
     End If
     If M <> 0 Then Err.Number = 4: GoTo EndIt:
    ElseIf H <= 0 Or H > 0 Then
     On Error Resume Next
     Randomize Timer
     RndNum = Int(Rnd * Vals(0))
     If Err.Number <> 0 Then: Err.Number = 2: Obj = "Rnd": Obj2 = "-?": Obj3 = "+?": GoTo EndIt:
    Else
     Err.Number = 2
     Obj = "Rnd"
     Obj2 = "-?"
     Obj3 = "+?"
     GoTo EndIt:
    End If
   '/////////////////////////////////////////
   ElseIf Found(K) = "absolute" Then
    If YCom(K) = "" Then Err.Number = 13: Obj = "Absolute": GoTo EndIt:
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 1 Or UBound(Vals) > 1 Then Err.Number = 13: Obj = "Absolute": GoTo EndIt:
    If Vals(0) = "" Then Err.Number = 15: Obj = "Absolute": GoTo EndIt:
    On Error GoTo ABSNum:
    If (Asc(Vals(0)) >= Asc("A") And Asc(Vals(0)) <= Asc("Z") Or Asc(Vals(0)) >= Asc("a") And Asc(Vals(0)) <= Asc("z") Or Asc(Vals(0)) = Asc("-")) And H >= 0 Then
     If Left(Vals(0), 1) = "-" Then
      'Check If This Variable Exists
      For J = Used To H
       If Vars(J) = Right(Vals(0), Len(Vals(0)) - 1) Then
        'Load Values From Variables
        Open App.Path & "\" & Vars(J) & ".set" For Input As #2
         Input #2, Vals(0)
         Absol = Abs(Vals(0))
         M = 0
         Exit For
        Close #2
       End If
      Next J
      If M < 0 Then Err.Number = 4: GoTo EndIt:
ABSNum:
     Else
      For J = Used To H
       If Vars(J) = Vals(0) Then
        'Load Values From Variables
        Open App.Path & "\" & Vals(0) & ".set" For Input As #2
         Input #2, Vals(0)
         Absol = Abs(Vals(0))
         M = 0
         Exit For
        Close #2
       End If
      Next J
      If M < 0 Then Err.Number = 4: GoTo EndIt:
     End If
     Used = Used + 1
    ElseIf H < 0 Or (Asc(Right(Vals(0), 1)) < Asc(0) Or Asc(Right(Vals(0), 1)) > Asc(9)) Then
     'This Variable Is Not Acceptable
     Err.Number = 2
     Obj = "Absolute"
     Obj2 = "-?"
     Obj3 = "+?"
     GoTo EndIt:
    Else
     'Simple Number
     Absol = Abs(Vals(0))
    End If
   '/////////////////////////////////////////
   ElseIf Found(K) = "add" Then
    If YCom(K) = "" Then Err.Number = 13: Obj = "Add": GoTo EndIt:
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 2 Then Err.Number = 13: Obj = "Add": GoTo EndIt:
    For I = 0 To UBound(Vals) - 1
     If Vals(I) = "" Then Err.Number = 2: Obj = "Add": Obj2 = "-?": Obj3 = "+?": GoTo EndIt:
     If Asc(Vals(I)) >= Asc("0") And Asc(Vals(I)) <= Asc("9") Or Asc(Vals(I)) <= Asc("0") And Asc(Vals(I)) >= Asc("-9") Then
      'Big Check For String Into Number
      For V = 1 To Len(Vals(I))
       If Right(Left(Vals(I), V), 1) > Asc("9") And Right(Left(Vals(I), V), 1) = Asc("-") Then
        Err.Number = 4
        GoTo EndIt:
       End If
      Next V
      Common = Common + Vals(I)
     ElseIf Asc(Vals(I)) >= Asc("A") And Asc(Vals(I)) <= Asc("Z") Or Asc(Vals(I)) >= Asc("a") And Asc(Vals(I)) <= Asc("z") And H >= 0 Then
      'Check If This Variable Exists
      For J = 0 To H
       If Vars(J) = Vals(I) Then
        M = M + 1
        FVars(M) = Vals(I)
        Used = Used + 1
        Exit For
       End If
      Next J
      If M = 0 Then
       'Get Previous Value
       Open App.Path & "\" & FVars(0) & ".set" For Input As #1
        Input #1, PVal
       Close #1
       'New Common
       Common = Common + PVal
       'Save Common To Variable
       Open App.Path & "\" & FVars(0) & ".set" For Output As #1
        Print #1, Common
       Close #1
      ElseIf M > 0 Then
       For N = Used - 1 To M
        'Get Previous Value
        Open App.Path & "\" & FVars(N) & ".set" For Input As #1
         Input #1, PVal
        Close #1
        'New Common
        Common = Common + PVal
        'Save Common To Variable
        Open App.Path & "\" & FVars(N) & ".set" For Output As #1
         Print #1, Common
        Close #1
       Next N
      Else
       Err.Number = 4
       GoTo EndIt:
      End If
     Else
      If H < 0 Then
       Err.Number = 4
      Else
       Err.Number = 2
       Obj = "Add"
       Obj2 = "-?"
       Obj3 = "+?"
      End If
      GoTo EndIt:
     End If
    Next I
    If M = -1 Then Err.Number = 4: GoTo EndIt:
   '/////////////////////////////////////////
   ElseIf Found(K) = "sub" Then
    If YCom(K) = "" Then Err.Number = 13: Obj = "Sub": GoTo EndIt:
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 2 Then Err.Number = 13: Obj = "Sub": GoTo EndIt:
    For I = 0 To UBound(Vals) - 1
     If Vals(I) = "" Then Err.Number = 2: Obj = "Sub": Obj2 = "-?": Obj3 = "+?": GoTo EndIt:
     If Asc(Vals(I)) >= Asc("0") And Asc(Vals(I)) <= Asc("9") Or Asc(Vals(I)) <= Asc("0") And Asc(Vals(I)) >= Asc("-9") Then
      'Big Check For String Into Number
      For V = 1 To Len(Vals(I))
       If Right(Left(Vals(I), V), 1) > Asc("9") And Right(Left(Vals(I), V), 1) = Asc("-") Then
        Err.Number = 4
        GoTo EndIt:
       End If
      Next V
      If Common = 0 Then
       Common = Vals(I)
      Else
       Common = Common - Vals(I)
      End If
     ElseIf Asc(Vals(I)) >= Asc("A") And Asc(Vals(I)) <= Asc("Z") Or Asc(Vals(I)) >= Asc("a") And Asc(Vals(I)) <= Asc("z") And H >= 0 Then
      'Check If This Variable Exists
      For J = 0 To H
       If Vars(J) = Vals(I) Then
        M = M + 1
        FVars(M) = Vals(I)
        Used = Used + 1
        Exit For
       End If
      Next J
      If M = 0 Then
       'Get Previous Value
       Open App.Path & "\" & FVars(0) & ".set" For Input As #1
        Input #1, PVal
       Close #1
       'New Common
       Common = PVal - Common
       'Save Common To Variable
       Open App.Path & "\" & FVars(0) & ".set" For Output As #1
        Print #1, Common
       Close #1
      ElseIf M > 0 Then
       For N = Used - 1 To M
        'Get Previous Value
        Open App.Path & "\" & FVars(N) & ".set" For Input As #1
         Input #1, PVal
        Close #1
        'New Common
        Common = PVal - Common
        'Save Common To Variable
        Open App.Path & "\" & FVars(N) & ".set" For Output As #1
         Print #1, Common
        Close #1
       Next N
      Else
       Err.Number = 4
       GoTo EndIt:
      End If
     Else
      If H < 0 Then
       Err.Number = 4
      Else
       Err.Number = 2
       Obj = "Sub"
       Obj2 = "-?"
       Obj3 = "+?"
      End If
      GoTo EndIt:
     End If
    Next I
    If M = -1 Then Err.Number = 4: GoTo EndIt:
   '/////////////////////////////////////////
   ElseIf Found(K) = "mul" Then
    If YCom(K) = "" Then Err.Number = 13: Obj = "Mul": GoTo EndIt:
    'Extra Initial Values
    Common = 1
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 2 Then Err.Number = 13: Obj = "Mul": GoTo EndIt:
    For I = 0 To UBound(Vals) - 1
     If Vals(I) = "" Then Err.Number = 2: Obj = "Mul": Obj2 = "-?": Obj3 = "+?": GoTo EndIt:
     If Asc(Vals(I)) >= Asc("0") And Asc(Vals(I)) <= Asc("9") Or Asc(Vals(I)) <= Asc("0") And Asc(Vals(I)) >= Asc("-9") Then
      'Big Check For String Into Number
      For V = 1 To Len(Vals(I))
       If Right(Left(Vals(I), V), 1) > Asc("9") And Right(Left(Vals(I), V), 1) = Asc("-") Then
        Err.Number = 4
        GoTo EndIt:
       End If
      Next V
      Common = Common * Vals(I)
     ElseIf Asc(Vals(I)) >= Asc("A") And Asc(Vals(I)) <= Asc("Z") Or Asc(Vals(I)) >= Asc("a") And Asc(Vals(I)) <= Asc("z") And H >= 0 Then
      'Check If This Variable Exists
      For J = 0 To H
       If Vars(J) = Vals(I) Then
        M = M + 1
        FVars(M) = Vals(I)
        Used = Used + 1
        Exit For
       End If
      Next J
      If M = 0 Then
       'Get Previous Value
       Open App.Path & "\" & FVars(0) & ".set" For Input As #1
        Input #1, PVal
       Close #1
       'New Common
       Common = Common * PVal
       'Save Common To Variable
       Open App.Path & "\" & FVars(0) & ".set" For Output As #1
        Print #1, Common
       Close #1
      ElseIf M > 0 Then
       For N = Used - 1 To M
        'Get Previous Value
        Open App.Path & "\" & FVars(N) & ".set" For Input As #1
         Input #1, PVal
        Close #1
        'New Common
        Common = Common * PVal
        'Save Common To Variable
        Open App.Path & "\" & FVars(N) & ".set" For Output As #1
         Print #1, Common
        Close #1
       Next N
      Else
       Err.Number = 4
       GoTo EndIt:
      End If
     Else
      If H < 0 Then
       Err.Number = 4
      Else
       Err.Number = 2
       Obj = "Mul"
       Obj2 = "-?"
       Obj3 = "+?"
      End If
      GoTo EndIt:
     End If
    Next I
    If M = -1 Then Err.Number = 4: GoTo EndIt:
   '/////////////////////////////////////////
   ElseIf Found(K) = "div" Then
    If YCom(K) = "" Then Err.Number = 13: Obj = "Div": GoTo EndIt:
    'Extra Initial Values
    Common = 1
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 2 Then Err.Number = 13: Obj = "Div": GoTo EndIt:
    For I = 0 To UBound(Vals) - 1
     If Vals(I) = "" Then Err.Number = 2: Obj = "Div": Obj2 = "-?": Obj3 = "+?": GoTo EndIt:
     If Asc(Vals(I)) >= Asc("0") And Asc(Vals(I)) <= Asc("9") Or Asc(Vals(I)) <= Asc("0") And Asc(Vals(I)) >= Asc("-9") Then
      'Big Check For String Into Number
      For V = 1 To Len(Vals(I))
       If Right(Left(Vals(I), V), 1) > Asc("9") And Right(Left(Vals(I), V), 1) = Asc("-") Then
        Err.Number = 4
        GoTo EndIt:
       End If
      Next V
      If Vals(I) = "0" And Ready = True Then
       Err.Number = 6
       GoTo EndIt:
      Else
       If Common = 0 Then
        Common = Vals(I)
       Else
        Common = Common / Vals(I)
       End If
      End If
      Ready = True
     ElseIf Asc(Vals(I)) >= Asc("A") And Asc(Vals(I)) <= Asc("Z") Or Asc(Vals(I)) >= Asc("a") And Asc(Vals(I)) <= Asc("z") And H >= 0 Then
      'Check If This Variable Exists
      For J = 0 To H
       If Vars(J) = Vals(I) Then
        M = M + 1
        FVars(M) = Vals(I)
        Used = Used + 1
        Exit For
       End If
      Next J
      If M = 0 Then
       'Get Previous Value
       Open App.Path & "\" & FVars(0) & ".set" For Input As #1
        Input #1, PVal
       Close #1
       'New Common
       Common = PVal / Common
       'Save Common To Variable
       Open App.Path & "\" & FVars(0) & ".set" For Output As #1
        Print #1, Common
       Close #1
      ElseIf M > 0 Then
       For N = Used - 1 To M
        'Get Previous Value
        Open App.Path & "\" & FVars(N) & ".set" For Input As #1
         Input #1, PVal
        Close #1
        'New Common
        Common = PVal / Common
        'Save Common To Variable
        Open App.Path & "\" & FVars(N) & ".set" For Output As #1
         Print #1, Common
        Close #1
       Next N
      Else
       Err.Number = 4
       GoTo EndIt:
      End If
     Else
      If H < 0 Then
       Err.Number = 4
      Else
       Err.Number = 2
       Obj = "Div"
       Obj2 = "-?"
       Obj3 = "+?"
      End If
      GoTo EndIt:
     End If
    Next I
    If M = -1 Then Err.Number = 4: GoTo EndIt:
   '/////////////////////////////////////////
   ElseIf Found(K) = "sdiv" Then
    If YCom(K) = "" Then Err.Number = 13: Obj = "SDiv": GoTo EndIt:
    'Extra Initial Values
    Common = 1
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 2 Then Err.Number = 13: Obj = "SDiv": GoTo EndIt:
    For I = 0 To UBound(Vals) - 1
     If Vals(I) = "" Then Err.Number = 2: Obj = "SDiv": Obj2 = "-?": Obj3 = "+?": GoTo EndIt:
     If Asc(Vals(I)) >= Asc("0") And Asc(Vals(I)) <= Asc("9") Or Asc(Vals(I)) <= Asc("0") And Asc(Vals(I)) >= Asc("-9") Then
      'Big Check For String Into Number
      For V = 1 To Len(Vals(I))
       If Right(Left(Vals(I), V), 1) > Asc("9") And Right(Left(Vals(I), V), 1) = Asc("-") Then
        Err.Number = 4
        GoTo EndIt:
       End If
      Next V
      If Vals(I) = "0" And Ready = True Then
       Err.Number = 6
       GoTo EndIt:
      Else
       If Common = 0 Then
        Common = Vals(I)
       Else
        Common = Int(Common / Vals(I))
       End If
      End If
      Ready = True
     ElseIf Asc(Vals(I)) >= Asc("A") And Asc(Vals(I)) <= Asc("Z") Or Asc(Vals(I)) >= Asc("a") And Asc(Vals(I)) <= Asc("z") And H >= 0 Then
      'Check If This Variable Exists
      For J = 0 To H
       If Vars(J) = Vals(I) Then
        M = M + 1
        FVars(M) = Vals(I)
        Used = Used + 1
        Exit For
       End If
      Next J
      If M = 0 Then
       'Get Previous Value
       Open App.Path & "\" & FVars(0) & ".set" For Input As #1
        Input #1, PVal
       Close #1
       'New Common
       Common = Int(PVal / Common)
       'Save Common To Variable
       Open App.Path & "\" & FVars(0) & ".set" For Output As #1
        Print #1, Common
       Close #1
      ElseIf M > 0 Then
       For N = Used - 1 To M
        'Get Previous Value
        Open App.Path & "\" & FVars(N) & ".set" For Input As #1
         Input #1, PVal
        Close #1
        'New Common
        Common = Int(PVal / Common)
        'Save Common To Variable
        Open App.Path & "\" & FVars(N) & ".set" For Output As #1
         Print #1, Common
        Close #1
       Next N
      Else
       Err.Number = 4
       GoTo EndIt:
      End If
     Else
      If H < 0 Then
       Err.Number = 4
      Else
       Err.Number = 2
       Obj = "SDiv"
       Obj2 = "-?"
       Obj3 = "+?"
      End If
      GoTo EndIt:
     End If
    Next I
    If M = -1 Then Err.Number = 4: GoTo EndIt:
   '/////////////////////////////////////////
   ElseIf Found(K) = "=" Then 'Equals
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) <= 1 Or UBound(Vals) > 2 Then: Err.Number = 13: Obj = "= (Equals)": GoTo EndIt:
    If Vals(0) = "" Or Vals(1) = "" Then Err.Number = 15: Obj = "= (Equals)": GoTo EndIt:
    For I = 0 To UBound(Vals) - 1
     If Asc(Left(Vals(I), Len(Vals(I)))) >= Asc("0") And Asc(Left(Vals(I), Len(Vals(I)))) <= Asc("9") Then
      Vals(1) = Vals(0)
     ElseIf Asc(Left(Vals(I), Len(Vals(I)))) >= Asc("A") And Asc(Left(Vals(I), Len(Vals(I)))) <= Asc("Z") Or Asc(Left(Vals(I), Len(Vals(I)))) >= Asc("a") And Asc(Left(Vals(I), Len(Vals(I)))) <= Asc("z") And H >= 0 Then
      'Check If This Variable Exists
      For J = 0 To H
       If Vars(J) = Vals(I) Then
        M = M + 1
        FVars(M) = Vals(I)
        Used = Used + 1
        Exit For
       End If
      Next J
      If M = 0 Then
       'Get Previous Value
       Open App.Path & "\" & FVars(0) & ".set" For Input As #1
        Input #1, PVal
       Close #1
       'Save Common To Variable
       Open App.Path & "\" & FVars(0) & ".set" For Output As #1
        Print #1, PVal
       Close #1
      ElseIf M > 0 Then
       For N = Used - 1 To M
        'Get Previous Value
        Open App.Path & "\" & FVars(N) & ".set" For Input As #1
         Input #1, PVal
        Close #1
        'New Common
        Common = PVal
        'Save Common To Variable
        Open App.Path & "\" & FVars(N) & ".set" For Output As #1
         Print #1, Common
        Close #1
       Next N
      Else
       Err.Number = 4
       GoTo EndIt:
      End If
     Else
      Err.Number = 2
      Obj = "= (Equals)"
      Obj2 = "-?"
      Obj3 = "+?"
      GoTo EndIt:
     End If
    Next I
   '/////////////////////////////////////////
   ElseIf Found(K) = "reset" Then
    If Right(YCom(K), 1) = "\" Then: Err.Number = 13: Obj = "ReSet": GoTo EndIt:
    'ReSet Variables
    On Error Resume Next
    For I = 0 To VarsToDel
     Kill App.Path & "\" & Vars(I) & ".set"
    Next I
   '/////////////////////////////////////////
   ElseIf Found(K) = "counter" Then
    'Values
    Vals() = Split(YCom(K), "\")
    If UBound(Vals) < 1 Or UBound(Vals) > 1 Then: Err.Number = 13: Obj = "Counter": GoTo EndIt:
    If Vals(0) = "" Then Err.Number = 15: Obj = "Counter": GoTo EndIt:
    'Loops Counter
    If Asc(Vals(0)) >= Asc("A") And Asc(Vals(0)) <= Asc("Z") Or Asc(Vals(0)) >= Asc("a") And Asc(Vals(0)) <= Asc("z") Then
     'Check If This Variable Exists
     For J = 0 To H
      If Vars(J) = Vals(0) Then
       M = M + 1
       FVars(M) = Vals(0)
       Used = Used + 1
       Exit For
      End If
     Next J
     If M < 0 Then
      Err.Number = 4
      GoTo EndIt:
     Else
      'Save Counter Value To Variable
      Open App.Path & "\" & FVars(0) & ".set" For Output As #1
       Print #1, Times
      Close #1
     End If
    Else
     Err.Number = 4
     GoTo EndIt:
    End If
   '/////////////////////////////////////////
   ElseIf Found(K) = "help" And HELPTrace = 0 Then
    If Right(YCom(K), 1) = "\" Then: Err.Number = 13: Obj = "Help": GoTo EndIt:
    'Clear Previous Help File
    THelp.Text1.Text = ""
    'Show Help File
    On Error GoTo FuckErr:
    Open App.Path & "\THelp.thp" For Input As #2
     While Not EOF(2)
      Line Input #2, File
      THelp.Text1.Text = THelp.Text1.Text + File
      THelp.Text1.Text = THelp.Text1.Text & vbCrLf
     Wend
    Close #2
    THelp.Show
    HELPTrace = 1
    GoTo Cool:
FuckErr:
    Err.Number = 1
    GoTo EndIt:
Cool:
   '/////////////////////////////////////////
   ElseIf Found(K) = "exit" And EXITrace = 0 Then
    If Right(YCom(K), 1) = "\" Then: Err.Number = 13: Obj = "Exit": GoTo EndIt:
    'Exit Program
    Unload Prog
    EXITrace = 1
   '/////////////////////////////////////////
   ElseIf Found(K) = "end" Then
    If Right(YCom(K), 1) = "\" Then: Err.Number = 13: Obj = "End": GoTo EndIt:
    If EndCom <> "end" Then Err.Number = 11: GoTo EndIt:
    'Go To Previous Loop If This Succesfully Completed
    If EndCom = "end" And (FOFL <> 0 And LOFL <> 0 And OtherL > 0) Then
     FTime = FOFL
     LTime = LOFL
     OtherL = OtherL - 1
     GoTo Looping:
    'Last Check To Continue With Next Commands
    ElseIf EndCom = "end" And Times = LTime Then
     EndCom = ""
     Looper = 0
     nTime = Founds - K
     nTime = Founds - nTime + 1
     GoTo Coms:
    End If
   '/////////////////////////////////////////
   
   End If
  Next K
  If Looper = 0 Then GoTo EndIt:
Next Times
'End Of Main Check
EndIt:
 'BUGS
 If Err.Number <> 0 And Err.Number <> 1 And Err.Number <> 2 And Err.Number <> 3 And Err.Number <> 4 And Err.Number <> 5 And Err.Number <> 6 And Err.Number <> 7 And Err.Number <> 8 And Err.Number <> 9 And Err.Number <> 10 And Err.Number <> 11 And Err.Number <> 12 And Err.Number <> 13 And Err.Number <> 14 And Err.Number <> 15 And Err.Number <> 340 And Err.Number <> 92 And Err.Number <> 76 Then
  Text2.Text = "Error:" & Err.Number & " " & Err.Description
 ElseIf Err.Number = 1 Then
  Text2.Text = "Error:" & Err.Number & " No Help File Found In Taurus Directory."
 ElseIf Err.Number = 2 Then
  Text2.Text = "Error:" & Err.Number & " Invalid Number For " & Obj & ".Number Must Be Between " & Obj2 & " And " & Obj3 & "."
 ElseIf Err.Number = 3 Then
  Text2.Text = "Error:" & Err.Number & " Invalid Set.You Can Set Only Letters Or Words With Numeric Charachters But The Numeric Charachters Should Not Be In The Begining Of The Variable."
 ElseIf Err.Number = 4 Then
  Text2.Text = "Error:" & Err.Number & " Invalid Variable.This Variable Does Not Exist.Please Choose Another Or Set A New One."
 ElseIf Err.Number = 5 Then
  Text2.Text = "Error:" & Err.Number & " Invalid Show.Show Command Needs One String,Two Positive Variables Or Numbers For The Position On Screen And One Positive Variable Or Number (0-15) For Drawing Color."
 ElseIf Err.Number = 6 Then
  Text2.Text = "Error:" & Err.Number & " Invalid Division.Division With Zero Is Not Acceptable."
 ElseIf Err.Number = 7 Then
  Text2.Text = "Error:" & Err.Number & " Invalid Mouse Value.Mouse Values Are On Or Off."
 ElseIf Err.Number = 8 Then
  Text2.Text = "Error:" & Err.Number & " Invalid Put.Put Needs A Variable To Add A Value Or A String."
 ElseIf Err.Number = 340 And Found(K) = "cs" Then
  Text2.Text = "Error:" & "17" & " You Had An Instant Exit.The Screen Could Not Clear."
 ElseIf Err.Number = 340 Then
  Text2.Text = "Error:" & "9" & " Too Big String."
 ElseIf Err.Number = 10 Then
  Text2.Text = "Error:" & Err.Number & " Invalid Loop.Loop Command Needs Two Values.One To Know Where To Start And One To Know Where To Finish."
 ElseIf Err.Number = 11 Then
  Text2.Text = "Error:" & Err.Number & " End Without Loop."
 ElseIf Err.Number = 12 Then
  Text2.Text = "Error:" & Err.Number & " Loop Without End."
 ElseIf Err.Number = 13 Then
  Text2.Text = "Error:" & Err.Number & " No Suitable Number Of Slashes For " & Obj & " Command."
 ElseIf Err.Number = 14 Then
  Text2.Text = "Error:" & Err.Number & " ."
 ElseIf Err.Number = 15 Then
  Text2.Text = "Error:" & Err.Number & " No Suitable Variables For " & Obj & " Command."
 ElseIf Err.Number = 76 Then
  Text2.Text = "Error:" & "16" & " '\' Is Not A Valid Charachter To Set As Variable Or To Be Contained In It."
 End If
 'Process Viewer Messages
 If C = 1 Then
  'Activities For BUGS
  If Err.Number <> 0 And Err.Number <> 92 Then
   'Exit Program
   Unload Prog
   Label9.Caption = "Interpretation...FAILED"
  Else
   'Successfull Interpretation
   Text2.Text = "Success"
   Label9.Caption = "Interpretation...OK"
  End If
  Timer1.Enabled = True
 Else
  'No Commands
  Text2.Text = "No Commands Found"
  Label9.Caption = "Interpretation...FAILED"
  Timer1.Enabled = True
 End If
End Sub
'Editor
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 'Main Save To Memory
 If KeyCode = vbKeyReturn And Label1.Caption < MaxLines Then
  Pinax(U) = Text1.Text
  Lnh(U) = Len(Pinax(U))
  If U = 0 Then
   Num(U) = Lnh(U)
   SCode(U) = Pinax(U)
  ElseIf U > 0 And U <= 1 Then
   Num(U) = Abs(Lnh(U) - Lnh(U - 1))
   Txt = Right(Pinax(U - 1), Lnh(U - 1))
   Txt2 = Right(Pinax(U), Num(U) - 1)
   SCode(U - 1) = Txt
   SCode(U) = Txt2
  ElseIf U > 1 Then
   Num(U) = Abs(Lnh(U) - Lnh(U - 1))
   Txt3 = Txt2
   Txt2 = Right(Pinax(U), Num(U) - 1)
   SCode(U - 1) = Txt3
   SCode(U) = Txt2
  End If
  U = U + 1
  Label1.Caption = Label1.Caption + 1
 ElseIf KeyCode = vbKeyReturn And Label1.Caption = MaxLines Then
  Text1.Locked = True
  Text2.Text = "WARNING! Max Number Of Lines.You Have Used MaxLines Lines Of Code.You Do Not Have Other Space To Use.Please Save Your Source File And Start A New Project."
 ElseIf KeyCode = vbKeyBack And Label1.Caption > 1 And Label1.Caption <= MaxLines Then
  Pinax(U) = Text1.Text & "¥"
  Lnh(U) = Len(Pinax(U))
  If U = 0 Then
   Num(U) = Lnh(U)
   SCode(U) = Pinax(U)
  ElseIf U > 0 And U <= 1 Then
   Num(U) = Abs(Lnh(U) - Lnh(U - 1))
   Txt = Right(Pinax(U - 1), Lnh(U - 1))
   Txt2 = Right(Pinax(U), Abs(Num(U) - 1))
   SCode(U - 1) = Txt
   SCode(U) = Txt2
  ElseIf U > 1 Then
   Num(U) = Abs(Lnh(U) - Lnh(U - 1))
   Txt3 = Txt2
   Txt2 = Right(Pinax(U), Num(U) - 1)
   SCode(U - 1) = Txt3
   SCode(U) = Txt2
  End If
  If Num(U) <= 3 Then
   U = U - 1
   Label1.Caption = Label1.Caption - 1
  End If
 End If
End Sub
'Load File
Private Sub Command2_Click()
 Me.Enabled = False
 TLoadSave.Caption = "Load Source File"
 TLoadSave.Label2.Caption = "LOAD FILE"
 TLoadSave.Show
End Sub
'Save File
Private Sub Command3_Click()
 Me.Enabled = False
 TLoadSave.Show
End Sub
'Make
Private Sub Command4_Click()
 'Make Special Executable
 
End Sub
'Open Help System
Private Sub Command5_Click()
 'Close All Open Files
 Close
 'Clear Previous Help File
 THelp.Text1.Text = ""
 'Show Help File
 On Error GoTo FuckIt:
 Open App.Path & "\THelp.thp" For Input As #2
  While Not EOF(2)
   Line Input #2, File
   THelp.Text1.Text = THelp.Text1.Text + File
   THelp.Text1.Text = THelp.Text1.Text & vbCrLf
  Wend
 Close #2
 THelp.Show
 HELPTrace = 1
 Text2.Text = ""
 GoTo OK:
FuckIt:
  Text2.Text = "Error:" & 1 & " No Help File Found In Taurus Directory."
OK:
End Sub
'End
Private Sub Form_Unload(Cancel As Integer)
 Cancel = 1
 'Close All Open Files
 Close
 'Clear All Variables
 On Error Resume Next
 For N = 0 To VarsToDel
  Kill App.Path & "\" & Vars(N) & ".set"
 Next N
 'Show Confirmation Dialog
 Dialog.Show
 Me.Enabled = False
End Sub
'Timer
Private Sub Timer1_Timer()
 'Clear Previous Messages
 Label9.Caption = ""
 Timer1.Enabled = False
End Sub
