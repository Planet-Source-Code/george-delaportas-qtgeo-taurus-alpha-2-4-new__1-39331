Attribute VB_Name = "Modules"
'Modules

'Declarations Section
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Integer) As Integer
'Dim Section
Public Const MaxLines = 10000 'Maximum Lines
Public Buffer As String
Public Vars(MaxLines) As String

'TAURUS Extensions

'.tef : Taurus Edit File (Load/Save)
'.tcf : Taurus Code File (Taurus Processes)
'.tdb : Taurus DataBase (Taurus Processes)
