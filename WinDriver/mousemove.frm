VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form2"
   ScaleHeight     =   5625
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "MoveUP"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Right Click"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Left Click"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'See: http://www.experts-exchange.com/Programming/Misc/Q_21597376.html
'http://msdn.microsoft.com/en-us/library/ms646260(VS.85).aspx
'http://www.vbforums.com/archive/index.php/t-15301.html

'pending:
'http://www.autohotkey.com/docs/misc/SendMessage.htm

Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, _
    ByVal Y As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
    ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, _
    ByVal dwExtraInfo As Long)
        

Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
'Private Const MOUSEEVENTF_ABSOLUTE = &H8000

Public Sub LeftMouseClick(ByVal X As Long, ByVal Y As Long)
    'SetCursorPos x, y
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Public Sub RightMouseClick(ByVal X As Long, ByVal Y As Long)
    'SetCursorPos x, y
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

Public Sub MouseMove(ByVal X As Long, ByVal Y As Long)
    'SetCursorPos x, y
    mouse_event MOUSEEVENTF_MOVE, X, Y, 0, 0
    'mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub


Private Sub Command2_Click()
    RightMouseClick 0, 0
End Sub

Private Sub Command5_Click()
    MouseMove 10, 10
End Sub

Private Sub Command1_Click()
    LeftMouseClick 0, 0
End Sub

