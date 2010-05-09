VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8235
   ClientLeft      =   7320
   ClientTop       =   2160
   ClientWidth     =   9810
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   9810
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Graph"
      Height          =   1335
      Left            =   2160
      TabIndex        =   1
      Top             =   6120
      Width           =   5415
      Begin VB.Label lblmsCnt 
         Alignment       =   2  'Center
         Caption         =   "msCnt"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblfnx 
         Alignment       =   2  'Center
         Caption         =   "fn"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblfny 
         Alignment       =   2  'Center
         Caption         =   "fn"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4575
      Left            =   2040
      ScaleHeight     =   4515
      ScaleMode       =   0  'User
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   960
      Width           =   5535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GRAPH_HEIGHT = 50
Private Const GRAPH_WIDTH = 5

Dim msCnt As Single

Public Sub ScaleAxes()
    Picture1.ScaleMode = vbUser
    Picture1.Scale (0, (GRAPH_HEIGHT / 2))-(GRAPH_WIDTH, -(GRAPH_HEIGHT / 2))
End Sub

'Scales and Draws the Axes
Public Sub DrawAxes()
Dim i As Integer
    
    ' Draw X axis.
    Picture1.ForeColor = vbBlack
    Picture1.Line (0, 0)-(Picture1.ScaleWidth, 0)
    For i = 0 To Picture1.ScaleWidth
        Picture1.Line (i, -0.5)-(i, 0.5)
    Next i

    ' Draw Y axis.
    Picture1.Line (0, Picture1.ScaleTop)-(0, Picture1.ScaleTop - Picture1.ScaleHeight)
    For i = Picture1.ScaleTop To (Picture1.ScaleTop - Picture1.ScaleHeight) Step -1
        'Picture1.Line (-0.5, i)-(0.5, i)
        Picture1.Line (-2, i)-(2, i)
    Next i
End Sub

Public Sub setAxes()
    ScaleAxes
    DrawAxes
End Sub

Public Sub ResetGraph()
    msCnt = msCnt + 0.01
    
    If (Picture1.CurrentX > Picture1.ScaleWidth) Or (Picture1.CurrentY > -Picture1.ScaleHeight) Then
        Picture1.CurrentX = 0
        Picture1.CurrentY = 0
        msCnt = 0
        Picture1.Cls
        Call DrawAxes
    End If

End Sub

Public Sub plotxy(ByVal x As Single, ByVal y As Single)
    'Picture1.PSet (msCnt, x), vbRed
    'Picture1.PSet (msCnt, y), vbBlue
    Static prevX, prevY As Single
    
    Picture1.Line (msCnt, prevX)-(msCnt, x), vbRed
    Picture1.Line (msCnt, prevY)-(msCnt, y), vbBlue
    prevX = x
    prevY = y
    
    lblmsCnt.Caption = "msCnt: " & Mid(CStr(msCnt), 1, 5)
    lblfnx.Caption = "x: " & CStr(x)
    lblfny.Caption = "y: " & CStr(y)
    
End Sub

Private Sub cmdExit_Click()
    Call Form1.cmdGraph_Click
End Sub

Private Sub Form_Load()
    Picture1.AutoRedraw = True
    Call setAxes
End Sub

