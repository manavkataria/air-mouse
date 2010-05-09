VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8220
   ClientLeft      =   7320
   ClientTop       =   2160
   ClientWidth     =   10740
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
   ScaleHeight     =   8220
   ScaleWidth      =   10740
   Begin VB.Frame frame1 
      Caption         =   "Display"
      Height          =   3255
      Left            =   8040
      TabIndex        =   7
      Top             =   960
      Width           =   1455
      Begin VB.CheckBox chkz 
         Caption         =   "Z-Input"
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chky 
         Caption         =   "Y-Input"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkx 
         Caption         =   "X-Input"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Graph"
      Height          =   1335
      Left            =   1920
      TabIndex        =   1
      Top             =   5880
      Width           =   5775
      Begin VB.Label lblfnz 
         Alignment       =   2  'Center
         Caption         =   "Z"
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
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
         Caption         =   "X"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblfny 
         Alignment       =   2  'Center
         Caption         =   "Y"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4575
      Left            =   1800
      ScaleHeight     =   4515
      ScaleMode       =   0  'User
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   840
      Width           =   5895
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

Public Sub plotxyz(ByVal X As Single, ByVal Y As Single, ByVal z As Single)
    'Picture1.PSet (msCnt, x), vbRed
    'Picture1.PSet (msCnt, y), vbBlue
    'Picture1.PSet (msCnt, z), vbGreen
    Static prevX, prevY, prevZ As Single
    
    If chkx.Value = 1 Then Picture1.Line (msCnt, prevX)-(msCnt, X), vbRed
    If chky.Value = 1 Then Picture1.Line (msCnt, prevY)-(msCnt, Y), vbBlue
    If chkz.Value = 1 Then Picture1.Line (msCnt, prevZ)-(msCnt, z), vbGreen
    
    prevX = X
    prevY = Y
    prevZ = z
    
    lblmsCnt.Caption = "msCnt: " & Mid(CStr(msCnt), 1, 5)
    lblfnx.Caption = "x: " & CStr(X)
    lblfny.Caption = "y: " & CStr(Y)
    lblfnz.Caption = "z: " & CStr(z)
    
End Sub

Private Sub cmdExit_Click()
    Call Form1.cmdGraph_Click
End Sub

Private Sub Form_Load()
    Picture1.AutoRedraw = True
    Call setAxes
End Sub

