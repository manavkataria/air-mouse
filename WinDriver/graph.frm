VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13275
   LinkTopic       =   "Form2"
   ScaleHeight     =   9000
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Graph"
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4815
      Left            =   2640
      ScaleHeight     =   4755
      ScaleWidth      =   8355
      TabIndex        =   1
      Top             =   2880
      Width           =   8415
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Accelerometer Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
Dim x As Single
Dim y As Single

    Picture1.Scale (-10, 10)-(10, -10)
    
    Picture1.ForeColor = vbBlack
    ' Draw X axis.
    Picture1.Line (-10, 0)-(10, 0)
    For i = -9 To 9
        Picture1.Line (i, -0.5)-(i, 0.5)
    Next i

    ' Draw Y axis.
    Picture1.Line (0, -10)-(0, 10)
    For i = -9 To 9
        Picture1.Line (-0.5, i)-(0.5, i)
    Next i
    
    ' Draw y = 4 * sin(x).
    Picture1.ForeColor = vbRed
    x = -10
    y = 4 * Sin(x)
    Picture1.CurrentX = x
    Picture1.CurrentY = y
    For x = -10 To 10 Step 0.25
        y = 4 * Sin(x)
        Picture1.Line -(x, y)
    Next x
    
    ' Draw y = x ^ 3 / 5 - 3 * x + 1.
    Picture1.ForeColor = vbBlue
    x = -10
    y = x ^ 3 / 5 - 3 * x + 1
    Picture1.CurrentX = x
    Picture1.CurrentY = y
    For x = -10 To 10 Step 0.25
        y = x ^ 3 / 5 - 3 * x + 1
        Picture1.Line -(x, y)
    Next x
End Sub

Private Sub Form_Load()

End Sub
