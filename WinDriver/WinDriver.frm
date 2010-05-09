VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   10080
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   10080
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox RXtxt 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton RTSBtn 
      Caption         =   "RTS Toggle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   2040
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Accelerometer Mouse Windows Driver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3105
      TabIndex        =   0
      Top             =   480
      Width           =   7365
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo handler

    MSComm.RTSEnable = False
    
    MSComm.CommPort = 1
    MSComm.PortOpen = True
    
    Exit Sub

handler:
    MsgBox Err.Description
    Exit Sub
    
End Sub

Private Sub RTSBtn_Click()
    MSComm.RTSEnable = Not MSComm.RTSEnable
    RTSBtn.FontBold = Not RTSBtn.FontBold
End Sub

Private Sub MSComm_oncomm()
    Me.RXtxt.Text = Me.RXtxt.Text & Asc(MSComm.Input) & " " '& vbNewLine
    
    '    SetCursorPos pt.X, pt.Y

    Select Case Me.MSComm.CommEvent
        Case comEvRecieve
            Dim buffer As Byte
            buffer = MSComm.Input
            
            Me.RXtxt.Text = Me.RXtxt.Text & "m" & Chr$(buffer) & Chr$(13)
    End Select
End Sub

Private Sub form_unload(Cancel As Integer)
    MSComm.PortOpen = False
End Sub

Private Sub cmdClear_Click()
    Me.Cls
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Circle (X, Y), 10
End Sub
