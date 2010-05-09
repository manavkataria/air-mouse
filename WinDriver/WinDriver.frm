VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
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
   ScaleHeight     =   8010
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMore 
      Caption         =   "More >>"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   3840
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
      Left            =   7320
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   1920
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   38400
      InputMode       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mouse Information"
      Height          =   1575
      Left            =   1320
      TabIndex        =   3
      Top             =   6120
      Width           =   8175
      Begin VB.TextBox RXtxt 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Text            =   "RX Data"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblLeft 
         Alignment       =   2  'Center
         Caption         =   "   Left    Click"
         Height          =   495
         Left            =   5280
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMiddle 
         Alignment       =   2  'Center
         Caption         =   "Middle Click"
         Height          =   495
         Left            =   5880
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblRight 
         Alignment       =   2  'Center
         Caption         =   "   Right  Click"
         Height          =   615
         Left            =   6600
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblY 
         Alignment       =   2  'Center
         Caption         =   "lblY"
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "lblX"
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Air Mouse Windows Driver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   600
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://parthasarathi.netfirms.com/Mscomm_control_.htm


Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, _
    ByVal y As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
    ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, _
    ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

'Private Const MOUSEEVENTF_ABSOLUTE = &H8000

Private Const MOUSE_PACKET_MARKER = &HAA
Private Const MOUSE_XDEAD = 70
Private Const MOUSE_YDEAD = 80

Dim flagMarkerFound

Public Sub LeftMouseClick()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Public Sub RightMouseClick()
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

Public Sub MouseMove(ByVal x As Long, ByVal y As Long)
    mouse_event MOUSEEVENTF_MOVE, x, y, 0, 0
End Sub

Private Sub doMouse(events() As Byte)
    Dim xdead, ydead As Long
    xdead = MOUSE_XDEAD
    ydead = MOUSE_YDEAD
     
    On Error GoTo handler
    
    
        
    If IsArray(events) Then
        Dim x, y As Long
        x = Val(events(1) - xdead)
        y = Val(events(2) - ydead)
        
        lblX.Caption = "X: " & Val(x)
        lblY.Caption = "Y: " & Val(y)
        MouseMove x, y
        
        If ((events(3) And &H3) = 3) Then
            lblMiddle.FontBold = Not lblMiddle.FontBold
        ElseIf (events(3) And &H2) Then
            lblLeft.FontBold = Not lblLeft.FontBold
            LeftMouseClick
        ElseIf (events(3) And &H1) Then
            lblRight.FontBold = Not lblRight.FontBold
            RightMouseClick
        End If
    
    End If
    
    Exit Sub
handler:
    MsgBox Err.Description
    Exit Sub
    
End Sub

Private Sub cmdMore_Click()
    If (cmdMore.Caption = "More >>") Then
        cmdMore.Caption = "<< Less"
        Me.Height = Frame1.Top + Frame1.Height
    Else
        cmdMore.Caption = "More >>"
        Me.Height = Frame1.Top + 500
    End If
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    On Error GoTo handler

    MSComm.RTSEnable = False
    flagMarkerFound = 0
    
    MSComm.CommPort = 1
    MSComm.PortOpen = True
    MSComm.RThreshold = 1   'Set to 1 initially and later to 4 after sync
    MSComm.InputLen = 1
    
    Exit Sub

handler:
    MsgBox "Form_Load()" & Err.Description
    Exit Sub
    
End Sub


Private Sub RTSBtn_Click()
    MSComm.RTSEnable = Not MSComm.RTSEnable
    RTSBtn.FontBold = Not RTSBtn.FontBold
End Sub

'Initially RThreshold set to 1;
'And eventually synchronized with marker and RThreshold is set to 4;
'Then forth this event is fired when there are 4 characters to be read in MSComm

Private Sub MSComm_oncomm()
    'Sync RThreshold
    If MSComm.RThreshold = 1 Then
        Dim buffer As Byte
        buffer = CByte(MSComm.Input(0))
        'Print buffer
        
        If (buffer = MOUSE_PACKET_MARKER) Then
            'Print "[" & MOUSE_PACKET_MARKER; " found" & vbNewLine
            flagMarkerFound = flagMarkerFound + 1
        Else: If (flagMarkerFound > 0) Then flagMarkerFound = flagMarkerFound + 1
        End If
        
        If (flagMarkerFound = 4) Then
            MSComm.RThreshold = 4
            MSComm.InputLen = 4
        End If
        
    Else
        Dim inbuffer() As Byte
        Dim i As Long

        ReDim inbuffer(MSComm.InBufferCount)
        inbuffer = MSComm.Input

        Me.RXtxt.Text = ""

        'Ubound(inbuffer) gives the upper bound of the array,
        'which is equal to the number of characters in the InputBuffer
        For i = 0 To UBound(inbuffer)
           Me.RXtxt.Text = Me.RXtxt.Text & "[" & i & "]" & inbuffer(i) & " "
        Next i
        
        'here we go!
        doMouse inbuffer
        
    End If
    
End Sub

Private Sub form_unload(Cancel As Integer)
    MSComm.PortOpen = False
End Sub

Private Sub cmdClear_Click()
    Me.Cls
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Circle (x, y), 10
End Sub
