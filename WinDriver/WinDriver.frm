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
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   3840
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
      Height          =   495
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6720
      Width           =   3015
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
      TabIndex        =   6
      Top             =   1440
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
      BaudRate        =   38400
      InputMode       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "RX Data"
      Height          =   1575
      Left            =   3000
      TabIndex        =   5
      Top             =   6120
      Width           =   3735
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
      Left            =   1440
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
'http://parthasarathi.netfirms.com/Mscomm_control_.htm


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

Public Sub MouseMove(ByVal X As Long, ByVal Y As Long)
    mouse_event MOUSEEVENTF_MOVE, X, Y, 0, 0
End Sub

Private Sub doMouse(events() As Byte)
    Dim xdead, ydead As Long
    xdead = MOUSE_XDEAD
    ydead = MOUSE_YDEAD
     
    On Error GoTo handler
    
    If IsArray(events) Then
        MouseMove Val(events(1) - xdead), Val(events(2) - ydead)
        If (events(3) * &H2) Then LeftMouseClick
        If (events(3) * &H1) Then RightMouseClick
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Circle (X, Y), 10
End Sub
