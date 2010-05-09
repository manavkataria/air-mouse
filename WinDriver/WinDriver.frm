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
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1800
      Width           =   3615
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
      BaudRate        =   38400
      InputMode       =   1
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
'http://parthasarathi.netfirms.com/Mscomm_control_.htm

Option Explicit
Private Const MOUSE_PACKET_MARKER = &HAA
Dim flagMarkerFound

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
    MsgBox Err.Description
    Exit Sub
    
End Sub

Private Sub RTSBtn_Click()
    MSComm.RTSEnable = Not MSComm.RTSEnable
    RTSBtn.FontBold = Not RTSBtn.FontBold
End Sub

'Initially RThreshold set to 1; then synchronized with marker and RThreshold set to 4;
'This event is fired when there are 4 characters to be read in MSComm



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
        Dim inbuffer() As Byte  'Declare an array of bytes
        Dim i As Long

        ReDim inbuffer(MSComm.InBufferCount) 'Specify the size of the array.
        inbuffer = MSComm.Input

        Me.RXtxt.Text = ""

        'Ubound(inbuffer) gives the upper bound of the array,
        'which is equal to the number of characters in the InputBuffer
        For i = 0 To UBound(inbuffer)
           Me.RXtxt.Text = Me.RXtxt.Text & " [" & i & "]" & (((inbuffer(i))))
        Next i
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
