VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
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
   ScaleHeight     =   8385
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame2 
      Caption         =   "Mouse Calibration"
      Height          =   3975
      Left            =   1080
      TabIndex        =   12
      Top             =   1320
      Width           =   5775
      Begin VB.TextBox txtCalibCount 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   4320
         TabIndex        =   15
         Text            =   "10"
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblY 
         Alignment       =   2  'Center
         Caption         =   "lblY"
         Height          =   855
         Left            =   1920
         TabIndex        =   20
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "lblX"
         Height          =   735
         Left            =   480
         TabIndex        =   19
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblydead 
         Alignment       =   2  'Center
         Caption         =   "Y DeadZone"
         Height          =   855
         Left            =   1920
         TabIndex        =   18
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblxdead 
         Alignment       =   2  'Center
         Caption         =   "X DeadZone"
         Height          =   855
         Left            =   480
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblCalibCount 
         Alignment       =   2  'Center
         Caption         =   "Calibration Count"
         Height          =   855
         Left            =   3840
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblycalib 
         Alignment       =   2  'Center
         Caption         =   "Y Calibration"
         Height          =   855
         Left            =   1920
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblxcalib 
         Alignment       =   2  'Center
         Caption         =   "X Calibration"
         Height          =   975
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "More >>"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdUncalibrate 
      Caption         =   "Uncalibrate"
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalibrate 
      Caption         =   "Calibrate"
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   9360
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   38400
      InputMode       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mouse Events"
      Height          =   1575
      Left            =   960
      TabIndex        =   3
      Top             =   5880
      Width           =   7935
      Begin VB.TextBox RXtxt 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "RX Data"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblLeft 
         Alignment       =   2  'Center
         Caption         =   "   Left    Click"
         Height          =   495
         Left            =   5040
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMiddle 
         Alignment       =   2  'Center
         Caption         =   "Middle Click"
         Height          =   495
         Left            =   5640
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblRight 
         Alignment       =   2  'Center
         Caption         =   "   Right  Click"
         Height          =   615
         Left            =   6240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblRxY 
         Alignment       =   2  'Center
         Caption         =   "lblRxY"
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblRxX 
         Alignment       =   2  'Center
         Caption         =   "lblRxX"
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Air Mouse Prototype Windows Driver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   10
      Top             =   360
      Width           =   3495
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

Private Const MOUSE_XCALIB = 80
Private Const MOUSE_YCALIB = 80
Private Const MOUSE_XDEAD = 3
Private Const MOUSE_YDEAD = 3
Private Const MOUSE_CALIBRATION_COUNT = 10

Public Enum MouseCalibrationState
    CALIB_NEVER = 0
    CALIB_YES = 1
    CALIB_NO = 2
End Enum

Dim flagMarkerFound As Byte
Dim MouseCalibrated As MouseCalibrationState
Dim MouseCalibCount, MouseXCalib, MouseYCalib
Dim MouseXDead, MouseYDead

'Dim MouseCalibrated As Boolean

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
    On Error GoTo handler
        
    If (IsArray(events) And UBound(events) = 3) Then
        Dim X, Y As Long
        X = Val(events(1) - MouseXCalib)
        Y = Val(events(2) - MouseYCalib)
        
        'Dead Zone
        If (Abs(X) < MouseXDead) Then X = 0
        If (Abs(Y) < MouseYDead) Then Y = 0
        
        lblX.Caption = "X: " & Val(-X)
        lblY.Caption = "Y: " & Val(Y)
        
        MouseMove -X, Y
        
        lblRxX.Caption = "X: " & events(1)
        lblRxY.Caption = "Y: " & events(2)
        
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
    MsgBox "doMouse() " & Err.Description
    Exit Sub
    
End Sub

Private Sub updateDisplayFrames()
        lblxcalib = "X Calibration: " & MouseXCalib
        lblycalib = "Y Calibration: " & MouseYCalib
        lblxdead = "X Deadzone: " & MouseXDead
        lblydead = "Y Deadzone: " & MouseYDead
 
End Sub

Private Sub cmdMore_Click()
    If (cmdMore.Caption = "More >>") Then
        cmdMore.Caption = "<< Less"
        Me.Height = 8430 'frame2.Top + frame2.Height + Frame1.Height * 1.5
    Else
        cmdMore.Caption = "More >>"
        Me.Height = 6060
    End If
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdCalibrate_Click()
    If (MouseCalibrated = CALIB_YES) Then MouseCalibrated = CALIB_NO
    cmdCalibrate.FontBold = True
    cmdUncalibrate.FontBold = False
End Sub

Private Sub restCalibration(inbuf() As Byte)
    Static i As Double, xrest As Long, yrest As Long
    Static xmin, xmax, ymin, ymax As Long
    
    If (i = 0) Then
        'i = 0
        xrest = 0
        yrest = 0
        xmin = inbuf(1)
        xmax = inbuf(1)
        ymin = inbuf(2)
        ymax = inbuf(2)
    End If
       
    If (i < MouseCalibCount And MouseCalibrated <> CALIB_YES) Then
        If (xmin > inbuf(1)) Then xmin = inbuf(1)
        If (xmax < inbuf(1)) Then xmax = inbuf(1)
        If (ymin > inbuf(2)) Then ymin = inbuf(2)
        If (ymax < inbuf(2)) Then ymax = inbuf(2)
    
        xrest = xrest + inbuf(1)
        yrest = yrest + inbuf(2)
        i = i + 1
    
    ElseIf (i = MouseCalibCount) Then
        xrest = xrest / i
        yrest = yrest / i
        
        MouseXCalib = xrest
        MouseYCalib = yrest
        MouseXDead = xmax - xmin
        MouseYDead = ymax - ymin
        
        Call updateDisplayFrames
              
        MouseCalibrated = CALIB_YES
        cmdCalibrate.FontBold = True
        
        i = 0
        xrest = 0
        yrest = 0
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
        
    End If
     
End Sub

'Initially RThreshold set to 1;
'And eventually synchronized with marker and RThreshold is set to 4;
'Then forth this event is fired when there are 4 characters to be read in MSComm

Private Sub syncMarker()
    Dim buffer As Byte
    buffer = CByte(MSComm.Input(0))
        
    If (buffer = MOUSE_PACKET_MARKER) Then
        flagMarkerFound = flagMarkerFound + 1
    ElseIf (flagMarkerFound > 0) Then
        flagMarkerFound = flagMarkerFound + 1
    End If
        
    If (flagMarkerFound = 4) Then
        MSComm.RThreshold = 4
        MSComm.InputLen = 4
    End If
End Sub

Private Sub reSyncMarker()
On Error GoTo handler
    flagMarkerFound = 0
    
    MSComm.CommPort = 1
    MSComm.PortOpen = True
    MSComm.RThreshold = 1   'Set to 1 initially and later to 4 after sync
    MSComm.InputLen = 1
  
    Call syncMarker
handler:
    
End Sub


Private Sub MSComm_oncomm()
    On Error GoTo handler
    
    'Sync with Marker
    If MSComm.RThreshold = 1 Then
        syncMarker
        
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
        If (MouseCalibrated = CALIB_YES) Then
            doMouse inbuffer
        Else
            restCalibration inbuffer
        End If
        
    End If
    Exit Sub
    
handler:
    'MsgBox Err.Description
    reSyncMarker
    Exit Sub
    
End Sub


Private Sub form_unload(Cancel As Integer)
    MSComm.PortOpen = False
End Sub

Private Sub cmdUncalibrate_Click()
    Dim result As VbMsgBoxResult
    result = MsgBox("Are you sure you want to uncalibrate the Air Mouse?", vbYesNo, "Uncalibrate AirMouse")
    
    If (result = vbYes) Then
        MouseXCalib = 0
        MouseYCalib = 0
        MouseXDead = 0
        MouseYDead = 0
        
        cmdUncalibrate.FontBold = True
        cmdCalibrate.FontBold = False
    End If
    
    Call updateDisplayFrames
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Circle (X, Y), 10
End Sub

Private Sub txtCalibCount_LostFocus()
   MouseCalibCount = txtCalibCount.Text
End Sub

Private Sub Form_Load()
    On Error GoTo handler

    MSComm.RTSEnable = False
    flagMarkerFound = 0
    
    MSComm.CommPort = 1
    MSComm.PortOpen = True
    MSComm.RThreshold = 1   'Set to 1 initially and later to 4 after syncMarker
    MSComm.InputLen = 1
    MouseCalibrated = CALIB_NEVER
    
    MouseCalibCount = MOUSE_CALIBRATION_COUNT
    'txtCalibCount = MOUSE_CALIBRATION_COUNT
    MouseXCalib = MOUSE_XCALIB
    MouseYCalib = MOUSE_YCALIB
    MouseXDead = MOUSE_XDEAD
    MouseYDead = MOUSE_YDEAD
    Exit Sub

handler:
    MsgBox "Form_Load()" & Err.Description
    Exit Sub
    
End Sub


