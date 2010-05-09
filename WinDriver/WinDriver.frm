VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   4590
   ClientTop       =   1920
   ClientWidth     =   9750
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
   ScaleWidth      =   9750
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   21
      Top             =   7890
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmr 
      Interval        =   1
      Left            =   7920
      Top             =   960
   End
   Begin VB.Frame frame2 
      Caption         =   "Mouse Calibration"
      Height          =   3375
      Left            =   1080
      TabIndex        =   12
      Top             =   1320
      Width           =   5775
      Begin VB.TextBox txtCalibCount 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   4320
         TabIndex        =   15
         Text            =   "100"
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblY 
         Alignment       =   2  'Center
         Caption         =   "lblY"
         Height          =   495
         Left            =   1920
         TabIndex        =   20
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "lblX"
         Height          =   495
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
      Left            =   7800
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   6
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   38400
      InputMode       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mouse Events"
      Height          =   1455
      Left            =   960
      TabIndex        =   3
      Top             =   5880
      Width           =   7935
      Begin VB.TextBox txtRXRaw 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Text            =   "RX Raw"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox RXtxt 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "RX Data"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblMiddle 
         Alignment       =   2  'Center
         Caption         =   "Middle Click"
         Height          =   735
         Left            =   6600
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblLeft 
         Alignment       =   2  'Center
         Caption         =   "   Left    Click"
         Height          =   735
         Left            =   5040
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblRight 
         Alignment       =   2  'Center
         Caption         =   "   Right  Click"
         Height          =   735
         Left            =   5760
         TabIndex        =   8
         Top             =   480
         Width           =   615
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
      Caption         =   "Air Mouse Windows Driver 2.8"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   360
      Width           =   3975
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
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40


'Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSE_PACKET_MARKER = &HAA
Private Const MOUSE_PACKET_SIZE = 4
Private Const MOUSE_XCALIB = 80
Private Const MOUSE_YCALIB = 80
Private Const MOUSE_XDEAD = 3
Private Const MOUSE_YDEAD = 3
Private Const MOUSE_CALIBRATION_COUNT = 100
Private Const MOUSE_DYNAMIC_CALIBRATION_COUNT = 10

Public Enum MouseCalibrationState
    CALIB_NEVER = 0
    CALIB_YES = 1
    CALIB_NO = 2
End Enum

Dim flagMarkerFound As Byte
Dim mouseLeftReport, mouseRightReport, mouseLeft, mouseRight As Byte
Dim MouseCalibrated As MouseCalibrationState
Dim MouseCalibCount, MouseXCalib, MouseYCalib As Long
Dim MouseXDead, MouseYDead As Long

Public Sub LeftMouseDown()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub

Public Sub LeftMouseUp()
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Public Sub LeftMouseClick()
    LeftMouseDown
    LeftMouseUp
End Sub

Public Sub RightMouseDown()
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
End Sub

Public Sub RightMouseUp()
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

Public Sub RightMouseClick()
    RightMouseDown
    RightMouseUp
End Sub

Public Sub MiddleMouseClick()
    mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
End Sub

Public Sub MouseMove(ByVal x As Long, ByVal y As Long)
    mouse_event MOUSEEVENTF_MOVE, x, y, 0, 0
End Sub

'Private Sub slope(inbuf() As Byte, ByRef slope)
'    Static bufHistory(MOUSE_PACKET_SIZE, MOUSE_DYNAMIC_CALIBRATION_COUNT) As Byte
'End Sub

Private Sub doMouse(events() As Byte)
    Static leftctr, rightctr, middlectr As Long
    
    On Error GoTo handler
        
    If (IsArray(events) And UBound(events) = 3) Then
        Dim x, y As Long
        x = Val(events(1) - MouseXCalib)
        y = Val(events(2) - MouseYCalib)
        
        'Dead Zone
        If (Abs(x) < MouseXDead) Then x = 0
        If (Abs(y) < MouseYDead) Then y = 0
        
        lblX.Caption = "X: " & Val(-x)
        lblY.Caption = "Y: " & Val(y)
        
        MouseMove -x, y
        
        lblRxX.Caption = "X: " & events(1)
        lblRxY.Caption = "Y: " & events(2)
        
        'check debounce; consult with Timer
        mouseLeftReport = ((events(3) And &H2) = 2)
        mouseRightReport = ((events(3) And &H1) = 1)
        
        If (mouseLeft And mouseRight) Then
            'mouseMiddle = False
            mouseLeft = False
            mouseRight = False
            
            MiddleMouseClick
            middlectr = middlectr + 1
            lblMiddle.FontBold = Not lblMiddle.FontBold
            lblMiddle.Caption = "Middle Click# " & Val(middlectr)
        ElseIf (mouseLeft) Then
            mouseLeft = False
            LeftMouseClick
            'LeftMouseDown
            leftctr = leftctr + 1
            lblLeft.FontBold = Not lblLeft.FontBold
            lblLeft.Caption = "Left Click# " & Val(leftctr)
        ElseIf (mouseRight) Then
            mouseRight = False
            RightMouseClick
            rightctr = rightctr + 1
            lblRight.FontBold = Not lblRight.FontBold
            lblRight.Caption = "Right Click# " & Val(rightctr)
        'Else
            'LeftMouseUp
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
    If (MouseCalibrated = CALIB_YES) Then
        MouseCalibrated = CALIB_NO
        cmdCalibrate.FontBold = False
    Else
        cmdCalibrate.FontBold = True
        cmdUncalibrate.FontBold = False
    End If
End Sub

Private Sub restCalibration(inbuf() As Byte)
    Static i As Double, xrest As Long, yrest As Long
    Static xmin, xmax, ymin, ymax As Long
    
    If (i = 0) Then
        'i = 0
        StatusBar1.Panels(1).Text = "Calibrating ..."
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
        MouseXDead = xmax - xmin + 1    '+1 for saftey
        MouseYDead = ymax - ymin + 1    '+1 for saftey;
        
        Call updateDisplayFrames
              
        MouseCalibrated = CALIB_YES
        cmdCalibrate.FontBold = True
        cmdUncalibrate.FontBold = False
        StatusBar1.Panels(1).Text = "Calib Done!"
        
        i = 0
        xrest = 0
        yrest = 0
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
        
    End If
    
    'ProgressBar1 = Val(i)
     
End Sub

'Initially RThreshold set to 1;
'And eventually synchronized with marker and RThreshold is set to MOUSE_PACKET_SIZE ;
'Then forth this event is fired when there are 4 characters to be read in MSComm

Private Sub syncMarker()
    Dim buffer As Byte
    buffer = CByte(MSComm.Input(0))
    txtRXRaw.Text = CStr(buffer)
    
    If (buffer = MOUSE_PACKET_MARKER) Then
        StatusBar1.Panels(1).Text = "Syncing ..."
        flagMarkerFound = flagMarkerFound + 1
    ElseIf (flagMarkerFound > 0) Then
        flagMarkerFound = flagMarkerFound + 1
    End If
        
    If (flagMarkerFound = MOUSE_PACKET_SIZE) Then
        MSComm.RThreshold = MOUSE_PACKET_SIZE
        MSComm.InputLen = MOUSE_PACKET_SIZE
        StatusBar1.Panels(1).Text = "Sync Done!"
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
            'dynamicCalibration inbuffer
            'slopeComputaion over past MOUSE_DYNAMIC_CALIBRATION_COUNT RX inputs
            doMouse inbuffer ', slope
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
        StatusBar1.Panels(1).Text = "Uncalibrating ..."
            
        MouseXCalib = 0
        MouseYCalib = 0
        MouseXDead = 0
        MouseYDead = 0
        
        cmdUncalibrate.FontBold = True
        cmdCalibrate.FontBold = False
    End If
    
    Call updateDisplayFrames
    StatusBar1.Panels(1).Text = "Calib Purged!"
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Circle (X, Y), 10
End Sub


Private Sub tmr_Timer()
    Static ctrDebounceLeft, ctrDebounceRight As Long
    If (IsEmpty(ctrDebounceLeft)) Then ctrDebounceLeft = 0
    If (IsEmpty(ctrDebounceRight)) Then ctrDebounceRight = 0
    
    If (mouseLeftReport = False) Then
        mouseLeft = False
        ctrDebounceLeft = 0
    Else
        ctrDebounceLeft = ctrDebounceLeft + 1
    End If
    If ctrDebounceLeft = 3 Then mouseLeft = True
    
    If (mouseRightReport = False) Then
        mouseRight = False
        ctrDebounceRight = 0
    Else
        ctrDebounceRight = ctrDebounceRight + 1
    End If
    If ctrDebounceRight = 3 Then mouseRight = True
    
End Sub

Private Sub txtCalibCount_LostFocus()
   MouseCalibCount = txtCalibCount.Text
End Sub

Private Sub Form_Load()
    On Error GoTo handler

    MSComm.RTSEnable = False
    flagMarkerFound = 0
    
    MSComm.CommPort = 6
    MSComm.PortOpen = True
    MSComm.RThreshold = 1   'Set to 1 initially and later to 4 after syncMarker
    MSComm.InputLen = 1
    MouseCalibrated = CALIB_NEVER
    
    MouseCalibCount = MOUSE_CALIBRATION_COUNT
    txtCalibCount = MOUSE_CALIBRATION_COUNT
    MouseXCalib = MOUSE_XCALIB
    MouseYCalib = MOUSE_YCALIB
    MouseXDead = MOUSE_XDEAD
    MouseYDead = MOUSE_YDEAD
    Exit Sub

handler:
    MsgBox "Form_Load()" & Err.Description
    Exit Sub
    
End Sub
