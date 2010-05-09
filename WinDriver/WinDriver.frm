VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Accelerometer Mouse WinDriver"
   ClientHeight    =   7965
   ClientLeft      =   1980
   ClientTop       =   1920
   ClientWidth     =   9210
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
   ScaleHeight     =   7965
   ScaleWidth      =   9210
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   6960
      TabIndex        =   25
      Top             =   7560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "Graph >>"
      Height          =   495
      Left            =   7320
      TabIndex        =   24
      Top             =   4560
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   7470
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      TabIndex        =   11
      Top             =   1320
      Width           =   5775
      Begin VB.TextBox txtCalibCount 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Text            =   "100"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblXPos 
         Alignment       =   2  'Center
         Caption         =   "lblXPos"
         Height          =   615
         Left            =   4440
         TabIndex        =   26
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblXVel 
         Alignment       =   2  'Center
         Caption         =   "lblXVel"
         Height          =   615
         Left            =   3360
         TabIndex        =   23
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblMode 
         Alignment       =   2  'Center
         Caption         =   "lblMode"
         Height          =   495
         Left            =   4080
         TabIndex        =   22
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblY 
         Alignment       =   2  'Center
         Caption         =   "lblY"
         Height          =   495
         Left            =   1920
         TabIndex        =   19
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "lblX"
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblydead 
         Alignment       =   2  'Center
         Caption         =   "Y DeadZone"
         Height          =   855
         Left            =   1920
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblxdead 
         Alignment       =   2  'Center
         Caption         =   "X DeadZone"
         Height          =   855
         Left            =   480
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblCalibCount 
         Alignment       =   2  'Center
         Caption         =   "Calibration Count"
         Height          =   855
         Left            =   3840
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblycalib 
         Alignment       =   2  'Center
         Caption         =   "Y Calibration"
         Height          =   855
         Left            =   1920
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblxcalib 
         Alignment       =   2  'Center
         Caption         =   "X Calibration"
         Height          =   975
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdUncalibrate 
      Caption         =   "Uncalibrate"
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
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
      TabIndex        =   3
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
      BaudRate        =   115200
      InputMode       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mouse Events"
      Height          =   1455
      Left            =   720
      TabIndex        =   2
      Top             =   5640
      Width           =   7935
      Begin VB.TextBox txtRXRaw 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   21
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
         TabIndex        =   4
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
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblLeft 
         Alignment       =   2  'Center
         Caption         =   "   Left    Click"
         Height          =   735
         Left            =   5040
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblRight 
         Alignment       =   2  'Center
         Caption         =   "   Right  Click"
         Height          =   735
         Left            =   5760
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblRxY 
         Alignment       =   2  'Center
         Caption         =   "lblRxY"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblRxX 
         Alignment       =   2  'Center
         Caption         =   "lblRxX"
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Air Mouse Windows Driver 2.14"
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
      TabIndex        =   9
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'For USB = 6; For Serial = 1
Private Const MOUSE_COM_PORT = 6

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

Public Enum MouseModeState
    MOUSE_MODE_MOVE = 0
    MOUSE_MODE_SCROLL
    MOUSE_MODE_EARTH
    MAX_MOUSE_MODE
End Enum

Dim flagMarkerFound As Byte
Dim mouseLeftReport, mouseRightReport, mouseLeft, mouseRight As Byte
Dim MouseCalibrated As MouseCalibrationState
Dim MouseCalibCount, MouseXCalib, MouseYCalib As Long
Dim MouseXDead, MouseYDead As Long
Dim MouseMode As Long
Dim xReport, yReport As Long

'InfoBars
Private defProgBarHwnd  As Long
Private Declare Function SetParent Lib "user32" _
  (ByVal hWndChild As Long, _
   ByVal hWndNewParent As Long) As Long
  
'NOTE: Only Dynamic Acceleratoin should be passed to me!
Private Sub calcPosition(accl)
    'Angular Velocity:
    Static xvel(1) As Long, xaccl(1) As Long, xpos(1) As Long
    
    xaccl(1) = accl
    
    If IsEmpty(xvel(0)) Then
        xvel(0) = 0
        xvel(1) = 0
    End If
    
    xvel(1) = xvel(0) + xaccl(0) + (xaccl(1) - xaccl(0)) / 2
    xvel(0) = xvel(1)
    
    xpos(1) = xpos(0) + xvel(0) + (xvel(1) - xvel(0)) / 2
    xpos(0) = xpos(1)
    
    xaccl(0) = xaccl(1)
    
    lblXVel.Caption = "XVelocity: " & xvel(1)
    lblXPos.Caption = "XPos: " & xpos(1)

End Sub


'TODO: 1) Add Invert X Button; x=-x
'TODO: 2) Fix Scroll Implementation

Private Sub doMouse(events() As Byte)
'On Error GoTo handler
Static leftctr, rightctr, middlectr As Long
        
    If (IsArray(events) And UBound(events) = (MOUSE_PACKET_SIZE - 1)) Then
        
        Dim x, y As Long
        'DC Cancellation aka Zero Gravity Cancellation Filter
        x = Val(events(1) - MouseXCalib)
        y = Val(events(2) - MouseYCalib)
        
        'Dead Zone Filter aka Mechanical Filter Zone
        If (Abs(x) < MouseXDead) Then x = 0
        If (Abs(y) < MouseYDead) Then y = 0
        
        lblRxX.Caption = "X: " & events(1)
        lblRxY.Caption = "Y: " & events(2)
        lblX.Caption = "X: " & -x
        lblY.Caption = "Y: " & y
        
        xReport = -x
        yReport = y
        
        'Position Estimation:
        'NOTE: Only Dynamic Acceleratoin should be passed here!
        'calcPosition (xReport)
        
        If MouseMode = MOUSE_MODE_MOVE Then
            MouseMove -x, y
        ElseIf MouseMode = MOUSE_MODE_SCROLL Then
            HorzMouseScroll (-x / 2)
            VertMouseScroll (y / 2)
        ElseIf MouseMode = MOUSE_MODE_EARTH Then
            HorzKeybScroll (-x / 2)
            VertKeybScroll (y / 2)
        End If
            
        If (GraphOn = True) Then Call Form2.plotxy(-x, y)
                
        'check debounce; consult with Timer
        mouseLeftReport = ((events(3) And &H2) = 2)
        mouseRightReport = ((events(3) And &H1) = 1)
        
        If (mouseLeft And mouseRight) Then
            mouseLeft = False
            mouseRight = False
            
            Call MiddleMouseClick
            middlectr = middlectr + 1
            lblMiddle.FontBold = Not lblMiddle.FontBold
            lblMiddle.Caption = "Middle Click# " & Val(middlectr)
        ElseIf (mouseLeft) Then
            mouseLeft = False
            Call LeftMouseClick
            leftctr = leftctr + 1
            lblLeft.FontBold = Not lblLeft.FontBold
            lblLeft.Caption = "Left Click# " & Val(leftctr)
        ElseIf (mouseRight) Then
            mouseRight = False
            rightctr = rightctr + 1
            lblRight.FontBold = Not lblRight.FontBold
            lblRight.Caption = "Right Click# " & Val(rightctr)
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

'Initially GraphOn = False
Public Sub cmdGraph_Click()
    cmdGraph.FontBold = Not cmdGraph.FontBold
    GraphOn = Not GraphOn
    
    If (GraphOn = True) Then
        cmdGraph.Caption = "Graph <<"
        Load Form2
        Form2.Show
    Else
        cmdGraph.Caption = "Graph >>"
        Unload Form2
        Set Form2 = Nothing
    End If
    
End Sub

Private Sub cmdExit_Click()
    Call RestoreParent
    Unload Me
    Unload Form2
End Sub

Private Sub cmdCalibrate_Click()
    If (MouseCalibrated = CALIB_YES) Then
        MouseCalibrated = CALIB_NO
        cmdCalibrate.FontBold = False
        tmr.Enabled = False
    Else
        cmdCalibrate.FontBold = True
        cmdUncalibrate.FontBold = False
    End If
End Sub

Private Sub restCalibration(inbuf() As Byte)
    Static i As Double, xrest As Long, yrest As Long
    Static xmin, xmax, ymin, ymax As Long
    
    'Initializations
    If (i = 0) Then
        With StatusBar1
            .Panels(1).Text = "Calibrating..."
            .Panels(1).AutoSize = sbrSpring
            .Panels(1).Bevel = sbrInset
        End With
        
        xrest = 0
        yrest = 0
        xmin = inbuf(1)
        xmax = inbuf(1)
        ymin = inbuf(2)
        ymax = inbuf(2)
    End If
    
    'DeadZone Computation
    If (i < MouseCalibCount And MouseCalibrated <> CALIB_YES) Then
        If (xmin > inbuf(1)) Then xmin = inbuf(1)
        If (xmax < inbuf(1)) Then xmax = inbuf(1)
        If (ymin > inbuf(2)) Then ymin = inbuf(2)
        If (ymax < inbuf(2)) Then ymax = inbuf(2)
    
        xrest = xrest + inbuf(1)
        yrest = yrest + inbuf(2)
        i = i + 1
        ProgressBar1.Value = i
        'needed to trap cancel click
        'DoEvents
   
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
        StatusBar1.Panels(1).Text = "Calibration Complete!"
        ProgressBar1.Value = 0
 
        i = 0
        xrest = 0
        yrest = 0
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
        
        tmr.Enabled = True
    End If

End Sub


'Initially RThreshold set to 1;
'And eventually synchronized with marker and RThreshold is set to MOUSE_PACKET_SIZE ;
'Then forth this event is fired when there are 4 characters to be read in MSComm
Private Sub syncMarker(inbuf() As Byte)
    Dim buffer As Byte
    
    If UBound(inbuf) = 0 Then
        buffer = CByte(inbuf(0))
    End If
    
    txtRXRaw.Text = CStr(buffer)
    
    If (buffer = MOUSE_PACKET_MARKER) Then
        StatusBar1.Panels(1).Text = "Syncing Packet..."
        flagMarkerFound = flagMarkerFound + 1
    ElseIf (flagMarkerFound > 0) Then
        flagMarkerFound = flagMarkerFound + 1
    End If
       
    If (flagMarkerFound = MOUSE_PACKET_SIZE) Then
        MSComm.RThreshold = MOUSE_PACKET_SIZE
        MSComm.InputLen = MOUSE_PACKET_SIZE
        StatusBar1.Panels(1).Text = "Syncing Packet Complete!"
    End If
End Sub

Private Sub reSyncMarker(inbuf() As Byte)
    flagMarkerFound = 0
    
    MouseCalibrated = CALIB_NO
    MSComm.PortOpen = False
    MSComm.CommPort = MOUSE_COM_PORT
    MSComm.PortOpen = True
    MSComm.RThreshold = 1   'Set to 1 initially and later to 4 after sync
    MSComm.InputLen = 1
    
    Call syncMarker(inbuf)
End Sub

Private Function InSync(inbuf() As Byte) As Boolean
    If (inbuf(0) = MOUSE_PACKET_MARKER) Then
        InSync = True
    Else
        InSync = False
   '     MsgBox "Out of Sync Packets! Forcing Packet Resync!", , "Packet Sync Error"
    End If
End Function

Private Sub MSComm_oncomm()
'On Error GoTo handler
    Dim inbuffer() As Byte
    Dim i As Long
    
    ReDim inbuffer(MSComm.InBufferCount)
    inbuffer = MSComm.Input
    
    'Sync with Marker
    If Not MSComm.RThreshold = MOUSE_PACKET_SIZE Then
        syncMarker inbuffer
        Exit Sub
    End If
    
    Me.RXtxt.Text = ""
    txtRXRaw.Text = ""

    'Ubound(inbuffer) gives the upper bound of the array,
    For i = 0 To UBound(inbuffer)
       Me.RXtxt.Text = Me.RXtxt.Text & "[" & i & "]" & inbuffer(i) & " "
       txtRXRaw.Text = txtRXRaw.Text & "[" & i & "]" & Hex(inbuffer(i)) & " "
    Next i
    
    'here we go!
    If (MSComm.RThreshold = MOUSE_PACKET_SIZE) And (Not InSync(inbuffer)) Then Call reSyncMarker(inbuffer)
    
    If Not MouseCalibrated = CALIB_YES Then
        restCalibration inbuffer
    Else
        'dynamicCalibration inbuffer
        'slopeComputaion over past MOUSE_DYNAMIC_CALIBRATION_COUNT RX inputs
        doMouse inbuffer ', slope
    End If
    Exit Sub
    
handler:
    MsgBox Err.Description
    Call reSyncMarker(inbuffer)
    Exit Sub
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
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
    StatusBar1.Panels(1).Text = "Calibration Purged!"
    
End Sub

'InfoBars
Private Sub CreateInfoBars()
   Dim pnl As Panel
   Dim x As Long
   
  'create statusbar
   With StatusBar1
        .Panels(1).AutoSize = sbrSpring
            For x = 1 To 1
                Set pnl = .Panels.Add(, , "", sbrText)
                pnl.Alignment = sbrLeft
                pnl.Bevel = sbrInset
            If x = 1 Then pnl.AutoSize = sbrSpring
        Next
   End With
   
   With ProgressBar1
      .Min = 0
      .Max = 100
      .Value = .Max
   End With
   
   Call SetProgressBar
End Sub

Private Sub SetProgressBar()
   Dim pading As Long
  'parent the progress bar in the status bar
   pading = 40
   AttachProgBar ProgressBar1, StatusBar1, 2, pading
   ProgressBar1.Value = 0
End Sub

Private Sub RestoreParent()
    If defProgBarHwnd <> 0 Then
        SetParent ProgressBar1.hWnd, defProgBarHwnd
    End If
End Sub

Private Function AttachProgBar(pb As ProgressBar, sb As StatusBar, nPanel As Long, pading As Long)
 If defProgBarHwnd = 0 Then
     'change the parent
      defProgBarHwnd = SetParent(pb.hWnd, sb.hWnd)
   
      With sb
         .Align = vbAlignTop
         .Visible = False
         
        'change, move, set size and re-show
        'the progress bar in the new parent
         With pb
            .Visible = False
            .Align = vbAlignNone
            .Appearance = ccFlat
            .BorderStyle = ccNone
            .Width = sb.Panels(nPanel).Width
            .Move (sb.Panels(nPanel).Left + pading), _
                 (sb.Top + pading), _
                 (sb.Panels(nPanel).Width - (pading * 2)), _
                 (sb.Height - (pading * 2))
                  
            .Visible = True
            .ZOrder 0
         End With
           
        'restore the statusbar to the
        'bottom of the form and show
         .Panels(nPanel).AutoSize = sbrNoAutoSize
         .Align = vbAlignBottom
         .Visible = True
         
       End With
    End If
End Function

Private Sub tmr_Timer()
    'Debounce:
    'Note: Mode logic below depends on this debounce code
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
    
    'Mode:
    'Note: Mode logic depends on above debounce code
    Dim threshold As Long
    threshold = 50
    
    If (ctrDebounceLeft >= threshold And ctrDebounceRight = threshold) Or (ctrDebounceLeft = threshold And ctrDebounceRight >= threshold) Then
        MouseMode = (MouseMode + 1) Mod MAX_MOUSE_MODE
        
        If MouseMode = MOUSE_MODE_MOVE Then
            lblMode.Caption = "Mode: Move"
        ElseIf MouseMode = MOUSE_MODE_SCROLL Then
            lblMode.Caption = "Mode: Scroll"
        ElseIf MouseMode = MOUSE_MODE_EARTH Then
            lblMode.Caption = "Mode: Earth"
        End If
        
    End If
    
    'Reset Graph
    Call Form2.ResetGraph
        
End Sub

Private Sub txtCalibCount_LostFocus()
   MouseCalibCount = txtCalibCount.Text
End Sub

Private Sub Form_Load()
On Error GoTo handler
        
    GraphOn = False
    tmr.Enabled = False
    MSComm.RTSEnable = False
    flagMarkerFound = 0
    
    MSComm.CommPort = MOUSE_COM_PORT
    MSComm.PortOpen = True
    MSComm.RThreshold = 1   'Set to 1 initially and later to 4 after syncMarker
    MSComm.InputLen = 1
    MouseCalibrated = CALIB_NEVER
    
    MouseMode = MOUSE_MODE_MOVE
    lblMode.Caption = "Mode: Move"
    
    MouseCalibCount = MOUSE_CALIBRATION_COUNT
    txtCalibCount = MOUSE_CALIBRATION_COUNT
    MouseXCalib = MOUSE_XCALIB
    MouseYCalib = MOUSE_YCALIB
    MouseXDead = MOUSE_XDEAD
    MouseYDead = MOUSE_YDEAD
       
    Call CreateInfoBars
    
    Exit Sub

handler:
    MsgBox Err.Description
    Exit Sub
    
End Sub
