Attribute VB_Name = "User32Events"
Option Explicit

'Set in Project Properties > Make > Conditional Compilation
'#Const DEF_3AXIS = 0
'#If DEF_3AXIS Then
'#Else
'#End If

'----------------------------------------------------------------------------------------------------------------
'                                           MOUSE METHODS DECLARATION
'----------------------------------------------------------------------------------------------------------------

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
        ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_WHEEL = &H800
Private Const MOUSEEVENTF_HWHEEL = &H1000
Private Const MOUSEWHEEL_DELTA = 120

'----------------------------------------------------------------------------------------------------------------
'                                          KEYBOARD METHODS DECLARATION
'----------------------------------------------------------------------------------------------------------------

Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

'Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
   ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'----------------------------------------------------------------------------------------------------------------
'                                            GLOBAL DATA DECLARATION
'----------------------------------------------------------------------------------------------------------------

Public GraphOn As Boolean

'----------------------------------------------------------------------------------------------------------------
'                                               MOUSE METHODS
'----------------------------------------------------------------------------------------------------------------

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

Public Sub VertMouseScroll(Optional ByVal clicks As Long = 1)
    Call mouse_event(MOUSEEVENTF_WHEEL, 0, 0, clicks * MOUSEWHEEL_DELTA, 0)
End Sub

Public Sub HorzMouseScroll(Optional ByVal clicks As Long = 1)
    Call mouse_event(MOUSEEVENTF_HWHEEL, 0, 0, clicks * MOUSEWHEEL_DELTA, 0)
End Sub

Public Sub MouseMove(ByVal X As Long, ByVal Y As Long)
    mouse_event MOUSEEVENTF_MOVE, X, Y, 0, 0
End Sub

'----------------------------------------------------------------------------------------------------------------
'                                               KEYBOARD METHODS
'----------------------------------------------------------------------------------------------------------------

Public Sub HorzKeybScroll(ByVal X As Long)
    'mouse_event MOUSEEVENTF_MOVE, x, y, 0, 0
End Sub

Public Sub VertKeybScroll(ByVal Y As Long)
        
End Sub

Public Function isControlKey() As Boolean

If GetAsyncKeyState(vbKeyUp) < 0 Then MsgBox "UP key pressed"
If GetAsyncKeyState(vbKeyDown) < 0 Then MsgBox "DOWN key pressed"
If GetAsyncKeyState(vbKeyLeft) < 0 Then MsgBox "LEFT key pressed"
If GetAsyncKeyState(vbKeyRight) < 0 Then MsgBox "RIGHT key pressed"
If GetAsyncKeyState(vbKeyControl) < 0 Then MsgBox "CONTROL key pressed"
If GetAsyncKeyState(vbKeyShift) < 0 Then Debug.Print "SHIFT key pressed"

End Function

Public Sub UpBhejo()
    keybd_event vbKeyUp, 0, 0, 0
End Sub

Public Sub DownBhejo()
    keybd_event vbKeyDown, 0, 0, 0
End Sub
'----------------------------------------------------------------------------------------------------------------
'                                               OTHER METHODS
'----------------------------------------------------------------------------------------------------------------

'Private Sub slope(inbuf() As Byte, ByRef slope)
'    Static bufHistory(MOUSE_PACKET_SIZE, MOUSE_DYNAMIC_CALIBRATION_COUNT) As Byte
'End Sub

