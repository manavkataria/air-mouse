Attribute VB_Name = "MouseEvents"
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
Private Const MOUSEEVENTF_WHEEL = &H800
Private Const MOUSEEVENTF_HWHEEL = &H1000
Private Const MOUSEWHEEL_DELTA = 120

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

Public Sub VertMouseScroll(Optional ByVal clicks As Byte = 1)
    Call mouse_event(MOUSEEVENTF_WHEEL, 0, 0, MOUSEWHEEL_DELTA, 0)
End Sub

Public Sub HorzMouseScroll(Optional ByVal clicks As Byte = 1)
    Call mouse_event(MOUSEEVENTF_HWHEEL, 0, 0, MOUSEWHEEL_DELTA, 0)
End Sub

Public Sub MouseMove(ByVal x As Long, ByVal y As Long)
    mouse_event MOUSEEVENTF_MOVE, x, y, 0, 0
End Sub

'Private Sub slope(inbuf() As Byte, ByRef slope)
'    Static bufHistory(MOUSE_PACKET_SIZE, MOUSE_DYNAMIC_CALIBRATION_COUNT) As Byte
'End Sub

