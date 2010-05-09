VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9435
   LinkTopic       =   "Form3"
   ScaleHeight     =   6690
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                 ByVal lpWindowName As String) As Long

Private Declare Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As Long) As Long
    
' Send a series of key presses to the Calculator application.
Private Sub Command1_Click()
    ' Get a handle to the Calculator application. The window class
    ' and window name were obtained using the Spy++ tool.
    
    'Dim calculatorHandle As IntPtr = FindWindow("SciCalc", "Calculator")
    Dim winHandle As Double
    'winHandle = FindWindow("SciCalc", "Calculator")
    'winHandle = FindWindow("QWidget", "Google Earth Pro                 ")
    winHandle = FindWindow("QWidget", "Google Earth")

    ' Verify that Calculator is a running process.
    If winHandle = 0 Then
        MsgBox ("'Google Earth' is not running.")
    End If
    
    SetForegroundWindow (winHandle)
    SendKeys "{UP 50}"
    SendKeys ("{LEFT 10}")
    SendKeys ("{DOWN 20}")
    SendKeys ("{RIGHT 10}")
End Sub

