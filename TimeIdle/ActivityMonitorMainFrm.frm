VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activity / Inactivity Monitor"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3120
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##################################################################################################
'   The accompanying module was created by Scott Morris (10/04/02)
'   The purpose of that module was to be able to detect user inactivity
'   Please see the module for specifics on how to use it
'##################################################################################################

Dim gTimerNumber As Long 'Timer Number
Dim gMaxTime As Long 'Max Time before timeout occurs

'Double-click the form to stop monitoring
Private Sub Form_DblClick()
    hStopMonitor gTimerNumber, Me.hwnd
    Timer1.Enabled = False
    MsgBox "Monitor stopped.", vbOKOnly, "STOPPED"
End Sub

'Set the max inactivity time limit, set a timer number, and start monitoring
Private Sub Form_Load()
    gMaxTime = 10000 'This just sets the maximum time you want to wait to have a 'hTimeoutOccurred' event fire
    gTimerNumber = 1369 'This is an identifier number for the timer you want to create (can be set to anything!)
    hStartMonitor Me.hwnd, gMaxTime, gTimerNumber 'Start the monitor
End Sub

'When the maximum inactivity time limit has been exceeded, the following subroutine gets called
Public Sub hTimeoutOccurred(pTimerNumber As Long)
    MsgBox "TIME LIMIT EXCEEDED ON TIMER " + Trim(Str(pTimerNumber))
End Sub

'Stop the monitor before the form completely terminates
Private Sub Form_Unload(Cancel As Integer)
    hStopMonitor gTimerNumber, Me.hwnd
End Sub

'Display the time since the last key or mouse button was pressed
Private Sub Timer1_Timer()
    Label1.Caption = Trim(Str(hTimePassed))
End Sub
