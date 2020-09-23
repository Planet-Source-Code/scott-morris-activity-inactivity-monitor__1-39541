Attribute VB_Name = "UserTimeout"
'####################################################################################################################################
'   USERTIMEOUT module created by Scott Morris (scottmmorris@mailbreak.com) on 10/04/02
'
'   The purpose of this module is to be able to detect user inactivity of keypresses and mousebutton clicks.
'   Feel free to use it in your own programs, as long as you give me some credit in the source code. :)
'   I realize that this isn't the 100% best way to do this, but it works. Please feel free to email me if
'   you have any suggestions or comments.  I will do what I can to help out.  Thanks for your interest!
'
'   TO USE THIS MODULE:
'
'   1) Attach the module to the project (duh...).
'   2) To start monitoring the user's activity/inactivity, in the main form of your project, you will need to
'       call the 'hStartMonitor' subroutine found below.
'
'       When you are calling it, you will need to pass in a few parameters.  They are a) a handle to the main
'       form's window, b) the maximum allowed time (in milliseconds) of inactivity before the timeout event
'       occurrs, and c) a unique numeric value used to identify the particular timer.
'
'   3) To find out how much time has passed since the last key or mouse button was pressed, you can call the
'       'hTimePassed' function.  It will return a Long value representing the time passed (in milliseconds) since
'       the last key or mouse button was pressed.
'
'   4) To terminate monitoring the user's activity/inactivity, call the 'hStopMonitor' subroutine.
'
'   5) In the hTimerProc subroutine, there is the following line:

'       MainFrm.hTimeoutOccurred uElapse 'uElapse is the timer number that fired the event
'
'       You need to make sure that the part that says 'MainFrm' is changed to whatever your main form's name is
'       and that the part that says 'hTimeoutOccurred' is an actual subroutine in that main form.  This will be the
'       subroutine called when the user's inactivity period exceeds the maximum allowed inactivity time.  It must
'       reside in your main form's code.  It also must receive a Long value (this is the timer number passed in
'       to the hStartMonitor subroutine when you first called it to start monitoring the user).
'
'
'   KNOWN ISSUES:
'
'   1) Does not currently monitor mouse cursor movement.  I plan to fix that, but until then, this will have to do. :)
'   2) Seems to miss some mouse clicks that switch widow focus - To see what I mean, open notepad, run this program,
'       click on the desktop and then your notepad window while watching this program's window.  The timer does
'       not reset.  However, if you click again or press another key, the timer will reset.  This is not true under one
'       condition.  That is if THIS program is inactive, and you reactivate it by clicking on the FORM, not the TITLEBAR,
'       the timer will reset; it catches the mouse click.  However, if THIS program is inactive, and you reactivate it
'       by clicking on the TITLEBAR, not the FORM, the timer will NOT reset; it misses the mouseclick.
'
'####################################################################################################################################

Const TIMER_TIMEOUT = 10 'Check every 1/100th of a second
Const VK_CAPITAL = &H14

Private Type KeyboardBytes 'Structure to hold our keyboard state
     kbByte(0 To 255) As Byte
End Type

Private Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Dim kbOld As KeyboardBytes 'Holds the last polled keyboard state
Dim gLastEvent As Long 'Holds the time in ms that the last event occurred since the computer was booted
Dim gMaxTime As Long

Public Sub hStartMonitor(ByVal whichWindow As Long, ByVal timeoutMax As Long, ByVal pTimerNumber)
    gMaxTime = timeoutMax
    gLastEvent = GetTickCount
    SetTimer whichWindow, pTimerNumber, TIMER_TIMEOUT, AddressOf hTimerProc
End Sub

Private Sub hTimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    Dim tCount As Integer
    Dim tKeyPressed As Boolean
    GetKeyboardState kbOld 'Put our keyboard state into our structure
    
    For tCount = 0 To 255
        tKeyPressed = False
        If kbOld.kbByte(tCount) > 1 Then tKeyPressed = True: Exit For
    Next
    
    DoEvents
    If tKeyPressed Or hKeyPressed Then
            gLastEvent = GetTickCount 'Event occurred - reset time
    Else
        'Check to see if the time limit has passed
        If GetTickCount - gLastEvent > gMaxTime Then 'time limit passed
            gLastEvent = GetTickCount 'reset the time so the above 'event' handler doesn't keep re-firing
            'This next line calls an 'event' handler in the main form.
            'In this case, my main form is named 'MainFrm', and my 'event' function is called 'hTimePassed'.
            'Feel free to rename any or all of the above, but make sure this next line calls the right handler!
            MainFrm.hTimeoutOccurred uElapse 'uElapse is the timer number that fired the event
        Else 'time limit not passed
            Debug.Print "Time limit not exceeded..."
        End If
    End If
    
End Sub

Public Function hTimePassed() As Long
    hTimePassed = GetTickCount - gLastEvent
End Function

Public Sub hStopMonitor(ByVal pTimerNumber As Long, ByVal pWhichWindow As Long)
    KillTimer pWhichWindow, pTimerNumber
End Sub

'Tests to see if caps lock is on
'This function was NOT coded by me.  If you know who it WAS coded by, please tell me, so I can credit them in my source code!
'The 'hKeyPressed' function is based upon another function that I got from this same person, although I have almost completely rewritten it.
Private Function CAPSLOCKON() As Boolean
    Static bInit As Boolean
    Static bOn As Boolean
    If Not bInit Then
        While GetAsyncKeyState(VK_CAPITAL)
        Wend
        bOn = GetKeyState(VK_CAPITAL)
        bInit = True
    Else
        If GetAsyncKeyState(VK_CAPITAL) Then
            While GetAsyncKeyState(VK_CAPITAL)
                DoEvents
            Wend
            bOn = Not bOn
        End If
    End If
    CAPSLOCKON = bOn
End Function

'Tests to see if keys or mouse buttons are pressed, even if this application is minimized, or otherwise is not the active window.
Private Function hKeyPressed() As Boolean
    Dim Shift As Long
    Dim keyState As Long
    
    Shift = GetAsyncKeyState(vbKeyShift)
    keyState = GetAsyncKeyState(vbLeftButton)
    If keyState = -32767 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbRightButton)
    If keyState = -32767 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbMiddleButton)
    If keyState = -32767 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyA)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyB)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyC)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyD)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyE)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyG)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyH)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyI)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyJ)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyK)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyL)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyM)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyN)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyO)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyP)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyQ)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyR)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyS)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyT)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyU)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyV)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyW)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyX)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyY)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyZ)
    If (CAPSLOCKON = True And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    If (CAPSLOCKON = False And Shift = 0 And (keyState And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keyState And &H1) = &H1) Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey1)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey2)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey3)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey4)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey5)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey6)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey7)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey8)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey9)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKey0)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyBack)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyTab)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyReturn)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyShift)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyControl)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyMenu)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyPause)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyEscape)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeySpace)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyEnd)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyHome)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyLeft)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyRight)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyUp)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyDown)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyInsert)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyDelete)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HBA)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HBB)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HBC)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HBD)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HBE)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HBF)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HC0)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HDB)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HDC)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HDD)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(&HDE)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    If Shift <> 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyMultiply)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyDivide)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyAdd)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeySubtract)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyDecimal)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF1)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF2)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF3)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF4)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF5)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF6)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF7)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF8)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF9)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF10)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF11)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyF12)
    If Shift = 0 And (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumlock)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyScrollLock)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyPrint)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyPageUp)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyPageDown)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad1)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad2)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad3)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad4)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad5)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad6)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad7)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad8)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad9)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    keyState = GetAsyncKeyState(vbKeyNumpad0)
    If (keyState And &H1) = &H1 Then GoTo keypressed:
    
    hKeyPressed = False: Exit Function
keypressed:
    hKeyPressed = True: Exit Function
End Function
