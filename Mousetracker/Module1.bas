Attribute VB_Name = "Module1"
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'************************************************************************************************
'*  Author: Pradeep Singh (Pioneer Soft.)                                                       *
'************************************************************************************************
'This is my first submission on planet source code...
'This an example that I wanted to share with everyone, I tried to search various sites to find out
'a cool way to track the mouse enter and mouse exit for a control but all the examples that I found
'were either non-working or they used timers to accomplish the tracking, some of them used subclassing
'but the code was very complex and buggy. So I finally tried to create my own routines and I was able
'to find a way out with a very simple subclassing technique. Timers Take a Lot of system resources
'and can be very resource consuming if you want to track several Controls.
'I have tried to comment everywhere i felt necessary but still if you face any problem you can
'e-mail me: pradeep.ansh.sumit@gmail.com,
'I have taken Ideas from several subclassing routines and
'if you find that some of the code matches your routines please accept my thanks and credits for
'Ideas.
' Please Do Vote If You Like The Code and Idea
'!!!!!Happy Coding!!!!!
'************************************************************************************************
'*                      Disclaimer                                                              *
'* Obviously no warranty or liability, or any responsibility in any way imaginable, is expressed*
'* or implied.                                                                                  *
'* Use this software in any way you wish under a relaxed GNU, but all responsibilities          *
'* are yours entirely. This example program is only a demonstration for educational purposes    *
'************************************************************************************************
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\



'************************************************************************************************
'Enumerators                                                                                    *
'************************************************************************************************
'this enum contains constants that we pass on to 'TrackMouseEvent' API Function and tells that function
'about the mouse events that we want to track
Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1             'Track the Mouse Hover in the control window
    TME_LEAVE = &H2             'Track the Mouse Leave Event in the control window
    TME_QUERY = &H40000000      'I don't Know about this Flag, do Tell me if you know or find out
    TME_CANCEL = &H80000000     'I don't Know about this Flag, do Tell me if you know or find out
End Enum

'************************************************************************************************
'Types/ Structures                                                                              *
'************************************************************************************************
'this structure contains variables that we pass on to 'TrackMouseEvent' API Function and tells that function
'about the window and other details that we want to add that we want to track
Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                   As Long                    'Pass The Size of This Variable using Len()Function
    dwFlags                  As TRACKMOUSEEVENT_FLAGS   'Flags for Tracking, see above section
    hwndTrack                As Long                    'Handle to the Window of Control hWnd
    dwHoverTime              As Long                    'Hover Time after which it send the message to control window
End Type

'************************************************************************************************
'Local Variables                                                                                *
'************************************************************************************************
Private boolHover As Boolean                'Keeps Track if Mouse is Hovering the Control or Not
Private hTME As TRACKMOUSEEVENT_STRUCT      'Structure for Mouse Hover Tracking
Private lTME As TRACKMOUSEEVENT_STRUCT      'Structure for Mouse Leave Tracking
Private PrevInstance As Long                'Address of Original WndProc of the control window
Private ParentFrm As Object                 'Control for which we want to track the Mouse in and Mouse Out

'************************************************************************************************
'API Declares                                                                                   *
'************************************************************************************************
Private Const WM_MOUSEMOVE = &H200                          'Message Received when mouse moves over the window
Private Const WM_MOUSELEAVE As Long = &H2A3                 'Message Received when mouse Leave the window
Private Const WM_MOUSEHOVER = &H2A1                         'Message Received when mouse Hovers over the window

Public Const GWL_WNDPROC = (-4)                             'Argument for Set window Long that tells the function that
                                                            'we want get/alter the oringinal WndProc of the Window
'Sets various attributes of the window
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'calls a WndProc of a window by it's Address
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Tracks the mouse events for the specified window and send messages to the window
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
'Replacement Function for Original WndProc of the control
'This function must always be a public function
Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Modify the lTme Structure to monitor MouseLeave events
With lTME
.cbSize = Len(lTME)     'This is basically to tell the Function about the Size of this structure
.dwFlags = TME_LEAVE    'Events that we want to track
.hwndTrack = hwnd       'hWnd of the Parent Control
End With
'Modify the hTme Structure to monitor MouseHover events
With hTME
.cbSize = Len(hTME)     'Same  as lTME (See above)
.dwFlags = TME_HOVER
.dwHoverTime = 1
.hwndTrack = hwnd
End With
'Process the Messages Received for the Window
Select Case uMsg
    Case WM_MOUSEMOVE:  'If mouse moves
    TrackMouseEvent lTME    'Start The Tracking Routine for Mouse Leave
    TrackMouseEvent hTME    'Start The Tracking Routine for Mouse Hover
    WndProc = CallWindowProc(PrevInstance, hwnd, uMsg, wParam, lParam)  'Pass on the received message to original wndProc for
                                                                        'Default Processing
    Case WM_MOUSELEAVE:
    boolHover = False               'Set The Flag to False Indicating that the Mouse Has Left The Window
    ParentFrm.Print "Mouse Left"    'You Can Place Your Custom Code here to do what ever you feel when the mouse leaves
    WndProc = CallWindowProc(PrevInstance, hwnd, uMsg, wParam, lParam)  'Pass on the received message to original wndProc for
                                                                        'Default Processing
    Case WM_MOUSEHOVER:
    If Not boolHover = True Then    'Check if the mouse was already Hovering or it has just started now
    boolHover = True                'If Mouse has just Entered set the Flag to True Indicating that the mouse has entered the window
    ParentFrm.Print "Mouse Hover"   'You Can Place Your Custom Code here to do what ever you feel when the mouse leaves
    WndProc = CallWindowProc(PrevInstance, hwnd, uMsg, wParam, lParam)  'Pass on the received message to original wndProc for
                                                                        'Default Processing
    End If
End Select
'Pass on every single message to Default(Original WndProc Function of the Window for Default Processing and to Protect
'from any possible errors (Altering or Hiding Messages can be dangerous!!!)
WndProc = CallWindowProc(PrevInstance, hwnd, uMsg, wParam, lParam)
End Function

'This Function Starts Subclassing of the Object Window
Public Function Attach(vObject As Object) As Boolean
'check if we have a valid Hwnd
If Not vObject.hwnd = 0 Then
Set ParentFrm = vObject     'set the local variable
PrevInstance = SetWindowLongA(vObject.hwnd, GWL_WNDPROC, AddressOf WndProc) 'store the address of original WndProc
                                                                            'and also set our WndProc Function as Message
                                                                            'Processing function for window
Attach = True   'Return True and Exit out of Function
Exit Function
End If
Attach = False  'You will reach here only if there is some error
End Function

'This Function Ends Subclassing of the Object Window
'Always Unhook/Unsublass the parent object before terminating the application or you might see a GPF error message
'and can be very dangerous as well
Public Function DeAttach(vObject As Object) As Boolean
'check if we have a valid Hwnd
If Not vObject.hwnd = 0 Then
SetWindowLongA vObject.hwnd, GWL_WNDPROC, PrevInstance                      'Restore the Original WndProc Function as Message
                                                                            'For the window
DeAttach = True         'Return True and Exit out of Function
Set vObject = Nothing   'Destroy the local variable
Exit Function
End If
DeAttach = False
Set vObject = Nothing   'You will reach here only if there is some error
End Function
