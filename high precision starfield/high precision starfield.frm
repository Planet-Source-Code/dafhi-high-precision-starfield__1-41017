VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Caption"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'dafhi's high precision starfield

Dim Tick&
Dim TickSum&
Dim fps!

Dim standardSpeedControl!

'Drag borderless form
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Private Const WM_NCLBUTTONDOWN& = &HA1
    Private Const HTCAPTION& = 2
    
Private Sub Form_Load()
   
 'applied to StandardSpeed in CalcSpeed and SystemTest
 standardSpeedControl = 0.058
   
 'Initial Form Position
 Top = 1000
 Left = 1000
 
 'print color
 ForeColor = vbWhite
   
 ScaleMode = vbPixels
 
 AutoRedraw = True
 
 'WindowState = vbMaximized
 
 Font.Size = 11

 Randomize

 blnRunning = True
 
End Sub
Private Sub Form_Activate()

 'Introduce yourself to some variables in here if
 'you feel up to it.
 InitThings
 
 GenerateBackground
 
 'SystemTest implements a rudimentary one-shot
 'time-based modelling method, which I prefer.
 'If you comment out CalcSpeed inside the loop,
 'there will still be movement
 SystemTest
 
 'To use traditional time-based modelling,
 'Pay attention to the CalcSpeed call inside While blnRunning,
 'and realize that I have a variable called StandardSpeed
 'that holds a floating point value which you can easily
 'tie to all your time-based variables.
 'Or, you may want to use Long value 'Elapsed'
 
 While blnRunning 'press Esc to terminate this loop
 
  Render 'DoEvents is in here
   
  Cls 'Refresh
  
  'Reset print height every frame
  CurrentY = 5
  CurrentX = 10
  
  ''Useful for debugging
  Print "FPS: " & Round(fps, 1)
  Print "   Spacebar changes background"
  
  CalcTick
  CalcFPS
  
  'traditional time-based modelling
  CalcSpeed
  
  EraseStars
 
 Wend
 
 'We've only reached this point if the user has quit
 
 'Erase Refresh buffer pointer
 CopyMemory ByVal VarPtrArray(bDib), 0&, 4
 
 Unload Me
 
 End
 
End Sub

Private Sub InitThings()
 
 'I use this as a parameter in bullet and ship blit calls
 'to show the usefulness of EraseSprites.  I put it in
 'the While loop in Form_Activate.  Set this to False
 'and see what happens
 b_Demo_LeaveNoTrail = True
 
 'See ProcessTiming() in modMainRendering.  I add new stars
 'each time starWaitFrame goes past starWaitFrameUB
 starWaitFrameUB = 0.9
 
 'This sets up some stacks for onscreen / offscreen recycling,
 'as well as some look up tables
 InitStars
 
End Sub
Public Sub InitStandardSpeedVars()
 
 travelSpeed = 1.25! * StandardSpeed
 
 iStarWait = StandardSpeed
 
End Sub
Private Sub GenerateBackground()
Dim XTrack&
Dim x_position!
Dim y_position!
Dim ix!
Dim iy!
Dim B&
Dim right_side!
Dim top_edge!
Dim left_side!
Dim bottom_edge!
   
  'position along scanline
  XTrack = 0
  
  left_side = 32& * (Rnd - 0.5)
  right_side = left_side + 16 * (Rnd - 0.5)
  
  bottom_edge = 32& * (Rnd - 0.5)
  top_edge = bottom_edge + 12 * (Rnd - 0.5)

  ix = (right_side - left_side) / SW
  iy = (top_edge - bottom_edge) / SH
  
  x_position = left_side
  y_position = bottom_edge
  
  '3 elements per pixel
  For N& = 0& To Last_Red_Byte Step 3&
  
   B& = 35& * Sin(x_position! + y_position - Cos(x_position * y_position!)) + 132!
   
   bDib(N&) = B 'N = Blue
   bDib(N& + 1&) = B 'N + 1 = Green
   bDib(N& + 2&) = B 'N + 2 = Red
   
   x_position! = x_position! + ix!
   
   XTrack& = XTrack& + 3&
   If XTrack& >= StepX& Then
    x_position! = left_side!
    y_position! = y_position! + iy!
    XTrack& = 0&
   End If
   
  Next
  
  'FirstFrameDib is what EraseSprites copies from to 'Erase'
  CopyMemory FirstFrameDib(0), bDib(0), Last_Red_Byte
  
End Sub
Public Sub InitStars()
Dim IPLoop&
Dim starDefinition!
Dim starDiameter!
Dim sng!, InverseOpacity As Byte

 'Every few frames, depending on cpu speed, new stars are
 'introduced.  Z value randomness is applied when this is
 'used in AddStar() in modVisualEffects
 InvGalaxySize = -1024
 
 StarsPerBurst = 5
 
 'This is look-up table stuff.  You can control how a star
 'looks as its 'z' moves closer to screen
 For IPLoop = 0& To 999&
  
  sng = IPLoop / 999&
  starDiameter = 4& * sng * sng + 0.5!
  SpecStarSize StarSizes(IPLoop), starDiameter
  
  InverseOpacity = 255
  starDefinition = 0.6! + IPLoop / 800&
  LumenStar StarSizes(IPLoop), _
   InverseOpacity, starDefinition
 
 Next
 
 MaxStars = 25000
 
 'critical
 StarsUB = MaxStars - 1
 StarVectStack.GroupsUB = StarsUB
 
 StarVectStack.Avail.StackUB = StarVectStack.GroupsUB
 StarVectStack.InUse.StackUB = -1
 
 ReDim StarVectStack.Avail.Stack(StarVectStack.GroupsUB)
 ReDim StarVectStack.InUse.Stack(StarVectStack.GroupsUB)
 
 For N = 0 To StarVectStack.GroupsUB
  StarVectStack.Avail.Stack(N) = N
 Next N
 
 ReDim StarVect(StarVectStack.GroupsUB)
 
End Sub

Private Sub CalcFPS()
 
 TickSum = TickSum + Elapsed
 If TickSum& > 1000& Then
  fps = 1000& * FrameCount / TickSum
  FrameCount = 0&
  TickSum = 0&
 End If

End Sub
Private Sub CalcSpeed()
 
 StandardSpeed = Elapsed * standardSpeedControl
 
 InitStandardSpeedVars
 
End Sub

Private Sub Form_Resize()
 WithResize Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
 Case vbKeySpace
  GenerateBackground
 Case vbKeyEscape
  blnRunning = False
 End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ret&
 
 If Button = vbRightButton Then
 
 Else 'Left button
 
  blnMouseDown = True
  ReleaseCapture
  ret& = SendMessageLong&(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  
 End If
 
 blnMouseDown = False
End Sub
Private Sub CalcTick()

 'Will need more accurate timer.  with 200Ghz systems,
 'Elapsed will = 0, and nothing will move.
 Tick = GetTickCount
 Elapsed = Tick - LastTic
 LastTic = Tick

End Sub


Private Sub SystemTest()
Dim NumPasses&

 Elapsed = 0
  
 LastTic = timeGetTime
 
 While Elapsed < 10&
  LoopTestCode
  Elapsed = timeGetTime - LastTic
  NumPasses = NumPasses + 1&
 Wend
 
 StandardSpeed = (Elapsed / NumPasses) * standardSpeedControl * 0.9!
 
 ''If puter is really slow, give some sort of non-frame-skip protection
 If StandardSpeed > 2! Then StandardSpeed = 2!
 
 InitStandardSpeedVars
   
End Sub
Private Sub LoopTestCode()
Dim N102&
 
 For N102 = 1& To 2&
 Refresh
 EraseStars
 Next N102
 
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
 blnRunning = False

End Sub
