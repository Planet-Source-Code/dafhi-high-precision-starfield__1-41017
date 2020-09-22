Attribute VB_Name = "modMainRendering"
Option Explicit

Public starWaitFrame!
Public starWaitFrameUB!
Public iStarWait!

Type StackedArray
 StackUB As Long
 Stack() As Long
End Type

Public FrameCount&
Public StarsPerBurst&

Public Sub Render()
  
 DoEvents
   
 ProcessTiming
  
 DrawStars
    
 FrameCount& = FrameCount& + 1&
   
End Sub

Public Sub ProcessTiming()
Dim chRed As Byte
Dim chGreen As Byte
Dim chBlue As Byte
Dim PBLoops&
  
  'starWaitFrameUB is initialized in InitThings in Form's code
  If starWaitFrame >= starWaitFrameUB! Then

   For PBLoops = 1& To StarsPerBurst&
    chRed = 255& - 128& * Rnd
    chGreen = chRed
    chBlue = 255&
    AddStar SW * (Rnd - 0.5!), SH * (Rnd - 0.5!), _
            chRed, chGreen, chBlue, b_Demo_LeaveNoTrail
   Next
    
   starWaitFrame = 0!

  Else
   starWaitFrame = starWaitFrame + iStarWait 'iBulletWait computed in InitStandardSpeedVars
  End If
   
End Sub



