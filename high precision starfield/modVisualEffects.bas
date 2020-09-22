Attribute VB_Name = "modVisualEffects"
Option Explicit

'common to module
Dim iR!
Dim iG!
Dim iB!

'blit vars
Dim SpriteX&
Dim SpriteY&
Dim SpriteArrayDepth&
Dim ClipWidthLeft&
Dim CLipWidthBot&
Dim CLipWidthTop&
Dim DrawWidth&
Dim Left_&
Dim Right_&
Dim Top_&
Dim Bot_&
Dim DrawY&
Dim DrawX&
Dim DrawBot&
Dim DrawTop&
Dim DrawLeft&
Dim DrawRight&
Dim AddDrawWidthBytes&
Dim DrawWidthBytes&

'compute star size
Dim initial_offset_X!
Dim initial_offset_Y!
Dim delta_x!
Dim delta_y!
Dim delta_ySq!
Dim half!
Dim Rounded&
Dim pp5!
Dim mult2!
Dim add2!
Dim baseleft!
Dim Bright&
Dim right_side_scale!
Dim left_side_scale!
Dim top_edge_scale!
Dim bottom_edge_scale!

Private Type ParticlePrecisionDims
 wWide As Single
 hHigh As Single
 wDiv2 As Single
 hDiv2 As Single
End Type

Public Type StarStruct
 chIntensity As Byte
 definition As Single
 intens As Single
 def As Single
 centerX As Single
 centerY As Single
 iRed As Single
 iGrn As Single
 iBlu As Single
End Type

Public Type StarSprite
 Dims As ParticlePrecisionDims
 LngWide As Long
 LngHigh As Long
 SHigh As Long
 SRight As Long
 AddHalfW As Long
 AddHalfH As Long
 iBriteX As Single
 iBriteY As Single
 wall_x As Single
 floor_y As Single
 IPS As StarStruct
End Type

Type StarVec
 px As Single
 py As Single
 pz As Single
 chRed As Byte
 chGreen As Byte
 chBlue As Byte
 bErasing As Boolean
End Type

Public InvGalaxySize&

Public travelSpeed As Single

Public StarSizes(999) As StarSprite

Dim WhichAry() As SAFEARRAY1D

Type StarVectStackType
 GroupsUB As Long
 Avail As StackedArray
 InUse As StackedArray
End Type

Public StarVectStack As StarVectStackType

Public StarVect() As StarVec

Public MaxStars As Long
Public StarsUB As Long

'EraseStars
Dim CheckErase() As Single
Private Type EraseInfo
 FirstEraseByt As Long
 TopLeftEraseByt As Long
 EraseWidthBytes As Long
 SprArrayDep As Long
 SprWidth As Long
End Type
Public EraseData() As EraseInfo

Dim StarLoop&

Public b_Demo_LeaveNoTrail As Boolean

Public Sub DrawStars()
Dim Z1&
Dim z!

 StarLoop = -1&
 Do While StarLoop < StarVectStack.InUse.StackUB
  StarLoop = StarLoop + 1&
  N = StarVectStack.InUse.Stack(StarLoop)
 
 'StarVect(N).px = 0 and .py = 0 would be screen center.
 
 'a .pz of 0 is 'farthest from screen' and a .pz of
 '1000 matches screen plane
  
  z = StarVect(N).pz
  Z1 = RealRound(z)
 'Z1 should be between 0 and 999 because it is used
 'as an array element for a star size look-up-table
  If Z1 < 1000& Then
   If Z1 > -1 Then
    BlitStar StarVect(N), StarSizes(Z1)
   End If
   StarVect(N).pz = StarVect(N).pz + travelSpeed 'computed in InitStandar in Form's code
  Else
   DeleteStar StarLoop
  End If
 Loop
 
End Sub
Public Sub AddStar(x!, y!, chRed As Byte, chGreen As Byte, chBlue As Byte, Optional blnErasing As Boolean = False)
Dim SLoop&

 If StarVectStack.Avail.StackUB > -1 Then
  
  ''Stack prep
  
  'Up pointer for InUse stack
  StarVectStack.InUse.StackUB = StarVectStack.InUse.StackUB + 1&
  
  'Read value off the top of the Avail stack
  N = StarVectStack.Avail.Stack(StarVectStack.Avail.StackUB)
  
  'Store value onto top of InUse stack
  StarVectStack.InUse.Stack(StarVectStack.InUse.StackUB) = N
  
  'Lower pointer for Avail stack
  StarVectStack.Avail.StackUB = StarVectStack.Avail.StackUB - 1&
  
  'new star info
  With StarVect(N)
  .px = x
  .py = y
  .pz = Rnd * InvGalaxySize&
  .chBlue = chBlue
  .chGreen = chGreen
  .chRed = chRed
  .bErasing = blnErasing
  End With
  
 End If
End Sub

Public Sub BlitStar(StarVect As StarVec, StarSize As StarSprite)
Dim BPLoop&, x!, y!, iBrightX!, iBrightY!, distanceratio!
  
  distanceratio = 1000& / (1010& - StarVect.pz)
  
  x = HalfWidth + StarVect.px * distanceratio
  If x > -1! Then
  If x < SW Then
  y = HalfHeight + StarVect.py * distanceratio
  If y > -1! Then
  If y < SH Then
  SpriteX = RealRound2(x)
  SpriteY = RealRound2(y)
 
  Left_& = SpriteX& - StarSize.AddHalfW
  Top_& = StarSize.AddHalfH - SpriteY
 
  Right_& = Left_& + StarSize.SRight
  Bot_& = Top_& - StarSize.SHigh
  
  If Bot_ < 0& Then
   CLipWidthBot& = -Bot_&
   DrawBot& = 0&
  Else
   CLipWidthBot = 0&
   DrawBot& = Bot_& * StepX
  End If
 
  If Right_ > ViewPort_Right Then
   DrawRight = ViewPort_Right
  Else
   DrawRight = Right_
  End If
 
  If Left_& < 0& Then
   ClipWidthLeft& = -Left_&
   DrawLeft = 0&
  Else
   ClipWidthLeft& = 0&
   DrawLeft = Left_
   DrawBot& = DrawBot + Left_& * 3&
  End If
  
  DrawWidth = DrawRight - DrawLeft + 1&
 
  CLipWidthTop& = Top_& - Ymax
  
  If CLipWidthTop& < 0& Then
   DrawTop& = DrawBot + StepX * (Top_& - Bot_&)
  Else
   DrawTop& = DrawBot + StepX * (Ymax - Bot_&)
  End If
 
  DrawWidthBytes& = DrawWidth& * 3&
  AddDrawWidthBytes& = DrawWidthBytes& - 3&
 
  If StarVect.bErasing Then
   EraseSpriteCount& = EraseSpriteCount& + 1&

   ReDim Preserve EraseData(1 To EraseSpriteCount&)
  
   EraseData(EraseSpriteCount&).FirstEraseByt = DrawBot&
   EraseData(EraseSpriteCount&).TopLeftEraseByt = DrawTop&
   EraseData(EraseSpriteCount&).EraseWidthBytes = DrawWidthBytes

  End If
    
  DrawX = RealRound2(x)
  initial_offset_X = (x - DrawX - ClipWidthLeft) / StarSize.Dims.wDiv2
 
  DrawY = RealRound2(y)
  initial_offset_Y = (y - DrawY + CLipWidthBot) / StarSize.Dims.hDiv2
  
  delta_y = StarSize.floor_y + initial_offset_Y
  baseleft = StarSize.wall_x - initial_offset_X
 
  iBrightX = StarSize.iBriteX
  iBrightY = StarSize.iBriteY
  
  For DrawY& = DrawBot& To DrawTop& Step StepX&
   DrawRight& = DrawY& + AddDrawWidthBytes&
   delta_ySq! = delta_y * delta_y
   delta_x! = baseleft
   For DrawX& = DrawY& To DrawRight& Step 3&
    Bright& = StarSize.IPS.def * _
    (1! - Sqr(delta_x! * delta_x! + delta_ySq!))
    If Bright& > 255& Then
     Bright& = 255&
    ElseIf Bright& < 0& Then
     Bright& = 0&
    End If
    If Bright& <> 0& Then
     Loca1& = DrawX& + 1&
     Loca2& = DrawX& + 2&
     iB! = bDib(DrawX&) + Bright& * StarVect.chBlue / 255&
     iG! = bDib(Loca1&) + Bright& * StarVect.chGreen / 255&
     iR! = bDib(Loca2&) + Bright& * StarVect.chRed / 255&
     If iB > 255! Then iB = 255!
     If iG > 255! Then iG = 255!
     If iR > 255! Then iR = 255!
     bDib(Loca2&) = iR!
     bDib(Loca1&) = iG!
     bDib(DrawX&) = iB!
    End If
    delta_x! = delta_x! + iBrightX!
   Next
   delta_y! = delta_y! + iBrightY!
  Next
  
  End If
  End If
  End If
  End If
  
End Sub
Public Sub SpecStarSize(AS1 As StarSprite, ByVal starDiameter!)

 AS1.Dims.wWide = starDiameter
 AS1.Dims.hHigh = starDiameter
 
 half = AS1.Dims.wWide / BYT2
 Rounded = RealRound2(half)
 pp5 = Rounded + 0.5!
 mult2 = pp5 * BYT2
 add2 = mult2 + 2!
 
 AS1.LngWide = RealRound2(add2)
 AS1.Dims.wDiv2 = half
 
 half = AS1.Dims.hHigh / BYT2
 Rounded = RealRound2(half)
 pp5 = Rounded + 0.5!
 mult2 = pp5 * BYT2
 add2 = mult2 + 2!
 
 AS1.LngHigh = RealRound2(add2)
 AS1.Dims.hDiv2 = half
 
 With AS1
  .SHigh = .LngHigh - 1&
  .SRight = .LngWide - 1&
  .AddHalfW = .SRight / 2&
  .AddHalfH = .SHigh / 2& + SH
 End With
 
 right_side_scale = (AS1.LngWide / AS1.Dims.wWide)
 left_side_scale = -right_side_scale
 
 top_edge_scale = (AS1.LngHigh / AS1.Dims.hHigh)
 bottom_edge_scale = -top_edge_scale

 AS1.iBriteX! = (right_side_scale! - left_side_scale!) / AS1.LngWide
 AS1.iBriteY! = (top_edge_scale! - bottom_edge_scale!) / AS1.LngHigh
  
 AS1.floor_y = (bottom_edge_scale - AS1.iBriteY / BYT2)
 AS1.wall_x = (left_side_scale - AS1.iBriteX / BYT2)

End Sub
Public Sub LumenStar(Star As StarSprite, chIntensity As Byte, definition!)
Dim tmpIntensity!
 
 Star.IPS.chIntensity = chIntensity 'Star.IPS.sBright
 Star.IPS.def = definition * Star.IPS.chIntensity
 
 tmpIntensity = Star.IPS.chIntensity / B255
 Star.IPS.intens = tmpIntensity / B255
 
End Sub

Public Function RealRound(ByVal sngValue!) As Long
Dim diff!
 'This function rounds .5 up
 
 RealRound = Int(sngValue)
 diff = sngValue - RealRound
 If diff >= 0.5! Then RealRound = RealRound + 1&

End Function
Public Function RealRound2(ByVal sngValue!) As Long
Dim diff!
 'This function rounds .5 down
 
 RealRound2 = Int(sngValue)
 diff = sngValue - RealRound2
 If diff > 0.5! Then RealRound2 = RealRound2 + 1&

End Function
Public Sub DeleteStar(Element&)
 'Increase the OffScreen bullet stack pointer
 StarVectStack.Avail.StackUB = StarVectStack.Avail.StackUB + 1&
 'Copy InUse.Stack(Element) to Avail.Stack(OffScrUBound)
 StarVectStack.Avail.Stack(StarVectStack.Avail.StackUB) = StarVectStack.InUse.Stack(Element)
 'Copy InUse.Stack(UBound) to InUse.Stack(Element)
 StarVectStack.InUse.Stack(Element) = StarVectStack.InUse.Stack(StarVectStack.InUse.StackUB)
 'Lower the InUse bullet stack pointer
 StarVectStack.InUse.StackUB = StarVectStack.InUse.StackUB - 1&
 Element = Element - 1&
End Sub
Private Sub AddToStars()
 StarVectStack.InUse.StackUB = StarVectStack.InUse.StackUB + 1&
 StarVectStack.InUse.Stack(StarVectStack.InUse.StackUB) = _
  StarVectStack.Avail.Stack(StarVectStack.Avail.StackUB)
 StarVectStack.Avail.StackUB = StarVectStack.Avail.StackUB - 1&
End Sub

Public Sub EraseStars()
 
 For N& = 1& To EraseSpriteCount&
  
  With EraseData(N&)
    
    If .EraseWidthBytes& > 0& Then
     DrawBot& = .FirstEraseByt&
     DrawTop& = .TopLeftEraseByt&
     DrawWidth = .EraseWidthBytes - 3&
     For Loca& = DrawBot& To DrawTop& Step StepX
      AddDrawWidthBytes = Loca + DrawWidth
      For DrawX = Loca To AddDrawWidthBytes Step 3&
       Loca1 = DrawX + 1&
       Loca2 = DrawX + 2&
       bDib(DrawX) = FirstFrameDib(DrawX)
       bDib(Loca1) = FirstFrameDib(Loca1)
       bDib(Loca2) = FirstFrameDib(Loca2)
      Next DrawX
     Next
    End If
   
  End With 'EraseData(N)
 
 Next N& 'Next Sprite
 
 EraseSpriteCount = 0&

End Sub

