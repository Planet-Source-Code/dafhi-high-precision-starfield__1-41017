Attribute VB_Name = "ModPublic"
Option Explicit

Public SW&        '& means  As Long
Public HalfWidth! '! means  As Single
Public SH&
Public HalfHeight!

Public Xmax&, Ymax&

Public blnRunning As Boolean

Public Elapsed&
Public LastTic&
Public StandardSpeed!

Public FirstFrameDib() As Byte 'also used by Blit subs

'1d processing
Public Loca&, Loca1&, Loca2&
Public StepX&
Public Last_Blue_Byte&
Public Last_Red_Byte&
Public ViewPort_Right&
Public ViewPort_TopLeft&
'Public ViewPort_Right_Blue&
Public c_PadBytes&

Public EraseSpriteCount&

'all-purpose
Public N&

Public Const B255 As Byte = 255
Public Const BYT2 As Byte = 2

Public blnMouseDown As Boolean

Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Sub WithResize(vform As Form)

 SW = vform.ScaleWidth
 SH = vform.ScaleHeight
 HalfWidth! = SW / 2
 HalfHeight! = SH / 2
  
 vform.Picture = CreatePicture(SW, SH, 24)
 GetObjectAPI vform.Picture, Len(BM), BM
 
 StepX = BM.bmWidthBytes
 ViewPort_Right = BM.bmWidth - 1
 Ymax = BM.bmHeight - 1: Xmax = ViewPort_Right * 3
 ViewPort_TopLeft = Ymax * StepX
 'ViewPort_Right_Blue = (ViewPort_Right) * 3
 Last_Blue_Byte = ViewPort_TopLeft + Xmax
 Last_Red_Byte = Last_Blue_Byte + 2
 c_PadBytes = BM.bmWidthBytes - BM.bmWidth * BM.bmBitsPixel / 8
 
 With tSA 'pointer to bitmap structure
  .cbElements = 1
  .cDims = 1
  .cElements = BM.bmHeight * BM.bmWidthBytes
  .pvData = BM.bmBits
 End With
 CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
 
 vform.Caption = SW & " x " & SH
 
 ReDim FirstFrameDib(Last_Red_Byte)
 
End Sub



' ==================== Storage =====================

 'CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
 'CopyMemory FirstFrameDib(0), bDib(0), Last_Red_Byte
 'CopyMemory ByVal VarPtrArray(bDib), 0&, 4

