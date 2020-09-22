Attribute VB_Name = "modFunc"
Option Explicit

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(1) As SAFEARRAYBOUND
End Type

Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Type RGBQUAD
    rgbRed As Byte
    rgbGreen As Byte
    rgbBlue As Byte
    rgbReserved As Byte
End Type

Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public bDib() As Byte
Public BM As BITMAP
Public tSA As SAFEARRAY1D

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)

Declare Function GetTickCount Lib "kernel32" () As Long

Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any) As Long
Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC&, pBitmapInfo As BITMAPINFO, ByVal un&, lplpVoid&, ByVal handle&, ByVal dw&) As Long

Declare Function RedrawWindow Lib "user32" (ByVal hwnd&, lprcUpdate As RECT, ByVal hrgnUpdate&, ByVal fuRedraw&) As Long

Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, IPic As IPicture) As Long

Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Function CreatePicture(ByVal nWidth&, ByVal nHeight&, ByVal BitDepth&) As Picture
Dim Pic As PicBmp, IID_IDispatch As GUID
Dim BMI As BITMAPINFO
With BMI.bmiHeader
.biSize = Len(BMI.bmiHeader)
.biWidth = nWidth
.biHeight = nHeight
.biPlanes = 1
.biBitCount = BitDepth
End With
Pic.hBmp = CreateDIBSection(0, BMI, 0, 0, 0, 0)
IID_IDispatch.Data1 = &H20400: IID_IDispatch.Data4(0) = &HC0: IID_IDispatch.Data4(7) = &H46
Pic.Size = Len(Pic)
Pic.Type = vbPicTypeBitmap
OleCreatePictureIndirect Pic, IID_IDispatch, 1, CreatePicture
If CreatePicture = 0 Then Set CreatePicture = Nothing
End Function

Function GetPicture(ByVal Pic&, outAry() As Byte) As Boolean
GetObjectAPI Pic, Len(BM), BM
ReDim outAry(BM.bmWidthBytes * BM.bmHeight - 1)
CopyMemory outAry(0), ByVal BM.bmBits, BM.bmWidthBytes * BM.bmHeight
GetPicture = True
End Function

Function SetPicture(ByVal Pic&, inAry() As Byte) As Boolean
GetObjectAPI Pic, Len(BM), BM
If LBound(inAry) <> 0 Or UBound(inAry) <> BM.bmWidthBytes * BM.bmHeight - 1 Then Exit Function
CopyMemory ByVal BM.bmBits, inAry(0), BM.bmWidthBytes * BM.bmHeight
SetPicture = True
End Function


