Attribute VB_Name = "mMYRGB"
Public Type tRGBcol
    r As Byte
    G As Byte
    b As Byte
    
    C As Long
    
End Type
'Public Const MAXIND As Integer = 10 '11 '11 '=(NumColor -1) Max is 255 unless you modify the steps for displaying the table

Public MAXIND As Integer

'Public Pale(0 To 2, 0 To MAXIND) As Integer 'Index of colors
'Public AvgTbl(0 To 2, 0 To MAXIND) As Long 'Summation of pixels for each index
'Public CntTbl(0 To MAXIND) As Long 'Number of pixels belonging to each index
Public PALE() As Integer 'Index of colors
Public AvgTbl() As Long 'Summation of pixels for each index
Public CntTbl() As Long 'Number of pixels belonging to each index


Public Kmul(-2 To 2, -2 To 2) As Single 'integer
Public KmulBLUR(-3 To 3, -3 To 3) As Single

Public KmulBlurD As Integer


Public mPIC() As tRGBcol
Public CopiaPic() As tRGBcol
Public BlurPic() As tRGBcol

Public ContoPic() As tRGBcol
Public QuantizedPic() As tRGBcol

Public ResizedPic() As tRGBcol


Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long

Public Const STRETCHMODE = vbPaletteModeNone 'You can find other modes in the "PaletteModeConstants" section of your Object Browser


Public aD(8) As Integer

'Public Const SS As Single = 1.61803 * 1.61803 '* 2 '* 1.61803
Public Const SS As Single = 1.61803 * 2

Public FastAVG(7650) As Integer
Public FastPower(-255 To 255) As Long
Public FastRoot(0 To 195075) As Integer

Public NearEst() As Long

Sub InitFastAVG()
Dim I As Integer
For I = 0 To 7650
    FastAVG(I) = I / 3
Next I
End Sub

Sub InitFastPower()
Dim I As Integer
For I = -255 To 255
    FastPower(I) = CLng(I) * CLng(I)
    Debug.Print FastPower(I)
    
Next
'Stop
End Sub

Sub InitFastRoot()
Dim I As Long
For I = 0 To 195075
    FastRoot(I) = Round(Sqr(I))
Next

End Sub

Sub LongToRGB(RGBcol As Long, ByRef r As Byte, ByRef G As Byte, ByRef b As Byte)

'If RGBcol < 0 Then RGBcol = 0: r = 0: G = 0: b = 0: Exit Sub ' Stop


r = RGBcol And &HFF ' set red
G = (RGBcol And &H100FF00) / &H100 ' set green
b = (RGBcol And &HFF0000) / &H10000 ' set blue

End Sub

