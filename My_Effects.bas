Attribute VB_Name = "My_Effects"
Public Type myBitmap
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Public bmpBits() As Byte
Public hBmp As myBitmap

Public Sub GetBits(pBox As PictureBox)
Dim iRet As Long
    'Get the bitmap header
    iRet = GetObject(pBox.Picture.Handle, Len(hBmp), hBmp)
    'Resize to hold image data
    ReDim bmpBits(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into bmpBits array
    iRet = GetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, bmpBits(0, 0, 0))
End Sub

Public Sub SetBits(pBox As PictureBox)
Dim iRet As Long
    'Set the new image data back onto pBox
    iRet = SetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, bmpBits(0, 0, 0))
    'Erase bmpBits because we finished with it now
    Erase bmpBits
End Sub
