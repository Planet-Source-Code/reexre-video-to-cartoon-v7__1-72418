VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "myEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
'--------------------------------------------------------------------------------
' Based on SetBitmapBits by DreamVB from Planet Source Code
'
'

Public Enum Effect
    eContour
    eEXblur
    eMyQuant
    eSepa
    ePGM
    ePGM_C
End Enum


Private Type Bitmap
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

Private Sbyte() As Byte
Private Sbyte2() As Byte
Private ContByte() As Byte
Private BlurByte() As Byte

Private SepaByte() As Byte

'Private PGMByte() As Byte
'Private PGM_C_Byte() As Byte



Private hBmp As Bitmap

Private pW As Integer
Private PH As Integer
Private PB As Integer

Public Sub SetSource(pbox As PictureBox)
'Public Sub GetBits(pBoxPicHand As Long)
Dim iRet As Long
'Get the bitmap header
iRet = GetObject(pbox.Image.Handle, Len(hBmp), hBmp)
'   iRet = GetObject(pBoxPicHand, Len(hBmp), hBmp)

'Resize to hold image data
ReDim Sbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
'Get the image data and store into Sbyte array
'iRet = GetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))
iRet = GetBitmapBits(pbox.Image.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))


pW = hBmp.bmWidth - 1
PH = hBmp.bmHeight - 1
PB = (hBmp.bmBitsPixel \ 8) - 1

'Stop


'ReDim PGMByte(0 To PB, 0 To pW, 0 To PH)


End Sub

Public Sub PutEffToPic(pbox As PictureBox, wEff As Effect)

'public Sub SetBits(pBoxPicHand As Long)

Dim iRet As Long
'Set the new image data back onto pBox

Select Case wEff
    Case 99
        iRet = SetBitmapBits(pbox.Image.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))
        Erase Sbyte: Stop
    Case eCountour
        
        iRet = SetBitmapBits(pbox.Image.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, ContByte(0, 0, 0))
        'Erase ContByte
        
    Case eEXblur, eMyQuant
        
        iRet = SetBitmapBits(pbox.Image.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, BlurByte(0, 0, 0))
        'Erase BlurByte
    Case eSepa
        
        iRet = SetBitmapBits(pbox.Image.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, SepaByte(0, 0, 0))
    
    Case ePGM, ePGM_C
    
        
        iRet = SetBitmapBits(pbox.Image.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))
      
        
End Select
'iRet = SetBitmapBits(pBoxPicHan, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))

'' 'Erase Sbyte because we finished with it now
'' Erase Sbyte
End Sub

Public Sub effNoise(ByVal Value As Integer)
Dim X As Long
Dim Y As Long
Dim R As Integer, G As Integer, B As Integer
Dim iRnd As Integer

'  Call GetBits(pDst)
'Now we can play with the image data
For X = 0 To hBmp.bmWidth - 1
    For Y = 0 To hBmp.bmHeight - 1
        'Noise
        iRnd = Int(Rnd * Value)
        'Invert the colors
        R = Sbyte(0, X, Y) + iRnd 'Red
        G = Sbyte(1, X, Y) + iRnd 'Green
        B = Sbyte(2, X, Y) + iRnd 'Blue
        
        If (R < 0) Then R = 0
        If (G < 0) Then G = 0
        If (B < 0) Then B = 0
        If (R > 255) Then R = 255
        If (G > 255) Then G = 255
        If (B > 255) Then B = 255
        
        'Set colors
        Sbyte(0, X, Y) = R
        Sbyte(1, X, Y) = G
        Sbyte(2, X, Y) = B
    Next Y
Next X
'Here we set the new bits
'  Call SetBits(pDst)
'And Refresh the picturebox
'  Call pDst.Refresh
End Sub


Public Sub effCONTOUR(Enhanced As Integer)
Dim X As Long
Dim Y As Long

Dim iX As Long
Dim iY As Long

Dim X2 As Long
Dim Y2 As Long

Dim vMinR As Integer
Dim vMinG As Integer
Dim vMinB As Integer

Dim vR As Integer
Dim vG As Integer
Dim vB As Integer

Dim R As Integer
Dim G As Integer
Dim B As Integer

Dim Xp3 As Long
Dim Xm3 As Long
Dim Xp2 As Long
Dim Xm2 As Long
Dim yp3 As Long
Dim ym3 As Long
Dim yp2 As Long
Dim ym2 As Long

'ReDim Sbyte2(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
'ReDim ContByte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
ReDim Sbyte2(0 To PB, 0 To pW, 0 To PH) As Byte
ReDim ContByte(0 To PB, 0 To pW, 0 To PH) As Byte
If Enhanced > 0 Then ReDim ContByte2(0 To PB, 0 To pW, 0 To PH) As Long


For Y = 1 To PH - 1
    For X = 1 To pW - 1
        
        
        ' Sbyte2(0, x, y) = Sbyte(0, x, y)
        ' Sbyte2(1, x, y) = Sbyte(1, x, y)
        ' Sbyte2(2, x, y) = Sbyte(2, x, y)
        
        
        vMinR = 255
        vMinG = 255
        vMinB = 255
        
        For iX = -1 To 1
            X2 = X + iX
            For iY = -1 To 1
                
                'If IX <> 0 And IY <> 0 Then
                vR = Sbyte(0, X2, Y + iY)
                If vR < vMinR Then vMinR = vR
                
                vG = Sbyte(1, X2, Y + iY)
                If vG < vMinG Then vMinG = vG
                
                vB = Sbyte(2, X2, Y + iY)
                If vB < vMinB Then vMinB = vB
                'End If
                
            Next iY
        Next iX
        
        ' With ContoPic(X, Y)
        '     .R = .R - vMinR
        '     .G = .G - vMinG
        '     .B = .B - vMinB
        
        R = Sbyte(0, X, Y)
        G = Sbyte(1, X, Y)
        B = Sbyte(2, X, Y)
        
        R = (R - vMinR)
        G = (G - vMinG)
        B = (B - vMinB)
        
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
        
        
        '*************
        R = R * 10 ' 12 '10
        G = G * 10 '12
        B = B * 10 '12
        '*************
        
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
        
        Sbyte2(0, X, Y) = 255 - R
        Sbyte2(1, X, Y) = 255 - G
        Sbyte2(2, X, Y) = 255 - B
        'r = 255 - r
        'G = 255 - G
        'b = 255 - b
        
        'r = FastAVG(r + G + b)
        
        'ContByte(0, x, y) = r
        'ContByte(1, x, y) = r
        'ContByte(2, x, y) = r
        
    Next X
Next Y


'Blurring contour

For Y = 1 To PH - 1
    For X = 1 To pW - 1
        
        R = 0
        G = 0
        B = 0
        
        
        For iX = -1 To 1
            X2 = X + iX
            For iY = -1 To 1
                Y2 = Y + iY
                R = R + Sbyte2(0, X2, Y2)
                G = G + Sbyte2(1, X2, Y2)
                B = B + Sbyte2(2, X2, Y2)
            Next iY
        Next iX
        
        'r = FastAVG(r + G + b) \ 9
        R = (R + G + B) \ 20 '27
        
        
        R = R + (CInt(Sbyte2(0, X, Y)) + CInt(Sbyte2(1, X, Y)) + CInt(Sbyte2(2, X, Y))) * 0.14
        
        'If r < 0 Then r = 0
        If R > 255 Then R = 255
        
        ContByte(0, X, Y) = R
        ContByte(1, X, Y) = R
        ContByte(2, X, Y) = R

        
    Next X
Next Y






'****************************************************************************

If Enhanced > 0 Then
For Y = 0 To PH - 1
    For X = 0 To pW - 1
        ContByte2(0, X, Y) = ContByte(0, X, Y)
    Next X
Next Y


For Y = 4 To PH - 5
    
    ym3 = Y - 3
    yp3 = Y + 3
    ym2 = Y - 2
    yp2 = Y + 2
    
    For X = 4 To pW - 5
        
        Xm3 = X - 3
        Xp3 = X + 3
        Xm2 = X - 2
        Xp2 = X + 2
        
        If ContByte2(0, Xm3, Y) + ContByte2(0, Xm2, ym2) + _
           ContByte2(0, X, ym3) + ContByte2(0, Xp2, ym2) + _
           ContByte2(0, Xp3, Y) + ContByte2(0, Xp2, yp2) + _
           ContByte2(0, X, yp3) + ContByte2(0, Xm2, yp2) < _
           Enhanced Then
            ContByte(0, X, Y) = 255
            ContByte(1, X, Y) = 255
            ContByte(2, X, Y) = 255
      
        End If
           
        

    Next X
Next Y
End If

'****************************************************************************



End Sub

Sub effEXTENDEDBlurOLD(Optional Sat As Single = 0)
Dim X As Long
Dim Y As Long

Dim iX As Long
Dim iY As Long

Dim X2 As Long

Dim Rt As Integer
Dim Gt As Integer
Dim Bt As Integer
Dim R As Integer
Dim G As Integer
Dim B As Integer

Dim V As Integer

ReDim BlurByte(0 To PB, 0 To pW, 0 To PH) As Byte
ReDim SepaByte(0 To PB, 0 To pW, 0 To PH) As Byte



If Sat <> 0 Then
    
    For Y = 0 To PH
        For X = 0 To pW
            
            R = Sbyte(0, X, Y)
            G = Sbyte(1, X, Y)
            B = Sbyte(2, X, Y)
            
            V = FastAVG(B + G + R)
            'B = B + Sat * (B - V)
            'G = G + Sat * (G - V)
            'R = R + Sat * (R - V)
            
            B = V + (B - V) * Sat
            G = V + (G - V) * Sat
            R = V + (R - V) * Sat
            
            'B = (V + B) / 2 + (B - V) * Sat
            'G = (V + G) / 2 + (G - V) * Sat
            'R = (V + R) / 2 + (R - V) * Sat
            If R > 255 Then R = 255
            If R < 0 Then R = 0
            If G > 255 Then G = 255
            If G < 0 Then G = 0
            If B > 255 Then B = 255
            If B < 0 Then B = 0
            
            Sbyte(0, X, Y) = R
            Sbyte(1, X, Y) = G
            Sbyte(2, X, Y) = B
            
        Next X
    Next Y
    
    
End If

For Y = 1 To PH - 1
    For X = 1 To pW - 1
        Rt = 0
        Gt = 0
        Bt = 0
        For iY = -1 To 1
            Y2 = Y + iY
            For iX = -1 To 1
                X2 = X + iX
                
                Rt = Rt + Sbyte(0, X2, Y2) '* KmulBLUR(iX, iY)
                Gt = Gt + Sbyte(1, X2, Y2) '* KmulBLUR(iX, iY)
                Bt = Bt + Sbyte(2, X2, Y2) '* KmulBLUR(iX, iY)
            Next iX
        Next iY
        
        Rt = Rt \ 9 'KmulBlurD '\ for rt integer
        Gt = Gt \ 9 'KmulBlurD
        Bt = Bt \ 9 'KmulBlurD
        
        '     If chSaturation = Checked Then Saturation Rt, Gt, Bt, SS '* SS
        '     Brightness Rt, Gt, Bt
        If Rt > 255 Then Rt = 255
        If Gt > 255 Then Gt = 255
        If Bt > 255 Then Bt = 255
        
        BlurByte(0, X, Y) = Rt
        BlurByte(1, X, Y) = Gt
        BlurByte(2, X, Y) = Bt
        
    Next X
    DoEvents
Next Y
End Sub

Private Sub Class_Terminate()
'Stop

Erase Sbyte
Erase Sbyte2
Erase ContByte
Erase BlurByte
Erase SepaByte

'Erase PGMByte
'Erase PGM_C_Byte

End Sub


Public Sub EFFQuantizeMy()
Dim C As Long
Dim R As Byte
Dim G As Byte
Dim B As Byte
Dim X As Integer
Dim Y As Integer
Dim P
Dim dR
Dim dG
Dim dB
Dim D
Dim Dmin
Dim iMin
Dim palI

For X = 0 To pW
    For Y = 0 To PH
        Dmin = 99999999999999#
        
        'R = BlurPic(X, y).R
        'G = BlurPic(X, y).G
        'B = BlurPic(X, y).B
        
        R = BlurByte(0, X, Y)
        G = BlurByte(1, X, Y)
        B = BlurByte(2, X, Y)
        
        For P = 0 To MAXIND
            dR = R - PALE(2, P)
            dG = G - PALE(1, P)
            dB = B - PALE(0, P)
            D = FastRoot(FastPower(dR) + FastPower(dG) + FastPower(dB))
            'D = Abs(dR) + Abs(dG) + Abs(dB)
            If D < Dmin Then Dmin = D: iMin = P
        Next P
        
        '        SetPixel ResPIC.hDC, X, y, RGB(PALE(0, iMin), PALE(1, iMin), PALE(2, iMin))
        
        BlurByte(0, X, Y) = PALE(2, iMin)
        BlurByte(1, X, Y) = PALE(1, iMin)
        BlurByte(2, X, Y) = PALE(0, iMin)
        
    Next Y
    
Next X


'For palI = 0 To MAXIND
'    PicPAL.Line (palI * PicPAL.Width / (MAXIND + 1), 0)- _
'            ((palI + 1) * PicPAL.Width / (MAXIND + 1), PicPAL.Height), _
'            RGB(PALE(0, palI), PALE(1, palI), PALE(2, palI)), BF
'
'Next
'PicPAL.Refresh

End Sub

Public Sub EFFSeparateMY(COLIndex As Long)

Dim X As Integer
Dim Y As Integer

For Y = 0 To PH
    For X = 0 To pW
        '            Stop
        
        If PALE(0, COLIndex) = BlurByte(2, X, Y) Then
            If PALE(1, COLIndex) = BlurByte(1, X, Y) Then
                If PALE(2, COLIndex) = BlurByte(0, X, Y) Then
                    SepaByte(0, X, Y) = 0
                    SepaByte(1, X, Y) = 0
                    SepaByte(2, X, Y) = 0
                Else
                    SepaByte(0, X, Y) = 255
                    SepaByte(1, X, Y) = 255
                    SepaByte(2, X, Y) = 255
                End If
            Else
                SepaByte(0, X, Y) = 255
                SepaByte(1, X, Y) = 255
                SepaByte(2, X, Y) = 255
            End If
        Else
            SepaByte(0, X, Y) = 255
            SepaByte(1, X, Y) = 255
            SepaByte(2, X, Y) = 255
        End If
        
    Next X
Next Y
End Sub
Public Sub EFFSeparateFREE(COLIndex As Long)

Dim X As Integer
Dim Y As Integer

For Y = 0 To PH
    For X = 0 To pW
        '            Stop
        
        If PALE(0, COLIndex) = Sbyte(2, X, Y) Then
            If PALE(1, COLIndex) = Sbyte(1, X, Y) Then
                If PALE(2, COLIndex) = Sbyte(0, X, Y) Then
                    SepaByte(0, X, Y) = 0
                    SepaByte(1, X, Y) = 0
                    SepaByte(2, X, Y) = 0
                Else
                    SepaByte(0, X, Y) = 255
                    SepaByte(1, X, Y) = 255
                    SepaByte(2, X, Y) = 255
                End If
            Else
                SepaByte(0, X, Y) = 255
                SepaByte(1, X, Y) = 255
                SepaByte(2, X, Y) = 255
            End If
        Else
            SepaByte(0, X, Y) = 255
            SepaByte(1, X, Y) = 255
            SepaByte(2, X, Y) = 255
        End If
        
    Next X
Next Y
End Sub

Sub EffDrawPGM(fName As String, COL)



Dim s As String
Dim C As Byte

Dim NewR As Integer
Dim NewG As Integer
Dim NewB As Integer
Dim CC As Long

Dim sR As Byte
Dim sG As Byte
Dim sB As Byte

Dim C1 As Single
Dim C2 As Single

Dim InFile() As Byte
Dim Pos As Integer
Dim G As Long

Dim W As Long
Dim H As Long

Dim X As Integer
Dim Y As Integer

Dim sBLEND As Single
Dim s1BLEND As Single

Dim BlendGS As Boolean

Dim GRAY As Integer


sBLEND = Form1.BLEND / 100
s1BLEND = 1 - sBLEND

BlendGS = IIf(Form1.chBW.Value = Checked, True, False)


Open App.path & "\Potrace\" & fName & ".pgm" For Binary Access Read As 1


Readline 10
Readline 10
W = CInt(Readline(Asc(" ")))
H = CInt(Readline(Asc(" ")))
Readline 10
'PIC.Width = W
'PIC.Height = H

ReDim InFile(W * H)
Get #1, , InFile
G = 0
For Y = 0 To H - 1
    For X = 0 To W - 1
        
        G = G + 1
        C = InFile(G)
        
        If C = 0 Then
            
            If sBLEND = 0 Then
                Sbyte(0, X, Y) = PALE(2, COL)
                Sbyte(1, X, Y) = PALE(1, COL)
                Sbyte(2, X, Y) = PALE(0, COL)
            Else
                
                If BlendGS Then
                                
                    GRAY = Sbyte(2, X, Y) * 0.241 + Sbyte(1, X, Y) * 0.691 + Sbyte(0, X, Y) * 0.068
                    GRAY = GRAY * sBLEND
                                
                    NewB = (GRAY + PALE(2, COL) * s1BLEND)
                    NewG = (GRAY + PALE(1, COL) * s1BLEND)
                    NewR = (GRAY + PALE(0, COL) * s1BLEND)
                
                Else
                
                    NewB = (Sbyte(0, X, Y) * sBLEND + PALE(2, COL) * s1BLEND)
                    NewG = (Sbyte(1, X, Y) * sBLEND + PALE(1, COL) * s1BLEND)
                    NewR = (Sbyte(2, X, Y) * sBLEND + PALE(0, COL) * s1BLEND)
                
                End If
                
                
                
                If NewR < 0 Then NewR = 0
                If NewG < 0 Then NewG = 0
                If NewB < 0 Then NewB = 0
            
                If NewR > 255 Then NewR = 255
                If NewG > 255 Then NewG = 255
                If NewB > 255 Then NewB = 255
            
                Sbyte(0, X, Y) = NewB
                Sbyte(1, X, Y) = NewG
                Sbyte(2, X, Y) = NewR
            
            End If
            
        Else
            If C < 255 Then
        
                C1 = C / 255
                C2 = 1 - C1
        
                'CC = GetPixel(PIC.hDC, x, y)
                'LongToRGB CC, sR, sG, sB
        
                sR = Sbyte(0, X, Y)
                sG = Sbyte(1, X, Y)
                sB = Sbyte(2, X, Y)
        
        
                NewR = PALE(2, COL) '
                NewG = PALE(1, COL) '
                NewB = PALE(0, COL) '
        
                NewR = CInt(NewR) * C2 + sR * C1
                NewG = CInt(NewG) * C2 + sG * C1
                NewB = CInt(NewB) * C2 + sB * C1
                If NewR < 0 Then NewR = 0
                If NewG < 0 Then NewG = 0
                If NewB < 0 Then NewB = 0
                '        NewR = ((255 - C) / 255) * NewR
                '        NewG = ((255 - C) / 255) * NewG
                '        NewB = ((255 - C) / 255) * NewB
        
                'SetPixel PIC.hDC, x, y, RGB(NewR, NewG, NewB)
         
                'PGMByte(0, x, y) = NewR
                'PGMByte(1, x, y) = NewG
                'PGMByte(2, x, y) = NewB
         
                 Sbyte(0, X, Y) = NewR
                 Sbyte(1, X, Y) = NewG
                 Sbyte(2, X, Y) = NewB
          
          
            End If
        End If
    Next X
Next Y


'Set bmp = New cDIB

'If bmp.CreateFromFile(App.Path & "\PIC.bmp") <> True Then
'        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
'     'GoTo Error
'End If


Close 1







End Sub
Sub EffDrawPGM_C(fName As String)

Dim s As String
Dim C As Byte
Dim C2 As Byte

Dim CC As Long

Dim NewR As Integer
Dim NewG As Integer
Dim NewB As Integer

Dim sR As Byte
Dim sG As Byte
Dim sB As Byte

Dim InFile() As Byte
Dim Pos As Integer
Dim G As Long

Dim W As Long
Dim H As Long

Dim Y As Integer
Dim X As Integer


Open App.path & "\Potrace\" & fName & ".pgm" For Binary Access Read As 1

Readline 10
Readline 10
W = CInt(Readline(Asc(" ")))
H = CInt(Readline(Asc(" ")))
Readline 10
'PIC.Width = W
'PIC.Height = H

ReDim InFile(W * H)
Get #1, , InFile
G = 0
For Y = 0 To H - 1
    For X = 0 To W - 1
        G = G + 1
        C = InFile(G)
        
        If C < 255 Then
                       
            
            'CC = GetPixel(PIC.hDC, x, y)
            'LongToRGB CC, sR, sG, sB
            sR = Sbyte(0, X, Y)
            sG = Sbyte(1, X, Y)
            sB = Sbyte(2, X, Y)
            
            C2 = 255 - C
            
            NewR = CInt(sR) - C2 '(256 - C)
            NewG = CInt(sG) - C2 '(256 - C)
            NewB = CInt(sB) - C2 '(256 - C)
            
            If NewR < 0 Then NewR = 0
            If NewG < 0 Then NewG = 0
            If NewB < 0 Then NewB = 0
            
            'SetPixel PIC.hDC, x, y, RGB(NewR, NewG, NewB)
            Sbyte(0, X, Y) = NewR
            Sbyte(1, X, Y) = NewG
            Sbyte(2, X, Y) = NewB
            
        End If
        
    Next X
Next Y
Close 1
'PIC.Refresh

End Sub

Private Function Readline(StopCHR As Byte) As String
Dim SS As String
SS = vbNullString
Dim B As Byte

Do
    Get #1, , B
    SS = SS + Chr(B)
    
Loop While B <> StopCHR

Readline = SS
'Stop
End Function



Public Sub EFFsaturation(V)

'#define  Pr  .241
'#define  Pg  .691
'#define  Pb  .068
'
'
'
'//  public-domain function by Darel Rex Finley
'//
'//  The passed-in RGB values can be on any desired scale, such as 0 to
'//  to 1, or 0 to 255.  (But use the same scale for all three!)
'//
'//  The "change" parameter works like this:
'//    0.0 creates a black-and-white image.
'//    0.5 reduces the color saturation by half.
'//    1.0 causes no change.
'//    2.0 doubles the color saturation.
'//  Note:  A "change" value greater than 1.0 may project your RGB values
'//  beyond their normal range, in which case you probably should truncate
'//  them to the desired range before trying to use them in an image.''

'void changeSaturation(double *R, double *G, double *B, double change) {

'  double  P=sqrt(
'  (*R)*(*R)*Pr+
'  (*G)*(*G)*Pg+
'  (*B)*(*B)*Pb ) ;

'  *R=P+((*R)-P)*change;
'  *G=P+((*G)-P)*change;
'  *B=P+((*B)-P)*change; }
'
End Sub

Sub effEXTENDEDBlur(Optional Sat As Single = 0)
Const pRR = 0.241
Const pGG = 0.691
Const pBB = 0.068


Dim X As Long
Dim Y As Long

Dim iX As Long
Dim iY As Long

Dim X2 As Long

Dim Rt As Integer
Dim Gt As Integer
Dim Bt As Integer
Dim R As Long
Dim G As Long
Dim B As Long

Dim V As Integer

Dim P As Single
'Stop



ReDim BlurByte(0 To PB, 0 To pW, 0 To PH) As Byte
ReDim SepaByte(0 To PB, 0 To pW, 0 To PH) As Byte



If Sat <> 0 Then
    
    For Y = 0 To PH
        For X = 0 To pW
            
            R = Sbyte(0, X, Y)
            G = Sbyte(1, X, Y)
            B = Sbyte(2, X, Y)
            
'OLD Satutation
'            V = FastAVG(b + G + r)
'            b = V + (b - V) * Sat
'            G = V + (G - V) * Sat
'            r = V + (r - V) * Sat
            
' New Saturation
  P = Sqr(R * R * pRR + G * G * pGG + B * B * pBB)
  R = P + (R - P) * Sat
  G = P + (G - P) * Sat
  B = P + (B - P) * Sat



'-------------------

            If R > 255 Then R = 255
            If R < 0 Then R = 0
            If G > 255 Then G = 255
            If G < 0 Then G = 0
            If B > 255 Then B = 255
            If B < 0 Then B = 0
            
            Sbyte(0, X, Y) = R
            Sbyte(1, X, Y) = G
            Sbyte(2, X, Y) = B
            
        Next X
    Next Y
    
    
End If

For Y = 1 To PH - 1
    For X = 1 To pW - 1
        Rt = 0
        Gt = 0
        Bt = 0
        For iY = -1 To 1
            Y2 = Y + iY
            For iX = -1 To 1
                X2 = X + iX
                
                Rt = Rt + Sbyte(0, X2, Y2) '* KmulBLUR(iX, iY)
                Gt = Gt + Sbyte(1, X2, Y2) '* KmulBLUR(iX, iY)
                Bt = Bt + Sbyte(2, X2, Y2) '* KmulBLUR(iX, iY)
            Next iX
        Next iY
        
        Rt = Rt \ 9 'KmulBlurD '\ for rt integer
        Gt = Gt \ 9 'KmulBlurD
        Bt = Bt \ 9 'KmulBlurD
        
        '     If chSaturation = Checked Then Saturation Rt, Gt, Bt, SS '* SS
        '     Brightness Rt, Gt, Bt
        If Rt > 255 Then Rt = 255
        If Gt > 255 Then Gt = 255
        If Bt > 255 Then Bt = 255
        
        BlurByte(0, X, Y) = Rt
        BlurByte(1, X, Y) = Gt
        BlurByte(2, X, Y) = Bt
        
    Next X
    DoEvents
Next Y
End Sub

