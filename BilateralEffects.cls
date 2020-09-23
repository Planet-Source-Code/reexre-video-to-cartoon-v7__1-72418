VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BilateralEffect"
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
Option Explicit


Private Type tHSP
    H As Single
    S As Single
    P As Single
End Type

Private Type tVector
    x As Single
    Y As Single
    L As Single
    
End Type

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

Private BlurByte() As Byte

Private SepaByte() As Byte

'Private PGMByte() As Byte
'Private PGM_C_Byte() As Byte

Private BILAByte() As Byte
Private ContByte() As Byte
Private ContByte2() As Byte

Private hBmp As Bitmap

Private pW As Integer
Private PH As Integer
Private PB As Integer


Private FastExp() As Single

Private FastDIF() As Single

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, _
        source As Any, ByVal bytes As Long)


Private Sub ZInitFastExp(SigS)
    
    Dim V As Single
    Dim V2 As Single
    
    ReDim FastExp(255)
    
    For V = 0 To 255
        FastExp(V) = Exp(-((V / 255) / (SigS)))
    Next
    
    ReDim FastDIF(0 To 255, 0 To 255)
    For V = 0 To 255
    For V2 = 0 To 255
    FastDIF(V, V2) = FastExp(Abs(V - V2))
    Next
    Next
    
End Sub


Public Sub SetSource(pboxImageHandle As Long)
    'Public Sub GetBits(pBoxPicHand As Long)
    Dim iRet As Long
    'Get the bitmap header
    iRet = GetObject(pboxImageHandle, Len(hBmp), hBmp)
    '   iRet = GetObject(pBoxPicHand, Len(hBmp), hBmp)
    
    'Resize to hold image data
    ReDim Sbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into Sbyte array
    'iRet = GetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))
    iRet = GetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))
    
    
    pW = hBmp.bmWidth - 1
    PH = hBmp.bmHeight - 1
    PB = (hBmp.bmBitsPixel \ 8) - 1
    
    
    
    'ReDim PGMByte(0 To PB, 0 To pW, 0 To PH)
    
    
End Sub



Public Sub zEFF_CONTOUR2(Enhanced As Integer)
    Dim x As Long
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
    
    ReDim ContByte(0 To PB, 0 To pW, 0 To PH)
    ReDim ContByte2(0 To PB, 0 To pW, 0 To PH)
    
    For Y = 1 To PH - 1
        For x = 1 To pW - 1
            ContByte2(0, x, Y) = Sbyte(0, x, Y)
            ContByte2(1, x, Y) = Sbyte(1, x, Y)
            ContByte2(2, x, Y) = Sbyte(2, x, Y)
        Next
    Next
    
    
    If Enhanced > 0 Then ReDim ContByte2(0 To PB, 0 To pW, 0 To PH) 'As Long
    
    
    For Y = 1 To PH - 1
        For x = 1 To pW - 1
            
            
            ' contbyte2(0, x, y) = Sbyte(0, x, y)
            ' contbyte2(1, x, y) = Sbyte(1, x, y)
            ' contbyte2(2, x, y) = Sbyte(2, x, y)
            
            
            vMinR = 255
            vMinG = 255
            vMinB = 255
            
            For iX = -1 To 1
                X2 = x + iX
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
            
            
            R = Sbyte(0, x, Y)
            G = Sbyte(1, x, Y)
            B = Sbyte(2, x, Y)
            
            R = (R - vMinR)
            G = (G - vMinG)
            B = (B - vMinB)
            
            If R < 10 Then R = 0
            If G < 10 Then G = 0
            If B < 10 Then B = 0
            
            
            '*************
            '        r = r * 4 '10 ' 12 '10
            '        G = G * 4 '10 '12
            '        B = B * 4 '10 '12
            '*************
            
            If R > 255 Then R = 255
            If G > 255 Then G = 255
            If B > 255 Then B = 255
            
            ContByte2(0, x, Y) = R '255 - r
            ContByte2(1, x, Y) = G '255 - G
            ContByte2(2, x, Y) = B '255 - B
            
            
        Next x
    Next Y
    
    
    'Blurring contour
    
    For Y = 1 To PH - 1
        For x = 1 To pW - 1
            
            R = 0
            G = 0
            B = 0
            
            
            For iX = -1 To 1
                X2 = x + iX
                For iY = -1 To 1
                    Y2 = Y + iY
                    R = R + ContByte2(0, X2, Y2)
                    G = G + ContByte2(1, X2, Y2)
                    B = B + ContByte2(2, X2, Y2)
                Next iY
            Next iX
            
            'r = FastAVG(r + G + b) \ 9
            R = (R + G + B) \ 27
            
            
            ContByte(0, x, Y) = R
            ContByte(1, x, Y) = R
            ContByte(2, x, Y) = R
            
            
        Next x
    Next Y
    
    
    
    
    
    
    '****************************************************************************
    
    If Enhanced > 0 Then
        For Y = 0 To PH - 1
            For x = 0 To pW - 1
                ContByte2(0, x, Y) = ContByte(0, x, Y)
            Next x
        Next Y
        
        
        For Y = 4 To PH - 5
            
            ym3 = Y - 3
            yp3 = Y + 3
            ym2 = Y - 2
            yp2 = Y + 2
            
            For x = 4 To pW - 5
                
                Xm3 = x - 3
                Xp3 = x + 3
                Xm2 = x - 2
                Xp2 = x + 2
                
                If ContByte2(0, Xm3, Y) + ContByte2(0, Xm2, ym2) + _
                        ContByte2(0, x, ym3) + ContByte2(0, Xp2, ym2) + _
                        ContByte2(0, Xp3, Y) + ContByte2(0, Xp2, yp2) + _
                        ContByte2(0, x, yp3) + ContByte2(0, Xm2, yp2) < _
                        Enhanced Then
                ContByte(0, x, Y) = 255
                ContByte(1, x, Y) = 255
                ContByte(2, x, Y) = 255
                
            End If
            
            
            
        Next x
    Next Y
End If

'****************************************************************************



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






Public Function zNotMin0(V) As Byte
    If V < 0 Then zNotMin0 = 0 Else: zNotMin0 = V
End Function
Public Function zNotMax255(V As Single) As Byte
    If V > 255 Then zNotMax255 = 255 Else: zNotMax255 = CByte(V)
End Function

Public Sub zEFF_Contour_Apply()
Dim x As Long
Dim Y As Long

    For x = 0 + 1 To pW - 1
        For Y = 0 + 1 To PH - 1
            'BILAByte(0, X, Y) = zNotMin0(BILAByte(0, X, Y) \ 1 - 2 * (ContByte(0, X, Y) \ 1))
            'BILAByte(1, X, Y) = zNotMin0(BILAByte(1, X, Y) \ 1 - 2 * (ContByte(0, X, Y) \ 1))
            'BILAByte(2, X, Y) = zNotMin0(BILAByte(2, X, Y) \ 1 - 2 * (ContByte(0, X, Y) \ 1))
            '    Stop
            If ContByte(0, x, Y) > 0 Then
                
                BILAByte(0, x, Y) = zNotMin0(BILAByte(0, x, Y) \ 1 - (ContByte(0, x, Y) \ 1))
                BILAByte(1, x, Y) = zNotMin0(BILAByte(1, x, Y) \ 1 - (ContByte(0, x, Y) \ 1))
                BILAByte(2, x, Y) = zNotMin0(BILAByte(2, x, Y) \ 1 - (ContByte(0, x, Y) \ 1))
                
            End If
            'BILAByte(0, X, Y) = zNotMin0(1 * (ContByte(0, X, Y) \ 1))
            'BILAByte(1, X, Y) = zNotMin0(1 * (ContByte(0, X, Y) \ 1))
            'BILAByte(2, X, Y) = zNotMin0(1 * (ContByte(0, X, Y) \ 1))
            
        Next
    Next
    
End Sub
Public Sub zEFF_BilateralFilter(n As Long, Sigma As Single, Iterations As Long)
    'Author :Roberto Mior
    '     reexre@gmail.com
    '
    'If you use source code or part of it please cite the author
    'You can use this code however you like providing the above credits remain intact
    '
    '
    '
    '
    Dim I As Long
    
    Dim x As Long
    Dim Y As Long
    Dim B As Long
    
    Dim Xp As Long
    Dim Yp As Long
    Dim XmN As Long
    Dim XpN As Long
    Dim YmN As Long
    Dim YpN As Long
    
    
    Dim dR As Single
    Dim dG As Single
    Dim dB As Single
    Dim TR As Long
    Dim TG As Long
    Dim TB As Long
    
    Dim RDiv As Single
    Dim GDiv As Single
    Dim BDiv As Single
    
    
    ZInitFastExp 2 * Sigma * Sigma
    
    
    
    
    ReDim BILAByte(0 To PB, 0 To pW, 0 To PH)
    
    For I = 1 To Iterations
        
        For x = 0 + n To pW - n
            XmN = x - n
            XpN = x + n
            For Y = 0 + n To PH - n
                
                TR = 0
                TG = 0
                TB = 0
                RDiv = 0
                GDiv = 0
                BDiv = 0
                
                YmN = Y - n
                YpN = Y + n
                
                
                For Xp = XmN To XpN
                    For Yp = YmN To YpN
                        
                        'How to Speed up
                        'Everything inside these For Loops ?
                        
                        
                        '***** wich is FASTER?????
                        'dR = abs(Sbyte(2, Xp, Yp) \ 1 - Sbyte(2, X, Y) \ 1)
                        'dG = abs(Sbyte(1, Xp, Yp) \ 1 - Sbyte(1, X, Y) \ 1)
                        'dB = abs(Sbyte(0, Xp, Yp) \ 1 - Sbyte(0, X, Y) \ 1)
                        
                        dR = FastDIF(Sbyte(2, Xp, Yp), Sbyte(2, x, Y))
                        dG = FastDIF(Sbyte(1, Xp, Yp), Sbyte(1, x, Y))
                        dB = FastDIF(Sbyte(0, Xp, Yp), Sbyte(0, x, Y))
                        '***************
                        
                        'Pixels that are very different in intensity from the central pixel are weighted less
                        'dR = FastExp(dR)
                        'dG = FastExp(dG)
                        'dB = FastExp(dB)
                        
                        
                        TR = TR + CSng(Sbyte(2, Xp, Yp)) * dR
                        TG = TG + CSng(Sbyte(1, Xp, Yp)) * dG
                        TB = TB + CSng(Sbyte(0, Xp, Yp)) * dB
                        
                        RDiv = RDiv + dR
                        GDiv = GDiv + dG
                        BDiv = BDiv + dB
                        
                    Next
                Next
                
                
                BILAByte(2, x, Y) = zNotMax255(TR / RDiv)
                BILAByte(1, x, Y) = zNotMax255(TG / GDiv)
                BILAByte(0, x, Y) = zNotMax255(TB / BDiv)
                
                'PIC2.PSet (X, Y), RGB(COut(X, Y).R, COut(X, Y).G, COut(X, Y).B)
                
            Next
            DoEvents
        Next
        
        
        'For B = 0 To PB
        'For X = 0 To pW
        'For Y = 0 To PH
        'Sbyte(B, X, Y) = BILAByte(B, X, Y)
        'Next
        'Next
        'Next
        CopyMemory ByVal VarPtr(Sbyte(0, 0, 0)), ByVal VarPtr(BILAByte(0, 0, 0)), CLng(PB + 1) * CLng(pW + 1) * CLng(PH + 1)
        
    Next
    
End Sub

Public Sub zGet_Effect(pboxImageHandle As Long)
    Dim iRet As Long
    
    iRet = SetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, BILAByte(0, 0, 0))
    
    Erase BILAByte
End Sub

Public Sub zGet_Contour(pboxImageHandle As Long)
    Dim iRet As Long
    iRet = SetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, ContByte(0, 0, 0))
    
    Erase ContByte
End Sub

Public Sub zEFF_Contour(Contour_0_100 As Single)
    Dim x As Long
    Dim Y As Long
    
    Dim ContAmount As Single
    
    'Contour_0_100 = 25
    ContAmount = 0.00004 * Contour_0_100
    
        
    ReDim ContByte(0 To PB, 0 To pW, 0 To PH)
    ReDim ContByte2(0 To PB, 0 To pW, 0 To PH)
    
    Dim HSP() As tHSP
    Dim Vec() As tVector
    
    ReDim HSP(0 To pW, 0 To PH)
    ReDim Vec(0 To pW, 0 To PH)
    
'    frmMAIN.PIC2.Cls
    
    For x = 0 To pW
    For Y = 0 To PH
    With HSP(x, Y)
    RGBtoHSP BILAByte(2, x, Y), BILAByte(1, x, Y), BILAByte(0, x, Y), .H, .S, .P
    End With
    Next
    Next

    
    For x = 1 To pW - 1
    For Y = 1 To PH - 1
    
    With Vec(x, Y)
    .Y = -(-HSP(x - 1, Y - 1).P - 2 * HSP(x - 1, Y).P - HSP(x - 1, Y + 1).P + HSP(x + 1, Y - 1).P + 2 * HSP(x + 1, Y).P + HSP(x + 1, Y + 1).P)
    .x = (-HSP(x - 1, Y - 1).P - 2 * HSP(x, Y - 1).P - HSP(x + 1, Y - 1).P + HSP(x - 1, Y + 1).P + 2 * HSP(x, Y + 1).P + HSP(x + 1, Y + 1).P)
    .L = (.x * .x + .Y * .Y)
    '.L = .L * 0.001
    .L = .L * ContAmount
    
    'frmMAIN.PIC2.PSet (X, Y), RGB(.L, .L, .L)
    ContByte(0, x, Y) = zNotMax255(.L)
    
    DoEvents
    
    End With
    
    Next
    Next
    
        
End Sub



