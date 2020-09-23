Attribute VB_Name = "zzColorDialog"
'''How to show Color Dialog using API
'''Monday, January 5th, 2009

'System & API - How to show Color Dialog using API
Option Explicit

Type CHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'ChooseColor flags:
Public Const CC_ANYCOLOR = &H100
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_RGBINIT = &H1
Public Const CC_SHOWHELP = &H8
Public Const CC_SOLIDCOLOR = &H80

Global CC As CHOOSECOLOR

Dim CustomColors() As Byte

Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" _
        (pChoosecolor As CHOOSECOLOR) As Long

'Use the OR operator for multiple flags
'ex: nFlags = CC_FULLOPEN Or CC_SOLIDCOLOR
Public Function ShowColor(hWndOwner As Long, Optional nFlags As Long) As Long
Dim Custcolor(16) As Long, lReturn As Long, I As Integer
ReDim CustomColors(0 To 16 * 4 - 1) As Byte
For I = LBound(CustomColors) To UBound(CustomColors)
    CustomColors(I) = 0
Next I
CC.lStructSize = Len(CC)
CC.hWndOwner = hWndOwner
CC.hInstance = App.hInstance
CC.lpCustColors = StrConv(CustomColors, vbUnicode)
CC.flags = nFlags
If CHOOSECOLOR(CC) <> 0 Then
    ShowColor = CC.rgbResult
    CustomColors = StrConv(CC.lpCustColors, vbFromUnicode)
Else
    ShowColor = -1
End If
End Function



