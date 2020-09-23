VERSION 5.00
Begin VB.Form frmHELP 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "h e l p"
   ClientHeight    =   5490
   ClientLeft      =   1260
   ClientTop       =   2100
   ClientWidth     =   7245
   FillColor       =   &H00404040&
   Icon            =   "frmHELP.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   Begin VB.CheckBox Check1 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "down"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "up"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6600
      Top             =   1800
   End
   Begin VB.Label L 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4995
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5970
   End
End
Attribute VB_Name = "frmHELP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Speed As Single
Dim IncSpeed As Single
Dim PH As Integer
Dim H As Integer

Private Sub Check1_Click()
Timer1.Enabled = IIf(Check1, False, True)

End Sub

Private Sub Command1_Click()
IncSpeed = IncSpeed - 7

End Sub

Private Sub Command2_Click()
IncSpeed = IncSpeed + 7
End Sub







Private Sub Form_Activate()
Timer1.Enabled = True

End Sub

Private Sub Form_Load()
Speed = 0
IncSpeed = 0

L.Top = frmHELP.ScaleHeight

L.Height = 1780 '1640
L = "VIDEO TO CARTOON V" & App.Major & "." & App.Minor & vbCrLf & vbCrLf & vbCrLf
L = L + "HELP" & vbCrLf & vbCrLf

''''
L = L + "------------------------------" & vbCrLf & vbCrLf
L = L + "Version 7 introduce a new Cartoonization Mode: by Iteration Bilateral Filtering." & vbCrLf
L = L + "Cartoonization mode are Two: BILATERAL - NOT BILATERAL" & vbCrLf & vbCrLf
L = L + "BILATERAL MODE  - Apply Bilateral Filtering [Very Slow :( ]. Bilateral Mode Parameters are only these: 'Color Paramters': Brightness and Contrast, plus main 'CONTOUR Params' slider. (To indicate the Amount of Contour)" & vbCrLf & vbCrLf
L = L + "------------------------------" & vbCrLf & vbCrLf & vbCrLf
''''



L = L + "OPEN AVI  - Select the Source Video to Convert to Cartoon. If you want to select another video it is not needed to close it, so click again Open Avi." & vbCrLf & vbCrLf

L = L + "GO TO FRAME   - Go to a specific Frame." & vbCrLf & vbCrLf
L = L + "START  - Select current Frame as the First of the Sequence." & vbCrLf & vbCrLf
L = L + "END  - Select current Frame as the Last of the Sequence." & vbCrLf & vbCrLf
L = L + "OUTPUT WIDTH  - Set the Video Output Width. The height is base on Input aspect Ratio." & vbCrLf & vbCrLf
L = L + "FPS  - Choose the Video Output Frames Per Seconds. (Min 1 , Max 30) " & vbCrLf & vbCrLf




L = L + "QUANTIZE MODE  - There are Two Quantization Mode (Ways to reduce the Colors Number): WUQANT and NNQUANT. Unchecked is WUQUANT, this mode is more based on Colors (Hue). Checked is NNQUANT, this mode is more based on Brightness. Quantization is done by FreeImageLib." & vbCrLf & vbCrLf
L = L + "COLORS PARAMETERS  - " & vbCrLf
L = L + "       Saturation - 0 kinda Sepia - 1 Black and white - 10 No Variation - 20,30,40... 2,3,4 times Saturation." & vbCrLf & vbCrLf
L = L + "CONTOUR PARAMETERS  - " & vbCrLf & vbCrLf
L = L + "COLOR NUMBER Slidebar  - Here you can choose how many colors will be used in the Cartoonized Image. Usually a big number is not needed. Values between 5 and 10 are Suggested." & vbCrLf & vbCrLf
L = L + "PALETTE  - The Colors to use are displaied here. To manually customize palette click a Color." & vbCrLf & vbCrLf
L = L + "GLOBAL PALETTE CheckBox  - Unchecked means that Quantization (and so a new palette) will be done for each frame Cartoonized, so the Palette will be computed for each frame." & vbCrLf
L = L + "Checked means that the Dispalyed Palette will be used do Cartoonize each Frame." & vbCrLf & vbCrLf
L = L + "FIND GLOBAL PALETTE  - Find a Platte that should be good for the selected sequence of frames. Notice, To Use this Palette 'Global Palette' must be checked. (It is Drawn a big image with NxN frames between the Sequence, then it is Quantized). ONE preview frame will be Cartoonized After Global Palette Creation." & vbCrLf & vbCrLf
L = L + "PREVIEW THIS FRAME  - Cartoonize for preview Current Frame. (Bug to solve: if it Dosen't stop click Abort.)" & vbCrLf & vbCrLf
L = L + "CARTOONIZE ALL  - Begin the cartoonization of Sequence. See START and END button for sequence selection." & vbCrLf & vbCrLf
L = L + "ABORT  - Stop Cartoonization Process. (Click once)." & vbCrLf & vbCrLf
L = L + "SAVE AVI...  - After all frames Creation click here to Save all frames as AVI. It will be promped for Avi FileName and Video Compression. You Can click here again to save it again. All the frames will not be deleted until you begin a new Cartoonization or you open a new Input Avi file." & vbCrLf & vbCrLf
L = L + "AUTO SAVE AVI  - After all frames Creation 'Save Avi...' is Run. Have Bug, leave Unchecked." & vbCrLf & vbCrLf
L = L + "PLAYER...  - If you Check AutoPlayAvi it is needed to click here to select the Player." & vbCrLf & vbCrLf
L = L + "AUTO PLAY AVI  - When Avi is finished then it will be played. Click 'Player...' to select your Avi Player." & vbCrLf & vbCrLf
L = L + "EXTRA FRAMES  - Number of Extra Frames between Each Frame. E.G. Ouput FPS=12, ExtraFrame=1 will create an AVI with 24 FPS with Each (12fps) Frame repeated 2 times. (Should be useful for improved compression Quality or Not to be quality killed by youtube.)" & vbCrLf & vbCrLf
L = L + "SAVE/LOAD SETTINGS - Save Load Current Setting; Selected AviFile, First Frame, Last Frame, Output Width, Output FPS, Quantization Mode, Color Parameters, Contour Parameters, Number of Colors and Palette Colors." & vbCrLf & vbCrLf

L = L + vbCrLf + vbCrLf
L = L + "Still Bugs... hope to fix..." & vbCrLf & vbCrLf
L = L + "Thanks to" & vbCrLf & vbCrLf
L = L + "POTRACE.EXE" & vbCrLf
L = L + "http://potrace.sourceforge.net/" & vbCrLf & vbCrLf
L = L + "FreeImage.dll" & vbCrLf
L = L + "http://freeimage.sourceforge.net/download.html" & vbCrLf & vbCrLf
L = L + "http://www.shrinkwrapvb.com/avihelp/avihelp.htm" + vbCrLf + vbCrLf
L = L + Chr$(34) + "SetBitmapBits" + Chr$(34) + " by DreamVB from Planet Source Code" + vbCrLf + vbCrLf + vbCrLf + vbCrLf
L = L + "Roberto Mior  |  reexre@gmail.com" & vbCrLf & vbCrLf


Open App.path & "\HELP.txt" For Output As 33
Print #33, L.Caption
Close 33

End Sub

Private Sub Timer1_Timer()
Speed = -1 - IncSpeed

H = H + Speed

If Abs(H - PH) > 0 Then
L.Top = L.Top + Speed
PH = H
End If

If L.Top + L.Height < 1 Then L.Top = frmHELP.ScaleHeight

IncSpeed = IncSpeed * 0.92



End Sub
