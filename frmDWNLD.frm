VERSION 5.00
Begin VB.Form frmDWNLD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VTC - Download DLL and EXE required."
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "frmDWNLD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmDWNLD.frx":000C
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   2250
      Picture         =   "frmDWNLD.frx":0316
      ScaleHeight     =   2475
      ScaleWidth      =   4515
      TabIndex        =   5
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   4
      Left            =   210
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmDWNLD.frx":24AF0
      Top             =   2760
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   210
      Locked          =   -1  'True
      MouseIcon       =   "frmDWNLD.frx":24AF6
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmDWNLD.frx":24E00
      Top             =   6600
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   210
      Locked          =   -1  'True
      MouseIcon       =   "frmDWNLD.frx":24E06
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmDWNLD.frx":25110
      Top             =   6120
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   210
      Locked          =   -1  'True
      MouseIcon       =   "frmDWNLD.frx":25116
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmDWNLD.frx":25420
      Top             =   5160
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   210
      Locked          =   -1  'True
      MouseIcon       =   "frmDWNLD.frx":25426
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmDWNLD.frx":25730
      Top             =   4680
      Width           =   8655
   End
End
Attribute VB_Name = "frmDWNLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author : Creator Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
'--------------------------------------------------------------------------------

Dim S1(4)
Dim SL(4)


Private Sub Form_Load()

S1(0) = "FreeImage.dll       WEBSITE"
SL(0) = "http://freeimage.sourceforge.net/download.html"

S1(1) = "FreeImage.dll       DOWNLOAD"
SL(1) = "http://sourceforge.net/projects/freeimage/files/Binary%20Distribution/3.12.0/FreeImage3120Win32.zip/download"

S1(2) = "Potrace.exe         WEBSITE"
SL(2) = "http://potrace.sourceforge.net/"

S1(3) = "Potrace.exe         DOWNLOAD"
SL(3) = "http://potrace.sourceforge.net/download/potrace-1.8.win32-i386.zip"


For i = 0 To 3
Text1(i).Text = S1(i) & vbTab & vbTab & vbTab & SL(i)
Next

Text1(4).Text = "VIDEO TO CARTOON Application uses FreeImage.dll (FreeImageLib) for Color Quantization and " & vbCrLf & _
"Potrace.exe for 'Vector Graphics'. " & vbCrLf & vbCrLf & _
"FreeImage.dll must be in C:\windows\system\ and should be registered so: Start - Run - Regsvr32 C:\windows\system\freeimage.dll" & vbCrLf & vbCrLf & _
"Potrace.exe must be in the Application \Potrace\ Folder. (no registration needed)" & vbCrLf & vbCrLf & _
"Enjoy!"



End Sub

Private Sub Picture1_Click()
Shell "rundll32.exe url.dll,FileProtocolHandler http://www.youtube.com/watch?v=y5ECMGNq5mA", 3

End Sub

Private Sub Text1_Click(Index As Integer)
If Index = 4 Then Exit Sub

Shell "rundll32.exe url.dll,FileProtocolHandler " & SL(Index), 3

End Sub
