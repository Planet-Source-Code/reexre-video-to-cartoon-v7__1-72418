VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "VIDEO TO CARTOON"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   642
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chBILATcont 
      Caption         =   "Bilateral Contour"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chGlobalMODE 
      BackColor       =   &H000080FF&
      Caption         =   "BILATERAL MODE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   1560
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   8280
      Pattern         =   "*.ini"
      TabIndex        =   49
      ToolTipText     =   "Double Click to Load Settings"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.Slider BLEND 
      Height          =   315
      Left            =   12240
      TabIndex        =   56
      ToolTipText     =   "BLEND - % of Blend with Original Frame"
      Top             =   5595
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      Max             =   30
      TickFrequency   =   5
   End
   Begin VB.ComboBox cmbEXTRA 
      Height          =   315
      Left            =   11040
      TabIndex        =   55
      Text            =   "Combo1"
      ToolTipText     =   "EXTRA FRAMES"
      Top             =   8760
      Width           =   495
   End
   Begin MSComctlLib.Slider Ccontra 
      Height          =   300
      Left            =   12960
      TabIndex        =   44
      ToolTipText     =   "Contour Contrast"
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   393216
      SmallChange     =   5
      Max             =   6000
      TickFrequency   =   1000
   End
   Begin MSComctlLib.Slider Cbright 
      Height          =   300
      Left            =   12960
      TabIndex        =   45
      ToolTipText     =   "Contour Brightness"
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   393216
      SmallChange     =   5
      Min             =   -50
      Max             =   100
      TickFrequency   =   25
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14040
      TabIndex        =   54
      ToolTipText     =   "HELP"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Status 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Height          =   1095
      Left            =   10080
      MultiLine       =   -1  'True
      TabIndex        =   53
      Text            =   "frmMAIN.frx":0000
      Top             =   5940
      Width           =   1335
   End
   Begin VB.CommandButton cmdPLAY 
      Caption         =   "Player ..."
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
      Left            =   11640
      TabIndex        =   52
      ToolTipText     =   "Select Your AVI Player (.exe)"
      Top             =   8760
      Width           =   855
   End
   Begin VB.CheckBox chPLAY 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Auto PLAY Avi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12600
      TabIndex        =   51
      ToolTipText     =   "AutoPlay AVI when It's Created"
      Top             =   9000
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LOAD Settings"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   14040
      TabIndex        =   50
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveParam 
      Caption         =   "SAVE Settings"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   14040
      TabIndex        =   48
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox txtCurFrame 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "0"
      ToolTipText     =   "Current FRAME"
      Top             =   1845
      Width           =   2295
   End
   Begin MSComctlLib.Slider sWWWcol 
      Height          =   975
      Left            =   11520
      TabIndex        =   38
      ToolTipText     =   "Colors Detail"
      Top             =   3255
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1720
      _Version        =   393216
      Orientation     =   1
      SmallChange     =   5
      Min             =   200
      Max             =   520
      SelStart        =   240
      TickFrequency   =   40
      Value           =   240
   End
   Begin MSComctlLib.Slider Contra 
      Height          =   300
      Left            =   12960
      TabIndex        =   13
      ToolTipText     =   "Colors Contrast"
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   393216
      SmallChange     =   5
      Max             =   200
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider Bright 
      Height          =   300
      Left            =   12960
      TabIndex        =   11
      ToolTipText     =   "Colors Brightness"
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   393216
      SmallChange     =   5
      Max             =   200
      SelStart        =   15
      TickFrequency   =   25
   End
   Begin MSComctlLib.Slider Satura 
      Height          =   300
      Left            =   12960
      TabIndex        =   12
      ToolTipText     =   "Colors Saturation "
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   393216
      SmallChange     =   5
      Max             =   60
      SelStart        =   20
      TickFrequency   =   10
      Value           =   20
   End
   Begin MSComctlLib.Slider sWWWcont 
      Height          =   975
      Left            =   11520
      TabIndex        =   39
      ToolTipText     =   "Contour Detail"
      Top             =   4320
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1720
      _Version        =   393216
      Orientation     =   1
      SmallChange     =   5
      Min             =   380
      Max             =   820
      SelStart        =   380
      TickFrequency   =   41
      Value           =   380
   End
   Begin VB.CommandButton FrameF 
      Caption         =   ">>>"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12840
      TabIndex        =   37
      ToolTipText     =   "1 Frame Forward"
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton frameB 
      Caption         =   "<<<"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12480
      TabIndex        =   36
      ToolTipText     =   "1 Frame Back"
      Top             =   840
      Width           =   255
   End
   Begin VB.Timer Timer_Move 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   13320
      Top             =   0
   End
   Begin MSComctlLib.Slider sMOVE 
      Height          =   375
      Left            =   11640
      TabIndex        =   32
      ToolTipText     =   "<< >>  Frame Mover"
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Max             =   1000
      SelStart        =   500
      Value           =   500
   End
   Begin VB.TextBox txtFrameTime 
      Height          =   285
      Left            =   9360
      TabIndex        =   30
      Text            =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdBuildAVI 
      Caption         =   "save AVI..."
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
      Left            =   10080
      TabIndex        =   29
      ToolTipText     =   "Save AVI file from pervious Cartoonized frames. (All frames in \Frames folder)"
      Top             =   8760
      Width           =   975
   End
   Begin VB.CheckBox chSaveAvi 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Auto SAVE Avi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12600
      TabIndex        =   28
      Top             =   8760
      Width           =   1695
   End
   Begin VB.Timer TIMERDO_Sepa_Continue 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6840
      Top             =   3240
   End
   Begin VB.Timer TIMERDo_Cont_Continue 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   6840
      Top             =   3840
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "ABORT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14040
      TabIndex        =   27
      ToolTipText     =   "Abort Sequence Cartoonization."
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFPS 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13440
      TabIndex        =   26
      Text            =   "12"
      ToolTipText     =   "Output FPS"
      Top             =   2640
      Width           =   495
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   210
      Left            =   120
      TabIndex        =   25
      Top             =   9360
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   9,99999e5
   End
   Begin VB.CommandButton cmdDoSequence 
      Caption         =   "C A R T O O N I Z E   A L L  !"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      TabIndex        =   24
      ToolTipText     =   "Cartoonize All Frames !"
      Top             =   7920
      Width           =   3855
   End
   Begin VB.CommandButton cmdSetEnd 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   13080
      TabIndex        =   20
      ToolTipText     =   "Set This Frame as the LAST"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdSETStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11640
      TabIndex        =   19
      ToolTipText     =   "Set This Frame as the FIRST"
      Top             =   840
      Width           =   855
   End
   Begin VB.CheckBox ChQM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quantize Mode"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12000
      TabIndex        =   7
      ToolTipText     =   "Quantization Mode: Checked=NNQUANT, Unchecked=WUQUANT.  BEST=UNCHECKED"
      Top             =   2895
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "go to frame"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O P E N          A V I "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox OutSize 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13080
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton RNDpale 
      Caption         =   "RND Palette"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   10
      ToolTipText     =   "Randomize Palette"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chMYPale 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GLOBAL Palette"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11520
      TabIndex        =   18
      ToolTipText     =   "If Not Checked Palette is computed for every frame. its BAD! (Not Checked is good for Preview)"
      Top             =   6690
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdDoFRAME 
      Caption         =   "Preview this Frame"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   17
      ToolTipText     =   "Cartoonize This Frame .. Useful for setting up Parameters."
      Top             =   7080
      Width           =   2415
   End
   Begin MSComctlLib.Slider MAXindBAR 
      Height          =   315
      Left            =   11520
      TabIndex        =   9
      ToolTipText     =   "Number Of Colors to USE. some values ..7,9,11,14,17"
      Top             =   5955
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Min             =   1
      Max             =   23
      SelStart        =   6
      Value           =   6
   End
   Begin VB.PictureBox PicPAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11520
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   8
      ToolTipText     =   "Palette. Click To Customize"
      Top             =   6300
      Width           =   2415
   End
   Begin VB.CommandButton cmdDoGlobalPAL 
      Caption         =   "Find GLOBAL Palette"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      TabIndex        =   15
      ToolTipText     =   "Find Global Palette for ALL Frames to Cartoonize."
      Top             =   7080
      Width           =   1335
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.PictureBox PicEffC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3720
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   16
      Top             =   6360
      Width           =   1095
   End
   Begin VB.PictureBox PICeff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   3
      Top             =   6360
      Width           =   1095
   End
   Begin VB.PictureBox PICtmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3000
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   12480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Slider sVpos 
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   767
      _Version        =   393216
      SmallChange     =   24
   End
   Begin VB.PictureBox SEPA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   2880
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   14
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chBW 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gray Scaled"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13350
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "GrayScaled Blend"
      Top             =   5580
      Value           =   1  'Checked
      Width           =   615
   End
   Begin MSComctlLib.Slider cEnhanced 
      Height          =   300
      Left            =   12960
      TabIndex        =   59
      ToolTipText     =   "Contour Enhance. Suggested = 800-1000"
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   393216
      LargeChange     =   100
      SmallChange     =   5
      Max             =   1000
      TickFrequency   =   300
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enhanced"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11730
      TabIndex        =   60
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "BLEND "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   57
      Top             =   5610
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   1230
      Left            =   10680
      Top             =   4275
      Width           =   3300
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1095
      Left            =   10680
      Top             =   3195
      Width           =   3300
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "COLOR params"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   41
      Top             =   3480
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   888
      X2              =   888
      Y1              =   344
      Y2              =   288
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Brightness"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11730
      TabIndex        =   47
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contrast"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11730
      TabIndex        =   46
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Output Width:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   43
      Top             =   2310
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Saturation"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11730
      TabIndex        =   35
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contrast"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11730
      TabIndex        =   34
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Brightness"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11730
      TabIndex        =   33
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "L5"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FPS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   23
      Top             =   2670
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "1st Frame"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   21
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Last frme"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13080
      TabIndex        =   22
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTOUR params"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   40
      Top             =   4560
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
Public Eff As New myEffects
Public FX As New BilateralEffect

Public MaxOutW As Integer

Public ScalaX As Single
Public ScalaY As Single

Public curFRAME As Long

Public AViLoaded As Boolean

Public x As Long
Public Y As Long

Public ToWaitPotrace As Boolean

Private FIdib As Long
Private FIdibQ As Long

Public EndF As Long
Public StartF As Long
Public FloatFRAME As Single

Public MyStep As Single

Public IndexC As Long

Public WWWcols As Single
Public WWWcont As Single

Public OneFrameTime As Boolean
Public FrameTime As Single

Public INI As New clsINI

Public ParamOPENAVI As String

Public AVIPLAYER As String
Public OutputAVIName As String

Public TEMPO As Single

Public BckColor As Long


Private Sub chGlobalMODE_Click()
chBILATcont.Value = Unchecked
End Sub

Private Sub ChPLAY_Click()
If chPLAY And (AVIPLAYER = "") Then MsgBox "Select Avi Player": cmdPLAY_Click

End Sub

Private Sub cmdAbort_Click()
FloatFRAME = EndF

End Sub

Private Sub cmdBuildAVI_Click()
BUILD_AVI


If chPLAY.Value = Checked Then
    If (AVIPLAYER <> "") Then
        If OutputAVIName <> "" Then
            Shell AVIPLAYER & " " & Chr$(34) & OutputAVIName, vbNormalFocus
        End If
    Else
        MsgBox "No Avi Player Selected!", vbCritical
    End If
    
End If

End Sub

Private Sub cmdDoFRAME_Click()

cmdAbort.Visible = True


OneFrameTime = True
FrameTime = Timer
If chGlobalMODE.Value = Unchecked Then

Me.Caption = "Blurring..."
Do_effblur
Me.Caption = "Separating..."
Do_SepaAndPut

Else

    Set PIC.Picture = FreeImage_AdjustBrightnessIOP(PIC.Image, Bright) '-37' - 33
    Set PIC.Picture = FreeImage_AdjustContrastIOP(PIC.Image, Contra) '-37' - 33
  
    'Call SetStretchBltMode(PICeff.hDC, STRETCHMODE)
    'Call StretchBlt(PICeff.hDC, 0, 0, PICeff.Width - 1, PICeff.Height - 1, _
    '           PIC.hDC, 0, 0, PIC.Width - 1, PIC.Height - 1, vbSrcCopy)
    'PICeff.Refresh
    FX.SetSource PIC.Image.Handle
    'FX.SetSource PICeff.Image.Handle
    
    FX.zEFF_BilateralFilter 2, 0.11, 10 '2, 0.105, 8 '3, 0.105, 9
    
    FX.zEFF_Contour 200 * (sWWWcont.Value - sWWWcont.Min) / (sWWWcont.Max - sWWWcont.Min)
        
    FX.zEFF_Contour_Apply
    
    FX.zGet_Effect PIC.Image.Handle
    'FX.zGet_Effect PICeff.Image.Handle
    'Call SetStretchBltMode(PIC.hDC, STRETCHMODE)
    'Call StretchBlt(PIC.hDC, 0, 0, PIC.Width - 1, PIC.Height - 1, _
               PICeff.hDC, 0, 0, PICeff.Width, PICeff.Height, vbSrcCopy)
    'PICeff.Refresh
    PIC.Refresh
    
    
    
    'If chBILATcont.Value = Checked Then
    '
    '    Me.Caption = "Countour... "
    '    Do_Contour
    'Else
    TIMERDo_Cont_Continue.Enabled = True
    'End If
    
End If

End Sub

Private Sub cmdDoGlobalPAL_Click()
Dim F As Long

Dim X2 As Integer
Dim Y2 As Integer
Dim WW As Integer
Dim HH As Integer

Dim DIV As Integer

DIV = InputBox("Late N x N", , 4)


PICeff.Width = 1200
PICeff.Height = PIC.Height / PIC.Width * PICeff.Width

If StartF = EndF Then Exit Sub

WW = PICeff.Width / DIV
HH = PICeff.Height / DIV

F = StartF
For x = 1 To DIV
    X2 = (x - 2) * WW
    
    For Y = 1 To DIV
        F = F + (EndF - StartF) / (DIV * DIV)
        PutFrameToPIC F
        
        Y2 = (Y - 2) * HH
        
        Call SetStretchBltMode(PICeff.hDC, STRETCHMODE)
        Call StretchBlt(PICeff.hDC, X2 + WW, Y2 + HH, WW, HH, _
                PIC.hDC, 0, 0, PIC.Width - 1, PIC.Height - 1, vbSrcCopy)
        PICeff.Refresh
        DoEvents
        
    Next
    
Next
'-------------------------------------------------
Set PICeff.Picture = FreeImage_AdjustBrightnessIOP(PICeff.Image, Bright) '-37' - 33
Set PICeff.Picture = FreeImage_AdjustContrastIOP(PICeff.Image, Contra) '-37' - 33

PICeff.Refresh


Eff.SetSource PICeff

'Eff.effEXTENDEDBlurOLD Satura / 10
Eff.effEXTENDEDBlur Satura / 10

Eff.PutEffToPic PICeff, eEXblur
PICeff.Refresh

'-------------------------------------------------
EffQuantizeFreeImage


'-------------------------------------------------



PICeff.Width = WWWcols
PICeff.Height = PIC.Height / PIC.Width * WWWcols '- 1


cmdDoFRAME_Click


End Sub

Private Sub cmdDoSequence_Click()

If Dir(App.path & "\Potrace\*.bmp") <> "" Then Kill App.path & "\Potrace\*.bmp"
If Dir(App.path & "\Potrace\*.pgm") <> "" Then Kill App.path & "\Potrace\*.pgm"



TEMPO = Timer



PB.Min = 0
PB.Value = 0
PB.Max = EndF '+ AVIin.FPS
PB.Value = EndF
PB.Min = StartF



If Dir(App.path & "\frames\*.bmp") <> "" Then Kill App.path & "\frames\*.bmp"

MyStep = AVIin.FPS / Val(txtFPS)

'For FloatFRAME = StartF To EndF Step MyStep
FloatFRAME = StartF

PB.Value = FloatFRAME
curFRAME = CLng(Int(FloatFRAME))

Status = "Frame " & curFRAME & "   " & curFRAME - StartF & "/" & EndF - StartF '& vbCrLf & _
'   "Elapsed " & Format((Timer - TEMPO) / 86400, "HH:MM:SS") & "  Remain: " & Format((((Timer - TEMPO) / 86400) / (FloatFRAME - StartF)) * (EndF - FloatFRAME ), "HH:MM:SS")


DoEvents

PutFrameToPIC curFRAME



cmdDoFRAME_Click


'Next


'Kill App.path & "\frames\0*.*"

End Sub

Private Sub cmdPLAY_Click()
With CMD
    .FileName = ""
    .InitDir = "c:\"
    .Filter = "AVI Player|*.EXE" ';*.mpg"
    .DialogTitle = "Select AVI PLAYER"
End With
CMD.Action = 1

If CMD.FileName <> "" Then
    AVIPLAYER = CMD.FileName
    Open App.path & "\Player.txt" For Output As 22
    Print #22, AVIPLAYER
    Close 22
End If

End Sub

Private Sub cmdSaveParam_Click()
SaveSetting

End Sub

Private Sub cmdSeparateAndPut_Click()
'Do_SepaAndPut



End Sub



Private Sub cmdSetEnd_Click()
EndF = curFRAME
Label2 = EndF


Label5.Caption = (Val(Replace(txtFrameTime, ",", ".")) * (EndF - StartF) / (AVIin.FPS / txtFPS)) / 60



End Sub

Private Sub cmdSETStart_Click()
StartF = curFRAME
FloatFRAME = StartF

Label1 = StartF

Label5.Caption = (Val(Replace(txtFrameTime, ",", ".")) * (EndF - StartF) / (AVIin.FPS / txtFPS)) / 60



End Sub



Private Sub Command3_Click()
File1.Refresh
File1.Visible = True

End Sub

Private Sub Command4_Click()
Load frmHELP
frmHELP.Show


End Sub

Private Sub File1_DblClick()
LoadSetting (File1.FileName)
File1.Visible = False
End Sub

Private Sub Form_Activate()
frmHELP.Timer1.Enabled = False

End Sub

Private Sub Form_Resize()

If Me.WindowState = 1 Then Exit Sub


PB.Width = Me.ScaleWidth - 16

PB.Left = Me.ScaleWidth / 2 - PB.Width / 2

PB.Top = Me.ScaleHeight - 20

sVpos.Width = PB.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmHELP


If pAVIStream <> 0 Then AVIStreamClose (pAVIStream)
If pAVIFile <> 0 Then AVIFileClose (pAVIFile)
AVIFileExit
End Sub

Private Sub frameB_Click()
curFRAME = curFRAME - 1


sVpos.Value = curFRAME / 10

PutFrameToPIC curFRAME
FloatFRAME = curFRAME

txtCurFrame = curFRAME
End Sub

Private Sub FrameF_Click()
curFRAME = curFRAME + 1


sVpos.Value = curFRAME / 10

PutFrameToPIC curFRAME
FloatFRAME = curFRAME

txtCurFrame = curFRAME
End Sub



Private Sub MAXindBAR_Click()
MAXIND = CInt(MAXindBAR.Value)
ReDim PALE(0 To 2, 0 To MAXIND + 1) As Integer
ReDim AvgTbl(0 To 2, 0 To MAXIND) ' As Long
ReDim AvgPREC(0 To 2, 0 To MAXIND) ' As Long
ReDim CntTbl(0 To MAXIND) 'As Long
RNDpale_Click

SetupSomething
End Sub

Private Sub MAXindBAR_Scroll()
MAXIND = CInt(MAXindBAR.Value)
ReDim PALE(0 To 2, 0 To MAXIND + 1) As Integer
ReDim AvgTbl(0 To 2, 0 To MAXIND) ' As Long
ReDim AvgPREC(0 To 2, 0 To MAXIND) ' As Long
ReDim CntTbl(0 To MAXIND) 'As Long
RNDpale_Click

SetupSomething

End Sub

Sub InitKmultBlur()
Debug.Print "____BLUR_________"
KmulBlurD = 0
For x = -3 To 3
    For Y = -3 To 3
        'KmulBLUR(x, y) = Round(4.6 - Sqr(x * x + y * y)) 'xy -3 to 3
        
        'KmulBLUR(x, y) = Cos((Abs(x) + Abs(y)) / 6 * 1.57) * 1  '4
        KmulBLUR(x, Y) = Cos((Sqr(x * x + Y * Y)) / Sqr(3 * 3 + 3 * 3) * 1.57) * 1
        
        
        If KmulBLUR(x, Y) < 0 Then KmulBLUR(x, Y) = 0
        
        Me.Line (PIC.Width + PIC.Left + 20 + x * 4, Y * 4 + PIC.Top)- _
                (PIC.Width + PIC.Left + 20 + (x + 1) * 4, (Y + 1) * 4 + PIC.Top), RGB(255 * KmulBLUR(x, Y), 255 * KmulBLUR(x, Y), 255 * KmulBLUR(x, Y)), BF
        
        
        'KmulBLUR(x, y) = Round(4 - Sqr(x * x + y * y))   'xy -2 to 2
        'KmulBLUR(x, y) = Round(1 - Sqr(x * x + y * y)) 'xy -1 to 1
        
        Debug.Print KmulBLUR(x, Y)
        KmulBlurD = KmulBlurD + KmulBLUR(x, Y)
    Next Y
    Debug.Print
Next x

End Sub


Private Sub cmdEFFcontou_Click()

End Sub

Private Sub Command1_Click()
AViLoaded = False



Call AVIFileInit '// opens AVIFile library

If pGetFrameObj <> 0 Then
    Call AVIStreamGetFrameClose(pGetFrameObj) '//deallocates the GetFrame resources and interface
End If
If pAVIStream <> 0 Then
    Call AVIStreamRelease(pAVIStream) '//closes video stream
End If
If pAVIFile <> 0 Then
    Call AVIFileRelease(pAVIFile) '// closes the file
End If

If (res <> AVIERR_OK) Then 'if there was an error then show feedback to user
    MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.Title
End If 'Stop



If ParamOPENAVI = "" Then
    With CMD
        .Filter = "AVI Files|*.avi;*.mpg"
        .DialogTitle = "Open AVI File"
    End With
    CMD.Action = 1
Else
    CMD.FileName = ParamOPENAVI
End If




'res = ofd.VBGetOpenFileNamePreview(szFile)
'  If res = False Then GoTo ErrorOut
szFile = CMD.FileName
'  Stop




'Open the AVI File and get a file interface pointer (PAVIFILE)
res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
'res = AVIFileOpen(pAVIFile, szFile, OF_READ, 0&)

If res <> AVIERR_OK Then GoTo ErrorOUT
' Stop

'Get the first available video stream (PAVISTREAM)
res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
If res <> AVIERR_OK Then GoTo ErrorOUT

'get the starting position of the stream (some streams may not start simultaneously)
firstFrame = AVIStreamStart(pAVIStream)
If firstFrame = -1 Then GoTo ErrorOUT 'this function returns -1 on error

'get the length of video stream in frames
numFrames = AVIStreamLength(pAVIStream)
If numFrames = -1 Then GoTo ErrorOUT ' this function returns -1 on error

'get file info struct (UDT)
res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
If res <> AVIERR_OK Then GoTo ErrorOUT

'print file info to Debug Window
Call DebugPrintAVIFileInfo(fileInfo)

'get stream info struct (UDT)
res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
If res <> AVIERR_OK Then GoTo ErrorOUT



'--------------------------hhhhh


'print stream info to Debug Window
Call DebugPrintAVIStreamInfo(streamInfo)


'set bih attributes which we want GetFrame functions to return
With bih
    .biBitCount = 24
    .biClrImportant = 0
    .biClrUsed = 0
    .biCompression = BI_RGB
    .biHeight = streamInfo.rcFrame.Bottom - streamInfo.rcFrame.Top
    .biPlanes = 1
    .biSize = 40
    .biWidth = streamInfo.rcFrame.Right - streamInfo.rcFrame.Left
    .biXPelsPerMeter = 0
    .biYPelsPerMeter = 0
    .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight 'calculate total size of RGBQUAD scanlines (DWORD aligned)
End With


'init AVISTreamGetFrame* functions and create GETFRAME object
'pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, ByVal AVIGETFRAMEF_BESTDISPLAYFMT) 'tell AVIStream API what format we expect and input stream
pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, bih) 'force function to return 24bit DIBS
If pGetFrameObj = 0 Then
    MsgBox "No suitable decompressor found for this video stream!", vbInformation, App.Title
    GoTo ErrorOUT
End If


AVIin.W_in = fileInfo.dwWidth
AVIin.H_in = fileInfo.dwHeight

'ScalaX = OutSize / AVIin.W_in '720
'ScalaY = ScalaX / 1 ' (Stretch.Value / 10)
'AVIin.W_out = Round(AVIin.W_in * ScalaX) 'SCALA 1.125
'AVIin.H_out = Round(AVIin.H_in * ScalaY) 'SCALA
'PIC.Width = AVIin.W_out
'PIC.Height = AVIin.H_out
OutSize_Click

AVIin.MaxFrames = streamInfo.dwLength
AVIin.FPS = streamInfo.dwRate / streamInfo.dwScale
AVIin.FileName = streamInfo.szName
sVpos.Max = AVIin.MaxFrames / 10 - 1
sVpos.TickFrequency = AVIin.FPS * 6 '30


GoTo NOERROR


ErrorOUT:
MsgBox "AVI ERROR !!!"

ParamOPENAVI = ""
Exit Sub
NOERROR:


ParamOPENAVI = ""

AViLoaded = True

curFRAME = 1
PutFrameToPIC curFRAME

Label3.Caption = "Input FPS = " & AVIin.FPS

PicEffC.Cls
PicEffC.Width = WWWcont
PicEffC.Height = PIC.Height / PIC.Width * WWWcont - 1
PicEffC.Refresh


End Sub


Sub PutFrameToPIC(Frame As Long)


Dim dib As cDIB

'create a DIB class to load the frames into
Set dib = New cDIB


pDIB = AVIStreamGetFrame(pGetFrameObj, Frame) 'returns "packed DIB"

If dib.CreateFromPackedDIBPointer(pDIB) Then
    
    
    Call dib.WriteToFile(OutFileName(Frame))
    
    
    PICtmp = LoadPicture(OutFileName(Frame))
    
    
    'PIC.Width = AVIin.W_out
    'PIC.Height = AVIin.H_out
    
    Call SetStretchBltMode(PIC.hDC, STRETCHMODE)
    Call StretchBlt(PIC.hDC, 0, 0, PIC.Width, PIC.Height, _
            PICtmp.hDC, 0, 0, PICtmp.Width - 1, PICtmp.Height - 1, vbSrcCopy)
    
    
    
    PIC.Refresh
    
    '    SavePicture PIC.Image, OUTFileName(Frame)
    '''   'SaveJPG PIC.Image, OUTFileName(frame), 98
    
    
Else
    
End If
Set dib = Nothing

'CurrentFRAME = frame

End Sub

Function OutFileName(F As Long) As String

OutFileName = App.path & "\Frames\" & Format(F, "00000000") & ".BMP"

End Function

Private Sub Command2_Click()
Dim Frame As Long

Frame = InputBox("go to frame  N")
sVpos = Frame / 10
curFRAME = Frame
txtCurFrame = curFRAME
PutFrameToPIC curFRAME

End Sub

Private Sub cmdEFFcontour_Click()
'Do_Contour
End Sub

Private Sub cmdBLURandQUNAT_Click()
'Do_effblur
End Sub

Private Sub Form_Load()
Dim x As Single
Dim Y As Single
Dim I As Single






If Dir(App.path & "\Data.dat") <> "" Then
    Name App.path & "\Data.dat" As App.path & "\DATA.ZIP"
    MsgBox _
        "This is the first time you Run this Process. " & vbCrLf & vbCrLf & _
        "Since  -FreeImage.dll-  and  -Potrace.exe-  are Required, you can Google and " & _
        "Download them or unzip 'DATA.ZIP' . " & vbCrLf & vbCrLf & _
        "To Get Download Links open project DownloadHelper.vbp ..." & vbCrLf & vbCrLf & vbCrLf & _
        "Notice that FreeImage.dll must be in System Folder!  (C:\windows\system\)" & vbCrLf & _
        "And Potrace.exe in Application folder or Application \POTRACE\ folder.", vbInformation: End
 
End If



If Dir(App.path & "\INIs", vbDirectory) = "" Then MkDir App.path & "\INIs"

If Dir(App.path & "\VIDEO", vbDirectory) = "" Then MkDir App.path & "\VIDEO"

If Dir(App.path & "\Potrace", vbDirectory) = "" Then MkDir App.path & "\Potrace"
If Dir(App.path & "\FRAMES", vbDirectory) = "" Then MkDir App.path & "\FRAMES"

If Dir(App.path & "\Potrace.exe") <> "" Then FileCopy App.path & "\Potrace.exe", App.path & "\Potrace\Potrace.exe": 'Kill App.path & "\Potrace.exe"

If Dir(App.path & "\Player.txt") <> "" Then
    Open App.path & "\Player.txt" For Input As 22
    Input #22, AVIPLAYER
    Close 22
End If


File1.path = App.path & "\inis"


BckColor = &H9CB09C
Me.BackColor = BckColor
ChQM.BackColor = BckColor
chMYPale.BackColor = BckColor
chSaveAvi.BackColor = BckColor
chPLAY.BackColor = BckColor
chBW.BackColor = BckColor

'chGlobalMODE.BackColor = BckColor

chBILATcont.BackColor = BckColor

Me.Caption = Me.Caption & "   " & App.Major & "." & App.Minor


ProcessPrioritySet , , ppidle 'ppbelownormal ' So While is Computing You Can to Other


WWWcols = sWWWcol '360
WWWcont = sWWWcont '600

sWWWcont.TickFrequency = (sWWWcont.Max - sWWWcont.Min) / ((sWWWcol.Max - sWWWcol.Min) / sWWWcol.TickFrequency)
Ccontra.TickFrequency = (Ccontra.Max - Ccontra.Min) / ((Cbright.Max - Cbright.Min) / Cbright.TickFrequency)

cEnhanced.TickFrequency = (cEnhanced.Max - cEnhanced.Min) / ((Cbright.Max - Cbright.Min) / Cbright.TickFrequency)


AVIin.W_in = 720
AVIin.H_in = 576

'x = 720 * 1.25
'y = 576 * 1.25
'For I = 3 To 1 Step -0.125 '-0.0625
'    'If i = 1 Then OutSize.AddItem "640x480"
'    OutSize.AddItem Round(x / I) ' & "x" & Round(Y / I)
'Next I
'
'OutSize.ListIndex = 8 '1
'OutSize_Click
OutSize.AddItem "320"
OutSize.AddItem "360"
OutSize.AddItem "400"
OutSize.AddItem "420"
OutSize.AddItem "480"
OutSize.AddItem "520"
OutSize.AddItem "600"
OutSize.AddItem "640"
OutSize.AddItem "704"
OutSize.AddItem "720"
OutSize.AddItem "800"
OutSize.AddItem "840"
OutSize.AddItem "900"
OutSize.AddItem "1024"
OutSize.AddItem "1280"

OutSize.ListIndex = 4
OutSize_Click

cmbEXTRA.Clear
cmbEXTRA.AddItem "0"
cmbEXTRA.AddItem "1"
cmbEXTRA.AddItem "2"
cmbEXTRA.AddItem "3"
cmbEXTRA.AddItem "4"
cmbEXTRA.AddItem "5"
cmbEXTRA.ListIndex = 0






InitFastAVG
InitKmultBlur
InitFastRoot
InitFastPower




MAXindBAR_Click
End Sub



Private Sub OutSize_Click()

ScalaX = OutSize / AVIin.W_in '720
ScalaY = ScalaX / 1 ' (Stretch.Value / 10)
AVIin.W_out = Round(AVIin.W_in * ScalaX)
AVIin.H_out = Round(AVIin.H_in * ScalaY) '
If AVIin.H_out \ 2 <> AVIin.H_out / 2 Then AVIin.H_out = AVIin.H_out + 1

PIC.Cls
PIC.Width = AVIin.W_out
PIC.Height = AVIin.H_out
PIC.Refresh

If AViLoaded Then PutFrameToPIC curFRAME

SetupSomething

End Sub




Private Sub PicPAL_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim C As Long
Dim P As Integer
Dim R As Byte
Dim G As Byte
Dim B As Byte

C = ShowColor(Me.hWnd, CC_FULLOPEN Or CC_RGBINIT)
P = Round((x / PicPAL.Width) * (MAXIND + 1) + 0.5) - 1

If C <> -1 Then
    LongToRGB C, R, G, B
    
    PALE(0, P) = R
    PALE(1, P) = G
    PALE(2, P) = B
    'PAL(P).rgbRed = R
    'PAL(P).rgbGreen = G
    'PAL(P).rgbBlue = B
    
    drawPALE
End If
End Sub

Private Sub RNDpale_Click()
Dim C As Long
Dim I
'Stop'

For C = 0 To 2
    For I = 0 To MAXIND
        PALE(C, I) = Int(Rnd * 255)
        
    Next
Next
drawPALE
End Sub
Sub drawPALE()
Dim palI
For palI = 0 To MAXIND
    
    
    PicPAL.Line (palI * PicPAL.Width / (MAXIND + 1), 0)- _
            ((palI + 1) * PicPAL.Width / (MAXIND + 1), PicPAL.Height), _
            RGB(PALE(0, palI), PALE(1, palI), PALE(2, palI)), BF
    
Next
PicPAL.Refresh
End Sub

Private Sub sMOVE_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Timer_Move.Enabled = True

End Sub

Private Sub sMOVE_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
sMOVE.Value = 500
Timer_Move.Enabled = False
End Sub

Private Sub sVpos_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

curFRAME = sVpos * 10
PutFrameToPIC curFRAME
FloatFRAME = curFRAME

txtCurFrame = curFRAME
End Sub
Sub POTRACE(FileName As String, Larghezza As Integer, Optional ByVal Turdsize As Single = 25, Optional ByVal GammaAntiA As Single = 2.2, Optional ByVal TurnPolicy As String = "min")
'g 1.5 t 25


Me.Caption = "creating files PGM (portable Gray Map) " & FileName

'Shell App.Path & "\mkbitmap.exe " & FileName & ".bmp -f 4 -s 2 -t 0.40"

'-t 50 meno fitto -t 20 piu fitto
'Shell App.Path & "\potrace.exe " & filename & ".bmp -g -G 1.5 -t 30  -W " & Larghezza, vbHide
'Shell App.Path & "\potrace.exe " & filename & ".bmp -g -z white -G 1.5 -t 25  -W " & Larghezza, vbHide
'pgm
'Shell App.path & "\potrace.exe " & filename & ".bmp -g " & _
" -z " & TurnPolicy & _
        " -G " & GammaAntiA & _
        " -t " & Turdsize & _
        " -W " & Larghezza ', vbHide

'Stop

ShellEx App.path & "\Potrace\potrace.exe", essSW_SHOWDEFAULT, FileName & ".bmp -g " & _
        " -z " & TurnPolicy & _
        " -G " & GammaAntiA & _
        " -t " & Turdsize & _
        " -W " & Larghezza, App.path & "\potrace\", , Me.hWnd
'TimerWaitPotrace.Enabled = True
'ToWaitPotrace = True

'svg
'Shell App.path & "\potrace.exe " & filename & ".bmp -s " & _
" -z " & TurnPolicy & _
        " -G " & GammaAntiA & _
        " -t " & Turdsize & _
        " -W " & Larghezza ', vbHide
'eps
'Shell App.path & "\potrace.exe " & filename & ".bmp -e -c " & _
" -z " & TurnPolicy & _
        " -G " & GammaAntiA & _
        " -t " & Turdsize & _
        " -W " & Larghezza ', vbHide


End Sub

Sub DrawPGM_C(fName As String)
Dim S As String
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
Dim x As Integer


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
    For x = 0 To W - 1
        G = G + 1
        C = InFile(G)
        
        If C < 255 Then
                       
            
            CC = GetPixel(PIC.hDC, x, Y)
            LongToRGB CC, sR, sG, sB
            
            C2 = 255 - C
            
            NewR = CInt(sR) - C2 '(256 - C)
            NewG = CInt(sG) - C2 '(256 - C)
            NewB = CInt(sB) - C2 '(256 - C)
            
            If NewR < 0 Then NewR = 0
            If NewG < 0 Then NewG = 0
            If NewB < 0 Then NewB = 0
            
            SetPixel PIC.hDC, x, Y, RGB(NewR, NewG, NewB)
            
        End If
        
    Next x
Next Y
Close 1
'PIC.Refresh

End Sub

Sub DrawPGM2(fName As String, COL)
Dim S As String
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

Dim x As Integer
Dim Y As Integer


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
    For x = 0 To W - 1
        
        G = G + 1
        C = InFile(G)
        
        If C = 0 Then
            SetPixel PIC.hDC, x, Y, RGB(PALE(0, COL), PALE(1, COL), PALE(2, COL))
        Else
            If C < 255 Then
                
                C1 = C / 255
                C2 = 1 - C1
                
                CC = GetPixel(PIC.hDC, x, Y)
                LongToRGB CC, sR, sG, sB
                
                NewR = PALE(0, COL) '
                NewG = PALE(1, COL) '
                NewB = PALE(2, COL) '
                
                NewR = CInt(NewR) * C2 + sR * C1
                NewG = CInt(NewG) * C2 + sG * C1
                NewB = CInt(NewB) * C2 + sB * C1
                
                If NewR < 0 Then NewR = 0
                If NewG < 0 Then NewG = 0
                If NewB < 0 Then NewB = 0
                '        NewR = ((255 - C) / 255) * NewR
                '        NewG = ((255 - C) / 255) * NewG
                '        NewB = ((255 - C) / 255) * NewB
                
                SetPixel PIC.hDC, x, Y, RGB(NewR, NewG, NewB)
                
            End If
        End If
    Next x
Next Y


'Set bmp = New cDIB

'If bmp.CreateFromFile(App.Path & "\PIC.bmp") <> True Then
'        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
'     'GoTo Error
'End If


Close 1
End Sub
 Function Readline(StopCHR As Byte) As String
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

Private Sub Timer1_Timer()


End Sub

Private Sub sWWWcol_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
WWWcols = sWWWcol '360

PICeff.Cls
PICeff.Width = WWWcols
PICeff.Height = PIC.Height / PIC.Width * WWWcols - 1
'If PICeff.Height / 2 <> PICeff.Height \ 2 Then PICeff.Height = PICeff.Height + 1

SetupSomething

End Sub



Private Sub sWWWcont_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
WWWcont = sWWWcont '600

PicEffC.Cls
PicEffC.Width = WWWcont
PicEffC.Height = PIC.Height / PIC.Width * WWWcont - 1
'If PicEffC.Height / 2 <> PicEffC.Height \ 2 Then PicEffC.Height = PicEffC.Height + 1


PicEffC.Refresh
End Sub

Private Sub Timer_Move_Timer()
curFRAME = curFRAME + (sMOVE - 500) / 7

sVpos.Value = curFRAME / 10

PutFrameToPIC curFRAME
FloatFRAME = curFRAME

txtCurFrame = curFRAME

End Sub

Private Sub TIMERDO_Sepa_Continue_Timer()

If IsProcessRunning("potrace.exe") Then Exit Sub
TIMERDO_Sepa_Continue.Enabled = False


If BLEND > 0 Then
Set PIC.Picture = FreeImage_AdjustContrastIOP(PIC.Image, 500) '3000
PIC.Refresh
End If

Eff.SetSource PIC

For IndexC = 0 To MAXIND
    Me.Caption = "ADDING color " & IndexC & " of " & MAXIND
    
    Eff.EffDrawPGM "OUT" & IndexC, IndexC
    Eff.PutEffToPic PIC, ePGM
    
    'DrawPGM2 "OUT" & IndexC, IndexC

Next
PIC.Refresh

Me.Caption = "Countour... "
Do_Contour



End Sub

Private Sub TIMERDo_Cont_Continue_Timer()
Dim TmpColor As Long

If chGlobalMODE.Value = Unchecked Then 'Or chBILATcont.Value = Checked Then
If IsProcessRunning("potrace.exe") Then Exit Sub
TIMERDo_Cont_Continue.Enabled = False

'Do
'DoEvents
'Loop While (ToWaitPotrace)
Eff.SetSource PIC
Eff.EffDrawPGM_C "CCC"
Eff.PutEffToPic PIC, ePGM_C

'DrawPGM_C "CCC"
'Stop
End If
TIMERDo_Cont_Continue.Enabled = False
If ((FloatFRAME - StartF) < AVIin.FPS * 5) Or (FloatFRAME > EndF - AVIin.FPS * 5) Then
PIC.CurrentX = 8
PIC.CurrentY = PIC.Height - 25
TmpColor = (FloatFRAME * AVIin.FPS - StartF) * 0.5
TmpColor = TmpColor Mod 510
If TmpColor > 255 Then TmpColor = 510 - TmpColor
PIC.ForeColor = RGB(TmpColor, TmpColor, TmpColor)
PIC.Print "Video to Cartoon V" & App.Major & "." & App.Minor & "   by Roberto Mior"
End If

PIC.Refresh
SavePicture PIC.Image, App.path & "\Frames\OUT_" & Format(curFRAME, "00000000") & ".bmp"


FloatFRAME = FloatFRAME + MyStep

If FloatFRAME < EndF Then
    
    
    PB.Value = FloatFRAME
    
    curFRAME = CLng(Int(FloatFRAME))
    
    Status = "Frame " & curFRAME & "   " & curFRAME - StartF & "/" & EndF - StartF & vbCrLf & _
            "Elapsed " & Format((Timer - TEMPO) / 86400, "HH:MM:SS") & "  Remain: " & Format((((Timer - TEMPO) / 86400) / (FloatFRAME - StartF)) * (EndF - FloatFRAME), "HH:MM:SS")
    
    
    
    'If Dir(App.path & "\Potrace\*.bmp") <> "" Then Kill App.path & "\Potrace\*.bmp"
    If Dir(App.path & "\Potrace\*.pgm") <> "" Then Kill App.path & "\Potrace\*.pgm"

    PutFrameToPIC curFRAME
    DoEvents
    
    cmdDoFRAME_Click
    
Else
    PB.Value = PB.Max
    
    If OneFrameTime Then
        FrameTime = Timer - FrameTime
        txtFrameTime = FrameTime
        OneFrameTime = False
    End If
    
    If Dir(App.path & "\frames\0*.*") <> "" Then
        Kill App.path & "\frames\0*.*"
        If chSaveAvi Then Beep: Beep: cmdBuildAVI_Click
    End If
    
    Me.Caption = "All Frames Done!"
    cmdAbort.Visible = False
    
    Status = "Done!"
    
End If




End Sub

Private Sub TimerWaitPotrace_Timer()

'ToWaitPotrace = IsProcessRunning("potrace.exe")
'If Not ToWaitPotrace Then TimerWaitPotrace.Enabled = False

End Sub

Sub EffQuantizeFreeImage()
Dim palI As Integer

Dim PAL() As RGBQUAD

Dim Qmode As FREE_IMAGE_QUANTIZE

SavePicture PICeff.Image, App.path & "\Potrace\Blurred.BMP"

FIdib = FreeImage_Load(FIF_BMP, App.path & "\Potrace\Blurred.BMP", 0)
If ChQM = Checked Then Qmode = FIQ_NNQUANT Else: Qmode = FIQ_WUQUANT
'                                        FIQ_NNQUANT 'seems better
FIdibQ = FreeImage_ColorQuantizeEx(FIdib, Qmode, MAXIND + 1, 0, 0)

FreeImage_Save FIF_BMP, FIdibQ, App.path & "\Potrace\blurredQ.BMP", 0

PICeff = LoadPicture(App.path & "\Potrace\blurredQ.BMP")

PAL = FreeImage_GetPaletteEx(FIdibQ)

For palI = 0 To MAXIND
    PALE(0, palI) = PAL(palI).rgbRed
    PALE(1, palI) = PAL(palI).rgbGreen
    PALE(2, palI) = PAL(palI).rgbBlue
    PicPAL.Line (palI * PicPAL.Width / (MAXIND + 1), 0)- _
            ((palI + 1) * PicPAL.Width / (MAXIND + 1), PicPAL.Height), _
            RGB(PAL(palI).rgbRed, PAL(palI).rgbGreen, PAL(palI).rgbBlue), BF
Next
PicPAL.Refresh

FreeImage_Unload (FIdib)
FreeImage_Unload (FIdibQ)
Erase PAL
End Sub

Public Sub SeparateColors()
Dim C As Long
Dim R As Byte
Dim G As Byte
Dim B As Byte
Dim Rgb1 As Long
Dim RGB2 As Long
Dim x As Integer
Dim Y As Integer
Dim COL As Long
Dim I As Long
Dim QPC As Long
'Stop


For I = 0 To MAXIND
    SEPA(I).Cls
Next I


For Y = 0 To PICeff.Height - 1
    For x = 0 To PICeff.Width - 1
        For COL = 0 To MAXIND
            
            Rgb1 = RGB(PALE(0, COL), PALE(1, COL), PALE(2, COL))
            QPC = GetPixel(PICeff.hDC, x, Y)
            If QPC = Rgb1 Then SetPixel SEPA(COL).hDC, x, Y, RGB(0, 0, 0)
            
        Next COL
    Next x
Next Y

For I = 0 To MAXIND
    SEPA(I).Refresh
    SavePicture SEPA(I).Image, App.path & "\Potrace\OUT" & I & ".bmp"
Next



End Sub

Sub SetupSomething()

PICeff.Cls
PICeff.Width = WWWcols
PICeff.Height = PIC.Height / PIC.Width * WWWcols - 1
'If PICeff.Height / 2 <> PICeff.Height \ 2 Then PICeff.Height = PICeff.Height + 1


SEPA(0).Width = PICeff.Width
SEPA(0).Height = PICeff.Height
SEPA(0).Cls

If SEPA.Count < MAXIND Then
    For I = SEPA.Count To MAXindBAR.Max '+ 1
        Load SEPA(I)
    Next I
End If

For I = 0 To MAXIND
    SEPA(I).Width = PICeff.Width
    SEPA(I).Height = PICeff.Height
    If I <> 0 Then SEPA(I).Left = SEPA(I - 1).Left + SEPA(I - 1).Width + 10
    'SEPA(I).Visible = True
    SEPA(I).Cls
    SEPA(I).Refresh
Next I
End Sub


Sub Do_Contour()
'PutFrameToPIC curFRAME

'Stop




Call SetStretchBltMode(PicEffC.hDC, STRETCHMODE)
Call StretchBlt(PicEffC.hDC, 0, 0, PicEffC.Width - 1, PicEffC.Height - 1, _
        PICtmp.hDC, 0, 0, PICtmp.Width - 1, PICtmp.Height - 1, vbSrcCopy)

PicEffC.Refresh


Eff.SetSource PicEffC
'Stop

Eff.effCONTOUR cEnhanced
Eff.PutEffToPic PicEffC, eContour
PICeff.Refresh


Set PicEffC.Picture = FreeImage_AdjustContrastIOP(PicEffC.Image, Ccontra) '3000
PicEffC.Refresh

Set PicEffC.Picture = FreeImage_AdjustBrightnessIOP(PicEffC.Image, Cbright) '100
PicEffC.Refresh



'Kill App.path & "\CCC.pgm"
SavePicture PicEffC.Image, App.path & "\Potrace\CCC.bmp"
'POTRACE "CCC", PIC.Width, 20
POTRACE "CCC", PIC.Width, 15

TIMERDo_Cont_Continue.Enabled = True

End Sub


Sub Do_effblur()
Dim palI As Long


PICeff.Width = WWWcols
PICeff.Height = PIC.Height / PIC.Width * WWWcols - 1



Call SetStretchBltMode(PICeff.hDC, STRETCHMODE)
Call StretchBlt(PICeff.hDC, 0, 0, PICeff.Width, PICeff.Height, _
        PICtmp.hDC, 0, 0, PICtmp.Width - 1, PICtmp.Height - 1, vbSrcCopy)


Set PICeff.Picture = FreeImage_AdjustBrightnessIOP(PICeff.Image, Bright) '-37' - 33
Set PICeff.Picture = FreeImage_AdjustContrastIOP(PICeff.Image, Contra) '-37' - 33

PICeff.Refresh


Eff.SetSource PICeff
Eff.effEXTENDEDBlur Satura / 10
Eff.PutEffToPic PICeff, eEXblur
PICeff.Refresh

If chMYPale.Value = Checked Then
    'Stop
    
    For I = 1 To MAXIND
        Debug.Print PALE(0, I)
    Next
    'Stop
    
    
    
    Eff.EFFQuantizeMy
    Eff.PutEffToPic PICeff, eMyQuant
    PICeff.Refresh
    
    For I = 1 To MAXIND
        Debug.Print PALE(0, I)
    Next
    'Stop
    
    For palI = 0 To MAXIND
        
        PicPAL.Line (palI * PicPAL.Width / (MAXIND + 1), 0)- _
                ((palI + 1) * PicPAL.Width / (MAXIND + 1), PicPAL.Height), _
                RGB(PALE(0, palI), PALE(1, palI), PALE(2, palI)), BF
        
    Next
    
    PicPAL.Refresh
Else
    EffQuantizeFreeImage
End If


End Sub


Sub Do_SepaAndPut()


'----
'SeparateColors
'----

'----
''''EFF.SetSource PICeff 'Use blurbYte
'----
If chMYPale.Value = Unchecked Then Eff.SetSource PICeff 'use sBYTE

For IndexC = 0 To MAXIND
    
    '----
    SEPA(IndexC).Cls
    
    If chMYPale Then
        Eff.EFFSeparateMY IndexC
    Else
        Eff.EFFSeparateFREE IndexC
    End If
    
    Eff.PutEffToPic SEPA(IndexC), eSepa
    SEPA(IndexC).Refresh
    SavePicture SEPA(IndexC).Image, App.path & "\Potrace\OUT" & IndexC & ".bmp"
    '----
Next

For IndexC = 0 To MAXIND
    POTRACE "OUT" & IndexC, PIC.Width, 15 ' 30
Next

TIMERDO_Sepa_Continue.Enabled = True



'Do
'DoEvents
'Loop While (ToWaitPotrace)

'MsgBox "POTRACE DONE! "

'SavePAL (MFN)

'For IndexC = 0 To MAXIND
'    Me.Caption = "ADDING color " & IndexC & " of " & MAXIND
'    DrawPGM2 "OUT" & IndexC, IndexC
'Next
'PIC.Refresh


End Sub



Sub BUILD_AVI()

OutputAVIName = ""

Dim fPATH As String

fPATH = App.path & "\frames\"

Dim S As String

Dim fLIST() As String
Dim C As Long

S = Dir(fPATH & "*.bmp")

If S = "" Then Exit Sub

ReDim Preserve fLIST(1)
Do
    fLIST(C) = fPATH & S
    C = C + 1
    ReDim Preserve fLIST(0 To C)
    S = Dir
Loop While S <> ""

'----------------------------------------------------------------------------------------
Dim file As cFileDlg
Dim InitDir As String
Dim szOutputAVIFile As String
Dim res As Long
Dim pfile As Long 'ptr PAVIFILE
Dim bmp As cDIB
Dim ps As Long 'ptr PAVISTREAM
Dim psCompressed As Long 'ptr PAVISTREAM
Dim strhdr As AVI_STREAM_INFO
Dim BI As BITMAPINFOHEADER
Dim opts As AVI_COMPRESS_OPTIONS
Dim pOpts As Long
Dim I As Long
Dim I2 As Long

Dim EXTRA As Integer 'Extra Frame

Debug.Print
Set file = New cFileDlg
'get an avi filename from user
With file
    .InitDirectory = App.path & "\VIDEO\"
    .DefaultExt = "avi"
    .DlgTitle = "Choose a filename to save AVI to..."
    .Filter = "AVI Files|*.avi"
    .OwnerHwnd = Me.hWnd
End With
szOutputAVIFile = "MyAVI.avi"
If file.VBGetSaveFileName(szOutputAVIFile) <> True Then Exit Sub


OutputAVIName = szOutputAVIFile
'Stop

'    Open the file for writing
res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
If (res <> AVIERR_OK) Then GoTo error

'Get the first bmp in the list for setting format
Set bmp = New cDIB
If bmp.CreateFromFile(fLIST(1)) <> True Then
    MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
    GoTo error
End If
'Stop

'   Fill in the header for the video stream
With strhdr
    .fccType = mmioStringToFOURCC("vids", 0&) '// stream type video
    .fccHandler = 0& '// default AVI handler
    .dwScale = 1
    .dwRate = Val(txtFPS) * (Val(cmbEXTRA) + 1) '// fps
    .dwSuggestedBufferSize = bmp.SizeImage '// size of one frame pixels
    Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height) '// rectangle for stream
End With

'validate user input
If strhdr.dwRate < 1 Then strhdr.dwRate = 1
If strhdr.dwRate > 30 Then strhdr.dwRate = 30

'   And create the stream
res = AVIFileCreateStream(pfile, ps, strhdr)
If (res <> AVIERR_OK) Then GoTo error

'get the compression options from the user
'Careful! this API requires a pointer to a pointer to a UDT
pOpts = VarPtr(opts)
res = AVISaveOptions(Me.hWnd, _
        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
        1, _
        ps, _
        pOpts) 'returns TRUE if User presses OK, FALSE if Cancel, or error code
If res <> 1 Then 'In C TRUE = 1
    Call AVISaveOptionsFree(1, pOpts)
    GoTo error
End If

'make compressed stream
res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
If res <> AVIERR_OK Then GoTo error

'set format of stream according to the bitmap
With BI
    .biBitCount = bmp.BitCount
    .biClrImportant = bmp.ClrImportant
    .biClrUsed = bmp.ClrUsed
    .biCompression = bmp.Compression
    .biHeight = bmp.Height
    .biWidth = bmp.Width
    .biPlanes = bmp.Planes
    .biSize = bmp.SizeInfoHeader
    .biSizeImage = bmp.SizeImage
    .biXPelsPerMeter = bmp.XPPM
    .biYPelsPerMeter = bmp.YPPM
End With

'set the format of the compressed stream
res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
If (res <> AVIERR_OK) Then GoTo error

'   Now write out each video frame
I2 = 0
For I = 0 To C - 1
    
    Me.Caption = "Creating AVI file...  Frame " & I & " of " & C - 1
    DoEvents
    
    
    For EXTRA = 0 To Val(cmbEXTRA)
     
    bmp.CreateFromFile (fLIST(I)) 'load the bitmap (ignore errors)
    
   
    res = AVIStreamWrite(psCompressed, _
            I2, _
            1, _
            bmp.PointerToBits, _
            bmp.SizeImage, _
            AVIIF_KEYFRAME, _
            ByVal 0&, _
            ByVal 0&)
    If res <> AVIERR_OK Then GoTo error
    'Show user feedback
    'imgPreview.Picture = LoadPicture(lstDIBList.Text)
    'imgPreview.Refresh
    'lblStatus = "Frame number " & i & " saved"
    'lblStatus.Refresh
    I2 = I2 + 1
    
    Next EXTRA


Next
Me.Caption = "Avi file  Created!"



error:
'   Now close the file
Set file = Nothing
Set bmp = Nothing

If (ps <> 0) Then Call AVIStreamClose(ps)

If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

If (pfile <> 0) Then Call AVIFileClose(pfile)

Call AVIFileExit

If (res <> AVIERR_OK) Then
    MsgBox "There was an error writing the file.", vbInformation, App.Title
End If

'Stop


End Sub

Sub SaveSetting()
Dim S$

S$ = InputBox("Type Name for this Set of Parameters", "?", "Setting")
If S$ = "" Then Exit Sub

If LCase(Right$(S$, 4)) <> ".ini" Then S$ = S$ & ".ini"

INI.FileName = App.path & "\INIs\" & S$

INI.WriteINI "Input File", "Long Path", CMD.FileName
INI.WriteINI "Input File", "Start Frame", CStr(StartF)
INI.WriteINI "Input File", "End Frame", CStr(EndF)

INI.WriteINI "Output", "Width", OutSize
INI.WriteINI "Output", "Fps", txtFPS
INI.WriteINI "Output", "QuantizeMode", ChQM

INI.WriteINI "Color Params", "Main", sWWWcol
INI.WriteINI "Color Params", "Brightness", Bright
INI.WriteINI "Color Params", "Contrast", Contra
INI.WriteINI "Color Params", "Saturation", Satura

INI.WriteINI "Contour Params", "Main", sWWWcont
INI.WriteINI "Contour Params", "Brightness", Cbright
INI.WriteINI "Contour Params", "Contrast", Ccontra

INI.WriteINI "Extra", "Extra Frame", cmbEXTRA
INI.WriteINI "Extra", "Blend", BLEND
INI.WriteINI "Extra", "Blend is Gray", chBW
INI.WriteINI "Extra", "Enhanced Countour", cEnhanced

INI.WriteINI "Palette", "MAXind", CStr(MAXIND)
For I = 0 To MAXindBAR 'MAXIND
    
    INI.WriteINI "Palette", "R" & I, CStr(PALE(0, I))
    INI.WriteINI "Palette", "G" & I, CStr(PALE(1, I))
    INI.WriteINI "Palette", "B" & I, CStr(PALE(2, I))
    
Next

End Sub

Sub LoadSetting(S$)

If S$ = "" Then Exit Sub
INI.FileName = App.path & "\INIs\" & S$

'Stop

ParamOPENAVI = INI.GetINI("Input File", "Long Path")

'Stop


StartF = INI.GetINI("Input File", "Start Frame")
EndF = INI.GetINI("Input File", "End Frame")
Label1 = StartF
Label2 = EndF

OutSize = INI.GetINI("Output", "Width")
ChQM = INI.GetINI("Output", "QuantizeMode")

txtFPS = INI.GetINI("Output", "fps")

sWWWcol = INI.GetINI("Color Params", "Main")
sWWWcol_MouseUp 1, 0, 0, 0
sWWWcont = INI.GetINI("Contour Params", "Main")
sWWWcont_MouseUp 1, 0, 0, 0


Bright = INI.GetINI("Color Params", "Brightness")
Contra = INI.GetINI("Color Params", "Contrast")
Satura = INI.GetINI("Color Params", "Saturation")

Cbright = INI.GetINI("Contour Params", "Brightness")
Ccontra = INI.GetINI("Contour Params", "Contrast")

cmbEXTRA = INI.GetINI("Extra", "Extra Frame")
BLEND = Val(INI.GetINI("Extra", "Blend"))
'Stop

chBW = INI.GetINI("Extra", "Blend is Gray")
cEnhanced = INI.GetINI("Extra", "Enhanced Countour")

MAXindBAR.Value = INI.GetINI("Palette", "MaxInd")
MAXIND = MAXindBAR.Value
MAXindBAR_Click

For I = 0 To MAXIND
    
    PALE(0, I) = INI.GetINI("Palette", "R" & I)
    PALE(1, I) = INI.GetINI("Palette", "G" & I)
    PALE(2, I) = INI.GetINI("Palette", "B" & I)
    
Next
drawPALE

If ParamOPENAVI <> "" Then
    If Dir(ParamOPENAVI) <> "" Then
        Command1_Click
        sVpos = (StartF + EndF) / 20
        sVpos_MouseUp 1, 0, 0, 0
    Else
        MsgBox "Can't Find " & vbCrLf & ParamOPENAVI, vbCritical
        
    End If
    
End If

End Sub

