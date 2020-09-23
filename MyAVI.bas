Attribute VB_Name = "MyAVI"
Public res As Long 'result code
Public ofd As cFileDlg 'OpenFileDialog class
Public szFile As String 'filename
Public pAVIFile As Long 'pointer to AVI file interface (PAVIFILE handle)
Public pAVIStream As Long 'pointer to AVI stream interface (PAVISTREAM handle)
Public numFrames As Long 'number of frames in video stream
Public firstFrame As Long 'position of the first video frame
Public fileInfo As AVI_FILE_INFO 'file info struct
Public streamInfo As AVI_STREAM_INFO 'stream info struct
Public dib As cDIB
Public pGetFrameObj As Long 'pointer to GetFrame interface
Public pDIB As Long 'pointer to packed DIB in memory
Public bih As BITMAPINFOHEADER 'infoheader to pass to GetFrame functions
Public i As Long



Public I2 As Long

Public ElaStart As Long
Public ElaEnd As Long

Public Bright As Integer


'simple UDT containing parameters of first BMP file user chooses
'all the following BMPs should be the same format so there will be no problems in writing the vidstream
Private Type PARAMS
    Init As Boolean
    Width As Long
    Height As Long
    bpp As Long
End Type

Public Declare Function SetRect Lib "user32.dll" _
        (ByRef lprc As AVI_RECT, ByVal xLeft As Long, ByVal yTop As Long, ByVal xRight As Long, ByVal yBottom As Long) As Long 'BOOL1
Public m_params As PARAMS




Type tmAVi
    
    W_in As Long
    H_in As Long
    W_out As Long
    H_out  As Long
    
    
    MaxFrames As Long
    FPS As Long
    
    filename As String
    
    
End Type



Public AVIin As tmAVi

'================================
'Api used to copy an Array!
'================================
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        Destination As Any, Source As Any, ByVal Length As Long)


Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" _
        (Var() As Any) As Long



Public FPSmultiplier
Public fpsREAL


