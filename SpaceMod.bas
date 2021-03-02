Attribute VB_Name = "SpaceMod"
Public Type POINTAPI
  X As Integer
  Y As Integer
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_RESOURCE = &H40004     '  name is a resource name or atom

Public Type t3DVector
  X As Integer
  Y As Integer
  z As Integer
End Type

Public Type tSpaceship
  zm As Integer
  RollAngle As Integer
  TurnAngle As Integer
  PitchAngle As Integer
End Type

Public NUMSTARS As Integer
Public CometsOn As Boolean
Public Const VIEWWIDTH = 200
Public Const VIEWHEIGHT = 150
Public Const VIEWDEPTH = 300
Public Star() As t3DVector

Public Ship As tSpaceship

Public i As Integer
Public i2 As Integer
Public X As Integer
Public Y As Integer
Public z As Integer

Public Const LENS = VIEWDEPTH
Public LensDivDist As Single

Public DispWidth As Integer
Public DispHeight As Integer
Public Const PBWIDTH = 600
Public Const PBHEIGHT = 450
Public Const CX = 300
Public Const CY = 225
Public Sine(0 To 359) As Single
Public Cosine(0 To 359) As Single

Public lpPoint As POINTAPI

Public Const PI = 3.14159265358979 'obvious
Public Const PIdiv180 = PI / 180

Public Sub BuildTrigTable()
For i = 0 To 359
  Sine(i) = Sin(i * PIdiv180)
  Cosine(i) = Cos(i * PIdiv180)
Next
End Sub
