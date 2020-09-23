VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnimator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private aniScreen As PictureBox
Private aniActive As Boolean
Private aniFile As IPictureDisp
Private FrameNow As Integer
Private FrameLast As Integer
Private FrameHeight As Long
Private FrameWidth As Long
Private FieldRect As RECT
Private bmpRect As RECT

Public Property Let animation_Screen(picBox As PictureBox)
Set aniScreen = picBox
End Property

Public Sub animation_File(sFile As String, FrameX As Long, FrameH As Long, xScale As Long, yScale As Long)

Set aniFile = LoadPicture(sFile)
FrameLast = FrameH
FrameHeight = FrameX
FrameNow = 1
FieldRect.Right = aniScreen.Width / xScale
FieldRect.Bottom = aniScreen.Height / yScale
End Sub

Public Sub ShowNextFrame(Optional FrameNumber As Integer, Optional bStopAfter As Boolean)
If aniActive = False And FrameNumber = 0 Then Exit Sub
If FrameNumber Then
    FrameNow = FrameNumber
    If bStopAfter = False Then aniActive = True
Else
    If FrameNow = FrameLast Then FrameNow = 0
    FrameNow = FrameNow + 1
End If
aniScreen.Cls
bmpRect.Top = (FrameNow - 1) * FrameHeight
bmpRect.Bottom = bmpRect.Top + FrameHeight
DrawTransparentBitmap aniScreen.hdc, FieldRect, aniFile.handle, bmpRect, -1
aniScreen.Refresh
If bStopAfter = True Then aniActive = False
End Sub

Private Sub Class_Initialize()
Set aniFile = Nothing
Set aniScreen = Nothing
End Sub
