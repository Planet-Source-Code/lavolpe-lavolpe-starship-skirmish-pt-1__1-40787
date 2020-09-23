Attribute VB_Name = "modMain"
Public Declare Function Rectangle Lib "GDI32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LR_COPYDELETEORG As Long = &H8
Public Const LR_COPYRETURNORG = &H4

Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

' Used to read INI files
Private Declare Function GetPrivateProfileSection Lib "Kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long

Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long

Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFilename As String) As Long

Private Declare Function WritePrivateProfileSection Lib "Kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFilename As String) As Long


Public Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type PlayerData
    Name As String
    ID() As Integer
    Location() As RECT
    BMPxy() As RECT
    Shield(1 To 3) As Integer
    Cloak As Integer
    CloakRevealed As Byte
    Strength() As Integer
    Ammo(0 To 3) As Byte
    Scans(0 To 2) As Byte
    Size() As Byte
    Color() As Long
End Type

Public Const Ammo_Laser As Byte = 50
Public Const Ammo_Cannon As Byte = 20
Public Const Ammo_Torpedo As Byte = 10

Public GP As Integer
Public Interval As Long
Public PlayerID As Integer
Public Player() As PlayerData
Public sourceBmp As IPictureDisp
Public miscBmp As IPictureDisp
Public GridItems() As RECT
Public GameMode As Integer
    ' 1 = 8 Ship & shields
    ' 2 = 5 ship & shields
    ' 3 = 8 Ship & no shields
    ' 4 = 5 ship & no shields
    ' 5 = 8 ship & salvo & no shields
    ' 6 = 5 Ship & salvo & no shields

Public Sub MaxSizeMe(ObjID As Form)
    Dim UserWindow As RECT
    ' Returns UserWindow in twips
    SystemParametersInfo SPI_GETWORKAREA, 0, UserWindow, 0
    ObjID.Move UserWindow.Left * Screen.TwipsPerPixelX, _
               UserWindow.Top * Screen.TwipsPerPixelY, _
               UserWindow.Right * Screen.TwipsPerPixelX, _
               UserWindow.Bottom * Screen.TwipsPerPixelY
End Sub

Public Function GetGridCoordsString(X As Single, Y As Single) As String
Dim XY(0 To 1) As Integer
XY(0) = Int(Y / Interval)
XY(1) = Int(X / Interval)
GetGridCoordsString = Chr$(XY(0) + 65) & ":" & XY(1) + 1
End Function


Public Function ConvertGridCoord2Integer(X As Long, Y As Long, Optional GridID As Integer) As Integer
If GridID Then
   Y = Int(GridID / 10.1) * Interval
   X = (GridID - (Int(GridID / 10.1) * 10) - 1) * Interval
Else
    Dim X1 As Long, Y1 As Long
    X1 = X \ Interval
    Y1 = Y \ Interval
    ConvertGridCoord2Integer = X1 + (Y1 * 10) + 1
End If
End Function

Public Function GetGridCoordsXY(X As Single, Y As Single, Optional GridID As String, Optional Index As Integer = -1) As RECT
Dim XY(0 To 1) As Integer
If Len(GridID) Then
    XY(1) = Val(Mid(GridID, 3)) - 1
    XY(0) = Asc(Left(GridID, 1)) - 65
Else
    XY(0) = Int(Y / Interval)
    XY(1) = Int(X / Interval)
End If
With GetGridCoordsXY
    .Left = XY(1) * Interval
    .Top = XY(0) * Interval
    If Index > -1 Then
        .Left = .Left + 1
        .Top = .Top + 1
        With Player(PlayerID).Location(Index)
            GetGridCoordsXY.Right = .Right - .Left + GetGridCoordsXY.Left
            GetGridCoordsXY.Bottom = .Bottom - .Top + GetGridCoordsXY.Top
        End With
    Else
        .Right = .Left + Interval - 2
        .Bottom = .Top + Interval - 2
    End If
End With
End Function

Public Sub BeginPlaySound(ResourceId As String, Optional bStop As Boolean = False, Optional bWait As Boolean = False)

If Not frmBoard.mnuOpts(0).Checked Then Exit Sub
If bStop Then
    sndPlaySound ByVal vbNullString, &H1
    Exit Sub
End If

Dim sndFlags As Long
' &H0 = Sync (halts program until sound done)
' &H1 = Async (returns immediately)
' &H2 = No_Default otherwise you get the computer beep
' &H2000 = No_Wait, plays immediately
' &H20000 = FileName
' &H4 = Memory
sndFlags = Abs(CInt(bWait) + 1) Or &H2 Or &H2000

If IsNumeric(ResourceId) Then
    Dim SoundBuffer() As Byte
    SoundBuffer = LoadResData(Val(ResourceId), "Custom")
    sndPlaySound SoundBuffer(0), sndFlags Or &H4
    Erase SoundBuffer
Else
    sndPlaySound ByVal ResourceId, sndFlags Or &H20000
End If
End Sub

Public Sub CalculateRatio(rtnX As Long, rtnY As Long, destX As Long, destY As Long, imgX As Long, imgY As Long)
Dim Ratio(0 To 1) As Single
On Error Resume Next
Ratio(0) = destX / imgX
Ratio(1) = destY / imgY
If Ratio(1) < Ratio(0) Then Ratio(0) = Ratio(1)
If Ratio(0) > 1 Then Ratio(0) = 1
rtnX = imgX * Ratio(0)
rtnY = imgY * Ratio(0)
End Sub

Public Function MakeRectangle(rLeft As Long, rTop As Long, rWidth As Long, rHeight As Long) As RECT
MakeRectangle.Left = rLeft
MakeRectangle.Top = rTop
If rWidth Then MakeRectangle.Right = rLeft + rWidth
If rHeight Then MakeRectangle.Bottom = rTop + rHeight
End Function

Public Function ShipsRemaining(ForWho) As Integer
Dim Looper As Integer, iCount As Integer
For Looper = 0 To UBound(Player(1).ID)
    iCount = iCount + CInt(Player(ForWho).Strength(Looper) > 0)
Next
ShipsRemaining = Abs(iCount)
End Function

Public Sub DownLoadImages()
Dim sFile As String, Looper As Integer, fNr As Integer, vData() As Byte
For Looper = 1 To 8
    sFile = Choose(Looper, "aniShip1.gif", "aniShip2.gif", "aniDish.gif", "Boom.gif", "GridMisc.gif", "Ships.gif", "blast.wav", "StarshipSkirmish.hlp")
    If Len(Dir(App.Path & "\" & sFile)) = 0 Then
        vData = LoadResData(103 + Looper, "Custom")
        fNr = FreeFile()
        Open App.Path & "\" & sFile For Binary As #fNr
        Put #fNr, , vData()
        Close #fNr
    End If
Next
End Sub


Public Sub LoadShipCoords(shipRect() As RECT)
ReDim shipRect(0 To 14)
shipRect(0).Left = 0: shipRect(0).Top = 0
shipRect(0).Right = 82: shipRect(0).Bottom = 55
shipRect(1).Left = 0: shipRect(1).Top = 55
shipRect(1).Right = 130: shipRect(1).Bottom = 110
shipRect(2).Left = 0: shipRect(2).Top = 110
shipRect(2).Right = 95: shipRect(2).Bottom = 165
shipRect(3).Left = 0: shipRect(3).Top = 165
shipRect(3).Right = 84: shipRect(3).Bottom = 220
shipRect(4).Left = 0: shipRect(4).Top = 220
shipRect(4).Right = 128: shipRect(4).Bottom = 275
shipRect(5).Left = 0: shipRect(5).Top = 275
shipRect(5).Right = 101: shipRect(5).Bottom = 330
shipRect(6).Left = 0: shipRect(6).Top = 330
shipRect(6).Right = 66: shipRect(6).Bottom = 385
shipRect(7).Left = 0: shipRect(7).Top = 385
shipRect(7).Right = 95: shipRect(7).Bottom = 440
shipRect(8).Left = 0: shipRect(8).Top = 440
shipRect(8).Right = 84: shipRect(8).Bottom = 495
shipRect(9).Left = 0: shipRect(9).Top = 495
shipRect(9).Right = 104: shipRect(9).Bottom = 550
shipRect(10).Left = 0: shipRect(10).Top = 550
shipRect(10).Right = 86: shipRect(10).Bottom = 605
shipRect(11).Left = 0: shipRect(11).Top = 605
shipRect(11).Right = 131: shipRect(11).Bottom = 655
shipRect(12).Left = 0: shipRect(12).Top = 655
shipRect(12).Right = 56: shipRect(12).Bottom = 710
shipRect(13).Left = 0: shipRect(13).Top = 710
shipRect(13).Right = 105: shipRect(13).Bottom = 765
shipRect(14).Left = 0: shipRect(14).Top = 765
shipRect(14).Right = 106: shipRect(14).Bottom = 820
End Sub

Public Sub LoadGridMiscCoords()
ReDim GridItems(0 To 6)
GridItems(0).Left = 0: GridItems(0).Top = 0         ' Blue shield
GridItems(0).Right = 32: GridItems(0).Bottom = 32
GridItems(1).Left = 0: GridItems(1).Top = 32        ' Green shield
GridItems(1).Right = 32: GridItems(1).Bottom = 64
GridItems(2).Left = 0: GridItems(2).Top = 64        ' Yellow shield
GridItems(2).Right = 32: GridItems(2).Bottom = 96
GridItems(3).Left = 0: GridItems(3).Top = 96        ' Red shield
GridItems(3).Right = 32: GridItems(3).Bottom = 128
GridItems(4).Left = 0: GridItems(4).Top = 128       ' Skull
GridItems(4).Right = 45: GridItems(4).Bottom = 190
GridItems(5).Left = 0: GridItems(5).Top = 190       ' Cloak
GridItems(5).Right = 29: GridItems(5).Bottom = 217
GridItems(6).Left = 0: GridItems(6).Top = 217       ' Mine
GridItems(6).Right = 19: GridItems(6).Bottom = 235
End Sub

Sub ShellSortNumbers(vArray As Variant)
  Dim lLoop1 As Long
  Dim lHold As Long
  Dim lHValue As Long
  Dim lTemp As Variant

  lHValue = LBound(vArray)
  Do
    lHValue = 3 * lHValue + 1
  Loop Until lHValue > UBound(vArray)
  Do
    lHValue = lHValue / 3
    For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
      lTemp = vArray(lLoop1)
      lHold = lLoop1
      Do While vArray(lHold - lHValue) > lTemp
        vArray(lHold) = vArray(lHold - lHValue)
        lHold = lHold - lHValue
        If lHold < lHValue Then Exit Do
      Loop
      vArray(lHold) = lTemp
    Next lLoop1
  Loop Until lHValue = LBound(vArray)
End Sub

Public Function ReadWriteINI(Mode As String, BKfile As String, tmpSecName As String, tmpKeyname As String, _
    Optional tmpKeyValue As String = "*****", _
    Optional DeleteSection As Boolean = False) As String

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo ReadWriteINI_General_ErrTrap

If DeleteSection = True Then
    WritePrivateProfileSection tmpSecName, "", BKfile
    Exit Function
End If
Dim tmpString As String, tmpCounter As Integer
Dim secname As String, tmpTimer As Single
Dim KeyName As String
Dim keyvalue As String
Dim anInt
Dim defaultkey As String
On Error GoTo ReadWriteINIError

ReadWriteINI = tmpKeyValue
If IsNull(Mode) Or Len(Mode) = 0 Then Exit Function
If IsNull(tmpSecName) Or Len(tmpSecName) = 0 Then Exit Function
If IsNull(tmpKeyname) Or Len(tmpKeyname) = 0 Then Exit Function

secname = tmpSecName
KeyName = tmpKeyname
keyvalue = tmpKeyValue
defaultkey = tmpKeyValue
tmpString = tmpKeyValue
' ******* WRITE MODE *************************************
  If UCase(Mode) = "WRITE" Then
        If keyvalue = "" Then keyvalue = vbNullString
      anInt = WritePrivateProfileString(secname, KeyName, keyvalue, BKfile)
      If anInt > 0 Then anInt = 1
  Else
  ' *******  READ MODE *************************************
    If UCase(Mode) = "GET" Then
ReadFileNow:
      keyvalue = String$(255, 32)
      anInt = GetPrivateProfileString(secname, KeyName, defaultkey, keyvalue, Len(keyvalue), BKfile)
      If Left(keyvalue, Len(tmpKeyValue) + 1) <> tmpKeyValue & Chr$(0) Then     ' *** got it
         tmpString = keyvalue
         tmpString = RTrim(tmpString)
         If Len(tmpString) Then tmpString = Left(tmpString, Len(tmpString) - 1)
      End If
   End If
  End If
If anInt > 0 Then ReadWriteINI = tmpString
Exit Function
  ' *******
ReadWriteINIError:
Exit Function

' Inserted by LaVolpe OnError Insertion Program.
ReadWriteINI_General_ErrTrap:
MsgBox "Err: " & Err.Number & " - Procedure: ReadWriteINI" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
End Function



