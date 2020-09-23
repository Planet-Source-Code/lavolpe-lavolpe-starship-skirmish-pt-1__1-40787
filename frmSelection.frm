VERSION 5.00
Begin VB.Form frmSelection 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ship Selection for Player - "
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picShield 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DragMode        =   1  'Automatic
      Height          =   360
      Index           =   2
      Left            =   7290
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2970
      Width           =   360
   End
   Begin VB.PictureBox picShield 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DragMode        =   1  'Automatic
      Height          =   360
      Index           =   1
      Left            =   6705
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2970
      Width           =   360
   End
   Begin VB.PictureBox picShield 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DragMode        =   1  'Automatic
      Height          =   360
      Index           =   0
      Left            =   6075
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2970
      Width           =   360
   End
   Begin VB.PictureBox picCloak 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DragMode        =   1  'Automatic
      Height          =   360
      Left            =   6645
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5295
      Width           =   360
   End
   Begin VB.Frame frameMode 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   0
      Left            =   5910
      TabIndex        =   15
      Top             =   1200
      Width           =   1950
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " Cloaking Device "
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   29
         Top             =   2340
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   " Shields "
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   28
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSelection.frx":0000
         ForeColor       =   &H00FFFFFF&
         Height          =   1470
         Index           =   5
         Left            =   90
         TabIndex        =   31
         Top             =   2625
         Width           =   1710
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   2100
         Index           =   1
         Left            =   0
         Top             =   2430
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSelection.frx":00A0
         ForeColor       =   &H00FFFFFF&
         Height          =   1470
         Index           =   4
         Left            =   105
         TabIndex        =   30
         Top             =   285
         Width           =   1710
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   2130
         Index           =   0
         Left            =   15
         Top             =   90
         Width           =   1905
      End
   End
   Begin VB.Frame frameMode 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   4710
      Index           =   2
      Left            =   5910
      TabIndex        =   38
      Top             =   1200
      Width           =   1950
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "The player with the most ships on the board always has the advantage since that player gets more shots each turn they shoot."
         ForeColor       =   &H00FFFFFF&
         Height          =   1290
         Index           =   9
         Left            =   120
         TabIndex        =   41
         Top             =   2670
         Width           =   1725
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "With the Salvo version, you get to fire one shot for each ship you still have that hasn't been destroyed."
         ForeColor       =   &H00FFFFFF&
         Height          =   960
         Index           =   8
         Left            =   135
         TabIndex        =   40
         Top             =   1380
         Width           =   1725
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This version of the game requires only one hit to destroy any section of any ship."
         ForeColor       =   &H00FFFFFF&
         Height          =   960
         Index           =   7
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1725
      End
   End
   Begin VB.Frame frameMode 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   4710
      Index           =   1
      Left            =   5910
      TabIndex        =   36
      Top             =   1200
      Width           =   1950
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This version of the game requires only one hit to destroy any section of any ship."
         ForeColor       =   &H00FFFFFF&
         Height          =   960
         Index           =   6
         Left            =   120
         TabIndex        =   37
         Top             =   1725
         Width           =   1725
      End
   End
   Begin VB.PictureBox picShipSelect 
      AutoRedraw      =   -1  'True
      Height          =   885
      Index           =   4
      Left            =   3435
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3630
      Width           =   885
   End
   Begin VB.PictureBox picShipSelect 
      AutoRedraw      =   -1  'True
      Height          =   885
      Index           =   7
      Left            =   3435
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4980
      Width           =   885
   End
   Begin VB.PictureBox picShipSelect 
      AutoRedraw      =   -1  'True
      Height          =   885
      Index           =   6
      Left            =   2520
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4980
      Width           =   885
   End
   Begin VB.PictureBox picShipSelect 
      AutoRedraw      =   -1  'True
      Height          =   885
      Index           =   5
      Left            =   1590
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4980
      Width           =   885
   End
   Begin VB.PictureBox picShipSelect 
      AutoRedraw      =   -1  'True
      Height          =   885
      Index           =   3
      Left            =   2520
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3630
      Width           =   885
   End
   Begin VB.PictureBox picShipSelect 
      AutoRedraw      =   -1  'True
      Height          =   885
      Index           =   2
      Left            =   1590
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3630
      Width           =   885
   End
   Begin VB.PictureBox picShipSelect 
      AutoRedraw      =   -1  'True
      Height          =   885
      Index           =   1
      Left            =   1590
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2295
      Width           =   885
   End
   Begin VB.PictureBox picShipSelect 
      AutoRedraw      =   -1  'True
      Height          =   885
      Index           =   0
      Left            =   1620
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   960
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Finished - Close"
      Height          =   420
      Left            =   5985
      TabIndex        =   1
      Top             =   735
      Width           =   1830
   End
   Begin VB.CommandButton cmdAutoSelect 
      Caption         =   "Auto-Select"
      Default         =   -1  'True
      Height          =   420
      Left            =   5955
      TabIndex        =   0
      Top             =   120
      Width           =   1830
   End
   Begin VB.VScrollBar vscShip 
      Height          =   1140
      Index           =   3
      Left            =   1305
      Max             =   5
      Min             =   1
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4680
      Value           =   1
      Width           =   225
   End
   Begin VB.VScrollBar vscShip 
      Height          =   1140
      Index           =   2
      Left            =   1305
      Max             =   5
      Min             =   1
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3330
      Value           =   1
      Width           =   225
   End
   Begin VB.VScrollBar vscShip 
      Height          =   1140
      Index           =   1
      Left            =   1305
      Max             =   3
      Min             =   1
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1995
      Value           =   1
      Width           =   225
   End
   Begin VB.VScrollBar vscShip 
      Height          =   1140
      Index           =   0
      Left            =   1305
      Max             =   2
      Min             =   1
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   660
      Value           =   1
      Width           =   225
   End
   Begin VB.PictureBox picShip 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C000&
      DragMode        =   1  'Automatic
      Height          =   1140
      Index           =   3
      Left            =   150
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   72
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1140
   End
   Begin VB.PictureBox picShip 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      DragMode        =   1  'Automatic
      Height          =   1140
      Index           =   2
      Left            =   150
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   72
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3330
      Width           =   1140
   End
   Begin VB.PictureBox picShip 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      DragMode        =   1  'Automatic
      Height          =   1140
      Index           =   1
      Left            =   150
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   72
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1995
      Width           =   1140
   End
   Begin VB.PictureBox picShip 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      DragMode        =   1  'Automatic
      Height          =   1140
      Index           =   0
      Left            =   150
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   72
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   660
      Width           =   1140
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   13
      Left            =   5265
      Shape           =   3  'Circle
      Top             =   1425
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   12
      Left            =   4830
      Shape           =   3  'Circle
      Top             =   1425
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   11
      Left            =   4380
      Shape           =   3  'Circle
      Top             =   1425
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   10
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   1425
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   9
      Left            =   3540
      Shape           =   3  'Circle
      Top             =   1425
      Width           =   300
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFF80&
      Height          =   525
      Index           =   3
      Left            =   3405
      Shape           =   4  'Rounded Rectangle
      Top             =   1305
      Width           =   2355
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   8
      Left            =   5295
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   7
      Left            =   4845
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   6
      Left            =   4425
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   5
      Left            =   4005
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   300
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFF80&
      Height          =   525
      Index           =   2
      Left            =   3870
      Shape           =   4  'Rounded Rectangle
      Top             =   2625
      Width           =   1890
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   4
      Left            =   5295
      Shape           =   3  'Circle
      Top             =   5325
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   3
      Left            =   4875
      Shape           =   3  'Circle
      Top             =   5325
      Width           =   300
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFF80&
      Height          =   525
      Index           =   1
      Left            =   4740
      Shape           =   4  'Rounded Rectangle
      Top             =   5205
      Width           =   1020
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   2
      Left            =   5385
      Shape           =   3  'Circle
      Top             =   4110
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   1
      Left            =   4980
      Shape           =   3  'Circle
      Top             =   4110
      Width           =   300
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Index           =   0
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4110
      Width           =   300
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFF80&
      Height          =   525
      Index           =   0
      Left            =   4425
      Shape           =   4  'Rounded Rectangle
      Top             =   3990
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select 3 - This class has 2 sections that must be destroyed"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1635
      TabIndex        =   27
      Top             =   4740
      Width           =   4380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select 3 - This class has 3 sections that must be destroyed"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1650
      TabIndex        =   26
      Top             =   3390
      Width           =   4380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select 1 - This class has 4 sections that must be destroyed"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1635
      TabIndex        =   25
      Top             =   2055
      Width           =   4380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select 1 - This class has 5 sections that must be destroyed"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1665
      TabIndex        =   24
      Top             =   720
      Width           =   4320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSelection.frx":013E
      ForeColor       =   &H0080FFFF&
      Height          =   405
      Left            =   150
      TabIndex        =   14
      Tag             =   "Select your ships below by dragging and dropping them onto the empty squares."
      Top             =   30
      Width           =   6000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Battlehsip Class"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   165
      TabIndex        =   13
      Top             =   465
      Width           =   1350
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Figher Class"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   12
      Top             =   4485
      Width           =   1350
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cruiser Class"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   11
      Top             =   3135
      Width           =   1350
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Destroyer Class"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   10
      Top             =   1785
      Width           =   1350
   End
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dRect As RECT, bmpRect As RECT
Private sourceRect() As RECT

Private Sub cmdAutoSelect_Click()
Dim iSelect As Integer, Looper As Integer, I As Integer, iShip As Integer
Dim i3Holes As Integer, i2Holes As Integer
' function simply chooses ships, shields and cloaking device randomly
If GameMode Mod 2 Then  ' 8 ship games
    i3Holes = 3
    i2Holes = 3
Else                    ' 5 ship games
    i3Holes = 2
    i2Holes = 1
End If
Randomize Timer
For Looper = 0 To 3
    ' loop thru each ship category
    For I = 1 To Choose(Looper + 1, 1, 1, i3Holes, i2Holes)
        ' loop thru each destination picBox and place a ship there
        iShip = Choose(Looper + 1, 0, 1, 2, 5) + I - 1  ' destination
        iSelect = Int(Rnd * vscShip(Looper).Max + 1)    ' ship to select
        ' show the ship in the sample picBox
        If Val(vscShip(Looper).Tag) <> iSelect Then vscShip(Looper).Value = iSelect
        ' call function which will eventually draw the ship
        Call picShipSelect_DragDrop(iShip, picShip(Looper), 0, 0)
    Next
Next
If GameMode < 3 Then        ' Shield games
    Dim sRandom As String
    ' give larger ships an advantage when placing shields
    If GameMode Mod 2 Then  ' 8 ship game
        sRandom = "0123450123401601234701"
    Else                    ' 5 ship game
        sRandom = "01235012301012301"
    End If
    For Looper = 0 To 2
        picShield(Looper).Tag = ""
    Next
    For Looper = 0 To 2
        ' loop thru each shield
        iSelect = Int(Rnd * Len(sRandom) + 1)       ' ship to shield
        iShip = Val(Mid$(sRandom, iSelect, 1))
        sRandom = Replace$(sRandom, CStr(iShip), "") ' remove ship from string
        ' call function which will eventually draw the ship
        Call picShipSelect_DragDrop(iShip, picShield(Looper), 0, 0)
    Next
    ' now we do the same for the cloaking device
    ' give smaller ships an advantage when cloaking
    If GameMode Mod 2 Then              ' 8 ship games
        sRandom = "567256735674567"
    Else                                ' 5 ship games
        sRandom = "525355"
    End If
    iSelect = Int(Rnd * Len(sRandom) + 1)   ' ship to cloak
    iShip = Val(Mid$(sRandom, iSelect, 1))
    ' call function which will eventually draw the ship
    picCloak.Tag = ""
    Call picShipSelect_DragDrop(iShip, picCloak, 0, 0)
End If
End Sub

Private Sub Command1_Click()
' The save function
Dim Looper As Integer
' ensure all ships have been selected. If so their Tag property won't be empty
For Looper = picShipSelect.LBound To picShipSelect.UBound
    If picShipSelect(Looper).Visible = True And picShipSelect(Looper).Tag = "" Then
        MsgBox "You haven't selected all of your ships. Try again.", vbInformation + vbOKOnly
        Exit Sub
    End If
Next
If GameMode < 3 Then        ' games with shields
    ' ensure sheilds and cloak have been placed
    For Looper = picShield.LBound To picShield.UBound
        If picShield(Looper).Tag = "" Then
            MsgBox "You haven't placed all of your shields. Try again.", vbInformation + vbOKOnly
            Exit Sub
        End If
    Next
    If picCloak.Tag = "" Then
        MsgBox "You haven't placed your cloaking device.", vbInformation + vbOKOnly
        Exit Sub
    End If
End If
' now we simply populate the coords in the GIF for the ship selected
' and also assign the backcolor of the ship
For Looper = 0 To UBound(Player(PlayerID).ID)
    Player(PlayerID).BMPxy(Looper) = sourceRect(Player(PlayerID).ID(Looper))
    Player(PlayerID).Size(Looper) = Choose(Player(PlayerID).ID(Looper) + 1, 5, 5, 4, 4, 4, 3, 3, 3, 3, 3, 2, 2, 2, 2, 2)
Next
GP = 1          ' flag indicating all choices selected
Unload Me
End Sub

Private Sub Form_Load()
' initial display of the form
Caption = Caption & Player(PlayerID).Name
' function loads the coords in the GIF for each ship
LoadShipCoords sourceRect
' set the backcolors here. They can be changed to whatever we want
picShip(0).BackColor = vbRed
    picShipSelect(0).BackColor = vbRed
picShip(1).BackColor = vbBlue
    picShipSelect(1).BackColor = vbBlue
picShip(2).BackColor = vbYellow
    picShipSelect(2).BackColor = vbYellow
    picShipSelect(3).BackColor = vbYellow
    picShipSelect(4).BackColor = vbYellow
picShip(3).BackColor = &HC0C000
    picShipSelect(5).BackColor = &HC0C000
    picShipSelect(6).BackColor = &HC0C000
    picShipSelect(7).BackColor = &HC0C000
Dim Looper As Integer
' load the 1st ship for each of the 4 categories
For Looper = 0 To 3
    Call vscShip_Change(Looper)
Next
If GameMode < 3 Then        ' games with shields
    ' display the 3 blue shields and the cloak
    For Looper = 0 To 2
        CreateSample picShield(Looper)
    Next
    CreateSample picCloak
Else                        ' games without shields
    ' we hide & disable the shields and cloak
    For Looper = 0 To 2
        picShield(Looper).Enabled = False
        picShield(Looper).Visible = False
    Next
    picCloak.Enabled = False
    picCloak.Visible = False
    Label3.Caption = Label3.Tag
End If
If GameMode Mod 2 = 0 Then      ' 8 ship games
    ' update the labels indicating how many ships to select & hide extra destination picBoxes
    For Looper = 3 To 4
        Label1(Looper - 1).Caption = Replace$(Label1(Looper - 1).Caption, "Select " & Choose(Looper - 2, 3, 3), "Select " & Choose(Looper - 2, 2, 1))
        picShipSelect(Looper + 3).Enabled = False
        picShipSelect(Looper + 3).Visible = False
    Next
    picShipSelect(4).Enabled = False
    picShipSelect(4).Visible = False
End If
Select Case GameMode
    Case Is < 3 ' do nothing
    Case Is < 5 ' show the plain jane instructions (no shields)
        frameMode(0).Enabled = False
        frameMode(1).ZOrder
        frameMode(1).Enabled = True
    Case Else   ' show the salvo instructions
        frameMode(0).Enabled = False
        frameMode(2).ZOrder
        frameMode(2).Enabled = True
End Select

If PlayerID = 2 And frmBoard.optOpponent(1) = True Then
    ' computer is playing & it's the computer's turn to select
    Call cmdAutoSelect_Click
    Call Command1_Click
Else
    ' if ships were already loaded, display the player's selections
    If frmBoard.picPlacement.UBound Then LoadPreviousSelections
    GP = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase sourceRect
End Sub

Private Sub picShipSelect_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
' Assigns ship information to the Player() array

' Since there are several choices for each category of ships, we need afew
' indexes to align things up.
' Battleship Class: 2 ship choices, 1 selection
' Destroyer Class: 3 ship choices, 1 selection
' Cruiser Class: 5 ship choices, 3 selections (8 ship game), else 2 selections
' Fighter Class: 5 ship choices, 3 selections (8 ship game), else 1 selection
Dim imgIdxOffset As Integer, I As Integer, newIndex As Integer
If GameMode Mod 2 Then  ' 8 game ships
    newIndex = Index
Else                    ' 5 game ships
    newIndex = Choose(Index + 1, 0, 1, 2, 3, 3, 4, 4, 4)
End If
' All objects are dropped on the picShipSelect picBoxes
Select Case Source.Name
    ' making a ship choice
Case "picShip"
    ' ensure the wrong class of ships aren't dropped here
    Select Case Index
    Case 0  ' Battleship
        If Source.Index > 0 Then Exit Sub
    Case 1  ' Destroyer
        If Source.Index <> 1 Then Exit Sub
    Case 2, 3, 4    ' Cruiser
        If Source.Index <> 2 Then Exit Sub
    Case Else       ' Fighter
        If Source.Index <> 3 Then Exit Sub
    End Select
    Dim iStart As Integer
    picShipSelect(Index).Tag = "Placed" ' indicate ship was placed
    ' this offset aligns the ship selection's Scroll bar index to a
    ' Player().ID reference which may be Ubound at 7 or 4
    iStart = Choose(Index + 1, -1, 1, 4, 4, 4, 9, 9, 9)
    Player(PlayerID).ID(newIndex) = iStart + vscShip(Source.Index).Value
    ' Create the dropped ship & update the backcolor for the ship
    CreateSample picShipSelect(Index), CLng(Player(PlayerID).ID(newIndex))
    Player(PlayerID).Color(newIndex) = picShipSelect(Index).BackColor
Case "picShield"
    For I = 0 To 2
        If I <> Source.Index Then
            If Index + 1 = Val(picShield(I).Tag) Then
                MsgBox "That ship is already shielded. Choose another ship to shield.", vbInformation + vbOKOnly
                Exit Sub
            End If
        End If
    Next
    Source.Tag = Index + 1 ' indicate shield placed
    'align shields in bottom left corner of the selection
    Source.Left = picShipSelect(Index).Left
    Source.Top = picShipSelect(Index).Top + picShipSelect(Index).Height - Source.Height
    ' update which ship was shielded
    Player(PlayerID).Shield(Source.Index + 1) = newIndex
Case "picCloak"
    Source.Tag = "Placed"   ' indicate cloak was placed
    ' align cloak in bottom right corner of the selection
    Source.Left = picShipSelect(Index).Width - Source.Width + picShipSelect(Index).Left
    Source.Top = picShipSelect(Index).Top + picShipSelect(Index).Height - Source.Height
    ' update which ship was cloaked
    Player(PlayerID).Cloak = newIndex
End Select

End Sub

Private Sub vscShip_Change(Index As Integer)
' changes choice of which ship to select
If vscShip(Index).Value = Val(vscShip(Index).Tag) Then Exit Sub
Dim iStart As Integer
iStart = Choose(Index + 1, -1, 1, 4, 9)
CreateSample picShip(Index), (iStart + vscShip(Index).Value)
vscShip(Index).Tag = vscShip(Index).Value
End Sub

Private Sub CreateSample(ImageID As PictureBox, Optional Index As Long)

Dim bmpX As Long, bmpY As Long, NewX As Long, NewY As Long
Dim bmpHandle As Long

Select Case ImageID.Name
Case "picShip", "picShipSelect"
    ' drawing ships
    bmpX = sourceRect(Index).Right - sourceRect(Index).Left
    bmpY = sourceRect(Index).Bottom - sourceRect(Index).Top
    bmpRect = MakeRectangle(0, sourceRect(Index).Top, bmpX, bmpY)
    bmpHandle = sourceBmp.Handle
Case "picShield"
    ' drawing shields
    bmpRect = GridItems(0)
    bmpX = GridItems(0).Right - GridItems(0).Left
    bmpY = GridItems(0).Bottom - GridItems(0).Top
    bmpHandle = miscBmp.Handle
Case "picCloak"
    ' drawing cloaking device
    bmpRect = GridItems(5)
    bmpX = GridItems(5).Right - GridItems(5).Left
    bmpY = GridItems(5).Bottom - GridItems(5).Top
    bmpHandle = miscBmp.Handle
End Select
With ImageID
    CalculateRatio NewX, NewY, (.Width - 2), (.Height - 2), bmpX, bmpY
    dRect = MakeRectangle(((.Width - 2) - NewX) / 2, ((.Height - 2) - NewY) / 2, .Width, .Height)
    .Cls
    DrawTransparentBitmap .hdc, dRect, bmpHandle, bmpRect, -1, NewX, NewY
End With
DoEvents
End Sub

Private Sub LoadPreviousSelections()
Dim Looper As Integer, iOffset As Integer, scrollOffset As Integer
Dim imgIdx As Integer, destOffset As Integer
' When player wants to change ships before a game starts, we need to
' display which ships were already selected and which were shielded
' or cloaked, if needed
For Looper = 0 To UBound(Player(PlayerID).ID)
    If GameMode Mod 2 Then      ' 8 ship games
        iOffset = Choose(Looper + 1, -1, 1, 4, 4, 4, 9, 9, 9)
        scrollOffset = Choose(Looper + 1, 0, 1, 2, 2, 2, 3, 3, 3)
        imgIdx = Choose(Looper + 1, 0, 1, 2, 2, 2, 3, 3, 3)
        destOffset = Looper
    Else                        ' 5 ship games
        iOffset = Choose(Looper + 1, -1, 1, 4, 4, 9)
        scrollOffset = Choose(Looper + 1, 0, 1, 2, 2, 3)
        imgIdx = Choose(Looper + 1, 0, 1, 2, 2, 3)
        destOffset = Choose(Looper + 1, 0, 1, 2, 3, 5)
    End If
    ' set scroll value to ship value to display the ship
    vscShip(scrollOffset).Value = Player(PlayerID).ID(Looper) - iOffset
    ' draw the sample ship
    Call picShipSelect_DragDrop(destOffset, picShip(imgIdx), 0, 0)
Next
If GameMode < 3 Then    ' shielded and cloaked games
    ' we drag & drop the cloak on the appropriate ship
    If GameMode Mod 2 = 0 Then
        destOffset = Choose(Player(PlayerID).Cloak + 1, 0, 0, 0, 0, 1)
    Else
        destOffset = 0
    End If
    Call picShipSelect_DragDrop(Player(PlayerID).Cloak + destOffset, picCloak, 0, 0)
    ' now we drag & drop the 3 shields on the appropriate ships
    For Looper = 1 To 3
        If GameMode Mod 2 Then
            destOffset = 0
        Else
            destOffset = Choose(Player(PlayerID).Shield(Looper) + 1, 0, 0, 0, 0, 1)
        End If
        Call picShipSelect_DragDrop(Player(PlayerID).Shield(Looper) + destOffset, picShield(Looper - 1), 0, 0)
    Next
End If
End Sub
