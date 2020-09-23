VERSION 5.00
Begin VB.Form frmBoard 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Starship Skirmish"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10365
   Icon            =   "frmBoard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   691
   Begin VB.Timer TimerScanner 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1200
      Top             =   195
   End
   Begin VB.Timer TimerAnimation 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   765
      Top             =   180
   End
   Begin VB.PictureBox picAniShip 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000D&
      Height          =   690
      Left            =   4875
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   54
      Top             =   240
      Width           =   675
   End
   Begin VB.CommandButton cmdNewGame 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New Game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   135
      Width           =   2730
   End
   Begin VB.Timer TimerMain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   315
      Top             =   180
   End
   Begin VB.PictureBox picLegend 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   1
      Top             =   7770
      Width           =   10365
      Begin VB.Label lblLegend 
         Caption         =   "Missed Shot"
         Height          =   420
         Index           =   2
         Left            =   6705
         TabIndex        =   27
         Top             =   15
         Width           =   765
      End
      Begin VB.Label lblLegend 
         Caption         =   "Section has been destroyed"
         Height          =   420
         Index           =   1
         Left            =   4800
         TabIndex        =   26
         Top             =   15
         Width           =   1395
      End
      Begin VB.Label lblLegend 
         Caption         =   "Shield Strength:  Green = Strong Yellow = Good    Red = Weak"
         Height          =   420
         Index           =   0
         Left            =   1665
         TabIndex        =   25
         Top             =   15
         Width           =   2400
      End
      Begin VB.Label lblGridID 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7725
         TabIndex        =   23
         Top             =   105
         Width           =   75
      End
   End
   Begin VB.PictureBox picField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   4710
      Left            =   255
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   314
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   5355
      Begin VB.PictureBox picPlacement 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   0
         Left            =   3270
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2940
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.PictureBox picDrawingBoard 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   360
         Left            =   660
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   24
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   855
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Shape shpGrid 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         Height          =   435
         Left            =   900
         Shape           =   4  'Rounded Rectangle
         Top             =   2340
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape bdrScan 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H00FF80FF&
         BorderStyle     =   6  'Inside Solid
         Height          =   390
         Left            =   3075
         Shape           =   4  'Rounded Rectangle
         Top             =   2715
         Visible         =   0   'False
         Width           =   390
      End
   End
   Begin VB.Frame frameStep 
      BackColor       =   &H00000000&
      Caption         =   "GAME OVER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   6225
      Index           =   2
      Left            =   5700
      TabIndex        =   56
      Top             =   675
      Width           =   3915
      Begin VB.OptionButton optShow 
         BackColor       =   &H00C0C000&
         Caption         =   "View Ship Placement for Player #2"
         Height          =   465
         Index           =   2
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   5130
         Width           =   3240
      End
      Begin VB.OptionButton optShow 
         BackColor       =   &H00C0C000&
         Caption         =   "View Ship Placement for Player #1"
         Height          =   465
         Index           =   1
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   4185
         Width           =   3240
      End
      Begin VB.Label lblGameOver 
         BackColor       =   &H00000000&
         Caption         =   "Better luck next time. The computer was just a little bit too good this game.  Very good match!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   1
         Left            =   360
         TabIndex        =   59
         Top             =   420
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.Label lblGameOver 
         BackColor       =   &H00000000&
         Caption         =   $"frmBoard.frx":030A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1005
         Index           =   2
         Left            =   270
         TabIndex        =   58
         Top             =   1620
         Width           =   3465
      End
      Begin VB.Label lblGameOver 
         BackColor       =   &H00000000&
         Caption         =   "To view where each player positioned their ships, click on the buttons below."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   750
         Index           =   3
         Left            =   270
         TabIndex        =   57
         Top             =   3270
         Width           =   3495
      End
      Begin VB.Label lblGameOver 
         BackStyle       =   0  'Transparent
         Caption         =   "Congratulations you won the game!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   0
         Left            =   360
         TabIndex        =   60
         Tag             =   "Congratulations #, you won the game!"
         Top             =   420
         Width           =   3165
      End
   End
   Begin VB.Frame frameStep 
      BackColor       =   &H00000000&
      Caption         =   "Player Name"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   6225
      Index           =   1
      Left            =   5700
      TabIndex        =   42
      Top             =   675
      Width           =   3915
      Begin VB.Frame frameSalvo 
         BackColor       =   &H00000000&
         Caption         =   "Salvo Instructions"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   3735
         Left            =   135
         TabIndex        =   63
         Top             =   2415
         Visible         =   0   'False
         Width           =   3645
         Begin VB.CommandButton cmdFireSalvo 
            BackColor       =   &H00FFFF00&
            Caption         =   "Click to Fire the Salvo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   3135
            Width           =   3435
         End
         Begin VB.PictureBox picSalvo 
            BackColor       =   &H00C0C000&
            Height          =   1005
            Left            =   90
            ScaleHeight     =   63
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   226
            TabIndex        =   66
            Top             =   1995
            Width           =   3450
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of shots you have available..."
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   135
            TabIndex        =   67
            Top             =   1785
            Width           =   3330
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "To move a shot, click on the one you want to move and it will disappear.  Click on the grid where you want that shot moved."
            ForeColor       =   &H00FFFFFF&
            Height          =   720
            Index           =   1
            Left            =   165
            TabIndex        =   65
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "You get one shot for each undestroyed ship you have.  Simply click on the grid where you want you shot placed.  "
            ForeColor       =   &H00FFFFFF&
            Height          =   720
            Index           =   0
            Left            =   150
            TabIndex        =   64
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.PictureBox picScanner 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1050
         Left            =   2700
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   70
         TabIndex        =   55
         Top             =   5010
         Width           =   1050
      End
      Begin VB.PictureBox picEnemyStat 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   1565
         Left            =   435
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   202
         TabIndex        =   43
         Top             =   585
         Width           =   3090
      End
      Begin VB.Shape shpAmmo 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   3
         Left            =   3375
         Shape           =   3  'Circle
         Top             =   3060
         Width           =   345
      End
      Begin VB.Image imgScan 
         DragMode        =   1  'Automatic
         Height          =   405
         Index           =   1
         Left            =   2115
         OLEDragMode     =   1  'Automatic
         Top             =   5625
         Width           =   405
      End
      Begin VB.Image imgScan 
         DragMode        =   1  'Automatic
         Height          =   405
         Index           =   0
         Left            =   2115
         OLEDragMode     =   1  'Automatic
         Top             =   5085
         Width           =   405
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "When needed drag && drop the scanning device on 1 of your enemy ships above."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   915
         Index           =   9
         Left            =   270
         TabIndex        =   53
         ToolTipText     =   "Causes 3 point of damage"
         Top             =   5145
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Scanning Devices"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Index           =   8
         Left            =   330
         TabIndex        =   52
         Top             =   4860
         Width           =   1515
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1185
         Index           =   2
         Left            =   150
         Top             =   4950
         Width           =   3630
      End
      Begin VB.Label lblAmmo 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   345
         TabIndex        =   51
         ToolTipText     =   "Causes 5 point of damage"
         Top             =   4365
         Width           =   2970
      End
      Begin VB.Label lblAmmo 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   345
         TabIndex        =   50
         ToolTipText     =   "Causes 3 points of damage"
         Top             =   3720
         Width           =   2970
      End
      Begin VB.Label lblAmmo 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   345
         TabIndex        =   49
         ToolTipText     =   "Causes 1 point of damage"
         Top             =   3060
         Width           =   2970
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Photon Torpedo  - Click to Use"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   7
         Left            =   360
         TabIndex        =   48
         ToolTipText     =   "Causes 5 point of damage"
         Top             =   4095
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Pulse Cannons - Click to Use"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   6
         Left            =   360
         TabIndex        =   47
         ToolTipText     =   "Causes 3 points of damage"
         Top             =   3435
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Laser Blasters - Click to Use"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   5
         Left            =   360
         TabIndex        =   46
         ToolTipText     =   "Causes 1 point of damage"
         Top             =   2790
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   " Available Ammunition"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Index           =   4
         Left            =   330
         TabIndex        =   45
         Top             =   2460
         Width           =   1875
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   2175
         Index           =   1
         Left            =   150
         Top             =   2565
         Width           =   3630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   " Enemy Fleet Status "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Index           =   3
         Left            =   330
         TabIndex        =   44
         Top             =   315
         Width           =   1710
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1875
         Index           =   0
         Left            =   150
         Top             =   390
         Width           =   3630
      End
      Begin VB.Shape shpAmmo 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   4  'Upward Diagonal
         Height          =   300
         Index           =   0
         Left            =   345
         Top             =   3060
         Width           =   15
      End
      Begin VB.Shape shpAmmo 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFC0&
         FillStyle       =   4  'Upward Diagonal
         Height          =   300
         Index           =   1
         Left            =   345
         Top             =   3720
         Width           =   1725
      End
      Begin VB.Shape shpAmmo 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFFF&
         FillStyle       =   4  'Upward Diagonal
         Height          =   300
         Index           =   2
         Left            =   345
         Top             =   4365
         Width           =   15
      End
   End
   Begin VB.Frame frameStep 
      BackColor       =   &H00000000&
      Caption         =   "Game Setup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   6225
      Index           =   0
      Left            =   5700
      TabIndex        =   28
      Top             =   675
      Width           =   3915
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Let's Play"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4665
         Width           =   2970
      End
      Begin VB.CommandButton cmdAutoPlace 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Automatically Place Ships"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4245
         Width           =   2970
      End
      Begin VB.CommandButton cmdShipSelect 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select Your Ships"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3825
         Width           =   2970
      End
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Finished. Let Player #2 Go"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2880
         Width           =   2970
      End
      Begin VB.CommandButton cmdAutoPlace 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Automatically Place Ships"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2460
         Width           =   2970
      End
      Begin VB.CommandButton cmdShipSelect 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select Your Ships"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2040
         Width           =   2970
      End
      Begin VB.CommandButton cmdNames 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter your names"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1080
         Width           =   2970
      End
      Begin VB.OptionButton optOpponent 
         BackColor       =   &H00000000&
         Caption         =   "The Computer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   2175
         TabIndex        =   30
         Top             =   585
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.OptionButton optOpponent 
         BackColor       =   &H00000000&
         Caption         =   "Another Person"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   315
         TabIndex        =   29
         Top             =   585
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   " Player #2 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   37
         Top             =   3495
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   1590
         Index           =   1
         Left            =   315
         Top             =   3600
         Width           =   3315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   " Player #1 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Index           =   1
         Left            =   435
         TabIndex        =   33
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   1590
         Index           =   0
         Left            =   330
         Top             =   1785
         Width           =   3315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Your Opponent is..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   31
         Top             =   285
         Width           =   3555
      End
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   19
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   18
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblGrid 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape bdrField 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   390
      Left            =   5115
      Top             =   240
      Width           =   420
   End
   Begin VB.Menu mnuShipRotate 
      Caption         =   "popupShipRotate"
      Visible         =   0   'False
      Begin VB.Menu mnuFlipShip 
         Caption         =   "Flip Ship"
         Index           =   0
      End
      Begin VB.Menu mnuFlipShip 
         Caption         =   "Cancel"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit      Alt+F4"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Game Type"
      Index           =   1
      Begin VB.Menu mnuType 
         Caption         =   "With Shields and Cloaking Devices"
         Index           =   0
         Begin VB.Menu mnuShields 
            Caption         =   "Play with 8 Ships"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuShields 
            Caption         =   "Play with 5 Ships"
            Index           =   1
         End
      End
      Begin VB.Menu mnuType 
         Caption         =   "Without Shields or Cloaking Devices"
         Index           =   1
         Begin VB.Menu mnuNoShields 
            Caption         =   "Play with 8 Ships"
            Index           =   0
         End
         Begin VB.Menu mnuNoShields 
            Caption         =   "Play with 5 Ships"
            Index           =   1
         End
      End
      Begin VB.Menu mnuType 
         Caption         =   "Salvo - Speed Play"
         Index           =   2
         Begin VB.Menu mnuSalvo 
            Caption         =   "Play with 8 Ships"
            Index           =   0
         End
         Begin VB.Menu mnuSalvo 
            Caption         =   "Play with 5 Ships"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "New Game"
      Index           =   2
      Begin VB.Menu mnuNewGame 
         Caption         =   "Play Against another Person"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuNewGame 
         Caption         =   "Play Against the Computer"
         Index           =   1
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Options"
      Index           =   3
      Begin VB.Menu mnuOpts 
         Caption         =   "Sound"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOpts 
         Caption         =   "Missed Shot Color"
         Index           =   1
         Begin VB.Menu mnuMissed 
            Caption         =   "Black"
            Index           =   0
         End
         Begin VB.Menu mnuMissed 
            Caption         =   "White"
            Index           =   1
         End
         Begin VB.Menu mnuMissed 
            Caption         =   "Red"
            Index           =   2
         End
         Begin VB.Menu mnuMissed 
            Caption         =   "Green"
            Index           =   3
         End
         Begin VB.Menu mnuMissed 
            Caption         =   "Magenta"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Help"
      Index           =   4
      Begin VB.Menu mnuHelp 
         Caption         =   "Help Contents   F1"
      End
   End
End
Attribute VB_Name = "frmBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type ComputerPlayerData ' used when playing against computer
    ActiveHits As String        ' string of grids of current hits
    LastHit As Integer          ' last successful hit if any
    Pattern As Integer          ' 1=horizontal pattern, -1=vertical
    MaxShipSize As Byte         ' largest ship still on the board
    ShotsTaken As Integer       ' number shots taken--used to determine when to scan
    SalvoTurns As String
    Index As Integer
End Type
Private HullStrength() As Byte      ' used to help determine what ammo to use
Private PCdata() As ComputerPlayerData
Private bComputersTurn As Boolean       ' flag indicating computer's turn to play
Private bShowSelectedGrid As Boolean    ' flag sets grid hiliter on/off
Private ShipAnchor As PictureBox        ' used when positioning ships
Private PlayerAnimation(1 To 2) As New clsAnimator  ' ship animation
Private ScannerAnimation As New clsAnimator         ' radar animation
Private dRect As RECT, bmpRect As RECT  ' rectangles used for drawing
Private SalvoXY(1 To 4, 0 To 8) As Integer  ' used when placing salvo charges
Private Grid() As Integer                   ' grid status for each player
Private MissedColor(0 To 1) As Long

Private Sub cmdAutoPlace_Click(Index As Integer)
' Automatically place ships on the playing field in random locations

Dim Looper As Integer, X As Integer, Y As Integer, rRect As RECT, I As Integer
Dim bIsComputer As Boolean, bIsCollision As Boolean, bAllowAdj As Boolean

' Determine whether the computer is player #2 or not
bIsComputer = (optOpponent(1) = True And Index = 1)
' Initially align all ships off the screen so they don't intersect each other
' Note: Not physically relocating the picBoxes, just their coords
For Looper = 0 To UBound(Player(PlayerID).ID)
    Player(Index + 1).Location(Looper) = _
        MakeRectangle(0, 0, Player(PlayerID).Size(Looper) * Interval, Interval)
    Player(Index + 1).Location(Looper) = GetGridCoordsXY(0, 0, Chr$(Looper + 65) & ":11", Looper)
    picPlacement(Looper).Visible = False    ' ensure each is hidden
Next
Randomize Timer
bAllowAdj = ((Int(Rnd * 100 + 1) Mod 4) = 0)
' Now we get a random X,Y coord, see if their is a collision with another
' ship, resolve that collision if needed, and physically locate the picBox
For Looper = 0 To UBound(Player(PlayerID).ID)
    Set ShipAnchor = picPlacement(Looper)   ' flag for current picBox
    bIsCollision = True                         ' assume collision
    Do While bIsCollision
        X = Int(Rnd * (10 * Interval))          ' Random x,y coord
        Y = Int(Rnd * (10 * Interval))
        I = Int(Rnd * 100) + 1                  ' Random Horizontal/Vertical
        With Player(Index + 1).Location(Looper)
            ' depending on Hor/Ver position, calculate ship dimensions
            If I Mod 2 Then                     ' Vertical
                .Bottom = Y + Player(PlayerID).Size(Looper) * Interval - 2
                .Right = X + Interval - 2
            Else                                ' Horizonal
                .Right = X + Player(PlayerID).Size(Looper) * Interval - 2
                .Bottom = Y + Interval - 2
            End If          ' offset ship size to fit within grids
            .Top = Y + 1
            .Left = X + 1
        End With
        ' Call function to ensure the ship is aligned on the grid
        Snap2Grid CSng(X), CSng(Y), Not bIsComputer
        ' Now determine if ship collides with other ships already on the board
        bIsCollision = ShipCollision(Player(Index + 1).Location(Looper), Looper)
        ' if collision is present, pick another X,Y coord & try again
        If bAllowAdj = False And GameMode > 4 And bIsCollision = False Then ' don't allow adjoining ships
            With Player(Index + 1).Location(Looper)
                dRect = MakeRectangle(.Left - Interval / 2, .Top - Interval / 2, .Right + Interval / 2, .Bottom + Interval / 2)
            End With
            If Looper Then
                For I = 0 To Looper - 1
                    bmpRect = Player(Index + 1).Location(I)
                    IntersectRect rRect, dRect, bmpRect
                    If rRect.Right > rRect.Left And rRect.Bottom > rRect.Top Then
                        bIsCollision = True
                        Exit For
                    End If
                Next
            End If
        End If
    Loop
    ' if the computer is not the current player, then we want to display
    ' the ships after positioning to allow player to manually position if wanted
    If Not bIsComputer Then
        With Player(Index + 1).Location(Looper)
            X = CInt(.Right - .Left + 2) / Interval
            Y = CInt(.Bottom - .Top + 2) / Interval
        End With
        ' Call function to graphically draw ship on picBox & ensure visible
        CreateShip picPlacement(Looper), X, Y, Looper
        picPlacement(Looper).Visible = True
    End If
Next
End Sub

Private Sub cmdFireSalvo_Click()
' button used to fire the salvo
If SalvoXY(PlayerID, 0) = 0 And bComputersTurn = False Then
    MsgBox "You haven't placed any of your mines.", vbInformation + vbOKOnly
    Exit Sub
End If
If SalvoXY(PlayerID, 0) < ShipsRemaining(PlayerID) And bComputersTurn = False Then
    ' if user didn't select all available mines, offer a second chance
    If MsgBox("You haven't placed all of your mines. Do you want to fire the salvo anyway?", vbExclamation + vbYesNo + vbDefaultButton2, "Not a Full Salvo") = vbNo Then Exit Sub
End If
Dim Looper As Integer, sTimer As Single, bDelaySound As Boolean, bNoSound As Boolean
Enabled = False
bShowSelectedGrid = False
ShowGrid False
PCdata(0).ActiveHits = ""
If SalvoXY(PlayerID, 0) > 1 Then
    BeginPlaySound "110", , True
Else
    BeginPlaySound "101", , True
End If
For Looper = 1 To SalvoXY(PlayerID, 0)
    ' for each mine fire on the grid, if TakeShot is False, game is over
    If GameMode > 4 Then PCdata(0).Index = SalvoXY(3, Looper)
    picPlacement(Looper - 1).Visible = False
    If GameMode > 4 Then bDelaySound = True
    If TakeShot(SalvoXY(PlayerID, Looper)) = False Then GoTo ReEnableForm
Next
' game is not over
' we pause for a fraction of a second and then change players
If Len(PCdata(0).ActiveHits) Then BeginPlaySound App.Path & "\Hullgone.wav"
sTimer = Timer
Do While Abs(Timer - sTimer) < 1.5
    DoEvents
Loop
If bComputersTurn = True And GameMode > 4 Then PostSalvoCleanup
bShowSelectedGrid = True
ChangePlayer
' set/release flags depending on if the computer's turn is next/now
If PlayerID = 2 And optOpponent(1) = True Then
    bComputersTurn = True
    TimerMain.Enabled = True
Else
    bComputersTurn = False
    Enabled = True
    SendMessage picField.hwnd, &H200, 0&, 0&
End If
ReEnableForm:
Enabled = True
End Sub

Private Sub cmdGo_Click(Index As Integer)
' button either starts the game (index of 1) or allows player 2 to set up
Dim Looper As Integer

TimerMain.Enabled = False
If Index = 0 Then   ' player#1 done, let player#2 go
    ' Don't allow player to have ships overlapping each other
    GoSub EnsureNoCollisions
    ' disable player#1's butttons & hide his/her ships
    cmdAutoPlace(0).Enabled = False
    cmdGo(0).Enabled = False
    cmdShipSelect(0).Enabled = False
    GoSub RemovePlaceholders
    PlayerID = 2 ' flag indicating player 2 is now active
    If optOpponent(1) = True Then   ' computer is player#2
        SelectShips 2, False        ' call function to allow ship choices
        Call cmdAutoPlace_Click(1)  ' auto place the ship
        GoSub ShowNextScreen        ' begin the game
        MsgBox "The computer has selected its ships." & vbCrLf & _
            Label1(1).Caption & ", you will go first", vbInformation + vbOKOnly, "Computer is Ready"
    Else    ' player#2 is human
        MsgBox Player(2).Name & "," & vbCrLf & "On the next screen, select your ships and click Finished when done.", vbInformation + vbOKOnly, "Ship Selection"
        ' enable player#2 ship select button & prompt for ships
        cmdShipSelect(1).Enabled = True
        If SelectShips(Index + 1, False) = False Then Exit Sub
        ' if ships were selected, enable the auto-position & go button
        cmdAutoPlace(1).Enabled = True
        cmdGo(1).Enabled = True
    End If
Else        ' player#2 done, let's begin the game
    GoSub EnsureNoCollisions   ' don't allow overlapping ships
    GoSub RemovePlaceholders   ' hide player#2's ships
    GoSub ShowNextScreen       ' start the game
    MsgBox Label1(1).Caption & ", let's begin the fight. You will go first.", vbInformation + vbOKOnly, "Start the Battle"
End If
Exit Sub

EnsureNoCollisions:
' Routine checks to make sure no ships overlap each other
For Looper = 0 To UBound(Player(PlayerID).ID)
    ' call function to check for overlapping of each ship
    If ShipCollision(Player(PlayerID).Location(Looper), Looper, True) Then
        MsgBox "Sorry. You cannot overlap any of your ships. " & vbCrLf & _
        "The pink rectangle indicates where the overlap is." & vbCrLf & vbCrLf & _
        "If you cannot see the rectangle, move this message window.", vbInformation + vbOKOnly, "Overlapped Ships"
        GP = -1     ' flag indicating to flash a pink square
        TimerMain.Interval = 600
        TimerMain.Enabled = True
        Exit Sub
    End If
Next
Return
RemovePlaceholders:
' Routine simply hides the player's positioned ships
For Looper = picPlacement.UBound To 1 Step -1
    ' we unload these 'cause picBoxes only needed during setup
    Unload picPlacement(Looper)
Next
picPlacement(Looper).Visible = False
Return
ShowNextScreen:
    GoSub RemovePlaceholders
    ' shows the normal PLAY screen
    frameStep(0).Enabled = False
    frameStep(1).Enabled = True
    frameStep(1).ZOrder
    ' ensure player ammo selections and scan info are reset
    For Looper = 1 To 2
        Player(Looper).Ammo(3) = 0
        Player(Looper).Ammo(0) = Ammo_Laser
        Player(Looper).Ammo(1) = Ammo_Cannon
        Player(Looper).Ammo(2) = Ammo_Torpedo
        Player(Looper).Scans(0) = Abs(CInt(GameMode < 3)) * 2
        Player(Looper).Scans(1) = 0
        Player(Looper).Scans(2) = 0
        Player(Looper).CloakRevealed = 0
    Next
    ' reset this picBox which will be used later in the game
    With picDrawingBoard
        .Width = picField.Width
        .Height = picField.Height
        .Cls
        .AutoRedraw = True
    End With
    If GameMode > 4 Then        ' salvo game
        ReDim PCdata(0 To UBound(Player(1).ID) + 1)
        With picPlacement(0)    ' set up template to display mines
            .Cls
            .BackColor = vbRed
            If 32 > Interval Then .Width = Interval Else .Width = 32
            .Height = .Width
            .AutoRedraw = True
            .Visible = False
            .Enabled = True
            .ToolTipText = "Double Click to remove"
        End With
        For Looper = 0 To UBound(Player(PlayerID).ID)
            If Looper Then  ' load additional picBoxes for mines
                Load picPlacement(Looper)
                picPlacement(Looper).ZOrder
                picPlacement(Looper).ToolTipText = "Double Click to remove"
            End If
            ' draw each mine on its own picBox
            dRect = MakeRectangle(0, 0, picPlacement(0).Width, picPlacement(0).Height)
            bmpRect = GridItems(6)
            DrawTransparentBitmap picPlacement(Looper).hdc, dRect, miscBmp.Handle, bmpRect, -1, picPlacement(0).Width - 4, picPlacement(0).Height - 4
        Next
    Else
        ReDim PCdata(0 To 0)
    End If
    SetStrengths                ' calculate each player's ship strengths
    ChangePlayer True           ' start with player #1
    bShowSelectedGrid = True    ' show grid selection when mouse hovers
    PCdata(0).MaxShipSize = 5 ' largest ship on the board
    PCdata(0).Index = 0
    If UBound(Player(1).ID) > 5 Then PCdata(0).SalvoTurns = "12345678" Else PCdata(0).SalvoTurns = "12345"
    ReDim HullStrength(1 To 100)   ' redimension array
    TimerMain.Interval = 1000                   ' time lag between players
    cmdNewGame.Enabled = True   ' allow New Game button
Return
End Sub

Private Sub cmdNames_Click()
' Function simply allows users to change or select their names
Dim sVal As String
' start with player#1
sVal = InputBox("Player #1. Please enter your name.", "Player #1 Name", Label1(1).Caption)
If sVal = "" Then Exit Sub
' Change captions to match player's name
    Label1(1).Caption = Left(sVal, 30)
    Label1(1).Left = Shape1(0).Left + 90
    Player(1).Name = sVal
    frameStep(1).Caption = Label1(1).Caption
If optOpponent(1) Then  ' playing against the computer
    ' just in case player#1 selected "The Computer" as their game name
    sVal = "The Computer"
    If Player(1).Name = "The Computer" Then sVal = "Admiral PC"
Else                    ' playing against another person
    ' get that player's name
    If Label1(2).Caption = "The Computer" Then Label1(2).Caption = "Player #2"
    sVal = InputBox("Player #2. Please enter your name.", "Player #2 Name", Label1(2).Caption)
    If sVal = "" Then Exit Sub
End If
Label1(2).Caption = Left(sVal, 30)
Label1(2).Left = Shape1(0).Left + 90
Player(2).Name = sVal
cmdShipSelect(0).Enabled = True ' allow ship select for player#1
If picPlacement.UBound = 0 Then
    MsgBox Player(1).Name & "," & vbCrLf & "On the next screen, select your ships and click Finished when done.", vbInformation + vbOKOnly, "Ship Selection"
    ' prompt for ship selections & enable auto-position & Go buttons if done
    Call cmdShipSelect_Click(0)
End If
cmdNewGame.Enabled = True
End Sub

Private Function SelectShips(iOpponent As Integer, bRetry As Boolean) As Boolean
' Function simply shows the ship selection form & initially places selected
' ships on the board

Dim Looper As Integer, bIsComputer As Boolean
' determine if player#2 is the computer
bIsComputer = (optOpponent(1) = True And iOpponent = 2)
' if cmdShipSelect was hit, then bRetry is true
If Not bRetry Then
    ' 1st time selecting ships, so ensure any previous selections are hidden
    For Looper = picPlacement.LBound To picPlacement.UBound
        picPlacement(Looper).Visible = False
    Next
End If
GP = 0  ' reset flag
If bIsComputer Then
    ' if playing against computer, hide the ship selection form & load it
    On Error Resume Next
    Load frmSelection
    ' the form will close automatically
    On Error GoTo 0
Else    ' show ship selection form
    frmSelection.Show 1, Me
End If
If GP <> 1 Then Exit Function   ' user cancelled out selection of ships
Dim X As Integer, rRect As RECT
' now we redraw and/or re-position ships
If bRetry Then
    ' user simply decided to change which ships they wanted
    ' but they are already positioned, so we don't reposition them
    For Looper = 0 To UBound(Player(PlayerID).ID)
        ' determine size of each ship & reset the tooltiptext
        X = Player(PlayerID).Size(Looper)
        picPlacement(Looper).ToolTipText = ""
        ' depending on hor/ver, draw the ship
        If picPlacement(Looper).Width > picPlacement(Looper).Height Then
            CreateShip picPlacement(Looper), X, 1, Looper
        Else
            CreateShip picPlacement(Looper), 1, X, Looper
        End If
    Next
    Exit Function
End If
' initial selection of ships, we need to position them initially
For Looper = 0 To UBound(Player(PlayerID).ID)
    ' determine ship size of picBoxes
    X = Player(PlayerID).Size(Looper)
    ' all ships are initially displayed horizontal
    ' if needed, load a temporary picBox to use
    If picPlacement.UBound < Looper Then Load picPlacement(Looper)
    ' initially start each ship in top left corner
    rRect = MakeRectangle(0, 0, Interval * X - 2, Interval - 2)
    Player(PlayerID).Location(Looper) = rRect
    ' now update the memory coords of the ship, relocate it, & display it if it's not the computer positioning its ships
    Player(PlayerID).Location(Looper) = GetGridCoordsXY(0, 0, Chr(Looper + 66) & ":" & Looper + 1 + (Abs(CInt(bIsComputer)) * 10), Looper)
    With picPlacement(Looper)
        .Left = Player(PlayerID).Location(Looper).Left
        .Top = Player(PlayerID).Location(Looper).Top
        ' no need to create graphical ships if the computer is positioning
        If Not bIsComputer Then
            picPlacement(Looper).ToolTipText = ""
            CreateShip picPlacement(Looper), X, 1, Looper
        End If
        .ZOrder
    End With
Next
If Not bIsComputer Then
    ' if player is human, display the ships for manual positioning
    For Looper = 0 To UBound(Player(PlayerID).ID)
        picPlacement(Looper).Visible = True
    Next
End If
SelectShips = True
End Function

Private Sub cmdNewGame_Click()
' Resets screen for a new game
TimerMain.Enabled = False       ' disable main timer (computer's turn)
bComputersTurn = False
bShowSelectedGrid = False       ' hide the grid indicator
ShowGrid False
Dim Looper As Integer, Names(1 To 2) As String
' unload any extra picBoxes
For Looper = frameStep.UBound To frameStep.LBound + 1 Step -1
    frameStep(Looper).Enabled = False
Next
' ensure the first frame is displayed & appropriate buttons disabled for now
cmdAutoPlace(0).Enabled = False: cmdAutoPlace(1).Enabled = False
cmdGo(0).Enabled = False: cmdGo(1).Enabled = False
cmdShipSelect(0).Enabled = True: cmdShipSelect(1).Enabled = False
' reset the computer data should the game be played against the computer
bdrScan.Visible = False ' hide the scanner rectangle
InitializePlayingField  ' redraw a blank playing field
' store the players names, since the array will be erased
If Player(1).Name = "" Then Names(1) = "Player #1" Else Names(1) = Player(1).Name
If Player(2).Name = "" Then Names(2) = "Player #1" Else Names(2) = Player(2).Name
' erase the array & load defaults
ReDim Player(1 To 2)
For Looper = 1 To 2
    ' we redim the player data to fit number of ships being played with
    ReDim Player(Looper).BMPxy(0 To Choose(GameMode, 7, 4, 7, 4, 7, 4))
    ReDim Player(Looper).ID(0 To Choose(GameMode, 7, 4, 7, 4, 7, 4))
    ReDim Player(Looper).Location(0 To Choose(GameMode, 7, 4, 7, 4, 7, 4))
    ReDim Player(Looper).Strength(0 To Choose(GameMode, 7, 4, 7, 4, 7, 4))
    ReDim Player(Looper).Size(0 To Choose(GameMode, 7, 4, 7, 4, 7, 4))
    ReDim Player(Looper).Color(0 To Choose(GameMode, 7, 4, 7, 4, 7, 4))
    Player(Looper).Name = Names(Looper)
Next
' finish up
Erase Names
PlayerID = 1
picDrawingBoard.Cls
MousePointer = vbDefault
End Sub

Private Sub cmdShipSelect_Click(Index As Integer)
' Button allows player to change which ships they want to play with
' during the game setup
PlayerID = Index + 1
If SelectShips(PlayerID, cmdGo(Index).Enabled) = True Then
    ' ensure these buttons are active if ships were selected
    cmdGo(Index).Enabled = True
    cmdAutoPlace(Index).Enabled = True
End If
End Sub

Private Sub Form_Load()
DownLoadImages
Dim Looper As Integer
App.HelpFile = App.Path & "\StarshipSkirmish.hlp"
ReDim Player(1 To 2)        ' initialize public array
For Looper = 1 To 2         ' initialize its sub arrays
    ReDim Player(Looper).BMPxy(0 To 7)
    ReDim Player(Looper).ID(0 To 7)
    ReDim Player(Looper).Location(0 To 7)
    ReDim Player(Looper).Strength(0 To 7)
    ReDim Player(Looper).Size(0 To 7)
    ReDim Player(Looper).Color(0 To 7)
Next
Player(1).Name = "Player #1"
Player(2).Name = "Player #2"
GameMode = 1    ' default is 8-ship with shields
' max out the size of the app window with relation to screen size
MaxSizeMe Me
' load the images containing ships & shields
Set sourceBmp = LoadPicture(App.Path & "\Ships.gif")
Set miscBmp = LoadPicture(App.Path & "\GridMisc.gif")
' load the coordinates for the image containing shields
LoadGridMiscCoords
' make the playing field as large as possible
ResizePlayingField
' draw any graphics & reset variables needed at beginning of a new game
InitializePlayingField
' show Player#1's animated ship
TimerAnimation.Enabled = True
End Sub

Private Sub ResizePlayingField()
' Function calculates size of playing field & utlmately positions all
' controls on the app window

Dim X As Single, Y As Single, lTwips As Long, hImage As IPictureDisp, sFile As String

frameStep(0).Left = (Width / Screen.TwipsPerPixelX) - frameStep(0).Width - 6
' Calculate a width/height that are evenly divisible by 10
' fudge in offsets for the game frames & space for the animated ships
X = frameStep(0).Left - 22
Y = picLegend.Top - 32
If X > Y Then X = Y
Do Until X Mod 10 = 0 And X + X / 11 < frameStep(0).Left - 8
    X = X - 1
Loop
Interval = X \ 10       ' variable indicates size of a grid square
' reposition the playing field picBox
picField.Move Left + 20, (picLegend.Top - X + 12) \ 2, X, X
' center the gam frames vertically on the playing field picBox
For Y = frameStep.LBound To frameStep.UBound
    frameStep(Y).Move frameStep(0).Left, (Abs(picField.Height - frameStep(0).Height)) / 2 + picField.Top
Next
' align the New Game button
cmdNewGame.Move frameStep(0).Left + cmdGo(0).Left / Screen.TwipsPerPixelX, frameStep(0).Top - cmdNewGame.Height - 6

' Legends............................
' - Show each shield on the legend bar
    For X = 1 To 3
        bmpRect = GridItems(X)
        dRect = MakeRectangle(lblLegend(0).Left - (X * 30), (picLegend.Height - 24) / 2, 24, 24)
        DrawTransparentBitmap picLegend.hdc, dRect, miscBmp.Handle, bmpRect, -1, 24, 24
    Next
' - Show the Skull on the legend bar
    dRect = MakeRectangle(lblLegend(1).Left - 30, 1, 0, 0)
    bmpRect = GridItems(4)
    DrawTransparentBitmap picLegend.hdc, dRect, miscBmp.Handle, bmpRect, -1, 22, 28
' Align the animated ship
X = frameStep(0).Left - ((frameStep(0).Left - (picField.Width + picField.Left) - Interval) \ 2) - Interval + 6
picAniShip.Move X, picField.Top, Interval, Interval
' Place the grid labels around the edge of the playing field
For Y = 0 To 9
    X = Y * Interval
    With lblGrid(Y) ' Alpha labels (A-J)
        .ForeColor = 65535
        .BackColor = BackColor
        .Caption = Chr$(65 + Y)
        .Left = bdrField.Left - .Width - 5
        .Top = X + ((Interval - .Height) / 2) + bdrField.Top
    End With
    With lblGrid(Y + 10)    ' Numerical labels (1-10)
        .ForeColor = 65535
        .BackColor = BackColor
        .Caption = Y + 1
        .Top = bdrField.Top - .Height - 2
        .Left = X + ((Interval - .Width) / 2) + bdrField.Left
    End With
Next
' reset drawing board for now
With picDrawingBoard
    .Picture = Nothing
    .Cls
    .Height = 1
    .Width = 1
    .AutoRedraw = False
    .AutoSize = False
End With
' Let's update the two mini-radar images
Set imgScan(0).Picture = LoadResPicture(103, vbResIcon)
Set imgScan(1).Picture = imgScan(0).Picture
' Now we initialize the two animated ships
PlayerAnimation(1).animation_Screen = picAniShip
PlayerAnimation(1).animation_File App.Path & "\aniShip1.gif", 48, 34, 1, 1
PlayerAnimation(1).ShowNextFrame 1  ' start animation on ship #1
PlayerAnimation(2).animation_Screen = picAniShip
PlayerAnimation(2).animation_File App.Path & "\aniShip2.gif", 48, 34, 1, 1
ScannerAnimation.animation_Screen = picScanner
' Initialize the animated radar
ScannerAnimation.animation_File App.Path & "\aniDish.gif", 79, 9, Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
' Size shapes used to highlight grids
shpGrid.Width = Interval
shpGrid.Height = Interval
bdrScan.Height = Interval * 4
bdrScan.Width = Interval * 4
PlayerID = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' prior to exiting, ensure timers are off
If Cancel Then Exit Sub
On Error Resume Next
ReleaseCapture
TimerMain.Enabled = False
TimerAnimation.Enabled = False
TimerMain.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Quiting, so let's dump a lot of memory items
Set sourceBmp = Nothing
Set miscBmp = Nothing
Erase PCdata
Erase HullStrength
Erase GridItems
Erase Grid
Erase SalvoXY
Set ShipAnchor = Nothing
Dim Looper As Integer
For Looper = picPlacement.UBound To 1 Step -1
    Unload picPlacement(Looper)
Next
Erase Player
Set PlayerAnimation(1) = Nothing
Set PlayerAnimation(2) = Nothing
Set ScannerAnimation = Nothing
End Sub


Private Sub lblAmmo_Click(Index As Integer)
' Ammo labels. When user clicks then update which ammo to use for the
' player's current/next turn
If bComputersTurn Then
    ' for the computer, automatically select next available if needed
    Do Until Player(2).Ammo(Index) > 0
        If Index Then Index = Index - 1 Else Player(2).Ammo(Index) = Ammo_Laser
    Loop
End If
If Index = 0 And Player(PlayerID).Ammo(0) < 2 Then Player(PlayerID).Ammo(0) = Ammo_Laser
If Player(PlayerID).Ammo(Index) = 0 Then
    ' if player is out of that ammo, notify
    MsgBox "You no longer have any more " & Choose(Index + 1, "Lasers", "Pulse Cannons", "Photon Torpedoes") & ".", vbInformation + vbOKOnly, "Out of Ammo!"
Else    ' otherwise, update
    Player(PlayerID).Ammo(3) = Index    ' flag says which is currently selected
    shpAmmo(3).Top = lblAmmo(Index).Top ' move the ammo selector
End If
DoEvents
End Sub

Private Sub mnuFile_Click()
Unload Me
End Sub

Private Sub mnuFlipShip_Click(Index As Integer)
' menu item: swaps a ship's orientation from vertical to horizontal & vice versa

If Index = 1 Then Exit Sub  ' the cancel menu item
Dim NewX As Integer, NewY As Integer
With ShipAnchor ' flag to indicate which ship is being flipped
    .Visible = False
    .Cls
    NewX = (.Height + 1) / Interval     ' calculate new ship dimensions
    NewY = (.Width + 1) / Interval
    ' graphically draw the ship
    CreateShip ShipAnchor, NewX, NewY, ShipAnchor.Index
    ' update the memory coords for the placement
    With Player(PlayerID).Location(ShipAnchor.Index)
        .Bottom = .Top + NewY * Interval - 2
        .Right = .Left + NewX * Interval - 2
    End With
    ' ensure ship correctly on grid & display it
    Snap2Grid CSng(Player(PlayerID).Location(ShipAnchor.Index).Left), CSng(Player(PlayerID).Location(ShipAnchor.Index).Top), True
    .Visible = True
End With
End Sub

Private Sub mnuMissed_Click(Index As Integer)
If mnuMissed(Index).Checked = True Then Exit Sub
Dim I As Integer, C1 As Long, C2 As Long
For I = 0 To mnuMissed.UBound
    mnuMissed(I).Checked = False
Next
mnuMissed(Index).Checked = True
C1 = Choose(Index + 1, vbBlack, vbWhite, vbRed, vbGreen, vbMagenta)
C2 = Choose(Index + 1, vbBlue, vbBlue, vbWhite, vbBlue, vbWhite)
If frameStep(0).Enabled = False Then
    MsgBox "The color change will happen when a new game is started.", vbInformation + vbOKOnly
Else
    MissedColor(0) = C1
    MissedColor(1) = C2
    With picLegend
        .FillColor = MissedColor(0)
        .FillStyle = 0
        picLegend.Circle (lblLegend(2).Left - (.Height / 2) - 3, .Height / 2 - 3), .Height / 4, MissedColor(1)
    End With
End If
ReadWriteINI "Write", App.Path & "\SSkirmish.ini", "Defaults", "Missed", CStr(C1)
ReadWriteINI "Write", App.Path & "\SSkirmish.ini", "Defaults", "MissedOutline", CStr(C2)
End Sub

Private Sub mnuNewGame_Click(Index As Integer)
' menu option to start a new game vs person or computer
Call cmdNewGame_Click
optOpponent(1) = Index
optOpponent(0) = (Index - 1)
If Not cmdNewGame.Enabled Then Call cmdNames_Click
End Sub

Private Sub mnuNoShields_Click(Index As Integer)
' menu selects these games
If mnuNoShields(Index).Checked Then Exit Sub
If frameStep(2).Enabled = False And cmdNewGame.Enabled = True Then
    If MsgBox("Changing game modes will force a new game. Continue?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
End If
mnuNoShields(Abs(Index - 1)).Checked = False
mnuNoShields(Index).Checked = True
'-------------------------------
mnuShields(0).Checked = False
mnuShields(1).Checked = False
mnuSalvo(0).Checked = False
mnuSalvo(1).Checked = False
GameMode = Index + 3
Call cmdNewGame_Click
frameSalvo.Enabled = False
frameSalvo.Visible = False
optOpponent(1).Enabled = True
End Sub

Private Sub mnuOpts_Click(Index As Integer)
If Index = 0 Then mnuOpts(Index).Checked = Not mnuOpts(Index).Checked
End Sub

Private Sub mnuSalvo_Click(Index As Integer)
' mnu selects these games
If mnuSalvo(Index).Checked Then Exit Sub
If frameStep(2).Enabled = False And cmdNewGame.Enabled = True Then
    If MsgBox("Changing game modes will force a new game. Continue?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
End If
mnuSalvo(Abs(Index - 1)).Checked = False
mnuSalvo(Index).Checked = True
'-------------------------------
mnuShields(0).Checked = False
mnuShields(1).Checked = False
mnuNoShields(0).Checked = False
mnuNoShields(1).Checked = False
GameMode = Index + 5
frameSalvo.Enabled = True
frameSalvo.Visible = True
Call cmdNewGame_Click
End Sub

Private Sub mnuShields_Click(Index As Integer)
' menu selects default games
If mnuShields(Index).Checked Then Exit Sub
If frameStep(2).Enabled = False And cmdNewGame.Enabled = True Then
    If MsgBox("Changing game modes will force a new game. Continue?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then Exit Sub
End If
mnuShields(Abs(Index - 1)).Checked = False
mnuShields(Index).Checked = True
'-------------------------------
mnuNoShields(0).Checked = False
mnuNoShields(1).Checked = False
mnuSalvo(0).Checked = False
mnuSalvo(1).Checked = False
GameMode = Index + 1
optOpponent(1).Enabled = True
Call cmdNewGame_Click
frameSalvo.Enabled = False
frameSalvo.Visible = False
End Sub

Private Sub optOpponent_Click(Index As Integer)
' when toggling between computer & person, update player #2 as needed
If optOpponent(1) = False Then
    If Label1(2).Caption = "The Computer" Or Label1(2) = "Admiral PC" Then
        Label1(2).Caption = "Player #2"
        Player(2).Name = "Player #2"
    End If
    Label1(2).Left = Shape1(0).Left + 90
Else
    If Player(PlayerID).Name = "The Computer" Then
        Label1(2).Caption = "Admiral PC"
    Else
        Label1(2).Caption = "The Computer"
    End If
    Player(2).Name = Label1(2).Caption
    Label1(2).Left = Shape1(0).Left + 90
    If cmdShipSelect(1).Enabled Then Call cmdGo_Click(0)
End If
End Sub

Private Sub optShow_Click(Index As Integer)
' shows players ship positions at end of game
If Len(optShow(1).Tag) Then Exit Sub
DoEndofGame True
End Sub

Private Sub picEnemyStat_DragDrop(Source As Control, X As Single, Y As Single)
' when player finishes dragging a ship, ensure it is properly displayed
' on the grids. First check to ensure what is being dragged is from this game
If Source.Name <> "imgScan" Then Exit Sub
If Player(PlayerID).Scans(1) Then
    MsgBox "You are currently scanning a ship. You need to sink that one first" & vbCrLf & "before trying to scan another ship.", vbInformation + vbOKOnly, "Scan In Effect"
    Exit Sub
End If
Dim Index As Integer, ShipID As Integer, sTimer As Single
If X < 0 Then
    ShipID = Abs(X) - 1     ' called by the computer
Else
    Index = Int((X - 7) / 48)
    If Y < 50 Then
        ShipID = Choose(Index + 1, 0, 2, 3, 4)
    Else
        ShipID = Choose(Index + 1, 1, 5, 6, 7)
    End If
End If
If ShipID > UBound(Player(PlayerID).ID) Then
    ' when playing with 5 ships, we check
    Beep
    Exit Sub
End If
' check to make sure not dropped on a sunk ship
If Player(Abs((PlayerID - 2) - 1)).Strength(ShipID) = 0 Then
    MsgBox "That ship has already been destroyed. Select another ship to scan for.", vbInformation + vbOKOnly, "Can't Scan that Ship"
    Exit Sub
End If
' check to see if ship dropped on already has been cloaked
If Player(PlayerID).CloakRevealed - 1 = ShipID Then
    MsgBox "That ship has already been scanned and was found to be cloaked. Select another ship.", vbInformation + vbOKOnly, "Enemy is Cloaked"
    Exit Sub
End If
Player(PlayerID).Scans(1) = ShipID + 1  ' grid where left corner of scan rectangle is
If Not Player(PlayerID).CloakRevealed Then  ' check to see if player selected the cloaked ship
    If Player(Abs((PlayerID - 2) - 1)).Cloak = ShipID Then  ' scanned cloaked ship
        Player(PlayerID).CloakRevealed = ShipID + 1 ' identify which ship is cloaked
        Player(PlayerID).Scans(1) = 0               ' reset scan rectangle top left corner
        ' send Wav telling user that ship was cloaked
        If Len(Dir(App.Path & "\Cloak" & PlayerID & ".wav")) And mnuOpts(0).Checked = True Then
            BeginPlaySound App.Path & "\Cloak" & PlayerID & ".wav", , True
        Else
            If bComputersTurn = False Then MsgBox "That ship is cloaked and cannot be scanned.", vbInformation + vbOKOnly, "Cloaked"
        End If
    End If
End If
DrawEnemyStats True ' update the enemy stats showning cloak or scan icons
' now we update the scan information
' determine which mini-radar image to disappear
If Player(PlayerID).Scans(0) = 2 Then Index = 0 Else Index = 1
imgScan(Index).Visible = False
imgScan(Index).Enabled = False
' update the number of scans left for this player
Player(PlayerID).Scans(0) = Player(PlayerID).Scans(0) - 1
If Player(PlayerID).Scans(1) = 0 Then Exit Sub
' send Wav telling player where ship can be found
If Len(Dir(App.Path & "\Scan" & PlayerID & ".wav")) And mnuOpts(0).Checked = True Then
    BeginPlaySound App.Path & "\Scan" & PlayerID & ".wav"
Else
    If bComputersTurn = False Then
        MsgBox "The ship can be found somewhere within the rectangle that will be " & _
            vbCrLf & "shown on screen after you click Ok.", vbInformation + vbOKOnly, "Scan Successful"
    End If
End If
ShowScanArea ShipID + 1     ' function to display the scan rectangle
If bComputersTurn Then
    ' here we pause to allow the Wav to finish
    sTimer = Timer
    Do While Abs(Timer - sTimer) < 1.75
        DoEvents
    Loop
End If
End Sub

Private Sub picField_Click()
' When this field is clicked, it is either to fire a shot or
' setup salvo mines

' Prevent actions when field is clicked outside an active game
If frameStep(1).Enabled = False And bComputersTurn = False Then Exit Sub
If GameMode > 4 And bComputersTurn = True Then Exit Sub
Dim X As Long, Y As Long, GridID As Integer, Looper As Integer
' just in case the grid indicator didn't appear, we send a
' mouse move message back to the program to trigger it
If shpGrid.Visible = False And bComputersTurn = False Then
    SendMessage picField.hwnd, &H200, 0&, 0&
    DoEvents
    If shpGrid.Visible = False Then
        ' if it still didn't appear, display a message
        MsgBox "Please move your mouse over the grid section you want to fire on.", vbInformation + vbOKOnly
        Exit Sub
    End If
End If
' we want a grid reference to the location of the grid indicators left/top coords
X = shpGrid.Left: Y = shpGrid.Top
GridID = ConvertGridCoord2Integer(X, Y)
' if the grid clicked on is already a destroyed ship section or a miss, then....
If Grid(Abs((PlayerID - 2) - 1), GridID) < 0 Then
    Enabled = True
    Beep
    Exit Sub
End If
' before firing hide the grid indicator, except if computer is firing
If Not bComputersTurn Then ShowGrid False
If GameMode > 4 Then        ' salvo game
    ' we want to see if grid clicked on already has a salvo mine
    ' if so, we remove it
    For Looper = 1 To SalvoXY(PlayerID, 0)
        If SalvoXY(PlayerID, Looper) = GridID Then
            RearrangeSalvos Looper - 1
            Exit Sub
        End If
    Next
    ' now we see if trying to place more mines that available
    If SalvoXY(PlayerID, 0) = ShipsRemaining(PlayerID) Then
        Beep
        Exit Sub
    End If
    ' good to go, let's place the mine
    ' get x,y coords of the Grid clicked
    ConvertGridCoord2Integer X, Y, CInt(GridID)
    ' set which grid is associated with this mine & increment mine placement count
    SalvoXY(PlayerID, SalvoXY(PlayerID, 0) + 1) = GridID
    SalvoXY(PlayerID, 0) = SalvoXY(PlayerID, 0) + 1
    ' move the mine in position
    With picPlacement(SalvoXY(PlayerID, 0) - 1)
        .Move ((Interval - .Width) \ 2) + X, ((Interval - .Height) \ 2) + Y
        .Visible = True
    End With
    RearrangeSalvos -1      ' updates graphically number of mines left to place
    If ShipsRemaining(PlayerID) = 1 Then Call cmdFireSalvo_Click
    Exit Sub
End If
TakeShot GridID
End Sub

Private Sub picField_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Basically used to display the grid indicator
If frameStep(1).Enabled = False Then Exit Sub
If (X < 0 Or X > picField.Width) Or (Y < 0 Or Y > picField.Height) Then
    ' when cursor goes off the playing field, remove selected grid border
    ShowGrid False
    lblGridID.Caption = ""
    ReleaseCapture  ' releae mouse capture
Else
    Dim sGridID As String, tRect As RECT
    ' get current grid coordinate in Alpha & Numeric format
    sGridID = GetGridCoordsString(X, Y)
    If sGridID <> lblGridID.Caption Or Len(lblGridID.Caption) = 0 Then
        ' display the grid coordinate & show the grid selection border
        lblGridID.Caption = sGridID
        lblGridID.Left = picLegend.Width - lblGridID.Width - 6
        ShowGrid True
        SetCapture picField.hwnd
    End If
End If
End Sub

Private Sub picField_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
' when player finishes dragging a ship, ensure it is properly displayed
' on the grids. First check to ensure what is being dragged is from this game
If Data.GetData(vbCFText) <> "Battleship" Then Exit Sub
Snap2Grid X, Y, True    ' function to force ship to be fully displayed on grid
End Sub

Private Sub picField_Resize()
' align the outer edge border of the playing field
bdrField.Move picField.Left, picField.Top, picField.Width + 1, picField.Height + 1
End Sub

Private Sub InitializePlayingField()
' This is called each time a new game is started

Dim X As Long, Y As Long, Looper As Integer, sFile As String
Dim hImage As IPictureDisp, NewX As Single, NewY As Single

frameStep(0).Enabled = True
frameStep(0).ZOrder
picAniShip.Top = picField.Top   ' put animated ship at top of screen
' ensure all extraneous/memory picBoxes are unloaded
' should already be done--except after playing a Salvo game
For Looper = picPlacement.UBound To picPlacement.LBound + 1 Step -1
    Unload picPlacement(Looper)
Next
' set up the only hardcopy picBox which is used as a template
picPlacement(0).Cls
picPlacement(0).Visible = False
picPlacement(0).ToolTipText = ""

' Now load and resize the playing field background
If Len(Dir(App.Path & "\StarFld.jpg")) = 0 Then
    MsgBox "There is no bacground image for this game." & vbCrLf & vbCrLf _
        & "Please read the help file, help topic: Miscellaneous", vbInformation + vbOKOnly
    picField.Cls
Else
    Set hImage = LoadPicture(App.Path & "\StarFld.jpg")
    With picField
        NewX = .Width / (.ScaleX(hImage.Width, vbHimetric, vbPixels))
        NewY = .Height / (.ScaleY(hImage.Height, vbHimetric, vbPixels))
        .Cls
        ' call function to stretch & paste image
        ResizeBMP hdc, .hdc, hImage.Handle, NewX, NewY
    End With
    ' erase the image & set the grid lines color
    Set hImage = Nothing
End If
picField.ForeColor = &H808080
' draw the grid lines
For Looper = 2 To 10
    X = (Looper - 1) * Interval
    picField.Line (X, 0)-(X, picField.Height)
    picField.Line (0, X)-(picField.Width, X)
Next
' Ensure the New Game frames is shown
For Looper = frameStep.LBound + 1 To frameStep.UBound
    frameStep(Looper).Enabled = False
Next
MissedColor(0) = Val(ReadWriteINI("Get", App.Path & "\SSkirmish.ini", "Defaults", "Missed", CStr(vbBlack)))
MissedColor(1) = Val(ReadWriteINI("Get", App.Path & "\SSkirmish.ini", "Defaults", "MissedOutline", CStr(vbBlue)))
' - Show the "Missed" peg on the legend bar
With picLegend
    .FillColor = MissedColor(0)
    .FillStyle = 0
    picLegend.Circle (lblLegend(2).Left - (.Height / 2) - 3, .Height / 2 - 3), .Height / 4, MissedColor(1)
End With
For Looper = 0 To mnuMissed.UBound
    mnuMissed(Looper).Checked = False
Next
Select Case MissedColor(0)
    Case vbBlack: mnuMissed(0).Checked = True
        mnuMissed(0).Checked = True
    Case vbWhite: mnuMissed(1).Checked = True
        mnuMissed(1).Checked = True
    Case vbRed: mnuMissed(2).Checked = True
        mnuMissed(2).Checked = True
    Case vbGreen: mnuMissed(3).Checked = True
        mnuMissed(3).Checked = True
    Case vbMagenta: mnuMissed(4).Checked = True
        mnuMissed(4).Checked = True
End Select
PlayerID = 1    ' reset to ensure 1st animated ship is shown
TimerScanner.Enabled = False    ' stop animated radar
bShowSelectedGrid = False       ' prevent grid selection border
End Sub

Private Sub ShowGrid(bVisible As Boolean)
' show the grid selection border
shpGrid.Visible = (bVisible And bShowSelectedGrid)
Dim rRect As RECT
rRect = GetGridCoordsXY(0, 0, lblGridID.Caption)
shpGrid.Move rRect.Left, rRect.Top
End Sub

Private Sub ShowScanArea(ShipID As Integer)
' show the scan selection area
Dim rRect As RECT, GridID As Integer, sTimer As Single
If ShipID Then
    ' Show scan area for the first time
    Dim X As Integer, Y As Integer, tRect As RECT, Cx As Long, Cy As Long
    ' get actual location of opponent's ship being scanned
    rRect = Player(Abs((PlayerID - 2) - 1)).Location(ShipID - 1)
    ' randomly offset the scan rectangle based on the top/left edge of the ship being scanned
    If rRect.Right - rRect.Left > rRect.Bottom - rRect.Top Then ' horizontal
        X = Int(Rnd * (Abs(5 - Player(PlayerID).Size(ShipID - 1))))
        Y = Int(Rnd * 4)
    Else
        X = Int(Rnd * 4)
        Y = Int(Rnd * (Abs(5 - Player(PlayerID).Size(ShipID - 1))))
    End If
    ' adjust rectangle with random selections
    rRect.Top = rRect.Top - (Y * Interval)
    rRect.Left = rRect.Left - (X * Interval)
    ' now we make sure the rectangle is on the playing field
    If rRect.Top < 0 Then rRect.Top = 0
    If rRect.Left < 0 Then rRect.Left = 0
    If rRect.Left + Interval * 3 > Interval * 10 Then rRect.Left = Interval * 6
    If rRect.Top + Interval * 3 > Interval * 10 Then rRect.Top = Interval * 6
    ' good, let's ensure the top/left edge is snapped to a grid
    rRect = GetGridCoordsXY(CSng(rRect.Left), CSng(rRect.Top))
    ' now we are going to animate the grid, moving it clockwise from the
    ' bottom right corner, full circle, then to the middle then to the
    ' actual location we want it displayed at
    bdrScan.Move 6 * Interval, 6 * Interval
    bdrScan.Visible = True
    For X = 1 To 9
        bdrScan.Move Choose(X, 6, 3, 0, 0, 0, 3, 6, 6, 3) * Interval, _
            Choose(X, 6, 6, 6, 3, 0, 0, 0, 3, 3) * Interval
        sTimer = Timer
        Do While Abs(Timer - sTimer) < 0.15
            DoEvents
        Loop
    Next
    bdrScan.Move rRect.Left, rRect.Top
    ' return the Grid reference & update the Scan array
    GridID = ConvertGridCoord2Integer(rRect.Left, rRect.Top)
    Player(PlayerID).Scans(2) = GridID
Else
    ' pre-exisitng scan, let's get the actual X,Y coords from the
    ' Scan array reference, then move it in place
    GridID = ConvertGridCoord2Integer(rRect.Left, rRect.Top, CLng(Player(PlayerID).Scans(2)))
    With bdrScan
        .Left = rRect.Left
        .Top = rRect.Top
        .Visible = True
    End With
End If
End Sub

Private Sub picPlacement_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Used to flip ships or drag and drop. Also used to remove a salvo mine selection
If GameMode > 4 And frameStep(1).Enabled = True Then
    ' remove the selected salvo mine & update number of mines available
    RearrangeSalvos Index
    Exit Sub
End If
If Index > 7 Then Exit Sub
' always identify which ship is being selected
Set ShipAnchor = picPlacement(Index)
If Button = vbRightButton Then  ' flip ship option
    PopupMenu mnuShipRotate
Else                            ' drag option
    picPlacement(Index).OLEDrag
End If
' if the pink collision square is visible, let's remove it now
If picPlacement.UBound > 7 Then
   Unload picPlacement(picPlacement.UBound)
   TimerMain.Enabled = False
   GP = 0
End If
End Sub

Private Sub picPlacement_OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
' Beginning of drag operation
AllowedEffects = vbDropEffectMove
Data.SetData "Battleship", vbCFText
End Sub

Private Sub Snap2Grid(X As Single, Y As Single, bPlace As Boolean)
' function ensures ships are displayed properly on the playing field
' and also ensures the edge of the ship doesn't go off the field

Dim rRect As RECT
' get the actual playing field coordinates closest to where the ship
' is now positioned
rRect = GetGridCoordsXY(X, Y, "", ShipAnchor.Index)
If rRect.Right > picField.Width Then    ' hanging off the right edge
    ' adjust the coords to ensure entire ship is on the screen
    rRect = GetGridCoordsXY(picField.Width - (rRect.Right - rRect.Left), CSng(rRect.Top), "", ShipAnchor.Index)
End If
If rRect.Bottom > picField.Height Then  ' hanging off the bottom edge
    ' adjust the coords to ensure entire ship is on the screen
    rRect = GetGridCoordsXY(CSng(rRect.Left), picField.Height - (rRect.Bottom - rRect.Top), "", ShipAnchor.Index)
End If
' return snapped to grid coords
Player(PlayerID).Location(ShipAnchor.Index) = rRect

If bPlace Then  ' option to physically reposition the picBox
    ShipAnchor.Left = rRect.Left
    ShipAnchor.Top = rRect.Top
End If

End Sub

Private Sub CreateShip(ImageID As PictureBox, SizeX As Integer, SizeY As Integer, Index As Integer)
' Function draws the ship on a picBox and is used during gameboard setup
' and displaying a sunk ship
Dim bmpX As Long, bmpY As Long
Dim NewX As Long, NewY As Long

' load the ship in memory
'Set hBitmap = LoadPicture(App.Path & "\Player" & PlayerID & ".bmp")
With ImageID
    .Cls    ' resize the physical dimensions of the picBox
    .Width = SizeX * Interval - 2
    .Height = SizeY * Interval - 2
    .AutoRedraw = True
    bmpX = Player(PlayerID).BMPxy(Index).Right - Player(PlayerID).BMPxy(Index).Left
    bmpY = Player(PlayerID).BMPxy(Index).Bottom - Player(PlayerID).BMPxy(Index).Top
    ' Call function to calculate sizes needed
    CalculateRatio NewX, NewY, SizeX * Interval - 2, SizeY * Interval - 2, bmpX, bmpY
    ' setp the destination picBox coords for placement of image
    .BackColor = Player(PlayerID).Color(Index)
    dRect = MakeRectangle((.Width - NewX) / 2, (.Height - NewY) / 2, NewX, NewY)
    bmpRect = MakeRectangle(0, Player(PlayerID).BMPxy(Index).Top, bmpX, bmpY)
    DrawTransparentBitmap .hdc, dRect, sourceBmp.Handle, bmpRect, -1, NewX, NewY
    If GameMode < 3 Then
        ' if the selected ship is "Cloaked" then display the cloak icon
        If Player(PlayerID).Cloak = Index Then
            dRect = MakeRectangle(ImageID.Width - 24, ImageID.Height - 24, 24, 24)
            bmpRect = GridItems(5)
            DrawTransparentBitmap .hdc, dRect, miscBmp.Handle, bmpRect, -1, 24, 24
            picPlacement(Index).ToolTipText = "This ship is cloaked against enemy scanning devices"
        Else
            picPlacement(Index).ToolTipText = "Right click on any ship to rotate it 90 degrees"
        End If
        ' if the ship is "Shielded" then display the shield icon
        If Player(PlayerID).Shield(1) = Index Or _
            Player(PlayerID).Shield(2) = Index Or _
                Player(PlayerID).Shield(3) = Index Then
            dRect = MakeRectangle(1, ImageID.Height - 24, 24, 24)
            bmpRect = GridItems(0)
            DrawTransparentBitmap .hdc, dRect, miscBmp.Handle, bmpRect, -1, 24, 24
            If Player(PlayerID).Cloak = Index Then
                picPlacement(Index).ToolTipText = "This ship has additional strength through its shields & is cloaked against enemy scan devices"
            Else
                picPlacement(Index).ToolTipText = "This ship has additional strength through its shields"
            End If
        End If
    Else
        picPlacement(Index).ToolTipText = "Right click on any ship to rotate it 90 degrees"
    End If
End With
DoEvents
End Sub

Private Sub SinkShip(sRect As RECT, iShipID As Integer, Optional bEndofGame As Boolean)
' Here we sink a ship and update several variables and screen displays
Dim bmpX As Long, bmpY As Long, eRect As RECT
Dim NewX As Long, NewY As Long, iOpponent As Integer
Dim GridID As Integer, sPattern As String
Dim X As Long, Y As Long

iOpponent = Abs((PlayerID - 2) - 1)     ' opponent player ID
' if ship sunk was that scanned, reset the reference to that ship
If Player(PlayerID).Scans(1) - 1 = iShipID Then Player(PlayerID).Scans(1) = 0
' let's determine the actual rectangle of the ship being sunk
bmpX = Player(iOpponent).BMPxy(iShipID).Right - Player(iOpponent).BMPxy(iShipID).Left
bmpY = Player(iOpponent).BMPxy(iShipID).Bottom - Player(iOpponent).BMPxy(iShipID).Top
' now we load a temporary picBox to be used as a background for the explosion
Load picPlacement(picPlacement.UBound + 1)
With picPlacement(picPlacement.UBound)
    .Visible = False
    .AutoRedraw = True
    .Width = sRect.Right - sRect.Left
    .Height = sRect.Bottom - sRect.Top
    ' copy the background from the playing field to the temp picBox
    BitBlt .hdc, 0, 0, .Width, .Height, picField.hdc, sRect.Left, sRect.Top, vbSrcCopy
    ' Call function to calculate sizes needed
    CalculateRatio NewX, NewY, .Width - 2, .Height - 2, bmpX, bmpY
    .BackColor = Player(PlayerID).Color(iShipID)
    ' draw the image on the temp picBox
    dRect = MakeRectangle((.Width - NewX + 1) / 2, (.Height - NewY + 1) / 2, .Width - 2, .Height - 2)
    bmpRect = MakeRectangle(0, Player(iOpponent).BMPxy(iShipID).Top, bmpX, bmpY)
    DrawTransparentBitmap .hdc, dRect, sourceBmp.Handle, bmpRect, -1, NewX, NewY
    If bEndofGame Then
        ' let's display any destroyed sections on the winner's ships
        GridID = ConvertGridCoord2Integer(Player(iOpponent).Location(iShipID).Left, Player(iOpponent).Location(iShipID).Top)
        bmpX = Player(iOpponent).Location(iShipID).Right - Player(iOpponent).Location(iShipID).Left
        bmpY = Player(iOpponent).Location(iShipID).Bottom - Player(iOpponent).Location(iShipID).Top
        If bmpX > bmpY Then NewX = 1 Else NewX = 10
        For NewY = 0 To Player(iOpponent).Size(iShipID) - 1
            If Grid(iOpponent, GridID + (NewY * NewX)) < 0 Then
                bmpRect = GridItems(4)
                If bmpX > bmpY Then
                    dRect = MakeRectangle((Interval - 24) / 2 + (Interval * NewY), (.Height - 30) / 2, 24, 30)
                Else
                    dRect = MakeRectangle((.Width - 24) / 2, (Interval - 30) / 2 + (Interval * NewY), 24, 30)
                End If
                DrawTransparentBitmap .hdc, dRect, miscBmp.Handle, bmpRect, -1, 24, 30
            End If
        Next
    End If
    ' now copy the drawn ship onto the playing field
    BitBlt picField.hdc, sRect.Left, sRect.Top, .Width, .Height, .hdc, 0, 0, vbSrcCopy
End With
' this routine can be called at the end of the game to draw opponents ships
' that were not sunk by the end of the game. If so, we skip the rest
If Not bEndofGame Then
    ' here we determine where on the EnemyStat screen we want to reverse
    ' colors to indicate which ship is being sunk
    If GameMode Mod 2 Then  ' offsets depend on 8 or 5 ship game
        eRect.Left = Choose(iShipID + 1, 0, 0, 1, 2, 3, 1, 2, 3) * 48 + 6
        eRect.Top = Choose(iShipID + 1, 1, 50, 1, 1, 1, 50, 50, 50)
    Else
        eRect.Left = Choose(iShipID + 1, 0, 0, 1, 2, 3) * 48 + 6
        eRect.Top = Choose(iShipID + 1, 1, 50, 1, 1, 1)
    End If
    eRect.Right = eRect.Left + 48
    eRect.Bottom = eRect.Top + 48
    ' invert the colors & update the screen
    InvertRect picEnemyStat.hdc, eRect
    InvalidateRect picEnemyStat.hwnd, eRect, 0
    ' now show the fire ball of an explosion
    ShowExplosion sRect
    ' since the ship on temp picBox is in full color, we redraw it with
    ' a brown background (indicating sunk)
    With picPlacement(picPlacement.UBound)
        .Cls
        .BackColor = &H65&        '&H40&
        ' draw the image on the picBox
        dRect = MakeRectangle((.Width - NewX + 1) / 2, (.Height - NewY + 1) / 2, .Width - 2, .Height - 2)
        bmpRect = MakeRectangle(0, Player(iOpponent).BMPxy(iShipID).Top, bmpX, bmpY)
        DrawTransparentBitmap .hdc, dRect, sourceBmp.Handle, bmpRect, -1, NewX, NewY
        ' now copy the brown ship to the playing field & update
        BitBlt picField.hdc, sRect.Left, sRect.Top, .Width, .Height, .hdc, 0, 0, vbSrcCopy
        InvalidateRect picField.hwnd, sRect, 0
    End With
    If bComputersTurn = True Or GameMode > 4 Then
        ' the computer's turn, so we need to set a lot of variables
        ' get a grid reference to the 1st section of the ship
        GridID = ConvertGridCoord2Integer(Player(iOpponent).Location(iShipID).Left, Player(iOpponent).Location(iShipID).Top)
        With PCdata(PCdata(0).Index)
            If bComputersTurn Then
                .LastHit = 0        ' no current hit
                .Pattern = 0        ' no current pattern
            End If
            ' now we need to remove references to the grid sections that were destroyed
            ' first determine whether the ship being sunk was horizontal or vertical
            If picPlacement(picPlacement.UBound).Width > picPlacement(picPlacement.UBound).Height Then bmpY = 1 Else bmpY = 10
            For bmpX = GridID To GridID + (bmpY * (Player(iOpponent).Size(iShipID) - 1)) Step bmpY
                ' remove those references
                .ActiveHits = Replace$(.ActiveHits, Format(bmpX, "000."), "")
                PCdata(0).ActiveHits = Replace$(PCdata(0).ActiveHits, Format(bmpX, "000."), "")
                ' update the value on the grid array to indicate sunk/missed shot
                Grid(iOpponent, bmpX) = -5
            Next
            ' if we have other hits still active after the sinking, then...
            If Len(.ActiveHits) > 0 And bComputersTurn = True Then
                ' let's randomly select one to shoot at next turn
                bmpX = Int(Rnd * Len(.ActiveHits) / 4 + 1)
                .LastHit = Val(Mid(.ActiveHits, (bmpX - 1) * 4 + 1, 4))
                ' if that was not the only active hit, let's see if a pattern exits with the randomly selected one
                If Len(.ActiveHits) > 4 Then
                    For bmpY = 1 To 4
                        ' check each surrounding grid to see if a destroyed section is there
                        GridID = .LastHit + Choose(bmpY, -1, 1, -10, 10)
                        If GridID > 0 And GridID < 101 Then
                            ' destroyed sections have a value of -1
                            ' missed & sunk have -5
                            ' open space have 0 & ships with health between 1 and 9
                            If Grid(iOpponent, GridID) = -1 Then sPattern = sPattern & bmpY
                            ' add to pattern string if applicable indicating direction of pattern
                        End If
                    Next
                    If Len(sPattern) Then
                        ' we have a destroyed section adjacent to our next hit
                        ' so we randomly select a pattern
                        bmpX = Int(Rnd * Len(sPattern) + 1)
                        .Pattern = Choose(bmpX, 1, 1, -1, -1)
                    End If
                End If
            End If
        End With
        If bComputersTurn Then
            ' now to update largest opponent ship still on board
            For bmpX = 0 To UBound(Player(iOpponent).ID)
                If Player(iOpponent).Strength(bmpX) Then Exit For
            Next
            ' the value of bmpX should never be zero, otherwise the game would
            ' be over. So, if bmpX is zero, we have a problem. So we set minimum size to 1
            If bmpX > UBound(Player(iOpponent).ID) Then
                PCdata(0).MaxShipSize = 1
            Else
                PCdata(0).MaxShipSize = Player(iOpponent).Size(bmpX)
            End If
        End If
    End If
End If
' remove the temp picBox used for background masking
Unload picPlacement(picPlacement.UBound)
End Sub

Private Function ShipCollision(rRect As RECT, ShipNr As Integer, Optional bShowCollision As Boolean) As Boolean
' Function calculates whether the passed ship overlaps any other ships

Dim iShip As Integer, bCollide As Boolean, dRect As RECT
For iShip = 0 To UBound(Player(PlayerID).ID)  ' loop thru all ships
    If ShipNr <> iShip Then ' obviously--don't include the passed ship
        ' dRect will have an geometrical area if a collision occurred
        IntersectRect dRect, rRect, Player(PlayerID).Location(iShip)
        bCollide = ((dRect.Right > dRect.Left) Or (dRect.Bottom > dRect.Top))
        If bCollide Then    ' collision found
            If bShowCollision Then  ' option to display collision rectangle
                Load picPlacement(picPlacement.UBound + 1)
                With picPlacement(picPlacement.UBound)
                    .ToolTipText = "Ships at this position are overlapped"
                    .Width = dRect.Right - dRect.Left
                    .Height = dRect.Bottom - dRect.Top
                    .Top = dRect.Top
                    .Left = dRect.Left
                    .BackColor = vbMagenta
                    .ZOrder
                    .Visible = True
                End With
            End If
            Exit For
        End If
    End If
Next
ShipCollision = bCollide
End Function

Private Sub TimerMain_Timer()
' Timer's only function is to flash the collision rectangle
If GP = -1 Then
    With picPlacement(picPlacement.UBound)
        If .BackColor = vbMagenta Then .BackColor = vbBlack Else .BackColor = vbMagenta
        .Refresh
    End With
Else
    ' used to let computer shoot when playing against the computer
    If frameStep(0).Visible Then
        TimerMain.Enabled = False
        LetComputerShoot
    End If
End If
End Sub

Private Sub SetStrengths()
' This routine sets the shield valus for each player's ships

Dim Looper As Integer, I As Integer, P As Integer, iSection As Integer
Dim Strength As Integer, GridID As Integer

Randomize Timer
ReDim Grid(1 To 2, 0 To 100)
For P = 1 To 2  ' loop thru each player
    For Looper = 0 To UBound(Player(PlayerID).ID)   ' loop thru each ship
        iSection = Player(PlayerID).Size(Looper)    ' size of ship
        With Player(P).Location(Looper)
            ' get grid ref to each section of a ship
            GridID = ConvertGridCoord2Integer(.Left, .Top)
            For I = 1 To iSection                   ' loop thru each section
                If I - 1 Then
                    ' move to the next section
                    If .Right - .Left > .Bottom - .Top Then ' horizontal
                        GridID = GridID + 1
                    Else                                    ' vertical
                        GridID = GridID + 10
                    End If
                End If
                ' determine if ship is shielded
                If Player(P).Shield(1) = Looper Or _
                    Player(P).Shield(2) = Looper Or _
                        Player(P).Shield(3) = Looper Then
                        ' if so randomly set a section strength
                        Strength = Int(Rnd * 6 + 3)
                Else    ' not shielded
                    ' but we give extra strength 25% of the time anyway
                    Strength = Int(Rnd * 75 + 1)
                    If Strength < 76 Then Strength = 1 Else Strength = 2
                End If
                ' however, if playing without shields we set strength to 1
                If GameMode > 2 Then Strength = 1
                Grid(P, GridID) = Strength
            Next
        End With
        ' here we set the non-destroyed hull count. This is the same as
        ' the .Size(Looper) value, but this value gets reduced as each
        ' section is destroyed
        Player(P).Strength(Looper) = iSection
    Next
Next
End Sub

Private Sub ChangePlayer(Optional bStart As Boolean = False, Optional bEndofGame As Boolean)
' Function resets screen for the current player
Dim Looper As Integer
Dim NewX As Long, NewY As Long, lColor As Long
Dim bmpX As Long, bmpY As Long, hBitmap As IPictureDisp, iShip As Integer
Dim StatXY As POINTAPI, iOpponent As Integer
Dim tDC As Long, tBMP As Long, oldBMP As Long
If Not bStart Then  ' game in play
    ' place the current screen in the clipboard
    'Clipboard.Clear                -- the following works well on '98 but
    'OpenClipboard hwnd         -- didn't seem to work with NT
    'SetClipboardData 2, CopyImage(picField.Image.handle, 0, 0, 0, LR_COPYRETURNORG)
    'CloseClipboard
    tDC = CreateCompatibleDC(picField.hdc)
    tBMP = CreateCompatibleBitmap(picField.hdc, picField.Width, picField.Height)
    oldBMP = SelectObject(tDC, tBMP)
    SetBkColor tDC, GetBkColor(picField.hdc)
    SetTextColor tDC, GetTextColor(picField.hdc)
    BitBlt tDC, 0, 0, picField.Width, picField.Height, picField.hdc, 0, 0, vbSrcCopy
    ' now copy the next player's screen to the playing area
    BitBlt picField.hdc, 0, 0, picField.Width, picField.Height, picDrawingBoard.hdc, 0, 0, vbSrcCopy
    ' copy the current player's screen into the picBox & clear clipboard
    'picDrawingBoard.Picture = Clipboard.GetData(vbCFBitmap)
    BitBlt picDrawingBoard.hdc, 0, 0, picField.Width, picField.Height, tDC, 0, 0, vbSrcCopy
    DeleteObject SelectObject(tDC, oldBMP)
    DeleteObject tBMP
    DeleteDC tDC
    DoEvents
    picField.Refresh
    'Clipboard.Clear
Else    ' first screen of the game
    ' copy the blank screen to be used for player#2
    picDrawingBoard.AutoRedraw = True
    BitBlt picDrawingBoard.hdc, 0, 0, picField.Width, picField.Height, picField.hdc, 0, 0, vbSrcCopy
    PlayerID = 2
End If
PlayerAnimation(PlayerID).ShowNextFrame 1, True ' stop ship animation
iOpponent = PlayerID            ' identify the opponent
If iOpponent = 1 Then PlayerID = 2 Else PlayerID = 1
PlayerAnimation(PlayerID).ShowNextFrame 1, False ' start animation for current player
If bEndofGame Then Exit Sub
frameStep(1).Caption = Label1(PlayerID).Caption
For Looper = 0 To 2
    ' reset ammo count for current player
    shpAmmo(Looper).Width = (Player(PlayerID).Ammo(Looper) / Choose(Looper + 1, Ammo_Laser, Ammo_Cannon, Ammo_Torpedo)) * lblAmmo(Looper).Width
Next
' with the scanners, ensure proper number displayed
imgScan(0).Visible = (Player(PlayerID).Scans(0) = 2)
imgScan(1).Visible = CBool(Player(PlayerID).Scans(0))
imgScan(0).Enabled = imgScan(0).Visible
imgScan(1).Enabled = imgScan(1).Visible
' start of stop radar animation, depending on player scans available
If Player(PlayerID).Scans(0) Then
    ScannerAnimation.ShowNextFrame 1
    TimerScanner.Enabled = True
Else
    picScanner.Cls
End If
If GameMode > 4 Then    ' salvo game
    ' reset number of mines positioned & update the total mines available
    SalvoXY(PlayerID, 0) = 0
    RearrangeSalvos -1
End If
DrawEnemyStats  ' function updates the EnemyStatus screen
' change cursor to match player number
Set MouseIcon = LoadResPicture(100 + PlayerID, vbResIcon)
MousePointer = vbCustom
lblGridID.Caption = ""
' move the ammo selector to player's last selection
shpAmmo(3).Top = lblAmmo(Player(PlayerID).Ammo(3)).Top
' activate the scan rectangle if needed
If Player(PlayerID).Scans(1) Then ShowScanArea 0 Else bdrScan.Visible = False
End Sub

Private Sub TimerAnimation_Timer()
' calls class action to display next animation frame
PlayerAnimation(PlayerID).ShowNextFrame
End Sub

Private Sub TimerScanner_Timer()
' calls class action to display next radar animation frame
ScannerAnimation.ShowNextFrame
End Sub

Private Function TakeShot(GridID As Integer) As Boolean
' Function takes a shot based off of the grid selection border
Dim ammoIndex As Long, X As Long, Y As Long, sTimer As Single
' get actual X,Y for grid being shot at
ConvertGridCoord2Integer X, Y, GridID
' play the wav for the appropriate ammo used
    ammoIndex = Player(PlayerID).Ammo(3)
    PlayerAnimation(PlayerID).ShowNextFrame 29, True ' move animated ship to frame 1
    picAniShip.Top = picField.Top + Y     ' move ship on screen
    DoEvents
    If GameMode < 5 Then
        BeginPlaySound CStr(Val(ammoIndex) + 101), , True ' Fire! make program wait
    End If
' after sound effect...
' - adjust ammo count and visual count
Player(PlayerID).Ammo(ammoIndex) = Player(PlayerID).Ammo(ammoIndex) - 1
shpAmmo(ammoIndex).Width = (Player(PlayerID).Ammo(ammoIndex) / Choose(ammoIndex + 1, Ammo_Laser, Ammo_Cannon, Ammo_Torpedo)) * lblAmmo(ammoIndex).Width
' if ammo is now out, force selection to the Laser (pretty much unlimited ammo count)
If Player(PlayerID).Ammo(ammoIndex) = 0 Then Call lblAmmo_Click(0)
' hide the grid indicator if needed & prevent it from displaying for now
ShowGrid False
bShowSelectedGrid = False
' disable the form to prevent unwanted clicks until after routines are done
Enabled = False
' call function to display results of the firing
If ShowShot(GridID, X, Y) Then
    ' end of game -- reset stuff
    bComputersTurn = False
    ReleaseCapture
    ShowGrid False
    DoEndofGame False   ' function displays end of game frame
    Enabled = True
    bdrScan.Visible = False
    TakeShot = False    ' used for Salvo games
Else
    TakeShot = True     ' used for Salvo games
    If GameMode > 4 Then Exit Function
    ' after shot takes place, delay for a second
    sTimer = Timer
    Do While Abs(Timer - sTimer) < 1.5
        DoEvents
    Loop
    bShowSelectedGrid = True    ' allow grid indicator to display
    ChangePlayer                ' change player screen & stats
    ' if against computer & computer's turn, then...
    If PlayerID = 2 And optOpponent(1) = True Then
        bComputersTurn = True
        TimerMain.Enabled = True
    Else                              ' otherwise...
        bComputersTurn = False
        Enabled = True
        SendMessage picField.hwnd, &H200, 0&, 0&
    End If
End If
End Function

Private Sub ShowExplosion(sRect As RECT)
' Function simply animates a fireball over a background mask
Dim hBitmap As IPictureDisp, sTimer As Single, I As Integer
Dim NewX As Long, NewY As Long
' load the explosion bitmap in memory
Set hBitmap = LoadPicture(App.Path & "\Boom.gif")
' the image has 15 frames, each 62 x 84 pixels
CalculateRatio NewX, NewY, sRect.Right - sRect.Left - 2, sRect.Bottom - sRect.Top, 62, 84
For I = 1 To 14
    ' set up bitmap coordinates for the explosion bitmap
    ' calculate the Y coordinate of the next explosion frame
    dRect = MakeRectangle(((sRect.Right - sRect.Left) - NewX) / 2 + sRect.Left, ((sRect.Bottom - sRect.Top) - NewY) / 2 + sRect.Top, NewX, NewY)
    bmpRect = MakeRectangle(0, (I - 1) * 84, 62, 84)
    ' draw the explosion on the field of play
    DrawTransparentBitmap picField.hdc, dRect, hBitmap.Handle, bmpRect, -1, NewX, NewY, _
        picPlacement(picPlacement.UBound).hdc, _
        ((sRect.Right - sRect.Left) - NewX) / 2, _
        ((sRect.Bottom - sRect.Top) - NewY) / 2
    ' delay the program to show the explosion
    InvalidateRect picField.hwnd, sRect, 0
    sTimer = Timer
    Do While Abs(Timer - sTimer) < 0.14
        DoEvents
    Loop
Next
' unload the explosion bitmap
Set hBitmap = Nothing
End Sub

Private Function ShowShot(GridID As Integer, X As Long, Y As Long) As Boolean

' Function displays result of a grid being fired on

Dim hBitmap As IPictureDisp, iOpponent As Integer, sTimer As Single, sRect As RECT
Dim ShipID As Integer, dRect As RECT, Looper As Integer
Dim NewX As Long, NewY As Long

iOpponent = Abs((PlayerID - 2) - 1)     ' ref to opponent ID number
bmpRect = MakeRectangle(0, 0, 0, 0)     ' blank rectangle
CalculateShipStatus:
Select Case Grid(iOpponent, GridID)
Case 0  ' Miss, draw the circle & updated grid array value
    picField.FillColor = MissedColor(0) ' circle outline color
    picField.FillStyle = 0              ' solid vs transparent
    picField.Circle (Interval / 2 + X, Interval / 2 + Y), Interval / 5, MissedColor(1)
    Grid(iOpponent, GridID) = -5
    ' increment shots taken
    If bComputersTurn Then PCdata(0).ShotsTaken = PCdata(0).ShotsTaken + 1
Case 1  ' destroyed section & maybe sunk
    ' determine which ship was hit
    ' 1. make rectangle of grid being fired at
    sRect = MakeRectangle(X, Y, Interval, Interval)
    ' 2. see which actual X,Y coords of each ship intersects with this rectangle
    For ShipID = 0 To UBound(Player(PlayerID).ID)
        IntersectRect dRect, sRect, Player(iOpponent).Location(ShipID)
        ' if we have an intersection, then we know which ship was hit
        If dRect.Right > dRect.Left Or dRect.Bottom > dRect.Top Then Exit For
    Next
    ' reduce the number of sections still in tact
    Player(iOpponent).Strength(ShipID) = Player(iOpponent).Strength(ShipID) - 1
    bmpRect = GridItems(4)  ' get the skull image coords
    ' calculate size of image & placement in center of shot grid
    CalculateRatio NewX, NewY, Interval / 6 * 5, Interval / 6 * 5, _
        GridItems(4).Right - GridItems(4).Left, GridItems(4).Bottom - GridItems(4).Top
    ' update the number of shots taken
    If bComputersTurn Then PCdata(0).ShotsTaken = PCdata(0).ShotsTaken + 1
    ' if following value is zero, then ship was sunk, otherwise....
    If Player(iOpponent).Strength(ShipID) Then
        ' sound off hull explosion
        If GameMode < 5 Then
            BeginPlaySound App.Path & "\HullGone.wav"
        Else
            PCdata(0).ActiveHits = PCdata(0).ActiveHits & Format(GridID, "000.")
        End If
        If bComputersTurn Then
            PCdata(PCdata(0).Index).LastHit = GridID   ' update last hit to a positive number
                                            ' negative numbers indicate hull section still in tact
            If InStr(PCdata(PCdata(0).Index).ActiveHits, Format(GridID, "000.")) = 0 Then
                ' ensure this section is added to the active hits
                PCdata(PCdata(0).Index).ActiveHits = Format(GridID, "000.") & PCdata(PCdata(0).Index).ActiveHits
            End If
        End If
    Else    ' sunk ship
        ' return a value indicating whether game is over or not
        ShowShot = (ShipsRemaining(iOpponent) = 0)
        sTimer = 99 ' flag indicating game is over
        MousePointer = vbDefault
    End If
    ' update grid array value. If sunk it will be changed to -5 later
    Grid(iOpponent, GridID) = -1
Case Is > 1 ' shield value may still be in tact
        If bComputersTurn Then
            ' ensure hit is in active hit string
            If InStr(PCdata(PCdata(0).Index).ActiveHits, Format(GridID, "000.")) = 0 Then
                PCdata(PCdata(0).Index).ActiveHits = Format(GridID, "000.") & PCdata(PCdata(0).Index).ActiveHits
                ' flag indicating this ship takes more than 1 hit
                HullStrength(GridID) = 1
            End If
            ' update last hit grid ID. Negative value indicates section not destroyed
            PCdata(PCdata(0).Index).LastHit = GridID * -1
        End If
        ' subtract the number of hit points based off of the ammo used
        Grid(iOpponent, GridID) = Grid(iOpponent, GridID) - Choose(Player(PlayerID).Ammo(3) + 1, 1, 3, 5)
        ' now based on that resulting value....
        Select Case Grid(iOpponent, GridID)
        Case Is < 1 ' section destroyed
            ' send this back up top to be handled there
            Grid(iOpponent, GridID) = 1
            GoTo CalculateShipStatus
        Case Is < 4 ' section is weak (red shield)
            bmpRect = GridItems(3)
        Case Is < 7 ' section is well sheilded (yellow shield)
            bmpRect = GridItems(2)
        Case Else   ' section is strong (green shield)
            bmpRect = GridItems(1)
        End Select
        ' send hit sound & calculate size of shield to be displayed
        BeginPlaySound App.Path & "\Hit.wav"
        CalculateRatio NewX, NewY, Interval / 4 * 5, Interval / 4 * 5, 32, 32
End Select
If bmpRect.Right > 0 Then
    ' location where shield/skull will be drawn
    dRect = MakeRectangle(X + (Interval - NewX) \ 2, Y + (Interval - NewY) \ 2, Interval, Interval)
    DrawTransparentBitmap picField.hdc, dRect, miscBmp.Handle, bmpRect, -1, NewX, NewY
    InvalidateRect picField.hwnd, dRect, 0
    If sTimer Then  ' sinking to take place
        BeginPlaySound App.Path & "\Explode.wav"
        ' delay before showing the explosion
        sTimer = Timer
        Do While Abs(Timer - sTimer) < 1
            DoEvents
        Loop
        ' show the explosion
        SinkShip Player(iOpponent).Location(ShipID), ShipID
    End If
End If
End Function

Private Sub DrawEnemyStats(Optional bNewScan As Boolean = False)
Dim StatXY As POINTAPI, hBitmap As IPictureDisp
Dim iShip As Integer, lColor As Long, Looper As Integer
Dim NewX As Long, NewY As Long
Dim bmpX As Long, bmpY As Long, iOpponent As Integer

iOpponent = Abs((PlayerID - 2) - 1)
If Not bNewScan Then
    ' load the current player's ship selections in memory
    picEnemyStat.Cls
    For Looper = 1 To UBound(Player(PlayerID).ID) + 1
        ' draw the opponents ship status on the screen
        lColor = -1
        Select Case Looper
        Case 1:
            iShip = 0
        Case 2, 3:
            iShip = Looper
        Case 4
            iShip = Looper
        Case 5:
            lColor = vbBlue
            iShip = 1
        Case Else:
            If GameMode Mod 2 Then
                iShip = Looper - 1
            Else
                lColor = vbBlack: iShip = -1
            End If
        End Select
        If lColor < 0 Then lColor = Player(PlayerID).Color(iShip)
        ' don't display sunk ships (or maybe display as disabled?)
        ' calculate image size needed
        If iShip > -1 Then
            bmpX = Player(iOpponent).BMPxy(iShip).Right - Player(iOpponent).BMPxy(iShip).Left
            bmpY = Player(iOpponent).BMPxy(iShip).Bottom - Player(iOpponent).BMPxy(iShip).Top
            CalculateRatio NewX, NewY, 48, 48, bmpX, bmpY
            ' set the X,Y coord to place ship on status picBox
            If Looper < 5 Then
                StatXY.X = (Looper - 1) * 48 + 6
                StatXY.Y = 1
            Else
                StatXY.X = (Looper - 5) * 48 + 6
                StatXY.Y = 50
            End If
            dRect = MakeRectangle(((48 - NewX) / 2) + StatXY.X, _
                ((48 - NewY) / 2) + StatXY.Y, 48, 48)
            bmpRect = MakeRectangle(0, Player(iOpponent).BMPxy(iShip).Top, bmpX, bmpY)
            If Player(iOpponent).Strength(iShip) Then
                picEnemyStat.FillColor = lColor ' set background color for next function
                ' draw a rectangle on the picBox
                Rectangle picEnemyStat.hdc, StatXY.X, StatXY.Y, StatXY.X + 48, StatXY.Y + 48
            End If
            ' now draw the ship on the picBox
            DrawTransparentBitmap picEnemyStat.hdc, dRect, sourceBmp.Handle, bmpRect, -1, NewX, NewY
        End If
    Next
End If
DrawMiniIcons:
If Player(PlayerID).Scans(1) Then
    iShip = Player(PlayerID).Scans(1) - 1
    If GameMode Mod 2 Then
        StatXY.Y = Choose(iShip + 1, 1, 50, 1, 1, 1, 50, 50, 50) + 28
        StatXY.X = Choose(iShip + 1, 1, 1, 2, 3, 4, 2, 3, 4) * 48 - 14
    Else
        StatXY.Y = Choose(iShip + 1, 1, 50, 1, 1, 1) + 28
        StatXY.X = Choose(iShip + 1, 1, 1, 2, 3, 4) * 48 - 14
    End If
    picEnemyStat.FillColor = vbWhite
    Rectangle picEnemyStat.hdc, StatXY.X, StatXY.Y, StatXY.X + 20, StatXY.Y + 20
    picEnemyStat.PaintPicture imgScan(0).Picture, StatXY.X + 1, StatXY.Y + 1, 18, 18
End If
If Player(PlayerID).CloakRevealed Then
    iShip = Player(PlayerID).CloakRevealed - 1
    ' Don't draw miniCloak icon if ship was already destroyed
    If Player(iOpponent).Strength(iShip) Then
        picEnemyStat.FillColor = vbWhite
        If GameMode Mod 2 Then
            dRect.Left = Choose(iShip + 1, 1, 1, 2, 3, 4, 2, 3, 4) * 48 - 14
            dRect.Top = Choose(iShip + 1, 1, 50, 1, 1, 1, 50, 50, 50) + 28
        Else
            dRect.Left = Choose(iShip + 1, 1, 1, 2, 3, 4) * 48 - 14
            dRect.Top = Choose(iShip + 1, 1, 50, 1, 1, 1) + 28
        End If
        Rectangle picEnemyStat.hdc, dRect.Left, dRect.Top, dRect.Left + 20, dRect.Top + 20
        dRect = MakeRectangle(dRect.Left + 1, dRect.Top + 1, 20, 20)
        bmpRect = GridItems(5)
        DrawTransparentBitmap picEnemyStat.hdc, dRect, miscBmp.Handle, bmpRect, -1, 18, 18
        Set hBitmap = Nothing
    End If
End If
picEnemyStat.Refresh
End Sub

Private Sub DoEndofGame(bShowOpponentShips As Boolean)
Dim Looper As Integer
If bShowOpponentShips Then
    ChangePlayer , True
    If Len(optShow(2).Tag) Then
        For Looper = 0 To UBound(Player(PlayerID).ID)
            If Player(Abs((PlayerID - 2) - 1)).Strength(Looper) Then
                SinkShip Player(Abs((PlayerID - 2) - 1)).Location(Looper), Looper, True
            End If
        Next
        picField.Refresh
        optShow(2).Tag = ""
    End If
Else
    picPlacement(0).Visible = False
    For Looper = picPlacement.UBound To 1 Step -1
        Unload picPlacement(Looper)
    Next
    frameStep(1).Enabled = False
    frameStep(2).Enabled = True
    frameStep(2).ZOrder
    FlashSunkenShips
    If optOpponent(1) = True And PlayerID = 2 Then ' the computer won
        lblGameOver(1).Visible = True
        lblGameOver(0).Visible = False
        If Len(Dir(App.Path & "\EndGame.wav")) = 0 Or mnuOpts(0).Checked = False Then
            MsgBox "And the Winner Is..." & vbCrLf & vbCrLf & Player(PlayerID).Name, vbInformation + vbOKOnly, "Game Over"
        End If
    Else
        lblGameOver(0).Caption = Replace$(lblGameOver(0).Tag, "#", Player(PlayerID).Name)
        lblGameOver(0).Visible = True
        lblGameOver(1).Visible = False
        MsgBox "And the Winner Is..." & vbCrLf & vbCrLf & Player(PlayerID).Name, vbInformation + vbOKOnly, "Game Over"
    End If
    optShow(2).Tag = "Display"
    optShow(1).Tag = "NoSync"
    optShow(Abs((PlayerID - 2) - 1)) = True
    optShow(1).Tag = ""
End If
End Sub

' -------------------------------------------------------------------
' Following functions and subroutines are for the computer opponent |
' -------------------------------------------------------------------
Private Sub LetComputerShoot()
Dim GridID As Integer, X As Long, Y As Long
If GameMode < 5 Then
    If PCdata(PCdata(0).Index).LastHit Then
        GridID = GetNextHit
    Else
        GridID = GetBestGrid
    End If
    Dim Looper As Integer, sTimer As Single
    shpGrid.Visible = True
    For Looper = 1 To 5
        ConvertGridCoord2Integer X, Y, Int(Rnd * 100 + 1)
        shpGrid.Move X, Y
        sTimer = Timer
        Do While Abs(Timer - sTimer) < 0.25
            DoEvents
        Loop
    Next
    ConvertGridCoord2Integer X, Y, GridID
    shpGrid.Move X, Y
    sTimer = Timer
    Do While Abs(Timer - sTimer) < 0.25
        DoEvents
    Loop
    Call picField_Click
    ShowGrid False
    Exit Sub
End If
Dim pcSalvo(1 To 100), nrEnemy As Integer, nrShot As Integer, nrMines As Integer
Dim I As Integer, J As Integer, iPattern As Integer, idxSalvo As Integer, sDelay As Single
Dim sHits As String, sOrder As String, nrShips As Integer, PatternCount As Integer

If GameMode Mod 2 Then nrShips = 8 Else nrShips = 5
nrEnemy = ShipsRemaining(1)
nrMines = ShipsRemaining(2)
For I = 1 To 100
    pcSalvo(I) = Grid(1, I)
Next
GoSub SetSalvoOrder
For I = 1 To nrMines
    idxSalvo = Val(Mid(sOrder, I, 1))
    iPattern = PCdata(idxSalvo).Pattern
    PCdata(0).Index = idxSalvo
    If PCdata(idxSalvo).LastHit Then
        GridID = GetNextHit(PatternCount)
        If PatternCount + 1 > PCdata(0).MaxShipSize And nrEnemy > 1 Then GridID = GetBestGrid
    Else
        GridID = GetBestGrid
    End If
    If Grid(1, GridID) < 0 Then GridID = GetBestGrid
    If PCdata(idxSalvo).Pattern Then
        If iPattern Then        ' previous pattern existed, assume pattern continued
            Grid(1, GridID) = -1
        Else                       ' no previous pattern, remove pattern otherwise it's cheating
            Grid(1, GridID) = -5
        End If
    Else
        Grid(1, GridID) = -5
    End If
    PCdata(idxSalvo).Pattern = iPattern
    SalvoXY(3, I) = idxSalvo
    SalvoXY(2, I) = GridID
    ConvertGridCoord2Integer X, Y, GridID
    With picPlacement(I - 1)
        .Move ((Interval - .Width) \ 2) + X, ((Interval - .Height) \ 2) + Y
        .Visible = True
    End With
Next
For I = 1 To 100
    Grid(1, I) = pcSalvo(I)
Next
Erase pcSalvo
SalvoXY(2, 0) = nrMines
picSalvo.Cls
If nrMines > 1 Then sDelay = 1.5 Else sDelay = 0.66
sTimer = Timer
Do While Abs(Timer - sTimer) < sDelay
    DoEvents
Loop
Call cmdFireSalvo_Click
Exit Sub

SetSalvoOrder:
sOrder = "": sHits = ""
nrShot = 0
PCdata(0).ActiveHits = PCdata(0).SalvoTurns
For I = 1 To Len(PCdata(0).SalvoTurns)
    J = Int(Rnd * Len(PCdata(0).ActiveHits) + 1)
    J = Val(Mid$(PCdata(0).ActiveHits, J, 1))
    If PCdata(J).LastHit Then
        nrShot = nrShot + 1
        sHits = sHits & J
        sOrder = J & J & sOrder
    Else
        sOrder = sOrder & J
    End If
    PCdata(0).ActiveHits = Replace(PCdata(0).ActiveHits, J, "")
Next
If nrEnemy < nrMines Then
    If Len(sHits) Then sOrder = Left(sOrder, Len(sHits) * 2)
    Do While Len(sOrder) < nrMines
        sOrder = sOrder & Int(Rnd * Len(sHits) + 1)
    Loop
End If
Return
End Sub

Private Function GetBestGrid() As Integer
' Use when playing against the computer. This is the computers way of
' choosing a grid to shoot at when no ships are actively being hit
Dim X As Long, Y As Long, GridID As Integer, iCount As Byte, iTotal As Byte
Dim sChoice As String, GridChoice() As Byte, iMax As Byte, iShip As Integer
Dim sGrid As String, GridAttr(1 To 100, 0 To 4) As Byte, iBestSection As Integer
Dim GridStart As Integer, GridStop As Integer

Randomize Timer
If Player(2).Scans(1) = 0 And Player(2).Scans(0) > 0 Then
    ' if the computer hasn't scanned an enemy ship, see if now would be a good time
    For Y = 0 To UBound(Player(PlayerID).ID)
        ' we determine which ships are still active
        If Player(1).Strength(Y) Then iCount = iCount + 1
    Next
    Select Case Player(2).Scans(0)
    Case 2  ' both scans still available
        ' fire the 1st scan around the 20-shot range or if only 2 ships are left
        If PCdata(0).ShotsTaken > 20 Or iCount < 3 Then iShip = 1
    Case 1  ' one scan available
        ' fire the last scan around the 35 shot range or if only 1 ship is left
        If PCdata(0).ShotsTaken > 35 Or iCount = 1 Or (Count = 2 And Player(2).CloakRevealed > 0) Then iShip = 2
    End Select
    If iShip Then
        ' if we are going to scan, let's decide which ship to scan
        For Y = 0 To UBound(Player(PlayerID).ID)
            ' loop thru each non-destroyed ship
            If Player(1).Strength(Y) > 0 And Player(2).CloakRevealed <> Y + 1 Then
                Select Case Y
                Case 0, 1       ' large ships
                    sGrid = sGrid & Y + 1
                Case 2, 3, 4    ' middle size ships
                    If GameMode Mod 2 = 0 And Y = 4 Then
                        sGrid = sGrid & "5555555"
                    Else
                        sGrid = sGrid & String$(3, CStr(Y + 1))
                    End If
                Case Else       ' small ships
                    sGrid = sGrid & String$(7, CStr(Y + 1))
                End Select
            End If
        Next
        If Len(sGrid) Then
            ' lets remove large ships from the string unless that is the only ship in the string
            If sGrid <> "1" Then sGrid = Replace$(sGrid, "1", "")
            If sGrid <> "2" Then sGrid = Replace$(sGrid, "2", "")
            ' randomly select one of the remaining ships
            X = Val(Mid$(sGrid, Int(Rnd * Len(sGrid) + 1), 1))
            ' call function to initiate the scan
            Call picEnemyStat_DragDrop(imgScan(iShip - 1), -X, 0)
        Else
            ' no ships to scan, lets reset this value to prevent running
            ' thru this routine every shot
            Player(2).Scans(0) = 0
            Player(2).Scans(2) = 0
        End If
    End If
End If
If bdrScan.Visible Then ' actively scanning enemy ship
    ' we want to look for a grid to shoot at only in within the scan area
    GridStart = Player(2).Scans(2)
    GridStop = GridStart + 33
    ' size of ship we are scanning for
    iShip = Player(PlayerID).Size(Player(2).Scans(1) - 1)
Else
    ' otherwise we want to use the entire board to select the next shot
    GridStart = 1: GridStop = 100
    ' size of ship ware scanning for
    iShip = PCdata(0).MaxShipSize
End If
BeginInitialSearch:
sGrid = "": X = 0
For GridID = GridStart To GridStop
    ' now we loop thru each grid to attempt the best shot
    If Grid(1, GridID) > -1 Then        ' this grid is open
        ' determine how many horizontal spaces (open or destroyed hull) along this grid
        GridAttr(GridID, 0) = NumberGridsHorizontal(GridID, GridAttr(GridID, 1), GridAttr(GridID, 2))
        ' now we check vertical spaces & if either meet/exceed ship size, we add it to a string
        If GridAttr(GridID, 0) > iShip - 1 Or _
            NumberGridsVertical(GridID, GridAttr(GridID, 3), GridAttr(GridID, 4)) > iShip - 1 Then sGrid = sGrid & Chr$(50 + GridID)
    End If
    If bdrScan.Visible Then
        ' scanning. We need to keep track of grid ID to ensure we are not
        ' looking outside the grid
        X = X + 1
        If X = 4 Then   ' hit right edge of scan area
            GridID = GridID + 6 ' move down to the next line, left edge of scan
            X = 0       ' reset grid counter
        End If
    End If
Next
If Len(sGrid) = 0 Then
    ' should never happen, but just in case
    iShip = iShip - 1           ' reduce ship size & try again
    If iShip = 0 Then
        MsgBox "Sorry. You need to restart the game.", vbInformation + vbOKOnly
        Exit Function
    End If
    GoTo BeginInitialSearch
End If
' now we set up an array to relook at the grid we found above
' this array will hold individual counts of adjacent spaces
' above, below, right & left of the grid we are looking at
ReDim GridChoice(1 To Len(sGrid), 0 To 2)
For X = 1 To Len(sGrid)
       GridID = Asc(Mid$(sGrid, X, 1)) - 50 ' get grid ref out of string
       iTotal = 0                           ' reset total
        For Y = 1 To 4
            ' check each direction around the grid
            iCount = GridAttr(GridID, Y)    ' total count of all spaces
            If iCount > iShip Then iCount = iShip ' don't count more than target ship size
            iTotal = iTotal + iCount        ' new total
        Next
        GridChoice(X, 0) = CByte(GridID)    ' grid ID
        GridChoice(X, 1) = iTotal           ' best count
        If iTotal > iMax Then iMax = iTotal ' keep track of best count
Next
' now lets only keep those grids = to the best count
For X = 1 To Len(sGrid)
    If GridChoice(X, 1) = iMax Then
        ' get total number of open spaces horizontal and vertical
        GridChoice(X, 2) = NumberGridsHorizontal(GridID, GridAttr(GridID, 1), GridAttr(GridID, 2)) + NumberGridsVertical(GridID, GridAttr(GridID, 1), GridAttr(GridID, 2))
        If GridChoice(X, 2) > iBestSection Then iBestSection = GridChoice(X, 2)
    End If
Next
If PCdata(0).ShotsTaken < 8 Then
    ' this is to ensure truly random selections for the 1st 7 shots
    ' otherwise the computer is too predictable in its first series of shots
    sChoice = sGrid
Else
    ' after the first seven, we want only those grids with the largest number of adjacent spaces
    For X = 1 To Len(sGrid)
        If GridChoice(X, 2) = iBestSection Then
            sChoice = sChoice & Chr$(50 + GridChoice(X, 0))
        End If
    Next
End If
' extract a randomly selected grid to shoot at
GridID = Asc(Mid$(sChoice, Int(Rnd * Len(sChoice) + 1), 1)) - 50
If GameMode < 4 Then Call lblAmmo_Click(0) ' use the laser
GetBestGrid = GridID    ' return the grid ID
End Function

Private Function NumberGridsHorizontal(GridID As Integer, Optional NrLt As Byte, Optional NrRt As Byte, Optional MinValue As Integer = -1) As Integer
' function returns the number of open spaces to the left & right of a target grid and also
' returns the sum of the left & right open spaces
Dim X As Integer
NrLt = 0
NrRt = 0
For X = GridID - 1 To (((GridID - 1) \ 10)) * 10 + 1 Step -1
    If Grid(1, X) > MinValue Then NrLt = NrLt + 1 Else Exit For
Next
For X = GridID + 1 To ((GridID - 1) \ 10 + 1) * 10
    If Grid(1, X) > MinValue Then NrRt = NrRt + 1 Else Exit For
Next
NumberGridsHorizontal = NrRt + NrLt + 1
End Function

Private Function NumberGridsVertical(GridID As Integer, Optional NrUp As Byte, Optional NrDn As Byte, Optional MinValue As Integer = -1) As Integer
' function returns the number of open spaces to the top & bottom of a target grid and also
' returns the sum of the top & bottom open spaces
Dim X As Integer
NrUp = 0
NrDn = 0
For X = GridID - 10 To 1 Step -10
    If Grid(1, X) > MinValue Then NrUp = NrUp + 1 Else Exit For
Next
For X = GridID + 10 To 100 Step 10
    If Grid(1, X) > MinValue Then NrDn = NrDn + 1 Else Exit For
Next
NumberGridsVertical = NrDn + NrUp + 1
End Function

Private Function GetNextHit(Optional NrInPattern As Integer) As Integer
' This function is called when we have a ship being hit but hasn't been sunk yet

Dim X As Long, Y As Long, GridID As Integer, iTotal As Byte, sChoice As String
Dim sGrid As String, GridAttr(1 To 4, 0 To 2) As Byte, iMax As Byte
Dim bSwap As Boolean, iMinMax(0 To 1) As Integer, iPattern As Integer
Dim EndHits(0 To 3) As Integer, bScan As Boolean, bIgnoreScan As Boolean
Dim gridOffset As Integer

If PCdata(PCdata(0).Index).LastHit < 0 Then
    ' when this value is less than zero then we are shooting at a shielded section
    ' so let's finish it off
    GridID = Abs(PCdata(PCdata(0).Index).LastHit)  ' get actual grid ref
    If Grid(1, GridID) > 3 Then Y = 2 Else Y = 1    ' choose ammo
    Call lblAmmo_Click(CInt(Y))         ' select the ammo
    GoTo FinishUp                       ' return the grid ID to shoot at
End If
With PCdata(PCdata(0).Index)
' patterns: if we have sunk a section but not the ship and then subsequently
' hit another occupied adjacent space, we have a pattern.  We try to stay
' with the pattern until either the ship is sunk or the pattern can no
' longer by followed because both ends of the pattern are blocked
CheckPattern:
    If .Pattern Then
        ' pattern of 1=horizontal, -1=vertical
        If .Pattern < 0 Then Y = 10 Else Y = 1  ' direction of pattern
        For X = .LastHit To 0 Step -Y
            ' determine last hit section within the pattern (left or top)
            If Grid(1, X) <> -1 Then Exit For
        Next
        EndHits(0) = X + Y
        For X = .LastHit To 100 Step Y
            ' determine last hit section within the pattern (right or bottom)
            If Grid(1, X) <> -1 Then Exit For
        Next
        EndHits(1) = X - Y
        NrInPattern = (EndHits(1) - EndHits(0) + 1) / Y
        For X = 1 To 2
            ' we loop thru each direction of the pattern to see if the pattern can be continued
            GridID = EndHits(X - 1)
            ' GridOffset is the direction to test
            If .Pattern < 0 Then gridOffset = Choose(X, -10, 10) Else gridOffset = Choose(X, -1, 1)
            ' ensure the grid we are testing is on the playing field
            If GridID + gridOffset > 0 And GridID + gridOffset < 101 Then
                GridAttr(X, 0) = GridID + gridOffset
                ' test each direction of the pattern
                Select Case X * .Pattern
                Case -1:
                    If Grid(1, GridAttr(X, 0)) > -2 Then NumberGridsVertical GridID, GridAttr(X, 1)
                Case -2:
                    If Grid(1, GridAttr(X, 0)) > -2 Then NumberGridsVertical GridID, 0, GridAttr(X, 1)
                Case 1:
                    If Grid(1, GridAttr(X, 0)) > -2 Then NumberGridsHorizontal GridID, GridAttr(X, 1)
                Case 2:
                    If Grid(1, GridAttr(X, 0)) > -2 Then NumberGridsHorizontal GridID, 0, GridAttr(X, 1)
                End Select
                ' keep track of which end of the pattern has the most open spaces
                If GridAttr(X, 1) > iTotal Then iTotal = GridAttr(X, 1)
            End If
            EndHits(X + 1) = iTotal
            iTotal = 0
        Next
        If EndHits(2) = 0 And EndHits(3) = 0 Then
            ' pattern is blocked on both ends, so we need to clear the pattern,
            ' choose a section to begin with to test for a new pattern
            X = Int(Rnd * (Len(.ActiveHits) / 4) + 1)
            .LastHit = Val(Mid(.ActiveHits, (X - 1) * 4 + 1, 4))
            .Pattern = 0
            GoTo CheckPattern
        Else
            If bdrScan.Visible = True And EndHits(2) > 0 And EndHits(3) > 0 Then
                ' when an active scan is present, if the choice is to shoot outside the
                ' scan area or inside, we want to force an inside shot if possible
                Dim sRect1 As RECT, sRect2 As RECT, gRect As RECT
                ' create rectangles for both ends of the pattern & of the scan area
                ' then test for a collision. Collision indicates grid within scan area
                ConvertGridCoord2Integer X, Y, EndHits(0) - gridOffset
                sRect1 = MakeRectangle(X, Y, Interval, Interval)
                ConvertGridCoord2Integer X, Y, EndHits(1) + gridOffset
                sRect2 = MakeRectangle(X, Y, Interval, Interval)
                ConvertGridCoord2Integer X, Y, CInt(Player(2).Scans(2))
                gRect = MakeRectangle(X, Y, Interval * 4, Interval * 4)
                IntersectRect sRect1, sRect1, gRect
                IntersectRect sRect2, sRect2, gRect
                ' now we see if collisions exist
                If sRect1.Right > sRect1.Left And sRect2.Right > sRect2.Left Then
                    ' both ends of pattern are in the scan area, randomly choose one
                    If Int(Rnd * 100) < 50 Then X = 0 Else X = 1
                Else
                    ' only one end is in scan area, select it
                    If sRect1.Right > sRect1.Left Then X = 0 Else X = 1
                End If
            Else    ' no active scans, randomly select one of the two ends
                If Int(Rnd * 100) < 50 Then X = 0 Else X = 1
            End If
            ' here we ensure the selected grid is not blocked & if so, choose the other one
            If EndHits(X + 2) = 0 Then X = Abs(X - 1)
            ' pick the actual grid we want to shoot at
            GridID = EndHits(X) + Choose(X + 1, -gridOffset, gridOffset)
            iTotal = EndHits(X + 2)
        End If
    Else                ' We have destroyed sections but no pattern!
LookforAdjacentGrid:
        ' No pattern exists. We simply go off the last successfully hit grid & test around it
        ' to see if we can shoot in any direction & if so, choose a direction
        For X = 1 To 4
            ' loop thru each direction
            Y = Choose(X, -10, 10, -1, 1)
            ' identify the grid boundaries to search in
            If Abs(Y) = 10 Then     ' testing vertically
                iMinMax(0) = 0: iMinMax(1) = 101
                If Player(2).Scans(1) And bIgnoreScan = False Then
                    ' active scan, let's limit boundaries
                    iMinMax(1) = Player(2).Scans(2) + 35
                    iMinMax(0) = Player(2).Scans(2) - 1
                End If
            Else                    ' testing horizontally
                iMinMax(0) = (Val(Left(Format(.LastHit - 1, "00"), 1))) * 10
                iMinMax(1) = iMinMax(0) + 11
                If Player(2).Scans(1) And bIgnoreScan = False Then
                    ' active scan, let's limit boundaries
                    iMinMax(0) = (Val(Right(Player(2).Scans(2), 1))) + (Val(Left(Format(.LastHit - 1, "00"), 1))) * 10 - 1
                    iMinMax(1) = iMinMax(0) + 5
                End If
            End If
            ' if the grid to test in is within boundaries, let's continue
            If .LastHit + Y > iMinMax(0) And .LastHit + Y < iMinMax(1) Then
                GridAttr(X, 0) = .LastHit + Y
                ' for each of the 2 directions, get number of open spaces, left & right. Sunk sections are included
                Select Case X
                Case 1:
                    If Grid(1, GridAttr(X, 0)) > -2 Then NumberGridsVertical .LastHit, GridAttr(X, 1), GridAttr(X, 2)
                Case 2:
                    If Grid(1, GridAttr(X, 0)) > -2 Then NumberGridsVertical .LastHit, GridAttr(X, 1), GridAttr(X, 2)
                Case 3:
                    If Grid(1, GridAttr(X, 0)) > -2 Then NumberGridsHorizontal .LastHit, GridAttr(X, 1), GridAttr(X, 2)
                Case 4:
                    If Grid(1, GridAttr(X, 0)) > -2 Then NumberGridsHorizontal .LastHit, GridAttr(X, 1), GridAttr(X, 2)
                End Select
                If GridAttr(X, 1) + GridAttr(X, 2) > iTotal Then iTotal = GridAttr(X, 1) + GridAttr(X, 2)
            End If
        Next
        If iTotal = 0 And bIgnoreScan = False And bdrScan.Visible = True Then
            ' no adjacent open spaces in the scan area. This can happen if accidentally hitting another
            ' ship but the one we're looking for and that ship has its 1st section on the edge of
            ' the scan and extending out of the scan. In this case, we want to ignore the scanarea
            ' and try again. We'll come back to the scan area after this ship is sunk
            bIgnoreScan = True
            GoTo LookforAdjacentGrid
        End If
    End If
    If iTotal = 0 Then
        ' could not find any open spaces. This should never happen unless ships extend
        ' outside the playing field or overlap. But since those checks were done before
        ' the game started, this is a critical error & we simply reset everything to allow
        ' play to continue. The computer will never win if this happens
        If GameMode < 5 Then .LastHit = 0
        .Pattern = 0
        GetNextHit = GetBestGrid
        Exit Function
    End If
    ' if more than one direction, choose the one with the largest number of open spaces
    For X = 1 To 4
        If GridAttr(X, 1) + GridAttr(X, 2) = iTotal Then sGrid = sGrid & Chr$(GridAttr(X, 0) + 50)
    Next
    ' select the grid & choose appropriate ammo
    GridID = Asc(Mid$(sGrid, Int(Rnd * Len(sGrid) + 1), 1)) - 50
    ComputerPickAmmo GridID, .LastHit
    If Grid(1, GridID) > 0 And .Pattern = 0 Then
        ' this is looking forward, but since we are going to shoot here anyway, determine if the
        ' grid we are shooting at is a hit & if so set the pattern (horizontally or vertically)
        If Abs(GridID - .LastHit) < 10 Then .Pattern = 1 Else .Pattern = -1
    End If
End With
FinishUp:
GetNextHit = GridID
End Function

Private Sub ComputerPickAmmo(GridID As Integer, iLastHit As Integer)
' here we select ammo. With shields in play, under estimating the ammo needed means
' you need to shoot again at the target. This in effect gives your opponent an extra shot
' each time. So we want to overestimate and hope to sink a section on the first hit
If GameMode > 2 Then        ' no shields, randomly select ammo (only for sound effect)
    Call lblAmmo_Click(Int(Rnd * 2 + 1))
    Exit Sub
End If
' This routine is only called when we are targetting a pattern or off of a destroyed section
Dim NrAlive As Integer, Looper As Integer, GridAdj As Integer
NrAlive = ShipsRemaining(1)     ' number of opponent's active ships
If NrAlive = 1 Then
    ' if this is the last ship to sink, let's use maximum firepower
    Call lblAmmo_Click(2)
    Exit Sub
End If
For Looper = 1 To 4
    ' let's check around the area we are shooting at
    ' if its a pattern & any 1 of the sections in the pattern required
    ' more than the basic ammo, we max out the ammo with the theory
    ' that if one section was shielded, then the others must be too
    GridAdj = GridID + Choose(Looper, -1, 1, -10, 10)
    If GridAdj > 0 And GridAdj < 101 Then
        If HullStrength(GridAdj) Then
            ' check horizontally
            If Looper < 3 And PCdata(PCdata(0).Index).Pattern = 1 Then
                Call lblAmmo_Click(2)
                Exit Sub
            Else
                ' check vertically
                If Looper > 2 And PCdata(PCdata(0).Index).Pattern = -1 Then
                    Call lblAmmo_Click(2)
                    Exit Sub
                End If
            End If
        End If
    End If
Next
Call lblAmmo_Click(0)
If iLastHit = 0 Then Exit Sub
    ' with no pattern existing and being forced in one direction cause the
    ' other three sides of the previously destroyed section are blocked,
    ' we want to max out the ammo if the previously destroyed section was
    ' shielded
    NrAlive = 4
    GridAdj = (Val(Left(Format(iLastHit - 1, "00"), 1))) * 10
    If GridAdj < 0 Then
        NrAlive = NrAlive - 1
    Else
        If Grid(1, iLastHit - 1) < 0 Then NrAlive = NrAlive - 1
    End If
    GridAdj = (Val(Left(Format(iLastHit - 1, "00"), 1))) * 10 + 11
    If GridAdj > 100 Then
        NrAlive = NrAlive - 1
    Else
        If Grid(1, iLastHit + 1) < 0 Then NrAlive = NrAlive - 1
    End If
    If iLastHit - 10 < 1 Then
        NrAlive = NrAlive - 1
    Else
        If Grid(1, iLastHit - 10) < 0 Then NrAlive = NrAlive - 1
    End If
    If iLastHit + 10 > 100 Then
        NrAlive = NrAlive - 1
    Else
        If Grid(1, iLastHit + 10) < 0 Then NrAlive = NrAlive - 1
    End If
    If NrAlive = 1 Then Call lblAmmo_Click(2)
End Sub

Private Sub RearrangeSalvos(SalvoID As Integer)
Dim Looper As Integer, X As Long, Y As Long
If SalvoID > -1 Then
    For Looper = SalvoID + 1 To SalvoXY(PlayerID, 0) - 1
        SalvoXY(PlayerID, Looper) = SalvoXY(PlayerID, Looper + 1)
    Next
    For Looper = 1 To SalvoXY(PlayerID, 0) - 1
        ConvertGridCoord2Integer X, Y, SalvoXY(PlayerID, Looper)
        With picPlacement(Looper - 1)
            .Move ((Interval - .Width) \ 2) + X, ((Interval - .Height) \ 2) + Y
            .Visible = True
        End With
    Next
    SalvoXY(PlayerID, 0) = SalvoXY(PlayerID, 0) - 1
    picPlacement(Looper - 1).Visible = False
End If
With picSalvo
    .Cls
    .AutoRedraw = True
    For Looper = 1 To ShipsRemaining(PlayerID) - SalvoXY(PlayerID, 0)
        X = (.Width / Screen.TwipsPerPixelX) / 4 + 5
        X = X * Choose(Looper, 1, 2, 3, 4, 1, 2, 3, 4) - X
        dRect = MakeRectangle(X + 5, IIf(Looper < 5, 5, 34), 24, 24)
        bmpRect = GridItems(6)
        DrawTransparentBitmap .hdc, dRect, miscBmp.Handle, bmpRect, -1, 24, 24
    Next
    .Refresh
End With
Exit Sub

End Sub

Private Sub PostSalvoCleanup()
Dim I As Integer, J As Integer, GridID As Integer, sGridID As String
Dim adjGrid(1 To 4) As Integer, iMin As Integer, Hits() As Integer
Dim iPatternStart As Integer, iPatternStop As Integer
' After a computer salvo, patterns could be messed up.
' Each shot is like a different turn and some turns may share the
' same pattern. We need to loop thru the active hits & recreate
' patterns as needed
PCdata(0).ActiveHits = ""
For I = 1 To UBound(PCdata)
    ' loop thru each turn the computer took
     For J = 1 To Len(PCdata(I).ActiveHits) Step 4
        ' for each hit that is not a sinking, se add it to the (0).ActiveHits string
        GridID = Val(Mid$(PCdata(I).ActiveHits, J, 3))
        If Grid(1, GridID) > -2 Then
            sGridID = Format(GridID, "000.")
            If InStr(PCdata(0).ActiveHits, sGridID) = 0 Then
                PCdata(0).ActiveHits = PCdata(0).ActiveHits & sGridID
                ReDim Preserve Hits(0 To Len(PCdata(0).ActiveHits) / 4 - 1)
                Hits(UBound(Hits)) = Val(sGridID)
            End If
        End If
    Next
    PCdata(I).ActiveHits = ""       ' clear the active hits from this turn
    PCdata(I).LastHit = 0           ' remove ref to last successful hit
Next
' if we have not active hits, exit now
If Len(PCdata(0).ActiveHits) = 0 Then Exit Sub
ShellSortNumbers Hits
PCdata(0).ActiveHits = ""
For I = 0 To UBound(Hits)
    PCdata(0).ActiveHits = PCdata(0).ActiveHits & Format(Hits(I), "000.")
Next
Erase Hits
' otherwise we start creating patterns if they exist
I = 1
Do While Len(PCdata(0).ActiveHits)      ' contains all active hits for now
    ' pattern search initially in all 4 directions
    iPatternStart = 1: iPatternStop = 4
    ' get the first grid to check
    GridID = Val(Left(PCdata(0).ActiveHits, 3))
    GoSub GetAdjacents  ' routine to calculate what grids are around current grid
    ' add the current grid to the next turn
    PCdata(I).ActiveHits = Format(GridID, "000.")
    ' remove the current grid from the active hits
    PCdata(0).ActiveHits = Replace$(PCdata(0).ActiveHits, Format(GridID, "000."), "")
    For J = iPatternStart To iPatternStop
        ' loop thru each surrounding grid to see if it was hit
        sGridID = Format(adjGrid(J), "000.")
        If InStr(PCdata(0).ActiveHits, sGridID) Then
            If J < 3 Then       ' horizontal grids
                ' yep one was hit next to the current, we have a pattern
                PCdata(I).Pattern = 1           ' pattern direction
                PCdata(I).LastHit = GridID      ' ensure a ref to a current hit
                iPatternStart = 1: iPatternStop = 2 ' set pattern checks to horizontal
                J = 0                               ' reset to check only horizontal hits
            Else
                ' got a vertical pattern
                PCdata(I).LastHit = GridID      ' ensure ref to current hit
                PCdata(I).Pattern = -1          ' pattern direction
                iPatternStart = 3: iPatternStop = 4 ' set pattern checks to vertical
                J = 2                               ' reset to check only vertical hits
            End If
            ' let's add this grid to the next turn's active hits
            PCdata(I).ActiveHits = PCdata(I).ActiveHits & sGridID
            ' we remove the ref from (0).ActiveHits
            PCdata(0).ActiveHits = Replace$(PCdata(0).ActiveHits, sGridID, "")
            ' set the next grid to check & find surrounding grids
            GridID = Val(sGridID)
            GoSub GetAdjacents
        End If
    Next
    If PCdata(I).LastHit = 0 And Len(PCdata(I).ActiveHits) Then PCdata(I).LastHit = Val(Left(PCdata(I).ActiveHits, 3))
    If UBound(PCdata) > I Then
        I = I + 1
    Else
        If Len(PCdata(0).ActiveHits) Then PCdata(1).ActiveHits = PCdata(1).ActiveHits & PCdata(0).ActiveHits
        PCdata(0).ActiveHits = ""
    End If
Loop
Exit Sub

GetAdjacents:
    adjGrid(1) = GridID - 1
    adjGrid(2) = GridID + 1
    iMin = (Val(Left(Format(GridID - 1, "00"), 1))) * 10 + 1
    If adjGrid(1) < iMin Then adjGrid(1) = 0
    If adjGrid(2) > iMin + 9 Then adjGrid(2) = 0
    adjGrid(3) = GridID - 10
    adjGrid(4) = GridID + 10
    If adjGrid(3) < 0 Then adjGrid(3) = 0
    If adjGrid(4) > 100 Then adjGrid(4) = 0
Return
End Sub

Private Sub FlashSunkenShips()
Dim sTimer As Single, iShip As Integer, iOpponent As Integer, Looper As Integer
If PlayerID = 1 Then iOpponent = 2 Else iOpponent = 1
If Len(Dir(App.Path & "\EndGame.wav")) > 0 And _
    mnuOpts(0).Checked = True And _
        optOpponent(1) = True And iOpponent = 1 Then
            BeginPlaySound App.Path & "\EndGame.wav"
End If
For iShip = 0 To UBound(Player(iOpponent).ID)
    dRect = Player(iOpponent).Location(iShip)
    InvertRect picField.hdc, dRect
    picField.Refresh
    sTimer = Timer
    Do While Abs(Timer - sTimer) < 0.4
        DoEvents
    Loop
    InvertRect picField.hdc, dRect
Next
picField.Refresh
For Looper = 1 To 3
    For iShip = 0 To UBound(Player(iOpponent).ID)
        dRect = Player(iOpponent).Location(iShip)
        InvertRect picField.hdc, dRect
    Next
    picField.Refresh
    sTimer = Timer
    Do While Abs(Timer - sTimer) < 0.4
        DoEvents
    Loop
    For iShip = 0 To UBound(Player(iOpponent).ID)
        dRect = Player(iOpponent).Location(iShip)
        InvertRect picField.hdc, dRect
    Next
    picField.Refresh
    sTimer = Timer
    Do While Abs(Timer - sTimer) < 0.4
        DoEvents
    Loop
Next
End Sub
