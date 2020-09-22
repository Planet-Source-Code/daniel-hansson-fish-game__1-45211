VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Fish Test 4"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicStim 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   6240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   448
      TabIndex        =   36
      Top             =   3000
      Width           =   6720
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicStimMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   6240
      Picture         =   "Form1.frx":29B02
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   448
      TabIndex        =   35
      Top             =   4920
      Width           =   6720
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicTreasure 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   2520
      Picture         =   "Form1.frx":53604
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   34
      Top             =   4920
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicTreasureMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   2520
      Picture         =   "Form1.frx":5527E
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   33
      Top             =   5520
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicBottomMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   120
      Picture         =   "Form1.frx":56EF8
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   32
      Top             =   7680
      Width           =   6135
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicBottom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   120
      Picture         =   "Form1.frx":5F0BE
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   31
      Top             =   7200
      Width           =   6135
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicBubble 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5040
      Picture         =   "Form1.frx":67284
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   30
      Top             =   3960
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicBubbleMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5040
      Picture         =   "Form1.frx":67776
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   29
      Top             =   4320
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicSmokeMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   2400
      Picture         =   "Form1.frx":67C68
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   28
      Top             =   4200
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicPMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      Picture         =   "Form1.frx":67E12
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   27
      Top             =   4320
      Width           =   570
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   6240
      Picture         =   "Form1.frx":68608
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   26
      Top             =   120
      Width           =   2775
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicShipMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   6240
      Picture         =   "Form1.frx":77BB6
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   25
      Top             =   1560
      Width           =   2775
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Index           =   6
      Left            =   5280
      Picture         =   "Form1.frx":87164
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   24
      Top             =   5760
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrassMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Index           =   6
      Left            =   2040
      Picture         =   "Form1.frx":87DD6
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   23
      Top             =   6120
      Width           =   300
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrassMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   5
      Left            =   1680
      Picture         =   "Form1.frx":88A48
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   22
      Top             =   6240
      Width           =   345
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrassMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   4
      Left            =   1440
      Picture         =   "Form1.frx":89582
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   21
      Top             =   6480
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrassMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   3
      Left            =   1200
      Picture         =   "Form1.frx":89AA4
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   20
      Top             =   5760
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrassMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   960
      Picture         =   "Form1.frx":8A8F6
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   19
      Top             =   5880
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrassMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   885
      Index           =   1
      Left            =   720
      Picture         =   "Form1.frx":8B568
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   18
      Top             =   6000
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrassMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1770
      Index           =   0
      Left            =   480
      Picture         =   "Form1.frx":8C1A6
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   17
      Top             =   5160
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicHelicopter2Mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   3480
      Picture         =   "Form1.frx":8D9E0
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   16
      Top             =   4440
      Width           =   885
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicHelicopterMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   2640
      Picture         =   "Form1.frx":8EDD2
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   15
      Top             =   4440
      Width           =   885
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicSharkMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      Picture         =   "Form1.frx":901C4
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   138
      TabIndex        =   14
      Top             =   4560
      Width           =   2070
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   5
      Left            =   4920
      Picture         =   "Form1.frx":951A6
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   13
      Top             =   6000
      Width           =   345
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Index           =   4
      Left            =   4680
      Picture         =   "Form1.frx":95CE0
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   12
      Top             =   6120
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   3
      Left            =   4440
      Picture         =   "Form1.frx":96202
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   11
      Top             =   5520
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   4200
      Picture         =   "Form1.frx":97054
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   10
      Top             =   5640
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   885
      Index           =   1
      Left            =   3960
      Picture         =   "Form1.frx":97CC6
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   9
      Top             =   5760
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicGrass 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1770
      Index           =   0
      Left            =   3720
      Picture         =   "Form1.frx":98904
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   8
      Top             =   4800
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicShark 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      Picture         =   "Form1.frx":9A13E
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   138
      TabIndex        =   7
      Top             =   3960
      Width           =   2070
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      Picture         =   "Form1.frx":9F120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   5
      Top             =   3960
      Width           =   570
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicHelicopter2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   3480
      Picture         =   "Form1.frx":9F916
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   4
      Top             =   3960
      Width           =   885
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicHelicopter 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   2640
      Picture         =   "Form1.frx":A0D08
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   2
      Top             =   3960
      Width           =   885
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicSmoke 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   2400
      Picture         =   "Form1.frx":A20FA
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   1
      Top             =   3960
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicDisp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF00FF&
      Height          =   3720
      Left            =   0
      Picture         =   "Form1.frx":A22A4
      ScaleHeight     =   248
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click to start!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   1665
         TabIndex        =   6
         Top             =   0
         Width           =   2805
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Best: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   3360
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim FirstTime As Boolean

Private Type HelicopterType 'properties of the helicopter
    Y As Single
    StartX As Single
    Speed As Single
    Width As Integer
    Height As Integer
    Range As Single
End Type

Private Type SmokeType 'properties of the smoke
    Y As Single
    X As Single
    Width As Integer
    Height As Integer
End Type

Private Type GrassType 'properties of the smokgrass
    Y As Single
    X As Single
    Width As Integer
    Height As Integer
    Pic As Integer
End Type

Private Type BoxType 'properties of the boxes
    Y As Single
    X As Single
    Width As Integer
    Height As Integer
End Type

Private Type BonusType 'properties of the bonus
    Y As Single
    X As Single
    Width As Integer
    Height As Integer
    Show As Long
End Type

Private Type TreasureType 'properties of the treasure
    Y As Single
    X As Single
    Width As Integer
    Height As Integer
    Show As Long
End Type

Private Type StimType 'properties of the stim
    Y As Single
    X As Single
    Width As Integer
    Height As Integer
    Show As Long
End Type

Private Type ShipType 'properties of the boxes
    Y As Single
    X As Single
    Width As Integer
    Height As Integer
    Show As Long
End Type

Private Type BubbleType 'properties of the bubble
    Y As Single
    X As Single
    Draw As Boolean
    Show As Integer
End Type

Dim Helicopter As HelicopterType
Dim Smoke(9) As SmokeType
Dim Box(1) As BoxType
Dim Bonus As BonusType
Dim Grass(6) As GrassType
Dim Ship As ShipType
Dim Bubble As BubbleType
Dim Treasure As TreasureType
Dim Stim As StimType

Dim GoingUp As Boolean 'check if mouse is clicked
Dim Running As Boolean 'check if game is on
Dim Crash As Boolean 'check if helicopter has crashed
Dim DrawBonus As Boolean
Dim Paused As Boolean
Dim MaxRange As Long
Dim TempFrame As Integer
Dim TempFrame2 As Single
Dim TempFrame3 As Integer
Dim TempFrame4 As Single
Dim TempFrameStim As Integer
Dim TempFrameStim2 As Single
Dim TempFrameBottom As Integer
Private Sub Form_Activate()
Dim i As Integer
Helicopter.StartX = PicDisp.ScaleWidth / 2 - 80 'set the startx of the helicopter
Helicopter.Y = 0 'set the Y to center of the screen
For i = 0 To 9
    'set height and width of smoke
    Smoke(i).Width = PicSmoke.ScaleWidth
    Smoke(i).Height = PicSmoke.ScaleHeight
    'set the X position of the smoke
    Smoke(i).X = (Helicopter.StartX - Helicopter.Width / 2) / 10 * (i + 1) - Smoke(i).Width / 2
    'move the smoke out of the screen so it will not be visible
    Smoke(i).Y = -(Smoke(i).Width * 2)
Next

For i = 0 To PicGrass.UBound
    'set height and width of smoke
    Grass(i).Pic = Rnd * PicGrass.UBound
    Grass(i).Width = PicGrass(Grass(i).Pic).ScaleWidth
    Grass(i).Height = PicGrass(Grass(i).Pic).ScaleHeight
    'set the X position of the smoke
    Grass(i).X = PicDisp.ScaleWidth / (PicGrass.UBound + 1) * (i + 1)
    'move the smoke out of the screen so it will not be visible
    Grass(i).Y = PicDisp.ScaleHeight - PicGrass(Grass(i).Pic).ScaleHeight
Next
    
    Helicopter.Width = PicHelicopter.ScaleWidth
    Helicopter.Height = PicHelicopter.ScaleHeight

For i = 0 To 1
    Box(i).Width = PicShark.ScaleWidth
    Box(i).Height = PicShark.ScaleHeight
    Randomize
    Box(i).Y = Rnd * (PicDisp.ScaleHeight - Box(i).Height) 'move box to random Y position
Next
    Box(0).X = PicDisp.ScaleWidth 'move box out of screen
    Box(1).X = -Box(1).Width
    
Bonus.Width = PicP.ScaleWidth / 2
Bonus.Height = PicP.ScaleWidth / 2
Bonus.X = PicDisp.ScaleWidth + 25
Bonus.Y = Rnd * (PicDisp.ScaleHeight - Bonus.Height)
Bonus.Show = Rnd * 20000 + 5000 + GetTickCount
Helicopter.Speed = 4

Ship.Width = PicShip.ScaleWidth
Ship.Height = PicShip.ScaleHeight
Ship.X = PicDisp.ScaleWidth
Ship.Y = PicDisp.ScaleHeight - Ship.Height - 5
Ship.Show = Rnd * 20000 + 5000 + GetTickCount

Bubble.Draw = False
Bubble.Show = 0

Treasure.Width = PicTreasure.ScaleWidth
Treasure.Height = PicTreasure.ScaleHeight
Treasure.Show = Rnd * 20000 + 5000 + GetTickCount
Treasure.X = PicDisp.ScaleWidth
Treasure.Y = PicDisp.ScaleHeight - Treasure.Height - 2

Stim.Show = Rnd * 20000 + 5000 + GetTickCount
Stim.X = -(PicStim.ScaleWidth / 4)
Stim.Y = PicDisp.ScaleHeight - PicStim.ScaleHeight - 60

SetStars
End Sub

Private Sub Main()
Dim TempTime As Long
Dim TempTime2 As Long
Dim TempTime3 As Long
Dim X As Integer
Dim Y As Integer
Dim i As Integer

    While Running = True
        If TempTime < GetTickCount Then 'timer
            TempTime = GetTickCount + 20 'set timer to 20 ms
            
            If TempTime3 < GetTickCount Then
                TempTime3 = GetTickCount + 5000
                PlayTheWav "water3.wav"
            End If
            
            Helicopter.Range = Helicopter.Range + 0.5 'change range points
            If Helicopter.Range > MaxRange Then MaxRange = Helicopter.Range
            Label1.Caption = "Best: " & MaxRange
            
            If GoingUp = True Then 'if mouse is clicked, move up
                Helicopter.Speed = Helicopter.Speed + 0.4
                If Helicopter.Speed > 6 Then Helicopter.Speed = 6
                Helicopter.Y = Helicopter.Y - Helicopter.Speed
            Else 'else move down
                Helicopter.Speed = Helicopter.Speed - 0.4
                If Helicopter.Speed < -6 Then Helicopter.Speed = -6
                Helicopter.Y = Helicopter.Y - Helicopter.Speed
            End If
            
            PicDisp.Cls 'clear picture before drawing
            
            If Ship.X + Ship.Width <= 0 Then
                Ship.Show = Rnd * 20000 + 5000 + GetTickCount
                Ship.X = PicDisp.ScaleWidth
            End If
            If Ship.X + Ship.Width > 0 And Ship.Show < GetTickCount Then
                Ship.X = Ship.X - 1
                TransBlt PicShip.hdc, 0, 0, PicShipMask.hdc, 0, 0, PicDisp.hdc, Ship.X, Ship.Y, Ship.Width, Ship.Height
            End If
            
            If Stim.X >= PicDisp.ScaleWidth Then
                Stim.Show = Rnd * 20000 + 5000 + GetTickCount
                Stim.X = -(PicStim.ScaleWidth / 4)
            End If
            If Stim.X < PicDisp.ScaleWidth And Stim.Show < GetTickCount Then
                Stim.X = Stim.X + 3
                DrawStim
            End If
            
            ReDrawStars
            
            DrawBottom
            
            For i = 0 To PicGrass.UBound
                Grass(i).X = Grass(i).X - 2
                
                If Grass(i).X + Grass(i).Width < 0 Then
                    Grass(i).Pic = Rnd * PicGrass.UBound
                    Grass(i).Width = PicGrass(Grass(i).Pic).ScaleWidth
                    Grass(i).Height = PicGrass(Grass(i).Pic).ScaleHeight
                    Grass(i).Y = PicDisp.ScaleHeight - PicGrass(Grass(i).Pic).ScaleHeight
                    Grass(i).X = PicDisp.ScaleWidth
                End If
                
                TransBlt PicGrass(Grass(i).Pic).hdc, 0, 0, PicGrassMask(Grass(i).Pic).hdc, 0, 0, PicDisp.hdc, Grass(i).X, Grass(i).Y, Grass(i).Width, Grass(i).Height
            Next
            
            If Treasure.X + Treasure.Width <= 0 Then
                Treasure.Show = Rnd * 20000 + 5000 + GetTickCount
                Treasure.X = PicDisp.ScaleWidth
            End If
            If Treasure.X + Treasure.Width > 0 And Treasure.Show < GetTickCount Then
                Treasure.X = Treasure.X - 2
                TransBlt PicTreasure.hdc, 0, 0, PicTreasureMask.hdc, 0, 0, PicDisp.hdc, Treasure.X, Treasure.Y, Treasure.Width, Treasure.Height
            End If
            
            'Bonus
            
            If Bonus.X + Bonus.Width <= 0 Then
                Bonus.Show = Rnd * 20000 + 5000 + GetTickCount
                Bonus.X = PicDisp.ScaleWidth + 25
            End If
            If Bonus.X + Bonus.Width > 0 And Bonus.Show < GetTickCount Then
                Bonus.X = Bonus.X - 6
                'TransBlt PicP.hdc, 0, 0, PicPMask.hdc, 0, 0, PicDisp.hdc, Bonus.X, Bonus.Y, Bonus.Width, Bonus.Height
                DrawWorm
            End If
            
            'boxes
            For i = 0 To 1
                Box(i).X = Box(i).X - 8 'move box to the left
                
                If Box(0).X < 50 And Box(0).X > 45 Then 'if a box is here, move the other box to the end of the screen (it appears to be another box)
                    Box(1).X = PicDisp.ScaleWidth
                    Box(1).Y = Rnd * (PicDisp.ScaleHeight - Box(1).Height)
                End If
                
                If Box(1).X < 50 And Box(1).X > 45 Then 'if a box is here, move the other box to the end of the screen (it appears to be another box)
                    Box(0).X = PicDisp.ScaleWidth
                    Box(0).Y = Rnd * (PicDisp.ScaleHeight - Box(0).Height)
                End If
                
                TransBlt PicShark.hdc, 0, 0, PicSharkMask.hdc, 0, 0, PicDisp.hdc, Box(i).X, Box(i).Y, Box(i).Width, Box(i).Height
            Next
            
            'smoke
            For i = 0 To 9
                If Smoke(i).X < 0 Then 'if smoke is out of screen, move it to the helicopters position (it appears to be new smoke)
                    Smoke(i).X = Helicopter.StartX - Helicopter.Width / 2
                    Smoke(i).Y = PicDisp.ScaleHeight / 2 + Helicopter.Y
                End If
                Smoke(i).X = Smoke(i).X - 4 'move smoke to the left
                'draw smoke
                TransBlt PicSmoke.hdc, 0, 0, PicSmokeMask.hdc, 0, 0, PicDisp.hdc, Smoke(i).X - Smoke(i).Width / 2, Smoke(i).Y - Smoke(i).Height / 2, PicSmoke.ScaleWidth, PicSmoke.ScaleHeight
            Next

            'draw helicopter
            If GoingUp = True Then
                TransBlt PicHelicopter2.hdc, 0, 0, PicHelicopter2Mask.hdc, 0, 0, PicDisp.hdc, Helicopter.StartX - Helicopter.Width / 2, Helicopter.Y - Helicopter.Height / 2 + PicDisp.ScaleHeight / 2, Helicopter.Width, Helicopter.Height
            Else
                TransBlt PicHelicopter.hdc, 0, 0, PicHelicopterMask.hdc, 0, 0, PicDisp.hdc, Helicopter.StartX - Helicopter.Width / 2, Helicopter.Y - Helicopter.Height / 2 + PicDisp.ScaleHeight / 2, Helicopter.Width, Helicopter.Height
            End If
            
            If Bubble.Draw = False Then
                Randomize
                Bubble.Show = Rnd * 50
                If Bubble.Show = 1 Then
                    Bubble.Draw = True
                    Bubble.X = Rnd * (PicDisp.ScaleWidth) - PicBubble.ScaleWidth
                    Bubble.Y = PicDisp.ScaleHeight
                    'PlayTheWav "bubble.wav"
                End If
            End If
            
            If Bubble.Draw = True Then
                Bubble.Y = Bubble.Y - 4
                Bubble.X = Bubble.X - ((PicDisp.ScaleHeight - Bubble.Y) / 40 + 2)
                TransBlt PicBubble.hdc, 0, 0, PicBubbleMask.hdc, 0, 0, PicDisp.hdc, Bubble.X, Bubble.Y, PicBubble.ScaleWidth, PicBubble.ScaleHeight
                If Bubble.Y < 0 Then
                    Bubble.Draw = False
                End If
            End If
            
            'set printing position, and print range
            PicDisp.CurrentX = 0
            PicDisp.FontSize = 12
            PicDisp.FontBold = True
            PicDisp.CurrentY = PicDisp.ScaleHeight - 25
            PicDisp.Print "Range: " & Round(Helicopter.Range)
            
            If TempTime2 > 0 Then
                TempTime2 = TempTime2 - 20
            Else
                Label2.Visible = False
            End If
            
            PicDisp.Refresh
            
            'collision detecting
            
            'bonus
            If Helicopter.Y + Helicopter.Height / 2 + PicDisp.ScaleHeight / 2 >= Bonus.Y And Helicopter.Y - Helicopter.Height / 2 + PicDisp.ScaleHeight / 2 <= Bonus.Y + Bonus.Height And Helicopter.StartX + Helicopter.Width / 2 >= Bonus.X And Helicopter.StartX - Helicopter.Width / 2 <= Bonus.X + Bonus.Width Then 'check if the helicopter is inside a Bonus
                Bonus.Show = Rnd * 20000 + 5000 + GetTickCount
                Bonus.X = -50
                Label2.Caption = "Worm!"
                PlayTheWav "runekey.wav"
                TempTime3 = GetTickCount + 1000
                Label2.Visible = True
                TempTime2 = 1000
                Helicopter.Range = Helicopter.Range + 200
            End If
            
            'box
            For i = 0 To 1
                'If Helicopter.Y + Helicopter.Height / 2 + PicDisp.ScaleHeight / 2 >= Box(i).Y And Helicopter.Y - Helicopter.Height / 2 + PicDisp.ScaleHeight / 2 <= Box(i).Y + Box(i).Height And Helicopter.StartX + Helicopter.Width / 2 >= Box(i).X And Helicopter.StartX - Helicopter.Width / 2 <= Box(i).X + Box(i).Width Then 'check if the helicopter is inside a box
                If Helicopter.Y + Helicopter.Height / 2 + PicDisp.ScaleHeight / 2 - 4 >= Box(i).Y + 4 And Helicopter.Y - Helicopter.Height / 2 + PicDisp.ScaleHeight / 2 + 4 <= Box(i).Y + Box(i).Height And Helicopter.StartX + Helicopter.Width / 2 - 4 >= Box(i).X And Helicopter.StartX - Helicopter.Width / 2 + 4 <= Box(i).X + Box(i).Width Then 'check if the helicopter is inside a box
                    PlayTheWav "drone6.wav"
                    Running = False 'stop running
                    Crash = True 'start crash
                    Crashing 'draw some crash stuff
                End If
                'botton
                If Helicopter.Y + Helicopter.Height / 2 + PicDisp.ScaleHeight / 2 >= PicDisp.ScaleHeight Then
                    PlayTheWav "h2ohit1.wav"
                    Running = False
                    Crash = True
                    Crashing
                End If
                'top
                If Helicopter.Y - Helicopter.Height / 2 + PicDisp.ScaleHeight / 2 <= 0 Then
                    PlayTheWav "h2ohit1.wav"
                    Running = False
                    Crash = True
                    Crashing
                End If
            Next
        End If
        DoEvents 'let windows handle other stuff
    Wend
    
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicDisp_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicDisp_MouseUp Button, Shift, X, Y
End Sub

Private Sub PicDisp_Click()
    If Running = False And Crash = False And Paused = False And FirstTime = False Then  'if helicopter is stopped and the program is not running the crashing code
        Label2.Visible = False
        Form_Activate 'reset all properties
        Running = True 'start running
        Main 'do the main code
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Running = False
    Crash = False
    FirstTime = False
End Sub

Private Sub Form_Terminate()
    Running = False
    Crash = False
    FirstTime = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Running = False
    Crash = False
    FirstTime = False
End Sub

Private Sub PicDisp_DblClick()
    GoingUp = True
End Sub

Private Sub PicDisp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyP And Running = True Then
    Running = False
    Label2.Caption = "Paused!"
    Label2.Visible = True
    Paused = True
ElseIf KeyCode = vbKeyP And Running = False Then
    Label2.Visible = False
    Running = True
    Paused = False
    Main
End If
End Sub

Private Sub PicDisp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GoingUp = True
End Sub

Private Sub PicDisp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GoingUp = False
End Sub

Private Sub Crashing()
Dim TempTime As Long

Helicopter.Range = 0 'reset range
Label2.Caption = "You Died!"
Label2.Visible = True

'only pause for 1 sec
    While Crash = True
        If TempTime < GetTickCount Then
            If FirstTime = True Then Crash = False
            TempTime = GetTickCount + 1000
            If FirstTime = False Then FirstTime = True
        End If
        'DoEvents
    Wend
FirstTime = False
Label2.Caption = "Click to start!"
End Sub

Private Sub TransBlt(hPicture As Long, iFromX As Integer, iFromY As Integer, _
  hMask As Long, iFromMaskX As Integer, iFromMaskY As Integer, hTo As Long, _
  X As Single, Y As Single, iWidth As Integer, iHeight As Integer)
  
  BitBlt hTo, X, Y, iWidth, iHeight, hMask, iFromX, iFromY, vbSrcAnd
  BitBlt hTo, X, Y, iWidth, iHeight, hPicture, iFromX, iFromY, vbSrcPaint
End Sub


Private Sub DrawBottom2()
TransBlt PicBottom.hdc, TempFrame * PicDisp.ScaleWidth, 0, PicBottomMask.hdc, TempFrame * PicDisp.ScaleWidth, 0, PicDisp.hdc, 0, PicDisp.ScaleHeight - PicBottom.ScaleHeight, PicDisp.ScaleWidth, PicBottom.ScaleHeight
TempFrame2 = TempFrame2 + 0.2
TempFrame = Round(TempFrame2)
If TempFrame2 > 39 Then TempFrame2 = 0
End Sub

Private Sub DrawWorm()
TransBlt PicP.hdc, TempFrame3 * PicP.ScaleWidth / 2, 0, PicPMask.hdc, TempFrame3 * (PicP.ScaleWidth / 2), 0, PicDisp.hdc, Bonus.X, Bonus.Y, Bonus.Width, Bonus.Height
TempFrame4 = TempFrame4 + 0.1
TempFrame3 = Round(TempFrame4)
If TempFrame4 > 1 Then TempFrame4 = 0
End Sub

Private Sub DrawStim()
TransBlt PicStim.hdc, TempFrameStim * PicStim.ScaleWidth / 4, 0, PicStimMask.hdc, TempFrameStim * PicStim.ScaleWidth / 4, 0, PicDisp.hdc, Stim.X, Stim.Y, PicStim.ScaleWidth / 4, PicStim.ScaleHeight
'TransBlt PicStim.hdc, TempFrameStim * PicStim.ScaleWidth / 4, 0, PicStimMask.hdc, TempFrameStim * PicStim.ScaleWidth / 4, 0, PicDisp.hdc, PicDisp.ScaleWidth / 2, PicDisp.ScaleHeight / 2, PicStim.ScaleWidth / 4, PicStim.ScaleHeight
TempFrameStim2 = TempFrameStim2 + 0.1
TempFrameStim = Round(TempFrameStim2)
If TempFrameStim2 > 3 Then TempFrameStim2 = 0
End Sub








Private Sub DrawBottom()
Dim NewXSrc As Integer
Dim NewWidth As Integer
Dim NewX As Single

    TempFrameBottom = TempFrameBottom + 2
    
    NewXSrc = 0
    NewWidth = (PicBottom.ScaleWidth * (TempFrameBottom - 1) / 409)
    NewX = PicBottom.ScaleWidth - PicBottom.ScaleWidth / 409 * (TempFrameBottom - 1) - 1

    TransBlt PicBottom.hdc, NewXSrc, 0, PicBottomMask.hdc, NewXSrc, 0, PicDisp.hdc, NewX, PicDisp.ScaleHeight - PicBottom.ScaleHeight, NewWidth, PicBottom.ScaleHeight
    
    NewXSrc = PicBottom.ScaleWidth / 409 * (TempFrameBottom)
    NewWidth = PicBottom.ScaleWidth - PicBottom.ScaleWidth / 409 * (TempFrameBottom - 1)
    
    TransBlt PicBottom.hdc, NewXSrc, 0, PicBottomMask.hdc, NewXSrc, 0, PicDisp.hdc, 0, PicDisp.ScaleHeight - PicBottom.ScaleHeight, NewWidth, PicBottom.ScaleHeight
    If TempFrameBottom > 409 Then TempFrameBottom = 0
End Sub
