VERSION 5.00
Begin VB.Form frmGame 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planet Lander by Mental Soft"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSettings 
      Caption         =   "Settings"
      Height          =   2895
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
      Begin VB.HScrollBar ScrSize 
         Height          =   255
         Left            =   840
         Max             =   30
         Min             =   1
         TabIndex        =   4
         Top             =   600
         Value           =   1
         Width           =   1575
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Simulation"
         Default         =   -1  'True
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   3120
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label2 
         Caption         =   "(Number of screens, default is 3)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblSize 
         Alignment       =   2  'Center
         Caption         =   "3"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   840
         X2              =   3000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Map size"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Staus"
      Height          =   3135
      Left            =   3240
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmdEnd 
         Caption         =   "End"
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New game"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblTotal 
         Caption         =   "50"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblBonus 
         Caption         =   "200"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTimeScore 
         Caption         =   "100"
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Total score:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Landing bonus:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Time score:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblTime 
         Caption         =   "30 s"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "Success !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2820
      End
   End
   Begin VB.PictureBox picShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   4440
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblLoad 
      Alignment       =   2  'Center
      Caption         =   "Loading ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   9255
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
  End
End Sub

Private Sub cmdNew_Click()
  fraStatus.Visible = False
  fraSettings.Visible = True
  Cls
End Sub

Private Sub Form_Load()
  Show
  DoEvents
  Randomize Timer
  
  Dim DDColor As DDCOLORKEY
  Dim PrimaryDesc As DDSURFACEDESC2
  Dim BackDesc As DDSURFACEDESC2
  Dim TMPDesc As DDSURFACEDESC2
  
  Set DirX = New DirectX7
  
  Set DDraw = DirX.DirectDrawCreate("")
  DDraw.SetCooperativeLevel frmGame.hWnd, DDSCL_NORMAL
  
  DirX.GetWindowRect frmGame.hWnd, PrimaryRect
  
  DDColor.low = 0
  DDColor.high = 0
  
  PrimaryDesc.lFlags = DDSD_CAPS
  PrimaryDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
  PrimaryRect.Top = PrimaryRect.Top + 22
  PrimaryRect.Left = PrimaryRect.Left + 4
  PrimaryRect.Right = PrimaryRect.Right - 4
  PrimaryRect.Bottom = PrimaryRect.Bottom - 4
  PrimaryDesc.lWidth = PrimaryRect.Right - PrimaryRect.Left
  PrimaryDesc.lHeight = PrimaryRect.Bottom - PrimaryRect.Top
    
  Set sPrimary = DDraw.CreateSurface(PrimaryDesc)
    
  BackDesc.lFlags = 7
  BackDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
  BackDesc.lWidth = 640
  BackDesc.lHeight = 480
  
  BackRect.Right = 640
  BackRect.Bottom = 480
  
  TMPDesc.lFlags = 7
  TMPDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
  TMPDesc.lWidth = 0
  TMPDesc.lHeight = 0
  
  Set sBack = DDraw.CreateSurface(BackDesc)
  Set sShip = DDraw.CreateSurfaceFromFile(App.Path & "\Data\Ship2.bmp", TMPDesc)
  sShip.SetColorKey DDCKEY_SRCBLT, DDColor
  
  picShip.Picture = LoadPicture(App.Path & "\Data\Ship.bmp")
  GetShip
  
  lblLoad.Visible = False
  fraSettings.Visible = True
  ScrSize.Value = 5
End Sub

Sub SetDefault()
  gLander.Angle = 0
  gLander.MoveX = 0
  gLander.MoveY = 0
  gLander.Thrust = 0
  gLander.Y = 200
  gLander.X = MapSize / 2
  GameTime = 0
  gCamX = gLander.X - 320
  gMoveX = 0
End Sub

Private Sub GetShip()
  Dim a As Integer, b As Integer, c As Integer, d As Integer
  Dim IsCorner As Boolean, Leg As Byte
  Dim LegPos(1 To 2) As POINTAPI
  
  With gLander
   .HalfX = picShip.ScaleWidth / 2
   .HalfY = picShip.ScaleHeight / 2
   
   For a = 0 To picShip.ScaleWidth
   For b = 0 To picShip.ScaleHeight
    IsCorner = False
    If picShip.Point(a, b) = RGB(255, 0, 255) Then
     Leg = Leg + 1
     LegPos(Leg).X = a
     LegPos(Leg).Y = b
     .Leg(Leg).Coll.Angle = GetAngle(.HalfX, .HalfY, a, b)
     .Leg(Leg).Coll.Dist = GetDist(.HalfX, .HalfY, a, b)
    ElseIf picShip.Point(a, b) > 0 Then
    
     For c = a - 1 To a + 1
     If IsCorner = True Then Exit For
     For d = b - 1 To b + 1
      If picShip.Point(c, d) = 0 Then
       .Count = .Count + 1
       .Coll(.Count).Angle = GetAngle(.HalfX, .HalfY, a, b)
       .Coll(.Count).Dist = GetDist(.HalfX, .HalfY, a, b)
       IsCorner = True
       Exit For
      End If
     Next d
     Next c
    
    End If
   Next b
   Next a
   
   .Leg(1).Angle = GetAngle(LegPos(1).X, LegPos(1).Y, LegPos(2).X, LegPos(2).Y)
   .Leg(2).Angle = RepairAngle(.Leg(1).Angle + 180)
   .LegDist = GetDist(LegPos(1).X, LegPos(1).Y, LegPos(2).X, LegPos(2).Y)
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub ScrSize_Change()
  lblSize = ScrSize.Value
End Sub

Private Sub cmdStart_Click()
  
  fraSettings.Visible = False
  Dim TMPDesc As DDSURFACEDESC2
  
  lblLoad.Visible = True
  fraSettings.Visible = False
  lblLoad = "Generating Terrain ..."
  DoEvents
  
  TMPDesc.lFlags = 7
  TMPDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
  
  Size = ScrSize.Value
  MapSize = Size * 640
  TMPDesc.lWidth = 640
  ReDim Terrain(MapSize)
  ReDim sTerrain(Size)
  TMPDesc.lHeight = 300
  TerrainRect.Right = 640
  TerrainRect.Bottom = 300
  
  For a = 1 To Size
   Set sTerrain(a) = DDraw.CreateSurface(TMPDesc)
   sTerrain(a).setDrawWidth 1
   sTerrain(a).SetForeColor vbGreen
   sTerrain(a).BltColorFill TerrainRect, 0
  Next a
  
  GenerateTerrain
  
  lblLoad.Visible = False
  OldTime = GetTickCount
  SetDefault
  
  Do Until EndGame
   GameTick
  Loop
  EndGame = False
  
End Sub
