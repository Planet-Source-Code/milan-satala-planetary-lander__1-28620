Attribute VB_Name = "modGrafika"
Public DirX As DirectX7
Public DDraw As DirectDraw7

Public sPrimary As DirectDrawSurface7
Public PrimaryRect As RECT
Public sBack As DirectDrawSurface7
Public BackRect As RECT

Public sTerrain() As DirectDrawSurface7
Public TerrainRect As RECT

Public sShip As DirectDrawSurface7
Public ShipRect As RECT

Public Size As Long

Public gMoveX As Single, gCamX As Single, OnScrX As Single

Public Sub Render()
  Dim TRect1 As RECT, TRect2 As RECT
  Dim SRect As RECT
  Dim X As Integer, Y As Integer, Angle As Single
  Dim T1 As Byte, T2 As Byte
  
  sBack.BltColorFill BackRect, 0
  
  TRect1.Bottom = 300
  TRect2.Bottom = 300
  Y = Int(gLander.RealAngle / 90)
  SRect.Left = gLander.RealAngle * 5 - Y * 450
  SRect.Top = Y * 42
  
  SRect.Bottom = SRect.Top + 42
  SRect.Right = SRect.Left + 42
  T1 = GetTerrainNum(gCamX)
  T2 = GetTerrainNum(gCamX + 640)
  
  TRect1.Left = gCamX - (T1 - 1) * 640
  TRect1.Right = 640
  TRect2.Left = 0
  TRect2.Right = TRect1.Left
  
  sBack.BltFast 0, 180, sTerrain(T1), TRect1, DDBLTFAST_WAIT
  sBack.BltFast 640 - TRect1.Left, 180, sTerrain(T2), TRect2, DDBLTFAST_WAIT
  
  With gLander
   OnScrX = .X - gCamX
   
   sBack.BltFast OnScrX, 480 - gLander.Y, sShip, SRect, DDBLTFAST_SRCCOLORKEY
   
   If .Thrust Then
    
    Angle = RepairAngle(180 + .RealAngle)
    X = OnScrX + .HalfX + Posun_X(Angle, 10) + 4
    Y = 480 - (.Y - .HalfY - Posun_Y(Angle, 10) - 6)
    If Rnd * 100 > 50 Then
     sBack.SetForeColor RGB(255, 0, 0)
    Else
     sBack.SetForeColor RGB(255, 255, 0)
    End If
    sBack.setDrawWidth 5
    sBack.DrawLine X, Y, X + Posun_X(Angle, 5), Y - Posun_Y(Angle, -5)
   End If
  End With
  
  sBack.SetForeColor vbWhite
  sBack.DrawText 10, 10, "FPS: " & GetFPS, False
  DirX.GetWindowRect frmGame.hWnd, PrimaryRect

  PrimaryRect.Top = PrimaryRect.Top + 22
  PrimaryRect.Left = PrimaryRect.Left + 4
  PrimaryRect.Right = PrimaryRect.Right - 4
  PrimaryRect.Bottom = PrimaryRect.Bottom - 4
 
  sPrimary.Blt PrimaryRect, sBack, BackRect, DDBLT_WAIT
  
End Sub
