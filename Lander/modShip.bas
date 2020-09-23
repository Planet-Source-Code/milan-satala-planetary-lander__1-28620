Attribute VB_Name = "modShip"
Option Explicit

Private Type typLanderColl
  Angle As Single
  Dist As Single
End Type

Private Type typLanderLeg
  Angle As Single
  Coll As typLanderColl
End Type

Public Type typLander
  X As Single
  Y As Single
  MoveX As Single
  MoveY As Single
  Angle As Single
  RealAngle As Single
  Thrust As Single
  Speed As Single
  HalfX As Single
  HalfY As Single
  Coll(1 To 200) As typLanderColl
  Count As Integer
  Leg(1 To 2) As typLanderLeg
  LegDist As Single
End Type

Public gLander As typLander
Const Gravity As Single = 9.6

Public Sub CalculateShip()
 With gLander
  Dim XUPD As Single, a As Byte
  
  .RealAngle = Int(.Angle / 10) * 10
  
  XUPD = (.Thrust * Sin(.RealAngle * 3.14159 / 180))
  If Abs(.MoveX) > Abs(XUPD) Then
   .MoveX = .MoveX + GS((XUPD - .MoveX) / 5)
  Else
   .MoveX = .MoveX + GS((XUPD - .MoveX))
  End If
  
  .MoveY = .MoveY + GS(.Thrust) * Sin((90 - .RealAngle) * 3.14159 / 180)
  If .MoveY < -120 Then
   .MoveY = -120
  ElseIf .MoveY > 50 Then
   .MoveY = 50
  End If
  
  If .MoveX < 0 Then
   .MoveX = .MoveX + GS(5)
  ElseIf .MoveX > 0 Then
   .MoveX = .MoveX - GS(5)
  End If
  
  If .MoveY < -Gravity * 10 Then
   .MoveY = -Gravity * 10
  Else
   .MoveY = .MoveY - GS(2 * Gravity)
  End If
  
  If .Y > 480 Then .MoveY = -50
  
  .Speed = GetDist(0, 0, .MoveX, .MoveY)
  
  If .X + 50 > MapSize Then
   .MoveX = -Abs(.MoveX)
  ElseIf .X - 1 < 0 Then
   .MoveX = Abs(.MoveX)
  End If
  
  .X = .X + GS(.MoveX)
  .Y = .Y + GS(.MoveY)
  
  For a = 1 To .Count
   If TestColl(.Coll(a)) Then
    ShowSkore "Crash !!!", -100
   End If
  Next a
  
  Dim RealX As Integer
  For a = 1 To 2
   TestLeg a
  Next a
 End With
End Sub

Sub TestLeg(a As Byte)
  Dim RealX As Integer, RealX2 As Integer, RealY As Integer
  Dim Angle As Single, Angle2 As Single
  Dim Dir1 As Integer, Dir2 As Integer
  
  
  With gLander
   Angle = RepairAngle(.Leg(a).Coll.Angle + Int(.Angle / 10) * 10)
   RealX = .X + .HalfX + Posun_X(Angle, .Leg(a).Coll.Dist) + 6
   RealY = .Y - .HalfY - Posun_Y(Angle, .Leg(a).Coll.Dist) - 6
   
   If Terrain(RealX) >= RealY Then
    RealX2 = RealX + Posun_X(RepairAngle(.Leg(a).Angle + .RealAngle), .LegDist)
    Angle = GetAngle(RealX, RealY, RealX2, Terrain(RealX2))
    Dir1 = Abs(GetDir(0, Angle))
    
    If Dir1 < 60 Or .Speed > 30 Then
     ShowSkore "Failed !!!", 0
    Else
     
     Angle2 = RepairAngle(.RealAngle + .Leg(a).Angle)
     Dir2 = Abs(GetDir(Angle, Angle2))
     If Dir2 > 30 Then ShowSkore "Failed !!!", 90 - Dir1 Else ShowSkore "Success !!!", Dir1 + (30 - Dir2) * 2
  
    End If
   End If
  End With
End Sub

Sub ShowSkore(Caption As String, Bonus As Integer)
  frmGame.lblStatus = Caption
  frmGame.lblTime = Int(GameTime) & " sec"
  frmGame.lblTimeScore = Int(60 - GameTime)
  frmGame.lblBonus = Bonus
  frmGame.lblTotal = Bonus + Int(60 - GameTime)
  frmGame.fraStatus.Visible = True
  frmGame.cmdNew.SetFocus
  EndGame = True
End Sub

Function TestColl(Coll As typLanderColl) As Boolean
  Dim RealX As Integer, RealY As Integer
  Dim Angle As Single
  
  
  With gLander
   Angle = RepairAngle(Coll.Angle + .RealAngle)
   RealX = .X + .HalfX + Posun_X(Angle, Coll.Dist) + 6
   RealY = .Y - .HalfY - Posun_Y(Angle, Coll.Dist) + 6
   
   If Terrain(RealX) > RealY Then TestColl = True
  End With
End Function

Function Posun_X(ByVal Uhol As Integer, ByVal Rychlost As Integer) As Double
  Posun_X = Rychlost * Sin(Uhol * 3.14159 / 180)
End Function

Function Posun_Y(ByVal Uhol As Integer, ByVal Rychlost As Integer) As Double
  Posun_Y = -Rychlost * Sin((90 - Uhol) * 3.14159 / 180)
End Function

