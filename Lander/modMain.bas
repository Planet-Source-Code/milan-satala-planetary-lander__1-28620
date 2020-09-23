Attribute VB_Name = "modMain"
Option Explicit

Public MapSize As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Const KEY_TOGGLED As Integer = &H1
Private Const KEY_DOWN As Integer = &H1000

Dim LastState(0 To 255) As Byte

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public OldTime As Long, DrawTime As Long, GameSpeed As Single
Public Terrain() As Integer
Public EndGame As Boolean, GameTime As Single

Public StickLeft As Boolean, StickRight As Boolean, StickUp As Boolean
Dim OldTick As Long, OldFPS As Integer, CurrentFPS As Integer

Public Sub GameTick()
  Dim Time As Long
  Time = GetTickCount - OldTime
  Dim a, b
  b = GetTickCount
  OldTime = GetTickCount
  
  If Time > 1000 Then Exit Sub
  
  GameSpeed = Time / 1000
  
  If GetTickCount - DrawTime > 10 Then
   DrawTime = GetTickCount
   DoEvents
   Render
  End If
  
  GameTime = GameTime + GS(1)
  CalculateShip
  
  If KeyState(vbKeyEscape) = 1 Then End
  
  StickRight = False
  StickLeft = False
  StickUp = False
   
  If KeyState(vbKeyRight) = 2 Then
   StickRight = True
  End If
  If KeyState(vbKeyLeft) = 2 Then
   StickLeft = True
  End If
 
  If KeyState(vbKeyUp) = 2 Then
   StickUp = True
  End If
  
  
  If StickRight Then
   gLander.Angle = RepairAngle(gLander.Angle + GS(150))
  End If
  If StickLeft = True Then
   gLander.Angle = RepairAngle(gLander.Angle - GS(150))
  End If
  
  If StickUp = True Then
   gLander.Thrust = 100
  Else
   gLander.Thrust = 0
  End If
  
  If OnScrX > 400 Then gMoveX = 100
  If OnScrX < 240 Then gMoveX = -100
  
  If gMoveX > 0 Then
   gMoveX = gMoveX - GS(50)
  ElseIf gMoveX < 0 Then
   gMoveX = gMoveX + GS(50)
  End If
  gCamX = gCamX + GS(gMoveX)
  
  If gCamX + 640 > MapSize Then
   gCamX = MapSize - 640
   gMoveX = 0
  ElseIf gCamX < 0 Then
   gCamX = 0
   
   gMoveX = 0
  End If
End Sub

Public Function GS(ByVal Var As Single) As Single
  GS = Var * GameSpeed
End Function

Public Function KeyState(ByVal m_Key As Byte) As Byte
 KeyState = 0
 If (GetKeyState(m_Key) And KEY_DOWN) Then KeyState = 1
 If LastState(m_Key) > 0 And KeyState = 1 Then KeyState = 2
 If LastState(m_Key) = 3 Then LastState(m_Key) = 0
 If LastState(m_Key) > 0 And KeyState = 0 Then KeyState = 3
 LastState(m_Key) = KeyState
End Function

Public Sub GenerateTerrain()
  Dim a As Byte, T As Byte, X As Long
  Dim SizeLeft As Long, Size As Long, Tall As Long, NextTall As Long
  Dim CurX As Long, CurTall As Single, Move As Single
  
  SizeLeft = MapSize
  NextTall = Rnd * 300
  CurTall = Rnd * 300
  
  Do Until SizeLeft <= 0
   Size = 0
   Do Until Size > 20
    Size = Rnd * 100
   Loop
   If SizeLeft - 2 * Size < 0 Then
   Size = SizeLeft / 2
   End If
   
   SizeLeft = SizeLeft - 2 * Size
   
   Tall = NextTall
   NextTall = 0
   Do Until NextTall > 10 And NextTall < 300
    NextTall = CurTall - 100 + Rnd * 200
   Loop
   
   Move = (Tall - CurTall) / Size
   
   For a = 1 To Size
    CurX = CurX + 1
    T = GetTerrainNum(CurX)
    
    CurTall = CurTall + Move
    X = CurX - (T - 1) * 640
    sTerrain(T).DrawLine X, 300 - CurTall, X, 300
    Terrain(CurX) = CurTall
   Next a
   
   Move = ((NextTall / 2) - CurTall) / Size
   
   For a = 1 To Size
    CurX = CurX + 1
    T = GetTerrainNum(CurX)
    
    CurTall = CurTall + Move
    X = CurX - (T - 1) * 640
    sTerrain(T).DrawLine X, 300, X, 300 - CurTall
    Terrain(CurX) = CurTall
   Next a
  Loop
End Sub

Public Function GetTerrainNum(ByVal Pos As Long) As Byte
  GetTerrainNum = 1 + Int(Pos / 640)
  If GetTerrainNum > Size Then GetTerrainNum = Size
End Function

Public Function RepairAngle(Angle As Single) As Single
  RepairAngle = Angle
  If Angle > 359 Then RepairAngle = Angle - 359
  If Angle < 0 Then
  RepairAngle = 359 + Angle
  End If
End Function

Function GetAngle(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
  Dim Cislo1 As Long
  Dim Cislo2 As Long
  Dim Uhol As Double
  Dim Poloha As Integer
  
  If X1 = X2 And Y1 < Y2 Then
   Cislo2 = 0
   Poloha = 180
  
  
  ElseIf X1 = X2 And Y1 > Y2 Then
   Cislo2 = 0
   Poloha = 0
  ElseIf X1 < X2 And Y1 = Y2 Then
   Cislo2 = 0
   Poloha = 90
  ElseIf X1 > X2 And Y1 = Y2 Then
   Cislo2 = 0
   Poloha = 270
  ElseIf X1 < X2 And Y1 > Y2 Then
   Cislo1 = Abs(X2 - X1)
   Cislo2 = Abs(Y2 - Y1)
   Poloha = 0
  ElseIf X1 < X2 And Y1 < Y2 Then
   Cislo1 = Abs(Y1 - Y2)
   Cislo2 = Abs(X2 - X1)
   Poloha = 90
  ElseIf X1 > X2 And Y1 < Y2 Then
   Cislo1 = Abs(X1 - X2)
   Cislo2 = Abs(Y1 - Y2)
   Poloha = 180
  ElseIf X1 > X2 And Y1 > Y2 Then
   Cislo1 = Abs(Y2 - Y1)
   Cislo2 = Abs(X1 - X2)
   Poloha = 270
  End If
  
On Error GoTo Chyba
  Uhol = Atn(Cislo1 / Cislo2) * 57
Chyba:

  GetAngle = Uhol + Poloha
End Function

Function GetDist(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
  GetDist = Sqr((X1 - X2) * (X1 - X2) + (Y1 - Y2) * (Y1 - Y2))
End Function

Function GetDir(ByVal Angle1 As Integer, ByVal Angle2 As Integer) As Integer
  Dim Plus As Integer
  Dim Minus As Integer
  Dim PlusKurz As Integer
  Dim MinusKurz As Integer
  
  PlusKurz = Angle1
  MinusKurz = Angle1
  
  If Angle2 < Angle1 Then Plus = 360 - Angle1: PlusKurz = 0
  If Angle2 > Angle1 Then Minus = -Angle1: MinusKurz = 359
  
  Plus = Plus + (Angle2 - PlusKurz)
  Minus = Minus + (Angle2 - MinusKurz)
  
  If Plus > Abs(Minus) Then GetDir = Minus Else GetDir = Plus
  
End Function

Public Function GetFPS() As Integer
  Dim Rozdiel As Long
  
  Rozdiel = GetTickCount - OldTick
  CurrentFPS = CurrentFPS + 1
  
  If Rozdiel >= 1000 Then
   OldFPS = CurrentFPS
   OldTick = GetTickCount
   CurrentFPS = 0
  End If
  GetFPS = OldFPS
End Function
