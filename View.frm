VERSION 4.00
Begin VB.Form View 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D View"
   ClientHeight    =   3600
   ClientLeft      =   4605
   ClientTop       =   3015
   ClientWidth     =   4800
   Height          =   4005
   Icon            =   "View.frx":0000
   Left            =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Top             =   2670
   Width           =   4920
   Begin VB.Line linWalls 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   160
      Y2              =   161
   End
   Begin VB.Image Sprites 
      Height          =   135
      Index           =   0
      Left            =   720
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Ground 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   0
      Top             =   1560
      Width           =   4815
   End
End
Attribute VB_Name = "View"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Const ViewW = 320
Const ViewH = 240
Const ViewHHalf = ViewH / 2

Const WallSideNone = 0
Const WallSideNS = 1
Const WallSideEW = 2

Dim WallColorsNS(5) As Long
Dim WallColorsEW(5) As Long

Private Type TileRow
    Cells(10) As Integer
End Type

Private Type TileMap
    Rows(10) As TileRow
End Type

Dim PlayerX As Single
Dim PlayerY As Single
Dim PlayerDirX As Single
Dim PlayerDirY As Single
Dim CameraLensX As Single
Dim CameraLensY As Single
Dim Map As TileMap

Private Function RayWallDist(RayDir As Single)
    If 0 = RayDir Then
        Rem Use a large number so we don't divide by zero later.
        RayWallDist = 1E+32
    Else
        RayWallDist = Abs(1 / RayDir)
    End If
End Function

Private Sub RotateView(ByVal PlayerCurrentDirX As Single, ByVal CameraCurrentDirX As Single, RotateSpeed As Single)
    Rem Pass the old dir in by value so we can use it in the rotation multiplications below.
    PlayerDirX = (PlayerCurrentDirX * Cos(RotateSpeed)) - (PlayerDirY * Sin(RotateSpeed))
    PlayerDirY = (PlayerCurrentDirX * Sin(RotateSpeed)) + (PlayerDirY * Cos(RotateSpeed))
    CameraLensX = (CameraCurrentDirX * Cos(RotateSpeed)) - (CameraLensY * Sin(RotateSpeed))
    CameraLensY = (CameraCurrentDirX * Sin(RotateSpeed)) + (CameraLensY * Cos(RotateSpeed))
    UpdateView
End Sub


Public Sub UpdateView()
    Dim VStripeX As Integer
    
    Rem Cast a ray for each pixel-wide vertical line.
    For VStripeX = 0 To ViewW - 1
        UpdateViewRay VStripeX
    Next VStripeX
End Sub

Private Sub UpdateViewRay(VStripeX As Integer)
    Dim RayVector As Single
    Dim CameraLensVStripeX As Single
    Dim RayDirX As Single
    Dim RayDirY As Single
    Dim RayTilemapX As Single
    Dim RayTilemapY As Single
    Dim RayStepX As Integer
    Dim RayStepY As Integer
    Dim RayTileDistX As Single
    Dim RayTileDistY As Single
    Dim RayLenX As Single
    Dim RayLenY As Single
    Dim WallDist As Single
    Dim WallLineHeight As Integer
    Dim WallColorIdx As Integer
    Dim RayWallSide As Integer
    
    Rem No wall hit yet!
    RayWallSide = WallSideNone
    
    Rem Translate pixel screen vertical coord into camera plane vertical coord.
    CameraLensVStripeX = ((2 * VStripeX) / ViewW) - 1
    
    Rem Setup ray for this vertical stripe's initial position.
    RayDirX = PlayerDirX + (CameraLensX * CameraLensVStripeX)
    RayDirY = PlayerDirY + (CameraLensY * CameraLensVStripeX)
    
    Rem Set tilemap tile ray is in based on player position.
    RayTilemapX = PlayerX
    RayTilemapY = PlayerY
    MiniMap.RayStart VStripeX, PlayerX, PlayerY
    
    Rem Set initial distance to next wall based on ray angle hypoteneuse.
    RayTileDistX = RayWallDist(RayDirX)
    RayTileDistY = RayWallDist(RayDirY)

    If 0 > RayDirX Then
        Rem Moving to the west.
        RayStepX = -1
        RayLenX = (PlayerX - RayTilemapX) * RayTileDistX
    Else
        Rem Moving to the east.
        RayStepX = 1
        RayLenX = (RayTilemapX + (1# - PlayerX)) * RayTileDistX
    End If
    
    If 0 > RayDirY Then
        Rem Moving to the north.
        RayStepY = -1
        RayLenY = (PlayerY - RayTilemapY) * RayTileDistY
    Else
        Rem Moving to the south.
        RayStepY = 1
        RayLenY = (RayTilemapY + (1# - PlayerY)) * RayTileDistY
    End If
    
    Rem Perform the raycast!
    While WallSideNone = RayWallSide
        Rem Move the ray forward depending on whether last time we moved map tile by X or Y.
        If RayLenX < RayLenY Then
            RayLenX = RayLenX + RayTileDistX
            RayTilemapX = RayTilemapX + RayStepX
            RayWallSide = WallSideEW
        Else
            RayLenY = RayLenY + RayTileDistY
            RayTilemapY = RayTilemapY + RayStepY
            RayWallSide = WallSideNS
        End If
        
        Rem Check if there was actually a collision.
        If 0 <= RayTilemapX And 10 > RayTilemapX And 0 <= RayTilemapY And 10 > RayTilemapY Then
            WallColorIdx = Map.Rows(Int(RayTilemapX)).Cells(Int(RayTilemapY))
            If 0 = WallColorIdx Then
                Rem In a cell with no wall.
                RayWallSide = 0
            End If
        Else
            Rem Virtual wall of type 1 around the map.
            WallColorIdx = 1
        End If
    Wend
    
    Rem Draw the wall that we eventually encountered.
    If WallSideEW = RayWallSide Then
        WallDist = (RayLenX - RayTileDistX)
        If WallDist < 0.01 Then
            WallDist = 0.01
        End If
        VertLine VStripeX, ViewH / WallDist, WallColorsEW(WallColorIdx)
        MiniMap.RayEnd VStripeX, RayTilemapX, RayTilemapY, WallColorsEW(WallColorIdx)
    Else
        WallDist = (RayLenY - RayTileDistY)
        If WallDist < 0.01 Then
            WallDist = 0.01
        End If
        VertLine VStripeX, ViewH / WallDist, WallColorsNS(WallColorIdx)
        MiniMap.RayEnd VStripeX, RayTilemapX, RayTilemapY, WallColorsNS(WallColorIdx)
    End If
    
End Sub
Public Sub VertLine(XOff As Integer, YHeight As Single, ByVal Color As Long)
    linWalls(XOff).Y1 = ViewHHalf - (YHeight / 2)
    linWalls(XOff).Y2 = ViewHHalf + (YHeight / 2)
    linWalls(XOff).X1 = XOff
    linWalls(XOff).X2 = XOff
    linWalls(XOff).Visible = True
    linWalls(XOff).BorderColor = Color
End Sub

Public Sub WalkView(Distance As Single)
    Dim DistanceX As Single
    Dim DistanceY As Single
    Dim NewX As Single
    Dim NewY As Single
    
    Rem PlayerDir* are precalculated to increment rays, so walking is just another "ray"!
    NewX = PlayerX + (Distance * PlayerDirX)
    NewY = PlayerY + (Distance * PlayerDirY)
    
    If Map.Rows(Int(NewX)).Cells(Int(NewY)) = 0 Then
        PlayerX = NewX
        PlayerY = NewY
    End If
    
    UpdateView
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim PrevX As Single
    If KeyAscii = 97 Then
        Rem 'a'
        RotateView PlayerDirX, CameraLensX, 0.33
    
    ElseIf KeyAscii = 100 Then
        Rem 'd'
        RotateView PlayerDirX, CameraLensX, -0.33
    
    ElseIf KeyAscii = 119 Then
        Rem 'w'
        WalkView 1
    
    ElseIf KeyAscii = 115 Then
        Rem 's'
        WalkView -1
    End If
End Sub

Private Sub Form_Load()
    Dim XOff As Integer
    
    Rem Set player position.
    PlayerX = 4
    PlayerY = 4
    PlayerDirX = -1
    PlayerDirY = 0
    CameraLensX = 0
    CameraLensY = 0.66
    
    Rem Setup wall colors.
    WallColorsNS(1) = &HFF0000
    WallColorsEW(1) = &H800000
    WallColorsNS(2) = &HFF00&
    WallColorsEW(2) = &H8000&
    WallColorsNS(3) = &HFF&
    WallColorsEW(3) = &H80&
    WallColorsNS(4) = &HFF00FF
    WallColorsEW(4) = &H800080
    
    Rem Generate the tilemap.
    For Y = 0 To 9
        If Y = 0 Or Y = 9 Then
            Rem Fill in entire top and bottom rows.
            For X = 0 To 9
                Map.Rows(Y).Cells(X) = 1
            Next X
        Else
            Rem Just fill in side walls.
            Map.Rows(Y).Cells(0) = 1
            Map.Rows(Y).Cells(9) = 1
        End If
    Next Y
    Map.Rows(2).Cells(2) = 3
    Map.Rows(7).Cells(7) = 4
    
    Rem Setup the wall lines.
    For XOff = 0 To ViewW - 1
        Rem Expand control array as needed.
        If XOff > 0 Then
            Load linWalls(XOff)
        End If
        Rem Bring to front.
        linWalls(XOff).ZOrder
        VertLine XOff, 0, 0
    Next XOff
    
    UpdateView
    
    MiniMap.Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload MiniMap
End Sub


