VERSION 4.00
Begin VB.Form View 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D View"
   ClientHeight    =   3600
   ClientLeft      =   5265
   ClientTop       =   3375
   ClientWidth     =   4800
   Height          =   4005
   Icon            =   "View.frx":0000
   Left            =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Top             =   3030
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
Option Explicit

Const ViewW = 320
Const ViewH = 240
Const ViewHHalf = ViewH / 2

Const WalkSpeed = 1

Const MaxSprites = 10
Const TilemapWidth = 10
Const TilemapHeight = 10

Const WallSideNone = 0
Const WallSideNS = 2
Const WallSideEW = 1

Const XIdx = 1
Const YIdx = 2

Dim WallColors(3, 5) As Long

Private Type TileRow
    Cells(10) As Integer
End Type

Private Type Tilemap
    Rows(10) As TileRow
End Type

Private Type Sprite
    PictureIdx As Integer
    TilemapX As Integer
    TilemapY As Integer
End Type

Private Type Ray
    Dir(3) As Single
    Tilemap(3) As Single
    Step(3) As Integer
    TileDist(3) As Single
    Reach(3) As Single
    WallColorIdx As Integer
    WallSide As Integer
    SpriteIdx As Integer
End Type

Dim PlayerX As Single
Dim PlayerY As Single
Dim PlayerDirX As Single
Dim PlayerDirY As Single
Dim CameraLensX As Single
Dim CameraLensY As Single
Dim Map As Tilemap
Dim SpriteObjects(MaxSprites) As Sprite

Private Sub DrawWall(Ray As Ray, ByVal VStripeX As Integer, ByVal CoordIdx As Integer)
    Dim WallDist As Single

    WallDist = (Ray.Reach(CoordIdx) - Ray.TileDist(CoordIdx))
    
    Rem Make sure we don't divide by zero when too close to a wall below!
    If WallDist < 0.0001 Then
        WallDist = 0.0001
    End If
    
    Rem Draw the wall line and minimap ray.
    VertLine VStripeX, ViewH / WallDist, WallColors(CoordIdx, Ray.WallColorIdx)
    MiniMap.RayEnd VStripeX, Ray.Tilemap(XIdx), Ray.Tilemap(YIdx), WallColors(CoordIdx, Ray.WallColorIdx)
End Sub


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
    Dim Rays(ViewW) As Ray
    
    Rem Cast a ray for each pixel-wide vertical line.
    For VStripeX = 0 To ViewW - 1
        UpdateViewRay VStripeX, Rays(VStripeX)
    Next VStripeX
End Sub

Private Sub UpdateViewRay(VStripeX As Integer, Ray As Ray)
    Dim CameraLensVStripeX As Single
    
    Rem No wall hit yet!
    Ray.WallSide = WallSideNone
    
    Rem Translate pixel screen vertical coord into camera plane vertical coord.
    CameraLensVStripeX = ((2 * VStripeX) / ViewW) - 1
    
    Rem Setup ray for this vertical stripe's initial position.
    Ray.Dir(XIdx) = PlayerDirX + (CameraLensX * CameraLensVStripeX)
    Ray.Dir(YIdx) = PlayerDirY + (CameraLensY * CameraLensVStripeX)
    
    Rem Set tilemap tile ray is in based on player position.
    Ray.Tilemap(XIdx) = PlayerX
    Ray.Tilemap(YIdx) = PlayerY
    MiniMap.RayStart VStripeX, PlayerX, PlayerY
    
    Rem Set initial distance to next wall based on ray angle hypoteneuse.
    Ray.TileDist(XIdx) = RayWallDist(Ray.Dir(XIdx))
    Ray.TileDist(YIdx) = RayWallDist(Ray.Dir(YIdx))

    If 0 > Ray.Dir(XIdx) Then
        Rem Moving to the west.
        Ray.Step(XIdx) = -1
        Ray.Reach(XIdx) = (PlayerX - Ray.Tilemap(XIdx)) * Ray.TileDist(XIdx)
    Else
        Rem Moving to the east.
        Ray.Step(XIdx) = 1
        Ray.Reach(XIdx) = (Ray.Tilemap(XIdx) + (1# - PlayerX)) * Ray.TileDist(XIdx)
    End If
    
    If 0 > Ray.Dir(YIdx) Then
        Rem Moving to the north.
        Ray.Step(YIdx) = -1
        Ray.Reach(YIdx) = (PlayerY - Ray.Tilemap(YIdx)) * Ray.TileDist(YIdx)
    Else
        Rem Moving to the south.
        Ray.Step(YIdx) = 1
        Ray.Reach(YIdx) = (Ray.Tilemap(YIdx) + (1# - PlayerY)) * Ray.TileDist(YIdx)
    End If
    
    Rem Perform the raycast!
    While WallSideNone = Ray.WallSide
        Rem Move the ray forward depending on whether last time we moved map tile by X or Y.
        If Ray.Reach(XIdx) < Ray.Reach(YIdx) Then
            Ray.Reach(XIdx) = Ray.Reach(XIdx) + Ray.TileDist(XIdx)
            Ray.Tilemap(XIdx) = Ray.Tilemap(XIdx) + Ray.Step(XIdx)
            Ray.WallSide = WallSideEW
        Else
            Ray.Reach(YIdx) = Ray.Reach(YIdx) + Ray.TileDist(YIdx)
            Ray.Tilemap(YIdx) = Ray.Tilemap(YIdx) + Ray.Step(YIdx)
            Ray.WallSide = WallSideNS
        End If
        
        Rem Check if there was actually a collision.
        If 0 <= Ray.Tilemap(XIdx) And TilemapWidth > Ray.Tilemap(XIdx) And 0 <= Ray.Tilemap(YIdx) And TilemapHeight > Ray.Tilemap(YIdx) Then
            Ray.WallColorIdx = Map.Rows(Int(Ray.Tilemap(XIdx))).Cells(Int(Ray.Tilemap(YIdx)))
            If 0 = Ray.WallColorIdx Then
                Rem In a cell with no wall.
                Ray.WallSide = 0
            End If
        Else
            Rem Virtual wall of type 1 around the map.
            Ray.WallColorIdx = 1
        End If
    Wend
    
    Rem Draw the wall that we eventually encountered.
    DrawWall Ray, VStripeX, Ray.WallSide
    
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
    Dim Y As Integer
    Dim X As Integer
    
    Rem Set player position.
    PlayerX = 4
    PlayerY = 4
    PlayerDirX = -1
    PlayerDirY = 0
    CameraLensX = 0
    CameraLensY = 0.66
    
    Rem Setup wall colors.
    WallColors(WallSideNS, 1) = &HFF0000
    WallColors(WallSideEW, 1) = &H800000
    WallColors(WallSideNS, 2) = &HFF00&
    WallColors(WallSideEW, 2) = &H8000&
    WallColors(WallSideNS, 3) = &HFF&
    WallColors(WallSideEW, 3) = &H80&
    WallColors(WallSideNS, 4) = &HFF00FF
    WallColors(WallSideEW, 4) = &H800080
    
    Rem Generate the tilemap.
    For Y = 0 To TilemapHeight - 1
        If Y = 0 Or Y = TilemapHeight - 1 Then
            Rem Fill in entire top and bottom rows.
            For X = 0 To TilemapWidth - 1
                Map.Rows(Y).Cells(X) = 1
            Next X
        Else
            Rem Just fill in side walls.
            Map.Rows(Y).Cells(0) = 1
            Map.Rows(Y).Cells(TilemapHeight - 1) = 1
        End If
    Next Y
    Map.Rows(2).Cells(2) = 3
    Map.Rows(7).Cells(7) = 4
    
    Rem Generate sprites.
    
    
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
    
    Rem Setup Minimap.
    MiniMap.ScaleWidth = TilemapWidth * 3
    MiniMap.ScaleHeight = TilemapHeight * 3
    MiniMap.Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload MiniMap
End Sub


