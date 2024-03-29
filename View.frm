VERSION 4.00
Begin VB.Form View 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D View"
   ClientHeight    =   3600
   ClientLeft      =   5235
   ClientTop       =   2475
   ClientWidth     =   4800
   Height          =   4290
   Icon            =   "View.frx":0000
   Left            =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Top             =   1845
   Width           =   4920
   Begin VB.Timer TimerUnlockKeyboard 
      Interval        =   50
      Left            =   3000
      Top             =   840
   End
   Begin VB.Timer TimerUnlockMouse 
      Interval        =   250
      Left            =   2400
      Top             =   840
   End
   Begin VB.Timer TimerAnimate 
      Interval        =   500
      Left            =   1800
      Top             =   840
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   4080
      Top             =   240
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Image SpriteStorage 
      Height          =   240
      Index           =   0
      Left            =   3840
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line linWalls 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   160
      Y2              =   161
   End
   Begin VB.Image Sprites 
      Height          =   975
      Index           =   0
      Left            =   -150
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Ground 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu MenuOpenTilemap 
         Caption         =   "&Open Tilemap"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MenuWindow 
      Caption         =   "&Window"
      Begin VB.Menu MenuMiniMap 
         Caption         =   "&MiniMap"
      End
      Begin VB.Menu MenuLog 
         Caption         =   "&Log"
      End
   End
   Begin VB.Menu MenuOptions 
      Caption         =   "&Options"
      Begin VB.Menu MenuHalfResolution 
         Caption         =   "&Half Resolution"
      End
      Begin VB.Menu MenuQuarterResolution 
         Caption         =   "&Quarter Resolution"
      End
      Begin VB.Menu MenuOptionsDiv1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuDebugLog 
         Caption         =   "&Debug Log"
      End
      Begin VB.Menu MenuMouseNav 
         Caption         =   "&Mouse Navigation"
      End
   End
End
Attribute VB_Name = "View"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Const WalkSpeed = 1

Const WallSideNone = 0
Const WallSideNS = 2
Const WallSideEW = 1

Const XIdx = 1
Const YIdx = 2

Const MouseDeadZone = 5

Private Type TileRow
    Tiles(100) As Integer
End Type

Private Type Tilemap
    Rows(100) As TileRow
    Width As Integer
    Length As Integer
End Type

Private Type Mobile
    Name As String
    SpriteIdx As Integer
    WalkFrameIdxs(2) As Integer
    TilemapX As Integer
    TilemapY As Integer
    VXStart As Integer
    VXEnd As Integer
    Visible As Boolean
    Frame As Integer
    TalkText As String
End Type

Private Type Warp
    TargetTilemap As String
    TargetX As Integer
    TargetY As Integer
    TargetDirX As Single
    TargetDirY As Single
    TargetCameraLensX As Single
    TargetCameraLensY As Single
    SourceX As Integer
    SourceY As Integer
End Type

Private Type Ray
    Dir(3) As Single
    Tilemap(3) As Integer
    Step(3) As Integer
    TileDist(3) As Single
    Reach(3) As Single
    WallColorIdx As Integer
    WallSide As Integer
    SpriteIdx As Integer
End Type

Dim ViewW As Integer
Dim ViewH As Integer
Dim ViewHHalf As Integer
Dim Overscan As Integer
Dim LineWMult As Integer
Dim Rays() As Ray

Dim LastMouseX As Single
Dim LastMouseY As Single
Dim MouseLocked As Boolean
Dim KeyboardLocked As Boolean

Rem Player properties.
Dim PlayerX As Single
Dim PlayerY As Single
Dim PlayerDirX As Single
Dim PlayerDirY As Single
Dim CameraLensX As Single
Dim CameraLensY As Single

Rem Other mobile properties.
Dim Mobiles() As Mobile
Dim MobilesActive As Integer
Dim SpritesStored As Integer
Dim SpritesActive As Integer

Rem Tilemap properties.
Dim Map As Tilemap
Dim WallColors() As Long
Dim WallColorsActive As Integer

Rem Warps currently on tilemap.
Dim Warps() As Warp
Dim WarpsActive As Integer

Rem Used to set player X/Y on warp.
Dim Warping As Boolean
Dim WarpTilemap As String
Dim ForceStartX As Integer
Dim ForceStartY As Integer
Dim ForceStartDirX As Single
Dim ForceStartDirY As Single
Dim ForceStartCameraLensX As Single
Dim ForceStartCameraLensY As Single


Private Sub DrawWall(Ray As Ray, ByVal VStripeX As Integer, ByVal CoordIdx As Integer)
    Dim WallDist As Single

    WallDist = (Ray.Reach(CoordIdx) - Ray.TileDist(CoordIdx))
    
    Rem Make sure we don't divide by zero when too close to a wall below!
    If WallDist < 0.0001 Then
        WallDist = 0.0001
    End If
    
    If Ray.WallColorIdx > WallColorsActive Then
        Exit Sub
    End If
    
    Rem Draw the wall line and minimap ray.
    DrawVertLine VStripeX, ViewH / WallDist, WallColors(CoordIdx, Ray.WallColorIdx)
    If MenuMiniMap.Checked Then
        Rem Reverse X for minimap.
        MiniMap.RayEnd VStripeX, Map.Width - Ray.Tilemap(XIdx), Ray.Tilemap(YIdx), WallColors(CoordIdx, Ray.WallColorIdx)
    End If
    Rem Log.LogLine "Ended at: " & Ray.Tilemap(XIdx) & ", " & Ray.Tilemap(YIdx)
End Sub


Private Function InitRayWallDist(RayDir As Single)
    If 0 = RayDir Then
        Rem Use a large number so we don't divide by zero later.
        InitRayWallDist = 1E+32
    Else
        InitRayWallDist = Abs(1 / RayDir)
    End If
End Function

Public Function LoadStoredSprite(SpritePath As String) As Integer
    SpriteStorage(SpritesStored).Picture = LoadPicture(SpritePath)
    
    Rem Return the current sprite index and increment the count.
    LoadStoredSprite = SpritesStored
    Log.LogDebug "Loaded " & SpritePath & " as stored sprite: " & SpritesStored
    SpritesStored = SpritesStored + 1
    Load SpriteStorage(SpritesStored)
    Log.LogDebug "Incremented stored sprites to: " & SpritesStored
End Function

Public Sub LoadTilemap(Filename As String)
    Dim FileNo As Long
    Dim LineIn As String
    Dim LineArr() As String
    Dim TileIdx As Integer
    Dim RowIdx As Integer
    
    TimerAnimate.Enabled = False
    UnloadTilemap
    
    FileNo = FreeFile
    Open Filename For Input Access Read Shared As FileNo
        
    RowIdx = 0
    Do Until EOF(FileNo)
        Line Input #FileNo, LineIn
        StringSplit LineIn, ",", LineArr

        Rem Parse each line based on what kind of line it is.
        If "ground" = LineArr(0) Then
            Log.LogDebug "Ground color: " & LineArr(1)
            View.Ground.FillColor = LineArr(1)
            View.Ground.Visible = True
            
        ElseIf "sky" = LineArr(0) Then
            Log.LogDebug "Sky color: " & LineArr(1)
            View.BackColor = LineArr(1)
            
        ElseIf "width" = LineArr(0) Then
            Log.LogDebug "Map width: " & LineArr(1)
            Map.Width = LineArr(1)
            
        ElseIf "height" = LineArr(0) Then
            Log.LogDebug "Map height: " & LineArr(1)
            Map.Length = LineArr(1)
            
        ElseIf "wall" = LineArr(0) Then
            Rem Load a wall color.
            If LineArr(1) >= WallColorsActive Then
                WallColorsActive = LineArr(1) + 1
                ReDim Preserve WallColors(3, WallColorsActive) As Long
            End If
            WallColors(WallSideNS, LineArr(1)) = LineArr(2)
            WallColors(WallSideEW, LineArr(1)) = LineArr(3)
            Log.LogDebug "Added wall " & LineArr(1) & _
                ": NS: " & WallColors(WallSideNS, LineArr(1)) & _
                ", EW: " & WallColors(WallSideEW, LineArr(1))
            
        ElseIf "map" = LineArr(0) Then
            For TileIdx = 1 To Map.Width
                Rem TODO: Verify that the array is really Map.Width + 1 tiles long first.
                Map.Rows(RowIdx).Tiles(TileIdx - 1) = LineArr(TileIdx)
            Next TileIdx
            RowIdx = RowIdx + 1
            
        ElseIf "mobile" = LineArr(0) Then
            Rem Create a new mobile.
            ReDim Preserve Mobiles(MobilesActive + 1) As Mobile
            Mobiles(MobilesActive).SpriteIdx = SpritesActive
            If 0 < SpritesActive Then
                Log.LogDebug "Creating sprite: " & Mobiles(MobilesActive).SpriteIdx
                Load Sprites(Mobiles(MobilesActive).SpriteIdx)
            End If
            Mobiles(MobilesActive).Frame = 0
            Mobiles(MobilesActive).Name = LineArr(1)
            Mobiles(MobilesActive).TilemapX = LineArr(2)
            Mobiles(MobilesActive).TilemapY = LineArr(3)
            Mobiles(MobilesActive).WalkFrameIdxs(0) = LoadStoredSprite(LineArr(4))
            Mobiles(MobilesActive).WalkFrameIdxs(1) = LoadStoredSprite(LineArr(5))
            Mobiles(MobilesActive).TalkText = LineArr(6)
            Log.LogDebug "Loaded mobile " & MobilesActive & "(Sprite " & SpritesActive & "), " & _
                LineArr(4) & " (" & Mobiles(MobilesActive).WalkFrameIdxs(0) & _
                ")/" & LineArr(5) & " (" & Mobiles(MobilesActive).WalkFrameIdxs(1) & ") at " & _
                Mobiles(MobilesActive).TilemapX & ", " & Mobiles(MobilesActive).TilemapY
            
            SpritesActive = SpritesActive + 1
            MobilesActive = MobilesActive + 1
            
        ElseIf "warp" = LineArr(0) Then
            ReDim Preserve Warps(WarpsActive + 1) As Warp
            Warps(WarpsActive).SourceX = LineArr(1)
            Warps(WarpsActive).SourceY = LineArr(2)
            Warps(WarpsActive).TargetX = LineArr(4)
            Warps(WarpsActive).TargetY = LineArr(5)
            Warps(WarpsActive).TargetDirX = LineArr(6)
            Warps(WarpsActive).TargetDirY = LineArr(7)
            Warps(WarpsActive).TargetCameraLensX = LineArr(8)
            Warps(WarpsActive).TargetCameraLensY = LineArr(9)
            Warps(WarpsActive).TargetTilemap = LineArr(3)
            Log.LogDebug "Loaded warp (" & Warps(WarpsActive).SourceX & ", " & Warps(WarpsActive).SourceY & _
                ") to " & Warps(WarpsActive).TargetTilemap & " (" & Warps(WarpsActive).TargetX & ", " & Warps(WarpsActive).TargetY & ")"
            WarpsActive = WarpsActive + 1
            
        ElseIf "start" = LineArr(0) Then
            Rem Set player starting position.
            If Warping Then
                Log.LogDebug "Starting at warp target: " & WarpTilemap & ": " & ForceStartX & ", " & ForceStartY
                PlayerX = ForceStartX
                PlayerY = ForceStartY
                PlayerDirX = ForceStartDirX
                PlayerDirY = ForceStartDirY
                CameraLensX = ForceStartCameraLensX
                CameraLensY = ForceStartCameraLensY
            Else
                Log.LogDebug "Starting at: " & LineArr(1) & ", " & LineArr(2)
                PlayerX = LineArr(1)
                PlayerY = LineArr(2)
                PlayerDirX = LineArr(3)
                PlayerDirY = LineArr(4)
                CameraLensX = LineArr(5)
                CameraLensY = LineArr(6)
            End If
        End If
    Loop
    
    Rem Set warping false so UpdateView below works, but do it after load or
    Rem "start" lines above won't work properly.
    Warping = False
    
    If MenuMiniMap.Checked Then
        Rem Rescale the mini map if it's open.
        MiniMap.SetupLines ViewW, Map.Width, Map.Length, True
        MiniMap.SetupLines ViewW, Map.Width, Map.Length, False
    End If
    
    UpdateView
    
    TimerAnimate.Enabled = True
End Sub
Private Sub RotateView(ByVal PlayerCurrentDirX As Single, ByVal CameraCurrentDirX As Single, RotateSpeed As Single)
    Rem Pass the old dir in by value so we can use it in the rotation multiplications below.
    PlayerDirX = (PlayerCurrentDirX * Cos(RotateSpeed)) - (PlayerDirY * Sin(RotateSpeed))
    PlayerDirY = (PlayerCurrentDirX * Sin(RotateSpeed)) + (PlayerDirY * Cos(RotateSpeed))
    CameraLensX = (CameraCurrentDirX * Cos(RotateSpeed)) - (CameraLensY * Sin(RotateSpeed))
    CameraLensY = (CameraCurrentDirX * Sin(RotateSpeed)) + (CameraLensY * Cos(RotateSpeed))
    Log.LogDebug "New DirX: " & PlayerDirX & ", DirY: " & PlayerDirY & ", " & _
        "CameraLensX: " & CameraLensX & ", CameraLensY: " & CameraLensY
    UpdateView
End Sub


Public Sub SetupLines(UnloadLines As Boolean)
    Dim XOff As Integer

    Rem Setup the wall lines.
    For XOff = 0 To ViewW - 1
        Rem Expand control array as needed.
        If XOff > 0 Then
            If UnloadLines Then
                Unload linWalls(XOff)
            Else
                Load linWalls(XOff)
            End If
        End If
        If Not UnloadLines Then
            Rem Bring to front.
            linWalls(XOff).ZOrder
            DrawVertLine XOff, 0, 0
        End If
    Next XOff
    
    If MenuMiniMap.Checked Then
        MiniMap.SetupLines ViewW, Map.Width, Map.Length, UnloadLines
    End If
End Sub

Public Sub SetupScreen(ViewWIn As Integer, ViewHIn As Integer, OverscanIn As Integer, LineWMultIn As Integer)
    SetupLines True
    ViewW = ViewWIn
    ViewH = ViewHIn
    ViewHHalf = ViewH / 2
    Overscan = OverscanIn
    LineWMult = LineWMultIn
    ReDim Rays(ViewW + (2 * Overscan))
    SetupLines False
End Sub

Public Sub StringSplit(Haystack As String, Needle As String, StringsOut() As String)
    Dim NewHaystack As String
    Dim LastNeedle As Integer
    Dim ThisNeedle As Integer
    Dim StringsFound As Integer
    
    StringsFound = 0
    ThisNeedle = 1
    LastNeedle = 1
    Do
        Rem Find the next comma.
        ThisNeedle = InStr(LastNeedle, Haystack, Needle, 1)
        
        Rem This is either the next or last substring, but a substring regardless.
        ReDim Preserve StringsOut(StringsFound) As String
        
        If 0 = ThisNeedle Then
            Rem This is the last substring, so just grab the rest of the string into it.
            StringsOut(StringsFound) = Mid(Haystack, LastNeedle)
            StringsFound = StringsFound + 1
            Exit Do
        Else
            Rem The length of a string between commas is the last needle minus this needle.
            StringsOut(StringsFound) = Mid(Haystack, LastNeedle, ThisNeedle - LastNeedle)
            StringsFound = StringsFound + 1
        End If
        LastNeedle = ThisNeedle + 1
    Loop
End Sub

Public Sub UnloadTilemap()
    Dim Iter As Integer
    
    Rem Clear out sprite storage.
    For Iter = 1 To SpritesStored
        Unload SpriteStorage(Iter)
    Next Iter
    SpritesStored = 0
    
    Rem Return sprite display images to initial state.
    Sprites(0).Visible = False
    For Iter = 1 To SpritesActive - 1
        Unload Sprites(Iter)
    Next Iter
    SpritesActive = 0
    
    Rem Clear out mobiles.
    MobilesActive = 0
    ReDim Mobiles(MobilesActive) As Mobile
    
    Rem Clear out wall colors.
    WallColorsActive = 0
    ReDim WallColors(3, WallColorsActive) As Long
End Sub

Public Sub UpdateView()
    Dim VStripeX As Integer
    Dim MobileIter As Integer
    Dim MobileWidth As Integer
    
    Rem Reset all mobiles to off-screen.
    For MobileIter = 0 To MobilesActive - 1
        Mobiles(MobileIter).VXStart = 0
        Mobiles(MobileIter).VXEnd = 0
        Mobiles(MobileIter).Visible = False
        Sprites(Mobiles(MobileIter).SpriteIdx).Visible = False
    Next MobileIter
    
    Rem Cast a ray for each pixel-wide vertical line.
    Rem We use an overscan here, processing a few X pixels to the left/right of the visible
    Rem field, so that sprite picture boxes that start off-screen to the left or right
    Rem because they're partially off-screen due to e.g. being too close, don't vanish
    Rem entirely.
    For VStripeX = -1 * Overscan To ViewW + Overscan
        UpdateViewRay VStripeX, Rays(VStripeX + Overscan)
    Next VStripeX
    
    Rem Place picture boxes for visible mobiles.
    For MobileIter = 0 To MobilesActive - 1
        If Mobiles(MobileIter).Visible Then
            MobileWidth = (Mobiles(MobileIter).VXEnd - Mobiles(MobileIter).VXStart) * LineWMult
            Sprites(Mobiles(MobileIter).SpriteIdx).Left = Mobiles(MobileIter).VXStart * LineWMult
            Sprites(Mobiles(MobileIter).SpriteIdx).Width = MobileWidth
            Sprites(Mobiles(MobileIter).SpriteIdx).Height = MobileWidth
            Sprites(Mobiles(MobileIter).SpriteIdx).Top = ViewHHalf - (MobileWidth / 2)
            Sprites(Mobiles(MobileIter).SpriteIdx).Visible = True
            Sprites(Mobiles(MobileIter).SpriteIdx).ZOrder
        End If
    Next MobileIter
End Sub

Private Sub UpdateViewRay(VStripeX As Integer, Ray As Ray)
    Dim CameraLensVStripeX As Single
    Dim MobileIter As Integer
    
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
    If 0 <= VStripeX And VStripeX < ViewW And MenuMiniMap.Checked Then
        Rem Reverse X for minimap.
        MiniMap.RayStart VStripeX, Map.Width - PlayerX, PlayerY
    End If
    
    Rem Set initial distance to next wall based on ray angle hypoteneuse.
    Ray.TileDist(XIdx) = InitRayWallDist(Ray.Dir(XIdx))
    Ray.TileDist(YIdx) = InitRayWallDist(Ray.Dir(YIdx))

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
        If 0 <= Ray.Tilemap(XIdx) And Map.Width > Ray.Tilemap(XIdx) And 0 <= Ray.Tilemap(YIdx) And Map.Length > Ray.Tilemap(YIdx) Then
            Ray.WallColorIdx = Map.Rows(Int(Ray.Tilemap(XIdx))).Tiles(Int(Ray.Tilemap(YIdx)))
            If 0 = Ray.WallColorIdx Then
                Rem In a cell with no wall.
                Ray.WallSide = 0
            End If
        Else
            Rem Virtual wall of type 1 around the map.
            Ray.WallColorIdx = 1
        End If
        
        Rem Check if this ray passes through a mobile tile.
        For MobileIter = 0 To MobilesActive - 1
            If Mobiles(MobileIter).TilemapX = Ray.Tilemap(XIdx) And _
            Mobiles(MobileIter).TilemapY = Ray.Tilemap(YIdx) Then
                If Not Mobiles(MobileIter).Visible Then
                    Rem This is the first vertical X stripe this mobile appears in.
                    Mobiles(MobileIter).VXStart = VStripeX
                    Mobiles(MobileIter).Visible = True
                End If
                If VStripeX > Mobiles(MobileIter).VXEnd Then
                    Rem This is the first vertical X stripe this mobile appears in.
                    Mobiles(MobileIter).VXEnd = VStripeX
                End If
            End If
        Next MobileIter
    Wend
    
    If 0 <= VStripeX And VStripeX < ViewW Then
        Rem Draw the wall that we eventually encountered (if it's on-screen).
        DrawWall Ray, VStripeX, Ray.WallSide
    End If
End Sub
Public Sub DrawVertLine(XOff As Integer, YHeight As Single, ByVal Color As Long)
    linWalls(XOff).Y1 = ViewHHalf - (YHeight / 2)
    linWalls(XOff).Y2 = ViewHHalf + (YHeight / 2)
    linWalls(XOff).X1 = XOff * LineWMult
    linWalls(XOff).X2 = XOff * LineWMult
    linWalls(XOff).BorderWidth = LineWMult
    linWalls(XOff).Visible = True
    linWalls(XOff).BorderColor = Color
End Sub

Public Sub WalkView(Distance As Single)
    Dim DistanceX As Single
    Dim DistanceY As Single
    Dim NewX As Single
    Dim NewY As Single
    Dim WarpIter As Integer
    
    Rem PlayerDir* are precalculated to increment rays, so walking is just another "ray"!
    NewX = PlayerX + (Distance * PlayerDirX)
    NewY = PlayerY + (Distance * PlayerDirY)
    
    If Map.Rows(Int(NewX)).Tiles(Int(NewY)) = 0 Then
        PlayerX = NewX
        PlayerY = NewY
    End If
    
    For WarpIter = 0 To WarpsActive - 1
        If Int(NewX) = Warps(WarpIter).SourceX And Int(NewY) = Warps(WarpIter).SourceY Then
            ForceStartX = Warps(WarpIter).TargetX
            ForceStartY = Warps(WarpIter).TargetY
            ForceStartDirX = Warps(WarpIter).TargetDirX
            ForceStartDirY = Warps(WarpIter).TargetDirY
            ForceStartCameraLensX = Warps(WarpIter).TargetCameraLensX
            ForceStartCameraLensY = Warps(WarpIter).TargetCameraLensY
            WarpTilemap = Warps(WarpIter).TargetTilemap
            Log.LogDebug "Set warp target to " & WarpTilemap & ": " & ForceStartX & ", " & ForceStartY
            Warping = True
        End If
    Next WarpIter
    
    If Warping Then
        LoadTilemap WarpTilemap
    Else
        UpdateView
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim PrevX As Single
    
    If KeyboardLocked Then
        Exit Sub
    End If
    
    If KeyAscii = 97 Then
        Rem 'a'
        RotateView PlayerDirX, CameraLensX, 0.33
        KeyboardLocked = True
        MouseLocked = True
    
    ElseIf KeyAscii = 100 Then
        Rem 'd'
        RotateView PlayerDirX, CameraLensX, -0.33
        KeyboardLocked = True
        MouseLocked = True
    
    ElseIf KeyAscii = 119 Then
        Rem 'w'
        WalkView 1
        KeyboardLocked = True
        MouseLocked = True
    
    ElseIf KeyAscii = 115 Then
        Rem 's'
        WalkView -1
        KeyboardLocked = True
        MouseLocked = True
    End If
End Sub

Private Sub Form_Load()
    
    Rem No pictures loaded yet.
    SpritesStored = 0
    SpritesActive = 0
    MobilesActive = 0
    WallColorsActive = 0
    Warping = False
    LastMouseX = 0
    LastMouseY = 0
    Map.Width = 0
    Map.Length = 0
    
    SetupScreen 320, 240, 80, 1

    Log.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastMouseX = X
    LastMouseY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseLocked Or Not MenuMouseNav.Checked Then
        Exit Sub
    End If

    If 1 = Button Then
        If Y < LastMouseY - MouseDeadZone Then
            WalkView 1
        ElseIf Y > LastMouseY + MouseDeadZone Then
            WalkView -1
        ElseIf X < LastMouseX - MouseDeadZone Then
            RotateView PlayerDirX, CameraLensX, 0.33
        ElseIf X > LastMouseX + MouseDeadZone Then
            RotateView PlayerDirX, CameraLensX, -0.33
        End If
        MouseLocked = True
        KeyboardLocked = True
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If MenuMiniMap.Checked Then
        Unload MiniMap
    End If
    If MenuLog.Checked Then
        Unload Log
    End If
End Sub


Private Sub MenuDebugLog_Click()
    If MenuDebugLog.Checked Then
        MenuDebugLog.Checked = False
    Else
        MenuDebugLog.Checked = True
    End If
End Sub

Private Sub MenuExit_Click()
    Unload View
End Sub


Private Sub MenuHalfResolution_Click()
    If MenuHalfResolution.Checked Then
        SetupScreen 320, 240, 80, 1
        MenuHalfResolution.Checked = False
    Else
        MenuQuarterResolution.Checked = False
        SetupScreen 160, 240, 40, 2
        MenuHalfResolution.Checked = True
    End If
    Log.LogDebug "Horizontal resolution changed to: " & ViewW
    UpdateView
End Sub

Private Sub MenuLog_Click()
    If MenuLog.Checked Then
        Unload Log
        MenuLog.Checked = False
    Else
        Log.Show
    End If
End Sub

Private Sub MenuMiniMap_Click()
    If MenuMiniMap.Checked Then
        Unload MiniMap
        MenuMiniMap.Checked = False
    Else
        MiniMap.Show
        MiniMap.SetupLines ViewW, Map.Width, Map.Length, False
        
        Rem Refresh the view so the minimap redraws its lines.
        UpdateView
    End If
End Sub


Private Sub MenuMouseNav_Click()
    If Not MenuMouseNav.Checked Then
        MenuMouseNav.Checked = True
    Else
        MenuMouseNav.Checked = False
    End If
End Sub

Private Sub MenuOpenTilemap_Click()
    dialog.DialogTitle = "Open Tilemap"
    dialog.Filter = "Comma-Separated Values (*.csv)|*.csv"
    dialog.ShowOpen
    If "" <> dialog.Filename Then
        LoadTilemap dialog.Filename
    End If
End Sub

Private Sub MenuQuarterResolution_Click()
    If MenuQuarterResolution.Checked Then
        SetupScreen 320, 240, 80, 1
        MenuQuarterResolution.Checked = False
    Else
        MenuHalfResolution.Checked = False
        SetupScreen 80, 240, 40, 4
        MenuQuarterResolution.Checked = True
    End If
    Log.LogDebug "Horizontal resolution changed to: " & ViewW
    UpdateView
End Sub

Private Sub Sprites_Click(Index As Integer)
    Dim MobileIter As Integer
    
    For MobileIter = 0 To MobilesActive - 1
        If Index = Mobiles(MobileIter).SpriteIdx Then
            Log.LogTalk Mobiles(MobileIter).Name, Mobiles(MobileIter).TalkText
            Exit Sub
        End If
    Next MobileIter
End Sub

Private Sub TimerAnimate_Timer()
    Dim MobIter As Integer
       
    For MobIter = 0 To MobilesActive - 1
        If 0 = Mobiles(MobIter).Frame Then
            Mobiles(MobIter).Frame = 1
        Else
            Mobiles(MobIter).Frame = 0
        End If
        Rem Log.LogDebug "Setting Sprite " & Mobiles(MobIter).SpriteIdx & " to Stored Sprite " & Mobiles(MobIter).WalkFrameIdxs(Mobiles(MobIter).Frame)
        Sprites(Mobiles(MobIter).SpriteIdx).Picture = _
            SpriteStorage(Mobiles(MobIter).WalkFrameIdxs(Mobiles(MobIter).Frame)).Picture
    Next MobIter
End Sub


Private Sub TimerUnlockKeyboard_Timer()
    KeyboardLocked = False
End Sub

Private Sub TimerUnlockMouse_Timer()
    MouseLocked = False
End Sub


