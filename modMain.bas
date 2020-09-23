Attribute VB_Name = "modMain"
'------------------------------
'         modMain.bas
'------------------------------

Public DirectX As New DirectX7
Public CameraFrame As Direct3DRMFrame3
Public SceneFrame As Direct3DRMFrame3
Public DDraw As DirectDraw4
Public Clipper As DirectDrawClipper
Public D3DRM As Direct3DRM3
Public Device As Direct3DRMDevice3
Public Viewport As Direct3DRMViewport2

Public StartGameTime As Long, NowTime As Long, Delay As Integer
Public StartTick As Long, LastTick As Long

'General Declarations
Public QuitFlag As Boolean
Public UseJoystick As Boolean

Public XSpeed
Public YSpeed
Public ZSpeed
Public RotateSpeed

Public OnGround As Boolean
Public Dead As Boolean

Public CameraView As Integer
Public Const CHASE = 0
Public Const INSIDE = 1
Public Const SKY = 2
Public Const FREE = 3

'Lights
Public Light As Direct3DRMLight
Public LightFrame As Direct3DRMFrame3

'Objects
Public ChopperFrame As Direct3DRMFrame3
Public ChopperMesh As Direct3DRMMeshBuilder3
Public ChopperLives As Integer
Public ChopperPos As D3DVECTOR
Public ChopperShoot As Boolean
Public PeopleInChopper As Integer

Public ShowPos As Boolean

Public TextFrame As Direct3DRMFrame3
Public TextMesh As Direct3DRMMeshBuilder3
Public TextMoveCount As Integer
Public PlayingIntro As Boolean
Public PlayingEnding As Boolean

Public BladeFrame As Direct3DRMFrame3
Public BladeMesh As Direct3DRMMeshBuilder3

Public SkyFrame As Direct3DRMFrame3
Public SkyMesh As Direct3DRMMeshBuilder3
Public SkyTexture As Direct3DRMTexture3

Public GroundFrame As Direct3DRMFrame3
Public GroundMesh As Direct3DRMMeshBuilder3
Public RiverMesh As Direct3DRMMeshBuilder3

Public Const NumTrees = 9
Public TreeFrame(NumTrees) As Direct3DRMFrame3
Public TreeMesh As Direct3DRMMeshBuilder3

Public ExplodeFrame As Direct3DRMFrame3
Public ExplodeMesh As Direct3DRMMeshBuilder3
Public ExplodeCount As Integer

Public BaseFrame As Direct3DRMFrame3
Public BaseMesh As Direct3DRMMeshBuilder3
Public AtBase As Boolean

Public HouseFrame(3) As Direct3DRMFrame3
Public HouseMesh As Direct3DRMMeshBuilder3

Public TankFrame(3) As Direct3DRMFrame3
Public TankMesh As Direct3DRMMeshBuilder3
Public TankDead(3) As Boolean
Public TankExplodeCount(3) As Integer

Public PlaneFrame As Direct3DRMFrame3
Public PlaneMesh As Direct3DRMMeshBuilder3
Public PlaneWait As Integer

Public BombFrame As Direct3DRMFrame3
Public BombMesh As Direct3DRMMeshBuilder3
Public BombDistance As Integer

Public PersonFrame(15) As Direct3DRMFrame3
Public PersonMesh As Direct3DRMMeshBuilder3
Public PersonDead(15) As Boolean
Public PersonInHouse(15) As Boolean
Public PersonInChopper(15) As Boolean
Public PersonHome(15) As Boolean
Public PeopleHome As Integer

Public GraveFrame(15) As Direct3DRMFrame3
Public GraveMesh As Direct3DRMMeshBuilder3

Public AmmoFrame(4) As Direct3DRMFrame3
Public AmmoMesh As Direct3DRMMeshBuilder3
Public AmmoDistance(4) As Integer

Public TempFrame As Direct3DRMFrame3

Public Fullscreen As Boolean
Public ScreenX As Long
Public ScreenY As Long
Public UseFiltering As Boolean
Public UseFlat As Boolean
Public UseSky As Boolean
Public Use3D As Boolean

Public Sub SetupDirectX()
    ChDir App.Path
    
    Set DDraw = DirectX.DirectDraw4Create("")

    Set Clipper = DDraw.CreateClipper(0)

    frmGame.Width = ScreenX * Screen.TwipsPerPixelX
    frmGame.Height = ScreenY * Screen.TwipsPerPixelY
    If Not Fullscreen Then frmGame.BorderStyle = 1: frmGame.Caption = "Chopper"
    If Fullscreen Then DDraw.SetDisplayMode ScreenX, ScreenY, 16, 0, DDSDM_DEFAULT

    Clipper.SetHWnd frmGame.Hwnd
    
    ' init d3drm and main frames
    Set D3DRM = DirectX.Direct3DRMCreate()
    Set SceneFrame = D3DRM.CreateFrame(Nothing)
    Set CameraFrame = D3DRM.CreateFrame(SceneFrame)
    
    SceneFrame.SetSceneBackgroundRGB 0.5, 0.9, 1     'make background blue
    
    On Error GoTo ErrHandler
    
    If Not Use3D Then
        Set Device = D3DRM.CreateDeviceFromClipper(Clipper, "IID_IDirect3DRGBDevice", ScreenX, ScreenY)
    Else
        Set Device = D3DRM.CreateDeviceFromClipper(Clipper, "IID_IDirect3DHALDevice", ScreenX, ScreenY)
    End If
    
    Device.SetQuality D3DRMFILL_SOLID + D3DRMLIGHT_ON + D3DRMSHADE_GOURAUD
    Device.SetDither D_TRUE
    'Device.SetShades 1

    Set Viewport = D3DRM.CreateViewport(Device, CameraFrame, 0, 0, ScreenX, ScreenY)
    Viewport.SetBack 1050
                  
    If UseFiltering Then Device.SetTextureQuality D3DRMTEXTURE_LINEAR  'smooth out textures
    If UseFlat Then Device.SetQuality D3DRMRENDER_FLAT
    
    SetupDSound
    SetupDInput
    If UseJoystick Then Call JoyStick_Init(JOYSTICKID1)
    SetupLights
    frmGame.Show

    Exit Sub
ErrHandler:
    MsgBox "There was an error initializing DirectX. Your computer my not have a compatible 3D card installed. If this is true, go in the Config app and turn off the Use 3D Card option.", vbCritical, "DirectX Error"
End Sub

Private Sub SetupLights()
    Set Light = D3DRM.CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 1, 1, 1)
    Set LightFrame = D3DRM.CreateFrame(SceneFrame)
    LightFrame.SetPosition Nothing, 0, 0, 0
    Light.SetRange 1000!
    LightFrame.AddLight Light
    
    SceneFrame.AddLight D3DRM.CreateLightRGB(D3DRMLIGHT_AMBIENT, 1, 1, 1)
End Sub

Public Sub DelayGame()
    StartTick = DirectX.TickCount
    NowTime = DirectX.TickCount
    Do Until NowTime - LastTick > Delay
        DoEvents
        NowTime = DirectX.TickCount
    Loop
    LastTick = NowTime
End Sub

Public Sub SetupWorld()
    Set GroundFrame = D3DRM.CreateFrame(SceneFrame)
    Set GroundMesh = D3DRM.CreateMeshBuilder
    Set RiverMesh = D3DRM.CreateMeshBuilder
    RiverMesh.LoadFromFile "river.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    Call MakeWall(GroundMesh, -1000, 0, 2150, 1000, 0, 2100, 1000, 0, -2000, -1000, 0, -2000, 50, 25, 0, 0, 0, "grass.bmp")
    GroundFrame.AddVisual GroundMesh
    RiverMesh.ScaleMesh 21, 20, 20
    GroundFrame.AddVisual RiverMesh
    
    Set TreeMesh = D3DRM.CreateMeshBuilder
    TreeMesh.LoadFromFile "tree.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    
    For u = 0 To NumTrees - 1
        Randomize Timer
        Set TreeFrame(u) = D3DRM.CreateFrame(SceneFrame)
        TreeFrame(u).AddVisual TreeMesh
        TreeFrame(u).SetPosition Nothing, (Rnd * 800) - 400, 0, (Rnd * 950) + 900
    Next
        
    Set ChopperFrame = D3DRM.CreateFrame(SceneFrame)
    Set ChopperMesh = D3DRM.CreateMeshBuilder
    ChopperMesh.LoadFromFile "chopper.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    ChopperFrame.AddVisual ChopperMesh
    ChopperFrame.SetPosition Nothing, 0, 6, 1841
    ChopperMesh.ScaleMesh 0.3, 0.3, 0.3 'make chopper smaller
    
    Set SkyFrame = D3DRM.CreateFrame(GroundFrame)
    Set SkyMesh = D3DRM.CreateMeshBuilder
    
    If Not UseSky Then GoTo SkipSky
    Set SkyTexture = D3DRM.LoadTexture("sky.bmp")
    SkyMesh.LoadFromFile "sky.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    SkyMesh.SetTexture SkyTexture
    SkyFrame.AddVisual SkyMesh
   
    SkyMesh.ScaleMesh 55, 70, 55
SkipSky:
    
    Set BladeFrame = D3DRM.CreateFrame(SceneFrame)
    Set BladeMesh = D3DRM.CreateMeshBuilder
    BladeMesh.LoadFromFile "blade.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    BladeFrame.AddVisual BladeMesh
    ChopperFrame.AddVisual BladeFrame
    BladeFrame.SetPosition BladeFrame, 0, 4, 2
    BladeMesh.ScaleMesh 0.3, 0.3, 0.3
    
    Set BaseFrame = D3DRM.CreateFrame(GroundFrame)
    Set BaseMesh = D3DRM.CreateMeshBuilder
    BaseMesh.LoadFromFile "base.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    BaseFrame.AddVisual BaseMesh
    BaseFrame.SetPosition Nothing, 0, 0, 1900
    BaseMesh.ScaleMesh 3, 3, 3
    
    Set PlaneFrame = D3DRM.CreateFrame(GroundFrame)
    Set PlaneMesh = D3DRM.CreateMeshBuilder
    PlaneMesh.LoadFromFile "plane.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    PlaneFrame.AddVisual PlaneMesh
    PlaneFrame.SetPosition Nothing, 0, 150, -1000
    PlaneMesh.ScaleMesh 0.6, 0.6, 0.7
    PlaneWait = 700
    
    Set BombFrame = D3DRM.CreateFrame(GroundFrame)
    Set BombMesh = D3DRM.CreateMeshBuilder
    BombMesh.LoadFromFile "bomb.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    BombFrame.AddVisual BombMesh
    BombFrame.SetPosition Nothing, 0, 0, 1700
    BombMesh.ScaleMesh 0.3, 0.3, 0.3
    
    Set TextFrame = D3DRM.CreateFrame(CameraFrame)
    Set TextMesh = D3DRM.CreateMeshBuilder
    TextMesh.LoadFromFile "intro.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    TextFrame.AddVisual TextMesh
    TextFrame.SetPosition CameraFrame, 0, 0, 50
    TextMesh.ScaleMesh 0.1, 0.1, 0.1
    
    Set HouseMesh = D3DRM.CreateMeshBuilder
    HouseMesh.LoadFromFile "house.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    HouseMesh.ScaleMesh 2, 2, 2
    Randomize
    For y = 0 To 3
        Set HouseFrame(y) = D3DRM.CreateFrame(GroundFrame)
        HouseFrame(y).AddVisual HouseMesh
        HouseFrame(y).SetPosition Nothing, (Rnd * 800) - 400, 0, (Rnd * 1700) - 1500
    Next
    
    Set ExplodeFrame = D3DRM.CreateFrame(GroundFrame)
    Set ExplodeMesh = D3DRM.CreateMeshBuilder
    ExplodeMesh.LoadFromFile "Explode.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    ExplodeFrame.AddVisual ExplodeMesh
    ExplodeFrame.SetPosition Nothing, 0, -1000, 0
    ExplodeMesh.ScaleMesh 1, 1, 1
    
    Set TankMesh = D3DRM.CreateMeshBuilder
    TankMesh.LoadFromFile "tank.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    TankMesh.ScaleMesh 0.3, 0.3, 0.3
    Randomize
    For y = 0 To 3
        Set TankFrame(y) = D3DRM.CreateFrame(GroundFrame)
        TankFrame(y).AddVisual TankMesh
        TankFrame(y).SetPosition Nothing, (Rnd * 800) - 800, 0, (Rnd * 1700) - 1500
    Next
    Set TempFrame = D3DRM.CreateFrame(GroundFrame)
    
    Set AmmoMesh = D3DRM.CreateMeshBuilder
    AmmoMesh.LoadFromFile "ammo.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    AmmoMesh.ScaleMesh 0.1, 0.1, 0.2
    For r = 0 To 4
        Set AmmoFrame(r) = D3DRM.CreateFrame(GroundFrame)
        AmmoFrame(r).AddVisual AmmoMesh
        AmmoFrame(r).SetPosition TankFrame(t), 0, -10, 0
    Next r
    
    AmmoFrame(4).SetPosition Nothing, 0, -100, 0
    
    Set GraveMesh = D3DRM.CreateMeshBuilder
    GraveMesh.LoadFromFile "grave.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    GraveMesh.ScaleMesh 0.16, 0.18, 0.16
    
    For r = 0 To 15
        Set GraveFrame(r) = D3DRM.CreateFrame(GroundFrame)
        GraveFrame(r).AddVisual GraveMesh
        GraveFrame(r).SetPosition Nothing, 0, -10000, 0
    Next r
    
    Set PersonMesh = D3DRM.CreateMeshBuilder
    PersonMesh.LoadFromFile "person.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    PersonMesh.ScaleMesh 0.2, 0.2, 0.2
    
    For r = 0 To 3
        Set PersonFrame(r) = D3DRM.CreateFrame(GroundFrame)
        PersonFrame(r).AddVisual PersonMesh
        PersonFrame(r).SetPosition HouseFrame(0), 0, 0, 0
    Next r

    For r = 4 To 7
        Set PersonFrame(r) = D3DRM.CreateFrame(GroundFrame)
        PersonFrame(r).AddVisual PersonMesh
        PersonFrame(r).SetPosition HouseFrame(1), 0, 0, 0
    Next r
    
    For r = 8 To 11
        Set PersonFrame(r) = D3DRM.CreateFrame(GroundFrame)
        PersonFrame(r).AddVisual PersonMesh
        PersonFrame(r).SetPosition HouseFrame(2), 0, 0, 0
    Next r
    
    Randomize
    For r = 12 To 15
        Set PersonFrame(r) = D3DRM.CreateFrame(GroundFrame)
        PersonFrame(r).AddVisual PersonMesh
        PersonFrame(r).SetPosition HouseFrame(3), 0, 0, 0
    Next r
    
    For r = 0 To 15
        PersonInHouse(r) = True
    Next
    
    PlayingIntro = True
    ChopperLives = 3
End Sub

Public Sub MakeWall(mesh As Direct3DRMMeshBuilder3, X1 As Single, Y1 As Single, z1 As Single, x2 As Single, y2 As Single, z2 As Single, x3 As Single, y3 As Single, z3 As Single, x4 As Single, y4 As Single, z4 As Single, TileX As Single, TileY As Single, r As Single, g As Single, b As Single, texfile As String)
    Dim Face As Direct3DRMFace2
    Dim texture As Direct3DRMTexture3
    Set Face = D3DRM.CreateFace()

    Face.AddVertex X1, Y1, z1
    Face.AddVertex x2, y2, z2
    Face.AddVertex x3, y3, z3
    Face.AddVertex x4, y4, z4
    Face.AddVertex x4, y4, z4
    Face.AddVertex x3, y3, z3
    Face.AddVertex x2, y2, z2
    Face.AddVertex X1, Y1, z1
    
    If texfile = "" Then
        Face.SetColorRGB r, g, b
    Else
        Set tex = D3DRM.LoadTexture(texfile)
        Face.SetTextureCoordinates 0, 0, TileY
        Face.SetTextureCoordinates 1, 0, 0
        Face.SetTextureCoordinates 2, TileX, 0
        Face.SetTextureCoordinates 3, TileX, TileY
        Face.SetTextureCoordinates 4, TileX, TileY
        Face.SetTextureCoordinates 5, TileX, 0
        Face.SetTextureCoordinates 6, 0, 0
        Face.SetTextureCoordinates 7, 0, TileY
        Face.SetTexture tex
    End If
    mesh.AddFace Face
End Sub
