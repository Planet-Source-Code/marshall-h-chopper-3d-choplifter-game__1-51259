VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   0  'None
   Caption         =   "Chopper"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------'
'                   Chopper                 '
'              (c) 2003 GameSoft            '
'-------------------------------------------'
'  All models except helicopter made by me  '
'    Helicopter model from DirectX 8 SDK    '
'-------------------------------------------'

Private Sub Form_Load()
    'get settings
    ReadConfig
    'setup DirectX
    SetupDirectX
    'build world
    SetupWorld
    'go to the main loop
    GameLoop
    'reset all buffers
    CleanUp
End Sub

Private Sub GameLoop()
    StartGameTime = DirectX.TickCount
    Delay = 5
    On Error Resume Next
    
    Do Until QuitFlag
        DelayGame
        DoEvents
        CheckKeyboard
        CheckJoystick
        CheckCamera
        MoveExplosion
        MoveChopper
        MovePeople
        MoveTanks
        MovePlane
        CheckCollisions
        
        Viewport.Clear D3DRMCLEAR_ALL
        Viewport.Render SceneFrame
        Device.Update 'render scene!
     Loop
End Sub

Private Sub CheckKeyboard()
    DInputDevice.GetDeviceStateKeyboard Keyboard
       
    If Keyboard.Key(DIK_LEFT) <> 0 Then
        If Not OnGround Then If RotateSpeed > -0.02 Then RotateSpeed = RotateSpeed - 0.005
    End If
    If Keyboard.Key(DIK_RIGHT) <> 0 Then
        If Not OnGround Then If RotateSpeed < 0.02 Then RotateSpeed = RotateSpeed + 0.005
    End If
        
    If Keyboard.Key(DIK_UP) <> 0 Then If Not OnGround Then If ZSpeed < 1.5 Then ZSpeed = ZSpeed + 0.03
    If Keyboard.Key(DIK_DOWN) <> 0 Then If Not OnGround Then If ZSpeed > -1 Then ZSpeed = ZSpeed - 0.02
    
    If Keyboard.Key(DIK_RETURN) <> 0 Then If YSpeed < 0.2 Then YSpeed = YSpeed + 0.2
    If Keyboard.Key(DIK_RSHIFT) <> 0 Then If YSpeed > -0.3 Then YSpeed = YSpeed - 0.3
    
    'Check for camera buttons
    If Keyboard.Key(DIK_INSERT) <> 0 Then CameraView = CHASE
    If Keyboard.Key(DIK_DELETE) <> 0 Then CameraView = INSIDE
    If Keyboard.Key(DIK_HOME) <> 0 Then CameraView = SKY
    If Keyboard.Key(DIK_END) <> 0 Then CameraView = FREE
    
    If Keyboard.Key(DIK_ESCAPE) <> 0 Then CleanUp
    
    If Keyboard.Key(DIK_S) <> 0 Then ShowPos = Not (ShowPos)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then ChopperShoot = True
End Sub

Private Sub CheckJoystick()
    If UseJoystick = False Then Exit Sub
    Dim Xpos
    Dim Ypos
    
    Call JoyStick_GetPos(JOYSTICKID1)
      
    Xpos = JsInfo.dwXpos
    Ypos = JsInfo.dwYpos
    
    Xpos = Xpos - 32768
    Xpos = Xpos / 32768
    Xpos = Round(Xpos, 1)
    
    Ypos = Ypos - 32768
    Ypos = Ypos / 32768
    Ypos = Round(Ypos, 1)
    
    If JsInfo.dwButtons = 4 Then If YSpeed < 0.2 Then YSpeed = YSpeed + 0.2
    If JsInfo.dwButtons = 8 Then If YSpeed > -0.3 Then YSpeed = YSpeed - 0.3
    If JsInfo.dwButtons = 1 Then ChopperShoot = True
    
    If Ypos < -0.5 Then If Not OnGround Then If ZSpeed < 1.5 Then ZSpeed = ZSpeed + 0.03
    If Ypos > 0.5 Then If Not OnGround Then If ZSpeed > -1 Then ZSpeed = ZSpeed - 0.02
    If Xpos < -0.5 Then If Not OnGround Then If RotateSpeed > -0.02 Then RotateSpeed = RotateSpeed - 0.005
    If Xpos > 0.5 Then If Not OnGround Then If RotateSpeed < 0.02 Then RotateSpeed = RotateSpeed + 0.005
End Sub

Private Sub CheckCamera()
    Dim CameraPos As D3DVECTOR
    
    If CameraView = CHASE Then CameraFrame.SetPosition ChopperFrame, 0, 20, -50
    If CameraView = SKY Then CameraFrame.SetPosition ChopperFrame, 0, 100, 0
    CameraFrame.LookAt ChopperFrame, Nothing, D3DRMCONSTRAIN_Z
    If CameraView = INSIDE Then
        CameraFrame.SetPosition ChopperFrame, 0, 1, -10
        CameraFrame.LookAt ChopperFrame, Nothing, D3DRMCONSTRAIN_Z
        CameraFrame.SetPosition ChopperFrame, 0, 0, 8
    End If
    
    'keep camera from going inside base
    CameraFrame.GetPosition Nothing, CameraPos
    If CameraView <> INSIDE Then
        If CameraPos.x < 21 And CameraPos.x > -28 Then
            If CameraPos.z > 1860 And CameraPos.z < 1926 Then
                If CameraPos.z < 1929 Then CameraPos.z = 1869 Else CameraPos.z = 1930
            End If
        End If
    End If
    
    CameraFrame.SetPosition Nothing, CameraPos.x, CameraPos.y, CameraPos.z
    
    'move 3d text
    If PlayingIntro Then
        If TextMoveCount <= 50 Then
            TextFrame.SetPosition TextFrame, 0, 0, -0.9
            TextMoveCount = TextMoveCount + 1
        ElseIf TextMoveCount >= 50 Then
            TextMoveCount = TextMoveCount + 1
        End If
        If TextMoveCount < 250 And TextMoveCount > 174 Then
            TextFrame.SetPosition TextFrame, 0, -0.3, 0.6
            TextMoveCount = TextMoveCount + 1
        End If
        If TextMoveCount >= 250 Then TextMoveCount = 0: PlayingIntro = False: TextMesh.Empty: TextFrame.SetPosition CameraFrame, 0, 0, 50
    End If
    
    If PlayingEnding Then
        If TextMoveCount <= 50 Then
            TextFrame.SetPosition TextFrame, 0, 0, -0.9
            TextMoveCount = TextMoveCount + 1
        ElseIf TextMoveCount >= 50 Then
            TextMoveCount = TextMoveCount + 1
        End If
        If TextMoveCount < 400 And TextMoveCount > 324 Then
            TextFrame.SetPosition TextFrame, 0.4, -0.2, 0.6
            TextMoveCount = TextMoveCount + 1
        End If
        If TextMoveCount >= 400 Then CleanUp
    End If
End Sub

Private Sub MoveChopper()
    If Dead Then Exit Sub
    ChopperFrame.GetPosition Nothing, ChopperPos
    If ShowPos Then Caption = "Chopper - XPos = " & Round(ChopperPos.x, 0) & " ZPos = " & Round(ChopperPos.z, 0)
        
    OnGround = False
    If ChopperPos.x > 850 Then ChopperPos.x = 850
    If ChopperPos.x < -850 Then ChopperPos.x = -850
    If ChopperPos.z > 1970 Then ChopperPos.z = 1970
    If ChopperPos.z < -1850 Then ChopperPos.z = -1850
    If ChopperPos.y > 200 Then ChopperPos.y = 200
    If ChopperPos.y < 6.1 Then ChopperPos.y = 6: OnGround = True: If ZSpeed > 0.8 Or ZSpeed < -0.8 Then ChopperDie
    
    ChopperFrame.SetPosition Nothing, ChopperPos.x, ChopperPos.y, ChopperPos.z

    If ZSpeed >= 0.01 Then ZSpeed = ZSpeed - 0.01 'simulate deceleration
    If ZSpeed <= -0.01 Then ZSpeed = ZSpeed + 0.01   'keep Chopper from rolling backwards
    
    If YSpeed < 0.02 And YSpeed > -0.02 Then YSpeed = 0 'prevent drift
    If YSpeed >= 0.01 Then YSpeed = YSpeed - 0.01 'simulate slowing of vertical climb
    If YSpeed <= -0.01 Then YSpeed = YSpeed + 0.01
    
    If RotateSpeed > -0.003 And RotateSpeed < 0.003 Then RotateSpeed = 0 'prevent drift
    If RotateSpeed >= 0.003 Then RotateSpeed = RotateSpeed - 0.003
    If RotateSpeed <= -0.003 Then RotateSpeed = RotateSpeed + 0.003
    
    If ZSpeed <> 0 Then ChopperFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, RotateSpeed
    ChopperFrame.SetPosition ChopperFrame, 0, 0, ZSpeed
    
    ChopperFrame.SetPosition ChopperFrame, 0, YSpeed, 0
    For p = 0 To 15
        If PersonInChopper(p) And Not PersonDead(p) And Not AtBase And Not PersonHome(p) Then PersonFrame(p).SetPosition ChopperFrame, 0, -4, 2
    Next
    
    If ChopperPos.y > 9 Then Idle = 0.03 Else Idle = 0
    RotAmount = Abs(ZSpeed / 4) + 0.1 + Idle
    ChopperBuffer.SetFrequency RotAmount * 25000 + 20000
    
    BladeFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, RotAmount
    
    If ChopperShoot Then
        'init shot
        ChopperShoot = False
        AmmoDistance(4) = 1
        ShootDelay = 0
        AmmoFrame(4).SetPosition ChopperFrame, 0, 0.2, -2 'increase 0.2 to make gun point more down
        AmmoFrame(4).LookAt ChopperFrame, Nothing, D3DRMCONSTRAIN_Z
        AmmoFrame(4).SetPosition ChopperFrame, 0, 0, 3
    ElseIf Not ChopperShoot And AmmoDistance(4) > 0 Then
        'move shot
        AmmoFrame(4).SetPosition AmmoFrame(4), 0, 0, 3
        AmmoDistance(4) = AmmoDistance(4) + 1
    ElseIf AmmoDistance(4) > 200 Then
        'reset shot
        AmmoFrame(4).SetPosition Nothing, 0, -100, 0
        AmmoDistance(4) = 0
    End If
End Sub

Private Sub MoveExplosion()
    If Dead And ExplodeCount < 10 Then
        'move explosion
        ExplodeMesh.ScaleMesh 1.1, 1.1, 1.1
        ExplodeFrame.SetPosition ExplodeFrame, 0, 3, 0
        ChopperFrame.SetPosition ChopperFrame, 0, -(ChopperPos.y / 9), 0
        ChopperFrame.GetPosition Nothing, ChopperPos
        ExplodeCount = ExplodeCount + 1
    ElseIf ExplodeCount > 120 And Dead Then
        'reset explosion
        Dead = False
        ExplodeCount = 0
        ChopperFrame.SetPosition Nothing, 0, 6, 1841
        ExplodeMesh.Empty
        ExplodeMesh.LoadFromFile "explode.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        ExplodeFrame.SetPosition Nothing, 0, -100, 0
        PlaySound ChopperBuffer, False, True
        ZSpeed = 0
        YSpeed = 0
    ElseIf Dead Then
        'increment counter
        ExplodeCount = ExplodeCount + 1
        ExplodeFrame.SetPosition ExplodeFrame, 0, 3, 0
    End If

    For a = 0 To 3
        If TankDead(a) And TankExplodeCount(a) < 10 Then
            ExplodeMesh.ScaleMesh 1.05, 1.05, 1.05
            ExplodeFrame.SetPosition ExplodeFrame, 0, 2, 0
            TankExplodeCount(a) = TankExplodeCount(a) + 1
            TankFrame(a).SetPosition Nothing, 0, -100, 0
        ElseIf TankExplodeCount(a) > 100 And TankDead(a) Then
            'reset explosion
            TankExplodeCount(a) = 0
            ExplodeMesh.Empty
            ExplodeMesh.LoadFromFile "explode.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
            ExplodeFrame.SetPosition Nothing, 0, -1100, 0
        ElseIf TankDead(a) Then
            TankExplodeCount(a) = TankExplodeCount(a) + 1
            ExplodeFrame.SetPosition ExplodeFrame, 0, 2, 0
        End If
    Next

End Sub

Private Sub MoveTanks()
    Dim DestVec As D3DVECTOR
    Dim TankPos(3) As D3DVECTOR
    ChopperFrame.GetPosition Nothing, DestVec
    For t = 0 To 3
        If TankDead(t) Then GoTo SkipTank
        TempFrame.SetPosition Nothing, DestVec.x, 0, DestVec.z
        TankFrame(t).LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
        TankFrame(t).SetPosition TankFrame(t), 0, 0, (Rnd * 0.2) + 0.1
        TankFrame(t).GetPosition Nothing, TankPos(t)
        If TankPos(t).z > 0 Then TankPos(t).z = 0 'keep tanks from going into US territory
        TankFrame(t).SetPosition Nothing, TankPos(t).x, TankPos(t).y, TankPos(t).z
        
        If AmmoDistance(t) < 300 And AmmoDistance(t) <> 0 Then
            AmmoDistance(t) = AmmoDistance(t) + 1
            AmmoFrame(t).SetPosition AmmoFrame(t), 0, 0, 3
        ElseIf AmmoDistance(t) = 0 Then
            AmmoDistance(t) = 1
            AmmoFrame(t).SetPosition TankFrame(t), 0, 7, 2
            AmmoFrame(t).LookAt ChopperFrame, Nothing, D3DRMCONSTRAIN_Z
        ElseIf AmmoDistance(t) >= 300 Then
            AmmoDistance(t) = 0
        End If
SkipTank:
    Next t
    CameraFrame.GetPosition Nothing, DestVec
    SkyFrame.SetPosition Nothing, DestVec.x, 0, DestVec.z
    Dim SkyPos As D3DVECTOR
    SkyFrame.GetPosition Nothing, SkyPos

    If SkyPos.x > 500 Then SkyPos.x = 500
    If SkyPos.x < -500 Then SkyPos.x = -500
    If SkyPos.z > 1500 Then SkyPos.z = 1500
    If SkyPos.z < -1500 Then SkyPos.z = -1500
    SkyFrame.SetPosition Nothing, SkyPos.x, 0, SkyPos.z
End Sub

Private Sub MovePlane()
    Dim PlanePos As D3DVECTOR
    Dim BombPos As D3DVECTOR
    PlaneFrame.GetPosition Nothing, PlanePos
    BombFrame.GetPosition Nothing, BombPos
    
    ChopperFrame.GetPosition Nothing, ChopperPos
    TempFrame.SetPosition Nothing, ChopperPos.x, 150, ChopperPos.z
    If PlanePos.z > 100 Then TempFrame.SetPosition Nothing, 0, 150, -1000
    If ChopperPos.x < 100 Then PlaneWait = PlaneWait + 1 Else BombFrame.SetPosition Nothing, 0, -1000, 0: Exit Sub
    If PlaneWait < 10 Then PlaneFrame.LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z: BombFrame.SetPosition Nothing, 0, -1000, 0: Exit Sub
    If PlaneWait > 2200 Then PlaneWait = 0
        
    PlaneFrame.SetPosition PlaneFrame, 0, 0, 2
        
    If BombDistance < 300 And BombDistance <> 0 Then
        BombDistance = BombDistance + 1
        BombFrame.SetPosition BombFrame, 0, 0, 3
        If BombPos.z > 500 Then BombDistance = 999
    ElseIf BombDistance = 0 Then
        BombDistance = 1
        BombFrame.SetPosition PlaneFrame, 0, 0, 0
        BombFrame.LookAt ChopperFrame, Nothing, D3DRMCONSTRAIN_Z
    ElseIf BombDistance >= 300 Then
        BombDistance = 0
    End If
End Sub

Private Sub MovePeople()
    Dim DestVec As D3DVECTOR
    Dim PersonPos(15) As D3DVECTOR
    ChopperFrame.GetPosition Nothing, DestVec
    For t = 0 To 15
        If PersonHome(t) Or PersonDead(t) Then d = d + 1
        If PersonDead(t) Or PersonInHouse(t) Or PersonInChopper(t) Or PersonHome(t) Then GoTo SkipPerson
        If (Rnd * 10) + 1 < 3 Then GoTo SkipPerson
        TempFrame.SetPosition Nothing, DestVec.x, 0, DestVec.z
        PersonFrame(t).LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
        PersonFrame(t).SetPosition PersonFrame(t), (Rnd * 0.1) - 0.1, 0, (Rnd * 0.5) + 0.1
        PersonFrame(t).GetPosition Nothing, PersonPos(t)
        If PersonPos(t).z > 400 Then PersonPos(t).z = 400 'keep person from going across the river
        PersonFrame(t).SetPosition Nothing, PersonPos(t).x, PersonPos(t).y, PersonPos(t).z
        
        If PersonPos(t).x > DestVec.x - 2 And PersonPos(t).x < DestVec.x + 2 Then
            If PersonPos(t).z > DestVec.z - 2 And PersonPos(t).z < DestVec.z + 2 Then
                If OnGround And Not Dead And Not PersonDead(t) And Not PersonHome(t) Then
                    If PeopleInChopper < 4 Then PersonInChopper(t) = True: PeopleInChopper = PeopleInChopper + 1

                End If
            End If
        End If
        
SkipPerson:
    If AtBase And PersonInChopper(t) And Not PersonDead(t) And Not PersonHome(t) And Not PersonInHouse(t) Then
        TempFrame.SetPosition Nothing, 0, 0, 1885
        PersonFrame(t).LookAt TempFrame, Nothing, D3DRMCONSTRAIN_Z
        PersonFrame(t).SetPosition PersonFrame(t), (Rnd * 0.1) - 0.1, 0, (Rnd * 0.5) + 0.1
        PersonFrame(t).GetPosition Nothing, PersonPos(t)
        If PersonPos(t).z > 1876 And Not PersonHome(t) Then
            'person reached safety
            PersonHome(t) = True
            PeopleInChopper = PeopleInChopper - 1
            PeopleHome = PeopleHome + 1
        End If
    End If
    Next t
    
    If d = 16 Then
        'player rescued all people
        If PlayingEnding Then GoTo SkipEnding
        For w = 0 To 15
            If PersonDead(w) Then pd = True
        Next
        If pd = False Then 'if all people rescued with no dead
            TextMesh.LoadFromFile "youwin.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
            TextMesh.ScaleMesh 0.06, 0.06, 0.06
        ElseIf pd = True Then 'if a person was hurt
            TextMesh.LoadFromFile "gameover.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
            TextMesh.ScaleMesh 0.045, 0.06, 0.06
        End If
        PlayingEnding = True
    End If
SkipEnding:
End Sub

Private Sub CheckCollisions()
    If Dead Then Exit Sub
    Dim AmmoPos(4) As D3DVECTOR
    Dim BombPos As D3DVECTOR
    Dim TankPos(3) As D3DVECTOR
    Dim PersonPos(15) As D3DVECTOR
    Dim HousePos(3) As D3DVECTOR
    
    ChopperFrame.GetPosition Nothing, ChopperPos
    BombFrame.GetPosition Nothing, BombPos
    
    For e = 0 To 4
        AmmoFrame(e).GetPosition Nothing, AmmoPos(e)
        If e <> 4 Then TankFrame(e).GetPosition Nothing, TankPos(e)
        If e <> 4 Then HouseFrame(e).GetPosition Nothing, HousePos(e)
    Next
    
    For e = 0 To 15
        PersonFrame(e).GetPosition Nothing, PersonPos(e)
    Next
    
    For w = 0 To 3
        If ChopperPos.x > AmmoPos(w).x - 2 And ChopperPos.x < AmmoPos(w).x + 2 Then
            If ChopperPos.z > AmmoPos(w).z - 2 And ChopperPos.z < AmmoPos(w).z + 2 Then
                If ChopperPos.y > AmmoPos(w).y - 2 And ChopperPos.y < AmmoPos(w).y + 2 Then
                    AmmoDistance(w) = 999 'destroy bullet
                    ChopperDie
                End If
            End If
        End If
    Next
    
    If ChopperPos.x > BombPos.x - 2 And ChopperPos.x < BombPos.x + 2 Then
        If ChopperPos.z > BombPos.z - 2 And ChopperPos.z < BombPos.z + 2 Then
            If ChopperPos.y > BombPos.y - 2 And ChopperPos.y < BombPos.y + 2 Then
                BombDistance = 999 'destroy bomb
                ChopperDie
            End If
        End If
    End If
    
    For q = 0 To 3
        If TankPos(q).x > AmmoPos(4).x - 4 And TankPos(q).x < AmmoPos(4).x + 4 Then
            If TankPos(q).z > AmmoPos(4).z - 7 And TankPos(q).z < AmmoPos(4).z + 7 Then
                If TankPos(q).y > AmmoPos(4).y - 6 And TankPos(q).y < AmmoPos(4).y + 6 Then
                    AmmoDistance(4) = 999 'destroy bullet
                    AmmoFrame(q).SetPosition Nothing, 0, -100, 0
                    TankDie (q)
                End If
            End If
        End If
    Next
    
    For q = 0 To 15
        If PersonPos(q).x > AmmoPos(4).x - 1.4 And PersonPos(q).x < AmmoPos(4).x + 1.4 Then
            If PersonPos(q).z > AmmoPos(4).z - 1.4 And PersonPos(q).z < AmmoPos(4).z + 1.4 Then
                If PersonPos(q).y > AmmoPos(4).y - 1 And PersonPos(q).y < AmmoPos(4).y + 16 Then
                    AmmoDistance(4) = 999
                    AmmoFrame(4).SetPosition AmmoFrame(4), 0, -1000, 0
                    If Not PersonDead(q) Or Not PersonInHouse(q) Then PersonDie (q)
                End If
            End If
        End If
    Next
    
    For q = 0 To 3
        If HousePos(q).x > AmmoPos(4).x - 5 And HousePos(q).x < AmmoPos(4).x + 5 Then
            If HousePos(q).z > AmmoPos(4).z - 20 And HousePos(q).z < AmmoPos(4).z + 20 Then
                If HousePos(q).y > AmmoPos(4).y - 1 And HousePos(q).y < AmmoPos(4).y + 14 Then
                    AmmoDistance(4) = 999
                    AmmoFrame(4).SetPosition AmmoFrame(4), 0, -1000, 0
                    PlaySound DoorBuffer, True, False
                    For u = q * 4 To q * 4 + 3
                        PersonInHouse(u) = False
                    Next
                End If
            End If
        End If
    Next
    
    AtBase = False
    If ChopperPos.x > -30 And ChopperPos.x < 40 Then
        If ChopperPos.z > 1810 And ChopperPos.z < 1855 Then
            If OnGround Then
                AtBase = True
            End If
        End If
    End If
    
    For y = 0 To 3
        If ChopperPos.z > 700 And TankDead(y) Then
            're-create tanks
            TankDead(y) = False
            TankFrame(y).SetPosition Nothing, (Rnd * 800) - 800, 0, (Rnd * 1700) - 1500
            TankExplodeCount(y) = 0
        End If
    Next
End Sub

Private Sub ChopperDie()
    ExplodeFrame.SetPosition ChopperFrame, 0, -6, 0
    Dead = True
    ChopperLives = ChopperLives - 1
    ChopperBuffer.Stop
    ChopperBuffer.SetCurrentPosition 0
    PlaySound ExplodeBuffer, True, False
    
    For y = 0 To 15
         If PersonInChopper(y) And Not PersonHome(y) Then PersonDie (y): PeopleInChopper = PeopleInChopper - 1
    Next y
    
    If ChopperLives = -1 Then
        PlayingEnding = True
        TextMesh.LoadFromFile "gameover.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        TextMesh.ScaleMesh 0.045, 0.06, 0.06
    End If
End Sub

Private Sub TankDie(TankNum As Integer)
    ExplodeFrame.SetPosition TankFrame(TankNum), 0, 0, 0
    TankDead(TankNum) = True
    PlaySound ExplodeBuffer, True, False
End Sub

Private Sub PersonDie(PersonNum As Integer)
    Dim PersonVec As D3DVECTOR
    PersonFrame(PersonNum).GetPosition Nothing, PersonVec
    GraveFrame(PersonNum).SetPosition Nothing, PersonVec.x, 0, PersonVec.z
    PersonFrame(PersonNum).SetPosition PersonFrame(PersonNum), 0, -1000, 0
    PersonInChopper(PersonNum) = False
    PersonDead(PersonNum) = True
End Sub

Private Sub ReadConfig()
    On Error GoTo NoFile 'if file exists, get setting
                         'else create file with setting
    Open "chopper.cfg" For Input As #1
        Input #1, TmpF
        Input #1, TmpS
        Input #1, TmpX
        Input #1, TmpY
        Input #1, TmpJ
        Input #1, TmpB
        Input #1, TmpR
        Input #1, Tmp3D
    Close #1
    If TmpF = "TRUE" Then Fullscreen = True
    If TmpJ = "TRUE" Then UseJoystick = True
    If TmpB = "TRUE" Then UseFiltering = True
    If TmpR = "TRUE" Then UseFlat = True
    If TmpS = "TRUE" Then UseSky = True
    If Tmp3D = "TRUE" Then Use3D = True

    ScreenX = Val(TmpX)
    ScreenY = Val(TmpY)
    Exit Sub
    
NoFile:
    Close #1
    Open "chopper.cfg" For Output As #1
        Print #1, "FALSE"
        Print #1, "FALSE"
        Print #1, "640"
        Print #1, "480"
        Print #1, "FALSE"
        Print #1, "TRUE"
        Print #1, "FALSE"
        Print #1, "TRUE"
    Close #1
    UseFiltering = True

    ScreenX = 640
    ScreenY = 480
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'exit the loop
    QuitFlag = True
End Sub

Private Sub CleanUp()
    'clean up our program
    Set D3DRM = Nothing
    Set ChopperFrame = Nothing
    Set ChopperMesh = Nothing
    
    Set BaseFrame = Nothing
    Set BaseMesh = Nothing
    
    Set BladeMesh = Nothing
    Set BladeFrame = Nothing
    
    Set PlaneMesh = Nothing
    Set PlaneFrame = Nothing
    
    Set BombMesh = Nothing
    Set BombFrame = Nothing

    Set GroundFrame = Nothing
    Set GroundMesh = Nothing
    
    Set TreeMesh = Nothing
    For i = 0 To NumTrees - 1
        Set TreeFrame(i) = Nothing
    Next

    Set ExplodeFrame = Nothing
    Set ExplodeMesh = Nothing
    
    Set TextFrame = Nothing
    Set TextMesh = Nothing

    Set Light = Nothing
    Set LightFrame = Nothing
    
    Set RiverMesh = Nothing
    Set RiverFrame = Nothing
    
    Set AmmoMesh = Nothing
    Set TankMesh = Nothing
    Set HouseMesh = Nothing
    For i = 0 To 4
        Set AmmoFrame(i) = Nothing
        If i <> 4 Then Set TankFrame(i) = Nothing
        If i <> 4 Then Set HouseFrame(i) = Nothing
    Next
    
    Set PersonMesh = Nothing
    For i = 0 To 15
        Set PersonFrame(i) = Nothing
    Next
    
    Unload Me
    End
End Sub

