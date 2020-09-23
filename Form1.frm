VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Johna DX7 particle engine Demo"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================
'  johna DX7 Engine version 1.02
'    particle engine beta Demo
'     new particle engine
'       handle with snow,smoke,fire,explosion,or texture particle
'       for explanation contact me at
'       Johna.pop@caramail.com
'
'
'=================================================














'main class engine
Dim DX7 As New johna_DX7

'for the sky
Dim SKY As New cDome_sky

'smoke particle
Dim SMOKE As New cJohna_Particle

'snow particle
Dim SNOW_FALL As New cJohna_Particle

'FIRE particle
Dim FIRE As New cJohna_Particle

'Fountain particle
Dim GEZER As New cJohna_Particle


'for the Floor
Dim FLOOR(3) As D3DVERTEX
Dim FLOOR_SURF As DirectDrawSurface7

'math class
Dim MATH As New cMATH

Dim Zangle As Single
Dim FIRE_rot As Single




Private Sub Form_Load()
Me.Refresh
Me.Show
'DX7.Initialize_DX_Windowed Me.hWnd  'initialize engine in windowed mode
'DX7.INiT_D3D Me.hWnd     'init the 3d engine

'step 1 init the engine
DX7.INIT_engineEX Me.hWnd
Call GAME_LOOP     'call the game loop

End Sub




Sub GAME_LOOP()

'step 2
'create the floor vertices and texture
     Dim V1 As D3DVECTOR
     Dim v2 As D3DVECTOR
     
     V1 = MATH.Vector(-1000, 0, -1000)
     v2 = MATH.Vector(1000, 0, 1000)
     
    DX7.DX_7.CreateD3DVertex V1.x, v2.y, v2.z, 0, 1, 0, 0, 0, FLOOR(0)
    DX7.DX_7.CreateD3DVertex v2.x, v2.y, v2.z, 0, 1, 0, 10, 0, FLOOR(1)
    DX7.DX_7.CreateD3DVertex V1.x, v2.y, V1.z, 0, 1, 0, 0, 10, FLOOR(2)
    DX7.DX_7.CreateD3DVertex v2.x, v2.y, V1.z, 0, 1, 0, 10, 10, FLOOR(3)
    'texture
    Set FLOOR_SURF = DX7.CreateTextureEX(App.Path + "\data\road.jpg", 256, 256)
     
'step 3
'create the sky
SKY.Init DX7, App.Path + "\data\ciel2.x", App.Path + "\data\sky03.jpg"
SKY.SKY_scale = 8


'step 4 for particles
'init snow PositionCenter,TextureFile ,numberof particle,
SNOW_FALL.Init DX7, MATH.Vector(0, 0, 0), App.Path + "\data\snow.bmp", 100, 4, Johna_SNOW, Johna_STATIC_LOCATION, 0, 400, 400, 550

'smoke
SMOKE.Init DX7, MATH.Vector(0, 0, 0), App.Path + "\data\smoke.jpg", 50, 1, Johna_SMOKE

'FIRE
FIRE.Init DX7, MATH.Vector(0, 5, 0), App.Path + "\data\fire3.bmp", 25, 1, Johna_SMOKE


'FIRE
GEZER.Init DX7, MATH.Vector(0, 5, 0), App.Path + "\data\water_a.bmp", 250, 1 / 5, Johna_FONTAIN, , , , , , 2




Randomize Timer
DX7.SET_camera MATH.Vector(0, 10, -80)
'LOOP game
Do
  Zangle = Zangle + 0.3
  If Zangle > 360 Then Zangle = 0
  
  FIRE_rot = FIRE_rot + 0.1
  If FIRE_rot > 360 Then FIRE_rot = 0

  DoEvents
  'til the user toggle ESCAPE key
  If DX7.GetKEY(Johna_KEY_ESCAPE) Then GoTo end_IT

  'check the keys
  Call KEY_check


 'Draw 3d objects
 DX7.D3D_DEV.BeginScene
 DX7.Clear_3D
         'render the sky
         SKY.Render DX7

        'render the floor
        DX7.D3D_DEV.SetTexture 0, FLOOR_SURF
        DX7.D3D_DEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, FLOOR(0), 4, D3DDP_DEFAULT
 
        
        
        'render FIRE
        FIRE.SetPosition MATH.Vector(Cos(Zangle * 3.14 / 108) * 80 + 15, 5, Sin(FIRE_rot * 3.14 / 108) * 80 - 25)
        FIRE.Render DX7
        
        
        'render snow
        SNOW_FALL.Render DX7
        
        'render smoke and animate it
        SMOKE.SetPosition MATH.Vector(Cos(Zangle * 3.14 / 108) * 80, 5, Sin(Zangle * 3.14 / 108) * 80)
        SMOKE.Render DX7
        
        'render GEZER
        GEZER.SetPosition MATH.Vector(Cos(Zangle * 3.14 / 108) * 20, 5, Sin(FIRE_rot * 3.14 / 108) * 20)
        GEZER.Render DX7
        
        
  
 DX7.D3D_DEV.EndScene
 'end drawind
 
 
 
 
 
 
 
  'print info FPS
  DX7.BAK.DrawText 10, 20, "Frames PER seconde=" + Str(DX7.FramesPerSec), 0
  DX7.BAK.DrawText 10, 35, "PRESS SPACE for reset camera", 0
  
  DX7.BAK.DrawText 10, 48, "Use key for walking", 0
  DX7.BAK.DrawText 10, 60, "Use Left ControlKey for more speed", 0
  
  
  
  
  
 'flipp to the screen
 DX7.FLIPP Me.hWnd
 
  
Loop
'end

end_IT:
DX7.FreeDX Me.hWnd
End


End Sub




Sub KEY_check()


If DX7.GetKEY(DIK_UP) Then
  DX7.Camera_Move_Foward 1 / 8
  
End If
  
If DX7.GetKEY(DIK_RCONTROL) Then
  DX7.Camera_Move_Foward 1
  
End If

If DX7.GetKEY(DIK_DOWN) Then
  DX7.Camera_Move_Backward 1 / 8
  
End If

If DX7.GetKEY(DIK_LEFT) Then _
  DX7.Camera_Move_Left 0.0005

If DX7.GetKEY(DIK_RIGHT) Then _
  DX7.Camera_Move_Right 0.0005


If DX7.GetKEY(DIK_NUMPAD8) Then _
  DX7.Camera_Move_UP 0.0005

If DX7.GetKEY(DIK_NUMPAD2) Then _
  DX7.Camera_Move_DOWN 0.0005



If DX7.GetKEY(DIK_NUMPAD4) Then _
  DX7.CAM_step_LEFT 1

If DX7.GetKEY(DIK_NUMPAD6) Then _
  DX7.CAM_step_RIGHT 1


If DX7.GetKEY(DIK_SPACE) Then _
  DX7.Camera_set_EYE DX7.johna_MakeVector(0, 10, -10)


'check World Bound limits
If DX7.GET_CameraEYE.x >= 1000 Or _
   DX7.GET_CameraEYE.x <= -1000 Or _
   DX7.GET_CameraEYE.z >= 1000 Or _
   DX7.GET_CameraEYE.z <= -1000 Then _
   DX7.REcallCAM


End Sub
