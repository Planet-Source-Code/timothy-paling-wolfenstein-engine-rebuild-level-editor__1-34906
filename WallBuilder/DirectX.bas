Attribute VB_Name = "Direct3D"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Type CUSTOMVERTEX
    Position As D3DVECTOR    '3d position for vertex
    normal As D3DVECTOR
    Color As Long           'color of the vertex
    tu1 As Single            'texture map coordinate
    tv1 As Single            'texture map coordinate
    tu2 As Single
    tv2 As Single
End Type

'Public Type SortStruct
'    SqD As Single
'    Li As Integer
'End Type
'
'Dim LightSort() As SortStruct

' Our custom FVF, which describes our custom vertex structure
Public Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_DIFFUSE Or D3DFVF_TEX2)

Public g_DX As New DirectX8
Public g_D3DX As New D3DX8
Public g_D3D As Direct3D8              ' Used to create the D3DDevice
Public g_D3DDevice As Direct3DDevice8  ' Our rendering device
Public g_VB As Direct3DVertexBuffer8   ' Holds our vertex data
'Public g_Texture As Direct3DTexture8   ' Our texture
Public d3dpp As D3DPRESENT_PARAMETERS

Public matWorld As D3DMATRIX
Public matWorldPuck As D3DMATRIX
Public matWorldPuck1 As D3DMATRIX
Public matSpriteRot As D3DMATRIX
Public matRollX As D3DMATRIX
Public matRollY As D3DMATRIX
Public matRollZ As D3DMATRIX
Public matTrans As D3DMATRIX

Public Const g_pi = 3.1415

Dim v As CUSTOMVERTEX
Dim sizeOfVertex As Long

Dim FrameCount As Long
Dim FrameCount2 As Long

Dim emptyText As Direct3DTexture8
Dim defMat As D3DMATERIAL8

Dim SpritesIn As Boolean
Public RenderSuspend As Boolean






Public Function InitD3D(hWnd As Long) As Boolean
    On Local Error Resume Next
    

    Set g_D3D = g_DX.Direct3DCreate()
    If g_D3D Is Nothing Then Exit Function
    
    Dim Mode As D3DDISPLAYMODE
    g_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode
   
    d3dpp.Windowed = 1
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = Mode.Format
    d3dpp.BackBufferCount = 1
    d3dpp.EnableAutoDepthStencil = 1
    d3dpp.AutoDepthStencilFormat = D3DFMT_D16
    
    Set g_D3DDevice = g_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, _
                                      D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If g_D3DDevice Is Nothing Then Exit Function
  
    g_D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    g_D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    g_D3DDevice.SetRenderState D3DRS_FILLMODE, 2
    
    SetTextureStages
    
    
    D3DXMatrixIdentity matWorld
    g_D3DDevice.SetTransform D3DTS_WORLD, matWorld

    Dim matView As D3DMATRIX
    D3DXMatrixLookAtLH matView, vec3(0#, 1#, -5#), vec3(0#, 0#, 0#), vec3(0#, 1#, 0#)
                                 
    g_D3DDevice.SetTransform D3DTS_VIEW, matView

    Dim matProj As D3DMATRIX
    D3DXMatrixPerspectiveFovLH matProj, g_pi / 6, 0.75, 0.01, 50
    g_D3DDevice.SetTransform D3DTS_PROJECTION, matProj

    sizeOfVertex = Len(v)
    ReDim LightSort(0)
    
    Dim mcol As D3DCOLORVALUE
    With mcol
        .a = 1
        .b = 1
        .g = 1
        .r = 1
    End With
    With defMat
        .Ambient = mcol
        .diffuse = mcol
        '.emissive = mcol
        .specular = mcol
        .power = 1
    End With
    
    g_D3DDevice.SetMaterial defMat

    InitD3D = True
End Function

Public Sub Render()
    
    If RenderSuspend = True Then Exit Sub
    
    FrameCount = FrameCount + 1
    FrameCount2 = FrameCount2 + 1
    SpritesIn = False
    
    If g_D3DDevice.TestCooperativeLevel <> D3D_OK Then
        Exit Sub 'Something's not working...
    ElseIf g_D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        g_D3DDevice.Reset d3dpp
        Exit Sub
    End If
    
    'And if all the above fails resort to a nice Win32 routine
    If GetForegroundWindow() <> Form1.hWnd Then
        Exit Sub
    End If
    'End If
    
'    If Form1.Check3.Value = 1 Then
'        g_D3DDevice.SetRenderState D3DRS_LIGHTING, 1
'
'    Else
        g_D3DDevice.SetRenderState D3DRS_LIGHTING, 0
'    End If
    
    
    
    
    
    If g_D3DDevice Is Nothing Then Exit Sub
    g_D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF000000, 1#, 0

    

If FrameCount2 Mod Form7.Slider2.Value = 0 Then
        If GetForegroundWindow() = Form1.hWnd Then
            ReadKyb
        End If
        
        If aKeys(Form4.Text11.Text) Then 'A - Roll Fwd
            DefaultVP.RotationX = DefaultVP.RotationX + 0.1
        ElseIf aKeys(Form4.Text12.Text) Then 'Z - Roll Bwd
            DefaultVP.RotationX = DefaultVP.RotationX - 0.1
        ElseIf aKeys(Form4.Text7.Text) Then 'E - Move Fwd
            Angle = DefaultVP.RotationY
            Distance = 0.2 * Sin(Angle)
            DefaultVP.Position.X = DefaultVP.Position.X + Distance
            Distance = 0.2 * Cos(Angle)
            DefaultVP.Position.z = DefaultVP.Position.z + Distance
            'DefaultVP.Position.X = DefaultVP.Position.X + 0.1
        ElseIf aKeys(Form4.Text8.Text) Then 'D - Move Bwd
            Angle = DefaultVP.RotationY
            Distance = 0.2 * Sin(Angle)
            DefaultVP.Position.X = DefaultVP.Position.X - Distance
            Distance = 0.2 * Cos(Angle)
            DefaultVP.Position.z = DefaultVP.Position.z - Distance
            'DefaultVP.Position.X = DefaultVP.Position.X - 0.1
        ElseIf aKeys(Form4.Text9.Text) Then 'S - Move Left
            DefaultVP.RotationY = DefaultVP.RotationY - 0.1
        ElseIf aKeys(Form4.Text10.Text) Then 'F - Move right
            DefaultVP.RotationY = DefaultVP.RotationY + 0.1
        End If
    FrameCount2 = 0
End If

If FrameCount Mod Form7.Slider1.Value = 0 Then

        'If Form1.Check6.Value = 1 Then
        '    MoveVal = 1
        'Else
        '    MoveVal = 0.1
        'End If
        MoveVal = Form1.Text1.Text
        If aKeys(Form4.Text1.Text) Then  'Up Arrow
            
                DefaultPuck.CurrentX = DefaultPuck.CurrentX + MoveVal
       
        ElseIf aKeys(Form4.Text4.Text) Then 'Left Arrow
            
                DefaultPuck.CurrentY = DefaultPuck.CurrentY - MoveVal
             
        ElseIf aKeys(Form4.Text3.Text) Then 'Right Arrow
            
                DefaultPuck.CurrentY = DefaultPuck.CurrentY + MoveVal
            
        ElseIf aKeys(Form4.Text2.Text) Then  'Down Arrow
            
                DefaultPuck.CurrentX = DefaultPuck.CurrentX - MoveVal
            
        ElseIf aKeys(Form4.Text5.Text) Then 'Home
            
            
               DefaultPuck.CurrentYY = DefaultPuck.CurrentYY + 0.1
            
        ElseIf aKeys(Form4.Text6.Text) Then 'End
           
                
                DefaultPuck.CurrentYY = DefaultPuck.CurrentYY - 0.1
            
        ElseIf aKeys(Form4.Text13.Text) Then
            Form1.Invoke_Delete
        End If
            
        If Form4.Text14.Text <> "#" Then
            If aKeys(Form4.Text14.Text) Then
                Form1.Invoke_MouseDown 1
                Form1.Invoke_MouseUp
            End If
        End If
        
        If Form4.Text15.Text <> "#" Then
            If aKeys(Form4.Text15.Text) Then
                Form1.Invoke_MouseDown 2
                Form1.Invoke_MouseUp
            End If
        End If
        PurgeKeyDat
        FrameCount = 0
    End If

    For a = 1 To UBound(WorldObjects())
        WorldObjects(a).TempCheck = False
        WorldObjects(a).UnderSelection = False
    Next a
    
    For a = 1 To UBound(WorldObjects())
        With WorldObjects(a)
            If .WorldX = DefaultPuck.CurrentX And .WorldZ = DefaultPuck.CurrentY And .WorldY = DefaultPuck.CurrentYY And .Rotation = DefaultPuck.AnglePI And .InUse = True And .CodeWord = Form1.Combo1.Text Then
                If Form1.TreeView1.Nodes.Item(a).Checked = False Then
                    Form1.TreeView1.Nodes.Item(.NodeIndex).Checked = True
                    .TempCheck = True
                End If
                .UnderSelection = True
                'Apply lighting and heights
                If Form1.Check3.Value = 1 Then
                    If Form1.v1red.Text <> "" And Form1.v1blue.Text <> "" And Form1.v1green.Text <> "" Then
                    If IsNumeric(Form1.v1red.Text) And IsNumeric(Form1.v1blue.Text) And IsNumeric(Form1.v1green.Text) Then
                    .Color1 = RGB(Form1.v1red.Text, Form1.v1blue.Text, Form1.v1green.Text)
                    End If
                    End If
                    
                    If Form1.v2red.Text <> "" And Form1.v2blue.Text <> "" And Form1.v2green.Text <> "" Then
                    If IsNumeric(Form1.v2red.Text) And IsNumeric(Form1.v2blue.Text) And IsNumeric(Form1.v2green.Text) Then
                    .Color2 = RGB(Form1.v2red.Text, Form1.v2blue.Text, Form1.v2green.Text)
                    End If
                    End If
                    
                    If Form1.v3red.Text <> "" And Form1.v3blue.Text <> "" And Form1.v3green.Text <> "" Then
                    If IsNumeric(Form1.v3red.Text) And IsNumeric(Form1.v3blue.Text) And IsNumeric(Form1.v3green.Text) Then
                    .Color3 = RGB(Form1.v3red.Text, Form1.v3blue.Text, Form1.v3green.Text)
                    End If
                    End If
                    
                    If Form1.v4red.Text <> "" And Form1.v4blue.Text <> "" And Form1.v4green.Text <> "" Then
                    If IsNumeric(Form1.v4red.Text) And IsNumeric(Form1.v4blue.Text) And IsNumeric(Form1.v4green.Text) Then
                    .Color4 = RGB(Form1.v4red.Text, Form1.v4blue.Text, Form1.v4green.Text)
                    End If
                    End If
                    
                End If
                
                If Form1.Check6.Value = 1 Then
                    If Form1.V1Y <> "" Then
                    .LSc1R = Form1.V1Y
                    End If
                    
                    If Form1.V2Y <> "" Then
                    .LSc2R = Form1.V2Y
                    End If
                    
                    If Form1.V3Y <> "" Then
                    .LSc3R = Form1.V3Y
                    End If
                    
                    If Form1.V4Y <> "" Then
                    .LSc4R = Form1.V4Y
                    End If
                End If
            End If
        End With
    Next a
    
    For a = 1 To UBound(WorldObjects())
        If WorldObjects(a).InUse = True Then
        If WorldObjects(a).TempCheck = False Then
            If Form1.TreeView1.Nodes.Item(WorldObjects(a).NodeIndex).Checked = True Then
                Form1.TreeView1.Nodes.Item(WorldObjects(a).NodeIndex).Checked = False
            End If
        End If
        End If
    Next a

    ' Setup the world, view, and projection matrices
    SetupMatrices
    'D3DXMatrixTranslation matWorldPuck, DefaultPuck.CurrentX, 0, DefaultPuck.CurrentY
    

    
    g_D3DDevice.BeginScene
        g_D3DDevice.SetRenderState D3DRS_FILLMODE, 2
        g_D3DDevice.SetTexture 0, Nothing
        g_D3DDevice.SetTexture 1, Nothing
        
        g_D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
        '---------------------------------------
        'Draw the gridlines...
        If Form1.Check1.Value = 1 Then
            g_D3DDevice.SetStreamSource 0, GuideVB, sizeOfVertex
            g_D3DDevice.DrawPrimitive D3DPT_LINELIST, 0, SizeConstant / 2
        End If
        
        '---------------------------------------
        'Draw the Puck!
       
       D3DXMatrixTranslation matWorldPuck, DefaultPuck.CurrentX, DefaultPuck.CurrentYY, DefaultPuck.CurrentY
       
        D3DXMatrixRotationY matSpriteRot, DefaultVP.RotationY
        D3DXMatrixMultiply matWorldPuck1, matSpriteRot, matWorldPuck
        g_D3DDevice.SetTransform D3DTS_WORLD, matWorldPuck
        

        Select Case Form1.Combo1.Text
            Case "Wall"
                                
                    If DefaultPuck.AnglePI = 1 Then
                        g_D3DDevice.SetStreamSource 0, PuckVB2, sizeOfVertex
                        g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
                    Else
                        g_D3DDevice.SetStreamSource 0, PuckVB, sizeOfVertex
                        g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
                    End If
                
            Case "Ceiling"
                g_D3DDevice.SetStreamSource 0, PuckVB4, sizeOfVertex
                g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
            Case "Floor"
                g_D3DDevice.SetStreamSource 0, PuckVB3, sizeOfVertex
                g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
            Case "Sprite"
                g_D3DDevice.SetTransform D3DTS_WORLD, matWorldPuck1
                g_D3DDevice.SetStreamSource 0, PuckVB5, sizeOfVertex
                g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
                g_D3DDevice.SetTransform D3DTS_WORLD, matWorldPuck
            
        End Select
        
        
        If Form1.Option1 Then
            g_D3DDevice.SetRenderState D3DRS_FILLMODE, 3
        Else
            g_D3DDevice.SetRenderState D3DRS_FILLMODE, 2
        End If
        
        '----------------------------------------
        'Draw normal objects
        
        
        
        
        For a = 1 To UBound(WorldObjects())
            If WorldObjects(a).InUse = True Then
                D3DXMatrixIdentity matWorldPuck
                
                D3DXMatrixTranslation matTrans, WorldObjects(a).WorldX, WorldObjects(a).WorldY, WorldObjects(a).WorldZ
                
                D3DXMatrixMultiply matWorldPuck, matWorldPuck, matTrans
                
                g_D3DDevice.SetTransform D3DTS_WORLD, matWorldPuck
               
                '************TEXTURING******************
                
                If WorldObjects(a).CodeWord <> "Custom" Then
                    g_D3DDevice.SetTexture 1, MainTextureLib(WorldObjects(a).TextureIndex).Surface
                    g_D3DDevice.SetTexture 0, MainTextureLib(WorldObjects(a).LightMapIndex).Surface
               End If

                
                Select Case WorldObjects(a).CodeWord
                    Case "Wall"
                        If WorldObjects(a).AlphaChannel = False Then
                            If WorldObjects(a).Rotation = 1 Then
                               
                                'g_D3DDevice.SetStreamSource 0, WallVB2, sizeOfVertex
                                'g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
                                If WorldObjects(a).UnderSelection = False Then
                                    WallV2(0).Color = WorldObjects(a).Color1
                                    WallV2(1).Color = WorldObjects(a).Color2
                                    WallV2(2).Color = WorldObjects(a).Color3
                                    WallV2(3).Color = WorldObjects(a).Color4
                                Else
                                    'unlucky to those with radioactive coloured walls
                                    WallV2(0).Color = &HFF33FF11
                                    WallV2(1).Color = &HFF33FF11
                                    WallV2(2).Color = &HFF33FF11
                                    WallV2(3).Color = &HFF33FF11
                                End If
                                
                                If WorldObjects(a).SwapTC = True Then
                                'Original Data
'                                    WallV2(0).tu2 = 1
'                                    WallV2(0).tv2 = 1
'                                    WallV2(1).tu2 = 1
'                                    WallV2(1).tv2 = 0
'                                    WallV2(2).tu2 = 0
'                                    WallV2(2).tv2 = 1
'                                    WallV2(3).tu2 = 0
'                                    WallV2(3).tv2 = 0
                                    'Inverse X
                                    WallV2(0).tu2 = 0
                                    WallV2(1).tu2 = 0
                                    WallV2(2).tu2 = 1
                                    WallV2(3).tu2 = 1
                                    
                                    
                                End If
                                If WorldObjects(a).SwapTC2 = True Then
                                    WallV2(0).tu1 = 0
                                    WallV2(1).tu1 = 0
                                    WallV2(2).tu1 = 1
                                    WallV2(3).tu1 = 1
                                End If
                                g_D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, WallV2(0), sizeOfVertex
                                    'Reset Data
                                    WallV2(0).tu2 = 1
                                    WallV2(0).tv2 = 1
                                    WallV2(1).tu2 = 1
                                    WallV2(1).tv2 = 0
                                    WallV2(2).tu2 = 0
                                    WallV2(2).tv2 = 1
                                    WallV2(3).tu2 = 0
                                    WallV2(3).tv2 = 0
                                    
                                    WallV2(0).tu1 = 1
                                    WallV2(0).tv1 = 1
                                    WallV2(1).tu1 = 1
                                    WallV2(1).tv1 = 0
                                    WallV2(2).tu1 = 0
                                    WallV2(2).tv1 = 1
                                    WallV2(3).tu1 = 0
                                    WallV2(3).tv1 = 0
                            Else
                                'g_D3DDevice.SetStreamSource 0, WallVB, sizeOfVertex
                                'g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
                                If WorldObjects(a).UnderSelection = False Then
                                    WallV(0).Color = WorldObjects(a).Color1
                                    WallV(1).Color = WorldObjects(a).Color2
                                    WallV(2).Color = WorldObjects(a).Color3
                                    WallV(3).Color = WorldObjects(a).Color4
                                Else
                                    'unlucky to those with limegreen walls
                                    WallV(0).Color = &HFF33FF11
                                    WallV(1).Color = &HFF33FF11
                                    WallV(2).Color = &HFF33FF11
                                    WallV(3).Color = &HFF33FF11
                                End If
                                
                                If WorldObjects(a).SwapTC = True Then
                                'Original Data
'
                                    'Inverse X
                                    WallV(0).tu2 = 0
                                    
                                    WallV(1).tu2 = 0
                                    
                                    WallV(2).tu2 = 1
                                    
                                    WallV(3).tu2 = 1
                                    
                                End If
                                g_D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, WallV(0), sizeOfVertex
                                'Reset Data
                                    WallV(0).tu2 = 1
                                    WallV(0).tv2 = 1
                                    WallV(1).tu2 = 1
                                    WallV(1).tv2 = 0
                                    WallV(2).tu2 = 0
                                    WallV(2).tv2 = 1
                                    WallV(3).tu2 = 0
                                    WallV(3).tv2 = 0
                            End If
                        Else
                            SpritesIn = True
                        End If
                    Case "Floor"
                        FloorV(1).Position.Y = FloorV(1).Position.Y + WorldObjects(a).LSc2R
                        FloorV(0).Position.Y = FloorV(0).Position.Y + WorldObjects(a).LSc1R
                        FloorV(2).Position.Y = FloorV(2).Position.Y + WorldObjects(a).LSc3R
                        FloorV(3).Position.Y = FloorV(3).Position.Y + WorldObjects(a).LSc4R
                        
                        If WorldObjects(a).UnderSelection = False Then
                            FloorV(0).Color = WorldObjects(a).Color1
                            FloorV(1).Color = WorldObjects(a).Color2
                            FloorV(2).Color = WorldObjects(a).Color3
                            FloorV(3).Color = WorldObjects(a).Color4
                        Else
                            'unlucky to those with radioactive coloured walls
                            FloorV(0).Color = &HFF33FF11
                            FloorV(1).Color = &HFF33FF11
                            FloorV(2).Color = &HFF33FF11
                            FloorV(3).Color = &HFF33FF11
                        End If
                        
                        g_D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, FloorV(0), sizeOfVertex
                        
                        
                        FloorV(0).Position.Y = FloorV(0).Position.Y - WorldObjects(a).LSc1R
                        
                        FloorV(1).Position.Y = FloorV(1).Position.Y - WorldObjects(a).LSc2R
                        
                        FloorV(2).Position.Y = FloorV(2).Position.Y - WorldObjects(a).LSc3R
                        
                        FloorV(3).Position.Y = FloorV(3).Position.Y - WorldObjects(a).LSc4R
                        
                    Case "Ceiling"
                        If WorldObjects(a).UnderSelection = False Then
                            CeilingV(0).Color = WorldObjects(a).Color1
                            CeilingV(1).Color = WorldObjects(a).Color2
                            CeilingV(2).Color = WorldObjects(a).Color3
                            CeilingV(3).Color = WorldObjects(a).Color4
                        Else
                           'unlucky to those with radioactive coloured walls
                            CeilingV(0).Color = &HFF33FF11
                            CeilingV(1).Color = &HFF33FF11
                            CeilingV(2).Color = &HFF33FF11
                            CeilingV(3).Color = &HFF33FF11
                        End If
                        g_D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CeilingV(0), sizeOfVertex
                        
                    Case "Sprite"
                        SpritesIn = True
                    Case "Custom"
                        FirstHandle = WorldObjects(a).CodeIndex
                        
                        g_D3DDevice.SetTexture 1, MainTextureLib(CObjects(FirstHandle).TextureIndex).Surface
                        g_D3DDevice.SetTexture 0, MainTextureLib(CObjects(FirstHandle).LightMapIndex).Surface
                        
                        g_D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CObjects(FirstHandle).Vertices(0), sizeOfVertex
                        'g_D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CeilingV(0), sizeOfVertex
                        Do
                        NextHandle = CObjects(FirstHandle).ChainNext
                        If NextHandle <> -1 Then
                            
                            g_D3DDevice.SetTexture 1, MainTextureLib(CObjects(FirstHandle).TextureIndex).Surface
                            g_D3DDevice.SetTexture 0, MainTextureLib(CObjects(FirstHandle).LightMapIndex).Surface
                        
                            g_D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, CObjects(NextHandle).Vertices(0), sizeOfVertex
                            FirstHandle = NextHandle
                        Else
                            Exit Do
                        End If
                        Loop
                        
                End Select
            End If
        Next a
        g_D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
        '---------------------------------------
        'Finally draw all transparent objects. This MUST be the last to be drawn!
        If SpritesIn Then
            g_D3DDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_DIFFUSE
            g_D3DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
            g_D3DDevice.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
            g_D3DDevice.SetRenderState D3DRS_ALPHAREF, 1
            'g_D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 0

            For a = 1 To UBound(WorldObjects())
                If WorldObjects(a).InUse = True Then
                    D3DXMatrixTranslation matWorldPuck, WorldObjects(a).WorldX, WorldObjects(a).WorldY, WorldObjects(a).WorldZ
                    D3DXMatrixRotationY matWorldPuck1, DefaultVP.RotationY
                    D3DXMatrixMultiply matWorldPuck1, matWorldPuck1, matWorldPuck
                    g_D3DDevice.SetTransform D3DTS_WORLD, matWorldPuck



                    g_D3DDevice.SetTexture 1, MainTextureLib(WorldObjects(a).TextureIndex).Surface
                    'g_D3DDevice.SetTexture 0, emptyText

                    Select Case WorldObjects(a).CodeWord
                        Case "Sprite"
                            g_D3DDevice.SetTransform D3DTS_WORLD, matWorldPuck1
                            g_D3DDevice.SetStreamSource 0, SpriteTile, sizeOfVertex
                            g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
                            g_D3DDevice.SetTransform D3DTS_WORLD, matWorldPuck
                        Case "Wall"
                            If WorldObjects(a).AlphaChannel = True Then
                               If WorldObjects(a).Rotation = 1 Then
        
                                    g_D3DDevice.SetStreamSource 0, WallVB2, sizeOfVertex
                                    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
                                Else
                                    g_D3DDevice.SetStreamSource 0, WallVB, sizeOfVertex
                                    g_D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
                                End If
                            End If
                    End Select
                End If
            Next a

            'End the temporary world (device) domination!
            g_D3DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 0
            'g_D3DDevice.SetRenderState D3DRS_ZWRITEENABLE, 1
            g_D3DDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT 'D3DTA_DIFFUSE
        End If


    g_D3DDevice.EndScene


    ' Present the backbuffer contents to the front buffer (screen)
    g_D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

    
    
    
End Sub

Function vec3(X As Single, Y As Single, z As Single) As D3DVECTOR
    vec3.X = X
    vec3.Y = Y
    vec3.z = z
End Function

Public Function InitGeometry() As Boolean

    Build_Guides 20, 20
    
    Build_BasicObjects
    
    Build_PositionPuck

    InitGeometry = True
End Function


Sub SetupMatrices()
    
    D3DXMatrixIdentity matWorld
    g_D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    Dim matView As D3DMATRIX
    D3DXMatrixLookAtLH matView, vec3(DefaultVP.Position.X, DefaultVP.Position.Y, DefaultVP.Position.z), vec3((Sin(DefaultVP.RotationY)) + DefaultVP.Position.X, Sin(DefaultVP.RotationX) + DefaultVP.Position.Y, Cos(DefaultVP.RotationY) + DefaultVP.Position.z), vec3(0, 1, 0)
    g_D3DDevice.SetTransform D3DTS_VIEW, matView

End Sub

Public Sub Cleanup()
    Set g_Texture = Nothing
    Set g_VB = Nothing
    Set g_D3DDevice = Nothing
    Set g_D3D = Nothing
End Sub

Private Sub SetTextureStages()

    g_D3DDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    g_D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
    g_D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
    g_D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    g_D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
  
    g_D3DDevice.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
    g_D3DDevice.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
    g_D3DDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT 'D3DTA_DIFFUSE

    g_D3DDevice.SetTextureStageState 1, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
    g_D3DDevice.SetTextureStageState 1, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
    
    g_D3DDevice.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
    g_D3DDevice.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_LINEAR
  
    g_D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    g_D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    g_D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
End Sub

