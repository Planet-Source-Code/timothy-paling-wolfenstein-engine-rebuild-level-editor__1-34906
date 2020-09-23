Attribute VB_Name = "BasicShapes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public WallVB As Direct3DVertexBuffer8
Public WallVB2 As Direct3DVertexBuffer8
Public FloorTile As Direct3DVertexBuffer8
Public CeilingTile As Direct3DVertexBuffer8
Public SpriteTile As Direct3DVertexBuffer8
Public WallV() As CUSTOMVERTEX
Public WallV2() As CUSTOMVERTEX
Public CeilingV() As CUSTOMVERTEX
Public FloorV() As CUSTOMVERTEX

Public Sub Build_BasicObjects(Optional ScaleSize As Integer = 1, Optional TextureRepeat As Integer = 1)
    
     Dim Vertices() As CUSTOMVERTEX
     ReDim Vertices(4)
     ReDim WallV(4)
     
     WallV(0).Position.X = 0
     WallV(0).Position.Y = 0
     WallV(0).Position.z = 0
     'WallV(0).normal.z = 1
     WallV(0).tu1 = 1
     WallV(0).tv1 = 1
     WallV(0).tu2 = 1
     WallV(0).tv2 = 1
     WallV(0).Color = &HFFFFFFFF
     
     
     WallV(1).Position.X = 0
     WallV(1).Position.Y = 1
     WallV(1).Position.z = 0
     'WallV(1).normal.z = 1
     WallV(1).tu1 = 1
     WallV(1).tv1 = 0
     WallV(1).tu2 = 1
     WallV(1).tv2 = 0
     WallV(1).Color = &HFFFFFFFF
     
     WallV(2).Position.X = 1
     WallV(2).Position.Y = 0
     WallV(2).Position.z = 0
     'WallV(2).normal.z = 1
     WallV(2).tu1 = 0
     WallV(2).tv1 = 1
     WallV(2).tu2 = 0
     WallV(2).tv2 = 1
     
     WallV(2).Color = &HFFFFFFFF
     
     WallV(3).Position.X = 1
     WallV(3).Position.Y = 1
     WallV(3).Position.z = 0
     'WallV(3).normal.z = 1
     WallV(3).tu1 = 0
     WallV(3).tv1 = 0
     WallV(3).tu2 = 0
     WallV(3).tv2 = 0
     WallV(3).Color = &HFFFFFFFF
     
     VertexSizeInBytes = Len(Vertices(0))
     
    ' Create the vertex buffer.
    Set WallVB = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If WallVB Is Nothing Then Exit Sub

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData WallVB, 0, VertexSizeInBytes * 4, 0, WallV(0)
    
    
    '************************
        ReDim WallV2(4)
     WallV2(0).Position.X = 0
     WallV2(0).Position.Y = 0
     WallV2(0).Position.z = 0
     WallV2(0).normal.X = 1
     WallV2(0).tu1 = 1
     WallV2(0).tv1 = 1
     WallV2(0).tu2 = 1
     WallV2(0).tv2 = 1
     WallV2(0).Color = &HFFFFFFFF
     
     WallV2(1).Position.X = 0
     WallV2(1).Position.Y = 1
     WallV2(1).Position.z = 0
     WallV2(1).normal.X = 1
     WallV2(1).tu1 = 1
     WallV2(1).tv1 = 0
     WallV2(1).tu2 = 1
     WallV2(1).tv2 = 0
     WallV2(1).Color = &HFFFFFFFF
     
     WallV2(2).Position.X = 0
     WallV2(2).Position.Y = 0
     WallV2(2).Position.z = 1
     WallV2(2).normal.X = 1
     WallV2(2).tu1 = 0
     WallV2(2).tv1 = 1
     WallV2(2).tu2 = 0
     WallV2(2).tv2 = 1
     WallV2(2).Color = &HFFFFFFFF
     
     WallV2(3).Position.X = 0
     WallV2(3).Position.Y = 1
     WallV2(3).Position.z = 1
     WallV2(3).normal.X = 1
     WallV2(3).tu1 = 0
     WallV2(3).tv1 = 0
     WallV2(3).tu2 = 0
     WallV2(3).tv2 = 0
     WallV2(3).Color = &HFFFFFFFF
     
     VertexSizeInBytes = Len(Vertices(0))
     
    ' Create the vertex buffer.
    Set WallVB2 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If WallVB2 Is Nothing Then Exit Sub

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData WallVB2, 0, VertexSizeInBytes * 4, 0, WallV2(0)
    
        '************************
    ReDim FloorV(4)
     FloorV(0).Position.X = 0
     FloorV(0).Position.Y = 0
     FloorV(0).Position.z = 0
     FloorV(0).normal.Y = 1
     FloorV(0).tu1 = 1
     FloorV(0).tv1 = 1
     FloorV(0).tu2 = 1
     FloorV(0).tv2 = 1
     FloorV(0).Color = &HFFFFFFFF
     
     FloorV(1).Position.X = 0
     FloorV(1).Position.Y = 0
     FloorV(1).Position.z = 1
     FloorV(1).normal.Y = 1
     FloorV(1).tu1 = 1
     FloorV(1).tv1 = 0
     FloorV(1).tu2 = 1
     FloorV(1).tv2 = 0
     FloorV(1).Color = &HFFFFFFFF
     
     FloorV(2).Position.X = 1
     FloorV(2).Position.Y = 0
     FloorV(2).Position.z = 0
     FloorV(2).normal.Y = 1
     FloorV(2).tu1 = 0
     FloorV(2).tv1 = 1
     FloorV(2).tu2 = 0
     FloorV(2).tv2 = 1
     FloorV(2).Color = &HFFFFFFFF
     
     FloorV(3).Position.X = 1
     FloorV(3).Position.Y = 0
     FloorV(3).Position.z = 1
     FloorV(3).normal.Y = 1
     FloorV(3).tu1 = 0
     FloorV(3).tv1 = 0
     FloorV(3).tu2 = 0
     FloorV(3).tv2 = 0
     FloorV(3).Color = &HFFFFFFFF
     
     VertexSizeInBytes = Len(Vertices(0))
     
    ' Create the vertex buffer.
    Set FloorTile = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If FloorTile Is Nothing Then Exit Sub

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData FloorTile, 0, VertexSizeInBytes * 4, 0, FloorV(0)
    
    '************************
    ReDim CeilingV(4)
     CeilingV(0).Position.X = 0
     CeilingV(0).Position.Y = 1
     CeilingV(0).Position.z = 0
     CeilingV(0).normal.Y = 1
     CeilingV(0).tu1 = 1
     CeilingV(0).tv1 = 1
     CeilingV(0).tu2 = 1
     CeilingV(0).tv2 = 1
     CeilingV(0).Color = &HFFFFFFFF
     
     CeilingV(1).Position.X = 0
     CeilingV(1).Position.Y = 1
     CeilingV(1).Position.z = 1
     CeilingV(1).normal.Y = 1
     CeilingV(1).tu1 = 1
     CeilingV(1).tv1 = 0
     CeilingV(1).tu2 = 1
     CeilingV(1).tv2 = 0
     CeilingV(1).Color = &HFFFFFFFF
     
     CeilingV(2).Position.X = 1
     CeilingV(2).Position.Y = 1
     CeilingV(2).Position.z = 0
     CeilingV(2).normal.Y = 1
     CeilingV(2).tu1 = 0
     CeilingV(2).tv1 = 1
     CeilingV(2).tu2 = 0
     CeilingV(2).tv2 = 1
     CeilingV(2).Color = &HFFFFFFFF
     
     CeilingV(3).Position.X = 1
     CeilingV(3).Position.Y = 1
     CeilingV(3).Position.z = 1
     CeilingV(3).normal.Y = 1
     CeilingV(3).tu1 = 0
     CeilingV(3).tv1 = 0
     CeilingV(3).tu2 = 0
     CeilingV(3).tv2 = 0
     CeilingV(3).Color = &HFFFFFFFF
     
     VertexSizeInBytes = Len(Vertices(0))
     
    ' Create the vertex buffer.
    Set CeilingTile = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If CeilingTile Is Nothing Then Exit Sub

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData CeilingTile, 0, VertexSizeInBytes * 4, 0, CeilingV(0)
    
    '************************************
    
     Vertices(0).Position.X = -0.5
     Vertices(0).Position.Y = 0
     Vertices(0).Position.z = 0
     Vertices(0).normal.z = 1
     Vertices(0).tu1 = 1
     Vertices(0).tv1 = 1
     Vertices(0).tu2 = 1
     Vertices(0).tv2 = 1
     Vertices(0).Color = &HFFFFFFFF
     
     Vertices(1).Position.X = -0.5
     Vertices(1).Position.Y = 1
     Vertices(1).Position.z = 0
     Vertices(1).normal.z = 1
     Vertices(1).tu1 = 1
     Vertices(1).tv1 = 0
     Vertices(1).tu2 = 1
     Vertices(1).tv2 = 0
     
     Vertices(1).Color = &HFFFFFFFF
     
     Vertices(2).Position.X = 0.5
     Vertices(2).Position.Y = 0
     Vertices(2).Position.z = 0
     Vertices(2).normal.z = 1
     Vertices(2).tu1 = 0
     Vertices(2).tv1 = 1
     Vertices(2).tu2 = 0
     Vertices(2).tv2 = 1
     Vertices(2).Color = &HFFFFFFFF
     
     Vertices(3).Position.X = 0.5
     Vertices(3).Position.Y = 1
     Vertices(3).Position.z = 0
     Vertices(3).normal.z = 1
     Vertices(3).tu1 = 0
     Vertices(3).tv1 = 0
     Vertices(3).tu2 = 0
     Vertices(3).tv2 = 0
     Vertices(3).Color = &HFFFFFFFF
     
     VertexSizeInBytes = Len(Vertices(0))
     
    ' Create the vertex buffer.
    Set SpriteTile = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If SpriteTile Is Nothing Then Exit Sub

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData SpriteTile, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    

End Sub
