Attribute VB_Name = "Guides"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public GuideVB As Direct3DVertexBuffer8
Public SizeConstant As Long

Public PuckVB As Direct3DVertexBuffer8
Public PuckVB2 As Direct3DVertexBuffer8
Public PuckVB3 As Direct3DVertexBuffer8
Public PuckVB4 As Direct3DVertexBuffer8
Public PuckVB5 As Direct3DVertexBuffer8

'Public LightPuck As Direct3DVertexBuffer8


Public Sub Build_Guides(XLength As Long, YLength As Long)
    SizeConstant = ((XLength + 1) * 2) + ((YLength + 1) * 2)
     Dim Vertices() As CUSTOMVERTEX
     ReDim Vertices(SizeConstant)
     
     VertexSizeInBytes = Len(Vertices(0))
     c = -1
    For a = 0 To XLength
        c = c + 1
        Vertices(c).Position.X = a
        Vertices(c).Color = &HFFC0C0C0
        c = c + 1
        Vertices(c).Position.X = a
        Vertices(c).Position.z = YLength
        Vertices(c).Color = &HFFC0C0C0
    Next a
    
     'c = -1
    For a = 0 To YLength
        c = c + 1
        Vertices(c).Position.z = a
        Vertices(c).Color = &HFFC0C0C0
        c = c + 1
        Vertices(c).Position.z = a
        Vertices(c).Position.X = XLength
        Vertices(c).Color = &HFFC0C0C0
    Next a
    
    
    

    ' Create the vertex buffer.
    Set GuideVB = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * SizeConstant, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If GuideVB Is Nothing Then Exit Sub

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData GuideVB, 0, VertexSizeInBytes * SizeConstant, 0, Vertices(0)
End Sub

Public Sub Build_PositionPuck(Optional ScaleSize As Integer = 1)
    Dim Vertices(4) As CUSTOMVERTEX
    VertexSizeInBytes = Len(Vertices(0))
    
    Vertices(0).Position.X = 0
    Vertices(0).Position.Y = 0
    Vertices(0).Position.z = 0
    Vertices(0).Color = &HFFFCFA4E '&HFFFF0000

    Vertices(1).Position.X = 0
    Vertices(1).Position.Y = 1
    Vertices(1).Position.z = 0
    Vertices(1).Color = &HFF4F5FFD '&HFF00FF00

    Vertices(2).Position.X = 1
    Vertices(2).Position.Y = 0
    Vertices(2).Position.z = 0
    Vertices(2).Color = &HFFFD4FF3 '&HFF00FF00
    
    Vertices(3).Position.X = 1
    Vertices(3).Position.Y = 1
    Vertices(3).Position.z = 0
    Vertices(3).Color = &HFF53FD4F '&HFFFF0000

    ' Create the vertex buffer.
    Set PuckVB = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If PuckVB Is Nothing Then
        MsgBox "Error"
        Exit Sub
    End If

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData PuckVB, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    '*************************************
    
    Vertices(0).Position.X = 0
    Vertices(0).Position.Y = 0
    Vertices(0).Position.z = 0
    Vertices(0).Color = &HFFFCFA4E '&HFFFF0000

    Vertices(1).Position.X = 0
    Vertices(1).Position.Y = 1
    Vertices(1).Position.z = 0
    Vertices(1).Color = &HFF4F5FFD '&HFF00FF00

    Vertices(2).Position.X = 0
    Vertices(2).Position.Y = 0
    Vertices(2).Position.z = 1
    Vertices(2).Color = &HFFFD4FF3 '&HFF00FF00

    Vertices(3).Position.X = 0
    Vertices(3).Position.Y = 1
    Vertices(3).Position.z = 1
    Vertices(3).Color = &HFF53FD4F '&HFFFF0000

    ' Create the vertex buffer.
    Set PuckVB2 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If PuckVB2 Is Nothing Then
        MsgBox "Error"
        Exit Sub
    End If

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData PuckVB2, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    '*******************************************
    
    Vertices(0).Position.X = 0
    Vertices(0).Position.Y = 0
    Vertices(0).Position.z = 0
    Vertices(0).Color = &HFFFCFA4E '&HFFFF0000

    Vertices(1).Position.X = 0
    Vertices(1).Position.Y = 0
    Vertices(1).Position.z = 1
    Vertices(1).Color = &HFF4F5FFD '&HFF00FF00

    Vertices(2).Position.X = 1
    Vertices(2).Position.Y = 0
    Vertices(2).Position.z = 0
    Vertices(2).Color = &HFFFD4FF3 '&HFF00FF00

    Vertices(3).Position.X = 1
    Vertices(3).Position.Y = 0
    Vertices(3).Position.z = 1
    Vertices(3).Color = &HFF53FD4F '&HFFFF0000

    ' Create the vertex buffer.
    Set PuckVB3 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If PuckVB3 Is Nothing Then
        MsgBox "Error"
        Exit Sub
    End If

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData PuckVB3, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    '***********************************************
    
    Vertices(0).Position.X = 0
    Vertices(0).Position.Y = 1
    Vertices(0).Position.z = 0
    Vertices(0).Color = &HFFFCFA4E '&HFFFF0000

    Vertices(1).Position.X = 0
    Vertices(1).Position.Y = 1
    Vertices(1).Position.z = 1
    Vertices(1).Color = &HFF4F5FFD '&HFF00FF00

    Vertices(2).Position.X = 1
    Vertices(2).Position.Y = 1
    Vertices(2).Position.z = 0
    Vertices(2).Color = &HFF00FF00

    Vertices(3).Position.X = 1
    Vertices(3).Position.Y = 1
    Vertices(3).Position.z = 1
    Vertices(3).Color = &HFFFF0000

    ' Create the vertex buffer.
    Set PuckVB4 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If PuckVB4 Is Nothing Then
        MsgBox "Error"
        Exit Sub
    End If

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData PuckVB4, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    '*************************************
    
     Vertices(0).Position.X = -0.5
     Vertices(0).Position.Y = 0
     Vertices(0).Position.z = 0
     'Vertices(0).tu = 1
     'Vertices(0).tv = 1
     Vertices(0).Color = &HFFFF0000
     
     Vertices(1).Position.X = -0.5
     Vertices(1).Position.Y = 1
     Vertices(1).Position.z = 0
     'Vertices(1).tu = 1
     'Vertices(1).tv = 0
     Vertices(1).Color = &HFF00FF00
     
     Vertices(2).Position.X = 0.5
     Vertices(2).Position.Y = 0
     Vertices(2).Position.z = 0
     'Vertices(2).tu = 0
     'Vertices(2).tv = 1
     Vertices(2).Color = &HFF00FF00
     
     Vertices(3).Position.X = 0.5
     Vertices(3).Position.Y = 1
     Vertices(3).Position.z = 0
     'Vertices(3).tu = 0
     'Vertices(3).tv = 0
     Vertices(3).Color = &HFFFF0000
     
     VertexSizeInBytes = Len(Vertices(0))
     
    ' Create the vertex buffer.
    Set PuckVB5 = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If PuckVB5 Is Nothing Then Exit Sub

    ' fill the vertex buffer from our array
    D3DVertexBuffer8SetData PuckVB5, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
    '********************************
    'Lighting Puck
'         Vertices(0).Position.X = -0.05
'     Vertices(0).Position.Y = 0
'     Vertices(0).Position.z = 0
'     'Vertices(0).tu = 1
'     'Vertices(0).tv = 1
'     Vertices(0).Color = &HFFFF0000
'
'     Vertices(1).Position.X = -0.05
'     Vertices(1).Position.Y = 0.1
'     Vertices(1).Position.z = 0
'     'Vertices(1).tu = 1
'     'Vertices(1).tv = 0
'     Vertices(1).Color = &HFF00FF00
'
'     Vertices(2).Position.X = 0.05
'     Vertices(2).Position.Y = 0
'     Vertices(2).Position.z = 0
'     'Vertices(2).tu = 0
'     'Vertices(2).tv = 1
'     Vertices(2).Color = &HFF00FF00
'
'     Vertices(3).Position.X = 0.05
'     Vertices(3).Position.Y = 0.1
'     Vertices(3).Position.z = 0
'     'Vertices(3).tu = 0
'     'Vertices(3).tv = 0
'     Vertices(3).Color = &HFFFF0000
'
'     VertexSizeInBytes = Len(Vertices(0))
'
'    ' Create the vertex buffer.
'    Set LightPuck = g_D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
'    If LightPuck Is Nothing Then Exit Sub
'
'    ' fill the vertex buffer from our array
'    D3DVertexBuffer8SetData LightPuck, 0, VertexSizeInBytes * 4, 0, Vertices(0)
    
End Sub
