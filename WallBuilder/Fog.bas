Attribute VB_Name = "Fog"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub SetupVertexFog(Color As String, Mode As CONST_D3DFOGMODE, UseRange As Boolean, FogStart As Single, FogEnd As Single, Optional Density As Single)

    Dim StartFog As Single, EndFog As Single
    
    Call g_D3DDevice.SetRenderState(D3DRS_FOGENABLE, True)
 
    Call g_D3DDevice.SetRenderState(D3DRS_FOGCOLOR, Color)
    
    If Mode = D3DFOG_LINEAR Then
        Call g_D3DDevice.SetRenderState(D3DRS_FOGVERTEXMODE, Mode)
        Call g_D3DDevice.SetRenderState(D3DRS_FOGSTART, FtoDW(FogStart))
        Call g_D3DDevice.SetRenderState(D3DRS_FOGEND, FtoDW(FogEnd))
    Else
        Call g_D3DDevice.SetRenderState(D3DRS_FOGVERTEXMODE, Mode)
        Call g_D3DDevice.SetRenderState(D3DRS_FOGDENSITY, FtoDW(Density))
    End If


    If UseRange = True Then
        Call g_D3DDevice.SetRenderState(D3DRS_RANGEFOGENABLE, True)
    End If
End Sub

Public Function FtoDW(f As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = g_D3DX.CreateBuffer(4)
    g_D3DX.BufferSetData buf, 0, 4, 1, f
    g_D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l
End Function
