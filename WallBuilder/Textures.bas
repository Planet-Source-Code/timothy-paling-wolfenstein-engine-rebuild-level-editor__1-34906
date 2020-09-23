Attribute VB_Name = "Textures"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Cheaty way of getting mutiple type arrays
Public Type TextureLib
    FileName As String
    Surface As Direct3DTexture8
End Type

Public MainTextureLib() As TextureLib




Public Function SearchTextures(SearchString As String) As Long
    For a = 1 To UBound(MainTextureLib())
        If MainTextureLib(a).FileName = SearchString Then
            SearchTextures = a
            Exit Function
        End If
    Next a
    SearchTextures = -1
End Function

Public Function AddTexture(FileName As String) As Long
    ReDim Preserve MainTextureLib(UBound(MainTextureLib()) + 1)
    
    AddTexture = UBound(MainTextureLib())
    
    With MainTextureLib(AddTexture)
        'Set .Surface = g_D3DX.CreateTextureFromFileEx(g_D3DDevice, FileName, 64, 64, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, ByVal 0, ByVal 0)
        Set .Surface = g_D3DX.CreateTextureFromFileEx(g_D3DDevice, FileName, 256, 256, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        'Set .Surface = g_D3DX.CreateTextureFromFile(g_D3DDevice, FileName)
        If .Surface Is Nothing Then
            MsgBox "Error loading texture " & FileName
            Exit Function
        End If
        .FileName = FileName
    End With
End Function

Public Function OpenRetFile() As String

    On Error GoTo Err2
    With Form1.CommonDialog1
        .CancelError = True
        .DefaultExt = "dds"
        .DialogTitle = "Open World"
        .FileName = "untitled"
        .Filter = "DirectX AlphaChannel Texture (*.dds)|*.dds"
        .ShowOpen
        OpenRetFile = .FileName
    End With
    Exit Function
Err2:
    
    OpenRetFile = "-1"
    
    
End Function

Public Function OpenRetFileEx(Filter As String, DefaultExt As String, dlgtitle As String) As String

    On Error GoTo Err2
    With Form1.CommonDialog1
        .CancelError = True
        .DefaultExt = DefaultExt
        .DialogTitle = dlgtitle
        .FileName = "untitled"
        .Filter = Filter
        .ShowOpen
        OpenRetFileEx = .FileName
    End With
    Exit Function
Err2:
    
    OpenRetFileEx = "-1"
    
    
End Function
