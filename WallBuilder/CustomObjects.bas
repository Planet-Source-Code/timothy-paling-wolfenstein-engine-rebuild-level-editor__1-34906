Attribute VB_Name = "CustomObjects"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'------------------------------------------------------------------------------------------------------------------------
'All custom objects are based on quads. However, L3D's implementation of object chains means quads complex objects can be
'made.
'------------------------------------------------------------------------------------------------------------------------
Public Type CustomObject
    Vertices(4) As CUSTOMVERTEX
    DataFile As String
    
    TextureFile As String
    TextureIndex As Long
    
    LightMapFile As String
    LightMapIndex As Long
    
    AlphaChannel As Boolean
    
    ChainNext As Long
    
End Type

Public CObjects() As CustomObject



Public Function Load_CObjectChain(FileName As String) As Long
    
    
    
    ReDim Preserve CObjects(UBound(CObjects()) + 1)
    Load_CObjectChain = UBound(CObjects())
    FileNum = FreeFile()
    Open FileName For Input As FileNum
        For a = 0 To 3
            Line Input #1, X
            Line Input #1, Y
            Line Input #1, z
            Line Input #1, c
            Line Input #1, tu1
            Line Input #1, tv1
            Line Input #1, tu2
            Line Input #1, tv2
            
            With CObjects(Load_CObjectChain).Vertices(a)
                .Position.X = CSng(X)
                .Position.Y = CSng(Y)
                .Position.z = CSng(z)
                
                .Color = CLng(c)
                
                .tu1 = CSng(tu1)
                .tv1 = CSng(tv1)
                .tu2 = CSng(tu2)
                .tv2 = CSng(tv2)
            End With
            
        Next a
        Line Input #1, TFile
        Line Input #1, LFile
        Line Input #1, AChan
        Line Input #1, FileChainNext
        Close FileNum
        With CObjects(Load_CObjectChain)
            .TextureFile = CStr(TFile)
            .LightMapFile = CStr(LFile)
            .AlphaChannel = CBool(AChan)
        End With
        If CStr(FileChainNext) <> "End" Then
            CObjects(Load_CObjectChain).ChainNext = Load_CObjectChain(CStr(FileChainNext))
        Else
            CObjects(Load_CObjectChain).ChainNext = -1
        End If
        
    
End Function
