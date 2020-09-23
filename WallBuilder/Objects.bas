Attribute VB_Name = "Objects"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type WorldObject
    
    CodeWord As String
    CodeIndex As Long
    
    TextureFile As String
    TextureIndex As Long
    
    LightMapFile As String
    LightMapIndex As Long
    
    WorldX As Single
    WorldY As Single
    WorldZ As Single
    
    LSc1R As Single
    LSc2R As Single
    LSc3R As Single
    LSc4R As Single
    
    Rotation As Byte
    
    AlphaChannel As Boolean
    
    Color1 As Long
    Color2 As Long
    Color3 As Long
    Color4 As Long
    
    SwapTC As Boolean
    SwapTC2 As Boolean
 
    DrawMode As CONST_D3DFILLMODE
    
    SqD As Single 'Not used?
    
    InUse As Boolean
    NodeIndex As Integer
    TempCheck As Boolean
    UnderSelection As Boolean
End Type
Public WorldObjects() As WorldObject

Public Sub Add_Object(CodeWord As String, TextureFile As String, LightMapFile As String, DrawMode As CONST_D3DFILLMODE, Position As D3DVECTOR, Rotation As Byte)
    For b = 1 To UBound(WorldObjects())
        If WorldObjects(b).InUse = False Then
            a = b
            b = -1

            Exit For
        End If
    Next b

    If b <> -1 Then
        ReDim Preserve WorldObjects(UBound(WorldObjects()) + 1)
        a = UBound(WorldObjects())

    End If
        With WorldObjects(a)
            If CodeWord <> "Custom" Then
                .CodeWord = CodeWord
                .DrawMode = DrawMode
                .Rotation = Rotation
                .TextureFile = TextureFile
                .LightMapFile = LightMapFile
                .WorldX = Position.X
                .WorldY = Position.Y
                .WorldZ = Position.z
                .Color1 = RGB(Form1.v1red.Text, Form1.v1blue.Text, Form1.v1green.Text)
                .Color2 = RGB(Form1.v2red.Text, Form1.v2blue.Text, Form1.v2green.Text)
                .Color3 = RGB(Form1.v3red.Text, Form1.v3blue.Text, Form1.v3green.Text)
                .Color4 = RGB(Form1.v4red.Text, Form1.v4blue.Text, Form1.v4green.Text)
                .LSc1R = Form1.V1Y
                .LSc2R = Form1.V2Y
                .LSc3R = Form1.V3Y
                .LSc4R = Form1.V4Y
                .UnderSelection = False
            Else
                .CodeWord = CodeWord
                .CodeIndex = Load_CObjectChain(TextureFile)
                .WorldX = Position.X
                .WorldY = Position.Y
                .WorldZ = Position.z
            End If
            
            .InUse = True
            
            If CodeWord <> "Light" Then
                If Form1.Check2.Value = 1 Then
                .AlphaChannel = True
                Else
                .AlphaChannel = False
                End If
            End If
            
            If Form1.Check4.Value = 1 Then
                .SwapTC = True
            Else
                .SwapTC = False
            End If
            
            If Form1.Check5.Value = 1 Then
                .SwapTC2 = True
            Else
                .SwapTC2 = False
            End If
        End With

        With Form1.TreeView1
            Dim TmpNode As Node

            Select Case WorldObjects(a).CodeWord
                Case "Wall"
                    Set TmpNode = .Nodes.Add("Walls", tvwChild, "Object" & a, WorldObjects(a).CodeWord)
                Case "Floor"
                    Set TmpNode = .Nodes.Add("Floor", tvwChild, "Object" & a, WorldObjects(a).CodeWord)
                Case "Ceiling"
                    Set TmpNode = .Nodes.Add("Ceiling", tvwChild, "Object" & a, WorldObjects(a).CodeWord)
                Case "Sprite"
                    Set TmpNode = .Nodes.Add("Objects", tvwChild, "Object" & a, WorldObjects(a).CodeWord)
                Case "Custom"
                    Set TmpNode = .Nodes.Add("Objects", tvwChild, "Object" & a, WorldObjects(a).CodeWord)
            End Select
            WorldObjects(a).NodeIndex = TmpNode.Index

            .Nodes.Add "Object" & a, tvwChild, "CObject" & a, "CodeWord=" & WorldObjects(a).CodeWord
            .Nodes.Add "Object" & a, tvwChild, "DObject" & a, "DrawMode=" & WorldObjects(a).DrawMode
            .Nodes.Add "Object" & a, tvwChild, "RObject" & a, "Rotation=" & WorldObjects(a).Rotation
            .Nodes.Add "Object" & a, tvwChild, "TObject" & a, "TextureFile=" & WorldObjects(a).TextureFile
            .Nodes.Add "Object" & a, tvwChild, "XObject" & a, "WorldX=" & WorldObjects(a).WorldX
            .Nodes.Add "Object" & a, tvwChild, "YObject" & a, "WorldY=" & WorldObjects(a).WorldY
            .Nodes.Add "Object" & a, tvwChild, "ZObject" & a, "WorldZ=" & WorldObjects(a).WorldZ
            .Nodes.Add "Object" & a, tvwChild, "IObject" & a, "Index=" & a
        End With
        If WorldObjects(a).CodeWord <> "Custom" Then
            Ret = SearchTextures(WorldObjects(a).TextureFile)
            If Ret = -1 Then
                Ret = AddTexture(WorldObjects(a).TextureFile)
            End If
            WorldObjects(a).TextureIndex = Ret
            
            
            Ret = SearchTextures(WorldObjects(a).LightMapFile)
            If Ret = -1 Then
                Ret = AddTexture(WorldObjects(a).LightMapFile)
            End If
            WorldObjects(a).LightMapIndex = Ret
        Else
            Index = WorldObjects(a).CodeIndex
            Ret = SearchTextures(CObjects(Index).TextureFile)
            If Ret = -1 Then
                Ret = AddTexture(CObjects(Index).TextureFile)
            End If
            CObjects(Index).TextureIndex = Ret
            
            Index = WorldObjects(a).CodeIndex
            Ret = SearchTextures(CObjects(Index).LightMapFile)
            If Ret = -1 Then
                Ret = AddTexture(CObjects(Index).LightMapFile)
            End If
            CObjects(Index).LightMapIndex = Ret
            
            
        End If
End Sub

Public Sub Delete_Object(Index As Integer)
    WorldObjects(Index).InUse = False
    Form1.TreeView1.Nodes.Remove WorldObjects(Index).NodeIndex
End Sub
