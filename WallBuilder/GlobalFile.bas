Attribute VB_Name = "GlobalFile"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveWorld()
On Error GoTo Err1
    With Form1.CommonDialog1
        .CancelError = True
        .DefaultExt = "wl3"
        .DialogTitle = "Save World"
        .FileName = "untitled"
        .Filter = "L3DWolfenstein Files (*.wl3)|*.wl3"
        .ShowSave
        
        Open .FileName For Output As #1
            Line = "L3DWolfWED" & "." & App.Major & "." & App.Minor & "." & App.Revision
            Print #1, Line
            
            Line = UBound(WorldObjects())
            Print #1, Trim(Line)
            
            For a = 1 To UBound(WorldObjects())
                Line = WorldObjects(a).CodeWord
                Print #1, Trim(Line)
                Line = WorldObjects(a).DrawMode
                Print #1, Trim(Line)
                Line = WorldObjects(a).Rotation
                Print #1, Trim(Line)
                Line = WorldObjects(a).TextureFile
                Print #1, Trim(Line)
                Line = WorldObjects(a).LightMapFile
                Print #1, Trim(Line)
                Line = WorldObjects(a).WorldX
                Print #1, Trim(Line)
                Line = WorldObjects(a).WorldY
                Print #1, Trim(Line)
                Line = WorldObjects(a).WorldZ
                Print #1, Trim(Line)
                Line = WorldObjects(a).LSc1R
                Print #1, Trim(Line)
                Line = WorldObjects(a).LSc2R
                Print #1, Trim(Line)
                Line = WorldObjects(a).LSc3R
                Print #1, Trim(Line)
                Line = WorldObjects(a).LSc4R
                Print #1, Trim(Line)
                Line = WorldObjects(a).Color1
                Print #1, Trim(Line)
                Line = WorldObjects(a).Color2
                Print #1, Trim(Line)
                Line = WorldObjects(a).Color3
                Print #1, Trim(Line)
                Line = WorldObjects(a).Color4
                Print #1, Trim(Line)
                If WorldObjects(a).SwapTC = True Then
                    Print #1, Trim("1")
                Else
                    Print #1, Trim("0")
                End If
                If WorldObjects(a).SwapTC2 = True Then
                    Print #1, Trim("1")
                Else
                    Print #1, Trim("0")
                End If
                If WorldObjects(a).AlphaChannel = True Then
                    Print #1, Trim("1")
                Else
                    Print #1, Trim("0")
                End If
                'Print #1, "//////////"
                
            Next a
            
            
        Close #1
        
    End With
Err1:

End Sub

Public Sub LoadWorld()
Dim WorldX As Single
Dim WorldY As Single
Dim WorldZ As Single
Dim sWorldX As String
Dim sWorldY As String
Dim sWorldZ As String
Dim CodeWord As String
Dim TextureFile As String
Dim LightMapFile As String
Dim DrawMode As CONST_D3DFILLMODE
Dim sDrawMode As String
Dim sRotation As String
Dim Rotation As Byte
Dim sAlpha As String
Dim sSwapTC As String
Dim sSwapTC2 As String
Dim sColor1 As String
Dim sColor2 As String
Dim sColor3 As String
Dim sColor4 As String
Dim sLSc1R As String
Dim sLSc2R As String
Dim sLSc3R As String
Dim sLSc4R As String

On Error GoTo Err2
    With Form1.CommonDialog1
        .CancelError = True
        .DefaultExt = "wl3"
        .DialogTitle = "Open World"
        .FileName = "untitled"
        .Filter = "L3DWolfenstein Files (*.wl3)|*.wl3"
        .ShowOpen
        
        Open .FileName For Input As #1
            Line Input #1, VerInfo
            Line Input #1, ObjectCount
            
            For a = 1 To ObjectCount
                Line Input #1, CodeWord
                Line Input #1, sDrawMode
                Line Input #1, sRotation
                Line Input #1, TextureFile
                Line Input #1, LightMapFile
                Line Input #1, sWorldX
                Line Input #1, sWorldY
                Line Input #1, sWorldZ
                Line Input #1, sLSc1R
                Line Input #1, sLSc2R
                Line Input #1, sLSc3R
                Line Input #1, sLSc4R
                Line Input #1, sColor1
                Line Input #1, sColor2
                Line Input #1, sColor3
                Line Input #1, sColor4
                Line Input #1, sSwapTC
                Line Input #1, sSwapTC2
                Line Input #1, sAlpha
                
                WorldX = sWorldX
                WorldY = sWorldY
                WorldZ = sWorldZ
                DrawMode = sDrawMode
                Rotation = sRotation
                
                Form1.Check2.Value = Int(sAlpha)
                Form1.Check4.Value = Int(sSwapTC)
                Form1.Check5.Value = Int(sSwapTC2)
                Form1.V1Y = CSng(sLSc1R)
                Form1.V2Y = CSng(sLSc2R)
                Form1.V3Y = CSng(sLSc3R)
                Form1.V4Y = CSng(sLSc4R)
                
                Add_Object CodeWord, TextureFile, LightMapFile, DrawMode, vec3(WorldX, WorldY, WorldZ), Rotation
                
            Next a
        Close #1
        
    End With
Err2:
End Sub
