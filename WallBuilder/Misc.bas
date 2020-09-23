Attribute VB_Name = "Misc"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type MouseData
    CurrentX As Single
    CurrentY As Single
    DX As Single
    DY As Single
End Type

Public DefaultMouseL As MouseData
Public DefaultMouseR As MouseData

Public Type PuckData
    CurrentX As Single
    CurrentY As Single
    CurrentYY As Single
    AnglePI As Byte
End Type
Public DefaultPuck As PuckData
'Public LightPuckDat As PuckData

Public Type ViewPoint
    Position As D3DVECTOR
    RotationY As Single
    RotationX As Single
End Type

Public DefaultVP As ViewPoint
Public FormHasFocus As Boolean

Public Declare Function GetForegroundWindow Lib "user32" () As Long
'Public LightI As Long




Public Sub Init_ObjectTree()
    With Form1.TreeView1
        .Nodes.Add , , "Walls", "Walls"
        '.Nodes.Add , , "Lights", "Lights"
        .Nodes.Add , , "Objects", "Objects"
        .Nodes.Add , , "Floor", "Floor"
        .Nodes.Add , , "Ceiling", "Ceiling"
    End With
End Sub
