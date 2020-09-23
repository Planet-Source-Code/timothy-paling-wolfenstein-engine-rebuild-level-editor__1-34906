VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Wolfenstein 1.5 Level Editor"
   ClientHeight    =   12030
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15825
   Icon            =   "Tut05_textures.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12030
   ScaleWidth      =   15825
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2160
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   65
      Top             =   11520
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2160
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   64
      Top             =   11160
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2160
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   63
      Top             =   10800
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2160
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   62
      Top             =   10440
      Width           =   255
   End
   Begin VB.CheckBox Check6 
      Caption         =   "IU Height"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9840
      TabIndex        =   61
      Top             =   0
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "IU Light"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8760
      TabIndex        =   60
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8160
      TabIndex        =   59
      Text            =   "1"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Height          =   495
      Left            =   1800
      Picture         =   "Tut05_textures.frx":0BC2
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6960
      Width           =   495
   End
   Begin VB.TextBox v1red 
      Height          =   285
      Left            =   1080
      TabIndex        =   48
      Text            =   "255"
      Top             =   8760
      Width           =   375
   End
   Begin VB.TextBox v2red 
      Height          =   285
      Left            =   1080
      TabIndex        =   47
      Text            =   "255"
      Top             =   9120
      Width           =   375
   End
   Begin VB.TextBox v3red 
      Height          =   285
      Left            =   1080
      TabIndex        =   46
      Text            =   "255"
      Top             =   9480
      Width           =   375
   End
   Begin VB.TextBox v4red 
      Height          =   285
      Left            =   1080
      TabIndex        =   45
      Text            =   "255"
      Top             =   9840
      Width           =   375
   End
   Begin VB.TextBox v1green 
      Height          =   285
      Left            =   1560
      TabIndex        =   44
      Text            =   "255"
      Top             =   8760
      Width           =   375
   End
   Begin VB.TextBox v1blue 
      Height          =   285
      Left            =   2040
      TabIndex        =   43
      Text            =   "255"
      Top             =   8760
      Width           =   375
   End
   Begin VB.TextBox v2green 
      Height          =   285
      Left            =   1560
      TabIndex        =   42
      Text            =   "255"
      Top             =   9120
      Width           =   375
   End
   Begin VB.TextBox v2blue 
      Height          =   285
      Left            =   2040
      TabIndex        =   41
      Text            =   "255"
      Top             =   9120
      Width           =   375
   End
   Begin VB.TextBox v3green 
      Height          =   285
      Left            =   1560
      TabIndex        =   40
      Text            =   "255"
      Top             =   9480
      Width           =   375
   End
   Begin VB.TextBox v3blue 
      Height          =   285
      Left            =   2040
      TabIndex        =   39
      Text            =   "255"
      Top             =   9480
      Width           =   375
   End
   Begin VB.TextBox v4green 
      Height          =   285
      Left            =   1560
      TabIndex        =   38
      Text            =   "255"
      Top             =   9840
      Width           =   375
   End
   Begin VB.TextBox v4blue 
      Height          =   285
      Left            =   2040
      TabIndex        =   37
      Text            =   "255"
      Top             =   9840
      Width           =   375
   End
   Begin VB.TextBox V1Y 
      Height          =   285
      Left            =   1440
      TabIndex        =   36
      Text            =   "0"
      Top             =   10440
      Width           =   615
   End
   Begin VB.TextBox V2Y 
      Height          =   285
      Left            =   1440
      TabIndex        =   35
      Text            =   "0"
      Top             =   10800
      Width           =   615
   End
   Begin VB.TextBox V3Y 
      Height          =   285
      Left            =   1440
      TabIndex        =   34
      Text            =   "0"
      Top             =   11160
      Width           =   615
   End
   Begin VB.TextBox V4Y 
      Height          =   285
      Left            =   1440
      TabIndex        =   33
      Text            =   "0"
      Top             =   11520
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Height          =   495
      Left            =   1560
      Picture         =   "Tut05_textures.frx":1004
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   1080
      Picture         =   "Tut05_textures.frx":1446
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hidden"
      Height          =   3735
      Left            =   2880
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   5895
      Begin VB.FileListBox File2 
         Height          =   870
         Left            =   2160
         Pattern         =   "*.bmp"
         TabIndex        =   31
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   2040
         Pattern         =   "*.bmp"
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1320
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   735
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         _Version        =   393217
         Indentation     =   176
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   2760
      ScaleHeight     =   8025
      ScaleWidth      =   10545
      TabIndex        =   27
      Top             =   480
      Width           =   10575
   End
   Begin VB.CommandButton Command10 
      Height          =   375
      Left            =   1920
      Picture         =   "Tut05_textures.frx":1888
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Height          =   375
      Left            =   1920
      Picture         =   "Tut05_textures.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton TextNext 
      Height          =   375
      Left            =   1320
      Picture         =   "Tut05_textures.frx":210C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   840
      Picture         =   "Tut05_textures.frx":254E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Swap Light Map"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton LightNext 
      Height          =   375
      Left            =   1320
      Picture         =   "Tut05_textures.frx":2990
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton LightSelect 
      Height          =   375
      Left            =   840
      Picture         =   "Tut05_textures.frx":2DD2
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton LightPrev 
      Height          =   375
      Left            =   240
      Picture         =   "Tut05_textures.frx":3214
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox LightPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1425
      ScaleWidth      =   2145
      TabIndex        =   17
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Height          =   495
      Left            =   1080
      Picture         =   "Tut05_textures.frx":3656
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6960
      Width           =   495
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Swap Texture "
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Height          =   495
      Left            =   1080
      Picture         =   "Tut05_textures.frx":3A98
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   1800
      Picture         =   "Tut05_textures.frx":3EDA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5760
      Width           =   2175
   End
   Begin VB.PictureBox TexturePreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1425
      ScaleWidth      =   2145
      TabIndex        =   10
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton TextPrev 
      Height          =   375
      Left            =   240
      Picture         =   "Tut05_textures.frx":431C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show guides"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Build Transparent Objects"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Render Solid"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   11040
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Render Wire"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   12360
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   360
      Picture         =   "Tut05_textures.frx":475E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   600
      Picture         =   "Tut05_textures.frx":4BA0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   360
      Picture         =   "Tut05_textures.frx":4FE2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   495
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   8640
      X2              =   8640
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Label Label12 
      Caption         =   "Puck Stop:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7200
      TabIndex        =   58
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   7080
      X2              =   7080
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00800000&
      Height          =   1575
      Left            =   120
      Top             =   10320
      Width           =   2415
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00800000&
      Height          =   1575
      Left            =   120
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Caption         =   "Vertex 1:"
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      Caption         =   "Vertex 2:"
      Height          =   255
      Left            =   240
      TabIndex        =   55
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF00FF&
      Caption         =   "Vertex 3:"
      Height          =   255
      Left            =   240
      TabIndex        =   54
      Top             =   9480
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Caption         =   "Vertex 4:"
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   9840
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 1 Y:"
      Height          =   255
      Left            =   480
      TabIndex        =   52
      Top             =   10440
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 2 Y:"
      Height          =   255
      Left            =   480
      TabIndex        =   51
      Top             =   10800
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 3 Y:"
      Height          =   255
      Left            =   480
      TabIndex        =   50
      Top             =   11160
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Vertex 4 Y:"
      Height          =   255
      Left            =   480
      TabIndex        =   49
      Top             =   11520
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   10920
      X2              =   10920
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   735
      Left            =   120
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   735
      Left            =   120
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Object Type:"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Light Map:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Base Texture:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00800000&
      Height          =   2295
      Left            =   120
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00800000&
      Height          =   1215
      Left            =   120
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800000&
      Height          =   2295
      Left            =   120
      Top             =   480
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Level"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Level"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FileIndex As Long
Dim Release As Boolean

Dim FileIndex2 As Long

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then g_D3DDevice.SetRenderState D3DRS_LIGHTING, 1 Else g_D3DDevice.SetRenderState D3DRS_LIGHTING, 0
End Sub

Private Sub Command10_Click()
    For a = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes.Item(a).Checked = True Then
            b = Right(TreeView1.Nodes.Item(a).Key, Len(TreeView1.Nodes.Item(a).Key) - 6)
            
            WorldObjects(b).LightMapFile = File2.Path & "\" & File2.List(FileIndex2 - 1)
            Ret = SearchTextures(WorldObjects(b).LightMapFile)
            If Ret = -1 Then
                Ret = AddTexture(WorldObjects(b).LightMapFile)
            End If
            WorldObjects(b).LightMapIndex = Ret
            Exit For
        End If
    Next a
End Sub

Private Sub Command11_Click()
    DefaultPuck.CurrentX = 0
    DefaultPuck.CurrentY = 0
    DefaultPuck.CurrentYY = 0
    DefaultVP.Position.X = 0
    DefaultVP.Position.Y = 0.5
    DefaultVP.Position.z = 0
End Sub

Private Sub Command12_Click()
    Form8.Show 1
End Sub

Private Sub Command2_Click()
    DefaultVP.Position.Y = DefaultVP.Position.Y + 0.1
    
End Sub

Private Sub Command3_Click()
    DefaultVP.Position.Y = DefaultVP.Position.Y - 0.1
End Sub

Private Sub Command4_Click()
Dim a As Integer
    For a = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes.Item(a).Checked = True Then
            b = Right(TreeView1.Nodes.Item(a).Key, Len(TreeView1.Nodes.Item(a).Key) - 6)
            WorldObjects(b).InUse = False
            TreeView1.Nodes.Remove WorldObjects(b).NodeIndex
            Exit For
        End If
    Next a
    Dim Index() As Integer
    For a = 1 To TreeView1.Nodes.Count
        If Len(TreeView1.Nodes.Item(a).Key) > 6 Then
            If TreeView1.Nodes.Item(a).Key <> "Objects" Then
            If Left(TreeView1.Nodes.Item(a).Key, 6) = "Object" Then
                b = Right(TreeView1.Nodes.Item(a).Key, Len(TreeView1.Nodes.Item(a).Key) - 6)
                WorldObjects(b).NodeIndex = a
            End If
            End If
        End If
    Next a
    
End Sub

Private Sub Command5_Click()
    Form4.Show
End Sub

Private Sub Command6_Click()
    Form5.Show
End Sub

Private Sub Command7_Click()
    Form6.Show
End Sub

Private Sub Command8_Click()
    Form7.Show
End Sub

Private Sub Command9_Click()
    For a = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes.Item(a).Checked = True Then
            b = Right(TreeView1.Nodes.Item(a).Key, Len(TreeView1.Nodes.Item(a).Key) - 6)
            
            WorldObjects(b).TextureFile = File1.Path & "\" & File1.List(FileIndex - 1)
            Ret = SearchTextures(WorldObjects(b).TextureFile)
            If Ret = -1 Then
                Ret = AddTexture(WorldObjects(b).TextureFile)
            End If
            WorldObjects(b).TextureIndex = Ret
            Exit For
        End If
    Next a
End Sub

Private Sub Form_GotFocus()
    FormHasFocus = True
End Sub

'-----------------------------------------------------------------------------
' Name: Form_Load()
'-----------------------------------------------------------------------------
Private Sub Form_Load()
    Dim b As Boolean
    
    ' Allow the form to become visible
    Me.Show
    DoEvents
    
    
    ' Initialize D3D and D3DDevice
    b = InitD3D(Picture1.hWnd)
    If Not b Then
        MsgBox "Unable to initialise a Direct3D compatible device. Cannot continue."
        End
    End If
    
    'Startup DirectInput
    InitDI
    
    
    ' Initialize vertex buffer with geometry and load our texture
    b = InitGeometry()
    If Not b Then
        MsgBox "Unable to Create VertexBuffer"
        End
    End If
    
    If File1.ListCount > 0 Then
        FileIndex = 1
        If Len(File1.Path) = 3 Then
            Set TexturePreview.Picture = LoadPicture(File1.Path & File1.List(FileIndex - 1))
        Else
            Set TexturePreview.Picture = LoadPicture(File1.Path & "\" & File1.List(FileIndex - 1))
        End If
    End If
    
    If File2.ListCount > 0 Then
        FileIndex2 = 1
        If Len(File1.Path) = 3 Then
            Set LightPreview.Picture = LoadPicture(File2.Path & File2.List(FileIndex2 - 1))
        Else
            Set LightPreview.Picture = LoadPicture(File2.Path & "\" & File2.List(FileIndex2 - 1))
        End If
    End If
    
    ReDim WorldObjects(0)
    ReDim MainTextureLib(0)
    ReDim LightsDat(0)
    ReDim CObjects(0)
    'Load_CObjectChain "C:\windows\desktop\chain1.coc"
    'Iface
    Combo1.AddItem "Wall"
    'Combo1.AddItem "Light"
    'Combo1.AddItem "Object"
    Combo1.AddItem "Custom"
    Combo1.AddItem "Ceiling"
    Combo1.AddItem "Floor"
    Combo1.AddItem "Sprite"
    Combo1.Text = "Wall"
    Check1.Value = 1
    Init_ObjectTree
    DefaultVP.Position.Y = 0.5
    ' Enable Timer to update
    Do
    Render
    DoEvents
    Loop
    
End Sub

Private Sub Form_LostFocus()
    FormHasFocus = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Picture1.Move 2760, 480, Me.Width - 3000, Me.Height - 1350
    'RenderSuspend = True
    
    
    'RenderSuspend = False
End Sub

'-----------------------------------------------------------------------------
' Name: Form_Unload()
'-----------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    EndDI
    Cleanup
    End
End Sub





Private Sub mnuAbout_Click()
    MsgBox "Live://3D Wolfenstein Level Editor. Copyright(c)2001 Tim Paling.", vbInformation, "Wolfenstein Editor"
End Sub

Private Sub mnuOpen_Click()
    LoadWorld
End Sub

Private Sub mnuSave_Click()
    SaveWorld
End Sub

'----------------------------------------------------------------------------
' Texture Preview and Selection
'----------------------------------------------------------------------------
Private Sub LightSelect_Click()
    Temp$ = File1.Path
    Form2.Show 1
    File2.Path = File1.Path
    File1.Path = Temp$
    If File2.ListCount > 0 Then
        FileIndex2 = 1
        If Len(File2.Path) = 3 Then
            Set LightPreview.Picture = LoadPicture(File2.Path & File2.List(FileIndex2 - 1))
        Else
            Set LightPreview.Picture = LoadPicture(File2.Path & "\" & File2.List(FileIndex2 - 1))
        End If
    Else
        Set LightPreview.Picture = Nothing
    End If
End Sub

Private Sub LightNext_Click()
    If File2.ListCount <> 0 Then
        FileIndex2 = FileIndex2 + 1
        If FileIndex2 > File2.ListCount Then FileIndex2 = File2.ListCount
        If Len(File1.Path) = 3 Then
            Set LightPreview.Picture = LoadPicture(File2.Path & File2.List(FileIndex2 - 1))
        Else
            Set LightPreview.Picture = LoadPicture(File2.Path & "\" & File2.List(FileIndex2 - 1))
        End If
    End If
End Sub

Private Sub LightPrev_Click()
    If File2.ListCount <> 0 Then
        FileIndex2 = FileIndex2 - 1
        If FileIndex2 < 1 Then FileIndex2 = 1
        If Len(File2.Path) = 3 Then
            Set LightPreview.Picture = LoadPicture(File2.Path & File2.List(FileIndex2 - 1))
            Label2.Caption = "Lightmap: " & File2.List(FileIndex2 - 1)
        Else
            Set LightPreview.Picture = LoadPicture(File2.Path & "\" & File2.List(FileIndex2 - 1))
            Label2.Caption = "Lightmap: " & File2.List(FileIndex2 - 1)
        End If
    End If

End Sub

Private Sub TextNext_Click()
    If File1.ListCount <> 0 Then
        FileIndex = FileIndex + 1
        If FileIndex > File1.ListCount Then FileIndex = File1.ListCount
        If Len(File1.Path) = 3 Then
            Set TexturePreview.Picture = LoadPicture(File1.Path & File1.List(FileIndex - 1))
            Label1.Caption = "Base Texture: " & File1.List(FileIndex - 1)
        Else
            Set TexturePreview.Picture = LoadPicture(File1.Path & "\" & File1.List(FileIndex - 1))
            Label1.Caption = "Base Texture: " & File1.List(FileIndex - 1)
        End If
    End If
End Sub


Private Sub TextPrev_Click()
    If File1.ListCount <> 0 Then
        FileIndex = FileIndex - 1
        If FileIndex < 1 Then FileIndex = 1
        If Len(File1.Path) = 3 Then
            Set TexturePreview.Picture = LoadPicture(File1.Path & File1.List(FileIndex - 1))
            Label1.Caption = "Base Texture: " & File1.List(FileIndex - 1)
        Else
            Set TexturePreview.Picture = LoadPicture(File1.Path & "\" & File1.List(FileIndex - 1))
            Label1.Caption = "Base Texture: " & File1.List(FileIndex - 1)
        End If
    End If
End Sub

Private Sub Command1_Click()
    Form2.Show 1
    If File1.ListCount > 0 Then
        FileIndex = 1
        If Len(File1.Path) = 3 Then
            Set TexturePreview.Picture = LoadPicture(File1.Path & File1.List(FileIndex - 1))
        Else
            Set TexturePreview.Picture = LoadPicture(File1.Path & "\" & File1.List(FileIndex - 1))
        End If
    Else
        Set TexturePreview.Picture = Nothing
    End If
End Sub

'-------------------------------------------------------------------
' Puck Movement
' ------------------------------------------------------------------

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        DefaultMouseL.DX = DefaultMouseL.CurrentX - X
        DefaultMouseL.DY = DefaultMouseL.CurrentY - Y
        DefaultMouseL.CurrentX = X
        DefaultMouseL.CurrentY = Y
        
    ElseIf Button = 2 Then
        DefaultMouseR.DX = DefaultMouseR.CurrentX - X
        DefaultMouseR.DY = DefaultMouseR.CurrentY - Y
        DefaultMouseR.CurrentX = X
        DefaultMouseR.CurrentY = Y
    End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    
    If Button = 2 Then
        If Form4.Text15.Text = "#" Or (X <> 0 And Y <> 0) Then
            MB2
        End If
    ElseIf Button = 1 Then
        If Form4.Text14.Text = "#" Or (X <> 0 And Y <> 0) Then
            MB1
         End If
    End If
End Sub
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Release = True
End Sub

'-------------------------------------------------------------------
' Interface
'-------------------------------------------------------------------
Private Sub Option1_Click()
    g_D3DDevice.SetRenderState D3DRS_FILLMODE, 3
End Sub

Private Sub Option2_Click()
    g_D3DDevice.SetRenderState D3DRS_FILLMODE, 2

End Sub

Private Sub ZoomOut_Click()

End Sub

Public Sub Invoke_Delete()
    Command4_Click
End Sub

Public Sub Invoke_MouseDown(Button As Integer)
    'Picture1_MouseDown Button, 0, 0, 0
    Select Case Button
        Case 1
            MB1
        Case 2
            MB2
    End Select
End Sub

Public Sub Invoke_MouseUp()
    Picture1_MouseUp 0, 0, 0, 0
    
End Sub

Public Sub MB1()
    Dim TexFileName As String
    'Dim LightFileName As String
    For a = 1 To UBound(WorldObjects())
            
                With WorldObjects(a)
                        If .WorldX = DefaultPuck.CurrentX And .WorldZ = DefaultPuck.CurrentY And .WorldY = DefaultPuck.CurrentYY And .Rotation = DefaultPuck.AnglePI And .InUse = True And .CodeWord = Combo1.Text Then
                            'Fail build
                            Exit Sub
                        End If
                End With
                
            Next a
            If (File1.ListCount > 0 And File2.ListCount > 0) Or Combo1.Text = "Custom" Then
                Select Case Combo1.Text
                    Case "Sprite"
                        TexFileName = OpenRetFile
                        If TexFileName <> "-1" Then
                        'Sprites need special DDS alpha channel textures...
                            Add_Object Combo1.Text, TexFileName, File2.Path & "\" & File2.List(FileIndex2 - 1), D3DFILL_SOLID, vec3(DefaultPuck.CurrentX, DefaultPuck.CurrentYY, DefaultPuck.CurrentY), DefaultPuck.AnglePI
                        End If
                   Case "Custom"
                        
                        TexFileName = OpenRetFileEx("L3D Object Chain|*.loc", "loc", "Open L3D Object Chain")
                        If TexFileName <> "-1" Then
                        Add_Object Combo1.Text, TexFileName, vbNullString, D3DFILL_SOLID, vec3(0, 0, 0), 0
                        End If
                    Case Else
                        If Check2.Value = 1 Then
                            TexFileName = OpenRetFile
                            If TexFileName <> "-1" Then
                                'Sprites need special DDS alpha channel textures...
                                Add_Object Combo1.Text, TexFileName, File2.Path & "\" & File2.List(FileIndex2 - 1), D3DFILL_SOLID, vec3(DefaultPuck.CurrentX, DefaultPuck.CurrentYY, DefaultPuck.CurrentY), DefaultPuck.AnglePI
                            End If
                        Else
                            Add_Object Combo1.Text, File1.Path & "\" & File1.List(FileIndex - 1), File2.Path & "\" & File2.List(FileIndex2 - 1), D3DFILL_SOLID, vec3(DefaultPuck.CurrentX, DefaultPuck.CurrentYY, DefaultPuck.CurrentY), DefaultPuck.AnglePI
                        End If
                End Select
            Else
                MsgBox "A base texture AND light map needs to be selected in order to build a wall.", vbExclamation, "Error"
                
            End If
End Sub

Public Sub MB2()
    If Release = True Then
        If DefaultPuck.AnglePI = 1 Then DefaultPuck.AnglePI = 0 Else DefaultPuck.AnglePI = 1
        Release = False
    End If
End Sub
