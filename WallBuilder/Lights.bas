Attribute VB_Name = "Lights"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public LightsDat() As D3DLIGHT8
'
'Public Function LTOpenRetFile() As String
'
'    On Error GoTo Err2
'    With Form1.CommonDialog1
'        .CancelError = True
'        .DefaultExt = "ldf"
'        .DialogTitle = "Open Light"
'        .FileName = "untitled"
'        .Filter = "Live://3D Light (*.ldf)|*.ldf"
'        .ShowOpen
'        LTOpenRetFile = .FileName
'    End With
'    Exit Function
'Err2:
'
'    LTOpenRetFile = "-1"
'
'
'End Function
'
'
'
'Public Function LoadLDF(PosX As Single, PosY As Single, PosZ As Single, TextureFile As String) As Integer
'    ReDim Preserve LightsDat(UBound(LightsDat()) + 1)
'    a = UBound(LightsDat())
'    Open TextureFile For Input As #1
'        Dim Temp As String
''        File = File & Text14.Text & vbNewLine 'AmbR
'        Line Input #1, Temp
'        'Text14.Text = Temp
'        LightsDat(a).Ambient.r = Temp
'
''        File = File & Text15.Text & vbNewLine 'AmbG
'        Line Input #1, Temp
'        'Text15.Text = Temp
'        LightsDat(a).Ambient.g = Temp
'
''        File = File & Text16.Text & vbNewLine 'AmbB
'        Line Input #1, Temp
'        'Text16.Text = Temp
'        LightsDat(a).Ambient.b = Temp
'
''        File = File & Text17.Text & vbNewLine 'DiffR
'        Line Input #1, Temp
'        'Text17.Text = Temp
'        LightsDat(a).diffuse.r = Temp
'
''        File = File & Text18.Text & vbNewLine 'DiffG
'        Line Input #1, Temp
'        'Text18.Text = Temp
'        LightsDat(a).diffuse.g = Temp
'
''        File = File & Text19.Text & vbNewLine 'DiffB
'        Line Input #1, Temp
'        'Text19.Text = Temp
'        LightsDat(a).diffuse.b = Temp
'
''        File = File & Text20.Text & vbNewLine 'SpecR
'        Line Input #1, Temp
'        'Text20.Text = Temp
'        LightsDat(a).specular.r = Temp
'
''        File = File & Text21.Text & vbNewLine 'SpecG
'        Line Input #1, Temp
'        'Text21.Text = Temp
'        LightsDat(a).specular.g = Temp
'
''        File = File & Text22.Text & vbNewLine 'SpecB
'        Line Input #1, Temp
'        'Text22.Text = Temp
'        LightsDat(a).specular.b = Temp
''
''        File = File & Text23.Text & vbNewLine 'Atten0
'        Line Input #1, Temp
'        'Text23.Text = Temp
'        LightsDat(a).Attenuation0 = Temp
'
''        File = File & Text24.Text & vbNewLine 'Atten1
'        Line Input #1, Temp
'        'Text24.Text = Temp
'        LightsDat(a).Attenuation1 = Temp
'
''        File = File & Text25.Text & vbNewLine 'Atten2
'        Line Input #1, Temp
'        'Text25.Text = Temp
'        LightsDat(a).Attenuation2 = Temp
''
''        File = File & Text29.Text & vbNewLine 'PosX
'        Line Input #1, Temp
'        'Text29.Text = Temp
'        LightsDat(a).Position.X = PosX
'
''        File = File & Text30.Text & vbNewLine 'PosY
'        Line Input #1, Temp
'        'Text30.Text = Temp
'        LightsDat(a).Position.Y = PosY
'
''        File = File & Text31.Text & vbNewLine 'PosZ
'        Line Input #1, Temp
'        'Text30.Text = Temp
'        LightsDat(a).Position.z = PosZ
'
''        File = File & Text34.Text & vbNewLine 'Range
'        Line Input #1, Temp
'        'Text34.Text = Temp
'        LightsDat(a).Range = Temp
'
''        File = File & Text26.Text & vbNewLine 'Ambient
'        'Text26.Text = Temp
'    Close #1
'    LoadLDF = a
'End Function
