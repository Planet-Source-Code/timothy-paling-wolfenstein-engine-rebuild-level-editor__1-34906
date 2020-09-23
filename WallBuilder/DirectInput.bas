Attribute VB_Name = "DirectInput"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Wolfenstein Conversion
'by Timothy Paling [timothy.paling@btinternet.com]
'See instructions.txt for info
'Visit :: www.live3d.btinternet.co.uk
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim di As DirectInput8
Dim diDEV As DirectInputDevice8
Dim diState As DIKEYBOARDSTATE
Dim iKeyCounter As Integer
Public aKeys(255) As Boolean



Public Sub InitDI()
    Set di = g_DX.DirectInputCreate()
        
    If Err.Number <> 0 Then
        MsgBox "Error starting Direct Input, please make sure you have DirectX installed", vbApplicationModal
        End
    End If
        
        
    Set diDEV = di.CreateDevice("GUID_SysKeyboard")
    
    diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    diDEV.SetCooperativeLevel Form1.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    
    
    diDEV.Acquire
End Sub

Public Sub EndDI()
    diDEV.Unacquire
End Sub


Public Sub ReadKyb()
    
    For iKeyCounter = 0 To 255
        aKeys(iKeyCounter) = False
    Next

    diDEV.GetDeviceStateKeyboard diState
    
    For iKeyCounter = 0 To 255
        If diState.Key(iKeyCounter) <> 0 Then
            aKeys(iKeyCounter) = True
        End If
    Next
    DoEvents
End Sub

Public Sub PurgeKeyDat()
    For iKeyCounter = 0 To 255
        aKeys(iKeyCounter) = False
    Next
End Sub
