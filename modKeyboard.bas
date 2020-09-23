Attribute VB_Name = "modKeyboard"
'------------------------------
'        modKeyboard.bas
'------------------------------

Public DInput As DirectInput
Public DInputDevice As DirectInputDevice
Public Keyboard As DIKEYBOARDSTATE

Public Sub SetupDInput()
    Set DInput = DirectX.DirectInputCreate
    Set DInputDevice = DInput.CreateDevice("GUID_SysKeyboard")
    
    DInputDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DInputDevice.SetCooperativeLevel frmGame.Hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DInputDevice.Acquire  'get control of the keyboard!
End Sub

