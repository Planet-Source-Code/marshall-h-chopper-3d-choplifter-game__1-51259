Attribute VB_Name = "modJoystick"
'------------------------------
'        modJoystick.bas
'------------------------------

Option Explicit

'Programmer: Raymond L. King
'Custom Software Designers

'-- General Constants.
Public Const MAXPNAMELEN = 32  ' Max Product Name Length (Including NULL)
Public Const MAXOEMVXD = 160

'-- Joystick ID Constants.
Public Const JOYSTICKID1 = 0
Public Const JOYSTICKID2 = 1

'-- Joystick Flag Constants.
Public Const JOY_CAL_READ3 = &H40000
Public Const JOY_CAL_READ4 = &H80000
Public Const JOY_CAL_READ5 = &H400000
Public Const JOY_CAL_READ6 = &H800000
Public Const JOY_CAL_READALWAYS = &H10000
Public Const JOY_CAL_READRONLY = &H2000000
Public Const JOY_CAL_READUONLY = &H4000000
Public Const JOY_CAL_READVONLY = &H8000000
Public Const JOY_CAL_READXONLY = &H100000
Public Const JOY_CAL_READXYONLY = &H20000
Public Const JOY_CAL_READYONLY = &H200000
Public Const JOY_CAL_READZONLY = &H1000000
Public Const JOY_POVBACKWARD = 18000
Public Const JOY_POVCENTERED = -1
Public Const JOY_POVFORWARD = 0
Public Const JOY_POVLEFT = 27000
Public Const JOY_POVRIGHT = 9000
Public Const JOY_RETURNBUTTONS = &H80&
Public Const JOY_RETURNCENTERED = &H400&
Public Const JOY_RETURNPOV = &H40&
Public Const JOY_RETURNPOVCTS = &H200&
Public Const JOY_RETURNR = &H8&
Public Const JOY_RETURNRAWDATA = &H100&
Public Const JOY_RETURNU = &H10         '  Axis 5
Public Const JOY_RETURNV = &H20         '  Axis 6
Public Const JOY_RETURNX = &H1&
Public Const JOY_RETURNY = &H2&
Public Const JOY_RETURNZ = &H4&
Public Const JOY_USEDEADZONE = &H800&
Public Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ _
                           Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV _
                           Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)

'-- JoyStick Error Constants.
Public Const JOYERR_BASE = 160                    ' Error Base
Public Const JOYERR_NOCANDO = (JOYERR_BASE + 6)   ' Request Not Completed
Public Const JOYERR_NOERROR = (0)                 ' No Error
Public Const JOYERR_PARMS = (JOYERR_BASE + 5)     ' Bad Parameters
Public Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7) ' JoyStick Is Unplugged

'-- JOYCAPS User Defined Type.
Public Type JOYCAPS
  wMid            As Integer  ' Manufacturer Identifier.
  wPid            As Integer  ' Product Identifier.
  szPname         As String * MAXPNAMELEN ' Null Terminated String - JoyStick Product Name.
  wXmin           As Long     ' Minimum X-coordinate.
  wXmax           As Long     ' Maximum X-coordinate.
  wYmin           As Long     ' Minimum Y-coordinate.
  wYmax           As Long     ' Maximum Y-coordinate.
  wZmin           As Long     ' Minimum Z-coordinate.
  wZmax           As Long     ' Maximum Z-coordinate.
  wNumButtons     As Long     ' Number Of JoyStick Buttons.
  wPeriodMin      As Long     ' Smallest Polling Frequency Supported By JoySetCapture.
  wPeriodMax      As Long     ' Largest Polling Frequency Supported By JoySetCapture.
  wRmin           As Long     ' Minimum Rudder Value. The Rudder Is A Fourth Axis Movement.
  wRmax           As Long     ' Maximum Rudder Value. The Rudder Is A Fourth Axis Movement.
  wUmin           As Long     ' Minimum U-coordinate (Fifth Axis) Values.
  wUmax           As Long     ' Maximum U-coordinate (Fifth Axis) Values.
  wVmin           As Long     ' Minimum V-coordinate (Sixth Axis) Values.
  wVmax           As Long     ' Maximum V-coordinate (Sixth Axis) Values.
  wCaps           As Long     ' JoyStick Capabilities.  Note: See JoyCaps Flags Below...
  wMaxAxes        As Long     ' Maximum Number Of Axes Supported By JoyStick.
  wNumAxes        As Long     ' Number Of Axes Currently In Use By JoyStick.
  wMaxButtons     As Long     ' Maximum Number Of Buttons Supported By The JoyStick.
  szRegKey        As String * MAXPNAMELEN ' Null-Terminated String Containing The Registry Key.
  szOEMVxD        As String * MAXOEMVXD ' Null-Terminated String Identifying The JoyStick Driver OEM.
End Type

'-- JOYINFOEX User Defined Type.
Public Type JOYINFOEX
  dwSize          As Long     ' Size, In Bytes, Of This User Defined Type.
  dwFlags         As Long     ' Flags See Below: JOYINFOEX Flags.
  dwXpos          As Long     ' Current X-coordinate.
  dwYpos          As Long     ' Current Y-coordinate.
  dwZpos          As Long     ' Current Z-coordinate.
  dwRpos          As Long     ' Current Position Of The Rudder Or Fourth JoyStick Axis.
  dwUpos          As Long     ' Current Fifth Axis Position.
  dwVpos          As Long     ' Current Sixth Axis Position.
  dwButtons       As Long     ' Current State Of The 32 JoyStick Buttons.
  dwButtonNumber  As Long     ' Current Button Number That Is Pressed.
  dwPOV           As Long     ' Current Position Of The Point-Of-View Control.
  dwReserved1     As Long     ' Reserved; Do Not Use.
  dwReserved2     As Long     ' Reserved; Do Not Use.
End Type

'-- For Accessing The JoyStick User Defined Types.
Public JsCaps As JOYCAPS
Public JsInfo As JOYINFOEX

'-- JoyStick API Declarations.
Private Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" _
  (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long

Private Declare Function joyGetNumDevs Lib "winmm.dll" () As Long

Private Declare Function joyGetPosEx Lib "winmm.dll" _
  (ByVal uJoyID As Long, pji As JOYINFOEX) As Long

Private Declare Function joyGetThreshold Lib "winmm.dll" _
  (ByVal id As Long, lpuThreshold As Long) As Long

Private Declare Function joyReleaseCapture Lib "winmm.dll" (ByVal id As Long) As Long

Private Declare Function joySetCapture Lib "winmm.dll" _
  (ByVal Hwnd As Long, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) _
  As Long

Private Declare Function joySetThreshold Lib "winmm.dll" _
  (ByVal id As Long, ByVal uThreshold As Long) As Long

'-- Initialize The JoyStick Structures
Public Sub JoyStick_Init(JoyStick As Integer)

  Dim lRtn   As Long
  Dim luSize As Long
  
  '-- Size Of User Defined Type
  luSize = Len(JsCaps)
  
  '-- Get The JoyStick Capabilities...
  lRtn = joyGetDevCaps(JoyStick, JsCaps, luSize)
    joySetThreshold lRtn, 32768
  '-- Check For Error...
  If lRtn <> JOYERR_NOERROR Then
    'MsgBox "JoyStick Initialization Error! " & Str(lRtn)
  End If
    
End Sub

'-- Get JoyStick Position
Public Sub JoyStick_GetPos(JoyStick As Integer)

  Dim lRtn  As Long
  Dim lSize As Long
  
  '-- Size Of User Defined Type
  lSize = Len(JsInfo)
  
  JsInfo.dwSize = lSize
  
  '-- Set Flag To Return All
  JsInfo.dwFlags = JOY_RETURNALL
  
  '-- Get JotStick Position
  lRtn = joyGetPosEx(JoyStick, JsInfo)
  
  '-- Check For An Error
  If lRtn <> JOYERR_NOERROR Then
    MsgBox "A Joystick Error Has Occurred! " & Str(lRtn)
  End If
  
End Sub


