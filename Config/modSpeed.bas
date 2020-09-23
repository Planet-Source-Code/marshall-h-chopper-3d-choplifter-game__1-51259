Attribute VB_Name = "modSpeed"
Public Declare Function GetProcessor Lib "getcpu.dll" (ByVal strCpu As String, ByVal strVendor As String, ByVal strL2Cache As String) As Long
Public Declare Function GetProcessorRawSpeed Lib "getcpu.dll" (ByVal RawSpeed As String) As Long
Public Declare Function GetProcessorNormSpeed Lib "getcpu.dll" (ByVal NormSpeed As String) As Long

Public Sub TestCPUSpeed()
    Dim sCpu As String, sVendor As String
    Dim sL2Cache As String
    Dim sRawSpeed As String
    Dim sNormSpeed As String
    
    sCpu = String(255, 0)
    sRawSpeed = String(255, 0)
    sNormSpeed = String(255, 0)
    sVendor = String(255, 0)
    sL2Cache = String(255, 0)
    
    GetProcessor sCpu, sVendor, sL2Cache
    GetProcessorNormSpeed sNormSpeed
    
    Dim CSpeed As Long
    CSpeed = Val(sNormSpeed)
    
    If CSpeed < 333 Then
        ScreenX = 320: ScreenY = 240
        Open "chopper.cfg" For Output As #1
            Print #1, "TRUE"
            Print #1, "FALSE"
            Print #1, ScreenX
            Print #1, ScreenY
            Print #1, "FALSE"
            Print #1, "FALSE"
            Print #1, "TRUE"
            Print #1, "FALSE"
        Close #1
    End If
    
    If CSpeed >= 333 And CSpeed < 600 Then
        ScreenX = 512: ScreenY = 384
        Open "chopper.cfg" For Output As #1
            Print #1, "TRUE"
            Print #1, "FALSE"
            Print #1, ScreenX
            Print #1, ScreenY
            Print #1, "FALSE"
            Print #1, "FALSE"
            Print #1, "FALSE"
            Print #1, "TRUE"
        Close #1
    End If
    
    If CSpeed > 650 Then
        ScreenX = 640: ScreenY = 480
        Open "chopper.cfg" For Output As #1
            Print #1, "TRUE"
            Print #1, "FALSE"
            Print #1, ScreenX
            Print #1, ScreenY
            Print #1, "FALSE"
            Print #1, "TRUE"
            Print #1, "FALSE"
            Print #1, "TRUE"
        Close #1
    End If
    
End Sub
    
