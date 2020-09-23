Attribute VB_Name = "modSound"
'------------------------------
'         modSound.bas
'------------------------------

Dim DSound As DirectSound

Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX

Public ChopperBuffer As DirectSoundBuffer
Public ExplodeBuffer As DirectSoundBuffer
Public DoorBuffer As DirectSoundBuffer

Public Sub SetupDSound()
    On Local Error Resume Next
    Set DSound = DirectX.DirectSoundCreate("")

    If Err.Number <> 0 Then
        MsgBox "Unable to continue, error creating Directsound object."
        Exit Sub
    End If
    
    DSound.SetCooperativeLevel frmGame.Hwnd, DSSCL_PRIORITY
    
    Set ChopperBuffer = CreateSound("chopper.wav")
    Set ExplodeBuffer = CreateSound("explode.wav")
    Set DoorBuffer = CreateSound("door.wav")
    
    PlaySound ChopperBuffer, False, True
End Sub

Function CreateSound(filename As String) As DirectSoundBuffer
  DsDesc.lFlags = (DSBCAPS_CTRLFREQUENCY)
  Set CreateSound = DSound.CreateSoundBufferFromFile(filename, DsDesc, DsWave)
  If Err.Number <> 0 Then
    MsgBox "Unable to find sound file"
    MsgBox Err.Description
    End
  End If
End Function

Sub PlaySound(Sound As DirectSoundBuffer, CloseFirst As Boolean, LoopSound As Boolean)
  
  If CloseFirst Then
    Sound.Stop
    Sound.SetCurrentPosition 0
  End If
  If LoopSound Then
    Sound.Play DSBPLAY_LOOPING
  Else
    Sound.Play DSBPLAY_DEFAULT
  End If
End Sub

